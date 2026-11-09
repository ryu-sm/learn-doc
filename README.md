
Option Explicit

' ============================================================
' XLSM Sheet Dependency Analyzer
'
' Source workbook: READ ONLY
' Output: new XLSX report
'
' Supports:
'   - Visible / Hidden / VeryHidden sheets
'   - Explicit formula references
'   - INDIRECT dynamic sheet names
'   - INDEX(INDIRECT(...))
'   - OFFSET cross-sheet references
'   - Defined names (basic)
'   - Shape.OnAction and hyperlinks
'   - VBA procedures and navigation
'   - Direct dependency matrices
'   - Evidence and unresolved details
'
' No source macro execution.
' ============================================================

Private gSource As Workbook
Private gReport As Workbook

Private gSheets As Object
Private gData As Object
Private gNav As Object
Private gEvidence As Object
Private gUnresolved As Object
Private gProcedures As Object
Private gProcTargets As Object
Private gCodeNames As Object

Private gEvidenceRows As Collection
Private gUnresolvedRows As Collection
Private gIndirectRows As Collection

Private gFormulaCount As Long
Private gIndirectCount As Long
Private gIndirectResolved As Long
Private gShapeCount As Long

Private gOldCalc As XlCalculation
Private gOldEvents As Boolean
Private gOldScreen As Boolean
Private gOldAlerts As Boolean
Private gOldSecurity As Long
Private gOldLinks As Boolean

Private Const SEC_FORCE_DISABLE As Long = 3

' ============================================================
' ENTRY POINT
' ============================================================

Public Sub AnalyzeXlsmDependencies()

    Dim selected As Variant
    Dim outputPath As String
    Dim completed As Boolean
    Dim errorText As String

    selected = Application.GetOpenFilename( _
        "Excel Macro Files (*.xlsm),*.xlsm", _
        , "Select XLSM to analyze")

    If VarType(selected) = vbBoolean Then Exit Sub

    InitAnalyzer

    gOldCalc = Application.Calculation
    gOldEvents = Application.EnableEvents
    gOldScreen = Application.ScreenUpdating
    gOldAlerts = Application.DisplayAlerts
    gOldSecurity = Application.AutomationSecurity
    gOldLinks = Application.AskToUpdateLinks

    On Error GoTo Failure

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = SEC_FORCE_DISABLE
    Application.Calculation = xlCalculationManual

    Set gSource = Workbooks.Open( _
        Filename:=CStr(selected), _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        AddToMru:=False, _
        IgnoreReadOnlyRecommended:=True)

    RegisterSheets

    Debug.Print "Scanning formulas..."
    ScanFormulas

    Debug.Print "Scanning defined names..."
    ScanDefinedNames

    Debug.Print "Scanning VBA..."
    ScanVba

    Debug.Print "Scanning shapes..."
    ScanShapes

    Debug.Print "Building report..."
    Set gReport = Workbooks.Add(xlWBATWorksheet)

    BuildReport

    outputPath = MakeOutputPath(CStr(selected))

    gReport.SaveAs _
        Filename:=outputPath, _
        FileFormat:=xlOpenXMLWorkbook

    completed = True

CleanExit:

    On Error Resume Next

    If Not gSource Is Nothing Then
        gSource.Close SaveChanges:=False
    End If

    If Not gReport Is Nothing Then
        If Not completed Then
            gReport.Close SaveChanges:=False
        End If
    End If

    Application.Calculation = gOldCalc
    Application.AutomationSecurity = gOldSecurity
    Application.AskToUpdateLinks = gOldLinks
    Application.DisplayAlerts = gOldAlerts
    Application.EnableEvents = gOldEvents
    Application.ScreenUpdating = gOldScreen
    Application.StatusBar = False

    On Error GoTo 0

    If completed Then

        MsgBox _
            "Analysis completed." & vbCrLf & _
            "Sheets: " & gSheets.Count & vbCrLf & _
            "Formula cells: " & gFormulaCount & vbCrLf & _
            "INDIRECT calls: " & gIndirectCount & vbCrLf & _
            "INDIRECT resolved: " & gIndirectResolved & vbCrLf & _
            "Shapes: " & gShapeCount & vbCrLf & _
            "Evidence: " & gEvidenceRows.Count & vbCrLf & _
            "Unresolved: " & gUnresolvedRows.Count & vbCrLf & _
            "Report: " & outputPath, _
            vbInformation

    ElseIf Len(errorText) > 0 Then

        MsgBox errorText, vbCritical

    End If

    Exit Sub

Failure:

    errorText = "Analysis failed: " & _
                Err.Number & " - " & Err.Description

    Resume CleanExit

End Sub

' ============================================================
' INITIALIZATION
' ============================================================

Private Sub InitAnalyzer()

    Set gSheets = NewDict()
    Set gData = NewDict()
    Set gNav = NewDict()
    Set gEvidence = NewDict()
    Set gUnresolved = NewDict()
    Set gProcedures = NewDict()
    Set gProcTargets = NewDict()
    Set gCodeNames = NewDict()

    Set gEvidenceRows = New Collection
    Set gUnresolvedRows = New Collection
    Set gIndirectRows = New Collection

    gFormulaCount = 0
    gIndirectCount = 0
    gIndirectResolved = 0
    gShapeCount = 0

End Sub

Private Function NewDict() As Object

    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = vbTextCompare

End Function

Private Sub RegisterSheets()

    Dim ws As Worksheet

    For Each ws In gSource.Worksheets

        gSheets(ws.Name) = ws.Name
        gCodeNames(ws.CodeName) = ws.Name

    Next ws

End Sub

Private Function NormalizeSheet(ByVal name As String) As String

    name = Trim$(name)

    If gSheets.Exists(name) Then
        NormalizeSheet = CStr(gSheets(name))
    Else
        NormalizeSheet = ""
    End If

End Function

' ============================================================
' RELATIONSHIP STORAGE
' ============================================================

Private Sub AddLink( _
    ByVal source As String, _
    ByVal target As String, _
    ByVal kind As String, _
    ByVal reason As String, _
    ByVal location As String, _
    ByVal expression As String)

    Dim graph As Object
    Dim targets As Object
    Dim key As String

    source = NormalizeSheet(source)
    target = NormalizeSheet(target)

    If source = "" Or target = "" Then Exit Sub

    If kind = "Navigation" Then
        Set graph = gNav
    Else
        Set graph = gData
    End If

    If Not graph.Exists(source) Then
        Set targets = NewDict()
        graph.Add source, targets
    Else
        Set targets = graph(source)
    End If

    targets(target) = True

    key = source & ChrW(30) & target & ChrW(30) & _
          kind & ChrW(30) & reason & ChrW(30) & _
          location & ChrW(30) & expression

    If Not gEvidence.Exists(key) Then

        gEvidence(key) = True

        gEvidenceRows.Add Array( _
            source, target, kind, reason, _
            location, expression)

    End If

End Sub

Private Sub AddUnresolved( _
    ByVal source As String, _
    ByVal location As String, _
    ByVal reason As String, _
    ByVal expression As String)

    Dim key As String

    key = source & ChrW(30) & location & _
          ChrW(30) & reason & ChrW(30) & expression

    If Not gUnresolved.Exists(key) Then

        gUnresolved(key) = True

        gUnresolvedRows.Add Array( _
            source, location, reason, expression)

    End If

End Sub

Private Function HasLink( _
    ByVal graph As Object, _
    ByVal source As String, _
    ByVal target As String) As Boolean

    Dim targets As Object

    If Not graph.Exists(source) Then Exit Function

    Set targets = graph(source)

    HasLink = targets.Exists(target)

End Function

' ============================================================
' REGULAR EXPRESSION HELPERS
' ============================================================

Private Function RegexMatches( _
    ByVal text As String, _
    ByVal pattern As String) As Object

    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = pattern

    Set RegexMatches = re.Execute(text)

End Function

Private Function RegexReplace( _
    ByVal text As String, _
    ByVal pattern As String, _
    ByVal replacement As String) As String

    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.Pattern = pattern

    RegexReplace = re.Replace(text, replacement)

End Function

' ============================================================
' EXTRACT EXPLICIT SHEET REFERENCES
' ============================================================

Private Function ExtractSheetRefs( _
    ByVal expression As String) As Object

    Dim result As Object
    Dim matches As Object
    Dim m As Object
    Dim target As String
    Dim cleaned As String

    Set result = NewDict()

    ' Remove Excel string literals.
    ' Quoted worksheet names are single-quoted,
    ' so they remain available for matching.
    cleaned = RegexReplace( _
        expression, _
        """(?:[^""]|"""")*""", _
        "")

    ' 'Sheet Name'!A1 or 'Sheet Name'!$A$1
    Set matches = RegexMatches( _
        cleaned, _
        "'(([^']|'')+)'!")

    For Each m In matches

        target = m.SubMatches(0)
        target = Replace(target, "''", "'")
        target = NormalizeSheet(target)

        If target <> "" Then result(target) = True

    Next m

    ' SheetName!A1
    Set matches = RegexMatches( _
        cleaned, _
        "(^|[^A-Za-z0-9_.'\]])" & _
        "([A-Za-z_0-9" & ChrW(128) & "-" & _
        ChrW(65535) & "]+)!")

    For Each m In matches

        target = NormalizeSheet(CStr(m.SubMatches(1)))

        If target <> "" Then result(target) = True

    Next m

    Set ExtractSheetRefs = result

End Function

' ============================================================
' BALANCED FUNCTION EXTRACTION
' ============================================================

Private Function FunctionArguments( _
    ByVal formula As String, _
    ByVal functionName As String) As Collection

    Dim result As New Collection
    Dim matches As Object
    Dim m As Object
    Dim startPos As Long
    Dim i As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim ch As String
    Dim body As String

    Set matches = RegexMatches( _
        formula, "\b" & functionName & "\s*\(")

    For Each m In matches

        startPos = m.FirstIndex + m.Length + 1
        depth = 1
        inString = False

        i = startPos

        Do While i <= Len(formula)

            ch = Mid$(formula, i, 1)

            If ch = """" Then

                If inString And i < Len(formula) Then

                    If Mid$(formula, i + 1, 1) = """" Then
                        i = i + 2
                        GoTo NextCharacter
                    End If

                End If

                inString = Not inString

            ElseIf Not inString Then

                If ch = "(" Then depth = depth + 1

                If ch = ")" Then

                    depth = depth - 1

                    If depth = 0 Then

                        body = Mid$( _
                            formula, startPos, i - startPos)

                        result.Add body

                        Exit Do

                    End If

                End If

            End If

            i = i + 1

NextCharacter:

        Loop

    Next m

    Set FunctionArguments = result

End Function

Private Function FirstArgument(ByVal body As String) As String

    Dim i As Long
    Dim depth As Long
    Dim inString As Boolean
    Dim ch As String

    For i = 1 To Len(body)

        ch = Mid$(body, i, 1)

        If ch = """" Then
            If i = 1 Or Mid$(body, i - 1, 1) <> """" Then
                inString = Not inString
            End If
        End If

        If Not inString Then

            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1

            If ch = "," And depth = 0 Then

                FirstArgument = Trim$(Left$(body, i - 1))
                Exit Function

            End If

        End If

    Next i

    FirstArgument = Trim$(body)

End Function

' ============================================================
' EVALUATE DYNAMIC SHEET EXPRESSIONS
' ============================================================

Private Function EvaluateText( _
    ByVal ws As Worksheet, _
    ByVal expression As String) As String

    Dim result As Variant

    On Error GoTo Failed

    expression = Trim$(expression)

    If Left$(expression, 1) <> "=" Then
        expression = "=" & expression
    End If

    ' Worksheet.Evaluate provides worksheet-relative context.
    ' Does not execute VBA procedures.
    result = ws.Evaluate(expression)

    If IsError(result) Then GoTo Failed
    If IsArray(result) Then GoTo Failed
    If IsObject(result) Then GoTo Failed
    If IsNull(result) Then GoTo Failed
    If IsEmpty(result) Then GoTo Failed

    EvaluateText = CStr(result)

    Exit Function

Failed:

    EvaluateText = ""

End Function

Private Function SheetFromAddress( _
    ByVal addressText As String) As String

    Dim p As Long
    Dim prefix As String

    addressText = Trim$(addressText)

    If Left$(addressText, 1) = "#" Then
        addressText = Mid$(addressText, 2)
    End If

    p = InStrRev(addressText, "!")

    If p = 0 Then Exit Function

    prefix = Left$(addressText, p - 1)

    If Left$(prefix, 1) = "'" And _
       Right$(prefix, 1) = "'" Then

        prefix = Mid$(prefix, 2, Len(prefix) - 2)
        prefix = Replace(prefix, "''", "'")

    End If

    ' Reject external workbook references.
    If InStr(prefix, "[") > 0 Then Exit Function

    SheetFromAddress = NormalizeSheet(prefix)

End Function

Private Function SheetFromCellValue( _
    ByVal ws As Worksheet, _
    ByVal expression As String) As String

    Dim matches As Object
    Dim m As Object
    Dim refSheet As Worksheet
    Dim refAddress As String
    Dim value As Variant
    Dim target As String

    expression = Trim$(expression)

    ' Current-sheet cell references such as $E36.
    Set matches = RegexMatches( _
        expression, _
        "^\$?[A-Z]{1,3}\$?[0-9]+$")

    If matches.Count > 0 Then

        On Error Resume Next
        value = ws.Range(expression).Value2
        On Error GoTo 0

        If Not IsError(value) Then

            If Not IsEmpty(value) Then

                target = NormalizeSheet(CStr(value))

                If target <> "" Then
                    SheetFromCellValue = target
                    Exit Function
                End If

            End If

        End If

    End If

    ' Explicitly qualified sheet!cell references.
    Set matches = RegexMatches( _
        expression, _
        "('([^']|'')+'|[^ !&()+,]+)!" & _
        "(\$?[A-Z]{1,3}\$?[0-9]+)")

    For Each m In matches

        target = SheetFromAddress(CStr(m.Value))

        If target <> "" Then

            Set refSheet = gSource.Worksheets(target)

            refAddress = CStr(m.SubMatches(2))

            On Error Resume Next
            value = refSheet.Range(refAddress).Value2
            On Error GoTo 0

            If Not IsError(value) Then

                If Not IsEmpty(value) Then

                    SheetFromCellValue = _
                        NormalizeSheet(CStr(value))

                    If SheetFromCellValue <> "" Then
                        Exit Function
                    End If

                End If

            End If

        End If

    Next m

End Function

Private Function DynamicCandidates( _
    ByVal ws As Worksheet, _
    ByVal expression As String) As Object

    Dim result As Object
    Dim matches As Object
    Dim m As Object
    Dim target As String
    Dim literal As String
    Dim name As Variant

    Set result = NewDict()

    ' Evaluate cell references that supply a sheet name.
    Set matches = RegexMatches( _
        expression, _
        "('([^']|'')+'|[^ !&()+,]+)!" & _
        "(\$?[A-Z]{1,3}\$?[0-9]+)" & _
        "|(^|[^A-Za-z0-9_!])" & _
        "(\$?[A-Z]{1,3}\$?[0-9]+)")

    For Each m In matches

        target = SheetFromCellValue(ws, CStr(m.Value))

        If target <> "" Then result(target) = True

    Next m

    ' Find literal worksheet names in strings containing '!'.
    Set matches = RegexMatches( _
        expression, _
        """(?:[^""]|"""")*""")

    For Each m In matches

        literal = Replace(CStr(m.Value), """", "")

        If InStr(literal, "!") > 0 Then

            For Each name In gSheets.Keys

                If InStr(1, literal, _
                    CStr(name) & "!", vbTextCompare) > 0 Then

                    result(CStr(name)) = True

                End If

            Next name

        End If

    Next m

    Set DynamicCandidates = result

End Function

' ============================================================
' FORMULA ANALYSIS
' ============================================================

Private Sub ScanFormulas()

    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim area As Range
    Dim cell As Range
    Dim formula As String
    Dim targets As Object
    Dim target As Variant
    Dim calls As Collection
    Dim body As Variant
    Dim argument As String
    Dim evaluated As String
    Dim resolved As String
    Dim candidates As Object
    Dim status As String
    Dim detail As String
    Dim location As String

    For Each ws In gSource.Worksheets

        Application.StatusBar = _
            "Scanning formulas: " & ws.Name

        Set formulaCells = Nothing

        On Error Resume Next
        Set formulaCells = ws.UsedRange.SpecialCells( _
            xlCellTypeFormulas)
        On Error GoTo 0

        If formulaCells Is Nothing Then GoTo NextSheet

        For Each area In formulaCells.Areas

            For Each cell In area.Cells

                gFormulaCount = gFormulaCount + 1

                formula = CStr(cell.Formula)
                location = ws.Name & "!" & cell.Address(False, False)

                ' 1. Explicit references
                Set targets = ExtractSheetRefs(formula)

                For Each target In targets.Keys

                    AddLink ws.Name, CStr(target), _
                        "Data", "Explicit Formula", _
                        location, formula

                Next target

                ' 2. INDIRECT
                Set calls = FunctionArguments(formula, "INDIRECT")

                For Each body In calls

                    gIndirectCount = gIndirectCount + 1

                    argument = FirstArgument(CStr(body))

                    evaluated = EvaluateText(ws, argument)

                    resolved = SheetFromAddress(evaluated)

                    status = "Unresolved"
                    detail = ""

                    If resolved <> "" Then

                        AddLink ws.Name, resolved, _
                            "Data", "INDIRECT Evaluated", _
                            location, formula

                        status = "Resolved"
                        detail = resolved

                        gIndirectResolved = gIndirectResolved + 1

                    Else

                        Set candidates = _
                            DynamicCandidates(ws, argument)

                        If candidates.Count > 0 Then

                            For Each target In candidates.Keys

                                AddLink ws.Name, CStr(target), _
                                    "Data", _
                                    "INDIRECT Sheet Candidate", _
                                    location, formula

                            Next target

                            status = "Candidate"
                            detail = Join(candidates.Keys, ", ")

                        Else

                            AddUnresolved ws.Name, location, _
                                "INDIRECT target unresolved", _
                                formula

                        End If

                    End If

                    gIndirectRows.Add Array( _
                        ws.Name, cell.Address(False, False), _
                        status, detail, argument, evaluated, formula)

                Next body

                ' 3. OFFSET
                ' Explicit cross-sheet references have already
                ' been collected above. A local OFFSET does not
                ' automatically imply another worksheet.
                If InStr(1, formula, "OFFSET(", _
                    vbTextCompare) > 0 Then

                    If targets.Count = 0 Then

                        AddUnresolved ws.Name, location, _
                            "OFFSET without explicit sheet target", _
                            formula

                    End If

                End If

                ' 4. HYPERLINK formulas
                Set calls = FunctionArguments(formula, "HYPERLINK")

                For Each body In calls

                    argument = FirstArgument(CStr(body))
                    evaluated = EvaluateText(ws, argument)
                    resolved = SheetFromAddress(evaluated)

                    If resolved <> "" Then

                        AddLink ws.Name, resolved, _
                            "Navigation", "HYPERLINK Formula", _
                            location, formula

                    Else

                        AddUnresolved ws.Name, location, _
                            "HYPERLINK target unresolved", formula

                    End If

                Next body

                If gFormulaCount Mod 10000 = 0 Then

                    Debug.Print "Formula cells: "; gFormulaCount
                    DoEvents

                End If

            Next cell

        Next area

NextSheet:

    Next ws

End Sub

' ============================================================
' DEFINED NAMES
' ============================================================

Private Sub ScanDefinedNames()

    Dim nm As Name
    Dim targets As Object
    Dim target As Variant
    Dim owner As String

    For Each nm In gSource.Names

        Set targets = ExtractSheetRefs(nm.RefersTo)

        owner = ""

        On Error Resume Next
        owner = nm.Parent.Name
        On Error GoTo 0

        ' Workbook-level names have no single source sheet.
        ' They are recorded as diagnostics, not assigned
        ' to every sheet automatically.
        If NormalizeSheet(owner) <> "" Then

            For Each target In targets.Keys

                AddLink owner, CStr(target), _
                    "Data", "Sheet Defined Name", _
                    nm.Name, nm.RefersTo

            Next target

        ElseIf targets.Count > 0 Then

            AddUnresolved "Workbook", nm.Name, _
                "Workbook name: consumers not traced", _
                nm.RefersTo

        End If

    Next nm

End Sub

' ============================================================
' VBA SOURCE ANALYSIS
' ============================================================

Private Sub ScanVba()

    Dim vbProject As Object
    Dim component As Object
    Dim codeModule As Object
    Dim lineNo As Long
    Dim lineText As String
    Dim proc As String
    Dim moduleName As String
    Dim match As Object
    Dim matches As Object
    Dim targets As Object
    Dim target As Variant
    Dim entry As Variant
    Dim lines As Collection
    Dim procName As Variant

    On Error GoTo AccessDenied

    Set vbProject = gSource.VBProject

    For Each component In vbProject.VBComponents

        Set codeModule = component.CodeModule
        moduleName = component.Name
        proc = ""

        For lineNo = 1 To codeModule.CountOfLines

            lineText = Trim$(codeModule.Lines(lineNo, 1))

            Set matches = RegexMatches( _
                lineText, _
                "^(Public |Private |Friend |Static )*" & _
                "(Sub|Function)\s+([^\s(]+)")

            If matches.Count > 0 Then

                Set match = matches(0)

                proc = CStr(match.SubMatches(2))

                If Not gProcedures.Exists(proc) Then
                    Set lines = New Collection
                    gProcedures.Add proc, lines
                End If

                GoTo NextLine

            End If

            If LCase$(Left$(lineText, 7)) = "end sub" Or _
               LCase$(Left$(lineText, 12)) = "end function" Then

                proc = ""
                GoTo NextLine

            End If

            If proc <> "" Then

                Set lines = gProcedures(proc)

                lines.Add Array( _
                    moduleName, lineNo, lineText)

            End If

NextLine:

        Next lineNo

    Next component

    ' First pass: discover direct navigation targets.
    For Each procName In gProcedures.Keys

        Set targets = NewDict()
        Set lines = gProcedures(procName)

        For Each entry In lines

            lineText = CStr(entry(2))

            If InStr(1, lineText, ".Activate", _
                vbTextCompare) > 0 Or _
               InStr(1, lineText, ".Select", _
                vbTextCompare) > 0 Or _
               InStr(1, lineText, "Application.Goto", _
                vbTextCompare) > 0 Then

                CollectVbaTargets lineText, targets

            End If

        Next entry

        gProcTargets(procName) = targets

    Next procName

    ' Second pass: resolve simple procedure calls.
    ResolveProcedureCalls

    Exit Sub

AccessDenied:

    AddUnresolved "VBA", "", _
        "VBA project access denied or unavailable", _
        Err.Description

    Err.Clear

End Sub

Private Sub CollectVbaTargets( _
    ByVal lineText As String, _
    ByVal targets As Object)

    Dim matches As Object
    Dim m As Object
    Dim target As String
    Dim name As Variant

    ' Sheets("Name") and Worksheets("Name")
    Set matches = RegexMatches( _
        lineText, _
        "(Worksheets|Sheets)\s*\(\s*""([^""]+)""\s*\)")

    For Each m In matches

        target = NormalizeSheet(CStr(m.SubMatches(1)))

        If target <> "" Then targets(target) = True

    Next m

    ' Worksheet CodeName.Activate
    For Each name In gCodeNames.Keys

        If InStr(1, lineText, _
            CStr(name) & ".Activate", vbTextCompare) > 0 Or _
           InStr(1, lineText, _
            CStr(name) & ".Select", vbTextCompare) > 0 Then

            target = CStr(gCodeNames(name))
            targets(target) = True

        End If

    Next name

End Sub

Private Sub ResolveProcedureCalls()

    Dim pass As Long
    Dim procName As Variant
    Dim otherProc As Variant
    Dim lines As Collection
    Dim entry As Variant
    Dim targets As Object
    Dim calledTargets As Object
    Dim target As Variant
    Dim lineText As String
    Dim pattern As String

    ' Bounded propagation prevents infinite loops in
    ' mutually recursive VBA procedures.
    For pass = 1 To 8

        For Each procName In gProcedures.Keys

            Set lines = gProcedures(procName)
            Set targets = gProcTargets(procName)

            For Each entry In lines

                lineText = Trim$(CStr(entry(2)))

                For Each otherProc In gProcedures.Keys

                    If StrComp(CStr(procName), _
                        CStr(otherProc), vbTextCompare) <> 0 Then

                        pattern = "^(Call\s+)?" & _
                            EscapeRegex(CStr(otherProc)) & _
                            "(\s|\(|$)"

                        If RegexMatches( _
                            lineText, pattern).Count > 0 Then

                            Set calledTargets = _
                                gProcTargets(otherProc)

                            For Each target In calledTargets.Keys
                                targets(CStr(target)) = True
                            Next target

                        End If

                    End If

                Next otherProc

            Next entry

        Next procName

    Next pass

End Sub

Private Function EscapeRegex(ByVal value As String) As String

    Dim chars As Variant
    Dim ch As Variant

    chars = Array("\", ".", "(", ")", "[", "]", _
                  "{", "}", "+", "*", "?", "^", "$", "|")

    For Each ch In chars
        value = Replace(value, CStr(ch), "\" & CStr(ch))
    Next ch

    EscapeRegex = value

End Function

' ============================================================
' SHAPES, BUTTONS, HYPERLINKS
' ============================================================

Private Sub ScanShapes()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim macro As String
    Dim proc As String
    Dim targets As Object
    Dim target As Variant
    Dim addressText As String
    Dim resolved As String
    Dim location As String
    Dim hl As Hyperlink

    For Each ws In gSource.Worksheets

        Application.StatusBar = _
            "Scanning shapes: " & ws.Name

        For Each shp In ws.Shapes

            gShapeCount = gShapeCount + 1

            location = ws.Name & " / " & shp.Name

            macro = ""

            On Error Resume Next
            macro = shp.OnAction
            On Error GoTo 0

            If macro <> "" Then

                proc = NormalizeMacroName(macro)

                If gProcTargets.Exists(proc) Then

                    Set targets = gProcTargets(proc)

                    If targets.Count > 0 Then

                        For Each target In targets.Keys

                            AddLink ws.Name, CStr(target), _
                                "Navigation", "Shape VBA Macro", _
                                location, macro

                        Next target

                    Else

                        AddUnresolved ws.Name, location, _
                            "Macro has no confirmed target", _
                            macro

                    End If

                Else

                    AddUnresolved ws.Name, location, _
                        "Shape macro not found", macro

                End If

            End If

            ' Direct hyperlink attached to shape
            addressText = ""

            On Error Resume Next
            addressText = shp.Hyperlink.SubAddress
            On Error GoTo 0

            resolved = SheetFromAddress(addressText)

            If resolved <> "" Then

                AddLink ws.Name, resolved, _
                    "Navigation", "Shape Hyperlink", _
                    location, addressText

            End If

        Next shp

        ' Worksheet hyperlinks, including cell hyperlinks.
        For Each hl In ws.Hyperlinks

            resolved = SheetFromAddress(hl.SubAddress)

            If resolved <> "" Then

                AddLink ws.Name, resolved, _
                    "Navigation", "Worksheet Hyperlink", _
                    hl.Range.Address(False, False), _
                    hl.SubAddress

            End If

        Next hl

    Next ws

End Sub

Private Function NormalizeMacroName( _
    ByVal macro As String) As String

    Dim p As Long

    macro = Trim$(macro)
    macro = Replace(macro, "'", "")

    p = InStrRev(macro, "!")

    If p > 0 Then macro = Mid$(macro, p + 1)

    p = InStrRev(macro, ".")

    If p > 0 Then macro = Mid$(macro, p + 1)

    NormalizeMacroName = macro

End Function

' ============================================================
' REPORT GENERATION
' ============================================================

Private Sub BuildReport()

    Dim ws As Worksheet

    Set ws = gReport.Worksheets(1)
    ws.Name = "Direct_Matrix"

    WriteMatrix ws, "Combined"

    Set ws = gReport.Worksheets.Add( _
        After:=gReport.Worksheets(gReport.Worksheets.Count))
    ws.Name = "Data_Matrix"

    WriteMatrix ws, "Data"

    Set ws = gReport.Worksheets.Add( _
        After:=gReport.Worksheets(gReport.Worksheets.Count))
    ws.Name = "Navigation_Matrix"

    WriteMatrix ws, "Navigation"

    WriteDetails "Dependency_Details", _
        Array("Source", "Target", "Kind", _
              "Reason", "Location", "Expression"), _
        gEvidenceRows

    WriteDetails "INDIRECT_Details", _
        Array("Sheet", "Cell", "Status", _
              "Target / Candidates", "Argument", _
              "Evaluated Address", "Formula"), _
        gIndirectRows

    WriteDetails "Unresolved_References", _
        Array("Source", "Location", _
              "Reason", "Expression"), _
        gUnresolvedRows

    WriteSheetInformation
    WriteStatistics

End Sub

Private Sub WriteMatrix( _
    ByVal ws As Worksheet, _
    ByVal matrixType As String)

    Dim i As Long
    Dim j As Long
    Dim source As String
    Dim target As String
    Dim value As Long
    Dim names As Variant
    Dim lastRow As Long
    Dim lastCol As Long

    names = gSheets.Keys

    ws.Cells(1, 1).Value2 = "Source / Target"

    For i = 0 To UBound(names)

        ws.Cells(1, i + 2).Value2 = CStr(names(i))
        ws.Cells(i + 2, 1).Value2 = CStr(names(i))

    Next i

    For i = 0 To UBound(names)

        source = CStr(names(i))

        For j = 0 To UBound(names)

            target = CStr(names(j))

            Select Case matrixType

                Case "Data"

                    value = Abs(CLng(HasLink( _
                        gData, source, target)))

                Case "Navigation"

                    value = Abs(CLng(HasLink( _
                        gNav, source, target)))

                Case Else

                    value = Abs(CLng( _
                        HasLink(gData, source, target) Or _
                        HasLink(gNav, source, target)))

            End Select

            With ws.Cells(i + 2, j + 2)

                .Value2 = value
                .HorizontalAlignment = xlCenter

                If value = 1 Then

                    .Interior.Color = RGB(198, 239, 206)
                    .Font.Color = RGB(0, 97, 0)

                End If

            End With

        Next j

    Next i

    lastRow = UBound(names) + 2
    lastCol = lastRow

    FormatHeader ws, lastCol

    ws.Columns(1).ColumnWidth = 38

    ws.Range( _
        ws.Cells(1, 2), _
        ws.Cells(1, lastCol)).ColumnWidth = 24

    ws.Rows(1).RowHeight = 35

    ws.Activate
    ws.Range("B2").Select

    ActiveWindow.FreezePanes = True

End Sub

Private Sub WriteDetails( _
    ByVal sheetName As String, _
    ByVal headers As Variant, _
    ByVal rows As Collection)

    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim record As Variant

    Set ws = gReport.Worksheets.Add( _
        After:=gReport.Worksheets(gReport.Worksheets.Count))

    ws.Name = sheetName

    For j = LBound(headers) To UBound(headers)
        WriteText ws.Cells(1, j + 1), CStr(headers(j))
    Next j

    For i = 1 To rows.Count

        record = rows(i)

        For j = LBound(record) To UBound(record)

            WriteText ws.Cells(i + 1, j + 1), _
                CStr(record(j))

        Next j

    Next i

    FormatHeader ws, UBound(headers) + 1

    ws.Columns(1).ColumnWidth = 35
    ws.Columns(2).ColumnWidth = 35
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 45

    If UBound(headers) >= 4 Then
        ws.Columns(5).ColumnWidth = 65
    End If

    If UBound(headers) >= 5 Then
        ws.Columns(6).ColumnWidth = 90
    End If

    If UBound(headers) >= 6 Then
        ws.Columns(7).ColumnWidth = 90
    End If

    ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(Application.Max(2, rows.Count + 1), _
                 UBound(headers) + 1)).AutoFilter

End Sub

Private Sub WriteSheetInformation()

    Dim ws As Worksheet
    Dim sourceWs As Worksheet
    Dim i As Long
    Dim state As String

    Set ws = gReport.Worksheets.Add( _
        After:=gReport.Worksheets(gReport.Worksheets.Count))

    ws.Name = "Sheet_Information"

    ws.Cells(1, 1).Value2 = "Sheet"
    ws.Cells(1, 2).Value2 = "Visibility"
    ws.Cells(1, 3).Value2 = "CodeName"

    i = 2

    For Each sourceWs In gSource.Worksheets

        Select Case sourceWs.Visible

            Case xlSheetVisible
                state = "Visible"

            Case xlSheetHidden
                state = "Hidden"

            Case xlSheetVeryHidden
                state = "VeryHidden"

        End Select

        WriteText ws.Cells(i, 1), sourceWs.Name
        WriteText ws.Cells(i, 2), state
        WriteText ws.Cells(i, 3), sourceWs.CodeName

        i = i + 1

    Next sourceWs

    FormatHeader ws, 3
    ws.Columns("A:C").ColumnWidth = 38

End Sub

Private Sub WriteStatistics()

    Dim ws As Worksheet
    Dim metrics As Variant
    Dim i As Long

    Set ws = gReport.Worksheets.Add( _
        After:=gReport.Worksheets(gReport.Worksheets.Count))

    ws.Name = "Statistics"

    metrics = Array( _
        Array("Sheets", gSheets.Count), _
        Array("Formula Cells", gFormulaCount), _
        Array("INDIRECT Calls", gIndirectCount), _
        Array("INDIRECT Fully Resolved", gIndirectResolved), _
        Array("Shapes", gShapeCount), _
        Array("Evidence Rows", gEvidenceRows.Count), _
        Array("Unresolved Rows", gUnresolvedRows.Count))

    ws.Cells(1, 1).Value2 = "Metric"
    ws.Cells(1, 2).Value2 = "Value"

    For i = LBound(metrics) To UBound(metrics)

        ws.Cells(i + 2, 1).Value2 = metrics(i)(0)
        ws.Cells(i + 2, 2).Value2 = metrics(i)(1)

    Next i

    FormatHeader ws, 2

    ws.Columns(1).ColumnWidth = 42
    ws.Columns(2).ColumnWidth = 25

End Sub

Private Sub WriteText( _
    ByVal cell As Range, _
    ByVal text As String)

    ' Explicitly store expressions as text.
    ' Prevent Excel interpreting =, +, -, @ as formulas.
    cell.NumberFormat = "@"

    If Len(text) > 32000 Then
        text = Left$(text, 32000)
    End If

    cell.Value2 = text

End Sub

Private Sub FormatHeader( _
    ByVal ws As Worksheet, _
    ByVal lastCol As Long)

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .WrapText = True

    End With

End Sub

Private Function MakeOutputPath( _
    ByVal sourcePath As String) As String

    Dim folder As String
    Dim name As String
    Dim p As Long
    Dim stamp As String

    p = InStrRev(sourcePath, "\")

    folder = Left$(sourcePath, p)
    name = Mid$(sourcePath, p + 1)

    p = InStrRev(name, ".")

    If p > 0 Then name = Left$(name, p - 1)

    stamp = Format$(Now, "yyyymmdd_hhnnss")

    MakeOutputPath = folder & name & _
        "_dependency_report_" & stamp & ".xlsx"

End Function
