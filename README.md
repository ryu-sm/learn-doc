
Option Explicit

' ============================================================
' XLSM Dependency Analyzer
' Version: 3.0
'
' Matrix direction:
' Row = Source Sheet
' Column = Referenced / Destination Sheet
'
' No source workbook modification.
' No source macros executed.
' ============================================================

Private gWB As Workbook
Private gReport As Workbook

Private gSheets As Object
Private gData As Object
Private gNav As Object
Private gEvidence As Collection
Private gUnresolved As Collection

Private gProcedures As Object
Private gMacroTargets As Object

Private gStage As String
Private gCurrentSheet As String
Private gCurrentCell As String

Private gFormulaCount As Long
Private gIndirectCount As Long
Private gShapeCount As Long

' ============================================================
' ENTRY POINT
' ============================================================

Public Sub AnalyzeXlsmDependencies()

    Dim selectedFile As Variant
    Dim outputFile As String
    Dim oldCalculation As XlCalculation
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    On Error GoTo FatalError

    selectedFile = Application.GetOpenFilename( _
        "Excel Macro Files (*.xlsm),*.xlsm", _
        , "Select XLSM File")

    If VarType(selectedFile) = vbBoolean Then Exit Sub

    oldCalculation = Application.Calculation
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    gStage = "Opening workbook"

    Set gWB = Workbooks.Open( _
        Filename:=CStr(selectedFile), _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True)

    InitializeAnalyzer

    gStage = "Loading sheets"
    LoadSheets

    gStage = "Scanning formulas"
    ScanFormulas

    gStage = "Scanning hyperlinks"
    ScanHyperlinks

    gStage = "Scanning VBA"
    ScanVba

    gStage = "Scanning shapes"
    ScanShapes

    gStage = "Generating report"
    GenerateReport

    outputFile = gWB.Path & Application.PathSeparator & _
        GetBaseName(gWB.Name) & _
        "_dependency_report_" & _
        Format(Now, "yyyymmdd_hhnnss") & ".xlsx"

    gReport.SaveAs _
        Filename:=outputFile, _
        FileFormat:=xlOpenXMLWorkbook

    gReport.Close SaveChanges:=False
    Set gReport = Nothing

    gWB.Close SaveChanges:=False
    Set gWB = Nothing

    RestoreApplication oldCalculation, oldEvents, _
                       oldScreen, oldAlerts

    MsgBox "Analysis completed." & vbCrLf & _
           "Sheets: " & gSheets.Count & vbCrLf & _
           "Formula cells: " & gFormulaCount & vbCrLf & _
           "INDIRECT calls: " & gIndirectCount & vbCrLf & _
           "Shapes: " & gShapeCount & vbCrLf & _
           "Evidence: " & gEvidence.Count & vbCrLf & _
           "Unresolved: " & gUnresolved.Count & vbCrLf & _
           "Report: " & outputFile, vbInformation

    Exit Sub

FatalError:

    Dim errorMessage As String

    errorMessage = _
        "Analysis failed." & vbCrLf & _
        "Error: " & Err.Number & vbCrLf & _
        "Description: " & Err.Description & vbCrLf & _
        "Stage: " & gStage & vbCrLf & _
        "Sheet: " & gCurrentSheet & vbCrLf & _
        "Cell: " & gCurrentCell

    On Error Resume Next

    If Not gReport Is Nothing Then
        gReport.Close SaveChanges:=False
    End If

    If Not gWB Is Nothing Then
        gWB.Close SaveChanges:=False
    End If

    RestoreApplication oldCalculation, oldEvents, _
                       oldScreen, oldAlerts

    MsgBox errorMessage, vbCritical

End Sub

' ============================================================
' INITIALIZATION
' ============================================================

Private Sub InitializeAnalyzer()

    Set gSheets = NewDict()
    Set gData = NewDict()
    Set gNav = NewDict()

    Set gEvidence = New Collection
    Set gUnresolved = New Collection

    Set gProcedures = NewDict()
    Set gMacroTargets = NewDict()

    gFormulaCount = 0
    gIndirectCount = 0
    gShapeCount = 0

    gCurrentSheet = ""
    gCurrentCell = ""

End Sub

Private Function NewDict() As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set NewDict = d

End Function

Private Sub RestoreApplication( _
    ByVal calculationMode As XlCalculation, _
    ByVal eventsEnabled As Boolean, _
    ByVal screenEnabled As Boolean, _
    ByVal alertsEnabled As Boolean)

    Application.Calculation = calculationMode
    Application.EnableEvents = eventsEnabled
    Application.ScreenUpdating = screenEnabled
    Application.DisplayAlerts = alertsEnabled

End Sub

Private Function GetBaseName(ByVal fileName As String) As String

    Dim p As Long

    p = InStrRev(fileName, ".")

    If p > 0 Then
        GetBaseName = Left$(fileName, p - 1)
    Else
        GetBaseName = fileName
    End If

End Function

' ============================================================
' SHEETS
' ============================================================

Private Sub LoadSheets()

    Dim ws As Worksheet

    For Each ws In gWB.Worksheets
        gSheets(ws.Name) = ws.Visible
    Next ws

End Sub

Private Function SheetExists(ByVal name As String) As Boolean

    SheetExists = gSheets.Exists(name)

End Function

Private Function NormalizeSheet(ByVal value As String) As String

    Dim s As String

    s = Trim$(value)

    If Left$(s, 1) = "'" And Right$(s, 1) = "'" Then
        If Len(s) >= 2 Then
            s = Mid$(s, 2, Len(s) - 2)
        End If
    End If

    s = Replace(s, "''", "'")

    NormalizeSheet = s

End Function

' ============================================================
' REGEX HELPERS
' ============================================================

Private Function RegexMatches( _
    ByVal value As String, _
    ByVal pattern As String) As Object

    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    re.Pattern = pattern

    Set RegexMatches = re.Execute(value)

End Function

Private Function RegexReplace( _
    ByVal value As String, _
    ByVal pattern As String, _
    ByVal replacement As String) As String

    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")

    re.Global = True
    re.IgnoreCase = True
    re.Pattern = pattern

    RegexReplace = re.Replace(value, replacement)

End Function

Private Function EscapeRegex(ByVal value As String) As String

    Dim i As Long
    Dim ch As String
    Dim result As String

    result = ""

    For i = 1 To Len(value)

        ch = Mid$(value, i, 1)

        If InStr(1, "\.^$|?*+()[]{}", ch, _
            vbBinaryCompare) > 0 Then
            result = result & "\"
        End If

        result = result & ch

    Next i

    EscapeRegex = result

End Function

' ============================================================
' RELATIONSHIP STORAGE
' ============================================================

Private Sub AddRelation( _
    ByVal sourceSheet As String, _
    ByVal targetSheet As String, _
    ByVal relationType As String, _
    ByVal location As String, _
    ByVal expression As String, _
    ByVal method As String)

    Dim key As String

    If Not SheetExists(sourceSheet) Then Exit Sub
    If Not SheetExists(targetSheet) Then Exit Sub

    key = sourceSheet & Chr$(30) & targetSheet

    If relationType = "DATA" Then
        gData(key) = True
    Else
        gNav(key) = True
    End If

    gEvidence.Add Array( _
        sourceSheet, targetSheet, relationType, _
        location, expression, method)

End Sub

Private Sub AddUnresolved( _
    ByVal sourceSheet As String, _
    ByVal location As String, _
    ByVal category As String, _
    ByVal expression As String, _
    ByVal reason As String)

    gUnresolved.Add Array( _
        sourceSheet, location, category, _
        expression, reason)

End Sub

' ============================================================
' EXPLICIT SHEET REFERENCES
' ============================================================

Private Function ExtractSheetRefs( _
    ByVal expression As String) As Object

    Dim result As Object
    Dim matches As Object
    Dim m As Object
    Dim target As String
    Dim cleaned As String

    Set result = NewDict()

    ' Remove ordinary Excel string literals.
    cleaned = RegexReplace( _
        expression, _
        """(?:[^""]|"""")*""", _
        "")

    ' Quoted sheet names.
    Set matches = RegexMatches( _
        cleaned, _
        "'(([^']|'')+)'!")

    For Each m In matches

        target = NormalizeSheet(CStr(m.SubMatches(0)))

        If SheetExists(target) Then
            result(target) = True
        End If

    Next m

    ' Unquoted names, including Japanese characters.
    Set matches = RegexMatches( _
        cleaned, _
        "(^|[^A-Za-z0-9_.'\]])" & _
        "([^ \t\r\n\+\-\*/\(\),=:!""']+)!")

    For Each m In matches

        target = NormalizeSheet(CStr(m.SubMatches(1)))

        If SheetExists(target) Then
            result(target) = True
        End If

    Next m

    Set ExtractSheetRefs = result

End Function

' ============================================================
' FORMULA SCANNING
' ============================================================

Private Sub ScanFormulas()

    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim cell As Range
    Dim formula As String
    Dim refs As Object
    Dim target As Variant

    For Each ws In gWB.Worksheets

        gCurrentSheet = ws.Name
        gStage = "Scanning formulas: " & ws.Name

        Set formulaCells = Nothing

        On Error Resume Next
        Set formulaCells = ws.UsedRange.SpecialCells( _
            xlCellTypeFormulas)
        On Error GoTo 0

        If Not formulaCells Is Nothing Then

            For Each cell In formulaCells.Cells

                gCurrentCell = cell.Address(False, False)
                formula = CStr(cell.Formula)

                gFormulaCount = gFormulaCount + 1

                Set refs = ExtractSheetRefs(formula)

                For Each target In refs.Keys

                    AddRelation ws.Name, CStr(target), _
                        "DATA", gCurrentCell, _
                        formula, "Explicit formula"

                Next target

                If InStr(1, formula, "INDIRECT(", _
                    vbTextCompare) > 0 Then

                    gIndirectCount = gIndirectCount + 1

                    AnalyzeIndirect ws, cell, formula

                End If

                If gFormulaCount Mod 10000 = 0 Then
                    Application.StatusBar = _
                        "Formula cells: " & gFormulaCount
                End If

            Next cell

        End If

    Next ws

    Application.StatusBar = False

End Sub

' ============================================================
' INDIRECT ANALYSIS
' ============================================================

Private Sub AnalyzeIndirect( _
    ByVal ws As Worksheet, _
    ByVal cell As Range, _
    ByVal formula As String)

    Dim refs As Object
    Dim target As Variant
    Dim matches As Object
    Dim m As Object
    Dim candidate As String
    Dim found As Boolean

    found = False

    ' Find literal sheet names inside string expressions.
    ' Example:
    ' INDIRECT("サマリ!A1")
    ' INDIRECT("サマリ!" & ADDRESS(...))

    Set matches = RegexMatches( _
        formula, _
        """([^""]+)!")

    For Each m In matches

        candidate = NormalizeSheet( _
            CStr(m.SubMatches(0)))

        If SheetExists(candidate) Then

            AddRelation ws.Name, candidate, _
                "DATA", cell.Address(False, False), _
                formula, "INDIRECT literal"

            found = True

        End If

    Next m

    ' Dynamic expressions may use a cell to hold
    ' the destination sheet name.
    '
    ' Example:
    ' INDIRECT("'" & $E36 & "'!$H$119:$W$123")
    '
    ' These are not automatically confirmed here.
    ' They are reported for manual verification.

    If Not found Then

        AddUnresolved ws.Name, _
            cell.Address(False, False), _
            "INDIRECT", formula, _
            "Dynamic target requires evaluation"

    End If

End Sub

' ============================================================
' HYPERLINKS
' ============================================================

Private Sub ScanHyperlinks()

    Dim ws As Worksheet
    Dim h As Hyperlink
    Dim target As String
    Dim subAddress As String
    Dim p As Long

    For Each ws In gWB.Worksheets

        gCurrentSheet = ws.Name
        gStage = "Scanning hyperlinks"

        For Each h In ws.Hyperlinks

            subAddress = h.SubAddress

            If Len(subAddress) > 0 Then

                p = InStr(1, subAddress, "!")

                If p > 0 Then

                    target = NormalizeSheet( _
                        Left$(subAddress, p - 1))

                    If SheetExists(target) Then

                        AddRelation ws.Name, target, _
                            "NAV", "Hyperlink", _
                            subAddress, "Hyperlink"

                    End If

                End If

            End If

        Next h

    Next ws

End Sub

' ============================================================
' VBA PROCEDURE ANALYSIS
' ============================================================

Private Sub ScanVba()

    Dim vbComponent As Object
    Dim codeModule As Object

    Dim i As Long
    Dim lineText As String
    Dim procName As String
    Dim procBody As String

    Dim matches As Object
    Dim m As Object
    Dim targets As Object
    Dim target As Variant

    gStage = "Scanning VBA modules"

    On Error GoTo VbaAccessError

    For Each vbComponent In gWB.VBProject.VBComponents

        Set codeModule = vbComponent.CodeModule

        procName = ""
        procBody = ""

        For i = 1 To codeModule.CountOfLines

            lineText = codeModule.Lines(i, 1)

            If IsProcedureStart(lineText) Then

                If Len(procName) > 0 Then
                    StoreProcedure procName, procBody
                End If

                procName = ExtractProcedureName(lineText)
                procBody = lineText & vbCrLf

            ElseIf Len(procName) > 0 Then

                procBody = procBody & lineText & vbCrLf

                If IsProcedureEnd(lineText) Then

                    StoreProcedure procName, procBody

                    procName = ""
                    procBody = ""

                End If

            End If

        Next i

        If Len(procName) > 0 Then
            StoreProcedure procName, procBody
        End If

    Next vbComponent

    Exit Sub

VbaAccessError:

    AddUnresolved "", "", "VBA", "", _
        "Cannot access VBA project: " & Err.Description

    Err.Clear

End Sub

Private Function IsProcedureStart( _
    ByVal lineText As String) As Boolean

    IsProcedureStart = RegexMatches( _
        Trim$(lineText), _
        "^(Public |Private |Friend )?(Static )?" & _
        "(Sub|Function) [A-Za-z_][A-Za-z0-9_]*").Count > 0

End Function

Private Function IsProcedureEnd( _
    ByVal lineText As String) As Boolean

    IsProcedureEnd = RegexMatches( _
        Trim$(lineText), _
        "^End (Sub|Function)").Count > 0

End Function

Private Function ExtractProcedureName( _
    ByVal lineText As String) As String

    Dim matches As Object

    Set matches = RegexMatches( _
        lineText, _
        "(Sub|Function) +([A-Za-z_][A-Za-z0-9_]*)")

    If matches.Count > 0 Then

        ExtractProcedureName = _
            CStr(matches(0).SubMatches(1))

    Else

        ExtractProcedureName = ""

    End If

End Function

Private Sub StoreProcedure( _
    ByVal procName As String, _
    ByVal body As String)

    Dim targets As Object
    Dim matches As Object
    Dim m As Object
    Dim target As String

    Set targets = NewDict()

    ' Detect:
    ' ThisWorkbook.Sheets("サマリ").Activate
    ' Worksheets("マスタ").Select

    Set matches = RegexMatches( _
        body, _
        "(Sheets|Worksheets)\s*\(\s*""([^""]+)""\s*\)" & _
        "\s*\.\s*(Activate|Select)")

    For Each m In matches

        target = CStr(m.SubMatches(1))

        If SheetExists(target) Then
            targets(target) = True
        End If

    Next m

    ' Also detect sheet-qualified Range access.
    Set matches = RegexMatches( _
        body, _
        "(Sheets|Worksheets)\s*\(\s*""([^""]+)""\s*\)")

    For Each m In matches

        target = CStr(m.SubMatches(1))

        If SheetExists(target) Then
            targets(target) = True
        End If

    Next m

    If gProcedures.Exists(procName) Then
        gProcedures.Remove procName
    End If

    gProcedures.Add procName, body

    If gMacroTargets.Exists(procName) Then
        gMacroTargets.Remove procName
    End If

    ' Store object with Set.
    gMacroTargets.Add procName, targets

End Sub

' ============================================================
' SHAPES AND BUTTONS
' ============================================================

Private Sub ScanShapes()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim macroName As String
    Dim procName As String
    Dim targets As Object
    Dim target As Variant

    For Each ws In gWB.Worksheets

        gCurrentSheet = ws.Name
        gStage = "Scanning shapes"

        For Each shp In ws.Shapes

            gShapeCount = gShapeCount + 1

            macroName = ""

            On Error Resume Next
            macroName = shp.OnAction
            On Error GoTo 0

            If Len(macroName) > 0 Then

                procName = NormalizeMacroName(macroName)

                If gMacroTargets.Exists(procName) Then

                    Set targets = gMacroTargets(procName)

                    For Each target In targets.Keys

                        AddRelation ws.Name, CStr(target), _
                            "NAV", shp.Name, _
                            macroName, "Shape VBA macro"

                    Next target

                Else

                    AddUnresolved ws.Name, _
                        shp.Name, "Shape Macro", _
                        macroName, _
                        "Macro not resolved"

                End If

            End If

        Next shp

    Next ws

End Sub

Private Function NormalizeMacroName( _
    ByVal macroName As String) As String

    Dim p As Long
    Dim s As String

    s = Replace(macroName, "'", "")

    p = InStrRev(s, "!")

    If p > 0 Then
        s = Mid$(s, p + 1)
    End If

    p = InStrRev(s, ".")

    If p > 0 Then
        s = Mid$(s, p + 1)
    End If

    NormalizeMacroName = Trim$(s)

End Function

' ============================================================
' REPORT GENERATION
' ============================================================

Private Sub GenerateReport()

    Set gReport = Workbooks.Add(xlWBATWorksheet)

    WriteMatrix "Direct_Matrix", "ALL"
    WriteMatrix "Data_Matrix", "DATA"
    WriteMatrix "Navigation_Matrix", "NAV"

    WriteEvidence
    WriteUnresolved
    WriteSheetInformation
    WriteStatistics

    Application.StatusBar = False

End Sub

Private Sub WriteMatrix( _
    ByVal sheetName As String, _
    ByVal matrixType As String)

    Dim ws As Worksheet
    Dim names As Variant
    Dim i As Long
    Dim j As Long
    Dim key As String
    Dim value As Long

    Set ws = GetReportSheet(sheetName)

    names = gSheets.Keys

    ws.Cells(1, 1).Value = "Source / Target"

    For i = 0 To UBound(names)

        ws.Cells(i + 2, 1).Value = names(i)
        ws.Cells(1, i + 2).Value = names(i)

    Next i

    For i = 0 To UBound(names)

        For j = 0 To UBound(names)

            key = CStr(names(i)) & Chr$(30) & _
                  CStr(names(j))

            value = 0

            Select Case matrixType

                Case "DATA"

                    If gData.Exists(key) Then value = 1

                Case "NAV"

                    If gNav.Exists(key) Then value = 1

                Case Else

                    If gData.Exists(key) Or _
                       gNav.Exists(key) Then
                        value = 1
                    End If

            End Select

            ws.Cells(i + 2, j + 2).Value = value

        Next j

    Next i

    FormatReportSheet ws

End Sub

Private Sub WriteEvidence()

    Dim ws As Worksheet
    Dim headers As Variant
    Dim item As Variant
    Dim i As Long
    Dim j As Long

    Set ws = GetReportSheet("Dependency_Details")

    headers = Array( _
        "Source Sheet", "Target Sheet", _
        "Type", "Location", _
        "Expression", "Detection Method")

    WriteHeaders ws, headers

    For i = 1 To gEvidence.Count

        item = gEvidence(i)

        For j = 0 To UBound(item)

            WriteTextCell ws.Cells(i + 1, j + 1), _
                          CStr(item(j))

        Next j

    Next i

    FormatReportSheet ws

End Sub

Private Sub WriteUnresolved()

    Dim ws As Worksheet
    Dim headers As Variant
    Dim item As Variant
    Dim i As Long
    Dim j As Long

    Set ws = GetReportSheet("Unresolved_References")

    headers = Array( _
        "Source Sheet", "Location", _
        "Category", "Expression", "Reason")

    WriteHeaders ws, headers

    For i = 1 To gUnresolved.Count

        item = gUnresolved(i)

        For j = 0 To UBound(item)

            WriteTextCell ws.Cells(i + 1, j + 1), _
                          CStr(item(j))

        Next j

    Next i

    FormatReportSheet ws

End Sub

Private Sub WriteSheetInformation()

    Dim ws As Worksheet
    Dim sourceWs As Worksheet
    Dim i As Long
    Dim visibility As String

    Set ws = GetReportSheet("Sheet_Information")

    WriteHeaders ws, Array( _
        "Sheet Name", "Visibility", "Index")

    i = 2

    For Each sourceWs In gWB.Worksheets

        Select Case sourceWs.Visible

            Case xlSheetVisible
                visibility = "Visible"

            Case xlSheetHidden
                visibility = "Hidden"

            Case xlSheetVeryHidden
                visibility = "VeryHidden"

        End Select

        ws.Cells(i, 1).Value = sourceWs.Name
        ws.Cells(i, 2).Value = visibility
        ws.Cells(i, 3).Value = sourceWs.Index

        i = i + 1

    Next sourceWs

    FormatReportSheet ws

End Sub

Private Sub WriteStatistics()

    Dim ws As Worksheet

    Set ws = GetReportSheet("Statistics")

    WriteHeaders ws, Array("Metric", "Value")

    ws.Cells(2, 1).Value = "Sheets"
    ws.Cells(2, 2).Value = gSheets.Count

    ws.Cells(3, 1).Value = "Formula cells"
    ws.Cells(3, 2).Value = gFormulaCount

    ws.Cells(4, 1).Value = "INDIRECT calls detected"
    ws.Cells(4, 2).Value = gIndirectCount

    ws.Cells(5, 1).Value = "Shapes"
    ws.Cells(5, 2).Value = gShapeCount

    ws.Cells(6, 1).Value = "Data relationships"
    ws.Cells(6, 2).Value = gData.Count

    ws.Cells(7, 1).Value = "Navigation relationships"
    ws.Cells(7, 2).Value = gNav.Count

    ws.Cells(8, 1).Value = "Evidence rows"
    ws.Cells(8, 2).Value = gEvidence.Count

    ws.Cells(9, 1).Value = "Unresolved rows"
    ws.Cells(9, 2).Value = gUnresolved.Count

    FormatReportSheet ws

End Sub

' ============================================================
' REPORT HELPERS
' ============================================================

Private Function GetReportSheet( _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = gReport.Worksheets(sheetName)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = gReport.Worksheets.Add( _
            After:=gReport.Worksheets( _
                gReport.Worksheets.Count))

        ws.Name = sheetName

    End If

    Set GetReportSheet = ws

End Function

Private Sub WriteHeaders( _
    ByVal ws As Worksheet, _
    ByVal headers As Variant)

    Dim i As Long

    For i = LBound(headers) To UBound(headers)

        ws.Cells(1, i + 1).Value = headers(i)

    Next i

End Sub

Private Sub WriteTextCell( _
    ByVal target As Range, _
    ByVal value As String)

    ' Force text storage.
    ' Prevent Excel from interpreting formula strings.

    target.NumberFormat = "@"

    ' Excel cell text limit is 32767 characters.
    If Len(value) > 32767 Then
        value = Left$(value, 32767)
    End If

    target.Value2 = value

End Sub

Private Sub FormatReportSheet(ByVal ws As Worksheet)

    With ws.Rows(1)

        .Font.Bold = True
        .Interior.Color = RGB(35, 65, 100)
        .Font.Color = RGB(255, 255, 255)

    End With

    ws.Rows(1).AutoFilter

    ws.Cells.EntireColumn.AutoFit

    If ws.Columns(1).ColumnWidth > 45 Then
        ws.Columns(1).ColumnWidth = 45
    End If

End Sub
