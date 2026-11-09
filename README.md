
import argparse
import os
import re
import sys
import zipfile
import traceback
from collections import defaultdict, Counter
from pathlib import Path
from xml.etree import ElementTree as ET

import pythoncom
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from oletools.olevba import VBA_Parser


# ============================================================
# CONFIG
# ============================================================

MAX_EVIDENCE = 250000
MAX_UNRESOLVED = 50000
PROGRESS_EVERY = 5000

# Never execute VBA.
# Never save the source workbook.
#
# Dynamic formula evaluation uses Excel's calculation engine
# in an isolated Excel process with macros disabled.
#
# IMPORTANT:
# Excel may still update in-memory formula values.
# The source file is opened read-only and never saved.

XL_CALC_MANUAL = -4135
XL_CELL_TYPE_FORMULAS = -4123
XL_CELL_TYPE_VISIBLE = 12
XL_LINK_TYPE_EXCEL = 1

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}


# ============================================================
# HELPERS
# ============================================================

def log(*args):
    print(*args, flush=True)


def safe_str(value):
    if value is None:
        return ""
    return str(value)


def normalize_sheet_name(value):
    value = safe_str(value).strip()
    if value.startswith("'") and value.endswith("'"):
        value = value[1:-1]
    return value.replace("''", "'").strip()


def normalize_macro_name(value):
    value = safe_str(value).strip()
    if "!" in value:
        value = value.rsplit("!", 1)[-1]
    return value.strip("'").lower()


def quote_sheet(name):
    return "'" + name.replace("'", "''") + "'"


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def write_text(cell, value):
    cell.number_format = "@"
    cell.value = safe_str(value)
    cell.data_type = "s"


def clean_expression(value):
    """
    Always write expressions as text to avoid Excel evaluating
    evidence strings beginning with '='.
    """
    return safe_str(value)


def unique_ordered(values):
    seen = set()
    result = []
    for v in values:
        if v not in seen:
            seen.add(v)
            result.append(v)
    return result


# ============================================================
# EXCEL FORMULA STRING PARSER
# ============================================================

def split_arguments(text):
    """
    Split arguments of a function call while respecting:
    - nested parentheses
    - Excel double-quoted strings
    - escaped double quotes
    """
    result = []
    buf = []
    depth = 0
    quoted = False
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == '"':
            buf.append(ch)

            if quoted and i + 1 < len(text) and text[i + 1] == '"':
                buf.append('"')
                i += 2
                continue

            quoted = not quoted

        elif not quoted and ch == "(":
            depth += 1
            buf.append(ch)

        elif not quoted and ch == ")":
            depth -= 1
            buf.append(ch)

        elif not quoted and depth == 0 and ch in ",;":
            result.append("".join(buf).strip())
            buf = []

        else:
            buf.append(ch)

        i += 1

    result.append("".join(buf).strip())
    return result


def extract_function_calls(formula, function_name):
    """
    Return the inner arguments of every function call,
    including nested calls.
    """
    result = []
    pattern = re.compile(
        r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?"
        + re.escape(function_name)
        + r"\s*\(",
        re.I,
    )

    for match in pattern.finditer(formula):
        start = match.end()
        depth = 1
        quoted = False
        i = start

        while i < len(formula):
            ch = formula[i]

            if ch == '"':
                if quoted and i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 2
                    continue
                quoted = not quoted

            elif not quoted and ch == "(":
                depth += 1

            elif not quoted and ch == ")":
                depth -= 1
                if depth == 0:
                    result.append(formula[start:i])
                    break

            i += 1

    return result


def strip_formula_strings(formula):
    """
    Replace string literal contents with spaces.
    Prevent false positives from textual mentions of Sheet!A1.
    """
    result = []
    quoted = False
    i = 0

    while i < len(formula):
        ch = formula[i]

        if ch == '"':
            if quoted and i + 1 < len(formula) and formula[i + 1] == '"':
                result.extend([" ", " "])
                i += 2
                continue

            quoted = not quoted
            result.append(" ")

        else:
            result.append(" " if quoted else ch)

        i += 1

    return "".join(result)


# ============================================================
# ANALYZER
# ============================================================

class XLSMAnalyzer:

    def __init__(self, source, output, include_self=False):
        self.source = Path(source).resolve()
        self.output = Path(output).resolve()
        self.include_self = include_self

        self.excel = None
        self.book = None

        self.sheet_names = []
        self.sheet_set = set()
        self.sheet_lookup = {}

        self.sheet_visibility = {}

        self.data_links = set()
        self.navigation_links = set()

        self.evidence = []
        self.unresolved = []

        self.stats = Counter()

        self.vba_procedures = {}
        self.vba_module_count = 0

        self.macro_bindings = []
        self.dynamic_cache = {}

    # --------------------------------------------------------
    # Evidence
    # --------------------------------------------------------

    def add_evidence(
        self,
        source,
        target,
        category,
        location,
        expression,
        details="",
        navigation=False,
    ):
        source = normalize_sheet_name(source)
        target = normalize_sheet_name(target)

        source = self.sheet_lookup.get(source.lower(), source)
        target = self.sheet_lookup.get(target.lower(), target)

        if source not in self.sheet_set:
            return False

        if target not in self.sheet_set:
            return False

        if source == target and not self.include_self:
            return False

        if navigation:
            self.navigation_links.add((source, target))
        else:
            self.data_links.add((source, target))

        self.stats["evidence_" + category] += 1

        if len(self.evidence) < MAX_EVIDENCE:
            self.evidence.append((
                source,
                target,
                category,
                location,
                clean_expression(expression),
                details,
            ))
        else:
            self.stats["truncated_evidence"] += 1

        return True

    def add_unresolved(
        self,
        sheet,
        location,
        category,
        expression,
        reason,
    ):
        self.stats["unresolved_" + category] += 1

        if len(self.unresolved) < MAX_UNRESOLVED:
            self.unresolved.append((
                sheet,
                location,
                category,
                clean_expression(expression),
                reason,
            ))
        else:
            self.stats["truncated_unresolved"] += 1

    # --------------------------------------------------------
    # Workbook
    # --------------------------------------------------------

    def open_excel(self):
        log("Starting isolated Excel process...")

        pythoncom.CoInitialize()

        self.excel = win32.DispatchEx("Excel.Application")

        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.EnableEvents = False
        self.excel.AskToUpdateLinks = False

        try:
            self.excel.AutomationSecurity = (
                MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            )
        except Exception:
            pass

        try:
            self.excel.Calculation = XL_CALC_MANUAL
        except Exception:
            pass

        self.book = self.excel.Workbooks.Open(
            Filename=str(self.source),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        for ws in self.book.Worksheets:
            name = str(ws.Name)
            self.sheet_names.append(name)
            self.sheet_visibility[name] = int(ws.Visible)

        self.sheet_set = set(self.sheet_names)
        self.sheet_lookup = {
            name.lower(): name
            for name in self.sheet_names
        }

        log("Sheets:", len(self.sheet_names))

    def close_excel(self):
        if self.book is not None:
            try:
                self.book.Close(SaveChanges=False)
            except Exception:
                pass
            self.book = None

        if self.excel is not None:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # --------------------------------------------------------
    # Static references
    # --------------------------------------------------------

    def find_explicit_sheets(self, formula):
        cleaned = strip_formula_strings(formula)

        result = set()

        # Quoted references: 'Sheet Name'!A1
        quoted = re.compile(
            r"'((?:[^']|'')+)'!",
            re.I,
        )

        for match in quoted.finditer(cleaned):
            name = normalize_sheet_name(match.group(1))
            actual = self.sheet_lookup.get(name.lower())
            if actual:
                result.add(actual)

        # Unquoted references: Sheet1!A1
        unquoted = re.compile(
            r"(?<![A-Za-z0-9_.])"
            r"([^\s+\-*/^&=(),;:{}\[\]'\"!]+)!",
            re.I,
        )

        for match in unquoted.finditer(cleaned):
            name = normalize_sheet_name(match.group(1))
            actual = self.sheet_lookup.get(name.lower())
            if actual:
                result.add(actual)

        return result

    # --------------------------------------------------------
    # Dynamic formula analysis
    # --------------------------------------------------------

    def evaluate_expression(self, ws, expression):
        """
        Evaluate the INDIRECT text argument using Excel itself.

        Example:
            "'" & $E36 & "'!$H$119:$W$123"

        Returns:
            '1023951_I_STAR現行売上'!$H$119:$W$123

        Does not evaluate the final INDIRECT reference.
        """
        expression = expression.strip()

        if not expression:
            return None

        cache_key = (ws.Name, expression)

        if cache_key in self.dynamic_cache:
            return self.dynamic_cache[cache_key]

        try:
            # Excel Evaluate uses the worksheet's context.
            # It may evaluate names, references and string functions.
            result = ws.Evaluate("=" + expression)

            if isinstance(result, str):
                self.dynamic_cache[cache_key] = result
                return result

        except Exception:
            pass

        return None

    def parse_reference_text(self, reference_text):
        """
        Extract sheet names from a resolved Excel reference.

        Example:
            'Sheet A'!$H$119:$W$123
        """
        text = safe_str(reference_text).strip()

        if not text:
            return set()

        result = set()

        quoted = re.compile(
            r"'((?:[^']|'')+)'!",
            re.I,
        )

        for match in quoted.finditer(text):
            name = normalize_sheet_name(match.group(1))
            actual = self.sheet_lookup.get(name.lower())
            if actual:
                result.add(actual)

        unquoted = re.compile(
            r"(?<![A-Za-z0-9_.])"
            r"([^\s+\-*/^&=(),;:{}\[\]'\"!]+)!",
            re.I,
        )

        for match in unquoted.finditer(text):
            name = normalize_sheet_name(match.group(1))
            actual = self.sheet_lookup.get(name.lower())
            if actual:
                result.add(actual)

        return result

    def analyze_indirect(self, ws, address, formula):
        calls = extract_function_calls(formula, "INDIRECT")

        if not calls:
            return

        self.stats["indirect_calls"] += len(calls)

        for call in calls:
            args = split_arguments(call)

            if not args:
                continue

            expression = args[0]

            resolved = self.evaluate_expression(
                ws,
                expression,
            )

            if resolved is None:
                self.add_unresolved(
                    ws.Name,
                    address,
                    "INDIRECT",
                    formula,
                    "Excel could not evaluate text argument: "
                    + expression,
                )
                continue

            targets = self.parse_reference_text(resolved)

            if not targets:
                # Local references without an explicit sheet name
                # do not create a cross-sheet dependency.
                if "!" not in resolved:
                    continue

                self.add_unresolved(
                    ws.Name,
                    address,
                    "INDIRECT",
                    formula,
                    "Resolved text has no matching workbook sheet: "
                    + resolved,
                )
                continue

            for target in targets:
                self.add_evidence(
                    ws.Name,
                    target,
                    "INDIRECT",
                    address,
                    formula,
                    "Resolved reference: " + resolved,
                )

    def analyze_offset(self, ws, address, formula):
        calls = extract_function_calls(formula, "OFFSET")

        self.stats["offset_calls"] += len(calls)

        for call in calls:
            args = split_arguments(call)

            if not args:
                continue

            first_argument = args[0]

            targets = self.find_explicit_sheets(
                first_argument
            )

            for target in targets:
                self.add_evidence(
                    ws.Name,
                    target,
                    "OFFSET",
                    address,
                    formula,
                    "OFFSET base: " + first_argument,
                )

            if not targets and (
                "INDIRECT(" in first_argument.upper()
                or "INDEX(" in first_argument.upper()
            ):
                self.add_unresolved(
                    ws.Name,
                    address,
                    "OFFSET",
                    formula,
                    "Dynamic OFFSET base requires deeper evaluation",
                )

    # --------------------------------------------------------
    # Formula scan
    # --------------------------------------------------------

    def scan_formulas(self):
        log("Scanning formulas...")

        for ws in self.book.Worksheets:
            sheet_name = str(ws.Name)

            log("Sheet:", sheet_name)

            try:
                formula_cells = ws.UsedRange.SpecialCells(
                    XL_CELL_TYPE_FORMULAS
                )
            except Exception:
                continue

            # Areas avoids reading the entire UsedRange as one huge
            # COM array and preserves formula cell coordinates.
            try:
                areas = formula_cells.Areas
                area_count = areas.Count
            except Exception:
                continue

            for area_index in range(1, area_count + 1):
                area = areas.Item(area_index)

                try:
                    row_count = area.Rows.Count
                    col_count = area.Columns.Count

                    values = area.Formula

                    if row_count == 1 and col_count == 1:
                        values = ((values,),)
                    elif row_count == 1:
                        if not isinstance(values[0], tuple):
                            values = (tuple(values),)
                    elif col_count == 1:
                        values = tuple(
                            (v,) if not isinstance(v, tuple) else v
                            for v in values
                        )

                    start_row = area.Row
                    start_col = area.Column

                    for r_index, row_values in enumerate(values):
                        for c_index, formula in enumerate(row_values):
                            if not is_formula(formula):
                                continue

                            self.stats["formula_cells"] += 1

                            row_num = start_row + r_index
                            col_num = start_col + c_index

                            address = (
                                get_column_letter(col_num)
                                + str(row_num)
                            )

                            targets = self.find_explicit_sheets(
                                formula
                            )

                            for target in targets:
                                self.add_evidence(
                                    sheet_name,
                                    target,
                                    "Formula",
                                    address,
                                    formula,
                                    "Explicit formula reference",
                                )

                            if "INDIRECT(" in formula.upper():
                                self.analyze_indirect(
                                    ws,
                                    address,
                                    formula,
                                )

                            if "OFFSET(" in formula.upper():
                                self.analyze_offset(
                                    ws,
                                    address,
                                    formula,
                                )

                            count = self.stats["formula_cells"]

                            if count % PROGRESS_EVERY == 0:
                                log("Processed:", count)

                except Exception as exc:
                    self.add_unresolved(
                        sheet_name,
                        safe_str(area.Address),
                        "FormulaArea",
                        "",
                        safe_str(exc),
                    )

    # --------------------------------------------------------
    # Defined names
    # --------------------------------------------------------

    def scan_defined_names(self):
        log("Scanning defined names...")

        try:
            names = self.book.Names

            for i in range(1, names.Count + 1):
                item = names.Item(i)

                name = safe_str(item.Name)
                expression = safe_str(item.RefersTo)

                targets = self.find_explicit_sheets(
                    expression
                )

                # A workbook-scoped name has no unique source
                # worksheet, so do not invent matrix edges.
                # Record it as evidence for further inspection.
                if targets:
                    self.stats["defined_names_with_targets"] += 1

                self.add_unresolved(
                    "(Defined Name)",
                    name,
                    "DefinedName",
                    expression,
                    "Workbook-level name; source sheet usage "
                    "must be traced from formulas",
                )

        except Exception as exc:
            log("Defined name warning:", exc)

    # --------------------------------------------------------
    # VBA parsing
    # --------------------------------------------------------

    def load_vba(self):
        log("Scanning VBA...")

        parser = VBA_Parser(str(self.source))

        try:
            if not parser.detect_vba_macros():
                return

            for _, stream_path, filename, code in (
                parser.extract_macros()
            ):
                self.vba_module_count += 1

                module = safe_str(filename or stream_path)

                self.extract_vba_procedures(
                    module,
                    code,
                )

        finally:
            parser.close()

    def extract_vba_procedures(self, module, code):
        lines = code.splitlines()

        start_re = re.compile(
            r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
            r"(?:Static\s+)?"
            r"(Sub|Function)\s+([^\s(]+)",
            re.I,
        )

        end_re = re.compile(
            r"^\s*End\s+(Sub|Function)\b",
            re.I,
        )

        current_name = None
        current_lines = []

        for line in lines:
            match = start_re.match(line)

            if match:
                current_name = match.group(2)
                current_lines = [line]
                continue

            if current_name is not None:
                current_lines.append(line)

                if end_re.match(line):
                    key = normalize_macro_name(
                        current_name
                    )

                    self.vba_procedures[key] = {
                        "module": module,
                        "name": current_name,
                        "code": "\n".join(current_lines),
                    }

                    current_name = None
                    current_lines = []

        self.stats["vba_procedures"] = len(
            self.vba_procedures
        )

    def extract_vba_targets(self, code):
        """
        Static VBA references:
            Sheets("Name")
            Worksheets("Name")
            ThisWorkbook.Sheets("Name")
        """
        result = set()

        pattern = re.compile(
            r"(?:Worksheets|Sheets)\s*\(\s*"
            r'["\']([^"\']+)["\']\s*\)',
            re.I,
        )

        for match in pattern.finditer(code):
            name = normalize_sheet_name(match.group(1))
            actual = self.sheet_lookup.get(name.lower())

            if actual:
                result.add(actual)

        return result

    def resolve_macro_targets(self, macro_name):
        """
        Resolve explicit sheet references in a macro and
        in simple statically identifiable procedure calls.

        This is bounded static traversal, not VBA execution.
        """
        result = set()
        visited = set()

        def walk(name, depth):
            if depth > 8:
                return

            key = normalize_macro_name(name)

            if key in visited:
                return

            visited.add(key)

            proc = self.vba_procedures.get(key)

            if proc is None:
                return

            code = proc["code"]

            result.update(
                self.extract_vba_targets(code)
            )

            for other_key in self.vba_procedures:
                if other_key == key:
                    continue

                pattern = (
                    r"(?<![A-Za-z0-9_])"
                    + re.escape(other_key)
                    + r"(?![A-Za-z0-9_])"
                )

                if re.search(pattern, code, re.I):
                    walk(other_key, depth + 1)

        walk(macro_name, 0)
        return result

    # --------------------------------------------------------
    # Shape macro bindings
    # --------------------------------------------------------

    def scan_shapes(self):
        log("Scanning shapes and macro buttons...")

        for ws in self.book.Worksheets:
            source = str(ws.Name)

            try:
                shapes = ws.Shapes
                shape_count = shapes.Count
            except Exception:
                continue

            for index in range(1, shape_count + 1):
                try:
                    shape = shapes.Item(index)
                    shape_name = safe_str(shape.Name)

                    try:
                        macro_name = safe_str(shape.OnAction)
                    except Exception:
                        macro_name = ""

                    if not macro_name:
                        continue

                    self.stats["macro_bindings"] += 1

                    targets = self.resolve_macro_targets(
                        macro_name
                    )

                    if not targets:
                        self.add_unresolved(
                            source,
                            shape_name,
                            "VBA_Button",
                            macro_name,
                            "Macro target unresolved",
                        )
                        continue

                    for target in targets:
                        self.add_evidence(
                            source,
                            target,
                            "VBA_Button",
                            shape_name,
                            macro_name,
                            "Static VBA sheet reference",
                            navigation=True,
                        )

                except Exception as exc:
                    self.add_unresolved(
                        source,
                        "Shape " + str(index),
                        "Shape",
                        "",
                        safe_str(exc),
                    )

    # --------------------------------------------------------
    # Hyperlinks
    # --------------------------------------------------------

    def scan_hyperlinks(self):
        log("Scanning hyperlinks...")

        for ws in self.book.Worksheets:
            source = str(ws.Name)

            try:
                hyperlinks = ws.Hyperlinks
                count = hyperlinks.Count
            except Exception:
                continue

            for index in range(1, count + 1):
                try:
                    link = hyperlinks.Item(index)

                    subaddress = safe_str(
                        link.SubAddress
                    )

                    if not subaddress:
                        continue

                    targets = self.parse_reference_text(
                        subaddress
                    )

                    for target in targets:
                        self.add_evidence(
                            source,
                            target,
                            "Hyperlink",
                            "Hyperlink " + str(index),
                            subaddress,
                            "Internal workbook hyperlink",
                            navigation=True,
                        )

                except Exception:
                    pass

    # --------------------------------------------------------
    # Report
    # --------------------------------------------------------

    def style_header(self, ws):
        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor="17365D",
            )
            cell.font = Font(
                color="FFFFFF",
                bold=True,
            )
            cell.alignment = Alignment(
                horizontal="center",
            )

        ws.freeze_panes = "B2"

    def make_matrix(self, wb, name, links):
        ws = wb.create_sheet(name)

        ws.cell(1, 1, "Source / Target")

        for index, sheet in enumerate(self.sheet_names, 2):
            write_text(ws.cell(1, index), sheet)
            write_text(ws.cell(index, 1), sheet)

        for row_index, source in enumerate(
            self.sheet_names, 2
        ):
            for col_index, target in enumerate(
                self.sheet_names, 2
            ):
                value = int((source, target) in links)

                cell = ws.cell(
                    row_index,
                    col_index,
                    value,
                )

                if value == 1:
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="C6EFCE",
                    )

        self.style_header(ws)
        ws.column_dimensions["A"].width = 38

        for col in range(2, len(self.sheet_names) + 2):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 18

        return ws

    def write_report(self):
        log("Generating report...")

        wb = Workbook()
        summary = wb.active
        summary.title = "Summary"

        metrics = [
            ("Source File", self.source.name),
            ("Total Sheets", len(self.sheet_names)),
            ("Confirmed Data Links", len(self.data_links)),
            ("Navigation Links", len(self.navigation_links)),
            ("Formula Cells", self.stats["formula_cells"]),
            ("INDIRECT Calls", self.stats["indirect_calls"]),
            ("OFFSET Calls", self.stats["offset_calls"]),
            ("VBA Modules", self.vba_module_count),
            ("VBA Procedures", self.stats["vba_procedures"]),
            ("Macro Bindings", self.stats["macro_bindings"]),
            ("Evidence Rows", len(self.evidence)),
            ("Unresolved Rows", len(self.unresolved)),
            ("Truncated Evidence", self.stats["truncated_evidence"]),
            ("Truncated Unresolved", self.stats["truncated_unresolved"]),
        ]

        summary.append(["Metric", "Value"])

        for item in metrics:
            summary.append(item)

        self.style_header(summary)
        summary.column_dimensions["A"].width = 32
        summary.column_dimensions["B"].width = 35

        self.make_matrix(
            wb,
            "Direct_Data_Matrix",
            self.data_links,
        )

        self.make_matrix(
            wb,
            "Navigation_Matrix",
            self.navigation_links,
        )

        self.make_matrix(
            wb,
            "Combined_Matrix",
            self.data_links | self.navigation_links,
        )

        evidence_ws = wb.create_sheet("Evidence")

        evidence_ws.append([
            "Source Sheet",
            "Target Sheet",
            "Category",
            "Location",
            "Expression",
            "Details",
        ])

        for row in self.evidence:
            evidence_ws.append([
                row[0],
                row[1],
                row[2],
                row[3],
                "",
                row[5],
            ])
            write_text(
                evidence_ws.cell(
                    evidence_ws.max_row,
                    5,
                ),
                row[4],
            )

        self.style_header(evidence_ws)

        unresolved_ws = wb.create_sheet(
            "Unresolved"
        )

        unresolved_ws.append([
            "Sheet",
            "Location",
            "Category",
            "Expression",
            "Reason",
        ])

        for row in self.unresolved:
            unresolved_ws.append([
                row[0],
                row[1],
                row[2],
                "",
                row[4],
            ])

            write_text(
                unresolved_ws.cell(
                    unresolved_ws.max_row,
                    4,
                ),
                row[3],
            )

        self.style_header(unresolved_ws)

        sheets_ws = wb.create_sheet("Sheets")
        sheets_ws.append([
            "Sheet Name",
            "Visibility",
        ])

        for name in self.sheet_names:
            visibility = self.sheet_visibility[name]

            label = {
                -1: "Visible",
                0: "Hidden",
                2: "VeryHidden",
            }.get(visibility, str(visibility))

            sheets_ws.append([name, label])

        self.style_header(sheets_ws)

        for ws in (evidence_ws, unresolved_ws):
            ws.column_dimensions["A"].width = 35
            ws.column_dimensions["B"].width = 35
            ws.column_dimensions["C"].width = 22
            ws.column_dimensions["D"].width = 30
            ws.column_dimensions["E"].width = 90
            ws.column_dimensions["F"].width = 65
            ws.auto_filter.ref = ws.dimensions

        sheets_ws.column_dimensions["A"].width = 45
        sheets_ws.column_dimensions["B"].width = 20

        wb.save(str(self.output))

    # --------------------------------------------------------
    # Run
    # --------------------------------------------------------

    def run(self):
        if not self.source.exists():
            raise FileNotFoundError(self.source)

        if self.source == self.output:
            raise ValueError(
                "Output must not overwrite source workbook"
            )

        try:
            self.open_excel()

            self.scan_formulas()
            self.scan_defined_names()
            self.load_vba()
            self.scan_shapes()
            self.scan_hyperlinks()

        finally:
            self.close_excel()

        self.write_report()

        log()
        log("========== ANALYSIS RESULT ==========")
        log("Sheets:", len(self.sheet_names))
        log("Formula Cells:", self.stats["formula_cells"])
        log("INDIRECT Calls:", self.stats["indirect_calls"])
        log("OFFSET Calls:", self.stats["offset_calls"])
        log("VBA Modules:", self.vba_module_count)
        log("VBA Procedures:", self.stats["vba_procedures"])
        log("Macro Bindings:", self.stats["macro_bindings"])
        log("Confirmed Data Links:", len(self.data_links))
        log("Navigation Links:", len(self.navigation_links))
        log("Evidence Rows:", len(self.evidence))
        log("Unresolved Rows:", len(self.unresolved))
        log("Report:", self.output)
        log("=====================================")


def main():
    parser = argparse.ArgumentParser(
        description="Analyze XLSM direct sheet dependencies"
    )

    parser.add_argument(
        "source",
        help="Source .xlsm workbook",
    )

    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output .xlsx report",
    )

    parser.add_argument(
        "--include-self",
        action="store_true",
        help="Include same-sheet references",
    )

    args = parser.parse_args()

    source = Path(args.source)

    if args.output:
        output = Path(args.output)
    else:
        output = source.with_name(
            source.stem + "_dependency_report.xlsx"
        )

    analyzer = XLSMAnalyzer(
        source,
        output,
        include_self=args.include_self,
    )

    analyzer.run()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
