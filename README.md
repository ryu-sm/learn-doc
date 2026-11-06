python -c "from openpyxl import load_workbook; from collections import Counter; w=load_workbook('input_dependency_report.xlsx',read_only=True,data_only=True); s=w['Unresolved_References']; c=Counter(r[3] for r in list(s.values)[1:]); print(*c.most_common(),sep='\n')"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

python -c "from openpyxl import load_workbook; w=load_workbook('input_dependency_report.xlsx',read_only=True,data_only=True); s=w['Unresolved_References']; rows=[r for r in list(s.values)[1:] if r[3] and 'Dynamic formula' in str(r[3])]; [print(str(r[1]),str(r[2]),repr(r[4])[:300]) for r in rows[:20]]"




"""
Excel XLSM Direct Sheet Dependency Analyzer

Features:
- Analyze Excel formulas
- Extract and analyze VBA macros
- Track simple worksheet variables
- Detect direct cross-sheet assignments
- Detect Copy Destination operations
- Detect simple With blocks
- Record unresolved dynamic references
- Generate a dependency matrix
- Preserve the original XLSM file

Usage:
    python3 analyze_xlsm.py business.xlsm
    python3 analyze_xlsm.py business.xlsm -o report.xlsx

Dependencies:
    pip install openpyxl oletools
"""

import argparse
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
)
from openpyxl.utils import get_column_letter
from oletools.olevba import VBA_Parser


# ============================================================
# Configuration
# ============================================================

MATRIX_SHEET = "Direct_Matrix"
DETAIL_SHEET = "Dependency_Details"
UNKNOWN_SHEET = "Unresolved_References"

SHEET_EXPR = (
    r'(?:'
    r'(?:ThisWorkbook|ActiveWorkbook)\s*\.\s*'
    r')?'
    r'(?:Worksheets|Sheets)\s*'
    r'\(\s*"([^"]+)"\s*\)'
)

SHEET_PATTERN = re.compile(SHEET_EXPR, re.I)

SET_PATTERN = re.compile(
    r'^\s*Set\s+(\w+)\s*=\s*(.+)$',
    re.I,
)

FORMULA_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(?:'((?:[^']|'')+)'|"
    r"([A-Za-z_\\][A-Za-z0-9_.]*))!"
)

PROC_PATTERN = re.compile(
    r'^\s*(?:(?:Public|Private|Friend)\s+)?'
    r'(?:Sub|Function|Property\s+(?:Get|Let|Set))'
    r'\s+(\w+)',
    re.I,
)

END_PROC_PATTERN = re.compile(
    r'^\s*End\s+(?:Sub|Function|Property)\b',
    re.I,
)

COPY_PATTERN = re.compile(
    r'\.\s*Copy\b', re.I
)

DEST_PATTERN = re.compile(
    r'\bDestination\s*:=\s*(.+)$',
    re.I,
)

DYNAMIC_PATTERN = re.compile(
    r'\b(?:Worksheets|Sheets)\s*\(\s*'
    r'(?!")([^)]*)\)',
    re.I,
)


# ============================================================
# Analyzer
# ============================================================

class DependencyAnalyzer:

    def __init__(self, input_file):
        self.input_file = Path(input_file)

        self.wb = load_workbook(
            self.input_file,
            data_only=False,
            read_only=True,
            keep_links=True,
        )

        self.sheet_names = list(self.wb.sheetnames)

        self.name_map = {
            name.lower(): name
            for name in self.sheet_names
        }

        self.graph = {
            name: set()
            for name in self.sheet_names
        }

        self.details = set()
        self.unresolved = set()

    # --------------------------------------------------------
    # Dependency recording
    # --------------------------------------------------------

    def normalize_sheet(self, name):
        return self.name_map.get(name.lower())

    def add_dependency(
        self,
        destination,
        source,
        dep_type,
        module,
        location,
        expression,
    ):
        destination = self.normalize_sheet(destination)
        source = self.normalize_sheet(source)

        if not destination or not source:
            return

        if destination == source:
            return

        self.graph[destination].add(source)

        self.details.add((
            destination,
            source,
            dep_type,
            module,
            str(location),
            str(expression)[:32000],
        ))

    def add_unresolved(
        self, dep_type, module,
        location, reason, expression
    ):
        self.unresolved.add((
            dep_type,
            module,
            str(location),
            reason,
            str(expression)[:32000],
        ))

    # --------------------------------------------------------
    # Formula analysis
    # --------------------------------------------------------

    def formula_references(self, formula):
        result = set()

        # Ignore Excel text literals
        clean = re.sub(
            r'"(?:[^"]|"")*"',
            '""',
            formula
        )

        for match in FORMULA_PATTERN.finditer(clean):
            raw = match.group(1) or match.group(2)
            raw = raw.replace("''", "'")

            # 3D references
            if ":" in raw:
                parts = raw.split(":")

                if len(parts) == 2:
                    start = self.normalize_sheet(parts[0])
                    end = self.normalize_sheet(parts[1])

                    if start and end:
                        a = self.sheet_names.index(start)
                        b = self.sheet_names.index(end)

                        result.update(
                            self.sheet_names[
                                min(a, b):max(a, b) + 1
                            ]
                        )
                continue

            sheet = self.normalize_sheet(raw)

            if sheet:
                result.add(sheet)

        return result

    def analyze_formulas(self):
        print("\n[1/3] Analyzing Excel formulas...")

        for source in self.sheet_names:
            ws = self.wb[source]
            count = 0

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    count += 1
                    formula = str(cell.value)

                    refs = self.formula_references(formula)

                    for target in refs:
                        self.add_dependency(
                            source,
                            target,
                            "Excel Formula",
                            source,
                            cell.coordinate,
                            formula,
                        )

                    if re.search(
                        r'\b(?:INDIRECT|OFFSET)\s*\(',
                        formula, re.I
                    ):
                        self.add_unresolved(
                            "Excel Formula",
                            source,
                            cell.coordinate,
                            "Dynamic formula",
                            formula,
                        )

            print(
                f"  {source}: {count} formulas"
            )

    # --------------------------------------------------------
    # VBA helpers
    # --------------------------------------------------------

    @staticmethod
    def remove_comment(line):
        """Remove VBA comments outside strings."""
        in_string = False
        i = 0

        while i < len(line):
            ch = line[i]

            if ch == '"':
                if (
                    in_string
                    and i + 1 < len(line)
                    and line[i + 1] == '"'
                ):
                    i += 2
                    continue

                in_string = not in_string

            if ch == "'" and not in_string:
                return line[:i]

            i += 1

        return line

    def get_sheet_refs(self, expression, aliases):
        refs = set()

        # Direct Worksheets("Name")
        for match in SHEET_PATTERN.finditer(expression):
            name = self.normalize_sheet(match.group(1))
            if name:
                refs.add(name)

        # Variable references: ws.Range(...)
        for variable, name in aliases.items():
            pattern = (
                r'\b'
                + re.escape(variable)
                + r'\s*[.!]'
            )

            if re.search(pattern, expression, re.I):
                refs.add(name)

        return refs

    def resolve_sheet_object(self, expression, aliases):
        expression = expression.strip()

        # Direct sheet object
        match = SHEET_PATTERN.fullmatch(expression)
        if match:
            return self.normalize_sheet(match.group(1))

        # Existing alias
        return aliases.get(expression.lower())

    @staticmethod
    def split_assignment(line):
        """
        Find assignment operator outside strings.
        Basic heuristic; not a full VBA parser.
        """
        in_string = False

        for i, ch in enumerate(line):
            if ch == '"':
                if (
                    in_string
                    and i + 1 < len(line)
                    and line[i + 1] == '"'
                ):
                    continue
                in_string = not in_string

            if ch == "=" and not in_string:
                before = line[:i]
                after = line[i + 1:]

                if before.rstrip().endswith(
                    ("<", ">", ":")
                ):
                    continue

                return before.strip(), after.strip()

        return None

    # --------------------------------------------------------
    # VBA analysis
    # --------------------------------------------------------

    def analyze_vba(self):
        print("\n[2/3] Analyzing VBA macros...")

        parser = VBA_Parser(str(self.input_file))

        try:
            if not parser.detect_vba_macros():
                print("  No VBA macros found.")
                return

            for _, _, module, code in parser.extract_macros():
                print(f"  Module: {module}")
                self.analyze_vba_module(module, code)

        finally:
            parser.close()

    def analyze_vba_module(self, module, code):
        aliases = {}
        procedure = "(module)"
        with_stack = []

        # Merge continued VBA lines
        code = re.sub(
            r'\s+_\s*\r?\n\s*',
            ' ',
            code
        )

        for lineno, raw in enumerate(code.splitlines(), 1):
            line = self.remove_comment(raw).strip()

            if not line:
                continue

            # Procedure scope
            proc_match = PROC_PATTERN.match(line)

            if proc_match:
                procedure = proc_match.group(1)
                aliases = {}
                with_stack = []
                continue

            if END_PROC_PATTERN.match(line):
                procedure = "(module)"
                aliases = {}
                with_stack = []
                continue

            location = f"{procedure}:{lineno}"

            # Track With blocks
            if re.match(r'^With\s+', line, re.I):
                expression = re.sub(
                    r'^With\s+', '', line, flags=re.I
                )
                refs = self.get_sheet_refs(
                    expression, aliases
                )

                with_stack.append(
                    next(iter(refs)) if len(refs) == 1 else None
                )
                continue

            if re.match(r'^End\s+With\b', line, re.I):
                if with_stack:
                    with_stack.pop()
                continue

            # Track Set aliases
            set_match = SET_PATTERN.match(line)

            if set_match:
                variable = set_match.group(1).lower()
                expression = set_match.group(2)

                resolved = self.resolve_sheet_object(
                    expression, aliases
                )

                if resolved:
                    aliases[variable] = resolved
                else:
                    aliases.pop(variable, None)

                    if re.search(
                        r'\b(?:Worksheets|Sheets)\s*\(',
                        expression, re.I
                    ):
                        self.add_unresolved(
                            "VBA",
                            module,
                            location,
                            "Dynamic or unsupported Set",
                            line,
                        )
                continue

            # Detect Copy Destination
            if COPY_PATTERN.search(line):
                self.analyze_copy(
                    line,
                    aliases,
                    module,
                    location,
                    with_stack,
                )
                continue

            # Detect simple assignments
            self.analyze_assignment(
                line,
                aliases,
                module,
                location,
                with_stack,
            )

            # Dynamic sheet access
            if DYNAMIC_PATTERN.search(line):
                self.add_unresolved(
                    "VBA",
                    module,
                    location,
                    "Dynamic worksheet name",
                    line,
                )

    # --------------------------------------------------------
    # VBA assignment analysis
    # --------------------------------------------------------

    def analyze_assignment(
        self,
        line,
        aliases,
        module,
        location,
        with_stack,
    ):
        if re.match(
            r'^(If|ElseIf|For|Do|While|Select|Case)\b',
            line, re.I
        ):
            return

        assignment = self.split_assignment(line)

        if not assignment:
            return

        left, right = assignment

        left = re.sub(
            r'^(?:Let\s+)',
            '',
            left,
            flags=re.I
        )

        destinations = self.get_sheet_refs(
            left, aliases
        )
        sources = self.get_sheet_refs(
            right, aliases
        )

        # With ws: .Range(...) = ...
        if with_stack and with_stack[-1]:
            if left.startswith("."):
                destinations.add(with_stack[-1])

            if right.startswith("."):
                sources.add(with_stack[-1])

        if destinations and sources:
            for dst in destinations:
                for src in sources:
                    self.add_dependency(
                        dst,
                        src,
                        "VBA Assignment",
                        module,
                        location,
                        line,
                    )

        elif destinations or sources:
            self.add_unresolved(
                "VBA",
                module,
                location,
                "One-sided or unsupported assignment",
                line,
            )

    # --------------------------------------------------------
    # VBA Copy analysis
    # --------------------------------------------------------

    def analyze_copy(
        self,
        line,
        aliases,
        module,
        location,
        with_stack,
    ):
        match = COPY_PATTERN.search(line)

        if not match:
            return

        source_expr = line[:match.start()]
        rest = line[match.end():]

        sources = self.get_sheet_refs(
            source_expr, aliases
        )

        if (
            not sources
            and source_expr.startswith(".")
            and with_stack
            and with_stack[-1]
        ):
            sources.add(with_stack[-1])

        destination_match = DEST_PATTERN.search(rest)

        if destination_match:
            destination_expr = destination_match.group(1)

            destinations = self.get_sheet_refs(
                destination_expr, aliases
            )

            if destinations and sources:
                for dst in destinations:
                    for src in sources:
                        self.add_dependency(
                            dst,
                            src,
                            "VBA Copy",
                            module,
                            location,
                            line,
                        )
                return

        self.add_unresolved(
            "VBA",
            module,
            location,
            "Copy destination not resolved",
            line,
        )

    # --------------------------------------------------------
    # Report generation
    # --------------------------------------------------------

    @staticmethod
    def format_header(ws):
        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78"
            )
            cell.font = Font(
                color="FFFFFF",
                bold=True
            )
            cell.alignment = Alignment(
                horizontal="center"
            )

    def create_matrix(self, report):
        ws = report.active
        ws.title = MATRIX_SHEET

        ws.cell(1, 1, "Source / Target")

        for i, name in enumerate(self.sheet_names, 2):
            ws.cell(1, i, name)
            ws.cell(i, 1, name)

        for i, source in enumerate(self.sheet_names, 2):
            for j, target in enumerate(self.sheet_names, 2):
                value = int(
                    target in self.graph[source]
                )

                cell = ws.cell(i, j, value)

                cell.alignment = Alignment(
                    horizontal="center"
                )

                if value == 1:
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="C6EFCE"
                    )
                    cell.font = Font(
                        color="006100",
                        bold=True
                    )

        self.format_header(ws)

        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 30

        for col in range(2, len(self.sheet_names) + 2):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 20

    def create_details(self, report):
        ws = report.create_sheet(DETAIL_SHEET)

        ws.append([
            "Source Sheet",
            "Target Sheet",
            "Dependency Type",
            "Module / Sheet",
            "Location",
            "Expression",
        ])

        for item in sorted(self.details):
            ws.append(item)

        self.format_header(ws)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = [25, 25, 22, 30, 20, 90]

        for i, width in enumerate(widths, 1):
            ws.column_dimensions[
                get_column_letter(i)
            ].width = width

    def create_unresolved(self, report):
        ws = report.create_sheet(UNKNOWN_SHEET)

        ws.append([
            "Type",
            "Module / Sheet",
            "Location",
            "Reason",
            "Expression",
        ])

        for item in sorted(self.unresolved):
            ws.append(item)

        self.format_header(ws)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = [20, 30, 20, 40, 90]

        for i, width in enumerate(widths, 1):
            ws.column_dimensions[
                get_column_letter(i)
            ].width = width

    def generate_report(self, output_file):
        print("\n[3/3] Generating report...")

        report = Workbook()

        self.create_matrix(report)
        self.create_details(report)
        self.create_unresolved(report)

        report.save(output_file)

        print("\n========== RESULT ==========")
        print("Sheets:", len(self.sheet_names))
        print(
            "Direct dependencies:",
            sum(len(v) for v in self.graph.values())
        )
        print("Evidence records:", len(self.details))
        print("Unresolved:", len(self.unresolved))
        print("Output:", output_file)
        print("============================")

    def run(self, output_file):
        try:
            self.analyze_formulas()
            self.analyze_vba()
            self.generate_report(output_file)
        finally:
            self.wb.close()


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Analyze direct sheet dependencies "
            "in XLSM/XLSX files"
        )
    )

    parser.add_argument(
        "input",
        help="Input Excel file"
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output report file"
    )

    args = parser.parse_args()

    input_path = Path(args.input)

    if not input_path.exists():
        parser.error(
            f"File not found: {input_path}"
        )

    if input_path.suffix.lower() not in (
        ".xlsx", ".xlsm"
    ):
        parser.error(
            "Only .xlsx and .xlsm are supported"
        )

    output_path = (
        Path(args.output)
        if args.output
        else input_path.with_name(
            input_path.stem
            + "_dependency_report.xlsx"
        )
    )

    if output_path.suffix.lower() != ".xlsx":
        parser.error(
            "Report output must be an .xlsx file"
        )

    if input_path.resolve() == output_path.resolve():
        parser.error(
            "Output must differ from input"
        )

    analyzer = DependencyAnalyzer(input_path)
    analyzer.run(output_path)


if __name__ == "__main__":
    main()







import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from oletools.olevba import VBA_Parser


def diagnose(file_path):
    path = Path(file_path)

    wb = load_workbook(
        path,
        read_only=False,
        data_only=False,
        keep_vba=True
    )

    print("\n========== SHEETS ==========")

    for ws in wb.worksheets:
        print(
            f"Name: {ws.title} | "
            f"CodeName: {ws.sheet_properties.codeName}"
        )

    print("\n========== VBA MODULES ==========")

    parser = VBA_Parser(str(path))

    patterns = {
        "Worksheets/Sheets":
            r"\b(?:Worksheets|Sheets)\s*\(",
        "CodeName":
            r"\bSheet\d+\s*\.",
        "Set Worksheet":
            r"\bSet\s+\w+\s*=",
        "Range":
            r"\bRange\s*\(",
        "Cells":
            r"\bCells\s*\(",
        "Copy":
            r"\.Copy\b",
        "Paste":
            r"\bPaste(?:Special)?\b",
        "With":
            r"\bWith\b",
        "ActiveSheet":
            r"\bActiveSheet\b",
    }

    try:
        has_macros = parser.detect_vba_macros()
        print("VBA detected:", has_macros)

        if has_macros:
            for _, _, module, code in parser.extract_macros():

                print(f"\nModule: {module}")
                print(f"Code length: {len(code)}")

                for name, pattern in patterns.items():
                    matches = re.findall(
                        pattern, code, re.I
                    )
                    print(
                        f"  {name}: {len(matches)}"
                    )

                print("  --- Code preview ---")
                print(code[:1500])

    finally:
        parser.close()
        wb.close()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(
            "Usage: python3 diagnose_xlsm.py file.xlsm"
        )
        sys.exit(1)

    diagnose(sys.argv[1])





#!/usr/bin/env python3
"""Conservative static direct-sheet dependency report for xlsx/xlsm.

The source workbook is read-only and never saved.  This is not a complete VBA
interpreter; unresolved patterns are reported rather than silently invented.
"""
import argparse
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from oletools.olevba import VBA_Parser

FORMULA_REF = re.compile(r"(?<![\w.])(?:'((?:[^']|'')+)'|([A-Za-z_\\][\w.]*))!")
SHEET_OBJ = re.compile(r'(?:(?:ThisWorkbook|ActiveWorkbook)\s*\.\s*)?(?:Worksheets|Sheets)\s*\(\s*"((?:[^"]|"")+)"\s*\)', re.I)
DYNAMIC_OBJ = re.compile(r'\b(?:Worksheets|Sheets)\s*\(\s*(?!")([^)]*)\)', re.I)
SET_RE = re.compile(r'^Set\s+(\w+)\s*=\s*(.+)$', re.I)
PROC_RE = re.compile(r'^(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)', re.I)
END_PROC = re.compile(r'^End\s+(?:Sub|Function|Property)\b', re.I)
COPY_RE = re.compile(r'\.Copy\b', re.I)
DEST_RE = re.compile(r'\bDestination\s*:=\s*(.+)', re.I)
REPORT_NAMES = {'Direct_Matrix', 'Dependency_Details', 'Unresolved_References'}


def safe_text(cell, value):
    """Force literal text even when the string starts with =, +, -, or @."""
    cell.value = str(value)[:32000]
    cell.data_type = 's'
    cell.number_format = '@'


def strip_vba_comment(line):
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        if not quoted and line[i] == "'":
            return line[:i]
        i += 1
    return line


def strip_formula_strings(s):
    return re.sub(r'"(?:[^"]|"")*"', '""', s)


def logical_lines(code):
    pending = ''
    first = 0
    for no, raw in enumerate(code.splitlines(), 1):
        s = strip_vba_comment(raw).strip()
        if not s:
            continue
        if not pending:
            first = no
        if re.search(r'\s+_$', s):
            pending += re.sub(r'\s+_$', ' ', s)
        else:
            yield first, pending + s
            pending = ''
    if pending:
        yield first, pending


def split_assignment(line):
    if re.match(r'^(?:If|ElseIf|For|Do|While|Select|Case|Const|Dim|ReDim|Debug|Call|Exit)\b', line, re.I):
        return None
    quoted = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if quoted and i + 1 < len(line) and line[i+1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif line[i] == '=' and not quoted:
            if (i and line[i-1] in '<>:') or (i+1 < len(line) and line[i+1] == '='):
                i += 1
                continue
            return line[:i].strip(), line[i+1:].strip()
        i += 1
    return None


class Analyzer:
    def __init__(self, path):
        self.path = Path(path)
        self.book = load_workbook(self.path, data_only=False, read_only=False, keep_links=True)
        self.names = list(self.book.sheetnames)
        self.by_lower = {s.casefold(): s for s in self.names}
        self.code_names = {}
        for ws in self.book.worksheets:
            codename = ws.sheet_properties.codeName
            if codename:
                self.code_names[codename.casefold()] = ws.title
        self.graph = {s: set() for s in self.names}
        self.details = set()
        self.unknown = set()
        self.stats = {'formula_cells': 0, 'vba_modules': 0, 'vba_lines': 0}

    def normalize(self, s):
        return self.by_lower.get(s.casefold()) if isinstance(s, str) else None

    def add(self, dst, src, typ, module, location, expression):
        dst, src = self.normalize(dst), self.normalize(src)
        if dst and src and dst != src:
            self.graph[dst].add(src)
            self.details.add((dst, src, typ, module, str(location), str(expression)))

    def unresolved(self, typ, module, location, reason, expression):
        self.unknown.add((typ, module, str(location), reason, str(expression)))

    def formula_refs(self, formula):
        refs = set()
        cleaned = strip_formula_strings(formula)
        for m in FORMULA_REF.finditer(cleaned):
            token = (m.group(1) or m.group(2)).replace("''", "'")
            if ':' in token:
                parts = token.split(':')
                if len(parts) == 2:
                    a, b = map(self.normalize, parts)
                    if a and b:
                        i, j = self.names.index(a), self.names.index(b)
                        refs.update(self.names[min(i,j):max(i,j)+1])
                    else:
                        refs.add(None)
            else:
                name = self.normalize(token)
                if name:
                    refs.add(name)
        return refs

    def formulas(self):
        for sheet in self.names:
            ws = self.book[sheet]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != 'f':
                        continue
                    self.stats['formula_cells'] += 1
                    expression = str(cell.value)
                    for target in self.formula_refs(expression):
                        if target is None:
                            self.unresolved('Formula', sheet, cell.coordinate, 'Unresolved 3D sheet range', expression)
                        else:
                            self.add(sheet, target, 'Excel formula', sheet, cell.coordinate, expression)
                    if re.search(r'\b(?:INDIRECT|OFFSET)\s*\(', expression, re.I):
                        self.unresolved('Formula', sheet, cell.coordinate, 'Dynamic formula; runtime target unknown', expression)

    def refs(self, expression, aliases, with_sheet=None):
        result = set()
        for m in SHEET_OBJ.finditer(expression):
            name = self.normalize(m.group(1).replace('""', '"'))
            if name:
                result.add(name)
        for code, sheet in self.code_names.items():
            if re.search(r'(?<![\w.])' + re.escape(code) + r'\s*[.!]', expression, re.I):
                result.add(sheet)
        for alias, sheet in aliases.items():
            if re.search(r'(?<![\w.])' + re.escape(alias) + r'\s*[.!]', expression, re.I):
                result.add(sheet)
        if with_sheet and expression.lstrip().startswith('.'):
            result.add(with_sheet)
        return result

    def resolve_object(self, expr, aliases):
        expr = expr.strip()
        m = SHEET_OBJ.fullmatch(expr)
        if m:
            return self.normalize(m.group(1).replace('""', '"'))
        return aliases.get(expr.casefold()) or self.code_names.get(expr.casefold())

    def vba(self):
        parser = VBA_Parser(str(self.path))
        try:
            if not parser.detect_vba_macros():
                self.unresolved('VBA', '(workbook)', '-', 'No extractable VBA macros detected', '')
                return
            for _, _, module, code in parser.extract_macros():
                self.stats['vba_modules'] += 1
                aliases = {}
                with_stack = []
                proc = '(module)'
                for line_no, line in logical_lines(code):
                    self.stats['vba_lines'] += 1
                    loc = f'{proc}:{line_no}'
                    m = PROC_RE.match(line)
                    if m:
                        proc = m.group(1)
                        aliases = {}
                        with_stack = []
                        continue
                    if END_PROC.match(line):
                        proc = '(module)'
                        aliases = {}
                        with_stack = []
                        continue
                    if re.match(r'^With\s+', line, re.I):
                        expr = re.sub(r'^With\s+', '', line, flags=re.I)
                        found = self.refs(expr, aliases, with_stack[-1] if with_stack else None)
                        with_stack.append(next(iter(found)) if len(found) == 1 else None)
                        if len(found) != 1:
                            self.unresolved('VBA', module, loc, 'With target not resolved', line)
                        continue
                    if re.match(r'^End\s+With\b', line, re.I):
                        if with_stack:
                            with_stack.pop()
                        continue
                    m = SET_RE.match(line)
                    if m:
                        var, expr = m.groups()
                        sheet = self.resolve_object(expr, aliases)
                        if sheet:
                            aliases[var.casefold()] = sheet
                        else:
                            aliases.pop(var.casefold(), None)
                            if re.search(r'\b(?:Sheets|Worksheets)\s*\(', expr, re.I):
                                self.unresolved('VBA', module, loc, 'Unresolved worksheet object assignment', line)
                        continue
                    with_sheet = with_stack[-1] if with_stack else None
                    copy = COPY_RE.search(line)
                    if copy:
                        src = self.refs(line[:copy.start()], aliases, with_sheet)
                        dest_match = DEST_RE.search(line[copy.end():])
                        dst = self.refs(dest_match.group(1), aliases, with_sheet) if dest_match else set()
                        if src and dst:
                            for d in dst:
                                for s in src:
                                    self.add(d, s, 'VBA Copy Destination', module, loc, line)
                        else:
                            self.unresolved('VBA', module, loc, 'Copy/Paste target or source unresolved', line)
                        continue
                    assignment = split_assignment(line)
                    if assignment:
                        left, right = assignment
                        left = re.sub(r'^Let\s+', '', left, flags=re.I)
                        dst = self.refs(left, aliases, with_sheet)
                        src = self.refs(right, aliases, with_sheet)
                        if dst and src:
                            for d in dst:
                                for s in src:
                                    self.add(d, s, 'VBA assignment', module, loc, line)
                        elif dst or src:
                            self.unresolved('VBA', module, loc, 'One-sided assignment; data flow not confirmed', line)
                    if DYNAMIC_OBJ.search(line):
                        self.unresolved('VBA', module, loc, 'Dynamic sheet selector', line)
        finally:
            parser.close()

    @staticmethod
    def style_header(ws):
        for cell in ws[1]:
            cell.fill = PatternFill('solid', fgColor='1F4E78')
            cell.font = Font(color='FFFFFF', bold=True)
            cell.alignment = Alignment(horizontal='center')

    def report(self, output):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Direct_Matrix'
        safe_text(ws.cell(1, 1), 'Source / Target')
        for i, name in enumerate(self.names, 2):
            safe_text(ws.cell(1, i), name)
            safe_text(ws.cell(i, 1), name)
        for i, source in enumerate(self.names, 2):
            for j, target in enumerate(self.names, 2):
                cell = ws.cell(i, j, int(target in self.graph[source]))
                cell.alignment = Alignment(horizontal='center')
                if cell.value:
                    cell.fill = PatternFill('solid', fgColor='C6EFCE')
                    cell.font = Font(color='006100', bold=True)
        self.style_header(ws)
        ws.freeze_panes = 'B2'
        ws.column_dimensions['A'].width = 32
        for col in range(2, len(self.names) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 22

        for title, headers, records, widths in [
            ('Dependency_Details', ['Source Sheet', 'Target Sheet', 'Type', 'Module / Sheet', 'Location', 'Expression'], self.details, [25,25,24,30,22,95]),
            ('Unresolved_References', ['Type', 'Module / Sheet', 'Location', 'Reason', 'Expression'], self.unknown, [22,30,22,50,95]),
        ]:
            sheet = wb.create_sheet(title)
            for j, h in enumerate(headers, 1):
                safe_text(sheet.cell(1, j), h)
            for i, record in enumerate(sorted(records), 2):
                for j, value in enumerate(record, 1):
                    safe_text(sheet.cell(i, j), value)
            self.style_header(sheet)
            sheet.freeze_panes = 'A2'
            sheet.auto_filter.ref = sheet.dimensions
            for j, width in enumerate(widths, 1):
                sheet.column_dimensions[get_column_letter(j)].width = width
        wb.save(output)

    def run(self, output):
        try:
            self.formulas()
            self.vba()
            self.report(output)
            print('Formula cells:', self.stats['formula_cells'])
            print('VBA modules:', self.stats['vba_modules'])
            print('VBA logical lines:', self.stats['vba_lines'])
            print('Confirmed direct sheet links:', sum(map(len, self.graph.values())))
            print('Evidence rows:', len(self.details))
            print('Unresolved rows:', len(self.unknown))
            print('Report:', output)
            if not self.details:
                print('WARNING: No confirmed direct dependencies. Check Unresolved_References and VBA extraction.')
        finally:
            self.book.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', help='Source .xlsx or .xlsm')
    parser.add_argument('-o', '--output', help='Output .xlsx report')
    args = parser.parse_args()
    source = Path(args.input)
    if not source.is_file() or source.suffix.lower() not in ('.xlsx', '.xlsm'):
        parser.error('Input must be an existing .xlsx or .xlsm file')
    output = Path(args.output) if args.output else source.with_name(source.stem + '_dependency_report.xlsx')
    if output.suffix.lower() != '.xlsx' or output.resolve() == source.resolve():
        parser.error('Output must be a different .xlsx path')
    Analyzer(source).run(output)


if __name__ == '__main__':
    main()

