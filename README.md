
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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

