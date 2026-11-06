
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM / XLSX Direct Sheet Dependency Analyzer

分析内容：
1. Excel 普通公式中的直接 Sheet 引用
2. 日文、中文及带空格的 Sheet 名称
3. OFFSET 中的显式 Sheet 引用
4. INDIRECT 中的固定字符串引用
5. VBA Worksheets / Sheets 引用
6. VBA Sheet CodeName 引用
7. VBA 简单变量赋值和跨 Sheet 数据传递
8. VBA Copy Destination
9. 无法确定的动态引用

输出：
- Direct_Matrix
- Dependency_Details
- Unresolved_References

说明：
- 只计算直接依赖，不计算间接依赖
- 不执行 VBA 宏
- 不修改原始 Excel 文件
- 静态分析无法保证覆盖所有复杂 VBA 数据流
"""

import argparse
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from oletools.olevba import VBA_Parser


# ============================================================
# Regular expressions
# ============================================================

FORMULA_REF = re.compile(
    r"(?<![\w.])"
    r"(?:"
    r"'((?:[^']|'')+)'"
    r"|"
    r"([^\s=+\-*/^&(),;:{}\[\]!'\"<>]+)"
    r")!",
    re.UNICODE,
)

SHEET_OBJ = re.compile(
    r'(?:(?:ThisWorkbook|ActiveWorkbook)\s*\.\s*)?'
    r'(?:Worksheets|Sheets)\s*'
    r'\(\s*"((?:[^"]|"")+)"\s*\)',
    re.I,
)

DYNAMIC_OBJ = re.compile(
    r'\b(?:Worksheets|Sheets)\s*'
    r'\(\s*(?!")([^)]*)\)',
    re.I,
)

SET_RE = re.compile(
    r'^Set\s+(\w+)\s*=\s*(.+)$',
    re.I,
)

PROC_RE = re.compile(
    r'^(?:(?:Public|Private|Friend|Static)\s+)*'
    r'(?:Sub|Function|Property\s+(?:Get|Let|Set))'
    r'\s+(\w+)',
    re.I,
)

END_PROC = re.compile(
    r'^End\s+(?:Sub|Function|Property)\b',
    re.I,
)

COPY_RE = re.compile(
    r'\.Copy\b',
    re.I,
)

DEST_RE = re.compile(
    r'\bDestination\s*:=\s*(.+)',
    re.I,
)


# ============================================================
# Utility functions
# ============================================================

def safe_text(cell, value):
    """
    Force literal text in Excel.

    Prevent strings beginning with '=' from being
    interpreted as Excel formulas.
    """
    cell.value = str(value)[:32000]
    cell.data_type = "s"
    cell.number_format = "@"


def strip_vba_comment(line):
    """Remove VBA comments outside quoted strings."""
    quoted = False
    i = 0

    while i < len(line):
        if line[i] == '"':
            if (
                quoted
                and i + 1 < len(line)
                and line[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted

        if not quoted and line[i] == "'":
            return line[:i]

        i += 1

    return line


def strip_formula_strings(formula):
    """Remove Excel string literals."""
    return re.sub(
        r'"(?:[^"]|"")*"',
        '""',
        formula,
    )


def logical_lines(code):
    """Merge VBA continuation lines."""
    pending = ""
    first_line = 0

    for number, raw in enumerate(
        code.splitlines(), 1
    ):
        line = strip_vba_comment(raw).strip()

        if not line:
            continue

        if not pending:
            first_line = number

        if re.search(r'\s+_$', line):
            pending += re.sub(
                r'\s+_$',
                ' ',
                line,
            )
        else:
            yield first_line, pending + line
            pending = ""

    if pending:
        yield first_line, pending


def split_assignment(line):
    """Detect basic VBA assignments."""
    if re.match(
        r'^(?:If|ElseIf|For|Do|While|Select|'
        r'Case|Const|Dim|ReDim|Debug|Call|Exit)\b',
        line,
        re.I,
    ):
        return None

    quoted = False
    i = 0

    while i < len(line):
        if line[i] == '"':
            if (
                quoted
                and i + 1 < len(line)
                and line[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted

        elif line[i] == "=" and not quoted:
            if (
                (i and line[i - 1] in "<>:")
                or (
                    i + 1 < len(line)
                    and line[i + 1] == "="
                )
            ):
                i += 1
                continue

            return (
                line[:i].strip(),
                line[i + 1:].strip(),
            )

        i += 1

    return None


# ============================================================
# Dependency Analyzer
# ============================================================

class Analyzer:

    def __init__(self, path):
        self.path = Path(path)

        self.book = load_workbook(
            self.path,
            data_only=False,
            read_only=False,
            keep_links=True,
        )

        self.names = list(self.book.sheetnames)

        self.by_lower = {
            name.casefold(): name
            for name in self.names
        }

        # VBA CodeName -> Sheet display name
        self.code_names = {}

        for ws in self.book.worksheets:
            code_name = ws.sheet_properties.codeName

            if code_name:
                self.code_names[
                    code_name.casefold()
                ] = ws.title

        self.graph = {
            name: set()
            for name in self.names
        }

        self.details = set()
        self.unknown = set()

        self.stats = {
            "formula_cells": 0,
            "formula_cross_sheet_cells": 0,
            "offset_cells": 0,
            "indirect_cells": 0,
            "indirect_literal_resolved": 0,
            "vba_modules": 0,
            "vba_lines": 0,
        }

    # --------------------------------------------------------
    # Common methods
    # --------------------------------------------------------

    def normalize(self, name):
        if isinstance(name, str):
            return self.by_lower.get(
                name.casefold()
            )
        return None

    def add(
        self,
        destination,
        source,
        dep_type,
        module,
        location,
        expression,
    ):
        destination = self.normalize(destination)
        source = self.normalize(source)

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
            str(expression),
        ))

    def unresolved(
        self,
        dep_type,
        module,
        location,
        reason,
        expression,
    ):
        self.unknown.add((
            dep_type,
            module,
            str(location),
            reason,
            str(expression),
        ))

    # --------------------------------------------------------
    # Formula reference extraction
    # --------------------------------------------------------

    def formula_refs(
        self,
        formula,
        strip_strings=True,
    ):
        refs = set()

        clean = (
            strip_formula_strings(formula)
            if strip_strings
            else formula
        )

        for match in FORMULA_REF.finditer(clean):
            token = (
                match.group(1)
                or match.group(2)
            )

            token = token.replace("''", "'")

            # Handle 3D sheet references
            if ":" in token:
                parts = token.split(":")

                if len(parts) == 2:
                    start = self.normalize(parts[0])
                    end = self.normalize(parts[1])

                    if start and end:
                        a = self.names.index(start)
                        b = self.names.index(end)

                        refs.update(
                            self.names[
                                min(a, b):max(a, b) + 1
                            ]
                        )
                    else:
                        refs.add(None)

                continue

            name = self.normalize(token)

            if name:
                refs.add(name)

        return refs

    # --------------------------------------------------------
    # Excel formula analysis
    # --------------------------------------------------------

    def formulas(self):
        print("[1/3] Analyzing Excel formulas...")

        for sheet in self.names:
            ws = self.book[sheet]

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["formula_cells"] += 1

                    expression = str(cell.value)

                    refs = self.formula_refs(expression)

                    cross_refs = refs - {None, sheet}

                    if cross_refs:
                        self.stats[
                            "formula_cross_sheet_cells"
                        ] += 1

                    for target in refs:
                        if target is None:
                            self.unresolved(
                                "Formula",
                                sheet,
                                cell.coordinate,
                                "Unresolved 3D sheet range",
                                expression,
                            )
                        else:
                            self.add(
                                sheet,
                                target,
                                "Excel formula",
                                sheet,
                                cell.coordinate,
                                expression,
                            )

                    # OFFSET with explicit sheet references
                    # is already handled by formula_refs().
                    if re.search(
                        r'\bOFFSET\s*\(',
                        expression,
                        re.I,
                    ):
                        self.stats["offset_cells"] += 1

                    # INDIRECT analysis
                    if re.search(
                        r'\bINDIRECT\s*\(',
                        expression,
                        re.I,
                    ):
                        self.stats["indirect_cells"] += 1

                        literal_refs = set()

                        literals = re.findall(
                            r'"((?:[^"]|"")*)"',
                            expression,
                        )

                        for literal in literals:
                            literal = literal.replace(
                                '""', '"'
                            )

                            found = self.formula_refs(
                                literal,
                                strip_strings=False,
                            )

                            literal_refs.update(
                                name
                                for name in found
                                if name is not None
                            )

                        for target in literal_refs:
                            self.add(
                                sheet,
                                target,
                                "INDIRECT literal",
                                sheet,
                                cell.coordinate,
                                expression,
                            )

                        if literal_refs:
                            self.stats[
                                "indirect_literal_resolved"
                            ] += 1
                        else:
                            self.unresolved(
                                "Formula",
                                sheet,
                                cell.coordinate,
                                "INDIRECT target not "
                                "statically resolved",
                                expression,
                            )

    # --------------------------------------------------------
    # VBA reference extraction
    # --------------------------------------------------------

    def refs(
        self,
        expression,
        aliases,
        with_sheet=None,
    ):
        result = set()

        # Worksheets("SheetName")
        for match in SHEET_OBJ.finditer(expression):
            name = self.normalize(
                match.group(1).replace('""', '"')
            )

            if name:
                result.add(name)

        # Sheet CodeName
        for code, sheet in self.code_names.items():
            pattern = (
                r'(?<![\w.])'
                + re.escape(code)
                + r'\s*[.!]'
            )

            if re.search(
                pattern,
                expression,
                re.I,
            ):
                result.add(sheet)

        # Worksheet variable aliases
        for alias, sheet in aliases.items():
            pattern = (
                r'(?<![\w.])'
                + re.escape(alias)
                + r'\s*[.!]'
            )

            if re.search(
                pattern,
                expression,
                re.I,
            ):
                result.add(sheet)

        # With ws ... End With
        if (
            with_sheet
            and expression.lstrip().startswith(".")
        ):
            result.add(with_sheet)

        return result

    def resolve_object(
        self,
        expression,
        aliases,
    ):
        expression = expression.strip()

        match = SHEET_OBJ.fullmatch(expression)

        if match:
            return self.normalize(
                match.group(1).replace('""', '"')
            )

        return (
            aliases.get(expression.casefold())
            or self.code_names.get(
                expression.casefold()
            )
        )

    # --------------------------------------------------------
    # VBA macro analysis
    # --------------------------------------------------------

    def vba(self):
        print("[2/3] Analyzing VBA macros...")

        parser = VBA_Parser(str(self.path))

        try:
            if not parser.detect_vba_macros():
                self.unresolved(
                    "VBA",
                    "(workbook)",
                    "-",
                    "No extractable VBA macros detected",
                    "",
                )
                return

            for _, _, module, code in parser.extract_macros():
                self.stats["vba_modules"] += 1

                aliases = {}
                with_stack = []
                procedure = "(module)"

                for line_no, line in logical_lines(code):
                    self.stats["vba_lines"] += 1

                    location = (
                        f"{procedure}:{line_no}"
                    )

                    # Procedure start
                    match = PROC_RE.match(line)

                    if match:
                        procedure = match.group(1)
                        aliases = {}
                        with_stack = []
                        continue

                    # Procedure end
                    if END_PROC.match(line):
                        procedure = "(module)"
                        aliases = {}
                        with_stack = []
                        continue

                    # With block
                    if re.match(
                        r'^With\s+',
                        line,
                        re.I,
                    ):
                        expression = re.sub(
                            r'^With\s+',
                            '',
                            line,
                            flags=re.I,
                        )

                        parent = (
                            with_stack[-1]
                            if with_stack
                            else None
                        )

                        found = self.refs(
                            expression,
                            aliases,
                            parent,
                        )

                        resolved = (
                            next(iter(found))
                            if len(found) == 1
                            else None
                        )

                        with_stack.append(resolved)

                        if resolved is None:
                            self.unresolved(
                                "VBA",
                                module,
                                location,
                                "With target not resolved",
                                line,
                            )

                        continue

                    if re.match(
                        r'^End\s+With\b',
                        line,
                        re.I,
                    ):
                        if with_stack:
                            with_stack.pop()
                        continue

                    # Set ws = Worksheets("...")
                    match = SET_RE.match(line)

                    if match:
                        variable, expression = match.groups()

                        sheet = self.resolve_object(
                            expression,
                            aliases,
                        )

                        if sheet:
                            aliases[
                                variable.casefold()
                            ] = sheet
                        else:
                            aliases.pop(
                                variable.casefold(),
                                None,
                            )

                            if re.search(
                                r'\b(?:Sheets|Worksheets)'
                                r'\s*\(',
                                expression,
                                re.I,
                            ):
                                self.unresolved(
                                    "VBA",
                                    module,
                                    location,
                                    "Unresolved worksheet "
                                    "object assignment",
                                    line,
                                )

                        continue

                    with_sheet = (
                        with_stack[-1]
                        if with_stack
                        else None
                    )

                    # Copy Destination
                    copy = COPY_RE.search(line)

                    if copy:
                        source_expr = line[:copy.start()]

                        sources = self.refs(
                            source_expr,
                            aliases,
                            with_sheet,
                        )

                        dest_match = DEST_RE.search(
                            line[copy.end():]
                        )

                        destinations = (
                            self.refs(
                                dest_match.group(1),
                                aliases,
                                with_sheet,
                            )
                            if dest_match
                            else set()
                        )

                        if sources and destinations:
                            for dst in destinations:
                                for src in sources:
                                    self.add(
                                        dst,
                                        src,
                                        "VBA Copy Destination",
                                        module,
                                        location,
                                        line,
                                    )
                        else:
                            self.unresolved(
                                "VBA",
                                module,
                                location,
                                "Copy/Paste target or "
                                "source unresolved",
                                line,
                            )

                        continue

                    # Simple VBA assignment
                    assignment = split_assignment(line)

                    if assignment:
                        left, right = assignment

                        left = re.sub(
                            r'^Let\s+',
                            '',
                            left,
                            flags=re.I,
                        )

                        destinations = self.refs(
                            left,
                            aliases,
                            with_sheet,
                        )

                        sources = self.refs(
                            right,
                            aliases,
                            with_sheet,
                        )

                        if destinations and sources:
                            for dst in destinations:
                                for src in sources:
                                    self.add(
                                        dst,
                                        src,
                                        "VBA assignment",
                                        module,
                                        location,
                                        line,
                                    )

                        elif destinations or sources:
                            self.unresolved(
                                "VBA",
                                module,
                                location,
                                "One-sided assignment; "
                                "data flow not confirmed",
                                line,
                            )

                    # Dynamic worksheet selector
                    if DYNAMIC_OBJ.search(line):
                        self.unresolved(
                            "VBA",
                            module,
                            location,
                            "Dynamic sheet selector",
                            line,
                        )

        finally:
            parser.close()

    # --------------------------------------------------------
    # Excel report formatting
    # --------------------------------------------------------

    @staticmethod
    def style_header(ws):
        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78",
            )

            cell.font = Font(
                color="FFFFFF",
                bold=True,
            )

            cell.alignment = Alignment(
                horizontal="center",
            )

    # --------------------------------------------------------
    # Generate report
    # --------------------------------------------------------

    def report(self, output):
        print("[3/3] Generating report...")

        wb = Workbook()

        # Direct Matrix
        ws = wb.active
        ws.title = "Direct_Matrix"

        safe_text(
            ws.cell(1, 1),
            "Source / Target",
        )

        for i, name in enumerate(self.names, 2):
            safe_text(ws.cell(1, i), name)
            safe_text(ws.cell(i, 1), name)

        for i, source in enumerate(self.names, 2):
            for j, target in enumerate(
                self.names, 2
            ):
                value = int(
                    target in self.graph[source]
                )

                cell = ws.cell(i, j, value)

                cell.alignment = Alignment(
                    horizontal="center",
                )

                if value:
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="C6EFCE",
                    )

                    cell.font = Font(
                        color="006100",
                        bold=True,
                    )

        self.style_header(ws)

        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 32

        for col in range(2, len(self.names) + 2):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 22

        # Details and unresolved references
        report_sheets = [
            (
                "Dependency_Details",
                [
                    "Source Sheet",
                    "Target Sheet",
                    "Type",
                    "Module / Sheet",
                    "Location",
                    "Expression",
                ],
                self.details,
                [25, 25, 24, 30, 22, 95],
            ),
            (
                "Unresolved_References",
                [
                    "Type",
                    "Module / Sheet",
                    "Location",
                    "Reason",
                    "Expression",
                ],
                self.unknown,
                [22, 30, 22, 50, 95],
            ),
        ]

        for title, headers, records, widths in report_sheets:
            sheet = wb.create_sheet(title)

            for j, header in enumerate(headers, 1):
                safe_text(
                    sheet.cell(1, j),
                    header,
                )

            for i, record in enumerate(
                sorted(records), 2
            ):
                for j, value in enumerate(
                    record, 1
                ):
                    # Critical fix:
                    # Force all source expressions to text.
                    safe_text(
                        sheet.cell(i, j),
                        value,
                    )

            self.style_header(sheet)

            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

            for j, width in enumerate(widths, 1):
                sheet.column_dimensions[
                    get_column_letter(j)
                ].width = width

        wb.save(output)

    # --------------------------------------------------------
    # Main execution
    # --------------------------------------------------------

    def run(self, output):
        try:
            self.formulas()
            self.vba()
            self.report(output)

            print()
            print("========== ANALYSIS RESULT ==========")

            print(
                "Formula cells:",
                self.stats["formula_cells"],
            )

            print(
                "Cross-sheet formula cells:",
                self.stats["formula_cross_sheet_cells"],
            )

            print(
                "OFFSET formulas:",
                self.stats["offset_cells"],
            )

            print(
                "INDIRECT formulas:",
                self.stats["indirect_cells"],
            )

            print(
                "INDIRECT literal resolved:",
                self.stats["indirect_literal_resolved"],
            )

            print(
                "VBA modules:",
                self.stats["vba_modules"],
            )

            print(
                "VBA logical lines:",
                self.stats["vba_lines"],
            )

            print(
                "Confirmed direct sheet links:",
                sum(
                    len(v)
                    for v in self.graph.values()
                ),
            )

            print(
                "Evidence rows:",
                len(self.details),
            )

            print(
                "Unresolved rows:",
                len(self.unknown),
            )

            print("Report:", output)
            print("=====================================")

            if not self.details:
                print(
                    "WARNING: No confirmed dependencies. "
                    "Check unresolved references."
                )

        finally:
            self.book.close()


# ============================================================
# CLI
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
        help="Source .xlsx or .xlsm file",
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output .xlsx report",
    )

    args = parser.parse_args()

    source = Path(args.input)

    if (
        not source.is_file()
        or source.suffix.lower()
        not in (".xlsx", ".xlsm")
    ):
        parser.error(
            "Input must be an existing "
            ".xlsx or .xlsm file"
        )

    output = (
        Path(args.output)
        if args.output
        else source.with_name(
            source.stem
            + "_dependency_report.xlsx"
        )
    )

    if (
        output.suffix.lower() != ".xlsx"
        or output.resolve() == source.resolve()
    ):
        parser.error(
            "Output must be a different .xlsx file"
        )

    Analyzer(source).run(output)


if __name__ == "__main__":
    main()
