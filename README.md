
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Sheet Dependency Analyzer

功能：
1. 包含 visible / hidden / veryHidden Sheet
2. 普通 Excel 公式直接引用
3. INDIRECT 动态 Sheet 引用
4. INDEX + INDIRECT 嵌套公式
5. OFFSET 显式 Sheet 引用
6. 从单元格获取动态 Sheet 名称
7. 简单公式计算与缓存值读取
8. VBA Sheets / Worksheets 引用
9. VBA Activate / Select / Goto
10. VBA 图形按钮绑定宏
11. DrawingML / VML 控件
12. Excel 内部超链接
13. 直接引用矩阵
14. 未解析引用报告
15. Expression 强制文本保存

矩阵方向：
行 = 引用来源 Sheet
列 = 被引用 Sheet

注意：
- 静态分析，不执行宏
- 不修改原始 XLSM
- 无法确定的动态关系单独记录
"""

import argparse
import posixpath
import re
import zipfile

from collections import defaultdict
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from oletools.olevba import VBA_Parser


# ============================================================
# XML Namespaces
# ============================================================

REL = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

RID = f"{{{REL}}}id"


# ============================================================
# Helpers
# ============================================================

def lname(tag):
    return tag.rsplit("}", 1)[-1]


def save_text(cell, value):
    """Always store expression as text."""
    cell.value = str(value)[:32000]
    cell.data_type = "s"
    cell.number_format = "@"


def clean_macro(value):
    if not value:
        return ""

    value = str(value).strip().strip("'")

    if "!" in value:
        value = value.rsplit("!", 1)[-1]

    if "." in value:
        value = value.rsplit(".", 1)[-1]

    return value.strip().casefold()


def resolve_part(base, target):
    target = unquote(target or "")

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(base),
            target
        )
    )


def read_rels(zf, part):
    folder = posixpath.dirname(part)
    filename = posixpath.basename(part)

    relpath = posixpath.join(
        folder,
        "_rels",
        filename + ".rels"
    )

    if relpath not in zf.namelist():
        return {}

    root = ET.fromstring(zf.read(relpath))
    result = {}

    for item in root:
        if lname(item.tag) != "Relationship":
            continue

        result[item.get("Id")] = {
            "target": item.get("Target", ""),
            "type": item.get("Type", ""),
            "mode": item.get("TargetMode", ""),
            "location": item.get("Location", "")
        }

    return result


def strip_comment(line):
    quoted = False
    result = []
    i = 0

    while i < len(line):
        ch = line[i]

        if ch == '"':
            if (
                quoted
                and i + 1 < len(line)
                and line[i + 1] == '"'
            ):
                result.extend(['"', '"'])
                i += 2
                continue

            quoted = not quoted

        if ch == "'" and not quoted:
            break

        result.append(ch)
        i += 1

    return "".join(result)


def vba_lines(code):
    pending = ""
    start = 0

    for no, raw in enumerate(code.splitlines(), 1):
        line = strip_comment(raw).strip()

        if not line:
            continue

        if not pending:
            start = no

        if re.search(r"\s+_$", line):
            pending += re.sub(
                r"\s+_$", " ", line
            )
        else:
            yield start, pending + line
            pending = ""

    if pending:
        yield start, pending


# ============================================================
# Formula Parsing Helpers
# ============================================================

CELL_REF = re.compile(
    r"^\$?([A-Z]{1,3})\$?(\d+)$",
    re.I
)

SHEET_CELL = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!]+))!"
    r"(\$?[A-Z]{1,3}\$?\d+)$",
    re.I
)

VBA_SHEET = re.compile(
    r'\b(?:ThisWorkbook\s*\.\s*|'
    r'ActiveWorkbook\s*\.\s*)?'
    r'(Sheets|Worksheets)\s*'
    r'\(\s*(?:"((?:[^"]|"")+)"|(\d+))\s*\)',
    re.I
)

VBA_PROC = re.compile(
    r"^(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function)\s+([^\s(]+)",
    re.I
)

VBA_SET = re.compile(
    r"^Set\s+(\w+)\s*=\s*(.+)$",
    re.I
)

VBA_NAV = re.compile(
    r"\.\s*(Activate|Select)\b",
    re.I
)


def split_top(text, separator):
    """
    Split expression at top level.
    Respect strings and parentheses.
    """
    result = []
    depth = 0
    quoted = False
    start = 0
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == '"':
            if (
                quoted
                and i + 1 < len(text)
                and text[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted

        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1

            elif ch == separator and depth == 0:
                result.append(text[start:i].strip())
                start = i + 1

        i += 1

    result.append(text[start:].strip())
    return result


def extract_function_args(formula, function_name):
    """
    Extract balanced function arguments.

    Example:
        INDEX(INDIRECT(...), 1, 2)

    Returns argument strings of matching functions.
    """
    pattern = re.compile(
        r"(?<![A-Z0-9_.])"
        + re.escape(function_name)
        + r"\s*\(",
        re.I
    )

    results = []

    for match in pattern.finditer(formula):
        start = match.end()
        depth = 1
        quoted = False
        i = start

        while i < len(formula):
            ch = formula[i]

            if ch == '"':
                if (
                    quoted
                    and i + 1 < len(formula)
                    and formula[i + 1] == '"'
                ):
                    i += 2
                    continue

                quoted = not quoted

            elif not quoted:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1

                    if depth == 0:
                        results.append(
                            split_top(
                                formula[start:i], ","
                            )
                        )
                        break

            i += 1

    return results


# ============================================================
# Main Analyzer
# ============================================================

class Analyzer:

    def __init__(self, filename):
        self.path = Path(filename)

        self.wb = load_workbook(
            self.path,
            data_only=False,
            read_only=False,
            keep_links=True
        )

        self.cached = load_workbook(
            self.path,
            data_only=True,
            read_only=True
        )

        # IMPORTANT:
        # Include ALL sheets, including hidden sheets.
        self.names = list(self.wb.sheetnames)

        self.lookup = {
            name.casefold(): name
            for name in self.names
        }

        self.worksheets = [
            ws.title for ws in self.wb.worksheets
        ]

        self.sheet_states = {
            ws.title: ws.sheet_state
            for ws in self.wb.worksheets
        }

        self.code_names = {}

        for ws in self.wb.worksheets:
            code = ws.sheet_properties.codeName

            if code:
                self.code_names[
                    code.casefold()
                ] = ws.title

        self.data = {
            name: set() for name in self.names
        }

        self.navigation = {
            name: set() for name in self.names
        }

        self.details = set()
        self.unresolved_rows = set()

        self.procedures = defaultdict(list)
        self.macro_targets = defaultdict(set)

        self.stats = defaultdict(int)

    # --------------------------------------------------------
    # Common
    # --------------------------------------------------------

    def normalize(self, name):
        if not isinstance(name, str):
            return None

        return self.lookup.get(
            name.casefold()
        )

    def add(
        self,
        source,
        target,
        kind,
        relation,
        location,
        expression
    ):
        source = self.normalize(source)
        target = self.normalize(target)

        if not source or not target:
            return

        if source == target:
            return

        graph = (
            self.navigation
            if kind == "Navigation"
            else self.data
        )

        graph[source].add(target)

        self.details.add((
            source,
            target,
            kind,
            relation,
            str(location),
            str(expression)
        ))

    def unresolved(
        self,
        source,
        location,
        reason,
        expression
    ):
        self.unresolved_rows.add((
            str(source),
            str(location),
            str(reason),
            str(expression)
        ))

    # --------------------------------------------------------
    # Resolve Cell Values
    # --------------------------------------------------------

    def cell_value(self, sheet, address, depth=0):
        """
        Resolve a cell value using:
        1. Literal value
        2. Excel cached result
        3. Simple string formula

        Does not execute Excel formulas.
        """
        if depth > 8:
            return None

        sheet = self.normalize(sheet)

        if not sheet:
            return None

        try:
            cell = self.wb[sheet][address]
            value = cell.value

            if cell.data_type != "f":
                return value

            cached_value = self.cached[sheet][address].value

            if cached_value is not None:
                return cached_value

            formula = str(value).lstrip("=")

            return self.eval_string(
                sheet,
                formula,
                depth + 1
            )

        except Exception:
            return None

    def eval_string(self, sheet, expression, depth=0):
        """
        Evaluate simple string expressions.

        Supported:
            "サマリ"
            $E36
            '設定'!A1
            "'" & $E36 & "'!A1"
            CONCATENATE(...)
            CONCAT(...)
            UPPER(...)
            LOWER(...)
            TRIM(...)
            IFERROR(expr, fallback)

        Complex formulas return None.
        """
        if depth > 8:
            return None

        expr = expression.strip()

        if expr.startswith("="):
            expr = expr[1:].strip()

        # Remove balanced outer parentheses
        while (
            expr.startswith("(")
            and expr.endswith(")")
        ):
            inner = expr[1:-1].strip()
            if not inner:
                break
            expr = inner

        # Concatenation
        parts = split_top(expr, "&")

        if len(parts) > 1:
            values = []

            for part in parts:
                value = self.eval_string(
                    sheet, part, depth + 1
                )

                if value is None:
                    return None

                values.append(str(value))

            return "".join(values)

        # Literal
        if (
            expr.startswith('"')
            and expr.endswith('"')
        ):
            return expr[1:-1].replace('""', '"')

        # Local cell
        if CELL_REF.fullmatch(expr):
            return self.cell_value(
                sheet, expr, depth + 1
            )

        # Cross-sheet cell
        match = SHEET_CELL.fullmatch(expr)

        if match:
            target = (
                match.group(1)
                or match.group(2)
            ).replace("''", "'")

            return self.cell_value(
                target,
                match.group(3),
                depth + 1
            )

        # Functions
        match = re.fullmatch(
            r"(CONCATENATE|CONCAT|UPPER|LOWER|"
            r"TRIM|IFERROR)\s*\((.*)\)",
            expr,
            re.I | re.S
        )

        if match:
            fn = match.group(1).upper()
            args = split_top(
                match.group(2), ","
            )

            if fn == "IFERROR":
                if not args:
                    return None

                value = self.eval_string(
                    sheet, args[0], depth + 1
                )

                if value is None and len(args) > 1:
                    return self.eval_string(
                        sheet, args[1], depth + 1
                    )

                return value

            values = []

            for arg in args:
                value = self.eval_string(
                    sheet, arg, depth + 1
                )

                if value is None:
                    return None

                values.append(str(value))

            if fn in ("CONCATENATE", "CONCAT"):
                return "".join(values)

            if len(values) != 1:
                return None

            if fn == "UPPER":
                return values[0].upper()

            if fn == "LOWER":
                return values[0].lower()

            if fn == "TRIM":
                return " ".join(values[0].split())

        return None

    # --------------------------------------------------------
    # Parse Sheet From Reference Text
    # --------------------------------------------------------

    def sheet_from_reference(self, value):
        if value is None:
            return None

        value = unquote(str(value)).strip()

        if value.startswith("#"):
            value = value[1:]

        if "!" not in value:
            return None

        name = value.split("!", 1)[0].strip()

        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")

        # Do not mistake external workbook refs
        # for internal sheet names.
        if "[" in name or "]" in name:
            return None

        return self.normalize(name)

    # --------------------------------------------------------
    # Ordinary Formula References
    # --------------------------------------------------------

    def formula_refs(self, formula):
        """
        Use Excel tokenizer to find explicit sheet
        references, including quoted Japanese names.
        """
        refs = set()

        try:
            tokens = Tokenizer(formula).items
        except Exception:
            self.stats["tokenizer_errors"] += 1
            return refs

        for token in tokens:
            if not (
                token.type == "OPERAND"
                and token.subtype == "RANGE"
            ):
                continue

            value = token.value

            if "!" not in value:
                continue

            sheet_part = value.rsplit("!", 1)[0]

            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                sheet_part = sheet_part[1:-1].replace(
                    "''", "'"
                )

            if "[" in sheet_part or "]" in sheet_part:
                continue

            # 3D reference: Sheet1:Sheet3!A1
            if ":" in sheet_part:
                first, last = sheet_part.split(":", 1)

                first = self.normalize(first)
                last = self.normalize(last)

                if first and last:
                    a = self.names.index(first)
                    b = self.names.index(last)

                    refs.update(
                        self.names[
                            min(a, b):max(a, b) + 1
                        ]
                    )

                continue

            target = self.normalize(sheet_part)

            if target:
                refs.add(target)

        return refs

    # --------------------------------------------------------
    # INDIRECT
    # --------------------------------------------------------

    def analyze_indirect(
        self,
        sheet,
        address,
        formula
    ):
        """
        Example:

        =INDEX(
            INDIRECT("'" & $E36 & "'!$H$119:$W$123"),
            ...
        )

        If E36 = マスタ:
            source -> マスタ
        """
        calls = extract_function_args(
            formula, "INDIRECT"
        )

        for args in calls:
            if not args:
                continue

            self.stats["indirect_calls"] += 1

            expression = args[0]

            resolved = self.eval_string(
                sheet, expression
            )

            target = self.sheet_from_reference(
                resolved
            )

            if target:
                self.add(
                    sheet,
                    target,
                    "Data",
                    "INDIRECT Resolved",
                    f"{sheet}!{address}",
                    f"{expression} => {resolved}"
                )

                self.stats[
                    "indirect_resolved"
                ] += 1

            else:
                self.unresolved(
                    sheet,
                    address,
                    "INDIRECT target unresolved",
                    f"{expression} | Result={resolved!r}"
                )

    # --------------------------------------------------------
    # Formula Analysis
    # --------------------------------------------------------

    def analyze_formulas(self):
        print("[1/5] Analyzing Excel formulas")

        for sheet in self.worksheets:
            ws = self.wb[sheet]

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["formula_cells"] += 1

                    formula = str(cell.value)

                    for target in self.formula_refs(formula):
                        self.add(
                            sheet,
                            target,
                            "Data",
                            "Explicit Formula",
                            f"{sheet}!{cell.coordinate}",
                            formula
                        )

                    if re.search(
                        r"\bINDIRECT\s*\(",
                        formula,
                        re.I
                    ):
                        self.analyze_indirect(
                            sheet,
                            cell.coordinate,
                            formula
                        )

                    if re.search(
                        r"\bOFFSET\s*\(",
                        formula,
                        re.I
                    ):
                        self.stats["offset_calls"] += 1

                    # HYPERLINK with static internal target
                    for args in extract_function_args(
                        formula, "HYPERLINK"
                    ):
                        if not args:
                            continue

                        value = self.eval_string(
                            sheet, args[0]
                        )

                        target = self.sheet_from_reference(
                            value
                        )

                        if target:
                            self.add(
                                sheet,
                                target,
                                "Navigation",
                                "HYPERLINK Formula",
                                f"{sheet}!{cell.coordinate}",
                                formula
                            )

    # --------------------------------------------------------
    # VBA References
    # --------------------------------------------------------

    def vba_refs(self, expression, aliases=None):
        aliases = aliases or {}
        refs = set()

        for match in VBA_SHEET.finditer(expression):
            kind, name, number = match.groups()

            if name:
                target = self.normalize(
                    name.replace('""', '"')
                )

            else:
                index = int(number)

                sequence = (
                    self.worksheets
                    if kind.casefold() == "worksheets"
                    else self.names
                )

                target = (
                    sequence[index - 1]
                    if 1 <= index <= len(sequence)
                    else None
                )

            if target:
                refs.add(target)

        for code, target in self.code_names.items():
            if re.search(
                r"(?<![\w.])"
                + re.escape(code)
                + r"\s*[.!]",
                expression,
                re.I
            ):
                refs.add(target)

        for alias, target in aliases.items():
            if re.search(
                r"(?<![\w.])"
                + re.escape(alias)
                + r"\s*[.!]",
                expression,
                re.I
            ):
                refs.add(target)

        return refs

    def extract_vba(self):
        print("[2/5] Extracting VBA")

        parser = VBA_Parser(str(self.path))

        try:
            if not parser.detect_vba_macros():
                return

            for _, _, module, code in parser.extract_macros():
                self.stats["vba_modules"] += 1
                current = None

                for no, line in vba_lines(code):
                    match = VBA_PROC.match(line)

                    if match:
                        current = clean_macro(
                            match.group(1)
                        )
                        continue

                    if re.match(
                        r"^End\s+(Sub|Function)\b",
                        line,
                        re.I
                    ):
                        current = None
                        continue

                    if current:
                        self.procedures[current].append(
                            (module, no, line)
                        )

        finally:
            parser.close()

    def analyze_vba(self):
        print("[3/5] Analyzing VBA")

        calls = defaultdict(set)
        direct_nav = defaultdict(set)

        for proc, lines in self.procedures.items():
            aliases = {}

            for module, no, line in lines:
                match = VBA_SET.match(line)

                if match:
                    variable, expression = match.groups()

                    refs = self.vba_refs(
                        expression, aliases
                    )

                    if len(refs) == 1:
                        aliases[
                            variable.casefold()
                        ] = next(iter(refs))

                    continue

                nav = VBA_NAV.search(line)

                if nav:
                    refs = self.vba_refs(
                        line[:nav.start()],
                        aliases
                    )

                    direct_nav[proc].update(refs)

                    if not refs:
                        self.unresolved(
                            module,
                            no,
                            "VBA navigation unresolved",
                            line
                        )

                    continue

                if re.search(
                    r"\b(?:Application\s*\.\s*)?Goto\b",
                    line,
                    re.I
                ):
                    refs = self.vba_refs(
                        line, aliases
                    )

                    direct_nav[proc].update(refs)

                # Basic assignment: target = source
                if (
                    "=" in line
                    and not re.match(
                        r"^(If|ElseIf|For|Do|While|"
                        r"Select|Case|Const|Dim)\b",
                        line,
                        re.I
                    )
                ):
                    left, right = line.split("=", 1)

                    targets = self.vba_refs(
                        left, aliases
                    )

                    sources = self.vba_refs(
                        right, aliases
                    )

                    for target in targets:
                        for source in sources:
                            self.add(
                                target,
                                source,
                                "Data",
                                "VBA Assignment",
                                f"{module}:{no}",
                                line
                            )

                # Simple procedure calls
                call = re.match(
                    r"^(?:Call\s+)?([^\s(]+)",
                    line,
                    re.I
                )

                if call:
                    called = clean_macro(
                        call.group(1)
                    )

                    if (
                        called != proc
                        and called in self.procedures
                    ):
                        calls[proc].add(called)

        # Resolve VBA procedure calls.
        # This is not transitive sheet dependency.
        resolved = {
            proc: set(direct_nav[proc])
            for proc in self.procedures
        }

        changed = True

        while changed:
            changed = False

            for proc in self.procedures:
                before = len(resolved[proc])

                for called in calls[proc]:
                    resolved[proc].update(
                        resolved.get(called, set())
                    )

                if len(resolved[proc]) != before:
                    changed = True

        self.macro_targets = resolved

    # --------------------------------------------------------
    # XML Sheet Mapping
    # --------------------------------------------------------

    def sheet_parts(self, zf):
        workbook = "xl/workbook.xml"

        root = ET.fromstring(
            zf.read(workbook)
        )

        rels = read_rels(zf, workbook)
        result = {}

        for elem in root.iter():
            if lname(elem.tag) != "sheet":
                continue

            name = elem.get("name")
            rid = elem.get(RID)

            if rid not in rels:
                continue

            result[name] = resolve_part(
                workbook,
                rels[rid]["target"]
            )

        return result

    # --------------------------------------------------------
    # Shape Macro
    # --------------------------------------------------------

    def shape_macro(
        self,
        source,
        macro,
        location
    ):
        self.stats["shape_macro_bindings"] += 1

        proc = clean_macro(macro)

        targets = self.macro_targets.get(
            proc, set()
        )

        if not targets:
            self.unresolved(
                source,
                location,
                "Shape macro target unresolved",
                macro
            )

            return

        for target in targets:
            self.add(
                source,
                target,
                "Navigation",
                "Shape VBA Macro",
                location,
                macro
            )

    # --------------------------------------------------------
    # DrawingML
    # --------------------------------------------------------

    def analyze_drawing(
        self,
        zf,
        source,
        part
    ):
        if part not in zf.namelist():
            return

        root = ET.fromstring(
            zf.read(part)
        )

        rels = read_rels(zf, part)

        for elem in root.iter():
            tag = lname(elem.tag)

            if tag in (
                "sp",
                "pic",
                "graphicFrame"
            ):
                macro = elem.get("macro")

                if macro:
                    shape = tag

                    for child in elem.iter():
                        if lname(child.tag) == "cNvPr":
                            shape = child.get(
                                "name", tag
                            )
                            break

                    self.shape_macro(
                        source,
                        macro,
                        f"{part}:{shape}"
                    )

            if tag == "hlinkClick":
                rid = elem.get(RID)

                rel = rels.get(
                    rid, {}
                )

                location = (
                    rel.get("location")
                    or rel.get("target", "")
                )

                target = self.sheet_from_reference(
                    location
                )

                if target:
                    self.add(
                        source,
                        target,
                        "Navigation",
                        "Shape Hyperlink",
                        part,
                        location
                    )

    # --------------------------------------------------------
    # VML Controls
    # --------------------------------------------------------

    def analyze_vml(
        self,
        zf,
        source,
        part
    ):
        if part not in zf.namelist():
            return

        root = ET.fromstring(
            zf.read(part)
        )

        for elem in root.iter():
            if lname(elem.tag) != "ClientData":
                continue

            for child in elem:
                if lname(child.tag) == "FmlaMacro":
                    macro = (
                        child.text or ""
                    ).strip()

                    if macro:
                        self.shape_macro(
                            source,
                            macro,
                            part
                        )

    # --------------------------------------------------------
    # Cell Hyperlinks
    # --------------------------------------------------------

    def analyze_cell_links(
        self,
        zf,
        source,
        part,
        rels
    ):
        root = ET.fromstring(
            zf.read(part)
        )

        for elem in root.iter():
            if lname(elem.tag) != "hyperlink":
                continue

            location = elem.get(
                "location", ""
            )

            rid = elem.get(RID)

            if not location and rid:
                location = rels.get(
                    rid, {}
                ).get("target", "")

            target = self.sheet_from_reference(
                location
            )

            if target:
                self.add(
                    source,
                    target,
                    "Navigation",
                    "Cell Hyperlink",
                    f"{part}:{elem.get('ref', '')}",
                    location
                )

    # --------------------------------------------------------
    # Shape Analysis
    # --------------------------------------------------------

    def analyze_shapes(self):
        print("[4/5] Analyzing shapes")

        with zipfile.ZipFile(
            self.path
        ) as zf:

            mapping = self.sheet_parts(zf)

            for source, sheet_part in mapping.items():
                # Include hidden sheets too.
                if sheet_part not in zf.namelist():
                    continue

                rels = read_rels(
                    zf, sheet_part
                )

                self.analyze_cell_links(
                    zf,
                    source,
                    sheet_part,
                    rels
                )

                for rel in rels.values():
                    if rel["mode"] == "External":
                        continue

                    target_part = resolve_part(
                        sheet_part,
                        rel["target"]
                    )

                    relation_type = rel["type"]

                    if relation_type.endswith("/drawing"):
                        self.analyze_drawing(
                            zf,
                            source,
                            target_part
                        )

                    elif relation_type.endswith(
                        "/vmlDrawing"
                    ):
                        self.analyze_vml(
                            zf,
                            source,
                            target_part
                        )

    # --------------------------------------------------------
    # Matrix
    # --------------------------------------------------------

    def create_matrix(
        self,
        wb,
        name,
        graph
    ):
        ws = wb.create_sheet(name)

        save_text(
            ws.cell(1, 1),
            "Source / Target"
        )

        for i, sheet in enumerate(
            self.names, 2
        ):
            save_text(
                ws.cell(1, i),
                sheet
            )

            save_text(
                ws.cell(i, 1),
                sheet
            )

        for i, source in enumerate(
            self.names, 2
        ):
            for j, target in enumerate(
                self.names, 2
            ):
                value = int(
                    target in graph[source]
                )

                cell = ws.cell(
                    i, j, value
                )

                cell.alignment = Alignment(
                    horizontal="center"
                )

                if value:
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="C6EFCE"
                    )

                    cell.font = Font(
                        bold=True,
                        color="006100"
                    )

        self.header(ws)

        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 35

        for col in range(
            2, len(self.names) + 2
        ):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 22

    @staticmethod
    def header(ws):
        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78"
            )

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

    # --------------------------------------------------------
    # Report
    # --------------------------------------------------------

    def report(self, output):
        print("[5/5] Generating report")

        wb = Workbook()
        wb.remove(wb.active)

        combined = {
            name: (
                self.data[name]
                | self.navigation[name]
            )
            for name in self.names
        }

        self.create_matrix(
            wb,
            "Direct_Matrix",
            combined
        )

        self.create_matrix(
            wb,
            "Data_Matrix",
            self.data
        )

        self.create_matrix(
            wb,
            "Navigation_Matrix",
            self.navigation
        )

        # Sheet visibility information
        ws = wb.create_sheet(
            "Sheet_Information"
        )

        ws.append([
            "Sheet Name",
            "Visibility"
        ])

        for name in self.names:
            ws.append([
                name,
                self.sheet_states.get(
                    name, "unknown"
                )
            ])

        self.header(ws)
        ws.column_dimensions["A"].width = 45
        ws.column_dimensions["B"].width = 20

        # Dependency details
        ws = wb.create_sheet(
            "Dependency_Details"
        )

        headers = [
            "Source Sheet",
            "Target Sheet",
            "Kind",
            "Relation",
            "Location",
            "Expression"
        ]

        for col, value in enumerate(headers, 1):
            save_text(
                ws.cell(1, col),
                value
            )

        for row, record in enumerate(
            sorted(self.details), 2
        ):
            for col, value in enumerate(
                record, 1
            ):
                save_text(
                    ws.cell(row, col),
                    value
                )

        self.header(ws)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col, width in enumerate(
            [28, 28, 20, 30, 40, 110],
            1
        ):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = width

        # Unresolved
        ws = wb.create_sheet(
            "Unresolved_References"
        )

        headers = [
            "Source",
            "Location",
            "Reason",
            "Expression"
        ]

        for col, value in enumerate(headers, 1):
            save_text(
                ws.cell(1, col),
                value
            )

        for row, record in enumerate(
            sorted(self.unresolved_rows), 2
        ):
            for col, value in enumerate(
                record, 1
            ):
                save_text(
                    ws.cell(row, col),
                    value
                )

        self.header(ws)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col, width in enumerate(
            [30, 35, 45, 110],
            1
        ):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = width

        wb.save(output)

    # --------------------------------------------------------
    # Run
    # --------------------------------------------------------

    def run(self, output):
        try:
            self.analyze_formulas()

            self.extract_vba()
            self.analyze_vba()

            self.analyze_shapes()

            self.report(output)

            data_count = sum(
                len(v)
                for v in self.data.values()
            )

            nav_count = sum(
                len(v)
                for v in self.navigation.values()
            )

            total_count = sum(
                len(
                    self.data[name]
                    | self.navigation[name]
                )
                for name in self.names
            )

            print()
            print("========== RESULT ==========")

            print(
                "Total Sheets:",
                len(self.names)
            )

            print(
                "Hidden Sheets:",
                sum(
                    state != "visible"
                    for state in self.sheet_states.values()
                )
            )

            print(
                "Formula Cells:",
                self.stats["formula_cells"]
            )

            print(
                "INDIRECT Calls:",
                self.stats["indirect_calls"]
            )

            print(
                "INDIRECT Resolved:",
                self.stats["indirect_resolved"]
            )

            print(
                "OFFSET Calls:",
                self.stats["offset_calls"]
            )

            print(
                "VBA Modules:",
                self.stats["vba_modules"]
            )

            print(
                "VBA Procedures:",
                len(self.procedures)
            )

            print(
                "Shape Macro Bindings:",
                self.stats["shape_macro_bindings"]
            )

            print("Data Links:", data_count)
            print("Navigation Links:", nav_count)
            print("Combined Links:", total_count)

            print(
                "Evidence Rows:",
                len(self.details)
            )

            print(
                "Unresolved Rows:",
                len(self.unresolved_rows)
            )

            print("Report:", output)
            print("============================")

        finally:
            self.wb.close()
            self.cached.close()


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Analyze XLSM direct sheet references "
            "including INDIRECT, VBA and shapes"
        )
    )

    parser.add_argument(
        "input",
        help="Input XLSM file"
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output XLSX report"
    )

    args = parser.parse_args()

    source = Path(args.input)

    if not source.is_file():
        parser.error(
            f"File not found: {source}"
        )

    if source.suffix.lower() not in (
        ".xlsm", ".xlsx"
    ):
        parser.error(
            "Only XLSM/XLSX supported"
        )

    output = (
        Path(args.output)
        if args.output
        else source.with_name(
            source.stem
            + "_dependency_report.xlsx"
        )
    )

    if output.suffix.lower() != ".xlsx":
        parser.error(
            "Output must be .xlsx"
        )

    if source.resolve() == output.resolve():
        parser.error(
            "Output cannot overwrite input"
        )

    analyzer = Analyzer(source)
    analyzer.run(output)


if __name__ == "__main__":
    main()
