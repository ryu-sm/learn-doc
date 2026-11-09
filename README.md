
# -*- coding: utf-8 -*-
"""
Excel XLSM Sheet Dependency Analyzer
====================================

支持：
1. 所有 Sheet（包括 Hidden / VeryHidden）
2. 普通公式跨 Sheet 直接引用
3. INDIRECT 动态引用：
   - INDIRECT("'Sheet1'!A1")
   - INDIRECT("'" & E36 & "'!A1")
   - INDIRECT("サマリ!$G" & ROW() & ...)
   - INDEX(INDIRECT(...), ...)
4. OFFSET、SUMIFS 等公式中的显式 Sheet 引用
5. 单元格值追踪（包括字符串公式的简单计算）
6. VBA 宏中的 Sheet 引用及 Activate / Select 导航
7. 图形按钮的宏绑定（VML / Drawing）
8. 单独生成数据引用矩阵和导航关系
9. 无法解析的动态引用单独记录
10. 不修改原始 XLSM

使用：
    python analyze_xlsm.py input.xlsm

可选：
    python analyze_xlsm.py input.xlsm --output result.xlsx

依赖：
    pip install openpyxl oletools

说明：
    本程序为静态分析器，不执行宏，不运行原文件。
    复杂运行时 VBA、外部连接和动态计算无法保证全部解析。
"""

import argparse
import re
import sys
import zipfile
import posixpath

from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
)
from openpyxl.utils import get_column_letter


# ============================================================
# 配置
# ============================================================

MAX_EVIDENCE = 150000
MAX_UNRESOLVED = 50000

# 是否统计 Sheet 对自身的引用
INCLUDE_SELF = True

# 是否扫描 VBA
SCAN_VBA = True

# 是否扫描图形按钮
SCAN_SHAPES = True

# 对存在未解析动态引用的源 Sheet：
# 未确认的矩阵单元格使用 ?，而不是错误地写成 0
MARK_UNKNOWN = True


# ============================================================
# 基础工具
# ============================================================

def norm(value):
    return str(value or "").strip().casefold()


def safe_text(value):
    if value is None:
        return ""
    return str(value)


def is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def sheet_lookup(names):
    return {norm(name): name for name in names}


def actual_sheet(value, lookup):
    return lookup.get(norm(value))


def unquote_excel(value):
    value = safe_text(value).strip()

    if value.startswith("'") and value.endswith("'"):
        value = value[1:-1].replace("''", "'")

    if value.startswith("[") and "]" in value:
        value = value.split("]", 1)[1]

    return value


def extract_target(text, lookup):
    """
    从形如以下的文本中提取目标 Sheet：
      Sheet1!A1
      'Sheet 1'!A1
      'Sheet1'!$H$119:$W$123
    """
    if not text:
        return None

    text = safe_text(text).strip()

    m = re.match(
        r"^\s*(?:'((?:[^']|'')+)'|([^!]+))!",
        text
    )

    if not m:
        return None

    name = m.group(1) if m.group(1) is not None else m.group(2)
    name = unquote_excel(name)

    return actual_sheet(name, lookup)


def split_args(text):
    """
    按最外层逗号分割 Excel 函数参数。
    忽略括号、字符串和数组内部的逗号。
    """
    result = []
    current = []
    depth = 0
    brace = 0
    in_string = False
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == '"':
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                current.append('""')
                i += 2
                continue

            in_string = not in_string
            current.append(ch)

        elif not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "{":
                brace += 1
            elif ch == "}":
                brace -= 1

            if ch in ",;" and depth == 0 and brace == 0:
                result.append("".join(current).strip())
                current = []
            else:
                current.append(ch)
        else:
            current.append(ch)

        i += 1

    result.append("".join(current).strip())
    return result


def split_concat(text):
    """
    分割 Excel & 字符串连接表达式。
    """
    result = []
    current = []
    depth = 0
    in_string = False
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == '"':
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                current.append('""')
                i += 2
                continue

            in_string = not in_string
            current.append(ch)

        elif not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1

            if ch == "&" and depth == 0:
                result.append("".join(current).strip())
                current = []
            else:
                current.append(ch)
        else:
            current.append(ch)

        i += 1

    result.append("".join(current).strip())
    return result


def function_calls(formula, name):
    """
    获取指定函数的完整参数字符串。
    支持嵌套括号及 Excel 字符串。
    """
    pattern = re.compile(
        r"(?<![A-Z0-9_.])(?:_xlfn\.)?"
        + re.escape(name)
        + r"\s*\(",
        re.I
    )

    results = []

    for match in pattern.finditer(formula):
        start = match.end()
        depth = 1
        in_string = False
        i = start

        while i < len(formula):
            ch = formula[i]

            if ch == '"':
                if in_string and i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 2
                    continue
                in_string = not in_string

            elif not in_string:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        results.append(formula[start:i])
                        break

            i += 1

    return results


def strip_strings(text):
    """
    删除双引号字符串内容，但保留字符串以外的公式。
    用于扫描真正的显式 Sheet!Cell 引用。
    """
    result = []
    in_string = False
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == '"':
            if in_string and i + 1 < len(text) and text[i + 1] == '"':
                result.extend([" ", " "])
                i += 2
                continue

            in_string = not in_string
            result.append(" ")

        else:
            result.append(" " if in_string else ch)

        i += 1

    return "".join(result)


def remove_outer_parentheses(text):
    text = text.strip()

    while text.startswith("(") and text.endswith(")"):
        depth = 0
        valid = True
        in_string = False

        for i, ch in enumerate(text):
            if ch == '"':
                in_string = not in_string
            elif not in_string:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0 and i != len(text) - 1:
                        valid = False
                        break

        if not valid or depth != 0:
            break

        text = text[1:-1].strip()

    return text


# ============================================================
# Excel 公式引用解析
# ============================================================

class DependencyAnalyzer:

    def __init__(self, source):
        self.source = Path(source)

        print("Loading workbook...", flush=True)

        # 普通模式：支持 sheet_state，避免 ReadOnlyWorksheet 错误
        self.wb = load_workbook(
            self.source,
            read_only=False,
            data_only=False,
            keep_vba=False,
            keep_links=True,
        )

        # 缓存值：用于追踪 E36 等动态 Sheet 名称
        self.cached_wb = load_workbook(
            self.source,
            read_only=True,
            data_only=True,
            keep_links=True,
        )

        self.names = self.wb.sheetnames
        self.lookup = sheet_lookup(self.names)

        self.links = defaultdict(set)
        self.evidence = []
        self.unresolved = []
        self.navigation = []
        self.stats = Counter()
        self.unknown_sources = set()

        self.truncated_evidence = 0
        self.truncated_unresolved = 0

        self.vba_procedures = {}
        self.macro_bindings = []

        self.sheet_code_names = {}

        for ws in self.wb.worksheets:
            code = getattr(
                getattr(ws, "sheet_properties", None),
                "codeName",
                None
            )
            if code:
                self.sheet_code_names[norm(code)] = ws.title

    def add_link(self, source, target, kind, location, expression):
        target = actual_sheet(target, self.lookup)

        if not target:
            return

        if source == target and not INCLUDE_SELF:
            return

        self.links[source].add(target)

        if len(self.evidence) < MAX_EVIDENCE:
            self.evidence.append((
                source,
                target,
                kind,
                location,
                safe_text(expression),
            ))
        else:
            self.truncated_evidence += 1

        self.stats["Evidence_" + kind] += 1

    def add_unresolved(self, sheet, location, category, reason, expression):
        self.unknown_sources.add(sheet)

        if len(self.unresolved) < MAX_UNRESOLVED:
            self.unresolved.append((
                sheet,
                location,
                category,
                reason,
                safe_text(expression),
            ))
        else:
            self.truncated_unresolved += 1

        self.stats["Unresolved_" + category] += 1

    def add_navigation(self, source, target, kind, location, expression):
        target = actual_sheet(target, self.lookup)

        if not target:
            return

        self.navigation.append((
            source,
            target,
            kind,
            location,
            safe_text(expression),
        ))

    def cached_value(self, sheet, address):
        try:
            return self.cached_wb[sheet][address].value
        except Exception:
            return None

    def cell_value(self, sheet, address, visited=None):
        """
        追踪 Sheet 名称所在的单元格。
        优先读取缓存值，再解析简单字符串公式。
        """
        visited = visited or set()
        key = (sheet, address.replace("$", "").upper())

        if key in visited or len(visited) > 12:
            return None

        visited.add(key)

        try:
            cell = self.wb[sheet][address.replace("$", "")]
            value = cell.value
        except Exception:
            return None

        if not is_formula(value):
            return value

        cached = self.cached_value(sheet, address.replace("$", ""))

        if cached is not None and not (
            isinstance(cached, str) and cached.startswith("#")
        ):
            return cached

        return self.eval_simple(
            value[1:],
            sheet,
            address,
            visited
        )

    def eval_simple(self, expression, sheet, cell_address, visited=None):
        """
        只计算确定性的简单表达式：
          "Sheet1"
          E36
          $E$36
          '設定'!A1
          "abc" & E36
          IF(条件,值1,值2) 的简单条件

        不执行 Excel 宏，不计算复杂工作簿。
        """
        visited = visited or set()

        text = remove_outer_parentheses(expression.strip())

        if text.startswith("="):
            text = text[1:].strip()

        parts = split_concat(text)

        if len(parts) > 1:
            values = []

            for part in parts:
                value = self.eval_simple(
                    part, sheet, cell_address, set(visited)
                )

                if value is None:
                    return None

                values.append(str(value))

            return "".join(values)

        if text.startswith('"') and text.endswith('"'):
            return text[1:-1].replace('""', '"')

        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
            try:
                return int(text)
            except ValueError:
                return float(text)

        m = re.fullmatch(
            r"(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?"
            r"(\$?[A-Z]{1,3}\$?\d+)",
            text,
            re.I
        )

        if m:
            target_sheet = sheet

            if m.group(1) or m.group(2):
                raw = m.group(1) or m.group(2)
                target_sheet = actual_sheet(
                    unquote_excel(raw),
                    self.lookup
                )

            if not target_sheet:
                return None

            return self.cell_value(
                target_sheet,
                m.group(3),
                set(visited)
            )

        # ROW() / COLUMN() 的简单计算
        if re.fullmatch(r"ROW\(\)", text, re.I):
            m = re.search(r"\d+", cell_address)
            return int(m.group()) if m else None

        if re.fullmatch(r"COLUMN\(\)", text, re.I):
            from openpyxl.utils.cell import coordinate_from_string
            from openpyxl.utils.column_index_from_string import (
                column_index_from_string,
            )

            try:
                col, _ = coordinate_from_string(cell_address)
                return column_index_from_string(col)
            except Exception:
                return None

        # IFERROR(value, fallback)：仅在 value 已可确定时返回
        if re.match(r"^IFERROR\s*\(", text, re.I):
            calls = function_calls(text, "IFERROR")
            if calls:
                args = split_args(calls[0])
                if args:
                    return self.eval_simple(
                        args[0], sheet, cell_address, set(visited)
                    )

        return None

    def explicit_references(self, formula):
        """
        从字符串以外的公式区域提取显式 Sheet 引用。
        使用实际 Sheet 名称匹配，降低误判。

        例如：
          =SUM(マスタ!A1:A10)
          =OFFSET('Sheet 2'!H2,B3-1,0)
        """
        clean = strip_strings(formula)
        found = set()

        # 按长度降序，避免名称前缀冲突
        for name in sorted(self.names, key=len, reverse=True):
            quoted = "'" + name.replace("'", "''") + "'"

            for representation in (quoted, name):
                pattern = (
                    r"(?<![A-Za-z0-9_\u0080-\uffff'])"
                    + re.escape(representation)
                    + r"\s*!"
                )

                if re.search(pattern, clean, re.I):
                    found.add(name)
                    break

        return found

    def resolve_indirect(self, argument, sheet, address):
        """
        解析 INDIRECT 第一个参数。

        返回：
            targets, method, reason

        支持：
          "'Sheet1'!A1"
          "'" & E36 & "'!A1"
          "サマリ!$G" & ROW() & ...
          "'" & E36 & "'!$H$119:$W$123"

        对于后续行列地址无法计算，但 Sheet 名称固定的情况，
        仍可确认目标 Sheet。
        """
        argument = argument.strip()

        # 1. 完整简单表达式求值
        evaluated = self.eval_simple(argument, sheet, address)

        if evaluated is not None:
            target = extract_target(evaluated, self.lookup)

            if target:
                return {target}, "EvaluatedText", ""

        # 2. 字符串拼接分段解析
        parts = split_concat(argument)

        # 仅需还原 ! 之前的 Sheet 名称
        prefix = ""
        unknown_before_bang = False

        for part in parts:
            part = remove_outer_parentheses(part)

            if "!" in prefix:
                break

            value = self.eval_simple(part, sheet, address)

            if value is None:
                # 如果前面的字符串已经包含 Sheet!
                # 则后面的地址计算失败不影响 Sheet 判定
                if "!" in prefix:
                    break

                unknown_before_bang = True
                break

            prefix += str(value)

            if "!" in prefix:
                break

        if "!" in prefix:
            target = extract_target(prefix, self.lookup)

            if target:
                return {target}, "PartialText", ""

        # 3. 特殊情况：
        # INDIRECT("サマリ!$G"&ROW()&...)
        # 第一段字符串已经明确目标 Sheet
        literals = re.findall(
            r'"((?:[^"]|"")*)"',
            argument
        )

        if literals:
            first = literals[0].replace('""', '"')

            if "!" in first:
                target = extract_target(first, self.lookup)

                if target:
                    return {target}, "LiteralPrefix", ""

        # 4. 如果表达式含有明确的 Sheet! 引用，
        # 这些属于参数计算的数据依赖
        # 但不一定是 INDIRECT 最终目标
        if unknown_before_bang:
            return set(), "", "Dynamic sheet selector unresolved"

        return set(), "", "INDIRECT text argument unresolved"

    def scan_formulas(self):
        print("Scanning formulas...", flush=True)

        for ws in self.wb.worksheets:
            print(
                f"  Sheet: {ws.title} [{ws.sheet_state}]",
                flush=True
            )

            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value

                    if not is_formula(value):
                        continue

                    self.stats["Formula Cells"] += 1
                    formula = str(value)
                    address = cell.coordinate

                    # A. 显式引用
                    targets = self.explicit_references(formula)

                    for target in targets:
                        self.add_link(
                            ws.title,
                            target,
                            "ExplicitFormula",
                            address,
                            formula
                        )

                    # B. INDIRECT
                    indirects = function_calls(formula, "INDIRECT")

                    for call in indirects:
                        self.stats["INDIRECT Calls"] += 1

                        args = split_args(call)

                        if not args:
                            self.add_unresolved(
                                ws.title,
                                address,
                                "INDIRECT",
                                "Missing argument",
                                formula
                            )
                            continue

                        resolved, method, reason = self.resolve_indirect(
                            args[0],
                            ws.title,
                            address
                        )

                        if resolved:
                            for target in resolved:
                                self.add_link(
                                    ws.title,
                                    target,
                                    "INDIRECT_" + method,
                                    address,
                                    formula
                                )
                        else:
                            self.add_unresolved(
                                ws.title,
                                address,
                                "INDIRECT",
                                reason,
                                formula
                            )

                    # C. OFFSET
                    offsets = function_calls(formula, "OFFSET")

                    for call in offsets:
                        self.stats["OFFSET Calls"] += 1

                        args = split_args(call)

                        if args:
                            offset_targets = self.explicit_references(
                                args[0]
                            )

                            for target in offset_targets:
                                if target not in targets:
                                    self.add_link(
                                        ws.title,
                                        target,
                                        "OFFSET",
                                        address,
                                        formula
                                    )

                            # OFFSET 的基准为本 Sheet 时，不新增跨 Sheet 关系
                            # 动态基准无法确定时单独记录
                            if not offset_targets:
                                base = args[0].strip()

                                if not re.fullmatch(
                                    r"\$?[A-Z]{1,3}\$?\d+"
                                    r"(?::\$?[A-Z]{1,3}\$?\d+)?",
                                    base,
                                    re.I
                                ):
                                    self.add_unresolved(
                                        ws.title,
                                        address,
                                        "OFFSET",
                                        "Dynamic OFFSET base",
                                        formula
                                    )

                    # D. 单元格超链接
                    if cell.hyperlink:
                        target_text = (
                            cell.hyperlink.location
                            or cell.hyperlink.target
                            or ""
                        )

                        target = extract_target(
                            target_text.lstrip("#"),
                            self.lookup
                        )

                        if target:
                            self.add_navigation(
                                ws.title,
                                target,
                                "CellHyperlink",
                                address,
                                target_text
                            )

                    count = self.stats["Formula Cells"]

                    if count % 10000 == 0:
                        print(
                            f"  Processed formulas: {count}",
                            flush=True
                        )

    # ========================================================
    # VBA 解析
    # ========================================================

    def scan_vba(self):
        if not SCAN_VBA:
            return

        print("Scanning VBA...", flush=True)

        try:
            from oletools.olevba import VBA_Parser
        except ImportError:
            print("WARNING: oletools not installed.")
            return

        parser = None

        try:
            parser = VBA_Parser(str(self.source))

            if not parser.detect_vba_macros():
                return

            for _, stream, module_name, code in parser.extract_macros():
                self.stats["VBA Modules"] += 1

                if isinstance(code, bytes):
                    code = code.decode("utf-8", errors="replace")

                self.parse_vba_module(
                    module_name,
                    code
                )

        except Exception as exc:
            print("VBA warning:", exc)

        finally:
            if parser is not None:
                try:
                    parser.close()
                except Exception:
                    pass

    def parse_vba_module(self, module_name, code):
        """
        解析 VBA 过程。

        注意：
        Activate / Select 是导航。
        Range.Value、Copy、Paste 等可能是数据流，
        不能仅凭出现 Sheets("X") 就确定数据依赖。
        """
        lines = code.splitlines()

        proc_pattern = re.compile(
            r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
            r"(?:Static\s+)?(?:Sub|Function)\s+"
            r"([^\s(]+)",
            re.I
        )

        sheet_pattern = re.compile(
            r"(?:ThisWorkbook\.)?"
            r"(?:Worksheets|Sheets)\s*\(\s*"
            r'"([^"]+)"\s*\)',
            re.I
        )

        current_proc = None
        current_lines = []

        def finish_procedure():
            if not current_proc:
                return

            body = "\n".join(current_lines)

            self.vba_procedures[norm(current_proc)] = {
                "module": module_name,
                "name": current_proc,
                "body": body,
            }

            refs = sheet_pattern.findall(body)

            for target_name in refs:
                target = actual_sheet(
                    target_name,
                    self.lookup
                )

                if not target:
                    continue

                self.stats["VBA Sheet References"] += 1

                # VBA 的数据流方向不能仅凭引用确定。
                # 因此先作为独立证据，不写入数据矩阵。
                self.add_navigation(
                    module_name,
                    target,
                    "VBA_SheetReference",
                    current_proc,
                    body[:1000]
                )

        for line in lines:
            match = proc_pattern.match(line)

            if match:
                finish_procedure()
                current_proc = match.group(1)
                current_lines = [line]
                self.stats["VBA Procedures"] += 1
                continue

            if current_proc:
                current_lines.append(line)

                if re.match(
                    r"^\s*End\s+(Sub|Function)\b",
                    line,
                    re.I
                ):
                    finish_procedure()
                    current_proc = None
                    current_lines = []

        finish_procedure()

    # ========================================================
    # 图形按钮 / VML / Drawing
    # ========================================================

    def scan_shapes(self):
        if not SCAN_SHAPES:
            return

        print("Scanning shape metadata...", flush=True)

        try:
            with zipfile.ZipFile(self.source) as z:
                names = set(z.namelist())

                # 读取 workbook sheet -> worksheet XML 的映射
                workbook_xml = ET.fromstring(
                    z.read("xl/workbook.xml")
                )

                rel_xml = ET.fromstring(
                    z.read("xl/_rels/workbook.xml.rels")
                )

                rels = {}

                for rel in rel_xml:
                    rels[rel.attrib.get("Id")] = rel.attrib.get(
                        "Target", ""
                    )

                ns = {
                    "m": (
                        "http://schemas.openxmlformats.org/"
                        "spreadsheetml/2006/main"
                    )
                }

                sheet_paths = {}

                for sheet in workbook_xml.findall(
                    ".//m:sheet", ns
                ):
                    name = sheet.attrib.get("name")
                    rid = sheet.attrib.get(
                        "{http://schemas.openxmlformats.org/"
                        "officeDocument/2006/relationships}id"
                    )

                    target = rels.get(rid, "")

                    if target.startswith("/"):
                        path = target.lstrip("/")
                    else:
                        path = posixpath.normpath(
                            posixpath.join("xl", target)
                        )

                    sheet_paths[path] = name

                # 每张 Sheet 的关系文件
                for path, sheet_name in sheet_paths.items():
                    if path not in names:
                        continue

                    rel_path = (
                        posixpath.dirname(path)
                        + "/_rels/"
                        + posixpath.basename(path)
                        + ".rels"
                    )

                    if rel_path not in names:
                        continue

                    root = ET.fromstring(z.read(rel_path))

                    for rel in root:
                        target = rel.attrib.get("Target", "")
                        typ = rel.attrib.get("Type", "")

                        if target.startswith("/"):
                            resource = target.lstrip("/")
                        else:
                            resource = posixpath.normpath(
                                posixpath.join(
                                    posixpath.dirname(path),
                                    target
                                )
                            )

                        if resource not in names:
                            continue

                        if "vmlDrawing" in typ:
                            self.parse_vml(
                                z.read(resource),
                                sheet_name,
                                resource
                            )

                        elif "drawing" in typ:
                            self.parse_drawing(
                                z.read(resource),
                                sheet_name,
                                resource
                            )

        except Exception as exc:
            print("Shape warning:", exc)

    def parse_vml(self, data, sheet_name, resource):
        """
        提取 VML Form Control 的 FmlaMacro。
        """
        try:
            root = ET.fromstring(data)
        except Exception:
            return

        for element in root.iter():
            if element.tag.split("}")[-1].lower() != "fmlamacro":
                continue

            macro = safe_text(element.text).strip()

            if macro:
                self.macro_bindings.append((
                    sheet_name,
                    macro,
                    resource
                ))

                self.stats["Macro Bindings"] += 1

    def parse_drawing(self, data, sheet_name, resource):
        """
        提取 Drawing 中可见的宏绑定属性。
        """
        try:
            root = ET.fromstring(data)
        except Exception:
            return

        for element in root.iter():
            macro = element.attrib.get("macro")

            if macro:
                self.macro_bindings.append((
                    sheet_name,
                    macro,
                    resource
                ))

                self.stats["Macro Bindings"] += 1

    def resolve_macro_bindings(self):
        print("Resolving macro bindings...", flush=True)

        for source_sheet, macro, location in self.macro_bindings:
            clean = macro.strip().lstrip("'")

            if "!" in clean:
                clean = clean.rsplit("!", 1)[-1]

            clean = clean.strip("'\" ")

            proc = self.vba_procedures.get(norm(clean))

            if not proc:
                self.add_unresolved(
                    source_sheet,
                    location,
                    "VBA_Button",
                    "Macro procedure not found",
                    macro
                )
                continue

            body = proc["body"]

            pattern = re.compile(
                r"(?:ThisWorkbook\.)?"
                r"(?:Worksheets|Sheets)\s*\(\s*"
                r'"([^"]+)"\s*\)'
                r"\s*\.\s*(?:Activate|Select)",
                re.I
            )

            targets = pattern.findall(body)

            if not targets:
                self.add_unresolved(
                    source_sheet,
                    location,
                    "VBA_Button",
                    "No static Activate/Select target",
                    macro
                )
                continue

            for target in targets:
                self.add_navigation(
                    source_sheet,
                    target,
                    "VBA_Button",
                    location,
                    macro
                )

    # ========================================================
    # 报告生成
    # ========================================================

    def write_table(self, wb, name, headers, rows):
        ws = wb.create_sheet(name)
        ws.append(headers)

        for row in rows:
            ws.append([
                self.safe_excel_value(value)
                for value in row
            ])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor="17365D"
            )
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.alignment = Alignment(
                horizontal="center"
            )

        for col in ws.columns:
            letter = col[0].column_letter

            if name in ("Evidence", "Unresolved", "Navigation"):
                width = {
                    "A": 28,
                    "B": 32,
                    "C": 24,
                    "D": 38,
                    "E": 100,
                }.get(letter, 22)
            else:
                width = 30

            ws.column_dimensions[letter].width = width

        # 重要：所有文本形式的公式均保存为文本，
        # 防止 Expression 出现 Excel 公式错误
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str):
                    cell.number_format = "@"
                    cell.data_type = "s"

        return ws

    @staticmethod
    def safe_excel_value(value):
        if value is None:
            return ""

        if isinstance(value, str):
            # Excel 单元格文本最大长度
            return value[:32767]

        return value

    def build_matrix(self, wb):
        ws = wb.create_sheet("Direct_Matrix")

        ws.cell(1, 1, "Source \\ Target")

        for i, name in enumerate(self.names, start=2):
            ws.cell(1, i, name)
            ws.cell(i, 1, name)

        for i, source in enumerate(self.names, start=2):
            confirmed = self.links.get(source, set())

            for j, target in enumerate(self.names, start=2):
                cell = ws.cell(i, j)

                if target in confirmed:
                    cell.value = 1
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="C6EFCE"
                    )

                elif MARK_UNKNOWN and source in self.unknown_sources:
                    cell.value = "?"
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="FFF2CC"
                    )

                else:
                    cell.value = 0

        ws.freeze_panes = "B2"

        ws.column_dimensions["A"].width = 38

        for i in range(2, len(self.names) + 2):
            ws.column_dimensions[
                get_column_letter(i)
            ].width = 17

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                "solid",
                fgColor="17365D"
            )
            cell.alignment = Alignment(
                textRotation=45
            )

        return ws

    def generate_report(self, output):
        print("Generating report...", flush=True)

        wb = Workbook()
        wb.remove(wb.active)

        summary_rows = [
            ("Source File", self.source.name),
            ("Total Sheets", len(self.names)),
            (
                "Confirmed Direct Links",
                sum(len(v) for v in self.links.values())
            ),
            (
                "Navigation Links",
                len(self.navigation)
            ),
            (
                "Formula Cells",
                self.stats["Formula Cells"]
            ),
            (
                "INDIRECT Calls",
                self.stats["INDIRECT Calls"]
            ),
            (
                "OFFSET Calls",
                self.stats["OFFSET Calls"]
            ),
            (
                "VBA Modules",
                self.stats["VBA Modules"]
            ),
            (
                "VBA Procedures",
                self.stats["VBA Procedures"]
            ),
            (
                "Macro Bindings",
                self.stats["Macro Bindings"]
            ),
            (
                "Evidence Rows",
                len(self.evidence)
            ),
            (
                "Unresolved Rows",
                len(self.unresolved)
            ),
            (
                "Unknown Source Sheets",
                len(self.unknown_sources)
            ),
            (
                "Truncated Evidence",
                self.truncated_evidence
            ),
            (
                "Truncated Unresolved",
                self.truncated_unresolved
            ),
        ]

        self.write_table(
            wb,
            "Summary",
            ["Metric", "Value"],
            summary_rows
        )

        self.build_matrix(wb)

        self.write_table(
            wb,
            "Data_Links",
            ["Source Sheet", "Target Sheet"],
            sorted(
                (source, target)
                for source, targets in self.links.items()
                for target in targets
            )
        )

        self.write_table(
            wb,
            "Evidence",
            [
                "Source Sheet",
                "Target Sheet",
                "Category",
                "Location",
                "Expression",
            ],
            self.evidence
        )

        self.write_table(
            wb,
            "Unresolved",
            [
                "Sheet",
                "Location",
                "Category",
                "Reason",
                "Expression",
            ],
            self.unresolved
        )

        self.write_table(
            wb,
            "Navigation",
            [
                "Source",
                "Target Sheet",
                "Category",
                "Location",
                "Expression",
            ],
            self.navigation
        )

        self.write_table(
            wb,
            "Sheet_Info",
            [
                "Sheet",
                "Visibility",
                "Confirmed Targets",
                "Has Unresolved",
            ],
            [
                (
                    ws.title,
                    ws.sheet_state,
                    len(self.links.get(ws.title, set())),
                    "YES" if ws.title in self.unknown_sources else "NO"
                )
                for ws in self.wb.worksheets
            ]
        )

        self.write_table(
            wb,
            "Method_Stats",
            ["Category", "Count"],
            sorted(self.stats.items())
        )

        wb.save(output)

        print()
        print("=" * 60)
        print("ANALYSIS COMPLETE")
        print("=" * 60)

        for key, value in summary_rows:
            print(f"{key}: {value}")

        print()
        print("Report:", output)
        print("=" * 60)

    def run(self, output):
        try:
            self.scan_formulas()
            self.scan_vba()
            self.scan_shapes()
            self.resolve_macro_bindings()
            self.generate_report(output)
        finally:
            try:
                self.wb.close()
            except Exception:
                pass

            try:
                self.cached_wb.close()
            except Exception:
                pass


# ============================================================
# 主程序
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Analyze Excel XLSM Sheet dependencies"
    )

    parser.add_argument(
        "input",
        help="Input .xlsm file"
    )

    parser.add_argument(
        "--output",
        help="Output .xlsx report",
        default=None
    )

    args = parser.parse_args()

    source = Path(args.input)

    if not source.exists():
        print("ERROR: Input file not found:", source)
        sys.exit(1)

    if source.suffix.lower() not in (".xlsm", ".xlsx"):
        print("ERROR: Only .xlsm / .xlsx supported")
        sys.exit(1)

    output = args.output

    if not output:
        output = str(
            source.with_name(
                source.stem + "_dependency_report.xlsx"
            )
        )

    if Path(output).resolve() == source.resolve():
        print("ERROR: Output cannot overwrite source.")
        sys.exit(1)

    analyzer = DependencyAnalyzer(source)
    analyzer.run(output)


if __name__ == "__main__":
    main()
