
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Direct Sheet Dependency Analyzer - Enhanced Version

Features
--------
1. All sheets, including hidden and veryHidden
2. Explicit cross-sheet formula references
3. INDIRECT with dynamic sheet names from cells
4. INDEX(INDIRECT(...)) and nested INDIRECT
5. Cached formula values
6. Static OFFSET references
7. VBA navigation macros and basic data assignments
8. Shape macros, VML controls and hyperlinks
9. Direct data/navigation/combined matrices
10. Detailed evidence, INDIRECT diagnostics and unresolved records
11. Safe text output for Excel formulas
12. Does not execute Excel macros or formulas

Usage
-----
python analyze_xlsm.py input.xlsm
python analyze_xlsm.py input.xlsm -o report.xlsx
python analyze_xlsm.py input.xlsm --exclude-self
"""

import argparse
import posixpath
import re
import zipfile

from pathlib import Path
from collections import defaultdict, Counter
from xml.etree import ElementTree as ET
from urllib.parse import unquote

from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import get_column_letter as col_letter

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    VBA_Parser = None


REL_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)
RID = "{" + REL_NS + "}id"

CELL_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$", re.I)
SHEET_CELL_RE = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!]+))!"
    r"(\$?[A-Z]{1,3}\$?\d+)$", re.I
)
VBA_SHEET_RE = re.compile(
    r'\b(Sheets|Worksheets)\s*\(\s*'
    r'(?:"((?:[^"]|"")*)"|(\d+))\s*\)', re.I
)
VBA_PROC_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function)\s+([^\s(]+)", re.I
)


def lname(tag):
    return tag.rsplit("}", 1)[-1]


def part_path(base, target):
    target = unquote(target or "")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base), target)
    )


def relationships(zf, part):
    folder = posixpath.dirname(part)
    name = posixpath.basename(part)
    relpart = posixpath.join(folder, "_rels", name + ".rels")
    result = {}
    if relpart not in zf.namelist():
        return result
    root = ET.fromstring(zf.read(relpart))
    for el in root:
        if lname(el.tag) == "Relationship":
            result[el.get("Id")] = {
                "target": el.get("Target", ""),
                "type": el.get("Type", ""),
                "mode": el.get("TargetMode", "")
            }
    return result


def split_top(expr, sep=","):
    result = []
    depth = 0
    quoted = False
    start = 0
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            if quoted and i + 1 < len(expr) and expr[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == sep and depth == 0:
                result.append(expr[start:i].strip())
                start = i + 1
        i += 1
    result.append(expr[start:].strip())
    return result


def function_calls(formula, function):
    pattern = re.compile(
        r"(?<![\w.])(?:_xlfn\.)?"
        + re.escape(function) + r"\s*\(", re.I
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
                if quoted and i + 1 < len(formula) and formula[i+1] == '"':
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
                            (formula[match.start():i+1],
                             split_top(formula[start:i]))
                        )
                        break
            i += 1
    return results


def strip_outer(expr):
    expr = expr.strip()
    while expr.startswith("(") and expr.endswith(")"):
        depth = 0
        quoted = False
        balanced = True
        for i, ch in enumerate(expr):
            if ch == '"':
                quoted = not quoted
            if quoted:
                continue
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0 and i != len(expr)-1:
                    balanced = False
                    break
        if not balanced or depth != 0:
            break
        expr = expr[1:-1].strip()
    return expr


def macro_name(value):
    value = str(value or "").strip().strip("'")
    if "!" in value:
        value = value.rsplit("!", 1)[-1]
    if "." in value:
        value = value.rsplit(".", 1)[-1]
    return value.casefold()


def logical_vba_lines(code):
    pending = ""
    start = 0
    for line_no, raw in enumerate(code.splitlines(), 1):
        # Remove comments without treating apostrophes in strings
        # as comment markers.
        chars = []
        quoted = False
        i = 0
        while i < len(raw):
            if raw[i] == '"':
                if quoted and i+1 < len(raw) and raw[i+1] == '"':
                    chars.extend(['"', '"'])
                    i += 2
                    continue
                quoted = not quoted
            if raw[i] == "'" and not quoted:
                break
            chars.append(raw[i])
            i += 1
        line = "".join(chars).strip()
        if not line:
            continue
        if not pending:
            start = line_no
        if re.search(r"\s+_$", line):
            pending += re.sub(r"\s+_$", " ", line)
        else:
            yield start, pending + line
            pending = ""
    if pending:
        yield start, pending


def report_text(cell, value):
    value = "" if value is None else str(value)
    # Force text storage, including expressions beginning with '='.
    cell.value = value[:32000]
    cell.data_type = "s"
    cell.number_format = "@"


class Analyzer:

    def __init__(self, filename, include_self=True):
        self.path = Path(filename)
        self.include_self = include_self

        self.wb = load_workbook(
            self.path, read_only=True,
            data_only=False, keep_links=True
        )
        self.cached = load_workbook(
            self.path, read_only=True,
            data_only=True, keep_links=True
        )

        self.names = list(self.wb.sheetnames)
        self.lookup = {n.casefold(): n for n in self.names}
        self.states = {
            ws.title: ws.sheet_state for ws in self.wb.worksheets
        }

        self.data = defaultdict(set)
        self.navigation = defaultdict(set)
        self.evidence = set()
        self.unresolved = set()
        self.indirect_details = []
        self.stats = Counter()
        self.procedures = defaultdict(list)
        self.proc_targets = defaultdict(set)
        self.code_names = {}
        self.cell_cache = {}
        self.value_cache = {}
        self.links_by_type = Counter()

        self.load_code_names()

    def norm(self, name):
        if not isinstance(name, str):
            return None
        return self.lookup.get(name.strip().casefold())

    def add_link(self, source, target, kind, reason, location, expr):
        source = self.norm(source)
        target = self.norm(target)
        if not source or not target:
            return False
        if source == target and not self.include_self:
            return False

        graph = self.navigation if kind == "Navigation" else self.data
        graph[source].add(target)

        row = (source, target, kind, reason, str(location), str(expr))
        if row not in self.evidence:
            self.evidence.add(row)
            self.links_by_type[reason] += 1
        return True

    def add_unresolved(self, sheet, location, reason, expr):
        self.unresolved.add(
            (str(sheet), str(location), str(reason), str(expr))
        )

    def sheet_parts(self, zf):
        part = "xl/workbook.xml"
        root = ET.fromstring(zf.read(part))
        rels = relationships(zf, part)
        mapping = {}
        for el in root.iter():
            if lname(el.tag) != "sheet":
                continue
            name, rid = el.get("name"), el.get(RID)
            if rid in rels:
                mapping[name] = part_path(part, rels[rid]["target"])
        return mapping

    def load_code_names(self):
        # Works with ReadOnlyWorksheet: no sheet_properties access.
        with zipfile.ZipFile(self.path) as zf:
            parts = self.sheet_parts(zf)
            available = set(zf.namelist())
            for sheet, part in parts.items():
                if part not in available:
                    continue
                with zf.open(part) as stream:
                    for _, el in ET.iterparse(stream, events=("start",)):
                        if lname(el.tag) == "sheetPr":
                            code = el.get("codeName")
                            if code:
                                self.code_names[code.casefold()] = sheet
                            break

    def cell_value(self, sheet, address, depth=0, seen=None):
        if depth > 12:
            return None
        sheet = self.norm(sheet)
        if not sheet:
            return None
        address = address.replace("$", "").upper()
        key = (sheet, address)
        seen = seen or set()
        if key in seen:
            return None

        if key in self.value_cache:
            return self.value_cache[key]

        try:
            cell = self.wb[sheet][address]
            if cell.data_type != "f":
                value = cell.value
            else:
                value = self.cached[sheet][address].value
                if value is None:
                    value = self.eval_expr(
                        sheet, str(cell.value),
                        depth + 1, seen | {key}
                    )
            self.value_cache[key] = value
            return value
        except Exception:
            return None

    def eval_expr(self, sheet, expr, depth=0, seen=None):
        """
        Evaluate a deliberately limited subset of Excel expressions.
        Never execute arbitrary functions or VBA.
        """
        if depth > 12 or expr is None:
            return None

        expr = strip_outer(str(expr).strip())
        if expr.startswith("="):
            expr = expr[1:].strip()

        parts = split_top(expr, "&")
        if len(parts) > 1:
            values = [
                self.eval_expr(sheet, p, depth+1, seen)
                for p in parts
            ]
            if any(v is None for v in values):
                return None
            return "".join(str(v) for v in values)

        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1].replace('""', '"')

        if re.fullmatch(r"-?\d+(?:\.\d+)?", expr):
            num = float(expr)
            return int(num) if num.is_integer() else num

        if CELL_RE.fullmatch(expr):
            return self.cell_value(sheet, expr, depth+1, seen)

        m = SHEET_CELL_RE.fullmatch(expr)
        if m:
            target = (m.group(1) or m.group(2)).replace("''", "'")
            return self.cell_value(target, m.group(3), depth+1, seen)

        m = re.fullmatch(
            r"(CONCATENATE|CONCAT|UPPER|LOWER|TRIM|IFERROR)"
            r"\s*\((.*)\)", expr, re.I | re.S
        )
        if m:
            func = m.group(1).upper()
            args = split_top(m.group(2))
            if func == "IFERROR":
                if not args:
                    return None
                val = self.eval_expr(sheet, args[0], depth+1, seen)
                if val is None and len(args) > 1:
                    return self.eval_expr(sheet, args[1], depth+1, seen)
                return val

            vals = [
                self.eval_expr(sheet, a, depth+1, seen)
                for a in args
            ]
            if any(v is None for v in vals):
                return None
            vals = [str(v) for v in vals]
            if func in ("CONCATENATE", "CONCAT"):
                return "".join(vals)
            if len(vals) != 1:
                return None
            if func == "UPPER":
                return vals[0].upper()
            if func == "LOWER":
                return vals[0].lower()
            return " ".join(vals[0].split())

        return None

    def sheet_from_ref(self, value):
        if not isinstance(value, str):
            return None
        value = value.strip().lstrip("#")
        if "!" not in value:
            return None
        prefix = value.rsplit("!", 1)[0].strip()
        if prefix.startswith("'") and prefix.endswith("'"):
            prefix = prefix[1:-1].replace("''", "'")
        if "[" in prefix or "]" in prefix:
            return None
        return self.norm(prefix)

    def explicit_targets(self, formula):
        targets = set()
        try:
            tokens = Tokenizer(formula).items
        except Exception:
            self.stats["tokenizer_errors"] += 1
            return targets

        for token in tokens:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            value = token.value
            if "!" not in value:
                continue
            prefix = value.rsplit("!", 1)[0].strip()
            if prefix.startswith("'") and prefix.endswith("'"):
                prefix = prefix[1:-1].replace("''", "'")
            if "[" in prefix or "]" in prefix:
                continue

            if ":" in prefix:
                a, b = prefix.split(":", 1)
                a, b = self.norm(a), self.norm(b)
                if a and b:
                    i, j = self.names.index(a), self.names.index(b)
                    targets.update(
                        self.names[min(i, j):max(i, j)+1]
                    )
            else:
                target = self.norm(prefix)
                if target:
                    targets.add(target)
        return targets

    def indirect_static_targets(self, expression):
        """
        Find explicit sheet-name literals inside INDIRECT arguments.
        Avoid matching arbitrary identifiers without '!'.
        """
        targets = set()
        literals = re.findall(
            r'"((?:[^"]|"")*)"', expression
        )
        for literal in literals:
            literal = literal.replace('""', '"')
            for name in self.names:
                candidates = (
                    name + "!",
                    "'" + name.replace("'", "''") + "'!"
                )
                if any(c.casefold() in literal.casefold()
                       for c in candidates):
                    targets.add(name)
        return targets

    def dynamic_sheet_candidates(self, sheet, expression):
        """
        Recognize:
            "'" & $E36 & "'!A1"
            E36 & "!A1"
            "'" & 'TOP'!E36 & "'!A1"

        Return confirmed targets only if the cell value
        matches an actual worksheet name.
        """
        targets = set()
        parts = split_top(expression, "&")

        if len(parts) < 2:
            return targets

        has_separator = any(
            "!" in p for p in parts
        )
        if not has_separator:
            return targets

        for p in parts:
            p = strip_outer(p.strip())
            value = None
            if CELL_RE.fullmatch(p):
                value = self.cell_value(sheet, p)
            else:
                m = SHEET_CELL_RE.fullmatch(p)
                if m:
                    src = (m.group(1) or m.group(2)).replace("''", "'")
                    value = self.cell_value(src, m.group(3))

            if isinstance(value, str):
                target = self.norm(value)
                if target:
                    targets.add(target)

        return targets

    def analyze_indirect(self, sheet, address, formula):
        for full_call, args in function_calls(formula, "INDIRECT"):
            self.stats["indirect_calls"] += 1
            if not args:
                continue

            expr = args[0]
            location = f"{sheet}!{address}"
            target = None
            status = "Unresolved"
            detail = ""

            value = self.eval_expr(sheet, expr)
            if isinstance(value, str):
                target = self.sheet_from_ref(value)

            if target:
                status = "Resolved expression"
                detail = str(value)
                self.add_link(
                    sheet, target, "Data",
                    "INDIRECT Resolved", location, full_call
                )
                self.stats["indirect_resolved"] += 1
            else:
                candidates = self.dynamic_sheet_candidates(sheet, expr)
                if not candidates:
                    candidates = self.indirect_static_targets(expr)

                if candidates:
                    status = "Resolved sheet only"
                    detail = ", ".join(sorted(candidates))
                    for candidate in candidates:
                        self.add_link(
                            sheet, candidate, "Data",
                            "INDIRECT Sheet Resolved",
                            location, full_call
                        )
                    self.stats["indirect_sheet_resolved"] += 1
                else:
                    status = "Unresolved"
                    detail = (
                        "Dynamic sheet name could not be "
                        "matched to workbook sheets"
                    )
                    self.stats["indirect_unresolved"] += 1
                    self.add_unresolved(
                        sheet, address, detail, full_call
                    )

            self.indirect_details.append((
                sheet, address, status,
                target or detail, expr, full_call
            ))

    def analyze_formulas(self):
        print("[1/5] Scanning formulas...", flush=True)

        for sheet in self.names:
            ws = self.wb[sheet]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["formula_cells"] += 1
                    formula = str(cell.value)
                    location = f"{sheet}!{cell.coordinate}"

                    for target in self.explicit_targets(formula):
                        self.add_link(
                            sheet, target, "Data",
                            "Explicit Formula", location, formula
                        )

                    if re.search(r"\bINDIRECT\s*\(", formula, re.I):
                        self.analyze_indirect(
                            sheet, cell.coordinate, formula
                        )

                    self.stats["offset_calls"] += len(
                        function_calls(formula, "OFFSET")
                    )

                    for _, args in function_calls(formula, "HYPERLINK"):
                        if not args:
                            continue
                        value = self.eval_expr(sheet, args[0])
                        target = self.sheet_from_ref(value)
                        if target:
                            self.add_link(
                                sheet, target, "Navigation",
                                "HYPERLINK Formula", location, formula
                            )

                    if self.stats["formula_cells"] % 10000 == 0:
                        print(
                            "  Formula cells:",
                            self.stats["formula_cells"],
                            flush=True
                        )

    def vba_targets(self, expr, aliases=None):
        aliases = aliases or {}
        result = set()

        for m in VBA_SHEET_RE.finditer(expr):
            kind, name, number = m.groups()
            if name is not None:
                target = self.norm(name.replace('""', '"'))
            else:
                index = int(number)
                sequence = (
                    self.names if kind.casefold() == "sheets"
                    else self.names
                )
                target = (
                    sequence[index-1]
                    if 1 <= index <= len(sequence)
                    else None
                )
            if target:
                result.add(target)

        for code, sheet in self.code_names.items():
            if re.search(
                r"(?<![\w.])" + re.escape(code) + r"\s*[.!]",
                expr, re.I
            ):
                result.add(sheet)

        for alias, sheet in aliases.items():
            if re.search(
                r"(?<![\w.])" + re.escape(alias) + r"\s*[.!]",
                expr, re.I
            ):
                result.add(sheet)
        return result

    def analyze_vba(self):
        print("[2/5] Scanning VBA...", flush=True)

        if VBA_Parser is None:
            self.add_unresolved(
                "VBA", "", "oletools not installed", ""
            )
            return

        parser = VBA_Parser(str(self.path))
        try:
            if parser.detect_vba_macros():
                for _, _, module, code in parser.extract_macros():
                    self.stats["vba_modules"] += 1
                    proc = None
                    for line_no, line in logical_vba_lines(code):
                        m = VBA_PROC_RE.match(line)
                        if m:
                            proc = macro_name(m.group(1))
                            continue
                        if re.match(
                            r"^\s*End\s+(Sub|Function)\b",
                            line, re.I
                        ):
                            proc = None
                            continue
                        if proc:
                            self.procedures[proc].append(
                                (module, line_no, line)
                            )
        finally:
            parser.close()

        direct = defaultdict(set)
        calls = defaultdict(set)

        for proc, lines in self.procedures.items():
            aliases = {}
            for module, line_no, line in lines:
                m = re.match(
                    r"^Set\s+(\w+)\s*=\s*(.+)$", line, re.I
                )
                if m:
                    refs = self.vba_targets(m.group(2), aliases)
                    if len(refs) == 1:
                        aliases[m.group(1).casefold()] = next(iter(refs))
                    continue

                nav = re.search(
                    r"\.\s*(Activate|Select)\b", line, re.I
                )
                if nav:
                    direct[proc].update(
                        self.vba_targets(line[:nav.start()], aliases)
                    )

                if re.search(
                    r"\b(?:Application\s*\.\s*)?Goto\b",
                    line, re.I
                ):
                    direct[proc].update(
                        self.vba_targets(line, aliases)
                    )

                if "=" in line and not re.match(
                    r"^(If|ElseIf|For|While|Do|Select|"
                    r"Case|Const|Dim)\b", line, re.I
                ):
                    left, right = line.split("=", 1)
                    left_refs = self.vba_targets(left, aliases)
                    right_refs = self.vba_targets(right, aliases)
                    for dest in left_refs:
                        for source in right_refs:
                            self.add_link(
                                dest, source, "Data",
                                "VBA Assignment",
                                f"{module}:{line_no}", line
                            )

                call = re.match(
                    r"^(?:Call\s+)?([^\s(]+)", line, re.I
                )
                if call:
                    called = macro_name(call.group(1))
                    if called in self.procedures and called != proc:
                        calls[proc].add(called)

        resolved = {
            p: set(direct[p]) for p in self.procedures
        }

        changed = True
        while changed:
            changed = False
            for proc in self.procedures:
                before = len(resolved[proc])
                for called in calls[proc]:
                    resolved[proc].update(resolved.get(called, set()))
                if len(resolved[proc]) != before:
                    changed = True

        self.proc_targets = resolved
        self.stats["vba_procedures"] = len(self.procedures)

    def shape_macro(self, source, macro, location):
        self.stats["shape_macro_bindings"] += 1
        targets = self.proc_targets.get(macro_name(macro), set())

        if not targets:
            self.add_unresolved(
                source, location,
                "Shape macro target unresolved", macro
            )
            return

        for target in targets:
            self.add_link(
                source, target, "Navigation",
                "Shape VBA Macro", location, macro
            )

    def analyze_shapes(self):
        print("[3/5] Scanning shapes and hyperlinks...", flush=True)

        with zipfile.ZipFile(self.path) as zf:
            available = set(zf.namelist())
            mapping = self.sheet_parts(zf)

            for source, sheet_part in mapping.items():
                if sheet_part not in available:
                    continue

                rels = relationships(zf, sheet_part)

                # Cell hyperlinks
                with zf.open(sheet_part) as stream:
                    for _, el in ET.iterparse(
                        stream, events=("end",)
                    ):
                        if lname(el.tag) == "hyperlink":
                            location = el.get("location", "")
                            rid = el.get(RID)
                            if not location and rid:
                                location = rels.get(
                                    rid, {}
                                ).get("target", "")
                            target = self.sheet_from_ref(location)
                            if target:
                                self.add_link(
                                    source, target, "Navigation",
                                    "Cell Hyperlink",
                                    el.get("ref", ""), location
                                )
                        el.clear()

                for relation in rels.values():
                    if relation["mode"] == "External":
                        continue
                    target_part = part_path(
                        sheet_part, relation["target"]
                    )
                    if target_part not in available:
                        continue

                    rtype = relation["type"]

                    if rtype.endswith("/drawing"):
                        self.scan_drawing(zf, source, target_part)

                    elif rtype.endswith("/vmlDrawing"):
                        self.scan_vml(zf, source, target_part)

    def scan_drawing(self, zf, source, part):
        root = ET.fromstring(zf.read(part))
        rels = relationships(zf, part)

        for el in root.iter():
            tag = lname(el.tag)

            if tag in ("sp", "pic", "graphicFrame"):
                macro = el.get("macro")
                if macro:
                    self.shape_macro(source, macro, part)

            if tag == "hlinkClick":
                location = el.get("location", "")
                rid = el.get(RID)
                if not location and rid:
                    location = rels.get(
                        rid, {}
                    ).get("target", "")
                target = self.sheet_from_ref(location)
                if target:
                    self.add_link(
                        source, target, "Navigation",
                        "Shape Hyperlink", part, location
                    )

    def scan_vml(self, zf, source, part):
        root = ET.fromstring(zf.read(part))
        for el in root.iter():
            if lname(el.tag) != "ClientData":
                continue
            for child in el:
                if lname(child.tag) == "FmlaMacro":
                    macro = (child.text or "").strip()
                    if macro:
                        self.shape_macro(source, macro, part)

    def matrix(self, wb, title, graph):
        ws = wb.create_sheet(title)
        report_text(ws.cell(1, 1), "Source / Target")

        for i, name in enumerate(self.names, 2):
            report_text(ws.cell(1, i), name)
            report_text(ws.cell(i, 1), name)

        for i, source in enumerate(self.names, 2):
            for j, target in enumerate(self.names, 2):
                value = int(target in graph[source])
                cell = ws.cell(i, j, value)
                cell.alignment = Alignment(horizontal="center")
                if value:
                    cell.fill = PatternFill(
                        "solid", fgColor="C6EFCE"
                    )
                    cell.font = Font(
                        color="006100", bold=True
                    )

        self.style(ws)
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 38
        for i in range(2, len(self.names)+2):
            ws.column_dimensions[
                get_column_letter(i)
            ].width = 25

    def style(self, ws):
        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid", fgColor="1F4E78"
            )
            cell.font = Font(
                color="FFFFFF", bold=True
            )

    def detail_sheet(self, wb, title, headers, rows, widths):
        ws = wb.create_sheet(title)

        for col, header in enumerate(headers, 1):
            report_text(ws.cell(1, col), header)

        for row_num, record in enumerate(rows, 2):
            for col, value in enumerate(record, 1):
                report_text(ws.cell(row_num, col), value)

        self.style(ws)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for i, width in enumerate(widths, 1):
            ws.column_dimensions[col_letter(i)].width = width

    def report(self, output):
        print("[4/5] Building matrices...", flush=True)

        combined = {
            name: self.data[name] | self.navigation[name]
            for name in self.names
        }

        wb = Workbook()
        wb.remove(wb.active)

        self.matrix(wb, "Direct_Matrix", combined)
        self.matrix(wb, "Data_Matrix", self.data)
        self.matrix(wb, "Navigation_Matrix", self.navigation)

        self.detail_sheet(
            wb, "Sheet_Information",
            ["Sheet", "Visibility"],
            [(n, self.states.get(n, "unknown")) for n in self.names],
            [45, 20]
        )

        self.detail_sheet(
            wb, "Dependency_Details",
            [
                "Source", "Target", "Kind",
                "Reason", "Location", "Expression"
            ],
            sorted(self.evidence),
            [35, 35, 18, 32, 40, 110]
        )

        self.detail_sheet(
            wb, "INDIRECT_Details",
            [
                "Source Sheet", "Cell", "Status",
                "Resolved Target / Detail",
                "INDIRECT Argument", "Full Call"
            ],
            self.indirect_details,
            [35, 16, 28, 55, 100, 100]
        )

        self.detail_sheet(
            wb, "Unresolved_References",
            ["Source", "Location", "Reason", "Expression"],
            sorted(self.unresolved),
            [35, 35, 55, 110]
        )

        stats = [
            ("Sheets", len(self.names)),
            ("Hidden Sheets", sum(
                v != "visible" for v in self.states.values()
            )),
            ("Formula Cells", self.stats["formula_cells"]),
            ("INDIRECT Calls", self.stats["indirect_calls"]),
            ("INDIRECT Resolved", self.stats["indirect_resolved"]),
            ("INDIRECT Sheet Resolved",
             self.stats["indirect_sheet_resolved"]),
            ("INDIRECT Unresolved",
             self.stats["indirect_unresolved"]),
            ("OFFSET Calls", self.stats["offset_calls"]),
            ("VBA Modules", self.stats["vba_modules"]),
            ("VBA Procedures", self.stats["vba_procedures"]),
            ("Shape Macro Bindings",
             self.stats["shape_macro_bindings"]),
            ("Data Link Pairs", sum(
                len(v) for v in self.data.values()
            )),
            ("Navigation Link Pairs", sum(
                len(v) for v in self.navigation.values()
            )),
            ("Combined Link Pairs", sum(
                len(v) for v in combined.values()
            )),
            ("Evidence Rows", len(self.evidence)),
            ("Unresolved Rows", len(self.unresolved)),
            ("Tokenizer Errors", self.stats["tokenizer_errors"])
        ]

        self.detail_sheet(
            wb, "Statistics",
            ["Metric", "Value"],
            stats, [42, 25]
        )

        print("[5/5] Saving report...", flush=True)
        wb.save(output)
        return stats

    def run(self, output):
        try:
            self.analyze_formulas()
            self.analyze_vba()
            self.analyze_shapes()
            stats = self.report(output)

            print("\n========== ANALYSIS RESULT ==========")
            for key, value in stats:
                print(f"{key}: {value}")
            print("Report:", output)
            print("=====================================")
        finally:
            self.wb.close()
            self.cached.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Input XLSM/XLSX file")
    parser.add_argument("-o", "--output")
    parser.add_argument("--exclude-self", action="store_true")
    args = parser.parse_args()

    source = Path(args.input)

    if not source.is_file():
        parser.error(f"File not found: {source}")

    if source.suffix.lower() not in (".xlsm", ".xlsx"):
        parser.error("Input must be XLSM or XLSX")

    output = (
        Path(args.output)
        if args.output
        else source.with_name(
            source.stem + "_dependency_report.xlsx"
        )
    )

    if output.suffix.lower() != ".xlsx":
        parser.error("Output must be .xlsx")

    if source.resolve() == output.resolve():
        parser.error("Output must not overwrite input")

    analyzer = Analyzer(
        source,
        include_self=not args.exclude_self
    )
    analyzer.run(output)


if __name__ == "__main__":
    main()
