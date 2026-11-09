
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Dependency Analyzer V4
===========================

Python 3.10+

Purpose:
    Analyze worksheet dependencies, dynamic references,
    drawing hyperlinks, shape macros, VBA navigation
    and hidden worksheets.

Output:
    *_dependency_report_v4.xlsx

Important:
    Source workbook is NEVER saved.
    VBA macros are NEVER executed.

Install:
    pip install openpyxl oletools

Run:
    python analyze_xlsm_v4.py input.xlsm

Optional:
    python analyze_xlsm_v4.py input.xlsm -o report.xlsx
"""

import argparse
import collections
import posixpath
import re
import sys
import traceback
import zipfile
import xml.etree.ElementTree as ET

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


# ============================================================
# XML namespaces
# ============================================================

MAIN = (
    "http://schemas.openxmlformats.org/"
    "spreadsheetml/2006/main"
)

OFFICE_REL = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

NS = {
    "m": MAIN,
    "r": OFFICE_REL,
}

MAX_TEXT = 32000


# ============================================================
# Basic utilities
# ============================================================

def local(tag):
    return tag.rsplit("}", 1)[-1]


def safe(value):
    if value is None:
        return ""
    return str(value)[:MAX_TEXT]


def key(value):
    return str(value).casefold().strip()


def normalize_path(base, target):
    if not target:
        return ""

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(base),
            target
        )
    )


def rels_path(part):
    return posixpath.join(
        posixpath.dirname(part),
        "_rels",
        posixpath.basename(part) + ".rels"
    )


def read_relationships(z, part):
    path = rels_path(part)

    if path not in z.namelist():
        return {}

    root = ET.fromstring(z.read(path))
    result = {}

    for item in root:
        result[item.get("Id", "")] = {
            "type": item.get("Type", "").rsplit("/", 1)[-1],
            "target": item.get("Target", ""),
            "mode": item.get("TargetMode", ""),
        }

    return result


def sheet_token(value):
    value = str(value).strip()

    if value.startswith("#"):
        value = value[1:]

    if value.startswith("="):
        value = value[1:]

    if "!" not in value:
        return None

    token = value.split("!", 1)[0].strip()

    if token.startswith("'") and token.endswith("'"):
        token = token[1:-1].replace("''", "'")

    if "[" in token or "]" in token:
        return None

    return token


def remove_string_literals(formula):
    return re.sub(
        r'"(?:[^"]|"")*"',
        lambda m: " " * len(m.group()),
        formula
    )


# ============================================================
# Formula parsing
# ============================================================

def function_calls(formula, name):
    pattern = re.compile(
        r"\b" + re.escape(name) + r"\s*\(",
        re.I
    )

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
                        yield formula[start:i]
                        break

            i += 1


def split_top(expression, separators):
    result = []
    buffer = []
    depth = 0
    quoted = False
    i = 0

    while i < len(expression):
        ch = expression[i]

        if ch == '"':
            if (
                quoted
                and i + 1 < len(expression)
                and expression[i + 1] == '"'
            ):
                buffer.extend(['"', '"'])
                i += 2
                continue

            quoted = not quoted

        if not quoted:
            if ch == "(":
                depth += 1

            elif ch == ")":
                depth -= 1

            elif ch in separators and depth == 0:
                result.append("".join(buffer).strip())
                buffer = []
                i += 1
                continue

        buffer.append(ch)
        i += 1

    result.append("".join(buffer).strip())
    return result


def function_args(expression):
    return split_top(expression, ",;")


def balanced(expr):
    depth = 0
    quoted = False

    for ch in expr:
        if ch == '"':
            quoted = not quoted

        if not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth < 0:
                    return False

    return depth == 0


def strip_parens(expr):
    expr = expr.strip()

    while (
        expr.startswith("(")
        and expr.endswith(")")
        and balanced(expr[1:-1])
    ):
        expr = expr[1:-1].strip()

    return expr


# ============================================================
# Workbook model
# ============================================================

class WorkbookModel:

    def __init__(self, source):
        self.source = source

        self.formula_wb = load_workbook(
            source,
            read_only=False,
            data_only=False,
            keep_links=False,
            keep_vba=False
        )

        self.value_wb = load_workbook(
            source,
            read_only=True,
            data_only=True,
            keep_links=False
        )

        self.sheets = self.formula_wb.sheetnames

        self.lookup = {
            key(name): name
            for name in self.sheets
        }

        self.merged_map = {}
        self.formulas = collections.defaultdict(dict)

        for ws in self.formula_wb.worksheets:
            ranges = []

            for area in ws.merged_cells.ranges:
                ranges.append((
                    area.min_row,
                    area.max_row,
                    area.min_col,
                    area.max_col,
                    ws.cell(
                        area.min_row,
                        area.min_col
                    ).coordinate
                ))

            self.merged_map[ws.title] = ranges

    def canonical(self, name):
        if name is None:
            return None

        return self.lookup.get(key(name))

    def value(self, sheet, address):
        sheet = self.canonical(sheet)

        if sheet is None:
            return None

        address = address.replace("$", "").upper()

        try:
            value = self.value_wb[sheet][address].value

            if value is not None:
                return value

            cell = self.formula_wb[sheet][address]

            if (
                cell.value is not None
                and cell.data_type != "f"
            ):
                return cell.value

            row = cell.row
            col = cell.column

            for (
                min_row, max_row,
                min_col, max_col,
                origin
            ) in self.merged_map[sheet]:

                if (
                    min_row <= row <= max_row
                    and min_col <= col <= max_col
                ):
                    return self.value_wb[sheet][origin].value

        except Exception:
            return None

        return None

    def close(self):
        self.formula_wb.close()
        self.value_wb.close()


# ============================================================
# Safe dynamic expression evaluator
# ============================================================

class Evaluator:

    def __init__(self, model):
        self.model = model

    def evaluate(
        self,
        expression,
        sheet,
        address,
        depth=0
    ):
        if depth > 15:
            return None

        expr = strip_parens(expression)

        if not expr:
            return None

        parts = split_top(expr, "&")

        if len(parts) > 1:
            values = []

            for part in parts:
                value = self.evaluate(
                    part, sheet, address, depth + 1
                )

                if value is None:
                    return None

                values.append(str(value))

            return "".join(values)

        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1].replace('""', '"')

        if re.fullmatch(r"-?\d+(?:\.\d+)?", expr):
            return (
                float(expr)
                if "." in expr
                else int(expr)
            )

        # Cell reference
        match = re.fullmatch(
            r"(?:(?:'((?:[^']|'')+)'|"
            r"([A-Za-z_\u0080-\uffff][\w.\u0080-\uffff]*))!)?"
            r"(\$?[A-Z]{1,3}\$?\d+)",
            expr,
            re.I
        )

        if match:
            target = (
                match.group(1)
                or match.group(2)
                or sheet
            )

            return self.model.value(
                target,
                match.group(3)
            )

        # ROW / COLUMN
        if re.fullmatch(r"ROW\(\s*\)", expr, re.I):
            return self.model.formula_wb[sheet][address].row

        if re.fullmatch(r"COLUMN\(\s*\)", expr, re.I):
            return self.model.formula_wb[sheet][address].column

        match = re.fullmatch(
            r"(ROW|COLUMN)\(\s*"
            r"(\$?[A-Z]{1,3}\$?\d+)"
            r"\s*\)",
            expr,
            re.I
        )

        if match:
            from openpyxl.utils.cell import (
                coordinate_from_string,
                column_index_from_string
            )

            col, row = coordinate_from_string(
                match.group(2).replace("$", "")
            )

            if match.group(1).upper() == "ROW":
                return row

            return column_index_from_string(col)

        # SUBSTITUTE
        for func in (
            "SUBSTITUTE",
            "ADDRESS",
            "IF",
            "IFERROR",
            "CONCATENATE",
            "TEXT"
        ):
            if not re.match(
                rf"^{func}\s*\(",
                expr,
                re.I
            ):
                continue

            calls = list(function_calls(expr, func))

            if not calls:
                return None

            args = function_args(calls[0])

            if func == "SUBSTITUTE" and len(args) >= 3:
                values = [
                    self.evaluate(
                        a, sheet, address, depth + 1
                    )
                    for a in args[:3]
                ]

                if all(v is not None for v in values):
                    return str(values[0]).replace(
                        str(values[1]),
                        str(values[2])
                    )

            if func == "CONCATENATE":
                values = [
                    self.evaluate(
                        a, sheet, address, depth + 1
                    )
                    for a in args
                ]

                if all(v is not None for v in values):
                    return "".join(map(str, values))

            if func == "IF" and len(args) >= 2:
                condition = self.evaluate(
                    args[0], sheet, address, depth + 1
                )

                if isinstance(condition, bool):
                    selected = (
                        args[1]
                        if condition
                        else (
                            args[2]
                            if len(args) > 2
                            else '""'
                        )
                    )

                    return self.evaluate(
                        selected,
                        sheet,
                        address,
                        depth + 1
                    )

            if func == "IFERROR" and len(args) >= 2:
                result = self.evaluate(
                    args[0], sheet, address, depth + 1
                )

                if result is not None:
                    return result

                return self.evaluate(
                    args[1], sheet, address, depth + 1
                )

            if func == "ADDRESS" and len(args) >= 2:
                row = self.evaluate(
                    args[0], sheet, address, depth + 1
                )

                col = self.evaluate(
                    args[1], sheet, address, depth + 1
                )

                if (
                    isinstance(row, (int, float))
                    and isinstance(col, (int, float))
                ):
                    row = int(row)
                    col = int(col)

                    if (
                        1 <= row <= 1048576
                        and 1 <= col <= 16384
                    ):
                        return (
                            f"${get_column_letter(col)}"
                            f"${row}"
                        )

            if func == "TEXT" and len(args) >= 2:
                value = self.evaluate(
                    args[0], sheet, address, depth + 1
                )

                if value is not None:
                    return str(value)

        return None


# ============================================================
# Evidence collector
# ============================================================

class Collector:

    def __init__(self, model):
        self.model = model

        self.data = set()
        self.navigation = set()

        self.possible_data = set()
        self.possible_navigation = set()

        self.evidence = []
        self.unresolved = []
        self.shapes = []
        self.macro_routes = []

        self.seen = set()
        self.stats = collections.Counter()

    def add(
        self,
        source,
        target,
        category,
        location,
        expression,
        detail="",
        confidence="Confirmed",
        relation="Data"
    ):
        source = self.model.canonical(source)
        target = self.model.canonical(target)

        if not source or not target:
            return

        record = (
            source,
            target,
            category,
            location,
            expression,
            detail,
            confidence,
            relation
        )

        if record in self.seen:
            return

        self.seen.add(record)

        if relation == "Data":
            if confidence == "Confirmed":
                self.data.add((source, target))
            else:
                self.possible_data.add((source, target))
        else:
            if confidence == "Confirmed":
                self.navigation.add((source, target))
            else:
                self.possible_navigation.add((source, target))

        self.evidence.append(record)

    def unresolved_add(
        self,
        sheet,
        address,
        category,
        expression,
        reason
    ):
        self.unresolved.append((
            sheet,
            address,
            category,
            expression,
            reason
        ))

    def finalize(self):
        self.possible_data -= self.data
        self.possible_navigation -= self.navigation


# ============================================================
# Formula analyzer
# ============================================================

DIRECT_REF = re.compile(
    r"(?:'((?:[^']|'')+)'|"
    r"([A-Za-z_\u0080-\uffff][\w.\u0080-\uffff]*))!"
    r"(?=\$?[A-Z]{1,3}\$?\d+"
    r"|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}"
    r"|\$?\d+:\$?\d+)",
    re.I
)


class FormulaAnalyzer:

    def __init__(self, model, collector):
        self.model = model
        self.collector = collector
        self.evaluator = Evaluator(model)

    def scan(self):
        print("[1/5] Scanning formulas...", flush=True)

        for ws in self.model.formula_wb.worksheets:
            count = 0

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    formula = cell.value

                    if not isinstance(formula, str):
                        # ArrayFormula or other formula objects
                        formula = getattr(
                            formula, "text", None
                        )

                    if not formula:
                        continue

                    count += 1
                    self.collector.stats["Formula Cells"] += 1

                    self.scan_one(
                        ws.title,
                        cell.coordinate,
                        formula
                    )

            print(
                f"  {ws.title}: {count} formulas",
                flush=True
            )

    def scan_one(self, sheet, address, formula):
        c = self.collector

        clean = remove_string_literals(formula)

        # Direct references
        for match in DIRECT_REF.finditer(clean):
            token = (
                match.group(1)
                or match.group(2)
            ).replace("''", "'")

            if "[" in token or "]" in token:
                continue

            target = self.model.canonical(token)

            if target:
                c.add(
                    sheet,
                    target,
                    "DIRECT_FORMULA",
                    address,
                    formula,
                    "Direct worksheet reference"
                )

        # INDIRECT
        indirects = list(
            function_calls(formula, "INDIRECT")
        )

        c.stats["INDIRECT Calls"] += len(indirects)

        for expression in indirects:
            args = function_args(expression)

            if not args:
                continue

            resolved = self.evaluator.evaluate(
                args[0], sheet, address
            )

            target = None

            if resolved is not None:
                token = sheet_token(resolved)

                if token:
                    target = self.model.canonical(token)

                elif re.match(
                    r"^\$?[A-Z]{1,3}\$?\d+",
                    str(resolved),
                    re.I
                ):
                    target = sheet

            if target:
                c.add(
                    sheet,
                    target,
                    "INDIRECT",
                    address,
                    formula,
                    "Resolved INDIRECT target"
                )

            else:
                reason = (
                    "Cannot evaluate INDIRECT argument"
                    if resolved is None
                    else "Target not resolved: " + safe(resolved)
                )

                c.unresolved_add(
                    sheet,
                    address,
                    "INDIRECT",
                    formula,
                    reason
                )

                # Candidate targets from quoted literals
                # are possible, never confirmed.
                for name in self.model.sheets:
                    if re.search(
                        re.escape(name) + r"!",
                        expression,
                        re.I
                    ):
                        c.add(
                            sheet,
                            name,
                            "INDIRECT_CANDIDATE",
                            address,
                            formula,
                            "Candidate only",
                            "Possible"
                        )

        # OFFSET
        offsets = list(
            function_calls(formula, "OFFSET")
        )

        c.stats["OFFSET Calls"] += len(offsets)

        for expression in offsets:
            c.unresolved_add(
                sheet,
                address,
                "OFFSET",
                formula,
                "OFFSET destination not fully evaluated"
            )

        # HYPERLINK
        for expression in function_calls(
            formula, "HYPERLINK"
        ):
            args = function_args(expression)

            if not args:
                continue

            destination = self.evaluator.evaluate(
                args[0], sheet, address
            )

            if destination is not None:
                target = sheet_token(destination)

                if target:
                    c.add(
                        sheet,
                        target,
                        "HYPERLINK_FORMULA",
                        address,
                        formula,
                        "Formula navigation",
                        relation="Navigation"
                    )


# ============================================================
# XLSM XML worksheet / drawing discovery
# ============================================================

def workbook_parts(z):
    root = ET.fromstring(
        z.read("xl/workbook.xml")
    )

    rels = read_relationships(
        z, "xl/workbook.xml"
    )

    result = []

    for sheet in root.findall(
        "m:sheets/m:sheet", NS
    ):
        rid = sheet.get(
            "{" + OFFICE_REL + "}id", ""
        )

        rel = rels.get(rid)

        if not rel:
            continue

        result.append({
            "name": sheet.get("name", ""),
            "state": sheet.get("state", "visible"),
            "part": normalize_path(
                "xl/workbook.xml",
                rel["target"]
            )
        })

    return result


# ============================================================
# Shape / hyperlink / button analyzer
# ============================================================

class ShapeAnalyzer:

    def __init__(self, source, model, collector):
        self.source = source
        self.model = model
        self.collector = collector

    def record(
        self,
        sheet,
        object_type,
        object_id,
        name,
        link,
        macro,
        xml_part,
        status="FOUND",
        note=""
    ):
        self.collector.shapes.append((
            sheet,
            object_type,
            object_id,
            name,
            link,
            macro,
            xml_part,
            status,
            note
        ))

        if link:
            token = sheet_token(link)

            if token:
                target = self.model.canonical(token)

                if target:
                    self.collector.add(
                        sheet,
                        target,
                        "SHAPE_LINK",
                        object_id,
                        link,
                        "Direct shape hyperlink",
                        relation="Navigation"
                    )

    def scan(self):
        print(
            "[2/5] Scanning shapes and hyperlinks...",
            flush=True
        )

        with zipfile.ZipFile(self.source) as z:
            files = set(z.namelist())
            infos = workbook_parts(z)

            for info in infos:
                sheet = info["name"]
                part = info["part"]

                if part not in files:
                    continue

                root = ET.fromstring(z.read(part))
                rels = read_relationships(z, part)

                # Cell hyperlinks
                for item in root.findall(
                    ".//m:hyperlink", NS
                ):
                    rid = item.get(
                        "{" + OFFICE_REL + "}id", ""
                    )

                    relation = rels.get(rid)

                    location = item.get(
                        "location", ""
                    )

                    if relation:
                        destination = relation["target"]
                    else:
                        destination = location

                    self.record(
                        sheet,
                        "CELL_HYPERLINK",
                        item.get("ref", ""),
                        item.get("display", ""),
                        destination,
                        "",
                        part
                    )

                # Drawing and VML relationships
                for rid, rel in rels.items():
                    typ = rel["type"]

                    if typ not in (
                        "drawing",
                        "vmlDrawing",
                        "ctrlProp",
                        "control",
                    ):
                        continue

                    target = normalize_path(
                        part,
                        rel["target"]
                    )

                    if target not in files:
                        self.record(
                            sheet,
                            typ,
                            rid,
                            "",
                            "",
                            "",
                            target,
                            "MISSING",
                            "Target part missing"
                        )
                        continue

                    if target.endswith(".vml"):
                        self.scan_vml(
                            z, sheet, target
                        )

                    elif target.endswith(".xml"):
                        if typ == "drawing":
                            self.scan_drawing(
                                z, sheet, target
                            )

                        else:
                            self.scan_control(
                                z, sheet, target
                            )

    def scan_control(self, z, sheet, part):
        root = ET.fromstring(z.read(part))

        for elem in root.iter():
            attrs = elem.attrib

            macro = (
                attrs.get("macro")
                or attrs.get("fmlaMacro")
                or attrs.get("onAction")
                or ""
            )

            if macro:
                self.record(
                    sheet,
                    "CONTROL",
                    attrs.get("id", ""),
                    attrs.get("name", ""),
                    "",
                    macro,
                    part
                )

    def scan_vml(self, z, sheet, part):
        root = ET.fromstring(z.read(part))

        for shape in root.iter():
            if local(shape.tag) != "shape":
                continue

            object_id = shape.get("id", "")
            macro = ""
            link = ""
            name = ""

            for elem in shape.iter():
                tag = local(elem.tag)

                if tag in ("FmlaMacro", "Macro"):
                    macro = safe(elem.text)

                elif tag in (
                    "FmlaLink",
                    "FmlaRange",
                    "FmlaTxbx"
                ):
                    link = safe(elem.text)

                elif tag == "ClientData":
                    name = elem.get(
                        "ObjectType", ""
                    )

            self.record(
                sheet,
                "VML_SHAPE",
                object_id,
                name,
                link,
                macro,
                part
            )

    def scan_drawing(self, z, sheet, part):
        root = ET.fromstring(z.read(part))
        rels = read_relationships(z, part)

        for anchor in root:
            if local(anchor.tag) not in (
                "twoCellAnchor",
                "oneCellAnchor",
                "absoluteAnchor"
            ):
                continue

            for obj in anchor:
                if local(obj.tag) not in (
                    "sp",
                    "pic",
                    "graphicFrame",
                    "cxnSp",
                    "grpSp"
                ):
                    continue

                object_id = ""
                name = ""
                macro = obj.get("macro", "")
                links = []

                for elem in obj.iter():
                    tag = local(elem.tag)

                    if tag == "cNvPr":
                        object_id = elem.get("id", "")
                        name = elem.get("name", "")

                        if elem.get("macro"):
                            macro = elem.get("macro")

                    if elem.get("macro"):
                        macro = elem.get("macro")

                    if tag in (
                        "hlinkClick",
                        "hlinkHover"
                    ):
                        rid = elem.get(
                            "{" + OFFICE_REL + "}id", ""
                        )

                        action = elem.get(
                            "action", ""
                        )

                        if action:
                            links.append(action)

                        if rid in rels:
                            rel = rels[rid]

                            if rel["mode"] == "External":
                                links.append(rel["target"])
                            else:
                                links.append(
                                    normalize_path(
                                        part,
                                        rel["target"]
                                    )
                                )

                self.record(
                    sheet,
                    local(obj.tag),
                    object_id,
                    name,
                    "; ".join(links),
                    macro,
                    part
                )


# ============================================================
# VBA analyzer
# ============================================================

VBA_PROC = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
    r"(?:Sub|Function)\s+([A-Za-z_]\w*)",
    re.I | re.M
)

VBA_SHEET = re.compile(
    r"(?:ThisWorkbook\.)?"
    r"(?:Worksheets|Sheets)\s*\(\s*"
    r'(?:"([^"]+)"|\'([^\']+)\')\s*\)',
    re.I
)

VBA_NAV = re.compile(
    r"\.\s*(?:Activate|Select)\b",
    re.I
)

VBA_CALL = re.compile(
    r"^\s*(?:Call\s+)?([A-Za-z_]\w*)\b",
    re.I | re.M
)


class VBAAnalyzer:

    def __init__(self, source, model, collector):
        self.source = source
        self.model = model
        self.collector = collector

        self.procedures = {}
        self.module_procs = collections.defaultdict(list)

    def scan(self):
        print("[3/5] Scanning VBA...", flush=True)

        try:
            from oletools.olevba import VBA_Parser
        except ImportError:
            self.collector.unresolved_add(
                "(Workbook)",
                "VBA",
                "VBA",
                "",
                "oletools not installed"
            )
            return

        parser = None

        try:
            parser = VBA_Parser(
                str(self.source)
            )

            if not parser.detect_vba_macros():
                return

            for _, stream, module, code in parser.extract_macros():
                if isinstance(code, bytes):
                    code = code.decode(
                        "utf-8",
                        errors="replace"
                    )

                matches = list(
                    VBA_PROC.finditer(code)
                )

                for i, match in enumerate(matches):
                    start = match.start()

                    end = (
                        matches[i + 1].start()
                        if i + 1 < len(matches)
                        else len(code)
                    )

                    name = match.group(1)
                    body = code[start:end]

                    identifier = (
                        key(module),
                        key(name)
                    )

                    self.procedures[identifier] = {
                        "module": module,
                        "name": name,
                        "body": body
                    }

                    self.module_procs[key(name)].append(
                        identifier
                    )

            self.resolve_macro_bindings()

        except Exception as exc:
            self.collector.unresolved_add(
                "(Workbook)",
                "VBA",
                "VBA",
                "",
                str(exc)
            )

        finally:
            if parser is not None:
                parser.close()

    def macro_name(self, value):
        value = str(value).strip()

        if value.startswith("="):
            value = value[1:]

        if "!" in value:
            value = value.rsplit("!", 1)[-1]

        value = value.strip("'\" ")

        return value

    def get_routes(self, macro):
        macro = self.macro_name(macro)

        if "." in macro:
            module, proc = macro.rsplit(".", 1)

            identifier = (
                key(module),
                key(proc)
            )

            roots = (
                [identifier]
                if identifier in self.procedures
                else []
            )
        else:
            roots = self.module_procs.get(
                key(macro), []
            )

        if not roots:
            return set(), False, "Macro not found"

        targets = set()
        navigation = False
        visited = set()

        def walk(identifier, depth=0):
            nonlocal navigation

            if depth > 8 or identifier in visited:
                return

            visited.add(identifier)

            proc = self.procedures.get(identifier)

            if not proc:
                return

            body = proc["body"]

            if VBA_NAV.search(body):
                navigation = True

            for match in VBA_SHEET.finditer(body):
                name = (
                    match.group(1)
                    or match.group(2)
                )

                target = self.model.canonical(name)

                if target:
                    targets.add(target)

            for match in VBA_CALL.finditer(body):
                called = key(match.group(1))

                for next_id in self.module_procs.get(
                    called, []
                ):
                    walk(next_id, depth + 1)

        for root in roots:
            walk(root)

        return targets, navigation, ""

    def resolve_macro_bindings(self):
        for row in self.collector.shapes:
            (
                sheet,
                obj_type,
                object_id,
                name,
                link,
                macro,
                part,
                status,
                note
            ) = row

            if not macro:
                continue

            targets, navigation, reason = (
                self.get_routes(macro)
            )

            if not targets:
                self.collector.macro_routes.append((
                    sheet,
                    object_id,
                    name,
                    macro,
                    "",
                    "Unresolved",
                    reason or "No literal sheet target"
                ))

                continue

            for target in sorted(targets):
                # VBA source reference is a candidate,
                # not proof of runtime navigation.
                confidence = "Possible"

                self.collector.add(
                    sheet,
                    target,
                    "VBA_MACRO_ROUTE",
                    object_id,
                    macro,
                    "Referenced by assigned macro",
                    confidence,
                    "Navigation"
                )

                self.collector.macro_routes.append((
                    sheet,
                    object_id,
                    name,
                    macro,
                    target,
                    confidence,
                    (
                        "Contains Activate/Select"
                        if navigation
                        else "Referenced by VBA"
                    )
                ))


# ============================================================
# Report
# ============================================================

def report_sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)

    ws.append(headers)

    for row in rows:
        values = []

        for value in row:
            if isinstance(value, str):
                value = value[:MAX_TEXT]

            values.append(value)

        ws.append(values)

        # Prevent source text beginning with "="
        # from becoming an executable Excel formula.
        for cell in ws[ws.max_row]:
            if (
                isinstance(cell.value, str)
                and cell.value.startswith(
                    ("=", "+", "-", "@")
                )
            ):
                cell.data_type = "s"

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid",
            fgColor="17365D"
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True
        )

    for index, header in enumerate(headers, 1):
        width = (
            65
            if header in (
                "Expression",
                "Formula",
                "Detail",
                "Reason",
                "Link",
                "Macro"
            )
            else 28
        )

        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    return ws


def create_matrix(
    wb,
    name,
    sheets,
    confirmed,
    possible
):
    ws = wb.create_sheet(name)

    ws.cell(1, 1, "Source / Target")

    for j, target in enumerate(sheets, 2):
        ws.cell(1, j, target)

    for i, source in enumerate(sheets, 2):
        ws.cell(i, 1, source)

        for j, target in enumerate(sheets, 2):
            pair = (source, target)

            if pair in confirmed:
                value = 1
                color = "C6EFCE"

            elif pair in possible:
                value = "?"
                color = "FFEB9C"

            else:
                value = 0
                color = "FFFFFF"

            cell = ws.cell(i, j, value)

            cell.fill = PatternFill(
                "solid",
                fgColor=color
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid",
            fgColor="17365D"
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True
        )

        cell.alignment = Alignment(
            textRotation=45,
            horizontal="center",
            vertical="bottom"
        )

    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 36
    ws.row_dimensions[1].height = 170

    for j in range(2, len(sheets) + 2):
        ws.column_dimensions[
            get_column_letter(j)
        ].width = 15

    return ws


def create_report(model, collector, output):
    print("[4/5] Creating report...", flush=True)

    collector.finalize()

    wb = Workbook()
    wb.remove(wb.active)

    combined_confirmed = (
        collector.data
        | collector.navigation
    )

    combined_possible = (
        collector.possible_data
        | collector.possible_navigation
    ) - combined_confirmed

    summary = [
        ("Source", str(model.source)),
        ("Total Sheets", len(model.sheets)),
        (
            "Hidden Sheets",
            sum(
                model.formula_wb[s].sheet_state
                != "visible"
                for s in model.sheets
            )
        ),
        (
            "Formula Cells",
            collector.stats["Formula Cells"]
        ),
        (
            "INDIRECT Calls",
            collector.stats["INDIRECT Calls"]
        ),
        (
            "OFFSET Calls",
            collector.stats["OFFSET Calls"]
        ),
        (
            "Confirmed Data Links",
            len(collector.data)
        ),
        (
            "Possible Data Links",
            len(collector.possible_data)
        ),
        (
            "Confirmed Navigation Links",
            len(collector.navigation)
        ),
        (
            "Possible Navigation Links",
            len(collector.possible_navigation)
        ),
        (
            "Shape Records",
            len(collector.shapes)
        ),
        (
            "Macro Routes",
            len(collector.macro_routes)
        ),
        (
            "Evidence Rows",
            len(collector.evidence)
        ),
        (
            "Unresolved Rows",
            len(collector.unresolved)
        ),
        (
            "Unresolved INDIRECT",
            sum(
                r[2] == "INDIRECT"
                for r in collector.unresolved
            )
        )
    ]

    report_sheet(
        wb,
        "Summary",
        ["Metric", "Value"],
        summary
    )

    report_sheet(
        wb,
        "Sheets",
        ["Sheet", "Visibility"],
        [
            (
                s,
                model.formula_wb[s].sheet_state
            )
            for s in model.sheets
        ]
    )

    create_matrix(
        wb,
        "Data_Matrix",
        model.sheets,
        collector.data,
        collector.possible_data
    )

    create_matrix(
        wb,
        "Navigation_Matrix",
        model.sheets,
        collector.navigation,
        collector.possible_navigation
    )

    create_matrix(
        wb,
        "Combined_Matrix",
        model.sheets,
        combined_confirmed,
        combined_possible
    )

    report_sheet(
        wb,
        "Evidence",
        [
            "Source",
            "Target",
            "Category",
            "Location",
            "Expression",
            "Detail",
            "Confidence",
            "Relation"
        ],
        collector.evidence
    )

    report_sheet(
        wb,
        "Unresolved",
        [
            "Sheet",
            "Cell",
            "Category",
            "Expression",
            "Reason"
        ],
        collector.unresolved
    )

    report_sheet(
        wb,
        "Shapes_Links",
        [
            "Sheet",
            "Object Type",
            "Object ID",
            "Name",
            "Link",
            "Macro",
            "XML Part",
            "Status",
            "Note"
        ],
        collector.shapes
    )

    report_sheet(
        wb,
        "Macro_Routes",
        [
            "Source Sheet",
            "Object ID",
            "Shape Name",
            "Macro",
            "Target Sheet",
            "Confidence",
            "Reason"
        ],
        collector.macro_routes
    )

    # Regression validation
    source = model.canonical("サマリ")
    target = model.canonical(
        "1023951_I_STAR現行売上"
    )

    if source and target:
        pair = (source, target)

        if pair in collector.data:
            result = "PASS"
        elif pair in collector.possible_data:
            result = "POSSIBLE"
        else:
            result = "NOT_RESOLVED"
    else:
        result = "SHEET_NOT_FOUND"

    top = model.canonical("TOP")

    top_shapes = [
        r for r in collector.shapes
        if top and key(r[0]) == key(top)
    ]

    top_routes = [
        r for r in collector.macro_routes
        if top and key(r[0]) == key(top)
    ]

    report_sheet(
        wb,
        "Regression_Check",
        ["Check", "Result", "Details"],
        [
            (
                "Summary -> STAR",
                result,
                "サマリ -> 1023951_I_STAR現行売上"
            ),
            (
                "TOP Sheet",
                "FOUND" if top else "NOT_FOUND",
                top or ""
            ),
            (
                "TOP Shape Count",
                len(top_shapes),
                "Drawing / VML records"
            ),
            (
                "TOP Macro Routes",
                len(top_routes),
                "Macro binding route records"
            ),
            (
                "Navigation Confirmed",
                len(collector.navigation),
                "Confirmed direct navigation"
            ),
            (
                "Navigation Possible",
                len(collector.possible_navigation),
                "VBA candidate routes"
            ),
            (
                "Unresolved INDIRECT",
                sum(
                    r[2] == "INDIRECT"
                    for r in collector.unresolved
                ),
                "Dynamic formula diagnostics"
            )
        ]
    )

    report_sheet(
        wb,
        "README",
        ["Topic", "Explanation"],
        [
            (
                "Data_Matrix",
                "Formula dependencies only"
            ),
            (
                "Navigation_Matrix",
                "Shape hyperlinks and VBA candidate routes"
            ),
            (
                "Combined_Matrix",
                "Union of data and navigation relations"
            ),
            (
                "1",
                "Confirmed relationship"
            ),
            (
                "?",
                "Possible relationship"
            ),
            (
                "0",
                "No confirmed or candidate evidence"
            ),
            (
                "Important",
                "0 does not prove absence of runtime links"
            ),
            (
                "VBA",
                "Macro references are candidates, not "
                "automatically confirmed destinations"
            ),
            (
                "Original",
                "Source XLSM is not saved"
            )
        ]
    )

    output.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    wb.save(output)

    print("[5/5] Report saved.", flush=True)

    return result


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="XLSM Dependency Analyzer V4"
    )

    parser.add_argument(
        "source",
        type=Path,
        help="Input XLSM workbook"
    )

    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None
    )

    args = parser.parse_args()

    source = args.source.resolve()

    if not source.is_file():
        parser.error(
            "Source file not found: " + str(source)
        )

    if not zipfile.is_zipfile(source):
        parser.error(
            "Source is not a valid XLSM/XLSX archive"
        )

    output = (
        args.output.resolve()
        if args.output
        else source.with_name(
            source.stem
            + "_dependency_report_v4.xlsx"
        )
    )

    if source == output:
        parser.error(
            "Output must not overwrite source"
        )

    print("=" * 65)
    print("XLSM DEPENDENCY ANALYZER V4")
    print("=" * 65)
    print("Source:", source)
    print("Output:", output)

    model = WorkbookModel(source)

    try:
        collector = Collector(model)

        FormulaAnalyzer(
            model, collector
        ).scan()

        ShapeAnalyzer(
            source, model, collector
        ).scan()

        VBAAnalyzer(
            source, model, collector
        ).scan()

        result = create_report(
            model,
            collector,
            output
        )

        print()
        print("=" * 65)
        print("ANALYSIS COMPLETE")
        print("=" * 65)

        print(
            "Total Sheets:",
            len(model.sheets)
        )

        print(
            "Formula Cells:",
            collector.stats["Formula Cells"]
        )

        print(
            "Confirmed Data Links:",
            len(collector.data)
        )

        print(
            "Confirmed Navigation Links:",
            len(collector.navigation)
        )

        print(
            "Possible Navigation Links:",
            len(collector.possible_navigation)
        )

        print(
            "Shape Records:",
            len(collector.shapes)
        )

        print(
            "Macro Routes:",
            len(collector.macro_routes)
        )

        print(
            "Unresolved INDIRECT:",
            sum(
                r[2] == "INDIRECT"
                for r in collector.unresolved
            )
        )

        print()
        print(
            "Regression Summary -> STAR:",
            result
        )

        print("Output:", output)

    finally:
        model.close()


if __name__ == "__main__":
    try:
        main()

    except Exception:
        traceback.print_exc()
        sys.exit(1)
