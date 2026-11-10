
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Dependency Analyzer V5
============================

Features:
- Formula dependencies
- INDIRECT dynamic sheet references
- DrawingML shapes and picture hyperlinks
- VML buttons and controls
- VBA macro route candidates
- Hidden / veryHidden sheets
- Data, navigation and combined matrices
- 1 / 0 / ? relationship states
- Streaming XLSX report with progress logging
- Original XLSM remains unchanged

Install:
    pip install openpyxl oletools

Run:
    python -u analyze_xlsm_v5.py input.xlsm

Optional:
    python -u analyze_xlsm_v5.py input.xlsm -o report.xlsx
"""

import argparse
import collections
import posixpath
import re
import sys
import time
import traceback
import zipfile
import xml.etree.ElementTree as ET

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill


MAIN = (
    "http://schemas.openxmlformats.org/"
    "spreadsheetml/2006/main"
)

RNS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

CELL = re.compile(
    r"^\$?[A-Z]{1,3}\$?\d+$",
    re.I
)

REF = re.compile(
    r"(?<![\w.])"
    r"(?:'((?:[^']|'')+)'|"
    r"([\w.\u0080-\uffff]+))!"
    r"(?=\$?[A-Z]{1,3}\$?\d+"
    r"|\$?[A-Z]{1,3}\s*:"
    r"|\$?\d+\s*:)",
    re.I
)

VBA_SHEET = re.compile(
    r'(?:ThisWorkbook\s*\.\s*)?'
    r'(?:Sheets|Worksheets)\s*\(\s*"([^"]+)"\s*\)',
    re.I
)

VBA_PROC = re.compile(
    r"^\s*"
    r"(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function)\s+"
    r"([\w\u0080-\uffff]+)",
    re.I | re.M
)

VBA_CALL = re.compile(
    r"^\s*(?:Call\s+)([\w\u0080-\uffff]+)\b",
    re.I | re.M
)


# ============================================================
# Utilities
# ============================================================

def loc(tag):
    return tag.rsplit("}", 1)[-1]


def txt(value):
    return "" if value is None else str(value)


def pathjoin(base, target):
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(base),
            target
        )
    )


def rels(z, part):
    p = posixpath.join(
        posixpath.dirname(part),
        "_rels",
        posixpath.basename(part) + ".rels"
    )

    if p not in z.namelist():
        return {}

    result = {}

    for x in ET.fromstring(z.read(p)):
        result[x.get("Id")] = (
            x.get("Type", "").rsplit("/", 1)[-1],
            x.get("Target", ""),
            x.get("TargetMode", "")
        )

    return result


def calls(formula, name):
    pattern = re.compile(
        r"\b" + re.escape(name) + r"\s*\(",
        re.I
    )

    i = 0
    quoted = False

    while i < len(formula):
        if formula[i] == '"':
            if (
                quoted
                and i + 1 < len(formula)
                and formula[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted
            i += 1
            continue

        match = (
            pattern.match(formula, i)
            if not quoted
            else None
        )

        if not match:
            i += 1
            continue

        start = match.end()
        j = start
        depth = 1
        q = False

        while j < len(formula):
            ch = formula[j]

            if ch == '"':
                if (
                    q
                    and j + 1 < len(formula)
                    and formula[j + 1] == '"'
                ):
                    j += 2
                    continue

                q = not q

            elif not q:
                if ch == "(":
                    depth += 1

                elif ch == ")":
                    depth -= 1

                    if depth == 0:
                        yield formula[start:j]
                        break

            j += 1

        i = j + 1


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

            elif depth == 0 and ch in separators:
                result.append("".join(buffer).strip())
                buffer = []
                i += 1
                continue

        buffer.append(ch)
        i += 1

    result.append("".join(buffer).strip())

    return result


def mask(formula):
    return re.sub(
        r'"(?:[^"]|"")*"',
        lambda m: " " * len(m.group()),
        formula
    )


# ============================================================
# Main analyzer
# ============================================================

class Audit:

    def __init__(self, source):
        self.src = Path(source)

        self.wbf = load_workbook(
            source,
            read_only=True,
            data_only=False,
            keep_links=False
        )

        self.wbv = load_workbook(
            source,
            read_only=True,
            data_only=True,
            keep_links=False
        )

        self.names = list(self.wbf.sheetnames)

        self.lookup = {
            name.casefold(): name
            for name in self.names
        }

        self.data = set()
        self.nav = set()

        self.possible_data = set()
        self.possible_nav = set()

        self.evidence = []
        self.unresolved = []
        self.shapes = []
        self.macros = []
        self.vba_rows = []
        self.external = []

        self.stats = collections.Counter()
        self.seen = set()

        self.procs = collections.defaultdict(list)

        self.cache = {}
        self.eval_active = set()

    def canonical(self, name):
        if name is None:
            return None

        return self.lookup.get(
            txt(name)
            .replace("''", "'")
            .casefold()
            .strip()
        )

    def add(
        self,
        source,
        target,
        kind,
        category,
        location,
        expression,
        detail="",
        confidence="Confirmed"
    ):
        source = self.canonical(source)
        target = self.canonical(target)

        if not source or not target:
            return

        pair = (source, target)

        if kind == "DATA":
            if confidence == "Confirmed":
                self.data.add(pair)
            else:
                self.possible_data.add(pair)

        else:
            if confidence == "Confirmed":
                self.nav.add(pair)
            else:
                self.possible_nav.add(pair)

        signature = (
            source,
            target,
            kind,
            category,
            location,
            confidence
        )

        if signature in self.seen:
            return

        self.seen.add(signature)

        self.evidence.append((
            source,
            target,
            kind,
            confidence,
            category,
            location,
            txt(expression)[:12000],
            txt(detail)[:1000]
        ))

    # --------------------------------------------------------
    # Cached values and dynamic expressions
    # --------------------------------------------------------

    def value(self, sheet, address, depth=0):
        sheet = self.canonical(sheet)
        address = address.replace("$", "").upper()

        if not sheet or depth > 12:
            return None

        cache_key = (sheet, address)

        if cache_key in self.cache:
            return self.cache[cache_key]

        if cache_key in self.eval_active:
            return None

        self.eval_active.add(cache_key)

        try:
            value = self.wbv[sheet][address].value

            if value is None:
                value = self.wbf[sheet][address].value

                if (
                    isinstance(value, str)
                    and value.startswith("=")
                ):
                    value = self.evaluate(
                        value[1:],
                        sheet,
                        address,
                        depth + 1
                    )

            if len(self.cache) < 50000:
                self.cache[cache_key] = value

            return value

        except Exception:
            return None

        finally:
            self.eval_active.discard(cache_key)

    def evaluate(
        self,
        expression,
        sheet,
        address,
        depth=0
    ):
        if depth > 12:
            return None

        expr = expression.strip()

        if (
            expr.startswith("(")
            and expr.endswith(")")
        ):
            inner = expr[1:-1]

            if inner.count("(") == inner.count(")"):
                return self.evaluate(
                    inner,
                    sheet,
                    address,
                    depth + 1
                )

        # Concatenation
        parts = split_top(expr, "&")

        if len(parts) > 1:
            values = [
                self.evaluate(
                    part,
                    sheet,
                    address,
                    depth + 1
                )
                for part in parts
            ]

            if any(v is None for v in values):
                return None

            return "".join(map(str, values))

        # Literal string
        if (
            len(expr) >= 2
            and expr[0] == '"'
            and expr[-1] == '"'
        ):
            return expr[1:-1].replace('""', '"')

        # Number
        if re.fullmatch(
            r"-?\d+(?:\.\d+)?",
            expr
        ):
            if "." in expr:
                return float(expr)

            return int(expr)

        # Cell reference
        match = re.fullmatch(
            r"(?:(?:'((?:[^']|'')+)'|"
            r"([\w.\u0080-\uffff]+))!)?"
            r"(\$?[A-Z]{1,3}\$?\d+)",
            expr,
            re.I
        )

        if match:
            target_sheet = (
                match.group(1)
                or match.group(2)
                or sheet
            )

            return self.value(
                target_sheet,
                match.group(3),
                depth + 1
            )

        # ROW() / COLUMN()
        match = re.fullmatch(
            r"(ROW|COLUMN)\s*\("
            r"\s*(\$?[A-Z]{1,3}\$?\d+)?\s*\)",
            expr,
            re.I
        )

        if match:
            reference = (
                match.group(2)
                or address
            ).replace("$", "")

            coordinate = re.match(
                r"([A-Z]+)(\d+)",
                reference,
                re.I
            )

            if coordinate:
                if match.group(1).upper() == "ROW":
                    return int(coordinate.group(2))

                column = 0

                for ch in coordinate.group(1).upper():
                    column = column * 26 + ord(ch) - 64

                return column

        # Supported text functions
        functions = (
            "SUBSTITUTE",
            "TRIM",
            "UPPER",
            "LOWER",
            "CONCATENATE",
            "IFERROR",
            "IF",
            "TEXT"
        )

        for function in functions:
            if not re.fullmatch(
                r"(?is)"
                + function
                + r"\s*\(.*\)",
                expr
            ):
                continue

            found = list(calls(expr, function))

            if not found:
                continue

            args = split_top(found[0], ",;")

            if function == "IFERROR" and len(args) > 1:
                value = self.evaluate(
                    args[0],
                    sheet,
                    address,
                    depth + 1
                )

                if value is not None:
                    return value

                return self.evaluate(
                    args[1],
                    sheet,
                    address,
                    depth + 1
                )

            if function == "IF" and len(args) > 2:
                condition = self.evaluate(
                    args[0],
                    sheet,
                    address,
                    depth + 1
                )

                if isinstance(
                    condition,
                    (int, float, bool)
                ):
                    selected = (
                        args[1]
                        if condition
                        else args[2]
                    )

                    return self.evaluate(
                        selected,
                        sheet,
                        address,
                        depth + 1
                    )

            values = [
                self.evaluate(
                    arg,
                    sheet,
                    address,
                    depth + 1
                )
                for arg in args
            ]

            if (
                function == "SUBSTITUTE"
                and len(values) >= 3
                and all(
                    value is not None
                    for value in values[:3]
                )
            ):
                return str(values[0]).replace(
                    str(values[1]),
                    str(values[2])
                )

            if (
                function == "TRIM"
                and values
                and values[0] is not None
            ):
                return " ".join(
                    str(values[0]).split()
                )

            if (
                function == "UPPER"
                and values
                and values[0] is not None
            ):
                return str(values[0]).upper()

            if (
                function == "LOWER"
                and values
                and values[0] is not None
            ):
                return str(values[0]).lower()

            if (
                function == "CONCATENATE"
                and all(
                    value is not None
                    for value in values
                )
            ):
                return "".join(map(str, values))

            if (
                function == "TEXT"
                and values
                and values[0] is not None
            ):
                return str(values[0])

        return None

    def target(self, link):
        value = (
            txt(link)
            .strip()
            .lstrip("#")
            .lstrip("=")
        )

        match = re.search(
            r"(?:'((?:[^']|'')+)'|"
            r"([\w.\u0080-\uffff]+))!",
            value
        )

        if not match:
            return None

        return self.canonical(
            match.group(1)
            or match.group(2)
        )

    # --------------------------------------------------------
    # Formula scan
    # --------------------------------------------------------

    def scan_formulas(self):
        print(
            "[1/5] Formula scan",
            flush=True
        )

        for ws in self.wbf.worksheets:
            count = 0

            for row in ws.iter_rows():
                for cell in row:
                    formula = cell.value

                    if not (
                        isinstance(formula, str)
                        and formula.startswith("=")
                    ):
                        continue

                    count += 1
                    self.stats["Formulas"] += 1

                    address = cell.coordinate

                    # Direct sheet references
                    for match in REF.finditer(
                        mask(formula)
                    ):
                        name = (
                            match.group(1)
                            or match.group(2)
                        )

                        if "[" in name or "]" in name:
                            self.external.append((
                                ws.title,
                                address,
                                name,
                                formula
                            ))
                            continue

                        target = self.canonical(name)

                        if target:
                            self.add(
                                ws.title,
                                target,
                                "DATA",
                                "DIRECT",
                                address,
                                formula
                            )

                    # INDIRECT
                    for expression in calls(
                        formula,
                        "INDIRECT"
                    ):
                        self.stats["INDIRECT"] += 1

                        args = split_top(
                            expression,
                            ",;"
                        )

                        value = (
                            self.evaluate(
                                args[0],
                                ws.title,
                                address
                            )
                            if args
                            else None
                        )

                        target = self.target(value)

                        if target:
                            self.add(
                                ws.title,
                                target,
                                "DATA",
                                "INDIRECT",
                                address,
                                formula,
                                "Resolved: " + txt(value)
                            )

                        elif (
                            value is not None
                            and CELL.fullmatch(txt(value))
                        ):
                            self.add(
                                ws.title,
                                ws.title,
                                "DATA",
                                "INDIRECT_LOCAL",
                                address,
                                formula
                            )

                        else:
                            self.unresolved.append((
                                ws.title,
                                address,
                                "INDIRECT",
                                formula,
                                "Argument not resolved; value="
                                + txt(value)[:200]
                            ))

                            # Literal candidates only.
                            for name in self.names:
                                if (
                                    "'" + name + "'!"
                                    in expression
                                    or
                                    '"' + name + "!"
                                    in expression
                                ):
                                    self.add(
                                        ws.title,
                                        name,
                                        "DATA",
                                        "INDIRECT_CANDIDATE",
                                        address,
                                        formula,
                                        "Literal candidate",
                                        "Possible"
                                    )

                    # OFFSET
                    for expression in calls(
                        formula,
                        "OFFSET"
                    ):
                        self.stats["OFFSET"] += 1

                        if not REF.search(
                            mask(expression)
                        ):
                            self.unresolved.append((
                                ws.title,
                                address,
                                "OFFSET",
                                formula,
                                "Dynamic offset not resolved"
                            ))

                    # HYPERLINK
                    for expression in calls(
                        formula,
                        "HYPERLINK"
                    ):
                        args = split_top(
                            expression,
                            ",;"
                        )

                        value = (
                            self.evaluate(
                                args[0],
                                ws.title,
                                address
                            )
                            if args
                            else None
                        )

                        target = self.target(value)

                        if target:
                            self.add(
                                ws.title,
                                target,
                                "NAV",
                                "HYPERLINK_FORMULA",
                                address,
                                formula
                            )

            print(
                f"  {ws.title}: {count:,} formulas",
                flush=True
            )

    # --------------------------------------------------------
    # Shape records
    # --------------------------------------------------------

    def record_shape(
        self,
        sheet,
        object_type,
        object_id,
        name,
        link,
        macro,
        part,
        detail=""
    ):
        self.shapes.append((
            sheet,
            object_type,
            txt(object_id),
            txt(name),
            txt(link),
            txt(macro),
            part,
            txt(detail)
        ))

        target = self.target(link)

        if target:
            self.add(
                sheet,
                target,
                "NAV",
                "SHAPE_LINK",
                txt(object_id),
                link,
                detail
            )

    # --------------------------------------------------------
    # Drawing / VML / controls
    # --------------------------------------------------------

    def scan_drawings(self):
        print(
            "[2/5] Drawing / VML / controls",
            flush=True
        )

        with zipfile.ZipFile(self.src) as z:
            files = set(z.namelist())

            workbook_rels = rels(
                z,
                "xl/workbook.xml"
            )

            root = ET.fromstring(
                z.read("xl/workbook.xml")
            )

            for element in root.iter():
                if loc(element.tag) != "sheet":
                    continue

                rid = element.get(
                    "{" + RNS + "}id"
                )

                relation = workbook_rels.get(rid)

                if not relation:
                    continue

                sheet = element.get("name", "")

                part = pathjoin(
                    "xl/workbook.xml",
                    relation[1]
                )

                if part not in files:
                    continue

                sheet_rels = rels(z, part)

                sheetroot = ET.fromstring(
                    z.read(part)
                )

                handled = set()

                # Worksheet hyperlinks and controls
                for item in sheetroot.iter():
                    tag = loc(item.tag)

                    if tag == "hyperlink":
                        rid = item.get(
                            "{" + RNS + "}id"
                        )

                        relation = sheet_rels.get(rid)

                        link = (
                            relation[1]
                            if relation
                            else item.get(
                                "location",
                                ""
                            )
                        )

                        self.record_shape(
                            sheet,
                            "CellHyperlink",
                            item.get("ref", ""),
                            item.get("display", ""),
                            link,
                            "",
                            part
                        )

                    if tag in (
                        "control",
                        "oleObject"
                    ):
                        macro = item.get(
                            "macro",
                            ""
                        )

                        object_id = (
                            item.get("shapeId")
                            or item.get("name")
                            or item.get("id")
                            or ""
                        )

                        if macro:
                            self.record_shape(
                                sheet,
                                tag,
                                object_id,
                                item.get("name", ""),
                                "",
                                macro,
                                part
                            )

                # Relationships to drawings and controls
                for rid, relation in sheet_rels.items():
                    typ, target, mode = relation

                    if typ not in (
                        "drawing",
                        "vmlDrawing",
                        "ctrlProp",
                        "control"
                    ):
                        continue

                    target_part = pathjoin(
                        part,
                        target
                    )

                    if target_part in handled:
                        continue

                    handled.add(target_part)

                    if target_part not in files:
                        self.record_shape(
                            sheet,
                            typ,
                            rid,
                            "",
                            target,
                            "",
                            target_part,
                            "Missing target"
                        )
                        continue

                    if target_part.endswith(".vml"):
                        self.parse_vml(
                            z,
                            sheet,
                            target_part
                        )

                    elif (
                        typ == "drawing"
                        and target_part.endswith(".xml")
                    ):
                        self.parse_drawing(
                            z,
                            sheet,
                            target_part
                        )

                    elif target_part.endswith(".xml"):
                        self.parse_control(
                            z,
                            sheet,
                            target_part
                        )

    def parse_vml(self, z, sheet, part):
        root = ET.fromstring(z.read(part))
        drawing_rels = rels(z, part)

        for shape in root.iter():
            if loc(shape.tag) != "shape":
                continue

            object_id = shape.get("id", "")
            macro = ""
            links = []
            kind = "VML"
            detail = []

            for element in shape.iter():
                tag = loc(element.tag)

                if tag in (
                    "FmlaMacro",
                    "Macro"
                ):
                    macro = txt(element.text)

                elif tag in (
                    "FmlaLink",
                    "FmlaRange",
                    "FmlaTxbx"
                ):
                    detail.append(
                        tag + ":" + txt(element.text)
                    )

                elif tag == "ClientData":
                    kind = (
                        "VML:"
                        + element.get(
                            "ObjectType",
                            ""
                        )
                    )

                elif tag in (
                    "hlinkClick",
                    "hlinkHover"
                ):
                    rid = element.get(
                        "{" + RNS + "}id"
                    )

                    relation = drawing_rels.get(rid)

                    if relation:
                        links.append(relation[1])

                    if element.get("action"):
                        links.append(
                            element.get("action")
                        )

            self.record_shape(
                sheet,
                kind,
                object_id,
                object_id,
                "; ".join(links),
                macro,
                part,
                "; ".join(detail)
            )

    def parse_control(self, z, sheet, part):
        root = ET.fromstring(z.read(part))

        for element in root.iter():
            macro = (
                element.get("macro")
                or element.get("fmlaMacro")
                or element.get("onAction")
            )

            if macro:
                self.record_shape(
                    sheet,
                    "Control",
                    element.get("id", ""),
                    element.get("name", ""),
                    "",
                    macro,
                    part
                )

    def parse_drawing(self, z, sheet, part):
        root = ET.fromstring(z.read(part))
        drawing_rels = rels(z, part)

        for anchor in root:
            if loc(anchor.tag) not in (
                "oneCellAnchor",
                "twoCellAnchor",
                "absoluteAnchor"
            ):
                continue

            for obj in anchor:
                if loc(obj.tag) not in (
                    "sp",
                    "pic",
                    "graphicFrame",
                    "cxnSp",
                    "grpSp"
                ):
                    continue

                objects = [obj]

                # Include nested shapes in groups.
                objects.extend(
                    element
                    for element in obj.iter()
                    if (
                        element is not obj
                        and loc(element.tag) in (
                            "sp",
                            "pic",
                            "graphicFrame",
                            "cxnSp"
                        )
                    )
                )

                for shape in objects:
                    object_id = ""
                    name = ""
                    macro = shape.get("macro", "")
                    links = []
                    details = []

                    for element in shape.iter():
                        tag = loc(element.tag)

                        if tag == "cNvPr":
                            object_id = element.get(
                                "id",
                                object_id
                            )

                            name = element.get(
                                "name",
                                name
                            )

                            if element.get("descr"):
                                details.append(
                                    element.get("descr")
                                )

                        if element.get("macro"):
                            macro = element.get("macro")

                        if tag in (
                            "hlinkClick",
                            "hlinkHover"
                        ):
                            rid = element.get(
                                "{" + RNS + "}id"
                            )

                            relation = drawing_rels.get(
                                rid
                            )

                            if relation:
                                if relation[2] == "External":
                                    links.append(
                                        relation[1]
                                    )
                                else:
                                    links.append(
                                        pathjoin(
                                            part,
                                            relation[1]
                                        )
                                    )

                            if element.get("action"):
                                links.append(
                                    element.get("action")
                                )

                            if element.get("tooltip"):
                                details.append(
                                    element.get("tooltip")
                                )

                    self.record_shape(
                        sheet,
                        loc(shape.tag),
                        object_id,
                        name,
                        "; ".join(
                            dict.fromkeys(links)
                        ),
                        macro,
                        part,
                        "; ".join(details)
                    )

    # --------------------------------------------------------
    # VBA macro analysis
    # --------------------------------------------------------

    def scan_vba(self):
        print(
            "[3/5] VBA macro routes",
            flush=True
        )

        try:
            from oletools.olevba import VBA_Parser

        except ImportError:
            self.unresolved.append((
                "Workbook",
                "VBA",
                "VBA",
                "",
                "oletools not installed"
            ))
            return

        parser = None

        try:
            parser = VBA_Parser(str(self.src))

            if parser.detect_vba_macros():
                for (
                    _,
                    stream,
                    module,
                    code
                ) in parser.extract_macros():

                    if isinstance(code, bytes):
                        code = code.decode(
                            "utf-8",
                            "replace"
                        )

                    matches = list(
                        VBA_PROC.finditer(code)
                    )

                    for i, match in enumerate(matches):
                        end = (
                            matches[i + 1].start()
                            if i + 1 < len(matches)
                            else len(code)
                        )

                        body = code[
                            match.start():end
                        ]

                        proc_name = match.group(1)

                        self.procs[
                            proc_name.casefold()
                        ].append((
                            module,
                            body
                        ))

                        for ref in VBA_SHEET.finditer(body):
                            self.vba_rows.append((
                                module,
                                proc_name,
                                ref.group(1),
                                body[:1000]
                            ))

        except Exception as exc:
            self.unresolved.append((
                "Workbook",
                "VBA",
                "VBA",
                "",
                str(exc)
            ))

        finally:
            if parser:
                try:
                    parser.close()
                except Exception:
                    pass

        # Connect shapes to assigned macros.
        for row in self.shapes:
            (
                sheet,
                typ,
                object_id,
                name,
                link,
                macro,
                part,
                detail
            ) = row

            if not macro:
                continue

            proc = (
                macro.strip()
                .lstrip("=")
                .split("!")[-1]
                .strip("'\" ")
                .split(".")[-1]
                .casefold()
            )

            roots = self.procs.get(proc, [])

            if not roots:
                self.macros.append((
                    sheet,
                    object_id,
                    name,
                    macro,
                    "",
                    "Unresolved",
                    "Procedure not found"
                ))
                continue

            visited = set()
            targets = set()
            has_navigation = False

            def walk(procname, depth=0):
                nonlocal has_navigation

                if (
                    depth > 8
                    or procname in visited
                ):
                    return

                visited.add(procname)

                for module, body in self.procs.get(
                    procname,
                    []
                ):
                    if re.search(
                        r"\.\s*(Activate|Select)\b"
                        r"|Application\.Goto\b",
                        body,
                        re.I
                    ):
                        has_navigation = True

                    for match in VBA_SHEET.finditer(body):
                        target = self.canonical(
                            match.group(1)
                        )

                        if target:
                            targets.add(target)

                    for match in VBA_CALL.finditer(body):
                        walk(
                            match.group(1).casefold(),
                            depth + 1
                        )

            walk(proc)

            if not targets:
                self.macros.append((
                    sheet,
                    object_id,
                    name,
                    macro,
                    "",
                    "Unresolved",
                    "No static target"
                ))
                continue

            for target in sorted(targets):
                self.add(
                    sheet,
                    target,
                    "NAV",
                    "VBA_MACRO_CANDIDATE",
                    object_id,
                    macro,
                    "Macro references target; "
                    "execution not proven",
                    "Possible"
                )

                self.macros.append((
                    sheet,
                    object_id,
                    name,
                    macro,
                    target,
                    "Possible",
                    (
                        "Activate/Select present"
                        if has_navigation
                        else "VBA sheet reference"
                    )
                ))

    # --------------------------------------------------------
    # Streaming XLSX report
    # --------------------------------------------------------

    def report(self, output):
        print(
            "[4/5] Creating streaming report...",
            flush=True
        )

        self.possible_data -= self.data
        self.possible_nav -= self.nav

        wb = Workbook(write_only=True)

        def tab(name, headers, rows):
            start = time.perf_counter()

            print(
                "  [START]",
                name,
                flush=True
            )

            ws = wb.create_sheet(name)

            header_cells = []

            for value in headers:
                cell = WriteOnlyCell(
                    ws,
                    value=value
                )

                cell.font = Font(
                    bold=True,
                    color="FFFFFF"
                )

                cell.fill = PatternFill(
                    "solid",
                    fgColor="17365D"
                )

                header_cells.append(cell)

            ws.append(header_cells)

            count = 0

            for row in rows:
                values = []

                for value in row:
                    if isinstance(value, str):
                        value = value[:32000]

                    if (
                        isinstance(value, str)
                        and value.startswith(
                            ("=", "+", "-", "@")
                        )
                    ):
                        cell = WriteOnlyCell(
                            ws,
                            value=value
                        )
                        cell.data_type = "s"
                        values.append(cell)

                    else:
                        values.append(value)

                ws.append(values)
                count += 1

                if count % 10000 == 0:
                    elapsed = (
                        time.perf_counter()
                        - start
                    )

                    print(
                        f"    {name}: "
                        f"{count:,} rows "
                        f"({elapsed:.1f}s)",
                        flush=True
                    )

            elapsed = (
                time.perf_counter()
                - start
            )

            print(
                f"  [DONE] {name}: "
                f"{count:,} rows "
                f"({elapsed:.1f}s)",
                flush=True
            )

        summary = [
            ("Source", str(self.src)),
            ("Sheets", len(self.names)),
            ("Formula cells", self.stats["Formulas"]),
            ("INDIRECT calls", self.stats["INDIRECT"]),
            ("OFFSET calls", self.stats["OFFSET"]),
            ("Data confirmed", len(self.data)),
            ("Data possible", len(self.possible_data)),
            ("Navigation confirmed", len(self.nav)),
            ("Navigation possible", len(self.possible_nav)),
            ("Shape records", len(self.shapes)),
            ("Macro routes", len(self.macros)),
            ("Unresolved", len(self.unresolved)),
            ("Evidence", len(self.evidence))
        ]

        tab(
            "Summary",
            ["Metric", "Value"],
            summary
        )

        tab(
            "Sheets",
            ["Sheet", "Visibility"],
            (
                (
                    sheet,
                    self.wbf[sheet].sheet_state
                )
                for sheet in self.names
            )
        )

        def matrix(name, confirmed, possible):
            def rows():
                for source in self.names:
                    values = [source]

                    for target in self.names:
                        pair = (source, target)

                        if pair in confirmed:
                            values.append(1)

                        elif pair in possible:
                            values.append("?")

                        else:
                            values.append(0)

                    yield values

            tab(
                name,
                ["Source / Target"] + self.names,
                rows()
            )

        matrix(
            "Data_Matrix",
            self.data,
            self.possible_data
        )

        matrix(
            "Navigation_Matrix",
            self.nav,
            self.possible_nav
        )

        combined_confirmed = (
            self.data | self.nav
        )

        combined_possible = (
            self.possible_data
            | self.possible_nav
        ) - combined_confirmed

        matrix(
            "Combined_Matrix",
            combined_confirmed,
            combined_possible
        )

        tab(
            "Evidence",
            [
                "Source",
                "Target",
                "Relation",
                "Confidence",
                "Category",
                "Cell or Object",
                "Expression",
                "Detail"
            ],
            self.evidence
        )

        tab(
            "Unresolved",
            [
                "Sheet",
                "Cell",
                "Category",
                "Formula",
                "Reason"
            ],
            self.unresolved
        )

        tab(
            "Shapes_Links",
            [
                "Sheet",
                "Object Type",
                "Object ID",
                "Name",
                "Hyperlink",
                "Macro",
                "XML Part",
                "Detail"
            ],
            self.shapes
        )

        tab(
            "Macro_Routes",
            [
                "Sheet",
                "Object ID",
                "Name",
                "Macro",
                "Target",
                "Confidence",
                "Reason"
            ],
            self.macros
        )

        tab(
            "VBA_References",
            [
                "Module",
                "Procedure",
                "Referenced Sheet",
                "Code"
            ],
            self.vba_rows
        )

        tab(
            "External_Refs",
            [
                "Sheet",
                "Cell",
                "External Token",
                "Formula"
            ],
            self.external
        )

        # Regression check
        source = self.canonical("サマリ")

        target = self.canonical(
            "1023951_I_STAR現行売上"
        )

        pair = (source, target)

        if pair in self.data:
            status = "PASS"

        elif pair in self.possible_data:
            status = "POSSIBLE"

        else:
            status = "NOT_RESOLVED"

        top = self.canonical("TOP")

        checks = [
            (
                "サマリ → 1023951_I_STAR現行売上",
                status,
                "Requires resolved data reference"
            ),
            (
                "TOP shapes",
                sum(
                    row[0] == top
                    for row in self.shapes
                ),
                "Drawing/VML objects"
            ),
            (
                "TOP macro bindings",
                sum(
                    row[0] == top and bool(row[5])
                    for row in self.shapes
                ),
                "Shape macro assignments"
            ),
            (
                "TOP confirmed navigation",
                sum(
                    source == top
                    for source, target in self.nav
                ),
                "Not necessarily TOP -> TOP"
            ),
            (
                "TOP possible navigation",
                sum(
                    source == top
                    for source, target in self.possible_nav
                ),
                "VBA candidates"
            )
        ]

        tab(
            "Regression_Check",
            [
                "Check",
                "Result",
                "Explanation"
            ],
            checks
        )

        tab(
            "README",
            ["Topic", "Meaning"],
            [
                ("1", "Confirmed relation"),
                ("?", "Candidate relation"),
                (
                    "0",
                    "No identified evidence; "
                    "not proof of absence"
                ),
                (
                    "Data_Matrix",
                    "Formula/data dependencies"
                ),
                (
                    "Navigation_Matrix",
                    "Direct hyperlinks and "
                    "candidate macro routes"
                ),
                (
                    "Macro_Routes",
                    "Static analysis only; "
                    "macros never executed"
                ),
                (
                    "Original",
                    "XLSM not modified"
                ),
                (
                    "Performance",
                    "Streaming output minimizes "
                    "report-memory usage"
                ),
                (
                    "Limitations",
                    "Cached values may be absent; "
                    "dynamic VBA and some Excel "
                    "drawing variants may be unresolved"
                )
            ]
        )

        print(
            "[5/5] Compressing/saving XLSX...",
            flush=True
        )

        start = time.perf_counter()

        wb.save(output)

        elapsed = (
            time.perf_counter()
            - start
        )

        print(
            f"[DONE] Saved in {elapsed:.1f}s: "
            f"{output}",
            flush=True
        )

        return status

    def close(self):
        self.wbf.close()
        self.wbv.close()


# ============================================================
# Entry point
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=__doc__
    )

    parser.add_argument(
        "source",
        type=Path
    )

    parser.add_argument(
        "-o",
        "--output",
        type=Path
    )

    args = parser.parse_args()

    source = args.source.resolve()

    if (
        not source.is_file()
        or not zipfile.is_zipfile(source)
    ):
        parser.error(
            "Input must be a valid XLSM/XLSX file"
        )

    output = (
        args.output
        or source.with_name(
            source.stem
            + "_dependency_report_v5.xlsx"
        )
    ).resolve()

    if source == output:
        parser.error(
            "Output cannot overwrite original"
        )

    print(
        "Source:",
        source,
        flush=True
    )

    print(
        "Output:",
        output,
        flush=True
    )

    start = time.perf_counter()

    audit = Audit(source)

    try:
        audit.scan_formulas()
        audit.scan_drawings()
        audit.scan_vba()

        output.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        status = audit.report(output)

        elapsed = (
            time.perf_counter()
            - start
        )

        print()
        print("=" * 60)
        print("FINISHED")
        print("=" * 60)

        print(
            f"Elapsed: {elapsed:.1f}s"
        )

        print(
            f"Summary -> STAR: {status}"
        )

        print(
            f"Output: {output}"
        )

    finally:
        audit.close()


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print(
            "\nInterrupted by user",
            file=sys.stderr
        )
        sys.exit(130)

    except Exception:
        traceback.print_exc()
        sys.exit(1)
