
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Dependency Analyzer V6.3

Install:
    pip install openpyxl oletools

Run:
    python -u analyze_xlsm_v6_3.py input.xlsm

Optional:
    python -u analyze_xlsm_v6_3.py input.xlsm -o report.xlsx

The original workbook is never modified.
VBA macros are inspected but never executed.
"""

import argparse
import collections
import posixpath
import re
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill


RNS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

SREF = re.compile(
    r"(?:'((?:[^']|'')+)'|"
    r"([\w.\u0080-\uffff]+))!",
    re.I
)

CELL = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|"
    r"([\w.\u0080-\uffff]+))!)?"
    r"(\$?[A-Z]{1,3}\$?\d+)$",
    re.I
)

VREF = re.compile(
    r'(?:Worksheets|Sheets)'
    r'\s*\(\s*"([^"]+)"\s*\)',
    re.I
)

PROC = re.compile(
    r"^\s*"
    r"(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function)\s+"
    r"([\w\u0080-\uffff]+)",
    re.I | re.M
)

NAV = re.compile(
    r"\.\s*(?:Activate|Select)\b"
    r"|\b(?:Application\s*\.\s*)?Goto\b",
    re.I
)


# ============================================================
# Utilities
# ============================================================

def local(tag):
    return tag.rsplit("}", 1)[-1]


def text(value):
    return "" if value is None else str(value)


def joinpart(part, target):
    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(part),
            target
        )
    )


def relationships(z, part):
    p = posixpath.join(
        posixpath.dirname(part),
        "_rels",
        posixpath.basename(part) + ".rels"
    )

    if p not in z.namelist():
        return {}

    root = ET.fromstring(z.read(p))

    return {
        x.get("Id"): (
            x.get("Type", "").rsplit("/", 1)[-1],
            x.get("Target", ""),
            x.get("TargetMode", "")
        )
        for x in root
    }


def split(expr, chars):
    out = []
    buf = []
    depth = 0
    quote = False
    i = 0

    while i < len(expr):
        ch = expr[i]

        if ch == '"':
            if (
                quote
                and i + 1 < len(expr)
                and expr[i + 1] == '"'
            ):
                buf.extend(['"', '"'])
                i += 2
                continue

            quote = not quote

        if not quote:
            if ch == "(":
                depth += 1

            elif ch == ")":
                depth -= 1

            elif depth == 0 and ch in chars:
                out.append(
                    "".join(buf).strip()
                )
                buf = []
                i += 1
                continue

        buf.append(ch)
        i += 1

    out.append(
        "".join(buf).strip()
    )

    return out


def calls(formula, name):
    pattern = re.compile(
        r"(?<![\w.])(?:_xlfn\.)?"
        + re.escape(name)
        + r"\s*\(",
        re.I
    )

    i = 0
    quote = False

    while i < len(formula):
        if formula[i] == '"':
            if (
                quote
                and i + 1 < len(formula)
                and formula[i + 1] == '"'
            ):
                i += 2
                continue

            quote = not quote
            i += 1
            continue

        match = (
            pattern.match(formula, i)
            if not quote
            else None
        )

        if not match:
            i += 1
            continue

        start = match.end()
        j = start
        depth = 1
        inside = False

        while j < len(formula):
            ch = formula[j]

            if ch == '"':
                if (
                    inside
                    and j + 1 < len(formula)
                    and formula[j + 1] == '"'
                ):
                    j += 2
                    continue

                inside = not inside

            elif not inside:
                if ch == "(":
                    depth += 1

                elif ch == ")":
                    depth -= 1

                    if depth == 0:
                        yield formula[start:j]
                        break

            j += 1

        i = j + 1


def normalize_formula(value):
    if value is None:
        return ""

    if not isinstance(value, str):
        value = (
            getattr(value, "text", None)
            or getattr(value, "formula", None)
            or text(value)
        )

    value = (
        value.strip()
        .replace("“", '"')
        .replace("”", '"')
        .replace("＂", '"')
    )

    if (
        value.startswith("{=")
        and value.endswith("}")
    ):
        value = value[1:-1]

    return (
        value[1:]
        if value.startswith("=")
        else value
    )


def unquoted(expr):
    return re.sub(
        r'"(?:[^"]|"")*"',
        lambda m: " " * len(m.group()),
        expr
    )


def sanitized_copy_without_pivot_views(
    source,
    destination
):
    """
    Create a temporary compatibility copy.

    Removes PivotTable view references from the
    temporary copy only. The original workbook
    remains untouched and is used for pivot analysis.
    """

    with zipfile.ZipFile(source) as src:
        with zipfile.ZipFile(
            destination,
            "w"
        ) as dst:

            for info in src.infolist():
                payload = src.read(info.filename)

                if re.fullmatch(
                    r"xl/worksheets/_rels/"
                    r"sheet\d+\.xml\.rels",
                    info.filename
                ):
                    root = ET.fromstring(payload)

                    for el in list(root):
                        if el.get(
                            "Type", ""
                        ).endswith("/pivotTable"):
                            root.remove(el)

                    payload = ET.tostring(
                        root,
                        encoding="utf-8",
                        xml_declaration=True
                    )

                elif re.fullmatch(
                    r"xl/worksheets/"
                    r"sheet\d+\.xml",
                    info.filename
                ):
                    root = ET.fromstring(payload)

                    for el in list(root):
                        if (
                            local(el.tag)
                            == "pivotTableParts"
                        ):
                            root.remove(el)

                    payload = ET.tostring(
                        root,
                        encoding="utf-8",
                        xml_declaration=True
                    )

                dst.writestr(
                    info,
                    payload
                )


# ============================================================
# Analyzer
# ============================================================

class Audit:

    def __init__(self, path):
        self.path = Path(path)
        self._temp = None

        try:
            self.wf = load_workbook(
                path,
                data_only=False,
                read_only=False,
                keep_links=False
            )

            self.wv = load_workbook(
                path,
                data_only=True,
                read_only=False,
                keep_links=False
            )

        except (TypeError, ValueError, KeyError) as exc:
            print(
                "[COMPAT] Pivot metadata issue:",
                str(exc)[:180],
                flush=True
            )

            self._temp = tempfile.TemporaryDirectory(
                prefix="xlsm_audit_"
            )

            clean = (
                Path(self._temp.name)
                / Path(path).name
            )

            sanitized_copy_without_pivot_views(
                path,
                clean
            )

            self.wf = load_workbook(
                clean,
                data_only=False,
                read_only=False,
                keep_links=False
            )

            self.wv = load_workbook(
                clean,
                data_only=True,
                read_only=False,
                keep_links=False
            )

            print(
                "[COMPAT] Reading formulas from "
                "temporary copy; pivot metadata "
                "from original.",
                flush=True
            )

        self.names = list(self.wf.sheetnames)

        self.lookup = {
            name.casefold(): name
            for name in self.names
        }

        self.data = set()
        self.nav = set()

        self.maybe_data = set()
        self.maybe_nav = set()

        self.evidence = []
        self.unresolved = []
        self.shapes = []
        self.routes = []
        self.indirect_debug = []
        self.pivots = []

        self.procs = collections.defaultdict(list)
        self.stats = collections.Counter()

        self.cache = {}
        self.active = set()
        self.seen = set()

        self.definitions = {}

        for name, item in (
            self.wf.defined_names.items()
        ):
            self.definitions[
                name.casefold()
            ] = text(item.attr_text)

    def canon(self, name):
        return self.lookup.get(
            text(name)
            .strip()
            .replace("''", "'")
            .casefold()
        )

    # --------------------------------------------------------
    # Record relationships
    # --------------------------------------------------------

    def add(
        self,
        src,
        dst,
        kind,
        reason,
        where,
        detail,
        confirmed=True
    ):
        src = self.canon(src)
        dst = self.canon(dst)

        if not src or not dst:
            return

        pair = (src, dst)

        if kind == "DATA":
            if confirmed:
                self.data.add(pair)
            else:
                self.maybe_data.add(pair)

        else:
            if confirmed:
                self.nav.add(pair)
            else:
                self.maybe_nav.add(pair)

        key = (
            src,
            dst,
            kind,
            reason,
            where,
            confirmed
        )

        if key not in self.seen:
            self.seen.add(key)

            self.evidence.append((
                src,
                dst,
                kind,
                (
                    "Confirmed"
                    if confirmed
                    else "Possible"
                ),
                reason,
                where,
                text(detail)[:16000]
            ))

    # --------------------------------------------------------
    # Cell value resolution
    # --------------------------------------------------------

    def cellvalue(
        self,
        sheet,
        coord,
        depth=0
    ):
        sheet = self.canon(sheet)

        if not sheet or depth > 18:
            return None

        key = (
            sheet,
            coord.replace(
                "$", ""
            ).upper()
        )

        if key in self.cache:
            return self.cache[key]

        if key in self.active:
            return None

        self.active.add(key)

        try:
            value = (
                self.wv[sheet][key[1]].value
            )

            if value is None:
                raw = (
                    self.wf[sheet][key[1]].value
                )

                if (
                    getattr(raw, "text", None)
                    or (
                        isinstance(raw, str)
                        and raw.startswith("=")
                    )
                ):
                    value = self.evaluate(
                        normalize_formula(raw),
                        sheet,
                        key[1],
                        depth + 1
                    )
                else:
                    value = raw

            self.cache[key] = value

            return value

        except Exception:
            return None

        finally:
            self.active.discard(key)

    # --------------------------------------------------------
    # Excel expression evaluator
    # --------------------------------------------------------

    def evaluate(
        self,
        expr,
        sheet,
        coord,
        depth=0
    ):
        if depth > 18:
            return None

        expr = normalize_formula(expr).strip()

        if not expr:
            return None

        while (
            expr.startswith("(")
            and expr.endswith(")")
        ):
            inside = expr[1:-1]

            if self._balanced(inside):
                expr = inside.strip()
            else:
                break

        # String concatenation
        pieces = split(expr, "&")

        if len(pieces) > 1:
            values = [
                self.evaluate(
                    part,
                    sheet,
                    coord,
                    depth + 1
                )
                for part in pieces
            ]

            if any(
                value is None
                for value in values
            ):
                return None

            return "".join(
                map(str, values)
            )

        # String literal
        if (
            len(expr) >= 2
            and expr[0] == '"'
            and expr[-1] == '"'
        ):
            return expr[1:-1].replace(
                '""',
                '"'
            )

        # Number
        if re.fullmatch(
            r"[+-]?\d+(?:\.\d+)?",
            expr
        ):
            return (
                float(expr)
                if "." in expr
                else int(expr)
            )

        # Cell reference
        match = CELL.fullmatch(expr)

        if match:
            return self.cellvalue(
                match.group(1)
                or match.group(2)
                or sheet,
                match.group(3),
                depth + 1
            )

        # Defined names
        if expr.casefold() in self.definitions:
            return self.evaluate(
                self.definitions[
                    expr.casefold()
                ],
                sheet,
                coord,
                depth + 1
            )

        supported = (
            "IFERROR",
            "IF",
            "SUBSTITUTE",
            "TRIM",
            "UPPER",
            "LOWER",
            "CONCAT",
            "CONCATENATE",
            "TEXT"
        )

        for fn in supported:
            if not re.match(
                r"^(?:_xlfn\.)?"
                + fn
                + r"\s*\(",
                expr,
                re.I
            ):
                continue

            found = list(
                calls(expr, fn)
            )

            if not found:
                continue

            args = split(
                found[0],
                ",;"
            )

            values = [
                self.evaluate(
                    arg,
                    sheet,
                    coord,
                    depth + 1
                )
                for arg in args
            ]

            if (
                fn == "IFERROR"
                and len(values) >= 2
            ):
                return (
                    values[0]
                    if values[0] is not None
                    else values[1]
                )

            if (
                fn == "IF"
                and len(values) >= 3
                and isinstance(
                    values[0],
                    (int, float, bool)
                )
            ):
                return (
                    values[1]
                    if values[0]
                    else values[2]
                )

            if (
                fn == "SUBSTITUTE"
                and len(values) >= 3
                and all(
                    v is not None
                    for v in values[:3]
                )
            ):
                return str(values[0]).replace(
                    str(values[1]),
                    str(values[2])
                )

            if (
                fn == "TRIM"
                and values
                and values[0] is not None
            ):
                return " ".join(
                    str(values[0]).split()
                )

            if (
                fn == "UPPER"
                and values
                and values[0] is not None
            ):
                return str(
                    values[0]
                ).upper()

            if (
                fn == "LOWER"
                and values
                and values[0] is not None
            ):
                return str(
                    values[0]
                ).lower()

            if (
                fn in (
                    "CONCAT",
                    "CONCATENATE"
                )
                and all(
                    v is not None
                    for v in values
                )
            ):
                return "".join(
                    map(str, values)
                )

            if (
                fn == "TEXT"
                and values
                and values[0] is not None
            ):
                return str(values[0])

        return None

    @staticmethod
    def _balanced(expr):
        depth = 0
        quote = False
        i = 0

        while i < len(expr):
            if expr[i] == '"':
                if (
                    quote
                    and i + 1 < len(expr)
                    and expr[i + 1] == '"'
                ):
                    i += 2
                    continue

                quote = not quote

            elif not quote:
                if expr[i] == "(":
                    depth += 1

                elif expr[i] == ")":
                    depth -= 1

                    if depth < 0:
                        return False

            i += 1

        return depth == 0 and not quote

    def target(self, value):
        value = (
            text(value)
            .strip()
            .lstrip("=#")
        )

        match = SREF.search(value)

        if match:
            return self.canon(
                match.group(1)
                or match.group(2)
            )

        if value.casefold() in self.definitions:
            return self.target(
                self.definitions[
                    value.casefold()
                ]
            )

        return None

    # --------------------------------------------------------
    # Formula scanning
    # --------------------------------------------------------

    def scan_formulas(self):
        print(
            "[1/6] Scanning formulas",
            flush=True
        )

        for ws in self.wf:
            count = 0

            for row in ws:
                for cell in row:
                    value = cell.value

                    if (
                        cell.data_type != "f"
                        and not (
                            isinstance(value, str)
                            and value.startswith("=")
                        )
                    ):
                        continue

                    formula = normalize_formula(
                        value
                    )

                    if not formula:
                        continue

                    count += 1
                    self.stats["Formulas"] += 1

                    # Direct references
                    for match in SREF.finditer(
                        unquoted(formula)
                    ):
                        self.add(
                            ws.title,
                            match.group(1)
                            or match.group(2),
                            "DATA",
                            "DIRECT",
                            cell.coordinate,
                            formula
                        )

                    # INDIRECT
                    for expr in calls(
                        formula,
                        "INDIRECT"
                    ):
                        self.stats["INDIRECT"] += 1

                        args = split(
                            expr,
                            ",;"
                        )

                        evaluated = (
                            self.evaluate(
                                args[0],
                                ws.title,
                                cell.coordinate
                            )
                            if args
                            else None
                        )

                        target_sheet = self.target(
                            evaluated
                        )

                        status = (
                            "RESOLVED"
                            if target_sheet
                            else "UNRESOLVED"
                        )

                        self.indirect_debug.append((
                            ws.title,
                            cell.coordinate,
                            expr,
                            text(evaluated),
                            text(target_sheet),
                            status,
                            (
                                text(
                                    self.cellvalue(
                                        ws.title,
                                        "J6"
                                    )
                                )
                                if re.search(
                                    r"\$?J\$?6\b",
                                    expr,
                                    re.I
                                )
                                else ""
                            ),
                            (
                                text(
                                    self.cellvalue(
                                        ws.title,
                                        "K6"
                                    )
                                )
                                if re.search(
                                    r"\$?K\$?6\b",
                                    expr,
                                    re.I
                                )
                                else ""
                            )
                        ))

                        if target_sheet:
                            self.add(
                                ws.title,
                                target_sheet,
                                "DATA",
                                "INDIRECT",
                                cell.coordinate,
                                (
                                    formula
                                    + " | "
                                    + text(evaluated)
                                )
                            )

                        elif (
                            evaluated is not None
                            and re.fullmatch(
                                r"\$?[A-Z]{1,3}"
                                r"\$?\d+"
                                r"(?::\$?[A-Z]{1,3}"
                                r"\$?\d+)?",
                                text(evaluated),
                                re.I
                            )
                        ):
                            self.add(
                                ws.title,
                                ws.title,
                                "DATA",
                                "INDIRECT_LOCAL",
                                cell.coordinate,
                                formula
                            )

                        else:
                            self.unresolved.append((
                                ws.title,
                                cell.coordinate,
                                "INDIRECT",
                                formula,
                                "Evaluated: "
                                + repr(evaluated)
                            ))

                    # OFFSET
                    for expr in calls(
                        formula,
                        "OFFSET"
                    ):
                        self.stats["OFFSET"] += 1

                    # HYPERLINK
                    for expr in calls(
                        formula,
                        "HYPERLINK"
                    ):
                        args = split(
                            expr,
                            ",;"
                        )

                        evaluated = (
                            self.evaluate(
                                args[0],
                                ws.title,
                                cell.coordinate
                            )
                            if args
                            else None
                        )

                        target_sheet = self.target(
                            evaluated
                        )

                        if target_sheet:
                            self.add(
                                ws.title,
                                target_sheet,
                                "NAV",
                                "HYPERLINK_FORMULA",
                                cell.coordinate,
                                formula
                            )

            print(
                f"  {ws.title}: "
                f"{count:,} formulas",
                flush=True
            )

    # --------------------------------------------------------
    # Shape and hyperlink analysis
    # --------------------------------------------------------

    def shape(
        self,
        sheet,
        kind,
        ident,
        name,
        link,
        macro,
        part
    ):
        self.shapes.append((
            sheet,
            kind,
            text(ident),
            text(name),
            text(link),
            text(macro),
            part
        ))

        for item in text(link).split(";"):
            target_sheet = self.target(item)

            if target_sheet:
                self.add(
                    sheet,
                    target_sheet,
                    "NAV",
                    "SHAPE_LINK",
                    ident,
                    item
                )

    def scan_shapes(self):
        print(
            "[2/6] Scanning shapes",
            flush=True
        )

        with zipfile.ZipFile(self.path) as z:
            files = set(z.namelist())

            workbook_rels = relationships(
                z,
                "xl/workbook.xml"
            )

            workbook_root = ET.fromstring(
                z.read("xl/workbook.xml")
            )

            for sh in workbook_root.iter():
                if local(sh.tag) != "sheet":
                    continue

                relation = workbook_rels.get(
                    sh.get(
                        "{" + RNS + "}id"
                    )
                )

                if not relation:
                    continue

                sheet = sh.get("name")

                part = joinpart(
                    "xl/workbook.xml",
                    relation[1]
                )

                if part not in files:
                    continue

                sheet_rels = relationships(
                    z,
                    part
                )

                sheet_root = ET.fromstring(
                    z.read(part)
                )

                for el in sheet_root.iter():
                    tag = local(el.tag)

                    if tag == "hyperlink":
                        relation = sheet_rels.get(
                            el.get(
                                "{" + RNS + "}id"
                            )
                        )

                        link = (
                            el.get("location")
                            or (
                                relation[1]
                                if relation
                                else ""
                            )
                        )

                        self.shape(
                            sheet,
                            "CELL_LINK",
                            el.get("ref"),
                            el.get("display"),
                            link,
                            "",
                            part
                        )

                    if (
                        tag in (
                            "control",
                            "oleObject"
                        )
                        and el.get("macro")
                    ):
                        self.shape(
                            sheet,
                            tag,
                            el.get("shapeId"),
                            el.get("name"),
                            "",
                            el.get("macro"),
                            part
                        )

                seen = set()

                for rid, relation in (
                    sheet_rels.items()
                ):
                    typ, target, mode = relation

                    if typ not in (
                        "drawing",
                        "vmlDrawing",
                        "ctrlProp",
                        "control"
                    ):
                        continue

                    child = joinpart(
                        part,
                        target
                    )

                    if (
                        child not in files
                        or child in seen
                    ):
                        continue

                    seen.add(child)

                    try:
                        root = ET.fromstring(
                            z.read(child)
                        )

                        drawing_rels = relationships(
                            z,
                            child
                        )

                        if typ == "drawing":
                            for obj in root.iter():
                                if local(obj.tag) not in (
                                    "sp",
                                    "pic",
                                    "graphicFrame",
                                    "cxnSp"
                                ):
                                    continue

                                ident = ""
                                name = ""
                                macro = ""
                                links = []

                                for x in obj.iter():
                                    tag = local(x.tag)

                                    if tag == "cNvPr":
                                        ident = x.get(
                                            "id", ""
                                        )
                                        name = x.get(
                                            "name", ""
                                        )

                                    macro = (
                                        macro
                                        or x.get("macro", "")
                                        or x.get(
                                            "onAction", ""
                                        )
                                    )

                                    if tag in (
                                        "hlinkClick",
                                        "hlinkHover"
                                    ):
                                        relation = (
                                            drawing_rels.get(
                                                x.get(
                                                    "{"
                                                    + RNS
                                                    + "}id"
                                                )
                                            )
                                        )

                                        if relation:
                                            links.append(
                                                relation[1]
                                            )

                                        for attr in (
                                            "action",
                                            "location"
                                        ):
                                            if x.get(attr):
                                                links.append(
                                                    x.get(attr)
                                                )

                                self.shape(
                                    sheet,
                                    local(obj.tag),
                                    ident,
                                    name,
                                    ";".join(links),
                                    macro,
                                    child
                                )

                        elif typ == "vmlDrawing":
                            for obj in root.iter():
                                if local(obj.tag) != "shape":
                                    continue

                                ident = obj.get(
                                    "id", ""
                                )

                                macro = obj.get(
                                    "macro", ""
                                )

                                links = []

                                for x in obj.iter():
                                    tag = local(x.tag)

                                    if (
                                        tag in (
                                            "FmlaMacro",
                                            "Macro"
                                        )
                                        and text(
                                            x.text
                                        ).strip()
                                    ):
                                        macro = text(
                                            x.text
                                        ).strip()

                                    if tag in (
                                        "hlinkClick",
                                        "hlinkHover"
                                    ):
                                        relation = (
                                            drawing_rels.get(
                                                x.get(
                                                    "{"
                                                    + RNS
                                                    + "}id"
                                                )
                                            )
                                        )

                                        if relation:
                                            links.append(
                                                relation[1]
                                            )

                                self.shape(
                                    sheet,
                                    "VML",
                                    ident,
                                    ident,
                                    ";".join(links),
                                    macro,
                                    child
                                )

                        else:
                            for obj in root.iter():
                                macro = (
                                    obj.get("macro")
                                    or obj.get(
                                        "fmlaMacro"
                                    )
                                    or obj.get(
                                        "onAction"
                                    )
                                )

                                if macro:
                                    self.shape(
                                        sheet,
                                        "CONTROL",
                                        obj.get("id"),
                                        obj.get("name"),
                                        "",
                                        macro,
                                        child
                                    )

                    except Exception as exc:
                        self.unresolved.append((
                            sheet,
                            child,
                            "XML",
                            "",
                            repr(exc)
                        ))

    # --------------------------------------------------------
    # PivotTable source analysis
    # --------------------------------------------------------

    def scan_pivots(self):
        print(
            "[3/6] Scanning PivotTables",
            flush=True
        )

        with zipfile.ZipFile(self.path) as z:
            files = set(z.namelist())

            workbook_part = "xl/workbook.xml"

            workbook_rels = relationships(
                z,
                workbook_part
            )

            workbook_root = ET.fromstring(
                z.read(workbook_part)
            )

            cache_parts = {}

            for item in workbook_root.iter():
                if local(item.tag) != "pivotCache":
                    continue

                cache_id = item.get(
                    "cacheId", ""
                )

                rel_id = item.get(
                    "{" + RNS + "}id"
                )

                relation = workbook_rels.get(
                    rel_id
                )

                if relation:
                    cache_parts[
                        cache_id
                    ] = joinpart(
                        workbook_part,
                        relation[1]
                    )

            cache_info = {}

            def read_cache(part):
                if part in cache_info:
                    return cache_info[part]

                info = (
                    "UNRESOLVED",
                    "",
                    "",
                    "",
                    ""
                )

                if part in files:
                    try:
                        root = ET.fromstring(
                            z.read(part)
                        )

                        for node in root.iter():
                            if (
                                local(node.tag)
                                != "cacheSource"
                            ):
                                continue

                            source_type = node.get(
                                "type", ""
                            )

                            worksheet_source = next(
                                (
                                    child
                                    for child in node
                                    if local(
                                        child.tag
                                    ) == "worksheetSource"
                                ),
                                None
                            )

                            if worksheet_source is not None:
                                info = (
                                    source_type,
                                    worksheet_source.get(
                                        "sheet", ""
                                    ),
                                    worksheet_source.get(
                                        "ref", ""
                                    ),
                                    worksheet_source.get(
                                        "name", ""
                                    ),
                                    ""
                                )

                            else:
                                info = (
                                    source_type,
                                    "",
                                    "",
                                    "",
                                    "No worksheetSource"
                                )

                            break

                    except Exception as exc:
                        info = (
                            "ERROR",
                            "",
                            "",
                            "",
                            repr(exc)
                        )

                cache_info[part] = info

                return info

            for sh in workbook_root.iter():
                if local(sh.tag) != "sheet":
                    continue

                sheet = sh.get(
                    "name", ""
                )

                relation = workbook_rels.get(
                    sh.get(
                        "{" + RNS + "}id"
                    )
                )

                if not relation:
                    continue

                sheet_part = joinpart(
                    workbook_part,
                    relation[1]
                )

                if sheet_part not in files:
                    continue

                sheet_rels = relationships(
                    z,
                    sheet_part
                )

                for rid, relation in (
                    sheet_rels.items()
                ):
                    typ, target, mode = relation

                    if typ != "pivotTable":
                        continue

                    pivot_part = joinpart(
                        sheet_part,
                        target
                    )

                    if pivot_part not in files:
                        self.unresolved.append((
                            sheet,
                            pivot_part,
                            "PIVOT",
                            "",
                            "Missing pivot definition"
                        ))
                        continue

                    try:
                        root = ET.fromstring(
                            z.read(pivot_part)
                        )

                        pivot_name = root.get(
                            "name", ""
                        )

                        cache_id = root.get(
                            "cacheId", ""
                        )

                        cache_part = ""

                        pivot_rels = relationships(
                            z,
                            pivot_part
                        )

                        for rel in pivot_rels.values():
                            if (
                                rel[0]
                                == "pivotCacheDefinition"
                            ):
                                cache_part = joinpart(
                                    pivot_part,
                                    rel[1]
                                )
                                break

                        if not cache_part:
                            cache_part = (
                                cache_parts.get(
                                    cache_id,
                                    ""
                                )
                            )

                        (
                            source_type,
                            source_sheet,
                            source_ref,
                            source_name,
                            issue
                        ) = read_cache(cache_part)

                        target_sheet = self.canon(
                            source_sheet
                        )

                        status = "UNRESOLVED"

                        if source_type != "worksheet":
                            issue = (
                                issue
                                + "; External/connection/"
                                "model source"
                            )

                        elif (
                            source_sheet
                            and not target_sheet
                        ):
                            issue = (
                                issue
                                + "; Source sheet "
                                "not found"
                            )

                        if (
                            source_type == "worksheet"
                            and target_sheet
                        ):
                            status = "CONFIRMED"

                            self.add(
                                sheet,
                                target_sheet,
                                "DATA",
                                "PIVOT_SOURCE",
                                pivot_name,
                                (
                                    "Pivot "
                                    + pivot_name
                                    + " source="
                                    + source_sheet
                                    + "!"
                                    + source_ref
                                )
                            )

                        elif (
                            source_type == "worksheet"
                            and source_name
                        ):
                            expression = (
                                self.definitions.get(
                                    source_name.casefold(),
                                    ""
                                )
                            )

                            if not expression:
                                local_names = (
                                    self.wf[
                                        sheet
                                    ].defined_names
                                )

                                for key, definition in (
                                    local_names.items()
                                ):
                                    if (
                                        key.casefold()
                                        == source_name.casefold()
                                    ):
                                        expression = text(
                                            definition.attr_text
                                        )
                                        break

                            if expression:
                                target_sheet = self.target(
                                    expression
                                )

                            if target_sheet:
                                status = "CONFIRMED_NAME"

                                self.add(
                                    sheet,
                                    target_sheet,
                                    "DATA",
                                    "PIVOT_NAMED_SOURCE",
                                    pivot_name,
                                    (
                                        source_name
                                        + " -> "
                                        + expression
                                    )
                                )

                            else:
                                matches = self._table_sources(
                                    z,
                                    files,
                                    source_name
                                )

                                if len(matches) == 1:
                                    target_sheet = matches[0]
                                    status = "CONFIRMED_TABLE"

                                    self.add(
                                        sheet,
                                        target_sheet,
                                        "DATA",
                                        "PIVOT_TABLE_SOURCE",
                                        pivot_name,
                                        source_name
                                    )

                                elif len(matches) > 1:
                                    status = "AMBIGUOUS_TABLE"

                                    for candidate in matches:
                                        self.add(
                                            sheet,
                                            candidate,
                                            "DATA",
                                            "PIVOT_TABLE_CANDIDATE",
                                            pivot_name,
                                            source_name,
                                            False
                                        )

                        self.pivots.append((
                            sheet,
                            pivot_name,
                            cache_id,
                            source_type,
                            source_sheet,
                            source_ref,
                            source_name,
                            target_sheet or "",
                            status,
                            cache_part,
                            issue
                        ))

                        self.stats["PivotTables"] += 1

                        if (
                            status.startswith("UNRESOLVED")
                            or status == "AMBIGUOUS_TABLE"
                        ):
                            self.unresolved.append((
                                sheet,
                                pivot_name,
                                "PIVOT",
                                source_name or source_ref,
                                (
                                    "Source unresolved: "
                                    + source_type
                                    + "; "
                                    + issue
                                )
                            ))

                    except Exception as exc:
                        self.unresolved.append((
                            sheet,
                            pivot_part,
                            "PIVOT",
                            "",
                            repr(exc)
                        ))

            print(
                "  Pivot tables:",
                self.stats["PivotTables"],
                flush=True
            )

    def _table_sources(
        self,
        z,
        files,
        table_name
    ):
        found = set()

        workbook_part = "xl/workbook.xml"

        workbook_rels = relationships(
            z,
            workbook_part
        )

        root = ET.fromstring(
            z.read(workbook_part)
        )

        for sh in root.iter():
            if local(sh.tag) != "sheet":
                continue

            relation = workbook_rels.get(
                sh.get(
                    "{" + RNS + "}id"
                )
            )

            if not relation:
                continue

            part = joinpart(
                workbook_part,
                relation[1]
            )

            if part not in files:
                continue

            for rel in relationships(
                z,
                part
            ).values():
                kind, target, mode = rel

                if kind != "table":
                    continue

                table_part = joinpart(
                    part,
                    target
                )

                if table_part not in files:
                    continue

                try:
                    table = ET.fromstring(
                        z.read(table_part)
                    )

                    display_name = table.get(
                        "displayName", ""
                    )

                    name = table.get(
                        "name", ""
                    )

                    if (
                        display_name.casefold()
                        == table_name.casefold()
                        or name.casefold()
                        == table_name.casefold()
                    ):
                        found.add(
                            sh.get("name", "")
                        )

                except Exception:
                    pass

        return sorted(found)

    # --------------------------------------------------------
    # VBA analysis
    # --------------------------------------------------------

    def scan_vba(self):
        print(
            "[4/6] Scanning VBA",
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
                "Install oletools"
            ))
            return

        parser = None

        try:
            parser = VBA_Parser(
                str(self.path)
            )

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
                        PROC.finditer(code)
                    )

                    for i, match in enumerate(
                        matches
                    ):
                        end = (
                            matches[i + 1].start()
                            if i + 1 < len(matches)
                            else len(code)
                        )

                        body = code[
                            match.start():end
                        ]

                        self.procs[
                            match.group(1).casefold()
                        ].append((
                            module,
                            body
                        ))

        except Exception as exc:
            self.unresolved.append((
                "Workbook",
                "VBA",
                "VBA",
                "",
                repr(exc)
            ))

        finally:
            if parser:
                try:
                    parser.close()
                except Exception:
                    pass

        for shape in self.shapes:
            (
                sheet,
                kind,
                ident,
                name,
                link,
                macro,
                part
            ) = shape

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

            seen = set()

            targets = collections.defaultdict(
                set
            )

            def walk(procname, depth=0):
                if (
                    depth > 10
                    or procname in seen
                ):
                    return

                seen.add(procname)

                for module, body in (
                    self.procs.get(
                        procname,
                        []
                    )
                ):
                    for line in body.splitlines():
                        if line.lstrip().startswith("'"):
                            continue

                        for match in VREF.finditer(
                            line
                        ):
                            dst = self.canon(
                                match.group(1)
                            )

                            if dst:
                                targets[dst].add(
                                    "NAV"
                                    if NAV.search(line)
                                    else "REFERENCE"
                                )

                        match = re.match(
                            r"(?i)^\s*"
                            r"(?:Call\s+)?"
                            r"([A-Za-z_]\w*)\b",
                            line
                        )

                        if match:
                            called = (
                                match.group(1)
                                .casefold()
                            )

                            if called in self.procs:
                                walk(
                                    called,
                                    depth + 1
                                )

            walk(proc)

            if not targets:
                self.routes.append((
                    sheet,
                    ident,
                    name,
                    macro,
                    "",
                    "UNRESOLVED",
                    "Macro/target not resolved"
                ))

            for dst, kinds in targets.items():
                status = (
                    "POSSIBLE_NAV"
                    if "NAV" in kinds
                    else "REFERENCE_ONLY"
                )

                if "NAV" in kinds:
                    self.add(
                        sheet,
                        dst,
                        "NAV",
                        "VBA_CANDIDATE",
                        ident,
                        macro,
                        False
                    )

                self.routes.append((
                    sheet,
                    ident,
                    name,
                    macro,
                    dst,
                    status,
                    ",".join(
                        sorted(kinds)
                    )
                ))

    # --------------------------------------------------------
    # Streaming report
    # --------------------------------------------------------

    def report(self, output):
        print(
            "[5/6] Streaming report",
            flush=True
        )

        self.maybe_data -= self.data
        self.maybe_nav -= self.nav

        wb = Workbook(
            write_only=True
        )

        def tab(
            name,
            headers,
            rows
        ):
            start = time.perf_counter()

            print(
                "  [START]",
                name,
                flush=True
            )

            ws = wb.create_sheet(name)

            header_cells = []

            for header in headers:
                cell = WriteOnlyCell(
                    ws,
                    value=header
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

                        if value.startswith(
                            ("=", "+", "-", "@")
                        ):
                            cell = WriteOnlyCell(
                                ws,
                                value=value
                            )

                            cell.data_type = "s"
                            value = cell

                    values.append(value)

                ws.append(values)

                count += 1

                if count % 10000 == 0:
                    print(
                        f"    {name}: "
                        f"{count:,} rows",
                        flush=True
                    )

            elapsed = (
                time.perf_counter()
                - start
            )

            print(
                f"  [DONE] {name}: "
                f"{count:,} rows, "
                f"{elapsed:.1f}s",
                flush=True
            )

        # Summary
        tab(
            "Summary",
            ["Metric", "Value"],
            [
                ("Source", str(self.path)),
                ("Sheets", len(self.names)),
                ("Formulas", self.stats["Formulas"]),
                ("INDIRECT", self.stats["INDIRECT"]),
                (
                    "INDIRECT resolved",
                    sum(
                        r[5] == "RESOLVED"
                        for r in self.indirect_debug
                    )
                ),
                ("OFFSET", self.stats["OFFSET"]),
                ("Data confirmed", len(self.data)),
                ("Data possible", len(self.maybe_data)),
                ("Navigation confirmed", len(self.nav)),
                ("Navigation possible", len(self.maybe_nav)),
                ("Shapes", len(self.shapes)),
                (
                    "Pivot tables",
                    self.stats["PivotTables"]
                ),
                (
                    "Pivot confirmed",
                    sum(
                        r[8].startswith("CONFIRMED")
                        for r in self.pivots
                    )
                ),
                (
                    "Pivot unresolved",
                    sum(
                        not r[8].startswith("CONFIRMED")
                        for r in self.pivots
                    )
                ),
                ("Macro routes", len(self.routes)),
                ("Unresolved", len(self.unresolved))
            ]
        )

        # Sheet inventory
        tab(
            "Sheets",
            ["Sheet", "Visibility"],
            (
                (
                    name,
                    self.wf[name].sheet_state
                )
                for name in self.names
            )
        )

        # Dependency matrices
        def matrix(name, yes, maybe):
            rows = (
                [src] + [
                    (
                        1
                        if (src, dst) in yes
                        else (
                            "?"
                            if (src, dst) in maybe
                            else 0
                        )
                    )
                    for dst in self.names
                ]
                for src in self.names
            )

            tab(
                name,
                ["Source / Target"] + self.names,
                rows
            )

        matrix(
            "Data_Matrix",
            self.data,
            self.maybe_data
        )

        matrix(
            "Navigation_Matrix",
            self.nav,
            self.maybe_nav
        )

        combined_yes = (
            self.data | self.nav
        )

        combined_maybe = (
            self.maybe_data
            | self.maybe_nav
        ) - combined_yes

        matrix(
            "Combined_Matrix",
            combined_yes,
            combined_maybe
        )

        # Detailed audit sheets
        tab(
            "Evidence",
            [
                "Source",
                "Target",
                "Type",
                "Confidence",
                "Category",
                "Cell/Object",
                "Evidence"
            ],
            self.evidence
        )

        tab(
            "Unresolved",
            [
                "Sheet",
                "Cell",
                "Type",
                "Formula",
                "Reason"
            ],
            self.unresolved
        )

        tab(
            "INDIRECT_Debug",
            [
                "Sheet",
                "Cell",
                "Expression",
                "Evaluated",
                "Target Sheet",
                "Status",
                "J6",
                "K6"
            ],
            self.indirect_debug
        )

        tab(
            "Shapes_Links",
            [
                "Sheet",
                "Type",
                "ID",
                "Name",
                "Link",
                "Macro",
                "Part"
            ],
            self.shapes
        )

        tab(
            "Pivot_Audit",
            [
                "Pivot Sheet",
                "Pivot Name",
                "Cache ID",
                "Source Type",
                "Source Sheet",
                "Source Range",
                "Source Name/Table",
                "Resolved Sheet",
                "Status",
                "Cache Definition Part",
                "Detail"
            ],
            self.pivots
        )

        tab(
            "Macro_Routes",
            [
                "Source",
                "ID",
                "Name",
                "Macro",
                "Target",
                "Status",
                "Reason"
            ],
            self.routes
        )

        # TOP button audit
        def top_rows():
            for row in self.shapes:
                if row[0].casefold() != "top":
                    continue

                matched = [
                    route
                    for route in self.routes
                    if (
                        route[0] == row[0]
                        and route[1] == row[2]
                    )
                ]

                targets = sorted({
                    route[4]
                    for route in matched
                    if route[4]
                })

                statuses = sorted({
                    route[5]
                    for route in matched
                })

                yield (
                    row[2],
                    row[3],
                    row[1],
                    row[5],
                    row[4],
                    ";".join(targets),
                    ";".join(statuses)
                )

        tab(
            "TOP_Button_Audit",
            [
                "ID",
                "Name",
                "Type",
                "Macro",
                "Link",
                "Targets",
                "Status"
            ],
            top_rows()
        )

        # Regression checks
        pair = (
            self.canon("サマリ"),
            self.canon(
                "1023951_I_STAR現行売上"
            )
        )

        check = (
            "PASS"
            if pair in self.data
            else (
                "POSSIBLE"
                if pair in self.maybe_data
                else "NOT_RESOLVED"
            )
        )

        tab(
            "Regression_Check",
            [
                "Check",
                "Result",
                "Detail"
            ],
            [
                (
                    "サマリ → "
                    "1023951_I_STAR現行売上",
                    check,
                    "Confirmed data dependency"
                ),
                (
                    "サマリ!E36",
                    repr(
                        self.cellvalue(
                            "サマリ",
                            "E36"
                        )
                    ),
                    "Source value"
                ),
                (
                    "TOP shapes",
                    sum(
                        r[0].casefold() == "top"
                        for r in self.shapes
                    ),
                    "Objects"
                ),
                (
                    "TOP macro bindings",
                    sum(
                        r[0].casefold() == "top"
                        and bool(r[5])
                        for r in self.shapes
                    ),
                    "Macro objects"
                ),
                (
                    "TOP confirmed navigation",
                    sum(
                        src.casefold() == "top"
                        for src, dst in self.nav
                    ),
                    "Direct links"
                ),
                (
                    "TOP possible navigation",
                    sum(
                        src.casefold() == "top"
                        for src, dst in self.maybe_nav
                    ),
                    "VBA candidates"
                ),
                (
                    "Pivot tables",
                    self.stats["PivotTables"],
                    "Pivot source relations"
                )
            ]
        )

        # README
        tab(
            "README",
            ["Item", "Meaning"],
            [
                (
                    "1",
                    "Confirmed static dependency"
                ),
                (
                    "?",
                    "Possible dependency, not proven"
                ),
                (
                    "0",
                    "No evidence found; "
                    "not proof of absence"
                ),
                (
                    "Data_Matrix",
                    "Data and formula dependencies"
                ),
                (
                    "Navigation_Matrix",
                    "Hyperlinks and VBA candidates"
                ),
                (
                    "Combined_Matrix",
                    "Combined relationships"
                ),
                (
                    "INDIRECT_Debug",
                    "Evaluated INDIRECT expressions"
                ),
                (
                    "Pivot_Audit",
                    "Pivot sheet to source sheet"
                ),
                (
                    "Source",
                    "Never modified"
                ),
                (
                    "VBA",
                    "Macros are never executed"
                )
            ]
        )

        print(
            "[6/6] Saving XLSX...",
            flush=True
        )

        wb.save(output)

        print(
            "[DONE]",
            output,
            flush=True
        )

    def close(self):
        self.wf.close()
        self.wv.close()

        if self._temp:
            self._temp.cleanup()


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=__doc__
    )

    parser.add_argument(
        "source"
    )

    parser.add_argument(
        "-o",
        "--output"
    )

    args = parser.parse_args()

    source = Path(
        args.source
    ).resolve()

    if (
        not source.is_file()
        or not zipfile.is_zipfile(source)
    ):
        parser.error(
            "Invalid XLSM/XLSX file"
        )

    output = (
        Path(args.output).resolve()
        if args.output
        else source.with_name(
            source.stem
            + "_dependency_report_v6_3.xlsx"
        )
    )

    if output == source:
        parser.error(
            "Output cannot overwrite input"
        )

    start = time.perf_counter()

    audit = Audit(source)

    try:
        audit.scan_formulas()
        audit.scan_shapes()
        audit.scan_pivots()
        audit.scan_vba()

        output.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        audit.report(output)

        print(
            "FINISHED in "
            + str(
                round(
                    time.perf_counter() - start,
                    1
                )
            )
            + "s",
            flush=True
        )

    finally:
        audit.close()


if __name__ == "__main__":
    main()
