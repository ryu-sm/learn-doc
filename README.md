
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Dependency Analyzer V6

Install:
    pip install openpyxl oletools

Run:
    python -u analyze_xlsm_v6.py input.xlsm

Optional:
    python -u analyze_xlsm_v6.py input.xlsm -o report.xlsx

The source XLSM is never modified.
VBA macros are inspected, never executed.
"""

import argparse
import collections
import posixpath
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill


REL_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

SHEET_REF = re.compile(
    r"(?<![\w.])"
    r"(?:'((?:[^']|'')+)'|([\w.\u0080-\uffff]+))!"
    r"\s*(?="
    r"\$?[A-Z]{1,3}\$?\d"
    r"|\$?[A-Z]{1,3}\s*:"
    r"|\$?\d+\s*:)",
    re.I
)

SIMPLE_CELL = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|"
    r"([\w.\u0080-\uffff]+))!)?"
    r"(\$?[A-Z]{1,3}\$?\d+)",
    re.I
)

VBA_SHEET = re.compile(
    r'(?:ThisWorkbook\s*\.\s*)?'
    r'(?:Worksheets|Sheets)'
    r'\s*\(\s*"([^"]+)"\s*\)',
    re.I
)

VBA_PROC = re.compile(
    r"^\s*"
    r"(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function)\s+"
    r"([\w\u0080-\uffff]+)",
    re.I | re.M
)

VBA_GOTO = re.compile(
    r"(?i)"
    r"\b(?:Application\s*\.\s*)?Goto\b"
    r"|\.\s*(?:Activate|Select)\b"
)


# ============================================================
# Common utilities
# ============================================================

def local(tag):
    return tag.rsplit("}", 1)[-1]


def s(value):
    return "" if value is None else str(value)


def xmlrels(z, part):
    rp = posixpath.join(
        posixpath.dirname(part),
        "_rels",
        posixpath.basename(part) + ".rels"
    )

    if rp not in z.namelist():
        return {}

    root = ET.fromstring(z.read(rp))
    result = {}

    for r in root:
        result[r.get("Id")] = (
            r.get("Type", "").rsplit("/", 1)[-1],
            r.get("Target", ""),
            r.get("TargetMode", "")
        )

    return result


def partpath(part, target):
    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(part),
            target
        )
    )


def split(expr, separators=",;"):
    result = []
    buf = []
    depth = 0
    quote = False
    i = 0

    while i < len(expr):
        c = expr[i]

        if c == '"':
            if (
                quote
                and i + 1 < len(expr)
                and expr[i + 1] == '"'
            ):
                buf.extend(('"', '"'))
                i += 2
                continue

            quote = not quote

        if not quote:
            if c == "(":
                depth += 1

            elif c == ")":
                depth -= 1

            elif c in separators and depth == 0:
                result.append(
                    "".join(buf).strip()
                )
                buf = []
                i += 1
                continue

        buf.append(c)
        i += 1

    result.append("".join(buf).strip())
    return result


def functions(expr, name):
    pat = re.compile(
        r"\b" + re.escape(name) + r"\s*\(",
        re.I
    )

    i = 0
    quoted = False

    while i < len(expr):
        if expr[i] == '"':
            if (
                quoted
                and i + 1 < len(expr)
                and expr[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted
            i += 1
            continue

        match = (
            pat.match(expr, i)
            if not quoted
            else None
        )

        if not match:
            i += 1
            continue

        start = match.end()
        j = start
        depth = 1
        inside = False

        while j < len(expr):
            c = expr[j]

            if c == '"':
                if (
                    inside
                    and j + 1 < len(expr)
                    and expr[j + 1] == '"'
                ):
                    j += 2
                    continue

                inside = not inside

            elif not inside:
                if c == "(":
                    depth += 1

                elif c == ")":
                    depth -= 1

                    if depth == 0:
                        yield expr[start:j]
                        break

            j += 1

        i = j + 1


def unquote_formula(expr):
    return re.sub(
        r'"(?:[^"]|"")*"',
        lambda m: " " * len(m.group()),
        expr
    )


def strip_outer(expr):
    expr = expr.strip()

    while (
        expr.startswith("(")
        and expr.endswith(")")
    ):
        depth = 0
        quote = False
        good = True

        for i, c in enumerate(expr):
            if c == '"':
                quote = not quote

            if quote:
                continue

            if c == "(":
                depth += 1

            elif c == ")":
                depth -= 1

                if (
                    depth == 0
                    and i != len(expr) - 1
                ):
                    good = False
                    break

        if not good:
            break

        expr = expr[1:-1].strip()

    return expr


# ============================================================
# Analyzer
# ============================================================

class Analyzer:

    def __init__(self, source):
        self.source = Path(source)

        # Normal mode enables fast random cell access.
        self.wf = load_workbook(
            source,
            read_only=False,
            data_only=False,
            keep_links=False
        )

        self.wv = load_workbook(
            source,
            read_only=False,
            data_only=True,
            keep_links=False
        )

        self.names = list(self.wf.sheetnames)

        self.names_map = {
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
        self.macro_routes = []
        self.vba_refs = []
        self.debug = []

        self.seen = set()
        self.stats = collections.Counter()

        self.value_cache = {}
        self.evaluating = set()

        self.procs = collections.defaultdict(list)
        self.defined = {}

        for key, definition in (
            self.wf.defined_names.items()
        ):
            self.defined[key.casefold()] = s(
                definition.attr_text
            )

    def canon(self, name):
        return self.names_map.get(
            s(name)
            .strip()
            .replace("''", "'")
            .casefold()
        )

    # --------------------------------------------------------
    # Record dependency
    # --------------------------------------------------------

    def add(
        self,
        src,
        dst,
        kind,
        category,
        at,
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
            category,
            at,
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
                category,
                at,
                s(detail)[:12000]
            ))

    # --------------------------------------------------------
    # Resolve cell values
    # --------------------------------------------------------

    def getcell(
        self,
        sheet,
        coord,
        depth=0
    ):
        sheet = self.canon(sheet)

        if not sheet or depth > 16:
            return None

        key = (
            sheet,
            coord.replace("$", "").upper()
        )

        if key in self.value_cache:
            return self.value_cache[key]

        if key in self.evaluating:
            return None

        self.evaluating.add(key)

        try:
            value = self.wv[sheet][key[1]].value

            if value is None:
                value = self.wf[sheet][key[1]].value

                if (
                    isinstance(value, str)
                    and value.startswith("=")
                ):
                    value = self.eval(
                        value[1:],
                        sheet,
                        key[1],
                        depth + 1
                    )

            if len(self.value_cache) < 100000:
                self.value_cache[key] = value

            return value

        except Exception:
            return None

        finally:
            self.evaluating.discard(key)

    # --------------------------------------------------------
    # Evaluate simple Excel expressions
    # --------------------------------------------------------

    def eval(
        self,
        expr,
        sheet,
        cell,
        depth=0
    ):
        if depth > 16:
            return None

        expr = strip_outer(
            s(expr).strip()
        )

        if not expr:
            return None

        # String concatenation:
        # "'" & $E36 & "'!$H$119:$W$123"
        parts = split(expr, "&")

        if len(parts) > 1:
            values = [
                self.eval(
                    part,
                    sheet,
                    cell,
                    depth + 1
                )
                for part in parts
            ]

            if any(
                value is None
                for value in values
            ):
                return None

            return "".join(
                str(value)
                for value in values
            )

        # String literal
        if (
            expr.startswith('"')
            and expr.endswith('"')
        ):
            return expr[1:-1].replace(
                '""',
                '"'
            )

        # Numeric literal
        if re.fullmatch(
            r"[+-]?\d+(?:\.\d+)?",
            expr
        ):
            if "." in expr:
                return float(expr)

            return int(expr)

        # Cell reference
        match = SIMPLE_CELL.fullmatch(expr)

        if match:
            target_sheet = (
                match.group(1)
                or match.group(2)
                or sheet
            )

            return self.getcell(
                target_sheet,
                match.group(3),
                depth + 1
            )

        # Defined names
        if expr.casefold() in self.defined:
            return self.eval(
                self.defined[expr.casefold()],
                sheet,
                cell,
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
            "TEXT",
            "ADDRESS"
        )

        for fname in supported:
            if not re.match(
                r"^" + fname + r"\s*\(",
                expr,
                re.I
            ):
                continue

            found = list(
                functions(expr, fname)
            )

            if not found:
                continue

            args = split(found[0])

            values = [
                self.eval(
                    arg,
                    sheet,
                    cell,
                    depth + 1
                )
                for arg in args
            ]

            if (
                fname == "IFERROR"
                and len(values) >= 2
            ):
                return (
                    values[0]
                    if values[0] is not None
                    else values[1]
                )

            if (
                fname == "IF"
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
                fname == "SUBSTITUTE"
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
                fname == "TRIM"
                and values
                and values[0] is not None
            ):
                return " ".join(
                    str(values[0]).split()
                )

            if (
                fname in ("UPPER", "LOWER")
                and values
                and values[0] is not None
            ):
                if fname == "UPPER":
                    return str(values[0]).upper()

                return str(values[0]).lower()

            if (
                fname in (
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
                fname == "TEXT"
                and values
                and values[0] is not None
            ):
                return str(values[0])

            return None

        return None

    # --------------------------------------------------------
    # Resolve target sheet from a link
    # --------------------------------------------------------

    def target(self, value):
        value = (
            s(value)
            .strip()
            .lstrip("=#")
        )

        match = re.search(
            r"(?:'((?:[^']|'')+)'|"
            r"([\w.\u0080-\uffff]+))!",
            value
        )

        if match:
            return self.canon(
                match.group(1)
                or match.group(2)
            )

        if value.casefold() in self.defined:
            return self.target(
                self.defined[
                    value.casefold()
                ]
            )

        return None

    # --------------------------------------------------------
    # Formula scanning
    # --------------------------------------------------------

    def scan_formulas(self):
        print(
            "[1/5] Formula scan",
            flush=True
        )

        for ws in self.wf.worksheets:
            count = 0

            for row in ws:
                for cell in row:
                    formula = cell.value

                    if (
                        cell.data_type != "f"
                        and not (
                            isinstance(formula, str)
                            and formula.startswith("=")
                        )
                    ):
                        continue

                    formula = s(formula)

                    count += 1
                    self.stats["Formulas"] += 1

                    # Direct cross-sheet references
                    for match in SHEET_REF.finditer(
                        unquote_formula(formula)
                    ):
                        target_sheet = (
                            match.group(1)
                            or match.group(2)
                        )

                        self.add(
                            ws.title,
                            target_sheet,
                            "DATA",
                            "DIRECT",
                            cell.coordinate,
                            formula
                        )

                    # INDIRECT
                    for expression in functions(
                        formula,
                        "INDIRECT"
                    ):
                        self.stats["INDIRECT"] += 1

                        args = split(expression)

                        value = (
                            self.eval(
                                args[0],
                                ws.title,
                                cell.coordinate
                            )
                            if args
                            else None
                        )

                        target_sheet = self.target(
                            value
                        )

                        if target_sheet:
                            self.add(
                                ws.title,
                                target_sheet,
                                "DATA",
                                "INDIRECT",
                                cell.coordinate,
                                (
                                    formula
                                    + " | resolved="
                                    + s(value)
                                )
                            )

                        elif (
                            value is not None
                            and re.fullmatch(
                                r"\$?[A-Z]{1,3}\$?\d+"
                                r"(?::\$?[A-Z]{1,3}"
                                r"\$?\d+)?",
                                s(value),
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
                                "Evaluated: " + repr(value)
                            ))

                            if ws.title == "サマリ":
                                self.debug.append((
                                    ws.title,
                                    cell.coordinate,
                                    "INDIRECT",
                                    (
                                        args[0]
                                        if args
                                        else ""
                                    ),
                                    repr(value),
                                    repr(
                                        self.getcell(
                                            ws.title,
                                            "E36"
                                        )
                                    )
                                ))

                    # OFFSET
                    for expression in functions(
                        formula,
                        "OFFSET"
                    ):
                        self.stats["OFFSET"] += 1

                        if not SHEET_REF.search(
                            unquote_formula(expression)
                        ):
                            self.unresolved.append((
                                ws.title,
                                cell.coordinate,
                                "OFFSET",
                                formula,
                                (
                                    "Dynamic/local OFFSET; "
                                    "no confirmed cross-sheet "
                                    "target"
                                )
                            ))

                    # HYPERLINK formula
                    for expression in functions(
                        formula,
                        "HYPERLINK"
                    ):
                        args = split(expression)

                        value = (
                            self.eval(
                                args[0],
                                ws.title,
                                cell.coordinate
                            )
                            if args
                            else None
                        )

                        target_sheet = self.target(
                            value
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

        # Special regression diagnostics
        if self.canon("サマリ"):
            e36 = self.getcell(
                "サマリ",
                "E36"
            )

            self.debug.insert(0, (
                "サマリ",
                "E36",
                "ACTUAL_VALUE",
                "",
                repr(e36),
                (
                    "Expected sheet: "
                    + s(self.canon(e36))
                )
            ))

    # --------------------------------------------------------
    # Shape recording
    # --------------------------------------------------------

    def shape(
        self,
        sheet,
        kind,
        ident,
        name,
        link,
        macro,
        part,
        note=""
    ):
        self.shapes.append((
            sheet,
            kind,
            s(ident),
            s(name),
            s(link),
            s(macro),
            part,
            s(note)
        ))

        for linkpart in s(link).split(";"):
            target_sheet = self.target(
                linkpart
            )

            if target_sheet:
                self.add(
                    sheet,
                    target_sheet,
                    "NAV",
                    "SHAPE_HYPERLINK",
                    s(ident),
                    linkpart
                )

    # --------------------------------------------------------
    # XML relationships
    # --------------------------------------------------------

    def scan_xml(self):
        print(
            "[2/5] Drawing / VML / hyperlinks",
            flush=True
        )

        with zipfile.ZipFile(
            self.source
        ) as z:
            files = set(z.namelist())

            workbook_rels = xmlrels(
                z,
                "xl/workbook.xml"
            )

            root = ET.fromstring(
                z.read("xl/workbook.xml")
            )

            for sh in root.iter():
                if local(sh.tag) != "sheet":
                    continue

                sheet = sh.get(
                    "name",
                    ""
                )

                relation = workbook_rels.get(
                    sh.get(
                        "{" + REL_NS + "}id"
                    )
                )

                if not relation:
                    continue

                part = partpath(
                    "xl/workbook.xml",
                    relation[1]
                )

                if part not in files:
                    continue

                sheet_rels = xmlrels(
                    z,
                    part
                )

                root2 = ET.fromstring(
                    z.read(part)
                )

                # Worksheet hyperlinks / controls
                for el in root2.iter():
                    tag = local(el.tag)

                    if tag == "hyperlink":
                        relation = sheet_rels.get(
                            el.get(
                                "{" + REL_NS + "}id"
                            )
                        )

                        link = (
                            el.get("location", "")
                            or (
                                relation[1]
                                if relation
                                else ""
                            )
                        )

                        self.shape(
                            sheet,
                            "CELL_LINK",
                            el.get("ref", ""),
                            el.get("display", ""),
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
                            el.get("shapeId", ""),
                            el.get("name", ""),
                            "",
                            el.get("macro"),
                            part
                        )

                done = set()

                for rid, relation in (
                    sheet_rels.items()
                ):
                    kind, target, mode = relation

                    if kind not in (
                        "drawing",
                        "vmlDrawing",
                        "ctrlProp",
                        "control"
                    ):
                        continue

                    child = partpath(
                        part,
                        target
                    )

                    if (
                        child in done
                        or child not in files
                    ):
                        continue

                    done.add(child)

                    try:
                        if kind == "drawing":
                            self.drawing(
                                z,
                                sheet,
                                child
                            )

                        elif kind == "vmlDrawing":
                            self.vml(
                                z,
                                sheet,
                                child
                            )

                        else:
                            self.control(
                                z,
                                sheet,
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
    # DrawingML
    # --------------------------------------------------------

    def drawing(
        self,
        z,
        sheet,
        part
    ):
        root = ET.fromstring(
            z.read(part)
        )

        drawing_rels = xmlrels(
            z,
            part
        )

        for el in root.iter():
            if local(el.tag) not in (
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
            notes = []

            for x in el.iter():
                tag = local(x.tag)

                if tag == "cNvPr":
                    ident = x.get(
                        "id",
                        ident
                    )

                    name = x.get(
                        "name",
                        name
                    )

                    description = x.get(
                        "descr",
                        ""
                    )

                    if description:
                        notes.append(description)

                macro = (
                    macro
                    or x.get("macro", "")
                    or x.get("onAction", "")
                )

                if tag in (
                    "hlinkClick",
                    "hlinkHover"
                ):
                    relation = drawing_rels.get(
                        x.get(
                            "{" + REL_NS + "}id"
                        )
                    )

                    if relation:
                        # Preserve #Sheet!A1 literally.
                        links.append(
                            relation[1]
                        )

                    if x.get("action"):
                        links.append(
                            x.get("action")
                        )

                    if x.get("location"):
                        links.append(
                            x.get("location")
                        )

            self.shape(
                sheet,
                local(el.tag),
                ident,
                name,
                ";".join(links),
                macro,
                part,
                ";".join(notes)
            )

    # --------------------------------------------------------
    # VML
    # --------------------------------------------------------

    def vml(
        self,
        z,
        sheet,
        part
    ):
        root = ET.fromstring(
            z.read(part)
        )

        drawing_rels = xmlrels(
            z,
            part
        )

        for el in root.iter():
            if local(el.tag) != "shape":
                continue

            ident = el.get(
                "id",
                ""
            )

            macro = el.get(
                "macro",
                ""
            )

            links = []
            notes = []

            kind = "VML"

            for x in el.iter():
                tag = local(x.tag)

                if (
                    tag in (
                        "FmlaMacro",
                        "Macro"
                    )
                    and s(x.text).strip()
                ):
                    macro = s(x.text).strip()

                if tag == "ClientData":
                    kind = (
                        "VML:"
                        + x.get(
                            "ObjectType",
                            ""
                        )
                    )

                if tag in (
                    "FmlaLink",
                    "FmlaRange"
                ):
                    notes.append(
                        tag + ":" + s(x.text)
                    )

                if tag in (
                    "hlinkClick",
                    "hlinkHover"
                ):
                    relation = drawing_rels.get(
                        x.get(
                            "{" + REL_NS + "}id"
                        )
                    )

                    if relation:
                        links.append(
                            relation[1]
                        )

                    if x.get("action"):
                        links.append(
                            x.get("action")
                        )

            self.shape(
                sheet,
                kind,
                ident,
                ident,
                ";".join(links),
                macro,
                part,
                ";".join(notes)
            )

    # --------------------------------------------------------
    # Control XML
    # --------------------------------------------------------

    def control(
        self,
        z,
        sheet,
        part
    ):
        root = ET.fromstring(
            z.read(part)
        )

        for el in root.iter():
            macro = (
                el.get("macro")
                or el.get("fmlaMacro")
                or el.get("onAction")
            )

            if macro:
                self.shape(
                    sheet,
                    "CONTROL",
                    el.get("id", ""),
                    el.get("name", ""),
                    "",
                    macro,
                    part
                )

    # --------------------------------------------------------
    # VBA analysis
    # --------------------------------------------------------

    def scan_vba(self):
        print(
            "[3/5] VBA macro analysis",
            flush=True
        )

        try:
            from oletools.olevba import (
                VBA_Parser
            )

        except ImportError:
            self.unresolved.append((
                "Workbook",
                "VBA",
                "VBA",
                "",
                (
                    "Install oletools: "
                    "pip install oletools"
                )
            ))
            return

        parser = None

        try:
            parser = VBA_Parser(
                str(self.source)
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
                        VBA_PROC.finditer(code)
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

                        name = match.group(1)

                        self.procs[
                            name.casefold()
                        ].append((
                            module,
                            body
                        ))

                        for ref in (
                            VBA_SHEET.finditer(body)
                        ):
                            self.vba_refs.append((
                                module,
                                name,
                                ref.group(1),
                                body[:700]
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

        # Trace each shape's assigned macro.
        for shape in self.shapes:
            (
                sheet,
                kind,
                ident,
                name,
                link,
                macro,
                part,
                note
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

            visited = set()

            destinations = (
                collections.defaultdict(set)
            )

            found = []

            def walk(procname, depth=0):
                if (
                    depth > 10
                    or procname in visited
                ):
                    return

                visited.add(procname)

                for module, body in (
                    self.procs.get(
                        procname,
                        []
                    )
                ):
                    found.append(
                        module + ":" + procname
                    )

                    for line in body.splitlines():
                        if line.lstrip().startswith("'"):
                            continue

                        refs = list(
                            VBA_SHEET.finditer(line)
                        )

                        if refs:
                            navigation = bool(
                                VBA_GOTO.search(line)
                            )

                            for ref in refs:
                                target_sheet = (
                                    self.canon(
                                        ref.group(1)
                                    )
                                )

                                if target_sheet:
                                    destinations[
                                        target_sheet
                                    ].add(
                                        "NAV"
                                        if navigation
                                        else "REFERENCE"
                                    )

                        # Follow simple procedure calls.
                        stripped = (
                            line.strip()
                            .split("'", 1)[0]
                        )

                        match = re.match(
                            r"(?i)^(?:Call\s+)?"
                            r"([A-Za-z_]\w*)\b",
                            stripped
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

            if not found:
                self.macro_routes.append((
                    sheet,
                    ident,
                    name,
                    macro,
                    "",
                    "UNRESOLVED",
                    "Macro procedure not found",
                    ""
                ))
                continue

            if not destinations:
                self.macro_routes.append((
                    sheet,
                    ident,
                    name,
                    macro,
                    "",
                    "UNRESOLVED",
                    (
                        "No static "
                        'Sheets("...") target'
                    ),
                    ";".join(found)
                ))
                continue

            for target_sheet, types in sorted(
                destinations.items()
            ):
                if "NAV" in types:
                    status = "POSSIBLE_NAV"

                    self.add(
                        sheet,
                        target_sheet,
                        "NAV",
                        "VBA_NAV_CANDIDATE",
                        ident,
                        (
                            macro
                            + " via "
                            + ";".join(found)
                        ),
                        False
                    )

                else:
                    status = (
                        "SHEET_REFERENCE_ONLY"
                    )

                self.macro_routes.append((
                    sheet,
                    ident,
                    name,
                    macro,
                    target_sheet,
                    status,
                    ",".join(sorted(types)),
                    ";".join(found)
                ))

    # --------------------------------------------------------
    # Streaming report
    # --------------------------------------------------------

    def write(self, output):
        print(
            "[4/5] Streaming report",
            flush=True
        )

        self.maybe_data -= self.data
        self.maybe_nav -= self.nav

        wb = Workbook(
            write_only=True
        )

        def tab(name, headers, rows):
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
                f"{count:,} rows "
                f"({elapsed:.1f}s)",
                flush=True
            )

        # Summary
        tab(
            "Summary",
            ["Metric", "Value"],
            [
                (
                    "Source",
                    str(self.source)
                ),
                (
                    "Sheets",
                    len(self.names)
                ),
                (
                    "Formulas",
                    self.stats["Formulas"]
                ),
                (
                    "INDIRECT",
                    self.stats["INDIRECT"]
                ),
                (
                    "OFFSET",
                    self.stats["OFFSET"]
                ),
                (
                    "Data confirmed pairs",
                    len(self.data)
                ),
                (
                    "Data possible pairs",
                    len(self.maybe_data)
                ),
                (
                    "Navigation confirmed pairs",
                    len(self.nav)
                ),
                (
                    "Navigation possible pairs",
                    len(self.maybe_nav)
                ),
                (
                    "Shape records",
                    len(self.shapes)
                ),
                (
                    "Macro routes",
                    len(self.macro_routes)
                ),
                (
                    "Unresolved records",
                    len(self.unresolved)
                ),
                (
                    "Evidence records",
                    len(self.evidence)
                )
            ]
        )

        # Sheet inventory
        tab(
            "Sheets",
            [
                "Sheet",
                "Visibility"
            ],
            (
                (
                    name,
                    self.wf[name].sheet_state
                )
                for name in self.names
            )
        )

        # Matrix generator
        def matrix(
            name,
            confirmed,
            possible
        ):
            def rows():
                for src in self.names:
                    values = [src]

                    for dst in self.names:
                        pair = (src, dst)

                        if pair in confirmed:
                            values.append(1)

                        elif pair in possible:
                            values.append("?")

                        else:
                            values.append(0)

                    yield values

            tab(
                name,
                ["Source / Target"]
                + self.names,
                rows()
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

        union = (
            self.data | self.nav
        )

        possible_union = (
            self.maybe_data
            | self.maybe_nav
        ) - union

        matrix(
            "Combined_Matrix",
            union,
            possible_union
        )

        # Evidence
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

        # Unresolved
        tab(
            "Unresolved",
            [
                "Sheet",
                "Cell/Object",
                "Type",
                "Formula",
                "Reason"
            ],
            self.unresolved
        )

        # Shape links
        tab(
            "Shapes_Links",
            [
                "Sheet",
                "ObjectType",
                "ID",
                "Name",
                "Link",
                "Macro",
                "Part",
                "Detail"
            ],
            self.shapes
        )

        # VBA macro routes
        tab(
            "Macro_Routes",
            [
                "Source",
                "ID",
                "Name",
                "Macro",
                "Target",
                "Status",
                "Reason",
                "Procedures"
            ],
            self.macro_routes
        )

        # VBA references
        tab(
            "VBA_References",
            [
                "Module",
                "Procedure",
                "Sheet",
                "Code excerpt"
            ],
            self.vba_refs
        )

        # INDIRECT diagnostics
        tab(
            "INDIRECT_Debug",
            [
                "Sheet",
                "Cell",
                "Type",
                "Expression",
                "Value",
                "E36 value"
            ],
            self.debug
        )

        # TOP button audit
        def top_rows():
            for shape in self.shapes:
                if shape[0].casefold() != "top":
                    continue

                routes = [
                    route
                    for route in self.macro_routes
                    if (
                        route[0] == shape[0]
                        and route[1] == shape[2]
                    )
                ]

                targets = sorted({
                    route[4]
                    for route in routes
                    if route[4]
                })

                statuses = sorted({
                    route[5]
                    for route in routes
                })

                if not statuses:
                    statuses = [
                        "LINK"
                        if shape[4]
                        else "UNRESOLVED"
                    ]

                yield (
                    shape[2],
                    shape[3],
                    shape[1],
                    shape[5],
                    shape[4],
                    "; ".join(targets),
                    "; ".join(statuses),
                    shape[6]
                )

        tab(
            "TOP_Button_Audit",
            [
                "Object ID",
                "Name",
                "Type",
                "Macro",
                "Link",
                "Targets",
                "Status",
                "Part"
            ],
            top_rows()
        )

        # Regression checks
        src = self.canon("サマリ")

        dst = self.canon(
            "1023951_I_STAR現行売上"
        )

        pair = (src, dst)

        if pair in self.data:
            check = "PASS"

        elif pair in self.maybe_data:
            check = "POSSIBLE"

        else:
            check = "NOT_RESOLVED"

        top = self.canon("TOP")

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
                    "Data reference"
                ),
                (
                    "サマリ!E36",
                    repr(
                        self.getcell(
                            "サマリ",
                            "E36"
                        )
                    ),
                    "Resolved source value"
                ),
                (
                    "TOP shapes",
                    sum(
                        shape[0] == top
                        for shape in self.shapes
                    ),
                    "Drawing/VML objects"
                ),
                (
                    "TOP macro bindings",
                    sum(
                        shape[0] == top
                        and bool(shape[5])
                        for shape in self.shapes
                    ),
                    "Bound macros"
                ),
                (
                    "TOP confirmed navigation",
                    sum(
                        source == top
                        for source, target
                        in self.nav
                    ),
                    "Direct hyperlinks"
                ),
                (
                    "TOP possible navigation",
                    sum(
                        source == top
                        for source, target
                        in self.maybe_nav
                    ),
                    "Macro candidates"
                )
            ]
        )

        # README
        tab(
            "README",
            [
                "Field",
                "Meaning"
            ],
            [
                (
                    "1",
                    "Confirmed formula reference "
                    "or direct hyperlink"
                ),
                (
                    "?",
                    "Possible relationship; "
                    "not verified execution"
                ),
                (
                    "0",
                    "No evidence found; "
                    "NOT proof of absence"
                ),
                (
                    "Data_Matrix",
                    "Formula dependencies"
                ),
                (
                    "Navigation_Matrix",
                    "Hyperlinks and macro "
                    "navigation candidates"
                ),
                (
                    "Combined_Matrix",
                    "Data and navigation"
                ),
                (
                    "TOP_Button_Audit",
                    "Detected TOP objects "
                    "and assigned macros"
                ),
                (
                    "INDIRECT_Debug",
                    "Dynamic reference diagnostics"
                ),
                (
                    "Source workbook",
                    "Never overwritten"
                ),
                (
                    "VBA",
                    "Static inspection only"
                )
            ]
        )

        # Save
        print(
            "[5/5] Saving XLSX...",
            flush=True
        )

        start = time.perf_counter()

        wb.save(output)

        print(
            f"[DONE] {output} "
            f"({time.perf_counter() - start:.1f}s)",
            flush=True
        )

        return check

    def close(self):
        self.wf.close()
        self.wv.close()


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
            "Invalid XLSX/XLSM input"
        )

    if args.output:
        output = Path(
            args.output
        ).resolve()

    else:
        output = source.with_name(
            source.stem
            + "_dependency_report_v6.xlsx"
        )

    if output == source:
        parser.error(
            "Output must not equal input"
        )

    print(
        "Input:",
        source,
        flush=True
    )

    print(
        "Output:",
        output,
        flush=True
    )

    start = time.perf_counter()

    audit = Analyzer(source)

    try:
        audit.scan_formulas()
        audit.scan_xml()
        audit.scan_vba()

        output.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        status = audit.write(output)

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
            f"Regression: {status}"
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
            "Interrupted",
            file=sys.stderr
        )
        sys.exit(130)
