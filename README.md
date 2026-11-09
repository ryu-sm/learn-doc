
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel XLSM Dependency Analyzer V2

Features:
- All worksheets including hidden / veryHidden
- Direct worksheet references
- INDIRECT / INDEX / OFFSET analysis
- Dynamic sheet names from cells and merged cells
- VBA procedures and navigation
- Drawing / Form button macro bindings
- External workbook references
- Three-state dependency matrix: 1 / 0 / ?
- Evidence and unresolved details
- Read-only source processing

Usage:
    python analyze_xlsm.py input.xlsm

Optional:
    python analyze_xlsm.py input.xlsm --com

Output:
    input_dependency_report.xlsx
"""

import argparse
import os
import re
import sys
import zipfile
import posixpath
import tempfile
from collections import defaultdict, Counter
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, range_boundaries


# ============================================================
# Configuration
# ============================================================

MAX_EVIDENCE = 500000
MAX_UNRESOLVED = 100000
MAX_FORMULA_LENGTH = 32000
MAX_XML_TEXT = 32000

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}

M = "{" + NS["m"] + "}"

DIRECT_REF = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(?:"
    r"'((?:[^']|'')+)'"
    r"|"
    r"([A-Za-z_\u0080-\uffff][A-Za-z0-9_.\u0080-\uffff]*)"
    r")!"
    r"(?=\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|"
    r"\$?\d+:\$?\d+)",
    re.I,
)

EXCEL_CELL = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$", re.I)

VBA_PROC = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
    r"(?:Sub|Function)\s+([^\s(]+)",
    re.I | re.M,
)

VBA_SHEET = re.compile(
    r"(?:ThisWorkbook\.)?"
    r"(?:Sheets|Worksheets)\s*\(\s*"
    r'(?:"([^"]+)"|\'([^\']+)\')\s*\)',
    re.I,
)

VBA_NAV = re.compile(
    r"\.\s*(?:Activate|Select)\b",
    re.I,
)

INDIRECT_RE = re.compile(r"\bINDIRECT\s*\(", re.I)
OFFSET_RE = re.compile(r"\bOFFSET\s*\(", re.I)

DYNAMIC_FUNCS = (
    "INDIRECT(",
    "OFFSET(",
)

ERROR_MARKER = object()


# ============================================================
# Utilities
# ============================================================

def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def safe_text(value):
    if value is None:
        return ""
    return str(value)[:MAX_XML_TEXT]


def sheet_key(name):
    return str(name).strip().casefold()


def split_external_sheet(token):
    """
    Returns:
        (external_book, sheet_name)
    """
    token = token.replace("''", "'")

    if "]" in token and "[" in token:
        left = token.find("[")
        right = token.find("]", left)
        if right > left:
            return token[left + 1:right], token[right + 1:]

    return "", token


def normalize_sheet(token, sheet_lookup):
    book, name = split_external_sheet(token)
    if book:
        return None, book
    return sheet_lookup.get(sheet_key(name)), ""


def quote_sheet(name):
    return "'" + name.replace("'", "''") + "'"


def normalize_target(path, target):
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return posixpath.normpath(target)
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(path), target)
    )


def extract_function_args(formula, function_name):
    """
    Extract top-level arguments from every occurrence of a function.
    Handles nested parentheses and quoted strings.
    """
    pattern = re.compile(
        r"\b" + re.escape(function_name) + r"\s*\(",
        re.I,
    )

    results = []

    for match in pattern.finditer(formula):
        start = match.end()
        depth = 1
        quoted = False
        args = []
        buf = []
        i = start

        while i < len(formula):
            ch = formula[i]

            if ch == '"':
                if quoted and i + 1 < len(formula) and formula[i + 1] == '"':
                    buf.append('""')
                    i += 2
                    continue
                quoted = not quoted

            if not quoted:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        args.append("".join(buf).strip())
                        results.append(args)
                        break
                elif ch in ",;" and depth == 1:
                    args.append("".join(buf).strip())
                    buf = []
                    i += 1
                    continue

            buf.append(ch)
            i += 1

    return results


def split_concat(expr):
    parts = []
    buf = []
    depth = 0
    quoted = False
    i = 0

    while i < len(expr):
        ch = expr[i]

        if ch == '"':
            if quoted and i + 1 < len(expr) and expr[i + 1] == '"':
                buf.append('""')
                i += 2
                continue
            quoted = not quoted

        if not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "&" and depth == 0:
                parts.append("".join(buf).strip())
                buf = []
                i += 1
                continue

        buf.append(ch)
        i += 1

    parts.append("".join(buf).strip())
    return parts


def is_balanced_parentheses(expr):
    depth = 0
    quoted = False
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
                if depth < 0:
                    return False
        i += 1

    return depth == 0 and not quoted


def strip_outer_parens(expr):
    expr = expr.strip()

    while (
        expr.startswith("(")
        and expr.endswith(")")
        and is_balanced_parentheses(expr[1:-1])
    ):
        expr = expr[1:-1].strip()

    return expr


def col_number(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def cell_row(address):
    m = EXCEL_CELL.match(address)
    return int(m.group(2)) if m else None


def cell_col(address):
    m = EXCEL_CELL.match(address)
    return col_number(m.group(1)) if m else None


def extract_sheet_candidates(text, lookup):
    """
    Conservative candidate extraction from quoted strings.
    Candidates are NOT automatically confirmed.
    """
    found = set()

    for name in lookup.values():
        if re.search(
            r"(?<![A-Za-z0-9_])" + re.escape(name) + r"!",
            text,
            re.I,
        ):
            found.add(name)

    return found


# ============================================================
# XLSM Reader (ZIP + XML)
# ============================================================

class XLSMReader:

    def __init__(self, filename):
        self.filename = filename
        self.zf = zipfile.ZipFile(filename, "r")
        self.names = set(self.zf.namelist())

        self.shared_strings = []
        self.sheets = []
        self.sheet_paths = {}
        self.sheet_states = {}
        self.sheet_lookup = {}
        self.cached = defaultdict(dict)
        self.formulas = defaultdict(dict)
        self.merged = defaultdict(list)
        self.defined_names = []
        self.external_links = []
        self.sheet_rels = {}

        self._load_shared_strings()
        self._load_workbook()

    def xml(self, path):
        if path not in self.names:
            return None
        return ET.fromstring(self.zf.read(path))

    def _load_shared_strings(self):
        root = self.xml("xl/sharedStrings.xml")
        if root is None:
            return

        for si in root.findall(M + "si"):
            pieces = [
                t.text or ""
                for t in si.iter(M + "t")
            ]
            self.shared_strings.append("".join(pieces))

    def _load_workbook(self):
        root = self.xml("xl/workbook.xml")
        relroot = self.xml("xl/_rels/workbook.xml.rels")

        if root is None:
            raise ValueError("Invalid XLSM: workbook.xml missing")

        relationships = {}

        if relroot is not None:
            for rel in relroot:
                relationships[rel.attrib.get("Id")] = rel.attrib.get(
                    "Target", ""
                )

        sheets_node = root.find(M + "sheets")

        if sheets_node is None:
            raise ValueError("No worksheets found")

        for item in sheets_node:
            name = item.attrib.get("name", "")
            state = item.attrib.get("state", "visible")
            rid = item.attrib.get("{" + NS["r"] + "}id")
            target = relationships.get(rid, "")

            path = normalize_target("xl/workbook.xml", target)

            if "worksheets/" not in path:
                continue

            self.sheets.append(name)
            self.sheet_paths[name] = path
            self.sheet_states[name] = state
            self.sheet_lookup[sheet_key(name)] = name

        dns = root.find(M + "definedNames")

        if dns is not None:
            for dn in dns:
                self.defined_names.append({
                    "name": dn.attrib.get("name", ""),
                    "scope": dn.attrib.get("localSheetId", ""),
                    "expression": dn.text or "",
                })

        for path in sorted(self.names):
            if path.startswith("xl/externalLinks/externalLink") and \
               path.endswith(".xml"):
                self.external_links.append(path)

    def load_cells(self, progress=None):
        count = 0

        for sheet in self.sheets:
            path = self.sheet_paths[sheet]

            if progress:
                progress("Sheet: " + sheet)

            with self.zf.open(path) as stream:
                context = ET.iterparse(stream, events=("end",))

                for _, elem in context:
                    tag = local_name(elem.tag)

                    if tag == "c":
                        address = elem.attrib.get("r", "")
                        dtype = elem.attrib.get("t", "")

                        formula = elem.find(M + "f")
                        value = elem.find(M + "v")
                        inline = elem.find(M + "is")

                        if formula is not None:
                            ftext = formula.text or ""

                            # Shared formula followers are handled later.
                            if ftext:
                                self.formulas[sheet][address] = ftext

                        if value is not None:
                            raw = value.text or ""

                            if dtype == "s":
                                try:
                                    val = self.shared_strings[int(raw)]
                                except (ValueError, IndexError):
                                    val = raw
                            elif dtype == "b":
                                val = raw == "1"
                            elif dtype in ("str", "e"):
                                val = raw
                            else:
                                try:
                                    val = float(raw)
                                    if val.is_integer():
                                        val = int(val)
                                except (ValueError, OverflowError):
                                    val = raw

                            self.cached[sheet][address] = val

                        elif inline is not None:
                            text = "".join(
                                n.text or ""
                                for n in inline.iter(M + "t")
                            )
                            self.cached[sheet][address] = text

                        count += 1

                        if count % 10000 == 0 and progress:
                            progress(f"Processed cells: {count}")

                        elem.clear()

                    elif tag == "mergeCell":
                        ref = elem.attrib.get("ref", "")
                        if ref:
                            self.merged[sheet].append(ref)
                        elem.clear()

        return count

    def get_value(self, sheet, address):
        if not sheet or sheet not in self.sheet_paths:
            return None

        address = address.replace("$", "").upper()

        value = self.cached[sheet].get(address)

        if value is not None:
            return value

        # Merged cell fallback
        row = cell_row(address)
        col = cell_col(address)

        if row is None or col is None:
            return None

        for ref in self.merged[sheet]:
            try:
                min_c, min_r, max_c, max_r = range_boundaries(ref)
            except ValueError:
                continue

            if min_r <= row <= max_r and min_c <= col <= max_c:
                top_left = f"{get_column_letter(min_c)}{min_r}"
                return self.cached[sheet].get(top_left)

        return None

    def close(self):
        self.zf.close()


# ============================================================
# Dynamic Expression Resolver
# ============================================================

class ExpressionResolver:

    def __init__(self, reader):
        self.reader = reader
        self.com_values = {}

    def set_com_values(self, values):
        self.com_values = values

    def value(self, sheet, address):
        address = address.replace("$", "").upper()
        key = (sheet, address)

        if key in self.com_values:
            val = self.com_values[key]
            if val is not None:
                return val

        return self.reader.get_value(sheet, address)

    def eval(self, expr, sheet, address, depth=0):
        """
        Limited safe evaluator.

        Returns:
            Python scalar or None.

        Never executes arbitrary formula or VBA code.
        """
        if depth > 12:
            return None

        expr = strip_outer_parens(expr)

        if not expr:
            return None

        parts = split_concat(expr)

        if len(parts) > 1:
            values = []

            for part in parts:
                value = self.eval(part, sheet, address, depth + 1)

                if value is None:
                    return None

                values.append(str(value))

            return "".join(values)

        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1].replace('""', '"')

        if re.fullmatch(r"-?\d+(?:\.\d+)?", expr):
            return float(expr) if "." in expr else int(expr)

        # Cell reference, optionally qualified by sheet.
        m = re.fullmatch(
            r"(?:(?:'((?:[^']|'')+)'|"
            r"([A-Za-z_\u0080-\uffff][\w.\u0080-\uffff]*))!)?"
            r"(\$?[A-Z]{1,3}\$?\d+)",
            expr,
            re.I,
        )

        if m:
            target = sheet

            if m.group(1) or m.group(2):
                token = m.group(1) or m.group(2)
                target, external = normalize_sheet(
                    token, self.reader.sheet_lookup
                )

                if external or not target:
                    return None

            return self.value(target, m.group(3))

        # ROW()
        if re.fullmatch(r"ROW\(\s*\)", expr, re.I):
            return cell_row(address)

        # COLUMN()
        if re.fullmatch(r"COLUMN\(\s*\)", expr, re.I):
            return cell_col(address)

        # ROW(A1), COLUMN(A1)
        m = re.fullmatch(
            r"(ROW|COLUMN)\(\s*(\$?[A-Z]{1,3}\$?\d+)\s*\)",
            expr,
            re.I,
        )

        if m:
            if m.group(1).upper() == "ROW":
                return cell_row(m.group(2))
            return cell_col(m.group(2))

        # SUBSTITUTE(text, old, new)
        for args in extract_function_args(expr, "SUBSTITUTE"):
            if len(args) >= 3 and expr.upper().startswith("SUBSTITUTE"):
                values = [
                    self.eval(a, sheet, address, depth + 1)
                    for a in args[:3]
                ]

                if all(v is not None for v in values):
                    return str(values[0]).replace(
                        str(values[1]), str(values[2])
                    )

        # ADDRESS(row, column, [abs_num])
        for args in extract_function_args(expr, "ADDRESS"):
            if len(args) >= 2 and expr.upper().startswith("ADDRESS"):
                row = self.eval(args[0], sheet, address, depth + 1)
                col = self.eval(args[1], sheet, address, depth + 1)

                if isinstance(row, (int, float)) and \
                   isinstance(col, (int, float)):
                    row, col = int(row), int(col)

                    if 1 <= row <= 1048576 and 1 <= col <= 16384:
                        return f"${get_column_letter(col)}${row}"

        return None

    def indirect_targets(self, formula, sheet, address):
        confirmed = set()
        candidates = set()
        reasons = []

        for args in extract_function_args(formula, "INDIRECT"):
            if not args:
                reasons.append("INDIRECT has no arguments")
                continue

            result = self.eval(args[0], sheet, address)

            if result is None:
                # Conservative extraction of candidate names.
                candidates.update(
                    extract_sheet_candidates(
                        args[0], self.reader.sheet_lookup
                    )
                )
                reasons.append(
                    "Cannot evaluate INDIRECT text argument"
                )
                continue

            text = str(result)

            # Examples:
            # 'Sheet Name'!$H$119:$W$123
            # Sheet1!A1
            m = re.match(
                r"^(?:'((?:[^']|'')+)'|([^!]+))!",
                text,
            )

            if m:
                token = m.group(1) or m.group(2)
                target, external = normalize_sheet(
                    token, self.reader.sheet_lookup
                )

                if target:
                    confirmed.add(target)
                elif external:
                    reasons.append(
                        "External INDIRECT reference: " + external
                    )
                else:
                    reasons.append(
                        "Unknown target sheet: " + token
                    )

            else:
                # INDIRECT("A1") refers to current sheet.
                if re.match(r"^\$?[A-Z]{1,3}\$?\d+", text, re.I):
                    confirmed.add(sheet)
                else:
                    reasons.append(
                        "Unsupported INDIRECT target: " + text[:200]
                    )

        return confirmed, candidates - confirmed, reasons


# ============================================================
# Dependency Analyzer
# ============================================================

class DependencyAnalyzer:

    def __init__(self, reader):
        self.reader = reader
        self.resolver = ExpressionResolver(reader)

        self.evidence = []
        self.unresolved = []
        self.external = []
        self.vba_rows = []
        self.button_rows = []

        self.confirmed = set()
        self.possible = set()

        self.stats = Counter()
        self.evidence_seen = set()
        self.unresolved_seen = set()

    def add_evidence(
        self, source, target, category, location,
        expression="", detail="", confidence="Confirmed"
    ):
        if source not in self.reader.sheets:
            return

        if target not in self.reader.sheets:
            return

        key = (
            source, target, category,
            location, expression[:500]
        )

        if key in self.evidence_seen:
            return

        self.evidence_seen.add(key)

        if confidence == "Confirmed":
            self.confirmed.add((source, target))
        else:
            self.possible.add((source, target))

        if len(self.evidence) < MAX_EVIDENCE:
            self.evidence.append([
                source, target, category, location,
                confidence, safe_text(expression),
                safe_text(detail),
            ])
        else:
            self.stats["Truncated Evidence"] += 1

    def add_unresolved(
        self, source, location, category,
        expression, reason, candidates=None
    ):
        candidates = set(candidates or [])

        for target in candidates:
            if target in self.reader.sheets:
                self.possible.add((source, target))

        key = (
            source, location, category,
            str(expression)[:500], reason
        )

        if key in self.unresolved_seen:
            return

        self.unresolved_seen.add(key)

        if len(self.unresolved) < MAX_UNRESOLVED:
            self.unresolved.append([
                source, location, category,
                safe_text(expression), safe_text(reason),
                ", ".join(sorted(candidates)),
            ])
        else:
            self.stats["Truncated Unresolved"] += 1

    def scan_direct(self, source, location, formula, category="Formula"):
        """
        Scan references outside Excel string literals.
        """
        masked = []
        quoted = False
        i = 0

        while i < len(formula):
            ch = formula[i]

            if ch == '"':
                if quoted and i + 1 < len(formula) and \
                   formula[i + 1] == '"':
                    masked.extend("  ")
                    i += 2
                    continue
                quoted = not quoted
                masked.append(" ")
            else:
                masked.append(" " if quoted else ch)

            i += 1

        clean = "".join(masked)

        for match in DIRECT_REF.finditer(clean):
            token = match.group(1) or match.group(2)

            # 3D reference: Sheet1:Sheet3!A1
            if ":" in token and "[" not in token:
                parts = token.split(":", 1)
                first = self.reader.sheet_lookup.get(sheet_key(parts[0]))
                last = self.reader.sheet_lookup.get(sheet_key(parts[1]))

                if first and last:
                    i1 = self.reader.sheets.index(first)
                    i2 = self.reader.sheets.index(last)

                    for target in self.reader.sheets[
                        min(i1, i2):max(i1, i2) + 1
                    ]:
                        self.add_evidence(
                            source, target, category,
                            location, formula, "3D reference"
                        )
                    continue

            target, external = normalize_sheet(
                token, self.reader.sheet_lookup
            )

            if target:
                self.add_evidence(
                    source, target, category,
                    location, formula, "Direct formula reference"
                )

            elif external:
                self.external.append([
                    source, location, external,
                    token, safe_text(formula),
                ])

            else:
                self.add_unresolved(
                    source, location, "UnknownSheet",
                    formula, "Referenced sheet not found: " + token
                )

    def scan_formulas(self):
        print("Scanning formulas...", flush=True)

        processed = 0

        for sheet in self.reader.sheets:
            formulas = self.reader.formulas[sheet]

            print(
                f"  {sheet}: {len(formulas)} formulas",
                flush=True
            )

            for address, formula in formulas.items():
                processed += 1
                self.stats["Formula Cells"] += 1

                self.scan_direct(sheet, address, formula)

                upper = formula.upper()

                if "INDIRECT(" in upper:
                    self.stats["INDIRECT Calls"] += len(
                        extract_function_args(formula, "INDIRECT")
                    )

                    confirmed, candidates, reasons = \
                        self.resolver.indirect_targets(
                            formula, sheet, address
                        )

                    for target in confirmed:
                        self.add_evidence(
                            sheet, target, "INDIRECT",
                            address, formula,
                            "Resolved dynamic sheet reference"
                        )

                    if reasons:
                        self.add_unresolved(
                            sheet, address, "INDIRECT",
                            formula, "; ".join(reasons),
                            candidates
                        )

                if "OFFSET(" in upper:
                    self.stats["OFFSET Calls"] += len(
                        extract_function_args(formula, "OFFSET")
                    )

                    # A direct reference inside OFFSET is already
                    # recorded by scan_direct.
                    # The actual offset destination may remain unknown.
                    self.add_unresolved(
                        sheet, address, "OFFSET",
                        formula,
                        "OFFSET destination not fully evaluated"
                    )

                if processed % 10000 == 0:
                    print(
                        f"  Processed formulas: {processed}",
                        flush=True
                    )

    def scan_defined_names(self):
        print("Scanning defined names...", flush=True)

        for dn in self.reader.defined_names:
            expression = dn["expression"]
            scope = dn["scope"]

            if scope != "":
                try:
                    source = self.reader.sheets[int(scope)]
                    sources = [source]
                except (ValueError, IndexError):
                    sources = []
            else:
                # Workbook-level names do not identify a single
                # source sheet. Preserve separately.
                sources = []

            for source in sources:
                self.scan_direct(
                    source,
                    "DefinedName:" + dn["name"],
                    expression,
                    "DefinedName",
                )

            if not sources:
                self.stats["Workbook Defined Names"] += 1

    def scan_external_links(self):
        print("Scanning external links...", flush=True)

        for path in self.reader.external_links:
            root = self.reader.xml(path)

            if root is None:
                continue

            self.external.append([
                "(Workbook)",
                path,
                "ExternalLinkPart",
                path,
                ET.tostring(root, encoding="unicode")[:MAX_XML_TEXT],
            ])

    def scan_shapes(self):
        """
        Scan drawing / VML / form control metadata.
        This is static XML extraction and does not execute macros.
        """
        print("Scanning shapes and buttons...", flush=True)

        paths = [
            p for p in self.reader.names
            if (
                p.startswith("xl/drawings/")
                or p.startswith("xl/ctrlProps/")
            ) and p.endswith(".xml")
        ]

        for path in paths:
            try:
                root = self.reader.xml(path)
            except ET.ParseError:
                continue

            if root is None:
                continue

            for elem in root.iter():
                attrs = elem.attrib

                macro = (
                    attrs.get("macro")
                    or attrs.get("fmlaMacro")
                    or attrs.get("onAction")
                    or attrs.get("OnAction")
                )

                if macro:
                    self.button_rows.append([
                        path,
                        local_name(elem.tag),
                        safe_text(macro),
                        "XML macro binding",
                    ])

        # VML is frequently used for legacy form buttons.
        for path in self.reader.names:
            if not (
                path.startswith("xl/drawings/")
                and path.endswith(".vml")
            ):
                continue

            try:
                root = self.reader.xml(path)
            except ET.ParseError:
                continue

            for elem in root.iter():
                if local_name(elem.tag).lower() in (
                    "fmlamacro", "macro"
                ):
                    macro = (elem.text or "").strip()

                    if macro:
                        self.button_rows.append([
                            path, "VML", macro,
                            "Legacy button macro binding"
                        ])

        self.stats["Macro Bindings"] = len(self.button_rows)

    def scan_vba(self):
        print("Scanning VBA...", flush=True)

        try:
            from oletools.olevba import VBA_Parser
        except ImportError:
            print(
                "  oletools unavailable; VBA scan skipped.",
                flush=True
            )
            return

        parser = None

        try:
            parser = VBA_Parser(self.reader.filename)

            if not parser.detect_vba_macros():
                return

            for _, stream, module, code in parser.extract_macros():
                self.stats["VBA Modules"] += 1

                procedures = list(VBA_PROC.finditer(code))

                self.stats["VBA Procedures"] += len(procedures)

                for index, match in enumerate(procedures):
                    start = match.start()
                    end = (
                        procedures[index + 1].start()
                        if index + 1 < len(procedures)
                        else len(code)
                    )

                    proc_name = match.group(1)
                    body = code[start:end]

                    targets = set()

                    for sm in VBA_SHEET.finditer(body):
                        name = sm.group(1) or sm.group(2)
                        target = self.reader.sheet_lookup.get(
                            sheet_key(name)
                        )

                        if target:
                            targets.add(target)

                    navigation = bool(VBA_NAV.search(body))

                    self.vba_rows.append([
                        module, proc_name,
                        ", ".join(sorted(targets)),
                        "Yes" if navigation else "No",
                        safe_text(body),
                    ])

                    # A VBA procedure referring to a sheet does
                    # not prove which worksheet invokes it.
                    # Therefore no matrix edge is created here.

        except Exception as exc:
            self.add_unresolved(
                "(Workbook)", "VBA", "VBA",
                "", str(exc)
            )
        finally:
            if parser is not None:
                parser.close()

    def finalize(self):
        # Confirmed edges always override possible edges.
        self.possible.difference_update(self.confirmed)

        self.stats["Confirmed Links"] = len(self.confirmed)
        self.stats["Possible Links"] = len(self.possible)
        self.stats["Evidence Rows"] = len(self.evidence)
        self.stats["Unresolved Rows"] = len(self.unresolved)
        self.stats["External Records"] = len(self.external)

    def run(self):
        self.scan_formulas()
        self.scan_defined_names()
        self.scan_external_links()
        self.scan_shapes()
        self.scan_vba()
        self.finalize()


# ============================================================
# Optional Excel COM cached-value assistance
# ============================================================

def load_com_values(filename, reader):
    """
    Open Excel in a separate process, read-only.

    Security:
    - Macros disabled where supported
    - Links not updated
    - No save operation
    - No formula recalculation requested

    Only retrieves values of cells referenced by
    INDIRECT text expressions.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("pywin32 not installed; skipping COM.")
        return {}

    print("Loading optional Excel COM values...", flush=True)

    values = {}
    excel = None
    book = None

    pythoncom.CoInitialize()

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        try:
            excel.Calculation = -4135  # Manual
        except Exception:
            pass

        book = excel.Workbooks.Open(
            os.path.abspath(filename),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        required = defaultdict(set)

        # Gather simple cell references used in INDIRECT arguments.
        for sheet, formulas in reader.formulas.items():
            for address, formula in formulas.items():
                if "INDIRECT(" not in formula.upper():
                    continue

                for args in extract_function_args(formula, "INDIRECT"):
                    if not args:
                        continue

                    for match in re.finditer(
                        r"(?<![A-Za-z0-9_.!])"
                        r"\$?[A-Z]{1,3}\$?\d+",
                        args[0],
                        re.I,
                    ):
                        required[sheet].add(
                            match.group(0).replace("$", "").upper()
                        )

        for sheet, addresses in required.items():
            print(
                f"  COM: {sheet}, {len(addresses)} cells",
                flush=True
            )

            try:
                ws = book.Worksheets(sheet)
            except Exception:
                continue

            for address in addresses:
                try:
                    cell = ws.Range(address)

                    if cell.MergeCells:
                        cell = cell.MergeArea.Cells(1, 1)

                    value = cell.Value2

                    if isinstance(value, (str, int, float, bool)):
                        values[(sheet, address)] = value

                except Exception:
                    continue

    except Exception as exc:
        print("COM warning:", exc, flush=True)

    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()

    return values


# ============================================================
# Report Writer
# ============================================================

def write_table(ws, headers, rows):
    ws.append(headers)

    for row in rows:
        ws.append([
            safe_text(v) if isinstance(v, str) else v
            for v in row
        ])

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid", fgColor="17365D"
        )
        cell.font = Font(
            bold=True, color="FFFFFF"
        )
        cell.alignment = Alignment(
            horizontal="center"
        )

    for i, header in enumerate(headers, start=1):
        width = 22

        if header in (
            "Expression", "Formula", "Reason",
            "VBA Code", "Detail"
        ):
            width = 65
        elif header in ("Source", "Target", "Sheet"):
            width = 32

        ws.column_dimensions[
            get_column_letter(i)
        ].width = width


def create_report(reader, analyzer, output):
    print("Generating report...", flush=True)

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    metrics = [
        ["Source File", os.path.basename(reader.filename)],
        ["Total Sheets", len(reader.sheets)],
        ["Hidden Sheets", sum(
            reader.sheet_states[s] != "visible"
            for s in reader.sheets
        )],
        ["Formula Cells", analyzer.stats["Formula Cells"]],
        ["Confirmed Links", len(analyzer.confirmed)],
        ["Possible Links", len(analyzer.possible)],
        ["INDIRECT Calls", analyzer.stats["INDIRECT Calls"]],
        ["OFFSET Calls", analyzer.stats["OFFSET Calls"]],
        ["VBA Modules", analyzer.stats["VBA Modules"]],
        ["VBA Procedures", analyzer.stats["VBA Procedures"]],
        ["Macro Bindings", analyzer.stats["Macro Bindings"]],
        ["External Records", len(analyzer.external)],
        ["Evidence Rows", len(analyzer.evidence)],
        ["Unresolved Rows", len(analyzer.unresolved)],
        ["Truncated Evidence", analyzer.stats["Truncated Evidence"]],
        ["Truncated Unresolved", analyzer.stats["Truncated Unresolved"]],
    ]

    write_table(summary, ["Metric", "Value"], metrics)

    # --------------------------------------------------------
    # Sheet inventory
    # --------------------------------------------------------

    ws = wb.create_sheet("Sheets")

    write_table(
        ws,
        ["Sheet", "Visibility", "Formula Count", "XML Path"],
        [
            [
                s,
                reader.sheet_states[s],
                len(reader.formulas[s]),
                reader.sheet_paths[s],
            ]
            for s in reader.sheets
        ],
    )

    # --------------------------------------------------------
    # Dependency Matrix
    # --------------------------------------------------------

    ws = wb.create_sheet("Dependency_Matrix")

    sheets = reader.sheets

    ws.cell(1, 1, "Source / Target")

    for j, target in enumerate(sheets, start=2):
        cell = ws.cell(1, j, target)
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(
            textRotation=45,
            horizontal="center",
            vertical="bottom",
        )

    for i, source in enumerate(sheets, start=2):
        ws.cell(i, 1, source)

        for j, target in enumerate(sheets, start=2):
            pair = (source, target)

            if pair in analyzer.confirmed:
                value = 1
                color = "C6EFCE"
            elif pair in analyzer.possible:
                value = "?"
                color = "FFEB9C"
            else:
                # No evidence found. This is NOT a proof
                # that no runtime dependency exists.
                value = 0
                color = "FFFFFF"

            cell = ws.cell(i, j, value)

            if color != "FFFFFF":
                cell.fill = PatternFill(
                    "solid", fgColor=color
                )

            cell.alignment = Alignment(
                horizontal="center"
            )

    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 36
    ws.row_dimensions[1].height = 170

    for col in range(2, len(sheets) + 2):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 15

    # --------------------------------------------------------
    # Evidence
    # --------------------------------------------------------

    ws = wb.create_sheet("Evidence")

    write_table(
        ws,
        [
            "Source", "Target", "Category",
            "Location", "Confidence",
            "Expression", "Detail",
        ],
        analyzer.evidence,
    )

    # --------------------------------------------------------
    # Unresolved
    # --------------------------------------------------------

    ws = wb.create_sheet("Unresolved")

    write_table(
        ws,
        [
            "Source", "Location", "Category",
            "Expression", "Reason", "Candidates",
        ],
        analyzer.unresolved,
    )

    # --------------------------------------------------------
    # External
    # --------------------------------------------------------

    ws = wb.create_sheet("External_Links")

    write_table(
        ws,
        [
            "Source", "Location",
            "External Book", "Reference", "Expression",
        ],
        analyzer.external,
    )

    # --------------------------------------------------------
    # VBA
    # --------------------------------------------------------

    ws = wb.create_sheet("VBA_Procedures")

    write_table(
        ws,
        [
            "Module", "Procedure",
            "Referenced Sheets", "Has Navigation",
            "VBA Code",
        ],
        analyzer.vba_rows,
    )

    ws = wb.create_sheet("Macro_Bindings")

    write_table(
        ws,
        [
            "XML Path", "Object Type",
            "Macro", "Detail",
        ],
        analyzer.button_rows,
    )

    # --------------------------------------------------------
    # Defined Names
    # --------------------------------------------------------

    ws = wb.create_sheet("Defined_Names")

    write_table(
        ws,
        ["Name", "Scope", "Expression"],
        [
            [
                item["name"],
                item["scope"],
                item["expression"],
            ]
            for item in reader.defined_names
        ],
    )

    # --------------------------------------------------------
    # Validation
    # --------------------------------------------------------

    ws = wb.create_sheet("Validation")

    validation = [
        [
            "Matrix Semantics",
            "1 = confirmed; ? = possible; 0 = no evidence found",
        ],
        [
            "Important",
            "0 does not prove that no runtime dependency exists",
        ],
        [
            "INDIRECT",
            "Unsupported expressions remain in Unresolved",
        ],
        [
            "VBA",
            "Macro source sheet is not inferred without evidence",
        ],
        [
            "External",
            "External workbook references are excluded from matrix",
        ],
        [
            "Source",
            "Original XLSM is never saved",
        ],
        [
            "Shared Formulas",
            "Shared formula followers may require additional analysis",
        ],
        [
            "Shape Bindings",
            "XML binding coverage varies by Excel object type",
        ],
    ]

    write_table(ws, ["Check", "Result"], validation)

    wb.save(output)

    print(f"Report saved: {output}", flush=True)


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Excel XLSM Dependency Analyzer V2"
    )

    parser.add_argument(
        "input",
        help="Source .xlsm file"
    )

    parser.add_argument(
        "--output",
        default=None,
        help="Output .xlsx file"
    )

    parser.add_argument(
        "--com",
        action="store_true",
        help="Use Excel COM to assist dynamic value resolution"
    )

    args = parser.parse_args()

    source = os.path.abspath(args.input)

    if not os.path.isfile(source):
        print("File not found:", source)
        sys.exit(1)

    if not zipfile.is_zipfile(source):
        print("Invalid XLSM/XLSX file.")
        sys.exit(1)

    output = args.output

    if not output:
        base, _ = os.path.splitext(source)
        output = base + "_dependency_report.xlsx"

    output = os.path.abspath(output)

    if os.path.normcase(source) == os.path.normcase(output):
        raise ValueError("Output cannot overwrite source file")

    print("=" * 65)
    print("XLSM Dependency Analyzer V2")
    print("Source:", source)
    print("Output:", output)
    print("=" * 65)

    reader = XLSMReader(source)

    try:
        print("Loading workbook...", flush=True)

        count = reader.load_cells(
            progress=lambda msg: print(msg, flush=True)
        )

        print("Cells loaded:", count, flush=True)
        print("Sheets:", len(reader.sheets), flush=True)

        analyzer = DependencyAnalyzer(reader)

        if args.com:
            values = load_com_values(source, reader)
            analyzer.resolver.set_com_values(values)

        analyzer.run()

        create_report(reader, analyzer, output)

        print()
        print("=" * 65)
        print("ANALYSIS COMPLETE")
        print("=" * 65)

        for key in (
            "Formula Cells",
            "Confirmed Links",
            "Possible Links",
            "INDIRECT Calls",
            "OFFSET Calls",
            "VBA Modules",
            "VBA Procedures",
            "Macro Bindings",
            "External Records",
            "Evidence Rows",
            "Unresolved Rows",
        ):
            print(f"{key}: {analyzer.stats[key]}")

        print()
        print("Output:", output)

    finally:
        reader.close()


if __name__ == "__main__":
    main()
