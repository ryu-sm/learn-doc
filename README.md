
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel XLSM Direct Sheet Dependency Analyzer

Usage:
    python analyze_xlsm.py input.xlsm

Optional:
    python analyze_xlsm.py input.xlsm -o report.xlsx
    python analyze_xlsm.py input.xlsm --include-self

Dependencies:
    pip install openpyxl oletools

Matrix:
    Row    = Source Sheet
    Column = Target Sheet
    1      = Confirmed direct relationship
    0      = No confirmed relationship

Note:
    0 does NOT prove that no runtime dependency exists.
"""

import argparse
import posixpath
import re
import zipfile

from pathlib import Path
from collections import defaultdict, Counter
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment
)
from openpyxl.utils import get_column_letter

from oletools.olevba import VBA_Parser


# ============================================================
# Configuration
# ============================================================

MAX_EXPRESSION_LENGTH = 32000
MAX_RECURSION = 12

REL_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

RID = "{" + REL_NS + "}id"

CELL_PATTERN = re.compile(
    r"^\$?([A-Z]{1,3})\$?(\d+)$",
    re.I
)

CROSS_CELL_PATTERN = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!]+))!"
    r"(\$?[A-Z]{1,3}\$?\d+)$",
    re.I
)

VBA_SHEET_PATTERN = re.compile(
    r"\b(Sheets|Worksheets)\s*\(\s*"
    r'(?:"((?:[^"]|"")*)"|(\d+))\s*\)',
    re.I
)

VBA_PROC_PATTERN = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*"
    r"(?:Sub|Function)\s+([^\s(]+)",
    re.I
)


# ============================================================
# General Helpers
# ============================================================

def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def normalize_part(base, target):
    target = unquote(target or "")

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(base),
            target
        )
    )


def read_relationships(zf, part):
    folder = posixpath.dirname(part)
    filename = posixpath.basename(part)

    rel_part = posixpath.join(
        folder,
        "_rels",
        filename + ".rels"
    )

    if rel_part not in zf.namelist():
        return {}

    root = ET.fromstring(zf.read(rel_part))
    result = {}

    for element in root:
        if local_name(element.tag) != "Relationship":
            continue

        result[element.get("Id")] = {
            "target": element.get("Target", ""),
            "type": element.get("Type", ""),
            "mode": element.get("TargetMode", "")
        }

    return result


def text_cell(cell, value):
    """
    Force text storage.

    Avoid interpreting formulas as Excel formulas
    in the generated report.
    """
    value = "" if value is None else str(value)

    # Prefix formula-like values with a literal apostrophe
    # so spreadsheet applications do not execute them.
    if value.startswith(("=", "+", "-", "@")):
        value = "'" + value

    cell.value = value[:MAX_EXPRESSION_LENGTH]
    cell.data_type = "s"
    cell.number_format = "@"


def split_top_level(expression, separator=","):
    """
    Split an Excel expression at top-level separators.

    Handles:
      - Nested parentheses
      - Quoted strings
      - Doubled quotes
    """
    result = []
    depth = 0
    quoted = False
    start = 0
    i = 0

    while i < len(expression):
        ch = expression[i]

        if ch == '"':
            if (
                quoted
                and i + 1 < len(expression)
                and expression[i + 1] == '"'
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
                result.append(
                    expression[start:i].strip()
                )
                start = i + 1

        i += 1

    result.append(expression[start:].strip())
    return result


def extract_function_calls(formula, function):
    """
    Extract balanced function calls.

    Returns a list of argument lists.
    """
    pattern = re.compile(
        r"(?<![\w.])"
        + re.escape(function)
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
                        args = split_top_level(
                            formula[start:i], ","
                        )
                        results.append(args)
                        break

            i += 1

    return results


def excel_string_literals(expression):
    """
    Extract actual Excel string literals.

    This avoids treating sheet names appearing
    outside quoted strings as static literals.
    """
    results = []
    i = 0

    while i < len(expression):
        if expression[i] != '"':
            i += 1
            continue

        i += 1
        chars = []

        while i < len(expression):
            if expression[i] == '"':
                if (
                    i + 1 < len(expression)
                    and expression[i + 1] == '"'
                ):
                    chars.append('"')
                    i += 2
                    continue

                i += 1
                break

            chars.append(expression[i])
            i += 1

        results.append("".join(chars))

    return results


def clean_macro_name(value):
    if not value:
        return ""

    value = str(value).strip().strip("'")

    if "!" in value:
        value = value.rsplit("!", 1)[-1]

    if "." in value:
        value = value.rsplit(".", 1)[-1]

    return value.casefold()


def strip_vba_comment(line):
    quoted = False
    output = []
    i = 0

    while i < len(line):
        ch = line[i]

        if ch == '"':
            if (
                quoted
                and i + 1 < len(line)
                and line[i + 1] == '"'
            ):
                output.extend(['"', '"'])
                i += 2
                continue

            quoted = not quoted

        if ch == "'" and not quoted:
            break

        output.append(ch)
        i += 1

    return "".join(output)


def logical_vba_lines(code):
    pending = ""
    start_line = 0

    for line_no, raw in enumerate(
        code.splitlines(), 1
    ):
        line = strip_vba_comment(raw).strip()

        if not line:
            continue

        if not pending:
            start_line = line_no

        if re.search(r"\s+_$", line):
            pending += re.sub(
                r"\s+_$", " ", line
            )
        else:
            yield start_line, pending + line
            pending = ""

    if pending:
        yield start_line, pending


# ============================================================
# Analyzer
# ============================================================

class XLSMAnalyzer:

    def __init__(self, filename, include_self=False):
        self.path = Path(filename)
        self.include_self = include_self

        self.wb = load_workbook(
            self.path,
            data_only=False,
            read_only=True,
            keep_links=True
        )

        self.cached_wb = load_workbook(
            self.path,
            data_only=True,
            read_only=True
        )

        # Include ALL sheets.
        self.sheet_names = list(self.wb.sheetnames)

        self.sheet_lookup = {
            name.casefold(): name
            for name in self.sheet_names
        }

        self.worksheet_names = [
            ws.title for ws in self.wb.worksheets
        ]

        self.sheet_states = {
            ws.title: ws.sheet_state
            for ws in self.wb.worksheets
        }

        self.code_names = {}

        for ws in self.wb.worksheets:
            code_name = ws.sheet_properties.codeName

            if code_name:
                self.code_names[
                    code_name.casefold()
                ] = ws.title

        self.data_links = {
            name: set()
            for name in self.sheet_names
        }

        self.navigation_links = {
            name: set()
            for name in self.sheet_names
        }

        self.evidence = set()
        self.unresolved = set()

        self.stats = Counter()
        self.procedures = defaultdict(list)
        self.procedure_navigation = defaultdict(set)

        self.link_counts = Counter()

    # --------------------------------------------------------
    # Sheet Utilities
    # --------------------------------------------------------

    def normalize_sheet(self, name):
        if not isinstance(name, str):
            return None

        return self.sheet_lookup.get(
            name.casefold()
        )

    def add_link(
        self,
        source,
        target,
        kind,
        relation,
        location,
        expression
    ):
        source = self.normalize_sheet(source)
        target = self.normalize_sheet(target)

        if not source or not target:
            return

        if source == target and not self.include_self:
            self.stats["self_links_excluded"] += 1
            return

        graph = (
            self.navigation_links
            if kind == "Navigation"
            else self.data_links
        )

        graph[source].add(target)

        record = (
            source,
            target,
            kind,
            relation,
            str(location),
            str(expression)
        )

        if record not in self.evidence:
            self.evidence.add(record)
            self.link_counts[
                (source, target, kind)
            ] += 1

    def add_unresolved(
        self,
        source,
        location,
        reason,
        expression
    ):
        self.unresolved.add((
            str(source),
            str(location),
            str(reason),
            str(expression)
        ))

    # --------------------------------------------------------
    # Cell Value Resolution
    # --------------------------------------------------------

    def resolve_cell(
        self,
        sheet,
        address,
        depth=0,
        visited=None
    ):
        if depth > MAX_RECURSION:
            return None

        sheet = self.normalize_sheet(sheet)

        if not sheet:
            return None

        address = address.replace("$", "").upper()

        visited = visited or set()
        key = (sheet, address)

        if key in visited:
            return None

        visited = visited | {key}

        try:
            cell = self.wb[sheet][address]
            value = cell.value

            if cell.data_type != "f":
                return value

            cached = self.cached_wb[
                sheet
            ][address].value

            if cached is not None:
                return cached

            return self.evaluate_expression(
                sheet,
                str(value).lstrip("="),
                depth + 1,
                visited
            )

        except Exception:
            return None

    # --------------------------------------------------------
    # Simple Excel Expression Evaluation
    # --------------------------------------------------------

    def evaluate_expression(
        self,
        sheet,
        expression,
        depth=0,
        visited=None
    ):
        """
        Evaluate only safe, supported expressions.

        Does not execute Excel or VBA.
        """
        if depth > MAX_RECURSION:
            return None

        expression = expression.strip()

        if expression.startswith("="):
            expression = expression[1:].strip()

        # Excel concatenation
        parts = split_top_level(expression, "&")

        if len(parts) > 1:
            values = []

            for part in parts:
                value = self.evaluate_expression(
                    sheet,
                    part,
                    depth + 1,
                    visited
                )

                if value is None:
                    return None

                values.append(str(value))

            return "".join(values)

        # String literal
        if (
            expression.startswith('"')
            and expression.endswith('"')
        ):
            return expression[1:-1].replace(
                '""', '"'
            )

        # Numeric literal
        if re.fullmatch(
            r"-?\d+(?:\.\d+)?",
            expression
        ):
            try:
                return float(expression)
            except ValueError:
                return None

        # Local cell
        if CELL_PATTERN.fullmatch(expression):
            return self.resolve_cell(
                sheet,
                expression,
                depth + 1,
                visited
            )

        # Cross-sheet cell
        match = CROSS_CELL_PATTERN.fullmatch(
            expression
        )

        if match:
            target = (
                match.group(1)
                or match.group(2)
            ).replace("''", "'")

            return self.resolve_cell(
                target,
                match.group(3),
                depth + 1,
                visited
            )

        # Simple string functions
        match = re.fullmatch(
            r"(CONCATENATE|CONCAT|UPPER|LOWER|"
            r"TRIM|IFERROR)\s*\((.*)\)",
            expression,
            re.I | re.S
        )

        if match:
            function = match.group(1).upper()

            args = split_top_level(
                match.group(2), ","
            )

            if function == "IFERROR":
                if not args:
                    return None

                value = self.evaluate_expression(
                    sheet,
                    args[0],
                    depth + 1,
                    visited
                )

                if value is None and len(args) > 1:
                    return self.evaluate_expression(
                        sheet,
                        args[1],
                        depth + 1,
                        visited
                    )

                return value

            values = []

            for arg in args:
                value = self.evaluate_expression(
                    sheet,
                    arg,
                    depth + 1,
                    visited
                )

                if value is None:
                    return None

                values.append(str(value))

            if function in ("CONCATENATE", "CONCAT"):
                return "".join(values)

            if len(values) != 1:
                return None

            if function == "UPPER":
                return values[0].upper()

            if function == "LOWER":
                return values[0].lower()

            if function == "TRIM":
                return " ".join(values[0].split())

        return None

    # --------------------------------------------------------
    # Sheet Name From Reference
    # --------------------------------------------------------

    def sheet_from_reference(self, value):
        if not isinstance(value, str):
            return None

        value = value.strip()

        if value.startswith("#"):
            value = value[1:]

        if "!" not in value:
            return None

        name = value.split("!", 1)[0].strip()

        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")

        if "[" in name or "]" in name:
            return None

        return self.normalize_sheet(name)

    # --------------------------------------------------------
    # Explicit Formula References
    # --------------------------------------------------------

    def explicit_formula_targets(self, formula):
        targets = set()

        try:
            tokens = Tokenizer(formula).items
        except Exception:
            self.stats["tokenizer_errors"] += 1
            return targets

        for token in tokens:
            if not (
                token.type == "OPERAND"
                and token.subtype == "RANGE"
            ):
                continue

            value = token.value

            if "!" not in value:
                continue

            sheet_part = value.rsplit(
                "!", 1
            )[0].strip()

            if (
                sheet_part.startswith("'")
                and sheet_part.endswith("'")
            ):
                sheet_part = sheet_part[1:-1].replace(
                    "''", "'"
                )

            if "[" in sheet_part or "]" in sheet_part:
                continue

            # 3D references
            if ":" in sheet_part:
                first, last = sheet_part.split(
                    ":", 1
                )

                first = self.normalize_sheet(first)
                last = self.normalize_sheet(last)

                if first and last:
                    a = self.sheet_names.index(first)
                    b = self.sheet_names.index(last)

                    targets.update(
                        self.sheet_names[
                            min(a, b):max(a, b) + 1
                        ]
                    )

                continue

            target = self.normalize_sheet(sheet_part)

            if target:
                targets.add(target)

        return targets

    # --------------------------------------------------------
    # Static INDIRECT Sheet Detection
    # --------------------------------------------------------

    def static_indirect_targets(self, expression):
        """
        Detect literal sheet names inside INDIRECT.

        Examples:
            INDIRECT("サマリ!A1")
            INDIRECT("'サマリ'!" & ADDRESS(...))
            INDIRECT("サマリ!" & something)

        This is a heuristic for embedded static names.
        """
        targets = set()

        literals = excel_string_literals(expression)

        for literal in literals:
            for sheet in self.sheet_names:
                candidates = [
                    sheet + "!",
                    "'" + sheet.replace("'", "''") + "'!"
                ]

                if any(
                    candidate.casefold()
                    in literal.casefold()
                    for candidate in candidates
                ):
                    targets.add(sheet)

        return targets

    # --------------------------------------------------------
    # INDIRECT Analysis
    # --------------------------------------------------------

    def analyze_indirect(
        self,
        sheet,
        address,
        formula
    ):
        calls = extract_function_calls(
            formula, "INDIRECT"
        )

        for args in calls:
            if not args:
                continue

            self.stats["indirect_calls"] += 1

            expression = args[0]

            # Method 1:
            # Fully resolve string expression
            resolved = self.evaluate_expression(
                sheet,
                expression
            )

            target = self.sheet_from_reference(
                resolved
            )

            if target:
                self.add_link(
                    sheet,
                    target,
                    "Data",
                    "INDIRECT Resolved",
                    f"{sheet}!{address}",
                    expression
                )

                self.stats[
                    "indirect_fully_resolved"
                ] += 1

                continue

            # Method 2:
            # Find explicit sheet name in string literals
            static_targets = self.static_indirect_targets(
                expression
            )

            if static_targets:
                for target in static_targets:
                    self.add_link(
                        sheet,
                        target,
                        "Data",
                        "INDIRECT Static Sheet",
                        f"{sheet}!{address}",
                        expression
                    )

                self.stats[
                    "indirect_static_resolved"
                ] += 1

                continue

            # Method 3:
            # No sheet specified in ADDRESS:
            # normally a current-sheet reference.
            #
            # Do not assume this for every expression.
            address_calls = extract_function_calls(
                expression, "ADDRESS"
            )

            if address_calls:
                all_local = all(
                    len(call) < 5
                    or not call[4].strip()
                    for call in address_calls
                )

                if all_local and not (
                    "!" in expression
                    or re.search(
                        r"\bINDIRECT\s*\(",
                        expression,
                        re.I
                    )
                ):
                    self.stats[
                        "indirect_local_candidates"
                    ] += 1

                    self.add_unresolved(
                        sheet,
                        address,
                        "INDIRECT ADDRESS local candidate",
                        expression
                    )

                    continue

            self.stats[
                "indirect_unresolved"
            ] += 1

            self.add_unresolved(
                sheet,
                address,
                "INDIRECT dynamic target unresolved",
                expression
            )

    # --------------------------------------------------------
    # Formula Analysis
    # --------------------------------------------------------

    def analyze_formulas(self):
        print("[1/5] Formula analysis...", flush=True)

        for sheet in self.worksheet_names:
            ws = self.wb[sheet]

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["formula_cells"] += 1

                    formula = str(cell.value)

                    # Explicit sheet references
                    targets = self.explicit_formula_targets(
                        formula
                    )

                    for target in targets:
                        self.add_link(
                            sheet,
                            target,
                            "Data",
                            "Explicit Formula",
                            f"{sheet}!{cell.coordinate}",
                            formula
                        )

                    # INDIRECT
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

                    # OFFSET
                    self.stats["offset_calls"] += len(
                        extract_function_calls(
                            formula, "OFFSET"
                        )
                    )

                    # HYPERLINK formula
                    for args in extract_function_calls(
                        formula, "HYPERLINK"
                    ):
                        if not args:
                            continue

                        value = self.evaluate_expression(
                            sheet, args[0]
                        )

                        target = self.sheet_from_reference(
                            value
                        )

                        if target:
                            self.add_link(
                                sheet,
                                target,
                                "Navigation",
                                "HYPERLINK Formula",
                                f"{sheet}!{cell.coordinate}",
                                formula
                            )

                    if (
                        self.stats["formula_cells"]
                        % 10000 == 0
                    ):
                        print(
                            "  Processed:",
                            self.stats["formula_cells"],
                            flush=True
                        )

    # --------------------------------------------------------
    # VBA Sheet References
    # --------------------------------------------------------

    def vba_sheet_targets(
        self,
        expression,
        aliases=None
    ):
        aliases = aliases or {}
        targets = set()

        for match in VBA_SHEET_PATTERN.finditer(
            expression
        ):
            kind, name, number = match.groups()

            if name is not None:
                target = self.normalize_sheet(
                    name.replace('""', '"')
                )
            else:
                index = int(number)

                sequence = (
                    self.worksheet_names
                    if kind.casefold() == "worksheets"
                    else self.sheet_names
                )

                target = (
                    sequence[index - 1]
                    if 1 <= index <= len(sequence)
                    else None
                )

            if target:
                targets.add(target)

        for code, sheet in self.code_names.items():
            if re.search(
                r"(?<![\w.])"
                + re.escape(code)
                + r"\s*[.!]",
                expression,
                re.I
            ):
                targets.add(sheet)

        for alias, sheet in aliases.items():
            if re.search(
                r"(?<![\w.])"
                + re.escape(alias)
                + r"\s*[.!]",
                expression,
                re.I
            ):
                targets.add(sheet)

        return targets

    # --------------------------------------------------------
    # Extract VBA
    # --------------------------------------------------------

    def extract_vba(self):
        print("[2/5] VBA extraction...", flush=True)

        parser = VBA_Parser(str(self.path))

        try:
            if not parser.detect_vba_macros():
                return

            for _, _, module, code in parser.extract_macros():
                self.stats["vba_modules"] += 1

                current_proc = None

                for line_no, line in logical_vba_lines(
                    code
                ):
                    match = VBA_PROC_PATTERN.match(line)

                    if match:
                        current_proc = clean_macro_name(
                            match.group(1)
                        )
                        continue

                    if re.match(
                        r"^\s*End\s+(Sub|Function)\b",
                        line,
                        re.I
                    ):
                        current_proc = None
                        continue

                    if current_proc:
                        self.procedures[
                            current_proc
                        ].append((
                            module,
                            line_no,
                            line
                        ))

        finally:
            parser.close()

    # --------------------------------------------------------
    # Analyze VBA
    # --------------------------------------------------------

    def analyze_vba(self):
        print("[3/5] VBA analysis...", flush=True)

        calls = defaultdict(set)
        direct_targets = defaultdict(set)

        for proc, lines in self.procedures.items():
            aliases = {}

            for module, line_no, line in lines:

                # Set ws = Sheets("xxx")
                match = re.match(
                    r"^Set\s+(\w+)\s*=\s*(.+)$",
                    line,
                    re.I
                )

                if match:
                    variable = match.group(1).casefold()
                    expression = match.group(2)

                    refs = self.vba_sheet_targets(
                        expression, aliases
                    )

                    if len(refs) == 1:
                        aliases[variable] = next(
                            iter(refs)
                        )

                    continue

                # Activate / Select
                nav = re.search(
                    r"\.\s*(Activate|Select)\b",
                    line,
                    re.I
                )

                if nav:
                    refs = self.vba_sheet_targets(
                        line[:nav.start()],
                        aliases
                    )

                    direct_targets[proc].update(refs)

                # Application.Goto
                if re.search(
                    r"\b(?:Application\s*\.\s*)?"
                    r"Goto\b",
                    line,
                    re.I
                ):
                    direct_targets[proc].update(
                        self.vba_sheet_targets(
                            line, aliases
                        )
                    )

                # Explicit VBA assignments
                if "=" in line and not re.match(
                    r"^(If|ElseIf|For|While|Do|"
                    r"Select|Case|Const|Dim)\b",
                    line,
                    re.I
                ):
                    left, right = line.split("=", 1)

                    left_refs = self.vba_sheet_targets(
                        left, aliases
                    )

                    right_refs = self.vba_sheet_targets(
                        right, aliases
                    )

                    for target in left_refs:
                        for source in right_refs:
                            self.add_link(
                                target,
                                source,
                                "Data",
                                "VBA Assignment",
                                f"{module}:{line_no}",
                                line
                            )

                # Procedure calls
                call = re.match(
                    r"^(?:Call\s+)?([^\s(]+)",
                    line,
                    re.I
                )

                if call:
                    called = clean_macro_name(
                        call.group(1)
                    )

                    if (
                        called in self.procedures
                        and called != proc
                    ):
                        calls[proc].add(called)

        # Resolve procedure navigation targets
        resolved = {
            proc: set(direct_targets[proc])
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

        self.procedure_navigation = resolved

        self.stats["vba_procedures"] = len(
            self.procedures
        )

    # --------------------------------------------------------
    # Shape Macro Binding
    # --------------------------------------------------------

    def add_shape_macro(
        self,
        source,
        macro,
        location
    ):
        self.stats["shape_macro_bindings"] += 1

        proc = clean_macro_name(macro)

        targets = self.procedure_navigation.get(
            proc, set()
        )

        if not targets:
            self.add_unresolved(
                source,
                location,
                "Shape macro target unresolved",
                macro
            )
            return

        for target in targets:
            self.add_link(
                source,
                target,
                "Navigation",
                "Shape VBA Macro",
                location,
                macro
            )

    # --------------------------------------------------------
    # Sheet XML Mapping
    # --------------------------------------------------------

    def sheet_xml_parts(self, zf):
        workbook_part = "xl/workbook.xml"

        root = ET.fromstring(
            zf.read(workbook_part)
        )

        rels = read_relationships(
            zf, workbook_part
        )

        mapping = {}

        for element in root.iter():
            if local_name(element.tag) != "sheet":
                continue

            name = element.get("name")
            rid = element.get(RID)

            if rid not in rels:
                continue

            mapping[name] = normalize_part(
                workbook_part,
                rels[rid]["target"]
            )

        return mapping

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

        rels = read_relationships(
            zf, part
        )

        for element in root.iter():
            tag = local_name(element.tag)

            if tag in (
                "sp",
                "pic",
                "graphicFrame"
            ):
                macro = element.get("macro")

                if macro:
                    self.add_shape_macro(
                        source,
                        macro,
                        part
                    )

            if tag == "hlinkClick":
                rid = element.get(RID)

                relation = rels.get(
                    rid, {}
                )

                location = relation.get(
                    "target", ""
                )

                target = self.sheet_from_reference(
                    location
                )

                if target:
                    self.add_link(
                        source,
                        target,
                        "Navigation",
                        "Shape Hyperlink",
                        part,
                        location
                    )

    # --------------------------------------------------------
    # VML
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

        for element in root.iter():
            if local_name(element.tag) != "ClientData":
                continue

            for child in element:
                if local_name(child.tag) != "FmlaMacro":
                    continue

                macro = (
                    child.text or ""
                ).strip()

                if macro:
                    self.add_shape_macro(
                        source,
                        macro,
                        part
                    )

    # --------------------------------------------------------
    # Cell Hyperlinks
    # --------------------------------------------------------

    def analyze_hyperlinks(
        self,
        zf,
        source,
        part,
        rels
    ):
        root = ET.fromstring(
            zf.read(part)
        )

        for element in root.iter():
            if local_name(element.tag) != "hyperlink":
                continue

            location = element.get(
                "location", ""
            )

            rid = element.get(RID)

            if not location and rid:
                location = rels.get(
                    rid, {}
                ).get("target", "")

            target = self.sheet_from_reference(
                location
            )

            if target:
                self.add_link(
                    source,
                    target,
                    "Navigation",
                    "Cell Hyperlink",
                    f"{part}:{element.get('ref', '')}",
                    location
                )

    # --------------------------------------------------------
    # Analyze All Shapes
    # --------------------------------------------------------

    def analyze_shapes(self):
        print("[4/5] Shape analysis...", flush=True)

        with zipfile.ZipFile(
            self.path
        ) as zf:
            mapping = self.sheet_xml_parts(zf)

            for source, sheet_part in mapping.items():
                if sheet_part not in zf.namelist():
                    continue

                rels = read_relationships(
                    zf,
                    sheet_part
                )

                self.analyze_hyperlinks(
                    zf,
                    source,
                    sheet_part,
                    rels
                )

                for relation in rels.values():
                    if relation["mode"] == "External":
                        continue

                    target_part = normalize_part(
                        sheet_part,
                        relation["target"]
                    )

                    relation_type = relation["type"]

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
    # Matrix Output
    # --------------------------------------------------------

    def create_matrix(
        self,
        wb,
        title,
        graph
    ):
        ws = wb.create_sheet(title)

        text_cell(
            ws.cell(1, 1),
            "Source / Target"
        )

        for index, name in enumerate(
            self.sheet_names, 2
        ):
            text_cell(
                ws.cell(1, index),
                name
            )

            text_cell(
                ws.cell(index, 1),
                name
            )

        for i, source in enumerate(
            self.sheet_names, 2
        ):
            for j, target in enumerate(
                self.sheet_names, 2
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

        self.style_header(ws)

        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 35

        for col in range(
            2, len(self.sheet_names) + 2
        ):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 24

    # --------------------------------------------------------
    # Count Matrix
    # --------------------------------------------------------

    def create_count_matrix(self, wb):
        ws = wb.create_sheet("Evidence_Count")

        text_cell(
            ws.cell(1, 1),
            "Source / Target"
        )

        for index, name in enumerate(
            self.sheet_names, 2
        ):
            text_cell(ws.cell(1, index), name)
            text_cell(ws.cell(index, 1), name)

        for i, source in enumerate(
            self.sheet_names, 2
        ):
            for j, target in enumerate(
                self.sheet_names, 2
            ):
                count = (
                    self.link_counts[
                        (source, target, "Data")
                    ]
                    + self.link_counts[
                        (source, target, "Navigation")
                    ]
                )

                ws.cell(i, j, count)

        self.style_header(ws)
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 35

        for col in range(
            2, len(self.sheet_names) + 2
        ):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 24

    # --------------------------------------------------------
    # Report Helpers
    # --------------------------------------------------------

    @staticmethod
    def style_header(ws):
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

    def create_detail_sheet(
        self,
        wb,
        title,
        headers,
        records,
        widths
    ):
        ws = wb.create_sheet(title)

        for col, value in enumerate(headers, 1):
            text_cell(
                ws.cell(1, col),
                value
            )

        for row_no, record in enumerate(
            sorted(records), 2
        ):
            for col, value in enumerate(
                record, 1
            ):
                text_cell(
                    ws.cell(row_no, col),
                    value
                )

        self.style_header(ws)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col, width in enumerate(widths, 1):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = width

    # --------------------------------------------------------
    # Generate Report
    # --------------------------------------------------------

    def generate_report(self, output):
        print("[5/5] Generating report...", flush=True)

        wb = Workbook()
        wb.remove(wb.active)

        combined = {
            sheet: (
                self.data_links[sheet]
                | self.navigation_links[sheet]
            )
            for sheet in self.sheet_names
        }

        self.create_matrix(
            wb,
            "Direct_Matrix",
            combined
        )

        self.create_matrix(
            wb,
            "Data_Matrix",
            self.data_links
        )

        self.create_matrix(
            wb,
            "Navigation_Matrix",
            self.navigation_links
        )

        self.create_count_matrix(wb)

        # Sheet Information
        self.create_detail_sheet(
            wb,
            "Sheet_Information",
            ["Sheet Name", "Visibility"],
            [
                (
                    name,
                    self.sheet_states.get(
                        name, "unknown"
                    )
                )
                for name in self.sheet_names
            ],
            [45, 20]
        )

        # Evidence
        self.create_detail_sheet(
            wb,
            "Dependency_Details",
            [
                "Source Sheet",
                "Target Sheet",
                "Kind",
                "Relation",
                "Location",
                "Expression"
            ],
            self.evidence,
            [30, 30, 20, 30, 40, 110]
        )

        # Unresolved
        self.create_detail_sheet(
            wb,
            "Unresolved_References",
            [
                "Source",
                "Location",
                "Reason",
                "Expression"
            ],
            self.unresolved,
            [30, 35, 45, 110]
        )

        # Statistics
        stats = [
            ("Total Sheets", len(self.sheet_names)),
            (
                "Hidden Sheets",
                sum(
                    state != "visible"
                    for state in self.sheet_states.values()
                )
            ),
            (
                "Formula Cells",
                self.stats["formula_cells"]
            ),
            (
                "INDIRECT Calls",
                self.stats["indirect_calls"]
            ),
            (
                "INDIRECT Fully Resolved",
                self.stats["indirect_fully_resolved"]
            ),
            (
                "INDIRECT Static Resolved",
                self.stats["indirect_static_resolved"]
            ),
            (
                "INDIRECT Local Candidates",
                self.stats["indirect_local_candidates"]
            ),
            (
                "INDIRECT Unresolved",
                self.stats["indirect_unresolved"]
            ),
            (
                "OFFSET Calls",
                self.stats["offset_calls"]
            ),
            (
                "VBA Modules",
                self.stats["vba_modules"]
            ),
            (
                "VBA Procedures",
                self.stats["vba_procedures"]
            ),
            (
                "Shape Macro Bindings",
                self.stats["shape_macro_bindings"]
            ),
            (
                "Self Links Excluded",
                self.stats["self_links_excluded"]
            ),
            (
                "Data Link Pairs",
                sum(
                    len(v)
                    for v in self.data_links.values()
                )
            ),
            (
                "Navigation Link Pairs",
                sum(
                    len(v)
                    for v in self.navigation_links.values()
                )
            ),
            (
                "Combined Link Pairs",
                sum(
                    len(v)
                    for v in combined.values()
                )
            ),
            (
                "Evidence Rows",
                len(self.evidence)
            ),
            (
                "Unresolved Rows",
                len(self.unresolved)
            )
        ]

        self.create_detail_sheet(
            wb,
            "Statistics",
            ["Metric", "Value"],
            stats,
            [40, 25]
        )

        wb.save(output)

        return stats

    # --------------------------------------------------------
    # Run
    # --------------------------------------------------------

    def run(self, output):
        try:
            self.analyze_formulas()
            self.extract_vba()
            self.analyze_vba()
            self.analyze_shapes()

            stats = self.generate_report(output)

            print()
            print("========== ANALYSIS RESULT ==========")

            for name, value in stats:
                print(f"{name}: {value}")

            print("Report:", output)
            print("=====================================")

        finally:
            self.wb.close()
            self.cached_wb.close()


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Analyze XLSM direct sheet dependencies"
        )
    )

    parser.add_argument(
        "input",
        help="Input XLSM/XLSX file"
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output report XLSX"
    )

    parser.add_argument(
        "--include-self",
        action="store_true",
        help="Include same-sheet references"
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

    analyzer = XLSMAnalyzer(
        source,
        include_self=args.include_self
    )

    analyzer.run(output)


if __name__ == "__main__":
    main()
