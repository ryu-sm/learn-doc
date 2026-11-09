
import argparse
import re
import sys
import zipfile
import posixpath
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.formula.tokenizer import Tokenizer


# ============================================================
# Configuration
# ============================================================

MAX_EVIDENCE = 500000
MAX_UNRESOLVED = 100000
PROGRESS_INTERVAL = 10000

SHEET_REF = re.compile(
    r"(?:'((?:[^']|'')+)'|([A-Za-z_][\w.\u0080-\uffff ]*))!"
    r"(?=\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}:|\$?\d+:)",
    re.I
)

CELL_REF = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$", re.I)

RANGE_REF = re.compile(
    r"^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$",
    re.I
)

FUNCTION_NAMES = (
    "INDIRECT", "OFFSET", "INDEX", "ADDRESS",
    "SUBSTITUTE", "HLOOKUP", "VLOOKUP",
    "MATCH", "IF", "ROW", "COLUMN",
    "CONCATENATE", "CHOOSE"
)


def norm(value):
    return str(value or "").strip().casefold()


def clean_sheet_name(value):
    value = str(value or "").strip()
    if value.startswith("'") and value.endswith("'"):
        value = value[1:-1].replace("''", "'")
    return value


def split_top(text, separator=","):
    """Split an Excel expression without splitting nested functions."""
    result = []
    start = 0
    depth = 0
    quoted = False
    i = 0

    while i < len(text):
        c = text[i]

        if c == '"':
            if quoted and i + 1 < len(text) and text[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted

        elif not quoted:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            elif c == separator and depth == 0:
                result.append(text[start:i].strip())
                start = i + 1

        i += 1

    result.append(text[start:].strip())
    return result


def strip_outer(expr):
    expr = expr.strip()

    while expr.startswith("(") and expr.endswith(")"):
        depth = 0
        valid = True
        quoted = False

        for i, c in enumerate(expr):
            if c == '"':
                quoted = not quoted
            elif not quoted:
                if c == "(":
                    depth += 1
                elif c == ")":
                    depth -= 1
                    if depth == 0 and i != len(expr) - 1:
                        valid = False
                        break

        if not valid or depth != 0:
            break

        expr = expr[1:-1].strip()

    return expr


def function_parts(expr):
    expr = strip_outer(expr)
    match = re.match(r"^([A-Z_][A-Z0-9_.]*)\(", expr, re.I)

    if not match or not expr.endswith(")"):
        return None, None

    name = match.group(1).upper().replace("_XLFN.", "")
    return name, split_top(expr[match.end():-1])


def find_calls(formula, function_name):
    """Extract function calls using balanced parentheses."""
    pattern = re.compile(r"\b" + re.escape(function_name) + r"\s*\(", re.I)
    pos = 0

    while True:
        match = pattern.search(formula, pos)
        if not match:
            return

        start = match.end()
        depth = 1
        quoted = False
        i = start

        while i < len(formula):
            c = formula[i]

            if c == '"':
                if quoted and i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 2
                    continue
                quoted = not quoted

            elif not quoted:
                if c == "(":
                    depth += 1
                elif c == ")":
                    depth -= 1
                    if depth == 0:
                        yield formula[start:i]
                        break

            i += 1

        pos = max(i + 1, match.end())


class Unknown(Exception):
    pass


class Analyzer:
    def __init__(self, filename):
        self.path = Path(filename).resolve()

        print("Loading workbook...", flush=True)

        self.wb = load_workbook(
            self.path,
            read_only=True,
            data_only=False,
            keep_links=False
        )

        self.values = load_workbook(
            self.path,
            read_only=True,
            data_only=True,
            keep_links=False
        )

        self.names = list(self.wb.sheetnames)
        self.name_map = {norm(n): n for n in self.names}

        self.direct = defaultdict(set)
        self.navigation = defaultdict(set)
        self.possible = defaultdict(set)

        self.evidence = []
        self.unresolved = []
        self.counts = Counter()
        self.unknown_sources = set()

        self.vba_code = []
        self.procedures = {}
        self.macro_bindings = []

    def canonical(self, name):
        return self.name_map.get(norm(clean_sheet_name(name)))

    def value(self, sheet, address):
        ws = self.values[sheet]
        return ws[address.replace("$", "")].value

    def add_link(self, source, target, kind, location, detail):
        target = self.canonical(target)

        if not target:
            self.add_unresolved(
                source, location, kind, detail,
                "Target sheet not found"
            )
            return

        self.direct[source].add(target)

        if len(self.evidence) < MAX_EVIDENCE:
            self.evidence.append(
                (source, target, kind, location, str(detail)[:1500])
            )
        else:
            self.counts["Truncated Evidence"] += 1

        self.counts[kind] += 1

    def add_navigation(self, source, target, kind, location, detail):
        target = self.canonical(target)

        if target:
            self.navigation[source].add(target)

            if len(self.evidence) < MAX_EVIDENCE:
                self.evidence.append(
                    (source, target, kind, location, str(detail)[:1500])
                )

            self.counts["Navigation Links"] += 1
        else:
            self.add_unresolved(
                source, location, kind, detail,
                "Navigation target not found"
            )

    def add_unresolved(self, source, location, kind, expression, reason):
        self.unknown_sources.add(source)
        self.counts["Unresolved Rows"] += 1

        if len(self.unresolved) < MAX_UNRESOLVED:
            self.unresolved.append(
                (source, location, kind, str(expression)[:3000], str(reason))
            )
        else:
            self.counts["Truncated Unresolved"] += 1

    # ========================================================
    # Expression evaluation
    # ========================================================

    def evaluate(self, expr, sheet, cell, depth=0):
        """
        Limited Excel expression evaluator.

        It does not execute Excel formulas.
        It resolves only expressions needed to identify references.
        """
        if depth > 16:
            raise Unknown("Maximum expression depth")

        expr = strip_outer(str(expr).strip())

        if expr.startswith("="):
            expr = expr[1:].strip()

        if not expr:
            return ""

        parts = split_top(expr, "&")

        if len(parts) > 1:
            return "".join(
                self.to_text(self.evaluate(p, sheet, cell, depth + 1))
                for p in parts
            )

        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1].replace('""', '"')

        if re.fullmatch(r"-?\d+(?:\.\d+)?", expr):
            return float(expr) if "." in expr else int(expr)

        if expr.upper() == "TRUE":
            return True

        if expr.upper() == "FALSE":
            return False

        if expr.upper() in ("ROW()", "ROW( )"):
            return cell.row

        if expr.upper() in ("COLUMN()", "COLUMN( )"):
            return cell.column

        match = CELL_REF.fullmatch(expr)

        if match:
            value = self.value(sheet, expr)
            if value is None:
                return ""
            return value

        # Cross-sheet cell reference
        match = re.fullmatch(
            r"(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]{1,3}\$?\d+)",
            expr,
            re.I
        )

        if match:
            target = self.canonical(match.group(1) or match.group(2))

            if not target:
                raise Unknown("Referenced sheet does not exist")

            return self.value(target, match.group(3))

        name, args = function_parts(expr)

        if name == "ROW":
            if not args or not args[0]:
                return cell.row

            arg = args[0].replace("$", "")
            m = re.match(r"[A-Z]+(\d+)", arg, re.I)

            if m:
                return int(m.group(1))

            raise Unknown("Unsupported ROW argument")

        if name == "COLUMN":
            if not args or not args[0]:
                return cell.column

            arg = args[0].replace("$", "")
            m = re.match(r"([A-Z]+)", arg, re.I)

            if m:
                from openpyxl.utils import column_index_from_string
                return column_index_from_string(m.group(1))

            raise Unknown("Unsupported COLUMN argument")

        if name == "SUBSTITUTE" and len(args) >= 3:
            text = self.to_text(self.evaluate(args[0], sheet, cell, depth + 1))
            old = self.to_text(self.evaluate(args[1], sheet, cell, depth + 1))
            new = self.to_text(self.evaluate(args[2], sheet, cell, depth + 1))

            if len(args) >= 4:
                occurrence = int(self.evaluate(args[3], sheet, cell, depth + 1))
                if occurrence < 1 or not old:
                    raise Unknown("Invalid SUBSTITUTE occurrence")

                segments = text.split(old)
                if occurrence >= len(segments):
                    return text

                return old.join(segments[:occurrence]) + new + old.join(
                    segments[occurrence:]
                )

            return text.replace(old, new)

        if name == "ADDRESS" and len(args) >= 2:
            row = int(self.evaluate(args[0], sheet, cell, depth + 1))
            col = int(self.evaluate(args[1], sheet, cell, depth + 1))

            mode = 1
            if len(args) >= 3:
                mode = int(self.evaluate(args[2], sheet, cell, depth + 1))

            if mode not in (1, 2, 3, 4):
                raise Unknown("Invalid ADDRESS mode")

            letter = get_column_letter(col)

            col_text = "$" + letter if mode in (1, 3) else letter
            row_text = "$" + str(row) if mode in (1, 2) else str(row)

            result = col_text + row_text

            if len(args) >= 5:
                target = self.to_text(
                    self.evaluate(args[4], sheet, cell, depth + 1)
                )
                result = "'" + target.replace("'", "''") + "'!" + result

            return result

        if name == "HLOOKUP" and len(args) >= 3:
            key = self.evaluate(args[0], sheet, cell, depth + 1)
            lookup_sheet, bounds = self.parse_range(args[1], sheet)

            row_index = int(
                self.evaluate(args[2], sheet, cell, depth + 1)
            )

            exact = True

            if len(args) >= 4:
                exact = not bool(
                    self.evaluate(args[3], sheet, cell, depth + 1)
                )

            if not exact:
                raise Unknown("Approximate HLOOKUP not supported")

            min_col, min_row, max_col, max_row = bounds
            target_row = min_row + row_index - 1

            if target_row > max_row or row_index < 1:
                raise Unknown("HLOOKUP row outside lookup range")

            ws = self.values[lookup_sheet]

            for col in range(min_col, max_col + 1):
                candidate = ws.cell(min_row, col).value

                if self.same_value(candidate, key):
                    result = ws.cell(target_row, col).value
                    if result is None:
                        raise Unknown("HLOOKUP result is empty")
                    return result

            raise Unknown("HLOOKUP lookup value not found")

        if name == "IF" and len(args) >= 3:
            condition = self.evaluate_condition(
                args[0], sheet, cell, depth + 1
            )
            branch = args[1] if condition else args[2]
            return self.evaluate(branch, sheet, cell, depth + 1)

        if name == "CONCATENATE":
            return "".join(
                self.to_text(self.evaluate(a, sheet, cell, depth + 1))
                for a in args
            )

        if name == "CHOOSE" and len(args) >= 2:
            index = int(self.evaluate(args[0], sheet, cell, depth + 1))

            if 1 <= index < len(args):
                return self.evaluate(args[index], sheet, cell, depth + 1)

            raise Unknown("CHOOSE index out of range")

        # Basic arithmetic
        for operator in ("+", "-", "*", "/"):
            parts = split_top(expr, operator)

            if len(parts) == 2 and parts[0] and parts[1]:
                a = self.evaluate(parts[0], sheet, cell, depth + 1)
                b = self.evaluate(parts[1], sheet, cell, depth + 1)

                try:
                    a, b = float(a), float(b)

                    if operator == "+":
                        return a + b
                    if operator == "-":
                        return a - b
                    if operator == "*":
                        return a * b
                    if operator == "/":
                        return a / b
                except (TypeError, ValueError, ZeroDivisionError):
                    raise Unknown("Arithmetic could not be evaluated")

        raise Unknown("Unsupported expression: " + expr[:180])

    @staticmethod
    def to_text(value):
        if value is None:
            return ""
        if isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    @staticmethod
    def same_value(a, b):
        if a == b:
            return True

        if a is None or b is None:
            return False

        return str(a).strip().casefold() == str(b).strip().casefold()

    def evaluate_condition(self, expr, sheet, cell, depth):
        for operator in ("<>", ">=", "<=", "=", ">", "<"):
            parts = split_top(expr, operator)

            if len(parts) == 2:
                a = self.evaluate(parts[0], sheet, cell, depth + 1)
                b = self.evaluate(parts[1], sheet, cell, depth + 1)

                if operator == "=":
                    return self.same_value(a, b)
                if operator == "<>":
                    return not self.same_value(a, b)

                try:
                    a, b = float(a), float(b)
                except (TypeError, ValueError):
                    a, b = str(a), str(b)

                if operator == ">":
                    return a > b
                if operator == "<":
                    return a < b
                if operator == ">=":
                    return a >= b
                if operator == "<=":
                    return a <= b

        return bool(self.evaluate(expr, sheet, cell, depth + 1))

    def parse_range(self, expression, default_sheet):
        expression = expression.strip()
        target = default_sheet

        if "!" in expression:
            sheet_part, address = expression.rsplit("!", 1)
            target = self.canonical(sheet_part)
        else:
            address = expression

        if not target:
            raise Unknown("Range target sheet not found")

        try:
            bounds = range_boundaries(address.replace("$", ""))
        except (ValueError, TypeError):
            raise Unknown("Invalid range: " + expression)

        if any(v is None for v in bounds):
            raise Unknown("Unbounded range")

        return target, bounds

    # ========================================================
    # Formula analysis
    # ========================================================

    def direct_references(self, formula):
        """Find explicit sheet references, ignoring string literals."""
        try:
            tokens = Tokenizer(formula).items
            pieces = []

            for token in tokens:
                if token.type == "OPERAND" and token.subtype == "RANGE":
                    pieces.append(token.value)

            text = " ".join(pieces)

        except Exception:
            # Do not scan literal strings as confirmed references.
            text = re.sub(r'"(?:[^"]|"")*"', '""', formula)

        found = set()

        for match in SHEET_REF.finditer(text):
            target = self.canonical(match.group(1) or match.group(2))
            if target:
                found.add(target)

        return found

    def resolve_indirect(self, source, cell, expression):
        args = split_top(expression)

        if not args:
            raise Unknown("Empty INDIRECT")

        if len(args) >= 2:
            mode = self.evaluate(args[1], source, cell)
            if not bool(mode):
                raise Unknown("INDIRECT R1C1 mode not supported")

        text = self.to_text(self.evaluate(args[0], source, cell))
        text = text.strip()

        if "!" not in text:
            raise Unknown("INDIRECT result has no sheet qualifier: " + text)

        target_name = text.rsplit("!", 1)[0]
        target = self.canonical(target_name)

        if not target:
            raise Unknown("INDIRECT target sheet not found: " + target_name)

        return target, text

    def resolve_offset(self, source, cell, expression):
        args = split_top(expression)

        if not args:
            raise Unknown("Empty OFFSET")

        first = args[0].strip()

        if "!" in first:
            target_name = first.rsplit("!", 1)[0]
            target = self.canonical(target_name)

            if target:
                return target

            raise Unknown("OFFSET target sheet not found")

        # Local OFFSET remains within the source worksheet.
        if CELL_REF.fullmatch(first):
            return source

        if RANGE_REF.fullmatch(first):
            return source

        raise Unknown("OFFSET base reference could not be resolved")

    def candidate_sheets(self, source, cell, expression):
        """Identify only specific possible targets, never all sheets."""
        candidates = set()

        # Sheet names appearing as literal text
        for name in self.names:
            quoted = "'" + name.replace("'", "''") + "'!"

            if quoted in expression or name + "!" in expression:
                candidates.add(name)

        # Common dynamic reference: "'" & $E36 & "'!..."
        for match in re.finditer(
            r"(?<![A-Z0-9_])\$?[A-Z]{1,3}\$?\d+(?![A-Z0-9_])",
            expression,
            re.I
        ):
            address = match.group(0)

            try:
                value = self.value(source, address)
                target = self.canonical(value)

                if target:
                    candidates.add(target)
            except Exception:
                pass

        return candidates

    def scan_formulas(self):
        print("Scanning formulas...", flush=True)

        for ws in self.wb.worksheets:
            source = ws.title
            print("Sheet:", source, flush=True)

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    formula = str(cell.value or "")
                    location = cell.coordinate

                    self.counts["Formula Cells"] += 1

                    if self.counts["Formula Cells"] % PROGRESS_INTERVAL == 0:
                        print(
                            "Processed:",
                            self.counts["Formula Cells"],
                            flush=True
                        )

                    # Explicit references
                    for target in self.direct_references(formula):
                        self.add_link(
                            source, target,
                            "Direct Formula",
                            location, formula
                        )

                    # INDIRECT
                    calls = list(find_calls(formula, "INDIRECT"))
                    self.counts["INDIRECT Calls"] += len(calls)

                    for expression in calls:
                        try:
                            target, resolved = self.resolve_indirect(
                                source, cell, expression
                            )

                            self.add_link(
                                source, target,
                                "INDIRECT",
                                location,
                                resolved + " | " + formula
                            )

                        except Exception as exc:
                            candidates = self.candidate_sheets(
                                source, cell, expression
                            )

                            for target in candidates:
                                self.possible[source].add(target)

                            self.add_unresolved(
                                source, location,
                                "INDIRECT",
                                formula,
                                str(exc)
                            )

                    # OFFSET
                    offsets = list(find_calls(formula, "OFFSET"))
                    self.counts["OFFSET Calls"] += len(offsets)

                    for expression in offsets:
                        try:
                            target = self.resolve_offset(
                                source, cell, expression
                            )

                            self.add_link(
                                source, target,
                                "OFFSET",
                                location, formula
                            )

                        except Exception as exc:
                            candidates = self.candidate_sheets(
                                source, cell, expression
                            )

                            for target in candidates:
                                self.possible[source].add(target)

                            self.add_unresolved(
                                source, location,
                                "OFFSET",
                                formula,
                                str(exc)
                            )

                    # External workbook references
                    if re.search(r"\[[^\]]+\.xls", formula, re.I):
                        self.add_unresolved(
                            source, location,
                            "External Workbook",
                            formula,
                            "External workbook dependency"
                        )

    # ========================================================
    # Hyperlinks
    # ========================================================

    def scan_hyperlinks(self):
        print("Scanning hyperlinks...", flush=True)

        # Read worksheet XML directly. This also covers hidden sheets.
        ns = {
            "m": "http://schemas.openxmlformats.org/"
                 "spreadsheetml/2006/main"
        }

        with zipfile.ZipFile(self.path) as archive:
            for i, source in enumerate(self.names, 1):
                path = f"xl/worksheets/sheet{i}.xml"

                if path not in archive.namelist():
                    continue

                try:
                    root = ET.fromstring(archive.read(path))
                except Exception:
                    continue

                for link in root.findall(".//m:hyperlink", ns):
                    location = link.attrib.get("ref", "")
                    target_text = link.attrib.get("location", "")

                    if not target_text:
                        continue

                    if "!" in target_text:
                        target = target_text.rsplit("!", 1)[0]

                        self.add_navigation(
                            source, target,
                            "Hyperlink",
                            location, target_text
                        )

    # ========================================================
    # VBA
    # ========================================================

    def scan_vba(self):
        print("Scanning VBA...", flush=True)

        try:
            from oletools.olevba import VBA_Parser
        except ImportError:
            print("oletools unavailable; VBA analysis skipped")
            return

        parser = None

        try:
            parser = VBA_Parser(str(self.path))

            if not parser.detect_vba_macros():
                return

            procedure_pattern = re.compile(
                r"(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?"
                r"(?:Sub|Function)\s+([^\s(]+)"
            )

            for _, stream, module, code in parser.extract_macros():
                self.counts["VBA Modules"] += 1

                if isinstance(code, bytes):
                    code = code.decode("utf-8", errors="replace")

                self.vba_code.append((module, code))

                matches = list(procedure_pattern.finditer(code))

                for i, match in enumerate(matches):
                    start = match.end()
                    end = (
                        matches[i + 1].start()
                        if i + 1 < len(matches)
                        else len(code)
                    )

                    proc_name = match.group(1)
                    body = code[start:end]

                    self.procedures[norm(proc_name)] = (
                        module, body
                    )

                    self.counts["VBA Procedures"] += 1

            # Detect sheet references and navigation in procedures
            sheet_pattern = re.compile(
                r'(?:ThisWorkbook\.)?(?:Worksheets|Sheets)'
                r'\s*\(\s*"((?:[^"]|"")*)"\s*\)',
                re.I
            )

            for proc_name, (module, body) in self.procedures.items():
                matches = list(sheet_pattern.finditer(body))

                for match in matches:
                    target = self.canonical(
                        match.group(1).replace('""', '"')
                    )

                    if not target:
                        continue

                    following = body[match.end():match.end() + 100]

                    if re.match(r"\s*\.(Activate|Select)\b", following, re.I):
                        # A procedure can be invoked from many sheets.
                        # Do not invent a source sheet here.
                        self.macro_bindings.append(
                            (module, proc_name, target, "Activate/Select")
                        )

                    else:
                        self.macro_bindings.append(
                            (module, proc_name, target, "Sheet Access")
                        )

        except Exception as exc:
            self.add_unresolved(
                "(VBA)", "", "VBA",
                "", str(exc)
            )

        finally:
            if parser is not None:
                try:
                    parser.close()
                except Exception:
                    pass

    # ========================================================
    # Shape / button bindings
    # ========================================================

    def scan_shapes(self):
        print("Scanning shape metadata...", flush=True)

        with zipfile.ZipFile(self.path) as archive:
            files = archive.namelist()

            # VML buttons
            for filename in files:
                if not filename.lower().endswith(".vml"):
                    continue

                try:
                    root = ET.fromstring(archive.read(filename))
                except Exception:
                    continue

                for node in root.iter():
                    if node.tag.rsplit("}", 1)[-1] != "FmlaMacro":
                        continue

                    macro = (node.text or "").strip()

                    if macro:
                        self.counts["Macro Bindings"] += 1
                        self.macro_bindings.append(
                            (filename, macro, "", "VML Button")
                        )

            # Drawing hyperlinks are reported as metadata evidence.
            for filename in files:
                if not re.search(r"xl/drawings/drawing\d+\.xml$", filename):
                    continue

                try:
                    root = ET.fromstring(archive.read(filename))
                except Exception:
                    continue

                for node in root.iter():
                    local = node.tag.rsplit("}", 1)[-1]

                    if local in ("hlinkClick", "hlinkHover"):
                        self.macro_bindings.append(
                            (filename, str(node.attrib), "", "Drawing Link")
                        )
                        self.counts["Drawing Links"] += 1

    # ========================================================
    # Report
    # ========================================================

    @staticmethod
    def make_table(ws, headers, rows):
        ws.append(headers)

        for row in rows:
            ws.append([
                str(v) if isinstance(v, (dict, list, set)) else v
                for v in row
            ])

        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(wrap_text=True)

        for col in ws.columns:
            letter = col[0].column_letter
            maximum = max(
                (len(str(c.value or "")) for c in list(col)[:100]),
                default=12
            )
            ws.column_dimensions[letter].width = min(
                max(maximum + 2, 14), 65
            )

    def make_matrix(self, wb, title, mapping, possible=False):
        ws = wb.create_sheet(title)

        ws.cell(1, 1, "Source \\ Target")

        for j, name in enumerate(self.names, 2):
            cell = ws.cell(1, j, name)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.alignment = Alignment(
                text_rotation=45,
                horizontal="center"
            )
            ws.column_dimensions[get_column_letter(j)].width = 17

        for i, source in enumerate(self.names, 2):
            ws.cell(i, 1, source)

            for j, target in enumerate(self.names, 2):
                cell = ws.cell(i, j)

                if target in mapping[source]:
                    cell.value = 1
                    cell.fill = PatternFill(
                        "solid", fgColor="C6EFCE"
                    )

                elif possible and target in self.possible[source]:
                    cell.value = "?"
                    cell.fill = PatternFill(
                        "solid", fgColor="FFF2CC"
                    )

                else:
                    cell.value = 0

                cell.alignment = Alignment(horizontal="center")

        ws.column_dimensions["A"].width = 42
        ws.row_dimensions[1].height = 130
        ws.freeze_panes = "B2"

    def write_report(self, output):
        print("Generating report...", flush=True)

        wb = Workbook()
        summary = wb.active
        summary.title = "Summary"

        confirmed_pairs = sum(
            len(targets) for targets in self.direct.values()
        )

        navigation_pairs = sum(
            len(targets) for targets in self.navigation.values()
        )

        metrics = [
            ("Source File", self.path.name),
            ("Total Sheets", len(self.names)),
            ("Confirmed Data Links", confirmed_pairs),
            ("Confirmed Navigation Links", navigation_pairs),
            ("Formula Cells", self.counts["Formula Cells"]),
            ("INDIRECT Calls", self.counts["INDIRECT Calls"]),
            ("OFFSET Calls", self.counts["OFFSET Calls"]),
            ("VBA Modules", self.counts["VBA Modules"]),
            ("VBA Procedures", self.counts["VBA Procedures"]),
            ("Macro Bindings", self.counts["Macro Bindings"]),
            ("Evidence Rows", len(self.evidence)),
            ("Unresolved Rows", self.counts["Unresolved Rows"]),
            ("Unknown Source Sheets", len(self.unknown_sources)),
            ("Truncated Evidence", self.counts["Truncated Evidence"]),
            ("Truncated Unresolved", self.counts["Truncated Unresolved"]),
        ]

        self.make_table(summary, ["Metric", "Value"], metrics)

        self.make_matrix(
            wb, "Data_Matrix",
            self.direct,
            possible=True
        )

        self.make_matrix(
            wb, "Navigation_Matrix",
            self.navigation
        )

        self.make_table(
            wb.create_sheet("Evidence"),
            ["Source", "Target", "Type", "Location", "Detail"],
            self.evidence
        )

        self.make_table(
            wb.create_sheet("Unresolved"),
            ["Sheet", "Location", "Category", "Expression", "Reason"],
            self.unresolved
        )

        self.make_table(
            wb.create_sheet("Sheet_List"),
            ["Sheet", "Visibility", "Data Targets", "Navigation Targets"],
            [
                (
                    name,
                    self.wb[name].sheet_state,
                    len(self.direct[name]),
                    len(self.navigation[name])
                )
                for name in self.names
            ]
        )

        self.make_table(
            wb.create_sheet("VBA_Bindings"),
            ["Module/Location", "Procedure/Macro", "Target", "Type"],
            self.macro_bindings
        )

        self.make_table(
            wb.create_sheet("Unknown_Sources"),
            ["Source", "Possible Targets", "Confirmed Targets"],
            [
                (
                    source,
                    ", ".join(sorted(self.possible[source])),
                    ", ".join(sorted(self.direct[source]))
                )
                for source in sorted(self.unknown_sources)
            ]
        )

        wb.save(output)

    def run(self, output):
        self.scan_formulas()
        self.scan_hyperlinks()
        self.scan_vba()
        self.scan_shapes()
        self.write_report(output)

        print()
        print("========== ANALYSIS COMPLETE ==========")
        print("Sheets:", len(self.names))
        print("Formula cells:", self.counts["Formula Cells"])
        print("INDIRECT calls:", self.counts["INDIRECT Calls"])
        print("OFFSET calls:", self.counts["OFFSET Calls"])
        print(
            "Confirmed data links:",
            sum(len(v) for v in self.direct.values())
        )
        print("Unresolved:", self.counts["Unresolved Rows"])
        print("Report:", output)
        print("=======================================")


def main():
    parser = argparse.ArgumentParser(
        description="Analyze XLSM worksheet dependencies"
    )

    parser.add_argument("input", help="Source .xlsm workbook")
    parser.add_argument(
        "-o", "--output",
        default="input_dependency_report.xlsx"
    )

    args = parser.parse_args()

    source = Path(args.input).resolve()
    output = Path(args.output).resolve()

    if not source.is_file():
        print("Input file not found:", source)
        sys.exit(1)

    if source == output:
        print("Output must differ from source")
        sys.exit(1)

    analyzer = Analyzer(source)

    try:
        analyzer.run(output)
    finally:
        analyzer.wb.close()
        analyzer.values.close()


if __name__ == "__main__":
    main()
