
import argparse
import re
import sys
import zipfile
import posixpath
from pathlib import Path
from collections import Counter, defaultdict
from xml.etree import ElementTree as ET

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.cell import range_boundaries, get_column_letter

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    VBA_Parser = None


# ============================================================
# CONFIG
# ============================================================

MAX_EXPRESSION_LENGTH = 32000
MAX_RECURSION = 12
MAX_EVIDENCE = 500000

SHEET_REF_RE = re.compile(
    r"(?:'((?:[^']|'')+)'|"
    r"([\w\u0080-\uffff.]+))!",
    re.UNICODE
)

CELL_REF_RE = re.compile(
    r"^\$?([A-Z]{1,3})\$?(\d+)$",
    re.I
)

VBA_SHEET_RE = re.compile(
    r'(?:Sheets|Worksheets)\s*\(\s*"([^"]+)"\s*\)',
    re.I
)

VBA_PROC_RE = re.compile(
    r'(?im)^\s*(?:(?:Public|Private|Friend|Static)\s+)*'
    r'(?:Sub|Function)\s+([A-Za-z_]\w*)\s*\('
)

VBA_CALL_RE = re.compile(
    r'(?im)^\s*(?:Call\s+)?([A-Za-z_]\w*)\b'
)

FUNC_RE = re.compile(r'\b(INDIRECT|OFFSET|INDEX)\s*\(', re.I)

STRING_RE = re.compile(r'"(?:[^"]|"")*"')

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


# ============================================================
# BASIC UTILITIES
# ============================================================

def norm_name(value):
    if value is None:
        return ""
    return str(value).strip().strip("'").replace("''", "'")


def split_top_level(expr, delimiter=","):
    result = []
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

        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == delimiter and depth == 0:
                result.append("".join(buf).strip())
                buf = []
                i += 1
                continue

        buf.append(ch)
        i += 1

    result.append("".join(buf).strip())
    return result


def extract_calls(formula, func_name):
    pattern = re.compile(r"\b" + re.escape(func_name) + r"\s*\(", re.I)
    result = []

    for m in pattern.finditer(formula):
        start = m.end()
        depth = 1
        quoted = False
        i = start

        while i < len(formula):
            ch = formula[i]

            if ch == '"':
                if quoted and i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 2
                    continue
                quoted = not quoted

            elif not quoted:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        result.append(formula[start:i])
                        break

            i += 1

    return result


def remove_string_literals(formula):
    return STRING_RE.sub('""', formula)


def extract_sheet_refs(formula, valid_names):
    cleaned = remove_string_literals(formula)
    found = set()

    for m in SHEET_REF_RE.finditer(cleaned):
        name = norm_name(m.group(1) or m.group(2))
        actual = valid_names.get(name.casefold())
        if actual:
            found.add(actual)

    return found


def text_cell(cell, value):
    cell.number_format = "@"
    s = "" if value is None else str(value)
    cell.value = s[:32767]


def normalize_macro_name(name):
    s = str(name or "").strip().strip("'")
    s = s.split("!")[-1]
    s = s.split(".")[-1]
    return s.strip().casefold()


# ============================================================
# MAIN ANALYZER
# ============================================================

class XLSMAnalyzer:

    def __init__(self, filename):
        self.path = Path(filename).resolve()

        print("Loading workbook...")

        self.wb = load_workbook(
            self.path,
            data_only=False,
            read_only=False,
            keep_links=False
        )

        self.cached_wb = load_workbook(
            self.path,
            data_only=True,
            read_only=True,
            keep_links=False
        )

        self.sheets = list(self.wb.sheetnames)

        self.sheet_lookup = {
            s.casefold(): s for s in self.sheets
        }

        self.data_links = set()
        self.nav_links = set()
        self.possible_links = set()

        self.evidence = []
        self.unresolved = []
        self.indirect_details = []

        self.procedures = {}
        self.macro_targets = defaultdict(set)
        self.macro_calls = defaultdict(set)

        self.shape_bindings = []
        self.stats = Counter()
        self.dedup = set()

    # --------------------------------------------------------
    # RELATION STORAGE
    # --------------------------------------------------------

    def add_link(
        self, source, target, category,
        location, expression, method,
        confidence="CONFIRMED"
    ):
        target = self.sheet_lookup.get(str(target).casefold())

        if not target or source not in self.sheets:
            return

        if confidence == "POSSIBLE":
            self.possible_links.add((source, target))
        elif category == "DATA":
            self.data_links.add((source, target))
        else:
            self.nav_links.add((source, target))

        key = (
            source, target, category,
            location, method, confidence
        )

        if key in self.dedup:
            return

        self.dedup.add(key)

        if len(self.evidence) < MAX_EVIDENCE:
            self.evidence.append((
                source, target, category,
                location, str(expression),
                method, confidence
            ))

    def add_unresolved(
        self, sheet, location,
        category, expression, reason
    ):
        self.unresolved.append((
            sheet, location, category,
            str(expression), reason
        ))

    # --------------------------------------------------------
    # VALUE ACCESS
    # --------------------------------------------------------

    def get_value(self, sheet, address):
        try:
            if sheet not in self.cached_wb.sheetnames:
                return None

            value = self.cached_wb[sheet][address].value

            if value is not None:
                return value

            return self.wb[sheet][address].value

        except Exception:
            return None

    def resolve_sheet(self, value):
        name = norm_name(value)
        return self.sheet_lookup.get(name.casefold())

    # --------------------------------------------------------
    # SIMPLE EXPRESSION EVALUATOR
    # --------------------------------------------------------

    def evaluate(
        self, expression, source_sheet,
        current_cell, depth=0, visited=None
    ):
        if depth > MAX_RECURSION:
            return None

        if visited is None:
            visited = set()

        expr = str(expression).strip()

        if expr.startswith("="):
            expr = expr[1:].strip()

        if not expr:
            return None

        # Parenthesized expression
        if expr.startswith("(") and expr.endswith(")"):
            inner = expr[1:-1]
            if len(split_top_level(inner, ",")) == 1:
                result = self.evaluate(
                    inner, source_sheet, current_cell,
                    depth + 1, visited
                )
                if result is not None:
                    return result

        # String concatenation
        parts = split_top_level(expr, "&")

        if len(parts) > 1:
            values = []

            for part in parts:
                v = self.evaluate(
                    part, source_sheet, current_cell,
                    depth + 1, visited
                )
                if v is None:
                    return None
                values.append(str(v))

            return "".join(values)

        # Literal string
        if re.fullmatch(r'"(?:[^"]|"")*"', expr):
            return expr[1:-1].replace('""', '"')

        # Number
        if re.fullmatch(r"-?\d+(?:\.\d+)?", expr):
            try:
                return int(expr)
            except ValueError:
                return float(expr)

        # ROW()
        if re.fullmatch(r"ROW\(\s*\)", expr, re.I):
            return self.wb[source_sheet][current_cell].row

        # COLUMN()
        if re.fullmatch(r"COLUMN\(\s*\)", expr, re.I):
            return self.wb[source_sheet][current_cell].column

        # Simple arithmetic: ROW()-ROW($E$35)
        m = re.fullmatch(
            r"ROW\(\s*\)\s*-\s*ROW\(\s*\$?[A-Z]+\$?(\d+)\s*\)",
            expr, re.I
        )
        if m:
            return (
                self.wb[source_sheet][current_cell].row
                - int(m.group(1))
            )

        # SUBSTITUTE(text, old, new)
        if re.match(r"^SUBSTITUTE\s*\(", expr, re.I):
            calls = extract_calls(expr, "SUBSTITUTE")

            if calls and expr.endswith(")"):
                args = split_top_level(calls[0])

                if len(args) >= 3:
                    vals = [
                        self.evaluate(
                            a, source_sheet, current_cell,
                            depth + 1, visited
                        )
                        for a in args[:3]
                    ]

                    if all(v is not None for v in vals):
                        return str(vals[0]).replace(
                            str(vals[1]), str(vals[2])
                        )

        # ADDRESS(row, column, [abs_num])
        if re.match(r"^ADDRESS\s*\(", expr, re.I):
            calls = extract_calls(expr, "ADDRESS")

            if calls:
                args = split_top_level(calls[0])

                if len(args) >= 2:
                    r = self.evaluate(
                        args[0], source_sheet,
                        current_cell, depth + 1, visited
                    )
                    c = self.evaluate(
                        args[1], source_sheet,
                        current_cell, depth + 1, visited
                    )

                    if isinstance(r, (int, float)) and \
                       isinstance(c, (int, float)):
                        if 1 <= int(c) <= 16384 and int(r) >= 1:
                            return (
                                "$" + get_column_letter(int(c))
                                + "$" + str(int(r))
                            )

        # Local or qualified cell reference
        m = re.fullmatch(
            r"(?:(?:'((?:[^']|'')+)'|([^!]+))!)?"
            r"(\$?[A-Z]{1,3}\$?\d+)",
            expr, re.I
        )

        if m:
            target_sheet = (
                self.resolve_sheet(m.group(1) or m.group(2))
                if (m.group(1) or m.group(2))
                else source_sheet
            )

            address = m.group(3).replace("$", "")

            if not target_sheet:
                return None

            visit_key = (target_sheet, address)

            if visit_key in visited:
                return None

            new_visited = visited | {visit_key}

            value = self.get_value(target_sheet, address)

            if isinstance(value, str) and value.startswith("="):
                return self.evaluate(
                    value, target_sheet, address,
                    depth + 1, new_visited
                )

            return value

        # Cached result for unsupported functions is not
        # evaluated here because there is no expression-cell
        # mapping for arbitrary nested expressions.
        return None

    # --------------------------------------------------------
    # INDIRECT
    # --------------------------------------------------------

    def analyze_indirect(self, sheet, cell, formula):
        calls = extract_calls(formula, "INDIRECT")

        for expression in calls:
            self.stats["INDIRECT calls"] += 1

            args = split_top_level(expression)

            if not args:
                continue

            target_expr = args[0]

            resolved = self.evaluate(
                target_expr, sheet, cell
            )

            target = None
            status = "UNRESOLVED"
            reason = ""

            if isinstance(resolved, str) and "!" in resolved:
                sheet_part = resolved.rsplit("!", 1)[0]

                # Remove external workbook prefix only when
                # present; do not assume external references
                # belong to this workbook.
                if "[" in sheet_part or "]" in sheet_part:
                    reason = "External workbook reference"
                else:
                    target = self.resolve_sheet(sheet_part)

                    if target:
                        self.add_link(
                            sheet, target, "DATA",
                            cell, formula,
                            "INDIRECT evaluated"
                        )
                        status = "CONFIRMED"
                    else:
                        reason = "Target sheet not found"

            elif isinstance(resolved, str):
                # INDIRECT("A1") is a local reference.
                if re.fullmatch(
                    r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?",
                    resolved, re.I
                ):
                    target = sheet
                    status = "CONFIRMED"

                    self.add_link(
                        sheet, sheet, "DATA",
                        cell, formula,
                        "INDIRECT local range"
                    )
                else:
                    reason = "No recognizable sheet target"

            else:
                reason = "Expression could not be evaluated"

            # Explicit sheet names in string fragments can
            # indicate possible targets, not confirmed targets.
            if status == "UNRESOLVED":
                for m in re.finditer(
                    r'"([^"]+)!', target_expr
                ):
                    candidate = self.resolve_sheet(m.group(1))

                    if candidate:
                        self.add_link(
                            sheet, candidate, "DATA",
                            cell, formula,
                            "INDIRECT literal candidate",
                            "POSSIBLE"
                        )

                self.add_unresolved(
                    sheet, cell, "INDIRECT",
                    formula, reason
                )

            self.indirect_details.append((
                sheet, cell, formula,
                target_expr,
                "" if resolved is None else str(resolved),
                target or "",
                status,
                reason
            ))

    # --------------------------------------------------------
    # FORMULAS
    # --------------------------------------------------------

    def scan_formulas(self):
        print("Scanning formulas...")

        for ws in self.wb.worksheets:
            print("  Sheet:", ws.title)

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["Formula cells"] += 1
                    formula = str(cell.value or "")

                    refs = extract_sheet_refs(
                        formula, self.sheet_lookup
                    )

                    for target in refs:
                        self.add_link(
                            ws.title, target, "DATA",
                            cell.coordinate, formula,
                            "Explicit formula"
                        )

                    if re.search(
                        r"\bINDIRECT\s*\(",
                        formula, re.I
                    ):
                        self.analyze_indirect(
                            ws.title, cell.coordinate,
                            formula
                        )

                    # OFFSET/INDEX have a reference
                    # argument. Explicit references are
                    # already captured above.
                    for fn in ("OFFSET", "INDEX"):
                        for call in extract_calls(formula, fn):
                            self.stats[fn + " calls"] += 1

                            args = split_top_level(call)

                            if args:
                                for target in extract_sheet_refs(
                                    args[0], self.sheet_lookup
                                ):
                                    self.add_link(
                                        ws.title, target,
                                        "DATA", cell.coordinate,
                                        formula,
                                        fn + " reference argument"
                                    )

                    if self.stats["Formula cells"] % 10000 == 0:
                        print(
                            "Processed:",
                            self.stats["Formula cells"]
                        )

    # --------------------------------------------------------
    # DEFINED NAMES
    # --------------------------------------------------------

    def scan_defined_names(self):
        print("Scanning defined names...")

        # Workbook-level defined names.
        try:
            definitions = list(
                self.wb.defined_names.values()
            )
        except Exception:
            definitions = []

        for definition in definitions:
            text = str(definition.attr_text or "")
            refs = extract_sheet_refs(
                text, self.sheet_lookup
            )

            # Workbook-level names do not have one
            # source sheet. Record diagnostic evidence.
            for target in refs:
                self.add_unresolved(
                    "", definition.name,
                    "Defined Name",
                    text,
                    "Workbook-scoped name references "
                    + target + "; usage requires tracing"
                )

    # --------------------------------------------------------
    # HYPERLINKS
    # --------------------------------------------------------

    def scan_hyperlinks(self):
        print("Scanning hyperlinks...")

        for ws in self.wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    link = cell.hyperlink

                    if link is None:
                        continue

                    sub = str(link.location or "")

                    if not sub and link.target:
                        sub = str(link.target)

                    if "!" not in sub:
                        continue

                    name = sub.rsplit("!", 1)[0]
                    target = self.resolve_sheet(name)

                    if target:
                        self.add_link(
                            ws.title, target, "NAV",
                            cell.coordinate, sub,
                            "Cell hyperlink"
                        )

    # --------------------------------------------------------
    # VBA
    # --------------------------------------------------------

    def scan_vba(self):
        print("Scanning VBA...")

        if VBA_Parser is None:
            self.add_unresolved(
                "", "", "VBA", "",
                "oletools not installed"
            )
            return

        parser = None

        try:
            parser = VBA_Parser(str(self.path))

            if not parser.detect_vba_macros():
                return

            for _, stream, module, code in parser.extract_macros():
                self.stats["VBA modules"] += 1

                if isinstance(code, bytes):
                    code = code.decode(
                        "utf-8", errors="replace"
                    )

                # Remove comment-only lines.
                code = re.sub(
                    r"(?m)^\s*'.*$", "", code
                )

                # Join VBA continuation lines.
                code = re.sub(
                    r"\s+_\s*\r?\n\s*",
                    " ", code
                )

                starts = list(
                    VBA_PROC_RE.finditer(code)
                )

                for idx, m in enumerate(starts):
                    name = m.group(1).casefold()

                    end = (
                        starts[idx + 1].start()
                        if idx + 1 < len(starts)
                        else len(code)
                    )

                    body = code[m.start():end]

                    # Prefer procedure bodies up to End Sub.
                    end_match = re.search(
                        r"(?im)^\s*End\s+(Sub|Function)\b",
                        body
                    )

                    if end_match:
                        body = body[:end_match.end()]

                    self.procedures[name] = (
                        self.procedures.get(name, "")
                        + "\n" + body
                    )

            # Direct target sheets.
            for proc, body in self.procedures.items():
                for m in VBA_SHEET_RE.finditer(body):
                    target = self.resolve_sheet(m.group(1))

                    if target:
                        self.macro_targets[proc].add(target)

                # Discover simple procedure calls.
                for m in VBA_CALL_RE.finditer(body):
                    called = m.group(1).casefold()

                    if called != proc:
                        self.macro_calls[proc].add(called)

            # Fixed-point propagation of macro calls.
            for _ in range(20):
                changed = False

                for proc, calls in self.macro_calls.items():
                    for called in calls:
                        before = len(self.macro_targets[proc])

                        self.macro_targets[proc].update(
                            self.macro_targets.get(
                                called, set()
                            )
                        )

                        if len(self.macro_targets[proc]) > before:
                            changed = True

                if not changed:
                    break

            self.stats["VBA procedures"] = len(
                self.procedures
            )

        except Exception as exc:
            self.add_unresolved(
                "", "", "VBA", "",
                "VBA parsing failed: " + str(exc)
            )

        finally:
            if parser is not None:
                try:
                    parser.close()
                except Exception:
                    pass

    # --------------------------------------------------------
    # SHAPE / VML METADATA
    # --------------------------------------------------------

    def scan_shapes(self):
        print("Scanning drawing and VML metadata...")

        with zipfile.ZipFile(self.path) as z:
            names = set(z.namelist())

            # Resolve worksheet XML to actual sheet names.
            workbook_xml = ET.fromstring(
                z.read("xl/workbook.xml")
            )

            rel_xml = ET.fromstring(
                z.read("xl/_rels/workbook.xml.rels")
            )

            rel_map = {}

            for rel in rel_xml:
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target", "")

                if target.startswith("/"):
                    target = target.lstrip("/")
                else:
                    target = posixpath.normpath(
                        posixpath.join("xl", target)
                    )

                rel_map[rid] = target

            sheet_paths = {}

            for item in workbook_xml.iter():
                if item.tag.endswith("}sheet"):
                    name = item.attrib.get("name")
                    rid = item.attrib.get(
                        "{" + NS_REL + "}id"
                    )

                    if rid in rel_map:
                        sheet_paths[rel_map[rid]] = name

            # Worksheet relationships connect each
            # drawing or VML file to its source sheet.
            drawing_sources = defaultdict(set)

            for sheet_path, sheet_name in sheet_paths.items():
                folder = posixpath.dirname(sheet_path)
                base = posixpath.basename(sheet_path)

                rel_path = (
                    folder + "/_rels/" + base + ".rels"
                )

                if rel_path not in names:
                    continue

                root = ET.fromstring(z.read(rel_path))

                for rel in root:
                    target = rel.attrib.get("Target", "")
                    kind = rel.attrib.get("Type", "")

                    if "drawing" not in kind.lower():
                        continue

                    if target.startswith("/"):
                        full = target.lstrip("/")
                    else:
                        full = posixpath.normpath(
                            posixpath.join(folder, target)
                        )

                    drawing_sources[full].add(sheet_name)

            for path, sources in drawing_sources.items():
                if path not in names:
                    continue

                try:
                    root = ET.fromstring(z.read(path))
                except ET.ParseError:
                    continue

                for node in root.iter():
                    tag = node.tag.split("}")[-1].lower()

                    # VML FmlaMacro or Drawing macro attribute.
                    macro = ""

                    if tag == "fmlamacro":
                        macro = node.text or ""

                    if not macro:
                        for k, v in node.attrib.items():
                            if k.split("}")[-1].lower() == "macro":
                                macro = v
                                break

                    if not macro:
                        continue

                    self.stats["Shape macro bindings"] += 1

                    for source in sources:
                        self.shape_bindings.append(
                            (source, path, macro)
                        )

                        self.resolve_shape_macro(
                            source, path, macro
                        )

    def resolve_shape_macro(
        self, source, location, macro
    ):
        proc = normalize_macro_name(macro)

        targets = self.macro_targets.get(proc, set())

        if targets:
            for target in targets:
                self.add_link(
                    source, target, "NAV",
                    location, macro,
                    "Shape macro binding"
                )
        else:
            self.add_unresolved(
                source, location,
                "Shape Macro", macro,
                "Macro target not statically resolved"
            )

    # --------------------------------------------------------
    # REPORT
    # --------------------------------------------------------

    def create_matrix(self, wb, title, links):
        ws = wb.create_sheet(title)

        ws.cell(1, 1, "Source / Target")

        for i, name in enumerate(self.sheets, 2):
            ws.cell(1, i, name)
            ws.cell(i, 1, name)

        for i, source in enumerate(self.sheets, 2):
            for j, target in enumerate(self.sheets, 2):
                ws.cell(
                    i, j,
                    1 if (source, target) in links else 0
                )

        ws.freeze_panes = "B2"
        return ws

    def write_rows(self, wb, title, headers, rows):
        ws = wb.create_sheet(title)
        ws.append(headers)

        for row in rows:
            for col, value in enumerate(row, 1):
                text_cell(
                    ws.cell(ws.max_row + (1 if col == 1 else 0), col)
                    if col == 1 else ws.cell(ws.max_row, col),
                    value
                )

        ws.freeze_panes = "A2"
        return ws

    def report(self):
        print("Generating report...")

        out = Workbook()
        out.remove(out.active)

        all_links = self.data_links | self.nav_links

        self.create_matrix(
            out, "Direct_Matrix", all_links
        )

        self.create_matrix(
            out, "Data_Matrix", self.data_links
        )

        self.create_matrix(
            out, "Navigation_Matrix", self.nav_links
        )

        self.create_matrix(
            out, "Possible_Matrix", self.possible_links
        )

        self.write_rows(
            out, "Dependency_Details",
            [
                "Source Sheet", "Target Sheet",
                "Category", "Location",
                "Expression", "Method", "Confidence"
            ],
            self.evidence
        )

        self.write_rows(
            out, "INDIRECT_Details",
            [
                "Sheet", "Cell", "Formula",
                "INDIRECT Argument", "Resolved Value",
                "Target Sheet", "Status", "Reason"
            ],
            self.indirect_details
        )

        self.write_rows(
            out, "Unresolved_References",
            [
                "Sheet", "Location", "Category",
                "Expression", "Reason"
            ],
            self.unresolved
        )

        self.write_rows(
            out, "Shape_Bindings",
            ["Source Sheet", "Drawing XML", "Macro"],
            self.shape_bindings
        )

        self.write_rows(
            out, "Sheet_Information",
            ["Sheet", "Visibility", "Index"],
            [
                (ws.title, ws.sheet_state, i)
                for i, ws in enumerate(
                    self.wb.worksheets, 1
                )
            ]
        )

        statistics = [
            ("Sheets", len(self.sheets)),
            ("Confirmed data links", len(self.data_links)),
            ("Confirmed navigation links", len(self.nav_links)),
            ("Possible links", len(self.possible_links)),
            ("Evidence rows", len(self.evidence)),
            ("Unresolved rows", len(self.unresolved))
        ]

        statistics.extend(
            sorted(self.stats.items())
        )

        self.write_rows(
            out, "Statistics",
            ["Metric", "Count"],
            statistics
        )

        # Formatting
        for ws in out.worksheets:
            ws.sheet_view.showGridLines = False

            for cell in ws[1]:
                cell.fill = PatternFill(
                    "solid", fgColor="17365D"
                )
                cell.font = Font(
                    bold=True, color="FFFFFF"
                )
                cell.alignment = Alignment(
                    vertical="center"
                )

            ws.row_dimensions[1].height = 28

            if ws.title.endswith("_Matrix"):
                ws.column_dimensions["A"].width = 38

                for c in range(2, ws.max_column + 1):
                    ws.column_dimensions[
                        get_column_letter(c)
                    ].width = 17

            else:
                for c in range(1, ws.max_column + 1):
                    ws.column_dimensions[
                        get_column_letter(c)
                    ].width = 32

            ws.auto_filter.ref = ws.dimensions

        output_path = self.path.with_name(
            self.path.stem
            + "_dependency_report.xlsx"
        )

        out.save(output_path)

        return output_path

    # --------------------------------------------------------
    # MAIN
    # --------------------------------------------------------

    def run(self):
        self.scan_formulas()
        self.scan_defined_names()
        self.scan_hyperlinks()
        self.scan_vba()
        self.scan_shapes()

        output = self.report()

        print()
        print("=" * 55)
        print("ANALYSIS COMPLETED")
        print("=" * 55)

        print("Sheets:", len(self.sheets))
        print(
            "Formula cells:",
            self.stats["Formula cells"]
        )
        print(
            "INDIRECT calls:",
            self.stats["INDIRECT calls"]
        )
        print(
            "Confirmed data links:",
            len(self.data_links)
        )
        print(
            "Confirmed navigation links:",
            len(self.nav_links)
        )
        print(
            "Possible links:",
            len(self.possible_links)
        )
        print(
            "Unresolved rows:",
            len(self.unresolved)
        )
        print("Report:", output)
        print("=" * 55)

        self.cached_wb.close()
        self.wb.close()


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Analyze XLSM sheet dependencies"
    )

    parser.add_argument(
        "file",
        help="Path to source XLSM file"
    )

    args = parser.parse_args()

    path = Path(args.file)

    if not path.exists():
        print("File not found:", path)
        sys.exit(1)

    if path.suffix.lower() != ".xlsm":
        print("Only XLSM files are supported.")
        sys.exit(1)

    analyzer = XLSMAnalyzer(path)
    analyzer.run()


if __name__ == "__main__":
    main()
