
import argparse
import re
import sys
import zipfile
from collections import defaultdict, Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import Cell
from oletools.olevba import VBA_Parser


# ============================================================
# Configuration
# ============================================================

MAX_DETAILS = 200000
MAX_UNRESOLVED = 50000
MAX_EXPRESSION = 16000
PROGRESS_INTERVAL = 10000

# True: include formulas referencing their own sheet
INCLUDE_SELF_REFERENCE = False

# True: add candidate references to matrix
# Recommended False to avoid false positives
INCLUDE_CANDIDATES_IN_MATRIX = False


# ============================================================
# Utilities
# ============================================================

def log(message):
    print(message, flush=True)


def clean_text(value):
    if value is None:
        return ""
    return str(value).replace("\x00", "")


def safe_text(value, limit=MAX_EXPRESSION):
    text = clean_text(value)[:limit]

    # Explicitly store as text, not an Excel formula.
    # Prefixing with apostrophe prevents formula interpretation
    # in spreadsheet applications.
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text

    return text


def normalized_sheet_name(value):
    value = clean_text(value).strip()
    value = value.replace("''", "'")
    return value


def make_lookup(names):
    return {name.casefold(): name for name in names}


def find_sheet(name, lookup):
    name = normalized_sheet_name(name)
    return lookup.get(name.casefold())


def strip_vba_comments(line):
    """Remove VBA comments outside quoted strings."""
    result = []
    quoted = False
    i = 0

    while i < len(line):
        ch = line[i]

        if ch == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                result.append('""')
                i += 2
                continue
            quoted = not quoted

        if not quoted and ch == "'":
            break

        result.append(ch)
        i += 1

    return "".join(result)


def split_arguments(text):
    """Split function arguments, respecting strings and parentheses."""
    parts = []
    depth = 0
    quoted = False
    start = 0
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == '"':
            if quoted and i + 1 < len(text) and text[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted

        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch in ",;" and depth == 0:
                parts.append(text[start:i].strip())
                start = i + 1

        i += 1

    parts.append(text[start:].strip())
    return parts


def function_calls(formula, function_name):
    """Extract nested function argument strings."""
    pattern = re.compile(
        r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?"
        + re.escape(function_name)
        + r"\s*\(",
        re.I,
    )

    for match in pattern.finditer(formula):
        start = match.end()
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
                        yield formula[start:i]
                        break

            i += 1


# ============================================================
# Dependency Analyzer
# ============================================================

class XLSMAnalyzer:

    def __init__(self, source):
        self.source = Path(source).resolve()

        if not self.source.is_file():
            raise FileNotFoundError(self.source)

        self.book = None
        self.names = []
        self.lookup = {}
        self.code_to_sheet = {}

        self.edges = defaultdict(set)
        self.details = []
        self.unresolved = []
        self.counts = Counter()

        self.vba_procedures = {}
        self.vba_modules = []
        self.macro_bindings = []
        self.cell_values = {}
        self.defined_names = defaultdict(set)

        self.truncated_details = 0
        self.truncated_unresolved = 0

    # --------------------------------------------------------
    # Recording
    # --------------------------------------------------------

    def add_edge(
        self,
        source,
        target,
        category,
        location,
        expression,
        confidence="CONFIRMED",
    ):
        source = find_sheet(source, self.lookup)
        target = find_sheet(target, self.lookup)

        if not source or not target:
            return

        if source == target and not INCLUDE_SELF_REFERENCE:
            return

        if (
            confidence == "CONFIRMED"
            or INCLUDE_CANDIDATES_IN_MATRIX
        ):
            self.edges[source].add(target)

        self.counts["evidence_" + category] += 1

        if len(self.details) < MAX_DETAILS:
            self.details.append((
                source,
                target,
                category,
                confidence,
                location,
                safe_text(expression),
            ))
        else:
            self.truncated_details += 1

    def add_unresolved(
        self,
        source,
        location,
        category,
        expression,
        reason,
    ):
        self.counts["unresolved_" + category] += 1

        if len(self.unresolved) < MAX_UNRESOLVED:
            self.unresolved.append((
                source,
                location,
                category,
                safe_text(expression),
                reason,
            ))
        else:
            self.truncated_unresolved += 1

    # --------------------------------------------------------
    # Load workbook
    # --------------------------------------------------------

    def load(self):
        log("Loading workbook...")

        # Normal mode is intentional:
        # ReadOnlyWorksheet does not expose all sheet metadata.
        self.book = load_workbook(
            self.source,
            read_only=False,
            data_only=False,
            keep_links=True,
        )

        self.names = list(self.book.sheetnames)
        self.lookup = make_lookup(self.names)

        for ws in self.book.worksheets:
            self.code_to_sheet[ws.title.casefold()] = ws.title

            try:
                self.code_to_sheet[
                    ws.sheet_properties.codeName.casefold()
                ] = ws.title
            except (AttributeError, TypeError):
                pass

        log(f"Sheets: {len(self.names)}")

        for ws in self.book.worksheets:
            log(f"  {ws.title} [{ws.sheet_state}]")

    # --------------------------------------------------------
    # Defined names
    # --------------------------------------------------------

    def scan_defined_names(self):
        log("Scanning defined names...")

        # Global names
        for name, definition in self.book.defined_names.items():
            self._process_defined_name(name, definition)

        # Worksheet-local names
        for ws in self.book.worksheets:
            for name, definition in ws.defined_names.items():
                self._process_defined_name(
                    name, definition, ws.title
                )

    def _process_defined_name(self, name, definition, owner=None):
        value = clean_text(
            getattr(definition, "attr_text", "")
        )

        targets = self.explicit_targets(value)

        for target in targets:
            self.defined_names[name.casefold()].add(target)

            if owner:
                self.add_edge(
                    owner, target,
                    "DefinedName",
                    name, value,
                )

    # --------------------------------------------------------
    # Explicit sheet references
    # --------------------------------------------------------

    def explicit_targets(self, expression):
        """
        Match:
          Sheet1!A1
          'Sheet A'!A1
          '[Book.xlsm]Sheet A'!A1
          Sheet1:Sheet3!A1

        Excludes sheet-name-looking text without !.
        """
        found = set()

        pattern = re.compile(
            r"(?:"
            r"'((?:[^']|'')+)'"
            r"|"
            r"(?<![A-Za-z0-9_.])"
            r"([^\s=+\-*/(),;:&\"'!]+)"
            r")\s*!",
            re.UNICODE,
        )

        for match in pattern.finditer(expression):
            raw = match.group(1) or match.group(2)

            if "]" in raw:
                raw = raw.rsplit("]", 1)[-1]

            target = find_sheet(raw, self.lookup)

            if target:
                found.add(target)
                continue

            # 3D references
            if ":" in raw:
                a, b = raw.split(":", 1)
                a = find_sheet(a, self.lookup)
                b = find_sheet(b, self.lookup)

                if a and b:
                    i = self.names.index(a)
                    j = self.names.index(b)

                    found.update(
                        self.names[min(i, j):max(i, j) + 1]
                    )

        return found

    # --------------------------------------------------------
    # Static resolution of dynamic formulas
    # --------------------------------------------------------

    def string_literals(self, expression):
        pattern = re.compile(r'"((?:[^"]|"")*)"')
        return [
            x.replace('""', '"')
            for x in pattern.findall(expression)
        ]

    def resolve_indirect(self, formula, current_sheet):
        """
        Conservative resolver.

        Handles literal sheet names and some cell-based
        sheet selectors. Does not execute Excel formulas.
        """
        confirmed = set()
        candidates = set()

        for call in function_calls(formula, "INDIRECT"):
            args = split_arguments(call)

            if not args:
                continue

            expr = args[0]

            # Literal sheet references
            confirmed.update(self.explicit_targets(expr))

            # Examples:
            # INDIRECT("'サマリ'!A1")
            # INDIRECT("サマリ!$G" & ROW())
            # INDIRECT("'サマリ'!$G" & ROW() & ...)
            for literal in self.string_literals(expr):
                confirmed.update(
                    self.explicit_targets(literal)
                )

            # Case:
            # INDIRECT("'" & $E36 & "'!$H$119:$W$123")
            #
            # Resolve selector from a literal cell value.
            selector_pattern = re.compile(
                r'(?:"[^"]*"\s*&\s*)?'
                r'(\$?[A-Z]{1,3}\$?\d+)'
                r'\s*&\s*"[^"]*!"',
                re.I,
            )

            for match in selector_pattern.finditer(expr):
                address = match.group(1).replace("$", "")
                value = self.cell_values.get(
                    (current_sheet, address)
                )

                if isinstance(value, str):
                    target = find_sheet(value, self.lookup)

                    if target:
                        candidates.add(target)

            # Also recognize literal text values that are
            # exactly worksheet names.
            for literal in self.string_literals(expr):
                target = find_sheet(literal, self.lookup)

                if target:
                    candidates.add(target)

        return confirmed, candidates

    def scan_formulas(self):
        log("Scanning formulas...")

        # Collect literal values used by selectors.
        for ws in self.book.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue

                    if cell.data_type != "f":
                        self.cell_values[
                            (ws.title, cell.coordinate)
                        ] = cell.value

        processed = 0

        for ws in self.book.worksheets:
            log(f"Sheet: {ws.title}")

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    formula = clean_text(cell.value)
                    location = cell.coordinate

                    processed += 1
                    self.counts["formula_cells"] += 1

                    if processed % PROGRESS_INTERVAL == 0:
                        log(f"Processed formulas: {processed}")

                    explicit = self.explicit_targets(formula)

                    for target in explicit:
                        self.add_edge(
                            ws.title, target,
                            "Formula",
                            location, formula,
                        )

                    # Defined-name references
                    for name, targets in self.defined_names.items():
                        if re.search(
                            r"(?<![\w.])"
                            + re.escape(name)
                            + r"(?![\w.])",
                            formula,
                            re.I,
                        ):
                            for target in targets:
                                self.add_edge(
                                    ws.title, target,
                                    "DefinedNameFormula",
                                    location, formula,
                                )

                    if re.search(
                        r"(?:_xlfn\.)?INDIRECT\s*\(",
                        formula,
                        re.I,
                    ):
                        self.counts["indirect_calls"] += 1

                        confirmed, candidates = (
                            self.resolve_indirect(
                                formula, ws.title
                            )
                        )

                        for target in confirmed - explicit:
                            self.add_edge(
                                ws.title, target,
                                "INDIRECT",
                                location, formula,
                            )

                        for target in candidates - confirmed:
                            self.add_edge(
                                ws.title, target,
                                "INDIRECT_SELECTOR",
                                location, formula,
                                "CANDIDATE",
                            )

                        if not confirmed and not candidates:
                            self.add_unresolved(
                                ws.title,
                                location,
                                "INDIRECT",
                                formula,
                                "Runtime expression cannot be "
                                "fully resolved statically",
                            )

                    if re.search(
                        r"(?:_xlfn\.)?OFFSET\s*\(",
                        formula,
                        re.I,
                    ):
                        self.counts["offset_calls"] += 1

                        # OFFSET normally uses its first
                        # argument as the base reference.
                        # If explicit_targets already finds
                        # the sheet, it is recorded above.
                        if not explicit:
                            self.add_unresolved(
                                ws.title,
                                location,
                                "OFFSET",
                                formula,
                                "No explicit cross-sheet "
                                "base reference",
                            )

                    if (
                        "INDIRECT(" not in formula.upper()
                        and "OFFSET(" not in formula.upper()
                        and not explicit
                    ):
                        self.counts["local_or_unknown"] += 1

        log(f"Formula cells: {processed}")

    # --------------------------------------------------------
    # Hyperlinks
    # --------------------------------------------------------

    def scan_hyperlinks(self):
        log("Scanning hyperlinks...")

        for ws in self.book.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    link = cell.hyperlink

                    if link is None:
                        continue

                    text = (
                        clean_text(link.location)
                        or clean_text(link.target)
                    )

                    targets = self.explicit_targets(text)

                    for target in targets:
                        self.add_edge(
                            ws.title, target,
                            "Hyperlink",
                            cell.coordinate,
                            text,
                        )

                    if not targets and "#" in text:
                        fragment = text.split("#", 1)[-1]

                        for target in self.explicit_targets(
                            fragment
                        ):
                            self.add_edge(
                                ws.title, target,
                                "Hyperlink",
                                cell.coordinate,
                                text,
                            )

    # --------------------------------------------------------
    # VBA extraction
    # --------------------------------------------------------

    def scan_vba(self):
        log("Scanning VBA...")

        parser = VBA_Parser(str(self.source))

        try:
            if not parser.detect_vba_macros():
                log("No VBA macros found.")
                return

            for _, stream, module_name, code in (
                parser.extract_macros()
            ):
                self.vba_modules.append((
                    clean_text(module_name),
                    clean_text(code),
                ))

        finally:
            parser.close()

        log(f"VBA modules: {len(self.vba_modules)}")

        proc_pattern = re.compile(
            r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
            r"(?:Static\s+)?"
            r"(Sub|Function)\s+"
            r"([^\s(]+)",
            re.I,
        )

        for module, code in self.vba_modules:
            lines = code.splitlines()

            current_name = None
            current_lines = []
            current_start = 0

            def store():
                if current_name:
                    self.vba_procedures[
                        current_name.casefold()
                    ] = (
                        module,
                        current_start,
                        "\n".join(current_lines),
                    )

            for line_no, line in enumerate(lines, 1):
                match = proc_pattern.match(line)

                if match:
                    store()

                    current_name = match.group(2)
                    current_start = line_no
                    current_lines = [line]
                    continue

                if current_name:
                    current_lines.append(line)

                    if re.match(
                        r"^\s*End\s+(Sub|Function)\b",
                        line,
                        re.I,
                    ):
                        store()
                        current_name = None
                        current_lines = []

            store()

        log(f"VBA procedures: {len(self.vba_procedures)}")

    # --------------------------------------------------------
    # VBA references and navigation
    # --------------------------------------------------------

    def vba_sheet_targets(self, code):
        targets = set()

        # Worksheets("サマリ")
        # Sheets("集計表")
        pattern = re.compile(
            r"\b(?:Worksheets|Sheets)\s*\(\s*"
            r'"((?:[^"]|"")*)"\s*\)',
            re.I,
        )

        for match in pattern.finditer(code):
            target = find_sheet(
                match.group(1).replace('""', '"'),
                self.lookup,
            )

            if target:
                targets.add(target)

        # Sheet CodeName references, e.g. Sheet3.Activate
        for code_name, target in self.code_to_sheet.items():
            if re.search(
                r"(?<![\w])"
                + re.escape(code_name)
                + r"\s*\.",
                code,
                re.I,
            ):
                targets.add(target)

        return targets

    def scan_vba_references(self):
        log("Analyzing VBA references...")

        for module, code in self.vba_modules:
            module_sheet = self.code_to_sheet.get(
                module.casefold()
            )

            # Also recognize Sheet module names such as:
            # Sheet1, Sheet2
            if not module_sheet:
                module_sheet = self.code_to_sheet.get(
                    module.split(".")[0].casefold()
                )

            if not module_sheet:
                continue

            for target in self.vba_sheet_targets(code):
                self.add_edge(
                    module_sheet, target,
                    "VBA_Module",
                    module,
                    code[:MAX_EXPRESSION],
                )

        # VBA routines assigned to shapes are processed
        # separately to establish the source sheet.
        for source, shape_name, macro in self.macro_bindings:
            macro_name = macro.split("!")[-1]
            macro_name = macro_name.strip("'\"")

            key = macro_name.casefold()
            visited = set()

            targets = self.resolve_macro_calls(
                key, visited, 0
            )

            if targets:
                for target in targets:
                    self.add_edge(
                        source, target,
                        "VBA_Button",
                        shape_name,
                        macro,
                    )
            else:
                self.add_unresolved(
                    source,
                    shape_name,
                    "VBA_Button",
                    macro,
                    "Macro not found or target unresolved",
                )

    def resolve_macro_calls(self, key, visited, depth):
        if depth > 8 or key in visited:
            return set()

        visited.add(key)

        procedure = self.vba_procedures.get(key)

        if not procedure:
            return set()

        _, _, code = procedure

        targets = self.vba_sheet_targets(code)

        # Follow direct calls to other procedures.
        # This is intentionally conservative.
        for line in code.splitlines():
            line = strip_vba_comments(line).strip()

            if not line:
                continue

            for other in self.vba_procedures:
                if other == key:
                    continue

                if re.match(
                    r"^(?:Call\s+)?"
                    + re.escape(other)
                    + r"(?:\s*\(|\s*$)",
                    line,
                    re.I,
                ):
                    targets.update(
                        self.resolve_macro_calls(
                            other, visited, depth + 1
                        )
                    )

        return targets

    # --------------------------------------------------------
    # Shape and VML metadata
    # --------------------------------------------------------

    def scan_shapes(self):
        log("Scanning drawing and VML metadata...")

        # Resolve drawing/VML part -> owner sheet through
        # workbook relationships and sheet relationships.
        with zipfile.ZipFile(self.source) as archive:
            names = set(archive.namelist())

            wb_xml = ET.fromstring(
                archive.read("xl/workbook.xml")
            )

            ns_main = (
                "http://schemas.openxmlformats.org/"
                "spreadsheetml/2006/main"
            )
            ns_rel = (
                "http://schemas.openxmlformats.org/"
                "officeDocument/2006/relationships"
            )

            rel_path = "xl/_rels/workbook.xml.rels"
            rel_map = {}

            if rel_path in names:
                root = ET.fromstring(
                    archive.read(rel_path)
                )

                for element in root:
                    rid = element.attrib.get("Id")
                    target = element.attrib.get("Target", "")

                    if rid:
                        rel_map[rid] = target

            import posixpath

            def resolve_path(base, target):
                target = target.lstrip("/")

                if target.startswith("xl/"):
                    return posixpath.normpath(target)

                return posixpath.normpath(
                    posixpath.join(
                        posixpath.dirname(base),
                        target,
                    )
                )

            for sheet_element in wb_xml.findall(
                f".//{{{ns_main}}}sheet"
            ):
                source = sheet_element.attrib.get("name")
                rid = sheet_element.attrib.get(
                    f"{{{ns_rel}}}id"
                )

                if not source or not rid:
                    continue

                sheet_path = resolve_path(
                    "xl/workbook.xml",
                    rel_map.get(rid, ""),
                )

                if sheet_path not in names:
                    continue

                sheet_rels_path = (
                    posixpath.dirname(sheet_path)
                    + "/_rels/"
                    + posixpath.basename(sheet_path)
                    + ".rels"
                )

                if sheet_rels_path not in names:
                    continue

                root = ET.fromstring(
                    archive.read(sheet_rels_path)
                )

                for rel in root:
                    target = rel.attrib.get("Target", "")
                    part = resolve_path(
                        sheet_path, target
                    )

                    if part not in names:
                        continue

                    if not (
                        "/drawings/" in part
                        or "/vmlDrawings/" in part
                    ):
                        continue

                    raw = archive.read(part)
                    content = raw.decode(
                        "utf-8", errors="replace"
                    )

                    # Drawing hyperlinks, macro attributes,
                    # and VML FmlaMacro metadata.
                    macro_values = set()

                    try:
                        tree = ET.fromstring(raw)

                        for node in tree.iter():
                            for attr, value in node.attrib.items():
                                local = attr.split("}")[-1].lower()

                                if local in (
                                    "macro",
                                    "fmlamacro",
                                ):
                                    macro_values.add(value)

                            if node.tag.split("}")[-1].lower() == "fmlamacro":
                                if node.text:
                                    macro_values.add(node.text)

                    except ET.ParseError:
                        pass

                    # Fallback for VML with unusual XML.
                    for match in re.finditer(
                        r"<[^>]*FmlaMacro[^>]*>"
                        r"(.*?)"
                        r"</[^>]*FmlaMacro>",
                        content,
                        re.I | re.S,
                    ):
                        macro_values.add(match.group(1))

                    for macro in macro_values:
                        macro = macro.strip()

                        if macro.startswith("="):
                            macro = macro[1:]

                        macro = macro.strip('"')

                        if macro:
                            self.macro_bindings.append((
                                source, part, macro
                            ))

                    # Static sheet references in shape XML
                    for target in self.explicit_targets(content):
                        self.add_edge(
                            source, target,
                            "ShapeReference",
                            part,
                            content[:MAX_EXPRESSION],
                            "CANDIDATE",
                        )

        log(
            "Shape macro bindings: "
            f"{len(self.macro_bindings)}"
        )

    # --------------------------------------------------------
    # Report writing
    # --------------------------------------------------------

    def write_sheet(self, wb, title, headers, rows):
        log(f"Writing report sheet: {title}")

        ws = wb.create_sheet(title)
        ws.append(headers)

        count = 0

        for row in rows:
            values = []

            for value in row:
                if value is None:
                    value = ""

                if isinstance(value, str):
                    value = value[:32767]

                values.append(value)

            ws.append(values)
            count += 1

            if count % 10000 == 0:
                log(f"  {title}: {count} rows")

        log(f"Completed {title}: {count} rows")

        return ws

    def create_report(self):
        log("Generating report...")

        output = self.source.with_name(
            self.source.stem
            + "_dependency_report.xlsx"
        )

        # Stream output to reduce memory usage.
        out = Workbook(write_only=True)

        matrix = []

        for source in self.names:
            matrix.append(
                [source]
                + [
                    int(target in self.edges[source])
                    for target in self.names
                ]
            )

        self.write_sheet(
            out,
            "Direct_Matrix",
            ["Source / Target"] + self.names,
            matrix,
        )

        edge_rows = [
            (source, target)
            for source in self.names
            for target in self.names
            if target in self.edges[source]
        ]

        self.write_sheet(
            out,
            "Direct_Links",
            ["Source Sheet", "Target Sheet"],
            edge_rows,
        )

        self.write_sheet(
            out,
            "Evidence",
            [
                "Source Sheet",
                "Target Sheet",
                "Category",
                "Confidence",
                "Location",
                "Expression",
            ],
            self.details,
        )

        self.write_sheet(
            out,
            "Unresolved",
            [
                "Source Sheet",
                "Location",
                "Category",
                "Expression",
                "Reason",
            ],
            self.unresolved,
        )

        self.write_sheet(
            out,
            "Sheet_List",
            ["Sheet", "Visibility", "Outgoing Links"],
            [
                (
                    ws.title,
                    ws.sheet_state,
                    len(self.edges[ws.title]),
                )
                for ws in self.book.worksheets
            ],
        )

        self.write_sheet(
            out,
            "Macro_Bindings",
            ["Source Sheet", "Shape / Part", "Macro"],
            self.macro_bindings,
        )

        summary = [
            ("Source File", self.source.name),
            ("Total Sheets", len(self.names)),
            (
                "Confirmed Direct Links",
                len(edge_rows),
            ),
            (
                "Formula Cells",
                self.counts["formula_cells"],
            ),
            (
                "INDIRECT Calls",
                self.counts["indirect_calls"],
            ),
            (
                "OFFSET Calls",
                self.counts["offset_calls"],
            ),
            (
                "VBA Modules",
                len(self.vba_modules),
            ),
            (
                "VBA Procedures",
                len(self.vba_procedures),
            ),
            (
                "Macro Bindings",
                len(self.macro_bindings),
            ),
            (
                "Evidence Rows",
                len(self.details),
            ),
            (
                "Unresolved Rows",
                len(self.unresolved),
            ),
            (
                "Truncated Evidence",
                self.truncated_details,
            ),
            (
                "Truncated Unresolved",
                self.truncated_unresolved,
            ),
        ]

        summary.extend(
            (key, value)
            for key, value in sorted(self.counts.items())
        )

        self.write_sheet(
            out,
            "Summary",
            ["Metric", "Value"],
            summary,
        )

        log("Saving XLSX report...")

        out.save(output)

        log("XLSX report saved successfully.")
        log(f"Report: {output}")

        return output

    # --------------------------------------------------------
    # Main
    # --------------------------------------------------------

    def run(self):
        self.load()
        self.scan_defined_names()
        self.scan_formulas()
        self.scan_hyperlinks()
        self.scan_vba()
        self.scan_shapes()
        self.scan_vba_references()

        output = self.create_report()

        log("")
        log("========== ANALYSIS RESULT ==========")
        log(f"Sheets: {len(self.names)}")
        log(
            "Formula cells: "
            f"{self.counts['formula_cells']}"
        )
        log(
            "INDIRECT calls: "
            f"{self.counts['indirect_calls']}"
        )
        log(
            "OFFSET calls: "
            f"{self.counts['offset_calls']}"
        )
        log(f"VBA modules: {len(self.vba_modules)}")
        log(
            "Macro bindings: "
            f"{len(self.macro_bindings)}"
        )
        log(
            "Confirmed direct links: "
            f"{sum(len(v) for v in self.edges.values())}"
        )
        log(f"Evidence rows: {len(self.details)}")
        log(f"Unresolved rows: {len(self.unresolved)}")
        log(f"Report: {output}")
        log("=====================================")


def main():
    parser = argparse.ArgumentParser(
        description="Analyze XLSM sheet dependencies"
    )

    parser.add_argument(
        "input",
        help="Path to XLSM workbook",
    )

    args = parser.parse_args()

    try:
        analyzer = XLSMAnalyzer(args.input)
        analyzer.run()

    except KeyboardInterrupt:
        log("Interrupted by user.")
        sys.exit(130)

    except Exception as exc:
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
