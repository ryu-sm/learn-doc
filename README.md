
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Direct Sheet Relationship Analyzer

Features:
- Visible worksheets only
- Excel formula references
- Japanese / Chinese sheet names
- OFFSET explicit references
- INDIRECT literal references
- VBA sheet references
- VBA CodeName references
- VBA basic assignments
- VBA Copy Destination
- Shape macro bindings
- VML form controls
- VBA Activate / Select / Goto
- Internal hyperlinks
- Direct dependency matrices

Output:
1. Direct_Matrix
2. Data_Matrix
3. Navigation_Matrix
4. Dependency_Details
5. Unresolved_References

Direction:
Row = Source Sheet
Column = Target Sheet

Hidden and veryHidden sheets are excluded.

Static analysis only.
VBA macros are never executed.
"""

import argparse
import posixpath
import re
import zipfile

from collections import defaultdict
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from oletools.olevba import VBA_Parser


# ============================================================
# XML Namespaces
# ============================================================

PKG = (
    "http://schemas.openxmlformats.org/"
    "package/2006/relationships"
)

REL = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)

RID = f"{{{REL}}}id"


# ============================================================
# Regular Expressions
# ============================================================

FORMULA_REF = re.compile(
    r"(?<![\w.])"
    r"(?:"
    r"'((?:[^']|'')+)'"
    r"|"
    r"([^\s=+\-*/^&(),;:{}\[\]!'\"<>]+)"
    r")!",
    re.UNICODE,
)

SHEET_OBJ = re.compile(
    r'(?:(?:ThisWorkbook|ActiveWorkbook)'
    r'\s*\.\s*)?'
    r'(?:Worksheets|Sheets)\s*'
    r'\(\s*"((?:[^"]|"")+)"\s*\)',
    re.I,
)

SHEET_INDEX = re.compile(
    r'\b(Worksheets|Sheets)\s*\(\s*(\d+)\s*\)',
    re.I,
)

SET_RE = re.compile(
    r'^Set\s+(\w+)\s*=\s*(.+)$',
    re.I,
)

PROC_RE = re.compile(
    r'^(?:(?:Public|Private|Friend|Static)\s+)*'
    r'(?:Sub|Function)\s+([^\s(]+)',
    re.I,
)

END_PROC = re.compile(
    r'^End\s+(?:Sub|Function)\b',
    re.I,
)

NAV_RE = re.compile(
    r'\.\s*(?:Activate|Select)\b',
    re.I,
)

GOTO_RE = re.compile(
    r'\b(?:Application\s*\.\s*)?Goto\b',
    re.I,
)

COPY_RE = re.compile(
    r'\.\s*Copy\b',
    re.I,
)

DEST_RE = re.compile(
    r'\bDestination\s*:=\s*(.+)',
    re.I,
)

CALL_RE = re.compile(
    r'^(?:Call\s+)?([^\s(]+)',
    re.I,
)


# ============================================================
# Utility Functions
# ============================================================

def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def safe_text(cell, value):
    """
    Force Excel string type.

    Prevent expressions beginning with '='
    from being evaluated as formulas.
    """
    cell.value = str(value)[:32000]
    cell.data_type = "s"
    cell.number_format = "@"


def clean_macro_name(name):
    """
    Example:
        input.xlsm!Module1.Button_Click
        -> button_click
    """
    if not name:
        return ""

    name = str(name).strip().strip("'")

    if "!" in name:
        name = name.rsplit("!", 1)[-1]

    if "." in name:
        name = name.rsplit(".", 1)[-1]

    return name.strip().casefold()


def strip_vba_comment(line):
    quoted = False
    i = 0

    while i < len(line):
        if line[i] == '"':
            if (
                quoted
                and i + 1 < len(line)
                and line[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted

        if line[i] == "'" and not quoted:
            return line[:i]

        i += 1

    return line


def logical_lines(code):
    pending = ""
    start = 0

    for number, raw in enumerate(
        code.splitlines(), 1
    ):
        line = strip_vba_comment(raw).strip()

        if not line:
            continue

        if not pending:
            start = number

        if re.search(r'\s+_$', line):
            pending += re.sub(
                r'\s+_$', ' ', line
            )
        else:
            yield start, pending + line
            pending = ""

    if pending:
        yield start, pending


def resolve_part(base, target):
    target = unquote(target or "")

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(
            posixpath.dirname(base),
            target,
        )
    )


def read_rels(zf, part):
    folder = posixpath.dirname(part)
    filename = posixpath.basename(part)

    rel_path = posixpath.join(
        folder,
        "_rels",
        filename + ".rels",
    )

    if rel_path not in zf.namelist():
        return {}

    root = ET.fromstring(zf.read(rel_path))
    result = {}

    for elem in root:
        if local_name(elem.tag) != "Relationship":
            continue

        result[elem.get("Id")] = {
            "target": elem.get("Target", ""),
            "type": elem.get("Type", ""),
            "mode": elem.get("TargetMode", ""),
            "location": elem.get("Location", ""),
        }

    return result


def split_assignment(line):
    if re.match(
        r'^(If|ElseIf|For|Do|While|Select|'
        r'Case|Dim|Const|Call|Exit)\b',
        line,
        re.I,
    ):
        return None

    quoted = False
    i = 0

    while i < len(line):
        ch = line[i]

        if ch == '"':
            if (
                quoted
                and i + 1 < len(line)
                and line[i + 1] == '"'
            ):
                i += 2
                continue

            quoted = not quoted

        if ch == "=" and not quoted:
            if i and line[i - 1] in "<>:":
                i += 1
                continue

            return (
                line[:i].strip(),
                line[i + 1:].strip(),
            )

        i += 1

    return None


# ============================================================
# Analyzer
# ============================================================

class Analyzer:

    def __init__(self, path):
        self.path = Path(path)

        self.book = load_workbook(
            self.path,
            data_only=False,
            read_only=False,
            keep_links=True,
        )

        # ----------------------------------------------------
        # IMPORTANT: Visible Sheets Only
        # ----------------------------------------------------

        # Original worksheet order
        # Includes hidden and veryHidden sheets
        self.all_sheet_names = [
            ws.title
            for ws in self.book.worksheets
        ]

        # Visible worksheet order
        self.names = [
            ws.title
            for ws in self.book.worksheets
            if ws.sheet_state == "visible"
        ]

        self.lookup = {
            name.casefold(): name
            for name in self.names
        }

        # VBA CodeName mapping
        # Only visible sheets
        self.code_names = {}

        for ws in self.book.worksheets:
            if ws.sheet_state != "visible":
                continue

            code = ws.sheet_properties.codeName

            if code:
                self.code_names[
                    code.casefold()
                ] = ws.title

        print(
            f"Visible Sheets: {len(self.names)} / "
            f"Total Sheets: {len(self.all_sheet_names)}"
        )

        self.data_graph = {
            name: set()
            for name in self.names
        }

        self.nav_graph = {
            name: set()
            for name in self.names
        }

        self.details = set()
        self.unknown = set()

        self.procedures = defaultdict(list)
        self.macro_navigation = defaultdict(set)

        self.stats = defaultdict(int)

    # --------------------------------------------------------
    # Common Methods
    # --------------------------------------------------------

    def normalize(self, name):
        if not isinstance(name, str):
            return None

        return self.lookup.get(
            name.casefold()
        )

    def add(
        self,
        source,
        target,
        relation,
        kind,
        module,
        location,
        expression,
    ):
        source = self.normalize(source)
        target = self.normalize(target)

        if not source or not target:
            return

        if source == target:
            return

        graph = (
            self.nav_graph
            if kind == "navigation"
            else self.data_graph
        )

        graph[source].add(target)

        self.details.add((
            source,
            target,
            kind,
            relation,
            str(module),
            str(location),
            str(expression),
        ))

    def unresolved(
        self,
        kind,
        module,
        location,
        reason,
        expression,
    ):
        self.unknown.add((
            str(kind),
            str(module),
            str(location),
            str(reason),
            str(expression),
        ))

    # --------------------------------------------------------
    # Formula References
    # --------------------------------------------------------

    def formula_refs(self, formula):
        refs = set()

        clean = re.sub(
            r'"(?:[^"]|"")*"',
            '""',
            formula,
        )

        for match in FORMULA_REF.finditer(clean):
            raw = (
                match.group(1)
                or match.group(2)
            ).replace("''", "'")

            if ":" in raw:
                parts = raw.split(":")

                if len(parts) == 2:
                    a = self.normalize(parts[0])
                    b = self.normalize(parts[1])

                    if a and b:
                        i = self.all_sheet_names.index(a)
                        j = self.all_sheet_names.index(b)

                        for name in self.all_sheet_names[
                            min(i, j):max(i, j) + 1
                        ]:
                            visible = self.normalize(name)

                            if visible:
                                refs.add(visible)

                continue

            target = self.normalize(raw)

            if target:
                refs.add(target)

        return refs

    # --------------------------------------------------------
    # Excel Formula Analysis
    # --------------------------------------------------------

    def analyze_formulas(self):
        print("[1/5] Excel formulas")

        for source in self.names:
            ws = self.book[source]

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["formula_cells"] += 1

                    formula = str(cell.value)
                    refs = self.formula_refs(formula)

                    if refs - {source}:
                        self.stats[
                            "cross_sheet_formulas"
                        ] += 1

                    for target in refs:
                        self.add(
                            source,
                            target,
                            "Excel Formula",
                            "data",
                            source,
                            cell.coordinate,
                            formula,
                        )

                    if re.search(
                        r'\bOFFSET\s*\(',
                        formula,
                        re.I,
                    ):
                        self.stats[
                            "offset_formulas"
                        ] += 1

                    if re.search(
                        r'\bINDIRECT\s*\(',
                        formula,
                        re.I,
                    ):
                        self.stats[
                            "indirect_formulas"
                        ] += 1

                        self.unresolved(
                            "Formula",
                            source,
                            cell.coordinate,
                            "INDIRECT dynamic target "
                            "requires review",
                            formula,
                        )

                    match = re.search(
                        r'HYPERLINK\s*\(\s*"([^"]+)"',
                        formula,
                        re.I,
                    )

                    if match:
                        target = self.parse_link(
                            match.group(1)
                        )

                        if target:
                            self.add(
                                source,
                                target,
                                "HYPERLINK Formula",
                                "navigation",
                                source,
                                cell.coordinate,
                                formula,
                            )

    # --------------------------------------------------------
    # VBA Sheet References
    # --------------------------------------------------------

    def vba_refs(self, expression, aliases=None):
        aliases = aliases or {}
        result = set()

        # Worksheets("Name")
        for match in SHEET_OBJ.finditer(expression):
            target = self.normalize(
                match.group(1).replace('""', '"')
            )

            if target:
                result.add(target)

        # Sheets(5) / Worksheets(5)
        for match in SHEET_INDEX.finditer(expression):
            index = int(match.group(2))

            if 1 <= index <= len(self.all_sheet_names):
                original_name = self.all_sheet_names[
                    index - 1
                ]

                visible = self.normalize(original_name)

                if visible:
                    result.add(visible)

        # VBA CodeName
        for code, target in self.code_names.items():
            if re.search(
                r'(?<![\w.])'
                + re.escape(code)
                + r'\s*[.!]',
                expression,
                re.I,
            ):
                result.add(target)

        # Worksheet aliases
        for alias, target in aliases.items():
            if re.search(
                r'(?<![\w.])'
                + re.escape(alias)
                + r'\s*[.!]',
                expression,
                re.I,
            ):
                result.add(target)

        return result

    def resolve_object(self, expression, aliases):
        expression = expression.strip()

        match = SHEET_OBJ.fullmatch(expression)

        if match:
            return self.normalize(
                match.group(1).replace('""', '"')
            )

        match = SHEET_INDEX.fullmatch(expression)

        if match:
            index = int(match.group(2))

            if 1 <= index <= len(self.all_sheet_names):
                original_name = self.all_sheet_names[
                    index - 1
                ]

                return self.normalize(original_name)

        return (
            aliases.get(expression.casefold())
            or self.code_names.get(
                expression.casefold()
            )
        )

    # --------------------------------------------------------
    # VBA Procedure Extraction
    # --------------------------------------------------------

    def extract_vba(self):
        print("[2/5] Extracting VBA")

        parser = VBA_Parser(str(self.path))

        try:
            if not parser.detect_vba_macros():
                return

            for _, _, module, code in parser.extract_macros():
                self.stats["vba_modules"] += 1

                current = None

                for line_no, line in logical_lines(code):
                    self.stats["vba_lines"] += 1

                    match = PROC_RE.match(line)

                    if match:
                        current = clean_macro_name(
                            match.group(1)
                        )
                        continue

                    if END_PROC.match(line):
                        current = None
                        continue

                    if current:
                        self.procedures[current].append(
                            (module, line_no, line)
                        )

        finally:
            parser.close()

    # --------------------------------------------------------
    # VBA Navigation
    # --------------------------------------------------------

    def analyze_macro_navigation(self):
        print("[3/5] VBA navigation")

        direct = defaultdict(set)
        calls = defaultdict(set)

        for proc, lines in self.procedures.items():
            aliases = {}

            for module, line_no, line in lines:
                match = SET_RE.match(line)

                if match:
                    variable, expr = match.groups()

                    target = self.resolve_object(
                        expr, aliases
                    )

                    if target:
                        aliases[
                            variable.casefold()
                        ] = target

                    continue

                if NAV_RE.search(line):
                    expression = NAV_RE.split(line)[0]

                    refs = self.vba_refs(
                        expression, aliases
                    )

                    if not refs:
                        self.unresolved(
                            "VBA Navigation",
                            module,
                            line_no,
                            "Activate/Select target unresolved",
                            line,
                        )

                    direct[proc].update(refs)

                if GOTO_RE.search(line):
                    refs = self.vba_refs(
                        line, aliases
                    )
                    direct[proc].update(refs)

                match = CALL_RE.match(line)

                if match:
                    called = clean_macro_name(
                        match.group(1)
                    )

                    if (
                        called != proc
                        and called in self.procedures
                    ):
                        calls[proc].add(called)

        resolved = {
            p: set(direct[p])
            for p in self.procedures
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

        self.macro_navigation = resolved

        print(
            "VBA procedures:",
            len(self.procedures)
        )

        print(
            "Navigation procedures:",
            sum(bool(v) for v in resolved.values())
        )

    # --------------------------------------------------------
    # VBA Data Analysis
    # --------------------------------------------------------

    def analyze_vba_data(self):
        for proc, lines in self.procedures.items():
            aliases = {}

            for module, line_no, line in lines:
                match = SET_RE.match(line)

                if match:
                    variable, expr = match.groups()

                    target = self.resolve_object(
                        expr, aliases
                    )

                    if target:
                        aliases[
                            variable.casefold()
                        ] = target

                    continue

                if NAV_RE.search(line):
                    continue

                if COPY_RE.search(line):
                    copy = COPY_RE.search(line)

                    sources = self.vba_refs(
                        line[:copy.start()],
                        aliases,
                    )

                    dest = DEST_RE.search(
                        line[copy.end():]
                    )

                    targets = (
                        self.vba_refs(
                            dest.group(1),
                            aliases,
                        )
                        if dest
                        else set()
                    )

                    if sources and targets:
                        for target in targets:
                            for source in sources:
                                self.add(
                                    target,
                                    source,
                                    "VBA Copy Destination",
                                    "data",
                                    module,
                                    line_no,
                                    line,
                                )
                    else:
                        self.unresolved(
                            "VBA",
                            module,
                            line_no,
                            "Copy target/source unresolved",
                            line,
                        )

                    continue

                assignment = split_assignment(line)

                if not assignment:
                    continue

                left, right = assignment

                targets = self.vba_refs(
                    left, aliases
                )

                sources = self.vba_refs(
                    right, aliases
                )

                for target in targets:
                    for source in sources:
                        self.add(
                            target,
                            source,
                            "VBA Assignment",
                            "data",
                            module,
                            line_no,
                            line,
                        )

    # --------------------------------------------------------
    # Hyperlink Target
    # --------------------------------------------------------

    def parse_link(self, target):
        target = unquote(
            str(target or "")
        ).strip()

        if target.startswith("#"):
            target = target[1:]

        if "!" not in target:
            return None

        name = target.split(
            "!", 1
        )[0].strip()

        if (
            name.startswith("'")
            and name.endswith("'")
        ):
            name = name[1:-1].replace(
                "''", "'"
            )

        return self.normalize(name)

    # --------------------------------------------------------
    # Worksheet XML Mapping
    # --------------------------------------------------------

    def worksheet_parts(self, zf):
        workbook_part = "xl/workbook.xml"

        root = ET.fromstring(
            zf.read(workbook_part)
        )

        rels = read_rels(
            zf, workbook_part
        )

        result = {}

        for elem in root.iter():
            if local_name(elem.tag) != "sheet":
                continue

            name = elem.get("name")
            rid = elem.get(RID)

            rel = rels.get(rid)

            if not rel:
                continue

            result[name] = resolve_part(
                workbook_part,
                rel["target"],
            )

        return result

    # --------------------------------------------------------
    # Shape Macro Binding
    # --------------------------------------------------------

    def record_shape_macro(
        self,
        source,
        macro,
        part,
        shape,
    ):
        self.stats[
            "shape_macro_bindings"
        ] += 1

        proc = clean_macro_name(macro)

        targets = self.macro_navigation.get(proc)

        if not targets:
            self.unresolved(
                "Shape Macro",
                source,
                shape,
                "Macro navigation target unresolved",
                f"{macro} @ {part}",
            )
            return

        for target in targets:
            self.add(
                source,
                target,
                "Shape VBA Navigation",
                "navigation",
                part,
                shape,
                f"Macro: {macro}",
            )

    # --------------------------------------------------------
    # DrawingML Shapes
    # --------------------------------------------------------

    def analyze_drawing(
        self,
        zf,
        source,
        drawing_part,
    ):
        if drawing_part not in zf.namelist():
            return

        root = ET.fromstring(
            zf.read(drawing_part)
        )

        rels = read_rels(
            zf, drawing_part
        )

        for elem in root.iter():
            tag = local_name(elem.tag)

            if tag in (
                "sp",
                "graphicFrame",
                "pic",
            ):
                macro = elem.get("macro")

                if macro:
                    shape = elem.get(
                        "name", tag
                    )

                    for child in elem.iter():
                        if local_name(
                            child.tag
                        ) == "cNvPr":
                            shape = child.get(
                                "name", shape
                            )
                            break

                    self.record_shape_macro(
                        source,
                        macro,
                        drawing_part,
                        shape,
                    )

            if tag == "hlinkClick":
                rid = elem.get(RID)

                rel = rels.get(
                    rid, {}
                )

                target_text = rel.get(
                    "target", ""
                )

                location = rel.get(
                    "location", ""
                )

                target = self.parse_link(
                    location or target_text
                )

                if target:
                    self.add(
                        source,
                        target,
                        "Shape Hyperlink",
                        "navigation",
                        drawing_part,
                        rid or "-",
                        location or target_text,
                    )

    # --------------------------------------------------------
    # VML Controls
    # --------------------------------------------------------

    def analyze_vml(
        self,
        zf,
        source,
        vml_part,
    ):
        if vml_part not in zf.namelist():
            return

        root = ET.fromstring(
            zf.read(vml_part)
        )

        for elem in root.iter():
            if local_name(
                elem.tag
            ) != "ClientData":
                continue

            macro = None

            for child in elem:
                if local_name(
                    child.tag
                ) == "FmlaMacro":
                    macro = (
                        child.text or ""
                    ).strip()
                    break

            if macro:
                self.record_shape_macro(
                    source,
                    macro,
                    vml_part,
                    "VML Form Control",
                )

    # --------------------------------------------------------
    # Cell Hyperlinks
    # --------------------------------------------------------

    def analyze_cell_links(
        self,
        zf,
        source,
        sheet_part,
        rels,
    ):
        root = ET.fromstring(
            zf.read(sheet_part)
        )

        for elem in root.iter():
            if local_name(
                elem.tag
            ) != "hyperlink":
                continue

            location = elem.get(
                "location", ""
            )

            rid = elem.get(RID)
            cell = elem.get(
                "ref", ""
            )

            if not location and rid:
                rel = rels.get(
                    rid, {}
                )

                location = rel.get(
                    "target", ""
                )

            target = self.parse_link(
                location
            )

            if target:
                self.add(
                    source,
                    target,
                    "Cell Hyperlink",
                    "navigation",
                    sheet_part,
                    cell,
                    location,
                )

    # --------------------------------------------------------
    # Shape Analysis
    # --------------------------------------------------------

    def analyze_shapes(self):
        print(
            "[4/5] Shape and hyperlink navigation"
        )

        with zipfile.ZipFile(
            self.path
        ) as zf:

            sheet_parts = self.worksheet_parts(zf)

            for source, sheet_part in sheet_parts.items():

                # IMPORTANT:
                # Skip hidden and veryHidden sheets
                if source not in self.names:
                    continue

                if sheet_part not in zf.namelist():
                    continue

                rels = read_rels(
                    zf,
                    sheet_part,
                )

                self.analyze_cell_links(
                    zf,
                    source,
                    sheet_part,
                    rels,
                )

                for rel in rels.values():
                    relation_type = rel["type"]

                    if rel["mode"] == "External":
                        continue

                    target_part = resolve_part(
                        sheet_part,
                        rel["target"],
                    )

                    if relation_type.endswith(
                        "/drawing"
                    ):
                        self.analyze_drawing(
                            zf,
                            source,
                            target_part,
                        )

                    elif relation_type.endswith(
                        "/vmlDrawing"
                    ):
                        self.analyze_vml(
                            zf,
                            source,
                            target_part,
                        )

    # --------------------------------------------------------
    # Matrix Creation
    # --------------------------------------------------------

    def create_matrix(
        self,
        wb,
        title,
        graph,
    ):
        ws = wb.create_sheet(title)

        safe_text(
            ws.cell(1, 1),
            "Source / Target",
        )

        for i, name in enumerate(
            self.names, 2
        ):
            safe_text(
                ws.cell(1, i),
                name,
            )

            safe_text(
                ws.cell(i, 1),
                name,
            )

        for i, source in enumerate(
            self.names, 2
        ):
            for j, target in enumerate(
                self.names, 2
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
                        fgColor="C6EFCE",
                    )

                    cell.font = Font(
                        bold=True,
                        color="006100",
                    )

        self.style_header(ws)

        ws.freeze_panes = "B2"

        ws.column_dimensions[
            "A"
        ].width = 32

        for col in range(
            2,
            len(self.names) + 2
        ):
            ws.column_dimensions[
                get_column_letter(col)
            ].width = 22

    @staticmethod
    def style_header(ws):
        for cell in ws[1]:
            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78",
            )

            cell.font = Font(
                bold=True,
                color="FFFFFF",
            )

            cell.alignment = Alignment(
                horizontal="center",
            )

    # --------------------------------------------------------
    # Report Generation
    # --------------------------------------------------------

    def report(self, output):
        print(
            "[5/5] Generating report"
        )

        wb = Workbook()
        wb.remove(wb.active)

        combined = {}

        for name in self.names:
            combined[name] = (
                self.data_graph[name]
                | self.nav_graph[name]
            )

        self.create_matrix(
            wb,
            "Direct_Matrix",
            combined,
        )

        self.create_matrix(
            wb,
            "Data_Matrix",
            self.data_graph,
        )

        self.create_matrix(
            wb,
            "Navigation_Matrix",
            self.nav_graph,
        )

        definitions = [
            (
                "Dependency_Details",
                [
                    "Source Sheet",
                    "Target Sheet",
                    "Kind",
                    "Relation Type",
                    "Module / XML",
                    "Location / Shape",
                    "Expression",
                ],
                self.details,
                [
                    25, 25, 18,
                    30, 42, 28, 95
                ],
            ),
            (
                "Unresolved_References",
                [
                    "Kind",
                    "Module / Sheet",
                    "Location",
                    "Reason",
                    "Expression",
                ],
                self.unknown,
                [
                    25, 35, 28,
                    48, 95
                ],
            ),
        ]

        for title, headers, records, widths in definitions:
            ws = wb.create_sheet(title)

            for col, header in enumerate(
                headers, 1
            ):
                safe_text(
                    ws.cell(1, col),
                    header,
                )

            for row, record in enumerate(
                sorted(records), 2
            ):
                for col, value in enumerate(
                    record, 1
                ):
                    safe_text(
                        ws.cell(row, col),
                        value,
                    )

            self.style_header(ws)

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col, width in enumerate(
                widths, 1
            ):
                ws.column_dimensions[
                    get_column_letter(col)
                ].width = width

        wb.save(output)

    # --------------------------------------------------------
    # Run
    # --------------------------------------------------------

    def run(self, output):
        try:
            self.analyze_formulas()

            self.extract_vba()
            self.analyze_macro_navigation()
            self.analyze_vba_data()

            self.analyze_shapes()

            self.report(output)

            data_links = sum(
                len(v)
                for v in self.data_graph.values()
            )

            nav_links = sum(
                len(v)
                for v in self.nav_graph.values()
            )

            combined_links = sum(
                len(
                    self.data_graph[name]
                    | self.nav_graph[name]
                )
                for name in self.names
            )

            print()
            print(
                "========== RESULT =========="
            )

            print(
                "Visible Sheets:",
                len(self.names),
            )

            print(
                "Total Sheets:",
                len(self.all_sheet_names),
            )

            print(
                "Formula cells:",
                self.stats["formula_cells"],
            )

            print(
                "Cross-sheet formula cells:",
                self.stats["cross_sheet_formulas"],
            )

            print(
                "OFFSET formulas:",
                self.stats["offset_formulas"],
            )

            print(
                "INDIRECT formulas:",
                self.stats["indirect_formulas"],
            )

            print(
                "VBA modules:",
                self.stats["vba_modules"],
            )

            print(
                "VBA procedures:",
                len(self.procedures),
            )

            print(
                "Shape macro bindings:",
                self.stats["shape_macro_bindings"],
            )

            print(
                "Data links:",
                data_links,
            )

            print(
                "Navigation links:",
                nav_links,
            )

            print(
                "Combined links:",
                combined_links,
            )

            print(
                "Evidence rows:",
                len(self.details),
            )

            print(
                "Unresolved rows:",
                len(self.unknown),
            )

            print(
                "Report:",
                output,
            )

            print(
                "============================"
            )

        finally:
            self.book.close()


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Analyze direct sheet relationships "
            "in XLSM/XLSX files"
        )
    )

    parser.add_argument(
        "input",
        help="Input XLSM/XLSX file",
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output XLSX report",
    )

    args = parser.parse_args()

    source = Path(args.input)

    if not source.is_file():
        parser.error(
            f"Input file not found: {source}"
        )

    if source.suffix.lower() not in (
        ".xlsm", ".xlsx"
    ):
        parser.error(
            "Only XLSM/XLSX files are supported"
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
            "Output must be XLSX"
        )

    if output.resolve() == source.resolve():
        parser.error(
            "Output must differ from input"
        )

    Analyzer(source).run(output)


if __name__ == "__main__":
    main()
