
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
XLSM Dependency Analyzer V3

Python 3.10+

Features:
- All sheets, including hidden and veryHidden
- Direct formula references
- INDIRECT dynamic references
- Cached cell value resolution
- Excel shapes and drawing hyperlinks
- VML buttons and macro bindings
- VBA reference extraction
- Separate data and navigation evidence
- Unresolved formula diagnostics
- Original XLSM remains unchanged
"""

import argparse
import collections
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


NS = {
    "m": "http://schemas.openxmlformats.org/"
         "spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/"
         "officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/"
         "package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/"
           "drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/"
         "drawingml/2006/main",
}

CELL = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$", re.I)

SHEET_TOKEN = re.compile(
    r"(?:'((?:[^']|'')+)'|"
    r"([A-Za-z_\u0080-\uffff][\w.\u0080-\uffff]*))!"
    r"(?=\$?[A-Z]{1,3}\$?\d+)",
    re.I,
)


def text(value):
    return "" if value is None else str(value)


def local(tag):
    return tag.rsplit("}", 1)[-1]


def resolve(base, target):
    if not target:
        return ""

    if target.startswith("/"):
        return target.lstrip("/")

    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base), target)
    )


def relationships(z, part):
    rp = posixpath.join(
        posixpath.dirname(part),
        "_rels",
        posixpath.basename(part) + ".rels"
    )

    if rp not in z.namelist():
        return {}

    root = ET.fromstring(z.read(rp))

    return {
        e.get("Id"): (
            e.get("Type", "").rsplit("/", 1)[-1],
            e.get("Target", ""),
            e.get("TargetMode", "")
        )
        for e in root
    }


def split_args(expression):
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
            elif ch in ",;" and depth == 0:
                result.append("".join(buffer).strip())
                buffer = []
                i += 1
                continue

        buffer.append(ch)
        i += 1

    result.append("".join(buffer).strip())
    return result


def calls(formula, function_name):
    pattern = re.compile(
        r"\b" + re.escape(function_name) + r"\s*\(",
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


def split_concat(expression):
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
            elif ch == "&" and depth == 0:
                result.append("".join(buffer).strip())
                buffer = []
                i += 1
                continue

        buffer.append(ch)
        i += 1

    result.append("".join(buffer).strip())
    return result


def literal_expr(expr, ws, cache_ws, depth=0):
    if depth > 12:
        return None

    expr = expr.strip()

    if expr.startswith("(") and expr.endswith(")"):
        return literal_expr(
            expr[1:-1], ws, cache_ws, depth + 1
        )

    parts = split_concat(expr)

    if len(parts) > 1:
        values = [
            literal_expr(p, ws, cache_ws, depth + 1)
            for p in parts
        ]

        if any(v is None for v in values):
            return None

        return "".join(str(v) for v in values)

    if expr.startswith('"') and expr.endswith('"'):
        return expr[1:-1].replace('""', '"')

    if re.fullmatch(r"-?\d+(?:\.\d+)?", expr):
        return expr

    if CELL.fullmatch(expr):
        address = expr.replace("$", "")

        try:
            cell = cache_ws[address]

            if cell.value is None:
                return None

            return str(cell.value)

        except Exception:
            return None

    # SUBSTITUTE
    if re.match(r"^SUBSTITUTE\s*\(", expr, re.I):
        found = list(calls(expr, "SUBSTITUTE"))

        if found:
            args = split_args(found[0])

            if len(args) >= 3:
                values = [
                    literal_expr(a, ws, cache_ws, depth + 1)
                    for a in args[:3]
                ]

                if all(v is not None for v in values):
                    return values[0].replace(
                        values[1], values[2]
                    )

    return None


def name_from_ref(value, names):
    if not value:
        return None

    value = str(value).strip().lstrip("=")

    if value.startswith("#"):
        value = value[1:]

    if "!" not in value:
        return None

    token = value.split("!", 1)[0].strip()

    if token.startswith("'") and token.endswith("'"):
        token = token[1:-1].replace("''", "'")

    if "[" in token and "]" in token:
        # External workbook references must not be
        # treated as internal worksheet references.
        return None

    lookup = {n.casefold(): n for n in names}
    return lookup.get(token.casefold())


def sheet_xml_info(z):
    root = ET.fromstring(z.read("xl/workbook.xml"))
    rels = relationships(z, "xl/workbook.xml")

    result = []

    for element in root.findall("m:sheets/m:sheet", NS):
        rid = element.get("{" + NS["r"] + "}id", "")
        item = rels.get(rid)

        if not item:
            continue

        result.append((
            element.get("name", ""),
            element.get("state", "visible"),
            resolve("xl/workbook.xml", item[1])
        ))

    return result


def drawing_audit(z, infos):
    """
    Extract DrawingML and VML references.

    Shape navigation and macro bindings are kept separate
    from formula data dependencies.
    """
    rows = []
    files = set(z.namelist())

    for sheet, state, part in infos:
        if part not in files:
            continue

        rels = relationships(z, part)
        root = ET.fromstring(z.read(part))

        targets = []

        for element in root.iter():
            kind = local(element.tag)

            if kind not in (
                "drawing",
                "legacyDrawing",
                "legacyDrawingHF",
                "control"
            ):
                continue

            rid = element.get("{" + NS["r"] + "}id", "")

            if rid in rels:
                typ, target, mode = rels[rid]

                targets.append((
                    kind,
                    resolve(part, target),
                    typ,
                    rid
                ))

        for kind, target, typ, rid in targets:
            if target not in files:
                rows.append([
                    sheet, kind, rid, "",
                    "", "", "", target,
                    "MISSING_PART",
                    "Relationship target missing"
                ])
                continue

            if target.endswith(".vml"):
                vml = ET.fromstring(z.read(target))

                for shape in vml.iter():
                    if local(shape.tag) != "shape":
                        continue

                    shape_id = shape.get("id", "")
                    macro = ""
                    link = ""
                    object_type = ""

                    for element in shape.iter():
                        tag = local(element.tag)

                        if tag in ("FmlaMacro", "Macro"):
                            macro = element.text or ""

                        if tag in (
                            "FmlaLink",
                            "FmlaRange",
                            "FmlaTxbx"
                        ):
                            link += " " + text(element.text)

                        if tag == "ClientData":
                            object_type = element.get(
                                "ObjectType", ""
                            )

                    rows.append([
                        sheet,
                        "VML",
                        shape_id,
                        object_type,
                        "",
                        macro,
                        link.strip(),
                        target,
                        "FOUND",
                        ""
                    ])

            elif (
                "/drawings/" in target
                and target.endswith(".xml")
            ):
                drawing = ET.fromstring(z.read(target))
                drawing_rels = relationships(z, target)

                for anchor in drawing:
                    if local(anchor.tag) not in (
                        "twoCellAnchor",
                        "oneCellAnchor",
                        "absoluteAnchor"
                    ):
                        continue

                    for obj in anchor:
                        if local(obj.tag) not in (
                            "sp", "pic", "graphicFrame",
                            "cxnSp", "grpSp"
                        ):
                            continue

                        name = ""
                        shape_id = ""
                        macro = obj.get("macro", "")
                        links = []
                        description = []

                        for element in obj.iter():
                            tag = local(element.tag)

                            if tag == "cNvPr":
                                shape_id = element.get("id", "")
                                name = element.get("name", "")

                                description.append(
                                    element.get("descr", "")
                                )

                            if tag in (
                                "hlinkClick",
                                "hlinkHover"
                            ):
                                link_id = element.get(
                                    "{" + NS["r"] + "}id", ""
                                )

                                if link_id in drawing_rels:
                                    lt, lv, mode = drawing_rels[
                                        link_id
                                    ]

                                    if mode == "External":
                                        links.append(lv)
                                    else:
                                        links.append(
                                            resolve(target, lv)
                                        )

                                if element.get("action"):
                                    links.append(
                                        element.get("action")
                                    )

                            if tag == "f":
                                description.append(
                                    text(element.text)
                                )

                            if element.get("macro"):
                                macro = element.get("macro")

                        rows.append([
                            sheet,
                            local(obj.tag),
                            shape_id,
                            name,
                            "; ".join(links),
                            macro,
                            "; ".join(
                                d for d in description if d
                            ),
                            target,
                            "FOUND",
                            ""
                        ])

            else:
                rows.append([
                    sheet,
                    kind,
                    rid,
                    typ,
                    "",
                    "",
                    "",
                    target,
                    "FOUND",
                    "Other linked object"
                ])

        # Cell hyperlinks
        for link in root.findall(".//m:hyperlink", NS):
            rid = link.get("{" + NS["r"] + "}id", "")
            item = rels.get(rid)

            destination = (
                item[1] if item
                else link.get("location", "")
            )

            rows.append([
                sheet,
                "CellHyperlink",
                link.get("ref", ""),
                link.get("display", ""),
                destination,
                "",
                "",
                part,
                "FOUND",
                ""
            ])

    return rows


def macro_audit(source):
    rows = []
    parser = None

    try:
        from oletools.olevba import VBA_Parser

        parser = VBA_Parser(str(source))

        if parser.detect_vba_macros():
            for _, stream, module, code in parser.extract_macros():
                if isinstance(code, bytes):
                    code = code.decode(
                        "utf-8", errors="replace"
                    )

                lines = code.splitlines()
                procedure = ""

                for number, line in enumerate(lines, 1):
                    match = re.match(
                        r"^\s*(?:Public |Private |Friend )?"
                        r"(?:Sub|Function)\s+"
                        r"([\w\u0080-\uffff]+)",
                        line,
                        re.I
                    )

                    if match:
                        procedure = match.group(1)

                    if re.search(
                        r"\.Sheets\s*\("
                        r"|\.Worksheets\s*\("
                        r"|\.Activate\b"
                        r"|\.Select\b"
                        r"|\.OnAction\b"
                        r"|\.Hyperlinks\b"
                        r"|\.Range\s*\(",
                        line,
                        re.I
                    ):
                        rows.append([
                            module,
                            procedure,
                            number,
                            line.strip()[:1500]
                        ])

    except ImportError:
        rows.append([
            "[oletools unavailable]",
            "",
            "",
            "Install with pip install oletools"
        ])

    except Exception as exc:
        rows.append([
            "[VBA extraction failed]",
            "",
            "",
            str(exc)
        ])

    finally:
        if parser is not None:
            try:
                parser.close()
            except Exception:
                pass

    return rows


def report_sheet(wb, title, headers, rows):
    ws = wb.create_sheet(title)
    ws.append(headers)

    for row in rows:
        safe = []

        for value in row:
            if isinstance(value, str):
                if len(value) > 32000:
                    value = value[:31980] + "...[truncated]"

                # Store expressions as text, not Excel formulas.
                if value.startswith(("=", "+", "-", "@")):
                    value = "'" + value

            safe.append(value)

        ws.append(safe)

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid", fgColor="17365D"
        )
        cell.font = Font(
            color="FFFFFF", bold=True
        )

    for index, header in enumerate(headers, 1):
        width = (
            55 if header in (
                "Formula", "Expression",
                "Reason", "Link", "Macro",
                "Evidence", "Code"
            )
            else 26
        )

        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    return ws


def main():
    parser = argparse.ArgumentParser(
        description="XLSM Dependency Analyzer V3"
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

    if not source.is_file():
        parser.error("Source file not found")

    output = (
        args.output
        or source.with_name(
            source.stem + "_dependency_report_v3.xlsx"
        )
    ).resolve()

    if output == source:
        parser.error("Output cannot overwrite source")

    print("Loading workbook...", flush=True)

    formula_wb = load_workbook(
        source,
        read_only=True,
        data_only=False,
        keep_links=True
    )

    cache_wb = load_workbook(
        source,
        read_only=True,
        data_only=True,
        keep_links=True
    )

    names = set(formula_wb.sheetnames)

    try:
        with zipfile.ZipFile(source) as z:
            infos = sheet_xml_info(z)

            print(
                "Scanning drawings, shapes and hyperlinks...",
                flush=True
            )

            drawing = drawing_audit(z, infos)

        evidence = []
        unresolved = []
        patterns = collections.Counter()

        # Data and navigation links are separate.
        data_links = set()
        navigation_links = set()

        formula_count = 0
        indirect_count = 0
        offset_count = 0

        for ws in formula_wb.worksheets:
            print("Sheet:", ws.title, flush=True)

            cache = cache_wb[ws.title]

            for row in ws.iter_rows():
                for cell in row:
                    if (
                        cell.data_type != "f"
                        or not isinstance(cell.value, str)
                    ):
                        continue

                    formula = cell.value
                    address = cell.coordinate

                    formula_count += 1

                    # Remove quoted string literals before
                    # scanning direct references.
                    masked = re.sub(
                        r'"(?:[^"]|"")*"',
                        '""',
                        formula
                    )

                    for match in SHEET_TOKEN.finditer(masked):
                        target = (
                            match.group(1)
                            or match.group(2)
                        ).replace("''", "'")

                        if "[" in target or "]" in target:
                            continue

                        if target in names:
                            data_links.add(
                                (ws.title, target)
                            )

                            evidence.append([
                                ws.title,
                                address,
                                target,
                                "DIRECT_FORMULA",
                                formula,
                                "Confirmed direct formula reference"
                            ])

                    # INDIRECT
                    indirect_calls = list(
                        calls(formula, "INDIRECT")
                    )

                    indirect_count += len(indirect_calls)

                    for expression in indirect_calls:
                        args = split_args(expression)

                        if not args:
                            continue

                        resolved = literal_expr(
                            args[0], ws, cache
                        )

                        target = name_from_ref(
                            resolved, names
                        ) if resolved else None

                        if target:
                            data_links.add(
                                (ws.title, target)
                            )

                            evidence.append([
                                ws.title,
                                address,
                                target,
                                "INDIRECT",
                                formula,
                                "Resolved dynamic reference"
                            ])

                        else:
                            reason = (
                                "Cannot evaluate INDIRECT "
                                "text expression"
                            )

                            if resolved:
                                reason += (
                                    ": " + str(resolved)[:150]
                                )

                            unresolved.append([
                                ws.title,
                                address,
                                "INDIRECT",
                                args[0],
                                reason,
                                formula
                            ])

                            upper = args[0].upper()

                            if "ADDRESS(" in upper:
                                pattern = "ADDRESS"
                            elif "HLOOKUP(" in upper:
                                pattern = "HLOOKUP"
                            elif "SUBSTITUTE(" in upper:
                                pattern = "SUBSTITUTE"
                            else:
                                pattern = "OTHER"

                            patterns[pattern] += 1

                    # OFFSET
                    offset_calls = list(
                        calls(formula, "OFFSET")
                    )

                    offset_count += len(offset_calls)

                    for expression in offset_calls:
                        unresolved.append([
                            ws.title,
                            address,
                            "OFFSET",
                            expression,
                            "Offset destination not fully evaluated",
                            formula
                        ])

                    # HYPERLINK formula
                    for expression in calls(
                        formula, "HYPERLINK"
                    ):
                        args = split_args(expression)

                        if not args:
                            continue

                        value = literal_expr(
                            args[0], ws, cache
                        )

                        target = name_from_ref(
                            value, names
                        ) if value else None

                        if target:
                            navigation_links.add(
                                (ws.title, target)
                            )

                        evidence.append([
                            ws.title,
                            address,
                            target or value or "",
                            "HYPERLINK_FORMULA",
                            formula,
                            "Navigation, not data dependency"
                        ])

                    if formula_count % 10000 == 0:
                        print(
                            "Processed formulas:",
                            formula_count,
                            flush=True
                        )

        print("Scanning VBA...", flush=True)
        vba = macro_audit(source)

        # Shape navigation and macros
        for item in drawing:
            (
                sheet, kind, shape_id, name,
                link, macro, other,
                part, status, note
            ) = item

            if link:
                target = name_from_ref(
                    link.replace("#", ""), names
                )

                if target:
                    navigation_links.add(
                        (sheet, target)
                    )

                evidence.append([
                    sheet,
                    shape_id,
                    target or link,
                    "SHAPE_HYPERLINK",
                    link,
                    "Shape navigation"
                ])

            if macro:
                evidence.append([
                    sheet,
                    shape_id,
                    macro,
                    "SHAPE_MACRO",
                    macro,
                    "Macro binding; destination not inferred"
                ])

        # VBA references are kept as evidence, not
        # automatically counted as data dependencies.
        for module, procedure, number, code in vba:
            for match in re.finditer(
                r'(?:ThisWorkbook\.)?'
                r'(?:Sheets|Worksheets)'
                r'\s*\(\s*"([^"]+)"\s*\)',
                text(code),
                re.I
            ):
                target = match.group(1)

                if target in names:
                    evidence.append([
                        module,
                        str(number),
                        target,
                        "VBA_SHEET_REFERENCE",
                        code,
                        "VBA reference requiring context"
                    ])

        print("Generating report...", flush=True)

        wb = Workbook()
        wb.remove(wb.active)

        sheet_names = formula_wb.sheetnames

        summary = [
            ["Source File", source.name],
            ["Total Sheets", len(sheet_names)],
            ["Formula Cells", formula_count],
            ["Confirmed Data Links", len(data_links)],
            ["Confirmed Navigation Links", len(navigation_links)],
            ["INDIRECT Calls", indirect_count],
            ["OFFSET Calls", offset_count],
            [
                "Unresolved INDIRECT",
                sum(
                    r[2] == "INDIRECT"
                    for r in unresolved
                )
            ],
            ["Drawing / Link Records", len(drawing)],
            ["VBA Audit Rows", len(vba)],
            ["Evidence Rows", len(evidence)],
            ["Unresolved Rows", len(unresolved)],
        ]

        report_sheet(
            wb, "Summary",
            ["Metric", "Value"],
            summary
        )

        report_sheet(
            wb, "Sheets",
            ["Sheet", "Visibility", "XML Part"],
            infos
        )

        # Data matrix: 1 or 0
        report_sheet(
            wb, "Data_Matrix",
            ["Source Sheet"] + sheet_names,
            [
                [source_sheet] + [
                    1 if (
                        source_sheet, target
                    ) in data_links else 0
                    for target in sheet_names
                ]
                for source_sheet in sheet_names
            ]
        )

        # Navigation matrix
        report_sheet(
            wb, "Navigation_Matrix",
            ["Source Sheet"] + sheet_names,
            [
                [source_sheet] + [
                    1 if (
                        source_sheet, target
                    ) in navigation_links else 0
                    for target in sheet_names
                ]
                for source_sheet in sheet_names
            ]
        )

        report_sheet(
            wb, "Evidence",
            [
                "Source",
                "Cell/Shape",
                "Target",
                "Category",
                "Expression",
                "Reason"
            ],
            evidence
        )

        report_sheet(
            wb, "Unresolved",
            [
                "Sheet",
                "Cell",
                "Category",
                "Expression",
                "Reason",
                "Formula"
            ],
            unresolved
        )

        report_sheet(
            wb, "INDIRECT_Patterns",
            ["Pattern", "Count"],
            patterns.most_common()
        )

        report_sheet(
            wb, "Shapes_Links",
            [
                "Sheet",
                "Object Type",
                "Object ID",
                "Name",
                "Link",
                "Macro",
                "Evidence",
                "XML Part",
                "Status",
                "Note"
            ],
            drawing
        )

        report_sheet(
            wb, "VBA_References",
            [
                "Module",
                "Procedure",
                "Line",
                "Code"
            ],
            vba
        )

        report_sheet(
            wb, "README",
            ["Topic", "Explanation"],
            [
                [
                    "Data Matrix",
                    "1 = confirmed data dependency; "
                    "0 = no confirmed evidence"
                ],
                [
                    "Navigation Matrix",
                    "Confirmed internal navigation links"
                ],
                [
                    "INDIRECT",
                    "Simple cached text expressions are resolved"
                ],
                [
                    "Shapes",
                    "DrawingML and VML metadata are inspected"
                ],
                [
                    "VBA",
                    "VBA source extraction requires oletools"
                ],
                [
                    "Original File",
                    "Source XLSM is never saved"
                ],
                [
                    "Limitations",
                    "Complex dynamic formulas and macro "
                    "execution paths may remain unresolved"
                ]
            ]
        )

        output.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        wb.save(output)

        print()
        print("=" * 60)
        print("ANALYSIS COMPLETE")
        print("=" * 60)
        print("Total Sheets:", len(sheet_names))
        print("Formula Cells:", formula_count)
        print("Confirmed Data Links:", len(data_links))
        print(
            "Confirmed Navigation Links:",
            len(navigation_links)
        )
        print("INDIRECT Calls:", indirect_count)
        print(
            "Unresolved INDIRECT:",
            sum(r[2] == "INDIRECT" for r in unresolved)
        )
        print("Shape Records:", len(drawing))
        print("VBA Rows:", len(vba))
        print("Output:", output)

    finally:
        formula_wb.close()
        cache_wb.close()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
