APEXの知識が不足しているため、早急に習得する。
また、ArgoCDの運用ルールを理解し、プロジェクト全体の構成を早めに完成させ、業務開発に着手する。

python -c "from openpyxl import load_workbook; from collections import Counter; w=load_workbook('input_diagnostic.xlsx',read_only=True); s=w['Formula_Diagnostics']; c=Counter(str(r[4]) for r in list(s.values)[1:] if r[2]=='INDIRECT'); print(*c.most_common(),sep='\n')"
import argparse
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.styles import Font, PatternFill


def extract_args(formula, function):
    pattern = re.compile(
        r"(?<![\w.])" + re.escape(function) + r"\s*\(",
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
                if quoted and i + 1 < len(formula) and formula[i+1] == '"':
                    i += 2
                    continue
                quoted = not quoted

            elif not quoted:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        results.append(formula[start:i])
                        break
            i += 1

    return results


def split_concat(expr):
    parts = []
    quoted = False
    depth = 0
    start = 0
    i = 0

    while i < len(expr):
        ch = expr[i]

        if ch == '"':
            if quoted and i + 1 < len(expr) and expr[i+1] == '"':
                i += 2
                continue
            quoted = not quoted

        elif not quoted:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "&" and depth == 0:
                parts.append(expr[start:i].strip())
                start = i + 1
        i += 1

    parts.append(expr[start:].strip())
    return parts


def safe_cell(ws, row):
    for value in row:
        c = ws.cell(ws.max_row + 1, 1) if False else None
    ws.append(["" if x is None else str(x) for x in row])
    for cell in ws[ws.max_row]:
        cell.data_type = "s"
        cell.number_format = "@"


class Diagnostic:
    def __init__(self, filename):
        self.path = Path(filename)
        self.formulas = load_workbook(
            filename, data_only=False, read_only=True
        )
        self.cached = load_workbook(
            filename, data_only=True, read_only=True
        )

        self.names = list(self.formulas.sheetnames)
        self.lookup = {x.casefold(): x for x in self.names}
        self.stats = Counter()
        self.rows = []
        self.shapes = []

    def normalize(self, name):
        return self.lookup.get(str(name).casefold())

    def value(self, sheet, addr, depth=0, visited=None):
        if depth > 12:
            return None, "depth_limit"

        visited = visited or set()
        key = (sheet, addr.upper().replace("$", ""))

        if key in visited:
            return None, "cycle"

        visited = visited | {key}

        try:
            cell = self.formulas[sheet][addr]
            value = cell.value
            cache = self.cached[sheet][addr].value

            if cell.data_type != "f":
                return value, "literal"

            if cache is not None:
                return cache, "cached"

            return self.evaluate(
                sheet, str(value)[1:], depth + 1, visited
            )
        except Exception as e:
            return None, "read_error:" + type(e).__name__

    def evaluate(self, sheet, expr, depth=0, visited=None):
        expr = expr.strip()

        if depth > 12:
            return None, "depth_limit"

        parts = split_concat(expr)

        if len(parts) > 1:
            values = []
            statuses = []

            for part in parts:
                val, status = self.evaluate(
                    sheet, part, depth + 1, visited
                )
                if val is None:
                    return None, status
                values.append(str(val))
                statuses.append(status)

            return "".join(values), "concat:" + ",".join(statuses)

        if expr.startswith('"') and expr.endswith('"'):
            return expr[1:-1].replace('""', '"'), "string"

        m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", expr, re.I)

        if m:
            return self.value(
                sheet, expr.replace("$", ""),
                depth + 1, visited
            )

        m = re.fullmatch(
            r"(?:'((?:[^']|'')+)'|([^'!]+))!"
            r"(\$?[A-Z]{1,3}\$?\d+)",
            expr, re.I
        )

        if m:
            target = self.normalize(
                (m.group(1) or m.group(2)).replace("''", "'")
            )
            if target:
                return self.value(
                    target, m.group(3).replace("$", ""),
                    depth + 1, visited
                )
            return None, "unknown_sheet"

        return None, "unsupported_expression"

    def target(self, value):
        if not isinstance(value, str) or "!" not in value:
            return None

        name = value.rsplit("!", 1)[0].strip()

        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")

        if "[" in name or "]" in name:
            return None

        return self.normalize(name)

    def scan_formulas(self):
        print("Scanning formulas...", flush=True)

        for sheet in self.names:
            ws = self.formulas[sheet]

            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != "f":
                        continue

                    self.stats["Formula cells"] += 1
                    formula = str(cell.value)

                    if self.stats["Formula cells"] % 10000 == 0:
                        print(
                            "Processed:",
                            self.stats["Formula cells"],
                            flush=True
                        )

                    try:
                        tokens = Tokenizer(formula).items
                        explicit = set()

                        for token in tokens:
                            if token.type == "OPERAND" and token.subtype == "RANGE":
                                text = token.value
                                if "!" in text:
                                    raw = text.rsplit("!", 1)[0]
                                    raw = raw.strip("'").replace("''", "'")
                                    target = self.normalize(raw)
                                    if target:
                                        explicit.add(target)

                        for target in explicit:
                            self.stats["Explicit references"] += 1
                            self.rows.append([
                                sheet, cell.coordinate,
                                "EXPLICIT", target,
                                "confirmed", formula
                            ])

                    except Exception as e:
                        self.stats["Tokenizer errors"] += 1
                        self.rows.append([
                            sheet, cell.coordinate,
                            "TOKENIZER_ERROR", "",
                            type(e).__name__, formula
                        ])

                    for fn in ("INDIRECT", "OFFSET", "HYPERLINK"):
                        calls = extract_args(formula, fn)

                        for args in calls:
                            self.stats[fn + " calls"] += 1

                            if fn != "INDIRECT":
                                continue

                            # First argument only.
                            # Split on comma outside strings and parentheses.
                            first = args
                            depth = 0
                            quoted = False

                            for i, ch in enumerate(args):
                                if ch == '"':
                                    quoted = not quoted
                                elif not quoted:
                                    if ch == "(":
                                        depth += 1
                                    elif ch == ")":
                                        depth -= 1
                                    elif ch == "," and depth == 0:
                                        first = args[:i]
                                        break

                            result, status = self.evaluate(
                                sheet, first
                            )

                            target = self.target(result)

                            if target:
                                self.stats["INDIRECT resolved"] += 1
                                if target == sheet:
                                    self.stats["INDIRECT self links"] += 1
                                else:
                                    self.stats["INDIRECT cross-sheet links"] += 1
                            else:
                                self.stats["INDIRECT unresolved"] += 1

                            self.rows.append([
                                sheet,
                                cell.coordinate,
                                "INDIRECT",
                                target or "",
                                status,
                                formula
                            ])

    def scan_shapes(self):
        print("Scanning shape metadata...", flush=True)

        with zipfile.ZipFile(self.path) as zf:
            for part in zf.namelist():
                if not (
                    part.startswith("xl/drawings/")
                    and part.endswith((".xml", ".vml"))
                ):
                    continue

                try:
                    root = ET.fromstring(zf.read(part))
                except Exception:
                    self.stats["Shape XML errors"] += 1
                    continue

                for elem in root.iter():
                    tag = elem.tag.rsplit("}", 1)[-1]

                    if tag in ("sp", "pic", "graphicFrame"):
                        macro = elem.get("macro")
                        if macro:
                            self.stats["Drawing macro bindings"] += 1
                            self.shapes.append([
                                part, "Drawing macro", macro
                            ])

                    if tag == "FmlaMacro" and elem.text:
                        self.stats["VML macro bindings"] += 1
                        self.shapes.append([
                            part, "VML macro", elem.text
                        ])

                    if tag == "hlinkClick":
                        self.stats["Drawing hyperlink elements"] += 1

    def report(self, output):
        wb = Workbook()
        summary = wb.active
        summary.title = "Summary"
        summary.append(["Metric", "Count"])

        for name, count in sorted(self.stats.items()):
            summary.append([name, count])

        sheets = wb.create_sheet("Sheets")
        sheets.append(["Sheet", "Visibility"])

        for name in self.names:
            sheets.append([
                name,
                self.formulas[name].sheet_state
            ])

        details = wb.create_sheet("Formula_Diagnostics")
        details.append([
            "Source Sheet",
            "Cell",
            "Type",
            "Resolved Target",
            "Status",
            "Expression"
        ])

        for row in self.rows:
            safe_cell(details, row)

        shapes = wb.create_sheet("Shape_Diagnostics")
        shapes.append(["XML Part", "Type", "Macro"])

        for row in self.shapes:
            safe_cell(shapes, row)

        for ws in wb:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.fill = PatternFill(
                    "solid", fgColor="1F4E78"
                )
                cell.font = Font(
                    color="FFFFFF", bold=True
                )

            ws.column_dimensions["A"].width = 38
            ws.column_dimensions["B"].width = 25
            ws.column_dimensions["C"].width = 25
            ws.column_dimensions["D"].width = 32
            ws.column_dimensions["E"].width = 38
            ws.column_dimensions["F"].width = 100

        wb.save(output)

    def run(self, output):
        try:
            self.scan_formulas()
            self.scan_shapes()
            self.report(output)

            print("\n===== DIAGNOSTIC RESULT =====")
            print("Sheets:", len(self.names))

            for name, count in sorted(self.stats.items()):
                print(f"{name}: {count}")

            print("Report:", output)
        finally:
            self.formulas.close()
            self.cached.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    parser.add_argument("-o", "--output")
    args = parser.parse_args()

    source = Path(args.input)

    if not source.is_file():
        parser.error("Input file not found")

    output = (
        Path(args.output)
        if args.output
        else source.with_name(
            source.stem + "_diagnostic.xlsx"
        )
    )

    Diagnostic(source).run(output)
