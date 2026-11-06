# learn-doc
学习资料

import argparse
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

MATRIX_NAME = "Direct_Matrix"

REF_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(?:'((?:[^']|'')+)'|([A-Za-z_\\][A-Za-z0-9_.]*))!"
)


def extract_references(formula, sheet_names):
    """Extract directly referenced sheet names."""
    result = set()

    # Ignore string literals
    formula = re.sub(r'"(?:[^"]|"")*"', '""', formula)

    name_map = {
        name.lower(): name for name in sheet_names
    }

    for match in REF_PATTERN.finditer(formula):
        name = match.group(1) or match.group(2)
        name = name.replace("''", "'")

        # Handle 3D sheet references
        if ":" in name:
            parts = name.split(":")

            if len(parts) == 2:
                start = name_map.get(parts[0].lower())
                end = name_map.get(parts[1].lower())

                if start and end:
                    a = sheet_names.index(start)
                    b = sheet_names.index(end)

                    result.update(
                        sheet_names[min(a, b):max(a, b) + 1]
                    )
        else:
            actual_name = name_map.get(name.lower())

            if actual_name:
                result.add(actual_name)

    return result


def analyze_dependencies(workbook, sheet_names):
    """Scan formulas and build direct dependencies."""
    graph = {
        name: set() for name in sheet_names
    }

    for source_name in sheet_names:
        ws = workbook[source_name]
        formula_count = 0

        for row in ws.iter_rows():
            for cell in row:

                if cell.data_type != "f":
                    continue

                formula_count += 1

                references = extract_references(
                    str(cell.value),
                    sheet_names
                )

                # Exclude self-reference
                references.discard(source_name)

                graph[source_name].update(references)

        print(
            f"{source_name}: "
            f"{formula_count} formulas, "
            f"{len(graph[source_name])} dependencies"
        )

    return graph


def create_matrix(workbook, sheet_names, graph):
    """Create direct dependency matrix."""
    if MATRIX_NAME in workbook.sheetnames:
        del workbook[MATRIX_NAME]

    ws = workbook.create_sheet(MATRIX_NAME)

    ws.cell(1, 1, "Source / Target")

    for index, name in enumerate(sheet_names, 2):
        ws.cell(1, index, name)
        ws.cell(index, 1, name)

    for i, source in enumerate(sheet_names, 2):
        for j, target in enumerate(sheet_names, 2):

            value = int(target in graph[source])
            cell = ws.cell(i, j, value)

            cell.alignment = Alignment(
                horizontal="center"
            )

            if value == 1:
                cell.fill = PatternFill(
                    "solid",
                    fgColor="C6EFCE"
                )
                cell.font = Font(
                    color="006100",
                    bold=True
                )

    # Header formatting
    for cell in ws[1]:
        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )
        cell.font = Font(
            color="FFFFFF",
            bold=True
        )

    ws.column_dimensions["A"].width = 30

    for col in range(2, len(sheet_names) + 2):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 20

    ws.freeze_panes = "B2"


def main(input_file, output_file):
    wb = load_workbook(
        input_file,
        data_only=False
    )

    sheet_names = [
        name for name in wb.sheetnames
        if name != MATRIX_NAME
    ]

    graph = analyze_dependencies(
        wb, sheet_names
    )

    create_matrix(
        wb, sheet_names, graph
    )

    wb.save(output_file)

    print("\nAnalysis completed.")
    print(f"Total sheets: {len(sheet_names)}")
    print(
        "Total direct dependencies:",
        sum(len(v) for v in graph.values())
    )
    print(f"Output: {output_file}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("input")
    parser.add_argument("-o", "--output")

    args = parser.parse_args()

    source = Path(args.input)

    output = (
        Path(args.output)
        if args.output
        else source.with_name(
            source.stem + "_direct_matrix.xlsx"
        )
    )

    if source.resolve() == output.resolve():
        parser.error(
            "Output file must differ from input file"
        )

    main(source, output)
