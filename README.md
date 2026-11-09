
from openpyxl import load_workbook

FILE = "input.xlsm"

wb_formula = load_workbook(
    FILE,
    read_only=True,
    data_only=False
)

wb_value = load_workbook(
    FILE,
    read_only=True,
    data_only=True
)

print("========== E36 DIAGNOSTIC ==========")

for sheet_name in wb_formula.sheetnames:
    ws = wb_formula[sheet_name]

    for address in ("E36", "E35", "E37"):
        cell = ws[address]

        if cell.value is None:
            continue

        cached = wb_value[sheet_name][address].value

        print()
        print("Sheet:", sheet_name)
        print("Cell:", address)
        print("Formula/Value:", repr(cell.value))
        print("Cached Value:", repr(cached))
        print("Cell Type:", cell.data_type)

wb_formula.close()
wb_value.close()
