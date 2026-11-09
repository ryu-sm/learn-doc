@'
from openpyxl import load_workbook
from collections import Counter

wb = load_workbook(
    "input_diagnostic.xlsx",
    read_only=True,
    data_only=True
)

ws = wb["Formula_Diagnostics"]

formulas = []

for row in list(ws.values)[1:]:
    if row[2] != "INDIRECT":
        continue

    formula = str(row[5] or "")

    if "!" in formula:
        formulas.append(
            (row[0], row[1], formula)
        )

print("INDIRECT formulas with !:", len(formulas))

patterns = Counter()

for sheet, cell, formula in formulas:
    f = formula.upper()

    if "ADDRESS(" in f:
        patterns["ADDRESS"] += 1

    if "HLOOKUP(" in f:
        patterns["HLOOKUP"] += 1

    if "&" in formula:
        patterns["CONCAT"] += 1

    if "$E" in f:
        patterns["COLUMN_E_REFERENCE"] += 1

print("\nPatterns:")
for k, v in patterns.items():
    print(k, v)

print("\nSample formulas:")

for sheet, cell, formula in formulas[:5]:
    print("\nSheet:", sheet)
    print("Cell:", cell)
    print("Formula:", formula[:500])
'@ | python -