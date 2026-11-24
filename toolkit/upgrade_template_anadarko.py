"""One-time upgrade of Non-Op Evaluation Template.xlsx (Nov 23, 2026): adds an "Anadarko / Midcon (OK)" basin.

  * Basins tab: new column H with the Anadarko / Midcon defaults (first used on the TNE_OK_2026-11-20 package: Mewbourne /
    Nadel & Gussman Cherokee and Tonkawa wells in Roger Mills and Custer Co OK); the Notes column moves from H to I.
  * Inputs C41:C57: the basin lookups read Basins D:H instead of D:G.
  * Inputs C19 dropdown: Anadarko / Midcon (OK) added.
The engine reads the recalculated Inputs values, so engine.py / write_model.py need no change. Recalculate afterwards
(LibreOffice headless or Excel) and confirm Summary shows "All checks OK" on the Sun 5899 example the template carries.

    python3 upgrade_template_anadarko.py "../Non-Op Evaluation Template.xlsx"
"""
import re
import sys
from copy import copy

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

BASIN = "Anadarko / Midcon (OK)"
BASINS = ["Williston (ND)", "DJ (CO)", "Permian (TX)", "Powder River (WY)", BASIN]
# Basins row: (value, note added to the row's note)
VALUES = {
    5: (-1.50, "Anadarko: Oklahoma sweet at the lease, close to Cushing; confirm on operator check stubs"),
    6: (1.50, "Anadarko: trucking / gathering to Cushing-area purchasers"),
    7: (-0.55, "Anadarko on Midcon pricing (NGPL Midcon / ONEOK / Panhandle prompt basis -$0.44 to -$0.57, Sep 2026)"),
    8: (0.20, "Anadarko: wet gas, plant shrink + fuel"),
    9: (100, "Anadarko Cherokee / Tonkawa gas is rich"),
    10: (0.28, "Anadarko: Conway-priced NGL after T&F"),
    11: (0.75, "Anadarko: Midcon gathering and processing fee"),
    12: (0.052, "OK: gross production tax 5% for the first 36 months (wells spudded after Jul 2018) + ~0.2% excise / conservation"),
    13: (0.072, "OK: 7% gross production tax after month 36 + ~0.2% excise / conservation"),
    14: (8000, None),
    15: (0.35, None),
    16: (3.0, None),
    17: (1.5, "Anadarko: Enverus water is blank for Oklahoma wells; 1.5 is an assumption"),
    18: (1.0, None),
    19: (125000, None),
}

f = sys.argv[1] if len(sys.argv) > 1 else "../Non-Op Evaluation Template.xlsx"
wb = openpyxl.load_workbook(f)
B, I = wb["Basins"], wb["Inputs"]
if B["H4"].value == BASIN:
    raise SystemExit("already upgraded")
assert B["H4"].value == "Notes" and B["G4"].value == "Powder River (WY)", "unexpected Basins layout"

# the B3 note is merged across B:H; unmerge, move columns, re-merge across B:I
merged = [str(m) for m in B.merged_cells.ranges if m.max_col == 8]
for m in merged:
    B.unmerge_cells(m)
# move Notes H -> I (values and styles), then fill H as a basin column styled like G
for r in range(4, B.max_row + 1):
    h, i = B.cell(r, 8), B.cell(r, 9)
    i.value = h.value
    if h.has_style:
        i._style = copy(h._style)
    g = B.cell(r, 7)
    h.value = None
    h._style = copy(g._style)
B.column_dimensions["I"].width = B.column_dimensions["H"].width or 90
B.column_dimensions["H"].width = 22
B["H4"] = BASIN
for r, (v, note) in VALUES.items():
    B.cell(r, 8).value = v
    B.cell(r, 8).number_format = B.cell(r, 7).number_format
    if note:
        old = B.cell(r, 9).value
        B.cell(r, 9).value = f"{old}. {note}" if old else note
for m in merged:
    B.merge_cells(m.replace("H", "I"))
B["B3"] = (B["B3"].value or "") + " Anadarko / Midcon (OK) added Nov 2026 (Oklahoma assumptions: confirm on JIBs, check stubs and the gas contract)."

# Inputs lookups: D:G -> D:H
n = 0
for r in range(36, 60):
    c = I.cell(r, 3)
    if isinstance(c.value, str) and "Basins!" in c.value:
        c.value = re.sub(r"(Basins!\$D\$\d+:\$)G(\$\d+)", r"\1H\2", c.value)
        n += 1
assert n == 15, n

# C19 dropdown
I.data_validations.dataValidation = [dv for dv in I.data_validations.dataValidation if "C19" not in str(dv.sqref)]
dv = DataValidation(type="list", formula1=f'"{",".join(BASINS)}"', allow_blank=False)
dv.add("C19")
I.add_data_validation(dv)
I["D19"] = "Defaults live on the Basins tab (Williston, DJ, Permian, Powder River, Anadarko / Midcon); type a number in column E next to any value to override it for this deal"

wb.save(f)
print("upgraded:", f, "| lookups changed:", n)
