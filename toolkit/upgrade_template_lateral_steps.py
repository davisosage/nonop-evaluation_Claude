"""One-time template upgrade (Nov 2026): two-step early-month schedule. Kept for the record.

Usage: python upgrade_template_lateral_steps.py <in.xlsx> <out.xlsx>
  * Inputs C26/C27: 92.5% at 5,000 ft away (3 miles vs 2); new C28/C29: 85% at 10,000 ft away (4 miles vs 2), held beyond
  * 'Lateral Curves' row 10: 100% within C24, straight lines through (C27, C26) and (C29, C28), C28 beyond
"""
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as L

src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
I, LC = wb["Inputs"], wb["Lateral Curves"]
for r in (28, 29):
    for col in "BCD":
        I[f"{col}{r}"]._style = copy(I[f"{col}{r - 2}"]._style)
    I.row_dimensions[r].height = 30
I["B26"], I["C26"] = "Early-month factor at the first step (% of linear)", 0.925
I["B27"], I["C27"] = "First step at (ft away from the curve's length)", 5000
I["B28"], I["C28"] = "Early-month factor at the second step and beyond", 0.85
I["B29"], I["C29"] = "Second step at (ft away)", 10000
for a in ("C26", "C28"):
    I[a].number_format = "0.0%"
I["C29"].number_format = "#,##0"
I["D26"] = ("92.5% at 5,000 ft away: a 3-mile well makes 1.5 x 92.5% of a 2-mile well in months 1-8. Going shorter it is "
            "linear / factor. From month 9 the hyperbolic decline is re-solved (same b) so the 30-year EUR is exactly linear.")
I["D27"] = ("The factor is 100% within C24 ft, falls in a straight line to C26 at this distance, then to C28 at C29 ft.")
I["D28"] = "85% at 10,000 ft away: a 4-mile well makes 2 x 85% of a 2-mile well in months 1-8."
I["D29"] = ("Held at C28 beyond this distance. Shown per well on 'Lateral Curves'; re-run write_model.py after changing "
            "C24:C29 or a lateral.")
for r in range(26, 30):
    I[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
LC["B10"] = "Early-month factor (x linear; stepped)"
for j in range(10):
    for k in (2, 3):
        c = L(3 + 4 * j + k)
        d = f"ABS({c}$7-10000)"
        f = (f"IF({d}<=Inputs!$C$27,1-(1-Inputs!$C$26)*({d}-Inputs!$C$24)/MAX(1,Inputs!$C$27-Inputs!$C$24),"
             f"IF({d}<=Inputs!$C$29,Inputs!$C$26+(Inputs!$C$28-Inputs!$C$26)*({d}-Inputs!$C$27)/MAX(1,Inputs!$C$29-Inputs!$C$27),"
             f"Inputs!$C$28))")
        LC[f"{c}10"] = f'=IF({c}$8="Shaped",IF({c}$9>1,{f},1/{f}),1)'
LC["B3"] = LC["B3"].value.replace(
    "(phased: 100% at C24 ft from 10,000, falling in a straight line to C26 at C27 ft, C26 beyond;",
    "(100% within C24 ft of 10,000, straight lines to C26 at C27 ft and C28 at C29 ft, C28 beyond;")
wb.save(out)
print("wrote", out)
