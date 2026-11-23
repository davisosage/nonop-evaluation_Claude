"""One-time template upgrade (Nov 2026): phase the lateral early-month factor in. Kept for the record.

Usage: python upgrade_template_lateral_phase.py <in.xlsx> <out.xlsx>
  * Inputs C27: ft away from the curve's length at which the full C26 factor applies (5,000); between C24 and C27 the
    factor falls in a straight line from 100% to C26
  * 'Lateral Curves' row 10 (early-month factor) uses the phased factor
"""
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as L

src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
I, LC = wb["Inputs"], wb["Lateral Curves"]
for col in "BCD":
    I[f"{col}27"]._style = copy(I[f"{col}26"]._style)
I["B27"], I["C27"] = "Full early-month factor from (ft away)", 5000
I["C27"].number_format = "#,##0"
I["D27"] = ("The early-month factor is phased in: 100% (linear) at C24 ft from the curve's length, falling in a straight line to "
            "C26 at this many ft, and C26 beyond. With 2,500 / 5,000: a 2.5-mile well (12,500 ft) is linear, 13,750 ft gets 95%, "
            "3 miles (15,000 ft) and longer get 90%.")
I["D27"].alignment = Alignment(wrap_text=True, vertical="top"); I.row_dimensions[27].height = 30
I["B26"] = "Early-month factor at full effect (% of linear scaling)"
I["D26"] = ("90%: a 4-mile well makes 1.8x (not 2x) a 2-mile well in months 1-8; going shorter it is linear / factor. From "
            "month 9 the hyperbolic decline is re-solved (same b, shallower going longer) so the 30-year EUR is exactly "
            "linear. Shown per well on 'Lateral Curves'; re-run write_model.py after changing these or a lateral.")
I["B25"] = "Early months held below linear"
LC["B10"] = "Early-month factor (x linear; phased)"
for j in range(10):
    for k in (2, 3):
        c = L(3 + 4 * j + k)
        LC[f"{c}10"] = (f'=IF({c}$8="Shaped",IF({c}$9>1,'
                        f'1-(1-Inputs!$C$26)*MIN(1,(ABS({c}$7-10000)-Inputs!$C$24)/MAX(1,Inputs!$C$27-Inputs!$C$24)),'
                        f'1/(1-(1-Inputs!$C$26)*MIN(1,(ABS({c}$7-10000)-Inputs!$C$24)/MAX(1,Inputs!$C$27-Inputs!$C$24)))),1)')
LC["B3"] = LC["B3"].value.replace(
    "x early factor (C26 going longer, 1 / C26 going shorter)",
    "x early factor f (phased: 100% at C24 ft from 10,000, falling in a straight line to C26 at C27 ft, C26 beyond; "
    "f going longer, 1 / f going shorter)")
wb.save(out)
print("wrote", out)
