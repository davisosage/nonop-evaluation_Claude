"""One-time template upgrade (Nov 23, 2026): lateral rule calibrated to Enverus. Kept for the record.

Usage: python upgrade_template_lateral_calibrated.py <in.xlsx> <out.xlsx>
  * Inputs C24 linear band 2,500 -> 0 ft; C25 early months 8 -> 3; C26 92.5% -> 75% at C27 5,000 ft (3 miles vs 2);
    C28 85% -> 67.5% at C29 10,000 ft (placeholder, held beyond). Notes in D24:D29 rewritten.
  * The 'Lateral Curves' formulas already read C24:C29, so no formula changes.
Calibration: Enverus lateral_stats, first production 2020-2024, 3-mile vs 2-mile per foot, matched by operator x year x
interval x county (see README, lateral length rule).
"""
import sys
import openpyxl
from openpyxl.styles import Alignment

src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
I = wb["Inputs"]
I["C24"], I["C25"], I["C26"], I["C27"], I["C28"], I["C29"] = 0, 3, 0.75, 5000, 0.675, 10000
I["D24"] = ("Type curves are per 10,000 ft. An offset (when building the curve) or a deal well (when using it) within this "
            "many ft of 10,000 scales linearly. 0 since the Nov 2026 Enverus calibration: per-foot output falls steadily "
            "with length, so every move off 10,000 ft holds the early months below linear (going longer) or above it "
            "(going shorter) and reshapes the decline so the EUR is still linear.")
I["D25"] = ("Months 1-3 on production (ramp included) are scaled at the factor below instead of linearly (was 8). Enverus "
            "3-mile vs 2-mile wells recover from about 70% of linear per foot at month 6 to 79% at month 12 and 88-90% at "
            "month 36; a 3-month hold then the re-solved decline tracks that.")
I["D26"] = ("75% at 5,000 ft away: a 3-mile well makes 1.5 x 75% of a 2-mile well in months 1-3 (was 92.5% for months 1-8). "
            "Matched Enverus data (2020-2024 first production) shows about 70% in both Williston Middle Bakken and DJ "
            "Niobrara/Codell; 75% allows for lighter DJ completions and where long laterals get drilled. Going shorter it "
            "is linear / factor. From month 4 the hyperbolic decline is re-solved (same b) so the 30-year EUR is linear.")
I["D27"] = "The factor is 100% within C24 ft, falls in a straight line to C26 at this distance, then to C28 at C29 ft."
I["D28"] = ("67.5% at 10,000 ft away and beyond (was 85%): a 4-mile well makes 2 x 67.5% of a 2-mile well in months 1-3. "
            "A placeholder (too few 4-mile wells in Enverus to calibrate) that keeps the old 7.5-point step from 3 to 4 "
            "miles, so a curve built on 3-mile offsets still carries a penalty onto 4-mile wells.")
I["D29"] = ("Held at C28 beyond this distance. Shown per well on 'Lateral Curves'; re-run write_model.py after changing "
            "C24:C29 or a lateral.")
for r in range(24, 30):
    I[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
I["C24"].number_format = "#,##0"
wb.save(out)
print("wrote", out)
