"""One-time upgrade of Non-Op Evaluation Template.xlsx: multi-zone type curves (kept for the record).

Usage: python upgrade_template_zones.py <in.xlsx> <out.xlsx>
  * new 'Zone Curves' tab: up to 6 zone curves per 10,000 ft (oil / gas pairs in C:N, zone names in row 4),
    a fit summary per zone (P4 down) and charts for zones 2-6 (written by write_model.py / plot_tc.py)
  * Wells column X = the well's zone (blank or zone 1 = the TypeCurve tab, including the Enverus P50 switch);
    column Y = 1 when the zone is found on 'Zone Curves'
  * CF W1..W10 oil / gas volumes look up the well's zone curve; Well TC factor (column R) still applies on top
  * Summary check D73 also fails when an active well names a zone that has no curve; Well Summary column U shows the zone
"""
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

src, out = sys.argv[1:3]
wb = openpyxl.load_workbook(src)
W, S, WS, TC = wb["Wells"], wb["Summary"], wb["Well Summary"], wb["TypeCurve"]
HDR_FONT = Font(color="FFFFFFFF", bold=True); HDR_FILL = PatternFill("solid", fgColor="FF1F3864")
BLUE = Font(color="FF0000FF")
NZ = 6

# ------------------------------------------------------------------ Zone Curves tab
if "Zone Curves" in wb.sheetnames:
    del wb["Zone Curves"]
Z = wb.create_sheet("Zone Curves", wb.sheetnames.index("TypeCurve") + 1)
Z["B2"] = "Zone type curves per 10,000 ft lateral (gross 8/8; oil bbl/month, gas Mcf/month; before the child haircut)"
Z["B2"].font = Font(bold=True, size=14)
Z["B3"] = ("One curve per producing zone, each fitted on its own offsets by build_tc.py (same multi-phase method as TypeCurve, b fitted per zone). "
           "Zone 1 is the TypeCurve tab. A well uses the curve named in Wells column X (blank = zone 1 / TypeCurve). Written by write_model.py; "
           "empty for a single-zone deal.")
Z["B3"].alignment = Alignment(wrap_text=True, vertical="top"); Z.merge_cells("B3:N3"); Z.row_dimensions[3].height = 45
Z["B4"] = "Zone"; Z["B5"] = "Month"
for c in ("B4", "B5"):
    Z[c].font = HDR_FONT; Z[c].fill = HDR_FILL
for k in range(NZ):
    co, cg = 3 + 2 * k, 4 + 2 * k
    for col, lab in ((co, "Oil (bbl)"), (cg, "Gas (Mcf)")):
        Z.cell(4, col).fill = HDR_FILL; Z.cell(4, col).font = HDR_FONT
        Z.cell(5, col, lab).font = HDR_FONT; Z.cell(5, col).fill = HDR_FILL
        Z.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 12
for m in range(360):
    Z.cell(6 + m, 2, m + 1)
heads = ["Zone", "Fit radius", "Fit wells", "Wells 24+ mo", "Low confidence", "Oil b", "Gas b", "Oil EUR (MBO / 10k ft)",
         "Gas EUR (Bcf / 10k ft)", "10-mi comparison wells", "10-mi oil EUR (MBO)", "10-mi gas EUR (Bcf)", "Notes"]
for j, h in enumerate(heads):
    c = Z.cell(4, 16 + j, h); c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = Alignment(wrap_text=True, vertical="top")
    Z.column_dimensions[openpyxl.utils.get_column_letter(16 + j)].width = 13
Z.column_dimensions["AB"].width = 60; Z.row_dimensions[4].height = 45
for k in range(NZ):
    Z.cell(5 + k, 16, f"=IF({openpyxl.utils.get_column_letter(3 + 2 * k)}4=\"\",\"\",{openpyxl.utils.get_column_letter(3 + 2 * k)}4)")
Z["P12"] = "Charts for zones 2-6 below (oil, then gas, per zone); zone 1 charts are on the TypeCurve tab."
Z["P12"].font = Font(italic=True, color="FF6B6A63")
Z.freeze_panes = "C6"

# ------------------------------------------------------------------ Wells: zone column
for col, lab in (("X", "Zone (type curve; blank = TypeCurve)"), ("Y", "Zone found (1 = yes)")):
    c = W[f"{col}6"]; c.value = lab
    c.font = copy(W["W6"].font); c.fill = copy(W["W6"].fill); c.alignment = copy(W["W6"].alignment); c.border = copy(W["W6"].border)
    W.column_dimensions[col].width = 18
dv = DataValidation(type="list", formula1="='Zone Curves'!$C$4:$N$4", allow_blank=True, showErrorMessage=False)
W.add_data_validation(dv)
for r in range(7, 17):
    W[f"X{r}"].font = BLUE; dv.add(f"X{r}")
    W[f"Y{r}"] = f"=IF(X{r}=\"\",1,IF(ISNUMBER(MATCH(X{r},'Zone Curves'!$C$4:$N$4,0)),1,0))"
W["B3"] = (W["B3"].value or "") + " Column X picks the well's zone curve on 'Zone Curves' (blank = the TypeCurve tab)."

# ------------------------------------------------------------------ CF sheets
for i in range(10):
    ws = wb[f"CF W{i + 1}"]; wr = 7 + i
    zx = f"Wells!$X${wr}"
    prim = f"OR({zx}=\"\",{zx}='Zone Curves'!$C$4)"
    match = f"MATCH({zx},'Zone Curves'!$C$4:$N$4,0)"
    for r in range(6, 406):
        for col, k, off in (("G", "2*Inputs!$C$13-1", ""), ("H", "2*Inputs!$C$13", "+1")):
            f = ws[f"{col}{r}"].value
            old = f"INDEX(TypeCurve!$C$6:$F$365,D{r},{k})"
            assert old in f, (ws.title, col, r)
            new = f"IF({prim},{old},IFERROR(INDEX('Zone Curves'!$C$6:$N$365,D{r},{match}{off}),0))"
            ws[f"{col}{r}"] = f.replace(old, new)

# ------------------------------------------------------------------ checks and Well Summary
S["B73"] = "Every active well has a lateral, a first-sales month and a zone with a curve"
S["D73"] = ('=IF(SUMPRODUCT((Wells!$V$7:$V$16=1)*((Wells!$E$7:$E$16=0)+(Wells!$P$7:$P$16=0)+(Wells!$Y$7:$Y$16=0)))=0,"OK","CHECK")')
c = WS["U4"]; c.value = "Zone (type curve)"
for a in ("font", "fill", "alignment", "border"):
    setattr(c, a, copy(getattr(WS["T4"], a)))
WS.column_dimensions["U"].width = 18
for i in range(10):
    WS[f"U{5 + i}"] = (f"=IF(C{5 + i}=\"\",\"\",IF(Wells!X{7 + i}=\"\",IF('Zone Curves'!$C$4=\"\",\"TypeCurve\",'Zone Curves'!$C$4),Wells!X{7 + i}))")
wb.save(out)
print("wrote", out)
