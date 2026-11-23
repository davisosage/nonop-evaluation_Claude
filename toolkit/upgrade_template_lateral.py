"""One-time template upgrade (Nov 2026): lateral-length rule. Kept for the record.

Usage: python upgrade_template_lateral.py <in.xlsx> <out.xlsx>
  * Inputs C24:C26: linear band (+/- ft, 2,500), early months (8), early-month factor (90% of linear)
  * new 'Lateral Curves' tab: each well's curve per 10,000 ft (its zone, TypeCurve switch C13) moved to its lateral,
    with s / Tsw / g written by write_model.py, and an EUR check against linear scaling
  * CF W1..W10 oil / gas read the well's column on 'Lateral Curves' instead of curve x lateral / 10,000
  * Summary: new check row 76 (each active well's EUR at its lateral ties to linear); Model status moves to row 77
"""
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
I, W, S = wb["Inputs"], wb["Wells"], wb["Summary"]

for ws in wb.worksheets:                         # nothing should point at the old Model status cell
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and ("Summary!C76" in c.value.replace("$", "") or "Summary!D7" in c.value.replace("$", "")):
                raise SystemExit(f"{ws.title}!{c.coordinate} refers to C76: {c.value}")

# ---- Inputs block
for col in "BCD":
    I[f"{col}23"]._style = copy(I[f"{col}17"]._style)
I["B23"], I["C23"], I["D23"] = "Lateral length scaling", "Value", "Basis"
I["C23"].number_format = "General"
rows = [(24, "Linear band (+/- ft from the length the curve is moved from)", 2500, "#,##0",
         "Type curves are per 10,000 ft. A well within this many ft of 10,000 (and an offset within this many ft of the "
         "modeled lateral) scales linearly; outside it the early months and the decline are reshaped (below)."),
        (25, "Early months held below linear", 8, "0",
         "Months 1-8 on production (ramp included) are scaled at the factor below instead of linearly."),
        (26, "Early-month factor (% of linear scaling)", 0.9, "0%",
         "90%: a 4-mile well makes 1.8x (not 2x) a 2-mile well in months 1-8; going shorter it is linear / 90%. From "
         "month 9 the hyperbolic decline is re-solved (same b, shallower going longer) so the 30-year EUR is exactly "
         "linear. Shown per well on 'Lateral Curves'; re-run write_model.py after changing these or a lateral.")]
for r, lab, v, fmt, note in rows:
    I[f"B{r}"]._style = copy(I["B18"]._style); I[f"D{r}"]._style = copy(I["D18"]._style)
    I[f"C{r}"]._style = copy(I["C11"]._style)
    I[f"B{r}"], I[f"C{r}"], I[f"D{r}"] = lab, v, note
    I[f"C{r}"].number_format = fmt
    I[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    I.row_dimensions[r].height = 30

# ---- Lateral Curves tab
if "Lateral Curves" in wb.sheetnames:
    del wb["Lateral Curves"]
LC = wb.create_sheet("Lateral Curves", wb.sheetnames.index("Zone Curves") + 1)
LC.sheet_properties.tabColor = "6B2FB3"
LC["B2"] = "Type curve at each well's lateral (gross 8/8; oil bbl/month, gas Mcf/month; before the TC multiplier, TC factor and child haircut)"
LC["B2"].font = Font(bold=True, size=14)
LC["B3"] = ("Base = the well's curve per 10,000 ft (its zone on 'Zone Curves', else TypeCurve; Inputs C13 picks base / Enverus P50). "
            "At lateral: within Inputs C24 ft of 10,000 = base x lateral / 10,000 (linear). Otherwise months 1..C25 = base x "
            "lateral / 10,000 x early factor (C26 going longer, 1 / C26 going shorter), and from the next month the decline is "
            "stretched in time by s (same b; until s is solved the well scales linearly and row 17 says RE-RUN; terminal 7% kept from month Tsw, monthly ratio g) so the 360-month EUR equals the "
            "linear EUR. s, Tsw and g are solved by write_model.py; row 17 flags a well whose EUR no longer ties (lateral or "
            "inputs changed): re-run write_model.py. The CF tabs read the 'at lateral' columns.")
LC["B3"].alignment = Alignment(wrap_text=True, vertical="top")
LC.merge_cells("B3:Z3"); LC.row_dimensions[3].height = 60
hdr = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="1F3864")
lab = Font(italic=True, color="6B6A63"); calc = Font(color="808080", italic=True)
labels = {5: "Well", 6: "Column", 7: "Lateral (ft)", 8: "Scaling", 9: "Lateral / 10,000", 10: "Early-month factor (x linear)",
          11: "Decline stretch s (write_model.py)", 12: "Terminal from month Tsw (write_model.py)",
          13: "Terminal monthly ratio g (write_model.py)", 14: "EUR at lateral (360 mo)", 15: "Linear EUR (base x lateral / 10,000)",
          16: "EUR / linear", 17: "Check", 19: "Month"}
for r, t in labels.items():
    LC.cell(r, 2, t).font = lab
LC.column_dimensions["B"].width = 34
LC.cell(19, 2).font = hdr; LC.cell(19, 2).fill = fill
R0, R1 = 20, 379
for m in range(360):
    LC.cell(R0 + m, 2, m + 1)
M1 = "(Inputs!$C$25+1)"
for j in range(10):
    wr = 7 + j
    for k, (kind, stream) in enumerate((("Base", "oil"), ("Base", "gas"), ("At lateral", "oil"), ("At lateral", "gas"))):
        col = 3 + 4 * j + k
        c = L(col)
        LC.column_dimensions[c].width = 13
        LC.cell(5, col, f'=IF(Wells!$C${wr}="","(empty slot {j + 1})",Wells!$C${wr})').font = Font(bold=True)
        h = LC.cell(6, col, f"{kind} {stream}" + (" /10k" if kind == "Base" else "")); h.font = hdr; h.fill = fill
        LC.cell(19, col).fill = fill
        if kind == "Base":
            zone = (f"IF(OR(Wells!$X${wr}=\"\",Wells!$X${wr}='Zone Curves'!$C$4),"
                    f"INDEX(TypeCurve!$C$6:$F$365,$B{{r}},2*Inputs!$C$13{'-1' if stream == 'oil' else ''}),"
                    f"IFERROR(INDEX('Zone Curves'!$C$6:$N$365,$B{{r}},MATCH(Wells!$X${wr},'Zone Curves'!$C$4:$N$4,0){'+1' if stream == 'gas' else ''}),0))")
            for m in range(360):
                LC.cell(R0 + m, col, "=" + zone.format(r=R0 + m)).number_format = "#,##0"
            continue
        b = L(col - 2)                              # the base column of the same stream
        LC.cell(7, col, f"=N(Wells!$E${wr})").number_format = "#,##0"
        LC.cell(8, col, f'=IF({c}$7=0,"",IF(ABS({c}$7-10000)<=Inputs!$C$24,"Linear","Shaped"))')
        LC.cell(9, col, f"={c}$7/10000").number_format = "0.000"
        LC.cell(10, col, f'=IF({c}$8="Shaped",IF({c}$9>1,Inputs!$C$26,1/Inputs!$C$26),1)').number_format = "0.000"
        for r, fmt in ((11, "0.0000"), (12, "0"), (13, "0.000000")):
            LC.cell(r, col).number_format = fmt; LC.cell(r, col).font = calc
        LC.cell(14, col, f"=SUM({c}{R0}:{c}{R1})").number_format = "#,##0"
        LC.cell(15, col, f"={c}$9*SUM({b}{R0}:{b}{R1})").number_format = "#,##0"
        LC.cell(16, col, f'=IF(N({c}$15)=0,"",{c}$14/{c}$15)').number_format = "0.000%"
        LC.cell(17, col, f'=IF(OR(Wells!$V${wr}<>1,{c}$7=0),"",IF(AND({c}$8="Shaped",N({c}$11)=0),"RE-RUN",IF(ABS(N({c}$16)-1)<=0.001,"OK","RE-RUN")))')
        P = f"({M1}+($B{{r}}-{M1})*{c}$11)"
        for m in range(360):
            r = R0 + m
            p = P.format(r=r)
            f = (f"=IF({c}$7=0,0,IF(OR({c}$8=\"Linear\",N({c}$11)=0),{b}{r}*{c}$9,IF($B{r}<=Inputs!$C$25,{b}{r}*{c}$9*{c}$10,"
                 f"{c}$9*{c}$10*IF({p}<={c}$12,INDEX({b}${R0}:{b}${R1},INT({p}))*(1-({p}-INT({p})))"
                 f"+INDEX({b}${R0}:{b}${R1},MIN(INT({p})+1,360))*({p}-INT({p})),"
                 f"INDEX({b}${R0}:{b}${R1},{c}$12)*{c}$13^(({p}-{c}$12)/{c}$11)))))")
            LC.cell(r, col, f).number_format = "#,##0"
LC.freeze_panes = "C20"

# ---- CF tabs read the 'at lateral' columns
for j in range(10):
    cf = wb[f"CF W{j + 1}"]
    wr = 7 + j
    oc, gc = L(3 + 4 * j + 2), L(3 + 4 * j + 3)
    for r in range(6, cf.max_row + 1):
        if not (isinstance(cf[f"G{r}"].value, str) and cf[f"G{r}"].value.startswith("=IF(AND(D")):
            continue
        for colx, lc in (("G", oc), ("H", gc)):
            cf[f"{colx}{r}"] = (f"=IF(AND(D{r}>=1,D{r}<=360),INDEX('Lateral Curves'!${lc}$20:${lc}$379,D{r})"
                                f"*(Inputs!$C$11*Wells!$R${wr})*(1-Wells!$Q${wr}),0)")
W["B3"] = W["B3"].value.replace(
    "Volumes use the TypeCurve sheet x lateral / 10,000 x global TC multiplier",
    "Volumes use the well's curve at its lateral ('Lateral Curves': linear within Inputs C24 ft of 10,000, otherwise the "
    "early-month rule) x global TC multiplier")

# ---- Summary check
for col in "BCDEF":
    S[f"{col}77"]._style = copy(S[f"{col}76"]._style)
    S[f"{col}76"]._style = copy(S[f"{col}75"]._style)
S["B77"], S["C77"] = "Model status", '=IF(COUNTIF(D72:D76,"OK")=5,"All checks OK","CHECK FAILED")'
S["B76"] = "Every active well's EUR at its lateral ties to linear scaling ('Lateral Curves' row 17; RE-RUN = run write_model.py again)"
S["C76"] = '=COUNTIF(\'Lateral Curves\'!$C$17:$AP$17,"RE-RUN")'
S["D76"] = '=IF(C76=0,"OK","CHECK")'
S["E76"] = S["F76"] = None
wb.save(out)
print("wrote", out)
