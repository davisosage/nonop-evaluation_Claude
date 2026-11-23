"""One-time upgrade of Non-Op Evaluation Template.xlsx: basin dropdown (Basins tab) and Minerals switch.

Usage: python upgrade_template.py <in.xlsx> <out.xlsx>
"""
import sys, datetime as dt
import openpyxl
from copy import copy
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

src, out = sys.argv[1:3]
wb = openpyxl.load_workbook(src)
I, W, S, WS = wb["Inputs"], wb["Wells"], wb["Summary"], wb["Well Summary"]
MIN = 'Inputs!$C$18="Minerals"'
BLUE = Font(color="FF0000FF"); BLACK = Font(color="FF000000"); GREY = Font(color="FF595959")
HDR_FONT = Font(color="FFFFFFFF", bold=True); HDR_FILL = PatternFill("solid", fgColor="FF1F3864")

# ------------------------------------------------------------------ Basins tab
BASINS = ["Williston (ND)", "DJ (CO)", "Permian (TX)", "Powder River (WY)"]
# (Inputs row, label, values by basin, number format, note)
PARAMS = [
    (41, "Oil differential to WTI ($/bbl)", [-1.50, -2.50, 0.75, -3.00], "0.00",
     "Wellhead-area crude vs NYMEX WTI before gathering. Chord 3Q26 Williston realized +$0.11 (Holmen deck kept at -$1.50); "
     "Midland trades at a premium to Cushing; DJ / PRB sell into Rockies-Guernsey-Cushing routes"),
    (42, "Oil gathering / trucking ($/bbl)", [2.75, 1.50, 1.25, 2.50], "0.00", "Pipeline-connected DJ / Permian vs more trucking in ND / PRB"),
    (43, "Residue gas differential to HH ($/MMBtu)", [-0.50, -0.60, -1.00, -0.60], "0.00",
     "Williston and DJ / PRB on Rockies (CIG) pricing; Permian on Waha (H1-26 Waha averaged -$2.19; forwards improve with 2026-27 pipelines)"),
    (44, "Gas shrink (flaring + plant)", [0.22, 0.25, 0.20, 0.20], "0%", "Share of wellhead gas lost to flaring, fuel and NGL extraction"),
    (45, "NGL yield (bbl/MMcf wellhead)", [87, 110, 120, 95], "0", "Rich-gas plays (DJ, Permian) yield more NGL"),
    (46, "NGL price (% of WTI)", [0.25, 0.30, 0.30, 0.27], "0%", "Chord 3Q26 Williston NGL realization 14% of WTI after fees"),
    (47, "Gas gathering and processing ($/Mcf wellhead)", [0.60, 0.85, 0.75, 0.80], "0.00", "Confirm against the operator's gas contract"),
    (50, "Production taxes, months 1-36", [0.096, 0.080, 0.075, 0.125], "0.0%",
     "ND: 5% oil extraction + 5% gross production (gas per Mcf), blended 9.6%. CO: severance up to 5% less the ad valorem credit "
     "(65.625% of gross income x mill levy in 2026) plus county ad valorem (~7% Weld) and state fees. TX: 4.6% oil / 7.5% gas severance "
     "plus ~2.5% ad valorem. WY: 6% severance + county ad valorem (statewide average ~6.4%) + 0.05% conservation"),
    (51, "Production taxes, after month 36", [0.096, 0.080, 0.075, 0.125], "0.0%", "Same as months 1-36 unless a state incentive expires"),
    (52, "Fixed LOE ($/well/month)", [8000, 6000, 7000, 8500], "#,##0", "8/8, per well"),
    (53, "Variable LOE ($/Mcf)", [0.25, 0.30, 0.30, 0.30], "0.00", ""),
    (54, "Variable LOE ($/bbl oil)", [4.00, 3.00, 3.50, 4.00], "0.00", ""),
    (55, "Water-oil ratio (bbl/bbl)", [2.4, 1.0, 3.0, 1.5], "0.0", "Delaware runs higher (4+), Midland 2-3; use offset water from Enverus when available"),
    (56, "Water disposal ($/bbl)", [1.00, 1.25, 0.85, 1.50], "0.00", ""),
    (57, "Plugging and abandonment, 8/8 ($)", [150000, 125000, 125000, 175000], "#,##0", ""),
]
if "Basins" in wb.sheetnames:
    del wb["Basins"]
B = wb.create_sheet("Basins", index=wb.sheetnames.index("Inputs") + 1)
B["B2"] = "Basin defaults for pricing, costs and taxes (8/8)"; B["B2"].font = Font(bold=True, size=14)
B["B3"] = ("Inputs C19 picks a column; each Inputs value uses it unless an override is typed in Inputs column E. "
           "ASSUMPTION: starting points from public sources (state tax law, operator disclosures, hub prices) as of Nov 2026, "
           "not a substitute for the operator's AFE, JIBs and contracts. Edit a column to change that basin's default for every future deal.")
B["B3"].alignment = Alignment(wrap_text=True, vertical="top"); B.merge_cells("B3:H3"); B.row_dimensions[3].height = 48
for j, h in enumerate(["Parameter", "Inputs cell"] + BASINS + ["Notes"]):
    c = B.cell(4, 2 + j, h); c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = Alignment(wrap_text=True)
ROW = {}
for i, (r, label, vals, fmt, note) in enumerate(PARAMS):
    rr = 5 + i; ROW[r] = rr
    B.cell(rr, 2, label); B.cell(rr, 3, f"C{r}")
    for j, v in enumerate(vals):
        c = B.cell(rr, 4 + j, v); c.font = BLUE; c.number_format = fmt
    B.cell(rr, 8, note).font = GREY; B.cell(rr, 8).alignment = Alignment(wrap_text=True, vertical="top")
for col, w in zip("BCDEFGH", (44, 10, 15, 15, 15, 17, 90)):
    B.column_dimensions[col].width = w
B.freeze_panes = "D5"

# ------------------------------------------------------------------ Inputs: deal setup block
hdr = [("B17", "Deal setup"), ("C17", "Value"), ("D17", "Basis")]
for a, v in hdr:
    I[a] = v; I[a].font = copy(I["B36"].font); I[a].fill = copy(I["B36"].fill); I[a].alignment = copy(I["B36"].alignment)
rows = [
    (18, "Deal type", "Working interest", "General",
     "Working interest = buy the WI and pay capex, LOE and the premium. Minerals = 0% WI: revenue at each well's mineral / royalty "
     "NRI (Wells column U), only production taxes come out, the purchase price (C20) is the only capital"),
    (19, "Basin (sets the price, cost and tax defaults below)", "Williston (ND)", "General",
     "Defaults live on the Basins tab; type a number in column E next to any value to override it for this deal"),
    (20, "Minerals: purchase price ($, total)", 0, '$#,##0;"($"#,##0);-',
     "Used only when Deal type = Minerals. Split across wells by mineral NRI x lateral (Wells column W)"),
    (21, "Minerals: closing month", None, "mmm-yy", "Used only for Minerals. Blank = the valuation month (C37)"),
]
for r, label, val, fmt, note in rows:
    I.cell(r, 2, label); I.cell(r, 3, val).font = BLUE; I.cell(r, 3).number_format = fmt
    I.cell(r, 4, note).font = GREY; I.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top")
I["C18"].fill = PatternFill("solid", fgColor="FFFFFF00"); I["C19"].fill = PatternFill("solid", fgColor="FFFFFF00")
dv1 = DataValidation(type="list", formula1='"Working interest,Minerals"', allow_blank=False); dv1.add("C18")
dv2 = DataValidation(type="list", formula1=f'"{",".join(BASINS)}"', allow_blank=False); dv2.add("C19")
I.add_data_validation(dv1); I.add_data_validation(dv2)

# ------------------------------------------------------------------ Inputs: values follow the basin unless overridden
for a in ("E36", "E49"):
    I[a] = "Override (blank = basin default)"; I[a].font = copy(I["B36"].font); I[a].fill = copy(I["B36"].fill)
    I[a].alignment = Alignment(wrap_text=True)
I.column_dimensions["E"].width = 16
for r, label, vals, fmt, note in PARAMS:
    old = I.cell(r, 3)
    I.cell(r, 5).number_format = old.number_format; I.cell(r, 5).font = BLUE
    I.cell(r, 3).value = (f'=IF(ISNUMBER($E{r}),$E{r},INDEX(Basins!$D${ROW[r]}:$G${ROW[r]},MATCH($C$19,Basins!$D$4:$G$4,0)))')
    I.cell(r, 3).font = copy(BLACK); I.cell(r, 3).fill = copy(old.fill)
    I.cell(r, 4).value = f'=IF(ISNUMBER($E{r}),"Override for this deal","Basin default ("&$C$19&"); see the Basins tab")'
    I.cell(r, 4).font = copy(GREY)

# ------------------------------------------------------------------ Wells: minerals NRI, active flag, price share
for col, h, w in (("U", "Minerals NRI (royalty / mineral decimal)", 16), ("V", "Active (1 = in the package)", 11),
                  ("W", "Minerals price share", 12)):
    W[f"{col}6"] = h; W[f"{col}6"].font = copy(W["T6"].font); W[f"{col}6"].fill = copy(W["T6"].fill)
    W[f"{col}6"].alignment = Alignment(wrap_text=True, vertical="top"); W.column_dimensions[col].width = w
for r in range(7, 17):
    W[f"U{r}"].font = BLUE; W[f"U{r}"].number_format = "0.000000"
    W[f"V{r}"] = f'=IF({MIN},IF(N(U{r})>0,1,0),IF(N(F{r})>0,1,0))'
    W[f"W{r}"] = f'=IF(V{r}=0,0,IFERROR(U{r}*E{r}/SUMPRODUCT($V$7:$V$16,$U$7:$U$16,$E$7:$E$16),0))'
    W[f"W{r}"].number_format = "0.0%"
W["B17"] = "=SUM(V7:V16)"
W["E17"] = "=SUMIF(V7:V16,1,E7:E16)"
W["N17"] = '=IF(B17=0,"",_xlfn.MINIFS(N7:N16,V7:V16,1))'
W["P17"] = '=IF(B17=0,"",_xlfn.MINIFS(P7:P16,V7:V16,1))'
W["F21"] = "=IF(B17=0,0,SUMPRODUCT((V7:V16=1)*Q7:Q16)/B17)"
W["U17"] = "=SUMPRODUCT(V7:V16,U7:U16)"; W["U17"].number_format = "0.000000"
W["B3"] = (str(W["B3"].value or "") + " Active wells (column V): WI > 0 for a working-interest deal, minerals NRI (column U) > 0 "
           "for a Minerals deal (Inputs C18).").strip()

# ------------------------------------------------------------------ CF W1..W10: minerals switch
for k in range(1, 11):
    ws = wb[f"CF W{k}"]; wr = 6 + k
    for r in range(6, 406):
        nri = f"IF({MIN},Wells!$U${wr},Wells!$H${wr})"
        ws[f"U{r}"] = f"={nri}*N{r}*T{r}"
        ws[f"V{r}"] = f"={nri}*P{r}*T{r}"
        for col in ("W", "X", "Y", "AA"):
            f = ws[f"{col}{r}"].value
            if isinstance(f, str) and f.startswith("=") and not f.startswith(f"=IF({MIN}"):
                ws[f"{col}{r}"] = f"=IF({MIN},0,{f[1:]})"
        f = ws[f"Z{r}"].value
        if isinstance(f, str) and not f.startswith(f"=IF({MIN}"):
            close = "MAX(N(Inputs!$C$21),$C$6)"
            ws[f"Z{r}"] = (f"=IF({MIN},IF(AND(YEAR(C{r})=YEAR({close}),MONTH(C{r})=MONTH({close})),"
                           f"Inputs!$C$20*Wells!$W${wr},0),{f[1:]})")
    ws["B3"] = (str(ws["B3"].value or "") + " Minerals (Inputs C18): revenue and tax at the mineral NRI (Wells U), "
                "no deducts, LOE, capex or P&A; the purchase price share is paid in the closing month.")

# ------------------------------------------------------------------ Well Summary
for r in range(5, 15):
    wr = r + 2
    for col in "IJKLMNOPQ":
        f = WS[f"{col}{r}"].value
        if isinstance(f, str):
            WS[f"{col}{r}"] = f.replace(f"Wells!$F${wr}>0", f"Wells!$V${wr}>0")
    WS[f"D{r}"] = f"=IF({MIN},0,Wells!F{wr})"
    WS[f"E{r}"] = f"=IF({MIN},Wells!U{wr},Wells!H{wr})"
    WS[f"G{r}"] = f"=IF({MIN},0,Wells!M{wr}*(1+Inputs!$C$59))"
    WS[f"H{r}"] = f"=IF({MIN},Inputs!$C$20*Wells!W{wr},Inputs!$C$10*Wells!M{wr})"
    WS[f"R{r}"] = f"=E{r}*Q{r}"
WS["H4"] = "Premium / purchase price ($)"

# ------------------------------------------------------------------ Summary
S["B8"] = "Working interest (average of wells); Minerals: average mineral NRI"
S["C8"] = f'=IF({MIN},IFERROR(AVERAGEIF(Wells!$U$7:$U$16,">0"),0),IFERROR(AVERAGEIF(Wells!$F$7:$F$16,">0"),0))'
S["C9"] = f"=IF({MIN},0,Wells!$M$17*(1+Inputs!$C$59))"
S["B10"] = "Premium paid ($); Minerals: purchase price"
S["C10"] = f"=IF({MIN},Inputs!$C$20,Inputs!$C$10*Wells!$M$17)"
S["F7"] = "Fund screen 20%; 0% = pay capex only. Not used for Minerals (price on Inputs C20)"
S["D72"] = f'=IF({MIN},"OK",IF(ABS(C72-Wells!$F$19)<1,"OK","CHECK"))'
S["B73"] = "Every active well has a lateral and a first-sales month"
S["D73"] = ('=IF(SUMPRODUCT((Wells!$V$7:$V$16=1)*((Wells!$E$7:$E$16=0)+(Wells!$P$7:$P$16=0)))=0,"OK","CHECK")')
S["D74"] = (f'=IF({MIN},"OK",IF(SUMPRODUCT((Wells!$V$7:$V$16=1)*((Wells!$P$7:$P$16<Wells!$O$7:$O$16)+'
            '(Wells!$O$7:$O$16<Wells!$N$7:$N$16)))=0,"OK","CHECK"))')
S["D75"] = (f'=IF(OR(C6<>"Strip",AND(NOT({MIN}),C7<>0.2),Inputs!$C$13<>1,Inputs!$C$11<>1,Inputs!$C$59<>0),"OK",'
            'IF(ABS(C75-$E$75)<500,"OK","CHECK"))')
S["B75"] = ("Package undiscounted cash ties to the Python engine (strip; 20% premium, or the purchase price for Minerals; "
            "base inputs); engine value in E75")
S["F75"] = "Written by write_model.py on each rebuild"
wb.save(out)
print("upgraded ->", out)
