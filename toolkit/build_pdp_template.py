"""Builds 'PDP Evaluation Template.xlsx' (producing-property evaluations) from scratch.

Usage: python build_pdp_template.py [<out.xlsx>]   (default: ../PDP Evaluation Template.xlsx)
Basins and the strip price deck are copied from 'Non-Op Evaluation Template.xlsx' so both models share one set of
basin defaults. Re-run after changing the layout here; deals are filled by write_pdp.py.

Layout (all live formulas; blue = input, grey = written by pdp_fit.py / write_pdp.py):
  Summary        price deck, value at the asking price, max price for the target IRR, PV10/15/20, reserves, metrics,
                 offset check for the package, checks, engine sensitivity grid, annual cash flow
  Inputs         deal setup, forecast settings, pricing / costs / taxes (basin defaults with overrides in column E)
  Basins         basin defaults (same table as the non-op template)
  Wells          one row per well (up to 30): interests, the fitted decline at the last reported month (qa, Da, b) for oil
                 and gas, override columns, the decline used, and results by well
  Offset Check   forecast vs the offsets' anchored P50 over the forecast months (12 / 24 / 36 / 60 / full overlap)
  Well Review    pick a well: history, fit, forecast and offset P25 / P50 / P75 on log charts
  Well Charts    one static chart per well from pdp_fit.py
  Package        monthly package cash flow from the effective date (480 months)
  Prices         strip (flat after the strip ends) and the flat decks
  M Oil / M Gas  gross monthly volumes by well from the effective date (before the economic limit)
  M RevAT        8/8 revenue after production taxes by well
  M On           1 while a well is economic on an 8/8 basis (operating cash flow > 0), 0 from the first month it is not
  M Net          our net cash flow by well (P&A in the month a well goes off)
  Fcst by Age    the forecast by month after the last report (1-120), for the offset check and charts
  History        monthly history by month on production (gross 8/8)
  Offset Data    the offsets' anchored P25 / P50 / P75 by forecast month (from pdp_fit.py)
"""
import os, sys, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.formatting.rule import FormulaRule

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
NW, NM, NK, NH = 30, 480, 120, 360          # wells, cash-flow months, forecast months checked, history months
W0 = 7                                       # first Wells row
M0, M1 = 25, 25 + NM - 1                     # matrix / Package data rows
K0, K1 = 8, 8 + NK - 1                       # Fcst by Age / Offset Data rows
H0, H1 = 7, 7 + NH - 1                       # History rows
MC0, MC1 = 3, 3 + NW - 1                     # matrix well columns C..AF
MCL0, MCL1 = L(MC0), L(MC1)
DECKS = ["Strip", "$65/$3.00", "$75/$3.50", "$85/$4.00"]

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
FIT_FILL = PatternFill("solid", fgColor="EDEDED")
INP_FILL = PatternFill("solid", fgColor="FFF2CC")
TITLE = Font(bold=True, size=14)
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))


def hdr(ws, row, col, text, width=None):
    c = ws.cell(row, col, text); c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = WRAP
    if width:
        ws.column_dimensions[L(col)].width = width
    return c


def sub(ws, row, text, c0=2, c1=5):
    for c in range(c0, c1 + 1):
        ws.cell(row, c).fill = SUB_FILL
    ws.cell(row, c0, text).font = BOLD


# ---------------------------------------------------------------- Wells columns
WCOLS = [  # key, header, width, number format, kind (in = input, fit = written by write_pdp, f = formula)
    ("num", "#", 4, "0", "f"), ("name", "Well", 30, None, "in"), ("api", "API", 13, "@", "in"), ("zone", "Zone (Enverus interval)", 14, None, "in"),
    ("lat", "Lateral (ft)", 9, "#,##0", "in"), ("first", "First production", 10, "mmm-yy", "in"), ("last", "Last reported month", 10, "mmm-yy", "in"),
    ("age", "Months on at last report", 8, "0", "f"), ("wi", "WI", 8, "0.000%", "in"), ("nri", "NRI (ours)", 8, "0.000%", "in"),
    ("lnri", "Lease NRI (8/8 economic limit)", 9, "0.0%", "in"), ("loe_ovr", "Fixed LOE override ($/mo 8/8)", 10, "$#,##0", "in"),
    ("mult", "Well forecast multiplier", 9, "0.00", "in"),
]
for s, S in (("o", "Oil"), ("g", "Gas")):
    u = "bbl/mo" if s == "o" else "Mcf/mo"
    WCOLS += [
        (f"{s}_qfit", f"{S}: anchor rate, fit ({u})", 10, "#,##0", "fit"), (f"{s}_qact", f"{S}: anchor rate, last 3 months ({u})", 10, "#,##0", "fit"),
        (f"{s}_da", f"{S}: Da fitted (nominal / yr)", 9, "0.0%", "fit"), (f"{s}_b", f"{S}: b fitted", 7, "0.00", "fit"),
        (f"{s}_r2", f"{S}: fit R2", 6, "0.00", "fit"), (f"{s}_note", f"{S}: fit notes", 34, None, "fit"),
        (f"{s}_q_ovr", f"{S}: anchor override ({u})", 10, "#,##0", "in"), (f"{s}_da_ovr", f"{S}: Da override", 9, "0.0%", "in"),
        (f"{s}_b_ovr", f"{S}: b override", 7, "0.00", "in"),
        (f"{s}_q", f"{S}: qa used ({u})", 10, "#,##0", "f"), (f"{s}_dau", f"{S}: Da used (nominal / yr)", 9, "0.0%", "f"),
        (f"{s}_bu", f"{S}: b used", 7, "0.00", "f"), (f"{s}_d12", f"{S}: decline over the next 12 months", 9, "0.0%", "f"),
        (f"{s}_tsw", f"{S}: years to terminal decline", 8, "0.0", "f"), (f"{s}_qsw", f"{S}: rate at terminal ({u})", 9, "#,##0", "f"),
    ]
WCOLS += [("k0", "Months from last report to effective date", 9, "0", "f"), ("active", "Active", 6, "0", "f"),
          ("wiu", "WI used (0 = Minerals)", 8, "0.000%", "f"), ("loe", "Fixed LOE used ($/mo 8/8)", 9, "$#,##0", "f"),
          ("dfac", "Deduct factor", 7, "0.00", "f")]
for s, S in (("o", "Oil"), ("g", "Gas")):
    WCOLS += [(f"{s}_start", f"{S} fit: start month", 7, "0", "fit"), (f"{s}_q0", f"{S} fit: q0", 9, "#,##0", "fit"),
              (f"{s}_D", f"{S} fit: D", 7, "0.000", "fit"), (f"{s}_bf", f"{S} fit: b", 7, "0.00", "fit"),
              (f"{s}_sh_age", f"{S} fit: level shift from month", 7, "0", "fit"), (f"{s}_sh", f"{S} fit: level factor", 7, "0.00", "fit")]
WCOLS += [("pv10", "Net PV10 ($)", 12, "$#,##0", "f"), ("pv_share", "Share of package PV10", 8, "0.0%", "f"),
          ("net_oil", "Net oil (bbl)", 10, "#,##0", "f"), ("net_gas", "Net gas (Mcf)", 11, "#,##0", "f"),
          ("life", "Economic life (months from effective date)", 9, "0", "f")]
WC = {k: 2 + i for i, (k, *_r) in enumerate(WCOLS)}       # key -> column number
WL = {k: L(v) for k, v in WC.items()}                     # key -> column letter


def wref(key, row, absolute=True):
    return f"Wells!${WL[key]}${row}" if absolute else f"Wells!{WL[key]}{row}"


# ---------------------------------------------------------------- Inputs cells
I = dict(name="C6", operator="C7", dtype="C8", basin="C9", eff="C10", price="C11", tirr="C12", anchor="C13", mult="C14",
         life_yr="C15", dmin="C16", econ="C17", tol="C18", deck="C19", flat_oil="C20", flat_gas="C21",
         odiff="C24", ogath="C25", gdiff="C26", shrink="C27", ngly="C28", nglp="C29", gp="C30",
         tax1="C33", tax2="C34", loe_fix="C35", loe_mcf="C36", loe_bbl="C37", wor="C38", swd="C39", pa="C40", roy="C41", disc="C42",
         dm="C45", life_m="C46")
IA = {k: f"Inputs!${v[0]}${v[1:]}" for k, v in I.items()}
BASIN_ROW = dict(odiff=5, ogath=6, gdiff=7, shrink=8, ngly=9, nglp=10, gp=11, tax1=12, tax2=13, loe_fix=14, loe_mcf=15, loe_bbl=16,
                 wor=17, swd=18, pa=19)


def build(out):
    src = openpyxl.load_workbook(os.path.join(ROOT, "Non-Op Evaluation Template.xlsx"))
    wb = openpyxl.Workbook()
    names = ["Summary", "Inputs", "Basins", "Wells", "Offset Check", "Well Review", "Well Charts", "Package", "Prices",
             "M Oil", "M Gas", "M RevAT", "M On", "M Net", "Fcst by Age", "History", "Offset Data"]
    wb.active.title = names[0]
    for n in names[1:]:
        wb.create_sheet(n)
    for n in names:
        wb[n].sheet_view.showGridLines = False if n in ("Summary", "Inputs", "Well Review", "Offset Check", "Well Charts") else True

    # ------------------------------------------------------------ Basins (copied)
    sb, b = src["Basins"], wb["Basins"]
    for row in sb.iter_rows(min_row=1, max_row=sb.max_row, max_col=9):
        for c in row:
            if c.value is not None:
                n = b.cell(c.row, c.column, c.value)
                n.number_format = c.number_format
                if c.row == 4:
                    n.font = HDR_FONT; n.fill = HDR_FILL; n.alignment = WRAP
                elif c.column >= 4 and c.column <= 8 and c.row >= 5:
                    n.font = BLUE
    b["B2"].font = TITLE
    b["B3"].alignment = WRAP; b.merge_cells("B3:I3"); b.row_dimensions[3].height = 60
    for col, w in zip("BCDEFGHI", [34, 9, 13, 11, 12, 15, 17, 90]):
        b.column_dimensions[col].width = w
    b["C4"] = "Inputs cell"
    for p, r in BASIN_ROW.items():
        b.cell(r, 3, I[p])
    for r in range(5, 20):
        b.cell(r, 9).alignment = WRAP

    # ------------------------------------------------------------ Inputs
    ws = wb["Inputs"]
    ws["B2"] = '="Inputs: "&C6'; ws["B2"].font = TITLE
    ws["B3"] = ("Blue = input. Grey on Wells = written by pdp_fit.py / write_pdp.py (re-run to refresh). Prices, costs and taxes use the "
                "basin column on the Basins tab unless a number is typed in column E. All volumes 8/8 gross unless marked net.")
    ws["B3"].alignment = WRAP; ws.merge_cells("B3:E3"); ws.row_dimensions[3].height = 30
    for col, w in zip("BCDE", [46, 16, 90, 14]):
        ws.column_dimensions[col].width = w
    sub(ws, 5, "Deal setup"); ws["C5"] = "Value"; ws["D5"] = "Basis"
    rows = [
        ("name", "Deal name", "Example PDP", "Shown on Summary"),
        ("operator", "Operator", "", ""),
        ("dtype", "Deal type", "Working interest", "Working interest = we pay our WI share of LOE, deducts and P&A and receive our NRI. Minerals = royalty / ORRI: "
         "our NRI (Wells column K) of revenue less production taxes only. Either way a well stops when it is uneconomic to the operator (8/8 basis)."),
        ("basin", "Basin (sets pricing, cost and tax defaults)", "DJ (CO)", "Defaults on the Basins tab; override any one in column E"),
        ("eff", "Effective date (first cash-flow month, price paid)", dt.datetime(2026, 12, 1), "Month 0 of the Package. Production between the last reported month (Wells G) and this date is forecast and belongs to the seller."),
        ("price", "Asking / offered price ($, total)", 0, "Paid in month 0. Summary also solves the max price for the target IRR (C12)"),
        ("tirr", "Target IRR", 0.20, "Fund screen 20%"),
        ("anchor", "Forecast anchor", "Fit", "Fit = the forecast starts from the fitted rate at the last reported month. Actual = from the last three good months' average carried to that month on the fitted shape. A per-well override (Wells) wins."),
        ("mult", "Package forecast multiplier", 1.0, "Scales every well's forecast (sensitivity); per-well multiplier on Wells"),
        ("life_yr", "Maximum life (years from the effective date)", 40, "Cash flows stop here even if still economic (480-month grid)"),
        ("dmin", "Terminal decline (effective, annual)", 0.07, "Hyperbolic declines switch to exponential at this decline (same 7% as the type curves)"),
        ("econ", "Economic limit on (1 = yes)", 1, "1 = each well stops the first month its 8/8 operating cash flow is not positive (lease NRI revenue less taxes, deducts and LOE). 0 = run to the maximum life"),
        ("tol", "Offset check tolerance (+/-)", 0.15, "Offset Check calls the forecast 'In line' when its cumulative is within this band of the offsets' anchored P50"),
        ("deck", "Price deck (1 = strip, 2 = flat); set on Summary C5", '=IF(Summary!$C$5="Strip",1,2)', "Strip on the Prices tab; flat decks follow the Summary dropdown"),
        ("flat_oil", "Flat oil price ($/bbl)", '=IF(Summary!$C$5="$75/$3.50",75,IF(Summary!$C$5="$85/$4.00",85,65))', "Follows the Summary dropdown"),
        ("flat_gas", "Flat gas price ($/MMBtu)", '=IF(Summary!$C$5="$75/$3.50",3.5,IF(Summary!$C$5="$85/$4.00",4,3))', "Follows the Summary dropdown"),
    ]
    for k, lab, v, note in rows:
        r = int(I[k][1:])
        ws.cell(r, 2, lab); c = ws.cell(r, 3, v); ws.cell(r, 4, note).alignment = WRAP
        if not (isinstance(v, str) and v.startswith("=")):
            c.font = BLUE
    ws[I["eff"]].number_format = "mmm-yy"; ws[I["price"]].number_format = "$#,##0"
    for k in ("tirr", "dmin", "tol"):
        ws[I[k]].number_format = "0.0%"
    ws[I["mult"]].number_format = "0.00"
    ws[I["flat_oil"]].number_format = '"$"0.00'; ws[I["flat_gas"]].number_format = '"$"0.00'
    sub(ws, 23, "Pricing (basin default unless an override is typed in column E)"); ws["C23"] = "Value"; ws["D23"] = "Basis"; ws["E23"] = "Override"
    sub(ws, 32, "Costs and taxes (8/8)"); ws["C32"] = "Value"; ws["D32"] = "Basis"; ws["E32"] = "Override"
    labels = dict(odiff="Oil differential to WTI ($/bbl)", ogath="Oil gathering / trucking ($/bbl)", gdiff="Residue gas differential to HH ($/MMBtu)",
                  shrink="Gas shrink (flaring + plant)", ngly="NGL yield (bbl/MMcf wellhead)", nglp="NGL price (% of WTI)",
                  gp="Gas gathering and processing ($/Mcf wellhead)", tax1="Production taxes, well months 1-36",
                  tax2="Production taxes, after well month 36", loe_fix="Fixed LOE ($/well/month)", loe_mcf="Variable LOE ($/Mcf)",
                  loe_bbl="Variable LOE ($/bbl oil)", wor="Water-oil ratio (bbl/bbl)", swd="Water disposal ($/bbl)", pa="Plugging and abandonment, 8/8 ($/well)")
    fmts = dict(odiff='"$"0.00', ogath='"$"0.00', gdiff='"$"0.00', shrink="0.0%", ngly="0", nglp="0.0%", gp='"$"0.00', tax1="0.0%", tax2="0.0%",
                loe_fix="$#,##0", loe_mcf='"$"0.00', loe_bbl='"$"0.00', wor="0.00", swd='"$"0.00', pa="$#,##0")
    for k, br in BASIN_ROW.items():
        r = int(I[k][1:])
        ws.cell(r, 2, labels[k])
        ws.cell(r, 3, f'=IF(ISNUMBER($E{r}),$E{r},INDEX(Basins!$D${br}:$H${br},MATCH({IA["basin"]},Basins!$D$4:$H$4,0)))').number_format = fmts[k]
        ws.cell(r, 4, f'=IF(ISNUMBER($E{r}),"Override for this deal","Basin default ("&{IA["basin"]}&"); see the Basins tab")')
        ws.cell(r, 5).font = BLUE; ws.cell(r, 5).fill = INP_FILL; ws.cell(r, 5).number_format = fmts[k]
    ws["B41"] = "Royalty owners bear deducts (1 = yes, 0 = no)"; ws["C41"] = 0; ws["C41"].font = BLUE
    ws["D41"] = "0 = the WI pays 100% of gathering / processing (conservative). Check the leases and check stubs."
    ws["B42"] = "Discount rate (PV)"; ws["C42"] = 0.10; ws["C42"].font = BLUE; ws["C42"].number_format = "0.0%"
    ws["D42"] = "Mid-month discounting from the effective date; PV15 and PV20 also shown"
    sub(ws, 44, "Derived")
    ws["B45"] = "Terminal decline, nominal (1/yr)"; ws["C45"] = f"=-LN(1-{I['dmin']})"; ws["C45"].number_format = "0.0000"
    ws["B46"] = "Maximum life (months)"; ws["C46"] = f"=MIN({NM},{I['life_yr']}*12)"
    dv = DataValidation(type="list", formula1='"Working interest,Minerals"'); ws.add_data_validation(dv); dv.add(I["dtype"])
    dv = DataValidation(type="list", formula1="=Basins!$D$4:$H$4"); ws.add_data_validation(dv); dv.add(I["basin"])
    dv = DataValidation(type="list", formula1='"Fit,Actual"'); ws.add_data_validation(dv); dv.add(I["anchor"])
    dv = DataValidation(type="list", formula1='"0,1"'); ws.add_data_validation(dv); dv.add(I["econ"])

    # ------------------------------------------------------------ Prices
    pr, sp = wb["Prices"], src["Prices"]
    pr["B2"] = "Price deck"; pr["B2"].font = TITLE
    pr["B3"] = (str(sp["B3"].value or "") + " Extended flat at the last strip month to 600 months. Paste a new strip into columns C and D "
                "(first row = the first strip month); the Package looks each month up by date.")
    pr["B3"].alignment = WRAP; pr.merge_cells("B3:H3"); pr.row_dimensions[3].height = 45
    for j, h in enumerate(["Month", "WTI strip", "HH strip", "WTI used", "HH used"]):
        hdr(pr, 5, 2 + j, h, 11)
    wti = [sp.cell(r, 4).value for r in range(6, 406)]; hh = [sp.cell(r, 5).value for r in range(6, 406)]
    d0 = dt.datetime(2026, 11, 1)
    for m in range(600):
        r = 6 + m
        y, mo = d0.year + (d0.month - 1 + m) // 12, (d0.month - 1 + m) % 12 + 1
        pr.cell(r, 2, dt.datetime(y, mo, 1)).number_format = "mmm-yy"
        pr.cell(r, 3, wti[min(m, 399)]).font = BLUE; pr.cell(r, 4, hh[min(m, 399)]).font = BLUE
        pr.cell(r, 5, f'=IF({IA["deck"]}=1,C{r},{IA["flat_oil"]})').number_format = '"$"0.00'
        pr.cell(r, 6, f'=IF({IA["deck"]}=1,D{r},{IA["flat_gas"]})').number_format = '"$"0.000'
    pr.freeze_panes = "C6"

    # ------------------------------------------------------------ Wells
    w = wb["Wells"]
    w["B2"] = "Wells: interests and the decline for each well"; w["B2"].font = TITLE
    w["B3"] = ("One row per well (up to 30). Grey = fitted by pdp_fit.py from the well's own production (write_pdp.py fills it). Blue = inputs; "
               "type in an override column to change a well's anchor rate, Da or b and everything downstream recalculates. The forecast is "
               "q = qa (1 + b Da t)^(-1/b), t in years after the last reported month, switching to the terminal decline (Inputs C16).")
    w["B3"].alignment = WRAP; w.merge_cells("B3:N3"); w.row_dimensions[3].height = 45
    groups = [("num", "lnri", "Well"), ("loe_ovr", "mult", "Inputs"), ("o_qfit", "o_note", "Oil: fitted"), ("o_q_ovr", "o_b_ovr", "Oil: overrides"),
              ("o_q", "o_qsw", "Oil: used"), ("g_qfit", "g_note", "Gas: fitted"), ("g_q_ovr", "g_b_ovr", "Gas: overrides"), ("g_q", "g_qsw", "Gas: used"),
              ("k0", "dfac", "Timing and interests used"), ("o_start", "g_sh", "Fit parameters (Well Review chart)"), ("pv10", "life", "Results")]
    for a, z, t in groups:
        w.merge_cells(start_row=5, end_row=5, start_column=WC[a], end_column=WC[z])
        c = w.cell(5, WC[a], t); c.font = BOLD; c.fill = SUB_FILL; c.alignment = Alignment(horizontal="center")
    for k, h, wd, f, kind in WCOLS:
        hdr(w, 6, WC[k], h, wd)
    w.row_dimensions[6].height = 58
    for i in range(1, NW + 1):
        r = W0 + i - 1
        f = {}
        f["num"] = i
        f["age"] = f'=IF({WL["name"]}{r}="",0,(YEAR({WL["last"]}{r})-YEAR({WL["first"]}{r}))*12+MONTH({WL["last"]}{r})-MONTH({WL["first"]}{r})+1)'
        f["k0"] = f'=IF({WL["name"]}{r}="",0,(YEAR({IA["eff"]})-YEAR({WL["last"]}{r}))*12+MONTH({IA["eff"]})-MONTH({WL["last"]}{r}))'
        f["active"] = f'=IF(AND({WL["name"]}{r}<>"",OR(N({WL["wi"]}{r})>0,N({WL["nri"]}{r})>0)),1,0)'
        f["wiu"] = f'=IF({IA["dtype"]}="Minerals",0,N({WL["wi"]}{r}))'
        f["loe"] = f'=IF(ISNUMBER({WL["loe_ovr"]}{r}),{WL["loe_ovr"]}{r},{IA["loe_fix"]})'
        f["dfac"] = f'=IF({IA["roy"]}=1,N({WL["lnri"]}{r}),1)'
        for s in ("o", "g"):
            q, da, bb = WL[f"{s}_q"], WL[f"{s}_dau"], WL[f"{s}_bu"]
            f[f"{s}_q"] = (f'=IF(ISNUMBER({WL[s + "_q_ovr"]}{r}),{WL[s + "_q_ovr"]}{r},IF({IA["anchor"]}="Actual",N({WL[s + "_qact"]}{r}),'
                           f'N({WL[s + "_qfit"]}{r})))')
            f[f"{s}_dau"] = f'=IF(ISNUMBER({WL[s + "_da_ovr"]}{r}),{WL[s + "_da_ovr"]}{r},N({WL[s + "_da"]}{r}))'
            f[f"{s}_bu"] = f'=MAX(0.001,IF(ISNUMBER({WL[s + "_b_ovr"]}{r}),{WL[s + "_b_ovr"]}{r},N({WL[s + "_b"]}{r})))'
            f[f"{s}_tsw"] = f'=IF({da}{r}<={IA["dm"]},0,({da}{r}/{IA["dm"]}-1)/({bb}{r}*{da}{r}))'
            f[f"{s}_qsw"] = f'={q}{r}*(1+{bb}{r}*{da}{r}*{WL[s + "_tsw"]}{r})^(-1/{bb}{r})'
            ts, qs = WL[f"{s}_tsw"], WL[f"{s}_qsw"]
            f[f"{s}_d12"] = (f'=IF({q}{r}<=0,"",1-IF(1<={ts}{r},{q}{r}*(1+{bb}{r}*{da}{r})^(-1/{bb}{r}),{qs}{r}*EXP(-{IA["dm"]}*(1-{ts}{r})))/{q}{r})')
        col_net = lambda m: f"INDEX('{m}'!${MCL0}${M0}:${MCL1}${M1},0,{WL['num']}{r})"
        f["pv10"] = f"=SUMPRODUCT({col_net('M Net')},Package!$V${M0}:$V${M1})"
        f["pv_share"] = f'=IF(Summary!$C$12=0,0,{WL["pv10"]}{r}/Summary!$C$12)'
        f["net_oil"] = f"=SUMPRODUCT({col_net('M Oil')},{col_net('M On')})*N({WL['nri']}{r})"
        f["net_gas"] = f"=SUMPRODUCT({col_net('M Gas')},{col_net('M On')})*N({WL['nri']}{r})"
        f["life"] = f"=SUM({col_net('M On')})"
        for k, h, wd, fmt, kind in WCOLS:
            c = w.cell(r, WC[k])
            if k in f:
                c.value = f[k]
            if fmt:
                c.number_format = fmt
            if kind == "in":
                c.font = BLUE
            elif kind == "fit":
                c.fill = FIT_FILL
        w.cell(r, WC["mult"], 1.0)
        w.cell(r, WC["lnri"], 0.80)
    w.freeze_panes = w.cell(W0, WC["wi"])
    tr = W0 + NW
    w.cell(tr, WC["name"], "Package").font = BOLD
    for k in ("pv10", "net_oil", "net_gas"):
        c = w.cell(tr, WC[k], f"=SUM({WL[k]}{W0}:{WL[k]}{tr - 1})"); c.font = BOLD; c.number_format = {x[0]: x[3] for x in WCOLS}[k]
    w.cell(tr, WC["life"], f"=MAX({WL['life']}{W0}:{WL['life']}{tr - 1})").font = BOLD
    w.conditional_formatting.add(f"{WL['o_note']}{W0}:{WL['o_note']}{tr - 1}",
                                 FormulaRule(formula=[f'LEN({WL["o_note"]}{W0})>0'], fill=PatternFill("solid", fgColor="FCE4D6")))
    w.conditional_formatting.add(f"{WL['g_note']}{W0}:{WL['g_note']}{tr - 1}",
                                 FormulaRule(formula=[f'LEN({WL["g_note"]}{W0})>0'], fill=PatternFill("solid", fgColor="FCE4D6")))

    # ------------------------------------------------------------ matrix sheets
    HELP = [(5, "Well #"), (6, "Active"), (7, "qa"), (8, "Da"), (9, "b"), (10, "Years to terminal"), (11, "Rate at terminal"),
            (12, "Months last report -> effective"), (13, "Well multiplier"), (14, "Months on at last report"), (15, "NRI"), (16, "WI used"),
            (17, "WI x deduct factor"), (18, "WI x fixed LOE"), (19, "Lease NRI"), (20, "Deduct factor"), (21, "Fixed LOE 8/8")]
    notes = {"M Oil": "Gross oil (bbl/month) by well from the effective date, before the economic limit (M On). Month k after the last report: "
                      "q = qa (1 + b Da k/12)^(-1/b), exponential at the terminal decline after 'Years to terminal'; actual history if the effective date is earlier.",
             "M Gas": "Gross gas (Mcf/month) by well from the effective date, before the economic limit (M On); same formula as M Oil on the gas decline.",
             "M RevAT": "8/8 revenue after production taxes ($) by well: oil, residue gas and NGL at the Package prices less taxes at the well's age (months 1-36 / after).",
             "M On": "1 while the well's 8/8 operating cash flow (lease NRI x revenue after tax - deducts - LOE) is positive; 0 from the first month it is not (economic limit).",
             "M Net": "Our net cash flow by well ($): NRI x revenue after tax - WI x (deducts + LOE) while on, less WI x P&A in the month it goes off. Minerals: WI used = 0."}
    for mname in ("M Oil", "M Gas", "M RevAT", "M On", "M Net"):
        m = wb[mname]
        m["A1"] = mname; m["A1"].font = TITLE
        m["A2"] = notes[mname]
        m["A24"] = "Month"; m["B24"] = "Date"; m["A24"].font = BOLD; m["B24"].font = BOLD
        m.column_dimensions["A"].width = 7; m.column_dimensions["B"].width = 9
        s = "g" if mname == "M Gas" else "o"
        for rr, lab in HELP:
            m.cell(rr, 2, lab).font = Font(italic=True, color="808080")
        for i in range(1, NW + 1):
            c, wr = MC0 + i - 1, W0 + i - 1
            X = L(c)
            vals = {5: i, 6: f"={wref('active', wr)}", 7: f"={wref(s + '_q', wr)}", 8: f"={wref(s + '_dau', wr)}", 9: f"={wref(s + '_bu', wr)}",
                    10: f"={wref(s + '_tsw', wr)}", 11: f"={wref(s + '_qsw', wr)}", 12: f"={wref('k0', wr)}", 13: f"=N({wref('mult', wr)})",
                    14: f"={wref('age', wr)}", 15: f"=N({wref('nri', wr)})", 16: f"={wref('wiu', wr)}", 17: f"={wref('wiu', wr)}*{wref('dfac', wr)}",
                    18: f"={wref('wiu', wr)}*{wref('loe', wr)}", 19: f"=N({wref('lnri', wr)})", 20: f"={wref('dfac', wr)}", 21: f"={wref('loe', wr)}"}
            for rr, v in vals.items():
                m.cell(rr, c, v).font = Font(color="808080")
            m.cell(24, c, f"={wref('name', wr)}").font = BOLD
            m.column_dimensions[X].width = 11
        for j in range(NM):
            r = M0 + j
            m.cell(r, 1, j)
            m.cell(r, 2, f"=EDATE({IA['eff']},A{r})").number_format = "mmm-yy"
        m.freeze_panes = f"C{M0}"
    hist = f"History!$C${H0}:${L(2 + 2 * NW)}${H1}"
    for j in range(NM):
        r = M0 + j
        for i in range(1, NW + 1):
            X = L(MC0 + i - 1)
            for mname, hc in (("M Oil", f"2*{X}$5-1"), ("M Gas", f"2*{X}$5")):
                K = f"({X}$12+$A{r})"
                wb[mname].cell(r, MC0 + i - 1,
                               f"=IF(OR({X}$6=0,$A{r}>={IA['life_m']}),0,{IA['mult']}*{X}$13*IF({K}<1,IFERROR(INDEX({hist},MAX(1,{X}$14+{K}),{hc}),0),"
                               f"IF({K}/12<={X}$10,{X}$7*(1+{X}$9*{X}$8*{K}/12)^(-1/{X}$9),{X}$11*EXP(-{IA['dm']}*({K}/12-{X}$10)))))")
            o, g = f"'M Oil'!{X}{r}", f"'M Gas'!{X}{r}"
            wb["M RevAT"].cell(r, MC0 + i - 1,
                               f"=IF({X}$6=0,0,({o}*(Package!$C{r}+{IA['odiff']})+{g}*(1-{IA['shrink']})*(Package!$D{r}+{IA['gdiff']})"
                               f"+{g}/1000*{IA['ngly']}*Package!$C{r}*{IA['nglp']})*(1-IF({X}$14+{X}$12+$A{r}>36,{IA['tax2']},{IA['tax1']})))")
            ded = f"({g}*{IA['gp']}+{o}*{IA['ogath']})"
            loe_var = f"({g}*{IA['loe_mcf']}+{o}*({IA['loe_bbl']}+{IA['wor']}*{IA['swd']}))"
            op8 = f"({X}$19*'M RevAT'!{X}{r}-{X}$20*{ded}-({X}$21+{loe_var}))"
            prev = "1" if j == 0 else f"{X}{r - 1}"
            wb["M On"].cell(r, MC0 + i - 1,
                            f"=IF({X}$6=0,0,IF({IA['econ']}=0,IF({o}+{g}>0,1,0),IF(AND({prev}=1,{op8}>0),1,0)))")
            ev = f"({X}$6-'M On'!{X}{r})" if j == 0 else f"('M On'!{X}{r - 1}-'M On'!{X}{r})"
            wb["M Net"].cell(r, MC0 + i - 1,
                             f"='M On'!{X}{r}*({X}$15*'M RevAT'!{X}{r}-{X}$17*{ded}-{X}$16*({X}$21+{loe_var}))-{X}$16*{IA['pa']}*{ev}")

    # ------------------------------------------------------------ Package
    p = wb["Package"]
    p["A1"] = "Package: monthly cash flow from the effective date (net to our interest)"; p["A1"].font = TITLE
    p["A2"] = ("Sums the well matrices. Volumes net = gross x NRI while the well is economic. Operating cash flow = revenue - taxes - deducts - LOE - P&A. "
               "Column U ties to the sum of M Net (must be 0).")
    heads = ["Month", "Date", "WTI used", "HH used", "Gross oil (bbl)", "Gross gas (Mcf)", "Net oil (bbl)", "Net gas (Mcf, wellhead)", "Net NGL (bbl)",
             "Oil revenue", "Residue gas revenue", "NGL revenue", "Total revenue", "Production taxes", "Deducts (G&P, oil gathering)", "LOE", "P&A",
             "Operating cash flow", "Purchase price", "Net cash flow", "Tie to M Net", "Discount factor (PV rate)", "Cumulative net cash flow", "Year"]
    for j, h in enumerate(heads):
        hdr(p, 24, 1 + j, h, 12 if j > 1 else 8)
    p.row_dimensions[24].height = 45
    rng = lambda m, r: f"'{m}'!${MCL0}{r}:${MCL1}{r}"
    hrow = lambda m, rr: f"'{m}'!${MCL0}${rr}:${MCL1}${rr}"
    for j in range(NM):
        r = M0 + j
        cells = {
            1: j, 2: f"=EDATE({IA['eff']},A{r})",
            3: f"=IFERROR(INDEX(Prices!$E$6:$E$605,MATCH($B{r},Prices!$B$6:$B$605,1)),Prices!$E$6)",
            4: f"=IFERROR(INDEX(Prices!$F$6:$F$605,MATCH($B{r},Prices!$B$6:$B$605,1)),Prices!$F$6)",
            5: f"=SUMPRODUCT({rng('M Oil', r)},{rng('M On', r)})", 6: f"=SUMPRODUCT({rng('M Gas', r)},{rng('M On', r)})",
            7: f"=SUMPRODUCT({rng('M Oil', r)},{rng('M On', r)},{hrow('M On', 15)})", 8: f"=SUMPRODUCT({rng('M Gas', r)},{rng('M On', r)},{hrow('M On', 15)})",
            9: f"=H{r}/1000*{IA['ngly']}", 10: f"=G{r}*(C{r}+{IA['odiff']})", 11: f"=H{r}*(1-{IA['shrink']})*(D{r}+{IA['gdiff']})",
            12: f"=I{r}*C{r}*{IA['nglp']}", 13: f"=J{r}+K{r}+L{r}",
            14: f"=M{r}-SUMPRODUCT({rng('M RevAT', r)},{rng('M On', r)},{hrow('M On', 15)})",
            15: f"={IA['gp']}*SUMPRODUCT({rng('M Gas', r)},{rng('M On', r)},{hrow('M On', 17)})+{IA['ogath']}*SUMPRODUCT({rng('M Oil', r)},{rng('M On', r)},{hrow('M On', 17)})",
            16: (f"=SUMPRODUCT({rng('M On', r)},{hrow('M On', 18)})+{IA['loe_mcf']}*SUMPRODUCT({rng('M Gas', r)},{rng('M On', r)},{hrow('M On', 16)})"
                 f"+({IA['loe_bbl']}+{IA['wor']}*{IA['swd']})*SUMPRODUCT({rng('M Oil', r)},{rng('M On', r)},{hrow('M On', 16)})"),
            17: (f"={IA['pa']}*SUMPRODUCT({hrow('M On', 6)}-{rng('M On', r)},{hrow('M On', 16)})" if j == 0 else
                 f"={IA['pa']}*SUMPRODUCT({rng('M On', r - 1)}-{rng('M On', r)},{hrow('M On', 16)})"),
            18: f"=M{r}-N{r}-O{r}-P{r}-Q{r}", 19: f"=IF(A{r}=0,-{IA['price']},0)", 20: f"=R{r}+S{r}",
            21: f"=SUM({rng('M Net', r)})-R{r}", 22: f"=1/(1+{IA['disc']})^((A{r}+0.5)/12)",
            23: f"=T{r}" if j == 0 else f"=W{r - 1}+T{r}", 24: f"=INT(A{r}/12)+1"}
        for c, v in cells.items():
            x = p.cell(r, c, v)
            x.number_format = {2: "mmm-yy", 3: '"$"0.00', 4: '"$"0.000', 22: "0.0000", 24: "0", 1: "0"}.get(c, "#,##0")
    p.freeze_panes = f"C{M0}"
    tot = 22
    p.cell(tot, 2, "Total").font = BOLD
    for c in list(range(5, 22)) + [23]:
        if c in (22, 23):
            continue
        x = p.cell(tot, c, f"=SUM({L(c)}{M0}:{L(c)}{M1})"); x.font = BOLD; x.number_format = "#,##0"

    # ------------------------------------------------------------ Fcst by Age
    fa = wb["Fcst by Age"]
    fa["A1"] = "Forecast by month after the last report (gross 8/8, before the economic limit)"; fa["A1"].font = TITLE
    fa["A2"] = "Month k = k months after the well's last reported month (its age is the Wells 'months on' + k). Feeds Offset Check and Well Review."
    hdr(fa, 7, 2, "k", 5)
    for i in range(1, NW + 1):
        wr = W0 + i - 1
        hdr(fa, 6, 2 * i + 1, f"={wref('name', wr)}", 11); fa.merge_cells(start_row=6, end_row=6, start_column=2 * i + 1, end_column=2 * i + 2)
        hdr(fa, 7, 2 * i + 1, "Oil", 11); hdr(fa, 7, 2 * i + 2, "Gas", 11)
    for k in range(1, NK + 1):
        r = K0 + k - 1
        fa.cell(r, 2, k)
        for i in range(1, NW + 1):
            wr = W0 + i - 1
            for s, c in (("o", 2 * i + 1), ("g", 2 * i + 2)):
                q, da, bb, ts, qs = (wref(s + x, wr) for x in ("_q", "_dau", "_bu", "_tsw", "_qsw"))
                fa.cell(r, c, f'=IF({wref("active", wr)}=0,"",{IA["mult"]}*N({wref("mult", wr)})*IF($B{r}/12<={ts},{q}*(1+{bb}*{da}*$B{r}/12)^(-1/{bb}),'
                              f'{qs}*EXP(-{IA["dm"]}*($B{r}/12-{ts}))))').number_format = "#,##0"
    fa.freeze_panes = f"C{K0}"

    # ------------------------------------------------------------ History, Offset Data (values from write_pdp.py)
    h = wb["History"]
    h["A1"] = "Monthly history by month on production (gross 8/8, Enverus); written by write_pdp.py"; h["A1"].font = TITLE
    hdr(h, 6, 2, "Month on", 7)
    for i in range(1, NW + 1):
        wr = W0 + i - 1
        hdr(h, 5, 2 * i + 1, f"={wref('name', wr)}", 10); h.merge_cells(start_row=5, end_row=5, start_column=2 * i + 1, end_column=2 * i + 2)
        hdr(h, 6, 2 * i + 1, "Oil", 10); hdr(h, 6, 2 * i + 2, "Gas", 10)
    for a in range(1, NH + 1):
        h.cell(H0 + a - 1, 2, a)
    h.freeze_panes = f"C{H0}"
    od = wb["Offset Data"]
    od["A1"] = "Offsets' view of each well's forecast months (anchored); written by write_pdp.py from pdp_fit.json"; od["A1"].font = TITLE
    od["A2"] = ("Each offset in the well's zone group is scaled so its average over the well's last three reported months of age equals the well's own, "
                "then its next months are read off. n = offsets reporting that month; percentiles shown where 3+ report.")
    hdr(od, 7, 2, "k", 5)
    for i in range(1, NW + 1):
        c0 = 3 + (i - 1) * 10
        hdr(od, 6, c0, f"={wref('name', W0 + i - 1)}"); od.merge_cells(start_row=6, end_row=6, start_column=c0, end_column=c0 + 9)
        for j, t in enumerate(["Oil n", "Oil P25", "Oil P50", "Oil P75", "Oil (spare)", "Gas n", "Gas P25", "Gas P50", "Gas P75", "Gas (spare)"]):
            hdr(od, 7, c0 + j, t, 8)
    for k in range(1, NK + 1):
        od.cell(K0 + k - 1, 2, k)
    od.freeze_panes = f"C{K0}"
    OD = f"'Offset Data'!$C${K0}:${L(2 + 10 * NW)}${K1}"
    FA = f"'Fcst by Age'!$C${K0}:${L(2 + 2 * NW)}${K1}"

    # ------------------------------------------------------------ Offset Check
    oc = wb["Offset Check"]
    oc["B2"] = "Offset check: our forecast vs what the offsets did over the same months of life"; oc["B2"].font = TITLE
    oc["B3"] = ("For each well, every offset in its zone group with production at the well's current age is scaled to the well's own last three months "
                "(anchored), and its following months are read off (Offset Data). Ratio = our forecast's cumulative / the offsets' P50 cumulative over "
                "the same months, where 3+ offsets report every month. 'In line' = within the tolerance on Inputs C18. Level = the well's last three months "
                "per 10,000 ft vs the offsets at the same age (how strong the well is), from pdp_fit.py.")
    oc["B3"].alignment = WRAP; oc.merge_cells("B3:Q3"); oc.row_dimensions[3].height = 58
    base = ["#", "Well", "Zone", "Months on"]
    per = ["Offsets at its age", "Months covered (3+ offsets)", "Level vs offsets (per 10k ft)", "Fcst cum 12 mo", "Offset P50 cum 12 mo",
           "Fcst cum 24 mo", "Offset P50 cum 24 mo", "Fcst cum (overlap)", "Offset P50 cum (overlap)", "Ratio 12 mo", "Ratio 24 mo", "Ratio 36 mo",
           "Ratio 60 mo", "Ratio full overlap", "Verdict"]
    for j, t in enumerate(base):
        hdr(oc, 6, 2 + j, t, [4, 28, 13, 7][j])
    OCC = {}
    for s, S, c0 in (("o", "Oil", 6), ("g", "Gas", 6 + len(per))):
        oc.merge_cells(start_row=5, end_row=5, start_column=c0, end_column=c0 + len(per) - 1)
        x = oc.cell(5, c0, S); x.font = BOLD; x.fill = SUB_FILL; x.alignment = Alignment(horizontal="center")
        for j, t in enumerate(per):
            hdr(oc, 6, c0 + j, t, 30 if t == "Verdict" else 10)
        OCC[s] = {t: L(c0 + j) for j, t in enumerate(per)}
    oc.row_dimensions[6].height = 45
    tol = IA["tol"]
    verdict = lambda x: f'IF(ABS({x}-1)<={tol},"In line",IF({x}>1,"Above offsets "&TEXT({x}-1,"+0%;-0%"),"Below offsets "&TEXT({x}-1,"+0%;-0%")))'
    for i in range(1, NW + 1):
        r, wr = W0 + i - 1, W0 + i - 1
        oc.cell(r, 2, i)
        oc.cell(r, 3, f'=IF({wref("name", wr)}="","",{wref("name", wr)})')
        oc.cell(r, 4, f'=IF({wref("name", wr)}="","",{wref("zone", wr)})')
        oc.cell(r, 5, f'=IF({wref("name", wr)}="","",{wref("age", wr)})')
        for s, fc, pc in (("o", f"2*$B{r}-1", f"($B{r}-1)*10+3"), ("g", f"2*$B{r}", f"($B{r}-1)*10+8")):
            C = OCC[s]
            cum = lambda R, col, hz: f"SUM(INDEX({R},1,{col}):INDEX({R},{hz},{col}))"
            cov = f"{C['Months covered (3+ offsets)']}{r}"
            cells = {
                "Fcst cum 12 mo": f'=IF(AND($C{r}<>"",N({cov})>=12),{cum(FA, fc, 12)},"")',
                "Offset P50 cum 12 mo": f'=IF(AND($C{r}<>"",N({cov})>=12),{cum(OD, pc, 12)},"")',
                "Fcst cum 24 mo": f'=IF(AND($C{r}<>"",N({cov})>=24),{cum(FA, fc, 24)},"")',
                "Offset P50 cum 24 mo": f'=IF(AND($C{r}<>"",N({cov})>=24),{cum(OD, pc, 24)},"")',
                "Fcst cum (overlap)": f'=IF(AND($C{r}<>"",N({cov})>=1),{cum(FA, fc, cov)},"")',
                "Offset P50 cum (overlap)": f'=IF(AND($C{r}<>"",N({cov})>=1),{cum(OD, pc, cov)},"")',
                "Ratio 12 mo": f'=IF($C{r}="","",IF(ISNUMBER({C["Fcst cum 12 mo"]}{r}),{C["Fcst cum 12 mo"]}{r}/{C["Offset P50 cum 12 mo"]}{r},"n/a"))',
                "Ratio 24 mo": f'=IF($C{r}="","",IF(ISNUMBER({C["Fcst cum 24 mo"]}{r}),{C["Fcst cum 24 mo"]}{r}/{C["Offset P50 cum 24 mo"]}{r},"n/a"))',
                "Ratio 36 mo": f'=IF($C{r}="","",IF(N({cov})>=36,{cum(FA, fc, 36)}/{cum(OD, pc, 36)},"n/a"))',
                "Ratio 60 mo": f'=IF($C{r}="","",IF(N({cov})>=60,{cum(FA, fc, 60)}/{cum(OD, pc, 60)},"n/a"))',
                "Ratio full overlap": f'=IF($C{r}="","",IF(ISNUMBER({C["Fcst cum (overlap)"]}{r}),{C["Fcst cum (overlap)"]}{r}/{C["Offset P50 cum (overlap)"]}{r},"n/a"))',
            }
            r24, rf = f"{C['Ratio 24 mo']}{r}", f"{C['Ratio full overlap']}{r}"
            cells["Verdict"] = (f'=IF($C{r}="","",IF(ISNUMBER({r24}),{verdict(r24)}&" (24 mo)",IF(ISNUMBER({rf}),{verdict(rf)}&" (first "&{cov}&" mo)",'
                                f'"No offsets with this much history")))')
            for t, v in cells.items():
                x = oc.cell(r, openpyxl.utils.column_index_from_string(C[t]), v)
                x.number_format = "0%" if t.startswith("Ratio") else "#,##0"
            for t in ("Offsets at its age", "Months covered (3+ offsets)", "Level vs offsets (per 10k ft)"):
                x = oc.cell(r, openpyxl.utils.column_index_from_string(C[t])); x.fill = FIT_FILL
                x.number_format = "0%" if t.startswith("Level") else "0"
    tr = W0 + NW
    oc.cell(tr, 3, "Package (wells with coverage)").font = BOLD
    for s in ("o", "g"):
        C = OCC[s]
        for t in ("Fcst cum 12 mo", "Offset P50 cum 12 mo", "Fcst cum 24 mo", "Offset P50 cum 24 mo", "Fcst cum (overlap)", "Offset P50 cum (overlap)"):
            x = oc.cell(tr, openpyxl.utils.column_index_from_string(C[t]), f"=SUM({C[t]}{W0}:{C[t]}{tr - 1})"); x.font = BOLD; x.number_format = "#,##0"
        for t, a, b_ in (("Ratio 12 mo", "Fcst cum 12 mo", "Offset P50 cum 12 mo"), ("Ratio 24 mo", "Fcst cum 24 mo", "Offset P50 cum 24 mo"),
                         ("Ratio full overlap", "Fcst cum (overlap)", "Offset P50 cum (overlap)")):
            x = oc.cell(tr, openpyxl.utils.column_index_from_string(C[t]), f'=IF(N({C[b_]}{tr})>0,{C[a]}{tr}/{C[b_]}{tr},"n/a")')
            x.font = BOLD; x.number_format = "0%"
        r24, rf = f"{C['Ratio 24 mo']}{tr}", f"{C['Ratio full overlap']}{tr}"
        x = oc.cell(tr, openpyxl.utils.column_index_from_string(C["Verdict"]),
                    f'=IF(ISNUMBER({r24}),{verdict(r24)}&" (24 mo)",IF(ISNUMBER({rf}),{verdict(rf)}&" (overlap)","n/a"))'); x.font = BOLD
        for t in ("Verdict",):
            rngv = f"{C[t]}{W0}:{C[t]}{tr}"
            oc.conditional_formatting.add(rngv, FormulaRule(formula=[f'LEFT({C[t]}{W0},5)="Above"'], fill=PatternFill("solid", fgColor="FCE4D6")))
            oc.conditional_formatting.add(rngv, FormulaRule(formula=[f'LEFT({C[t]}{W0},5)="Below"'], fill=PatternFill("solid", fgColor="DDEBF7")))
            oc.conditional_formatting.add(rngv, FormulaRule(formula=[f'LEFT({C[t]}{W0},7)="In line"'], fill=PatternFill("solid", fgColor="E2EFDA")))
    oc.freeze_panes = "F7"

    # ------------------------------------------------------------ Well Review
    wv = wb["Well Review"]
    wv["B2"] = "Well review: pick a well"; wv["B2"].font = TITLE
    wv["B3"] = "Well"; wv["C3"] = f"={wref('name', W0)}"; wv["C3"].font = Font(bold=True, color="0000FF", size=12); wv["C3"].fill = INP_FILL
    wv["D3"] = "Pick from the dropdown; the table and charts follow."
    dv = DataValidation(type="list", formula1=f"=Wells!${WL['name']}${W0}:${WL['name']}${W0 + NW - 1}"); wv.add_data_validation(dv); dv.add("C3")
    wv["B4"] = "Well #"; wv["C4"] = f"=MATCH(C3,Wells!${WL['name']}${W0}:${WL['name']}${W0 + NW - 1},0)"
    wix = lambda key: f"INDEX(Wells!${WL[key]}${W0}:${WL[key]}${W0 + NW - 1},$C$4)"
    info = [("Zone", wix("zone"), None), ("Lateral (ft)", wix("lat"), "#,##0"), ("Months on at last report", wix("age"), "0"),
            ("Oil: qa / Da / b used", f'=TEXT({wix("o_q")},"#,##0")&" bbl/mo, "&TEXT({wix("o_dau")},"0%")&", b "&TEXT({wix("o_bu")},"0.00")', None),
            ("Gas: qa / Da / b used", f'=TEXT({wix("g_q")},"#,##0")&" Mcf/mo, "&TEXT({wix("g_dau")},"0%")&", b "&TEXT({wix("g_bu")},"0.00")', None),
            ("Oil vs offsets", f"=INDEX('Offset Check'!${OCC['o']['Verdict']}${W0}:${OCC['o']['Verdict']}${W0 + NW - 1},$C$4)", None),
            ("Gas vs offsets", f"=INDEX('Offset Check'!${OCC['g']['Verdict']}${W0}:${OCC['g']['Verdict']}${W0 + NW - 1},$C$4)", None)]
    for j, (lab, f_, fmt) in enumerate(info):
        wv.cell(5 + j % 4, 2 if j < 4 else 5, lab)
        x = wv.cell(5 + j % 4, 3 if j < 4 else 6, f_ if str(f_).startswith("=") else "=" + f_)
        if fmt:
            x.number_format = fmt
    # fit helper cells (row 10): oil start, q0, D, b, tsw, qsw, shift age, shift factor; gas likewise
    HC = 27                                   # fit helpers in hidden columns AA:AH
    for jj, (s, rr) in enumerate((("o", 4), ("g", 5))):
        keys = ["_start", "_q0", "_D", "_bf", "_sh_age", "_sh"]
        for k, key in enumerate(keys):
            wv.cell(rr, HC + k, f"=N({wix(s + key)})").font = Font(color="808080")
        st, q0, D, bb, sa, sf = (f"${L(HC + k)}${rr}" for k in range(6))
        wv.cell(rr, HC + 6, f"=IF({D}<={IA['dm']},0,({D}/{IA['dm']}-1)/(MAX({bb},0.001)*{D}))").font = Font(color="808080")
        wv.cell(rr, HC + 7, f"={q0}*(1+MAX({bb},0.001)*{D}*${L(HC + 6)}${rr})^(-1/MAX({bb},0.001))").font = Font(color="808080")
    for k in range(8):
        wv.column_dimensions[L(HC + k)].hidden = True
    T0 = 12
    heads = ["Month on", "Oil actual", "Oil fit", "Oil forecast", "Oil offsets P25", "Oil offsets P50", "Oil offsets P75",
             "Gas actual", "Gas fit", "Gas forecast", "Gas offsets P25", "Gas offsets P50", "Gas offsets P75"]
    for j, t in enumerate(heads):
        hdr(wv, T0 - 1, 2 + j, t, 11 if j else 8)
    wv.column_dimensions["B"].width = 26; wv.column_dimensions["C"].width = 30; wv.column_dimensions["E"].width = 16
    A = wix("age")
    for n in range(180):
        r = T0 + n
        wv.cell(r, 2, f"=MAX(1,{A}-59)+{n}")
        for s, c0, hc, oc_ in (("o", 3, "2*$C$4-1", "($C$4-1)*10"), ("g", 9, "2*$C$4", "($C$4-1)*10+5")):
            rr = 4 if s == "o" else 5
            st, q0, D, bb, sa, sf, ts, qs = (f"${L(HC + k)}${rr}" for k in range(8))
            act = f"INDEX({hist},$B{r},{hc})"
            wv.cell(r, c0, f"=IF(OR($B{r}<1,$B{r}>{A}),NA(),IF(N({act})>0,{act},NA()))")
            t = f"(($B{r}-{st})/12)"
            wv.cell(r, c0 + 1, f"=IF(OR($B{r}<{st},$B{r}>{A},{st}=0),NA(),IF({t}<={ts},{q0}*(1+MAX({bb},0.001)*{D}*{t})^(-1/MAX({bb},0.001)),"
                               f"{qs}*EXP(-{IA['dm']}*({t}-{ts})))*IF(AND({sa}>0,$B{r}>={sa}),{sf},1))")
            wv.cell(r, c0 + 2, f"=IF(OR($B{r}<{A},$B{r}-{A}>{NK}),NA(),IF($B{r}={A},{IA['mult']}*N({wix('mult')})*{wix(s + '_q')},INDEX({FA},$B{r}-{A},{hc})))")
            for q, off in ((3, 2), (4, 3), (5, 4)):
                v = f"INDEX({OD},$B{r}-{A},{oc_}+{off})"
                wv.cell(r, c0 + q, f"=IF(OR($B{r}<={A},$B{r}-{A}>{NK}),NA(),IF(ISNUMBER({v}),{v},NA()))")
        for c in range(3, 15):
            wv.cell(r, c).number_format = "#,##0"
    add_review_charts(wv)
    wv.freeze_panes = f"A{T0}"

    # ------------------------------------------------------------ Well Charts
    wc = wb["Well Charts"]
    wc["B2"] = "Well charts (static, from pdp_fit.py): history, fit, forecast and the offsets' anchored P25-P75 / P50"; wc["B2"].font = TITLE
    wc["B3"] = "Grey crosses = months left out of the fit (shut-ins, short producing days, outliers). Live version of one well at a time: Well Review."

    # ------------------------------------------------------------ Summary
    S = wb["Summary"]
    S.column_dimensions["A"].width = 2; S.column_dimensions["B"].width = 46; S.column_dimensions["C"].width = 18
    S.column_dimensions["D"].width = 70
    for col in "EFGHIJKL":
        S.column_dimensions[col].width = 14
    S["B2"] = f'={IA["name"]}&" - PDP evaluation"'; S["B2"].font = TITLE
    S["B3"] = (f'={IA["operator"]}&" | "&{IA["basin"]}&" | "&{IA["dtype"]}&" | effective "&TEXT({IA["eff"]},"mmm yyyy")&" | "&'
               f'COUNTIF(Wells!${WL["active"]}${W0}:${WL["active"]}${W0 + NW - 1},1)&" wells"')
    S["B5"] = "Price deck"; S["C5"] = "Strip"; S["C5"].font = Font(bold=True, color="0000FF"); S["C5"].fill = INP_FILL
    dv = DataValidation(type="list", formula1='"' + ",".join(DECKS) + '"'); S.add_data_validation(dv); dv.add("C5")
    S["D5"] = "Strip = NYMEX strip on the Prices tab; flat decks hold WTI / HH flat"
    PK = lambda c: f"Package!${c}${M0}:${c}${M1}"
    tg = IA["tirr"]
    sub(S, 6, "Value", 2, 4); S["C6"] = "Value"; S["D6"] = "Basis"
    res = [
        (7, "Asking / offered price", f"={IA['price']}", "$#,##0", "Inputs C11, paid in month 0"),
        (8, "IRR at the asking price", f'=IF({IA["price"]}<=0,"n/a",IFERROR((1+IRR({PK("T")},0.02))^12-1,"n/a"))', "0.0%", "Monthly cash flows, annualized"),
        (9, "MOIC at the asking price", f'=IF({IA["price"]}<=0,"n/a",SUMIF({PK("T")},">0")/-SUMIF({PK("T")},"<0"))', "0.00x", "Total inflows / total outflows"),
        (10, "Payout (months from the effective date)", f'=IF({IA["price"]}<=0,"n/a",IF(MAX({PK("W")})<0,"never",COUNTIF({PK("W")},"<0")))', "0", "Months with cumulative net cash flow below zero"),
        (11, f'="Max price for a "&TEXT({tg},"0%")&" IRR"', f"=SUMPRODUCT({PK('R')}/(1+{tg})^({PK('A')}/12))", "$#,##0",
         "Operating cash flow discounted at the target IRR (monthly, month 0 undiscounted): paying this gives exactly the target IRR"),
        (12, "PV10 (operating cash flow, before the price)", f"=SUMPRODUCT({PK('R')},{PK('V')})", "$#,##0", "Mid-month discounting at Inputs C42"),
        (13, "PV15", f"=SUMPRODUCT({PK('R')}/(1.15)^(({PK('A')}+0.5)/12))", "$#,##0", ""),
        (14, "PV20", f"=SUMPRODUCT({PK('R')}/(1.2)^(({PK('A')}+0.5)/12))", "$#,##0", ""),
        (15, "Asking price as % of PV10", '=IF(C12<=0,"n/a",C7/C12)', "0%", ""),
        (16, "Undiscounted operating cash flow", f"=SUM({PK('R')})", "$#,##0", ""),
    ]
    sub(S, 18, "Reserves and metrics (net to our interest)", 2, 4)
    res += [
        (19, "Net oil (MBO)", f"=SUM({PK('G')})/1000", "#,##0.0", "Remaining from the effective date to the economic limit"),
        (20, "Net gas (MMcf, wellhead)", f"=SUM({PK('H')})/1000", "#,##0.0", ""),
        (21, "Net NGL (MBbl)", f"=SUM({PK('I')})/1000", "#,##0.0", ""),
        (22, "Net reserves (MBOE, oil + wellhead gas at 6:1)", "=C19+C20/6", "#,##0.0", ""),
        (23, "Gross remaining oil / gas (MBO / MMcf)", f'=TEXT(SUM({PK("E")})/1000,"#,##0")&" / "&TEXT(SUM({PK("F")})/1000,"#,##0")', None, "8/8, economic wells"),
        (24, "Longest economic life (years)", f"=Wells!${WL['life']}${W0 + NW}/12", "0.0", ""),
        (25, "Current net production (BOE/d, month 0)", f"=(Package!$G${M0}+Package!$H${M0}/6)/30.4", "#,##0.0", "Forecast for the effective month"),
        (26, "Asking price per flowing BOE/d", '=IF(OR(C7<=0,C25<=0),"n/a",C7/C25)', "$#,##0", ""),
        (27, "Max price per flowing BOE/d", '=IF(C25<=0,"n/a",C11/C25)', "$#,##0", ""),
        (28, "Next-12-month operating cash flow", f"=SUM(Package!$R${M0}:$R${M0 + 11})", "$#,##0", ""),
        (29, "Asking price / NTM cash flow", '=IF(OR(C7<=0,C28<=0),"n/a",C7/C28)', "0.0x", ""),
        (30, "Max price / NTM cash flow", '=IF(C28<=0,"n/a",C11/C28)', "0.0x", ""),
        (31, "PV10 per net BOE", '=IF(C22<=0,"n/a",C12/(C22*1000))', '"$"0.00', ""),
    ]
    sub(S, 33, "Offset check (package; detail on Offset Check)", 2, 4)
    o_, g_ = OCC["o"], OCC["g"]
    res += [
        (34, "Oil: forecast vs offsets' P50, next 24 months", f"='Offset Check'!{o_['Ratio 24 mo']}{tr}", "0%", f"='Offset Check'!{o_['Verdict']}{tr}"),
        (35, "Oil: forecast vs offsets' P50, full overlap", f"='Offset Check'!{o_['Ratio full overlap']}{tr}", "0%", ""),
        (36, "Gas: forecast vs offsets' P50, next 24 months", f"='Offset Check'!{g_['Ratio 24 mo']}{tr}", "0%", f"='Offset Check'!{g_['Verdict']}{tr}"),
        (37, "Gas: forecast vs offsets' P50, full overlap", f"='Offset Check'!{g_['Ratio full overlap']}{tr}", "0%", ""),
        (38, "Wells outside the tolerance (oil / gas)",
         f'=COUNTIF(\'Offset Check\'!{o_["Verdict"]}{W0}:{o_["Verdict"]}{tr - 1},"Above*")+COUNTIF(\'Offset Check\'!{o_["Verdict"]}{W0}:{o_["Verdict"]}{tr - 1},"Below*")'
         f'&" / "&(COUNTIF(\'Offset Check\'!{g_["Verdict"]}{W0}:{g_["Verdict"]}{tr - 1},"Above*")+COUNTIF(\'Offset Check\'!{g_["Verdict"]}{W0}:{g_["Verdict"]}{tr - 1},"Below*"))',
         None, "Look at these on Well Review; override qa / Da / b on Wells if the offsets are the better guide"),
    ]
    sub(S, 40, "Checks", 2, 4)
    notes_rng = lambda s: f"Wells!${WL[s + '_note']}${W0}:${WL[s + '_note']}${W0 + NW - 1}"
    res += [
        (41, "Formula errors", f'=IF(SUMPRODUCT(--ISERROR(Package!$A${M0}:$X${M1}))+SUMPRODUCT(--ISERROR(\'M Net\'!${MCL0}${M0}:${MCL1}${M1}))'
                                f'+SUMPRODUCT(--ISERROR(Wells!$B${W0}:${WL["life"]}${W0 + NW - 1}))=0,"OK","ERRORS")', None, ""),
        (42, "Package ties to the well cash flows", f'=IF(SUMPRODUCT(ABS({PK("U")}))<1,"OK","CHECK")', None, "Package column U"),
        (43, "Engine ties to the workbook (PV10)", '=IF(NOT(ISNUMBER(E43)),"not run",IF(ABS(E43-C12)<=MAX(1,0.001*ABS(C12)),"OK","RE-RUN"))', None,
         "E43 = PV10 from pdp_engine.py when write_pdp.py ran, at the Strip deck. RE-RUN after any override or input change if you want the grid below refreshed; the live numbers above are always current."),
        (44, "Fit notes (wells with a note: oil / gas)", f'=COUNTIF({notes_rng("o")},"?*")&" / "&COUNTIF({notes_rng("g")},"?*")', None,
         "b at its 2.0 cap or 0.01 floor, a level shift, short history or b from offsets: see Wells"),
        (45, "All checks", '=IF(AND(C41="OK",C42="OK",OR(C43="OK",C43="not run")),"All checks OK","CHECK")', None, ""),
    ]
    for r, lab, f_, fmt, note in res:
        S.cell(r, 2, lab); x = S.cell(r, 3, f_)
        if fmt:
            x.number_format = fmt
        S.cell(r, 4, note).alignment = Alignment(wrap_text=True, vertical="top")
    for r in (8, 11, 12, 45):
        S.cell(r, 3).font = BOLD
    S["C11"].fill = INP_FILL
    S["E43"].number_format = "$#,##0"; S["E43"].fill = FIT_FILL
    # sensitivity grid (values written by write_pdp.py)
    sub(S, 47, "Sensitivities (pdp_engine.py, written by write_pdp.py at its run; Max price for the target IRR / PV10)", 2, 12)
    S["B48"] = "Deck \\ forecast multiplier"
    for j, mlt in enumerate([0.85, 1.0, 1.15]):
        S.cell(48, 3 + j, mlt).number_format = "0.00x"; S.cell(48, 3 + j).font = BOLD
        S.cell(48, 7 + j, mlt).number_format = "0.00x"; S.cell(48, 7 + j).font = BOLD
    S["C47"].value = None
    S.cell(49, 2, "")
    for i, dk in enumerate(DECKS):
        S.cell(49 + i, 2, dk)
        for j in range(3):
            S.cell(49 + i, 3 + j).number_format = "$#,##0"; S.cell(49 + i, 3 + j).fill = FIT_FILL
            S.cell(49 + i, 7 + j).number_format = "$#,##0"; S.cell(49 + i, 7 + j).fill = FIT_FILL
    S["C53"] = "Max price for the target IRR"; S["G53"] = "PV10"
    S["B54"] = "Anchor: Fit / Actual (strip, max price)"; S["C54"].number_format = "$#,##0"; S["D54"].number_format = "$#,##0"
    S["C54"].fill = FIT_FILL; S["D54"].fill = FIT_FILL
    S["B55"] = "Economic limit off (strip, max price)"; S["C55"].number_format = "$#,##0"; S["C55"].fill = FIT_FILL
    # annual table
    sub(S, 57, "Annual cash flow (net)", 2, 12)
    ah = ["Year", "Net oil (bbl)", "Net gas (Mcf)", "Revenue", "Taxes", "Deducts + LOE", "P&A", "Operating cash flow", "Wells on (year end)", "Discounted (PV rate)"]
    for j, t in enumerate(ah):
        hdr(S, 58, 2 + j, t)
    yr = PK("X")
    for y in range(1, 16):
        r = 58 + y
        lab = y if y < 15 else "15+"
        S.cell(r, 2, lab)
        cond = f'{yr},{y}' if y < 15 else f'{yr},">=15"'
        for j, col in enumerate(["G", "H", "M", "N", None, "Q", "R"]):
            if col is None:
                S.cell(r, 3 + j, f"=SUMIFS({PK('O')},{cond})+SUMIFS({PK('P')},{cond})").number_format = "#,##0"
            else:
                S.cell(r, 3 + j, f"=SUMIFS({PK(col)},{cond})").number_format = "#,##0"
        S.cell(r, 10, f"=IFERROR(SUM(INDEX('M On'!${MCL0}${M0}:${MCL1}${M1},MIN({NM},{min(y, 15) * 12}),0)),0)" if y < 15 else "").number_format = "0"
        S.cell(r, 11, f"=SUMPRODUCT(({yr}{'=' + str(y) if y < 15 else '>=15'})*{PK('R')}*{PK('V')})").number_format = "#,##0"
    r = 74
    S.cell(r, 2, "Total").font = BOLD
    for c in (3, 4, 5, 6, 7, 8, 9, 11):
        S.cell(r, c, f"=SUM({L(c)}59:{L(c)}73)").number_format = "#,##0"; S.cell(r, c).font = BOLD
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    print("wrote", out)


def add_review_charts(wv, T0=12):
    """Well Review log charts (also re-added by write_pdp.py: openpyxl does not keep charts when it re-opens a workbook)."""
    for title, c0, anchor_cell, unit in (("Oil (bbl/month)", 3, "P10", "bbl/month"), ("Gas (Mcf/month)", 9, "P34", "Mcf/month")):
        ch = ScatterChart(); ch.title = title; ch.style = 13; ch.height = 11; ch.width = 22
        ch.y_axis.scaling.logBase = 10; ch.y_axis.title = unit; ch.x_axis.title = "Months on production"
        ch.x_axis.delete = False; ch.y_axis.delete = False
        xs = Reference(wv, min_col=2, min_row=T0, max_row=T0 + 179)
        spec = [(0, "Actual", "2E7D32" if c0 == 3 else "C62828", "marker"), (1, "Fit", "7F7F7F", "line"), (2, "Forecast", "EF6C00", "thick"),
                (3, "Offsets P25", "B39DDB", "dot"), (4, "Offsets P50", "7E57C2", "dash"), (5, "Offsets P75", "B39DDB", "dot")]
        for off, name, color, kind in spec:
            ys = Reference(wv, min_col=c0 + off, min_row=T0 - 1, max_row=T0 + 179)
            se = Series(ys, xs, title_from_data=True)
            if kind == "marker":
                se.marker.symbol = "circle"; se.marker.size = 4
                se.marker.graphicalProperties.solidFill = color; se.marker.graphicalProperties.line.solidFill = color
                se.graphicalProperties.line.noFill = True
            else:
                se.marker.symbol = "none"
                se.graphicalProperties.line.solidFill = color
                se.graphicalProperties.line.width = 28000 if kind == "thick" else 15000
                if kind in ("dot", "dash"):
                    se.graphicalProperties.line.dashStyle = "sysDot" if kind == "dot" else "dash"
            se.smooth = False
            ch.series.append(se)
        ch.display_blanks = "gap"
        wv.add_chart(ch, anchor_cell)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "PDP Evaluation Template.xlsx"))
