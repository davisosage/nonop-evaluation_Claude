"""Deal-specific workbook edits for Mystical Fed 1SXH / 2SXH (WRC Energy, Sussex, Converse Co WY), Citation (CILP-VI) interest.
    python3 fill_model.py pre  <template copy.xlsx>   # wells, inputs, strip, text; before write_model.py
    python3 fill_model.py post <model.xlsx>           # comps sheet, after write_model.py
Run from toolkit/. Recalculate after each step."""
import datetime as dt
import os
import sys

import numpy as np
import openpyxl
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "toolkit"))
import build_tc as bt  # noqa: E402

mode, f = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(f)
D = dt.datetime
DECK = ("NYMEX WTI strip Nov 20, 2026 (Nov 3 curve rescaled through the Nov 20 settles; latest deck in the repo), "
        "HH Nov 3, 2026; flat after 2030")
TITLE = "Mystical Fed 4075-25-24-13 1SXH / 2SXH (WRC Energy; Sussex, Converse Co WY)"

MYS_NOTE = ("WRC proposal 10/29/2026 (AFEs dated 10/22/2026): drilling and completion columns as on the AFE; the completion "
            "column includes the tangible equipment and facilities (no separate facility AFE). Election due 5:00 pm MST 12/3/2026; "
            "prepay the AFE within 30 days of WRC's spud notice; WOGCC pooling hearing 12/8/2026. ASSUMPTION: no spud date given; "
            "closing / drilling Mar-27, completion Jun-27, first sales Jul-27. No child haircut (fund standard 0%; WRC's Mystical "
            "Fed 1NH and 1SH have produced in these sections since Jun-24, in other zones; see the 15% haircut case on Summary).")
# WRC AFE (drilling, completion) by API; laterals from the AFE descriptions
WRC = {"49-009-50664": (3548351, 7814060, 15413), "49-009-50665": (3526408, 7785629, 15360),
       "49-009-38507": (4402099, 8733840, 10217), "49-009-49482": (3073362, 8860453, 10256),
       "49-009-50896": (2757628, 8731538, 10188)}
WELLS = [
    # name, api, lateral, WI, lease NRI, gross AFE, (drill, compl, fac month), status, zone, note
    ("MYSTICAL FED 4075-25-24-13-1SXH", "49-009-50664", 15906, 0.044525, 0.033340 / 0.044525, 11362411,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permitted", None,
     "Sussex (seller SSX; Enverus permit SUSSEX, TVD 10,179 ft). AFE: 15,413 ft Sussex lateral, 77 stages, 26,503 ft TMD (Enverus permit 15,906 ft); drilling $3,548,351, completion $7,814,060. "
     "Net AFE $505,911.35. Type curve: Sussex zone curve (6 offsets, 10 mi, low confidence). " + MYS_NOTE),
    ("MYSTICAL FED 4075-25-24-13-2SXH", "49-009-50665", 15745, 0.044525, 0.033340 / 0.044525, 11312038,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permitted", None,
     "Sussex (seller SSX; Enverus permit SUSSEX, TVD 10,169 ft). AFE: 15,360 ft lateral, 77 stages (Enverus permit 15,745 ft); drilling $3,526,408, completion $7,785,629. Net AFE $503,668.50. "
     "Type curve: Sussex zone curve. " + MYS_NOTE),
]

if mode == "pre":
    W = wb["Wells"]
    for k in range(10):
        r = 7 + k
        for col in "CDEFGIJKNOPSTUX":
            W[f"{col}{r}"] = None
        W[f"R{r}"] = 1
        W[f"Q{r}"] = "=Inputs!$C$12"
    for k, (nm, api, lat, wi, nri, gross, (dm, cm, fm), st, zone, note) in enumerate(WELLS):
        r = 7 + k
        d, c, lat = WRC[api]; fac = 0
        assert abs(d + c - gross) <= 1, api  # 2SXH columns sum $1 under the AFE total (rounding)
        vals = dict(C=nm, D=api, E=lat, F=wi, G=nri, I=d, J=c, K=fac, N=dm, O=cm, P=fm, S=st, T=note, U=0, X=zone)
        for col, v in vals.items():
            W[f"{col}{r}"] = v
        for col in "NOP":
            W[f"{col}{r}"].number_format = "mmm-yy"
    net = 505911.35 + 503668.50
    W["F19"] = round(net, 2)
    W["G19"] = ("Citation well table net AFEs: 1SXH $505,911.35 + 2SXH $503,668.50 at 4.4525% WI; gross matches the WRC "
                "AFEs ($22,674,449)")

    I = wb["Inputs"]
    I["B2"] = TITLE + ": Non-Op Participation Model Inputs"
    I["B3"] = ("Blue = hard-coded input. Black = formula. Yellow = key assumption to confirm with the seller (Citation 2002 Investment "
               "LLC / CILP-VI) or WRC Energy. ASSUMPTION = added by Osage Manka, not from the seller, the AFE or Enverus.")
    I["C19"] = "Powder River (WY)"
    I["D10"] = ("Paid in each well's drilling (closing) month on top of the net capex; no asking premium given, 20% assumed; "
                "the fund screen is 20%")
    I["B11"] = "Type curve multiplier (100% = 10-mile Sussex comp base)"
    I["B13"] = "Type curve (1 = 10-mile Sussex comp base, 2 = Enverus fit-set P50 on our shape)"
    I["D12"] = ("Default child haircut, 0% (fund standard). No producing Sussex well in the DSU (WRC's Mystical Fed 1NH and "
                "1SH are Niobrara A and Shannon); the 15% haircut case is on Summary. Override here or per well (Wells column Q)")
    I["D13"] = ("Base = 1 (Sussex: oil 563 MBO, gas 0.43 Bcf per 10,000 ft; 6 fit wells at 10 mi, LOW CONFIDENCE). Enverus EUR "
                "fields are blank on this subscription, so 2 = base. Both wells use the TypeCurve tab (Sussex)")
    I["D38"] = DECK

    # strip: latest deck in the repo (Nov 20, 2026 WTI rescale), copied from deals/sun5899-2bhp-3bhp
    src = openpyxl.load_workbook(os.path.join(HERE, "..", "sun5899-2bhp-3bhp", "Sun 5899 11-28 Williston Non-Op Model.xlsx"))
    P, SP = wb["Prices"], src["Prices"]
    for i in range(400):
        P.cell(6 + i, 4).value = SP.cell(6 + i, 4).value
        P.cell(6 + i, 5).value = SP.cell(6 + i, 5).value
    P["B3"] = DECK + ". Nov/Dec-26 use the Jan-27 contract. Flat decks follow the Summary C6 dropdown."

    S = wb["Summary"]
    S["B2"] = TITLE + ": Buy the Citation WI or Not"
    S["B3"] = ("Pure non-op economics at WRC's AFEs (10/22/2026) for Citation's (CILP-VI) 4.4525% WI in the two Sussex wells: "
               "1SXH ($11,362,411) and 2SXH ($11,312,038), net $1.01MM, plus the premium. The 1MH / 2NH / 3NH wells in WRC's "
               "proposal are left out. No asking premium given; 20% is the base. Live case follows Inputs, the deck selector (C6) "
               "and the premium (C7).")
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["C7"] = 0.2
    S["B78"] = "Confirm with the seller (Citation / CILP-VI) and WRC Energy"
    notes = [
        "- Asking price / premium: none given; the model uses 20% of net capex and shows the max premium for a 20% IRR",
        "- Mystical elections: WRC proposal 10/29/2026, elect by 5:00 pm MST 12/3/2026, prepay AFE within 30 days of the spud notice; "
        "non-consent is pooled with statutory risk penalties (W.S. 30-5-109(f)). No spud date given: Mar-27 drilling / Jul-27 first "
        "sales assumed",
        "- Sussex laterals: the AFEs plan 15,360-15,413 ft (3 sections, 77 stages); every Sussex offset is a ~10,000 ft lateral, "
        "so the curve is scaled linearly to the longer laterals (untested in this area)",
        "- Electing only the SXH wells: WRC lets Owner elect well by well; the other three wells would go to pooling as non-consent",
        "- Title to the 4.4525% WI / 74.9% lease NRI (3.3340% NRI); WRC proposes a participation agreement under pooling (no JOA): "
        "$15,683 / well / month drilling overhead and $1,568 producing overhead",
        "- Gas marketing, NGL netback and LOE: Powder River basin defaults used, not checked against WRC JIBs or contracts",
    ]
    for k in range(8):
        S[f"B{79 + k}"] = notes[k] if k < len(notes) else None

    T = wb["TypeCurve"]
    T["B2"] = "Sussex Type Curves per 10,000 ft lateral (gross 8/8; oil bbl/month, gas Mcf/month; before the child haircut)"
    T["B3"] = ("Comp standard (Nov 23, 2026). Enverus Developer API (Enverus connector): Sussex horizontals, first prod 2022-02 to "
               "2024-10, laterals 9,688-10,138 ft. 4 Sussex wells (2021+ completions) inside 5 mi of the Mystical DSU centroid "
               "(43.4277, -105.8075), so the set widens to 10 mi: 6 wells (Sage Butte Spruce 23W14-1SUH, Continental Ogalalla Fed "
               "1SXH / 3SXH and Solitude Fed 1SXH / 3SXH, WRC Mather Peaks 1SXH). LOW CONFIDENCE (< 15 wells). Excluded: Rusty "
               "Federal 4075-23-14-1UH (2018 legacy completion). Enverus EUR fields are blank on this subscription, so the "
               "guardrail is off. Partial first months scaled to a full month from Enverus producing days.")
    for sh in ("Comps", "5mi vs 10mi"):
        if sh in wb.sheetnames:
            del wb[sh]
else:
    import importlib.util
    spec = importlib.util.spec_from_file_location("dd", os.path.join(HERE, "deal_data.py"))
    dd = importlib.util.module_from_spec(spec); spec.loader.exec_module(dd)
    C = wb["Comps 5mi"]
    for row in C.iter_rows():
        for c in row:
            c.value = None
    C["B2"] = "Comp selection: Mystical Fed Sussex wells (WRC), DSU centroid 43.4277, -105.8075, Converse Co WY"
    C["B2"].font = Font(bold=True, size=14)
    C["B3"] = ("Distance = offset lateral midpoint to the DSU centroid. Screen: Sussex horizontals, first production 2017-01 or "
               "later, 12+ months by 2026-09. Radius: 5 mi (4 wells) -> 10 mi (6). Excluded: RUSTY FEDERAL 4075-23-14-1UH (2018 "
               "legacy completion). 12-month volumes per 10,000 ft after the partial-first-month scaling.")
    hdr = ["Fit set", "Well", "API", "Operator", "Enverus interval", "First prod", "Lateral (ft)", "Distance (mi)",
           "12-mo oil (MBO / 10k ft)", "12-mo gas (MMcf / 10k ft)"]
    for j, h in enumerate(hdr):
        C.cell(5, 2 + j, h).font = Font(bold=True)
    for r, a in enumerate(dd.FIT_SET, start=6):
        nm, lat, first, o, g, iv, dist, op = dd.WELLS[a]
        oc, gc = bt.clean(o, g, dd.DAYS.get(a), first)
        k = 10000 / lat
        vals = ["Sussex 10 mi", nm, a, op, iv.title(), first, lat, dist, round(oc[:12].sum() * k / 1000),
                round(gc[:12].sum() * k / 1000)]
        for j, v in enumerate(vals):
            C.cell(r, 2 + j, v)
wb.save(f)
print(mode, "done")
