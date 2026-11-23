"""Deal-specific workbook edits for Mystical Fed x5 (WRC Energy, Converse Co WY), Citation (CILP-VI) interest.
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
TITLE = "Mystical Fed x5 (WRC Energy; Sussex / Mowry / Niobrara, Converse Co WY)"

MYS_NOTE = ("WRC proposal 10/29/2026 (AFEs dated 10/22/2026): drilling and completion columns as on the AFE; the completion "
            "column includes the tangible equipment and facilities (no separate facility AFE). Election due 5:00 pm MST 12/3/2026; "
            "prepay the AFE within 30 days of WRC's spud notice; WOGCC pooling hearing 12/8/2026. ASSUMPTION: no spud date given; "
            "closing / drilling Mar-27, completion Jun-27, first sales Jul-27. No child haircut (fund standard 0%; WRC's Mystical "
            "Fed 1NH and 1SH have produced in these sections since Jun-24, see the 15% haircut case on Summary).")
# WRC AFE (drilling, completion) by API; laterals from the AFE descriptions
WRC = {"49-009-50664": (3548351, 7814060, 15413), "49-009-50665": (3526408, 7785629, 15360),
       "49-009-38507": (4402099, 8733840, 10217), "49-009-49482": (3073362, 8860453, 10256),
       "49-009-50896": (2757628, 8731538, 10188)}
WELLS = [
    # name, api, lateral, WI, lease NRI, gross AFE, (drill, compl, fac month), status, zone, note
    ("MYSTICAL FED 4075-25-24-13-1SXH", "49-009-50664", 15906, 0.044525, 0.033340 / 0.044525, 11362411,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permitted", "Sussex",
     "Sussex (seller SSX; Enverus permit SUSSEX, TVD 10,179 ft). AFE: 15,413 ft Sussex lateral, 77 stages, 26,503 ft TMD (Enverus permit 15,906 ft); drilling $3,548,351, completion $7,814,060. "
     "Net AFE $505,911.35. Type curve: Sussex zone curve (6 offsets, 10 mi, low confidence). " + MYS_NOTE),
    ("MYSTICAL FED 4075-25-24-13-2SXH", "49-009-50665", 15745, 0.044525, 0.033340 / 0.044525, 11312038,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permitted", "Sussex",
     "Sussex (seller SSX; Enverus permit SUSSEX, TVD 10,169 ft). AFE: 15,360 ft lateral, 77 stages (Enverus permit 15,745 ft); drilling $3,526,408, completion $7,785,629. Net AFE $503,668.50. "
     "Type curve: Sussex zone curve. " + MYS_NOTE),
    ("MYSTICAL FED 4075-24-13-1MH", "49-009-38507", 10607, 0.066787, 0.050100 / 0.066787, 13135939,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permitted", "Mowry (Mystical)",
     "Mowry (seller MWRY; Enverus permit MOWRY, TVD 13,225 ft). AFE: 10,217 ft Mowry lateral, 51 stages, 24,016 ft TMD (Enverus permit 10,607 ft); drilling $4,402,099, completion $8,733,840. "
     "Net AFE $877,309.96. Type curve: Mowry (Mystical) zone curve (7 offsets, 10 mi, low confidence). " + MYS_NOTE),
    ("MYSTICAL FED 4075-24-13-2NH", "49-009-49482", 10832, 0.066787, 0.050100 / 0.066787, 11933815,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permitted", "Niobrara C",
     "Niobrara C per the seller table (NIO C; listed as 'MYSTICAL FE'); the Enverus permit says NIOBRARA A, TVD 11,818 ft. "
     "AFE (target 'Niobrara'): 10,256 ft lateral, 51 stages (Enverus permit 10,832 ft); drilling $3,073,362, completion $8,860,453. Net AFE $797,023.70. Type curve: 'Niobrara C' = blended Niobrara A + B offsets inside 5 mi (15 wells; "
     "no Niobrara C well with 12+ months within 20 mi). " + MYS_NOTE),
    ("MYSTICAL FED 4075-24-13-3NH", "49-009-50896", 9921, 0.066787, 0.050100 / 0.066787, 11489166,
     (D(2027, 3, 1), D(2027, 6, 1), D(2027, 7, 1)), "Permit pending", "Niobrara B",
     "Niobrara B per the seller table (NIO B); the Enverus permit (pending approval) says NIOBRARA A, TVD 11,808 ft. "
     "AFE (target 'Niobrara'): 10,188 ft lateral, 51 stages (Enverus permit 9,921 ft); drilling $2,757,628, completion $8,731,538. Net AFE $767,326.93. Type curve: Niobrara B zone curve (8 offsets, 5 mi, low confidence; 30-well "
     "10-mile comparison). " + MYS_NOTE),
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
    net = 505911.35 + 503668.50 + 877309.96 + 797023.70 + 767326.93
    W["F19"] = round(net, 2)
    W["G19"] = ("Citation well table net AFEs: SXH x2 $1,009,579.85 (4.4525% WI), MH/NH x3 $2,441,660.59 (6.6787% WI); "
                "gross matches the WRC AFEs ($59,233,369)")

    I = wb["Inputs"]
    I["B2"] = TITLE + ": Non-Op Participation Model Inputs"
    I["B3"] = ("Blue = hard-coded input. Black = formula. Yellow = key assumption to confirm with the seller (Citation 2002 Investment "
               "LLC / CILP-VI) or WRC Energy. ASSUMPTION = added by Osage Manka, not from the seller, the AFE or Enverus.")
    I["C19"] = "Powder River (WY)"
    I["D10"] = ("Paid in each well's drilling (closing) month on top of the net capex; no asking premium given, 20% assumed; "
                "the fund screen is 20%")
    I["B11"] = "Type curve multiplier (100% = zone comp curves)"
    I["B13"] = "Type curve (1 = zone comp curves, 2 = Enverus fit-set P50 on our shape)"
    I["D12"] = ("Default child haircut, 0% (fund standard). WRC's Mystical Fed 1NH (Niobrara A) and 1SH (Shannon) already "
                "produce in the Mystical sections; the 15% haircut case is on Summary. Override here or per well (Wells column Q)")
    I["D13"] = ("Base = 1 (zone 1 Niobrara C: oil 523 MBO, gas 1.62 Bcf per 10,000 ft; 15 fit wells). Enverus EUR fields are blank "
                "on this subscription, so 2 = base. Multi-zone: 4 zone curves on 'Zone Curves'; each well uses its zone "
                "(Wells column X); the switch applies to zone 1 only")
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
    S["B3"] = ("Pure non-op economics at WRC's AFEs (10/22/2026) for Citation's (CILP-VI) interest: 1SXH / 2SXH (Sussex, 4.4525% WI) "
               "and 1MH (Mowry), 2NH (Niobrara C) and 3NH (Niobrara B) (6.6787% WI). Gross AFEs $59.2MM, net $3.45MM, plus the "
               "premium. No asking premium given; 20% is the base. Live case follows Inputs, the deck selector (C6) and the premium (C7).")
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["C7"] = 0.2
    S["B78"] = "Confirm with the seller (Citation / CILP-VI) and WRC Energy"
    notes = [
        "- Asking price / premium: none given; the model uses 20% of net capex and shows the max premium for a 20% IRR",
        "- Mystical elections: WRC proposal 10/29/2026, elect by 5:00 pm MST 12/3/2026, prepay AFE within 30 days of the spud notice; "
        "non-consent is pooled with statutory risk penalties (W.S. 30-5-109(f)). No spud date given: Mar-27 drilling / Jul-27 first "
        "sales assumed; 3NH permit still pending approval",
        "- Landing zones: the seller table lists 2NH as Niobrara C and 3NH as Niobrara B; WRC's AFEs say only 'Niobrara' and Enverus permits show both as Niobrara A. "
        "No Niobrara C well with 12+ months within 20 mi, so 2NH is priced on blended Niobrara A + B offsets",
        "- Title to the WIs / NRIs (74.9-75.0% lease NRI); WRC proposes a participation agreement under pooling (no JOA): "
        "$15,683 / well / month drilling overhead and $1,568 producing overhead",
        "- Gas marketing, NGL netback and LOE: Powder River basin defaults used, not checked against WRC JIBs or contracts",
    ]
    for k in range(8):
        S[f"B{79 + k}"] = notes[k] if k < len(notes) else None

    T = wb["TypeCurve"]
    T["B2"] = ("Niobrara C (zone 1) Type Curves per 10,000 ft lateral (gross 8/8; oil bbl/month, gas Mcf/month; before the child "
               "haircut)")
    T["B3"] = ("5-mile radius comp standard (Nov 23, 2026). Enverus Developer API (Enverus connector): no Niobrara C well with 12+ "
               "months within 20 mi, so zone 1 is fitted on the Niobrara A + B horizontals (first prod 2019-10 to 2025-05, "
               "laterals 9,593-10,175 ft) with lateral midpoint within 5 mi of the Mystical DSU centroid (43.4277, -105.8075): "
               "15 wells (WRC, Sage Butte, Titan). Enverus EUR fields are blank on this subscription, so the guardrail is off. "
               "Partial first months scaled to a full month from Enverus producing days.")
    for sh in ("Comps", "5mi vs 10mi"):
        if sh in wb.sheetnames:
            del wb[sh]
else:
    bt.load_deal(os.path.join(HERE))
    import json
    Z = json.load(open(os.path.join(HERE, "tc_zones.json")))["zones"]
    import importlib.util
    spec = importlib.util.spec_from_file_location("dd", os.path.join(HERE, "deal_data.py"))
    dd = importlib.util.module_from_spec(spec); spec.loader.exec_module(dd)
    C = wb["Comps 5mi"]
    for row in C.iter_rows():
        for c in row:
            c.value = None
    C["B2"] = "Comp selection (5-mile radius standard): Mystical Fed DSUs (WRC), centroid 43.4277, -105.8075, Converse Co WY"
    C["B2"].font = Font(bold=True, size=14)
    C["B3"] = ("Distance = offset lateral midpoint to the DSU centroid. Screen: horizontal, first production "
               "2017-01 or later, 12+ months by 2026-09, same Enverus interval. Radius rule per zone: 5 mi, widened to 10 / 15 / 20 mi "
               "until 6+ wells. Excluded: SPRUCE 21W33-2NH (bad header: SHL = BHL, TVD 10,283 ft) and RUSTY FEDERAL 4075-23-14-1UH "
               "(2018 legacy Sussex completion). 12-month volumes per 10,000 ft after the partial-first-month scaling.")
    hdr = ["Zone", "Fit radius", "Fit wells", "Oil EUR (MBO / 10k ft)", "Gas EUR (Bcf / 10k ft)", "Oil b", "Gas b",
           "Low confidence", "Note"]
    for j, h in enumerate(hdr):
        C.cell(5, 2 + j, h).font = Font(bold=True)
    for k, z in enumerate(Z):
        r = 6 + k
        vals = [z["name"], z["radius"], z["wells"], round(z["oil"]["eur"] / 1000), round(z["gas"]["eur"] / 1e6, 2),
                round(z["oil"]["b"], 2), round(z["gas"]["b"], 2), "Yes" if z["low_confidence"] else "No", z.get("note", "")]
        for j, v in enumerate(vals):
            C.cell(r, 2 + j, v)
    r0 = 6 + len(Z) + 1
    hdr2 = ["Fit set", "Well", "API", "Operator", "Enverus interval", "First prod", "Lateral (ft)", "Distance (mi)",
            "12-mo oil (MBO / 10k ft)", "12-mo gas (MMcf / 10k ft)"]
    for j, h in enumerate(hdr2):
        C.cell(r0, 2 + j, h).font = Font(bold=True)
    r = r0 + 1
    sets = []
    for name, z in dd.ZONES.items():
        sets.append((f"{name} {z['radius']}", z["fit"]))
        if z.get("fit_10mi"):
            sets.append((f"{name} 10 mi (comparison)", [a for a in z["fit_10mi"] if a not in z["fit"]]))
    for lab, apis in sets:
        for a in apis:
            nm, lat, first, o, g, iv, dist, op = dd.WELLS[a]
            oc, gc = bt.clean(o, g, dd.DAYS.get(a), first)
            k = 10000 / lat
            vals = [lab, nm, a, op, iv.title(), first, lat, dist, round(oc[:12].sum() * k / 1000), round(gc[:12].sum() * k / 1000)]
            for j, v in enumerate(vals):
                C.cell(r, 2 + j, v)
            r += 1
wb.save(f)
print(mode, "done")
