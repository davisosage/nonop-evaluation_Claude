"""Deal-specific workbook edits for Overtime Fed 4676-35-14-15 MH (Anschutz, Mowry, Campbell Co WY), Citation 2002 Investment interest.
    python3 fill_model.py pre  <template copy.xlsx>   # well, inputs, strip, text; before write_model.py
    python3 fill_model.py post <model.xlsx>           # comps sheet, after write_model.py
Run from toolkit/. Recalculate after each step."""
import datetime as dt
import importlib.util
import os
import sys

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
TITLE = "Overtime Fed 4676-35-14-15 MH (Anschutz; Mowry, Campbell Co WY)"

# Anschutz AFE (49-005-80253, $14,130,540): drilling + PDO/PDL, completion + lift, facilities
OT_D, OT_C, OT_F = 4579734 + 236042, 7709441 + 310111, 1295212
WI, NRI = 0.09214234, 0.075406 / 0.092142
NOTE = ("Mowry, Anschutz well proposal 11/5/2026 (11,565.56 ac DSU, T45N-R76W Secs 1-3, 10-15 and T46N-R76W Secs 22-27, 34-36). "
        "TMD 27,471 ft / TVD 11,622 ft; Enverus lateral 16,272 ft. AFE $14,130,540: drilling $4,579,734 + PDO/PDL $236,042; "
        "completion $7,709,441 + lift $310,111; facilities $1,295,212. Title WI 9.214234% (net AFE $1,302,021.02), subject to "
        "title examination; NRI 7.5406% (81.84% lease NRI) from the Citation well table. Spud ~12/4/2026. ASSUMPTION: "
        "completion Mar-27, first sales Apr-27. 15% child haircut is the fund default (no producing Mowry within 6.9 mi).")

if mode == "pre":
    W = wb["Wells"]
    for k in range(10):
        r = 7 + k
        for col in "CDEFGIJKNOPSTUX":
            W[f"{col}{r}"] = None
        W[f"R{r}"] = 1
        W[f"Q{r}"] = "=Inputs!$C$12"
    vals = dict(C="OVERTIME FED 4676-35-14-15 MH", D="49-005-80253", E=16272, F=WI, G=NRI, I=OT_D, J=OT_C, K=OT_F,
                N=D(2026, 12, 1), O=D(2027, 3, 1), P=D(2027, 4, 1), S="Permitted", T=NOTE, U=0)
    for col, v in vals.items():
        W[f"{col}7"] = v
    for col in "NOP":
        W[f"{col}7"].number_format = "mmm-yy"
    W["F19"] = 1302021.02
    W["G19"] = "Anschutz proposal: 9.214234% x $14,130,540 = $1,302,021.02"

    I = wb["Inputs"]
    I["B2"] = TITLE + ": Non-Op Participation Model Inputs"
    I["B3"] = ("Blue = hard-coded input. Black = formula. Yellow = key assumption to confirm with the seller (Citation 2002 Investment "
               "LLC / CILP-VI) or Anschutz. ASSUMPTION = added by Osage Manka, not from the seller, the AFE or Enverus.")
    I["C19"] = "Powder River (WY)"
    I["D10"] = ("Paid in the drilling (closing) month on top of the net capex; no asking premium given, 20% assumed; "
                "the fund screen is 20%")
    I["B11"] = "Type curve multiplier (100% = 15-mile Mowry comp base)"
    I["B13"] = "Type curve (1 = 15-mile Mowry comp base, 2 = Enverus fit-set P50 on our shape)"
    I["D12"] = "Default child haircut (fund standard). No producing Mowry within 6.9 mi of the DSU; see the 0% row on Summary"
    I["D13"] = ("Base = 1 (oil 508 MBO, gas 1.40 Bcf per 10,000 ft; 10 fit wells at 15 mi, LOW CONFIDENCE). Enverus EUR fields "
                "are blank on this subscription, so 2 = base")
    I["D38"] = DECK

    src = openpyxl.load_workbook(os.path.join(HERE, "..", "sun5899-2bhp-3bhp", "Sun 5899 11-28 Williston Non-Op Model.xlsx"))
    P, SP = wb["Prices"], src["Prices"]
    for i in range(400):
        P.cell(6 + i, 4).value = SP.cell(6 + i, 4).value
        P.cell(6 + i, 5).value = SP.cell(6 + i, 5).value
    P["B3"] = DECK + ". Nov/Dec-26 use the Jan-27 contract. Flat decks follow the Summary C6 dropdown."

    S = wb["Summary"]
    S["B2"] = TITLE + ": Buy the Citation 9.2142% WI or Not"
    S["B3"] = ("Pure non-op economics at Anschutz's AFE ($14,130,540 gross, $1,302,021 net at 9.214234% WI) plus the premium. "
               "No asking premium given; 20% is the base. Live case follows Inputs, the deck selector (C6) and the premium (C7).")
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["C7"] = 0.2
    S["F18"] = "After the 15% child haircut"
    S["B78"] = "Confirm with the seller (Citation / CILP-VI) and Anschutz"
    notes = [
        "- Asking price / premium: none given; the model uses 20% of net capex and shows the max premium for a 20% IRR",
        "- Election: Anschutz proposal dated 11/5/2026 (delivered 11/6), respond within 30 days (~12/6/2026); WOGCC pooling hearing "
        "12/8/2026. Non-consent goes to statutory pooling with the max risk penalty (W.S. 30-5-109(g)); closing must land first",
        "- Pay actual cost invoices within 30 days or Anschutz may deem the interest non-consent; conveyances are not recognized "
        "until invoiced costs are paid (proposal terms)",
        "- Type curve: no Mowry well inside 5 mi and 2 inside 10 mi; the 15-mile set is mostly Anschutz Cadet / Caterpillar wells "
        "with 14-34 months of history (4 wells with 24+ months), so b and EUR are low confidence",
        "- Completion and first-sales dates (model: completion Mar-27, first sales Apr-27 after a ~12/4/2026 spud)",
        "- Title to the 9.214234% WI / 81.84% lease NRI (WI subject to title examination per the proposal)",
        "- Gas marketing, NGL netback and LOE: Powder River basin defaults used, not checked against Anschutz JIBs or contracts",
    ]
    for k in range(8):
        S[f"B{79 + k}"] = notes[k] if k < len(notes) else None

    T = wb["TypeCurve"]
    T["B2"] = "Mowry Type Curves per 10,000 ft lateral (gross 8/8; oil bbl/month, gas Mcf/month; before the child haircut)"
    T["B3"] = ("Comp standard (Nov 23, 2026). Enverus Developer API (Enverus connector): Mowry horizontals, first prod 2021-11 to "
               "2025-07, laterals 9,420-13,102 ft. No Mowry well with 12+ months inside 5 mi of the Overtime DSU centroid "
               "(43.8896, -105.9591) and 2 inside 10 mi (Flying V, Case Fed), so the set widens to 15 mi: 10 wells (Anschutz "
               "Flying V / Case / Cadet x6 / Caterpillar, EOG Cutthroat). LOW CONFIDENCE (< 15 wells). Enverus EUR fields are "
               "blank on this subscription, so the guardrail is off. Partial first months scaled to a full month from Enverus "
               "producing days.")
    for sh in ("Comps", "5mi vs 10mi"):
        if sh in wb.sheetnames:
            del wb[sh]
else:
    spec = importlib.util.spec_from_file_location("dd", os.path.join(HERE, "deal_data.py"))
    dd = importlib.util.module_from_spec(spec); spec.loader.exec_module(dd)
    C = wb["Comps 5mi"]
    for row in C.iter_rows():
        for c in row:
            c.value = None
    C["B2"] = "Comp selection: Overtime Fed DSU (Anschutz), centroid 43.8896, -105.9591, Campbell Co WY"
    C["B2"].font = Font(bold=True, size=14)
    C["B3"] = ("Distance = offset lateral midpoint to the DSU centroid. Screen: Mowry horizontals, first production 2017-01 or later, "
               "12+ months. Radius: 5 mi (0 wells) -> 10 mi (2) -> 15 mi (10). 12-month volumes per 10,000 ft after the "
               "partial-first-month scaling.")
    hdr = ["Fit set", "Well", "API", "Operator", "Enverus interval", "First prod", "Lateral (ft)", "Distance (mi)",
           "12-mo oil (MBO / 10k ft)", "12-mo gas (MMcf / 10k ft)"]
    for j, h in enumerate(hdr):
        C.cell(5, 2 + j, h).font = Font(bold=True)
    for r, a in enumerate(dd.FIT_SET, start=6):
        nm, lat, first, o, g, iv, dist, op = dd.WELLS[a]
        oc, gc = bt.clean(o, g, dd.DAYS.get(a), first)
        k = 10000 / lat
        vals = ["Mowry 15 mi", nm, a, op, iv.title(), first, lat, dist, round(oc[:12].sum() * k / 1000),
                round(gc[:12].sum() * k / 1000)]
        for j, v in enumerate(vals):
            C.cell(r, 2 + j, v)
wb.save(f)
print(mode, "done")
