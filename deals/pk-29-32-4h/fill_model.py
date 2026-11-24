"""Deal-specific workbook edits for PK 29-32 4H-LL (Phoenix Operating, Middle Bakken, Divide Co ND).
    python3 fill_model.py pre  <template copy.xlsx>     # wells, inputs, strip, text; before write_model.py
    python3 fill_model.py post <model.xlsx>             # Comps 5mi tab, after write_model.py
Recalculate after each step."""
import datetime as dt
import importlib.util
import os
import sys

import openpyxl
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "toolkit"))
import build_tc as bt  # noqa: E402

_s = importlib.util.spec_from_file_location("dd", os.path.join(HERE, "deal_data.py"))
dd = importlib.util.module_from_spec(_s); _s.loader.exec_module(dd)

D = dt.datetime
NAME, OP, ZONE = "PK 29-32 4H-LL", "Phoenix Operating", "Middle Bakken"
TITLE = f"{NAME} ({OP}, {ZONE}, Divide Co ND)"
WI, LNRI = 0.01333858, 0.78
NRI = WI * LNRI
LAND, DRILL, COMP, PROD, FAC = 82500, 2873189, 2791000, 619000, 2946219
AFE = LAND + DRILL + COMP + PROD + FAC          # 9,311,908
CENTROID = (48.6475, -103.2325)
FIT_N = len(dd.FIT_SET)
DECK = ("NYMEX WTI strip Nov 20, 2026 (Nov 3 curve rescaled through the Nov 20 settles; latest deck in the repo), "
        "HH Nov 3, 2026; flat after 2030")
TIMING = ("Timing (ASSUMPTION): closing Dec-26 (election due ~30 days from Phoenix's Nov 5 delivery), drilling capex "
          "booked that month; completion and facilities Feb-27, first sales Mar-27. Enverus shows spud Oct 15, 2026; "
          "the proposal says on or about Nov 15, 2026.")
SRC = "Phoenix proposal and AFE N258DC (Nov 3, 2026)"

mode, f = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(f)

if mode == "pre":
    W = wb["Wells"]
    for k in range(10):
        r = 7 + k
        for col in "CDEFGIJKNOPSTUX":
            W[f"{col}{r}"] = None
        W[f"R{r}"] = 1
        W[f"Q{r}"] = "=Inputs!$C$12"
    note = (f"S29 / S32-T160N-R97W (2,560-ac DSU 28/29/32/33); {ZONE}. Enverus API 33-023-01894, lateral 9,989 ft, TVD "
            f"9,315 ft, MD 20,239 ft (AFE). WI {WI:.6%} per {SRC}; NRI delivered 78% of 8/8 (Davis) = {NRI:.6%}. AFE "
            f"${AFE:,.0f}: drilling column = drilling ${DRILL:,} + land / regulatory ${LAND:,}; completion ${COMP:,}; "
            f"facility column = facilities ${FAC:,} + artificial lift ${PROD:,}. {TIMING}")
    vals = dict(C=NAME, D="33-023-01894", E=9989, F=WI, G=LNRI, I=LAND + DRILL, J=COMP, K=FAC + PROD,
                N=D(2026, 12, 1), O=D(2027, 2, 1), P=D(2027, 3, 1), S="Spud (Enverus Oct 15, 2026)", T=note)
    for col, v in vals.items():
        W[f"{col}7"] = v
    for col in "NOP":
        W[f"{col}7"].number_format = "mmm-yy"
    W["F19"] = WI * AFE
    W["G19"] = f"Phoenix cash call: {WI:.6%} x AFE ${AFE:,.0f} = $124,207.66"

    I = wb["Inputs"]
    I["B2"] = TITLE + ": Non-Op Participation Model Inputs"
    I["B3"] = (f"Blue = hard-coded input. Black = formula. Yellow = key assumption to confirm with {OP} or the seller. "
               "ASSUMPTION = added by Osage Manka, not from the seller or Enverus.")
    I["B6"] = "Interest offered"
    I["D10"] = "Paid at closing (Dec-26) on top of the net capex; no premium asked, so C7 is solved to a 20% IRR"
    I["B11"] = f"Type curve multiplier (100% = 10 mi {ZONE} offsets base)"
    I["B13"] = f"Type curve (1 = 10 mi {ZONE} offsets base, 2 = Enverus fit-set P50 on our shape)"
    I["C18"] = "Working interest"
    I["C19"] = "Williston (ND)"
    I["C20"] = 0
    I["C21"] = None
    I["D12"] = "Default child haircut, 0% (fund standard). Override here or per well (Wells column Q)"
    I["D13"] = (f"Base = 1: {ZONE} offsets inside 10 mi ({FIT_N} wells; only 2 inside 5 mi with 12+ months). Enverus EUR "
                "fields are blank on this subscription, so 2 = base")
    I["D38"] = DECK

    src = openpyxl.load_workbook(os.path.join(HERE, "..", "friendly-skies", "Friendly Skies DJ Non-Op Model.xlsx"))
    P, SP = wb["Prices"], src["Prices"]
    for i in range(400):
        P.cell(6 + i, 4).value = SP.cell(6 + i, 4).value
        P.cell(6 + i, 5).value = SP.cell(6 + i, 5).value
    P["B3"] = DECK + ". Nov/Dec-26 use the Jan-27 contract. Flat decks follow the Summary C6 dropdown."

    S = wb["Summary"]
    S["B2"] = TITLE + ": Participate or Not"
    S["B3"] = (f"Non-op WI {WI:.6%} (NRI {NRI:.6%}, 78% lease NRI delivered) in one {ZONE} 2-mile well, S29 / S32-160N-97W, "
               f"operated by Phoenix (Phoenix Energy One in Enverus). AFE ${AFE:,.0f} gross, ${WI * AFE:,.0f} net. No "
               f"premium asked: Summary C7 is the premium (% of net capex) for a 20% IRR at the strip. {TIMING}")
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["B78"] = f"Confirm with the seller and {OP}"
    notes = [
        "- Premium: none given; C7 is solved to a 20% IRR at the strip",
        f"- WI {WI:.6%} is Phoenix's title estimate in the proposal ('may own'); 78% NRI delivered per Davis. Confirm title, "
        "any ORRI and the assignment's effective date vs the election deadline",
        f"- AFE N258DC ${AFE:,.0f} (drill ${DRILL:,}, complete ${COMP:,}, facilities ${FAC:,} incl. $955,000 frac-water "
        f"recovery, artificial lift ${PROD:,}, land ${LAND:,}); overruns not modeled (Inputs C59)",
        "- Timing: closing Dec-26, completion Feb-27, first sales Mar-27 (ASSUMPTION; Phoenix offsets ran 3 months spud to "
        "first sales)",
        "- Non-consent: ND risk penalty (NDCC 38-08-08) applies if not elected or not paid within 30 days",
        "- Williston basin defaults for pricing, costs and taxes; no Phoenix JIBs or check stubs seen",
        "- Type curve: 10-mile set (45 wells; 2 inside 5 mi). See the TypeCurve notes and Comps 5mi",
    ]
    for k in range(8):
        S[f"B{79 + k}"] = notes[k] if k < len(notes) else None

    T = wb["TypeCurve"]
    T["B2"] = f"{ZONE} Type Curve per 10,000 ft lateral (gross 8/8; oil bbl/month, gas Mcf/month)"
    T["B3"] = (f"Comp standard (Nov 23, 2026). Enverus Developer API (Enverus connector): {ZONE} horizontals whose lateral "
               f"midpoint is inside 10 mi of the DSU centroid {CENTROID}, first production 2023+, 12+ months by Sep 2026, "
               f"9,000-16,000 ft laterals, 900+ lb/ft: {FIT_N} wells (Phoenix 27, Kraken 10, Continental 7, Chevron 1). "
               "Only Stoltenberg 2-4HSL / 3-4H (Continental, 2.5 mi) pass inside 5 mi, so the radius was widened to 10 mi. "
               "Phoenix's nearest wells (Woods 35-26-23, Terry Nelson 34-27-22; 3.5-4.3 mi, 10-11 months) are checks only "
               "(Comps 5mi). Offset-frac shut-in months removed (deal_data.py); partial first months scaled on "
               "ProducingDays. Enverus EUR fields are blank on this subscription, so the guardrail is off.")
    for sh in ("Comps", "5mi vs 10mi"):
        if sh in wb.sheetnames:
            del wb[sh]
else:
    C = wb["Comps 5mi"]
    for row in C.iter_rows():
        for c in row:
            c.value = None
    C["B2"] = f"Comp selection: {TITLE}, DSU centroid {CENTROID[0]:.4f}, {CENTROID[1]:.4f}"
    C["B2"].font = Font(bold=True, size=14)
    C["B3"] = (f"Distance = offset lateral midpoint to the DSU centroid. Screen: {ZONE} horizontals (Enverus ENVInterval), "
               "first production 2023+ with 12+ months by Sep 2026, 9,000-16,000 ft, 900+ lb/ft. Fit = inside 10 mi (2 wells "
               "inside 5 mi). 12-month volumes per 10,000 ft (linear), after the cleaning in deal_data.py; partial first "
               "months scaled on ProducingDays.")
    hdr = ["Set", "Well", "API", "Operator", "Distance (mi)", "First prod", "Lateral (ft)", "12-mo oil (MBO / 10k ft)",
           "12-mo gas (MMcf / 10k ft)", "Months used"]
    for j, h in enumerate(hdr):
        C.cell(5, 2 + j, h).font = Font(bold=True)
    r = 6
    fit = sorted(dd.FIT_SET, key=lambda a: dd.DIST_MI[a])
    for lab, apis in (("Fit (10 mi)", fit), ("Check (<12 months)", dd.CHECK)):
        for a in apis:
            nm, lat, first, o, g = dd.WELLS[a][:5]
            oc, gc = bt.clean(o, g, dd.DAYS.get(a), first)
            k = 10000 / lat
            n = min(12, len(oc))
            vals = [lab + (" / 5 mi" if a in dd.FIVE_MI else ""), nm, a, dd.OPERATOR[a], dd.DIST_MI[a], first, lat,
                    round(oc[:n].sum() * k / 1000, 1), round(gc[:n].sum() * k / 1000), len(oc)]
            for j, v in enumerate(vals):
                C.cell(r, 2 + j, v)
            r += 1
    C.cell(r + 1, 2, "Check wells have fewer than 12 months; their 12-month columns sum the months available.")
wb.save(f)
print(mode, "done")
