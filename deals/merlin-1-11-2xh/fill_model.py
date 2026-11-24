"""Deal-specific workbook edits for the TNE_OK_2026-11-20 package (Mewbourne / Nadel & Gussman, Anadarko Basin, western OK).
Each well is its own evaluation (Davis, Nov 23, 2026). This same file sits in each of the five deal folders and picks its
deal from the folder name.
    python3 fill_model.py pre  <template copy.xlsx>     # wells, inputs, strip, text; before write_model.py
    python3 fill_model.py post <model.xlsx>             # Comps 5mi tab, after write_model.py
Recalculate after each step."""
import datetime as dt
import os
import sys

import openpyxl
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "toolkit"))
import build_tc as bt  # noqa: E402

SLUG = os.path.basename(HERE)
D = dt.datetime
DECK = ("NYMEX WTI strip Nov 20, 2026 (Nov 3 curve rescaled through the Nov 20 settles; latest deck in the repo), "
        "HH Nov 3, 2026; flat after 2030")
TIMING = ("Timing per Davis: closing Jan-27 (premium and drilling capex booked that month), spud Mar-27, completion and "
          "facilities Sep-27 (ASSUMPTION), first sales Oct-27.")
SRC = "Seller sheet TNE_OK_2026-11-20.xlsx"
DEALS_CENTROID = {"io-23-26": (35.6678, -99.6109), "bogges-29-20": (35.5815, -99.563), "lomas-7-18-1h": (35.7832, -99.3675),
                  "lomas-7-18-2h": (35.7832, -99.3675), "merlin-1-11-2xh": (35.6247, -99.183)}

DEALS = {
    "io-23-26": dict(
        name="IO 23/26", op="Mewbourne Oil", zone="Tonkawa", county="Roger Mills Co OK", str_="S26-T14N-R23W (1280-ac unit, 23/26)",
        api="35-129-00166", lat=9950, wi=0.0703125, nri=0.052734375, dr=5389500, cp=2852100, fac=643000, status="Permit",
        lat_note=("Enverus permit IO 26-23 PA 2HT (35-129-00166, approved Oct 19, 2026; Tonkawa / Low Tonkawa, TVD 8,765 ft). "
                  "Permit gives no lateral; 9,950 ft from SHL to BHL (2-section unit)"),
        sale="90 net acres of 1,280"),
    "bogges-29-20": dict(
        name="Bogges 29/20", op="Mewbourne Oil", zone="Cherokee", county="Roger Mills Co OK", str_="S20-T13N-R22W (1280-ac unit, 29/20)",
        api="", lat=10000, wi=0.01, nri=0.0075, dr=5389500, cp=2852100, fac=643000, status="Permit (not yet in Enverus)",
        lat_note=("Not yet permitted in Enverus (Mewbourne's BOGGS 19-30 DM 1H is a different unit, 13N-23W). ASSUMPTION: "
                  "10,000 ft lateral (2-section unit). The seller's AFE equals the IO 23/26 AFE line for line; confirm"),
        sale="12.8 net acres of 1,280"),
    "lomas-7-18-1h": dict(
        name="Lomas 7/18 DM #1H", op="Mewbourne Oil", zone="Cherokee", county="Custer Co OK", str_="S7-T15N-R20W (1280-ac unit, 7/18)",
        api="35-039-00137", lat=9873, wi=0.01, nri=0.0075, dr=5932800, cp=4315900, fac=593500, status="Proposal signed (permitted)",
        lat_note=("Enverus permit LOMAS 7-18 CN 1H (35-039-00137, approved Nov 12, 2026; Cherokee, TVD 10,928 ft), permitted "
                  "lateral 9,873 ft. The seller calls it DM #1H; the two Lomas permits are CN 1H and BO 2H"),
        sale="12.8 net acres of 1,280"),
    "lomas-7-18-2h": dict(
        name="Lomas 7/18 DM #2H", op="Mewbourne Oil", zone="Cherokee", county="Custer Co OK", str_="S7-T15N-R20W (1280-ac unit, 7/18)",
        api="35-039-00138", lat=9828, wi=0.01, nri=0.0075, dr=5932800, cp=4315900, fac=593500, status="Proposal signed (permitted)",
        lat_note=("Enverus permit LOMAS 7-18 BO 2H (35-039-00138, approved Nov 12, 2026; Cherokee, TVD 10,928 ft), permitted "
                  "lateral 9,828 ft. The seller calls it DM #2H; the two Lomas permits are CN 1H and BO 2H"),
        sale="12.8 net acres of 1,280"),
    "merlin-1-11-2xh": dict(
        name="Merlin 1-11-2XH", op="Nadel & Gussman", zone="Cherokee", county="Custer Co OK", str_="S2 / S11-T13N-R19W (1280-ac unit)",
        api="", lat=10000, wi=0.02890625, nri=0.021621875, dr=4960410, cp=6310750, fac=0, status="Waiting on pooling order",
        lat_note=("Not yet permitted in Enverus. ASSUMPTION: 10,000 ft lateral (2-section cross-unit XH; N&G's offsets run "
                  "9,800-10,500 ft). Seller AFE carries no facility line; confirm whether facilities are in the completion AFE"),
        sale="37 net acres of 1,280"),
}
d = DEALS[SLUG]
LNRI = d["nri"] / d["wi"]
spec_ = __import__("importlib.util").util
_s = spec_.spec_from_file_location("dd", os.path.join(HERE, "deal_data.py")); dd = spec_.module_from_spec(_s); _s.loader.exec_module(dd)
FIT_N = len(dd.FIT_SET)
RADIUS = {"io-23-26": "20 mi", "bogges-29-20": "15 mi", "merlin-1-11-2xh": "10 mi"}.get(SLUG, "5 mi")
TITLE = f"{d['name']} ({d['op']}, {d['zone']}, {d['county']})"

# Oklahoma (Anadarko) inputs: no Anadarko column on the Basins tab, so every line is overridden in Inputs column E.
OK_IN = {  # Inputs row: (value, note)
    41: (-1.50, "ASSUMPTION, Oklahoma: Oklahoma sweet at the lease vs NYMEX WTI before gathering; confirm on Mewbourne / N&G check stubs"),
    42: (1.50, "ASSUMPTION, Oklahoma: trucking / gathering to Cushing-area purchasers"),
    43: (-0.55, "Oklahoma: Midcon residue basis (NGPL Midcon / ONEOK / Panhandle prompt basis -$0.44 to -$0.57, Sep 2026)"),
    44: (0.20, "ASSUMPTION, Oklahoma: wet Anadarko gas, plant shrink + fuel"),
    45: (100, "ASSUMPTION, Oklahoma: rich Cherokee / Tonkawa gas; confirm on the gas contract statements"),
    46: (0.28, "ASSUMPTION, Oklahoma: Conway-priced NGL after T&F"),
    47: (0.75, "ASSUMPTION, Oklahoma: Midcon gathering and processing fee"),
    50: (0.052, "Oklahoma gross production tax 5% for the first 36 months of production (wells spudded after Jul 2018) + 0.2% excise / conservation"),
    51: (0.072, "Oklahoma gross production tax 7% after 36 months + 0.2% excise / conservation"),
    52: (8000, "ASSUMPTION, Oklahoma: fixed LOE per well per month (8/8); no Mewbourne / N&G JIBs seen"),
    53: (0.35, "ASSUMPTION, Oklahoma"),
    54: (3.00, "ASSUMPTION, Oklahoma"),
    55: (1.5, "ASSUMPTION: Enverus water is blank for these Oklahoma offsets"),
    56: (1.00, "ASSUMPTION, Oklahoma: disposal"),
    57: (125000, "ASSUMPTION"),
}

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
    note = (f"{d['str_']}; {d['zone']}. {d['lat_note']}. WI {d['wi']:.6%}, NRI {d['nri']:.6%} ({d['sale']}) and AFE from the "
            f"{SRC}; lease NRI = NRI / WI = {LNRI:.4%}. {TIMING}")
    vals = dict(C=d["name"].upper(), D=d["api"] or None, E=d["lat"], F=d["wi"], G=LNRI, I=d["dr"], J=d["cp"], K=d["fac"],
                N=D(2027, 1, 1), O=D(2027, 9, 1), P=D(2027, 10, 1), S=d["status"], T=note)
    for col, v in vals.items():
        W[f"{col}{7}"] = v
    for col in "NOP":
        W[f"{col}7"].number_format = "mmm-yy"
    W["F19"] = d["wi"] * (d["dr"] + d["cp"] + d["fac"])
    W["G19"] = f"Seller sheet: WI {d['wi']:.6%} x AFE ${d['dr'] + d['cp'] + d['fac']:,.0f}"

    I = wb["Inputs"]
    I["B2"] = TITLE + ": Non-Op Participation Model Inputs"
    I["B3"] = (f"Blue = hard-coded input. Black = formula. Yellow = key assumption to confirm with {d['op']} or the seller. "
               "ASSUMPTION = added by Osage Manka, not from the seller or Enverus.")
    I["B6"] = "Interest offered"
    I["D10"] = "Paid at closing (Jan-27) on top of the net capex; no premium asked, so C7 is solved to a 20% IRR"
    I["B11"] = f"Type curve multiplier (100% = {RADIUS} {d['zone']} offsets base)"
    I["B13"] = f"Type curve (1 = {RADIUS} {d['zone']} offsets base, 2 = Enverus fit-set P50 on our shape)"
    I["C18"] = "Working interest"
    I["C19"] = "Permian (TX)"
    I["D19"] = ("No Anadarko / Oklahoma column on the Basins tab: every price, cost and tax line below is overridden in column E "
                "with Oklahoma values, so the Permian defaults are not used")
    I["C20"] = 0
    I["C21"] = None
    I["D12"] = ("Default child haircut, 0% (fund standard). Override here or per well (Wells column Q)"
                + ("; the Lomas #1H and #2H are a 2-well unit (each evaluated on its own)" if "lomas" in SLUG else ""))
    I["D13"] = (f"Base = 1: {d['zone']} offsets inside {RADIUS} ({FIT_N} wells"
                + (", LOW CONFIDENCE" if FIT_N < 15 else "") + "). Enverus EUR fields are blank on this subscription, so 2 = base")
    I["D38"] = DECK
    for r, (v, n) in OK_IN.items():
        I[f"E{r}"] = v
        I[f"D{r}"] = f'=IF(ISNUMBER($E{r}),"{n}","Basin default ("&$C$19&"); see the Basins tab")'

    src = openpyxl.load_workbook(os.path.join(HERE, "..", "friendly-skies", "Friendly Skies DJ Non-Op Model.xlsx"))
    P, SP = wb["Prices"], src["Prices"]
    for i in range(400):
        P.cell(6 + i, 4).value = SP.cell(6 + i, 4).value
        P.cell(6 + i, 5).value = SP.cell(6 + i, 5).value
    P["B3"] = DECK + ". Nov/Dec-26 use the Jan-27 contract. Flat decks follow the Summary C6 dropdown."

    S = wb["Summary"]
    S["B2"] = TITLE + ": Participate or Not"
    S["B3"] = (f"Non-op WI {d['wi']:.6%} (NRI {d['nri']:.6%}) in one {d['zone']} well, {d['str_']}, operated by {d['op']}. "
               f"One of five wells in the TNE_OK_2026-11-20 package, each evaluated on its own. No premium asked: Summary C7 is "
               f"the premium (% of net capex) for a 20% IRR at the strip. {TIMING}")
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["B78"] = f"Confirm with the seller and {d['op']}"
    notes = [
        "- Premium: none given; C7 is solved to a 20% IRR at the strip",
        f"- AFE and interests as listed on the seller sheet (gross ${d['dr'] + d['cp'] + d['fac']:,.0f}); not checked against a signed AFE",
        f"- Lateral: {d['lat']:,} ft; " + ("permitted in Enverus" if d["api"] else "ASSUMPTION, no permit in Enverus yet"),
        "- Timing: closing Jan-27, spud Mar-27, first sales Oct-27 (Davis); completion Sep-27 assumed. Drilling capex is booked "
        "at closing, two months before spud (slightly conservative)",
        "- Oklahoma pricing, costs and taxes are Osage Manka assumptions (Inputs column E); no operator JIBs or check stubs seen",
        "- Type curve: young offsets (most first produced 2025); see the TypeCurve notes for the radius, b source and the gas b "
        "at the 0.05 floor",
    ]
    for k in range(8):
        S[f"B{79 + k}"] = notes[k] if k < len(notes) else None

    T = wb["TypeCurve"]
    T["B2"] = f"{d['zone']} Type Curve per 10,000 ft lateral (gross 8/8; oil bbl/month, gas Mcf/month)"
    head = (f"Comp standard (Nov 23, 2026). Enverus Developer API (Enverus connector): {d['zone']} horizontals whose lateral "
            f"midpoint is inside {RADIUS} of the DSU centroid {tuple(round(x, 4) for x in DEALS_CENTROID[SLUG])}, 12+ months "
            f"of production by Aug 2026: {FIT_N} wells")
    tail = {
        "io-23-26": (". No Tonkawa well with 12+ months inside 10 mi; widened to 20 mi (Sanguine, King, Charter Oak; mostly "
                     "1-mile laterals normalized to 10,000 ft). LOW CONFIDENCE. Excluded: Smith 30X31 #1HB / #2HC (Diversified, "
                     "2018, ~940 lb/ft). Strong City 1-30H19 (Sanguine, 1.6 mi, 10 months) is a check only. b fitted on the 3 "
                     "24+ month wells in the set; gas b sits at the 0.05 floor."),
        "bogges-29-20": (". Only Red Rooster (Mewbourne) has 12+ months inside 5 mi; widened to 15 mi (Mewbourne, Crawley, "
                         "Devon). No 24+ month well in the set: b fitted on 4 24+ month Cherokee wells in Roger Mills / Custer "
                         "(Sage Dragon, Mallard 1HX, Keahey x2) and applied to this set's mean; gas b sits at the 0.05 floor. "
                         "The 15-mile set leans east to oilier Custer wells; the nearest wells (Red Rooster, Grace Katelyn) are "
                         "gassier."),
        "merlin-1-11-2xh": (". 3 wells with 12+ months inside 5 mi (Osprey, Quetzal, Topaz); widened to 10 mi (N&G, Devon, "
                            "Duncan, Mewbourne). LOW CONFIDENCE. b fitted on 4 24+ month Cherokee wells in Roger Mills / Custer "
                            "(Sage Dragon, Mallard 1HX, Keahey x2); gas b sits at the 0.05 floor."),
    }.get(SLUG, (" (Mewbourne, Crawley). LOW CONFIDENCE; 10-mile comparison curve (19 wells). Only Sage Dragon has 24+ "
                 "months: b fitted on 4 24+ month Cherokee wells in Roger Mills / Custer (Sage Dragon, Mallard 1HX, Keahey x2); "
                 "gas b sits at the 0.05 floor."))
    T["B3"] = (head + tail + " Oklahoma rows: Mar-26 gas gaps filled from neighbouring months, shut-in months removed (see "
               "deal_data.py). Enverus EUR fields are blank on this subscription, so the guardrail is off.")
    for sh in ("Comps", "5mi vs 10mi"):
        if sh in wb.sheetnames:
            del wb[sh]
else:
    C = wb["Comps 5mi"]
    for row in C.iter_rows():
        for c in row:
            c.value = None
    C["B2"] = f"Comp selection: {TITLE}, DSU centroid {DEALS_CENTROID[SLUG][0]:.4f}, {DEALS_CENTROID[SLUG][1]:.4f}"
    C["B2"].font = Font(bold=True, size=14)
    C["B3"] = (f"Distance = offset lateral midpoint to the DSU centroid. Screen: {d['zone']} horizontals (Enverus ENVInterval), "
               f"first production 2018+ with 12+ months by Aug 2026. Fit = inside {RADIUS} (radius rule: widened from 5 mi "
               "until 6+ wells). 12-month volumes per 10,000 ft (linear), after the cleaning in deal_data.py.")
    hdr = ["Set", "Well", "API", "First prod", "Lateral (ft)", "12-mo oil (MBO / 10k ft)", "12-mo gas (MMcf / 10k ft)",
           "Months used"]
    for j, h in enumerate(hdr):
        C.cell(5, 2 + j, h).font = Font(bold=True)
    r = 6
    sets = [("Fit", dd.FIT_SET), ("10-mile comparison", [a for a in dd.FIT_SET_10MI if a not in dd.FIT_SET]),
            ("b donor (24+ months)", [a for a in getattr(dd, "B_SET", []) if a not in dd.FIT_SET])]
    for lab, apis in sets:
        for a in apis:
            nm, lat, first, o, g = dd.WELLS[a][:5]
            oc, gc = bt.clean(o, g, None, first)
            k = 10000 / lat
            vals = [lab, nm, a, first, lat, round(oc[:12].sum() * k / 1000, 1), round(gc[:12].sum() * k / 1000), len(oc)]
            for j, v in enumerate(vals):
                C.cell(r, 2 + j, v)
            r += 1
wb.save(f)
print(SLUG, mode, "done")
