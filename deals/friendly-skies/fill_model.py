"""Deal-specific workbook edits for Friendly Skies / 747 3-65 pad (Bison IV Operating, Niobrara A/B/C, Adams Co CO), Minerals.
    python3 fill_model.py pre  <template copy.xlsx> <purchase price $>   # wells, inputs, strip, text; before write_model.py
    python3 fill_model.py post <model.xlsx>                             # comps sheet, after write_model.py
Run from toolkit/. Recalculate after each step.

16 wells (permitted 8-5-1 excluded), template holds 10 rows: wells that share a bench and lateral direction are grouped into
one row. Each row carries the per-well NRI (3.783899%) and the well count in the Well TC factor (Wells R), so volumes and the
price split scale with the count (laterals inside a group differ by <= 86 ft)."""
import datetime as dt
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
UNIT_NRI = 0.03783899
DECK = ("NYMEX WTI strip Nov 20, 2026 (Nov 3 curve rescaled through the Nov 20 settles; latest deck in the repo), "
        "HH Nov 3, 2026; flat after 2030")
TITLE = "Friendly Skies / 747 3-65 pad: 16 Niobrara A/B/C wells (Bison IV Operating; Adams Co CO), minerals"
TIMING = ("Timing per Davis: first production Jul-27 for every well (all drilled / DUC, spud Apr-May 2026). "
          "Closing assumed Dec-26.")
# group: (row name, [(api, well, lateral)], zone, status)
GROUPS = [
    ("747 3-65 Niobrara B west (7 wells)", [("05-001-10674", "8-10-6", 17788), ("05-001-10671", "8-3-3", 17804),
        ("05-001-10684", "8-10-14", 17731), ("05-001-10677", "8-10-10", 17793), ("05-001-10682", "8-10-16", 17799),
        ("05-001-10669", "8-3-1", 17717), ("05-001-10670", "8-3-2", 17751)], "Niobrara B", "DUC",
     "west-running ~3.4-mile laterals"),
    ("747 3-65 8-5-4 (Niobrara B, north U-turn)", [("05-001-10673", "8-5-4", 15439)], "Niobrara B", "Drilled",
     "U-turn lateral north of the pad"),
    ("747 3-65 8-8-13 (Niobrara B, south U-turn)", [("05-001-10683", "8-8-13", 18045)], "Niobrara B", "DUC",
     "U-turn lateral south of the pad"),
    ("747 3-65 8-9-9 (Niobrara B, east)", [("05-001-10679", "8-9-9", 8753)], "Niobrara B", "DUC",
     "east-running lateral; Enverus 8,753 ft (MD 16,989 ft); the seller header shows 15,826 ft, which does not fit its 17,109 ft MD"),
    ("747 3-65 Niobrara A west (4 wells)", [("05-001-10678", "8-10-12", 17775), ("05-001-10681", "8-10-15", 17811),
        ("05-001-10672", "8-3-5", 17790), ("05-001-10675", "8-10-8", 17726)], "Niobrara A", "DUC",
     "west-running ~3.4-mile laterals"),
    ("747 3-65 8-5-7 (Niobrara C, north U-turn)", [("05-001-10676", "8-5-7", 15344)], "Niobrara C", "DUC",
     "U-turn lateral north of the pad"),
    ("747 3-65 8-8-11 (Niobrara C, south U-turn)", [("05-001-10680", "8-8-11", 18009)], "Niobrara C", "DUC",
     "U-turn lateral south of the pad"),
]

if mode == "pre":
    price = float(sys.argv[3])
    W = wb["Wells"]
    for k in range(10):
        r = 7 + k
        for col in "CDEFGIJKNOPSTUX":
            W[f"{col}{r}"] = None
        W[f"R{r}"] = 1
        W[f"Q{r}"] = "=Inputs!$C$12"
    for k, (nm, wells, zone, st, desc) in enumerate(GROUPS):
        r = 7 + k
        n = len(wells)
        lat = round(sum(w[2] for w in wells) / n)
        apis = ", ".join(f"{w[1]} {w[0]}" for w in wells)
        note = (f"{n} well{'s' if n > 1 else ''} ({apis}); {zone} per Enverus ENVInterval; {desc}. Lateral {lat:,} ft "
                f"(Enverus{', average' if n > 1 else ''}). Mineral NRI 3.783899% per well; Well TC factor = {n} (well count). {TIMING}")
        vals = dict(C=nm, D=wells[0][0], E=lat, F=0, G=0.8, I=0, J=0, K=0, N=D(2026, 12, 1), O=D(2027, 6, 1),
                    P=D(2027, 7, 1), R=n, S=st, T=note, U=UNIT_NRI, X=zone)
        for col, v in vals.items():
            W[f"{col}{r}"] = v
        for col in "NOP":
            W[f"{col}{r}"].number_format = "mmm-yy"
    W["F19"] = None
    W["G19"] = None

    I = wb["Inputs"]
    I["B2"] = TITLE + ": Minerals Model Inputs"
    I["B3"] = ("Blue = hard-coded input. Black = formula. Yellow = key assumption to confirm with the seller or Bison. "
               "ASSUMPTION = added by Osage Manka, not from the seller or Enverus.")
    I["B6"] = "Interest offered"
    I["C18"] = "Minerals"
    I["C19"] = "DJ (CO)"
    I["C20"] = price
    I["D20"] = ("No asking price given. Set to the most we could pay for a 20% IRR at the strip (solved by write_model.py / "
                "engine.max_price); split across the rows by mineral NRI x lateral x well count")
    I["C21"] = D(2026, 12, 1)
    I["D21"] = "ASSUMPTION: closing Dec-26"
    I["B11"] = "Type curve multiplier (100% = 5-mile zone curves, modern completions)"
    I["B13"] = "Type curve (1 = 5-mile zone curves, 2 = Enverus fit-set P50 on our shape)"
    I["D12"] = ("Default child haircut, 0% (fund standard). New pad (16 wells modeled); SM / Crestone wells 1.7-5 mi away. Override here "
                "or per row (Wells column Q)")
    I["D13"] = ("Base = 1: Niobrara B (TypeCurve tab) and Niobrara A / C on 'Zone Curves', each from its own 5-mile modern-"
                "completion offsets (9 / 12 / 12 wells, LOW CONFIDENCE). Enverus EUR fields are blank on this subscription, "
                "so 2 = base")
    I["D38"] = DECK

    src = openpyxl.load_workbook(os.path.join(HERE, "..", "mystical-fed-wrc", "Mystical Fed Powder River Non-Op Model.xlsx"))
    P, SP = wb["Prices"], src["Prices"]
    for i in range(400):
        P.cell(6 + i, 4).value = SP.cell(6 + i, 4).value
        P.cell(6 + i, 5).value = SP.cell(6 + i, 5).value
    P["B3"] = DECK + ". Nov/Dec-26 use the Jan-27 contract. Flat decks follow the Summary C6 dropdown."

    S = wb["Summary"]
    S["B2"] = TITLE + ": Buy or Not"
    S["B3"] = ("Mineral / royalty interest, unit NRI 3.783899% in each of 16 Bison IV Niobrara wells on the 747 3-65 pad (10 B, "
               "4 A, 2 C), grouped into 7 rows. The permitted 8-5-1 is excluded (Davis). Minerals: no capex or LOE; revenue at the mineral NRI less production taxes; "
               "purchase price paid at closing (Dec-26). No asking price given: Inputs C20 is set to the most we could pay for "
               "a 20% IRR at the strip. Live case follows Inputs, the deck selector (C6) and the price (Inputs C20).")
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["B78"] = "Confirm with the seller and Bison IV Operating"
    notes = [
        "- Asking price: none given; Inputs C20 = the max purchase price for a 20% IRR at the strip",
        "- Unit NRI 3.783899% applied to all 16 wells as given; confirm one unit covers the west laterals, the north / south "
        "U-turns and the east 8-9-9, or get NRI by well",
        "- Timing: first production Jul-27 for all 16 wells (Davis); Bison frac schedule not seen. The permitted 8-5-1 is excluded",
        "- Royalty deducts: modeled cost-free (gross proceeds less taxes); Colorado leases can bear post-production costs, "
        "check the lease / division order",
        "- Type curves: 5-mile modern-completion offsets are young (most first produced 2025); b for Niobrara A and B comes "
        "from their 10-mile sets; Niobrara B gas b sits at the 2.0 cap",
        "- Gas / NGL netback, CO taxes: DJ basin defaults used, not checked against Bison check stubs",
    ]
    for k in range(8):
        S[f"B{79 + k}"] = notes[k] if k < len(notes) else None

    T = wb["TypeCurve"]
    T["B2"] = "Niobrara B Type Curve per 10,000 ft lateral (zone 1; gross 8/8; oil bbl/month, gas Mcf/month)"
    T["B3"] = ("Comp standard (Nov 23, 2026). Enverus Developer API (Enverus connector): Niobrara B horizontals inside 5 mi of the "
               "747 3-65 DSU centroid (39.809, -104.715), modern completions (~1,150 lb/ft, first production 2024-06 to "
               "2025-06): 9 wells (SM Energy / Crestone Bijou, Invicta, King, Blue pads). LOW CONFIDENCE (< 15 wells); only 1 "
               "well has 24+ months, so b is fitted on the 40 24+ month wells of the 10-mile Niobrara B set (51 wells). "
               "Excluded: 2018-2021 legacy completions (590-730 lb/ft, shut in Mar-Sep 2024). Enverus EUR fields are blank on "
               "this subscription, so the guardrail is off. Niobrara A and C curves are on 'Zone Curves'.")
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
    C["B2"] = "Comp selection: 747 3-65 pad (Bison IV), DSU centroid 39.809, -104.715, Adams Co CO"
    C["B2"].font = Font(bold=True, size=14)
    C["B3"] = ("Distance = offset lateral midpoint to the DSU centroid. Screen: Niobrara A / B / C horizontals, first production "
               "2022-01 to 2025-08 (12+ months by 2026-08), modern completions. Fit = inside 5 mi; the 10-mile sets are the "
               "comparison curves. Excluded inside 5 mi: 2018-2021 legacy completions (Schuh, King 3xH, Bear 3xH, Blue 1H, "
               "Aspen 2DH) and Quicksilver / Aspen 15-14-13 (under 12 months). 12-month volumes per 10,000 ft (linear).")
    hdr = ["Zone", "In 5-mi fit", "Well", "API", "Operator", "First prod", "Lateral (ft)", "Proppant (lb/ft)", "Distance (mi)",
           "12-mo oil (MBO / 10k ft)", "12-mo gas (MMcf / 10k ft)"]
    for j, h in enumerate(hdr):
        C.cell(5, 2 + j, h).font = Font(bold=True)
    r = 6
    for zn, z in dd.ZONES.items():
        for a in sorted(set(z["fit"]) | set(z["fit_10mi"]), key=lambda a: dd.WELLS[a][6]):
            nm, lat, first, o, g, iv, dist, op, prop = dd.WELLS[a]
            oc, gc = bt.clean(o, g, dd.DAYS.get(a), first)
            k = 10000 / lat
            vals = [zn, "Yes" if a in z["fit"] else "10-mi only", nm, a, op, first, lat, prop, dist,
                    round(oc[:12].sum() * k / 1000, 1), round(gc[:12].sum() * k / 1000)]
            for j, v in enumerate(vals):
                C.cell(r, 2 + j, v)
            r += 1
wb.save(f)
print(mode, "done")
