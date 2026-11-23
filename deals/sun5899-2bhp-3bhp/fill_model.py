"""Deal-specific workbook edits for Sun 5899 11-28 2BHP / 3BHP.
    python3 fill_model.py pre  <template copy.xlsx>   # strip + AFE notes, before write_model.py
    python3 fill_model.py post <model.xlsx>           # Summary / Comps text, after write_model.py
Recalculate after each step."""
import sys
import numpy as np
import openpyxl

mode, f = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(f)
DECK = ("NYMEX WTI strip Nov 20, 2026 (rescaled from the Nov 3 curve through Nov 20 settles; back end tapers to Nov 3 by Dec-30), "
        "HH Nov 3, 2026; flat after 2030")

if mode == "pre":
    P = wb["Prices"]
    old = [P.cell(6 + i, 4).value for i in range(400)]
    # NYMEX WTI settlements, Nov 20, 2026 (month index from Nov-26 = 0); months 0-2 use the Jan-27 contract
    anchors = {0: 83.41, 1: 83.41, 2: 83.41, 3: 82.40, 4: 81.10, 5: 79.84, 6: 78.68, 7: 77.63, 25: 68.39}
    xs = sorted(anchors); r = [anchors[k] / old[k] for k in xs]
    for i in range(400):
        if i <= 25: q = float(np.interp(i, xs, r))
        elif i <= 49: q = r[-1] + (1 - r[-1]) * (i - 25) / 24
        else: q = 1.0
        P.cell(6 + i, 4).value = round(float(old[i]) * q, 2)
    P["B3"] = DECK + ". Nov/Dec-26 use the Jan-27 contract. Flat decks follow the Summary C6 dropdown."
    wb["Inputs"]["D38"] = DECK
    S = wb["Summary"]
    S["D6"] = '=IF(C6="Strip","' + DECK + '","Flat "&C6&" (oil / gas)")'
    S["B3"] = ("Pure non-op economics at Chord's executed AFEs (AFE29659 / AFE29660, $9,664,884 each) plus 1/4 of the CTB and pad AFEs "
               "(AFE29692 / AFE29693), plus the premium. Kraken elected to participate 9/9/2026. Live case follows Inputs, the deck selector (C6) "
               "and the premium (C7). Type curve: 5-mile set (8 wells, low confidence), oil adjusted per Davis (peak pinned to the offsets, "
               "b fixed 0.9); 10-mile comparison curve from 44 wells.")
    W = wb["Wells"]
    W["T7"] = ("Enverus 20,713 ft U-turn, TVD 10,080 ft. Spud 9/6/26. AFE29659 executed by Kraken 9/9/26: drilling $3,965,135, completion "
               "$5,699,749; facilities = 1/4 of CTB AFE29692 ($4,907,000) + 1/4 of pad AFE29693 ($1,923,000). ASSUMPTION: closing Dec-26, "
               "frac Jan-27, first sales Mar-27")
    W["T8"] = "Enverus 21,658 ft U-turn, TVD 10,090 ft. Spud 9/7/26. AFE29660 executed 9/9/26, same costs as 2BHP. Same timing assumptions"
    W["G19"] = "Matches the executed AFEs: $11,372,384 gross per well x 12.5% x 2"
else:
    S, C, T = wb["Summary"], wb["Comps"], wb["5mi vs 10mi"]
    S["B79"] = ("- AFEs confirmed from the executed election packet (AFE29659/29660 at $9,664,884 each; CTB AFE29692 $4.907MM and pad "
                "AFE29693 $1.923MM split 1/4 per well = $1,707,500). Cost overruns are the remaining capital risk")
    S["B80"] = ("- Frac and first-sales dates (model: frac Jan-27, first sales Mar-27). Both wells spud early Sep-26 (DUCs); see the "
                "+3 mo delay rows for the cost of slippage")
    S["B81"] = "- Completion design: AFE plans ~21.8 MMlb proppant per well (~1,000-1,050 lb/ft), in line with the 5-mile offsets (~1,000 lb/ft)"
    S["B82"] = ("- Kraken elected to participate in both wells 9/9/2026 (executed ballots and AFEs). Still confirm title to the 12.5% WI / "
                "81.25% NRI (NRI not stated in the packet), any ORRI or lease burdens, and the JOA")
    S["B83"] = "- Gas marketing and NGL netback in Williams Co (Williston basin defaults; not yet checked against Chord JIBs or contracts)"
    for c in ("H607", "H609"): C[c] = "Chord executed AFE (election packet, 9/9/2026)"
    for c in ("H608", "H610"): C[c] = "Chord executed AFE + 1/4 of CTB ($4.907MM) and pad ($1.923MM) AFEs"
    C["D624"] = "Chord executed AFEs AFE29659 / AFE29660 (election packet)"
    T["B3"] = ("Prior comparison (Nov 22 run, Nov 3 deck, 46-well 10-mile set incl. 2-mile laterals), kept for reference. The live 10-mile "
               "comparison for this evaluation (44 wells, 3-mile+ laterals) is on Summary rows 48 and 57 and the TypeCurve charts.")
wb.save(f)
print(mode, "done")
