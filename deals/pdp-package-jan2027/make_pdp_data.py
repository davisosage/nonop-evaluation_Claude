"""Builds pdp_data.py for the 128-well PDP package (Davis's Book1.xlsx, 2026-11-24; effective date Jan 2027).

Raw Enverus pulls (2026-11-24, Enverus connector) in raw/:
  package_well_list.xlsx / pkg.json   the package as received (API, lease, well no, operator, WI %, NRI %)
  subj_hdr.py                         subject headers (find_wells): interval, lateral, first production, SHL / BHL
  offsets_*.json                      offset_wells, 5 mi of each DJ area's lateral-midpoint centroid (nearest ~345 rows kept by the connector)
  wells_*.json                        headers of the hand-picked offsets for Benelli, Acacia, Gracie, CU Downs and Lee
  sel_files.json / sel_hand.json      the offset selection (screen below) and its headers
  prod_*.json                         production dataset (liquids, gas, ProducingDays) for subjects and offsets
Madison-Hamill offsets and their production come from deals/madison-hamill/raw (pulled the same day).

Offset screen: same zone group (DJ Niobrara A/B/C together, Codell; PRB Niobrara, Teapot; Bakken; Wolfcamp A), horizontal,
lateral >= 4,000 ft (DJ), lateral midpoint within 5 mi of the area centroid, first production at least 3 months before each
subject vintage in the area (12 months, and 2019+, for 2024-25 vintages), nearest 15 per vintage (20 for 2024-25 vintages).
Interests: Davis's figures are in percent (0.143686 = 0.143686%), stored as decimals.
Riley Fed 4776-15-22-15W MH (49-005-79263, no API in the list) has no first production in Enverus: left out (not PDP).
Lease NRI for the 8/8 economic limit: 80% (assumption), 75% for Madison-Hamill (Texas fee, as in deals/madison-hamill).
"""
import csv, glob, io, json, os, sys
H = os.path.dirname(os.path.abspath(__file__)); R = os.path.join(H, "raw")
sys.path.insert(0, R)
from subj_hdr import H as SH
MH = os.path.join(H, "..", "madison-hamill", "raw")

pkg = json.load(open(os.path.join(R, "pkg.json")))
BASIN = {"05": "DJ (CO)", "49": "Powder River (WY)", "42": "Permian (TX)", "33": "Williston (ND)"}
LEASE_AREA = {"MADISON-HAMILL 36A": "Madison-Hamill", "MADISON-HAMILL 36B": "Madison-Hamill", "MADISON-HAMILL 36C": "Madison-Hamill",
              "MADISON-HAMILL 36D": "Madison-Hamill"}
OFF_AREA_LEASES = {"Sullivan/Nelson/Carson": ["SULLIVAN", "NELSON", "CARSON"],
                   "Speak/Burge/FL Greens/Thermo/Carter/Tedford": ["SPEAK FED", "BURGE FED", "FL GREENS FEDERAL", "THERMO", "CARTER", "TEDFORD"],
                   "Dittmer": ["DITTMER"], "Bernhardt/Schneider": ["BERNHARDT", "SCHNEIDER HD"],
                   "Will/Dutch Lake/Barracuda/Whitetail/Triggerfish": ["WILL", "WILL TRUST", "DUTCH LAKE", "BARRACUDA 7-62", "WHITETAIL FED 7-62", "TRIGGERFISH 7-63"],
                   "Benelli": ["BENELLI 7-60"], "Acacia": ["ACACIA"], "Gracie": ["GRACIE FEE"], "CU Downs": ["CU DOWNS FED"], "Lee": ["LEE"]}
NICE = {"CU DOWNS FED": "CU Downs Fed", "FL GREENS FEDERAL": "FL Greens Federal", "SCHNEIDER HD": "Schneider HD"}
area_name = lambda lease: LEASE_AREA.get(lease, NICE.get(lease, lease.title()))

# ---- production (all pulls; later files win for duplicate months)
prod = {}
for f in sorted(glob.glob(os.path.join(R, "prod_*.json"))) + sorted(glob.glob(os.path.join(MH, "prod_*.json"))):
    t = json.load(open(f))[0]["text"].split("\n", 1)[1]
    for r in csv.DictReader(io.StringIO(t)):
        if not r.get("ProducingMonth"):
            continue
        d = r.get("ProducingDays")
        prod.setdefault(r["API_UWI_Unformatted"][:10], {})[r["ProducingMonth"][:7]] = (
            float(r["LiquidsProd_BBL"] or 0), float(r["GasProd_MCF"] or 0), int(float(d)) if d not in (None, "") else None)


def series(a):
    p = prod[a]
    months = sorted(p)
    first = min(m for m in months if p[m][0] + p[m][1] > 0); last = max(months)
    y, mo = int(first[:4]), int(first[5:7]); o, g, d = [], [], []
    while f"{y:04d}-{mo:02d}" <= last:
        v = p.get(f"{y:04d}-{mo:02d}", (0.0, 0.0, 0))
        o.append(round(v[0])); g.append(round(v[1])); d.append(v[2] if v[2] is not None else 0)
        mo += 1
        if mo > 12:
            y, mo = y + 1, 1
    return first, last, o, g, d


# ---- subjects
SUBJ, WB, WA, DAYS, missing, skipped = [], {}, {}, {}, [], []
for r in pkg:
    a = r["api"][:10]
    if not a:
        skipped.append(f"{r['lease']} {r['no']} (no API; not producing in Enverus)"); continue
    if a not in prod:
        missing.append(a); continue
    iv, lat, fp = SH[a][0], SH[a][1], SH[a][2]
    first, last, o, g, d = series(a)
    wi, nri = float(r["wi"] or 0) / 100, float(r["nri"] or 0) / 100
    lnri = 0.75 if a.startswith("42") else 0.80
    SUBJ.append((a, f"{r['lease']} {r['no']}", float(lat), first, o, g, iv, wi, nri, lnri, last))
    DAYS[a] = d
    WB[a] = BASIN[a[:2]]; WA[a] = area_name(r["lease"])

# ---- offsets: (api -> header, list of subject areas)
OFF, OA = {}, {}
sf, sh = json.load(open(os.path.join(R, "sel_files.json"))), json.load(open(os.path.join(R, "sel_hand.json")))
hdr = {**sf["hdr"], **sh["hdr"]}
for oarea, apis in list(sf["sel"].items()) + list(sh["sel"].items()):
    for a in apis:
        OA.setdefault(a, set()).update(area_name(l) for l in OFF_AREA_LEASES[oarea])
mh_sel = json.load(open(os.path.join(MH, "selection.json")))
mh_hdr = {x["API_UWI"].replace("-", "")[:10]: x for x in json.load(open(os.path.join(MH, "offsets_5mi.json")))}
for a in mh_sel["offsets"]:
    a = a[:10]; hdr[a] = mh_hdr[a]; OA.setdefault(a, set()).add("Madison-Hamill")
subj_apis = {s[0] for s in SUBJ}
for a in list(OA):
    if a in subj_apis or a not in prod:
        if a not in prod:
            missing.append(a)
        OA.pop(a); continue
    h = hdr[a]
    first, last, o, g, d = series(a)
    OFF[a] = (h["WellName"], float(h["LateralLength_FT"] or 0), first, o, g, h["ENVInterval"], float(h.get("DistanceMi") or 0))
    DAYS[a] = d

LAST = max(s[10] for s in SUBJ)
# Lee 1HS oil (12 months, R2 ~0 on its own history, flat fit 54% above the offsets over 24 months): the offsets are the
# better guide. Anchor at the last three months (17,815 bbl/mo), b 1.11 (offsets' shape), Da 123.8% solved so the next 24
# months equal the offsets' anchored P50 (213,755 bbl). Gas left as fitted (already below the offsets).
LEE_OVR = {"o_q": 17815, "o_da": 1.2377, "o_b": 1.1142}
# inactive wells: under 10 producing days a month on average over the last six reported months -> no forecast (multiplier 0)
OVR, INACTIVE = {}, []
for a, name, *_ in SUBJ:
    d6 = DAYS[a][-6:]
    if sum(d6) / len(d6) < 10:
        OVR[a] = {"mult": 0.0}; INACTIVE.append(f"{name} ({sum(d6)} producing days in its last 6 reported months)")
L = [f"# Generated by make_pdp_data.py (Enverus, pulled 2026-11-24; production through {LAST}). {len(SUBJ)} subject wells, {len(OFF)} offsets.",
     "# Mixed package: royalty / ORRI wells (WI 0) in the DJ and Powder River, WI wells in the Permian (Madison-Hamill) and Williston (Lee).",
     f"# Left out: {'; '.join(skipped) or 'none'}. Missing from the production pull: {missing or 'none'}.",
     'DEAL = dict(name="PDP Package (128 wells, Jan 2027)", basin="DJ (CO)", deal_type="Working interest", effective="2027-01",',
     '            asking_price=0, operator="Multiple (SM, Oxy, Bison, Chevron, Anschutz, Devon, XTO, Petro-Hunt)")',
     f'LAST_MONTH = "{LAST}"',
     "# api: (well, lateral_ft, first_month, oil_bbl[], gas_mcf[], zone, wi, nri, lease_nri)", "SUBJECT = {"]
for a, name, lat, first, o, g, iv, wi, nri, lnri, last in SUBJ:
    L.append(f'"{a}": ({name!r}, {lat}, "{first}", {o}, {g}, {iv!r}, {wi!r}, {nri!r}, {lnri}),')
L += ["}", "# api: (well, lateral_ft, first_month, oil_bbl[], gas_mcf[], zone, distance_mi)", "OFFSETS = {"]
for a, v in OFF.items():
    L.append(f'"{a}": ({v[0]!r}, {v[1]}, "{v[2]}", {v[3]}, {v[4]}, {v[5]!r}, {v[6]}),')
L += ["}", "DAYS = {" + ", ".join(f'"{a}": {d}' for a, d in DAYS.items()) + "}",
      'ZONE_GROUPS = {"NIOBRARA A": "NIOBRARA", "NIOBRARA B": "NIOBRARA", "NIOBRARA C": "NIOBRARA", "FORT HAYS": "CODELL", "CARLILE": "CODELL"}',
      "WELL_BASIN = " + json.dumps(WB), "WELL_AREA = " + json.dumps(WA),
      "OFFSET_AREAS = " + json.dumps({a: sorted(v) for a, v in OA.items() if a in OFF}),
      "# Inactive (shut in or near shut in): " + "; ".join(INACTIVE) + ". Forecast multiplier 0 (no value); confirm status with the operator.",
      "OVERRIDES = " + json.dumps({**OVR, "3306105494": LEE_OVR})]
open(os.path.join(H, "pdp_data.py"), "w").write("\n".join(L) + "\n")
print("inactive:", INACTIVE)
print(f"{len(SUBJ)} subjects, {len(OFF)} offsets, last month {LAST}; skipped {skipped}; missing {missing}")
from collections import Counter
print(Counter(WA.values()))
print("subjects by last month", Counter(s[10] for s in SUBJ))
