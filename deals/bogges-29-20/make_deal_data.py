"""Build deal_data.py for the five Mewbourne / Nadel & Gussman Anadarko wells (TNE_OK_2026-11-20 package).
Raw Enverus rows: raw/enverus_production_2026-11-23.csv  (api10, gas_mcf, oil_bbl, YYYY-MM), pulled 2026-11-23, last month 2026-08."""
import csv, math, os, sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, "raw", "enverus_production_2026-11-23.csv")
REPO = os.path.join(HERE, "..", "..")

META = {  # api10: (name, lateral ft, distance note)
    # Cherokee, around Lomas 7/18 (15N-20W)
    "3503922644": ("MOOREWOOD 30/19 PA 1H", 9637), "3503922669": ("QUARTERMASTER 28/21 PA 1H", 9797),
    "3512924183": ("CANDACE 36/25 PA 1H", 10069), "3503922645": ("STRICKLAN 31/6 AP 1H", 9955),
    "3512924147": ("SHANNON 34/27 PA 1H", 9762), "3503922627": ("SAGE DRAGON 32/5 AP 1H", 9655),
    "3512924175": ("DONITA 35/2 BO 1H", 9681), "3512924174": ("MONTE 26/23 OB 1H", 9604),
    "3503922667": ("WELDON 34/27 PA 1H", 9712), "3503922668": ("BRITTON 33/4 AP 1H", 9812),
    "3512924149": ("VIVIAN RUTH 1-28-33CHX", 9214),
    "3512924184": ("DONNY 1/12 AP 1H", 9804), "3503922642": ("MOORE 35/2 AP 1H", 9830),
    "3503922670": ("HYMAN 3/10 AP 1H", 9726), "3503922641": ("LEROY 36/1 DM 1H", 10018),
    "3512924189": ("CALEB 14/11 MD FEE COM 1H", 10434), "3503922658": ("EMERALD 1-5HX", 10339),
    "3512924173": ("HUGHES 1-22-15CHX", 9333), "3512924179": ("HUGHES 2-22-27-34CHX", 10184),
    # Cherokee, around Merlin 2/11 (13N-19W)
    "3503922640": ("OSPREY 1-10-15XH", 8018), "3503922656": ("QUETZAL 1-6-7XHR", 10086),
    "3503922666": ("TOPAZ 1-13HX(R)", 8966), "3503922628": ("MALLARD 1HX", 9849),
    "3503922663": ("COTINGA 1-14-23XH", 10041), "3503922647": ("GADWALL 27_34-13N-20W 1HX", 10061),
    "3503922648": ("GADWALL 27_34-13N-20W 3HX", 10077), "3503922649": ("GADWALL 27_34-13N-20W 2HX", 9914),
    # Cherokee, around Bogges 29/20 (13N-22W)
    "3512924197": ("RED ROOSTER 1/12 DM FEE COM 1H", 9876), "3512924167": ("GRACE KATELYN 1-36-1CHX", 9492),
    "3512924169": ("CARPENTER 12/1 OB 1H", 9275),
    # Cherokee 24+ month wells, Roger Mills / Custer (b donor set)
    "3512924146": ("KEAHEY 1-27-34CHX", 9384), "3512924145": ("KEAHEY 1-28-33CHX", 9435),
    # Tonkawa, around IO 26/23 (14N-23W)
    "3512924188": ("SHOW ME THE MONEY #1-16H", 5321), "3512924133": ("SHOWTIME 1-7H", 5116),
    "3512924140": ("SHOWDOWN 1-1H", 5959), "3512924125": ("BIG SHOW USA #1H-1213", 7917),
    "3512924180": ("BIG SHOW USA 2H-1213", 10098), "3512900023": ("STRONG CITY 1-30H19", 10729),
}

LOMAS_5 = ["3503922644", "3503922669", "3512924183", "3503922645", "3512924147", "3503922627", "3512924175",
           "3512924174", "3503922667", "3503922668", "3512924149"]
LOMAS_10 = LOMAS_5 + ["3512924184", "3503922642", "3503922670", "3503922641", "3512924189", "3503922658",
                      "3512924173", "3512924179"]
MERLIN_10 = ["3503922640", "3503922656", "3503922666", "3503922628", "3503922663", "3503922647", "3503922648",
             "3503922649", "3503922641", "3503922642", "3503922670"]
BOGGES_15 = ["3512924197", "3512924167", "3512924179", "3512924173", "3512924169", "3512924189", "3512924149",
             "3512924174", "3512924175", "3512924184", "3512924147", "3503922647", "3503922648", "3503922649",
             "3512924183"]
TONK_20 = ["3512924188", "3512924133", "3512924140", "3512924125", "3512924180"]
B_DONOR = ["3503922627", "3503922628", "3512924146", "3512924145"]

DEALS = {
    "lomas-7-18-1h": dict(
        title="Lomas 7/18 DM #1H (Mewbourne, Cherokee, Custer Co OK)", centroid=(35.7832, -99.3675), well="LOMAS 7-18 CN 1H",
        fit=LOMAS_5, fit10=LOMAS_10, bset=B_DONOR, radius="5 mi",
        note="11 Cherokee wells with 12+ months inside 5 mi (low confidence); 10-mile comparison set of 19"),
    "lomas-7-18-2h": dict(
        title="Lomas 7/18 DM #2H (Mewbourne, Cherokee, Custer Co OK)", centroid=(35.7832, -99.3675), well="LOMAS 7-18 BO 2H",
        fit=LOMAS_5, fit10=LOMAS_10, bset=B_DONOR, radius="5 mi",
        note="11 Cherokee wells with 12+ months inside 5 mi (low confidence); 10-mile comparison set of 19"),
    "merlin-1-11-2xh": dict(
        title="Merlin 1-11-2XH (Nadel & Gussman, Cherokee, Custer Co OK)", centroid=(35.6247, -99.183), well="MERLIN 1-11-2XH",
        fit=MERLIN_10, fit10=[], bset=B_DONOR, radius="10 mi",
        note="3 Cherokee wells with 12+ months inside 5 mi (Osprey, Quetzal, Topaz); widened to 10 mi (11 wells)"),
    "bogges-29-20": dict(
        title="Bogges 29/20 (Mewbourne, Cherokee, Roger Mills Co OK)", centroid=(35.5815, -99.563), well="BOGGES 29/20",
        fit=BOGGES_15, fit10=[], bset=B_DONOR, radius="15 mi",
        note="1 Cherokee well with 12+ months inside 5 mi (Red Rooster), 3 inside 10 mi; widened to 15 mi (15 wells)"),
    "io-23-26": dict(
        title="IO 23/26 (Mewbourne, Tonkawa, Roger Mills Co OK)", centroid=(35.6678, -99.6109), well="IO 26-23 PA 2HT",
        fit=TONK_20, fit10=[], bset=[], radius="20 mi",
        note="No Tonkawa well with 12+ months inside 5 or 10 mi; 5 inside 15 and 20 mi (Smith 30X31 #1HB / #2HC, 2018 "
             "Diversified, excluded: ~940 lb/ft, peak under 6,000 bbl/month). Strong City 1-30H19 (1.6 mi, 10 months) is "
             "a check only"),
}

EXCLUDED = {
    "3512924148": "JEANELDA 3/10 BO 1H: negative oil reported in months 1-2",
    "3512924142": "GLASS 1-25-24CHX: gas drops from 423 to 95 MMcf in one month (Aug 2024) plus shut-ins; reporting artifact",
    "3512924144": "TROUBADOUR 32-5 2H: same Aug 2024 gas cliff (270 to 130 MMcf)",
    "3512924115": "SMITH 30X31-16-21 #1HB: 2018 design, peak 6,252 bbl/month", "3512924116": "SMITH 30X31-16-21 #2HC: 2018 design",
}


def load():
    raw = defaultdict(dict)
    for api, gas, oil, ym in csv.reader(open(RAW)):
        raw[api][ym] = (max(0.0, float(oil)), max(0.0, float(gas)))
    return raw


def months(a, b):
    y, m = int(a[:4]), int(a[5:7]); out = []
    while f"{y:04d}-{m:02d}" <= b:
        out.append(f"{y:04d}-{m:02d}"); m += 1
        if m > 12: y, m = y + 1, 1
    return out


def gm(x, y):
    return math.sqrt(x * y) if x > 0 and y > 0 else max(x, y)


def series(api, rec, log):
    ks = sorted(rec)
    ms = months(ks[0], ks[-1])
    o = [rec.get(m, (0, 0))[0] for m in ms]; g = [rec.get(m, (0, 0))[1] for m in ms]
    while o and o[0] == 0 and g[0] == 0:
        o, g, ms = o[1:], g[1:], ms[1:]
    # 1) interior shut-in months (both streams under 35% of the month before: shut-ins, offset-frac downtime) are removed
    keep = [True] * len(o)
    for i in range(1, len(o) - 1):
        j = i - 1
        while j > 0 and not keep[j]:
            j -= 1
        if o[i] <= 0.35 * max(o[j], 1) and g[i] <= 0.35 * max(g[j], 1):
            keep[i] = False; log.append(f"{api} {ms[i]}: shut-in month removed (oil {o[i]:.0f}, gas {g[i]:.0f})")
    o = [v for v, k in zip(o, keep) if k]; g = [v for v, k in zip(g, keep) if k]; ms = [v for v, k in zip(ms, keep) if k]
    # 2) one stream missing in a month the other stream produced normally: filled with the neighbors' geometric mean
    for i in range(1, len(o) - 1):
        for s, t in ((g, o), (o, g)):
            ref_s, ref_t = gm(s[i - 1], s[i + 1]), gm(t[i - 1], t[i + 1])
            if ref_s > 0 and s[i] < 0.30 * ref_s and t[i] >= 0.60 * ref_t:
                log.append(f"{api} {ms[i]}: {'gas' if s is g else 'oil'} {s[i]:.0f} filled to {ref_s:.0f}")
                s[i] = ref_s
    return ms[0], [round(v) for v in o], [round(v) for v in g]


def main():
    raw = load()
    for slug, d in DEALS.items():
        log = []
        apis = list(dict.fromkeys(d["fit"] + d["fit10"] + d["bset"]))
        ddir = os.path.join(REPO, "deals", slug); os.makedirs(ddir, exist_ok=True)
        L = [f"# Enverus monthly production (gross 8/8), pulled 2026-11-23 via the Enverus connector; last reported month 2026-08.",
             f"# {d['title']}. DSU centroid {d['centroid']} (lateral midpoints measured to it).",
             f"# Fit set: {d['radius']} radius. {d['note']}.",
             "# api: (well, lateral_ft, first_month, oil_bbl[], gas_mcf[])",
             "WELLS = {"]
        for a in apis:
            first, o, g = series(a, raw[a], log)
            name, lat = META[a]
            L.append(f'"{a}": ("{name}", {lat}, "{first}",\n [{",".join(map(str, o))}],\n [{",".join(map(str, g))}]),')
        L.append("}")
        L.append('LAST_MONTH = "2026-08"\n')
        L.append(f"FIT_SET = {d['fit']!r}   # {d['radius']}: {d['note']}")
        L.append("ENV_EUR = {}   # Enverus EUR fields return blank on this subscription; guardrail off\n")
        L.append("# ProducingDays: Oklahoma rows in Enverus carry calendar days (30/31 every month), so they can't flag a partial")
        L.append("# first month; left empty so build_tc.py uses its month-2 test instead.")
        L.append("DAYS = {}\n")
        if d["fit10"]:
            L.append("# Fewer than 15 wells in the 5-mile set: 10-mile comparison curve")
        L.append(f"FIT_SET_10MI = {d['fit10']!r}\n")
        L.append("# b donor: fewer than 3 fit wells have 24+ months (young Cherokee development). b is fitted on the 24+ month")
        L.append("# Cherokee horizontals in Roger Mills / Custer (Sage Dragon, Mallard 1HX, Keahey 1-27-34 / 1-28-33) and applied to")
        L.append("# this fit set's mean (build_tc.py B_SET). Glass 1-25-24CHX and Troubadour 32-5 2H left out: Aug 2024 gas cliff.")
        L.append(f"B_SET = {d['bset']!r}\n")
        L.append("# Cleaning applied to the raw Enverus rows (make_deal_data.py): interior shut-in months (both streams under 35% of the month before) removed; a single")
        L.append("# stream reported near zero in a month the other stream produced normally (Oklahoma Mar 2026 gas gap) filled with")
        L.append("# the neighbouring months' geometric mean. Excluded from all sets:")
        for k, v in EXCLUDED.items():
            L.append(f"#   {k} {v}")
        L.append("CLEANING_LOG = [")
        L += [f"  {x!r}," for x in log]
        L.append("]")
        open(os.path.join(ddir, "deal_data.py"), "w").write("\n".join(L) + "\n")
        print(slug, len(apis), "wells;", len(log), "patches")
    print("\n".join(log))


if __name__ == "__main__":
    main()
