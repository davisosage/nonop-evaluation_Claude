"""Build deal_data.py for PK 29-32 4H-LL (Phoenix, Middle Bakken, Divide Co ND).
Raw Enverus rows: raw/prod_batch*.txt (api|first|oil|gas|days, chronological), pulled 2026-11-23, last month 2026-09."""
import glob
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# api10: (name, operator, lateral ft, distance mi lateral midpoint -> DSU centroid, first prod)
META = {
    "3310506523": ("STOLTENBERG 2-4HSL", "Continental", 15448, 2.51),
    "3310506524": ("STOLTENBERG 3-4H", "Continental", 15550, 2.51),
    "3310506596": ("GOODNIGHT 22-15 5H", "Phoenix", 10303, 5.99),
    "3310506595": ("GOODNIGHT 22-15 4H", "Phoenix", 10278, 6.06),
    "3310506542": ("GOODNIGHT 22-15 3H", "Phoenix", 10221, 6.14),
    "3310506594": ("GOODNIGHT 22-15 2H", "Phoenix", 10249, 6.22),
    "3310506593": ("GOODNIGHT 22-15 1H-LL", "Phoenix", 10277, 6.30),
    "3302301475": ("GOODSON 2-21H", "Continental", 14923, 6.34),
    "3302301617": ("SCHILKE 3-19H", "Continental", 15553, 6.72),
    "3302301474": ("DALIN 2-16H", "Continental", 15632, 7.11),
    "3302301558": ("JACOBSON 19-30-31 6H-LL", "Phoenix", 15548, 7.25),
    "3302301551": ("JACOBSON 19-30-31 5H", "Phoenix", 15539, 7.35),
    "3302301552": ("JACOBSON 19-30-31 4H", "Phoenix", 15476, 7.45),
    "3302301545": ("JACOBSON 19-30-31 3H", "Phoenix", 15384, 7.55),
    "3302301550": ("JACOBSON 19-30-31 2H", "Phoenix", 15510, 7.65),
    "3302301563": ("JACOBSON 19-30-31 1H-LLR", "Phoenix", 15278, 7.73),
    "3310506040": ("APOLLO 18-7-6 #3H", "Kraken", 15492, 8.00),
    "3310506039": ("APOLLO 18-7-6 #2H", "Kraken", 15492, 8.09),
    "3310506038": ("APOLLO 18-7-6 #1H", "Kraken", 15520, 8.18),
    "3310506585": ("WILLER 28-33 6H-LL", "Phoenix", 10349, 8.09),
    "3310506584": ("WILLER 28-33 5H", "Phoenix", 10155, 8.16),
    "3310506583": ("WILLER 28-33 4H", "Phoenix", 10264, 8.23),
    "3310506536": ("WILLER 28-33 3H", "Phoenix", 10399, 8.30),
    "3310506582": ("WILLER 28-33 2H", "Phoenix", 10250, 8.37),
    "3310506581": ("WILLER 28-33 1H-LL", "Phoenix", 10231, 8.44),
    "3302301544": ("AXEL FERRARI 25-36-1 5H", "Phoenix", 15413, 8.42),
    "3302301543": ("AXEL FERRARI 25-36-1 4H", "Phoenix", 15456, 8.51),
    "3302301540": ("AXEL FERRARI 25-36-1 3H", "Phoenix", 15629, 8.62),
    "3302301542": ("AXEL FERRARI 25-36-1 2H", "Phoenix", 15576, 8.73),
    "3302301541": ("AXEL FERRARI 25-36-1 1H-LL", "Phoenix", 15507, 8.83),
    "3310505925": ("ORBIT 24-13-12 #5H", "Kraken", 15415, 8.96),
    "3310505924": ("ORBIT 24-13-12 #4H", "Kraken", 15533, 9.05),
    "3310505922": ("ORBIT 24-13-12 #2H", "Kraken", 15501, 9.24),
    "3310505921": ("ORBIT 24-13-12 #1H", "Kraken", 15676, 9.33),
    "3302301596": ("KLEIST 3-35H", "Continental", 15499, 9.01),
    "3302301616": ("KLEIST 4-35HSL", "Continental", 15514, 9.12),
    "3302301529": ("DANIELE 26-35-2 5H", "Phoenix", 15542, 9.31),
    "3302301528": ("DANIELE 26-35-2 4H", "Phoenix", 15513, 9.41),
    "3302301522": ("DANIELE 26-35-2 3H", "Phoenix", 15498, 9.51),
    "3302301527": ("DANIELE 26-35-2 2H", "Phoenix", 15480, 9.61),
    "3302301526": ("DANIELE 26-35-2 1H-LL", "Phoenix", 15530, 9.72),
    "3310506037": ("TURBODIESEL 19-30-31 #3H", "Kraken", 15463, 9.38),
    "3310506036": ("TURBODIESEL 19-30-31 #2H", "Kraken", 15399, 9.45),
    "3310506035": ("TURBODIESEL 19-30-31 #1H", "Kraken", 15377, 9.52),
    "3310506392": ("GN-STUNDAL-LE- 158-97-1819H-1", "Chevron", 10016, 10.00),
    # checks only (under 12 months by Sep-26): Phoenix's newest wells nearest the PK
    "3302301644": ("WOODS 35-26-23 4H", "Phoenix", 15489, 3.45),
    "3302301643": ("WOODS 35-26-23 3H", "Phoenix", 15507, 3.55),
    "3302301624": ("WOODS 35-26-23 2H", "Phoenix", 15517, 3.65),
    "3302301642": ("WOODS 35-26-23 1H", "Phoenix", 15304, 3.74),
    "3310506617": ("TERRY NELSON 34-27-22 5H-LL", "Phoenix", 15491, 4.22),
    "3310506695": ("TERRY NELSON 34-27-22 4HR", "Phoenix", 15457, 4.32),
}
CHECK = ["3302301644", "3302301643", "3302301624", "3302301642", "3310506617", "3310506695"]
FIT = [a for a in META if a not in CHECK]
FIVE_MI = ["3310506523", "3310506524"]


def load():
    rows = {}
    for f in sorted(glob.glob(os.path.join(HERE, "raw", "prod_batch*.txt"))):
        for line in open(f):
            if line.startswith("#") or not line.strip():
                continue
            a, first, o, g, d = line.strip().split("|")
            rows[a] = (first, [float(x) for x in o.split(",")], [float(x) for x in g.split(",")],
                       [int(x) for x in d.split(",")])
    return rows


def series(api, first, o, g, d, log):
    assert len(o) == len(g) == len(d), api
    # interior shut-in months (both streams under 35% of the last kept month: offset-frac downtime) are removed,
    # with their producing days, so the curve isn't carved by neighbours' fracs
    keep = [True] * len(o)
    for i in range(1, len(o) - 1):
        j = i - 1
        while j > 0 and not keep[j]:
            j -= 1
        if o[i] <= 0.35 * max(o[j], 1) and g[i] <= 0.35 * max(g[j], 1):
            keep[i] = False
            log.append(f"{api} month {i + 1}: shut-in month removed (oil {o[i]:.0f}, gas {g[i]:.0f}, {d[i]} days)")
    o = [v for v, k in zip(o, keep) if k]; g = [v for v, k in zip(g, keep) if k]; d = [v for v, k in zip(d, keep) if k]
    return first, [round(v) for v in o], [round(v) for v in g], d


def main():
    raw = load()
    missing = [a for a in META if a not in raw]
    assert not missing, missing
    log = []
    L = ["# Enverus monthly production (gross 8/8) and ProducingDays, pulled 2026-11-23 via the Enverus connector; last month 2026-09.",
         "# PK 29-32 4H-LL (Phoenix Energy One, Middle Bakken, Divide Co ND; API 33-023-01894, lateral 9,989 ft).",
         "# DSU 28/29/32/33-160N-97W, centroid (48.6475, -103.2325); lateral midpoints measured to it.",
         "# Screen: Middle Bakken (ENVInterval), first production 2023+, 12+ months by Sep-26, lateral 9,000-16,000 ft (2- and",
         "# 3-mile), modern completions (900+ lb/ft), TF-named wells out. 5 mi: only Stoltenberg 2-4HSL / 3-4H (Continental)",
         "# pass (Woods and Terry Nelson, Phoenix, 3.5-4.3 mi, have 10-11 months; Larson 44-12-4H is a 2018 650 lb/ft well),",
         "# so the fit set is widened to 10 mi: 45 wells (Phoenix 27, Kraken 10, Continental 7, Chevron 1).",
         "# api: (well, lateral_ft, first_month, oil_bbl[], gas_mcf[])",
         "WELLS = {"]
    days = {}
    for a, (name, op, lat, dist) in META.items():
        first, o, g, d = series(a, *raw[a], log)
        days[a] = d
        L.append(f'"{a}": ("{name}", {lat}, "{first}",\n [{",".join(map(str, o))}],\n [{",".join(map(str, g))}]),')
    L.append("}")
    L.append('LAST_MONTH = "2026-09"\n')
    L.append(f"FIT_SET = {FIT!r}   # 10 mi (2 wells inside 5 mi)")
    L.append(f"FIVE_MI = {FIVE_MI!r}")
    L.append(f"CHECK = {CHECK!r}   # Woods / Terry Nelson: Phoenix, 3.5-4.3 mi, 10-11 months; shown against the curve only")
    L.append("DIST_MI = {" + ", ".join(f'"{a}": {m[3]}' for a, m in META.items()) + "}")
    L.append("OPERATOR = {" + ", ".join(f'"{a}": "{m[1]}"' for a, m in META.items()) + "}")
    L.append("ENV_EUR = {}   # Enverus EUR fields return blank on this subscription; guardrail off\n")
    L.append("# ProducingDays (Enverus production dataset): partial first months are scaled to a full month by build_tc.py")
    L.append("DAYS = {")
    L += [f'"{a}": [{",".join(map(str, d))}],' for a, d in days.items()]
    L.append("}\n")
    L.append("FIT_SET_10MI = []   # the fit set is already the 10-mile set\n")
    L.append("# Cleaning (make_deal_data.py): interior shut-in months (both streams under 35% of the month before; offset-frac")
    L.append("# downtime on Apollo, Turbodiesel, Jacobson, Orbit, Daniele) removed. Excluded: Bronson 31X-14A (Chord, 2019 DUC,")
    L.append("# gas-dominant from month 1), Orbit 24-13-12 #3TFH (TF-named), 4-mile laterals (Nystuen, Willow Gray), pre-2023 wells.")
    L.append("CLEANING_LOG = [")
    L += [f"  {x!r}," for x in log]
    L.append("]")
    open(os.path.join(HERE, "deal_data.py"), "w").write("\n".join(L) + "\n")
    print(len(META), "wells,", len(FIT), "fit;", len(log), "shut-in months removed")
    print("\n".join(log))


if __name__ == "__main__":
    main()
