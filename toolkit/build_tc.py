"""Type curve builder (TRUNORTH comp standard), same method as the template's TypeCurve sheet.

Usage: python build_tc.py <deal_dir> [--oil-b X] [--gas-b X]   (fixed b: sensitivities only)
Reads <deal_dir>/deal_data.py; writes <deal_dir>/tc_summary.json and tc_curves.csv.

Multi-zone deals (wells landed in different formations): deal_data.py defines
    ZONES = {"Wolfcamp B": dict(fit=[apis], fit_10mi=[apis], radius="5 mi", note=""), "Jo Mill": dict(...), ...}
(insertion order; up to 6 zones; the first zone is zone 1 = the TypeCurve tab and also supplies FIT_SET / FIT_SET_10MI).
Every zone gets its own curve from its own offsets with the same method (b fitted per zone), plus a 10-mile comparison
curve when its fit set has fewer than 15 wells. Writes tc_zones.json, tc_zone_curves.csv and tc_zone_curves_10mi.csv.

Lateral length (lateral.py, Davis's rule Nov 2026):
  * each offset is normalized straight to the 10,000 ft basis. Within +/- 2,500 ft of 10,000: linear (an 8,000 ft
    offset x 1.25 every month). Further away: months 1-8 get the linear factor x f (f phased from 100% at 2,500 ft away
    to 90% at 5,000 ft away and beyond) when the offset is shorter than 10,000 ft, and the linear factor / f when it is
    longer (mirror: a 5,000 ft offset x 2 x 0.9, a 20,000 ft offset x 0.5 / 0.9); from month 9 the offset's decline is
    re-shaped (time-stretched, same b) so its EUR still scales linearly. The stretch is solved on the fitted curve
    itself, so the fit is iterated a few times.
  * the engine / workbook move the 10,000 ft curve to each well's lateral on the same rule ('Lateral Curves').

Method (reverse-engineered from the model and its notes):
  * volumes per 10,000 ft of lateral (via the lateral rule above); month 1 = first producing month
  * a partial first month (oil < 50% of month 2) is dropped for both streams; a partial last reported month
    (oil < 50% of the month before) is dropped
  * arithmetic mean by month on production, months where at least half the fit set reports
  * multi-phase: if the mean rises after month 1 (peak in months 2-6 at least 10% above month 1),
    months before the peak are a straight-line ramp ending at the peak, and the decline starts at the
    peak (t = 0 at peak). Ramp start, qi and De (and b, for the b fit) are fitted together on all months.
    No rise -> peak month 1 and no ramp (a single-phase curve, as before)
  * oil and gas are fitted the same way, each on its own production (no GOR):
    modified hyperbolic, monthly vol = 30.4 * qi*(1+b*Di*t)^(-1/b), t in years from the peak,
    Di from the effective first-year decline (secant), 7% effective terminal decline;
    b fitted on the mean of the wells with 24+ months (from its peak), then qi / De fitted on the full mean with b fixed
  * guardrail: base EUR never above the Enverus fit-set median EUR per 10,000 ft
"""
import json, math, sys, os, importlib.util
import numpy as np
from scipy.optimize import least_squares
import lateral as LAT

# deal data: <deal_dir>/deal_data.py defining WELLS {api: (name, lateral_ft, first_month, oil[], gas[])},
# FIT_SET [apis], optional ENV_EUR {api: (oil MBO/10k, gas Bcf/10k)}
WELLS, FIT_SET, ENV_EUR, DEAL_DIR, OIL_B_FIXED, GAS_B_FIXED = {}, [], {}, ".", None, None
ZONES = {}                       # optional: multi-zone fit sets (see above)
DAYS, FIT_SET_10MI = {}, []      # optional: producing days by month {api: [days]}; 10-mile fit set [apis]
SERIES, ORIG_LEN = {}, {}        # offsets normalized to the 10,000 ft basis for the set being built; cleaned months reported
MIN_FIT_WELLS = 15               # below this the 5-mile curve is low confidence and the 10-mile curve is built too


def load_deal(deal_dir):
    global WELLS, FIT_SET, ENV_EUR, DEAL_DIR, DAYS, FIT_SET_10MI, ZONES
    spec = importlib.util.spec_from_file_location("deal_data", os.path.join(deal_dir, "deal_data.py"))
    dd = importlib.util.module_from_spec(spec); spec.loader.exec_module(dd)
    ZONES = dict(getattr(dd, "ZONES", {}) or {})
    z1 = next(iter(ZONES.values()), {})
    WELLS, ENV_EUR, DEAL_DIR = dd.WELLS, getattr(dd, "ENV_EUR", {}), deal_dir
    FIT_SET = list(getattr(dd, "FIT_SET", None) or z1.get("fit", []))
    DAYS = getattr(dd, "DAYS", {})
    FIT_SET_10MI = list(getattr(dd, "FIT_SET_10MI", None) or z1.get("fit_10mi", []))
    if len(ZONES) > 6:
        raise SystemExit("ZONES: the template holds 6 zone curves at most")


DMIN = 0.07
N = 360


def curve(qi, De, b, n=N, dmin=DMIN):
    b = max(b, 1e-6)
    Di = ((1 - De) ** (-b) - 1) / b
    dm = -math.log(1 - dmin)
    out, q_sw, t_sw = [], None, None
    for m in range(n):
        t = m / 12
        D = Di / (1 + b * Di * t)
        if q_sw is None and D <= dm:
            t_sw, q_sw = t, qi * (1 + b * Di * t) ** (-1 / b)
        q = qi * (1 + b * Di * t) ** (-1 / b) if q_sw is None else q_sw * math.exp(-dm * (t - t_sw))
        out.append(q * 30.4)
    return np.array(out)


MIN_DAYS = 7                     # a first month with fewer producing days than this is dropped


def _cal_days(first, m):
    y, mo = int(first[:4]), int(first[5:7]) + m
    y, mo = y + (mo - 1) // 12, (mo - 1) % 12 + 1
    import calendar
    return calendar.monthrange(y, mo)[1]


def clean(o, g, days=None, first=None):
    """Month 1 = first producing month. With producing days (Enverus ProducingDays): a partial first
    month is scaled up to a full calendar month (dropped if under MIN_DAYS days); later months keep
    their reported volumes, so normal downtime stays in the curve. Without days: a first month below
    50% of month 2 is treated as partial and dropped. A partial last reported month (below 50% of
    the month before) is dropped either way."""
    o, g = list(o), list(g)
    if days and first and len(days) >= 1 and days[0]:
        cal = _cal_days(first, 0)
        if days[0] < MIN_DAYS:
            o, g = o[1:], g[1:]
        elif days[0] < cal:
            f = cal / days[0]; o[0], g[0] = o[0] * f, g[0] * f
    elif len(o) > 1 and o[0] < 0.5 * o[1]:
        o, g = o[1:], g[1:]
    if len(o) > 1 and o[-1] < 0.5 * o[-2]:
        o, g = o[:-1], g[:-1]
    return np.array(o, float), np.array(g, float)


def prepare_series(apis, lref, ref=None):
    """Clean each offset and move it to the modeled lateral lref: linear within the band, otherwise the shaped rule
    with the moves taken from ref = (oil, gas) curves at lref (None = linear for every well, the first pass).
    Fills SERIES / ORIG_LEN; returns the APIs that were shaped."""
    global SERIES, ORIG_LEN
    SERIES, ORIG_LEN, shaped = {}, {}, []
    for a in apis:
        name, lat, first, o, g = WELLS[a][:5]
        o, g = clean(o, g, DAYS.get(a), first)
        ORIG_LEN[a] = len(o)
        if ref is None or abs(lref - lat) <= LAT.BAND:
            SERIES[a] = (o * lref / lat, g * lref / lat)
            continue
        out = []
        for d, c in ((o, ref[0]), (g, ref[1])):
            c_here = LAT.scale_curve(c, lref, lat)[0]             # the curve at this offset's lateral
            out.append(LAT.scale_data(d, LAT.scale_curve(c_here, lat, lref)[1]))
        n = min(len(out[0]), len(out[1]))
        SERIES[a] = (out[0][:n], out[1][:n]); shaped.append(a)
    return shaped


def to_basis(c, lref=LAT.BASIS):
    """A curve fitted at lref moved to the 10,000 ft basis (curves are now fitted on the basis, so lref = 10,000 and
    this returns the curve unchanged; kept for tc_summary files from the short-lived fit-at-deal-lateral version)."""
    return LAT.scale_curve(c, lref, LAT.BASIS)[0]


def mean_curve(apis):
    series = {a: SERIES[a] for a in apis}
    L = max(len(s[0]) for s in series.values())
    oil, gas, cnt = [], [], []
    for m in range(L):
        vo = [s[0][m] for s in series.values() if len(s[0]) > m]
        vg = [s[1][m] for s in series.values() if len(s[1]) > m]
        oil.append(np.mean(vo)); gas.append(np.mean(vg)); cnt.append(len(vo))
    need = math.ceil(len(apis) / 2)
    keep = [m for m in range(L) if cnt[m] >= need]
    M = max(keep) + 1
    return np.array(oil[:M]), np.array(gas[:M]), cnt[:M], series


def fit(vol, b=None, b_bounds=(0.05, 2.0)):
    """Least squares on log monthly volumes. Returns qi (per day), De, b."""
    y = np.log(vol)
    q0 = vol[0] / 30.4
    if b is None:
        f = lambda p: np.log(curve(p[0], p[1], p[2], n=len(vol))) - y
        r = least_squares(f, [q0, 0.65, 0.8], bounds=([q0 * 0.3, 0.2, b_bounds[0]], [q0 * 3, 0.95, b_bounds[1]]))
        return tuple(r.x)
    f = lambda p: np.log(curve(p[0], p[1], b, n=len(vol))) - y
    r = least_squares(f, [q0, 0.65], bounds=([q0 * 0.3, 0.2], [q0 * 3, 0.95]))
    return r.x[0], r.x[1], b


RAMP_SEARCH = 6        # the peak is looked for in months 1-6
RAMP_MIN_RISE = 0.10   # a ramp needs the peak at least 10% above month 1


def find_peak(vol):
    """0-based peak month; 0 = no ramp."""
    p = int(np.argmax(vol[:min(RAMP_SEARCH, len(vol) - 12)]))
    return p if p > 0 and vol[p] >= vol[0] * (1 + RAMP_MIN_RISE) else 0


def multiphase(qi, De, b, p, r1, n=N):
    """Linear ramp over months 1..p (0-based months 0..p-1), hyperbolic from month p+1 (t = 0 at peak)."""
    dec = curve(qi, De, b, n=n - p)
    ramp = [dec[0] * (r1 + (1 - r1) * m / p) for m in range(p)] if p else []
    return np.concatenate([ramp, dec])[:n]


def fit_tc(vol, b=None, p=None, b_bounds=(0.05, 2.0)):
    """Multi-phase fit on a mean monthly series: ramp start level, qi (per day at the peak), De and,
    when b is None, b are fitted together on log monthly volumes over all months (ramp and decline),
    with the peak month p taken from the data. Returns (360-month curve, dict(qi, De, b, peak, r1))."""
    p = find_peak(vol) if p is None else p
    vol = np.asarray(vol, float); n = len(vol); q0 = vol[p] / 30.4
    ok = vol > 0                     # months averaging to zero (shut-ins) are left out of the log fit
    y = np.log(np.where(ok, vol, 1.0))
    if b is None:
        f = lambda q: (np.log(multiphase(q[0], q[1], q[3], p, q[2], n=n)) - y)[ok]
        r = least_squares(f, [q0, 0.65, 0.7, 0.8], bounds=([q0 * 0.3, 0.2, 0.05, b_bounds[0]], [q0 * 3, 0.95, 1.0, b_bounds[1]]))
        qi, De, r1, bb = r.x
    else:
        f = lambda q: (np.log(multiphase(q[0], q[1], b, p, q[2], n=n)) - y)[ok]
        r = least_squares(f, [q0, 0.65, 0.7], bounds=([q0 * 0.3, 0.2, 0.05], [q0 * 3, 0.95, 1.0]))
        (qi, De, r1), bb = r.x, b
    r1 = r1 if p else 1.0
    return multiphase(qi, De, bb, p, r1), dict(qi=float(qi), De=float(De), b=float(bb), peak=int(p), r1=float(r1))


def stream_fit(k, apis, series, long_wells, mean_all, b_fixed=None):
    """Multi-phase fit for one stream (k = 0 oil, 1 gas). b from the 24+ month wells' mean (from its
    peak), checked across windows and by bootstrap; then ramp / qi / De on the full mean with b fixed."""
    long_mean = mean_curve(long_wells)[k]
    p_l = find_peak(long_mean)
    b_fitted = fit_tc(long_mean, p=p_l)[1]["b"]
    b_use = b_fitted if b_fixed is None else b_fixed
    windows = {w: fit_tc(long_mean[:w], p=p_l)[1]["b"] for w in (18, 24, 30, len(long_mean)) if w <= len(long_mean)}
    rng = np.random.default_rng(7)
    boots = []
    for _ in range(300):
        pick = list(rng.choice(long_wells, size=len(long_wells), replace=True))
        if len(set(pick)) < 3:
            continue
        v_l, L = [], max(len(series[a][k]) for a in pick)
        for m in range(L):
            v = [series[a][k][m] for a in pick if len(series[a][k]) > m]
            if len(v) < math.ceil(len(pick) / 2):
                break
            v_l.append(np.mean(v))
        boots.append(fit_tc(np.array(v_l))[1]["b"])
    c, P = fit_tc(mean_all, b=b_use)
    return c, dict(qi=P["qi"], De=P["De"], b=float(b_use), b_fitted=float(b_fitted),
                   b_pinned=bool(b_fitted <= 0.051 or b_fitted >= 1.999), b_windows=windows,
                   b_boot_p10_p90=[float(x) for x in np.percentile(boots, [10, 90])],
                   peak_month=P["peak"] + 1, ramp_start=P["r1"], b_fit_peak_month_long=p_l + 1)


OFFSETS_BASIS = {}                 # api -> (oil, gas) per 10,000 ft on the lateral rule, for plot_tc (saved to tc_offsets.json)


def build_set(apis):
    """Full type-curve build for one fit set. Returns (base_o, base_g, env_o_curve, env_g_curve, summary)."""
    lref = LAT.BASIS
    shaped = prepare_series(apis, lref)
    far = [a for a in apis if abs(WELLS[a][1] - lref) > LAT.BAND]
    if far:                            # shaped moves depend on the curve: fit, re-move the far offsets, refit
        for _ in range(3):
            m_o, m_g = mean_curve(apis)[:2]
            shaped = prepare_series(apis, lref, (fit_tc(m_o)[0], fit_tc(m_g)[0]))
    oil, gas, cnt, series = mean_curve(apis)
    long_wells = [a for a in apis if ORIG_LEN[a] >= 24]
    oil_r, PO = stream_fit(0, apis, series, long_wells, oil, OIL_B_FIXED)
    gas_r, PG = stream_fit(1, apis, series, long_wells, gas, GAS_B_FIXED)
    OFFSETS_BASIS.update({a: SERIES[a] for a in apis})
    oil_c, gas_c = oil_r, gas_r
    eur_o = oil_c.sum()

    have_env = all(a in ENV_EUR for a in apis)
    if have_env:
        env_o = float(np.median([ENV_EUR[a][0] for a in apis])) * 1000
        env_g = float(np.median([ENV_EUR[a][1] for a in apis])) * 1e6
        env_o_p = np.percentile([ENV_EUR[a][0] for a in apis], [90, 50, 10])
        env_g_p = np.percentile([ENV_EUR[a][1] for a in apis], [90, 50, 10])
    else:   # no Enverus EURs supplied: guardrail off, Enverus columns = base
        env_o, env_g = float(eur_o), float(gas_c.sum())
        env_o_p = np.array([env_o, env_o, env_o]) / 1000; env_g_p = np.array([env_g, env_g, env_g]) / 1e6
    guard_o = min(1.0, env_o / eur_o)
    guard_g = min(1.0, env_g / gas_c.sum())
    base_o = oil_c * guard_o
    base_g = gas_c * guard_g
    env_oil_c = base_o * (env_o / base_o.sum())
    env_gas_c = base_g * (env_g / base_g.sum())

    out = {
        "wells": len(apis), "long_wells": len(long_wells), "months_in_mean": len(oil),
        "lateral": {"model_lateral": lref, "offsets_linear": len(apis) - len(shaped), "offsets_shaped": shaped,
                    "offset_laterals": [float(WELLS[a][1]) for a in apis],
                    "rule": {"band_ft": LAT.BAND, "early_months": LAT.EARLY_MONTHS, "early_pct": LAT.EARLY_PCT, "full_ft": LAT.FULL_AT}},
        "wells_by_month": cnt,
        "fit_set": apis, "enverus_eurs_supplied": have_env,
        "oil": {**PO, "qi": PO["qi"] * guard_o, "b_fixed_sensitivity": OIL_B_FIXED,
                "eur": float(base_o.sum()), "eur_unguarded": float(eur_o)},
        "gas": {**PG, "qi": PG["qi"] * guard_g, "b_fixed_sensitivity": GAS_B_FIXED, "basis": "multi-phase fit",
                "eur": float(base_g.sum()), "eur_unguarded": float(gas_c.sum())},
        "enverus": {"oil_median": env_o, "gas_median": env_g,
                    "oil_p90_p50_p10": [float(x) for x in env_o_p],
                    "gas_p90_p50_p10": [float(x) for x in env_g_p]},
        "guardrail": {"oil": guard_o, "gas": guard_g},
        "fit_vs_median": {"oil": float(base_o.sum() / env_o), "gas": float(base_g.sum() / env_g)},
        "mean_oil": [float(x) for x in oil], "mean_gas": [float(x) for x in gas],
    }
    return base_o, base_g, env_oil_c, env_gas_c, out


ZONE_KEYS = ("wells", "long_wells", "months_in_mean", "lateral", "oil", "gas", "enverus", "guardrail", "fit_vs_median", "mean_oil", "mean_gas", "fit_set")


def build_zones():
    """One curve per zone (same method as the base), 10-mile comparison where the fit set is under 15 wells."""
    zs, cols, cols10 = [], [np.arange(1, N + 1)], [np.arange(1, N + 1)]
    hdr, hdr10 = ["month"], ["month"]
    for k, (name, z) in enumerate(ZONES.items()):
        fit = list(z["fit"])
        o, g, _, _, s = build_set(fit)
        rec = dict(name=name, zone=k + 1, radius=z.get("radius", "5 mi"), note=z.get("note", ""),
                   low_confidence=len(fit) < MIN_FIT_WELLS, **{kk: s[kk] for kk in ZONE_KEYS})
        t_o = t_g = np.zeros(N)
        if len(fit) < MIN_FIT_WELLS and z.get("fit_10mi"):
            t_o, t_g, _, _, t = build_set(list(z["fit_10mi"]))
            rec["ten_mile"] = {kk: t[kk] for kk in ZONE_KEYS}
        zs.append(rec)
        cols += [o, g]; cols10 += [t_o, t_g]; hdr += [f"z{k + 1}_oil", f"z{k + 1}_gas"]; hdr10 += [f"z{k + 1}_oil", f"z{k + 1}_gas"]
        print(f"zone {k + 1} {name}: {len(fit)} wells ({rec['radius']}){', LOW CONFIDENCE' if rec['low_confidence'] else ''}; "
              f"oil {o.sum()/1000:.0f} MBO (b {s['oil']['b']:.2f}{', at cap' if s['oil']['b_pinned'] else ''}), "
              f"gas {g.sum()/1e6:.2f} Bcf (b {s['gas']['b']:.2f}{', at cap' if s['gas']['b_pinned'] else ''}) per 10,000 ft"
              + (f"; 10-mile {rec['ten_mile']['wells']} wells: {t_o.sum()/1000:.0f} MBO, {t_g.sum()/1e6:.2f} Bcf" if "ten_mile" in rec else ""))
    json.dump(dict(zones=zs), open(os.path.join(DEAL_DIR, "tc_zones.json"), "w"), indent=1)
    np.savetxt(os.path.join(DEAL_DIR, "tc_zone_curves.csv"), np.column_stack(cols), delimiter=",", header=",".join(hdr), comments="", fmt="%.6f")
    np.savetxt(os.path.join(DEAL_DIR, "tc_zone_curves_10mi.csv"), np.column_stack(cols10), delimiter=",", header=",".join(hdr10), comments="", fmt="%.6f")


def main():
    for f in ("tc_zones.json", "tc_zone_curves.csv", "tc_zone_curves_10mi.csv"):
        if not ZONES and os.path.exists(os.path.join(DEAL_DIR, f)):
            os.remove(os.path.join(DEAL_DIR, f))
    if ZONES:
        build_zones()
    base_o, base_g, env_oil_c, env_gas_c, out = build_set(FIT_SET)
    _b5 = {a: OFFSETS_BASIS[a] for a in FIT_SET}
    out["days_used"] = sorted(a for a in FIT_SET if DAYS.get(a))
    if len(FIT_SET) < MIN_FIT_WELLS and FIT_SET_10MI:
        t_o, t_g, _, _, t = build_set(FIT_SET_10MI)
        OFFSETS_BASIS.update(_b5)                  # the 5-mile wells keep their 5-mile-fit view on the charts
        out["ten_mile"] = {k: t[k] for k in ("wells", "lateral", "long_wells", "months_in_mean", "oil", "gas", "enverus", "guardrail",
                                             "fit_vs_median", "mean_oil", "mean_gas", "fit_set")}
        np.savetxt(os.path.join(DEAL_DIR, "tc_curves_10mi.csv"), np.column_stack([np.arange(1, N + 1), t_o, t_g]),
                   delimiter=",", header="month,oil,gas", comments="", fmt="%.6f")
    elif os.path.exists(os.path.join(DEAL_DIR, "tc_curves_10mi.csv")):
        os.remove(os.path.join(DEAL_DIR, "tc_curves_10mi.csv"))
    json.dump(out, open(os.path.join(DEAL_DIR, "tc_summary.json"), "w"), indent=1)
    json.dump({a: [list(map(float, v[0])), list(map(float, v[1]))] for a, v in OFFSETS_BASIS.items()},
              open(os.path.join(DEAL_DIR, "tc_offsets.json"), "w"))
    np.savetxt(os.path.join(DEAL_DIR, "tc_curves.csv"),
               np.column_stack([np.arange(1, N + 1), base_o, base_g, env_oil_c, env_gas_c]),
               delimiter=",", header="month,base_oil,base_gas,env_p50_oil,env_p50_gas", comments="", fmt="%.6f")
    s = {k: v for k, v in out.items() if k not in ("mean_oil", "mean_gas", "wells_by_month", "ten_mile")}
    print(json.dumps(s, indent=1))
    print("wells by month:", out["wells_by_month"])
    lt = out["lateral"]
    print(f"lateral: offsets normalized to 10,000 ft; {lt['offsets_linear']} linear (within +/- {LAT.BAND:,.0f} ft), "
          f"{len(lt['offsets_shaped'])} shaped")
    if "ten_mile" in out:
        t = out["ten_mile"]
        print(f"10-mile curve: {t['wells']} wells, oil {t['oil']['eur']/1000:.0f} MBO, gas {t['gas']['eur']/1e6:.2f} Bcf per 10,000 ft")


if __name__ == "__main__":
    load_deal(sys.argv[1] if len(sys.argv) > 1 else ".")
    if "--oil-b" in sys.argv:            # sensitivity only; the base case always uses the fitted b
        OIL_B_FIXED = float(sys.argv[sys.argv.index("--oil-b") + 1])
    if "--gas-b" in sys.argv:
        GAS_B_FIXED = float(sys.argv[sys.argv.index("--gas-b") + 1])
    main()
