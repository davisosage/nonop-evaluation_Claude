"""PDP decline fits and the offset check (PDP Evaluation Template).

Usage: python pdp_fit.py <deal_dir>
Reads <deal_dir>/pdp_data.py; writes <deal_dir>/pdp_fit.json and one chart per well (pdp_<api>.png).

pdp_data.py (gross 8/8 Enverus monthly production; see deals/pdp-example):
    DEAL = dict(name=..., basin="DJ (CO)", deal_type="Working interest" | "Minerals", effective="2026-12",
                asking_price=0, operator=...)
    LAST_MONTH = "2026-08"                                   # last reported month in the Enverus pull
    SUBJECT = {api: (name, lateral_ft, first_month, oil[], gas[], zone, wi, nri, lease_nri)}
    OFFSETS = {api: (name, lateral_ft, first_month, oil[], gas[], zone, distance_mi)}
    DAYS = {api: [producing days by month]}                  # optional (Enverus ProducingDays), subject and offsets
    ZONE_GROUPS = {"NIOBRARA A": "Niobrara", ...}            # optional: Enverus intervals that compare with each other
    FIT_FROM = {api: "YYYY-MM"}                              # optional: start a well's fit later (after a regime change)
    LOE = {api: fixed LOE $/month 8/8}                       # optional per-well fixed LOE override

Decline fit (each well, oil and gas separately on their own production, no GOR):
  * month 1 = first producing month (age). Months left out of the fit: zero months, months with producing days under
    70% of the calendar month, a partial last month (under 50% of the month before, when there are no days), and
    outliers more than ~45% below or ~100% above the first-pass fit (offset-frac shut-ins, workovers).
  * fit window: from the peak (highest good month in months 1-6) to the last month, or from FIT_FROM. The last 12
    months carry triple weight, since they set the forecast.
  * level shift: when the last 3-12 months all sit 22%+ off the earlier trend, consistently (an allocation or metering
    change, curtailment, a choke change), the fit carries a level factor from that month on: the shape comes from all
    months, the level from the recent ones. Flagged on the Wells sheet; check it with the operator.
  * modified hyperbolic: q(t) = q0 (1 + b D t)^(-1/b), t in years from the start of the window, D nominal annual,
    switching to exponential at the 7% effective terminal decline. q0, D and b are all fitted (b in 0.01-2.0, never
    hard-coded). A well with fewer than 12 good months in its window takes b from its offsets' median shape instead.
  * the forecast restarts at the last reported month (the anchor, age A): qa = fitted rate at A ("Fit" anchor) or the
    last three good months' average carried to A on the fitted shape ("Actual" anchor); Da = the fitted decline at A.
    The workbook rebuilds the forecast from qa, Da, b, so overrides typed on the Wells sheet flow straight through.

Offset check (what the offsets did over the months we are forecasting):
  * offsets = OFFSETS in the same zone group, excluding the subject wells, with good production at the subject's age.
  * anchored: each offset is scaled so its average over ages A-2..A equals the subject's own average over the same ages,
    then its months A+1..A+120 are taken as that offset's view of the subject's future. P25 / P50 / P75 by month where
    3+ offsets report (MIN_OFFSETS). The workbook compares the forecast's cumulative to the P50's over 12 / 24 / 36 /
    60 months and over the full overlap.
  * level: the subject's ages A-2..A average per 10,000 ft vs the offsets' at the same age (how good the well is).
"""
import calendar, json, math, os, sys, importlib.util
import numpy as np
from scipy.optimize import least_squares

DMIN = 0.07                 # effective terminal decline
B_LO, B_HI = 0.01, 2.0
MIN_FIT_MONTHS = 12         # fewer good months in the window: b comes from the offsets' shape
DAYS_OK = 0.70              # a month with producing days under 70% of the calendar month is left out
OUT_LO, OUT_HI = -0.6, 0.7  # log residual limits for outliers (about -45% / +100% vs the fit)
RECENT_W = 3.0              # weight of the last 12 months
MIN_OFFSETS = 3
HORIZON = 120               # forecast months compared with offsets
PEAK_SEARCH = 6


def ym_add(ym, n):
    y, m = int(ym[:4]), int(ym[5:7]) - 1 + n
    return f"{y + m // 12:04d}-{m % 12 + 1:02d}"


def ym_diff(a, b):
    """months from a to b"""
    return (int(b[:4]) - int(a[:4])) * 12 + int(b[5:7]) - int(a[5:7])


def cal_days(ym):
    return calendar.monthrange(int(ym[:4]), int(ym[5:7]))[1]


def mh(t, q0, D, b, dmin=DMIN):
    """modified hyperbolic monthly volume at t years from the window start"""
    t = np.asarray(t, float)
    b = max(b, 1e-4); dm = -math.log(1 - dmin)
    tsw = max(0.0, (D / dm - 1) / (b * D)) if D > 0 else 0.0
    qsw = q0 * (1 + b * D * tsw) ** (-1 / b)
    return np.where(t <= tsw, q0 * (1 + b * D * np.minimum(t, tsw)) ** (-1 / b), qsw * np.exp(-dm * (t - tsw)))


def decline_at(t, D, b, dmin=DMIN):
    """nominal annual decline at t (the terminal decline once switched)"""
    dm = -math.log(1 - dmin); b = max(b, 1e-4)
    tsw = max(0.0, (D / dm - 1) / (b * D)) if D > 0 else 0.0
    return dm if t >= tsw else D / (1 + b * D * t)


def good_months(vol, days, first):
    """True where a month can be used (fit, anchor, offsets)"""
    vol = np.asarray(vol, float); n = len(vol)
    ok = vol > 0
    if days:
        for i in range(n):
            d = days[i] if i < len(days) else None
            if d is not None and d < DAYS_OK * cal_days(ym_add(first, i)):
                ok[i] = False
    elif n > 1 and vol[-1] < 0.5 * vol[-2]:
        ok[-1] = False
    return ok


def fit_window(vol, ok, start=None, b_fixed=None, step=None):
    """Weighted log fit of the modified hyperbolic from the window start (0-based) to the end. step (0-based month) adds a
    level factor exp(c) from that month on (a level shift: allocation or metering change, curtailment, choke change)
    so the shape is fitted on all months and the level on the recent ones. Returns dict(start, q0, D, b, c, step, used, r2)."""
    vol = np.asarray(vol, float); n = len(vol)
    if start is None:
        cand = [i for i in range(min(PEAK_SEARCH, n)) if ok[i]]
        start = max(cand, key=lambda i: vol[i]) if cand else int(np.argmax(ok))
    use = ok.copy(); use[:start] = False
    ages = np.arange(n)
    for it in range(3):
        idx = np.where(use)[0]
        if len(idx) < 3:
            return None
        t = (idx - start) / 12; y = np.log(vol[idx])
        w = np.sqrt(np.where(idx >= n - 12, RECENT_W, 1.0))
        q00 = max(vol[idx[0]], 1.0)
        best = None
        S = (idx >= step).astype(float) if step is not None else np.zeros(len(idx))
        for b0 in ([b_fixed] if b_fixed is not None else [0.3, 0.9, 1.5]):
            def f(p, S=S):
                b = p[2] if b_fixed is None else b_fixed
                c = p[-1] if step is not None else 0.0
                return w * (np.log(mh(t, math.exp(p[0]), math.exp(p[1]), b)) + c * S - y)
            x0, lo, hi = [math.log(q00), math.log(1.0)], [math.log(q00) - 3, math.log(0.01)], [math.log(q00) + 3, math.log(60)]
            if b_fixed is None:
                x0, lo, hi = x0 + [b0], lo + [B_LO], hi + [B_HI]
            if step is not None:
                x0, lo, hi = x0 + [0.0], lo + [-3.0], hi + [3.0]
            r = least_squares(f, x0, bounds=(lo, hi))
            if best is None or r.cost < best.cost:
                best = r
        q0, D = math.exp(best.x[0]), math.exp(best.x[1]); b = best.x[2] if b_fixed is None else b_fixed
        c = best.x[-1] if step is not None else 0.0
        res = np.log(np.where(ok, vol, 1)) - np.log(mh((ages - start) / 12, q0, D, b)) - c * (ages >= (step if step is not None else n))
        new = ok.copy(); new[:start] = False
        new &= (res > OUT_LO) & (res < OUT_HI)
        if (new == use).all():
            break
        use = new
    idx = np.where(use)[0]
    yhat = np.log(mh((idx - start) / 12, q0, D, b)) + c * (idx >= (step if step is not None else n)); y = np.log(vol[idx])
    r2 = 1 - np.sum((y - yhat) ** 2) / max(np.sum((y - y.mean()) ** 2), 1e-12)
    return dict(start=int(start), q0=float(q0), D=float(D), b=float(b), c=float(c), step=step, used=use, r2=float(r2), n_used=int(len(idx)))


SHIFT_MIN, SHIFT_SPREAD = 0.25, 0.15   # level shift: recent months all off the earlier trend by 22%+ (log 0.25), consistently


def find_shift(vol, ok, start):
    """A level shift in the last 3-12 months: the months from s on all sit on one side of the trend fitted before s,
    by more than SHIFT_MIN in log terms and within SHIFT_SPREAD of each other. Returns the earliest such s (0-based) or None."""
    n = len(vol); found = None
    for s in range(n - 3, max(n - 13, start + 12) - 1, -1):
        pre = ok.copy(); pre[s:] = False
        P = fit_window(vol, pre, start)
        if P is None:
            continue
        post = [i for i in range(s, n) if ok[i]]
        if len(post) < 3:
            continue
        r = np.log(vol[post]) - np.log(mh((np.array(post) - P["start"]) / 12, P["q0"], P["D"], P["b"]))
        if (np.all(r < -SHIFT_MIN) or np.all(r > SHIFT_MIN)) and np.std(r) < SHIFT_SPREAD:
            found = s
    return found


def anchor(vol, ok, P):
    """Forecast anchor at the last reported month A (0-based n-1): fitted rate and the actual-anchored rate."""
    n = len(vol); A = n - 1
    tA = (A - P["start"]) / 12
    q_fit = float(mh(tA, P["q0"], P["D"], P["b"]) * math.exp(P["c"]))
    Da = decline_at(tA, P["D"], P["b"])
    last = [i for i in range(n - 1, -1, -1) if ok[i] and P["used"][i]][:3] or [i for i in range(n - 1, -1, -1) if ok[i]][:3]
    shape = np.array([float(mh((i - P["start"]) / 12, P["q0"], P["D"], P["b"])) * (math.exp(P["c"]) if P["step"] is not None and i >= P["step"] else 1)
                      for i in last])
    q_act = float(np.mean(vol[last]) * q_fit / shape.mean()) if last else q_fit
    return dict(age=n, q_fit=q_fit, q_actual=q_act, Da=float(Da), b=P["b"],
                last3=float(np.mean(vol[[i for i in range(max(0, n - 3), n) if ok[i]]])) if any(ok[max(0, n - 3):]) else float("nan"))


def level(vol, ok, A):
    """average of good months over ages A-2..A (1-based ages) or None"""
    v = [vol[i] for i in range(A - 3, A) if 0 <= i < len(vol) and ok[i]]
    return float(np.mean(v)) if len(v) >= 2 else None


def offset_view(subj, offs, k_max=HORIZON):
    """anchored offset bands for the subject's forecast months. subj = (vol, ok, lateral, A); offs = [(vol, ok, lateral)]"""
    vol_s, ok_s, lat_s, A = subj
    lv_s = level(vol_s, ok_s, A)
    paths, per10 = [], []
    for vol, ok, lat in offs:
        lv = level(vol, ok, A)
        if lv is None or lv_s is None:
            continue
        p = np.full(k_max, np.nan)
        for k in range(1, k_max + 1):
            i = A - 1 + k
            if i < len(vol) and ok[i]:
                p[k - 1] = vol[i] / lv * lv_s
        if np.isfinite(p).any():
            paths.append(p)
        per10.append(lv * 10000 / lat)
    out = dict(n=[], p25=[], p50=[], p75=[], offsets_at_anchor=len(per10))
    P = np.array(paths) if paths else np.full((0, k_max), np.nan)
    for k in range(k_max):
        col = P[:, k][np.isfinite(P[:, k])] if len(P) else np.array([])
        out["n"].append(int(len(col)))
        if len(col) >= MIN_OFFSETS:
            a, b_, c = np.percentile(col, [25, 50, 75])
            out["p25"].append(float(a)); out["p50"].append(float(b_)); out["p75"].append(float(c))
        else:
            out["p25"].append(None); out["p50"].append(None); out["p75"].append(None)
    cov = 0
    while cov < k_max and out["n"][cov] >= MIN_OFFSETS:
        cov += 1
    out["covered_months"] = cov
    if lv_s is not None and per10:
        s10 = lv_s * 10000 / lat_s
        out["level_per10k"] = s10; out["offsets_level_p50_per10k"] = float(np.median(per10))
        out["level_vs_offsets"] = float(s10 / np.median(per10))
        out["level_percentile"] = float(np.mean(np.array(per10) <= s10))
    return out


def offset_b(series, k):
    """b from the offsets' median shape (each offset normalized to its peak, median by month where 3+ report)"""
    norm = []
    for vol, ok in series:
        cand = [i for i in range(min(PEAK_SEARCH, len(vol))) if ok[i]]
        if not cand:
            continue
        p = max(cand, key=lambda i: vol[i])
        norm.append([vol[i] / vol[p] if ok[i] else np.nan for i in range(p, len(vol))])
    if not norm:
        return None
    L = max(map(len, norm)); med = []
    for m in range(L):
        v = [s[m] for s in norm if len(s) > m and np.isfinite(s[m])]
        if len(v) < MIN_OFFSETS:
            break
        med.append(np.median(v))
    if len(med) < 24:
        return None
    med = np.array(med)
    P = fit_window(med, med > 0, start=0)
    return P and P["b"]


def load(deal_dir):
    spec = importlib.util.spec_from_file_location("pdp_data", os.path.join(deal_dir, "pdp_data.py"))
    d = importlib.util.module_from_spec(spec); spec.loader.exec_module(d)
    return d


def group(z, zg):
    z = (z or "").upper().strip()
    return zg.get(z, z)


def main(deal_dir):
    d = load(deal_dir)
    DAYS, ZG, FF = getattr(d, "DAYS", {}), {k.upper(): v for k, v in getattr(d, "ZONE_GROUPS", {}).items()}, getattr(d, "FIT_FROM", {})
    offs_by_group = {}
    for a, v in d.OFFSETS.items():
        if a in d.SUBJECT:
            continue
        o, g = np.array(v[3], float), np.array(v[4], float)
        offs_by_group.setdefault(group(v[5], ZG), []).append(
            dict(api=a, lat=float(v[1]), oil=o, gas=g, ok_o=good_months(o, DAYS.get(a), v[2]), ok_g=good_months(g, DAYS.get(a), v[2])))
    # offsets: also drop single-month shut-ins without days (under 40% of the neighbouring median)
    for lst in offs_by_group.values():
        for x in lst:
            for vk, okk in (("oil", "ok_o"), ("gas", "ok_g")):
                v, ok = x[vk], x[okk]
                for i in range(len(v)):
                    nb = [v[j] for j in range(max(0, i - 3), min(len(v), i + 4)) if j != i and ok[j]]
                    if ok[i] and nb and v[i] < 0.4 * np.median(nb):
                        ok[i] = False
    gb = {}
    wells = []
    for a, v in d.SUBJECT.items():
        name, lat, first, oil, gas, zone = v[:6]
        g = group(zone, ZG); offs = offs_by_group.get(g, [])
        rec = dict(api=a, name=name, lateral=float(lat), first=first, zone=zone, zone_group=g, months=len(oil),
                   last_month=ym_add(first, len(oil) - 1), offsets_in_group=len(offs))
        for key, vol, ok_key in (("oil", np.array(oil, float), "ok_o"), ("gas", np.array(gas, float), "ok_g")):
            ok = good_months(vol, DAYS.get(a), first)
            start = ym_diff(first, FF[a]) if a in FF else None
            P0 = fit_window(vol, ok, start)
            step = find_shift(vol, ok, P0["start"]) if P0 else None
            P = fit_window(vol, ok, start, step=step) if step is not None else P0
            b_src = "fitted on this well"
            if P is None or P["n_used"] < MIN_FIT_MONTHS:
                if (g, key) not in gb:
                    gb[(g, key)] = offset_b([(x[key], x[ok_key]) for x in offs], key)
                if gb[(g, key)] is not None:
                    P = fit_window(vol, ok, start, b_fixed=gb[(g, key)], step=step)
                    b_src = f"offsets' median shape ({len(offs)} offsets): short history"
                else:
                    b_src = "fitted on this well (short history, no offset shape available: LOW CONFIDENCE)"
            if P is None:
                raise SystemExit(f"{name}: not enough good {key} months to fit")
            an = anchor(vol, ok, P)
            # the well's own level uses the months its fit kept (outliers and shut-ins out), so one spike does not move the bands
            view = offset_view((vol, P["used"], float(lat), len(vol)), [(x[key], x[ok_key], x["lat"]) for x in offs])
            rec[key] = dict(start_age=P["start"] + 1, q0=P["q0"], D=P["D"], b=P["b"], b_source=b_src, b_at_cap=bool(P["b"] >= B_HI - 1e-3),
                            level_shift=(None if P["step"] is None else dict(from_age=P["step"] + 1, change=math.exp(P["c"]) - 1)),
                            r2=P["r2"], months_used=P["n_used"], months_excluded=[i + 1 for i in range(len(vol)) if not P["used"][i] and i >= P["start"]],
                            anchor=an, offsets=view, history=[float(x) for x in vol])
        wells.append(rec)
    out = dict(deal=getattr(d, "DEAL", {}), last_month=d.LAST_MONTH, wells=wells,
               rules=dict(dmin=DMIN, b_bounds=[B_LO, B_HI], min_fit_months=MIN_FIT_MONTHS, days_ok=DAYS_OK, min_offsets=MIN_OFFSETS))
    json.dump(out, open(os.path.join(deal_dir, "pdp_fit.json"), "w"), indent=1)
    for w in wells:
        plot_well(w, deal_dir)
        o, g = w["oil"], w["gas"]
        print(f"{w['name']:<32} age {w['months']:>3}  oil qa {o['anchor']['q_fit']:>7,.0f} bbl/mo Da {o['anchor']['Da']:.0%} b {o['b']:.2f}"
              f" (R2 {o['r2']:.2f}, {o['offsets']['covered_months']} mo of offsets)   gas qa {g['anchor']['q_fit']:>8,.0f} Da {g['anchor']['Da']:.0%} b {g['b']:.2f}"
              + ("  [b from offsets]" if "offsets'" in o["b_source"] else "")
              + "".join(f"  [{k} level shift {w[k]['level_shift']['change']:+.0%} from month {w[k]['level_shift']['from_age']}]" for k in ("oil", "gas") if w[k]["level_shift"]))


def plot_well(w, deal_dir):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.4))
    for ax, key, unit, col in ((axes[0], "oil", "bbl/month", "#2E7D32"), (axes[1], "gas", "Mcf/month", "#C62828")):
        s = w[key]; h = np.array(s["history"]); A = len(h)
        ages = np.arange(1, A + 1)
        used = np.array([i + 1 not in s["months_excluded"] and i + 1 >= s["start_age"] for i in range(A)])
        ax.scatter(ages[used], np.where(h[used] > 0, h[used], np.nan), s=12, color=col, label="Actual (fit)")
        ax.scatter(ages[~used], np.where(h[~used] > 0, h[~used], np.nan), s=12, color="#9E9E9E", marker="x", label="Actual (left out)")
        t = np.arange(s["start_age"], A + 1)
        ls_ = s.get("level_shift")
        fl = mh((t - s["start_age"]) / 12, s["q0"], s["D"], s["b"]) * np.where(ls_ is not None and t >= (ls_ or {}).get("from_age", 1e9), 1 + (ls_ or {}).get("change", 0), 1)
        ax.plot(t, fl, color=col, lw=1, alpha=0.6, label="Fit" + (f" (level shift {ls_['change']:+.0%} from month {ls_['from_age']})" if ls_ else ""))
        k = np.arange(1, HORIZON + 1)
        an = s["anchor"]
        ax.plot(np.r_[A, A + k], np.r_[an["q_fit"], mh(k / 12, an["q_fit"], an["Da"], an["b"])], color="#EF6C00", lw=2, label=f"Forecast (b {an['b']:.2f})")
        ov = s["offsets"]
        p25, p50, p75 = (np.array([np.nan if x is None else x for x in ov[q]]) for q in ("p25", "p50", "p75"))
        ax.fill_between(A + k, p25, p75, color="#7E57C2", alpha=0.18, label="Offsets P25-P75 (anchored)")
        ax.plot(A + k, p50, color="#7E57C2", lw=1.5, ls="--", label=f"Offsets P50 ({ov['covered_months']} mo with {MIN_OFFSETS}+)")
        ax.set_yscale("log"); ax.set_xlabel("Months on production"); ax.set_ylabel(unit); ax.grid(alpha=0.3, which="both")
        ax.set_title(f"{key.title()}"); ax.legend(fontsize=7, loc="lower left")
    fig.suptitle(f"{w['name']} ({w['zone']}, {w['lateral']:,.0f} ft)", fontsize=11)
    fig.tight_layout()
    fig.savefig(os.path.join(deal_dir, f"pdp_{w['api']}.png"), dpi=110)
    plt.close(fig)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
