"""Monthly cash-flow engine mirroring the multi-well template (Wells, CF W1..W10, Package sheets).

Usage:
    from engine import load_model, run, max_premium
    m = load_model("model.xlsx")            # reads Wells, Inputs, Prices, TypeCurve (values)
    r = run(m, m["oil"], m["gas"], premium=0.2, deck="Strip")
Per-well overrides: run(m, oil, gas, wells=[...]) with dicts like the ones load_model returns.
"""
import datetime as dt
import numpy as np
import openpyxl

FLAT = {"$65/$3.00": (65, 3.0), "$75/$3.50": (75, 3.5), "$85/$4.00": (85, 4.0)}


def _mi(d, v0):
    """month index of date d relative to the valuation month v0 (blank dates -> far future)"""
    if d is None:
        return 10 ** 6
    return (d.year - v0.year) * 12 + d.month - v0.month


def load_model(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path)                 # formulas (for cells LibreOffice/Excel may not have cached)
    inp, W, pr, tc = wb["Inputs"], wb["Wells"], wb["Prices"], wb["TypeCurve"]
    v0 = inp["C37"].value
    g = lambda a: inp[a].value
    glob = dict(premium=g("C10") if isinstance(g("C10"), (int, float)) else 0.2, tc_mult=g("C11") or 1.0,
                haircut_default=g("C12"), oil_diff=g("C41"), oil_gath=g("C42"), gas_diff=g("C43"), shrink=g("C44"),
                ngl_yield=g("C45"), ngl_pct=g("C46"), gp=g("C47"), tax1=g("C50"), tax2=g("C51"), loe_fixed=g("C52"),
                loe_mcf=g("C53"), loe_bbl=g("C54"), wor=g("C55"), swd=g("C56"), pa=g("C57"), roy_deducts=g("C58"),
                overrun=g("C59") or 0.0, disc=g("C60"), deck="Strip", months=400, life=360,
                minerals=(g("C18") == "Minerals"), basin=g("C19"), price=float(g("C20") or 0), price_mult=1.0)
    close = g("C21") if isinstance(g("C21"), dt.datetime) else None
    glob["close_m"] = max(0, _mi(close, v0)) if close else 0
    wells = []
    for r in range(7, 17):
        wi = W[f"F{r}"].value or 0
        mnri = W[f"U{r}"].value if isinstance(W[f"U{r}"].value, (int, float)) else 0
        if not (mnri if glob["minerals"] else wi):
            continue
        hc = W[f"Q{r}"].value
        if not isinstance(hc, (int, float)):
            hc = glob["haircut_default"]
        wells.append(dict(name=W[f"C{r}"].value, api=str(W[f"D{r}"].value or "").replace("-", ""), lat=float(W[f"E{r}"].value), wi=float(wi), lease_nri=float(W[f"G{r}"].value),
                          drill=float(W[f"I{r}"].value or 0), compl=float(W[f"J{r}"].value or 0), fac=float(W[f"K{r}"].value or 0),
                          drill_m=_mi(W[f"N{r}"].value, v0), compl_m=_mi(W[f"O{r}"].value, v0), first_m=_mi(W[f"P{r}"].value, v0),
                          haircut=float(hc), tc_factor=float(W[f"R{r}"].value or 1.0), mnri=float(mnri)))
    if glob["minerals"]:                 # purchase price split by mineral NRI x lateral (Wells column W)
        tot = sum(w["mnri"] * w["lat"] for w in wells) or 1.0
        for w in wells:
            w["share"] = w["mnri"] * w["lat"] / tot
    strip = (np.array([pr.cell(r, 4).value for r in range(6, 406)], float), np.array([pr.cell(r, 5).value for r in range(6, 406)], float))
    oil = np.array([tc.cell(r, 3).value for r in range(6, 366)], float)
    gas = np.array([tc.cell(r, 4).value for r in range(6, 366)], float)
    env_oil = np.array([tc.cell(r, 5).value for r in range(6, 366)], float)
    env_gas = np.array([tc.cell(r, 6).value for r in range(6, 366)], float)
    return dict(glob=glob, wells=wells, strip=strip, oil=oil, gas=gas, env_oil=env_oil, env_gas=env_gas)


def well_cf(oil_tc, gas_tc, w, p, strip):
    n = p["months"]
    if p["deck"] == "Strip":
        wti, hh = strip[0][:n], strip[1][:n]
    else:
        o, g = FLAT[p["deck"]]
        wti, hh = np.full(n, float(o)), np.full(n, float(g))
    M = p.get("minerals", False)
    nri = w["mnri"] if M else w["wi"] * w["lease_nri"]
    capex_net = w["wi"] * (w["drill"] + w["compl"] + w["fac"])
    cf = np.zeros(n); T_prev = 0
    shift = p.get("delay", 0)
    first_m, compl_m = w["first_m"] + shift, w["compl_m"] + shift
    for m in range(n):
        D = m - first_m + 1 if m >= first_m else 0
        if 1 <= D <= p["life"]:
            k = p["tc_mult"] * w["tc_factor"] * (1 - (w["haircut"] if p.get("haircut") is None else p["haircut"])) * w["lat"] / 10000
            oil, gas = oil_tc[D - 1] * k, gas_tc[D - 1] * k
        else:
            oil = gas = 0.0
        ngl = gas / 1000 * p["ngl_yield"]; res = gas * (1 - p["shrink"])
        rev = oil * (wti[m] + p["oil_diff"]) + res * (hh[m] + p["gas_diff"]) + ngl * wti[m] * p["ngl_pct"]
        tax = rev * (p["tax1"] if D <= 36 else p["tax2"])
        ded = gas * p["gp"] + oil * p["oil_gath"]
        loe = (p["loe_fixed"] if 1 <= D <= p["life"] else 0) + gas * p["loe_mcf"] + oil * p["loe_bbl"] + oil * p["wor"] * p["swd"]
        dfac = w["lease_nri"] if p["roy_deducts"] == 1 else 1.0
        S = w["lease_nri"] * (rev - tax) - dfac * ded - loe
        T = 0 if (D < 1 or D > p["life"]) else (1 if D == 1 else (1 if (T_prev == 1 and (D <= 24 or S > 0)) else 0))
        if M:                            # minerals: revenue less tax at the mineral NRI; price paid at closing
            prem = p["price"] * p["price_mult"] * w.get("share", 1.0) if m == p.get("close_m", 0) else 0.0
            cf[m] = nri * rev * T - nri * tax * T - prem
            T_prev = T
            continue
        capex = w["wi"] * ((m == w["drill_m"]) * w["drill"] + (m == compl_m) * (w["compl"] + w["fac"])) * (1 + p["overrun"])
        prem = p["premium"] * capex_net if m == w["drill_m"] else 0.0
        pa = w["wi"] * p["pa"] if (T_prev == 1 and T == 0) else 0.0
        cf[m] = nri * rev * T - nri * tax * T - w["wi"] * dfac * ded * T - w["wi"] * loe * T - capex - prem - pa
        T_prev = T
    return cf


def irr_m(cf):
    if cf.sum() < 0:
        return None
    lo, hi = 0.0, 1.0
    t = np.arange(len(cf))
    f = lambda r: float(np.sum(cf * np.exp(-t * np.log1p(r))))
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if f(lo) * f(mid) <= 0 else (mid, hi)
    return (1 + (lo + hi) / 2) ** 12 - 1


def run(model, oil_tc, gas_tc, wells=None, **over):
    p = dict(model["glob"], **over)
    wells = wells if wells is not None else model["wells"]
    cfs = [well_cf(oil_tc, gas_tc, w, p, model["strip"]) for w in wells]
    pkg = sum(cfs)
    t = np.arange(p["months"])
    df = 1 / (1 + p["disc"]) ** ((t + 0.5) / 12); df20 = 1 / 1.2 ** ((t + 0.5) / 12)
    first = min(w["first_m"] for w in wells) + p.get("delay", 0)
    cum = np.cumsum(pkg)
    payout = next((m - first + 1 for m in range(1, len(cum)) if m >= first and cum[m] >= 0 and cum[m - 1] < 0), None)
    pos, neg = pkg[pkg > 0].sum(), -pkg[pkg < 0].sum()
    wirr = [irr_m(c) for c in cfs]
    ok = [x for x in wirr if x is not None]
    return dict(irr=irr_m(pkg), moic=pos / neg if neg else float("inf"), npv10=float((pkg * df).sum()), pv20=float((pkg * df20).sum()),
                cash=float(pkg.sum()), payout=payout, well_irr=wirr,
                min_well_irr=min(ok) if len(ok) == len(wirr) and ok else None, max_well_irr=max(ok) if ok else None)


def max_price(model, oil_tc, gas_tc, target=0.20, wells=None, **over):
    """Minerals: the largest purchase price ($) that still earns the target IRR (bisection on dollars)."""
    def ok(price):
        r = run(model, oil_tc, gas_tc, wells=wells, **dict(over, price=price, price_mult=1.0))
        return r["irr"] >= target if r["irr"] is not None else r["cash"] > 0     # None + positive cash = IRR off the scale

    lo, hi = 0.0, 1000.0
    while ok(hi) and hi < 1e10:          # grow the bracket until the IRR falls below target
        lo, hi = hi, hi * 4
    for _ in range(60):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if ok(mid) else (lo, mid)
    return lo


def max_premium(model, oil_tc, gas_tc, target=0.20, wells=None, **over):
    lo, hi = -0.5, 2.0
    for _ in range(60):
        mid = (lo + hi) / 2
        r = run(model, oil_tc, gas_tc, wells=wells, **dict(over, premium=mid))["irr"]
        lo, hi = (mid, hi) if (r is not None and r >= target) else (lo, mid)
    return lo
