"""Monthly cash-flow engine mirroring the multi-well template (Wells, CF W1..W10, Package sheets).

Usage:
    from engine import load_model, run, max_premium
    m = load_model("model.xlsx")            # reads Wells, Inputs, Prices, TypeCurve (values)
    r = run(m, m["oil"], m["gas"], premium=0.2, deck="Strip")
Per-well overrides: run(m, oil, gas, wells=[...]) with dicts like the ones load_model returns.
Multi-zone: m["zones"] = {zone: (oil, gas)} from the 'Zone Curves' tab; a well with w["zone"] set uses that curve,
any other well uses the oil / gas curves passed to run (the TypeCurve tab = zone 1). run(..., zones={...}) overrides.
Lateral length: curves are per 10,000 ft; each well's curve is moved to its lateral with lateral.scale_curve (linear
within +/- the band, Inputs C24; otherwise months 1..C25 at a factor phased from 100% at C24 ft to C26 at C27 ft and C28 at C29 ft and the decline re-solved to the linear EUR),
the same arithmetic as the workbook's 'Lateral Curves' tab.
"""
import datetime as dt
import hashlib
import numpy as np
import openpyxl
import lateral as LAT

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
    num = lambda a, d: g(a) if isinstance(g(a), (int, float)) and not isinstance(g(a), bool) else d
    glob.update(lat_band=float(num("C24", LAT.BAND)), lat_months=int(num("C25", LAT.EARLY_MONTHS)), lat_pct=float(num("C26", LAT.EARLY_PCT)),
                lat_full=float(num("C27", LAT.FULL_AT)), lat_pct2=float(num("C28", LAT.EARLY_PCT2)),
                lat_full2=float(num("C29", LAT.FULL_AT2)))
    close = g("C21") if isinstance(g("C21"), dt.datetime) else None
    glob["close_m"] = max(0, _mi(close, v0)) if close else 0
    zones, zone1 = {}, None
    if "Zone Curves" in wb.sheetnames:
        zc = wb["Zone Curves"]
        for k in range(6):
            name = zc.cell(4, 3 + 2 * k).value
            if name in (None, ""):
                continue
            name = str(name).strip()
            zone1 = zone1 or (name if k == 0 else None)
            zones[name] = (np.array([zc.cell(r, 3 + 2 * k).value or 0 for r in range(6, 366)], float),
                           np.array([zc.cell(r, 4 + 2 * k).value or 0 for r in range(6, 366)], float))
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
                          haircut=float(hc), tc_factor=float(W[f"R{r}"].value or 1.0), mnri=float(mnri),
                          zone=_zone(W[f"X{r}"].value, zone1)))
    if glob["minerals"]:                 # purchase price split by mineral NRI x lateral (Wells column W)
        tot = sum(w["mnri"] * w["lat"] for w in wells) or 1.0
        for w in wells:
            w["share"] = w["mnri"] * w["lat"] / tot
    strip = (np.array([pr.cell(r, 4).value for r in range(6, 406)], float), np.array([pr.cell(r, 5).value for r in range(6, 406)], float))
    oil = np.array([tc.cell(r, 3).value for r in range(6, 366)], float)
    gas = np.array([tc.cell(r, 4).value for r in range(6, 366)], float)
    env_oil = np.array([tc.cell(r, 5).value for r in range(6, 366)], float)
    env_gas = np.array([tc.cell(r, 6).value for r in range(6, 366)], float)
    return dict(glob=glob, wells=wells, strip=strip, oil=oil, gas=gas, env_oil=env_oil, env_gas=env_gas, zones=zones, zone1=zone1)


def _zone(v, zone1):
    """Wells column X: blank or zone 1 = the TypeCurve curve (None); anything else = a 'Zone Curves' name."""
    if v in (None, ""):
        return None
    v = str(v).strip()
    return None if v == zone1 else v


_SC = {}


def scaled(c, lat, p):
    """Curve c (per 10,000 ft) moved to lateral lat on the lateral rule (cached)."""
    c = np.asarray(c, float)
    key = (hashlib.md5(c.tobytes()).hexdigest(), float(lat), p.get("lat_band"), p.get("lat_months"), p.get("lat_pct"), p.get("lat_full"), p.get("lat_pct2"), p.get("lat_full2"))
    if key not in _SC:
        if len(_SC) > 5000:
            _SC.clear()
        _SC[key] = LAT.scale_curve(c, LAT.BASIS, lat, p.get("lat_band"), p.get("lat_months"), p.get("lat_pct"), p.get("lat_full"),
                                   p.get("lat_pct2"), p.get("lat_full2"))
    return _SC[key]


def well_cf(oil_tc, gas_tc, w, p, strip, zones=None):
    n = p["months"]
    if w.get("zone"):
        oil_tc, gas_tc = (zones or {})[w["zone"]]        # KeyError = a zone with no curve (Summary check D73)
    if p["deck"] == "Strip":
        wti, hh = strip[0][:n], strip[1][:n]
    else:
        o, g = FLAT[p["deck"]]
        wti, hh = np.full(n, float(o)), np.full(n, float(g))
    oil_tc, gas_tc = scaled(oil_tc, w["lat"], p)[0], scaled(gas_tc, w["lat"], p)[0]
    M = p.get("minerals", False)
    nri = w["mnri"] if M else w["wi"] * w["lease_nri"]
    capex_net = w["wi"] * (w["drill"] + w["compl"] + w["fac"])
    cf = np.zeros(n); T_prev = 0
    shift = p.get("delay", 0)
    first_m, compl_m = w["first_m"] + shift, w["compl_m"] + shift
    for m in range(n):
        D = m - first_m + 1 if m >= first_m else 0
        if 1 <= D <= p["life"]:
            k = p["tc_mult"] * w["tc_factor"] * (1 - (w["haircut"] if p.get("haircut") is None else p["haircut"]))
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


def run(model, oil_tc, gas_tc, wells=None, zones=None, **over):
    p = dict(model["glob"], **over)
    wells = wells if wells is not None else model["wells"]
    zones = zones if zones is not None else model.get("zones", {})
    cfs = [well_cf(oil_tc, gas_tc, w, p, model["strip"], zones) for w in wells]
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


PREM_LO, PREM_HI = -1.0, 5.0             # premium solve range (-100% to +500% of net capex)


def max_premium(model, oil_tc, gas_tc, target=0.20, wells=None, **over):
    lo, hi = PREM_LO, PREM_HI
    for _ in range(60):
        mid = (lo + hi) / 2
        r = run(model, oil_tc, gas_tc, wells=wells, **dict(over, premium=mid))["irr"]
        lo, hi = (mid, hi) if (r is not None and r >= target) else (lo, mid)
    return lo
