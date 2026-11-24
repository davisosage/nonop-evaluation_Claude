"""Monthly cash-flow engine mirroring the PDP Evaluation Template (Wells, M Oil / M Gas / M RevAT / M On / M Net, Package).

    from pdp_engine import run
    r = run(inp, wells, prices, deck="Strip", mult=1.0)   # r: pv10, max_price, irr, moic, net_oil, ...

inp    dict of resolved Inputs values (keys as build_pdp_template.I: odiff, shrink, tax1, loe_fix, dm, life_m, ...)
wells  list of dicts: active, age, k0, mult, nri, wiu, dfac, loe, lnri, hist_oil, hist_gas, and for s in (o, g): s_q, s_da, s_b;
       optional p = the well's basin values (odiff, ogath, gdiff, shrink, ngly, nglp, gp, tax1, tax2, loe_mcf, loe_bbl, wor, swd, pa);
       without p the deal-level inp values are used
prices dict(start="YYYY-MM", wti=[...], hh=[...]) strip by month (flat after the end)

Check a recalculated workbook against the engine:  python pdp_engine.py <model.xlsx>
"""
import math, sys
import numpy as np

NM = 480
FLAT = {"$65/$3.00": (65, 3.0), "$75/$3.50": (75, 3.5), "$85/$4.00": (85, 4.0)}


def ym_diff(a, b):
    return (int(b[:4]) - int(a[:4])) * 12 + int(b[5:7]) - int(a[5:7])


def forecast(q, da, b, dm, k):
    """volume at month k (array, k >= 1) after the anchor: hyperbolic then terminal exponential"""
    b = max(b, 0.001)
    tsw = 0.0 if da <= dm else (da / dm - 1) / (b * da)
    qsw = q * (1 + b * da * tsw) ** (-1 / b)
    t = k / 12
    return np.where(t <= tsw, q * (1 + b * da * np.minimum(t, tsw)) ** (-1 / b), qsw * np.exp(-dm * (t - tsw)))


def price_path(inp, prices, deck):
    eff = inp["eff"]                                        # "YYYY-MM"
    i0 = ym_diff(prices["start"], eff)
    idx = np.clip(np.arange(NM) + i0, 0, len(prices["wti"]) - 1)
    if deck == "Strip":
        return np.array(prices["wti"], float)[idx], np.array(prices["hh"], float)[idx]
    o, g = FLAT[deck]
    return np.full(NM, float(o)), np.full(NM, float(g))


def irr_m(cf):
    if cf.sum() < 0:
        return None
    t = np.arange(len(cf))
    f = lambda r: float(np.sum(cf * np.exp(-t * np.log1p(r))))
    lo, hi = 0.0, 1.0
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        lo, hi = (lo, mid) if f(lo) * f(mid) <= 0 else (mid, hi)
    return (1 + (lo + hi) / 2) ** 12 - 1


def well_cf(w, inp, wti, hh, mult=1.0, econ=None):
    j = np.arange(NM)
    k = w["k0"] + j
    vols = {}
    for s, hk in (("o", "hist_oil"), ("g", "hist_gas")):
        v = np.zeros(NM)
        if w["active"]:
            kk = np.maximum(k, 1)
            v = forecast(w[s + "_q"], w[s + "_da"], w[s + "_b"], inp["dm"], kk)
            hist = w.get(hk) or []
            for jj in np.where(k < 1)[0]:
                a = max(1, w["age"] + k[jj])
                v[jj] = hist[a - 1] if a - 1 < len(hist) else 0.0
            v = v * mult * inp["mult"] * w["mult"]
            v[j >= inp["life_m"]] = 0.0
        vols[s] = v
    o, g = vols["o"], vols["g"]
    P = w.get("p") or inp
    rate = np.where(w["age"] + k > 36, P["tax2"], P["tax1"])
    rev_at = (o * (wti + P["odiff"]) + g * (1 - P["shrink"]) * (hh + P["gdiff"]) + g / 1000 * P["ngly"] * wti * P["nglp"]) * (1 - rate)
    ded = g * P["gp"] + o * P["ogath"]
    loe = w["loe"] + g * P["loe_mcf"] + o * (P["loe_bbl"] + P["wor"] * P["swd"])
    op8 = w["lnri"] * rev_at - w["dfac"] * ded - loe
    econ = inp["econ"] if econ is None else econ
    on = np.zeros(NM)
    if w["active"]:
        if not econ:
            on = ((o + g) > 0).astype(float)
        else:
            prev = 1.0
            for jj in range(NM):
                on[jj] = 1.0 if (prev == 1.0 and op8[jj] > 0) else 0.0
                prev = on[jj]
    ev = np.r_[w["active"] - on[0], on[:-1] - on[1:]]
    net = on * (w["nri"] * rev_at - w["wiu"] * w["dfac"] * ded - w["wiu"] * loe) - w["wiu"] * P["pa"] * ev
    return net, on, o, g


def run(inp, wells, prices, deck="Strip", mult=1.0, econ=None, price=None):
    wti, hh = price_path(inp, prices, deck)
    R = np.zeros(NM); net_o = net_g = 0.0; per = []
    for w in wells:
        net, on, o, g = well_cf(w, inp, wti, hh, mult, econ)
        R += net; per.append(net)
        net_o += float((o * on).sum() * w["nri"]); net_g += float((g * on).sum() * w["nri"])
    j = np.arange(NM)
    price = inp["price"] if price is None else price
    cf = R.copy(); cf[0] -= price
    pos, neg = cf[cf > 0].sum(), -cf[cf < 0].sum()
    return dict(pv10=float((R / (1 + inp["disc"]) ** ((j + 0.5) / 12)).sum()),
                pv15=float((R / 1.15 ** ((j + 0.5) / 12)).sum()), pv20=float((R / 1.2 ** ((j + 0.5) / 12)).sum()),
                max_price=float((R / (1 + inp["tirr"]) ** (j / 12)).sum()), cash=float(R.sum()),
                irr=irr_m(cf) if price > 0 else None, moic=(pos / neg) if (neg and price > 0) else None,
                net_oil=net_o, net_gas=net_g, well_pv10=[float((n / (1 + inp["disc"]) ** ((j + 0.5) / 12)).sum()) for n in per])


# ------------------------------------------------------------------ check a recalculated workbook
def from_workbook(path):
    import openpyxl
    import build_pdp_template as T
    wb = openpyxl.load_workbook(path, data_only=True)
    ws, W, H, P = wb["Inputs"], wb["Wells"], wb["History"], wb["Prices"]
    inp = {k: ws[c].value for k, c in T.I.items()}
    e = inp["eff"]; inp["eff"] = f"{e.year:04d}-{e.month:02d}"
    wells = []
    for i in range(T.NW):
        r = T.W0 + i
        g = lambda k: W.cell(r, T.WC[k]).value
        w = dict(active=int(g("active") or 0), age=int(g("age") or 0), k0=int(g("k0") or 0), mult=float(g("mult") or 0), nri=float(g("nri") or 0),
                 wiu=float(g("wiu") or 0), dfac=float(g("dfac") or 1), loe=float(g("loe") or 0), lnri=float(g("lnri") or 0),
                 hist_oil=[H.cell(T.H0 + a, 2 * (i + 1) + 1).value or 0 for a in range(T.NH)],
                 hist_gas=[H.cell(T.H0 + a, 2 * (i + 1) + 2).value or 0 for a in range(T.NH)])
        for s in ("o", "g"):
            w[s + "_q"], w[s + "_da"], w[s + "_b"] = float(g(s + "_q") or 0), float(g(s + "_dau") or 0), float(g(s + "_bu") or 0.001)
        w["p"] = {k: float(g("p_" + k) or 0) for k in T.PARAMS}
        wells.append(w)
    d0 = P["B6"].value
    prices = dict(start=f"{d0.year:04d}-{d0.month:02d}", wti=[P.cell(r, 5).value for r in range(6, 606)], hh=[P.cell(r, 6).value for r in range(6, 606)])
    return inp, wells, prices, wb


if __name__ == "__main__":
    inp, wells, prices, wb = from_workbook(sys.argv[1])
    deck = wb["Summary"]["C5"].value
    r = run(inp, wells, prices, deck="Strip" if inp["deck"] == 1 else deck)
    S = wb["Summary"]
    ok = True
    for lab, cell, key in (("PV10", "C12", "pv10"), ("Max price", "C11", "max_price"), ("Net oil MBO", "C19", "net_oil"), ("Net gas MMcf", "C20", "net_gas")):
        wbv = S[cell].value; ev = r[key] / (1000 if key in ("net_oil", "net_gas") else 1)
        d = abs(ev - wbv) / max(abs(wbv), 1)
        ok &= d < 1e-4
        print(f"{lab:<12} workbook {wbv:>16,.2f}  engine {ev:>16,.2f}  diff {d:.2e}")
    print("engine ties" if ok else "ENGINE DOES NOT TIE")
