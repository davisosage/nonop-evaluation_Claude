"""Downside for the PDP package: every stream forecast above its offsets is re-declined from its actual anchor so its
cumulative over the full overlap equals the offsets anchored P50. Run from toolkit/: python ../deals/pdp-package-jan2027/downside.py"""
import json, os, sys, numpy as np, openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "toolkit"))
import write_pdp as W, pdp_fit as F, pdp_engine as E
from scipy.optimize import brentq
D = "../deals/pdp-package-jan2027"
dd = F.load(D); fit = json.load(open(D + "/pdp_fit.json"))
wb = openpyxl.load_workbook(W.TEMPLATE)
setup = dict(basin=dd.DEAL["basin"], eff=W.d(dd.DEAL["effective"]), price=0.0, dtype=dd.DEAL["deal_type"])
inp, wells, prices = W.engine_inputs(wb, dd, fit, setup, {}, dd.OVERRIDES, {}, dd.WELL_BASIN, {})
base = E.run(inp, wells, prices, "Strip")
down = []; changed = []
for w, f in zip(wells, fit["wells"]):
    w = dict(w)
    for s, key in (("o", "oil"), ("g", "gas")):
        ov = f[key]["offsets"]; n = ov["covered_months"]
        if n < 12 or w["mult"] == 0: continue
        q, Da, b = w[s + "_q"], w[s + "_da"], w[s + "_b"]
        p50 = sum(ov["p50"][:n]); k = np.arange(1, n + 1) / 12
        r = F.mh(k, q, Da, b).sum() / p50
        if r <= 1: continue
        qa = f[key]["anchor"]["q_actual"]
        g = lambda x: F.mh(k, qa, x, b).sum() - p50
        if g(0.001) <= 0: qa = qa * p50 / F.mh(k, qa, 0.001, b).sum(); Dn = 0.001   # actual anchor already below: scale the level
        else: Dn = brentq(g, 0.001, 50)
        w[s + "_q"], w[s + "_da"] = qa, Dn
        changed.append((f["name"], key, round(r, 2), n))
    down.append(w)
dn = E.run(inp, down, prices, "Strip")
print(f"base max {base['max_price']:,.0f} pv10 {base['pv10']:,.0f} | downside max {dn['max_price']:,.0f} pv10 {dn['pv10']:,.0f} "
      f"({dn['max_price']/base['max_price']-1:+.1%}); streams changed {len(changed)}")
# by area
from collections import defaultdict
A = defaultdict(lambda: [0.0, 0.0]); j = np.arange(E.NM); wti, hh = E.price_path(inp, prices, "Strip")
for w0, w1, f in zip(wells, down, fit["wells"]):
    a = dd.WELL_AREA[f["api"]]
    for i, w in enumerate((w0, w1)):
        net = E.well_cf(w, inp, wti, hh)[0]; A[a][i] += float((net / 1.2 ** (j / 12)).sum())
for a, (b0, b1) in sorted(A.items(), key=lambda x: -x[1][0])[:8]: print(f"  {a:<20} {b0:>9,.0f} -> {b1:>9,.0f} ({b1/b0-1:+.0%})")
json.dump(dict(base=base["max_price"], down=dn["max_price"], down_pv10=dn["pv10"], changed=changed), open(D + "/downside.json", "w"), indent=1)
