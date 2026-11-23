"""Write a deal's type curve and scenario grid into a copy of the multi-well template.

Usage: python write_model.py <model_in.xlsx> <deal_dir> <model_out.xlsx> [--op-ratio OIL GAS]
  * <model_in.xlsx>: the template with the deal's Wells / Inputs / Prices already filled in
  * <deal_dir>: holds deal_data.py, tc_summary.json, tc_curves.csv (from build_tc.py)
Solves Summary C7 (premium) so the live-case IRR (C12) is 20%. Writes TypeCurve values and fit block, the TC vs offsets charts and monthly table (plot_tc.py), the Summary scenario grid (rows 25-54), the max-premium rows
(56-65) and the engine tie (Summary E75). Recalculate the output (LibreOffice / Excel) before use.
Multi-zone deals (tc_zones.json from build_tc.py): every zone curve goes on the 'Zone Curves' tab (zone 1 = TypeCurve), with
a fit summary and charts for zones 2-6; each well uses the zone named in Wells column X. Curve sensitivities (b, fixed b,
Enverus P10/P50/P90, 10-mile) are applied to every zone at once, each zone with its own fit.
Lateral length: offsets are normalized to 10,000 ft on the lateral rule before the fit (build_tc.py); each well's s / Tsw / g go on 'Lateral Curves' rows 11-13 (lateral.py).
The workbook must have been recalculated once (cached values) so the engine can read Wells and Inputs.
"""
import json, math, os, sys
from copy import copy as _copy
import numpy as np
import openpyxl
import engine as E
import build_tc as bt
import lateral as LAT
import plot_tc

src, deal, out = sys.argv[1], sys.argv[2], sys.argv[3]
op = None
if "--op-ratio" in sys.argv:
    k = sys.argv.index("--op-ratio"); op = (float(sys.argv[k + 1]), float(sys.argv[k + 2]))

bt.load_deal(deal)
S = json.load(open(os.path.join(deal, "tc_summary.json")))
C = np.loadtxt(os.path.join(deal, "tc_curves.csv"), delimiter=",", skiprows=1)
oil, gas, eo, eg = C[:, 1], C[:, 2], C[:, 3], C[:, 4]
M = E.load_model(src)
o, g = S["oil"], S["gas"]

# ---- zones (multi-zone deals): zone 1 = the TypeCurve curve; the others come from tc_zones.json
_zp = os.path.join(deal, "tc_zones.json")
ZL = json.load(open(_zp))["zones"] if os.path.exists(_zp) else []
ZC, Z10 = {}, {}
if ZL:
    _A = np.loadtxt(os.path.join(deal, "tc_zone_curves.csv"), delimiter=",", skiprows=1, ndmin=2)
    _A10 = np.loadtxt(os.path.join(deal, "tc_zone_curves_10mi.csv"), delimiter=",", skiprows=1, ndmin=2)
    for _k, _z in enumerate(ZL):
        ZC[_z["name"]] = (_A[:, 1 + 2 * _k], _A[:, 2 + 2 * _k])
        Z10[_z["name"]] = (_A10[:, 1 + 2 * _k], _A10[:, 2 + 2 * _k]) if "ten_mile" in _z else None
Z1 = ZL[0]["name"] if ZL else None
M["zones"], M["zone1"] = {n: c for n, c in ZC.items() if n != Z1}, Z1
for _w in M["wells"]:
    if _w.get("zone") == Z1:
        _w["zone"] = None
_missing = sorted({_w["zone"] for _w in M["wells"] if _w.get("zone") and _w["zone"] not in M["zones"]})
if _missing:
    raise SystemExit(f"Wells column X names zones with no curve: {_missing}; zones built: {list(ZC) or 'none (no ZONES in deal_data.py)'}")
ZOTHER = [z for z in ZL if z["name"] != Z1]


def zvar(fn):
    """A sensitivity applied to every non-primary zone: fn(zone record, (oil, gas)) -> (oil, gas)."""
    return {z["name"]: fn(z, ZC[z["name"]]) for z in ZOTHER}


def _lref(rec):
    return float((rec.get("lateral") or {}).get("model_lateral") or LAT.BASIS)


def _zb(z, mean, b):
    return bt.to_basis(bt.fit_tc(np.array(mean), b=b)[0], _lref(z))


def _cap(c, med):
    return c * min(1, med / c.sum()) if c.sum() else c


Z_BM = zvar(lambda z, c: (_zb(z, z["mean_oil"], max(z["oil"]["b"] - 0.3, 0.05)), _zb(z, z["mean_gas"], max(z["gas"]["b"] - 0.3, 0.05))))
Z_BP = zvar(lambda z, c: (_zb(z, z["mean_oil"], min(z["oil"]["b"] + 0.2, 2.0)), _zb(z, z["mean_gas"], min(z["gas"]["b"] + 0.2, 2.0))))
Z_B1 = zvar(lambda z, c: (_cap(_zb(z, z["mean_oil"], 1.0), z["enverus"]["oil_median"]), _cap(_zb(z, z["mean_gas"], 1.0), z["enverus"]["gas_median"])))
Z_P = {i: zvar(lambda z, c, i=i: (c[0] * z["enverus"]["oil_p90_p50_p10"][i] * 1000 / c[0].sum(),
                                  c[1] * z["enverus"]["gas_p90_p50_p10"][i] * 1e6 / c[1].sum())) for i in (0, 1, 2)}
Z_10 = zvar(lambda z, c: Z10[z["name"]] if Z10[z["name"]] is not None else c)

# ---- sensitivities on the curve
mo, mg = np.array(S["mean_oil"]), np.array(S["mean_gas"])


LREF = _lref(S)


def oil_b(b):                       # multi-phase refit (ramp + decline) with b fixed, at the modeled lateral -> 10,000 ft
    return bt.to_basis(bt.fit_tc(mo, b=b)[0], LREF)


def gas_b(b):
    return bt.to_basis(bt.fit_tc(mg, b=b)[0], LREF)


b0 = o["b"]
o_bm, o_bp = oil_b(max(b0 - 0.3, 0.05)), oil_b(min(b0 + 0.2, 2.0))
gb0 = g["b"]
g_bm, g_bp = gas_b(max(gb0 - 0.3, 0.05)), gas_b(min(gb0 + 0.2, 2.0))
o_b1 = oil_b(1.0); o_b1 = o_b1 * min(1, S["enverus"]["oil_median"] / o_b1.sum())
g_b1 = gas_b(1.0); g_b1 = g_b1 * min(1, S["enverus"]["gas_median"] / g_b1.sum())
ep_o, ep_g = S["enverus"]["oil_p90_p50_p10"], S["enverus"]["gas_p90_p50_p10"]
hi_o, hi_g = ep_o[0] * 1000 / oil.sum(), ep_g[0] * 1e6 / gas.sum()      # P10 (high)
lo_o, lo_g = ep_o[2] * 1000 / oil.sum(), ep_g[2] * 1e6 / gas.sum()      # P90 (low)

R = lambda **kw: E.run(M, kw.pop("o", oil), kw.pop("g", gas), **kw)
rows = [
    ("Strip, premium 0%", dict(premium=0)), ("Strip, premium 10%", dict(premium=0.1)), ("Strip, premium 20%", dict(premium=0.2)),
    ("Strip, premium 30%", dict(premium=0.3)),
    ("$65/$3.00, premium 0%", dict(deck="$65/$3.00", premium=0)), ("$65/$3.00, premium 10%", dict(deck="$65/$3.00", premium=0.1)),
    ("$65/$3.00, premium 20%", dict(deck="$65/$3.00", premium=0.2)), ("$65/$3.00, premium 30%", dict(deck="$65/$3.00", premium=0.3)),
    ("$75/$3.50, premium 0%", dict(deck="$75/$3.50", premium=0)), ("$75/$3.50, premium 20%", dict(deck="$75/$3.50", premium=0.2)),
    ("$85/$4.00, premium 0%", dict(deck="$85/$4.00", premium=0)), ("$85/$4.00, premium 20%", dict(deck="$85/$4.00", premium=0.2)),
    ("Strip, TC 120%, premium 20%", dict(tc_mult=1.2, premium=0.2)), ("Strip, TC 80%, premium 20%", dict(tc_mult=0.8, premium=0.2)),
    ("Strip, Cost +10%, premium 20%", dict(overrun=0.1, premium=0.2)),
    ("Strip, TC 80% + cost +10%, premium 20%", dict(tc_mult=0.8, overrun=0.1, premium=0.2)),
    ("$65/$3.00, TC 120%, premium 20%", dict(deck="$65/$3.00", tc_mult=1.2, premium=0.2)),
    ("$65/$3.00, TC 80%, premium 20%", dict(deck="$65/$3.00", tc_mult=0.8, premium=0.2)),
    ("$65/$3.00, Cost +10%, premium 20%", dict(deck="$65/$3.00", overrun=0.1, premium=0.2)),
    ("$65/$3.00, TC 80% + cost +10%, premium 20%", dict(deck="$65/$3.00", tc_mult=0.8, overrun=0.1, premium=0.2)),
    (f"Strip, Enverus P50 (oil {eo.sum()/oil.sum():.2f}x, gas {eg.sum()/gas.sum():.2f}x of base), premium 20%", dict(o=eo, g=eg, zones=Z_P[1], premium=0.2)),
    (f"Strip, Enverus P10 (oil {hi_o:.2f}x, gas {hi_g:.2f}x), premium 20%", dict(o=oil * hi_o, g=gas * hi_g, zones=Z_P[0], premium=0.2)),
    (f"Strip, Enverus P90 (oil {lo_o:.2f}x, gas {lo_g:.2f}x), premium 20%", dict(o=oil * lo_o, g=gas * lo_g, zones=Z_P[2], premium=0.2)),
    ("Strip, fixed b 1.0 both streams (capped at Enverus median), premium 20%", dict(o=o_b1, g=g_b1, zones=Z_B1, premium=0.2)),
    (f"Strip, b -0.3 (oil {max(b0-0.3,0.05):.2f} / gas {max(gb0-0.3,0.05):.2f}), premium 20%", dict(o=o_bm, g=g_bm, zones=Z_BM, premium=0.2)),
    (f"Strip, b +0.2 (oil {min(b0+0.2,2):.2f} / gas {min(gb0+0.2,2):.2f}), premium 20%", dict(o=o_bp, g=g_bp, zones=Z_BP, premium=0.2)),
    ("Strip, 15% child haircut (all wells), premium 20%", dict(haircut=0.15, premium=0.2)),
    ("Strip, +3 mo completion delay, premium 20%", dict(delay=3, premium=0.2)),
]
if op:
    rows.append((f"Strip, operator early rates (oil x{op[0]:.2f} / gas x{op[1]:.2f}), premium 20%", dict(o=oil * op[0], g=gas * op[1], premium=0.2)))
T10 = None
if "ten_mile" in S and os.path.exists(os.path.join(deal, "tc_curves_10mi.csv")):
    C10 = np.loadtxt(os.path.join(deal, "tc_curves_10mi.csv"), delimiter=",", skiprows=1)
    T10 = (C10[:, 1], C10[:, 2], S["ten_mile"])
    rows.insert(23, (f"Strip, 10-mile curve{'s (all zones' if ZOTHER else ' ('}{'; zone 1 ' if ZOTHER else ''}{T10[2]['wells']} wells; oil {C10[:,1].sum()/1000:.0f} MBO, gas {C10[:,2].sum()/1e6:.2f} Bcf), premium 20%",
                     dict(o=T10[0], g=T10[1], zones=Z_10, premium=0.2)))
rows = rows[:30]
MP = [("Max premium for 20% IRR, strip", 0.20, {}), ("Max premium for 10% IRR, $65/$3.00", 0.10, dict(deck="$65/$3.00")),
      ("Max premium for 20% IRR, $65/$3.00", 0.20, dict(deck="$65/$3.00")),
      ("Max premium for 20% IRR, strip, Enverus P90 curve", 0.20, dict(o=oil * lo_o, g=gas * lo_g, zones=Z_P[2])),
      ("Max premium for 20% IRR, strip, fixed b 1.0", 0.20, dict(o=o_b1, g=g_b1, zones=Z_B1)),
      ("Max premium for 20% IRR, strip, b -0.3", 0.20, dict(o=o_bm, g=g_bm, zones=Z_BM)),
      ("Max premium for 20% IRR, strip, +3 mo delay", 0.20, dict(delay=3)),
      ("Max premium for 20% IRR, strip, type curve 80%", 0.20, dict(tc_mult=0.8)),
      ("Max premium for 20% IRR, strip, 15% child haircut", 0.20, dict(haircut=0.15))]
if op:
    MP.append(("Max premium for 20% IRR, strip, operator early rates", 0.20, dict(o=oil * op[0], g=gas * op[1])))
if T10 is not None:
    MP.insert(1, ("Max premium for 20% IRR, strip, 10-mile curve" + ("s (all zones)" if ZOTHER else ""), 0.20, dict(o=T10[0], g=T10[1], zones=Z_10)))
MP = MP[:10]

wb = openpyxl.load_workbook(src)
# TypeCurve
tc = wb["TypeCurve"]
for i in range(360):
    tc.cell(6 + i, 3).value, tc.cell(6 + i, 4).value = float(oil[i]), float(gas[i])
    tc.cell(6 + i, 5).value, tc.cell(6 + i, 6).value = float(eo[i]), float(eg[i])
ro, rg = eo.sum() / oil.sum(), eg.sum() / gas.sum()
pk = o.get("peak_month", 1) - 1; r1 = o.get("ramp_start", 1.0)
pg = g.get("peak_month", 1) - 1; rg1 = g.get("ramp_start", 1.0)
blk = {6: [S["wells"]] * 4, 7: [S["months_in_mean"]] * 4,
       8: [o["qi"], g["qi"], o["qi"] * ro, g["qi"] * rg],
       9: [o["De"], g["De"], o["De"], g["De"]],
       10: [o["b"], g["b"], o["b"], g["b"]], 11: [0.07] * 4,
       12: [float(oil.sum()), float(gas.sum()), float(eo.sum()), float(eg.sum())],
       14: [S["enverus"]["oil_median"], S["enverus"]["gas_median"]] * 2,
       17: [pk + 1, pg + 1] * 2, 18: [r1, rg1] * 2}
for r, v in blk.items():
    for j, x in enumerate(v):
        tc.cell(r, 9 + j).value = x
import re as _re
_note = _re.sub(r" ?Multi-phase: .*?\)\.", "", tc["B3"].value or "")
def _bdesc(lab, d):
    w = list(d["b_windows"].values()); bp = d["b_boot_p10_p90"]
    return (f"{lab} b {d['b']:.2f} ({'at the 2.0 cap, ' if d.get('b_pinned') else ''}{min(w):.2f}-{max(w):.2f} across windows; "
            f"bootstrap P10-P90 {bp[0]:.2f}-{bp[1]:.2f})")
_fit = (f"{_bdesc('Oil', o)}; {_bdesc('gas', g)}; both fitted on the {S['long_wells']} wells with 24+ months. "
        f"Guardrail {'binding' if min(S['guardrail'].values()) < 1 else 'not binding'}: base is "
        f"{S['fit_vs_median']['oil']:.0%} (oil) / {S['fit_vs_median']['gas']:.0%} (gas) of the Enverus fit-set median. "
        f"Partial first months scaled to a full month from Enverus producing days for {len(S.get('days_used', []))} of the "
        f"{S['wells']} fit wells (others: a first month under half of month 2 is dropped). ")
if "Oil b " in _note and "Prior 10-mile" in _note:
    _note = _note[:_note.index("Oil b ")] + _fit + _note[_note.index("Prior 10-mile"):]
elif "Oil b " not in _note:
    _note = _note + " " + _fit
_note = _note.replace("Map at N2, plots at N40.", "Comp map and DSU detail pictures below the charts (N84 down).")
tc["B3"] = (_note + f" Multi-phase: oil and gas each fitted on their own production (no GOR): straight-line ramp to the peak"
             f" month (oil {pk + 1}, gas {pg + 1}; 1 = no ramp), then hyperbolic decline from the peak. Ramp start, qi, De fitted"
             f" together on all months; b from the 24+ month wells (oil {o['b']:.2f}, gas {g['b']:.2f}"
             f"{', gas b at the 2.0 cap' if g.get('b_pinned') else ''}).").strip()
c5 = wb["Comps 5mi"] if "Comps 5mi" in wb.sheetnames else None
if c5 is not None and str(c5["B9"].value or "").startswith("b used"):
    c5["C9"] = f"{o['b']:.2f} / {g['b']:.2f}"
    c5["E9"] = (f"Both streams multi-phase, b fitted on the {S['long_wells']} wells with 24+ months"
                f"{'; gas b at the 2.0 cap (data wants flatter)' if g.get('b_pinned') else ''}.")
# the old static type-curve picture (1400 x 500, oil + gas on the GOR method) is superseded by the plot_tc charts
tc._images = [im for im in tc._images if not (im.width == 1400 and im.height == 500)]
tc["H8"] = "qi at peak (per day)"
tc["H9"] = "First-year decline from peak (effective)"
tc["H17"] = "Peak month (months before it ramp up in a straight line; 1 = no ramp)"
tc["H18"] = "Ramp start: month 1 as % of peak (fitted)"
for j in range(4):
    tc.cell(17, 9 + j).number_format = "0"; tc.cell(18, 9 + j).number_format = "0%"
    for r in (17, 18):
        tc.cell(r, 9 + j).font = tc.cell(16, 9 + j).font.copy(); tc.cell(r, 9 + j).alignment = tc.cell(12, 9 + j).alignment.copy()
wb["Inputs"]["D13"].value = (f"Base = 1 (oil {oil.sum()/1000:.0f} MBO, gas {gas.sum()/1e6:.2f} Bcf per 10,000 ft; {S['wells']} fit wells"
                             f"{', LOW CONFIDENCE' if S['wells'] < 15 else ''}). 2 = Enverus P50 ({eo.sum()/1000:.0f} MBO, {eg.sum()/1e6:.2f} Bcf)")
if ZOTHER:
    wb["Inputs"]["D13"].value += (f". Multi-zone: {len(ZL)} zone curves on 'Zone Curves' (zone 1 = {Z1}); each well uses its zone "
                                  f"(Wells column X); the switch applies to zone 1 only")
    if "Zone Curves" not in tc["B3"].value:
        tc["B3"] = tc["B3"].value + f" Zone 1 of {len(ZL)} ({Z1}); the other zones are on 'Zone Curves'."
if "Zone Curves" in wb.sheetnames:
    zc = wb["Zone Curves"]
    for k in range(6):
        zc.cell(4, 3 + 2 * k).value = None
        for r in range(6, 366):
            zc.cell(r, 3 + 2 * k).value = zc.cell(r, 4 + 2 * k).value = None
        for j in range(1, 13):
            zc.cell(5 + k, 16 + j).value = None
    for k, z in enumerate(ZL if len(ZL) > 1 else []):
        zo, zg = (oil, gas) if k == 0 else ZC[z["name"]]
        zc.cell(4, 3 + 2 * k).value = z["name"]
        for i in range(360):
            zc.cell(6 + i, 3 + 2 * k).value, zc.cell(6 + i, 4 + 2 * k).value = float(zo[i]), float(zg[i])
            zc.cell(6 + i, 3 + 2 * k).number_format = zc.cell(6 + i, 4 + 2 * k).number_format = "#,##0"
        t = z.get("ten_mile")
        flags = [f"{s_} b at the 2.0 cap" for s_ in ("oil", "gas") if z[s_].get("b_pinned")]
        vals = [z.get("radius", "5 mi"), z["wells"], z["long_wells"], "Yes" if z.get("low_confidence") else "No",
                round(z["oil"]["b"], 2), round(z["gas"]["b"], 2), round(float(zo.sum()) / 1000), round(float(zg.sum()) / 1e6, 2),
                t["wells"] if t else None, round(t["oil"]["eur"] / 1000) if t else None, round(t["gas"]["eur"] / 1e6, 2) if t else None,
                "; ".join([x for x in [z.get("note", "")] + flags if x])]
        for j, v in enumerate(vals):
            zc.cell(5 + k, 17 + j).value = v
    if len(ZL) > 1:
        plot_tc.add_zone_charts(wb, deal, ZL, ZC, Z10, (wb["Summary"]["B2"].value or "Deal").split(":")[0])
    else:
        zc._images = []
# ---- Lateral Curves: s / Tsw / g per well slot (the workbook does the rest with formulas)
LAT_INFO = []
if "Lateral Curves" in wb.sheetnames:
    lc, Wf = wb["Lateral Curves"], wb["Wells"]
    gl = M["glob"]
    for j in range(10):
        wr = 7 + j
        lat = Wf[f"E{wr}"].value
        lat = float(lat) if isinstance(lat, (int, float)) else 0.0
        zn = str(Wf[f"X{wr}"].value or "").strip()
        co, cg = (oil, gas) if (not zn or zn == Z1) else ZC.get(zn, (oil, gas))
        for k, c in ((2, co), (3, cg)):
            col = 3 + 4 * j + k
            for r in (11, 12, 13):
                lc.cell(r, col).value = None
            if lat <= 0:
                continue
            _, inf = LAT.scale_curve(c, LAT.BASIS, lat, gl["lat_band"], gl["lat_months"], gl["lat_pct"], gl["lat_full"])
            if inf["mode"] == "shaped":
                lc.cell(11, col).value, lc.cell(12, col).value, lc.cell(13, col).value = inf["s"], inf["tsw"], inf["g"]
            if k == 2 and Wf[f"C{wr}"].value:
                LAT_INFO.append(dict(well=Wf[f"C{wr}"].value, lat_ft=lat, mode=inf["mode"], early_x_linear=inf["e"], oil_s=inf["s"]))
    _lt = S.get("lateral") or {}
    _sh = [w for w in LAT_INFO if w["mode"] == "shaped"]
    _txt = (f" Lateral rule: offsets normalized to this 10,000 ft basis "
            f"({_lt.get('offsets_linear', S['wells'])} linear within +/- {gl['lat_band']:,.0f} ft, {len(_lt.get('offsets_shaped', []))} shaped); "
            f"each well's curve at its lateral is on 'Lateral Curves' "
            f"({len(_sh)} of {len(LAT_INFO)} wells shaped: months 1-{gl['lat_months']} at 100% of linear {gl['lat_band']:,.0f} ft from 10,000 ft "
            f"phasing to {gl['lat_pct']:.0%} at {gl['lat_full']:,.0f} ft, decline re-solved to the linear EUR).")
    tc["B3"] = _re.sub(r" Lateral rule: .*?\)\.", "", tc["B3"].value or "") + _txt
# TC vs offsets charts (TypeCurve N2 / N42), monthly table (TypeCurve H18) and the Offset Production tab
plot_tc.add_to_workbook(wb, deal, oil, gas, S["fit_set"], (wb["Summary"]["B2"].value or "Deal").split(":")[0], S["wells"] < 15,
                        ten_mile=T10)
# Minerals: the grid is priced as a multiple of the asking price (Inputs C20) instead of a premium on capex
MINERALS = M["glob"].get("minerals", False)
if MINERALS:
    ask = M["glob"]["price"]
    PM = {0: 0.8, 0.1: 0.9, 0.2: 1.0, 0.3: 1.2}
    def _m(name, kw):
        kw = dict(kw); pm = PM.get(kw.pop("premium", 0.2), 1.0); kw["price_mult"] = pm
        lab = "at asking price" if pm == 1.0 else f"price {pm:.0%} of ask"
        for k in ("premium 0%", "premium 10%", "premium 20%", "premium 30%"):
            name = name.replace(k, lab)
        return name, kw
    rows = [_m(n, k) for n, k in rows if "overrun" not in k]          # no capex in a minerals deal
    MP = [(n.replace("Max premium", "Max purchase price ($)"), t, k) for n, t, k in MP]
    wb["Summary"]["B55"] = f"What it takes (max purchase price, $; asking ${ask:,.0f})"
else:
    wb["Summary"]["B55"] = "What it takes (max premium, % of net capex)"
# grid
sm = wb["Summary"]
for i in range(30):
    for c in range(2, 10):
        sm.cell(25 + i, c).value = None
f = lambda x: "n/a" if x is None else x
base_cash = None
for i, (name, kw) in enumerate(rows):
    r = R(**dict(kw))
    vals = [name, f(r["irr"]), r["moic"], r["npv10"], r["pv20"], f(r["payout"]), f(r["min_well_irr"]), f(r["max_well_irr"])]
    for j, v in enumerate(vals):
        sm.cell(25 + i, 2 + j).value = v
    if name == "Strip, premium 20%" or (MINERALS and name == "Strip, at asking price"):
        base_cash = r["cash"]
for i in range(10):
    for c in range(2, 6):
        sm.cell(56 + i, c).value = None
for i, (name, tgt, kw) in enumerate(MP):
    kw = dict(kw); oo, gg = kw.pop("o", oil), kw.pop("g", gas)
    if MINERALS:
        pkg = E.max_price(M, oo, gg, target=tgt, **kw)
        per = [E.max_price(M, oo, gg, target=tgt, wells=[dict(w, share=1.0)], **kw) for w in M["wells"]]
    else:
        pkg = E.max_premium(M, oo, gg, target=tgt, **kw)
        per = [E.max_premium(M, oo, gg, target=tgt, wells=[w], **kw) for w in M["wells"]]
    sm.cell(56 + i, 2).value, sm.cell(56 + i, 3).value = name, pkg
    sm.cell(56 + i, 4).value, sm.cell(56 + i, 5).value = min(per), max(per)
    for c in (3, 4, 5):
        if MINERALS:
            sm.cell(56 + i, c).number_format = '$#,##0'
        elif sm.cell(56 + i, c).number_format == '$#,##0':
            sm.cell(56 + i, c).number_format = '0.0%'
# ---- live case: solve Summary C7 (premium, % of net capex) so the base-case IRR (Summary C12) is 20%.
# Solved at the live deck (C6) and type-curve switch (Inputs C13) with the live Inputs (TC multiplier, haircut, overrun).
# The solved value is also kept in E7 so the engine tie (D75) knows C7 has not been typed over. Minerals: C7 is not used.
P_STAR = None
if not MINERALS:
    _deck = sm["C6"].value if (sm["C6"].value == "Strip" or sm["C6"].value in E.FLAT) else "Strip"
    _co, _cg = (eo, eg) if wb["Inputs"]["C13"].value == 2 else (oil, gas)
    P_STAR = E.max_premium(M, _co, _cg, target=0.20, deck=_deck)
    _pinned = P_STAR <= E.PREM_LO + 1e-6 or P_STAR >= E.PREM_HI - 1e-6
    sm["C7"].value, sm["E7"].value = P_STAR, P_STAR
    sm["E7"].number_format = sm["C7"].number_format = "0.0%"
    sm["E7"].font = _copy(sm["C10"].font)
    sm["B7"].value = "Premium, % of net capex (solved for a 20% IRR in C12; bid input)"
    sm["F7"].value = ("Solved by write_model.py so the live-case IRR (C12) = 20% (solve kept in E7). Type over C7 to test a bid. "
                      + ("20% IRR NOT REACHABLE within the solve range; C7 is at the limit. " if _pinned else "")
                      + ("Negative = misses 20% even paying capex only (needs a carry / discount). " if P_STAR < 0 and not _pinned else "")
                      + "Not used for Minerals (price on Inputs C20)")
    base_cash = R(premium=P_STAR)["cash"]
    sm["B75"].value = ("Package undiscounted cash ties to the Python engine (strip; the solved premium in C7, or the purchase price for "
                       "Minerals; base inputs); engine value in E75")
    sm["D75"].value = ('=IF(OR(C6<>"Strip",AND(NOT(Inputs!$C$18="Minerals"),ABS(C7-$E$7)>0.000001),Inputs!$C$13<>1,'
                       'Inputs!$C$11<>1,Inputs!$C$59<>0),"OK",IF(ABS(C75-$E$75)<500,"OK","CHECK"))')
sm["E75"].value = base_cash
wb.save(out)

# ---- deal record for the deal log and later lookbacks (<deal_dir>/deal.json); the decision fields are kept on rebuilds
import datetime as _dt
rec_path = os.path.join(deal, "deal.json")
old = json.load(open(rec_path)) if os.path.exists(rec_path) else {}
base_kw = {"price_mult": 1.0} if MINERALS else {"premium": 0.2}
rb = R(**base_kw)
mp0 = (E.max_price if MINERALS else E.max_premium)(M, oil, gas, target=0.20)
g0 = M["glob"]
rec = {
    "deal": (wb["Summary"]["B2"].value or "Deal").split(":")[0].strip(),
    "model_file": os.path.basename(out), "evaluated": _dt.date.today().isoformat(),
    "type": "Minerals" if MINERALS else "Working interest", "basin": g0.get("basin"),
    "wells": [{"name": w["name"], "api": w.get("api", ""), "lat_ft": w["lat"], "wi": w["wi"], "lease_nri": w["lease_nri"], "mineral_nri": w.get("mnri", 0),
               "first_sales_month_index": w["first_m"], "haircut": w["haircut"], "tc_factor": w["tc_factor"],
               "zone": w.get("zone") or Z1} for w in M["wells"]],
    "valuation_month": str(wb["Inputs"]["C37"].value)[:10],
    "lateral_rule": {"band_ft": g0["lat_band"], "early_months": g0["lat_months"], "early_pct": g0["lat_pct"], "full_ft": g0["lat_full"],
                     "model_lateral": (S.get("lateral") or {}).get("model_lateral"),
                     "offsets_shaped": len((S.get("lateral") or {}).get("offsets_shaped", [])), "wells": LAT_INFO},
    "net_capex": 0 if MINERALS else sum(w["wi"] * (w["drill"] + w["compl"] + w["fac"]) for w in M["wells"]),
    "asking_price": g0.get("price") if MINERALS else None,
    "base_case": "strip, asking price" if MINERALS else "strip, 20% premium",
    "irr": rb["irr"], "moic": rb["moic"], "npv10": rb["npv10"], "payout_months": rb["payout"],
    "max_for_20pct_irr": mp0, "summary_c7_solved_premium": P_STAR, "max_for_20pct_irr_units": "$" if MINERALS else "premium, % of net capex",
    "type_curve": {"fit_wells": S["wells"], "low_confidence": S["wells"] < 15 or any(z.get("low_confidence") for z in ZL),
                   "tc_multiplier": g0["tc_mult"],
                   "zones": [{"name": z["name"], "radius": z.get("radius"), "fit_wells": z["wells"], "low_confidence": z.get("low_confidence"),
                              "oil_eur_per_10k": float(ZC[z["name"]][0].sum()), "gas_eur_per_10k": float(ZC[z["name"]][1].sum()),
                              "oil_b": z["oil"]["b"], "gas_b": z["gas"]["b"],
                              "wells": sum(1 for w in M["wells"] if (w.get("zone") or Z1) == z["name"])} for z in ZL] if len(ZL) > 1 else [],
                   "oil_eur_per_10k": float(oil.sum()), "gas_eur_per_10k": float(gas.sum()),
                   "oil_b": o["b"], "gas_b": g["b"], "ten_mile_wells": S.get("ten_mile", {}).get("wells")},
    "decision": old.get("decision", "Pending"), "price_paid": old.get("price_paid"), "closed": old.get("closed"),
    "notes": old.get("notes", ""),
}
json.dump(rec, open(rec_path, "w"), indent=1, default=str)
print("wrote", out, "| mode", "Minerals" if MINERALS else "WI", "| C7 solved premium", P_STAR, "| IRR at 20% premium",
      R(**({"price_mult": 1.0} if MINERALS else {"premium": 0.2}))["irr"], "| cash", base_cash)
