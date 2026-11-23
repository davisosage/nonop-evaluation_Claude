"""Lookback: actual production of a deal's wells against the type curve the deal was priced on.

Usage: python lookback.py <deal_dir>
Needs in <deal_dir>:
  * deal.json        written by write_model.py when the deal was evaluated (wells, laterals, haircuts, valuation month)
  * tc_curves.csv    the type curve the deal was priced on (month, base_oil, base_gas, ...); multi-zone deals also
                     tc_zones.json / tc_zone_curves.csv, and each well is compared with its own zone's curve
  * actuals.py       ACTUALS = {api: (first_month "YYYY-MM", [oil bbl], [gas Mcf], [producing days] or None)}
                     pulled from Enverus (production dataset) for the deal's own wells
Writes <deal_dir>/lookback.json and lookback.png. deal_log.py reads lookback.json.

Comparison: each well's prediction is the type curve x lateral / 10,000 x (1 - haircut) x TC multiplier x well TC factor,
started in the well's actual first producing month (so the shape is compared on equal footing); the timing gap
(actual first production vs the model's first-sales month) is reported separately. A partial first month is
scaled to a full month when producing days are given, the same way the type curve was built.
"""
import json, os, sys, importlib.util, datetime as dt
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import build_tc as bt
import lateral as LAT


def month_add(ym, k):
    y, m = int(ym[:4]), int(ym[5:7]) + k
    return f"{y + (m - 1) // 12:04d}-{(m - 1) % 12 + 1:02d}"


def main(deal_dir):
    rec = json.load(open(os.path.join(deal_dir, "deal.json")))
    C = np.loadtxt(os.path.join(deal_dir, "tc_curves.csv"), delimiter=",", skiprows=1)
    ZC = {}
    if os.path.exists(os.path.join(deal_dir, "tc_zones.json")):
        A = np.loadtxt(os.path.join(deal_dir, "tc_zone_curves.csv"), delimiter=",", skiprows=1, ndmin=2)
        for k, z in enumerate(json.load(open(os.path.join(deal_dir, "tc_zones.json")))["zones"]):
            ZC[z["name"]] = A[:, [0, 1 + 2 * k, 2 + 2 * k]]
    spec = importlib.util.spec_from_file_location("actuals", os.path.join(deal_dir, "actuals.py"))
    A = importlib.util.module_from_spec(spec); spec.loader.exec_module(A)
    mult = rec["type_curve"].get("tc_multiplier", 1.0) or 1.0
    v0 = rec["valuation_month"][:7]
    wells, tot = [], dict(act_o=0.0, pred_o=0.0, act_g=0.0, pred_g=0.0)
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.8), dpi=100)
    for w in rec["wells"]:
        api = str(w.get("api", "")).replace("-", "")
        if api not in A.ACTUALS:
            wells.append(dict(name=w["name"], api=api, status="no production reported yet")); continue
        first, o, g, days = (list(A.ACTUALS[api]) + [None])[:4]
        o, g = np.array(o, float), np.array(g, float)
        if days and days[0] and days[0] < bt._cal_days(first, 0):       # partial first month -> full month
            f = bt._cal_days(first, 0) / max(days[0], 1); o[0] *= f; g[0] *= f
        n = len(o)
        k = (1 - w["haircut"]) * mult * w["tc_factor"]
        Cw = ZC.get(w.get("zone"), C)
        lr = rec.get("lateral_rule") or {"band_ft": 1e9}      # deals modeled before the rule (Nov 2026) were linear
        sc = lambda c: LAT.scale_curve(c, LAT.BASIS, w["lat_ft"], lr.get("band_ft"), lr.get("early_months"), lr.get("early_pct"), lr.get("full_ft", lr.get("band_ft")))[0]   # no full_ft = built on the earlier step rule
        po, pg = sc(Cw[:, 1])[:n] * k, sc(Cw[:, 2])[:n] * k
        model_first = month_add(v0, w["first_sales_month_index"])
        gap = (int(first[:4]) - int(model_first[:4])) * 12 + int(first[5:7]) - int(model_first[5:7])
        r = dict(name=w["name"], api=api, zone=w.get("zone"), months=n, actual_first=first, model_first_sales=model_first, timing_gap_months=gap,
                 cum_oil=float(o.sum()), pred_oil=float(po.sum()), oil_vs_curve=float(o.sum() / po.sum() - 1),
                 cum_gas=float(g.sum()), pred_gas=float(pg.sum()), gas_vs_curve=float(g.sum() / pg.sum() - 1) if pg.sum() else None)
        wells.append(r)
        tot["act_o"] += r["cum_oil"]; tot["pred_o"] += r["pred_oil"]; tot["act_g"] += r["cum_gas"]; tot["pred_g"] += r["pred_gas"]
        x = np.arange(1, n + 1)
        for ax, a, p_ in ((axes[0], o, po), (axes[1], g, pg)):
            l, = ax.plot(x, np.cumsum(a) / 1000, lw=2, label=f"{w['name']} actual")
            ax.plot(x, np.cumsum(p_) / 1000, lw=1.4, ls="--", color=l.get_color(), label=f"{w['name']} type curve")
    for ax, t, u in ((axes[0], "Cumulative oil", "MBO"), (axes[1], "Cumulative gas", "MMcf")):
        ax.set_title(f"{rec['deal']}: {t}, actual vs type curve", loc="left", fontsize=11)
        ax.set_xlabel("Months on production"); ax.set_ylabel(u); ax.grid(alpha=0.3); ax.legend(fontsize=7, frameon=False)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
    fig.tight_layout(); fig.savefig(os.path.join(deal_dir, "lookback.png")); plt.close(fig)
    out = dict(deal=rec["deal"], as_of=dt.date.today().isoformat(), wells=wells,
               package=dict(cum_oil=tot["act_o"], pred_oil=tot["pred_o"],
                            oil_vs_curve=tot["act_o"] / tot["pred_o"] - 1 if tot["pred_o"] else None,
                            cum_gas=tot["act_g"], pred_gas=tot["pred_g"],
                            gas_vs_curve=tot["act_g"] / tot["pred_g"] - 1 if tot["pred_g"] else None))
    json.dump(out, open(os.path.join(deal_dir, "lookback.json"), "w"), indent=1)
    print(json.dumps(out["package"], indent=1))
    for r in wells:
        print(r)


if __name__ == "__main__":
    main(sys.argv[1])
