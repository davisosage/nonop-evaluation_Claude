"""Per-zone TC factors for the Loki package (run from toolkit/: python3 ../deals/loki-upton/zone_factors.py).
The template carries one type curve (the 5-mile Wolfcamp B). Each zone gets a multi-phase curve fitted on its own offsets
mean (build_tc.fit_tc: ramp, qi, De fitted per zone), with b held at the values fitted on the 97-well 10-mile Wolfcamp B
set (its own 90 wells with 24+ months), because the zone sets are too thin for a stable b (13-21 wells; their own fits are
reported too). Wells-sheet TC factor = zone value / 10-mile Wolfcamp B value, where value per
10,000 ft = oil EUR x $64.50 + wellhead gas EUR x $3.59 (flat $65 / $3.50 at the Permian basin defaults: residue at HH-$1,
20% shrink, 120 bbl/MMcf NGL at 30% of WTI, $0.75 G&P). Writes zone_factors.json."""
import json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "toolkit"))
import build_tc as bt
bt.load_deal(HERE)
import importlib.util
spec = importlib.util.spec_from_file_location("dd", os.path.join(HERE, "deal_data.py")); dd = importlib.util.module_from_spec(spec); spec.loader.exec_module(dd)
OIL_V, GAS_V = 64.50, 3.59
S = json.load(open(os.path.join(HERE, "tc_summary.json")))
OB, GB = S["ten_mile"]["oil"]["b"], S["ten_mile"]["gas"]["b"]
out = {}
for z, apis in dd.ZONE_SETS.items():
    mo, mg, cnt, _ = bt.mean_curve(apis)
    oc, PO = bt.fit_tc(mo, b=OB); gc, PG = bt.fit_tc(mg, b=GB)
    own = bt.build_set(apis)[4]
    n = min(24, len(mo))
    out[z] = dict(wells=len(apis), months_in_mean=len(mo), oil_eur=float(oc.sum()), gas_eur=float(gc.sum()), oil_b=OB, gas_b=GB,
                  own_fit=dict(oil_eur=own["oil"]["eur"], gas_eur=own["gas"]["eur"], oil_b=own["oil"]["b"], gas_b=own["gas"]["b"]),
                  cum24_oil=float(mo[:n].sum()), cum24_gas=float(mg[:n].sum()),
                  value=float(oc.sum() * OIL_V + gc.sum() * GAS_V), value24=float(mo[:n].sum() * OIL_V + mg[:n].sum() * GAS_V))
ref, ref24 = out["WCB"]["value"], out["WCB"]["value24"]
for z in out:
    out[z]["factor"] = out[z]["value"] / ref
    out[z]["factor_24mo_check"] = out[z]["value24"] / ref24
    print(f"{z:5s} {out[z]['wells']:3d} wells  oil {out[z]['oil_eur']/1000:5.0f} MBO  gas {out[z]['gas_eur']/1e6:5.2f} Bcf /10k ft  "
          f"factor {out[z]['factor']:.3f}  (24-mo cum check {out[z]['factor_24mo_check']:.3f}; own-b fit {out[z]['own_fit']['oil_eur']/1000:.0f} MBO)")
json.dump(out, open(os.path.join(HERE, "zone_factors.json"), "w"), indent=1)
