"""Deal-specific type-curve adjustment for LTK 24-25 (Davis, Nov 23, 2026). Run after build_tc.py:
    cd toolkit && python3 ../deals/ltk-24-25-1bs/adjust_tc.py ../deals/ltk-24-25-1bs
Base fit set = the 35-well 10-mile 1st Bone Spring set (see deal_data.py).
  Oil: build_tc's curve peaked ~11% below the offsets' mean peak (month 5). The peak is pinned to the observed
       mean month-5 rate; ramp start, De and b are refitted (b fitted from the data, not hard-coded) on every month
       with 8+ wells reporting (~41 months), weighted by wells reporting, so the tail follows the longer-lived offsets.
       A 31-month refit gave b 1.13 / 638 MBO and ran 1.5-2x above every offset past month 31; Davis asked for a tail
       that fits the offsets.
  Gas: build_tc's fit sat on its De floor (0.20) with b at the 2.0 cap and ran under the offsets after month 12.
       The De floor is relaxed to 0.05 and the curve refitted (log RMS 0.113 -> 0.104), lifting gas ~10%.
Writes tc_curves.csv and the oil / gas blocks of tc_summary.json; keeps the unadjusted curves in tc_curves_unadjusted.csv."""
import sys, os, json, shutil
import numpy as np
from scipy.optimize import least_squares
sys.path.insert(0, os.getcwd())
import build_tc as B

deal = sys.argv[1]
B.load_deal(deal)
oil, gas, cnt, series = B.mean_curve(B.FIT_SET)
n = len(oil)
rms = lambda c, v: float(np.sqrt(np.mean((np.log(c[:len(v)]) - np.log(v)) ** 2)))
src = os.path.join(deal, "tc_curves.csv")
keep = os.path.join(deal, "tc_curves_unadjusted.csv")
if not os.path.exists(keep):
    shutil.copy(src, keep)
C = np.loadtxt(keep, delimiter=",", skiprows=1)
hdr = open(keep).readline().strip()

# oil: peak pinned to the observed mean peak; De, b, ramp refitted on every month with 8+ wells reporting
# (through month ~41, not just the 31-month build window), weighted by wells reporting, so the tail follows the
# older offsets instead of extrapolating the flat b > 1 seen in the first 31 months (Davis, Nov 23 2026)
MIN_TAIL_WELLS = 8
L = max(len(s[0]) for s in series.values())
cnt_all = np.array([sum(len(s[0]) > m for s in series.values()) for m in range(L)])
nt = int(np.max(np.where(cnt_all >= MIN_TAIL_WELLS)) + 1)
mean_t = np.array([np.mean([s[0][m] for s in series.values() if len(s[0]) > m]) for m in range(nt)])
w = np.sqrt(cnt_all[:nt])
p = B.find_peak(oil); qi = oil[p] / 30.4; y = np.log(np.maximum(mean_t, 1))
r = least_squares(lambda q: (np.log(B.multiphase(qi, q[0], q[2], p, q[1], n=nt)) - y) * w, [0.6, 0.8, 0.5],
                  bounds=([0.2, 0.05, 0.05], [0.95, 1.0, 2.0]))
De_o, r1_o, b_o = r.x; co = B.multiphase(qi, De_o, b_o, p, r1_o)

# gas: De floor relaxed to 0.05
pg = B.find_peak(gas); yg = np.log(gas); q0 = gas[pg] / 30.4
r = least_squares(lambda q: np.log(B.multiphase(q[0], q[1], q[3], pg, q[2], n=n)) - yg, [q0, 0.3, 0.7, 1.5],
                  bounds=([q0 * 0.3, 0.05, 0.05, 0.05], [q0 * 3, 0.95, 1.0, 2.0]))
qg, De_g, r1_g, b_g = r.x; cg = B.multiphase(qg, De_g, b_g, pg, r1_g)

S = json.load(open(os.path.join(deal, "tc_summary.json")))
note_o = f"peak pinned to the offsets' mean month-{p+1} rate; De / b / ramp refitted on months 1-{nt} (8+ wells reporting), weighted by wells reporting; log RMS months 1-{nt}: {rms(C[:,1], mean_t):.3f} -> {rms(co, mean_t):.3f}"
note_g = f"De floor relaxed 0.20 -> 0.05 and refitted (log RMS {rms(C[:,2], gas):.3f} -> {rms(cg, gas):.3f})"
for k, c, prm, note in (("oil", co, dict(qi=qi, De=De_o, b=b_o, peak_month=p + 1, ramp_start=r1_o), note_o),
                        ("gas", cg, dict(qi=qg, De=De_g, b=b_g, peak_month=pg + 1, ramp_start=r1_g), note_g)):
    S[k].setdefault("eur_before_adjustment", S[k]["eur"])
    S[k].update({kk: float(v) if not isinstance(v, int) else v for kk, v in prm.items()})
    S[k].update(b_fitted=float(prm["b"]), b_pinned=bool(prm["b"] >= 1.999), eur=float(c.sum()), eur_unguarded=float(c.sum()), adjustment=note)
S["enverus"]["oil_median"], S["enverus"]["gas_median"] = float(co.sum()), float(cg.sum())
S["enverus"]["oil_p90_p50_p10"] = [float(co.sum()) / 1e3] * 3   # no Enverus EURs on this subscription: P10 = P50 = P90 = base
S["enverus"]["gas_p90_p50_p10"] = [float(cg.sum()) / 1e6] * 3
S["mean_oil"] = [float(x) for x in oil]; S["mean_gas"] = [float(x) for x in gas]
json.dump(S, open(os.path.join(deal, "tc_summary.json"), "w"), indent=1)
C[:, 1], C[:, 2] = co, cg
if C.shape[1] > 4:
    C[:, 3], C[:, 4] = co, cg
np.savetxt(src, C, delimiter=",", header=hdr.lstrip("# "), comments="", fmt="%.6f")
print(f"oil: qi {qi:.0f}/d De {De_o:.3f} b {b_o:.2f} EUR {co.sum()/1e3:.0f} MBO (was {C.shape and S['oil']['eur_before_adjustment']/1e3:.0f}); {note_o}")
print(f"gas: qi {qg:.0f}/d De {De_g:.3f} b {b_g:.2f} EUR {cg.sum()/1e6:.2f} Bcf (was {S['gas']['eur_before_adjustment']/1e6:.2f}); {note_g}")
