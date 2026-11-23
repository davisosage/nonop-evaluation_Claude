"""Deal-specific oil type-curve adjustment for Sun 5899 11-28 2BHP / 3BHP (Davis, Nov 23, 2026). Run after build_tc.py:
    cd toolkit && python3 ../deals/sun5899-2bhp-3bhp/adjust_tc.py ../deals/sun5899-2bhp-3bhp
Base fit set = the 8-well 5-mile Middle Bakken set (see deal_data.py).
  Oil: Davis asked to bring the peak month down a little and flatten the decline, with b fixed at 0.9 (his call for this
       deal; overrides the fitted b of 0.96). The peak is pinned to the offsets' observed mean month-3 rate (build_tc's
       curve sat ~11% above it), b is held at 0.9, and De and the ramp start are refitted on every month with 4+ wells
       reporting (months 1-39), weighted by wells reporting.
  Gas: unchanged (Davis: gas looks good).
  10-mile comparison curve: unchanged (its own fitted b).
Writes tc_curves.csv and the oil block of tc_summary.json; keeps the unadjusted curves in tc_curves_unadjusted.csv."""
import sys, os, json, shutil
import numpy as np
from scipy.optimize import least_squares
sys.path.insert(0, os.getcwd())
import build_tc as B

B_OIL = 0.9          # Davis, Nov 23 2026
MIN_WELLS = 4

deal = sys.argv[1]
B.load_deal(deal)
oil, gas, cnt, series = B.mean_curve(B.FIT_SET)
n = len(oil)
src = os.path.join(deal, "tc_curves.csv")
keep = os.path.join(deal, "tc_curves_unadjusted.csv")
if not os.path.exists(keep):
    shutil.copy(src, keep)
C = np.loadtxt(keep, delimiter=",", skiprows=1)
hdr = open(keep).readline().strip()

L = max(len(s[0]) for s in series.values())
cnt_all = np.array([sum(len(s[0]) > m for s in series.values()) for m in range(L)])
nt = int(np.max(np.where(cnt_all >= MIN_WELLS)) + 1)
mean_t = np.array([np.mean([s[0][m] for s in series.values() if len(s[0]) > m]) for m in range(nt)])
w = np.sqrt(cnt_all[:nt]); y = np.log(np.maximum(mean_t, 1))
rms = lambda c: float(np.sqrt(np.mean((np.log(c[:nt]) - y) ** 2)))

p = B.find_peak(mean_t); qi = mean_t[p] / 30.4
r = least_squares(lambda q: (np.log(B.multiphase(qi, q[0], B_OIL, p, q[1], n=nt)) - y) * w, [0.6, 0.6],
                  bounds=([0.05, 0.05], [0.95, 1.0]))
De_o, r1_o = r.x
co = B.multiphase(qi, De_o, B_OIL, p, r1_o, n=C.shape[0])

S = json.load(open(os.path.join(deal, "tc_summary.json")))
note = (f"Davis: peak down, flatter decline, b fixed {B_OIL}. Peak pinned to the offsets' mean month-{p+1} rate; De and ramp "
        f"refitted on months 1-{nt} ({MIN_WELLS}+ wells), weighted by wells reporting; log RMS {rms(C[:,1]):.3f} -> {rms(co):.3f}")
o = S["oil"]
o.setdefault("eur_before_adjustment", o["eur"]); o.setdefault("b_fitted_before_adjustment", o["b"])
o.update(qi=float(qi), De=float(De_o), b=B_OIL, b_fitted=B_OIL, b_pinned=False, peak_month=p + 1, ramp_start=float(r1_o),
         b_fixed_sensitivity=B_OIL, eur=float(co.sum()), eur_unguarded=float(co.sum()), adjustment=note)
json.dump(S, open(os.path.join(deal, "tc_summary.json"), "w"), indent=1)
ratio = co / C[:, 1]
C[:, 1] = co
if C.shape[1] > 3:
    C[:, 3] = C[:, 3] * ratio     # Enverus P50 oil column keeps its ratio to the base shape
np.savetxt(src, C, delimiter=",", header=hdr.lstrip("# "), comments="", fmt="%.6f")
print(f"oil: qi {qi:.0f}/d De {De_o:.3f} b {B_OIL} ramp {r1_o:.2f} EUR {co.sum()/1e3:.1f} MBO/10k ft "
      f"(was {o['eur_before_adjustment']/1e3:.1f}, b {o['b_fitted_before_adjustment']:.2f}); {note}")
