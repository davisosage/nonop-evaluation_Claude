"""Fills the PDP Evaluation Template for a deal.

Usage: python write_pdp.py <deal_dir> "<deal_dir>/<Deal> <Basin> PDP Model.xlsx"
Needs <deal_dir>/pdp_data.py and pdp_fit.json (run pdp_fit.py first). Then recalculate the output (LibreOffice headless or
Excel) and run `python pdp_engine.py <output>` to confirm the engine ties; Summary C45 should read "All checks OK".

Writes: Inputs (deal name, operator, type, basin, effective date, asking price, any INPUTS overrides from pdp_data.py),
Wells (interests, the fitted decline per well, fit notes), History, Offset Data, Well Charts, the engine sensitivity grid
on Summary, and <deal_dir>/deal.json (decision fields kept if it already exists).

Optional in pdp_data.py:
    INPUTS = {"odiff": -3.0, "loe_fix": 9000, "anchor": "Actual", "tol": 0.2, ...}   keys as build_pdp_template.I; pricing / cost
             keys go to the override column E, setup keys to column C
    OVERRIDES = {api: {"o_b": 1.2, "g_da": 0.3, "o_q": 2500, "mult": 0.9}}           Wells override columns
"""
import datetime as dt, json, os, sys, shutil
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import build_pdp_template as T
import pdp_fit as F
import pdp_engine as E

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "PDP Evaluation Template.xlsx")
SETUP_KEYS = {"name", "operator", "dtype", "basin", "eff", "price", "tirr", "anchor", "mult", "life_yr", "dmin", "econ", "tol", "roy", "disc"}


def d(ym):
    return dt.datetime(int(ym[:4]), int(ym[5:7]), 1)


def notes(s):
    n = []
    if s["b_at_cap"]:
        n.append("b at the 2.0 cap")
    if s["b"] <= F.B_LO + 1e-3:
        n.append("b at the 0.01 floor (exponential): check vs offsets")
    if s.get("level_shift"):
        ls = s["level_shift"]; n.append(f"level shift {ls['change']:+.0%} from month {ls['from_age']}: confirm with the operator")
    if "offsets'" in s["b_source"]:
        n.append("b from offsets' shape (short history)")
    elif "LOW CONFIDENCE" in s["b_source"]:
        n.append("short history, low confidence")
    if s["r2"] < 0.8:
        n.append(f"loose fit (R2 {s['r2']:.2f})")
    return "; ".join(n)


def main(deal_dir, out):
    dd = F.load(deal_dir)
    fit = json.load(open(os.path.join(deal_dir, "pdp_fit.json")))
    DEAL = getattr(dd, "DEAL", {}); INP = getattr(dd, "INPUTS", {}); OVR = getattr(dd, "OVERRIDES", {}); LOE = getattr(dd, "LOE", {})
    if len(fit["wells"]) > T.NW:
        raise SystemExit(f"{len(fit['wells'])} wells: the template holds {T.NW}")
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Inputs"]
    setup = dict(name=DEAL.get("name", os.path.basename(deal_dir)), operator=DEAL.get("operator", ""), dtype=DEAL.get("deal_type", "Working interest"),
                 basin=DEAL.get("basin", "DJ (CO)"), eff=d(DEAL["effective"]), price=float(DEAL.get("asking_price", 0) or 0))
    for k, v in setup.items():
        ws[T.I[k]] = v
    for k, v in INP.items():
        if k in SETUP_KEYS:
            ws[T.I[k]] = v
        else:
            ws["E" + T.I[k][1:]] = v

    # --- Wells, History, Offset Data
    W, H, OD = wb["Wells"], wb["History"], wb["Offset Data"]
    for i, f in enumerate(fit["wells"]):
        r = T.W0 + i
        v = dd.SUBJECT[f["api"]]
        wi, nri, lnri = v[6], v[7], v[8] if len(v) > 8 else 0.8
        put = lambda k, x: W.cell(r, T.WC[k], x)
        put("name", f["name"]); put("api", f["api"]); put("zone", f["zone"]); put("lat", f["lateral"])
        put("first", d(f["first"])); put("last", d(f["last_month"])); put("wi", wi); put("nri", nri); put("lnri", lnri)
        if f["api"] in LOE:
            put("loe_ovr", LOE[f["api"]])
        for s, key in (("o", "oil"), ("g", "gas")):
            x = f[key]; an = x["anchor"]
            put(f"{s}_qfit", an["q_fit"]); put(f"{s}_qact", an["q_actual"]); put(f"{s}_da", an["Da"]); put(f"{s}_b", x["b"])
            put(f"{s}_r2", x["r2"]); put(f"{s}_note", notes(x))
            put(f"{s}_start", x["start_age"]); put(f"{s}_q0", x["q0"]); put(f"{s}_D", x["D"]); put(f"{s}_bf", x["b"])
            ls = x.get("level_shift")
            put(f"{s}_sh_age", ls["from_age"] if ls else 0); put(f"{s}_sh", 1 + ls["change"] if ls else 1)
            for a, val in enumerate(x["history"][:T.NH]):
                H.cell(T.H0 + a, 2 * (i + 1) + (1 if s == "o" else 2), val)
            ov = x["offsets"]; c0 = 3 + i * 10 + (0 if s == "o" else 5)
            for k in range(T.NK):
                OD.cell(T.K0 + k, c0, ov["n"][k])
                for q, nm in ((1, "p25"), (2, "p50"), (3, "p75")):
                    if ov[nm][k] is not None:
                        OD.cell(T.K0 + k, c0 + q, round(ov[nm][k], 2))
            oc = wb["Offset Check"]; C = {"o": 6, "g": 6 + 15}[s]
            oc.cell(r, C, ov["offsets_at_anchor"]); oc.cell(r, C + 1, ov["covered_months"]); oc.cell(r, C + 2, ov.get("level_vs_offsets"))
        for k, val in OVR.get(f["api"], {}).items():
            col = {"o_b": "o_b_ovr", "o_da": "o_da_ovr", "o_q": "o_q_ovr", "g_b": "g_b_ovr", "g_da": "g_da_ovr", "g_q": "g_q_ovr"}.get(k, k)
            put(col, val)
    wb["Well Review"]["C3"] = fit["wells"][0]["name"]
    T.add_review_charts(wb["Well Review"])

    # --- Well Charts
    wc = wb["Well Charts"]
    for i, f in enumerate(fit["wells"]):
        p = os.path.join(deal_dir, f"pdp_{f['api']}.png")
        if os.path.exists(p):
            im = XLImage(p); im.width, im.height = int(im.width * 0.72), int(im.height * 0.72)
            wc.add_image(im, f"B{5 + i * 25}")

    # --- engine (resolved inputs from the template's basin table and the deal)
    inp, wells, prices = engine_inputs(wb, dd, fit, setup, INP, OVR, LOE)
    base = E.run(inp, wells, prices, "Strip")
    S = wb["Summary"]
    S["E43"] = base["pv10"]
    for i, dk in enumerate(T.DECKS):
        for j, m in enumerate([0.85, 1.0, 1.15]):
            r = E.run(inp, wells, prices, dk, mult=m)
            S.cell(49 + i, 3 + j, round(r["max_price"])); S.cell(49 + i, 7 + j, round(r["pv10"]))
    for c, anc in (("C54", "Fit"), ("D54", "Actual")):
        S[c] = round(E.run(dict(inp, anchor=anc), anchor_wells(wells, fit, OVR, anc), prices, "Strip")["max_price"])
    S["C55"] = round(E.run(inp, wells, prices, "Strip", econ=0)["max_price"])
    wb.save(out)

    # --- deal.json
    dj = os.path.join(deal_dir, "deal.json")
    old = json.load(open(dj)) if os.path.exists(dj) else {}
    rec = dict(evaluated=dt.date.today().isoformat(), deal=setup["name"], type=f"PDP ({setup['dtype']})", basin=setup["basin"],
               operator=setup["operator"], effective=DEAL["effective"], model=os.path.basename(out),
               wells=[dict(api=f["api"], name=f["name"], zone=f["zone"]) for f in fit["wells"]],
               asking_price=setup["price"], irr=base["irr"], moic=base["moic"], pv10=base["pv10"], npv10=base["pv10"] - setup["price"],
               max_for_20pct_irr=base["max_price"] if abs(inp["tirr"] - 0.2) < 1e-9 else None, max_price_target_irr=base["max_price"],
               net_oil_bbl=base["net_oil"], net_gas_mcf=base["net_gas"],
               decision=old.get("decision", "Pending"), price_paid=old.get("price_paid"), closed=old.get("closed"), notes=old.get("notes", ""))
    json.dump(rec, open(dj, "w"), indent=1)
    irr = "n/a" if base["irr"] is None else f"{base['irr']:.1%}"
    print(f"wrote {out}\nstrip: max price for {inp['tirr']:.0%} IRR ${base['max_price']:,.0f}, PV10 ${base['pv10']:,.0f}, "
          f"IRR at ${setup['price']:,.0f} = {irr}")


def anchor_wells(wells, fit, OVR, anc):
    out = []
    for w, f in zip(wells, fit["wells"]):
        w = dict(w); o = OVR.get(f["api"], {})
        for s, key in (("o", "oil"), ("g", "gas")):
            if f"{s}_q" not in o:
                w[s + "_q"] = f[key]["anchor"]["q_actual" if anc == "Actual" else "q_fit"]
        out.append(w)
    return out


def engine_inputs(wb, dd, fit, setup, INP, OVR, LOE):
    ws, B = wb["Inputs"], wb["Basins"]
    basins = [B.cell(4, c).value for c in range(4, 9)]
    col = 4 + basins.index(setup["basin"])
    inp = {}
    for k, br in T.BASIN_ROW.items():
        inp[k] = float(INP[k]) if k in INP else float(B.cell(br, col).value)
    for k in ("tirr", "anchor", "mult", "life_yr", "dmin", "econ", "tol", "roy", "disc"):
        inp[k] = INP.get(k, ws[T.I[k]].value)
    inp["price"] = setup["price"]; inp["eff"] = f"{setup['eff'].year:04d}-{setup['eff'].month:02d}"
    inp["dm"] = -np.log(1 - inp["dmin"]); inp["life_m"] = min(T.NM, inp["life_yr"] * 12)
    P = wb["Prices"]
    d0 = P["B6"].value
    prices = dict(start=f"{d0.year:04d}-{d0.month:02d}", wti=[P.cell(r, 3).value for r in range(6, 606)], hh=[P.cell(r, 4).value for r in range(6, 606)])
    wells = []
    for f in fit["wells"]:
        v = dd.SUBJECT[f["api"]]; o = OVR.get(f["api"], {})
        wi, nri, lnri = v[6], v[7], v[8] if len(v) > 8 else 0.8
        w = dict(active=1 if (wi or nri) else 0, age=f["months"], k0=E.ym_diff(f["last_month"], inp["eff"]), mult=float(o.get("mult", 1.0)),
                 nri=nri, wiu=0.0 if setup["dtype"] == "Minerals" else wi, dfac=lnri if inp["roy"] == 1 else 1.0,
                 loe=float(LOE.get(f["api"], inp["loe_fix"])), lnri=lnri, hist_oil=f["oil"]["history"], hist_gas=f["gas"]["history"])
        for s, key in (("o", "oil"), ("g", "gas")):
            an = f[key]["anchor"]
            w[s + "_q"] = float(o.get(f"{s}_q", an["q_actual"] if inp["anchor"] == "Actual" else an["q_fit"]))
            w[s + "_da"] = float(o.get(f"{s}_da", an["Da"])); w[s + "_b"] = max(0.001, float(o.get(f"{s}_b", f[key]["b"])))
        wells.append(w)
    return inp, wells, prices


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
