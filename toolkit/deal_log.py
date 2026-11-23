"""Deal log: one row per evaluated deal, plus lookbacks for deals with production.

Usage: python deal_log.py <deals_root> <Deal Log.xlsx>
Reads every <deals_root>/<deal>/deal.json (written by write_model.py) and, where present, lookback.json
(written by lookback.py). Rebuilds the whole workbook each time, so edit decisions in deal.json
("decision": "Bought" / "Passed" / "Lost" / "Pending", "price_paid", "closed", "notes"), not in the workbook.
"""
import glob, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

HDR = dict(font=Font(bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor="1F3864"))


def header(ws, row, cols, widths):
    for j, (h, w) in enumerate(zip(cols, widths)):
        c = ws.cell(row, 2 + j, h); c.font = HDR["font"]; c.fill = HDR["fill"]
        c.alignment = Alignment(wrap_text=True, vertical="top"); ws.column_dimensions[L(2 + j)].width = w
    ws.row_dimensions[row].height = 32


def main(root, out):
    recs = []
    for p in sorted(glob.glob(os.path.join(root, "*", "deal.json"))):
        d = json.load(open(p)); d["_dir"] = os.path.basename(os.path.dirname(p))
        lb = os.path.join(os.path.dirname(p), "lookback.json")
        d["_lb"] = json.load(open(lb)) if os.path.exists(lb) else None
        recs.append(d)
    recs.sort(key=lambda d: d.get("evaluated", ""), reverse=True)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Deals"
    ws["B2"] = "Deal log: every deal evaluated with the Non-Op Evaluation Template"; ws["B2"].font = Font(bold=True, size=14)
    ws["B3"] = ("Rebuilt by deal_log.py from deals/*/deal.json. Base case = strip with a 20% premium (WI) or the asking price (Minerals). "
                "Edit decisions in each deal.json. Lookback columns fill in once the wells produce and lookback.py has run.")
    cols = ["Evaluated", "Deal", "Folder", "Type", "Basin", "Wells", "Net capex ($)", "Asking price ($)", "Base IRR", "MOIC", "NPV10 ($)",
            "Max for 20% IRR", "Fit wells", "Low confidence", "Oil EUR / 10k ft (MBO)", "Gas EUR / 10k ft (Bcf)", "Decision", "Price paid ($)",
            "Closed", "Lookback as of", "Oil vs curve", "Gas vs curve", "Notes"]
    header(ws, 5, cols, [11, 40, 14, 15, 15, 7, 13, 13, 9, 7, 13, 14, 8, 10, 12, 12, 11, 13, 11, 12, 10, 10, 50])
    for i, d in enumerate(recs):
        r = 6 + i; tc = d.get("type_curve", {}); lb = d["_lb"]
        mx = d.get("max_for_20pct_irr")
        vals = [d.get("evaluated"), d.get("deal"), d["_dir"], d.get("type"), d.get("basin"), len(d.get("wells", [])), d.get("net_capex"),
                d.get("asking_price"), d.get("irr"), d.get("moic"), d.get("npv10"), mx, tc.get("fit_wells"),
                "Yes" if tc.get("low_confidence") else "No", (tc.get("oil_eur_per_10k") or 0) / 1000, (tc.get("gas_eur_per_10k") or 0) / 1e6,
                d.get("decision"), d.get("price_paid"), d.get("closed"), lb and lb.get("as_of"),
                lb and lb["package"].get("oil_vs_curve"), lb and lb["package"].get("gas_vs_curve"), d.get("notes")]
        fmts = [None, None, None, None, None, "0", "$#,##0", "$#,##0", "0.0%", "0.00x", "$#,##0",
                "$#,##0" if d.get("type") == "Minerals" else "0.0%", "0", None, "0", "0.00", None, "$#,##0", None, None,
                "+0%;-0%", "+0%;-0%", None]
        for j, (v, f) in enumerate(zip(vals, fmts)):
            c = ws.cell(r, 2 + j, v)
            if f:
                c.number_format = f
    ws.freeze_panes = "D6"

    lw = wb.create_sheet("Lookbacks")
    lw["B2"] = "Lookbacks: actual production vs the type curve each deal was priced on"; lw["B2"].font = Font(bold=True, size=14)
    lw["B3"] = ("Per well, from lookback.py. Prediction starts in the well's actual first month; the timing gap is shown separately. "
                "A pattern across deals (curves running high or low in a basin) is the signal to adjust the type-curve method.")
    cols = ["Deal", "Basin", "Well", "API", "Model first sales", "Actual first prod", "Timing gap (months)", "Months on",
            "Cum oil (bbl)", "Type curve oil (bbl)", "Oil vs curve", "Cum gas (Mcf)", "Type curve gas (Mcf)", "Gas vs curve", "As of"]
    header(lw, 5, cols, [36, 15, 28, 13, 11, 11, 10, 8, 13, 13, 10, 13, 13, 10, 11])
    r = 6
    for d in recs:
        lb = d["_lb"]
        if not lb:
            continue
        for w in lb["wells"]:
            vals = [d.get("deal"), d.get("basin"), w.get("name"), w.get("api"), w.get("model_first_sales"), w.get("actual_first"),
                    w.get("timing_gap_months"), w.get("months"), w.get("cum_oil"), w.get("pred_oil"), w.get("oil_vs_curve"),
                    w.get("cum_gas"), w.get("pred_gas"), w.get("gas_vs_curve"), lb.get("as_of")]
            fmts = [None] * 7 + ["0", "#,##0", "#,##0", "+0%;-0%", "#,##0", "#,##0", "+0%;-0%", None]
            if "status" in w:
                vals[4] = w["status"]
            for j, (v, f) in enumerate(zip(vals, fmts)):
                c = lw.cell(r, 2 + j, v)
                if f:
                    c.number_format = f
            r += 1
    if r == 6:
        lw["B6"] = "No lookbacks yet: run lookback.py on a deal once its wells have production."
    lw.freeze_panes = "E6"
    wb.save(out)
    print(f"wrote {out}: {len(recs)} deals")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
