"""Offsets vs proposed type curve: charts, monthly table and the Offset Production tab.

Called by write_model.py (or alone: python plot_tc.py <model.xlsx> <deal_dir> <out.xlsx>).

For oil and for gas, one chart on the TypeCurve tab:
  * x = months on production (1, 2, 3 ...), y = monthly volume per 10,000 ft (log scale)
  * box and whisker per month across the fit-set wells (box = P25-P75, line = median,
    whiskers = min-max so every well is inside), each well's volume as a small gray dot
  * P50 (median of the wells reporting that month) as large purple circles
  * proposed type curve as small orange diamonds joined by a line
  * lower panel: number of wells above (up) and below (down) the type curve each month
  * months where fewer than half the fit set reports are shaded (outside the fit window)

Volumes use the same basis as build_tc.py: per 10,000 ft of lateral, month 1 = first producing
month (scaled to a full month when producing days are known; otherwise dropped if partial).

The workbook also gets an "Offset Production" tab (one column per fit well, oil and gas) and a
month-by-month table on TypeCurve (H20 down) with live formulas: wells reporting, P50,
type curve (linked to columns C / D), wells above and below. The charts are pictures, so they
are redrawn by this script; the table recalculates in Excel.
"""
import io, math, os, sys
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, MaxNLocator
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
import build_tc as bt

PURPLE, ORANGE = "#6B2FB3", "#F07A1A"
ABOVE, BELOW = "#2A78D6", "#B0ACA3"
INK, MUTED, GRID = "#2B2A27", "#6B6A63", "#E4E3DD"
MARKER = "TC vs offsets charts v1"          # TypeCurve!N1: charts at N2 / N42 belong to this script
ANCHORS = ("N2", "N42")
SHIFT = 82                                   # rows the template's older pictures move down (first run only)


def offsets(apis):
    """Per-well cleaned monthly oil and gas per 10,000 ft on the lateral rule, as build_tc.py fitted them
    (tc_offsets.json); falls back to linear normalization for an older deal folder."""
    import json
    path = os.path.join(bt.DEAL_DIR, "tc_offsets.json")
    saved = json.load(open(path)) if os.path.exists(path) else {}
    oil, gas = {}, {}
    for a in apis:
        if a in saved:
            oil[a], gas[a] = np.array(saved[a][0]), np.array(saved[a][1])
            continue
        name, lat, first, o, g = bt.WELLS[a][:5]
        o, g = bt.clean(o, g, bt.DAYS.get(a), first)
        oil[a], gas[a] = o * 10000 / lat, g * 10000 / lat
    return oil, gas


def month_stats(series, tc):
    n = max(len(v) for v in series.values())
    rows = []
    for m in range(n):
        v = np.array([s[m] for s in series.values() if len(s) > m])
        rows.append(dict(vals=v, n=len(v), p50=float(np.median(v)), t=float(tc[m]),
                         above=int((v > tc[m]).sum()), below=int((v < tc[m]).sum())))
    return rows


def _fmt(x, _):
    return f"{x/1000:g}k" if x >= 1000 else f"{x:g}"


def chart(series, tc, fit_n, title, unit, fname, tc10=None, n10=None):
    st = month_stats(series, tc)
    M = len(st); x = np.arange(1, M + 1)
    need = math.ceil(fit_n / 2)
    fig, (ax, bx) = plt.subplots(2, 1, figsize=(14, 7.8), dpi=100, sharex=True,
                                 gridspec_kw=dict(height_ratios=[3.2, 1], hspace=0.16))
    for a in (ax, bx):
        for s in a.spines.values():
            s.set_visible(False)
        a.tick_params(colors=MUTED, length=0, labelsize=9)
        a.grid(axis="y", color=GRID, lw=0.8); a.set_axisbelow(True)
    thin = [m + 1 for m in range(M) if st[m]["n"] < need]
    if thin:
        for a in (ax, bx):
            a.axvspan(thin[0] - 0.5, M + 0.5, color="#F3F2EE", zorder=0)
        ax.text(thin[0], 0.985, f"  Fewer than {need} of {fit_n} wells reporting (outside the fit window)",
                transform=ax.get_xaxis_transform(), va="top", fontsize=8.5, color=MUTED, wrap=True)

    data = [s["vals"] for s in st]
    ax.boxplot(data, positions=x, widths=0.62, whis=(0, 100), showfliers=False, patch_artist=True, manage_ticks=False,
               boxprops=dict(facecolor="#FFFFFF", edgecolor="#8C8A82", lw=0.9),
               medianprops=dict(color="#8C8A82", lw=0.9), whiskerprops=dict(color="#8C8A82", lw=0.9),
               capprops=dict(color="#8C8A82", lw=0.9), zorder=2)
    rng = np.random.default_rng(3)
    for m, v in enumerate(data):
        ax.scatter(m + 1 + rng.uniform(-0.18, 0.18, len(v)), v, s=7, color="#A3A199", lw=0, zorder=3)
    ax.scatter(x, [s["p50"] for s in st], s=110, color=PURPLE, edgecolor="white", lw=1.2, zorder=5,
               label="P50 of offset wells (median each month)")
    ax.plot(x, tc[:M], color=ORANGE, lw=1.6, marker="D", ms=4.5, mec="white", mew=0.6, zorder=6,
            label="Proposed type curve")
    if tc10 is not None:
        ax.plot(x, tc10[:M], color="#2A78D6", lw=1.4, ls="--", zorder=4, label=f"10-mile curve ({n10} wells), for comparison")
    ax.scatter([], [], s=7, color="#A3A199", label="Individual offset wells")
    ax.plot([], [], color="#8C8A82", lw=0.9, label="Box = P25-P75; whiskers = min-max")
    ax.set_yscale("log"); ax.yaxis.set_major_formatter(FuncFormatter(_fmt)); ax.yaxis.set_minor_formatter(FuncFormatter(lambda *_: ""))
    ax.set_ylabel(f"{unit} per month per 10,000 ft", color=MUTED, fontsize=10)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=5 if tc10 is not None else 4, fontsize=9, frameon=False, labelcolor=INK, handletextpad=0.4, columnspacing=1.6)
    ax.set_title(title, loc="left", fontsize=12, color=INK, pad=30)

    ab = np.array([s["above"] for s in st]); be = np.array([s["below"] for s in st])
    bx.bar(x, ab, width=0.72, color=ABOVE, label="Wells above type curve")
    bx.bar(x, -be, width=0.72, color=BELOW, label="Wells below type curve")
    bx.axhline(0, color=MUTED, lw=0.8)
    top = max(ab.max(), be.max(), 1)
    bx.set_ylim(-top - 0.6, top + 0.6)
    bx.yaxis.set_major_locator(MaxNLocator(integer=True, nbins=5))
    bx.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{abs(int(v))}"))
    bx.set_ylabel("Wells", color=MUTED, fontsize=10)
    bx.legend(loc="upper right", fontsize=8.5, frameon=False, ncol=2, labelcolor=INK)
    tot_a, tot_b = int(ab[:M if not thin else thin[0] - 1].sum()), int(be[:M if not thin else thin[0] - 1].sum())
    bx.set_title(f"Wells above / below the type curve each month. Inside the fit window: {tot_a} well-months above, "
                 f"{tot_b} below ({tot_a / max(tot_a + tot_b, 1):.0%} above)", loc="left", fontsize=9, color=MUTED, pad=4)
    bx.set_xlabel("Months on production", color=MUTED, fontsize=10)
    step = 3 if M <= 60 else 6
    ticks = np.r_[1, np.arange(step, M + 1, step)] if M > 12 else x
    bx.set_xticks(ticks); bx.set_xticklabels([str(t) for t in ticks]); bx.set_xlim(0.3, M + 0.7)
    fig.savefig(fname, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return st


def add_to_workbook(wb, deal_dir, oil_tc, gas_tc, fit_apis, title_prefix, low_conf, ten_mile=None):
    """ten_mile: optional (oil curve, gas curve, summary) of the 10-mile comparison curve, drawn dashed."""
    t_o, t_g, n10 = (ten_mile[0], ten_mile[1], ten_mile[2]["wells"]) if ten_mile else (None, None, None)
    oil, gas = offsets(fit_apis)
    n = len(fit_apis)
    tag = f"{n} fit wells{', LOW CONFIDENCE' if low_conf else ''}"
    so = chart(oil, oil_tc, n, f"{title_prefix} - Oil: offsets vs proposed type curve ({tag})", "bbl",
               os.path.join(deal_dir, "tc_vs_offsets_oil.png"), t_o, n10)
    sg = chart(gas, gas_tc, n, f"{title_prefix} - Gas: offsets vs proposed type curve ({tag})", "Mcf",
               os.path.join(deal_dir, "tc_vs_offsets_gas.png"), t_g, n10)

    # ---- Offset Production tab (values, per 10,000 ft)
    if "Offset Production" in wb.sheetnames:
        del wb["Offset Production"]
    op = wb.create_sheet("Offset Production")
    hdr = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="1F3864")
    blue = Font(color="0000FF")
    op["B2"] = "Offset fit-set monthly production per 10,000 ft of lateral (gross 8/8)"; op["B2"].font = Font(bold=True, size=14)
    op["B3"] = ("Source: Enverus monthly production (deal_data.py). Month 1 = first producing month, scaled to a full month "
                "from Enverus producing days where known (else dropped if partial); partial last month dropped. Each well is normalized to 10,000 ft: linear within +/- 2,500 ft; otherwise months 1-8 get the linear factor x 90% for a shorter well (/ 90% for a longer one; phased in from 2,500 to 5,000 ft away) and its decline is re-shaped so its EUR still scales linearly (a re-shaped well can show a few more or fewer months than it reported). Same basis as the type curve fit. Feeds the TypeCurve table (H20).")
    M = max(len(v) for v in oil.values())
    oc, gc = 3, 3 + n + 1            # oil block C.., gas block after one blank column
    op.cell(5, 2, "Month").font = hdr; op.cell(5, 2).fill = fill
    for j, a in enumerate(fit_apis):
        name, lat, first = bt.WELLS[a][:3]
        for col, label in ((oc + j, "Oil (bbl)"), (gc + j, "Gas (Mcf)")):
            op.cell(5, col, f"{name}\n{label}").font = hdr; op.cell(5, col).fill = fill
            op.cell(5, col).alignment = Alignment(wrap_text=True, vertical="top")
            op.cell(6, col, f"API {a}"); op.cell(7, col, lat); op.cell(8, col, first)
            op.column_dimensions[L(col)].width = 15
    for r, lab in ((6, "API"), (7, "Lateral (ft)"), (8, "First month")):
        op.cell(r, 2, lab).font = Font(italic=True, color="6B6A63")
    op.row_dimensions[5].height = 45
    r0 = 9
    for m in range(M):
        op.cell(r0 + m, 2, m + 1)
        for j, a in enumerate(fit_apis):
            for col, s in ((oc + j, oil[a]), (gc + j, gas[a])):
                if len(s) > m:
                    c = op.cell(r0 + m, col, round(float(s[m]), 1)); c.font = blue; c.number_format = "#,##0"
    op.freeze_panes = "C9"
    rlast = r0 + M - 1
    orng = f"'Offset Production'!${L(oc)}{{r}}:${L(oc + n - 1)}{{r}}"
    grng = f"'Offset Production'!${L(gc)}{{r}}:${L(gc + n - 1)}{{r}}"

    # ---- TypeCurve: monthly table at H19, picture anchors, marker
    tc = wb["TypeCurve"]
    for r in range(20, 20 + 2 * (M + 4) + 6):
        for c in range(8, 14):
            tc.cell(r, c).value = None
    tc["H20"] = "Offsets vs proposed type curve by month (per 10,000 ft; formulas on 'Offset Production')"
    tc["H20"].font = Font(bold=True)
    thin_b = Border(bottom=Side(style="thin", color="8C8A82"))
    row = 21
    for label, rng, col, unit in (("Oil", orng, "C", "bbl"), ("Gas", grng, "D", "Mcf")):
        heads = ["Month", "Wells reporting", f"P50 {label.lower()} ({unit})", f"Type curve {label.lower()} ({unit})",
                 "Wells above TC", "Wells below TC"]
        for j, h in enumerate(heads):
            c = tc.cell(row, 8 + j, h); c.font = hdr; c.fill = fill; c.alignment = Alignment(wrap_text=True)
        row += 1
        for m in range(M):
            src = r0 + m; rr = rng.format(r=src)
            tc.cell(row, 8, m + 1).alignment = Alignment(horizontal="left")
            tc.cell(row, 9, f"=COUNT({rr})")
            tc.cell(row, 10, f'=IF(I{row}=0,"",MEDIAN({rr}))').number_format = "#,##0"
            tc.cell(row, 11, f"={col}{6 + m}").number_format = "#,##0"
            tc.cell(row, 12, f'=COUNTIF({rr},">"&K{row})')
            tc.cell(row, 13, f'=COUNTIF({rr},"<"&K{row})')
            row += 1
        for j in range(6):
            tc.cell(row - 1, 8 + j).border = thin_b
        tc.cell(row, 8, "Total").font = Font(bold=True)
        tc.cell(row, 12, f"=SUM(L{row - M}:L{row - 1})").font = Font(bold=True)
        tc.cell(row, 13, f"=SUM(M{row - M}:M{row - 1})").font = Font(bold=True)
        row += 3
    for c in "JKLM":
        if tc.column_dimensions[c].width in (None, 13.0):
            tc.column_dimensions[c].width = 14

    first_run = tc["N1"].value != MARKER
    keep = []
    for im in tc._images:
        a = im.anchor._from
        if not first_run and (a.col, a.row) in ((13, 1), (13, 41)):
            continue                                    # our charts from an earlier run
        if first_run:
            a.row += SHIFT;
            if hasattr(im.anchor, "to") and im.anchor.to is not None:
                im.anchor.to.row += SHIFT
        keep.append(im)
    tc._images = keep
    for path, anc in zip(("tc_vs_offsets_oil.png", "tc_vs_offsets_gas.png"), ANCHORS):
        img = XLImage(os.path.join(deal_dir, path))   # native size, about 1,220 x 720 px (36 rows)
        tc.add_image(img, anc)
    tc["N1"] = MARKER; tc["N1"].font = Font(color="FFFFFF", size=6)
    note = tc["B3"].value or ""
    if "TC vs offsets" not in note:
        tc["B3"] = (note + " Charts at N2 (oil) and N42 (gas): TC vs offsets with monthly box plots; "
                    "month-by-month table at H20.").strip()
    return so, sg


ZONE_ROW0, ZONE_ROWS = 14, 80          # 'Zone Curves': zone 2 oil chart at P14, gas at P54, zone 3 at P94 ...


def add_zone_charts(wb, deal_dir, zones, curves, curves10, title_prefix):
    """Oil and gas charts (same layout as the TypeCurve charts) for zones 2-6 on the 'Zone Curves' tab.
    zones: tc_zones.json list; curves / curves10: {name: (oil, gas)} (10-mile entry None when not built)."""
    zc = wb["Zone Curves"]
    zc._images = []
    for k, z in enumerate(zones[1:]):
        name = z["name"]; n = len(z["fit_set"])
        oil, gas = offsets(z["fit_set"])
        tag = f"{n} fit wells, {z.get('radius', '5 mi')}{', LOW CONFIDENCE' if z.get('low_confidence') else ''}"
        t10 = curves10.get(name)
        n10 = z.get("ten_mile", {}).get("wells")
        slug = "".join(ch if ch.isalnum() else "_" for ch in name.lower()).strip("_")
        for j, (ser, tc, unit, lab) in enumerate(((oil, curves[name][0], "bbl", "Oil"), (gas, curves[name][1], "Mcf", "Gas"))):
            fname = os.path.join(deal_dir, f"tc_zone_{slug}_{lab.lower()}.png")
            chart(ser, tc, n, f"{title_prefix} - {name} {lab.lower()}: offsets vs zone type curve ({tag})", unit, fname,
                  None if t10 is None else t10[j], n10)
            zc.add_image(XLImage(fname), f"P{ZONE_ROW0 + k * ZONE_ROWS + j * 40}")


if __name__ == "__main__":
    import json, openpyxl
    src, deal, out = sys.argv[1:4]
    bt.load_deal(deal)
    C = np.loadtxt(os.path.join(deal, "tc_curves.csv"), delimiter=",", skiprows=1)
    S = json.load(open(os.path.join(deal, "tc_summary.json")))
    wb = openpyxl.load_workbook(src)
    title = (wb["Summary"]["B2"].value or "Deal").split(":")[0]
    add_to_workbook(wb, deal, C[:, 1], C[:, 2], S["fit_set"], title, S["wells"] < 15)
    wb.save(out)
    print("wrote", out)
