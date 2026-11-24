# Non-op evaluation (Osage Manka)

The deal model, the scripts that fill it, and a record of every deal evaluated.

| Path | What it is |
|---|---|
| `Non-Op Evaluation Template.xlsx` | The model. Up to 10 wells; Working interest or Minerals (Inputs C18); basin defaults (Inputs C19, Basins tab: Williston, DJ, Permian, Powder River, Anadarko / Midcon); up to 6 zone type curves ('Zone Curves' tab, Wells column X). Carries the Sun 5899 example. |
| `PDP Evaluation Template.xlsx` | Producing-property model: up to 30 wells, hyperbolic decline per well from its own history, offset check over the forecast months, max price for the target IRR. Built by `toolkit/build_pdp_template.py`, filled by `toolkit/write_pdp.py`. See "Evaluate a PDP package". |
| `examples/pdp-example/` | Worked PDP example (King 3-65 28-29 pad, DJ; hypothetical interests). |
| `Deal Log.xlsx` | One row per deal evaluated, plus lookbacks of actual production vs the type curve. Rebuilt by `toolkit/deal_log.py`. |
| `toolkit/` | `build_tc.py` (type curves), `engine.py` (cash flows, ties to the workbook), `write_model.py` (fills a model), `plot_tc.py` (TypeCurve charts and table), `lookback.py`, `deal_log.py`, `lateral.py` (lateral length rule), `upgrade_template.py` / `upgrade_template_zones.py` / `upgrade_template_lateral.py` / `_phase.py` / `_steps.py` / `upgrade_template_anadarko.py` (one-time template upgrades, kept for the record). |
| `deals/<deal>/` | Per deal: `deal_data.py` (Enverus production and fit sets), `tc_summary.json`, `tc_curves.csv`, `deal.json` (terms, results, decision), charts, the finished model, and later `actuals.py` / `lookback.json`. |

## Evaluate a deal
1. New folder `deals/<deal-slug>/`. Pull offsets and monthly production from Enverus (5-mile radius, lateral midpoint to DSU centroid, comp screen) and write `deal_data.py` in the format of `deals/sun5899/deal_data.py`, including `DAYS` (Enverus `ProducingDays`, production dataset). If fewer than 15 wells pass the 5-mile screen, also pull the 10-mile set and list it in `FIT_SET_10MI`.
   - **Multi-zone deals** (wells landed in different formations, e.g. a Midland wine-rack DSU): define `ZONES` in `deal_data.py`
     (see `deals/loki-upton`), one entry per zone with its own fit set, in Enverus interval terms. Zone 1 (first) is the TypeCurve tab.
     Radius rule per zone: 5 mi; if fewer than 6 screened wells, widen to 10 / 15 / 20 mi until there are 6+ and record `radius` and a
     `note`; a 5-mile fit under 15 wells also lists `fit_10mi` for the comparison curve. Up to 6 zones.
2. `cd toolkit && python3 build_tc.py ../deals/<deal-slug>`
   - If a fit set has fewer than 3 wells with 24+ months (young offsets), `build_tc.py` fits b on the 24+ month wells of the same
     zone's 10-mile set (`fit_10mi` / `FIT_SET_10MI`) and applies it to the 5-mile mean; the summary records `b_borrowed` and `b_source`. (multi-zone: also writes `tc_zones.json` and `tc_zone_curves*.csv`)
   - Young plays where even the 10-mile set lacks 3 wells with 24+ months (e.g. the 2025-26 Anadarko Cherokee): list same-formation
     24+ month wells from the wider area in `B_SET` in `deal_data.py`; `build_tc.py` fits b on them first (before the 10-mile
     fallback) and applies it to the fit set's mean (single-zone deals; see `deals/lomas-7-18-1h`). A fit-set mean of 13 months or
     less searches its ramp peak with 6+ months after it.
3. Copy the template, fill Wells (multi-zone: the zone name in column X, exactly as in `ZONES`; leave the TC factor at 1), Inputs (deal type, basin, overrides in column E) and the Prices strip; recalculate (LibreOffice headless or Excel).
4. `python3 write_model.py <filled.xlsx> ../deals/<deal-slug> "../deals/<deal-slug>/<Deal> <Basin> Non-Op Model.xlsx"`, then recalculate the output and confirm Summary shows "All checks OK". `write_model.py` solves Summary C7 (premium, % of net capex) so the live-case IRR (C12) is 20%; the solve is kept in E7, and you can type over C7 to test a bid.
5. `python3 deal_log.py ../deals "../Deal Log.xlsx"`, commit.

## Evaluate a PDP package (`PDP Evaluation Template.xlsx`)
For producing wells: each well's own history is fitted with a hyperbolic decline and forecast month by month, and the
forecast is checked against what offset wells actually did over the same months of life. Up to 30 wells; Working interest
or Minerals; same Basins table, strip and price-deck dropdown as the non-op template. Worked example: `examples/pdp-example`
(the King 3-65 28-29 pad priced as if offered as PDP; hypothetical interests).
1. New folder `deals/<deal-slug>/` with `pdp_data.py` (format in the `toolkit/pdp_fit.py` docstring and the example):
   `DEAL` (name, basin, deal type, effective month, asking price), `LAST_MONTH`, `SUBJECT` (Enverus monthly oil / gas, lateral,
   zone, WI, NRI, lease NRI per well), `OFFSETS` (same-formation wells within 5 mi with more history than the subject wells),
   `DAYS` (Enverus ProducingDays), `ZONE_GROUPS` (intervals that compare, e.g. Niobrara A/B/C). Optional `FIT_FROM`,
   `LOE`, `INPUTS` (Inputs overrides) and `OVERRIDES` (per-well anchor / Da / b).
2. `cd toolkit && python3 pdp_fit.py ../deals/<deal-slug>` - fits every well (oil and gas separately, b fitted, 7% terminal),
   builds the anchored offset bands and saves `pdp_fit.json` and a chart per well (`pdp_<api>.png`).
3. `python3 write_pdp.py ../deals/<deal-slug> "../deals/<deal-slug>/<Deal> <Basin> PDP Model.xlsx"`, recalculate it
   (`recalc.py` or Excel), then `python3 pdp_engine.py "<model>.xlsx"` must print "engine ties"; Summary C45 = "All checks OK".
4. Read Summary (max price for the target IRR, PV10/15/20, reserves, $/flowing BOE/d, price / NTM cash flow), then Offset Check
   and Well Review for any well outside the tolerance. Override qa / Da / b on Wells where the offsets are the better guide.
5. `python3 deal_log.py ../deals "../Deal Log.xlsx"`, commit.

Fit rules (`pdp_fit.py`): fit window from the peak (first 6 months) to the last month, last 12 months weighted 3x; months
left out: zero months, producing days under 70% of the calendar month, outliers ~45% below / ~100% above the trend. A level
shift (the last 3-12 months all 22%+ off the earlier trend, e.g. an allocation change or curtailment) gets a level factor so
the shape comes from all months and the level from the recent ones; it is noted on Wells. Under 12 good months: b from
the offsets' median shape. Forecast anchor: "Fit" (fitted rate at the last month) or "Actual" (last three good months on
the fitted shape), Inputs C13. Offset check: each offset is scaled to the well's own last three months at the same age and
its next 120 months read off; P25/P50/P75 where 3+ report; the forecast's cumulative vs the P50's over 12/24/36/60 months
and the full overlap, "In line" within +/-15% (Inputs C18). Economic limit: a well stops the first month its 8/8 operating
cash flow is not positive (also for Minerals, since the operator plugs it); WI pays P&A that month.

`toolkit/build_pdp_template.py` rebuilds the blank template (layout lives there; re-run `write_pdp.py` for deals after a change).

## Lateral length rule (`toolkit/lateral.py`)
Type curves are kept per 10,000 ft. A curve moved to another lateral keeps a linear EUR, but its early months are not
linear: months 1-3 (Inputs C25) get a factor of linear that falls from 100% at 10,000 ft (linear band C24 = 0 ft) in a
straight line to 75% (C26) at 5,000 ft away (C27, 3 miles vs 2), then to 67.5% (C28) at 10,000 ft away (C29, 4 miles vs 2),
67.5% beyond. The 4-mile factor is a placeholder (too few 4-mile wells to calibrate) that keeps the old 7.5-point step
from 3 to 4 miles; without it a curve built on 3-mile offsets would carry no penalty onto 4-mile wells. Going to a longer lateral it is linear x factor, going shorter linear / factor. From month 4
the hyperbolic decline is re-solved (same b, a time stretch s; 7% terminal kept) so the 360-month EUR is exactly linear.
- Calibration (Nov 23, 2026, Enverus `lateral_stats`, first production 2020-2024, 3-mile vs 2-mile per foot, matched by
  operator x year x interval x county): Williston Middle Bakken oil 0.69 / 0.70 / 0.74 / 0.79 / 0.88 and DJ Niobrara A-C +
  Codell oil 0.62 / 0.68 / 0.77 / 0.78 / 0.90 at 3 / 6 / 9 / 12 / 36 months. 70% for 3 months with a linear EUR
  reproduces that; 75% allows for lighter completions on DJ 3-mile wells and for where long laterals get drilled. EUR
  itself could not be tested (Enverus EUR blank). Previous rule: 2,500 ft band, months 1-8, 92.5% at 3 miles, 85% at 4.
- Offsets: each is normalized straight to 10,000 ft on the same rule: months 1-3 get linear x the factor for a shorter
  offset (5,000 ft: x 2 x 0.75) and linear / factor for a longer one (20,000 ft: x 0.5 / 0.75), and its decline from
  month 4 is re-shaped so its EUR is still linear. The fit is made on that basis; the charts and Offset Production tab
  show the same normalized offsets.
- Wells: the 'Lateral Curves' tab moves each well's curve from 10,000 ft to its lateral (s, Tsw, g solved by `write_model.py`);
  the CF tabs read it. Summary row 76 flags a well whose EUR no longer ties to linear: re-run `write_model.py`.
- Deals modeled before Nov 2026 were linear; `lookback.py` keeps them linear unless the deal is rebuilt.

## Record the decision and look back
- Edit `deals/<deal>/deal.json`: `decision` (Bought / Passed / Lost / Pending), `price_paid`, `closed`, `notes`. Rebuild the deal log.
- Once the wells produce: pull their monthly production (with producing days) into `deals/<deal>/actuals.py` as `ACTUALS = {api: ("YYYY-MM", [oil], [gas], [days])}`, run `python3 lookback.py ../deals/<deal>`, rebuild the deal log.

Requires Python 3 with numpy, scipy, matplotlib, Pillow and openpyxl. Run the scripts from `toolkit/` (they import each other).
