# Non-op evaluation (Osage Manka)

The deal model, the scripts that fill it, and a record of every deal evaluated.

| Path | What it is |
|---|---|
| `Non-Op Evaluation Template.xlsx` | The model. Up to 10 wells; Working interest or Minerals (Inputs C18); basin defaults (Inputs C19, Basins tab); up to 6 zone type curves ('Zone Curves' tab, Wells column X). Carries the Sun 5899 example. |
| `Deal Log.xlsx` | One row per deal evaluated, plus lookbacks of actual production vs the type curve. Rebuilt by `toolkit/deal_log.py`. |
| `toolkit/` | `build_tc.py` (type curves), `engine.py` (cash flows, ties to the workbook), `write_model.py` (fills a model), `plot_tc.py` (TypeCurve charts and table), `lookback.py`, `deal_log.py`, `lateral.py` (lateral length rule), `upgrade_template.py` / `upgrade_template_zones.py` / `upgrade_template_lateral.py` (one-time template upgrades, kept for the record). |
| `deals/<deal>/` | Per deal: `deal_data.py` (Enverus production and fit sets), `tc_summary.json`, `tc_curves.csv`, `deal.json` (terms, results, decision), charts, the finished model, and later `actuals.py` / `lookback.json`. |

## Evaluate a deal
1. New folder `deals/<deal-slug>/`. Pull offsets and monthly production from Enverus (5-mile radius, lateral midpoint to DSU centroid, comp screen) and write `deal_data.py` in the format of `deals/sun5899/deal_data.py`, including `DAYS` (Enverus `ProducingDays`, production dataset). If fewer than 15 wells pass the 5-mile screen, also pull the 10-mile set and list it in `FIT_SET_10MI`.
   - **Multi-zone deals** (wells landed in different formations, e.g. a Midland wine-rack DSU): define `ZONES` in `deal_data.py`
     (see `deals/loki-upton`), one entry per zone with its own fit set, in Enverus interval terms. Zone 1 (first) is the TypeCurve tab.
     Radius rule per zone: 5 mi; if fewer than 6 screened wells, widen to 10 / 15 / 20 mi until there are 6+ and record `radius` and a
     `note`; a 5-mile fit under 15 wells also lists `fit_10mi` for the comparison curve. Up to 6 zones.
2. `cd toolkit && python3 build_tc.py ../deals/<deal-slug>` (multi-zone: also writes `tc_zones.json` and `tc_zone_curves*.csv`)
3. Copy the template, fill Wells (multi-zone: the zone name in column X, exactly as in `ZONES`; leave the TC factor at 1), Inputs (deal type, basin, overrides in column E) and the Prices strip; recalculate (LibreOffice headless or Excel).
4. `python3 write_model.py <filled.xlsx> ../deals/<deal-slug> "../deals/<deal-slug>/<Deal> <Basin> Non-Op Model.xlsx"`, then recalculate the output and confirm Summary shows "All checks OK". `write_model.py` solves Summary C7 (premium, % of net capex) so the live-case IRR (C12) is 20%; the solve is kept in E7, and you can type over C7 to test a bid.
5. `python3 deal_log.py ../deals "../Deal Log.xlsx"`, commit.

## Lateral length rule (`toolkit/lateral.py`)
Type curves are kept per 10,000 ft, but a curve moved more than 2,500 ft (Inputs C24) is not scaled linearly: months 1-8
(C25) get a factor of linear phased from 100% at 2,500 ft away to 90% (C26) at 5,000 ft away (C27) and beyond, going to a
longer lateral (linear / factor going shorter), and from month 9 the hyperbolic decline
is re-solved (same b, a time stretch s; 7% terminal kept) so the 360-month EUR is exactly linear. Within the band it is linear.
- Offsets: each is normalized straight to 10,000 ft. Within 2,500 ft of 10,000 it is linear (8,000 ft x 1.25 every month);
  further away months 1-8 get linear x 90% for a shorter offset (5,000 ft: x 2 x 0.9) and linear / 90% for a longer one
  (20,000 ft: x 0.5 / 0.9), phased in between 2,500 and 5,000 ft, and its decline from month 9 is re-shaped so its EUR is
  still linear. The fit is made on that basis; the charts and Offset Production tab show the same normalized offsets.
- Wells: the 'Lateral Curves' tab moves each well's curve from 10,000 ft to its lateral (s, Tsw, g solved by `write_model.py`);
  the CF tabs read it. Summary row 76 flags a well whose EUR no longer ties to linear: re-run `write_model.py`.
- Deals modeled before Nov 2026 were linear; `lookback.py` keeps them linear unless the deal is rebuilt.

## Record the decision and look back
- Edit `deals/<deal>/deal.json`: `decision` (Bought / Passed / Lost / Pending), `price_paid`, `closed`, `notes`. Rebuild the deal log.
- Once the wells produce: pull their monthly production (with producing days) into `deals/<deal>/actuals.py` as `ACTUALS = {api: ("YYYY-MM", [oil], [gas], [days])}`, run `python3 lookback.py ../deals/<deal>`, rebuild the deal log.

Requires Python 3 with numpy, scipy, matplotlib, Pillow and openpyxl. Run the scripts from `toolkit/` (they import each other).
