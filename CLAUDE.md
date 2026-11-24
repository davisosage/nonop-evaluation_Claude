# Instructions for Claude in this repository

This repo is Davis's (Osage Manka) non-op / minerals deal evaluation system. The `nonop-deal-evaluation` skill describes the
workflow; this file covers the repository itself.

- The template and toolkit here are the current versions. Don't ask Davis to attach them.
- Each deal gets its own folder under `deals/`, named with a short slug (e.g. `leslie-fed-3670`). Put the finished model there.
- Never edit `Non-Op Evaluation Template.xlsx` for a single deal. Copy it into the deal folder first. Change the template only
  when Davis asks to change the model, and then also change the toolkit scripts it touches, rebuild the Sun 5899 example
  (`deals/sun5899`), recalculate, and confirm "All checks OK".
- After every evaluation, rebuild `Deal Log.xlsx` with `toolkit/deal_log.py` and commit the deal folder, the log and any
  toolkit change together, with a message that names the deal or the change. Push to `main`.
- Decisions live in each `deal.json` (`decision`, `price_paid`, `closed`, `notes`); rebuilds of the model keep them.
- When asked for lookbacks, pull actual production for bought deals' wells from Enverus into `deals/<deal>/actuals.py`
  and run `toolkit/lookback.py`, then rebuild the log and summarize which deals run above or below their curves.
- Run scripts from `toolkit/` without `python3 -I` (they import each other).
- PDP (producing-property) deals use `PDP Evaluation Template.xlsx` and `pdp_fit.py` / `write_pdp.py` / `pdp_engine.py`
  (README, "Evaluate a PDP package"). Change its layout in `toolkit/build_pdp_template.py`, rebuild the template, re-run
  `examples/pdp-example`, recalculate and confirm "All checks OK" and that `pdp_engine.py` prints "engine ties".
