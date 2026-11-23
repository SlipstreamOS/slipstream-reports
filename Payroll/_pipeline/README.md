# Slipstream Payroll Pipeline (v2)

Moved here from the old local Cowork session on 2026-09-30 so cloud sessions can use it.
Full rules/history are in Claude project memory: `reference_payroll_pipeline.md`.

## Layout
- `payroll/v2/build_template_v2.py` — source of truth: ROSTER (wages, ADP codes, Rate 2 flags), MIN_WAGE ($18.50), tipped min ($12), checks. Regenerates `template_v2.xlsx`.
- `payroll/v2/populate_v2.py` — fills a period workbook from a 16-col shifts CSV + cc_tips.json (auto weekly OT split; excludes Ines + Mo).
- `payroll/v2/finalize_v2.py` — builds the ADP ##GENERIC## V2.0 import CSV (IID 26642723) + prints FLAGS (missing codes, Rate 2, Event/Rate 3).
- `payroll/run_pull.py` — wage audit (Asana vs Square) + timecard review. Reads `cache/`.
- `cache/` — square_team.json, square_jobs.json, asana_team.json, periods/<period>/ (timecards, cc_tips).

## Each pay period (Sun–Sat, 2 weeks)
1. Refresh Asana — pull BOTH "Team Summary" (gid 1211816222497942) AND "Leadership Team Summary" (added 2026-10; leaders moved there) → merge into cache/asana_team.json → compare vs ROSTER, update drifts.
2. Refresh Square team (team.searchMembers) + timecards (labor.searchTimecards, filter key `start`).
3. Tips from `Weekly and Monthly Reports/_pipeline/shared/daily_metrics.csv` (tips column, per cafe per day; verified vs Oct cc_tips). Write cache/periods/<id>/cc_tips.json as {"Slipstream <Cafe>": {"M/D/YY": dollars}}.
4. Synthesize 16-col CSV → `python3 populate_v2.py <csv> <cc_tips.json> <YYYY-MM-DD> <out.xlsx>`
5. Adjustments tab: Colin OTH $480 (standing, unless told otherwise) + any PTO (VAC) from Ryan.
6. `python3 finalize_v2.py <out.xlsx> <YYYY-MM-DD> <out.csv>` → recalc, verify all Checks, deliver to `Payroll/<period>/`.
