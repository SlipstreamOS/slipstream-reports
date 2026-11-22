# Weekly Report Pipeline

Generates the Slipstream Weekly Business Review. Covers Sun–Sat for the 4 cafes (Downtown, Studio, 1899 Penn, NMWA).

## Outputs (land at the top of `Weekly and Monthly Reports/`)

- `Slipstream Weekly Data.xlsx` — data layer. Tabs: Daily-by-Cafe, Company Summary, History (appends each week, idempotent by week+cafe), README.
- `Slipstream Weekly Report - <dates> v0.1.docx` — narrative report.

## Folder layout

```
_pipeline/weekly/
  scripts/
    _paths.py                 # shared path resolution (portable across chats)
    projections.py            # loads sales + labor $ projections from Projections workbook
    aggregate_week.py         # Square orders → data/week_<label>.json
    aggregate_labor.py        # Square timecards → data/labor_<label>.json
    build_weekly_data.py      # JSONs + projections → xlsx
    build_weekly_report.py    # JSONs + projections → docx
  data/                       # intermediate JSONs (one per week)
  raw/                        # raw MCP captures (labor_week.json extracted from transcript)
```

## Run order (for a new Sun–Sat week)

1. **Pull orders from Square** via the Square MCP `orders.search` tool. The aggregator scans all tool-results folders visible to the current chat, so just run the pulls in the same chat where you'll build the report. Recommended slicing: one call per cafe (use location_id), per day, small page size (catering days on Downtown/Studio can have large tickets).
2b. **Pull loyalty events** (added Sep 2026 — replaces loyalty screenshots for the monthly). For the same Sun–Sat window, call `loyalty.searchEvents` per cafe-day (or per cafe-week for Studio/NMWA volume) with `limit: 30` and follow the `cursor` until absent:
   ```
   {"query":{"filter":{"date_time_filter":{"created_at":{"start_at":"<day>T00:00:00-04:00","end_at":"<day+1>T00:00:00-04:00"}},
                        "location_filter":{"location_ids":["<LOCATION_ID>"]}}},"limit":30,"cursor":"<from prev page>"}
   ```
   Location IDs: Downtown `LXM7V1QWDHKZJ`, Studio `LGVKCKX9B9DZT`, 1899 `L4ZBMGCWJ5C7R`, NMWA `LV2D18H0DD9ES`. `limit` is hard-capped at 30 (higher → `VALUE_TOO_HIGH`). Expect ~3 pages/cafe-day at Downtown, ~35 pages/week total.

   **Do not copy or re-type the pages anywhere.** The harness writes every tool result verbatim into the session transcript on disk, so once you have paginated to the end of each cafe the data is already saved. Harvest it with:
   ```
   python3 ../shared/scripts/loyalty_from_transcript.py --since <START> --until <END>
   ```
   That auto-finds this session's transcript, dedupes on event id, appends to `../shared/loyalty_events/YYYY-MM.jsonl`, and prints a per-cafe summary + store coverage. Idempotent — safe to re-run. Use `--dry-run` to preview. Fallback only if it can't find a transcript: write each page verbatim to `/tmp/loyalty_pages/<cafe>_p<N>.json` and run `python3 ../shared/scripts/loyalty_store.py ingest-dir /tmp/loyalty_pages`.

   NMWA has returned zero loyalty events since it opened (empty response, not an error) — expected; a zero at Downtown/Studio/1899 is a pull failure. The monthly's Section 6 then comes from `monthly/scripts/analyze_loyalty.py YYYY-MM` — no screenshots needed.

2. **Extract the labor response**. Call `labor.searchTimecards` for the full Sun–Sat window across all 4 locations. Save the tool response JSON to `raw/labor_week.json`.
3. **Edit WEEK constants** at the top of the four scripts (same values across all):
   - `aggregate_week.py`: `WEEK_START`, `WEEK_END` (datetimes, EDT)
   - `aggregate_labor.py`: `WEEK_START`, `WEEK_END`, `WEEK_LABEL` (e.g. `"apr19-25"`)
   - `build_weekly_data.py`: `WEEK_LABEL`, `DAYS` (list of 7 ISO dates), `WEEK_START_LABEL`, `WEEK_END_LABEL`
   - `build_weekly_report.py`: `WEEK_LABEL`, `WEEK_TITLE`, `WEEK_FILE_TAG`
4. **Run the pipeline** (from this folder):
   ```bash
   cd _pipeline/weekly/scripts
   python3 aggregate_week.py        # writes data/week_<label>.json
   python3 aggregate_labor.py       # writes data/labor_<label>.json
   python3 build_weekly_data.py     # writes/updates Slipstream Weekly Data.xlsx
   python3 build_weekly_report.py   # writes Slipstream Weekly Report - <dates> v0.1.docx
   ```
5. **Recalc xlsx formulas** (verifies no #REF/#DIV-0):
   ```bash
   python3 /sessions/<your-session>/mnt/.claude/skills/xlsx/scripts/recalc.py \
     "/sessions/<your-session>/mnt/Weekly and Monthly Reports/Slipstream Weekly Data.xlsx"
   ```

## Rules baked in

- **Core Net denominator**: every ratio (Tip %, Labor %, COGS %, Prime %) and the Budget-vs-Actual view is computed off Core Net Sales — i.e., ex-catering and ex-invoice-payments. Total Net is shown alongside for reference but is not the denominator. CEO spec (Apr 2026).
- **History tab is idempotent**: re-running a week overwrites the matching rows keyed on `(week_start, cafe)`. Safe to re-run.
- **Catering/invoice items** are defined in `CATERING_ITEMS` inside `aggregate_week.py`. Keep that list authoritative — if new catering SKUs appear, add them there.
- **FOH/BOH split**: `aggregate_labor.py`'s `ROLE_BUCKET` maps Square role titles to FOH (customer-facing) vs BOH (production/prep/commissary). Weekly report uses FOH labor only per v1 spec.
- **NMWA pre-open/gap handling**: NMWA Square feed has been sporadic through April 2026. If NMWA rows are empty or near-empty, don't force narrative — call it out as a data-collection follow-up.
- **Projections**: `projections.py` auto-picks the latest `Projections/2026 Projections v*.xlsx` in the workspace. Pulls per-cafe per-day sales (col G) and labor $ (col H) from each cafe's tab. Missing days → 0. Shared across the weekly-build pipeline (data + report) and across chats.

## Portable paths

`_paths.py` resolves everything from its own `__file__` — so the scripts work unchanged across chats (each chat has a different session slug, but all chats in this project share this workspace folder). Do not hardcode session paths.

## Projections (live)

Source: `Projections/2026 Projections v*.xlsx` (shared workspace folder, updated by Ryan frequently). `projections.py` auto-picks the highest v-number so future drops (v1.2, v1.3, ...) require no pipeline edits. When a cafe has no projection for a given day (closed, pre-open, unfilled row), loader returns 0 and the report flags the gap in the caveats section.

`load_forward_week_targets(...)` powers the "Forward Targets — Next 4 Weeks" section of the report: aggregated sales + labor $ per cafe per week, for the scheduling manager to pre-plan beyond Square Teams' 2-week horizon.

## Versioning

Bump `VERSION` in `build_weekly_report.py` when the spec materially changes:
- `v0.1` — initial build (dummy projections).
- `v1.0` — projections wired from the Projections workbook; 4-week forward-targets table added.
- `v1.1` — language aligned with Cafe Finance 101 handout v4 (Core Sales / Labor (FOH) / COGS / Prime Cost / Prime Margin); "The Formula" header box; Targets Scorecard (Labor % and Prime % vs handout bands, ✓/△/✗); cross-sectional tables (Cafe Financials, Weekly Sales by Category, Ops Metrics) flipped to metrics-as-rows × cafes-as-cols to match Monthly Report v1.3 (current).

## Sharing (Ryan, Nov 8 2026)

After Ryan approves a week's report, share it with the ALLs through the Asana **ALL Portfolio** (gid 1218029309936191): upload the docx to Google Drive, give the ALLs view access, and post a portfolio status update with the week's headline and the link. Never share before Ryan's OK.

The Admin/Training "Outside core" table is hidden for now: `SHOW_OUTSIDE_CORE = False` in `build_weekly_report.py`.
How (first done Nov 8 2026): the docx is too big to pass through the Drive connector as base64, so convert it to simple HTML (python-docx: headings, lists, tables) and upload it with Drive create_file (text/html → Google Doc) into Drive folder "Slipstream Weekly Reports" (id 1KBE76KkH4tDFJZYR-W2nr5OeCjmCKbf7). Share as reader with each ALL Portfolio member's Asana email, then post a portfolio status update (headline, per-cafe status, actions, link).

## PDF output (Ryan, Nov 15 2026)

The report goes out as a PDF, not a Google Doc. After `build_weekly_report.py`, run
`python3 scripts/build_weekly_pdf.py "../../Slipstream Weekly Report - <dates> v1.2.docx"`.
It lays out the docx's numbers and narrative as a clean PDF (KPI strip, scorecard with colored status, wins/watch-outs, then the tables); the docx stays the source.
Share the PDF (Drive upload by file reference, ALLs as readers) plus the ALL Portfolio status update, after Ryan approves.

## Ryan's report rules (Nov 17 2026)
- Timecard exceptions are NOT in the shared report. `build_weekly_pdf.py` writes them to a separate Ryan-only file, `Timecards to Fix - <week>.pdf`, next to the report. Keep timecard items out of the narrative watch-outs (any watch-out starting "TIMECARD" is routed to that file automatically).
- Section 4 "Filter Coffee Distribution": hot filter cups by cafe from the Square item descriptions in the transactions store. Cruise Control (house) vs Daily Selection vs single origin ([Bright]/[Fruit]/[Floral]/[Decaf]); Nitro and whole bean excluded. KPI = Beyond Cruise Control % = (Daily Selection + single origin) / total hot filter, same as the old monthly "Premium Hot Filter Share". Shows this week plus a 6-week trend, one cafe per line. No target set yet (Ryan to decide).
- Section 7 (monthly averages vs this week): one cafe per table, stacked, not side by side.
- Prime Cost uses CORE labor + COGS (Ryan, Nov 17 2026), the same basis as the labor scorecard. Admin and training stay on memo lines, outside Prime. Data workbook Labor rows and History Labor % / Prime are core from the Nov 8-14 week on (earlier History rows were total labor).
- 1899 Penn Prime band: healthy to 65%, watch to 67% (Ryan, Nov 17 2026) = labor range top/max + ~25% COGS. Other cafes keep 60/65. PRIME_BANDS in build_weekly_report.py.

## Saturday-night automation (Ryan, Nov 22 2026)
Scheduled task "Slipstream weekly report (Saturday night)" builds the report for the Sun-Sat week ending that night and sends Ryan the PDF plus his timecard file. It does NOT share with the ALLs; Ryan approves in chat, then the report is shared (Drive + ALL Portfolio). Sales come straight from Square orders (bundle script `--from-tenders`, `_pipeline/shared/scripts/slipstream_cloud_delta.py`), labor from searchTimecards, filter coffee from the Reporting API via `scripts/build_filter_mix.py`.
The pipeline lives in GitHub at SlipstreamOS/slipstream-reports (private) and runs entirely in the cloud: see `CLAUDE.md` at the repo root for the runbook and `run/build_week.sh` (data, then report). Projections come from Google Drive ("Slipstream Projections" / "2026 Projections.xlsx", Excel); `projections.py` prefers an unversioned `Projections/2026 Projections.xlsx` when present. The Sunday "weekly Square data pull" task is retired.
