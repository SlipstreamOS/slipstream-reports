# Slipstream reports, month close and payroll (cloud runbook)

This repo holds everything that used to live in the "Weekly and Monthly Reports" folder on Ryan's Mac. The repo root has the same layout as that folder, so every script runs from a clone with no dependency on Ryan's computer: the weekly report (below), month close, monthly and owner reports, and payroll (sections at the end).
Four cafes: Downtown (LXM7V1QWDHKZJ), Studio (LGVKCKX9B9DZT), 1899 Penn (L4ZBMGCWJ5C7R), NMWA (LV2D18H0DD9ES).
Report rules, bands and history: `_pipeline/weekly/README.md`. Read it before writing a narrative.

## Hard rules
- Never share the report, post to Asana, email, or give anyone Drive access. Ryan approves first, in chat.
- Timecard problems go only in the Ryan-only `Timecards to Fix` PDF, never in the shared report.
- No em dashes in anything Ryan or the ALLs read.
- If any pull fails or a store looks wrong, stop and tell Ryan what broke. Do not send a report built on partial data.

## Saturday-night run (week START = this past Sunday, END = today)
1. Repo: work in the clone of SlipstreamOS/slipstream-reports (`git pull` first).
2. Square pulls (Square connector, `make_api_request_read`). Times are America/New_York; in winter that is -05:00, in summer -04:00.
   - Orders, COMPLETED: `orders.search` (POST /v2/orders/search) with all four location_ids,
     `query.filter.state_filter.states=["COMPLETED"]`, `query.filter.date_time_filter.closed_at={start_at: START T00:00 local, end_at: END+1 T00:00 local}`,
     `query.sort={sort_field:"CLOSED_AT"}`, `limit 500`. Follow `cursor` to the end.
   - Orders, OPEN: same but `states=["OPEN"]` and `created_at` range, `sort_field "CREATED_AT"`.
   - Labor: `labor.searchTimecards` with `query.filter.location_ids` (all four) and `query.filter.start={start_at, end_at}` over the week, `limit 200`, follow cursor.
   - Filter coffee: Reporting `load` with the exact query in the docstring of `_pipeline/weekly/scripts/build_filter_mix.py`, dateRange [START, END].
   - If a method or argument shape is rejected, use `get_service_info` / `get_type_info`; do not guess.
   - Big pages spill to tool-results files. That is fine: the scripts read the session's transcripts and spilled files directly.
3. Projections: Google Drive `download_file_content` on file id `1_46lqQX8zcsyTOCNaSv7EPs5Gckybh44` ("2026 Projections.xlsx" in folder "Slipstream Projections"). `run/harvest.py` copies the downloaded xlsx to `Projections/2026 Projections.xlsx`. If the download did not spill to a file, save the bytes there yourself.
4. `run/build_week.sh data START` bundles and merges sales (tenders as payments), harvests labor and filter rows, rebuilds the data, and prints a first-draft report so the numbers can be read. Exit code 2 from harvest means labor or filter data is missing: go back to step 2.
5. Write the narrative to `_pipeline/weekly/data/narrative_<label>/` (label like `nov22-nov28`): `headline.txt`, `wins.txt`, `watches.txt`, `caveats.txt`. Each line is a Python string literal ending in a comma (headline lines without the trailing comma, see earlier weeks). Use plain quotes only; never `\"`. Follow the style of the last weeks' files. Timecard items start with `WATCH TIMECARD` (or `TIMECARD`) so the PDF builder routes them to Ryan's file. Use `data/labor_<label>.json` (exceptions, non_cafe_title_shifts) for timecard issues.
6. `run/build_week.sh report START` builds the docx and PDFs and copies them to `reports/START/`.
7. Check the PDF (pdftotext it; numbers match step 4, no em dashes, no timecard lines in the shared PDF).
8. Commit and push everything (`git add -A && git commit -m "Week START to END" && git push`).
9. Upload both PDFs to Ryan's private Drive folder "Weekly Reports - Drafts (Ryan only)" (id 1qqWmd0T9xn791kLMu6bVQHrXD0KOvO7h) with create_file (application/pdf, disableConversionToGoogleType true, base64Content as a file reference sha256:<hash>:<absolute path>). Never share that folder or its files; Ryan's EA reads them there. Do not use the "Slipstream Weekly Reports" folder for drafts: it is shared with an outside editor.
10. Send Ryan both PDFs (SendUserFile) and a short summary: labor % per cafe vs band, anything odd, questions for him (closures, events). Ask for his OK to share.

## After Ryan approves (only then)
Upload the report PDF to Drive folder "Slipstream Weekly Reports" (id 1KBE76KkH4tDFJZYR-W2nr5OeCjmCKbf7), share as reader with the six ALL Portfolio members, and post a status update on the ALL Portfolio (gid 1218029309936191). Details in `_pipeline/weekly/README.md`.

## Layout
- `_pipeline/shared/` sales stores: `transactions/YYYY-MM.csv`, `daily_metrics.csv`, `loyalty_events/`, scripts for bundling (`slipstream_cloud_delta.py --from-tenders`) and merging (`merge_deltas.py`).
- `_pipeline/weekly/` report scripts, labor overrides, per-week data and narratives.
- `Slipstream Weekly Data.xlsx` long-format history workbook (updated each run, committed).
- `deltas/` the week's sales bundle, kept for audit.
- `reports/<week start>/` the PDFs sent to Ryan.

## Month close, monthly reports, payroll (cloud, no Mac)
Moved off Ryan's Mac on 2026-11-22. Do not use device_bash or the Mac folder for any of this; work in the clone and push.

**Setup each session:** `git pull`. LibreOffice (`soffice`), python-docx, python-pptx and openpyxl are in the cloud container. For PDF renders of the owner reports install the fonts first: `mkdir -p ~/.fonts && cp _pipeline/fonts/*.otf ~/.fonts/ && fc-cache -f` (Noyh Black is not in the repo; ask Ryan for the file if a render needs it).

**Month close** (skill `slipstream-month-close`; process of record is `Month Close Playbook.md` at the repo root, read the version notes first):
- Inputs come from Drive folder "Month Close Files" (id 13tCRuMoMcybgZ-z2r1MmCFcWw2nxHtBZ) / `<Month YYYY>`, plus anything Ryan attaches. Save the Square sales CSVs to `_pipeline/monthly/data/sales/<YYYY-MM>/`.
- `python3 _pipeline/monthly/scripts/build_jes.py` (edit its CONFIG block). Outputs: `<Month> Sales JEs/` and `<Month> Sales JEs CSV/` at the repo root. No CLOSE_BASE / CLOSE_UP needed.
- Model: `sync_model.py <v_in> <v_out> <YYYY-MM>`, `recalc_copy.py <ver>` (recalculated copy in `~/rc/`), `validate_model.py <ver> <YYYY-MM> ~/rc`. Only the current model version is on `main`; add the new version, and `git rm` the one before last once the new one validates (old versions stay in history and on `mac-archive`).
- `Coefficient Output 2026.xlsx`: Ryan refreshes it in Excel (Coefficient add-in) and puts the refreshed file in the Drive month folder or the chat. Copy it over the repo copy byte for byte. Never openpyxl-save it.
- Reports: `analyze_patterns.py <YYYY-MM>`, `analyze_loyalty.py <YYYY-MM>`, `extract_fin.py`, `build_monthly_report_v2.py`, `MODEL_XLSX=~/rc/<model> build_1899_owner_report.py <ver> <YYYY-MM>`, `build_nmwa_report.py <YYYY-MM>`.
- Deliver to Ryan with SendUserFile and upload to the Drive month folder. Never share or send to anyone else without his OK in chat.
- Commit and push at the end of every working session, even mid-close, so nothing lives only in a container.

**Payroll** (skill `slipstream-payroll`; mechanics in `Payroll/_pipeline/README.md`): run `Payroll/_pipeline/payroll/run_pull.py`, `v2/populate_v2.py`, `v2/finalize_v2.py` from the clone. Deliverables go in `Payroll/<start>_to_<end>/`, to Ryan by SendUserFile, and to the private Drive folder "Slipstream Payroll (Ryan only)". Payroll files are confidential: never share that folder or this repo.

**Archive:** branch `mac-archive` is a complete copy of the Mac folder as of 2026-11-22 (1,129 files: Consolidated Model v6 to v19, pre-merge backups, raw pulls, old weekly report files). Fetch a file from it with `git fetch --depth 1 origin mac-archive && git show FETCH_HEAD:"<path>" > <file>`. Do not merge it into main.
