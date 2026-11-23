# Monthly Report Pipeline (v2 — data-driven, from August 2026; v2.1 rules from September)

Generates the Slipstream Monthly Business Review as an investor-style .docx.

**From 2026-11-22 this runs from the GitHub repo clone (repo root = the old "Weekly and Monthly Reports" folder), not Ryan's Mac.** Recalculate the model with `python3 _pipeline/monthly/scripts/recalc_copy.py <ver>` (writes `~/rc/`), then `extract_fin.py`. See the repo `CLAUDE.md`.

## Outputs (land at the top of `Weekly and Monthly Reports/`)

- `Slipstream Monthly Report - <Month Year> <version>.docx`

## Folder layout

```
_pipeline/monthly/
  scripts/
    build_monthly_report_v2.py   # v2 template — reads the two JSONs below; NARRATIVE block at top
    analyze_patterns.py          # transactions + labor → <yymm>_patterns.json
    analyze_loyalty.py           # loyalty-events store → <yymm>_loyalty.json (replaces screenshots)
    report_style.py              # shared docx formatting helpers (investor tables, callouts)
    build_monthly_report.py      # v1 template (Mar–Jul 2026), kept for reference
  data/
    fin_2026_v18.json            # Jan–Sep financials extracted from the recalculated model
    2609_patterns.json           # Sep sales patterns (daypart / DOW / core labor / basket / lead items)
    2609_loyalty.json            # Sep loyalty from the API ledger
    <mon>26_item_slices.json     # per-month item slices from Square Item Sales CSVs (v1 sections)
```

## Monthly refresh (three commands + narrative)

1. **Recalculate the model.** v14+ ships without cached 2026 values (formula-driven off the `QBO - *` tabs).
   ```
   cp "Slipstream Consolidated Model 2026 vNN.xlsx" /tmp/vNN_calc.xlsx
   python3 ~/.claude/skills/xlsx/scripts/recalc.py /tmp/vNN_calc.xlsx 420
   ```
   Then extract Jan..<month> to `data/fin_2026_vNN.json` (see the extract snippet in the Aug 2026 session; row map below).
2. **Run the pattern + loyalty analyses** for the month (needs `_pipeline/shared/transactions/YYYY-MM.csv` and the weekly `labor_*.json` files):
   ```
   python3 analyze_patterns.py 2026-09
   python3 analyze_loyalty.py 2026-09     # needs loyalty_events/2026-09.jsonl from the weekly pulls
   ```
3. **Edit `NARRATIVE` + config** at the top of `build_monthly_report_v2.py` (MONTH, MONTH_LABEL, VERSION, PREPARED, FIN_JSON, PAT_JSON), then:
   ```
   python3 build_monthly_report_v2.py
   ```
4. Validate (XML well-formed) → LibreOffice PDF → spot-check pages → deliver.

## Sections (v2)

1. **Cafe Financials** — 1a Core P&L (Gross Profit row, target-band callout), 1b Catering, 1c Prime/COGS/Labor trends, 1d Core Sales trend, 1e Actual vs Projected, 1f NMWA opening ramp.
2. **Operating Indicators** — weekday avg ticket, weekday txns/day, weekend metrics (Studio + NMWA).
3. **Sales Patterns & Staffing Alignment** *(new)* — 3a daypart curve + ticket/attach/tip by daypart; 3b day-of-week revenue vs labor hours with alignment gap; 3c basket (ticket distribution, attach proxy, dining mix).
4. **Item-Level Insights** — lead item by daypart, lead-category share, top lead items (from transactions). Square Item Sales CSVs still give the full top/bottom-10 + Filter/Premium Hot Filter views.
5. **Discount Report** — on hold.
6. **Loyalty Report** — from the Square Loyalty API event ledger (`_pipeline/shared/loyalty_events/`), filled by the Sunday weekly pull. Screenshots retired Sep 2026.
7. **Benchmarks & Prioritized Actions** *(new)* — scorecard vs typical specialty-coffee ranges (BENCH dict), then ranked actions with $ impact.
8. **Notes & Watch Items.**
Appendices A (support entities), B (consolidated), C (methodology).

## Model row map (v14+ Income Statement tabs; dates on row 14, 2026 Jan=S … Aug=Z)

| key | row | | key | row |
|---|---|---|---|---|
| Coffee Bev | 25 | | Coffee COGS | 37 |
| Retail Coffee | 26 | | Food COGS | 39 |
| Total Coffee | 27 | | Catering COGS | 41 |
| Food | 29 | | Merch purchases | 42 |
| Catering | 30 | | Total COGS | 43 |
| Discounts / Refunds | 33 / 34 | | Total Labor | 53 |
| Total Core Location Sales (incl. catering) | 35 | | Total Expenses / NOI / NI | 88 / 89 / 96 |
| WD/WE txns | 20 / 21 | | WD/WE ticket | 22 / 23 |

Consolidated: `TOTAL NET INCOME (all entities)` = row 106 (correct rollup from v14 onward — no restatement needed).
Support entities: Commissary rev 25 / cogs 30 / labor 37 / opex 43 / NI 45; SRP 24 / 27 / 34 / 45 / 47; Cafe Mgmt rev 21 / NI 22; Central labor 24 / opex 34 / NI 36.

## v2.1 rules (September 2026, per Ryan)

- **Labor = CORE hours only** everywhere (`core_hours` / `core_cents`; admin = Square title "Manager", training excluded). `analyze_patterns.py` computes SPLH, cost/hr, labor % and DOW alignment on **matched days only** (days with both sales and timecards) so partial coverage can't inflate ratios.
- **Section 2a avg ticket** = walk-in per-transaction average from the transactions store (`walkin_ticket_by_month` in the patterns JSON), not the model's Ops Metrics ticket.
- **Section 6 Loyalty** = `analyze_loyalty.py` → `<yymm>_loyalty.json` (API ledger). Builder reads current + prior month (`LOY_JSON`, `LOY_PRIOR_JSON`).
- **Loaded cost per core hour** shown next to SPLH in 3b and 7a.

## Conventions carried forward

- **Core-denominator rule**: all ratios off Core Sales = Total Core Location Sales − Catering.
- **NMWA** stays out of the "Open Cafes" blend while in ramp; shown in its own column + 1f.
- **1899 framing**: same goals and target bands as every cafe; never narrated as "managed separately".
- **Other Categories** rollup on Square exports is always flagged, never folded into Food.
- **Avg ticket** excludes invoice-sourced transactions.
- **Attach proxy**: share of tickets ≥ calibrated threshold (April full-description month); Jul/Aug exports truncate to first item.
- Investor-report styling: no cell grid, navy rules, shaded key/total rows; Calibri 10.
