# Slipstream Consolidated Model — v7 Audit & v8 Target Design

> **Status:** Draft for Ryan's sign-off (June 5, 2026). Audit complete; design recommended; open decisions at the end.
> **Scope:** Rebuild `Slipstream Consolidated Model 2026 v7.xlsx` → clean `v8`. v7 stays authoritative until cutover.
> **Companion:** This is the design sketch called for in the rebuild bootstrap (sequence step 3). Build does not start until decisions below are locked.

---

## Part 1 — v7 Audit

### 1.1 Tab inventory (22 tabs)

| # | Tab | Size | Role today | v8 disposition |
|---|-----|------|-----------|----------------|
| 0 | Control Panel | empty | Stub — never built | **Repurpose** as scenario/refresh control, or drop |
| 1–16 | `Monthly Paste - X` (8) + `Income Statement - X` (8) | per-loc | Coefficient pull → formula-driven IS | **Keep & clean** (8 locations) |
| 17 | Income Statement - Consolidated | 343×90 | Roll-up + toggle + legacy blocks | **Rework** (strip legacy, keep roll-up + toggle) |
| 18 | Income Statement - New Location | 256×48 | Proforma P&L stub (stale, ends 2026) | **Remove** → Pro-forma template |
| 19 | Budget | 111×23 | New-location construction sources/uses ("Location 5") | **Remove** → Pro-forma Inputs tab |
| 20 | 2023 Capital Raise | 183×18 | Dead use-of-proceeds (Cap Riv / Ballston / Logan) | **Remove** |
| 21 | Monthly Transaction Paste | **461,703×57** | Raw Square txns feeding ops-metric array formulas | **Remove** → Ops Metrics summary from `transactions_store` |

The 8 location/department tabs: Studio, Downtown, 1899, NMWA, Commissary, SRP, Cafe Mgmt, Central.

### 1.2 The per-location Income Statement pattern — the good bones

Each `Income Statement - X` tab is well-built and worth preserving conceptually. Column logic (Studio example):

- **Col C** — line-item labels (full P&L taxonomy: Sales → COGS → Labor → Expenses → Occupancy → NOI → Other → Net Income).
- **Col E** — assumption inputs (driver %s). E.g. coffee-mix 96%, catering COGS 27.5%, FOH labor 8%, general supplies 2.5%, rent 12.75%.
- **Cols F/G** — QBO account-number map (e.g. `40111`, `40112`). G = secondary account folded into the same line.
- **Cols K → AT** — monthly time series. Header dates in row 10; row 7 carries the column-index used by the VLOOKUPs; row 8 = year; row 9 = days-in-month.
- **Cols AW → AZ** — annual rollups (SUMIF by year).
- **Cols BB/BC/BD** — Projections / Actual / Difference snapshot block (budget-vs-actual).

**Three-zone column behavior (confirmed live in v7):**

| Zone | Studio columns | Cell logic | Example |
|------|---------------|-----------|---------|
| **Historical (frozen)** | prior FYs (2024, 2025) | hardcoded values | `AA15 = 32562` |
| **Current (live)** | current-FY YTD (Jan–May 2026) | VLOOKUP into `Monthly Paste` by account # | `AM15 = VLOOKUP($F15, 'Monthly Paste - Studio'!…, AM$7)` |
| **Projection (driver)** | future months (Jun 2026 →) | driver formula | `AN15 = $E15*AN17` (mix × total coffee) |

The projection drivers Ryan described **already exist**:
- Sales: `Total Core Sales = AvgDailyTxns × AvgTicket × Days + Catering` (`AN24 = AN41*AN42*AN9 + AN20`)
- Labor: floor + % of sales (`AN126 = MAX(12000, AN127*(AN24-AN20))`)
- Rent: % of sales (`AN203 = $E$203*AN24`)
- COGS: % of category sales

**Historical depth varies by tab (a feature, per requirements):** Downtown May 2021→; Studio Jan 2024→; 1899/Comm/SRP/Mgmt/Central Jan 2025→; NMWA Jan 2026→. All currently end at annual-2027.

### 1.3 The data-pull layer (`Monthly Paste - X`)

Coefficient drops a QBO P&L into each Paste tab: col B = indented account name with embedded number, col A = `NUMBERVALUE(LEFT(TRIM(B),5))` extracting the account number, cols C…N = the 12 months of the **current year only**, col O = Total. The IS tab VLOOKUPs match IS col F (account #) against Paste col A.

**Issues:** single-year design means each new FY needs a re-point; the index row (row 7) keeps computing for already-frozen years (dead logic); and several rows carry `#REF!` inside the account-map formula (Studio Paste rows 57–60). This is the "Coefficient pull layer adds friction" Ryan flagged.

### 1.4 The transaction-metrics problem (`Monthly Transaction Paste`)

- 461,703 rows × 57 cols of raw Square payment data = **~39 MB**, the bulk of the file.
- Consumed by **~160 array formulas** (Studio 55, 1899 50, NMWA 35, Downtown 20) that each scan all ~466K rows to compute daypart transaction counts, to-go %, and avg tickets (filter by month, time-of-day vs dinner-start, weekday/weekend, dining option, net>0, location).
- This is the single biggest cause of file bloat and slow recalc.

**The same data already lives, canonically, in the weekly pipeline:** `_pipeline/shared/transactions/` (one CSV per month, ~25K rows/cafe) with a clean API in `transactions_store.py` (`counts_and_totals`, `avg_ticket_cents`, `avg_daily_transactions`, `dining_option_mix`). No reason to duplicate it in the model.

### 1.5 The Consolidated tab

- **Toggle already exists:** rows 8–15, an "Include?" flag column (D8:D15 = 1/0) for SRP, Commissary, Downtown, Studio, 1899, NMWA, Management Income, Central Admin.
- **Legacy block 1 — embedded mini-IS for Commissary (rows 250–269) and SRP (271–292).** Redundant now that both have full tabs. Leftover from when they shared a tab.
- **Legacy block 2 — the entire 2023 capital-raise loan-coverage analysis (rows 316–337):** "Loan Payment – Existing Locations", "…Total Request", "Payment Coverage", "Net Income (Consolidated w/New Location)", "Logan Demo Cost", "Cashflow Estimate" — references `Budget!` and the New Location tab. Belongs in the Pro-forma Returns tab, not here.

### 1.6 Label-hygiene findings (clean up in v8)

- SRP tab: `Total Commissary COGS` / `Total Commissary Expenses` — should read SRP/Roasting (tab was repurposed).
- Consolidated SRP block row 279: `Total Commissary COGS` (mislabel under SRP).
- `Staff Medial Expenses` typo → "Medical" (Studio C187, Consolidated C194; also a duplicate "Staff Medical Expenses" line at Studio C197).
- Consolidated `Net Income (ex Shared Labor)` appears twice (rows 309 & 313) — legacy from combined-tab era.
- `#REF!` errors: `Monthly Paste - Studio` rows 57–60; `Income Statement - Studio` K41.
- Evening daypart labeled "after 4pm" on Consolidated (C83) but "after 6pm" on Studio (C74) — reconcile to the dinner-start cell.

### 1.7 External dependencies (confirmed accessible)

| Dependency | What it is | Use in v8 |
|---|---|---|
| `2026 Projections v1.2.xlsx` (Projections/) | Daily, per-cafe (1899/Studio/Downtown/NMWA), **top-line + labor only**, 2026 only. YoY-growth + labor-target assumption cells; holiday/day-of-week matched; frozen "Original Projections" drift snapshot. | Source of monthly top-line for the in-progress year |
| `transactions_store.py` (`_pipeline/shared/`) | Canonical payment-level Square store, by cafe + date range | Feeds the new Ops Metrics summary |
| QBO MCP `profit_loss_generator` | Direct P&L pull per location/class | Replaces Coefficient for the live pull |

---

## Part 2 — v8 Target Design (recommended)

### 2.1 Tab list

**Location/department (8):** Studio · Downtown · 1899 · NMWA · Commissary · SRP · Cafe Mgmt · Central
**Roll-up (1):** Consolidated (with toggles)
**Support (2):** Ops Metrics (precomputed daypart/ticket summary) · Data Pulls (the QBO P&L drops, one block per location — successor to `Monthly Paste`)
**Pro-forma (3):** Inputs & Budget · Income Statement · Returns
**Optional (1):** Assumptions (central) — *see open decision*

**Removed:** Monthly Transaction Paste, Budget, Income Statement – New Location, 2023 Capital Raise.

### 2.2 Column strategy (keep the v7 pattern, formalize it)

`[ Labels | Assumptions | Account map ] [ Historical — frozen values ] [ Current FY — live QBO pull ] [ Projections — driver formulas ] [ Annual rollups ]`

- **Historical:** hardcoded after each FY close (already the v7 behavior). Never re-pulled.
- **Current:** live VLOOKUP against the Data Pulls tab for the open FY.
- **Projections:** monthly through **Dec 2027**, then **annual 2028 + 2029** *(open decision — confirm horizon)*.
- Tabs need not align column-for-column; each starts at its own first month.

### 2.3 Ops Metrics replacement

A compact tab — per cafe × per month: total transactions, to-go transactions, weekday-AM / weekend-AM / evening splits, and avg tickets — generated by a small script calling `transactions_store.py` (needs a minor daypart helper added, splitting on each cafe's dinner-start time). IS tabs VLOOKUP these the same way they VLOOKUP P&L. **Result:** drop the 461K-row tab and all ~160 array formulas; file shrinks to a few MB; recalc is instant; one source of truth for transactions.

### 2.4 Projection driver structure

| Line | Driver | Source |
|---|---|---|
| Top-line sales (current/near year) | Monthly total from Projections workbook | `2026 Projections` (daily → monthly) |
| Top-line sales (beyond Projections coverage) | Avg guests/day × avg ticket, grown | Driver cells (ramp/growth assumptions) |
| Product mix (coffee/bar/food/etc.) | LTM average distribution | Derived in location tab from trailing 12 mo actuals |
| Variable costs (COGS, supplies, merchant fees) | % of (category) sales | Assumption %s |
| Overhead / fixed supplies | Inflation step-up | Annual inflation assumption *(open decision)* |
| Rent / occupancy | Contractual schedule | Known step dates (e.g. DT $2,112.09 → $2,217.69 May 2027) |

### 2.5 Data refresh — REVISED after testing the QBO MCP (June 5, 2026)

**Constraint discovered:** the connected QBO MCP (`profit_loss_quickbooks_account`) returns a **whole-company P&L only — no class/location/department dimension**, and its rolled-up totals are unreliable (top-level net income returns 0; income figures don't reconcile across the payload). It therefore **cannot** reproduce the per-location P&Ls the model needs. The other QBO tools (balance sheet, AR/AP aging, sales-by-product) don't help; the only generic API gateway available is **Square**, not QBO.

**Consequence:** we cannot fully drop Coefficient with today's tools. Coefficient's role — pulling a QBO P&L filtered by class (location) — is the one thing nothing else on hand can do.

**Recommended path (keep the engine, kill the friction):**
- **Keep Coefficient** as the per-location pull into the Data Pulls tab.
- Standardize the Data Pulls tab format (consistent columns, `MMM-YY` string headers, account-number ref column) so the IS-tab VLOOKUPs never break — this removes the friction Ryan actually feels (format drift, the v6 NMWA mismatch, the `#REF!` rows), without removing the data source.
- **Claude automates the manual post-refresh steps** each close (formula-copy prior-month → close-month, and the Paste↔IS net-income tie-out), so Coefficient becomes a one-click refresh for Ryan and nothing else.
- Frozen-history rule still means only the open FY is ever pulled.

**Alternative (only if Ryan wants Coefficient fully gone):** source a QBO connector that exposes the ProfitAndLoss report with a `class`/`department` filter (QBO's API supports it; this MCP just doesn't surface it). To be scoped separately — not assumed for v8.

### 2.6 Consolidated roll-up + toggles

Keep the include-flag mechanism but make it scenario-friendly: a small control block where each location is on/off, plus named presets (e.g. **Operating**, **Operating + Pipeline**, **All incl. Pro-forma**). The Pro-forma column can be toggled into any view. *(UX is an open decision.)*

### 2.7 Pro-forma template (3 tabs) — built after Ryan shares reference examples

- **Inputs & Budget** — sources & uses, construction/FF&E/equipment line items (successor to the Budget tab), operating assumptions, ramp.
- **Income Statement** — pro-forma P&L over the analysis horizon, same taxonomy as the location tabs.
- **Returns** — sources & uses summary, debt waterfall, IRR, equity multiple, sensitivity grid (absorbs the legacy loan-coverage analysis).

---

## Part 3 — Decisions (locked June 5, 2026)

1. **Projection horizon** — ✅ **Monthly through Dec 2027, then annual columns for 2028 & 2029.**
2. **Toggle UX** — ✅ **Simple per-location 1/0 include flags** (no named presets). Keep it simple.
3. **Assumptions location** — ✅ **Central Assumptions tab**; each location tab references it.
4. **Inflation default %** — ✅ **3.0%** annual step-up for overhead/fixed-supply lines.
5. **Product mix & Projections workbook** — ✅ **Do NOT recreate the projections workbook.** Canonical source = `Projections/2026 Projections v1.2.xlsx` (per-cafe daily tabs with a monthly rollup in cols O–T; top-line + labor only, 2026). The model reads monthly top-line from it and carries a **product-mix assumption (from LTM actuals)** on the Assumptions tab to split top-line into categories → COGS.
6. **Coefficient** — ✅ **Keep for now** (Ryan, June 5 2026): keep Coefficient as the per-location pull engine, standardize the Data Pulls tab format, and automate the post-refresh formula-copy + tie-out. Removing it later (via a class-capable QBO connector) is a future exploration, not part of the v8 build.

### Assumptions-tab contents (from locked decisions)
- Global inflation: **3.0%/yr** (overhead, fixed supplies).
- Per-location product-mix %s (coffee bev / retail / bar / food / catering / market) derived from **LTM actuals**.
- Per-location driver assumptions migrated out of col E: COGS %s, labor %/floor, rent %/schedule, variable-expense %s.
- Per-location growth / guests-per-day / avg-ticket assumptions for the post-2026 projection years.

---

## Validation gate (before rollout)

POC on Studio must reproduce v7's Studio Net Income for ≥1 closed month (Apr or May 2026) within rounding, or any delta must be an explained, signed-off methodology change. Then roll the validated pattern to the other 7 tabs.
