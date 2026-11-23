# Slipstream Monthly Close — Canonical System

> **Version**: 2.9 (Nov 5, 2026): reports step. New scripts in `_pipeline/monthly/scripts/`: `add_model_rows_xml.py` (adds IS rows to the model by XML patch: copies a template row, adds it to the total row, asserts only those parts changed; edit its CONFIG ROWS/CELLS) and `extract_fin.py <ver> <YYYY-MM> <recalculated projections xlsx>` (writes `data/fin_2026_v<N>.json` for the monthly report). Every reader of `_pipeline/shared/transactions/` (analyze_patterns, analyze_loyalty, build_nmwa_report) now reads the prior, current and next month files, filters to the close month and dedupes, because Sun-Sat bundles put a month's first days in the prior month's file. Monthly report wording lives in `scripts/nar_<yymm>.txt` (patched into NARRATIVE by `patch_report_<mon>.py`); no em dashes anywhere in the generated report. JLL Downtown report: new file version per month (`Sales Report Downtown <Mon> <YYYY>.xlsx`), XML-patched; Coffee/Food = category Net Sales less Invoice rows, Retail = net, Other = Register Catering & Events + Other Categories net; Services (1899 mgmt fee invoices) excluded.
>
> **Version 2.8** (Nov 2, 2026): **Claude now categorizes the bank feed** through the QBO Labs connector. Process: post every row Claude is sure of (rules in `QBO Categorization Guide.md`), match both halves of transfers and card payments, then send Ryan one xlsx (`<Mon> Sales JEs/<Mon> Bank Feed Review for Ryan.xlsx`) listing the rows left pending plus posted items to confirm, with a blank comment column. No QBO review cards for Ryan. Ryan's comments come back on the sheet and Claude posts them.
> v2.7 (Nov 2, 2026) — **Make.com retired for food bills.** Claude fills the Food Bills tab and posts the food bills directly in QuickBooks through the Intuit QuickBooks Labs connector (company Glass House Coffee Three, realm 9130350482895726): one bill per Food Bills row, vendor + A/P 175 + location + class 554641/554643 + account id from the row, dated the month end, memo naming the breakdown file. Tested Nov 2 with the October IGF Downtown bills (10121 food 79.95, 10122 coffee 342.90). IGF and Baldor payments go to their vendor clearing accounts; Halperns, Fells Point and Keany match to the bill.
> v2.6 (Oct 31, 2026) — **Inventory counts now feed core COGS.** Month-end counts per location (sheets in each location's Checklists and Par Sheets folder; tasks in each location's Finance project in Asana). Valued at last price paid (vendor bills; Commissary and SRP items at Mo's transfer prices). Accounts: 12110 Beverages:Coffee Inventory and 12150 Food Inventory, at the location tag. **October 2026 = baseline only** (Ryan, Oct 31): Oct 31 counts book DR 12110/12150 / CR equity opening balance, so October COGS stays purchases-based and 1899's October NI/incentive are unaffected. **From November**, one Inv-<Loc>-MonYY JE per location books the change (closing minus opening): coffee side to 50110 Coffee Purchases, food side to 50155 Food Purchases (Commissary to 50310); no new accounts. Count transcription + valuation workbook: `<Mon> Sales JEs/<Mon> Inventory Counts v<N>.xlsx`. **Tea is counted and valued in grams** from November 2026 (Ryan, Nov 2); the count sheets still say oz until Ryan's next sheet update, so convert any oz entries to grams (x 28.3495).
> v2.5 (Oct 30, 2026) — Grok bots now gather the close inputs (Square exports, vendor usage reports, ADP PDFs), save them to the Drive folder **Month Close Files / <Month YYYY>** (§1), update the Month Close Sheet, and help categorize the bank register (`QBO Categorization Guide.md`; briefing for the bots: `Slipstream Month Close - Briefing for Grok.md`). Claude reads inputs from the Drive folder at kickoff.
> v2.4 (Sept 3, 2026) — post-August-close. Model v17 = August final (v14→v15 sync, v16 added 60130/60145/82100/50100 rows on cafe IS tabs + Consolidated, SRP 62140, Central 62140/62150/62151/82100; labor drivers D48 set to targets Studio .30 / DT .32 / 1899 .35, admin .03/.03/.02, training .01 — `_pipeline/monthly/scripts/add_model_rows_v16.py`; v17 added Studio 61118 + Ops Metrics). Cloud-session mechanics: run scripts on-device via `CLOSE_BASE="$HOME/mnt" CLOSE_UP="../data/sales/<YYYY-MM>/"`; sales CSVs saved under `_pipeline/monthly/data/sales/`. Cash on hand is reconciled to the physical count each close (difference → Member's Draw or expense). NMWA Location ID = 1000000001. Transactions feed gap is Studio/DT only (~5%, Square Online / Mobile Order orders) — weekly assembler to fix.
> v2.3.1 (Sept 2, 2026) — NMWA QBO Location ID confirmed = 1000000001 (was listed as 9). 
> v2.3 (Sept 2, 2026) — pre-August-close update. NEW: **Core/Admin/Training payroll split** (Phase 7 — `labor_split.py`, admin → 60130 @ location, training → new "Training Labor" account), **two labor lines on cafe IS tabs** + **projections tie-out** (Phase 12), close runs in a **fresh chat per month** via the saved skill `slipstream-month-close` + `Bootstrap Prompt — Month Close.md` (§7). NMWA fee: July ran it inside the NMWA Sales JE as ACH "Other" tender (→ GHC3 Cash on hand) — that's the default now. Standing adjustment values: Linen $150, Catering COGS 40% of 1899 catering.
> v2.2 (July 30, 2026) — post-June-close update. Standard JE count is now **16** (added NMWA-MgmtEquipFee — NMWA no longer carved out of Downtown). Payroll JE count varies with pay dates in the month (July 2026 has THREE: 7/3, 7/17, 7/31 → 17 JEs). Model workflow replaced by two-file Coefficient architecture (§Phase 12). Reusable close scripts live in `_pipeline/monthly/scripts/`.
> v2.1 (June 5, 2026) — DT-Rent & SRP-Rent confirmed as Claude-built; count was 15.
> **Owner**: Ryan Fleming, CEO, Slipstream
> **Audience**: Ryan + any Claude model picking up this process
> **Goal**: Anyone — Ryan, a new Claude model, a future operator — should be able to execute a clean monthly close end-to-end using just this document + the linked memory files + the source-of-truth sheets it references.

---

## 0. Quick orientation — read this first

The monthly close turns one month of Slipstream operations into clean financial statements. It takes ~30–45 minutes of focused chat once all inputs are in hand, plus Ryan's bank register categorization (separate workstream, often days).

**Who does what at a glance:**
- **Ryan** gathers source files, runs bank/categorization in QBO, imports JE CSV, posts a few specific JEs himself, runs Coefficient refresh, owns the Monthly Transaction Paste tab.
- **Claude** parses source files, builds the JE CSV (15 standard JEs + 1 occasional), flags carve-outs, validates the model post-refresh.

**The 15 standard JEs Claude builds each close:**
1. Sales-Studio-`MonYY`
2. Sales-1899-`MonYY`
3. Sales-NMWA-`MonYY`
4. Sales-DT-`MonYY`
5. Transfers-`MonYY`
6. Adjustment-`MonYY`
7. Payroll-`MMDD`-`YY` (1st pay date)
8. Payroll-`MMDD`-`YY` (2nd pay date)
9. Payroll-Accr-`MonYY` (last day of close month)
10. Payroll-Rev-`NextMonYY` (1st of following month)
11. MgmtFee-`MonYY`
12. SRP-GreenCoffee-`MonYY`
13. ST-Rent-`MonYY`
14. DT-Rent-`MonYY`
15. SRP-Rent-`MonYY`
16. NMWA-MgmtEquipFee-`MonYY` — **$5,041/mo** (Mgmt $3,950 + Equip $1,091): DR GHC3 Cash on hand / CR 40710 Management Fees, all at NMWA tag. Standing since June 2026 (NMWA books at its own location now — no DT carve-out). Alternative: Ryan categorizes the museum's deposit straight to Management Fees @ NMWA in the bank register, then skip this JE — confirm which each close.

**Payroll count note:** JEs 7–8 assume 2 pay dates; months with 3 pay dates (e.g. July 2026: 7/3, 7/17, 7/31) get a third Payroll JE → 17 total.

**1 occasional JE (when applicable):**
17. NN-COGS-Trans-`MonYY` — when there's a NMWA catering event (~8 months/year, gap June–August).

**3 post-import JEs Ryan posts directly in QBO** (one-by-one, not via CSV):
- "Incent Fee X" — Profit Pass routing (after he reviews final adjusted 1899 NI)
- "Incent Fee X+1" — Incentive Fee income recognition (paired with the above)
- "Cash move X" — clears Inter-company AP-GHC6 / AR-GHC3 after the cash transfer

**Other recurring entries Ryan handles in QBO** (these are *not* part of Claude's CSV):
- Downtown Rent **bill** (Claude builds the paired JE)
- Studio Rent **bill** (Claude builds the paired JE)
- Commissary rent payment only
- Make.com food bills (IGF, Baldor, Keany, Halperns, Fells Point — Ryan triggers Make.com scenario after pasting totals into Month Close Sheet)
- Bank register transaction categorization (ongoing)

---

## 1. Pre-flight checklist — what Ryan provides at kick-off

**Where the inputs live (from the October 2026 close):** the Grok bots save every close input to the Google Drive folder **Month Close Files** (Drive id `13tCRuMoMcybgZ-z2r1MmCFcWw2nxHtBZ`) in a subfolder per month, `<Month YYYY>` (e.g. `October 2026`, id `1GIZISe8qkrKGncc-ZcvZ_Zxac1NMiCSG`). The Month Close Sheets workbook also lives in that folder. Claude reads the inputs from there at kickoff instead of waiting for attachments; Ryan may still attach files directly. Naming seen so far: `ADP Payroll Details <YYYY-MM-DD>.pdf`, `Keany Usage Report <YYYY-MM>.csv`, `Fells Point Usage Report <YYYY-MM>.pdf`.

Ryan sends a single message to Claude to start the close. The inputs needed are:

| # | Item | Format | Notes |
|---|---|---|---|
| 1 | 4× Square Sales Reports | CSV exports | Studio, 1899 Penn, NMWA, Downtown |
| 2 | IGF Statement | `SP Customer_Item Statistics (N).xlsx` | Single file with SLIP47 (Studio) / SLIP60 (Commissary) / SLIP62 (Downtown) sections |
| 3 | 4× Baldor Usage Reports | xlsx | 1899, Downtown, Commissary, Studio |
| 4 | 2× ADP Payroll PDFs | one per pay date | Typically 2 pay dates per close month |
| 5 | Downtown carve-out detail | text in message | See carve-out list below |
| 6 | NMWA catering event? | yes/no in message | If yes: food cost from Commissary records |
| 7 | Confirm SRP - MONTHLY INVENTORY 2026 sheet is updated | "yes" in message | Needed for SRP Green Coffee JE |
| 8 | Confirm Slipstream Studio Rent Sheet is updated | "yes" in message | Needed for Studio Rent JE |

**Downtown carve-out detail (SIMPLIFIED as of June 2026 — NMWA books at its own location now):**
- 1899 Management Fee payment count + total (e.g., `3 × $5,100 = $15,300` — CC payments processed through DT POS; Claude splits into $5K Clearing + $100 Merchant Fees offset per payment). **This is the ONLY remaining carve-out**, and only in months where payments actually flowed through DT (June 2026: none).
- Cash breakdown of the "Other" category (check portion vs other electronic)
- Adjustment JE inputs — now read from Month Close Sheet **Transfers tab** (Tech/Insurance rows) + Ryan confirms Linen + Catering COGS amounts (June 2026: Linen $100, Catering COGS $225, both DT→1899)
- Anything else one-off
- ~~NMWA Mgmt+Equip fee / NMWA catering sales / Loss Pass / $5K wash~~ — retired from carve-out; NMWA fee is now standard JE #16

**Derive, don't ask (from Sept 2026 — kickoff is one line, "Let's start [Month] close"):** items 5–8 above are no longer asked. Claude derives pay dates from the ADP PDFs, checks the three sheets itself (flags blanks), scans the DT report for $5,100 mgmt-fee items, books the NMWA $5,041 fee from its "Other – Services" line inside the NMWA Sales JE (ACH tender → GHC3 Cash on hand), treats any "Other" tender at any location as check/ACH → GHC3 Cash on hand, and applies the adjustment defaults (Linen $150, Catering COGS 40% of 1899 catering, Tech/Insr from the Transfers tab). The ONE thing still asked when it applies: Commissary food cost for a NMWA catering event (detected from the NMWA report). Anything anomalous gets flagged, not guessed.

---

## 2. Accesses Claude needs

A new Claude model starting cold needs working access to:

| System | Why | How to verify |
|---|---|---|
| Google Drive MCP | Read Month Close Sheets, SRP Inventory, Studio Rent, Consolidated Model | Run `search_files` with a known title |
| QBO MCP (Intuit) | Read CoA, A/P aging, P&L, Balance Sheet, Sales reports | Run `company-info`, then any report tool |
| Workspace folder | Read/write JE xlsx, combined CSV, breakdowns | Should be at `/Documents/Claude/Projects/Weekly and Monthly Reports/` |
| Memory files | Standing conventions, prior decisions | Index lives at `MEMORY.md` |

**Note on QBO write limits**: the QBO MCP does NOT expose bill creation, vendor creation, or JE creation tools. All bills and JEs go through Ryan (via Make.com for food bills, manual posting for rent bills, CSV import for the 15 standard JEs). Note the DT and Studio *rent bills* are Ryan's; the paired *rent JEs* (DT-Rent, ST-Rent) and the standalone SRP-Rent JE are Claude's, built into the CSV.

**Key Drive file IDs** (cache here so a new model doesn't have to search):

| File | Drive ID |
|---|---|
| Month Close Files folder (inputs by month, saved by the Grok bots) | `13tCRuMoMcybgZ-z2r1MmCFcWw2nxHtBZ` |
| Month Close Sheets 2025_26 | `1RuAWIPTODNSwjmhSuuntd7q7llwJ8Ez4TJFRw1_2PN0` |
| SRP - MONTHLY INVENTORY 2026 | `1zo3PbVqZf0TH0OgeQs_gDirQqRsDnr265GLg9y0JPZo` |
| Slipstream Studio Rent Sheet | `1qr_q790tGFNmtuvsDSbjKoKosOph1J6S6bAbtD6B2ok` |
| Consolidated Model | local xlsx — `Slipstream Consolidated Model 2026 v<N>.xlsx` in the workspace folder |

---

## 3. Standing conventions (read once, apply every close)

These are factored out into memory files so this doc stays scannable. A new model should read ALL of these on first pickup, then refer back as needed:

| Memory file | Contents |
|---|---|
| `reference_qbo_je_import.md` | CSV format, account-numbers-toggle gotcha, header rules |
| `feedback_je_naming.md` | Use descriptive JournalNo (`Sales-Studio-MonYY`), not integers |
| `reference_qbo_ids.md` | Location IDs (2–9), Vendor IDs (Baldor 35, Keany 36, IGF 43, Halperns 56, Fells Point 59), CoA file location |
| `feedback_other_categories_rule.md` | Always flag the Other Categories rollup on Square reports |
| `feedback_spreadsheet_color_protocol.md` | Blue=input, green=pull, black=formula (hex codes included) |
| `feedback_1899_framing.md` | Don't narrate 1899 as "managed separately" in reports |
| `reference_consolidated_model.md` | Tab structure, conventions |
| `project_consolidated_model_conventions.md` | Mgmt fee/incentive/NMWA/1899 P&L pass conventions |
| `feedback_consolidated_model_validation.md` | NI tie-out: Monthly Paste tab vs Income Statement tab |
| `reference_month_close_workbook.md` | Month Close Sheets tab/column map |
| `reference_payroll_adp_sheet.md` | ADP PDF → sheet mapping (Overall = FOH only) |
| `reference_payroll_pipeline.md` | Payroll automation pipeline, location maps |
| `reference_payroll_tip_source.md` | Daily CC tip source (Square Sales Summary, not CSV) |
| `reference_ryan_recurring_qbo_entries.md` | **Critical**: Rent, Mgmt Fee, SRP Green Coffee, Studio Rent, 1899 Profit Pass + Incentive Fee, post-Incentive cash transfer, Penn 1899 Payroll Clearing settlement, 1899 Mgmt Fee CC handling |

The two most critical ones for a new model starting cold: `reference_ryan_recurring_qbo_entries.md` (everything JE-related) and `project_consolidated_model_conventions.md` (everything model-related).

---

## 4. Phase-by-phase execution

Each phase has: **Owner**, **Trigger**, **Inputs**, **Steps**, **Output**, **Review point**.

### Phase 1 — Source intake & parsing
**Owner**: Claude. **Trigger**: Ryan sends the kick-off message.
**Inputs**: 4 Sales CSVs, IGF xlsx, 4 Baldor xlsx, 2 ADP PDFs, carve-out text.
**Steps**:
1. Read all 4 Sales CSVs. Extract Gross Sales, Returns, Discounts, Tax, Tips, Refunds-by-Amount, Total Collected, Cash/Card/Other/Gift Card splits, per-category gross/refund/discount/net.
2. Apply the "Other Categories" rule (Lunch + Grab-and-Go → Food, Uncategorized → Coffee). Flag anything else.
3. Parse IGF + 4 Baldor xlsx. Apply Coffee/Food classification (flag new items).
4. Parse both ADP PDFs. ADP "Overall" totals = Dept 001 FOH only (ER liability for taxes; deductions for medical). Mgmt = Dept 003.
**Output**: parsed numbers held in memory; breakdown xlsx files for IGF and Baldor saved to `/May Sales JEs/` (or month equivalent).
**Review**: any unrecognized categories, any IGF/Baldor items not previously mapped — flag to Ryan in Phase 2.

### Phase 2 — Clarifying ask
**Owner**: Claude → Ryan. **Trigger**: end of Phase 1.
**Steps**: Send one `AskUserQuestion` covering:
- Downtown carve-out specifics not in the kick-off message
- Adjustment JE inputs (Linen, Catering COGS, Tech, Insurance amounts)
- NMWA catering event yes/no + food cost
- Sheets-updated confirmations (SRP Inventory, Studio Rent)
- Anything ambiguous from parsing
**Output**: Ryan's answers captured for downstream phases.

### Phase 3 — Sales JEs (4)
**Owner**: Claude. **Trigger**: Phase 2 answers received. **Date**: last day of month.
**Pattern**: one xlsx per location with 6 cols (ACCOUNT, DEBITS, CREDITS, DESCRIPTION, NAME, LOCATION). Full hierarchical account paths. Revenue accounts CR, expense/offset DR. Each JE balances DR=CR to the penny.
**Downtown-specific carve-out**:
- 1899 Mgmt Fee CC payments: `CR Management Income Clearing N × $5,000` at Cafe Management + `CR Other Operating Expenses:Merchant Fees N × $100` at GHC3-Downtown (where N = number of $5,100 payments).
- NMWA Mgmt + Equipment: CR Sales:Management Income:Management Fees at NMWA tag.
- Cash split: GHC3 Cash on hand (`Cash` portion) + GHC3 Cash on hand (`check` portion of Other) + GHC3 Credit Card Receivables (CC + other electronic, net of fees).
**Output**: 4 xlsx files: `Studio Sales JE.xlsx`, `1899 Penn Sales JE.xlsx`, `NMWA Sales JE.xlsx`, `Downtown Sales JE.xlsx`.
**Review**: each balances DR=CR to the penny.

### Phase 4 — Food Bills inputs
**Owner**: Claude builds breakdowns; Ryan posts via Make.com.
**Steps (Claude)**: produce `IGF Breakdown <MonYY>.xlsx` and `Baldor Breakdown <MonYY>.xlsx` showing Coffee/Food split per vendor per location, and a summary `<Mon> Food Bills Posting.xlsx`.
**Steps (Ryan)**: paste Claude's totals into Month Close Sheet → Food Bills tab. Update Config tab cell B1 to the close-month EOM date (e.g. `5/31/26`). Run Make.com bill scenario → bills auto-created in QBO. Ryan also handles Keany/Halperns/Fells Point separately in the same Food Bills tab.
**Output**: breakdown xlsx (Claude); QBO bills posted (Ryan via Make.com).

### Phase 5 — Transfers JE
**Owner**: Ryan updates Transfers tab in Month Close Sheet → Claude builds JE.
**Output**: `Transfers JE.xlsx`, JournalNo `Transfers-MonYY`. Touches Inter-co AR-GHC3 / Inter-co AP-GHC6 lines for the 1899 share.

### Phase 6 — Adjustment JE
**Owner**: Claude.
**Pattern**: Tech & Insurance allocation across cafes; uses GHC3 ↔ GHC6 AR/AP only (the GHC6 share total).
**Output**: `Adjustment JE.xlsx`, JournalNo `Adjustment-MonYY`.

### Phase 6.25 — NMWA Catering Trans JE (conditional)
**Owner**: Claude, only when there was an event.
**Pattern** (post May 2026 — no labor adjustment needed since NMWA now has its own payroll allocation):
```
CR  50310 Commissary COGS:Commissary Food     = Event food cost  Commissary
DR  50170 Core Location COGS:Catering         = Event food cost  NMWA
```
**Output**: `NN COGS Trans JE.xlsx`, JournalNo `NN-COGS-Trans-MonYY`. Skip June/July/August.

### Phase 6.5 — Recurring Operational JEs (5)
**Owner**: Claude.

**Mgmt Fee** (fixed): DR Management Income Clearing $5,000 / CR 40710 Management Fees $5,000, both Cafe Management. JournalNo `MgmtFee-MonYY`.

**SRP Green Coffee** (variable, pre-flight): pull "Implied [CloseMonth] Coffee Cost" from NEXT-month tab of SRP Inventory sheet. DR 50025 Green Coffee COGS / CR 12201 Green Coffee Inventory, both SRP. JournalNo `SRP-GreenCoffee-MonYY`. **Pre-flight**: confirm Ryan updated the Inventory sheet.

**Green inventory true-up (Ryan rule, Nov 4 2026):** green purchases hit 12201 only when paid, so the sheet's Purchases line often runs ahead of QBO at close. After the month's green purchases are paid (usually 1-2 weeks after close), true up 12201 to the sheet's count for that month-end: difference to 50025 Green Coffee COGS @ SRP, dated the month-end, JournalNo `SRP-GreenTrueUp-MonYY`. Bags and packaging on the sheet are inventory too: if they were expensed to Roasting Paper Goods, reclass them into 12201 (Jul 2026 pattern, `SRP-Bags-Inv-Jul26`, $1,428).

**Studio Rent** (variable, pre-flight): pull from close-month column of Studio Rent Sheet. 6-line JE all at GHC5-Studio per the table in `reference_ryan_recurring_qbo_entries.md`. JournalNo `ST-Rent-MonYY`. Core Rent = `ROUND(Net Sales × 12.75%, 2)`. **Pre-flight**: confirm Ryan updated the Rent sheet.

**Downtown Rent** (fixed): DR 63100 Occupancy:Rent / CR 23200 Rent Liability, both GHC3 - Downtown. `$2,112.09` through April 2027; `$2,217.69` from May 2027 onward. JournalNo `DT-Rent-MonYY`. No pre-flight.

**SRP Rent** (fixed): DR 63100 Occupancy:Rent / CR 23200 Rent Liability, both SRP. `$1,900.00` every month. JournalNo `SRP-Rent-MonYY` (Ryan's existing template uses `Ryan 102 SRP Rent`). No pre-flight.

### Phase 7 — Payroll JEs (4)
**Owner**: Claude. **Source**: Month Close Sheets Payroll Detail tab (Ryan pastes ADP outputs there first).
**Pattern per Pay JE**: DR per-location wages + tips + per-location tax allocations + IC AR-GHC3 (1899 share) // CR Mgmt Tax residual + Healthcare Adj + Payroll Clearing + Penn 1899 Payroll Clearing (1899 share).
**Accrual + Reversal**: simpler (no Tips, no Healthcare Adj, no 1899 IC split — Mgmt Tax residual absorbs healthcare).
**Critical deliverable to Ryan**: report the precise 1899 share for each pay date (= L6 wages + L6 tax + L6 tips). Ryan transfers this exact amount from GHC6 bank → GHC3 bank on payroll day, then books a clearing JE at month-close that sums both pay periods' 1899 shares (see Phase 11.5).
**Output**: 4 xlsx files, JournalNos `Payroll-MMDD-YY`, `Payroll-Accr-MonYY`, `Payroll-Rev-NextMonYY`.

**Core / Admin / Training split (from the August 2026 close):** each regular Payroll JE books cafe wages in three lines per location instead of one:
- `FOH Labor` @ location = ADP location gross − admin − training (core, schedulable labor — what the labor targets measure)
- `Management Labor:Shared Management Labor` (60130) @ **the cafe's location tag** = wages of shifts under Square job title **"Manager"** (ALL admin/coverage blocks)
- `Training Labor` @ location (new account, Ryan created Aug 2026) = wages of shifts under Square role **"Training"**
- Admin/training dollars come from the Square timecard pulls: `python3 _pipeline/monthly/scripts/labor_split.py <period-start> <period-end> --json` per pay period (needs the weekly labor pulls to cover the period — it warns on gaps). Paste the JSON as `SPLIT` in that pay date's dict in `build_jes.py` CONFIG. ADP gross remains the control total; the JE still balances to Payroll Clearing. Accrual/Reversal JEs stay single-line FOH (immaterial, nets out).
- Central Admin's existing 60130 line (HQ salaries) is unchanged — the location-tagged 60130 lines are additive and reconcile via Location tracking.

### Phase 8 — Combined CSV
**Owner**: Claude. Combine all 15 (or 16 with NMWA Catering) JEs into one CSV.
**Format rules**:
- Headers exactly: `*JournalNo,*JournalDate,*AccountName,*Debits,*Credits,Description,Name,Currency,Location,Class`
- `*JournalDate`: only on the FIRST line of each JE
- `*AccountName`: full hierarchical path
- `*Debits` / `*Credits`: integer when whole, decimal when cents
- `Currency`: `USD` on first line of each JE
- `Location`: on every line
- `Name` / `Class`: blank
- Encoding: UTF-8 with BOM, CRLF line endings
- Zero-value placeholder rows filtered out
**Output**: `All <Mon> JEs (names).csv` in `/May Sales JEs CSV/` (or month equivalent).
**Review**: balance check printed per JE, all must show `Δ=+0.00 OK`.

### Phase 9 — QBO import
**Owner**: Ryan.
**Steps**: Settings → Account and Settings → Advanced → Chart of Accounts → uncheck "Enable account numbers" → Save. Then Settings → Import Data → Journal Entries → upload `(names).csv`. Toggle account numbers back ON after success.
**Review**: QBO confirms imported count = 15 (or 16 with NMWA Catering).

### Phase 10 — Bank register categorization
**Owner**: Claude (since Nov 2026), with Ryan reviewing one spreadsheet. Post the sure rows, match transfer and card-payment halves, then send the review xlsx (see v2.8 note). Mo's Amex x2005 cross-location buys (Amazon, Target and similar) always go on the review sheet for Ryan to assign.
**Notes**: NMWA mgmt fees → Cash on Hand. CC deposits for DT → handled via the Phase 3 carve-out (no separate bank register split needed).

### Phase 11 — 1899 Profit Pass + Incentive Fee (2 JEs)
**Owner**: Ryan posts directly in QBO. **Trigger**: after Phase 10 categorization is done AND Ryan has done any post-import line adjustments to QBO so 1899's NI is final.

**JE A "Incent Fee N"** (Profit Pass + AP routing):
```
DR  83000 Owner Profit Share / Loss Coverage   = 1899 NI         GHC6 - 1899
CR  23650 Inter-company AP - GHC6              = Incentive Fee   Cafe Management
CR  24700 Due to Owner / Pass-through Liability = NI − Incentive Cafe Management
```
Where **Incentive Fee = MIN(5% × 1899 Total Sales, 1899 NI)**.

**JE B "Incent Fee N+1"** (Income recognition):
```
DR  23650 Inter-company AP - GHC6              = Incentive Fee   GHC6 - 1899
CR  40720 Sales:Management Income:Incentive Fees = Incentive Fee Cafe Management
```

After both, 1899 P&L NI = $0; Cafe Mgmt recognizes Incentive Fee revenue; Inter-co AP - GHC6 nets to zero.

**Loss months**: JE A sides flip (CR 83000 at GHC6 absorbs the loss; DR 24700 at Cafe Mgmt for owner share); skip JE B (no incentive). Re-confirm structure first time it happens.

### Phase 11.5 — Cash transfers & inter-co cleanup
**Owner**: Ryan, real-world bank actions.

**Transfer 1 — Incentive Fee**: Move the Incentive Fee amount from GHC6 bank → HQ bank. Record as bank transfer in QBO: `DR HQ Bank / CR GHC6 Bank`. No additional clearing JE needed (Phase 11 JEs already cleared Inter-co AP-GHC6).

**Transfer 2 — Inter-co accumulated**: If Inter-co AP-GHC6 has any other balance (from Transfers, Adjustment, prior-month rollover), do a second bank transfer for that amount and book a paired JE:
```
DR  23650 Inter-company AP - GHC6   $X    GHC6 - 1899
CR  Inter-company AR - GHC3         $X    GHC3 - Downtown
```
JournalNo style `Cash move N`.

**Transfer 3 — Penn 1899 Payroll Clearing**: This account accumulates the 1899 share of each pay period (CR at GHC6). Ryan does a same-day bank transfer GHC6 → GHC3 for each pay period as it happens. At month close, book a single clearing JE summing both pay periods:
```
DR  Penn 1899 Payroll Clearing   = Sum of both pay periods' 1899 shares   GHC6 - 1899
CR  Inter-company AR - GHC3      = Sum of both pay periods' 1899 shares   GHC3 - Downtown
```

After all three, both Inter-co AR-GHC3 and Inter-co AP-GHC6 and Penn 1899 Payroll Clearing should be at $0 (or carry a small reconciling residual).

### Phase 12 — Model update (v8+ two-file architecture) + Ops Metrics
**Owner**: Ryan refreshes Coefficient; Claude does the paste + Ops Metrics.

**Model actuals (QBO):**
1. Refresh `Coefficient Output 2026.xlsx` — Coefficient sidebar pulls each cafe/department "Profit and Loss by Month" from QBO into its `QBO - <tab>`. This file is the ONLY place Coefficient lives. **Never put Coefficient in the master model** — any rebuild/resave of the master strips the add-in connection.
2. Paste the 8 QBO tabs from the Output file into the master's `QBO - <Location>` tabs — Claude runs `_pipeline/monthly/scripts/sync_model.py <inVer> <outVer> <YYYY-MM>` (does the paste, col-P acct# helper, C5 bump, and saves v+1 in one step; asserts account numbers are present — Ryan must keep numbers ON in Coefficient settings). Then recalc (LibreOffice headless) + `validate_model.py <ver> <YYYY-MM>` (per-location NI, 1899=$0 assert, full error scan). Name map: `DT`→Downtown, `Café Mgmt`→Cafe Mgmt, `Central Admin`→Central; the rest match.
3. On the `Assumptions` tab, set **"Last closed month" (C5)** to the close month-end date. Every Income Statement, the Consolidated, the actual/projection split, grey shading, and green pull-fonts all roll from this ONE cell — **no formula copying, no projection cleanup.**
4. Save the master as v+1. (Patch the master in place — never regenerate it over Ryan's saved assumptions.)

**Ops Metrics (weekday/weekend drivers — feeds the 2027 projection only; actuals come from QBO):**
5. Ensure the close month's Square **Transactions CSV** is in `_pipeline/shared/transactions/<YYYY-MM>.csv` (the monthly export — this is the source-of-truth daily transaction history read by `transactions_store.py`).
6. Run: `python3 _pipeline/monthly/scripts/ops_metrics_from_txstore.py <year> <month> "Slipstream Consolidated Model 2026 v<N>.xlsx"`
   - It computes, per cafe, weekday (Mon–Fri) and weekend (Sat–Sun) avg daily transactions + avg ticket at **daily resolution** (so a month ending mid-week cuts correctly), and writes all four into the `Ops Metrics` blocks (Studio 6–17, Downtown 20–31, 1899 34–45, NMWA 48–59; cols C–F, matched on month-end date). **All four cafes, automatically — nothing to hand-enter, no cafe skipped.**
   - Do NOT hand-pull per cafe from Square and do NOT aggregate weekly averages — that's what caused Studio to be skipped in the June close.

**Labor lines on the cafe IS tabs (first done at the August 2026 sync):**
7. Each cafe IS tab needs rows for **60130 Shared Management Labor** and **Training Labor** (pull formula pattern from any existing row; see v12's r98–101 additions for the mechanics). Consolidated rolls cafes line-by-line, so add the same rows there. Validate: cafe NI on IS tab == QBO NI for the month (`validate_model.py`) — a gap of exactly the admin+training $ means the rows are missing.

**Projections tie-out (added Sept 2026):**
8. After the sync, compare the close month's **Projections sheet actuals** (col J daily sum per cafe, walk-in ex-invoice) to **QBO walk-in sales** (Sales JE total less invoice-sourced items). Report the variance per cafe. The Square transaction feed is known to run ~5% light vs the Square Sales Summary (July: Studio −$2.2k, DT −$3.9k) — do NOT silently scale daily values; report it, and if the gap exceeds ~3% flag the weekly pipeline (assembler) for diagnosis. The projections workbook of record is `Projections/2026 Projections v<N>.xlsx` — new version per edit, Sun–Sat weeks, no hard-coded dates in summary formulas.

### Phase 13 — Model validation
**Owner**: Claude.
**Steps**:
1. For each per-location Income Statement tab, tie close-month Net Income to the Monthly Paste tab's NI for that location. Mismatches usually signal Coefficient pull formatting issues or IS formula gaps.
2. Verify 1899 NI = $0 (post-Profit Pass).
3. Verify NMWA's $5,041 Mgmt+Equip Fee shows under NMWA tag, not 1899/Slipstream.
4. Verify 1899 $5K Mgmt Fee shows at Cafe Mgmt tag.
5. Flag any "Other Categories" rollups, any new account types not previously mapped.
6. Spot-check operational metrics (avg ticket, daily transactions) for sanity.
**Output**: validation summary to Ryan with green/yellow/red on each location.

---

## 5. End-to-end verification checklist (use to double-check nothing was missed)

Run through this at the end of every close. Each line should be a clear ✓.

**Pre-import (Claude):**
- [ ] All 15 JEs (or 16 with NMWA Catering) built
- [ ] Each JE balances DR = CR to the penny in the combined CSV (printed at build time)
- [ ] All Downtown carve-outs accounted for (1899 mgmt fee, NMWA fees, NMWA catering, Loss Pass, $5K wash if any)
- [ ] 1899 share of each pay period reported to Ryan for same-day bank transfer
- [ ] SRP Green Coffee value pulled from NEXT-month tab of Inventory sheet (not current-month)
- [ ] Studio Rent JE pulled from correct column (close-month) of Studio Rent Sheet

**Import (Ryan):**
- [ ] Account numbers toggled OFF before import
- [ ] CSV imported, QBO shows 15 (or 16) journal entries
- [ ] Account numbers toggled back ON

**Post-import (Ryan):**
- [ ] Bank register categorization complete
- [ ] Final 1899 NI confirmed (after any line-level adjustments)
- [ ] Profit Pass + Incentive Fee JEs posted (2 entries)
- [ ] Cash move JE posted (Inter-co AP-GHC6 / AR-GHC3 cleanup)
- [ ] Penn 1899 Payroll Clearing settled with clearing JE
- [ ] DT Rent bill posted, SRP Rent JE posted, Studio Rent bill posted, Commissary rent payment recorded
- [ ] Food bills run through Make.com

**Post-import (Claude — optional but recommended):**
- [ ] Pull QBO Journal report for close month; diff against combined CSV row-by-row
- [ ] Pull QBO Balance Sheet; verify Inter-co AR-GHC3, Inter-co AP-GHC6, Penn 1899 Payroll Clearing all at $0 (or reconciling residual)
- [ ] Pull QBO P&L; verify 1899 NI = $0, Cafe Mgmt has Mgmt Fee + Incentive Fee revenue

**Model (Ryan + Claude):**
- [ ] Coefficient refresh done
- [ ] Monthly column formulas copied (April → May, May → June, etc.)
- [ ] Monthly Transaction Paste tab updated
- [ ] Model saved as v+1
- [ ] Claude NI tie-out: every per-location IS NI matches Monthly Paste NI (or residual is explained)
- [ ] No new uncategorized accounts surfaced from Coefficient

---

## 6. Known gaps / future-state automations

- **SRP green coffee — add a purchases-to-12201 check each close (Sept 30 2026):** the sheet's "Implied" cost moves whenever Mo re-counts or adds purchases after the close (Aug 2026: July 2,200.17→3,150.78, Aug 5,691.17→4,901.21; true-ups drafted). Each close: (1) pull the implied cost from the next-month tab, (2) confirm the tab's Purchases line equals the green-coffee bills that hit 12201 in QBO for the month (Aug: only 1 of 3 did), (3) compare 12201 balance to the sheet's count total and report the gap.
- **Post-August-2026 cleanup list (Sept 3, 2026):** (a) Q1 2026 payroll was booked on pay dates with no accrual/reversal until April — Jan carries 3 pay dates (incl. the Dec 14–27 period), Feb/Mar carry 2 with no accrual; confirm whether a 12/31/25 accrual existed, else post a one-time Q1 true-up. (b) 1899 January 2026 loss (NOI −1,069.53) was never passed to the owner and a $422 incentive was recognized — 1899 Jan NI sits at −1,491.52; correct with a Jan-dated (or Aug-dated) loss-pass + incentive reversal. (c) DT rent April 4,774 / May 555 — RESOLVED: % of sales true-up, correct as booked. (d) Tax/benefit allocation rate (Payroll Detail C2 = 12%) over-allocates vs actual ER tax+medical ≈10.2% — Central payroll tax runs negative every month since April; Ryan: change to 10% starting 2027. (e) Catering COGS 40% transfer credits DT (50170) for costs DT never incurred — credit Commissary 50310 instead (like NN-COGS-Trans). (f) NMWA unnumbered accounts "Supplies & Materials" (7,402.68 Feb + 1,906.68 Mar) and "Licenses and Fees" (82.50) are STARTUP costs — Ryan looks at them separately, so they stay out of the model on purpose; remind him they exist when reading NMWA Feb/Mar. (g) Inter-co: Ryan's decision Sept 3 — cash moves credit Inter-co AR-GHC3 only; he intends to collapse AR-COMM / AR-SRP / AP-GHC3 / AP-GHC5 / AP-NMWA into AR-GHC3 eventually (one LLC). Don't propose the AR-COMM/AR-SRP split again. (h) ST-Rent JE concentrate line must hit 50110 Coffee Purchases, not the 50100 parent — fix build_jes.py before the September close (Ryan reclassed August's $80 manually). (i) DT labor (Sept 3 finding): Square timecards show DT hours flat at ~150–180/wk Apr→mid-Aug while walk-in sales fell ~35% (≈$10k/wk spring → $6–7k/wk Aug); core labor % 31–35% → 40–59%. Not an allocation issue. QBO's FOH line is wages NET of tips (AJ = L − M) plus accrual timing, so it also inflates when tips drop. (j) GHC5 Cash on Hand deactivated after the Aug fold; Studio cash deposits → GHC3 Cash on hand @ Studio (bank rule).

- ~~Income Statement formula copy / Monthly Transaction Paste~~ — RETIRED with v8 two-file architecture (Assumptions!C5 auto-roll + Ops Metrics from daily store).
- **Penn 1899 Payroll Clearing residual**: **$18,034.47 at 7/31/26** (Ryan's July clearing JE was not reflected / not dated in-month). Reconcile at the August close: pull the account's transaction detail, post the true-up dated in-month, then the standing monthly clearing JE (Phase 11.5 Transfer 3).
- **AP-NMWA** ($1,471.50 at 7/31): decide settle-by-transfer vs keep accruing; document the choice in `project_consolidated_model_conventions.md`.
- **Square transaction feed ~5% light** vs Square Sales Summary (see Phase 12 step 8) — owned by the weekly pipeline; monthly close only reports the variance.
- **Model forecast labor drivers stale** (~26–28% vs actual ~36% loaded): recalibrate on the Assumptions tab at the August sync now that targets are final (Studio 30 / DT 32 / 1899 35 core + admin + training + ~7% load).
- **Catering forward block seeds** on the Projections tab (DT Oct $14.1k / Nov $22.0k) — Ryan to review at the August sync.
- **SRP Inventory sheet label lag**: the "Implied … Coffee Cost" label on each month tab reads one month behind (July tab says "Implied May"); value is correct. Fix labels when convenient.
- **Non-inventory green coffee usage** (e.g. June's bag-for-rent): the implied-cost calc treats ALL inventory decline as roasting COGS. Any sale/barter of green coffee must be flagged at kickoff → Claude books the reclass JE (June pattern: DR Rent Liability / CR Green Coffee COGS).
- **1899 Loss-month structure**: never tested; first time it happens, confirm direction of all lines before posting.
- **Owner-share rebate mechanism**: open workstream — how to credit owner's share back recurringly (likely net against future bills to owner).
- **DT Rent percentage rent**: if DT ever switches from fixed → variable, mirror the Studio Rent structure.

---

## 7. Model portability — how to onboard a new Claude model

A new Claude model with no prior context picks up this process as follows:

1. **Read this document** (`Month Close Playbook.md`) end-to-end first.
2. **Read all memory files listed in Section 3** — start with `reference_ryan_recurring_qbo_entries.md` and `project_consolidated_model_conventions.md` since they carry the most volatile / current-state info.
3. **Verify accesses** — Drive search for one known title, QBO `company-info` call, list of files in `/Documents/Claude/Projects/Weekly and Monthly Reports/`.
4. **Wait for Ryan's kick-off message** (per Section 1 pre-flight). Pause if any of the 8 required items missing.
5. **Execute phase-by-phase** per Section 4.
6. **Run verification checklist** per Section 5.
7. **Update memory + this doc** if anything changed structurally during the close (e.g., new account type, new carve-out, methodology change). Always note the date of the change so future readers can sequence updates.

**Fresh chat every close (from Sept 2026):** Ryan starts each close in a new chat with `Bootstrap Prompt — Month Close.md`, which invokes the saved skill **`slipstream-month-close`**. The skill's startup sequence is steps 1–3 above; the Bootstrap Prompt carries the month's variables (pay-period ranges, NMWA fee handling, adjustments, new accounts). Everything a new chat needs is in: this Playbook, the memory files in §3 (+ `project_labor_targets_framework.md` for the labor split), `_pipeline/monthly/scripts/`, and the prior month's `<Mon> Sales JEs/` folder as the worked example.

**Environment gotchas a new chat must know:**
- The sandbox's `/tmp` is wiped between tool calls in some sessions — keep every script in `_pipeline/…`, never only in `/tmp`. Node scripts need `NODE_PATH=/tmp/node_modules` after `npm install docx` in `/tmp`.
- The workspace mount blocks deleting/overwriting some files (e.g. PDFs) — write renders to `/tmp` and read from there.
- Oversized MCP results (Month Close Sheet, big P&L pulls, SRP sheet) spill to the host's `tool-results/*.txt`; decode the base64 `content` there, or delegate the pull to a subagent.
- Never openpyxl-save `Coefficient Output 2026.xlsx` (strips the add-in). openpyxl saves drop cached values — recalc via LibreOffice headless before reading numbers back.
- `searchTimecards` filter key is `filter.start` (TimeRange), not `start_at`.

**What NOT to assume**:
- Don't assume any Drive file ID is unchanged — verify via `search_files` if a stale reference fails.
- Don't assume any account name / number is unchanged — when in doubt, pull the CoA and verify.
- Don't assume balance-sheet conventions match what's documented — when in doubt, ask Ryan.

---

## 8. Appendix — Quick reference

### Location IDs (QBO Location tracking)
| ID | Location tag in JE LOCATION column |
|---:|---|
| 2 | GHC3 - Downtown |
| 3 | GHC5 - Studio |
| 4 | GHC6 - 1899 |
| 5 | Cafe Management |
| 6 | Commissary |
| 7 | SRP |
| 8 | Central Admin |
| 1000000001 | NMWA *(confirmed Sept 2 2026 — Make.com bills need this number; JE imports use the string "NMWA")* |

### Vendor IDs (food/coffee)
| ID | Vendor |
|---:|---|
| 35 | Baldor |
| 36 | Keany Produce |
| 43 | IGF |
| 56 | Halperns |
| 59 | Fells Point Wholesale Meats |

### Inter-company account routing
| JE | Account |
|---|---|
| Transfers, Adjustment, Cash Move | Inter-company AP - GHC6 ↔ Inter-company AR - GHC3 |
| Payroll | Penn 1899 Payroll Clearing ↔ Inter-company AR - GHC3 |
| Profit Pass + Incentive | Inter-company AP - GHC6 (at both GHC6 and Cafe Mgmt, nets to zero after both JEs) |

### CSV header (exact)
```
*JournalNo,*JournalDate,*AccountName,*Debits,*Credits,Description,Name,Currency,Location,Class
```

### File path conventions
- JE source xlsx → `/Documents/Claude/Projects/Weekly and Monthly Reports/<Mon> Sales JEs/`
- Combined CSV → `/Documents/Claude/Projects/Weekly and Monthly Reports/<Mon> Sales JEs CSV/All <Mon> JEs (names).csv`
- Consolidated Model → `/Documents/Claude/Projects/Weekly and Monthly Reports/Slipstream Consolidated Model 2026 v<N>.xlsx`

### Sheet pull mapping
| What | Source | Where |
|---|---|---|
| Payroll Detail (per-location wages/tips/tax) | Month Close Sheets | Drive id `1RuAWIPTODNSwjmhSuuntd7q7llwJ8Ez4TJFRw1_2PN0`, Payroll Detail tab |
| SRP Green Coffee monthly value | SRP - MONTHLY INVENTORY 2026 | Drive id `1zo3PbVqZf0TH0OgeQs_gDirQqRsDnr265GLg9y0JPZo`, NEXT-month tab |
| Studio Rent value | Slipstream Studio Rent Sheet | Drive id `1qr_q790tGFNmtuvsDSbjKoKosOph1J6S6bAbtD6B2ok`, close-month column |
| Food Bills tab | Month Close Sheets | Same Drive id, Food Bills tab |
| Transfers tab | Month Close Sheets | Same Drive id, Transfers tab |

---

*End of canonical system document. Update this file directly when the process changes; always bump the version + date in the front matter. The detail behind each phase lives in the memory files listed in Section 3 — keep those current too.*
