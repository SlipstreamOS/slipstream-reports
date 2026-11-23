# Slipstream Month Close — JE Conventions (from March 2026 examples)

## Location tags used in JE LOCATION column
- `GHC3 - Downtown` (QB id 2)
- `GHC5 - Studio` (id 3)
- `GHC6 - 1899` (id 4)
- `NMWA` (id 9, TBD-confirm)
- `Commissary` (id 6)
- `SRP` (id 7) — roasting
- `Central Admin` (id 8)
- `Cafe Management` (id 5) — not seen in March JEs

## File naming pattern
- `Downtown Sales JE.xls`
- `Studio Sales JE.xls`
- `1899 Penn Sales JE.xls`
- (need) `NMWA Sales JE.xls`
- `Payroll JE [Pay Date e.g. March 27].xls`
- `Payroll JE Accr [Last Day of Month e.g. March 31].xls`
- `Payroll JE Reversal [First Day of Next Month e.g. April 1].xls` (implied — reverse of accrual)
- `Transfers JE.xls`
- `Adjustment JE [Month].xls`

All JEs dated last day of the month, except payroll-period JEs (pay date) and the accrual reversal (first of next month).

## Column layout in every JE workbook
Single sheet named "File" with columns:
1. ACCOUNT (full path, colon-delimited, e.g. `Sales:Core Location Sales:Food Sales`)
2. DEBITS\n
3. CREDITS\n
4. DESCRIPTION
5. NAME (always blank in samples)
6. LOCATION

## Sales JE pattern (standard cafe — Studio/1899)
Credits (revenue & liabilities collected):
- Sales:Core Location Sales:Food Sales (gross food)
- Sales:Core Location Sales:Catering and Event Sales (catering + service charge)
- Sales:Core Location Sales:Coffee Sales:Retail Coffee Sales
- Sales:Core Location Sales:Coffee Sales:Coffee Beverage Sales
- Sales Tax Payable
- Employee Tips Payable
- Gift Card Liability (sales)

Debits (offsets):
- Sales:Core Location Sales:Discounts (positive number = reduces revenue)
- Sales:Core Location Sales:Coffee Sales:* — Returns lines per category (each with a 0.00 row even if empty)
- Sales:Core Location Sales:Food Sales — Returns
- Sales:Core Location Sales:Bar Sales — Returns + 0.00 placeholder
- Sales:Core Location Sales:Catering and Event Sales — Returns + service charge placeholder
- Other Operating Expenses:Merchant Fees (Square fees)
- Gift Card Liability claimed (debit when redeemed)
- GHC3 Credit Card Receivables (single account used for ALL locations, location tag determines sub-ledger)
- GHC3 Cash on hand (same — single account, tagged by location)
- Sales:Core Location Sales:Refunds and Voids

## Downtown Sales JE — extra complexity (NMWA passthrough)
Catering and Event Sales is split into two lines:
- "Catering and Other" tagged GHC3 - Downtown (Downtown's own catering)
- "NMWA - Catering and Other" tagged NMWA (NMWA catering events run through Downtown's POS)

Discounts split:
- "GHC3 Discounts" tagged GHC3 - Downtown
- "NMWA Discounts" tagged NMWA

NMWA fee lines (tagged NMWA):
- Sales:Management Income:Management Fees — "Mgmt Fee (NMWA)"
- Sales:Management Income:Management Fees — "Equip Fee (NMWA)"

1899 Mgmt fee (tagged GHC3 - Downtown):
- Management Income Clearing — "Mgmt Fee (1899)"

Loss pass (tagged GHC3 - Downtown):
- Due to Owner / Pass-through Liability — "Loss Pass (L)"

## Payroll JE — Regular (per pay date)
Debits — Labor split by location:
- Employee Labor:Salaries & Wages:FOH Labor (one row per cafe: L3, L5, L6, L7=NMWA)
- Employee Labor:Salaries & Wages:Commissary Labor → Commissary
- Operating Expenses:Commissary Delivery → Commissary
- Employee Labor:Salaries & Wages:Roasting Labor → SRP
- Operating Expenses:Roasting Expenses:Roasting Shipping and Delivery → SRP
- Employee Labor:Salaries & Wages:Management Labor:Shared Management Labor → Central Admin

Debits — Tips payout (only on regular, not accrual):
- Employee Tips Payable → all consolidated to GHC3 - Downtown

Debits — Payroll Tax (one row per location bucket: L3, L5, L6, L7=NMWA, Comm, SRP, Mgmt)

Credits:
- Payroll Clearing Account → all to GHC3 - Downtown
- Employee Labor:Health Insurance Expense → GHC3 - Downtown (only on regular when applicable)

Inter-company netting (1899 example):
- Inter-company AR - GHC3 (debit) tagged GHC3 - Downtown
- Penn 1899 Payroll Clearing (credit) tagged GHC6 - 1899

## Payroll JE — Accrual (last day of month)
Same as regular but:
- NO tips payable line
- NO health insurance line
- NO inter-company AR/AP
- Just labor + payroll tax debits → Payroll Clearing Account credit

Reversal (1st of next month) is exact opposite (debits become credits, vice versa).

## Transfers JE (intercompany sales of goods)
Commissary → Cafes:
- Sales:Commissary Sales:Commissary Food Sales (cr) — Commissary
- Sales:Commissary Sales:Commissary Coffee Sales (cr) — Commissary
- COGS:Core Location COGS:Location Food COGS:Food Purchases (dr) — per cafe
- COGS:Core Location COGS:Location Coffee COGS:Coffee Purchases (dr) — per cafe

SRP → Cafes:
- Sales:Roasting Sales:Roasted Coffee Sales (cr) — SRP
- Sales:Roasting Sales:Concentrate Sales (cr) — SRP
- Cost of Goods Sold:Core Location COGS:Location Coffee COGS:Coffee Purchases (dr) — per cafe (description "L# SRP Purchases")
- Cost of Goods Sold:Roasting COGS:Roasting Cost (placeholder dr 0.00) — SRP

Inter-company AR/AP (separate blocks for COMM and SRP):
- Inter-company AR - COMM (dr) — Commissary
- Inter-company AP - GHC3/5/6 (cr) per cafe (matching food+coffee from commissary)
- Inter-company AR - SRP (dr) — SRP
- Inter-company AP - GHC3/5/6 (cr) per cafe (matching SRP coffee purchases)

## Adjustment JE (intercompany splits, allocations)
Recorded last day of month. Pattern: balance one location vs others.
Examples seen:
- POS & Technology, Insurance, Linen/Uniforms — paid centrally (credit GHC3), split as debits to GHC5/GHC6
- Catering & Event COGS — moves $ from GHC3 to GHC6
- Inter-company AR/AP between GHC3 and GHC6 (for shared bills paid out of one location's bank)

## Description naming patterns
- `L3/L5/L6/L7 Adj FOH` — labor by location (L7 = NMWA)
- `L3 Payroll Tax`, etc.
- `L3 Food Purchases from Com`, `L3 Coffee Purchases from Com`, `L3 SRP Purchases`
- `Mgmt Fee (NMWA)`, `Equip Fee (NMWA)`, `Mgmt Fee (1899)`
- `NMWA Discounts`, `GHC3 Discounts`
- `Loss Pass (L)`
- `service charge`, `Returns - Coffee/Food/bar/Retail/Catering and events`
- `gift card liability claimed`, `merchant fees`
- `AR/AP Adj` (transfers JE inter-company lines)
- `Cash`, `check`, `Other` (cash on hand split in sales JE)

## Month Close Sheets 2025_26 — column maps (verified against March JEs)

Workbook id: `1RuAWIPTODNSwjmhSuuntd7q7llwJ8Ez4TJFRw1_2PN0` (Google Drive)

### Config tab
- `Target Close` (B1): last day of month being closed
- `Column Index` (B2) and `Column Letter` (B3): the column in Food Bills/Transfers used as "this month"
- Location Map repeated (Cafe Mgmt=5, Commissary=6, SRP=7, Central=8)

### Food Bills tab (max 22 rows, 34 cols as of now)
- Row 3 = column headers
- Cols A–H: VENDOR, Vendor ID, LOCATION, Location Id, CLASS, Class ID, CATEGORY NUMBER, CATEGORY
- Cols U–Z: month-end dates
  - U = 2025-12-31, V = 2026-01-31, W = 2026-02-28, **X = 2026-03-31** (March), Y = 2026-04-30, Z = 2026-05-31
- Each data row = one vendor × location × category bill; dollar value in the month column = that month's bill total
- Existing locations: GHC3, GHC5, GHC6, Commissary. **NMWA rows still need to be added before April close.**
- make.com reads B1 (Config!Target Close) and the column index to push bills into QuickBooks

### Transfers tab (max 27 rows, 31 cols)
- Row 3 = column headers
- Cols A–F: DESCRIPTION, LOCATION, CLASS, CREDIT/DEBIT, CATEGORY NUMBER, CATEGORY
- Cols Q–W: month-end dates; **U = 2026-03-31**
- Drives the Transfers JE — output JE just pulls each row's category/location/CR-DR with the month column $ amount

### Payroll Detail tab (max 973 rows, 81 cols) — column maps verified vs March JEs
Tax/benefits charge rate (B2) = 12%. Header rows are 1–4. Data starts row 5+ keyed by Payroll Date in col A.

**Inputs (ADP report + Hardcoded assumptions):**
- A: Payroll Date, B: Pay Start, C: Pay End
- ADP block ~D–N: Gross/CC Tips/Taxes/Medical (Overall + management split)
- Locations block: per-location Gross / Tips / Tax Charge for L3, L5, L6, L7=NMWA, Commissary, SRP, Management
- Hardcoded block: Ines Share + Tax (Comm), Mo Share + Tax (SRP), monthly fixed delivery/SRP labor

**Output block — Regular JE (cols AI–BB), 20 cols:**
| Col | Field | JE account |
|---|---|---|
| AI | L3 FOH wages | EmpLab:S&W:FOH Labor / GHC3 |
| AJ | L5 FOH wages | … / GHC5 |
| AK | L6 FOH wages | … / GHC6 |
| AL | L7 (NMWA) FOH wages | … / NMWA |
| AM | Commissary wages | EmpLab:S&W:Commissary Labor / Commissary |
| AN | Commissary delivery | OpEx:Commissary Delivery / Commissary |
| AO | SRP wages | EmpLab:S&W:Roasting Labor / SRP |
| AP | SRP delivery | OpEx:Roasting Exp:Roasting Shipping & Delivery / SRP |
| AQ | Shared Mgmt wages | EmpLab:S&W:Mgmt Labor:Shared Mgmt Labor / Central Admin |
| AR | Tips total | Employee Tips Payable (debit) / GHC3 |
| AS | L3 Payroll Tax | EmpLab:Payroll Tax / GHC3 |
| AT | L5 Payroll Tax | … / GHC5 |
| AU | L6 Payroll Tax | … / GHC6 |
| AV | L7 (NMWA) Payroll Tax | … / NMWA |
| AW | Commissary Payroll Tax | … / Commissary |
| AX | SRP Payroll Tax | … / SRP |
| AY | Mgmt Payroll Tax | … / Central Admin |
| AZ | Healthcare Adj | EmpLab:Health Insurance Expense (credit) / GHC3 |
| BA | Payroll Clearing | Payroll Clearing Account (credit, total) / GHC3 |
| BB | 1899 Inter-co Clearing | Inter-company AR–GHC3 (dr) and Penn 1899 Payroll Clearing (cr) / GHC6 |

JE balance check (March 27 row):
Sum of AI–AQ + AR + sum(AS–AY) − AZ = BA. ✓ Verified $29,463.75.

**Output block — Accrual JE (cols BE–BU), used on the row of the LAST payroll of the month:**
- BD: Days until EOM (e.g. 10 for March = pay-end 3/21 → 3/31). Driver for accrual = (Days/14) × that pay-period's adjusted pay.
| Col | Field |
|---|---|
| BE | L3 wages accrual |
| BF | L5 wages accrual |
| BG | L6 wages accrual |
| BH | L7 (NMWA) wages accrual |
| BI | Commissary wages accrual |
| BJ | Commissary delivery accrual |
| BK | SRP wages accrual |
| BL | SRP delivery accrual |
| BM | Mgmt wages accrual |
| BN | L3 Payroll Tax accrual |
| BO | L5 Payroll Tax accrual |
| BP | L6 Payroll Tax accrual |
| BQ | L7 (NMWA) Payroll Tax accrual |
| BR | Commissary Payroll Tax accrual |
| BS | SRP Payroll Tax accrual |
| BT | Mgmt Payroll Tax accrual |
| BU | Payroll Clearing accrual (credit, total) |

Accrual JE: same accounts as regular but no Tips, no Health Insurance, no inter-co. Reversal = same JE with debits/credits flipped, dated 1st of next month.

## "Other Categories" rule (from Ryan, Apr 2026)
Standard Square category rollup leaks Lunch and Grab-and-Go into "Other Categories". Slipstream's category mapping in Square is being fixed going forward, but for April close I MUST review the "Other Categories" line on each sales report and ask Ryan how to reclassify (typically into Food). Going forward there should be nothing in "Other Categories"; if there is, flag and ask.

This explains the JE Food vs CSV Food discrepancies in March:
- Studio: CSV Food gross 18,611.00 + Other Categories 2,582.25 = **21,193.25 = JE Food** ✓
- 1899: CSV Food gross 1,752.50 + Other Categories 356.25 = **2,108.75 = JE Food** ✓
- Downtown: Other Categories ($23,636) is largely Mgmt Fees / Loss Pass / NMWA carve-outs — needs line-by-line review (which is why we ask each month)

## March 2026 reconciliation notes
- Downtown Sales Report.csv ↔ Downtown Sales JE.xls: amounts tie (Coffee 26,681.25, Food 13,030.50 = 10,272 + 2,758.50? — need to verify Food number; CSV says Gross Food $10,272, JE says $13,030.50. Difference may be sandwiches/grab-n-go reclassed? Worth asking Ryan in April).
- Studio: Coffee 34,380.25, Food 21,193.25 (CSV gross was 18,611 — difference again worth asking).
- 1899: Coffee 11,821, Food 2,108.75 (CSV gross 1,752.50).

The JE Food/Coffee numbers don't match CSV "Category Rollup" — suggests Ryan is using a different report grouping (Square category mapping → JE category mapping is different from the simple Coffee/Food rollup). Need to ask which custom report he runs.
