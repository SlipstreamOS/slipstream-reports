# QBO Bank Register — Categorization Quick Guide

> Companion to Month Close Playbook Phase 10. Updated Nov 8 2026 (car payment identification; cash on hand note). Claude posts the sure rows through the QuickBooks Labs connector and sends Ryan one review xlsx for the rest.

## Deposits
| Item | Category |
|---|---|
| Square deposits (all locations) | GHC3/GHC5 Credit Card Receivables (revenue already in Sales JEs) |
| NMWA museum ACH payments (mgmt/equip fee, closure, invoices) | **GHC3 Cash on hand** — it's the clearing account: the NMWA Sales JE debits Cash on hand for the ACH tender, the deposit credits it back to zero. **Never Management Fees** — revenue is already recognized in the Sales JE |
| Any "Other" tender deposit (checks/ACH at any location) | **GHC3 Cash on hand** — same clearing pattern (Sales JE debits it, deposit clears it) |
| 1899's $5,100 CC mgmt-fee payments inside DT deposits | Normal CC receivables — the split ($5K Clearing + $100 offset) lives in the DT Sales JE |
| Check deposits | GHC3 Cash on hand (clears the Sales JE "check" line) |

## Payments
| Item | Category |
|---|---|
| ADP debits (net pay, taxes) | Payroll Clearing Account |
| IGF and Baldor payments | **Food Vendor Clearing - IGF / Food Vendor Clearing - Baldor** (no one-by-one bill matching, since Sept 2026 close) |
| Halperns, Fells Point, Keany payments | Match to the food bill (A/P) |
| DT + Studio rent | Match to rent bills (A/P) |
| SRP rent ($1,900) | Rent Liability (23200) directly — no bill |
| Commissary (Mess Hall) rent | Occupancy:Rent @ Commissary directly — no bill, no JE |
| GHC6 card charge $5,100 (1899 paying mgmt fee) | Management Fee expense @ GHC6 - 1899 (same as prior months) |
| Sales tax remittance | Sales Tax Payable |
| Paper payroll checks (5-digit check numbers, 5XXXX) | Payroll Clearing Account |
| ADP service fees (small debits) | Payroll Processing Fees |
| Car payment ($936.93; shows as a JPMorgan Chase external transfer out of GHC3 checking) | Auto @ Central |
| Ceremony Coffee | Repairs & Maintenance (repairs only, no coffee bought there) |
| Keylocksmith | Repairs @ Downtown |
| Microsoft | Downtown |
| Amazon on GHC3 Chase 9876 | SRP for now |
| Grubhub, Andy's, Smash on Mo's card | Employee Meals (company party, Oct 2026) |
| Noissue | Packaging (stickers) for all locations |
| Refund of unused legal fees | Credit Legal & Professional Services |
| Small monthly Chase autopay from GHC5 FNB x1569 (card not in QBO) | General Supplies @ Studio |
| FNB line of credit (loan 47759595) | Paid off Apr 2026. Liability account "FNB Line of Credit": principal to it, interest to Interest Expense @ Central |

## Transfers — record as bank Transfers, never categorize
- Payroll-day 1899 share moves GHC6 → GHC3 (amounts from Claude each close)
- Post-close incentive fee GHC6 → HQ
- Any other inter-company cash moves (the JEs do the inter-co clearing; the register just moves cash bank-to-bank)

## Card ownership (sets the location)
- GHC6 Amex x1009: everything is GHC6 - 1899 (Sp Small Planes, DC Govt, Brandywine and the rest).
- GHC3 cards (Chase 9876 and others): central costs, location by item.
- Mo's Amex x2005: used across locations. Cross-location buys go on Ryan's review sheet. Any 1899 share on a GHC3 card is carried to the 1899 adjustment JE through IC AP-GHC6 / IC AR-GHC3.

## Cash on hand
- No physical cash counts. The balance was written off to zero at the October 2026 close (`Cash-WriteOff-Oct26`). From then on it should only hold Sales JE cash/check/ACH lines waiting for their deposits; anything else left in it at month end goes on the review sheet.

## Rules of thumb
- If a Sales JE or Payroll JE already booked the P&L side, the register entry clears a **balance-sheet** account (receivable/clearing/liability) — never a revenue/expense account twice.
- If there's a bill, **match** — don't categorize. Exception: IGF and Baldor go to their vendor clearing accounts.
- When unsure whether revenue was already recognized, check the month's Sales JE before categorizing a deposit to income.
