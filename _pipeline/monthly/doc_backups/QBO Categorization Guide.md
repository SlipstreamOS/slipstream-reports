# QBO Bank Register — Categorization Quick Guide

> Companion to Month Close Cheat Sheet (step C). Updated July 2026.

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
| Food vendors (IGF/Baldor/Keany/Halperns/Fells Point) | Match to Make.com bills (A/P) |
| DT + Studio rent | Match to rent bills (A/P) |
| SRP rent ($1,900) | Rent Liability (23200) directly — no bill |
| Commissary (Mess Hall) rent | Occupancy:Rent @ Commissary directly — no bill, no JE |
| GHC6 card charge $5,100 (1899 paying mgmt fee) | Management Fee expense @ GHC6 - 1899 (same as prior months) |
| Sales tax remittance | Sales Tax Payable |

## Transfers — record as bank Transfers, never categorize
- Payroll-day 1899 share moves GHC6 → GHC3 (amounts from Claude each close)
- Post-close incentive fee GHC6 → HQ
- Any other inter-company cash moves (the JEs do the inter-co clearing; the register just moves cash bank-to-bank)

## Rules of thumb
- If a Sales JE or Payroll JE already booked the P&L side, the register entry clears a **balance-sheet** account (receivable/clearing/liability) — never a revenue/expense account twice.
- If there's a bill, **match** — don't categorize.
- When unsure whether revenue was already recognized, check the month's Sales JE before categorizing a deposit to income.
