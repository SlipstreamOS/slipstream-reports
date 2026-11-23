# November 2026 Close — Pre-flight (written Nov 7, 2026 after the October report review)

Read at the November kickoff. These are approved by Ryan (Nov 7) and already written into Playbook v2.10.

## New at this close
1. **Payroll accrual by work date.** November's unpaid days at close: 11/29 to 11/30 (period 11/29 to 12/12, paid 12/18). Run `accrual_from_timecards.py 2026-11-30 2026-11-29 "<Nov 11/15-11/28 pay JE xlsx>" 2026-11-15 2026-11-28` and paste L3/L5/L6/L7 + TL values into ACCR. The October accrual reversal (dated 11/1) is the actual 11/6 check, so November has no catch-up swing.
2. **Catering COGS 35%** at Downtown, Studio, NMWA (Playbook Phase 6). New JE `CatCOGS-Nov26`. 1899 keeps 40%.
3. **Event-invoice discounts to 40300** (Playbook Phase 3). Pull discounts on Catering and Event items by location from Square Reporting.
4. **Inventory change** flows into COGS for the first time (v2.6): opening = Oct 31 counts. Note: tea bought centrally sits on Downtown's count ($1,988 at 10/31); if tea moved to other cafes without an invoice, Downtown's count drop will read as Downtown usage. Ask Ryan before the Inv JEs if tea transfers were invoiced.

## Before or at this close, for Ryan
- Import `October Sales JEs/NMWA Event Discount Reclass Sep-Oct26.csv` (moves $1,280.30 Sep and $1,240.48 Oct of NMWA event discounts to Catering and Event Sales). The November model sync then restates NMWA Sep/Oct core sales.
- Sep 2026 `Transfers-Adj-Sep26` (JE 9977): Ryan decided Nov 7 it stands as booked (likely an accumulation of missed items from past months). No reversal.
- Studio's September labor (24.6%) is not explained by the accrual (Sept accrual was within $75 of actual). Trace before reading November.
