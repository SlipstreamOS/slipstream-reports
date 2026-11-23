# Month Close — Ryan's Checklist (v3, Sept 2026 · first use: August 2026 close)

> Your steps only. Full process: **Month Close Playbook.md** (v2.3). Start each close in a **new chat**: "Let's start [Month] close" + attachments — the `slipstream-month-close` skill does the rest.
> **August 2026 specifics:** 2 pay dates (8/14, 8/28) → ~15 JEs + NMWA fee handling · **first close with the Core/Admin/Training labor split** · Penn 1899 Payroll Clearing residual ($18,034.47 at 7/31) still open.

---

## ⓪ Before the chat (no AI)

- [ ] Export **4 Square Sales Reports** (CSV): Studio, 1899, NMWA, DT
- [ ] **IGF Statement** (xlsx) + **Baldor Usage Reports** (1899, DT, Commissary, Studio — plus NMWA if it ordered)
- [ ] **ADP PDFs**, one per pay date — note each **pay-period date range** (needed for the labor split)
- [ ] **SRP Inventory**: next-month tab populated
- [ ] **Studio Rent Sheet**: close-month column
- [ ] **Month Close Sheet**: Config B1 = close EOM · Transfers tab month column (incl. Tech/Insr)
- [ ] Decide: **AP-NMWA** ($1,471.50) settle or keep accruing · **Penn 1899 clearing** true-up JE dated in-month

## ① Open ONE new chat → type "Let's start [Month] close" → attach files → Food first

- [ ] Send IGF + Baldor files → Claude returns breakdowns + Food Bills posting column (locked classifications reused; only NEW items get flagged)
- [ ] Confirm flagged items → paste column into **Food Bills tab** → run **Make.com** (also Keany/Halperns/Fells Point)
- *(Optional: if food statements arrive days before payroll, do this step in its own short chat and start a fresh one for ②.)*

## ② Same chat → JEs

- [ ] Send sales CSVs + ADP PDFs — nothing else to tell it (pay dates, NMWA fee, "Other" tender, Linen/Catering, sheet status are all derived or standing rules; it only asks for a NMWA catering food cost if there was an event)
- [ ] Claude computes per-location gross/tips (`payroll_inputs.py`) **and the admin/training split per pay period** (`labor_split.py`) → **you type gross/tips into Payroll Detail** → tell Claude
- [ ] Claude builds all JEs + one combined CSV (balanced) → reports each pay date's **1899 share** for your GHC6→GHC3 transfers
- [ ] **Import**: account numbers OFF → Import Journal Entries → verify count → numbers ON
- [ ] Post **DT + Studio rent bills**; record **Commissary rent**

## ③ Bank register + post-close (your pace, no AI)

- [ ] Categorize + reconcile (see **QBO Categorization Guide.md**; NMWA ACH → GHC3 Cash on hand)
- [ ] 1899 NI final → **Profit Pass + Incentive** JEs (Incentive = MIN(5% × sales, NI); loss month flips, no incentive)
- [ ] **Cash move** JE + transfers (incentive GHC6→HQ; clear Inter-co AP-GHC6) · **Penn 1899 clearing** JE

## ④ Verify + model (same chat, or a short new one)

- [ ] Say "books done" → Claude pulls QBO P&L + BS and checks fees / COGS / inter-co / 1899 = $0 / labor split landed at locations
- [ ] Refresh **Coefficient Output 2026.xlsx** in Excel (account numbers ON) → save → tell Claude
- [ ] Claude: `sync_model.py` (new version + C5 bump) → recalc → `validate_model.py` → `ops_metrics_from_txstore.py` → **adds the two labor lines to cafe IS tabs (first time only)** → **projections tie-out** (month J vs QBO walk-in sales; flags feed undercount) → reviews per-location NI with you

---

### Fixed amounts
Mgmt Fee **$5,000** · NMWA Mgmt+Equip **$5,041** ($3,950+$1,091) · DT Rent **$2,112.09** (→$2,217.69 May 2027) · SRP Rent **$1,900** · Linen **$150** · Catering COGS **40% of 1899 catering** · Incentive **5% of 1899 sales, capped at NI**

### Scripts (`_pipeline/monthly/scripts/`)
`payroll_inputs.py <period>` · `labor_split.py <start> <end> --json` · `build_jes.py` (edit CONFIG) · `sync_model.py <inV> <outV> <YYYY-MM>` · `validate_model.py <ver> <YYYY-MM>` · `ops_metrics_from_txstore.py <yr> <mo> <model>`
