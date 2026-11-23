# Slipstream Coffee Month Close: Briefing for Grok

Prepared October 30, 2026 for the Grok bots that will gather close inputs, update the Month Close Sheet, and categorize QuickBooks transactions. Ryan Fleming (ryan@slipstreamdc.com) owns the process. Claude builds the journal entries, syncs the model, and produces the monthly reports. Read this whole document before touching anything.

## 1. The business in one page

Slipstream Coffee is a specialty coffee company in Washington, DC. All books are in one QuickBooks Online company, "Glass House Coffee Three, LLC", and every transaction is tagged with a Location (QuickBooks "Department"). Locations are how the business is run and reported, so a transaction without the right location tag is wrong even if the account is right.

| Location tag in QuickBooks | QBO ID | What it is | Square location |
|---|---|---|---|
| GHC3 - Downtown | 2 | Cafe, 1201 New York Ave NW Suite 260. Open 7a to 4p Mon to Fri. Owned. | Slipstream Downtown, `LXM7V1QWDHKZJ` |
| GHC5 - Studio | 3 | Cafe at Studio Theatre, 1333 P St NW. Open 7 days. Owned. Rent is a percentage of sales. | Slipstream Studio, `LGVKCKX9B9DZT` |
| GHC6 - 1899 | 4 | Cafe at 1899 Pennsylvania Ave NW. Open 7a to 3p Mon to Fri. Managed for the building owner (MRP / 1899 Penn Operating Company) under an operating agreement. Not owned. Has its own bank account (GHC6 Checking x1446). | Slipstream 1899 Penn, `L4ZBMGCWJ5C7R` |
| NMWA | 1000000001 | Cafe inside the National Museum of Women in the Arts, 1250 New York Ave NW. Open 10a to 2p Thu to Sun plus community Wednesdays. Opened July 2026. Same legal entity as GHC3. | Slipstream NMWA, `LV2D18H0DD9ES` |
| Commissary | 6 | Central food production kitchen (703 Edgewood St NE, shared space). Sells food to the cafes. | CR Commissary (Food), `LZCA3EY4GB7BM` |
| SRP | 7 | Slipstream Roasting Project, 2400 T St NE. Roasts coffee and makes concentrate, sells to the cafes. | Slipstream Roasting, `LFXQF0X8CKYF5` |
| Cafe Management | 5 | Where Slipstream's income from managing 1899 lands: $5,000 monthly management fee plus an incentive fee. | none |
| Central Admin | 8 | HQ overhead, salaried management, shared costs. | none |

Inactive Square locations you may see in history: Ballston, Capitol Riverfront, Logan Circle. Ignore them.

Two things about 1899 that explain a lot of entries:

- 1899 always ends every month at $0 net income in the books. Its profit or loss passes to the owner through account 24700 Due to Owner / Pass-through Liability and account 83000 Owner Profit Share / Loss Coverage. Slipstream's earnings from 1899 are only the $5,000 fee and the incentive fee (5% of 1899 sales, capped at 1899 net income), both booked at Cafe Management.
- Costs that Slipstream pays on 1899's behalf (labor, supplies, transfers) run through Inter-company AP - GHC6 and Inter-company AR - GHC3, and Ryan settles them with bank transfers from the 1899 account to the main account.

NMWA is similar but simpler: the museum pays Slipstream $5,041 a month ($3,950 management fee plus $1,091 equipment fee), booked as Management Fees income at the NMWA tag. It is intentional that this income sits at NMWA.

## 2. The close in order

Ryan's rule: QuickBooks first, then Coefficient, then the Consolidated Model, then the reports. Nothing downstream starts until the books are right.

1. Gather inputs (Grok). Sales exports, food statements, payroll PDFs, Square invoices for the Transfers tab. See section 3.
2. Update the Month Close Sheet (Grok). Food Bills tab, Transfers tab, Config B1, Payroll Detail inputs. See section 4.
3. Build the journal entries (Claude). About 15 standard JEs in one CSV, imported by Ryan. See section 5.
4. Post bills and run Make.com (Ryan). Food bills are created in QuickBooks by a Make.com scenario reading the Food Bills tab. Rent bills for Downtown and Studio are posted by hand.
5. Categorize the bank register (Grok, with Ryan reviewing). See section 6. This is the step Ryan most wants help with.
6. Final 1899 entries, cash transfers, payroll clearing (Ryan). Profit pass and incentive fee JEs, then transfers.
7. Verify the books (Claude). P&L and balance sheet checks, inter-company accounts at zero, 1899 at $0.
8. Refresh Coefficient (Ryan), sync the model (Claude), build the reports (Claude).

Process documents of record, all in the Weekly and Monthly Reports folder: `Month Close Playbook.md` (the full choreography, 13 phases), `Month Close Cheat Sheet.md` (Ryan's checklist), `QBO Categorization Guide.md` (the register rules, reproduced in section 6 here).

## 3. Inputs to gather each month

Save every input in Google Drive under **Month Close Files** (https://drive.google.com/drive/folders/13tCRuMoMcybgZ-z2r1MmCFcWw2nxHtBZ) in a subfolder named for the month, for example `October 2026`. The Month Close Sheets workbook lives in the same folder. Name files with the month or pay date so they sort: `Sales Report 2026-10 Downtown.csv`, `IGF Statement 2026-10.xlsx`, `Baldor Usage Report 2026-10 Studio.xlsx`, `Keany Usage Report 2026-10.csv`, `ADP Payroll Details 2026-10-09.pdf`. Claude reads from this folder at kickoff, so a file not saved there is a file that does not exist for the close. The "Where it goes" column below is where Claude files its working copy on Ryan's Mac; you do not need to put anything there.

| Input | Source | Where it goes | Notes |
|---|---|---|---|
| 4 Square Sales Reports (CSV) | Square Dashboard, Reports, Sales Summary with Category Sales, one export per cafe: Studio, 1899 Penn, NMWA, Downtown. Full close month. | `_pipeline/monthly/data/sales/<YYYY-MM>/` as `Sales Report <Month> <Cafe>.csv` | Must include the Category Sales by Source block. Anything in "Other Categories" has to be listed out, never lumped. |
| IGF statement | International Gourmet Foods customer item statistics (xlsx). One file with sections SLIP47 = Studio, SLIP60 = Commissary, SLIP62 = Downtown. | Attach to the close chat | Claude splits each line into Coffee or Food and returns a breakdown. |
| Baldor usage reports | One xlsx per account: 1899, Downtown, Commissary, Studio, and NMWA if it ordered | Attach to the close chat | Same treatment as IGF. |
| ADP payroll PDFs | One per pay date (usually 2, sometimes 3 in a month) | Attach to the close chat | Department 001 = FOH hourly, 003 = Management. Taxes used are employer liabilities only, never employee withholdings. |
| Slipstream Roasting invoices (Square) | One invoice per entity per month, issued by Mo from the Roasting location | Transfers tab, SRP rows | Concentrate = ICED COFFEE (KEG) + CONCENTRATE (KEG) items; everything else is roasted coffee. |
| CR Commissary invoices (Square) | Invoices from the Commissary location; customer = cafe; title "BEV COMMISSARY" = coffee, plain "COMMISSARY" = food | Transfers tab, Commissary rows | |
| Physical cash count | Ryan | Tell Claude | Standing rule: Cash on hand in the books equals physical cash at close; the difference goes to Member's Draw or an expense. |

Three Google Sheets Ryan maintains must be current before the JEs are built. Check them and report blanks rather than filling them yourself unless Ryan says to:

- Month Close Sheets 2025_26 (Drive id `1RuAWIPTODNSwjmhSuuntd7q7llwJ8Ez4TJFRw1_2PN0`)
- SRP - MONTHLY INVENTORY 2026 (Drive id `1zo3PbVqZf0TH0OgeQs_gDirQqRsDnr265GLg9y0JPZo`). The close month's coffee cost is read from the NEXT month's tab, line "Implied [Month] Coffee Cost".
- Slipstream Studio Rent Sheet (Drive id `1qr_q790tGFNmtuvsDSbjKoKosOph1J6S6bAbtD6B2ok`). One column per month. Core Rent = Net Sales x 12.75%.

## 4. Month Close Sheet: what to update

Workbook: Month Close Sheets 2025_26.

- Config tab: B1 = last day of the close month (for example `10/31/26`). Make.com reads this to date the bills. B2 and B3 are the month column index and letter.
- Food Bills tab: vendor x location x category. Row 3 is the header (Vendor, Vendor ID, Location, Location ID, Class, Class ID, Category number, Category). One column per month end. Rows 5 to 10 IGF, 11 to 16 Baldor, 17 to 18 Keany, 19 Halperns, 20 Fells Point, 21 to 22 Baldor 1899, 23 to 24 Baldor NMWA. Paste the Coffee and Food totals from Claude's IGF and Baldor breakdowns. Keany, Halperns and Fells Point come from their own statements.
- Transfers tab: rows 5 to 6 Commissary Food and Coffee Sales, 7 to 14 each location's purchases from Commissary, 15 to 16 SRP Sales (Coffee, Concentrate), 17 to 20 each location's SRP purchases, 25 to 28 Tech adjustments, 29 to 32 Insurance adjustments. Build from the Square invoices described above. Tech and Insurance rows copy forward unless something changed.
- Payroll Detail tab: one row per pay date. Blue cells are inputs: ADP Overall (FOH only, Department 001) gross, CC tips, employer taxes, medical; ADP Management (Department 003) gross, taxes, medical; per-location gross and tips from the Payroll Workbook Location Totals. Standing amounts: Ines $1,000, Commissary driver $300, Mo $700, SRP driver $180, healthcare adjustment $418.97. Formulas do the rest. Claude computes the per-location numbers with `payroll_inputs.py` and the admin/training split with `labor_split.py`; Ryan or Grok types them in.

Vendor IDs used on the Food Bills tab: Baldor 35, Keany 36, IGF 43, Halperns 56, Fells Point 59. NMWA's location ID is 1000000001, not 9; the wrong number breaks the Make.com NMWA bills.

## 5. The journal entries Claude builds (so you know what is already booked)

All dated the last day of the month unless noted. JournalNo is descriptive, at most 21 characters, repeated on every line.

| JE | What it books |
|---|---|
| Sales-Studio, Sales-1899, Sales-NMWA, Sales-DT | Revenue by category, discounts, refunds, sales tax payable, tips payable, and the tender side: card to Credit Card Receivables, cash and check/ACH to GHC3 Cash on hand. The DT JE also carves out any $5,100 1899 management fee card payments ($5,000 to Management Income Clearing at Cafe Management, $100 to merchant fees). The NMWA JE books the $5,041 museum fee from the "Other - Services" line. |
| Transfers | Commissary and SRP sales to the cafes, with the 1899 share through Inter-company AP - GHC6 / AR - GHC3. |
| Adjustment | Tech and insurance allocation across cafes, linen ($150) and 1899 catering COGS (40% of 1899 catering sales, out of Downtown Food Purchases). |
| Payroll (one per pay date) | Wages by location split into FOH Labor, Shared Management Labor (60130) and Training Labor (60145), tips, employer tax allocation, credit to Payroll Clearing Account, and the 1899 share to Penn 1899 Payroll Clearing. |
| Payroll-Accr and Payroll-Rev | Accrual for days after the last pay date, reversed on the 1st. |
| MgmtFee | DR Management Income Clearing / CR Management Fees $5,000 at Cafe Management. |
| SRP-GreenCoffee | DR Green Coffee COGS / CR Green Coffee Inventory at SRP, from the inventory sheet. |
| ST-Rent, DT-Rent, SRP-Rent | Rent expense to Rent Liability. DT $2,112.09 (to $2,217.69 from May 2027). SRP $1,900. Studio from the rent sheet. |
| NN-COGS-Trans (occasional) | NMWA catering event food cost moved from Commissary to NMWA. |

Ryan posts three more by hand after the register is done: the 1899 profit pass, the incentive fee recognition, and the cash move that clears Inter-company AP - GHC6.

What this means for categorization: the P&L side of sales, payroll, rent and vendor food is already booked by the time you see the bank feed. Bank lines clear balance sheet accounts. Booking a deposit to a revenue account or an ADP debit to a wage account double counts.

## 6. Bank register categorization rules

Ryan categorizes the QuickBooks bank feed himself today and considers it the worst part of the close. These are his rules. Every line needs the right Location tag as well as the right account.

Deposits

| Bank line | Category | Location |
|---|---|---|
| Square deposits | GHC3 Credit Card Receivables (GHC5 Credit Card Receivables was folded into GHC3 in September 2026) | The cafe the Square location belongs to |
| NMWA museum ACH payments (fee, closure, invoices) | GHC3 Cash on hand. Never Management Fees; the Sales JE already recognized the income | NMWA |
| Any check or ACH deposit for an invoice paid by "Other" tender in Square | GHC3 Cash on hand | The cafe |
| 1899 owner (MRP) payments into the main account | Mostly management fee invoices already recognized through Management Income Clearing; loss coverage portions go to Due to Owner. Ask before categorizing; these have been double counted before. | Cafe Management |
| Transfers between Slipstream bank accounts | Record as a Transfer, never categorize | |

Payments

| Bank line | Category | Location |
|---|---|---|
| ADP debits (net pay, taxes, fees) | Payroll Clearing Account | Follow the existing bank rule |
| IGF, Baldor, Keany, Halperns, Fells Point | Match to the bill (A/P). Never categorize to an expense; the bills carry the expense. IGF and Baldor each also have a vendor clearing account if the match is not one to one | |
| Downtown and Studio rent | Match to the rent bill | |
| SRP rent $1,900 | Rent Liability (23200) directly, no bill | SRP |
| Commissary (Mess Hall) rent | Occupancy:Rent directly, no bill, no JE | Commissary |
| 1899's $5,100 card charge for the management fee (seen in the GHC6 account) | Management Fee expense | GHC6 - 1899 |
| DC sales tax remittances | Sales Tax Payable | The cafe |
| Everything else (supplies, repairs, software, utilities, Amex payments) | Normal expense accounts by the chart of accounts, tagged to the location that used it. Existing bank rules cover the recurring vendors. | |

Rules of thumb

- If a Sales JE or Payroll JE already booked the P&L side, the register entry clears a balance sheet account (receivable, clearing, liability). Never a revenue or expense account a second time.
- If there is a bill, match, do not categorize.
- When unsure whether revenue was already recognized, check the month's Sales JE first.
- Payroll day transfers from GHC6 to GHC3 (the 1899 share of each payroll) and the post-close incentive fee transfer are bank Transfers. The JEs do the inter-company clearing.
- Flag, do not guess: a new vendor, a round-number deposit that is not Square, anything over $2,000 without an obvious home, anything in the GHC6 account that is not payroll, Amex, sales tax, or a transfer.

Current open balance sheet items to be aware of (as of the September 2026 close): Cash on hand needs a physical count true-up; Rent Liability is negative; old Baldor and IGF vendor clearing balances remain; Payroll Clearing and Penn 1899 Payroll Clearing carry balances Ryan clears himself.

## 7. Standing conventions that trip up newcomers

- Square "Other Categories" is never trusted. List each item. Coffee Drinks and Uncategorized go to Coffee; Grab and Go, Sides and Commissary Food Prep go to Food; Services are fee carve-outs. Anything else gets flagged.
- Average ticket and daily transaction metrics exclude invoice-sourced Square transactions (catering and billed events).
- Walk-in sales means Square net sales excluding invoices; the JLL Downtown rent report uses net sales (after discounts and returns), not gross.
- 1899 catering COGS comes out of Downtown Food Purchases, not Downtown catering COGS.
- Shared Management Labor (60130) tagged to a cafe is admin coverage by the cafe's own leads under the Square job title "Manager"; Training Labor is shifts under the "Training" role. Central Admin's 60130 line is HQ salaries.
- Account numbers are enabled in QuickBooks. For the bulk JE import they must be toggled off (Account and Settings, Advanced, Chart of accounts), then back on.
- The QBO import CSV header is exactly `*JournalNo,*JournalDate,*AccountName,*Debits,*Credits,Description,Name,Currency,Location,Class`, UTF-8 with BOM, CRLF line endings, account names as full colon paths.
- Two files are never to be opened and saved by a script: `Coefficient Output 2026.xlsx` (saving it breaks the Coefficient connection) and the Consolidated Model (Ryan edits it offline; it is patched in place, never rebuilt).

## 8. What Grok should and should not do

Do

- Export and file the four Square sales reports in the Drive folder, named consistently. The daily transactions file is built by Claude's weekly Square pull and is not needed.
- Pull the IGF, Baldor, Keany, Halperns and Fells Point statements and hand IGF and Baldor to Claude for the Coffee/Food split.
- Update Config B1, the Transfers tab (from Square invoices) and the Food Bills tab (from Claude's breakdowns and the other vendor statements).
- Check the SRP Inventory and Studio Rent sheets are populated for the month and report blanks.
- Categorize bank feed lines following section 6, with the right Location on every line, and keep a list of every line you were unsure about for Ryan.
- Match vendor payments to existing bills.

Do not

- Post, edit or delete journal entries. Claude builds them, Ryan imports and posts.
- Create bills by hand for IGF, Baldor, Keany, Halperns or Fells Point. Make.com does that from the Food Bills tab.
- Categorize any deposit to a revenue account or any ADP debit to wages.
- Touch anything dated in a prior closed month, or anything in the Due to Owner, Inter-company, Payroll Clearing or Penn 1899 Payroll Clearing accounts. Those are Ryan's.
- Run Make.com, run Coefficient, or open the Consolidated Model.
- Change bank rules without Ryan's approval.

## 9. Handoffs

To Claude: the four sales CSVs, the transactions CSV, IGF and Baldor files, ADP PDFs, and confirmation that the three Google Sheets are current. Claude answers with the food breakdowns, the payroll inputs to type into Payroll Detail, the combined JE CSV, the 1899 share per pay date for Ryan's transfers, then after import the book verification and model sync.

To Ryan: the list of uncategorized or uncertain bank lines, any new vendors or items, blanks in the three sheets, and anything in the Square "Other Categories" rollup.

From Ryan: the physical cash count, decisions on flagged items, confirmation when categorization is done ("books done"), and the Coefficient refresh.

## 10. Reference

Fixed amounts: 1899 management fee $5,000; 1899 incentive fee 5% of 1899 sales capped at 1899 net income; NMWA fee $5,041 ($3,950 + $1,091); DT rent $2,112.09 ($2,217.69 from May 2027); SRP rent $1,900; linen adjustment $150; 1899 catering COGS 40% of 1899 catering sales; healthcare adjustment $418.97; standing payroll fees Ines $1,000, Commissary driver $300, Mo $700, SRP driver $180; 1899 owner reserve held in the operating account $7,500.

Key accounts: 10100 GHC3 FNB Checking (3911); GHC5 FNB Checking x1569; GHC6 Checking x1446; NMWA Business Checking (3929); GHC3 Cash on hand; GHC3 Credit Card Receivables; Management Income Clearing; Payroll Clearing Account; Penn 1899 Payroll Clearing; 23200 Rent Liability; 23650 Inter-company AP - GHC6; Inter-company AR - GHC3; 24700 Due to Owner / Pass-through Liability; 40710 Management Fees; 40720 Incentive Fees; 60130 Shared Management Labor; 60145 Training Labor; 83000 Owner Profit Share / Loss Coverage. The full chart of accounts (254 accounts) is saved as `Glass House Coffee Three - Chart of Accounts (May 2026).xlsx` in the project folder.

People: Ryan (owner, runs the close); Mo (SRP, issues the Roasting and Commissary invoices, maintains SRP inventory); Claude (JEs, model, reports). NMWA contact for cafe matters is Carin Johnson.
