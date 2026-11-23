#!/usr/bin/env python3
"""Slipstream Monthly Business Review: v2 template (data-driven).

Reads two JSONs produced upstream and renders the full report:
  data/fin_2026_v<N>.json    : financials Jan..<month> from the recalculated
                                 Consolidated Model (see extract in the Aug session)
  data/<yymm>_patterns.json  : transaction-level sales patterns
                                 (analyze_patterns.py)

Everything numeric in the tables comes from those files; the narrative blocks
(exec summary, insights, actions) are written per month at the top of this file
in NARRATIVE so the analysis is reviewable in one place.

Usage:  python3 build_monthly_report_v2.py
"""
import os, sys, json
from docx import Document
from docx.shared import Pt, Inches, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from report_style import (add_heading, add_para, add_bullet, add_caveat, add_callout,
                          build_investor_table, _ensure_zoom_percent, COLOR_MUTED, COLOR_ACCENT,
                          FONT_NAME, _set_run_font, _set_para_spacing)

PIPELINE_DIR = os.path.dirname(HERE)
WORKSPACE = os.path.dirname(os.path.dirname(PIPELINE_DIR))
DATA_DIR = os.path.join(PIPELINE_DIR, "data")

# ------------------------------------------------------------------ config
MONTH = "Oct"
MONTH_LABEL = "October 2026"
VERSION = "v1.0"
PREPARED = "November 5, 2026"
MODEL_VER = "v20"
FIN_JSON = os.path.join(DATA_DIR, "fin_2026_v20.json")
PAT_JSON = os.path.join(DATA_DIR, "2610_patterns.json")
LOY_JSON = os.path.join(DATA_DIR, "2610_loyalty.json")
LOY_PRIOR_JSON = os.path.join(DATA_DIR, "2609_loyalty.json")
DD_JSON = os.path.join(DATA_DIR, "labor_deep_dive_2026-09.json")
OUT = os.path.join(WORKSPACE, f"Slipstream Monthly Report - {MONTH_LABEL} {VERSION}.docx")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"]
_FULL = {"Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April", "May": "May", "Jun": "June", "Jul": "July",
         "Aug": "August", "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December"}
PM = MONTHS[MONTHS.index(MONTH) - 1]                    # prior month key
MONTH_NAME, PM_NAME = _FULL[MONTH], _FULL[PM]
NEXT_NAME = list(_FULL.values())[(list(_FULL).index(MONTH) + 1) % 12]
RANGE = f"Jan → {MONTH} 2026"
CAFES = ["Downtown", "Studio", "1899", "NMWA"]
CAFE_LABEL = {"Downtown": "Downtown", "Studio": "Studio", "1899": "1899 Penn", "NMWA": "NMWA"}
OPEN3 = ["Downtown", "Studio", "1899"]
DAYPARTS = ["Early (6-8a)", "Rush (8-10a)", "Mid-morning (10-12p)", "Lunch (12-2p)", "Afternoon (2-5p)", "Late (5p+)"]
DOW = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# Typical ranges for independent specialty coffee in high-cost US metros.
# Directional benchmarks (operator experience + trade-association norms), not audited citations.
BENCH = {
    "prime_pct":        ("55–60%",   "Slipstream target; 60–65% warning, >65% problem"),
    "labor_pct":        ("28–33%",   "high-wage metro specialty; <28% is excellent"),
    "coffee_cogs_pct":  ("22–27%",   "beans + dairy + syrups, vertically-integrated roaster at low end"),
    "food_cogs_pct":    ("28–35%",   "wholesale-sourced pastry / prepared food"),
    "splh":             ("$60–80",   "sales per CORE labor hour; <$50 signals over-staffing or under-traffic"),
    "cost_hr":          ("$17–20",   "loaded wage + payroll tax + benefits per hour; role mix drives it"),
    "tplh":             ("9–13",     "transactions per labor hour; rush-hour peaks 15+"),
    "avg_ticket":       ("$9–12",    "with food attach; coffee-only sits $6–7"),
    "attach":           ("35–45%",   "share of tickets with a second item; 50%+ is strong"),
    "morning_share":    ("45–60%",   "office-district cafes, 7–10a; neighborhood cafes 35–45%"),
    "tip_rate":         ("10–15%",   "counter-service card tips"),
    "pickup_share":     ("35–55%",   "urban; higher in office districts"),
}


# ------------------------------------------------------------------ helpers
def fmt_money(v, paren=True):
    if v is None:
        return "–"
    if v < 0 and paren:
        return f"(${abs(v):,.0f})"
    return f"${v:,.0f}"


def fmt_pct(n, d, dp=1):
    if not d:
        return "–"
    v = n / d * 100
    return f"({abs(v):.{dp}f}%)" if v < 0 else f"{v:.{dp}f}%"


def pct_str(v, dp=1):
    if v is None:
        return "–"
    return f"({abs(v):.{dp}f}%)" if v < 0 else f"{v:.{dp}f}%"


def delta_pct(cur, prev):
    if not prev:
        return "–"
    v = (cur - prev) / abs(prev) * 100
    return f"({abs(v):.1f}%)" if v < 0 else f"+{v:.1f}%"


def blend(fin, m, cafes=OPEN3):
    keys = ["core", "coffee", "food", "cbev", "cret", "disc_ref", "ccogs", "fcogs", "merch", "cogs", "gp", "labor", "prime", "pm", "cat", "ni"]
    return {k: sum(fin["cafes"][c][m][k] for c in cafes) for k in keys}


# ------------------------------------------------------------------ narrative (edit per month)
NARRATIVE = dict(
    masthead=(
        f"Prepared {PREPARED}  |  {VERSION}: October results from Consolidated Model {MODEL_VER}. Changes this issue: "
        "(1) timecards cover the full month, so every Section 3 ratio is on all 31 days; (2) the transactions read now "
        "includes October 1 to 3, which sit in the September store file (earlier months read slightly short); "
        "(3) inventory counts started: October 31 counts are booked as a $14,767 opening balance (no P&L effect this "
        "month; from November the monthly change in counts flows into COGS); (4) prior months are restated slightly for "
        "late bank-feed categorizations and the FNB line of credit cleanup (Jan to Sep total net income +$830)."
    ),
    exec_para=(
        "October was a strong sales month and an expensive labor month, and the labor half is partly an accounting "
        "echo of September. Total net income across all entities was $21,429 (September restated to $27,788), "
        "bringing 2026 to $132.9K. Open-cafe core sales rose 8.8% to $112.7K, the third-best month of the year, led "
        "by Studio at $56.0K (+18.9%, within $300 of its April high). Prime cost rose from 58.7% to 63.7% blended, "
        "and all of it was labor (27.6% to 32.5%); COGS held at 31.2%. Part of that swing is timing: September's "
        "payroll accrual ran about $3.1K light at Downtown and Studio, and that cost landed in October when the "
        "accrual reversed. Read the two months together and the open cafes ran 30.2% labor on $216K of core sales, "
        "against a ~31% weighted target. The real October issues are narrower: Downtown's coffee COGS jumped to "
        "32.6%, Downtown's week is mis-staffed at both ends (Wednesday short, Monday and Friday heavy), and Studio's "
        "weekend is still under-staffed on its two biggest days. NMWA had its best month: $15.3K net income from "
        "$9.9K of events and $8.2K of management and closure fees, with walk-in sales nearly doubling."
    ),
    exec_bullets=[
        ("Studio carried the month.",
         "Core $47.1K → $56.0K (+18.9%), 14% over its projection and its second-best month of 2026. Weekday "
         "transactions 141 → 159/day and weekend 229 → 275/day; ticket up to $9.50. Sales per core labor hour "
         "crossed $60 ($62)."),
        ("Labor: read September and October together.",
         "Blended labor 27.6% → 32.5%. September's accrual was a proration that came in ~$1.6K light at Downtown and "
         "~$1.5K light at Studio; the October reversal put that cost here. Two-month labor: Downtown 33.5% (target "
         "32), Studio 26.0% (target 30), 1899 34.6% (target 35). From this close the accrual is booked from the "
         "actual next payroll, so this swing should not repeat."),
        ("Downtown: coffee COGS and the shape of the week.",
         "Prime cost 58.7% → 72.9%. Labor explains most of it, but coffee COGS rose from 28.2% to 32.6% of coffee "
         "sales on flat sales. Wednesday ($2,482/day, the best day) is 7.2 points under-staffed while Friday "
         "($1,068) is +4.9 and Monday ($945) +1.9."),
        ("1899 improved on cost.",
         "Prime 70.8% → 66.0%: coffee COGS 31.9% → 24.2% and food COGS 71.5% → 58.9%. Core sales $19.5K (−2%) "
         "were still 11% over projection. Food COGS remains the one number above its band."),
        ("NMWA: best month yet, and the café is growing.",
         "Walk-in core $1.3K → $2.5K and weekday transactions 18 → 30/day. Events ($9.9K) and $8.2K of management "
         "and closure fees drove $15.3K net income. Walk-in is still half of its $5K projection."),
        ("Support entities cost more.",
         "Commissary ($4.6K) on $11.5K of sales as labor rose to $6.9K; Central ($14.6K) including a one-time "
         "$857 cash on hand write-off and higher payroll. SRP near break-even; Cafe Management $6.0K."),
        ("Guests: bigger baskets, steady loyalty.",
         "Walk-in ticket $8.60 → $8.82 at the open cafes and attach 38.6% → 40.4%. Loyalty: 4,758 visits from "
         "1,816 members (2.62 visits each), 663 of them new to the ledger."),
    ],
    cav_1a=("1899 Net Income is a pass-through (owner profit share) and shows $0 at the entity level; its operating result "
            "is in Prime Margin. NMWA ratios sit on a $2.5K walk-in base next to $9.9K of events; see 1f."),
    callout_1a=("Labor drove the whole move this month (32.5% of core vs 27.6%) and includes about $3.1K of September cost "
                "that landed in October through the accrual reversal; COGS was flat at 31.2%."),
    cav_1b=("October catering was $15,213 company-wide: NMWA $9,926 (museum events), Downtown $3,906, Studio $962 and "
            "1899 $419. Catering is excluded from every core ratio above."),
    cav_1c=("Labor % rose 27.6% → 32.5% blended while COGS % held at 31.2% (September 31.1%). September was flattered "
            "by a light accrual and October carries the catch-up; the two-month blend is 30.2%. Prime cost 63.7% is in "
            "the warning band; the two-month blend is 61.3%."),
    cav_1d=("Studio's +18.9% is the story: its second-best month of 2026 after April. Downtown held its September level; "
            "1899 eased 2% on slightly fewer weekday transactions (132 → 129/day)."),
    src_1e=("Source: 2026 Projections v1.10, sales projection column. September to December were reset in v1.9 at the "
            "summer run-rate (Studio −13% YoY), which proved too low; November and December were raised this week "
            "(Studio −1%, Downtown +5%, holiday days set individually). Labor target compliance is in Section 1a."),
    cav_1e=("All three established cafes beat projection: Studio +$6.7K (+13.7%), 1899 +$1.9K (+10.9%), Downtown +$2.2K "
            "(+6.2%). Studio's gap is mostly the projection, not the month: October ran about 2% below last October. "
            "NMWA walk-in was $2.5K against a $5.0K projection."),
    cav_1f=("NMWA's best month: $15.3K net income, built on $9.9K of events and $8.2K of management and closure fees "
            "(booked at NMWA by convention). The café itself grew: walk-in core $1,323 → $2,536 and weekday "
            "transactions 18 → 30/day. P&L labor ($1.8K) is understated by accrual timing; timecards show $3.4K."),
    cav_2a=("Ticket rose at Downtown ($8.57 → $8.81) and Studio ($9.28 → $9.50), the first move up since spring; 1899 "
            "steady at $7.34. October was a traffic month at Studio and a ticket month at Downtown."),
    cav_2c=("Studio weekend 229 → 275 transactions/day, its strongest weekend since April. Weekday traffic: "
            "Downtown 194 → 202, Studio 141 → 159, 1899 132 → 129, NMWA 18 → 30."),
    src_3_coverage=("Timecards cover the full month at every cafe (all open days matched), so these ratios are complete "
                    "this month."),
    cav_3a=("Share of the month's walk-in net sales by daypart. The 7–10a window is 40.5% at Downtown and 34.8% at Studio; "
            "Studio's peak is 10a–12p (32.9%) and 1899's lunch (29.8%) is as large as its rush (30.2%)."),
    insights_3a=(
        "The daypart shapes held for a third month, which confirms they are structural. Downtown is a morning cafe "
        "(36% in the 8–10a rush, 14% after 2p). Studio peaks mid-morning (33%) and carries lunch (24%). 1899 splits "
        "evenly between rush and lunch. NMWA does 85% of its sales between 10a and 2p. Each cafe needs its own "
        "staffing curve, and the tier templates from the deep dive already reflect these shapes."
    ),
    insights_3b=(
        "With a full month of timecards the alignment picture is clean. Studio's weekday gaps closed (Tue/Wed +2.1/+2.6, "
        "from +5.1/+4.4 in September) but the weekend gap did not move: Saturday −5.4 and Sunday −2.5 on its two "
        "$2,650–2,850 days. Downtown flipped: Friday improved (+4.9 from +9.2) but Wednesday, its best day at "
        "$2,482, is now 7.2 points short while Monday ($945) and Friday ($1,068) run heavy. 1899 is close to "
        "aligned except Friday (+4.9) and Thursday (−4.3). Sales per core hour: Studio $62, 1899 $57, Downtown $53."
    ),
    insights_3c=(
        "The basket moved for the first time in months: attach proxy 38.6% → 40.4% blended, walk-in ticket up "
        "$0.24 at Downtown and $0.22 at Studio. Studio's pickup share rose from 49% to 56% with the traffic gain. "
        "The $5–8 single-coffee band is still 40% of all tickets, so the 10a–2p attach prompt remains the cheapest "
        "sales lever."
    ),
    cav_4=("Send the October Square Item Sales CSVs (Downtown / Studio / 1899 / NMWA) and this section gains the full "
           "top/bottom-10 by category, pastry detail and the Filter Coffee view."),
    loyalty_bullets=[
        ("Visits flat, members up.",
         "4,758 loyalty visits (September 4,726) from 1,816 members (1,750). Studio drove it: 2,164 visits (+10%) and "
         "864 members (+12%), matching its traffic gain."),
        ("Frequency eased slightly.",
         "2.62 visits per member (2.70). More new and occasional members in a growth month pull the average down; the "
         "5+ visit core held at 256 members (14%)."),
        ("New members: 663.",
         "Downtown 205, Studio 346, 1899 119, down from 854 as the ledger's history builds (members first seen in "
         "August or September now count as repeat)."),
        ("About 38% of walk-in tickets carry loyalty.",
         "Downtown 39%, Studio 37%, 1899 36%. From October the ticket count is month-exact, so this is the baseline "
         "going forward; September's 41% used a slightly different ticket base."),
    ],
    cav_6b=("Spend per visit ($8.16 / $8.78 / $6.81) sits at each cafe's walk-in ticket, so members do not out-spend "
            "other guests per visit; their value is frequency (2.5–2.6 visits a month)."),
    cav_6c=("Measured directly from the ledger (members who visited in the month). 5-yr = annualized × 5, no retention "
            "decay. A member at today's 2.6 visits a month is worth about $1,280 over five years."),
    cav_6d=("Same buckets every month. The 5+ visit regulars are 14% of members (256); Downtown and 1899 "
            "each keep about 15% of members in that group, Studio 12.5%."),
    cav_7a=("Studio passed $60 sales per core hour ($62) at $20.78 loaded cost per hour. Downtown fell "
            "to $53 at $21.74, the result of Monday and Friday hours, and 1899 is $57 at $22.09, the highest cost per "
            "hour of the three. Average ticket sits below the $9 to $12 range at Downtown ($8.81) and 1899 ($7.34), and transactions per labor hour trail the range at all three."),
    actions=[
        ("Rebalance Downtown's week toward Wednesday.",
         "Wednesday is the best day ($2,482) and 7.2 points short on labor; Monday ($945) and Friday ($1,068) are "
         "+1.9 and +4.9 heavy. Moving one shift from Monday and the Friday third start to Wednesday is labor-neutral "
         "and protects the highest-volume day.",
         "labor-neutral; revenue-protective"),
        ("Find the Downtown coffee COGS jump.",
         "Coffee COGS 28.2% → 32.6% on flat coffee sales is about $1.2K in one month. Check October's SRP, IGF and "
         "Baldor beverage purchases against usage; the November inventory change will show whether it was stock "
         "build or real usage.",
         "$800–1,200 / mo"),
        ("Staff Studio's weekend to its sales.",
         "Saturday −5.4 and Sunday −2.5 on $2,650–2,850 days, with weekend traffic up to 275/day. Move 6–8 hours a "
         "week from Tue–Thu to Sat/Sun. Studio's $62 per hour shows the room.",
         "labor-neutral; throughput on the best days"),
        ("Look at Commissary labor against volume.",
         "Labor $6.9K on $11.5K of sales (60%), up $2.8K on September while sales rose $1.8K. Some is accrual timing; "
         "check production hours against the October invoice volume.",
         "$1,000–2,000 / mo"),
        ("Finish the 1899 Friday template.",
         "Friday is +4.9 at both weekday cafes (better than September's +7 to +9 but still the widest positive gap at "
         "1899).",
         "$300–500 / mo"),
        ("Set NMWA café targets separate from events.",
         "The café nearly doubled but is at half its projection. A transactions-per-day ramp target for November and "
         "December keeps strong event months from hiding the café trend.",
         "decision"),
    ],
    data_caveats=[
        f"Financials from Consolidated Model {MODEL_VER}, recalculated before reading; every location's October net income "
        "ties to QuickBooks to the cent (Cafe Management within $0.01 of rounding). Four accounts with October activity were "
        "added to the Income Statement tabs (Studio Travel & Entertainment, Central Legal Services, Cash Over/Short and "
        "Interest Expense, Commissary Bank Charges). Two #N/A cells in the Pro-forma tabs are not used here.",
        "Prior months restated in QuickBooks: bank-feed catch-up (Studio card charges Jan–Sep), FNB line of credit interest "
        "(Central, Jan–Apr), coffee bags moved from paper goods into SRP inventory (July, +$1,428) and a January 1899 fix. "
        "Jan–Sep total net income moved from $110,610 to $111,440.",
        "Payroll accrual: September's accrual was a proration and ran light (Downtown ~$1.6K, Studio ~$1.5K vs the payroll "
        "bridge); October absorbed it when the accrual reversed. October's own accrual was booked from the actual 11/6 "
        "payroll (10/18–10/31), so November should not carry a swing.",
        "Section 3 labor % (timecard cost / walk-in sales: Downtown 41%, Studio 34%, 1899 39%) is higher than the P&L "
        "because timecards are gross wages, while P&L labor nets out card tips paid through payroll and adds employer taxes. "
        "Use the P&L for the target test and Section 3 for the shape of the week.",
        "Transactions store vs QuickBooks walk-in sales: Studio −1.0%, 1899 −0.5%, NMWA 0.0%, Downtown −3.5% (Square Online "
        "and mobile orders missing from the feed; flagged for the weekly assembler).",
        "Square exports since July truncate multi-item descriptions to the first item: attach is a calibrated proxy (3c) and "
        "Section 4 counts lead items, not units. Section 5 (Discounts) on hold.",
    ],
    watch=[
        ("Downtown coffee COGS.", "32.6% in October. November's inventory change will separate stock build from usage."),
        ("Labor without the accrual noise.", "November is the first clean month on the new accrual; the targets test is Downtown ≤32%, Studio ≤30%, 1899 ≤35%."),
        ("Inventory COGS starts.", "From November each cafe's COGS includes the change in its month-end count. Expect one-month swings while the counting settles."),
        ("Holiday weeks.", "Veterans Day (Wed 11/11, open) and Thanksgiving week (Downtown and 1899 closed Thu–Sat) are set day by day in the projections."),
        ("Commissary.", "Two months of losses (−$2.1K, −$4.6K) as labor climbs; Action 4."),
        ("NMWA café ramp.", "30 weekday transactions a day in October; the November read with events stripped out is the one to watch."),
    ],
    cav_A1=("Commissary ($2.1K) → ($4.6K): sales +18% but labor +$2.8K, part of it September accrual catch-up. Central "
            "($10.6K) → ($14.6K): payroll +$4.0K (same accrual effect), the one-time $857 cash on hand write-off, offset by "
            "an $800 legal refund. SRP ($0.6K) and Cafe Management $6.0K steady."),
    cav_B=("Total net income $21,429, down $6.4K from a restated September ($27,788). Cafes contributed $35,149, of which "
           "NMWA $15,334; support entities ($13,720). Jan–Oct total: $132,869."),
)


# ------------------------------------------------------------------ build
def build():
    fin = json.load(open(FIN_JSON))
    pat = json.load(open(PAT_JSON))
    loy = json.load(open(LOY_JSON))
    loyp = json.load(open(LOY_PRIOR_JSON))
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9)
    sec.top_margin = sec.bottom_margin = Inches(0.85)
    doc.styles["Normal"].font.name = FONT_NAME
    doc.styles["Normal"].font.size = Pt(10)

    # masthead
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0),
                                          ("Monthly Business Review", 22, True, COLOR_ACCENT, 0),
                                          (MONTH_LABEL, 14, True, None, 2)]:
        p = doc.add_paragraph(); r = p.add_run(txt)
        _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(p, before=0, after=after)
    add_para(doc, NARRATIVE["masthead"], italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------------- Executive summary
    add_heading(doc, "Executive Summary", level=1)
    add_callout(doc, "The Formula",
                "Control COGS  +  Protect Labor Hours  +  Grow Sales  =  Better Business + Better Pay. "
                "Every headline below maps back to one of these three levers.", body_size=10, title_size=11)
    add_para(doc, NARRATIVE["exec_para"], after=8)
    for lead, body in NARRATIVE["exec_bullets"]:
        add_bullet(doc, body, bold_lead=lead)

    # ---------------- 1. Financials
    m, pm = MONTH, MONTHS[MONTHS.index(MONTH) - 1]
    add_heading(doc, f"1. Cafe Financials: {MONTH_LABEL}", level=1)
    add_para(doc, f"Source: Consolidated Model {MODEL_VER}. Core = cafe operations net of catering. Prime Cost = "
                  "COGS + Labor; Prime Margin = Sales − Prime Cost. 'Open Cafes' = Downtown + Studio + 1899. "
                  "NMWA is shown in its own column and in 1f but excluded from blended ratios while in ramp.",
             italic=True, size=9, color=COLOR_MUTED, after=10)

    add_heading(doc, f"1a. Core Cafe P&L: {MONTH_LABEL}", level=2)
    F = {c: fin["cafes"][c][m] for c in CAFES}; O = blend(fin, m)
    def row(label, key, money=True, denom=None):
        cells = [label]
        for c in CAFES + ["O"]:
            d = O if c == "O" else F[c]
            if money:
                cells.append(fmt_money(d[key]))
            else:
                cells.append(fmt_pct(d[key], d[denom]))
        return cells
    rows_1a = [
        ["Sales", "", "", "", "", ""],
        row("  Coffee: Beverage", "cbev"), row("  Coffee: Retail (bean)", "cret"), row("  Food", "food"),
        row("  Less: Discounts / Voids / Refunds", "disc_ref"),
        row("Total Core Sales (Revenue)", "core"),
        row("  Coffee % of Core Sales", "coffee", False, "core"), row("  Food % of Core Sales", "food", False, "core"),
        ["Cost of Goods Sold", "", "", "", "", ""],
        row("  Coffee COGS", "ccogs"), row("  Coffee COGS % of Coffee", "ccogs", False, "coffee"),
        row("  Food COGS", "fcogs"), row("  Food COGS % of Food", "fcogs", False, "food"),
        row("  Merchandise / Retail Purchases", "merch"),
        row("Total Core COGS", "cogs"), row("  COGS % of Core Sales", "cogs", False, "core"),
        row("Gross Profit (Revenue − COGS)", "gp"), row("  Gross Profit % of Core Sales", "gp", False, "core"),
        ["Labor", "", "", "", "", ""],
        row("  Total Labor", "labor"), row("  Labor % of Core Sales", "labor", False, "core"),
        row("Core Prime Cost", "prime"), row("  Prime Cost % of Core Sales", "prime", False, "core"),
        row("Core Prime Margin", "pm"), row("  Prime Margin % of Core Sales", "pm", False, "core"),
    ]
    build_investor_table(doc, ["Metric (Core, ex-catering)"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"], rows_1a,
                         col_widths=[Inches(2.55), Inches(1.0), Inches(1.0), Inches(1.0), Inches(0.85), Inches(1.1)],
                         key_rows=[0, 8, 17, 18], total_rows=[5, 14, 16, 21, 23])
    add_caveat(doc, NARRATIVE["cav_1a"])
    def band(v, lo, hi):
        return "in band" if v <= hi else ("above" if v <= hi + 8 else "well above")
    def pband(v):
        return "healthy" if v <= 60 else ("warning" if v <= 65 else "problem")
    def pv(c, n, d):
        return F[c][n] / F[c][d] * 100 if F[c][d] else 0
    add_callout(doc, "How to read this against the Cafe Finance 101 targets",
                "Coffee COGS target 22–27% → " + ", ".join(f"{CAFE_LABEL[c]} {pv(c, 'ccogs', 'coffee'):.1f}% ({band(pv(c, 'ccogs', 'coffee'), 22, 27)})" for c in OPEN3) + ". "
                "Food COGS target 28–35% → " + ", ".join(f"{CAFE_LABEL[c]} {pv(c, 'fcogs', 'food'):.1f}% ({band(pv(c, 'fcogs', 'food'), 28, 35)})" for c in OPEN3) + ". "
                "Prime Cost 55–60% healthy / 60–65% warning / >65% problem → " + ", ".join(f"{CAFE_LABEL[c]} {pv(c, 'prime', 'core'):.1f}% ({pband(pv(c, 'prime', 'core'))})" for c in OPEN3) + ". "
                f"Blended {fmt_pct(O['prime'], O['core'])}. " + NARRATIVE["callout_1a"], body_size=9, title_size=10)

    # 1b catering
    add_heading(doc, f"1b. Catering Sales: {RANGE}", level=2)
    rows = [[f"{mm} 2026"] + [fmt_money(fin["cafes"][c][mm]["cat"]) for c in CAFES] + [fmt_money(sum(fin["cafes"][c][mm]["cat"] for c in CAFES))] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["All Cafes"], rows, col_widths=[Inches(1.1)] + [Inches(1.2)] * 5, total_rows=[len(rows) - 1])
    add_caveat(doc, NARRATIVE["cav_1b"])

    # 1c trends
    add_heading(doc, f"1c. Prime Cost & Component Trends: {RANGE} (ex-catering)", level=2)
    add_para(doc, "Prime Cost = COGS + Labor, as % of Core Sales. Splitting the components shows which lever is moving.", italic=True, size=9, color=COLOR_MUTED, after=6)
    for title, num, den in [("Prime Cost % of Core Sales", "prime", "core"), ("COGS % of Core Sales", "cogs", "core"), ("Labor % of Core Sales", "labor", "core")]:
        add_heading(doc, title, level=3)
        rows = []
        for mm in MONTHS:
            b = blend(fin, mm)
            rows.append([f"{mm} 2026"] + [fmt_pct(fin["cafes"][c][mm][num], fin["cafes"][c][mm][den]) for c in OPEN3] + [fmt_pct(b[num], b[den])])
        build_investor_table(doc, ["Period", "Downtown", "Studio", "1899 Penn", "Open Cafes Blended"], rows, col_widths=[Inches(1.15)] + [Inches(1.3)] * 3 + [Inches(1.7)], total_rows=[len(rows) - 1])
    add_caveat(doc, NARRATIVE["cav_1c"])

    # 1d core sales trend
    add_heading(doc, f"1d. Core Sales Trend: {RANGE} (ex-catering)", level=2)
    rows = [[f"{mm} 2026"] + [fmt_money(fin["cafes"][c][mm]["core"]) for c in CAFES] + [fmt_money(blend(fin, mm)["core"])] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"], rows, col_widths=[Inches(1.35)] + [Inches(1.15)] * 5, total_rows=[len(rows) - 1])
    oa, oj = blend(fin, m)["core"], blend(fin, pm)["core"]
    add_caveat(doc, f"Open Cafes {fmt_money(oa)} in {MONTH_NAME}, {delta_pct(oa, oj)} vs {PM_NAME} and {delta_pct(oa, blend(fin, 'Apr')['core'])} vs the April peak. "
                    f"Downtown {delta_pct(F['Downtown']['core'], fin['cafes']['Downtown'][pm]['core'])}, Studio {delta_pct(F['Studio']['core'], fin['cafes']['Studio'][pm]['core'])}, "
                    f"1899 {delta_pct(F['1899']['core'], fin['cafes']['1899'][pm]['core'])}. " + NARRATIVE["cav_1d"])

    # 1e actual vs projected
    add_heading(doc, f"1e. Actual vs. Projected: {MONTH_LABEL}", level=2)
    add_para(doc, NARRATIVE["src_1e"], italic=True, size=9, color=COLOR_MUTED, after=6)
    rows = []
    tot_a = tot_p = 0
    for c in CAFES:
        a, p_ = F[c]["core"], fin["proj_sales"][c][m]
        rows.append([CAFE_LABEL[c], fmt_money(a), fmt_money(p_), fmt_money(a - p_) if a - p_ >= 0 else f"({fmt_money(p_ - a, False)})", delta_pct(a, p_) if p_ else "–"])
        if c in OPEN3:
            tot_a += a; tot_p += p_
    rows.append(["Open Cafes", fmt_money(tot_a), fmt_money(tot_p), f"({fmt_money(tot_p - tot_a, False)})" if tot_a < tot_p else fmt_money(tot_a - tot_p), delta_pct(tot_a, tot_p)])
    build_investor_table(doc, ["Cafe", "Sales Actual", "Sales Proj", "Δ ($)", "Δ (%)"], rows, col_aligns=["left"] + ["right"] * 4, col_widths=[Inches(1.6), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.1)], total_rows=[4])
    add_caveat(doc, NARRATIVE["cav_1e"])

    # 1f NMWA ramp
    add_heading(doc, "1f. NMWA: Opening Ramp", level=2)
    rows = [["Core Sales (ex-catering)"] + [fmt_money(fin["cafes"]["NMWA"][mm]["core"]) for mm in MONTHS[-4:]],
            ["Catering & Events"] + [fmt_money(fin["cafes"]["NMWA"][mm]["cat"]) for mm in MONTHS[-4:]],
            ["Weekday avg transactions/day"] + [f"{fin['cafes']['NMWA'][mm]['wd_txn']:.1f}" for mm in MONTHS[-4:]],
            ["Weekend avg transactions/day"] + [f"{fin['cafes']['NMWA'][mm]['we_txn']:.1f}" for mm in MONTHS[-4:]],
            ["Weekday avg ticket"] + [f"${fin['cafes']['NMWA'][mm]['wd_tkt']:.2f}" for mm in MONTHS[-4:]],
            ["Total Labor"] + [fmt_money(fin["cafes"]["NMWA"][mm]["labor"]) for mm in MONTHS[-4:]],
            ["Net Operating Income"] + [fmt_money(fin["cafes"]["NMWA"][mm]["noi"]) for mm in MONTHS[-4:]],
            ["Net Income (incl. mgmt fee)"] + [fmt_money(fin["cafes"]["NMWA"][mm]["ni"]) for mm in MONTHS[-4:]]]
    build_investor_table(doc, ["Metric"] + [f"{mm} 2026" for mm in MONTHS[-4:]], rows, col_aligns=["left"] + ["right"] * 4, col_widths=[Inches(2.5)] + [Inches(1.2)] * 4, key_rows=[2, 3], total_rows=[7])
    add_caveat(doc, NARRATIVE["cav_1f"])

    # ---------------- 2. Operating indicators
    add_heading(doc, "2. Operating Indicators", level=1)
    add_para(doc, f"Source: Consolidated Model {MODEL_VER} 'Ops Metrics'. Weekday / weekend split. Downtown and 1899 are weekday-only.", italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, f"2a. Average Ticket ($, walk-in): {RANGE}", level=2)
    add_para(doc, "Ticket is the per-transaction walk-in average from the Square transactions store (invoice / catering excluded), not the model's Ops Metrics ticket.", italic=True, size=9, color=COLOR_MUTED, after=4)
    wt = pat.get("walkin_ticket_by_month", {})
    rows = [[f"{mm} 2026"] + [(f"${wt[c][mm]:.2f}" if wt.get(c, {}).get(mm) else "–") for c in CAFES] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES], rows, col_widths=[Inches(1.2)] + [Inches(1.4)] * 4, total_rows=[len(rows) - 1])
    add_caveat(doc, NARRATIVE["cav_2a"])
    add_heading(doc, f"2b. Weekday Average Daily Transactions: {RANGE}", level=2)
    rows = [[f"{mm} 2026"] + [(f"{fin['cafes'][c][mm]['wd_txn']:.0f}" if fin['cafes'][c][mm]['wd_txn'] else "–") for c in CAFES] + [f"{sum(fin['cafes'][c][mm]['wd_txn'] for c in CAFES):.0f}"] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["Total/Day"], rows, col_widths=[Inches(1.15)] + [Inches(1.2)] * 5, total_rows=[len(rows) - 1])
    add_heading(doc, "2c. Weekend Metrics: Studio & NMWA", level=2)
    rows = [[f"{mm} 2026", f"{fin['cafes']['Studio'][mm]['we_txn']:.0f}", f"${fin['cafes']['Studio'][mm]['we_tkt']:.2f}",
             (f"{fin['cafes']['NMWA'][mm]['we_txn']:.0f}" if fin['cafes']['NMWA'][mm]['we_txn'] else "–"),
             (f"${fin['cafes']['NMWA'][mm]['we_tkt']:.2f}" if fin['cafes']['NMWA'][mm]['we_tkt'] else "–")] for mm in MONTHS]
    build_investor_table(doc, ["Period", "Studio: Txns/Day", "Studio: Avg Ticket", "NMWA: Txns/Day", "NMWA: Avg Ticket"], rows, col_widths=[Inches(1.1), Inches(1.4), Inches(1.4), Inches(1.35), Inches(1.35)], total_rows=[len(rows) - 1])
    add_caveat(doc, NARRATIVE["cav_2c"])

    # ---------------- 3. Sales patterns (NEW)
    add_heading(doc, "3. Sales Patterns & Staffing Alignment", level=1)
    cal = pat["attach_calibration"]
    add_para(doc, "Source: payment-level Square transactions (walk-in only: invoices and $0 test transactions excluded) joined to daily FOH labor hours "
                  "from the weekly pipeline. This section answers three questions the P&L cannot: when in the day and week the revenue lands, whether labor hours "
                  "are placed where the revenue is, and what the basket looks like. CORE labor hours only (admin + training excluded). " + NARRATIVE["src_3_coverage"],
             italic=True, size=9, color=COLOR_MUTED, after=6)

    # 3a daypart
    add_heading(doc, "3a. Daypart Curve: Where the Day's Revenue Lands", level=2)
    add_para(doc, NARRATIVE["insights_3a"], after=6)
    hdr = ["Daypart"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"]
    rows = []
    for dp in DAYPARTS:
        r = [dp]
        for c in CAFES + ["open3"]:
            b = pat["open3"] if c == "open3" else pat["cafes"][c]
            r.append(pct_str(b["daypart"][dp]["share"]))
        rows.append(r)
    build_investor_table(doc, hdr, rows, col_widths=[Inches(1.8)] + [Inches(1.05)] * 5, key_rows=[1, 2])
    add_caveat(doc, NARRATIVE["cav_3a"])
    # avg ticket + attach by daypart, Open3
    add_heading(doc, "Avg ticket, attach and tip rate by daypart: Open Cafes", level=3)
    o3 = pat["open3"]
    rows = [[dp, f"${o3['daypart'][dp]['avg_ticket']:.2f}" if o3['daypart'][dp]['avg_ticket'] else "–",
             pct_str(o3["attach_by_daypart"][dp]), pct_str(o3["daypart"][dp]["tip_rate"]),
             pct_str(o3["daypart"][dp]["share"]), pct_str(o3["daypart"][dp]["share_prior"])] for dp in DAYPARTS]
    build_investor_table(doc, ["Daypart", "Avg Ticket", "Attach (proxy)", "Tip Rate", f"Share {MONTH}", f"Share {PM}"], rows,
                         col_widths=[Inches(1.8), Inches(1.0), Inches(1.15), Inches(1.0), Inches(1.0), Inches(1.0)], key_rows=[2, 3])
    add_caveat(doc, "Attach and ticket both peak mid-morning and lunch and trough in the 8–10a rush: the rush guest buys one coffee and goes. "
                    "That is normal and not a problem to solve; the opportunity is the 10a–2p window where the guest is already inclined to add food.")

    # 3b DOW + labor
    add_heading(doc, "3b. Day of Week: Revenue vs Labor Hours", level=2)
    add_para(doc, NARRATIVE["insights_3b"], after=6)
    for c in CAFES:
        b = pat["cafes"][c]
        add_heading(doc, f"{CAFE_LABEL[c]}", level=3)
        rows = []
        for d in DOW:
            if d not in b["dow"]:
                continue
            r = b["dow"][d]
            if r["days"] < 2 or r["txns"] < 10:      # skip stray single-day / test-transaction rows
                continue
            rows.append([d, fmt_money(r["net_per_day"]), f"{r['txns_per_day']:.0f}", f"${r['avg_ticket']:.2f}",
                         f"{r['labor_hours_per_day']:.1f}" if r.get("labor_hours_per_day") else "–",
                         f"${r['sales_per_labor_hour']:.0f}" if r.get("sales_per_labor_hour") else "–",
                         f"${r['cost_per_hour']:.0f}" if r.get("cost_per_hour") else "–",
                         pct_str(r["share"]), pct_str(r.get("labor_share")),
                         (f"+{r['alignment_gap']:.1f}" if (r.get("alignment_gap") or 0) > 0 else f"{r.get('alignment_gap', 0):.1f}") if r.get("alignment_gap") is not None else "–"])
        build_investor_table(doc, ["Day", "Sales/Day", "Txns/Day", "Avg Tkt", "Core Hrs/Day", "Sales / Hr", "$ / Hr", "Sales Share", "Labor Share", "Gap"], rows,
                             col_aligns=["left"] + ["right"] * 9,
                             col_widths=[Inches(0.5), Inches(0.8), Inches(0.65), Inches(0.65), Inches(0.8), Inches(0.75), Inches(0.6), Inches(0.75), Inches(0.75), Inches(0.55)])
    add_caveat(doc, "Gap = labor share of the week's hours minus sales share of the week's revenue. Positive = more labor than revenue would justify on that day; "
                    "negative = under-staffed relative to demand. A well-aligned schedule keeps every day within ±2 points. Sales / Labor Hr uses FOH hours from "
                    "Square timecards; the $60–80 range is where a 28–33% labor target lands at DC wage levels.")

    # 3c basket
    add_heading(doc, "3c. Basket: Ticket Size, Attach and Dining Mix", level=2)
    add_para(doc, NARRATIVE["insights_3c"], after=6)
    hdr = ["Metric"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"]
    def prow(label, fn):
        return [label] + [fn(pat["cafes"][c]) for c in CAFES] + [fn(pat["open3"])]
    rows = [
        prow("Avg ticket (walk-in)", lambda b: f"${b['avg_ticket']:.2f}"),
        prow("Median ticket", lambda b: f"${b['median_ticket']:.2f}"),
        prow(f"Attach proxy (tickets ≥ ${cal['threshold']:.2f})", lambda b: pct_str(b["attach_proxy"])),
        prow("  prior month", lambda b: pct_str(b["attach_proxy_prior"])),
        prow("Tip rate", lambda b: pct_str(b["tip_rate"])),
        prow("Pickup share", lambda b: pct_str(b["dining"].get("PICKUP", {}).get("share"))),
        prow("  Pickup avg ticket", lambda b: f"${b['dining']['PICKUP']['avg_ticket']:.2f}" if "PICKUP" in b["dining"] else "–"),
        prow("  In-store avg ticket", lambda b: f"${b['dining']['IN_STORE']['avg_ticket']:.2f}" if "IN_STORE" in b["dining"] else "–"),
    ]
    build_investor_table(doc, hdr, rows, col_widths=[Inches(2.2)] + [Inches(0.95)] * 5, key_rows=[2, 5])
    add_heading(doc, "Ticket-size distribution: Open Cafes", level=3)
    rows = [[k, f"{v['txns']:,}", pct_str(v["share"])] for k, v in pat["open3"]["ticket_dist"].items()]
    build_investor_table(doc, ["Ticket", "Transactions", "Share"], rows, col_aligns=["left", "right", "right"], col_widths=[Inches(1.5), Inches(1.5), Inches(1.2)], key_rows=[1])
    add_caveat(doc, f"Attach proxy calibration: in April (the last export with full multi-item descriptions) {cal['april_multi_item_rate']}% of tickets had 2+ items; "
                    f"tickets ≥ ${cal['threshold']:.2f} matched that at {cal['april_rate_at_threshold']}%. The proxy is applied from July on, where descriptions are truncated to the first item. "
                    "The $5–8 band is the single-coffee ticket and is 45–50% of all transactions: the size of the attach opportunity.")

    # ---------------- 4. Items (from transactions)
    add_heading(doc, f"4. Item-Level Insights: {MONTH_LABEL}", level=1)
    add_para(doc, "Source: transaction lead item (first item on the ticket). Square exports since July truncate multi-item descriptions, so unit counts here "
                  "UNDER-count items in 2+ item baskets; the Square Item Sales CSV remains the source of truth for units. What this view adds that the CSV cannot: "
                  "which item anchors the purchase in each daypart.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, "4a. Lead item by daypart: Open Cafes", level=2)
    rows = []
    for dp in DAYPARTS[:5]:
        items = pat["open3"]["lead_item_by_daypart"].get(dp, [])
        rows.append([dp] + [f"{it['item']} ({it['txns']:,})" if i < len(items) else "" for i, it in enumerate(items[:4])] if items else [dp, "", "", "", ""])
    build_investor_table(doc, ["Daypart", "#1", "#2", "#3", "#4"], rows, col_aligns=["left"] * 5, col_widths=[Inches(1.6)] + [Inches(1.4)] * 4)
    add_heading(doc, "4b. Lead-item category share by daypart: Open Cafes", level=2)
    cats = ["coffee_bev", "food", "retail_coffee", "other"]
    rows = [[dp] + [pct_str(pat["open3"]["lead_category_by_daypart"].get(dp, {}).get(c)) for c in cats] for dp in DAYPARTS[:5]]
    build_investor_table(doc, ["Daypart", "Coffee / Bev", "Food", "Retail Bean", "Other"], rows, col_widths=[Inches(1.8)] + [Inches(1.2)] * 4, key_rows=[1])
    add_caveat(doc, "Food as the LEAD item is highest at lunch and mid-morning. Because only the first item is visible, a coffee-led ticket that also has a pastry shows as "
                    "'coffee' here: read this as 'what the guest came for', not a unit count.")
    add_heading(doc, "4c. Top 15 lead items: Open Cafes", level=2)
    rows = [[i + 1, it["item"], f"{it['txns']:,}"] for i, it in enumerate(pat["open3"]["top_lead_items"])]
    build_investor_table(doc, ["#", "Item", "Tickets led"], rows, col_aligns=["right", "left", "right"], col_widths=[Inches(0.45), Inches(3.8), Inches(1.4)])
    add_caveat(doc, NARRATIVE["cav_4"])

    # ---------------- 5. Discounts (hold)
    add_heading(doc, f"5. Discount Report: {MONTH_LABEL}", level=1)
    add_para(doc, "On hold per Ryan. Discounts as booked: " + ", ".join(f"{CAFE_LABEL[c]} {fmt_money(-F[c]['disc_ref'], False)}" for c in OPEN3) +
                  f"; open cafes {fmt_money(-O['disc_ref'], False)} / {fmt_pct(-O['disc_ref'], O['core'])} of open-cafe core sales. The transaction export does not carry discount names, so a by-type view still needs a Square Orders export.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------------- 6. Loyalty (from Square Loyalty API ledger: NEW source)
    add_heading(doc, f"6. Loyalty Report: {MONTH_LABEL}", level=1)
    add_para(doc, "NEW SOURCE. Built from the Square Loyalty API event ledger (every points-earning visit and reward redemption, "
                  "with account, cafe and timestamp), pulled by the Sunday data job. Replaces the dashboard screenshots. "
                  "Visit-frequency buckets are now fixed at 1 / 2 / 3–4 / 5–7 / 8+ every month. A 'visit' = one "
                  "ACCUMULATE_POINTS event; loyalty spend = points ÷ 7 (pre-tax). 'First-time' = account with no prior "
                  "event in the store (store starts Aug 1). NMWA is not "
                  "enrolled in the program.", italic=True, size=9, color=COLOR_MUTED, after=6)
    L, LP = loy["cafes"], loyp["cafes"]
    L3, LP3 = loy["open3"], loyp["open3"]
    add_heading(doc, "6a. Headline Insights", level=2)
    for lead, body in NARRATIVE["loyalty_bullets"]:
        add_bullet(doc, body, bold_lead=lead)

    add_heading(doc, f"6b. Loyalty Activity by Cafe: {MONTH_NAME} vs {PM_NAME}", level=2)
    def lrow(label, fn, fmt=lambda v: f"{v:,}" if isinstance(v, int) else (f"{v:.2f}" if isinstance(v, float) else "–")):
        cells = [label]
        for c in OPEN3:
            cells.append(f"{fmt(fn(L[c]))}  ({fmt(fn(LP[c]))})")
        cells.append(f"{fmt(fn(L3))}  ({fmt(fn(LP3))})")
        return cells
    money = lambda v: f"${v:,.0f}"
    rows = [
        lrow("Loyalty visits", lambda b: b["visits"]),
        lrow("Unique members visiting", lambda b: b["members"]),
        lrow("Avg visits / member", lambda b: b["avg_visits_per_member"]),
        lrow("First-time members (no prior event in ledger)", lambda b: b["first_time_members"]),
        lrow("Repeat members", lambda b: b["repeat_members"]),
        lrow("Loyalty share of walk-in tickets", lambda b: b["loyalty_share_of_tickets"], lambda v: pct_str(v)),
        lrow("Loyalty spend (pre-tax)", lambda b: b["loyalty_spend_pretax"], money),
        lrow("Avg spend / visit", lambda b: b["avg_spend_per_visit"], lambda v: f"${v:.2f}"),
        lrow("Avg spend / member / month", lambda b: b["avg_spend_per_member"], lambda v: f"${v:.2f}"),
        lrow("Rewards redeemed", lambda b: b["redemptions"]),
    ]
    build_investor_table(doc, [f"Metric  ({MONTH}  ({PM}))", "Downtown", "Studio", "1899 Penn", "Open Cafes"], rows,
                         col_widths=[Inches(2.3)] + [Inches(1.15)] * 4, key_rows=[0, 5])
    add_caveat(doc, NARRATIVE["cav_6b"])

    add_heading(doc, "6c. Loyalty Member Value: In Dollars", level=2)
    rows = []
    for c in OPEN3 + ["open3"]:
        bb = L3 if c == "open3" else L[c]
        mo = bb["avg_spend_per_member"] or 0
        rows.append([("Open Cafes" if c == "open3" else CAFE_LABEL[c]), f"{bb['avg_visits_per_member']:.2f}", f"${bb['avg_spend_per_visit']:.2f}",
                     f"${mo:.2f}", f"${mo * 12:,.0f}", f"${mo * 60:,.0f}"])
    build_investor_table(doc, ["Cafe", "Visits / mo", "$ / visit", "$ / member / mo", "Annualized", "5-yr value"], rows,
                         col_aligns=["left"] + ["right"] * 5, col_widths=[Inches(1.5), Inches(1.0), Inches(1.0), Inches(1.3), Inches(1.1), Inches(1.1)], total_rows=[3])
    add_caveat(doc, NARRATIVE["cav_6c"])

    add_heading(doc, f"6d. Visit-Frequency Distribution: {MONTH_NAME} (members by visits in month)", level=2)
    BK = ["1 visit", "2 visits", "3–4 visits", "5–7 visits", "8+ visits"]
    rows = []
    for k in BK:
        r = [k]
        for c in OPEN3 + ["open3"]:
            bb = L3 if c == "open3" else L[c]
            n = bb["buckets"][k]; r.append(f"{n:,}  ({pct_str(n / bb['members'] * 100, 0)})")
        rows.append(r)
    build_investor_table(doc, ["Bucket  (members  (share))", "Downtown", "Studio", "1899 Penn", "Open Cafes"], rows,
                         col_widths=[Inches(2.0)] + [Inches(1.2)] * 4, key_rows=[3, 4])
    add_caveat(doc, NARRATIVE["cav_6d"])

    # ---------------- 7. Benchmarks + actions (NEW)
    add_heading(doc, "7. Benchmarks & Prioritized Actions", level=1)
    add_para(doc, "NEW. Each cafe against typical ranges for independent specialty coffee in high-cost US metros. The ranges are directional operator "
                  "benchmarks, not audited statistics: they are useful for knowing which number to work on first, not for judging a single month.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, f"7a. Scorecard: {MONTH_LABEL}", level=2)
    def sc(label, key, fn, note):
        return [label, BENCH[key][0]] + [fn(c) for c in OPEN3] + [fn("open3")]
    def fin_pct(c, num, den):
        d = O if c == "open3" else F[c]
        return fmt_pct(d[num], d[den])
    def pat_v(c, fn):
        b = pat["open3"] if c == "open3" else pat["cafes"][c]
        return fn(b)
    rows = [
        sc("Prime cost % of core", "prime_pct", lambda c: fin_pct(c, "prime", "core"), ""),
        sc("Labor % of core", "labor_pct", lambda c: fin_pct(c, "labor", "core"), ""),
        sc("Coffee COGS % of coffee", "coffee_cogs_pct", lambda c: fin_pct(c, "ccogs", "coffee"), ""),
        sc("Food COGS % of food", "food_cogs_pct", lambda c: fin_pct(c, "fcogs", "food"), ""),
        sc("Sales per CORE labor hour (matched days)", "splh", lambda c: pat_v(c, lambda b: f"${b['labor']['sales_per_labor_hour']:.0f}" if b['labor']['sales_per_labor_hour'] else "–"), ""),
        sc("Loaded cost per core labor hour", "cost_hr", lambda c: pat_v(c, lambda b: f"${b['labor']['cost_per_hour']:.2f}" if b['labor'].get('cost_per_hour') else "–"), ""),
        sc("Transactions per labor hour", "tplh", lambda c: pat_v(c, lambda b: f"{b['labor']['txns_per_labor_hour']:.1f}" if b['labor']['txns_per_labor_hour'] else "–"), ""),
        sc("Avg ticket (walk-in)", "avg_ticket", lambda c: pat_v(c, lambda b: f"${b['avg_ticket']:.2f}"), ""),
        sc("Attach rate (proxy)", "attach", lambda c: pat_v(c, lambda b: pct_str(b["attach_proxy"])), ""),
        sc("7–10a share of sales", "morning_share", lambda c: pat_v(c, lambda b: pct_str(round((b['daypart']['Early (6-8a)']['share'] or 0) + (b['daypart']['Rush (8-10a)']['share'] or 0), 1))), ""),
        sc("Tip rate", "tip_rate", lambda c: pat_v(c, lambda b: pct_str(b["tip_rate"])), ""),
        sc("Pickup share", "pickup_share", lambda c: pat_v(c, lambda b: pct_str(b["dining"].get("PICKUP", {}).get("share"))), ""),
    ]
    build_investor_table(doc, ["Metric", "Typical range", "Downtown", "Studio", "1899 Penn", "Open Cafes"], rows,
                         col_aligns=["left", "left", "right", "right", "right", "right"],
                         col_widths=[Inches(1.9), Inches(1.0), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95)], key_rows=[0, 1, 4])
    add_caveat(doc, NARRATIVE["cav_7a"] + " " + "Benchmark notes: " + "; ".join(f"{k.replace('_', ' ')}: {v[1]}" for k, v in list(BENCH.items())[:7]) + ".")

    add_heading(doc, "7b. Prioritized actions", level=2)
    add_para(doc, "Ranked by estimated impact against effort. Each is specific enough to be assigned and checked in the next weekly report.", italic=True, size=9, color=COLOR_MUTED, after=6)
    for i, (title, body, impact) in enumerate(NARRATIVE["actions"], 1):
        add_bullet(doc, f"{body}  [Est. impact: {impact}]", bold_lead=f"{i}. {title}")

    # ---------------- 8. Notes
    add_heading(doc, "8. Notes & Watch Items", level=1)
    add_heading(doc, "Data caveats", level=3)
    for b in NARRATIVE["data_caveats"]:
        add_bullet(doc, b)
    add_heading(doc, f"Watch items going into {NEXT_NAME}", level=3)
    for lead, body in NARRATIVE["watch"]:
        add_bullet(doc, body, bold_lead=lead)

    # ---------------- Appendices
    doc.add_paragraph().add_run().add_break()
    p = doc.add_paragraph(); r = p.add_run("LEADERSHIP TEAM ONLY: APPENDICES"); _set_run_font(r, size=12, bold=True, color=COLOR_MUTED); _set_para_spacing(p, before=10, after=6)
    S = fin["support"]
    add_heading(doc, f"A. Support Entities: {MONTH_LABEL}", level=2)
    rows = [["Revenue", fmt_money(S["Commissary"][m]["rev"]), "$0", fmt_money(S["Cafe Mgmt"][m]["rev"]), fmt_money(S["SRP"][m]["rev"]), fmt_money(S["Commissary"][m]["rev"] + S["Cafe Mgmt"][m]["rev"] + S["SRP"][m]["rev"])],
            ["COGS", fmt_money(S["Commissary"][m]["cogs"]), "$0", "$0", fmt_money(S["SRP"][m]["cogs"]), fmt_money(S["Commissary"][m]["cogs"] + S["SRP"][m]["cogs"])],
            ["Total Labor", fmt_money(S["Commissary"][m]["labor"]), fmt_money(S["Central"][m]["labor"]), "$0", fmt_money(S["SRP"][m]["labor"]), fmt_money(S["Commissary"][m]["labor"] + S["Central"][m]["labor"] + S["SRP"][m]["labor"])],
            ["Expenses", fmt_money(S["Commissary"][m]["opex"]), fmt_money(S["Central"][m]["opex"]), "$0", fmt_money(S["SRP"][m]["opex"]), fmt_money(S["Commissary"][m]["opex"] + S["Central"][m]["opex"] + S["SRP"][m]["opex"])],
            ["Net Income", fmt_money(S["Commissary"][m]["ni"]), fmt_money(S["Central"][m]["ni"]), fmt_money(S["Cafe Mgmt"][m]["ni"]), fmt_money(S["SRP"][m]["ni"]), fmt_money(sum(S[e][m]["ni"] for e in S))]]
    build_investor_table(doc, ["Metric", "Commissary", "Central Admin", "Cafe Mgmt", "SRP (Roasting)", "Support Total"], rows, col_widths=[Inches(1.5), Inches(1.15), Inches(1.2), Inches(1.05), Inches(1.25), Inches(1.15)], total_rows=[4])
    add_heading(doc, "A1. Support Entities: Net Income Trend", level=3)
    tm = MONTHS[-6:]
    rows = [[e] + [fmt_money(S[e][mm]["ni"]) for mm in tm] + [fmt_money(S[e][m]["ni"] - S[e][pm]["ni"])] for e in ["Commissary", "Central", "Cafe Mgmt", "SRP"]]
    build_investor_table(doc, ["Entity"] + tm + ["MoM Δ"], rows, col_widths=[Inches(1.35)] + [Inches(0.9)] * 7)
    add_caveat(doc, NARRATIVE["cav_A1"])

    add_heading(doc, f"B. Consolidated P&L: {MONTH_LABEL}", level=2)
    C = fin["consolidated"]
    tm = MONTHS[-6:]
    def crow(label, key):
        return [label] + [fmt_money(C[mm][key]) for mm in tm] + [delta_pct(C[m][key], C[pm][key])]
    rows = [crow("Total Core Location Sales", "tcls"), crow("Total COGS", "tcogs"), crow("Gross Profit", "gp"), crow("Total Labor", "labor"),
            crow("Total Expenses", "texp"), crow("Net Operating Income", "noi"), crow("Cafes: Net Income", "cafes_ni"), crow("TOTAL NET INCOME (all entities)", "total_ni")]
    build_investor_table(doc, ["Metric"] + tm + ["MoM Δ"], rows, col_widths=[Inches(2.1)] + [Inches(0.85)] * 6 + [Inches(0.8)], key_rows=[5, 6], total_rows=[7])
    add_caveat(doc, NARRATIVE["cav_B"])

    add_heading(doc, "C. Methodology & Data Sources", level=2)
    add_bullet(doc, f"Financials: Consolidated Model {MODEL_VER} → {os.path.basename(FIN_JSON)} (extract_fin.py: Jan–{MONTH} per-entity IS tabs, Consolidated, Projections).")
    add_bullet(doc, f"Sales patterns: _pipeline/shared/transactions (Square Transactions export, {MONTH_NAME} rows from the month file plus spillover in the neighbouring files) + weekly labor JSONs → analyze_patterns.py → {os.path.basename(PAT_JSON)}.")
    add_bullet(doc, "Report: build_monthly_report_v2.py reads both JSONs; narrative blocks live in NARRATIVE at the top of that file.")
    add_bullet(doc, "Definitions: Core Sales = Total Core Location Sales − Catering. Core COGS = Total COGS − Catering COGS. Prime = Core COGS + Total Labor. Open Cafes = Downtown + Studio + 1899.")
    add_bullet(doc, "Daypart boundaries: Early 6–8a, Rush 8–10a, Mid-morning 10a–12p, Lunch 12–2p, Afternoon 2–5p, Late 5p+. Alignment gap = labor-hour share − sales share by day of week.")

    _ensure_zoom_percent(doc)
    doc.save(OUT)
    print("Saved →", OUT)


if __name__ == "__main__":
    build()
