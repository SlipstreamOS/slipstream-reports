#!/usr/bin/env python3
"""Slipstream Monthly Business Review — v2 template (data-driven).

Reads two JSONs produced upstream and renders the full report:
  data/fin_2026_v<N>.json     — financials Jan..<month> from the recalculated
                                 Consolidated Model (see extract in the Aug session)
  data/<yymm>_patterns.json   — transaction-level sales patterns
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
MONTH = "Aug"
MONTH_LABEL = "August 2026"
VERSION = "v1.1"
PREPARED = "September 4, 2026"
MODEL_VER = "v17"
FIN_JSON = os.path.join(DATA_DIR, "fin_2026_v17.json")
PAT_JSON = os.path.join(DATA_DIR, "2608_patterns.json")
OUT = os.path.join(WORKSPACE, f"Slipstream Monthly Report - {MONTH_LABEL} {VERSION}.docx")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"]
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
    "splh":             ("$60–80",   "sales per FOH labor hour; <$50 signals over-staffing or under-traffic"),
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
        return "—"
    if v < 0 and paren:
        return f"(${abs(v):,.0f})"
    return f"${v:,.0f}"


def fmt_pct(n, d, dp=1):
    if not d:
        return "—"
    v = n / d * 100
    return f"({abs(v):.{dp}f}%)" if v < 0 else f"{v:.{dp}f}%"


def pct_str(v, dp=1):
    if v is None:
        return "—"
    return f"({abs(v):.{dp}f}%)" if v < 0 else f"{v:.{dp}f}%"


def delta_pct(cur, prev):
    if not prev:
        return "—"
    v = (cur - prev) / abs(prev) * 100
    return f"({abs(v):.1f}%)" if v < 0 else f"+{v:.1f}%"


def blend(fin, m, cafes=OPEN3):
    keys = ["core", "coffee", "food", "cbev", "cret", "disc_ref", "ccogs", "fcogs", "merch", "cogs", "gp", "labor", "prime", "pm", "cat", "ni"]
    return {k: sum(fin["cafes"][c][m][k] for c in cafes) for k in keys}


# ------------------------------------------------------------------ narrative (edit per month)
NARRATIVE = dict(
    masthead=(
        f"Prepared {PREPARED}  |  {VERSION} — August results from Consolidated Model {MODEL_VER}. "
        "NEW in this template: Section 3 'Sales Patterns & Staffing Alignment' (daypart curve, "
        "day-of-week revenue vs labor hours, basket and attach), Section 4 item velocity by "
        "daypart from transaction data, and Section 7 'Benchmarks & Prioritized Actions'. The "
        "report is now generated from two data files, so September is a three-command refresh."
    ),
    exec_para=(
        "August was the weakest month of 2026 on every operating measure. Core sales at the three "
        "established cafes fell to $93,559 (-8.2% vs July, -21% vs the April peak), and the "
        "company posted its first consolidated loss since February: total net income across all "
        "entities was ($12,671). Blended prime cost hit 75.3% — deep in the problem zone — with "
        "labor at 42.2% of core sales the primary driver (July: 33.9%). Downtown recorded its "
        "first monthly loss (-$3,236). Central Admin carried ~$5,700 of one-time tax prep and tax payment; normalized for that, the consolidated result is roughly ($7,000). Two "
        "things need to be said together: August is structurally the softest month for a DC "
        "office-district coffee business (Congress in recess, federal leave, school out), so a "
        "sales dip was expected — but labor hours did not come down with it, and that, not the "
        "sales dip, is what turned the month negative. The pattern analysis in Section 3 shows "
        "exactly where the hours went."
    ),
    exec_bullets=[
        ("Labor is the whole story.",
         "Labor $ rose to $41,858 consolidated (+$4,772 vs July) while core sales fell $8,333. "
         "Downtown labor 48.0% of core, 1899 52.7%, Studio 35.2%. At the transaction level, "
         "sales per FOH labor hour was $46 blended — well under the $60–80 range a specialty "
         "cafe needs to cover a 30% labor target."),
        ("Friday at Downtown and 1899 is the single biggest scheduling miss.",
         "Downtown Friday averaged $671/day vs $1,725 on Tuesday — 39% of a mid-week day — yet "
         "carried 7.5 points MORE labor share than sales share. 1899 Friday: $441 vs $953 "
         "Wednesday, labor +6.1 points over. August Fridays in a federal district are close to "
         "a half-day; the schedule treated them as full."),
        ("Studio's weekend is under-staffed relative to demand; Mon/Tue over.",
         "Saturday $2,419/day and Sunday $1,883 are Studio's two best days by a wide margin, "
         "yet weekend labor share runs 4–5 points BELOW sales share while Monday/Tuesday run "
         "3 points above. That is the inverse of the demand curve."),
        ("Coffee COGS held; food COGS drifted up.",
         "Coffee COGS 28.6% blended (Studio 25.1% in band; Downtown 31.7%, 1899 31.9% above). "
         "Food COGS 40.3% — Downtown jumped to 46.7% on a shrinking food base, 1899 63.9%. "
         "Food COGS at Downtown has doubled its share of the problem since June."),
        ("Central Admin: ~$5,700 one-time.",
         "($20,196) vs ($11,362) in July. Per Ryan, ~$5,700 was tax prep and tax payment — "
         "non-recurring. Normalized Central Admin is ~($14,500), still ~$3,100 above July; "
         "and normalized consolidated NI is ~($7,000) rather than ($12,671). The month is still "
         "negative on operations, but by about half the headline figure."),
        ("NMWA is now a real café — and its best day is Wednesday.",
         "$3,606 core sales, 29 weekday transactions/day. Wednesday averages $713/day vs "
         "$85–226 the rest of the week (community-day traffic), and it is the day with the "
         "LEAST labor share relative to sales. That is the easiest schedule fix in the company."),
    ],
    insights_3a=(
        "Three different businesses. Downtown is an office-district cafe: 8–10a is 35% of the "
        "day and it fades hard after 2p (13%). Studio is a neighborhood cafe: the peak is "
        "10–12p (34%), not the rush, and weekend outperforms weekday. 1899 is the only cafe "
        "with a real lunch (30% of the day, its largest daypart). NMWA is a museum cafe — no "
        "rush at all, 48% of sales at lunch. These call for four different staffing templates, "
        "not one."
    ),
    insights_3b=(
        "Two patterns that are costing money right now. First, Friday: at both weekday-only "
        "cafes Friday revenue is 40–45% of a mid-week day in August, and both carry more labor "
        "share on Friday than sales share. Second, Studio's weekend: it earns 33% of the week's "
        "revenue on Saturday and Sunday but gets 24% of the labor hours; Monday and Tuesday "
        "earn 20% of revenue and get 26% of the hours. Sales per labor hour is $43–49 across "
        "the three established cafes — every one of them is below the $60 floor a 30% labor "
        "target implies at current wages."
    ),
    insights_3c=(
        "The attach proxy (share of tickets above $7.75, calibrated to April's actual 41% "
        "multi-item rate) sits at 39% blended — Studio 43%, Downtown 39%, 1899 32%. Attach is "
        "highest in the mid-morning and lunch dayparts and lowest in the 8–10a rush at every "
        "cafe, which is the industry-normal pattern (rush guests want speed). In-store tickets "
        "run $1.50 higher than pickup at Studio ($10.30 vs $8.83) — the dine-in guest is the "
        "one buying the pastry. Studio's 50/50 dine-in / pickup split is its structural "
        "advantage on ticket size; Downtown is 94% pickup."
    ),
    actions=[
        ("Re-base Friday schedules at Downtown and 1899 to Friday demand.",
         "Friday runs 40–45% of Tue–Thu volume in August at both cafes. Moving Friday to a "
         "2-person close-early template (vs the current mid-week template) is worth roughly "
         "$1,200–1,600 / month in labor at current hours, with no revenue risk given the "
         "demand curve. Revisit in October when Congress is back.",
         "$1,200–1,600 / mo"),
        ("Shift Studio hours from Mon/Tue to Sat/Sun.",
         "Move ~10–12 labor hours per week from Monday/Tuesday to the weekend. Weekend sales "
         "per labor hour will still be higher than weekday after the move, and the "
         "under-staffed weekend is where ticket size and attach are strongest — lost weekend "
         "throughput is the most expensive kind. Net labor cost is neutral; the gain is "
         "revenue protected at the highest-value daypart of the week.",
         "revenue-protective; labor-neutral"),
        ("Fix NMWA Wednesday staffing.",
         "Wednesday is 4x any other NMWA day and carries the lowest labor share relative to "
         "sales (-12 points). Add one shift on Wednesday, drop one on Saturday (+12 points "
         "over). Same hours, matched to the museum's own calendar.",
         "labor-neutral; throughput gain"),
        ("Attack food COGS at Downtown before it becomes structural.",
         "46.7% of food sales in August vs 30–36% Jan–Jul. On a $4,278 food base that is ~$500 / "
         "month of margin, but the direction is the problem: par levels were set for July "
         "volume and August was 24% lower. Reset pastry pars to trailing-2-week sales, and "
         "re-check the 1899 menu costing (63.9%) at the same time.",
         "$500–800 / mo"),
        ("Set a sales-per-labor-hour floor and put it on the weekly report.",
         "Every established cafe is running $43–49 SPLH. A 30% labor target at ~$20 loaded "
         "cost implies ~$65 SPLH. Adding SPLH by cafe by week to the Sunday report gives the "
         "scheduling manager a leading indicator instead of a month-end surprise.",
         "enabler for all of the above"),
    ],
    watch=[
        ("Central Admin run-rate.", "~$5,700 of August was one-time tax prep + payment. Confirm the remaining ~$3,100 step-up vs July is also non-recurring; if not, it is a new ~$14.5K/mo baseline."),
        ("September rebound.", "Federal workforce returns after Labor Day; Downtown and 1899 should recover 10–20% of August's shortfall in the first two weeks. If they do not, the softness is not seasonal."),
        ("Downtown food COGS.", "Two months of deterioration on a shrinking base. Track weekly."),
        ("SRP.", "Back to ($1,744) after July's +$4,178. Third reversal this year — the roasting P&L is not yet predictable month to month."),
        ("Studio weekend transactions.", "197/day in July → 214/day in August is a partial recovery from June's 258. Watch whether the staffing shift (Action 2) moves it."),
    ],
)


# ------------------------------------------------------------------ build
def build():
    fin = json.load(open(FIN_JSON))
    pat = json.load(open(PAT_JSON))
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
    add_heading(doc, f"1. Cafe Financials — {MONTH_LABEL}", level=1)
    add_para(doc, f"Source: Consolidated Model {MODEL_VER}. Core = cafe operations net of catering. Prime Cost = "
                  "COGS + Labor; Prime Margin = Sales − Prime Cost. 'Open Cafes' = Downtown + Studio + 1899. "
                  "NMWA is shown in its own column and in 1f but excluded from blended ratios while in ramp.",
             italic=True, size=9, color=COLOR_MUTED, after=10)

    add_heading(doc, f"1a. Core Cafe P&L — {MONTH_LABEL}", level=2)
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
        row("  Coffee — Beverage", "cbev"), row("  Coffee — Retail (bean)", "cret"), row("  Food", "food"),
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
    add_caveat(doc, "1899 Net Income is a pass-through (owner profit share) and shows $0 at the entity level; "
                    "its operating result is visible in Prime Margin. NMWA ratios are for completeness only.")
    add_callout(doc, "How to read this against the Cafe Finance 101 targets",
                f"Coffee COGS target 22–27% → Studio {fmt_pct(F['Studio']['ccogs'], F['Studio']['coffee'])} (in band), "
                f"Downtown {fmt_pct(F['Downtown']['ccogs'], F['Downtown']['coffee'])} and 1899 {fmt_pct(F['1899']['ccogs'], F['1899']['coffee'])} (above). "
                f"Food COGS target 28–35% → Studio {fmt_pct(F['Studio']['fcogs'], F['Studio']['food'])} (in band), "
                f"Downtown {fmt_pct(F['Downtown']['fcogs'], F['Downtown']['food'])} and 1899 {fmt_pct(F['1899']['fcogs'], F['1899']['food'])} (well above). "
                f"Prime Cost 55–60% healthy / 60–65% warning / >65% problem → Studio {fmt_pct(F['Studio']['prime'], F['Studio']['core'])} (warning), "
                f"Downtown {fmt_pct(F['Downtown']['prime'], F['Downtown']['core'])} and 1899 {fmt_pct(F['1899']['prime'], F['1899']['core'])} (problem). "
                f"Blended {fmt_pct(O['prime'], O['core'])} — the worst reading since January, and the first month where labor alone "
                f"({fmt_pct(O['labor'], O['core'])}) exceeds the entire healthy prime-cost band.", body_size=9, title_size=10)

    # 1b catering
    add_heading(doc, "1b. Catering Sales — Jan → Aug 2026", level=2)
    rows = [[f"{mm} 2026"] + [fmt_money(fin["cafes"][c][mm]["cat"]) for c in CAFES] + [fmt_money(sum(fin["cafes"][c][mm]["cat"] for c in CAFES))] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["All Cafes"], rows, col_widths=[Inches(1.1)] + [Inches(1.2)] * 5, total_rows=[len(rows) - 1])
    add_caveat(doc, "August catering essentially stopped ($289 company-wide) — the lowest month of the year. Consistent with the August calendar; watch for the September event pipeline.")

    # 1c trends
    add_heading(doc, "1c. Prime Cost & Component Trends — Jan → Aug 2026 (ex-catering)", level=2)
    add_para(doc, "Prime Cost = COGS + Labor, as % of Core Sales. Splitting the components shows which lever is moving.", italic=True, size=9, color=COLOR_MUTED, after=6)
    for title, num, den in [("Prime Cost % of Core Sales", "prime", "core"), ("COGS % of Core Sales", "cogs", "core"), ("Labor % of Core Sales", "labor", "core")]:
        add_heading(doc, title, level=3)
        rows = []
        for mm in MONTHS:
            b = blend(fin, mm)
            rows.append([f"{mm} 2026"] + [fmt_pct(fin["cafes"][c][mm][num], fin["cafes"][c][mm][den]) for c in OPEN3] + [fmt_pct(b[num], b[den])])
        build_investor_table(doc, ["Period", "Downtown", "Studio", "1899 Penn", "Open Cafes Blended"], rows, col_widths=[Inches(1.15)] + [Inches(1.3)] * 3 + [Inches(1.7)], total_rows=[len(rows) - 1])
    add_caveat(doc, "Labor % is the mover. It has risen every month since April (26.7% → 30.1% → 30.1% → 33.9% → 42.2%) while sales fell from the April peak. "
                    "COGS % has been range-bound at 29–35%. Prime cost is now 75.3% blended — Section 3 shows the schedule is not tracking the demand curve by day of week or by cafe.")

    # 1d core sales trend
    add_heading(doc, "1d. Core Sales Trend — Jan → Aug 2026 (ex-catering)", level=2)
    rows = [[f"{mm} 2026"] + [fmt_money(fin["cafes"][c][mm]["core"]) for c in CAFES] + [fmt_money(blend(fin, mm)["core"])] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"], rows, col_widths=[Inches(1.35)] + [Inches(1.15)] * 5, total_rows=[len(rows) - 1])
    oa, oj = blend(fin, m)["core"], blend(fin, pm)["core"]
    add_caveat(doc, f"Open Cafes {fmt_money(oa)} in August, {delta_pct(oa, oj)} vs July and {delta_pct(oa, blend(fin, 'Apr')['core'])} vs the April peak. "
                    f"Downtown {delta_pct(F['Downtown']['core'], fin['cafes']['Downtown'][pm]['core'])}, Studio {delta_pct(F['Studio']['core'], fin['cafes']['Studio'][pm]['core'])}, "
                    f"1899 {delta_pct(F['1899']['core'], fin['cafes']['1899'][pm]['core'])}. Downtown's decline is the largest in dollars and reflects the federal August calendar.")

    # 1e actual vs projected
    add_heading(doc, f"1e. Actual vs. Projected — {MONTH_LABEL}", level=2)
    add_para(doc, "Source: 2026 Projections v1.7 (FY26 budget) sales projections. Labor projections were not carried in v17's Projections tab; the sales gap alone tells the story this month.", italic=True, size=9, color=COLOR_MUTED, after=6)
    rows = []
    tot_a = tot_p = 0
    for c in CAFES:
        a, p_ = F[c]["core"], fin["proj_sales"][c][m]
        rows.append([CAFE_LABEL[c], fmt_money(a), fmt_money(p_), fmt_money(a - p_) if a - p_ >= 0 else f"({fmt_money(p_ - a, False)})", delta_pct(a, p_) if p_ else "—"])
        if c in OPEN3:
            tot_a += a; tot_p += p_
    rows.append(["Open Cafes", fmt_money(tot_a), fmt_money(tot_p), f"({fmt_money(tot_p - tot_a, False)})" if tot_a < tot_p else fmt_money(tot_a - tot_p), delta_pct(tot_a, tot_p)])
    build_investor_table(doc, ["Cafe", "Sales Actual", "Sales Proj", "Δ ($)", "Δ (%)"], rows, col_aligns=["left"] + ["right"] * 4, col_widths=[Inches(1.6), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.1)], total_rows=[4])
    add_caveat(doc, "The FY26 budget already assumed an August dip (Downtown projected at $29,657, its lowest month of the plan). Downtown landed within $125 of that number — "
                    "the August softness at Downtown was PLANNED. Studio and 1899 came in below their plans. NMWA beat its first budgeted month by $1,051.")

    # 1f NMWA ramp
    add_heading(doc, "1f. NMWA — Opening Ramp", level=2)
    rows = [["Core Sales (ex-catering)"] + [fmt_money(fin["cafes"]["NMWA"][mm]["core"]) for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Catering & Events"] + [fmt_money(fin["cafes"]["NMWA"][mm]["cat"]) for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Weekday avg transactions/day"] + [f"{fin['cafes']['NMWA'][mm]['wd_txn']:.1f}" for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Weekend avg transactions/day"] + [f"{fin['cafes']['NMWA'][mm]['we_txn']:.1f}" for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Weekday avg ticket"] + [f"${fin['cafes']['NMWA'][mm]['wd_tkt']:.2f}" for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Total Labor"] + [fmt_money(fin["cafes"]["NMWA"][mm]["labor"]) for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Net Operating Income"] + [fmt_money(fin["cafes"]["NMWA"][mm]["noi"]) for mm in ["May", "Jun", "Jul", "Aug"]],
            ["Net Income (incl. mgmt fee)"] + [fmt_money(fin["cafes"]["NMWA"][mm]["ni"]) for mm in ["May", "Jun", "Jul", "Aug"]]]
    build_investor_table(doc, ["Metric", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026"], rows, col_aligns=["left"] + ["right"] * 4, col_widths=[Inches(2.5)] + [Inches(1.2)] * 4, key_rows=[2, 3], total_rows=[7])
    add_caveat(doc, "Core sales +24% MoM to $3,606 on a museum-calendar month. Weekday transactions plateaued (31 → 29/day) while weekend rose (15 → 17/day). "
                    "Operating loss held at ~($1.9K); Section 3 shows the Wednesday community-day pattern that the schedule is not yet built around.")

    # ---------------- 2. Operating indicators
    add_heading(doc, "2. Operating Indicators", level=1)
    add_para(doc, f"Source: Consolidated Model {MODEL_VER} 'Ops Metrics'. Weekday / weekend split. Downtown and 1899 are weekday-only.", italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, "2a. Weekday Average Ticket ($) — Jan → Aug 2026", level=2)
    rows = [[f"{mm} 2026"] + [(f"${fin['cafes'][c][mm]['wd_tkt']:.2f}" if fin['cafes'][c][mm]['wd_tkt'] else "—") for c in CAFES] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES], rows, col_widths=[Inches(1.2)] + [Inches(1.4)] * 4, total_rows=[len(rows) - 1])
    add_heading(doc, "2b. Weekday Average Daily Transactions — Jan → Aug 2026", level=2)
    rows = [[f"{mm} 2026"] + [(f"{fin['cafes'][c][mm]['wd_txn']:.0f}" if fin['cafes'][c][mm]['wd_txn'] else "—") for c in CAFES] + [f"{sum(fin['cafes'][c][mm]['wd_txn'] for c in CAFES):.0f}"] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["Total/Day"], rows, col_widths=[Inches(1.15)] + [Inches(1.2)] * 5, total_rows=[len(rows) - 1])
    add_heading(doc, "2c. Weekend Metrics — Studio & NMWA", level=2)
    rows = [[f"{mm} 2026", f"{fin['cafes']['Studio'][mm]['we_txn']:.0f}", f"${fin['cafes']['Studio'][mm]['we_tkt']:.2f}",
             (f"{fin['cafes']['NMWA'][mm]['we_txn']:.0f}" if fin['cafes']['NMWA'][mm]['we_txn'] else "—"),
             (f"${fin['cafes']['NMWA'][mm]['we_tkt']:.2f}" if fin['cafes']['NMWA'][mm]['we_tkt'] else "—")] for mm in MONTHS]
    build_investor_table(doc, ["Period", "Studio — Txns/Day", "Studio — Avg Ticket", "NMWA — Txns/Day", "NMWA — Avg Ticket"], rows, col_widths=[Inches(1.1), Inches(1.4), Inches(1.4), Inches(1.35), Inches(1.35)], total_rows=[len(rows) - 1])
    add_caveat(doc, "Weekday transactions fell at all three established cafes (Downtown 193 → 158, Studio 133 → 129, 1899 113 → 102). Studio's weekend recovered "
                    "partially (197 → 214/day) after July's dip. Tickets were stable — August was a traffic story, not a ticket story.")

    # ---------------- 3. Sales patterns (NEW)
    add_heading(doc, "3. Sales Patterns & Staffing Alignment", level=1)
    cal = pat["attach_calibration"]
    add_para(doc, "NEW. Source: payment-level Square transactions (walk-in only — invoices and $0 test transactions excluded) joined to daily FOH labor hours "
                  "from the weekly pipeline. This section answers three questions the P&L cannot: when in the day and week the revenue lands, whether labor hours "
                  "are placed where the revenue is, and what the basket looks like. Labor hours cover Aug 1–29 (Aug 30–31 timecards not yet pulled).",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    # 3a daypart
    add_heading(doc, "3a. Daypart Curve — Where the Day's Revenue Lands", level=2)
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
    add_caveat(doc, "Share of the month's walk-in net sales by daypart. Highlighted rows are the two dayparts that together carry ~60% of revenue at the "
                    "established cafes. The 7–10a window (Early + Rush) is 40% at Downtown vs 36% at Studio — Downtown is the closer fit to an "
                    "office-district profile (typical 45–60%), Studio to a neighborhood profile (35–45%).")
    # avg ticket + attach by daypart, Open3
    add_heading(doc, "Avg ticket, attach and tip rate by daypart — Open Cafes", level=3)
    o3 = pat["open3"]
    rows = [[dp, f"${o3['daypart'][dp]['avg_ticket']:.2f}" if o3['daypart'][dp]['avg_ticket'] else "—",
             pct_str(o3["attach_by_daypart"][dp]), pct_str(o3["daypart"][dp]["tip_rate"]),
             pct_str(o3["daypart"][dp]["share"]), pct_str(o3["daypart"][dp]["share_prior"])] for dp in DAYPARTS]
    build_investor_table(doc, ["Daypart", "Avg Ticket", "Attach (proxy)", "Tip Rate", "Share Aug", "Share Jul"], rows,
                         col_widths=[Inches(1.8), Inches(1.0), Inches(1.15), Inches(1.0), Inches(1.0), Inches(1.0)], key_rows=[2, 3])
    add_caveat(doc, "Attach and ticket both peak mid-morning and lunch and trough in the 8–10a rush — the rush guest buys one coffee and goes. "
                    "That is normal and not a problem to solve; the opportunity is the 10a–2p window where the guest is already inclined to add food.")

    # 3b DOW + labor
    add_heading(doc, "3b. Day of Week — Revenue vs Labor Hours", level=2)
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
                         f"{r['labor_hours_per_day']:.1f}" if r.get("labor_hours_per_day") else "—",
                         f"${r['sales_per_labor_hour']:.0f}" if r.get("sales_per_labor_hour") else "—",
                         pct_str(r["share"]), pct_str(r.get("labor_share")),
                         (f"+{r['alignment_gap']:.1f}" if (r.get("alignment_gap") or 0) > 0 else f"{r.get('alignment_gap', 0):.1f}") if r.get("alignment_gap") is not None else "—"])
        build_investor_table(doc, ["Day", "Sales/Day", "Txns/Day", "Avg Tkt", "Labor Hrs/Day", "Sales / Labor Hr", "Sales Share", "Labor Share", "Gap (pts)"], rows,
                             col_aligns=["left"] + ["right"] * 8,
                             col_widths=[Inches(0.55), Inches(0.85), Inches(0.7), Inches(0.7), Inches(0.9), Inches(0.95), Inches(0.8), Inches(0.8), Inches(0.7)])
    add_caveat(doc, "Gap = labor share of the week's hours minus sales share of the week's revenue. Positive = more labor than revenue would justify on that day; "
                    "negative = under-staffed relative to demand. A well-aligned schedule keeps every day within ±2 points. Sales / Labor Hr uses FOH hours from "
                    "Square timecards; the $60–80 range is where a 28–33% labor target lands at DC wage levels.")

    # 3c basket
    add_heading(doc, "3c. Basket — Ticket Size, Attach and Dining Mix", level=2)
    add_para(doc, NARRATIVE["insights_3c"], after=6)
    hdr = ["Metric"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"]
    def prow(label, fn):
        return [label] + [fn(pat["cafes"][c]) for c in CAFES] + [fn(pat["open3"])]
    rows = [
        prow("Avg ticket (walk-in)", lambda b: f"${b['avg_ticket']:.2f}"),
        prow("Median ticket", lambda b: f"${b['median_ticket']:.2f}"),
        prow(f"Attach proxy (tickets ≥ ${cal['threshold']:.2f})", lambda b: pct_str(b["attach_proxy"])),
        prow("  — prior month", lambda b: pct_str(b["attach_proxy_prior"])),
        prow("Tip rate", lambda b: pct_str(b["tip_rate"])),
        prow("Pickup share", lambda b: pct_str(b["dining"].get("PICKUP", {}).get("share"))),
        prow("  Pickup avg ticket", lambda b: f"${b['dining']['PICKUP']['avg_ticket']:.2f}" if "PICKUP" in b["dining"] else "—"),
        prow("  In-store avg ticket", lambda b: f"${b['dining']['IN_STORE']['avg_ticket']:.2f}" if "IN_STORE" in b["dining"] else "—"),
    ]
    build_investor_table(doc, hdr, rows, col_widths=[Inches(2.2)] + [Inches(0.95)] * 5, key_rows=[2, 5])
    add_heading(doc, "Ticket-size distribution — Open Cafes", level=3)
    rows = [[k, f"{v['txns']:,}", pct_str(v["share"])] for k, v in pat["open3"]["ticket_dist"].items()]
    build_investor_table(doc, ["Ticket", "Transactions", "Share"], rows, col_aligns=["left", "right", "right"], col_widths=[Inches(1.5), Inches(1.5), Inches(1.2)], key_rows=[1])
    add_caveat(doc, f"Attach proxy calibration: in April (the last export with full multi-item descriptions) {cal['april_multi_item_rate']}% of tickets had 2+ items; "
                    f"tickets ≥ ${cal['threshold']:.2f} matched that at {cal['april_rate_at_threshold']}%. The proxy is applied to Jul/Aug where descriptions are truncated to the first item. "
                    "The $5–8 band is the single-coffee ticket and is 45–50% of all transactions — the size of the attach opportunity.")

    # ---------------- 4. Items (from transactions)
    add_heading(doc, f"4. Item-Level Insights — {MONTH_LABEL}", level=1)
    add_para(doc, "Source: transaction lead item (first item on the ticket). The Jul/Aug Square exports truncate multi-item descriptions, so unit counts here "
                  "UNDER-count items in 2+ item baskets; the Square Item Sales CSV remains the source of truth for units. What this view adds that the CSV cannot: "
                  "which item anchors the purchase in each daypart.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, "4a. Lead item by daypart — Open Cafes", level=2)
    rows = []
    for dp in DAYPARTS[:5]:
        items = pat["open3"]["lead_item_by_daypart"].get(dp, [])
        rows.append([dp] + [f"{it['item']} ({it['txns']:,})" if i < len(items) else "" for i, it in enumerate(items[:4])] if items else [dp, "", "", "", ""])
    build_investor_table(doc, ["Daypart", "#1", "#2", "#3", "#4"], rows, col_aligns=["left"] * 5, col_widths=[Inches(1.6)] + [Inches(1.4)] * 4)
    add_heading(doc, "4b. Lead-item category share by daypart — Open Cafes", level=2)
    cats = ["coffee_bev", "food", "retail_coffee", "other"]
    rows = [[dp] + [pct_str(pat["open3"]["lead_category_by_daypart"].get(dp, {}).get(c)) for c in cats] for dp in DAYPARTS[:5]]
    build_investor_table(doc, ["Daypart", "Coffee / Bev", "Food", "Retail Bean", "Other"], rows, col_widths=[Inches(1.8)] + [Inches(1.2)] * 4, key_rows=[1])
    add_caveat(doc, "Food as the LEAD item is highest at lunch and mid-morning. Because only the first item is visible, a coffee-led ticket that also has a pastry shows as "
                    "'coffee' here — read this as 'what the guest came for', not a unit count.")
    add_heading(doc, "4c. Top 15 lead items — Open Cafes", level=2)
    rows = [[i + 1, it["item"], f"{it['txns']:,}"] for i, it in enumerate(pat["open3"]["top_lead_items"])]
    build_investor_table(doc, ["#", "Item", "Tickets led"], rows, col_aligns=["right", "left", "right"], col_widths=[Inches(0.45), Inches(3.8), Inches(1.4)])
    add_caveat(doc, "Send the August Square Item Sales CSVs (Downtown / Studio / 1899 / NMWA) and this section gains the full top/bottom-10 by category, pastry detail and the Filter Coffee / Premium Hot Filter view from prior months.")

    # ---------------- 5. Discounts (hold)
    add_heading(doc, f"5. Discount Report — {MONTH_LABEL}", level=1)
    add_para(doc, "On hold per Ryan. Discounts as booked: " + ", ".join(f"{CAFE_LABEL[c]} {fmt_money(-F[c]['disc_ref'], False)}" for c in OPEN3) +
                  f" — {fmt_money(-O['disc_ref'], False)} / {fmt_pct(-O['disc_ref'], O['core'])} of open-cafe core sales. The August transaction export does not carry discount names, so a by-type view still needs a Square Orders export.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------------- 6. Loyalty (slipped this month)
    add_heading(doc, f"6. Loyalty Report — {MONTH_LABEL}", level=1)
    add_para(doc, "Slipped this month per Ryan. The Square Loyalty dashboard screenshots only expose a few labeled figures; the rest "
                  "(visit-frequency buckets, spend distribution) are unlabeled bar charts and can't be read reliably. A better input path is "
                  "being set up for September — see the note in Section 8. Only 6% of August transactions carry a Customer ID in the export, "
                  "so repeat-guest analysis from transaction data is not yet a substitute.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------------- 7. Benchmarks + actions (NEW)
    add_heading(doc, "7. Benchmarks & Prioritized Actions", level=1)
    add_para(doc, "NEW. Each cafe against typical ranges for independent specialty coffee in high-cost US metros. The ranges are directional operator "
                  "benchmarks, not audited statistics — they are useful for knowing which number to work on first, not for judging a single month.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, "7a. Scorecard — August 2026", level=2)
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
        sc("Sales per FOH labor hour", "splh", lambda c: pat_v(c, lambda b: f"${b['labor']['sales_per_labor_hour']:.0f}" if b['labor']['sales_per_labor_hour'] else "—"), ""),
        sc("Transactions per labor hour", "tplh", lambda c: pat_v(c, lambda b: f"{b['labor']['txns_per_labor_hour']:.1f}" if b['labor']['txns_per_labor_hour'] else "—"), ""),
        sc("Avg ticket (walk-in)", "avg_ticket", lambda c: pat_v(c, lambda b: f"${b['avg_ticket']:.2f}"), ""),
        sc("Attach rate (proxy)", "attach", lambda c: pat_v(c, lambda b: pct_str(b["attach_proxy"])), ""),
        sc("7–10a share of sales", "morning_share", lambda c: pat_v(c, lambda b: pct_str(round((b['daypart']['Early (6-8a)']['share'] or 0) + (b['daypart']['Rush (8-10a)']['share'] or 0), 1))), ""),
        sc("Tip rate", "tip_rate", lambda c: pat_v(c, lambda b: pct_str(b["tip_rate"])), ""),
        sc("Pickup share", "pickup_share", lambda c: pat_v(c, lambda b: pct_str(b["dining"].get("PICKUP", {}).get("share"))), ""),
    ]
    build_investor_table(doc, ["Metric", "Typical range", "Downtown", "Studio", "1899 Penn", "Open Cafes"], rows,
                         col_aligns=["left", "left", "right", "right", "right", "right"],
                         col_widths=[Inches(1.9), Inches(1.0), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95)], key_rows=[0, 1, 4])
    add_caveat(doc, "Reading across the rows: sales per labor hour is the metric furthest from range at every cafe and is the mechanical cause of the labor % and prime cost misses. "
                    "Ticket, attach, tip and dining mix are all inside or close to typical ranges — the guest-facing business is fine; the schedule is not. "
                    "Benchmark notes: " + "; ".join(f"{k.replace('_', ' ')}: {v[1]}" for k, v in list(BENCH.items())[:6]) + ".")

    add_heading(doc, "7b. Prioritized actions", level=2)
    add_para(doc, "Ranked by estimated impact against effort. Each is specific enough to be assigned and checked in the next weekly report.", italic=True, size=9, color=COLOR_MUTED, after=6)
    for i, (title, body, impact) in enumerate(NARRATIVE["actions"], 1):
        add_bullet(doc, f"{body}  [Est. impact: {impact}]", bold_lead=f"{i}. {title}")

    # ---------------- 8. Notes
    add_heading(doc, "8. Notes & Watch Items", level=1)
    add_heading(doc, "Data caveats", level=3)
    add_bullet(doc, f"Financials from Consolidated Model {MODEL_VER} (recalculated before reading; 18,536 formulas, zero errors). The v14 rebuild's corrected 'TOTAL NET INCOME (all entities)' rollup is used directly — no manual restatement.")
    add_bullet(doc, "Section 3 uses walk-in transactions only (invoice-sourced and $0 test transactions excluded). Transaction net sales reconcile to the model's core sales within 3–5% per cafe; the difference is timing (deposit vs order date) and refunds netting. Ratios and shares are unaffected.")
    add_bullet(doc, "Labor hours cover Aug 1–29 from Square timecards; Aug 30–31 will be added when the next weekly pull runs. Day-of-week alignment for Sat/Sun at Studio and NMWA is slightly understated as a result.")
    add_bullet(doc, "Jul/Aug transaction exports truncate multi-item descriptions to the first item. Attach rate is therefore a calibrated proxy (see 3c) and Section 4 counts lead items, not units.")
    add_bullet(doc, "Section 6 (Loyalty) slipped this month — screenshot input is too lossy. Section 5 (Discounts) on hold.")
    add_heading(doc, "Watch items for September", level=3)
    for lead, body in NARRATIVE["watch"]:
        add_bullet(doc, body, bold_lead=lead)

    # ---------------- Appendices
    doc.add_paragraph().add_run().add_break()
    p = doc.add_paragraph(); r = p.add_run("LEADERSHIP TEAM ONLY — APPENDICES"); _set_run_font(r, size=12, bold=True, color=COLOR_MUTED); _set_para_spacing(p, before=10, after=6)
    S = fin["support"]
    add_heading(doc, f"A. Support Entities — {MONTH_LABEL}", level=2)
    rows = [["Revenue", fmt_money(S["Commissary"][m]["rev"]), "$0", fmt_money(S["Cafe Mgmt"][m]["rev"]), fmt_money(S["SRP"][m]["rev"]), fmt_money(S["Commissary"][m]["rev"] + S["Cafe Mgmt"][m]["rev"] + S["SRP"][m]["rev"])],
            ["COGS", fmt_money(S["Commissary"][m]["cogs"]), "$0", "$0", fmt_money(S["SRP"][m]["cogs"]), fmt_money(S["Commissary"][m]["cogs"] + S["SRP"][m]["cogs"])],
            ["Total Labor", fmt_money(S["Commissary"][m]["labor"]), fmt_money(S["Central"][m]["labor"]), "$0", fmt_money(S["SRP"][m]["labor"]), fmt_money(S["Commissary"][m]["labor"] + S["Central"][m]["labor"] + S["SRP"][m]["labor"])],
            ["Expenses", fmt_money(S["Commissary"][m]["opex"]), fmt_money(S["Central"][m]["opex"]), "$0", fmt_money(S["SRP"][m]["opex"]), fmt_money(S["Commissary"][m]["opex"] + S["Central"][m]["opex"] + S["SRP"][m]["opex"])],
            ["Net Income", fmt_money(S["Commissary"][m]["ni"]), fmt_money(S["Central"][m]["ni"]), fmt_money(S["Cafe Mgmt"][m]["ni"]), fmt_money(S["SRP"][m]["ni"]), fmt_money(sum(S[e][m]["ni"] for e in S))]]
    build_investor_table(doc, ["Metric", "Commissary", "Central Admin", "Cafe Mgmt", "SRP (Roasting)", "Support Total"], rows, col_widths=[Inches(1.5), Inches(1.15), Inches(1.2), Inches(1.05), Inches(1.25), Inches(1.15)], total_rows=[4])
    add_heading(doc, "A1. Support Entities — Net Income Trend", level=3)
    tm = MONTHS[-6:]
    rows = [[e] + [fmt_money(S[e][mm]["ni"]) for mm in tm] + [fmt_money(S[e][m]["ni"] - S[e][pm]["ni"])] for e in ["Commissary", "Central", "Cafe Mgmt", "SRP"]]
    build_investor_table(doc, ["Entity"] + tm + ["MoM Δ"], rows, col_widths=[Inches(1.35)] + [Inches(0.9)] * 7)
    add_caveat(doc, f"Central Admin {fmt_money(S['Central'][pm]['ni'])} → {fmt_money(S['Central'][m]['ni'])} includes ~$5,700 of one-time tax prep and tax payment (per Ryan). Normalized: ~($14,500).")

    add_heading(doc, f"B. Consolidated P&L — {MONTH_LABEL}", level=2)
    C = fin["consolidated"]
    tm = MONTHS[-6:]
    def crow(label, key):
        return [label] + [fmt_money(C[mm][key]) for mm in tm] + [delta_pct(C[m][key], C[pm][key])]
    rows = [crow("Total Core Location Sales", "tcls"), crow("Total COGS", "tcogs"), crow("Gross Profit", "gp"), crow("Total Labor", "labor"),
            crow("Total Expenses", "texp"), crow("Net Operating Income", "noi"), crow("Cafes — Net Income", "cafes_ni"), crow("TOTAL NET INCOME (all entities)", "total_ni")]
    build_investor_table(doc, ["Metric"] + tm + ["MoM Δ"], rows, col_widths=[Inches(2.1)] + [Inches(0.85)] * 6 + [Inches(0.8)], key_rows=[5, 6], total_rows=[7])
    add_caveat(doc, f"Total Net Income {fmt_money(C[m]['total_ni'])} — first negative month since February. Cafes contributed +{fmt_money(C[m]['cafes_ni'], False)}; "
                    f"support entities −{fmt_money(-sum(S[e][m]['ni'] for e in S), False)}, of which Central Admin is {fmt_money(S['Central'][m]['ni'])} (~$5,700 one-time tax items; normalized consolidated NI ≈ ($7,000)).")

    add_heading(doc, "C. Methodology & Data Sources", level=2)
    add_bullet(doc, f"Financials: Consolidated Model {MODEL_VER} → data/fin_2026_v17.json (extract of Jan–Aug per-entity IS tabs, Consolidated, Projections).")
    add_bullet(doc, "Sales patterns: _pipeline/shared/transactions/2026-08.csv (Square Transactions export) + weekly labor JSONs → analyze_patterns.py → data/2608_patterns.json.")
    add_bullet(doc, "Report: build_monthly_report_v2.py reads both JSONs; narrative blocks live in NARRATIVE at the top of that file.")
    add_bullet(doc, "Definitions: Core Sales = Total Core Location Sales − Catering. Core COGS = Total COGS − Catering COGS. Prime = Core COGS + Total Labor. Open Cafes = Downtown + Studio + 1899.")
    add_bullet(doc, "Daypart boundaries: Early 6–8a, Rush 8–10a, Mid-morning 10a–12p, Lunch 12–2p, Afternoon 2–5p, Late 5p+. Alignment gap = labor-hour share − sales share by day of week.")

    _ensure_zoom_percent(doc)
    doc.save(OUT)
    print("Saved →", OUT)


if __name__ == "__main__":
    build()
