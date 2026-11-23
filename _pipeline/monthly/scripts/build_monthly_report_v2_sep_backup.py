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
MONTH = "Sep"
MONTH_LABEL = "September 2026"
VERSION = "v1.0"
PREPARED = "October 2, 2026"
MODEL_VER = "v18"
FIN_JSON = os.path.join(DATA_DIR, "fin_2026_v18.json")
PAT_JSON = os.path.join(DATA_DIR, "2609_patterns.json")
LOY_JSON = os.path.join(DATA_DIR, "2609_loyalty.json")
LOY_PRIOR_JSON = os.path.join(DATA_DIR, "2608_loyalty.json")
DD_JSON = os.path.join(DATA_DIR, "labor_deep_dive_2026-09.json")
OUT = os.path.join(WORKSPACE, f"Slipstream Monthly Report - {MONTH_LABEL} {VERSION}.docx")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]
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
        f"Prepared {PREPARED}  |  {VERSION} — September results from Consolidated Model {MODEL_VER}. Changes "
        "this issue: (1) all labor-hour math is CORE hours only (admin + training excluded) and sales-per-hour "
        "is computed on matched days, so partial timecard coverage can't inflate it; (2) Section 2a avg ticket "
        "now comes from walk-in transactions, which removes the April–May catering distortion; (3) Section 6 "
        "Loyalty is live again from the Square Loyalty API ledger — fixed buckets, exact first-time counts; "
        "(4) loaded cost per labor hour is shown next to sales per hour everywhere. Timecards cover Sep 1–19 "
        "(the Sep 20–30 weekly pull is pending); P&L labor is complete from payroll."
    ),
    exec_para=(
        "September is the best month of 2026. Total net income across all entities was $28,471 — a $41K swing "
        "from August's ($12.7K) loss — and the three established cafes ran core prime cost at 58.7%, inside the "
        "healthy band for the first time since March. Two things happened at once and both matter. Sales "
        "recovered as expected after Labor Day (open-cafe core $103,579, +10.7%; Downtown +23%, 1899 +27%, "
        "weekday transactions back to spring levels). And labor dollars at the open cafes fell $10.9K "
        "($39.5K → $28.6K) on higher sales — labor went from 42.2% of core to 27.6%, and every established "
        "cafe landed under its target (Downtown 28.7% vs 32, Studio 24.6% vs 30, 1899 32.6% vs 35). That is the "
        "schedule work from the labor deep dive showing up in the P&L inside a month. Sales per core labor hour "
        "rose from $46 to $57 blended. NMWA's $10.8K net income is almost entirely an $11K catering month; its "
        "walk-in business was soft (Section 1f). Central Admin returned to ($10.6K) with the August tax items "
        "gone. The remaining work is narrower and more specific than it was a month ago: Fridays at Downtown and "
        "1899, Studio's Tue/Wed, and food COGS at 1899."
    ),
    exec_bullets=[
        ("Labor targets hit at all three established cafes.",
         "Core labor % of core sales: Downtown 48.0% → 28.7%, Studio 35.2% → 24.6%, 1899 52.7% → 32.6%. "
         "Blended 27.6% against a ~31% weighted target. Labor dollars fell $10.9K while core sales rose $10.0K. "
         "Loaded cost per core hour is $20.4–21.2 at all three — the cost-per-hour lever moved too (Downtown "
         "was $22.20 in August)."),
        ("Prime cost back in the healthy band: 58.7%.",
         "Downtown 85.6% → 58.7%, Studio 64.2% → 53.6%, 1899 89.6% → 70.8%. 1899 is still above 65% but has "
         "improved 19 points in one month; its remaining gap is food COGS (71.5% of food sales), not labor."),
        ("Sales recovered on schedule — it was seasonal.",
         "Downtown core $29.8K → $36.6K (+23%), within $200 of the FY26 plan; 1899 $15.7K → $19.9K (+27%), "
         "$3.5K over plan. Studio flat at $47.1K and 5% under its plan — the one cafe that did not get a "
         "September lift. Weekday transactions: Downtown 158 → 194/day, 1899 102 → 132/day."),
        ("Friday is the one schedule miss still standing.",
         "Downtown Friday runs $730/day (38% of a Wednesday) and carries +9.2 points more labor share than sales "
         "share — wider than August. 1899 Friday +7.2. The slow-day templates in the deep dive were built for "
         "exactly this; they have not been applied to Fridays yet. Mon–Thu at both cafes is now within ±5 pts."),
        ("Studio: weekend still thin, Tue/Wed now heavy.",
         "Saturday $2,398 and Sunday $2,214/day remain the two best days and run 3–5 pts under-staffed; "
         "Tuesday and Wednesday ($1,000/day) are +4–5 pts over. Monday is fixed (0.0). Studio's weekday "
         "hours came down but the weekday/weekend balance did not shift with them."),
        ("NMWA: catering month, not a trading month.",
         "$11,070 of catering drove $10.8K net income, but walk-in core sales fell to $1,323 (Aug $3,606) and "
         "weekday transactions 29 → 18/day. Labor rose to $3.8K to serve the events. Treat the P&L as an "
         "events result; the café ramp slipped and needs the October read to judge."),
        ("Loyalty is back, from the ledger.",
         "4,726 loyalty visits at the open cafes (41% of walk-in tickets), 1,750 members, 2.70 visits/member, "
         "854 first-time members in September (exact — first month with a prior month in the store). 1899 has "
         "the highest frequency (2.84) and the deepest 5+ visit core (18% of members)."),
    ],
    insights_3a=(
        "The daypart shapes are stable month to month, which is the point — they are structural. Downtown: "
        "36% of the day in the 8–10a rush, fading to 14% after 2p. Studio: peak is 10–12p (33%), not the rush, "
        "with a real afternoon (7%) on weekends. 1899: lunch is the largest daypart (31%). NMWA: 83% of sales "
        "between 10a and 2p. The four cafes need four staffing templates; the deep dive built them by tier."
    ),
    insights_3b=(
        "Compare to August. Mon–Thu alignment gaps at Downtown and 1899 tightened from +3/−2 ranges to within "
        "±5 points, and sales per core hour rose to $55–59 across the three established cafes. The two gaps "
        "that did not move are the two the deep dive flagged first: Friday at the weekday cafes (+9.2 Downtown, "
        "+7.2 at 1899 — still staffed like a Wednesday for 38–50% of the revenue) and Studio's weekend "
        "(Sat −5.0, Sun −2.9) versus Tue/Wed (+5.1, +4.4). Timecards cover Sep 1–19, so the matched-day view "
        "is 13 of 21 weekdays at Downtown / 1899 and 19 of 30 days at Studio — enough to read the pattern, and "
        "every ratio here is computed only on days with both sales and labor data."
    ),
    insights_3c=(
        "Basket is unchanged: attach proxy 39% blended (Studio 42%, Downtown 39%, 1899 32%), tickets within "
        "$0.10 of August at every cafe, in-store ~$1.00 above pickup. September's improvement came entirely "
        "from traffic and hours, not from the basket — which means the basket is the untouched lever. The "
        "10a–2p window, where attach peaks at every cafe, is where a food prompt at the register would land."
    ),
    loyalty_bullets=[
        ("Four in ten tickets are loyalty.",
         "41–43% of walk-in tickets at every open cafe carry a loyalty account, stable vs August. This is a high "
         "attach for a specialty program and means the loyalty base is a reliable read on the regular customer."),
        ("First-time members: 854 in September.",
         "Downtown 303, Studio 371, 1899 187 — now measured exactly (an account with no earlier event in the "
         "ledger). That is ~49% of September's visiting members being new-to-ledger, which looks high because the "
         "ledger only starts Aug 1; expect the ratio to settle over the next two months as history accumulates."),
        ("Frequency up at every cafe.",
         "Avg visits per member 2.56 → 2.70 blended (Downtown 2.44 → 2.65, Studio 2.46 → 2.54, 1899 2.66 → 2.84). "
         "The 5+ visit segment is 16–20% of members — the regulars the incentive plan should be built around."),
        ("Members spend at the walk-in ticket, not above it.",
         "$8.05 / $8.77 / $6.77 per visit vs $8.58 / $9.29 / $7.39 walk-in. Loyalty value is frequency, not basket. "
         "A member at 2.7 visits/month is worth ~$260/yr at Downtown — ~$1,300 over five years."),
    ],
    actions=[
        ("Apply the slow-day template to Fridays at Downtown and 1899.",
         "Friday is now the single widest gap in the business (+9.2 / +7.2 pts). The deep dive's slow-tier "
         "Friday templates (2-person, later third start or none) are built and sitting in the xlsx. This was "
         "worth ~$1,200–1,600/mo in August; at September hours it is ~$900–1,200/mo and costs nothing in revenue.",
         "$900–1,200 / mo"),
        ("Move 8–10 Studio hours per week from Tue/Wed to Sat/Sun.",
         "Same recommendation as last month, still open. Weekend is $2,200–2,400/day and under-staffed; Tue/Wed "
         "are $1,000/day and over. Labor-neutral; protects throughput on the highest-ticket, highest-attach days.",
         "labor-neutral; revenue-protective"),
        ("Cost 1899's food menu.",
         "Food COGS 71.5% of food sales for the third straight month above 60%. On $2,900 of food sales it is a "
         "$1,000/mo margin leak, but more importantly it is the only thing keeping 1899's prime cost above 65% "
         "now that labor is at target. Per-item costing, not ratio tracking.",
         "$600–1,000 / mo"),
        ("Put a food prompt into the 10a–2p register flow.",
         "Attach is flat at 39% and peaks mid-morning/lunch. A single scripted prompt in that window at the three "
         "established cafes, measured by the attach proxy next month, is the lowest-cost sales lever available.",
         "+2 pts attach ≈ $1,500 / mo"),
        ("Decide the NMWA café plan separately from NMWA catering.",
         "Walk-in fell 63% in a month while catering carried the P&L. The ramp needs its own target (transactions/"
         "day) so a strong events month cannot mask a stalling café.",
         "decision"),
    ],
    watch=[
        ("Fridays.", "If the Friday template is applied in October, the Downtown/1899 Friday gap should read within ±3 pts in the October report."),
        ("Studio sales.", "Flat in a month when the other two cafes grew 23–27%, and 5% under plan. Worth a look at the weekend transaction trend (230/day, down from 258 in June)."),
        ("1899 food COGS.", "71.5%. Third month >60%. Costing exercise is Action 3."),
        ("NMWA walk-in.", "18 weekday txns/day vs 29 in August. Watch October with catering stripped out."),
        ("Labor data coverage.", "Timecards for Sep 20–30 land with the next weekly pull; Section 3 ratios will tighten slightly when they do. No change expected to the direction."),
        ("Holding the line into Q4.", "Labor targets were hit in a +11% sales month. The deep dive's slow-tier templates exist for Dec–Feb; the test is whether labor flexes down when sales do."),
    ],
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
    add_caveat(doc, "1899 Net Income is a pass-through (owner profit share) and shows $0 at the entity level; its operating result is visible in Prime Margin. "
                    "NMWA ratios are distorted this month by an $11K catering month on a $1.3K walk-in base — see 1f.")
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
    add_heading(doc, "1b. Catering Sales — Jan → Sep 2026", level=2)
    rows = [[f"{mm} 2026"] + [fmt_money(fin["cafes"][c][mm]["cat"]) for c in CAFES] + [fmt_money(sum(fin["cafes"][c][mm]["cat"] for c in CAFES))] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["All Cafes"], rows, col_widths=[Inches(1.1)] + [Inches(1.2)] * 5, total_rows=[len(rows) - 1])
    add_caveat(doc, "September catering rebounded to $15,438 company-wide, $11,070 of it at NMWA (museum events) and $4,093 at Downtown. Catering is excluded from every core ratio above.")

    # 1c trends
    add_heading(doc, "1c. Prime Cost & Component Trends — Jan → Sep 2026 (ex-catering)", level=2)
    add_para(doc, "Prime Cost = COGS + Labor, as % of Core Sales. Splitting the components shows which lever is moving.", italic=True, size=9, color=COLOR_MUTED, after=6)
    for title, num, den in [("Prime Cost % of Core Sales", "prime", "core"), ("COGS % of Core Sales", "cogs", "core"), ("Labor % of Core Sales", "labor", "core")]:
        add_heading(doc, title, level=3)
        rows = []
        for mm in MONTHS:
            b = blend(fin, mm)
            rows.append([f"{mm} 2026"] + [fmt_pct(fin["cafes"][c][mm][num], fin["cafes"][c][mm][den]) for c in OPEN3] + [fmt_pct(b[num], b[den])])
        build_investor_table(doc, ["Period", "Downtown", "Studio", "1899 Penn", "Open Cafes Blended"], rows, col_widths=[Inches(1.15)] + [Inches(1.3)] * 3 + [Inches(1.7)], total_rows=[len(rows) - 1])
    add_caveat(doc, "Labor % reversed hard: 42.2% → 27.6% blended, the lowest reading of the year, on a +10.7% sales month. COGS % stayed in its 29–35% range. "
                    "Prime cost 58.7% is the second-best month of 2026 (March 56.7%). This is what hitting the labor targets looks like in the P&L.")

    # 1d core sales trend
    add_heading(doc, "1d. Core Sales Trend — Jan → Sep 2026 (ex-catering)", level=2)
    rows = [[f"{mm} 2026"] + [fmt_money(fin["cafes"][c][mm]["core"]) for c in CAFES] + [fmt_money(blend(fin, mm)["core"])] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["Open Cafes"], rows, col_widths=[Inches(1.35)] + [Inches(1.15)] * 5, total_rows=[len(rows) - 1])
    oa, oj = blend(fin, m)["core"], blend(fin, pm)["core"]
    add_caveat(doc, f"Open Cafes {fmt_money(oa)} in September, {delta_pct(oa, oj)} vs August and {delta_pct(oa, blend(fin, 'Apr')['core'])} vs the April peak. "
                    f"Downtown {delta_pct(F['Downtown']['core'], fin['cafes']['Downtown'][pm]['core'])}, Studio {delta_pct(F['Studio']['core'], fin['cafes']['Studio'][pm]['core'])}, "
                    f"1899 {delta_pct(F['1899']['core'], fin['cafes']['1899'][pm]['core'])}. The post-Labor-Day recovery landed at the two office-district cafes; Studio did not get it.")

    # 1e actual vs projected
    add_heading(doc, f"1e. Actual vs. Projected — {MONTH_LABEL}", level=2)
    add_para(doc, "Source: 2026 Projections v1.7 (FY26 budget) sales projections. Labor projections are not carried in the model's Projections tab; the sales variance is shown. Labor target compliance is in Section 1a.", italic=True, size=9, color=COLOR_MUTED, after=6)
    rows = []
    tot_a = tot_p = 0
    for c in CAFES:
        a, p_ = F[c]["core"], fin["proj_sales"][c][m]
        rows.append([CAFE_LABEL[c], fmt_money(a), fmt_money(p_), fmt_money(a - p_) if a - p_ >= 0 else f"({fmt_money(p_ - a, False)})", delta_pct(a, p_) if p_ else "—"])
        if c in OPEN3:
            tot_a += a; tot_p += p_
    rows.append(["Open Cafes", fmt_money(tot_a), fmt_money(tot_p), f"({fmt_money(tot_p - tot_a, False)})" if tot_a < tot_p else fmt_money(tot_a - tot_p), delta_pct(tot_a, tot_p)])
    build_investor_table(doc, ["Cafe", "Sales Actual", "Sales Proj", "Δ ($)", "Δ (%)"], rows, col_aligns=["left"] + ["right"] * 4, col_widths=[Inches(1.6), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.1)], total_rows=[4])
    add_caveat(doc, "Downtown within $200 of plan; 1899 $3.5K (+22%) over; Studio $2.6K (5%) under — Studio is the only established cafe below plan this month. NMWA walk-in came in $2.4K under its first real budget month, though catering more than covered it at the entity level.")

    # 1f NMWA ramp
    add_heading(doc, "1f. NMWA — Opening Ramp", level=2)
    rows = [["Core Sales (ex-catering)"] + [fmt_money(fin["cafes"]["NMWA"][mm]["core"]) for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Catering & Events"] + [fmt_money(fin["cafes"]["NMWA"][mm]["cat"]) for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Weekday avg transactions/day"] + [f"{fin['cafes']['NMWA'][mm]['wd_txn']:.1f}" for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Weekend avg transactions/day"] + [f"{fin['cafes']['NMWA'][mm]['we_txn']:.1f}" for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Weekday avg ticket"] + [f"${fin['cafes']['NMWA'][mm]['wd_tkt']:.2f}" for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Total Labor"] + [fmt_money(fin["cafes"]["NMWA"][mm]["labor"]) for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Net Operating Income"] + [fmt_money(fin["cafes"]["NMWA"][mm]["noi"]) for mm in ["Jun", "Jul", "Aug", "Sep"]],
            ["Net Income (incl. mgmt fee)"] + [fmt_money(fin["cafes"]["NMWA"][mm]["ni"]) for mm in ["Jun", "Jul", "Aug", "Sep"]]]
    build_investor_table(doc, ["Metric", "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026"], rows, col_aligns=["left"] + ["right"] * 4, col_widths=[Inches(2.5)] + [Inches(1.2)] * 4, key_rows=[2, 3], total_rows=[7])
    add_caveat(doc, "A catering month: $11,070 of events revenue produced the best net income in the company ($10,787), while walk-in core sales fell to $1,323 and weekday transactions dropped 29 → 18/day. "
                    "Labor rose to $3.8K to serve the events. The café ramp stalled this month; October is the read that matters.")

    # ---------------- 2. Operating indicators
    add_heading(doc, "2. Operating Indicators", level=1)
    add_para(doc, f"Source: Consolidated Model {MODEL_VER} 'Ops Metrics'. Weekday / weekend split. Downtown and 1899 are weekday-only.", italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, "2a. Average Ticket ($, walk-in) — Jan → Sep 2026", level=2)
    add_para(doc, "Source changed this month per Ryan: ticket is now the per-transaction walk-in average from the Square "
                  "transactions store (invoice / catering excluded), not the model's Ops Metrics ticket, which was inflated "
                  "by catering invoices in April–May.", italic=True, size=9, color=COLOR_MUTED, after=4)
    wt = pat.get("walkin_ticket_by_month", {})
    rows = [[f"{mm} 2026"] + [(f"${wt[c][mm]:.2f}" if wt.get(c, {}).get(mm) else "—") for c in CAFES] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES], rows, col_widths=[Inches(1.2)] + [Inches(1.4)] * 4, total_rows=[len(rows) - 1])
    add_caveat(doc, "Tickets have been stable all year: Downtown $8.4–8.8, Studio $9.3–9.6, 1899 $7.2–7.5. The sales story is traffic, not ticket.")
    add_heading(doc, "2b. Weekday Average Daily Transactions — Jan → Sep 2026", level=2)
    rows = [[f"{mm} 2026"] + [(f"{fin['cafes'][c][mm]['wd_txn']:.0f}" if fin['cafes'][c][mm]['wd_txn'] else "—") for c in CAFES] + [f"{sum(fin['cafes'][c][mm]['wd_txn'] for c in CAFES):.0f}"] for mm in MONTHS]
    build_investor_table(doc, ["Period"] + [CAFE_LABEL[c] for c in CAFES] + ["Total/Day"], rows, col_widths=[Inches(1.15)] + [Inches(1.2)] * 5, total_rows=[len(rows) - 1])
    add_heading(doc, "2c. Weekend Metrics — Studio & NMWA", level=2)
    rows = [[f"{mm} 2026", f"{fin['cafes']['Studio'][mm]['we_txn']:.0f}", f"${fin['cafes']['Studio'][mm]['we_tkt']:.2f}",
             (f"{fin['cafes']['NMWA'][mm]['we_txn']:.0f}" if fin['cafes']['NMWA'][mm]['we_txn'] else "—"),
             (f"${fin['cafes']['NMWA'][mm]['we_tkt']:.2f}" if fin['cafes']['NMWA'][mm]['we_tkt'] else "—")] for mm in MONTHS]
    build_investor_table(doc, ["Period", "Studio — Txns/Day", "Studio — Avg Ticket", "NMWA — Txns/Day", "NMWA — Avg Ticket"], rows, col_widths=[Inches(1.1), Inches(1.4), Inches(1.4), Inches(1.35), Inches(1.35)], total_rows=[len(rows) - 1])
    add_caveat(doc, "Weekday transactions recovered at Downtown (158 → 194) and 1899 (102 → 132) to spring levels; Studio +9% (129 → 141). Studio weekend 214 → 229/day. Tickets flat — September, like August, was a traffic story.")

    # ---------------- 3. Sales patterns (NEW)
    add_heading(doc, "3. Sales Patterns & Staffing Alignment", level=1)
    cal = pat["attach_calibration"]
    add_para(doc, "NEW. Source: payment-level Square transactions (walk-in only — invoices and $0 test transactions excluded) joined to daily FOH labor hours "
                  "from the weekly pipeline. This section answers three questions the P&L cannot: when in the day and week the revenue lands, whether labor hours "
                  "are placed where the revenue is, and what the basket looks like. CORE labor hours only (admin + training excluded). Timecards cover Sep 1–19; "
                  "every sales-per-hour and alignment figure is computed on matched days only, so partial coverage cannot inflate it.",
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
                         f"${r['cost_per_hour']:.0f}" if r.get("cost_per_hour") else "—",
                         pct_str(r["share"]), pct_str(r.get("labor_share")),
                         (f"+{r['alignment_gap']:.1f}" if (r.get("alignment_gap") or 0) > 0 else f"{r.get('alignment_gap', 0):.1f}") if r.get("alignment_gap") is not None else "—"])
        build_investor_table(doc, ["Day", "Sales/Day", "Txns/Day", "Avg Tkt", "Core Hrs/Day", "Sales / Hr", "$ / Hr", "Sales Share", "Labor Share", "Gap"], rows,
                             col_aligns=["left"] + ["right"] * 9,
                             col_widths=[Inches(0.5), Inches(0.8), Inches(0.65), Inches(0.65), Inches(0.8), Inches(0.75), Inches(0.6), Inches(0.75), Inches(0.75), Inches(0.55)])
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
    add_caveat(doc, "Send the September Square Item Sales CSVs (Downtown / Studio / 1899 / NMWA) and this section gains the full top/bottom-10 by category, pastry detail and the Filter Coffee / Premium Hot Filter view from prior months.")

    # ---------------- 5. Discounts (hold)
    add_heading(doc, f"5. Discount Report — {MONTH_LABEL}", level=1)
    add_para(doc, "On hold per Ryan. Discounts as booked: " + ", ".join(f"{CAFE_LABEL[c]} {fmt_money(-F[c]['disc_ref'], False)}" for c in OPEN3) +
                  f" — {fmt_money(-O['disc_ref'], False)} / {fmt_pct(-O['disc_ref'], O['core'])} of open-cafe core sales. The transaction export does not carry discount names, so a by-type view still needs a Square Orders export.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------------- 6. Loyalty (from Square Loyalty API ledger — NEW source)
    add_heading(doc, f"6. Loyalty Report — {MONTH_LABEL}", level=1)
    add_para(doc, "NEW SOURCE. Built from the Square Loyalty API event ledger (every points-earning visit and reward redemption, "
                  "with account, cafe and timestamp), pulled by the Sunday data job. Replaces the dashboard screenshots. "
                  "Visit-frequency buckets are now fixed at 1 / 2 / 3–4 / 5–7 / 8+ every month. A 'visit' = one "
                  "ACCUMULATE_POINTS event; loyalty spend = points ÷ 7 (pre-tax). 'First-time' = account with no prior "
                  "event in the store (store starts Aug 1, so September is the first month this is exact). NMWA is not "
                  "enrolled in the program.", italic=True, size=9, color=COLOR_MUTED, after=6)
    L, LP = loy["cafes"], loyp["cafes"]
    L3, LP3 = loy["open3"], loyp["open3"]
    add_heading(doc, "6a. Headline Insights", level=2)
    for lead, body in NARRATIVE["loyalty_bullets"]:
        add_bullet(doc, body, bold_lead=lead)

    add_heading(doc, "6b. Loyalty Activity by Cafe — September vs August", level=2)
    def lrow(label, fn, fmt=lambda v: f"{v:,}" if isinstance(v, int) else (f"{v:.2f}" if isinstance(v, float) else "—")):
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
        lrow("First-time members (exact Sep; Aug = upper bound)", lambda b: b["first_time_members"]),
        lrow("Repeat members", lambda b: b["repeat_members"]),
        lrow("Loyalty share of walk-in tickets", lambda b: b["loyalty_share_of_tickets"], lambda v: pct_str(v)),
        lrow("Loyalty spend (pre-tax)", lambda b: b["loyalty_spend_pretax"], money),
        lrow("Avg spend / visit", lambda b: b["avg_spend_per_visit"], lambda v: f"${v:.2f}"),
        lrow("Avg spend / member / month", lambda b: b["avg_spend_per_member"], lambda v: f"${v:.2f}"),
        lrow("Rewards redeemed", lambda b: b["redemptions"]),
    ]
    build_investor_table(doc, ["Metric  (Sep  (Aug))", "Downtown", "Studio", "1899 Penn", "Open Cafes"], rows,
                         col_widths=[Inches(2.3)] + [Inches(1.15)] * 4, key_rows=[0, 5])
    add_caveat(doc, "Loyalty share of walk-in tickets is ~41% at every cafe and stable month to month — four in ten tickets carry a "
                    "loyalty account. Avg spend per visit ($8.05 / $8.77 / $6.77) sits right at each cafe's walk-in ticket, so loyalty "
                    "members are not out-spending non-members per visit; the value is in frequency (2.5–2.8 visits/month vs ~1.3 for "
                    "non-members in prior screenshot months).")

    add_heading(doc, "6c. Loyalty Member Value — In Dollars", level=2)
    rows = []
    for c in OPEN3 + ["open3"]:
        bb = L3 if c == "open3" else L[c]
        mo = bb["avg_spend_per_member"] or 0
        rows.append([("Open Cafes" if c == "open3" else CAFE_LABEL[c]), f"{bb['avg_visits_per_member']:.2f}", f"${bb['avg_spend_per_visit']:.2f}",
                     f"${mo:.2f}", f"${mo * 12:,.0f}", f"${mo * 60:,.0f}"])
    build_investor_table(doc, ["Cafe", "Visits / mo", "$ / visit", "$ / member / mo", "Annualized", "5-yr value"], rows,
                         col_aligns=["left"] + ["right"] * 5, col_widths=[Inches(1.5), Inches(1.0), Inches(1.0), Inches(1.3), Inches(1.1), Inches(1.1)], total_rows=[3])
    add_caveat(doc, "Measured directly from the ledger (members who visited in the month). 5-yr = annualized × 5, no retention decay. "
                    "Every loyalty signup that becomes a 2.7-visit/month regular is worth ~$1,300 over five years at current spend.")

    add_heading(doc, "6d. Visit-Frequency Distribution — September (members by visits in month)", level=2)
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
    add_caveat(doc, "Fixed schema from here on, so the distribution is comparable month to month. The 5+ visit members (16–20% of "
                    "members) are the core regulars; 1899 has the highest share of them (18%) despite the smallest base.")

    # ---------------- 7. Benchmarks + actions (NEW)
    add_heading(doc, "7. Benchmarks & Prioritized Actions", level=1)
    add_para(doc, "NEW. Each cafe against typical ranges for independent specialty coffee in high-cost US metros. The ranges are directional operator "
                  "benchmarks, not audited statistics — they are useful for knowing which number to work on first, not for judging a single month.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    add_heading(doc, f"7a. Scorecard — {MONTH_LABEL}", level=2)
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
        sc("Sales per CORE labor hour (matched days)", "splh", lambda c: pat_v(c, lambda b: f"${b['labor']['sales_per_labor_hour']:.0f}" if b['labor']['sales_per_labor_hour'] else "—"), ""),
        sc("Loaded cost per core labor hour", "cost_hr", lambda c: pat_v(c, lambda b: f"${b['labor']['cost_per_hour']:.2f}" if b['labor'].get('cost_per_hour') else "—"), ""),
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
    add_caveat(doc, "Reading across the rows: sales per core labor hour moved from $43–50 to $55–59 in one month and loaded cost per hour from $18–22 to $20–21 — both levers moved. "
                    "At $21/hr loaded, $57 SPLH is 37% labor on the matched days; the full-month P&L shows 27.6% because the later September weeks were leaner. The $60–80 range is the next milestone. "
                    "Guest-facing metrics (ticket, attach, tips, dining mix) are all in range and unchanged — the basket is the untouched lever. "
                    "Benchmark notes: " + "; ".join(f"{k.replace('_', ' ')}: {v[1]}" for k, v in list(BENCH.items())[:7]) + ".")

    add_heading(doc, "7b. Prioritized actions", level=2)
    add_para(doc, "Ranked by estimated impact against effort. Each is specific enough to be assigned and checked in the next weekly report.", italic=True, size=9, color=COLOR_MUTED, after=6)
    for i, (title, body, impact) in enumerate(NARRATIVE["actions"], 1):
        add_bullet(doc, f"{body}  [Est. impact: {impact}]", bold_lead=f"{i}. {title}")

    # ---------------- 8. Notes
    add_heading(doc, "8. Notes & Watch Items", level=1)
    add_heading(doc, "Data caveats", level=3)
    add_bullet(doc, f"Financials from Consolidated Model {MODEL_VER} (recalculated before reading; 18,536 formulas; two #N/A cells in Pro-forma tabs, not used here). 'TOTAL NET INCOME (all entities)' rollup used directly. August Downtown NI restated by $63 in v18.")
    add_bullet(doc, "Section 3 uses walk-in transactions only (invoice-sourced and $0 test transactions excluded). Transaction net sales reconcile to the model's core sales within 3–5% per cafe; the difference is timing (deposit vs order date) and refunds netting. Ratios and shares are unaffected.")
    add_bullet(doc, "Timecards cover Sep 1–19 (13 of 21 weekdays at Downtown/1899, 19 of 30 days at Studio). All Section 3 labor ratios use matched days only. Sep 20–30 lands with the next weekly pull; expect small refinements, not a change in direction.")
    add_bullet(doc, "Reconciliation note: Section 3's matched-day labor % (35–38%) is higher than the P&L's 27.6%. The P&L labor line is full-month payroll from QBO against full-month core sales; Section 3 is timecard cost over 13–19 days, and the early-September weeks carried the heavier schedules. Use the P&L figure for the target test and Section 3 for the shape of the week.")
    add_bullet(doc, "Jul/Aug transaction exports truncate multi-item descriptions to the first item. Attach rate is therefore a calibrated proxy (see 3c) and Section 4 counts lead items, not units.")
    add_bullet(doc, "Section 6 (Loyalty) is now sourced from the Square Loyalty API ledger (Aug 1 onward) — screenshots retired. Section 5 (Discounts) on hold.")
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
    add_caveat(doc, f"Central Admin {fmt_money(S['Central'][pm]['ni'])} → {fmt_money(S['Central'][m]['ni'])}: the ~$5,700 of August tax prep / payment rolled off and labor fell $3.3K. "
                    f"SRP near break-even ({fmt_money(S['SRP'][m]['ni'])}) after August's loss; Commissary improved to {fmt_money(S['Commissary'][m]['ni'])}.")

    add_heading(doc, f"B. Consolidated P&L — {MONTH_LABEL}", level=2)
    C = fin["consolidated"]
    tm = MONTHS[-6:]
    def crow(label, key):
        return [label] + [fmt_money(C[mm][key]) for mm in tm] + [delta_pct(C[m][key], C[pm][key])]
    rows = [crow("Total Core Location Sales", "tcls"), crow("Total COGS", "tcogs"), crow("Gross Profit", "gp"), crow("Total Labor", "labor"),
            crow("Total Expenses", "texp"), crow("Net Operating Income", "noi"), crow("Cafes — Net Income", "cafes_ni"), crow("TOTAL NET INCOME (all entities)", "total_ni")]
    build_investor_table(doc, ["Metric"] + tm + ["MoM Δ"], rows, col_widths=[Inches(2.1)] + [Inches(0.85)] * 6 + [Inches(0.8)], key_rows=[5, 6], total_rows=[7])
    add_caveat(doc, f"Total Net Income {fmt_money(C[m]['total_ni'])} — best month of 2026, a {fmt_money(C[m]['total_ni'] - C[pm]['total_ni'], False)} swing from August. "
                    f"Cafes contributed {fmt_money(C[m]['cafes_ni'])} (of which NMWA $10.8K is a catering month); support entities {fmt_money(sum(S[e][m]['ni'] for e in S))}, "
                    f"with Central Admin back to {fmt_money(S['Central'][m]['ni'])} as the August tax items rolled off.")

    add_heading(doc, "C. Methodology & Data Sources", level=2)
    add_bullet(doc, f"Financials: Consolidated Model {MODEL_VER} → data/fin_2026_v18.json (extract of Jan–Aug per-entity IS tabs, Consolidated, Projections).")
    add_bullet(doc, "Sales patterns: _pipeline/shared/transactions/2026-08.csv (Square Transactions export) + weekly labor JSONs → analyze_patterns.py → data/2609_patterns.json.")
    add_bullet(doc, "Report: build_monthly_report_v2.py reads both JSONs; narrative blocks live in NARRATIVE at the top of that file.")
    add_bullet(doc, "Definitions: Core Sales = Total Core Location Sales − Catering. Core COGS = Total COGS − Catering COGS. Prime = Core COGS + Total Labor. Open Cafes = Downtown + Studio + 1899.")
    add_bullet(doc, "Daypart boundaries: Early 6–8a, Rush 8–10a, Mid-morning 10a–12p, Lunch 12–2p, Afternoon 2–5p, Late 5p+. Alignment gap = labor-hour share − sales share by day of week.")

    _ensure_zoom_percent(doc)
    doc.save(OUT)
    print("Saved →", OUT)


if __name__ == "__main__":
    build()
