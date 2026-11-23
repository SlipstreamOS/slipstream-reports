#!/usr/bin/env python3
"""Build 'Slipstream Pay & Incentive Review - Sep 2026 v1.0.xlsx' — Pay Review, Incentive (calculator + bench test) and Affordability sheets split out of the schedule workbook.
Slipstream color protocol: blue = input (rates, targets, capacity, typical sales, start/end times), green = pulled data
(need/actual curves, first-cut transcription, simulation results), black = formulas/labels.
Sheets: README, Pay Review, Incentive, Affordability.
"""
import os, json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Pay & Incentive Review - Sep 2026 v1.0.xlsx")
T = json.load(open(os.path.join(DD, "templates_v13.json"))); T0 = json.load(open(os.path.join(DD, "templates.json"))); PAY = json.load(open(os.path.join(DD, "pay_review_v13.json"))); CAL19 = json.load(open(os.path.join(DD, "tiers_v19_sep_dec.json")))
NV = json.load(open(os.path.join(DD, "need_v13.json"))); S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
INC = json.load(open(os.path.join(DD, "incentive_v12.json"))); PC = json.load(open(os.path.join(DD, "projections_check_v12.json"))); TIERS = json.load(open(os.path.join(DD, "tiers.json"))); STAFF = json.load(open(os.path.join(DD, "staffing.json")))
MAIN = "std"

BLUE = Font(name="Calibri", size=10, color="1F3864"); BLUE_FILL = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
YELLOW = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GREEN = Font(name="Calibri", size=10, color="006100"); BLACK = Font(name="Calibri", size=10, color="000000")
BOLD = Font(name="Calibri", size=10, bold=True); HDR = Font(name="Calibri", size=10, bold=True, color="1F3A5F")
HDR_FILL = PatternFill("solid", start_color="E8ECEF", end_color="E8ECEF"); TOT_FILL = PatternFill("solid", start_color="D9E2EC", end_color="D9E2EC")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F3A5F"); NOTE = Font(name="Calibri", size=9, italic=True, color="606060")
THIN = Side(style="thin", color="1F3A5F")
CAFES = ["Downtown", "Studio", "1899 Penn", "NMWA"]; OPEN3 = ["Downtown", "Studio", "1899 Penn"]
DOW7 = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
ROLE_CODES = [("SL", "Shift Lead"), ("PL", "Production Line"), ("PB", "Production Barista"), ("B", "Barista"), ("HR", "Hospitality Register"), ("CL", "Closer (keyholder barista)"), ("FOH", "FOH")]
TIER_LABEL = {"slow": "Slow", "mid": "Mid", "busy": "Busy"}
RATES = dict(T["rates"])

def hdr(ws, row, headers, widths=None):
    for j, t in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=t); c.font = HDR; c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = Border(bottom=THIN)
    if widths:
        for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w
def inp(c, fmt=None, key=False):
    c.font = BLUE; c.fill = YELLOW if key else BLUE_FILL
    if fmt: c.number_format = fmt
def pull(c, fmt=None):
    c.font = GREEN
    if fmt: c.number_format = fmt
def fx(c, fmt=None, bold=False):
    c.font = BOLD if bold else BLACK
    if fmt: c.number_format = fmt
def tm(s):
    hh, mm = map(int, s.split(":")); return dt.time(hh, mm)
def dec2time(x):
    return dt.time(int(x), int(round((x % 1) * 60)))


def main():
    wb = Workbook(); AF = json.load(open(os.path.join(DD, "affordability_v13.json"))); BT = json.load(open(os.path.join(DD, "backtest_v13.json")))
    ws = wb.active; ws.title = "README"
    ws["A1"] = "Slipstream Pay & Incentive Review — September 2026 v1.0"; ws["A1"].font = TITLE
    lines = ["Companion to 'Slipstream Pay & Incentive Review - Sep 2026 v1.0.docx'. Internal — split out of the schedule workbook so the templates can be shared without pay and incentive detail.", "",
             "COLOR PROTOCOL — blue text on light blue = input you can change (yellow = key assumptions); green = pulled from the analysis data; black = formula or label.", "",
             "Pay Review: DC peer comparison (base + tips), Slipstream's August bands, the proposed ladder (edit the blue cells to re-cost) and its monthly cost.",
             "Incentive: Option 4 calculator (enter a month's core sales, core wages and template wages per cafe), the Apr–Aug 2026 backtest of all four options, and the 2027 simulation.",
             "Affordability: pay scenario (today's rates, parity floors, full curve, DC peer all-in, peer + 5%) × COGS scenario (actual, 30, 28.5, 27%), profit after support entities and a $120K founder salary. Uplift %, salary, support, load and the COGS scenarios are blue inputs; everything else recomputes."]
    for i, t in enumerate(lines, 3):
        ws.cell(row=i, column=1, value=t).font = NOTE if t.startswith(("Companion", "COLOR")) else BLACK; ws.cell(row=i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 150
    LOAD_CELL = "Incentive!$B$19"

    # ---------------- Pay Review ----------------
    wpay = wb.create_sheet("Pay Review")
    wpay["A1"] = "Pay review — Slipstream's guaranteed ladder vs DC peers (base + tips), and the cost of a steeper curve"; wpay["A1"].font = TITLE
    wpay["A2"] = "Green = pulled (Aug 2026 timecards, transactions, public postings); blue = the proposed ladder (edit to re-cost); black = formulas."; wpay["A2"].font = NOTE
    hdr(wpay, 4, ["Peer", "Base", "Tips / hr", "Barista all-in", "Shift lead", "Source"], [34, 30, 12, 16, 18, 44])
    r = 5
    for x in PAY["peers"]:
        for j, k in enumerate(["who", "base", "tips", "all_in", "sl", "src"]): pull(wpay.cell(row=r, column=1 + j, value=x[k]))
        r += 1
    r += 1
    wpay.cell(row=r, column=1, value="Slipstream today (Aug 2026 core hours)").font = BOLD; r += 1
    hdr(wpay, r, ["Band", "People", "Share of core hours", "Avg rate", "Hours / month"]); r += 1
    for b in PAY["slipstream_now"]["bands"]:
        pull(wpay.cell(row=r, column=1, value=b["band"])); pull(wpay.cell(row=r, column=2, value=b["people"])); pull(wpay.cell(row=r, column=3, value=b["share_hours"] / 100), "0%"); pull(wpay.cell(row=r, column=4, value=b["avg_rate"]), '"$"#,##0.00'); pull(wpay.cell(row=r, column=5, value=b["hours_month"]), "#,##0")
        r += 1
    r += 1
    wpay.cell(row=r, column=1, value="Tips per core hour (Aug 2026) — applied toward the guaranteed rate").font = BOLD; r += 1
    for cafe, t in PAY["slipstream_now"]["tips"].items():
        wpay.cell(row=r, column=1, value=cafe).font = BLACK; pull(wpay.cell(row=r, column=2, value=t["tips_per_core_hour"]), '"$"#,##0.00'); pull(wpay.cell(row=r, column=3, value=t["tip_pct_of_net"] / 100), "0.0%"); r += 1
    r += 1
    wpay.cell(row=r, column=1, value="Proposed ladder (edit the blue cells) and its cost on the August roster").font = BOLD; r += 1
    hdr(wpay, r, ["Level", "Proposed low", "Proposed high", "Why"]); r += 1
    for x in PAY["ladder_proposed"]:
        wpay.cell(row=r, column=1, value=x["level"]).font = BLACK; inp(wpay.cell(row=r, column=2, value=x["low"]), '"$"#,##0.00', key=True); inp(wpay.cell(row=r, column=3, value=x["high"]), '"$"#,##0.00', key=True); wpay.cell(row=r, column=4, value=x["note"]).font = BLACK; r += 1
    r += 1
    hdr(wpay, r, ["Band today", "People", "Avg rate now", "Proposed (full curve)", "Hours / month", "Cost / month (full curve)"]); r += 1
    tot_row0 = r
    for b in PAY["slipstream_now"]["bands"]:
        bb = PAY["cost_of_proposed"]["by_band"][b["band"]]
        pull(wpay.cell(row=r, column=1, value=b["band"])); pull(wpay.cell(row=r, column=2, value=bb["people"])); pull(wpay.cell(row=r, column=3, value=bb["avg_now"]), '"$"#,##0.00')
        inp(wpay.cell(row=r, column=4, value=bb["proposed"]), '"$"#,##0.00', key=True); pull(wpay.cell(row=r, column=5, value=b["hours_month"]), "#,##0"); fx(wpay.cell(row=r, column=6, value=f"=(D{r}-C{r})*E{r}"), '"$"#,##0')
        r += 1
    wpay.cell(row=r, column=1, value="Total").font = BOLD; fx(wpay.cell(row=r, column=6, value=f"=SUM(F{tot_row0}:F{r-1})"), '"$"#,##0', bold=True)
    for col in range(1, 7): wpay.cell(row=r, column=col).fill = TOT_FILL
    r += 1
    wpay.cell(row=r, column=1, value=f"Step 1 (parity floors only: SL ≥ $22.50, ALL ≥ $25.50): ${PAY['cost_step1']['month']:,} / month, {PAY['cost_step1']['pct_of_core_sales_aug']} points of core sales. Full curve: ${PAY['cost_of_proposed']['month']:,} / month, {PAY['cost_of_proposed']['pct_of_core_sales_aug']} points.").font = NOTE
    wpay.freeze_panes = "A5"

    # ---------------- Incentive ----------------
    wc = wb.create_sheet("Incentive")
    wc["A1"] = "Incentive — Option 4 calculator, the four options bench-tested (2026 backtest, 2027 simulation)"; wc["A1"].font = TITLE
    wc["A2"] = "Calculator: enter a month's core sales, core wages and template wages per cafe (blue); on-template fixed amounts, the beat share, cap and payout are formulas. Below: backtest and simulation results (green)."; wc["A2"].font = NOTE
    wc["A4"] = "Option 4 calculator — on-template fixed $ + 50% of the beat below template, cap 1.5% of sales"; wc["A4"].font = BOLD
    hdr(wc, 5, ["Cafe", "Core sales (month)", "Core wages (month)", "Template wages (month)", "On-template $", "Tolerance", "Qualifies?", "Fixed earned", "Beat $ (template − actual)"], [22, 16, 16, 18, 13, 10, 10, 12, 20])
    FIXED = {"Downtown": 300, "Studio": 300, "1899 Penn": 150, "NMWA": 75}
    for i, cafe in enumerate(["Downtown", "Studio", "1899 Penn", "NMWA"]):
        r = 6 + i
        wc.cell(row=r, column=1, value=cafe).font = BLACK
        if cafe != "NMWA":
            wk = T["cafes"][cafe]["mid"]["week"]; inp(wc.cell(row=r, column=2, value=round(wk["sales_typical"] * 52 / 12)), '"$"#,##0', key=True)
            inp(wc.cell(row=r, column=3, value=round(wk["wages"] * 52 / 12 * 0.98)), '"$"#,##0', key=True); inp(wc.cell(row=r, column=4, value=round(wk["wages"] * 52 / 12)), '"$"#,##0', key=True)
        else:
            inp(wc.cell(row=r, column=2, value=4593), '"$"#,##0', key=True); inp(wc.cell(row=r, column=3, value=2250), '"$"#,##0', key=True); inp(wc.cell(row=r, column=4, value=2210), '"$"#,##0', key=True)
        inp(wc.cell(row=r, column=5, value=FIXED[cafe]), '"$"#,##0', key=True); inp(wc.cell(row=r, column=6, value=0.01), "0%")
        fx(wc.cell(row=r, column=7, value=f'=IF(C{r}<=D{r}*(1+F{r}),"yes","no")')); fx(wc.cell(row=r, column=8, value=f'=IF(G{r}="yes",E{r},0)'), '"$"#,##0'); fx(wc.cell(row=r, column=9, value=f"=D{r}-C{r}"), '"$"#,##0;("$"#,##0)')
    wc.cell(row=10, column=1, value="All cafes").font = BOLD
    for col in (2, 3, 4, 8, 9): fx(wc.cell(row=10, column=col, value=f"=SUM({L(col)}6:{L(col)}9)"), '"$"#,##0;("$"#,##0)', bold=True)
    for col in range(1, 10): wc.cell(row=10, column=col).fill = TOT_FILL
    for i, (lab, val, fmt) in enumerate([("Share of the beat", 0.5, "0%"), ("Cap (% of all-cafe core sales)", 0.015, "0.0%"), ("Pool before cap = fixed earned + share × max(0, beat)", "=H10+B12*MAX(0,I10)", '"$"#,##0'),
                                          ("Cap $", "=B13*B10", '"$"#,##0'), ("Payout (capped)", "=MIN(B14,B15)", '"$"#,##0'), ("Bonus % of core wages (percentage-of-wages split)", "=IF(C10>0,B16/C10,0)", "0.00%"),
                                          ("Employer cost incl. load", f"=B16*{LOAD_CELL}", '"$"#,##0'), ("Load factor on gross wages", 1.12, "0.00")]):
        r = 12 + i; wc.cell(row=r, column=1, value=lab).font = BLACK; c = wc.cell(row=r, column=2, value=val)
        if isinstance(val, (int, float)): inp(c, fmt, key=True)
        else: fx(c, fmt, bold=lab.startswith("Payout"))
    wc.cell(row=21, column=1, value="Guardrail (not modelled): schedules met the floors (two open / two close, no unplanned solo), core sales ≥ 90% of the re-based projection, admin hours tracked alongside.").font = NOTE
    r = 22
    wc.cell(row=r, column=1, value="Backtest — Apr–Aug 2026 actual months (what each option would have paid)").font = BOLD; r += 1
    hdr(wc, r, ["Month", "Downtown tier / core % / template $ / actual $", "Studio", "1899 (NI %)", "Opt 1 Open-cafe pool", "Opt 2 Own-metric", "Opt 3 Template beat", "Opt 4 On-template + beat"]); r += 1
    for b in INC["backtest_2026"]:
        wc.cell(row=r, column=1, value=b["month"]).font = BLACK
        for j, x in enumerate(b["cafes"]):
            pull(wc.cell(row=r, column=2 + j, value=f"{x['tier']} / {x['core_pct']}% / ${x['tmpl_wages']:,} / ${x['wages']:,}" + (f" / NI {x['ni_pct']}%" if x["cafe"] == "1899 Penn" else "")))
        for j, k in enumerate(["opt1", "opt2", "opt3", "opt4"]): pull(wc.cell(row=r, column=5 + j, value=b[k]), '"$"#,##0')
        r += 1
    r += 1
    sim = INC["simulation_2027"]
    wc.cell(row=r, column=1, value=f"2027 simulation — {sim['runs']} runs; labor saved vs Apr–Aug run-rate p50 ${sim['labor_saving_vs_aprAug_run_rate']['p50']:,}").font = BOLD; r += 1
    hdr(wc, r, ["Option", "Annual payout p10", "p50", "p90", "Months paying (p50)", "Share of labor saved", "Per 30 h/wk person (p50, year)"]); r += 1
    for k in ["opt1", "opt2", "opt3", "opt4"]:
        v = sim["options"][k]; wc.cell(row=r, column=1, value=INC["definitions"][k]).font = BLACK
        pull(wc.cell(row=r, column=2, value=v["p10"]), '"$"#,##0'); pull(wc.cell(row=r, column=3, value=v["p50"]), '"$"#,##0'); pull(wc.cell(row=r, column=4, value=v["p90"]), '"$"#,##0')
        pull(wc.cell(row=r, column=5, value=v["months_paying_p50"]), "0"); pull(wc.cell(row=r, column=6, value=v["share_of_saving_p50"] / 100), "0.0%"); pull(wc.cell(row=r, column=7, value=INC["per_person_30h_year_p50"][k]), '"$"#,##0')
        r += 1
    wc.column_dimensions["A"].width = 70
    for col in "BCDEFGHI": wc.column_dimensions[col].width = max(wc.column_dimensions[col].width or 0, 18)
    wc.freeze_panes = "A5"


    # ---------------- Affordability ----------------
    wa = wb.create_sheet("Affordability")
    PS = AF["pay_scenarios_12mo"]; EC = AF["slipstream_economics_12mo"]; SC = EC["scenarios"]; LN = EC["lines"]; OC = EC["own_cafes"]
    wa["A1"] = "What we can afford — Slipstream's own economics, twelve months, by pay scenario × COGS scenario"; wa["A1"].font = TITLE
    wa["A2"] = "Downtown + Studio on templates (backtest, catering kept, core COGS re-run) + 1899 management income + NMWA run-rate (incl. $5,041/mo fee) + NMWA Nights − Commissary/SRP/Central − founder salary − extra labor of the pay scenario. 1899's own NOI is passed to the owner and excluded. Green = from the backtest/model; blue = inputs; black = formulas."; wa["A2"].font = NOTE
    inputs = [("Downtown + Studio core sales, 12 months", OC["core_sales"], '"$"#,##0', False), ("Downtown + Studio core COGS actual, 12 months", OC["cogs_actual"], '"$"#,##0', False),
              ("Downtown + Studio NOI on templates, 12 months (reported basis, today's rates)", LN["downtown_studio_noi_on_templates_12mo"], '"$"#,##0', False),
              ("Downtown + Studio template wage base, 12 months (core + admin/training)", OC["wage_base"], '"$"#,##0', False), ("Downtown + Studio P&L labor on templates, 12 months", OC["labor_template"], '"$"#,##0', False),
              ("Load factor on gross wages", PS["load"], "0.00", True),
              ("1899 management income, per year ($5K + 5% of sales; 2026 avg × 12)", LN["mgmt_income_1899_year"], '"$"#,##0', True),
              ("NMWA cafe, per year (Jul–Aug 2026 NI run-rate × 12, includes the $5,041/mo fee)", LN["nmwa_cafe_run_rate_year_incl_fee"], '"$"#,##0', True),
              ("NMWA Nights events, per year (8 × $4,500 net)", LN["nmwa_nights_year"], '"$"#,##0', True),
              ("Commissary, per year (2026 avg × 12)", LN["commissary_year"], '"$"#,##0;("$"#,##0)', True), ("SRP, per year", LN["srp_year"], '"$"#,##0;("$"#,##0)', True), ("Central, per year", LN["central_year"], '"$"#,##0;("$"#,##0)', True),
              ("Founder salary, per year", -LN["founder_salary_year"], '"$"#,##0', True), ("Incentive Option 4, per year (median) — not deducted below", AF["costs"]["incentive_opt4_p50"], '"$"#,##0', True)]
    for i, (lab, val, fmt, is_inp) in enumerate(inputs):
        r = 4 + i; wa.cell(row=r, column=1, value=lab).font = BLACK; c = wa.cell(row=r, column=2, value=val)
        (inp(c, fmt, key=True) if is_inp else pull(c, fmt))
    # B4 core, B5 cogs, B6 noi, B7 base, B8 labor, B9 load, B10 mgmt, B11 nmwa, B12 nights, B13 comm, B14 srp, B15 central, B16 salary, B17 incentive
    wa["A18"] = "Fixed lines, per year (mgmt + NMWA + Nights + Commissary + SRP + Central − salary)"; wa["A18"].font = BOLD; fx(wa["B18"], '"$"#,##0;("$"#,##0)', bold=True); wa["B18"] = "=B10+B11+B12+B13+B14+B15-B16"
    wa["A20"] = "Slipstream profit after everything, per year — pay scenario × COGS scenario (blue COGS % in row 21)"; wa["A20"].font = BOLD
    hdr(wa, 21, ["Pay scenario (Downtown + Studio)", "Uplift % on wage bill", "Roster avg rate", "Extra P&L labor / yr", "Labor % of core", "COGS →", "actual", 0.30, 0.285, 0.27], [52, 14, 12, 16, 12, 8, 14, 14, 14, 14])
    for col in (8, 9, 10): inp(wa.cell(row=21, column=col), "0.0%", key=True)
    r = 22
    for kk in ("planning", "parity_floors", "full_curve", "curve_a", "peer_avg_curve", "peer_avg_curve_plus5", "peer", "peer_plus5"):
        v = SC[kk]; wa.cell(row=r, column=1, value=v["label"]).font = BLACK
        inp(wa.cell(row=r, column=2, value=v["uplift_pct"] / 100), "0.0%", key=True); pull(wa.cell(row=r, column=3, value=v["roster_avg_rate"]), '"$"#,##0.00')
        fx(wa.cell(row=r, column=4, value=f"=$B$7*B{r}*$B$9"), '"$"#,##0'); fx(wa.cell(row=r, column=5, value=f"=($B$8+D{r})/$B$4"), "0.0%")
        fx(wa.cell(row=r, column=7, value=f"=$B$6-D{r}+$B$18"), '"$"#,##0;("$"#,##0)')
        for col in (8, 9, 10): fx(wa.cell(row=r, column=col, value=f"=$B$6+($B$5-{L(col)}$21*$B$4)-$D{r}+$B$18"), '"$"#,##0;("$"#,##0)')
        r += 1
    wa.cell(row=r, column=1, value="Uplift % = every August roster member moved from their current band to the scenario rate (nobody down), hours-weighted incl. admin/training. Peer average $24.25 = a DC peer's hours-weighted take-home with tips. Curves: midpoint $18.50/$20.50/$23.50/$27; band-top $18.50/$21/$24.50/$28; peer-average $18.50/$21.50/$26/$29.50 (+5% keeps the trainee at $18.50); flat peer $23/$24/$25.50/$27. Raises at 1899 are the owner's cost (pass-through) and are not included. Subtract B17 from any cell for the incentive.").font = NOTE
    r += 2
    wa.cell(row=r, column=1, value="Reference — three-cafe view that keeps 1899's NOI (superseded), profit after support ($146K) and salary").font = BOLD; r += 1
    hdr(wa, r, ["Pay scenario", "", "", "", "", "", "actual COGS", "30%", "28.5%", "27%"]); r += 1
    for kk in ("planning", "parity_floors", "full_curve", "curve_a", "peer_avg_curve", "peer_avg_curve_plus5", "peer", "peer_plus5"):
        v = PS["scenarios"][kk]; wa.cell(row=r, column=1, value=v["label"]).font = BLACK
        for j, sc in enumerate(("actual_cogs", "30", "28.5", "27")): pull(wa.cell(row=r, column=7 + j, value=v["ladder"][sc]["after_support_salary"]), '"$"#,##0;("$"#,##0)')
        r += 1
    r += 1
    wa.cell(row=r, column=1, value="Levers, per year").font = BOLD; r += 1
    hdr(wa, r, ["Lever", "Last 12 months", "Per year going forward"]); r += 1
    sv = AF["template_saving"]["total"]
    for lab, a, b in [("Templates vs what was run", sv["12mo"], sv["annualized_2026"]), ("  of which Jan–Aug 2026", sv["jan_aug_2026"], None), ("  of which Sep–Dec 2025 (2025 rates were lower)", sv["sep_dec_2025"], None),
                      ("COGS per point (on 12-month core sales)", None, AF["cogs"]["per_point_year"]), ("COGS 33% → 30%", None, AF["cogs"]["gain_vs_2026_blend"]["30"]), ("COGS 33% → 28.5%", None, AF["cogs"]["gain_vs_2026_blend"]["28.5"]), ("COGS 33% → 27%", None, AF["cogs"]["gain_vs_2026_blend"]["27"]),
                      ("Price: 1% move, before volume response", None, PS["price_lever"]["pct_1_of_core_sales"]),
                      ("Parity floors", None, -SC["parity_floors"]["extra_pl_labor_year"]), ("Full curve", None, -SC["full_curve"]["extra_pl_labor_year"]), ("Peer all-in", None, -SC["peer"]["extra_pl_labor_year"]), ("Peer + 5%", None, -SC["peer_plus5"]["extra_pl_labor_year"]),
                      ("Incentive Option 4 (median)", None, -AF["costs"]["incentive_opt4_p50"]), ("Founder salary", None, -PS["founder_salary_year"]), ("Support entities (2026 pace)", None, AF["headroom_12mo"]["support_burden_year"])]:
        wa.cell(row=r, column=1, value=lab).font = BLACK
        if a is not None: pull(wa.cell(row=r, column=2, value=a), '"$"#,##0;("$"#,##0)')
        if b is not None: pull(wa.cell(row=r, column=3, value=b), '"$"#,##0;("$"#,##0)')
        r += 1
    r += 1
    wa.cell(row=r, column=1, value="Template saving by cafe (12 months / Jan–Aug 2026)").font = BOLD; r += 1
    for cafe, v in AF["template_saving"]["by_cafe"].items():
        wa.cell(row=r, column=1, value=cafe).font = BLACK; pull(wa.cell(row=r, column=2, value=v["12mo"]), '"$"#,##0;("$"#,##0)'); pull(wa.cell(row=r, column=3, value=v["jan_aug_2026"]), '"$"#,##0;("$"#,##0)'); r += 1
    wa.freeze_panes = "A22"
    wb.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    main()
