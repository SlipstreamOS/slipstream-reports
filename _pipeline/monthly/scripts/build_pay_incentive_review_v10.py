#!/usr/bin/env python3
"""Build 'Slipstream Pay & Incentive Review - Sep 2026 v1.0.docx' — the pay review and incentive options split out of the
Labor Deep Dive (so the schedule templates can be shared without them), plus the 'what we can afford' ladder from the
12-month backtest."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, p, h, sm

DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Pay & Incentive Review - Sep 2026 v1.0.docx")
J = lambda n: json.load(open(os.path.join(DD, n)))
PAY = J("pay_review_v13.json"); INC = J("incentive_v12.json"); BT = J("backtest_v13.json"); AF = J("affordability_v13.json")
def k(x): return ("−" if x < 0 else "") + f"${abs(x)/1000:.1f}K"

def build():
    doc = Document(); sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.85)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Pay & Incentive Review", 22, True, COLOR_ACCENT, 0), ("Labor Deep Dive  |  September 2026  |  v1.0", 14, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "Prepared September 10, 2026 for Ryan Fleming — internal, not for the wider team. Companion to 'Schedule Templates by Demand Tier v1.4'. "
                  "Sources: Square timecards Apr–Aug 2026, DC DOES minimum-wage notice, public pay postings (Indeed, Glassdoor, ZipRecruiter, Sep 2026), Consolidated Model v17, the 12-month template backtest. "
                  "The incentive options are the v1.2 bench test and will be re-run once you have commented.", italic=True, size=9, color=COLOR_MUTED, after=12)
    s1, sf = PAY["cost_step1"], PAY["cost_of_proposed"]; sim = INC["simulation_2027"]["options"]; L = AF["headroom_12mo"]["ladder"]; sv = AF["template_saving"]["total"]

    add_heading(doc, "1. Summary", level=1)
    add_bullet(doc, "Peers pay the DC minimum ($18.40) plus tips; Slipstream pays a guaranteed rate with tips applied toward it. On that basis a new or certified barista here is $2–5 an hour below a peer barista, a Shift Lead at $20.50–22.50 is level with a peer barista and below a peer shift lead, and ALLs at $23–25 match a busy-cafe barista. Your read is right.", bold_lead="Where pay sits.")
    add_bullet(doc, f"Trainee at the minimum and time-boxed, then a steeper ladder than peers: parity floors first (every SL ≥ $22.50, every ALL ≥ $25.50), the band-top curve ($21 / $24.50 / $28, average $23.32), and — the stated goal — a peer-average curve ($21.50 / $26 / $29.50, average $24.23). Section 4 prices each.", bold_lead="Proposed curve.")
    add_bullet(doc, f"Target-based bonuses do not pay under the two-open/two-close rules (about two months a year). Option 4 — a fixed amount for every cafe-month run on template plus half of any beat — pays in ~11 months and costs {k(sim['opt4']['p50'])} a year ({k(sim['opt4']['p90'])} at the 90th percentile), about $0.30 an hour for a 30-hour person.", bold_lead="Incentive.")
    EC0 = AF["slipstream_economics_12mo"]["scenarios"]
    add_bullet(doc, f"Counting what Slipstream actually keeps — Downtown and Studio on templates, the 1899 management income, NMWA with its fee and the Nights events, less Commissary / SRP / Central and a $120K founder salary — the year is {k(EC0['planning']['ladder']['actual_cogs']['after_everything'])} at today's COGS and {k(EC0['planning']['ladder']['28.5']['after_everything'])} at 28.5%. "
                    f"Parity floors ({k(EC0['parity_floors']['extra_pl_labor_year'])}) and the band-top curve ({k(EC0['curve_a']['extra_pl_labor_year'])}) fit today. A curve that matches the peer AVERAGE of ${AF['peer_avg_all_in']:.2f} with entry held at $18.50 (certified $21.50 / SL $26 / ALL $29.50) costs {k(EC0['peer_avg_curve']['extra_pl_labor_year'])} a year — "
                    f"{k(EC0['peer_avg_curve']['ladder']['actual_cogs']['after_everything'])} at today's COGS, {k(EC0['peer_avg_curve']['ladder']['28.5']['after_everything'])} at 28.5%; five percent above that needs COGS at 28.5% just to break even. The average sets the cost; the shape only moves it between people. Section 4.", bold_lead="What we can afford.")
    add_bullet(doc, "Templates and parity floors together in October; incentive from the first full on-template month; the band-top curve once COGS is holding under 30% for two months; the peer-average curve once COGS is at 28.5% and Central has been looked at.", bold_lead="Sequence.")

    # ================= 2. PAY =================
    add_heading(doc, "2. Pay review — where the guaranteed ladder sits against DC peers", level=1)
    add_para(doc, "Peers pay DC minimum (or a little more) plus tips; Slipstream pays a guaranteed rate with tips applied toward it. So the honest comparison is guaranteed rate vs base + tips. Slipstream's own tip pool is "
                  f"${PAY['slipstream_now']['tips']['Downtown']['tips_per_core_hour']:.2f} / ${PAY['slipstream_now']['tips']['Studio']['tips_per_core_hour']:.2f} / ${PAY['slipstream_now']['tips']['1899 Penn']['tips_per_core_hour']:.2f} per core hour (Downtown / Studio / 1899; 10–12% of sales) — "
                  "the same money a peer barista would keep on top of $18.40.", after=6)
    rows = [[x["who"], x["base"], x["tips"], x["all_in"], x["sl"]] for x in PAY["peers"]]
    table(doc, ["Peer", "Base", "Tips / hr", "Barista all-in / hr", "Shift lead"], rows, widths=[1.9, 1.4, 0.8, 1.2, 1.4], size=8)
    add_caveat(doc, "Public postings and aggregators, Sep 2026 (Indeed, Glassdoor, ZipRecruiter); directional. Ryan's rule of thumb — minimum + tips at a busy cafe ≈ $24–25, shift leads +$1–2 — is consistent with Compass and with Slipstream's own tip rate.")
    add_heading(doc, "Slipstream today (August core hours)", level=3)
    rows = [[b["band"], str(b["people"]), f"{b['share_hours']}%", f"${b['avg_rate']:.2f}"] for b in PAY["slipstream_now"]["bands"]]
    table(doc, ["Band", "People", "Share of core hours", "Avg rate"], rows, widths=[2.4, 0.8, 1.4, 1.0], size=9)
    add_caveat(doc, f"Average guaranteed rate ${PAY['slipstream_now']['avg_core_rate_aug']:.2f}. Read against the peer table: a new or certified barista here is $2–5 an hour below a peer barista with tips; a Shift Lead at $20.50–22.50 is at parity with a peer barista, below a peer shift lead; ALLs at $23–25 match a busy-cafe barista and only just exceed a peer shift lead.")
    add_heading(doc, "A steeper, more predictable curve", level=3)
    rows = [[x["level"], f"${x['low']:.2f}" + (f" – ${x['high']:.2f}" if x["high"] != x["low"] else ""), x["note"]] for x in PAY["ladder_proposed"]]
    table(doc, ["Level", "Proposed rate", "Why"], rows, widths=[2.0, 1.2, 3.5], size=8)
    add_bullet(doc, f"Step 1 — parity floors now: every Shift Lead to at least $22.50 and every ALL to at least $25.50. Cost on the August roster ≈ {m(s1['month'])} a month ({m(s1['year'])} a year, {s1['pct_of_core_sales_aug']} points of core sales, ${s1['per_core_hour']:.2f} per core hour).", bold_lead="Cost.")
    add_bullet(doc, f"Band midpoints (certified $20.50, SL $23.50, ALL $27): ≈ {m(sf['month'])} a month ({m(sf['year'])} a year, {sf['pct_of_core_sales_aug']} points of core sales, ${sf['per_core_hour']:.2f} per core hour) on August's hours. Section 4 prices the band tops and the peer-average curve on the template hours.", bold_lead="")
    add_bullet(doc, "The trainee rate has to track the DC minimum ($18.40 since July 1, 2026, CPI-indexed each July); four sampled timecards were at $18.00. Keep it explicitly time-boxed (60–90 days to Essentials) so 'below peers' is a stage, not a status.", bold_lead="Floor.")
    add_bullet(doc, "The incentive ($0.50–1.50 an hour in on-template months) is what puts a Shift Lead above a peer shift lead; without it parity is the most the base curve buys at a sustainable labor %.", bold_lead="Incentive.")

    # ================= 3. INCENTIVE =================
    add_heading(doc, "3. Incentive options", level=1)
    add_para(doc, "Constraints: 1899 and NMWA cannot hit profit-style targets for a while; a plan that never pays is worse than none; the monthly piece must stay small so the annual retirement share is real. "
                  "All options are pooled across cafes and paid as a percentage of each person's core wages (which keeps overtime math clean), with a guardrail: the cafe's schedules met the floors and sales came in at 90% or more of projection.", after=6)
    for lab, txt in [("Option 1 — beat the target.", "Downtown and Studio pool half of any wages saved against 32% / 30% of core sales; everyone with core hours shares it. Capped at 1.5% of sales."),
                     ("Option 2 — own metric per cafe.", "Option 1, plus $200 to the pool in any month 1899 books NI of 5% or more, and $100 when NMWA runs at or under its staffing-rule hours."),
                     ("Option 3 — beat the template.", "Half of the difference between template wages and actual core wages, all four cafes, capped at 1.5% of core sales."),
                     ("Option 4 — on template, plus beat (recommended).", "A fixed $300 (Downtown) / $300 (Studio) / $150 (1899) / $75 (NMWA) for every cafe-month run at or under its template (1% tolerance), plus half of any beat, capped at 1.5%. Pays for running the plan, not for luck.")]:
        add_bullet(doc, txt, bold_lead=lab)
    pp = INC["per_person_30h_year_p50"]
    NAMES = {"opt1": "1. Beat the target", "opt2": "2. Own metric per cafe", "opt3": "3. Beat the template", "opt4": "4. On template + beat"}
    rows = [[NAMES[x], f"{sim[x]['months_paying_p50']} of 12", f"{m(sim[x]['p10'])} – {m(sim[x]['p90'])}", m(pp[x]), f"{INC['backtest_summary'][x]['months_paying']} of 5 ({m(INC['backtest_summary'][x]['total'])})"] for x in ["opt1", "opt2", "opt3", "opt4"]]
    table(doc, ["Option", "Months that pay (2027)", "Annual cost (range)", "Per 30 h/wk person / yr", "Apr–Aug 2026 as run"], rows, widths=[2.0, 1.2, 1.3, 1.2, 1.3], size=8)
    add_caveat(doc, "Simulated 2027 on the tier calendar with the v1.2 rule-based templates (2,000 runs); the last column is what each option would have paid on the actual Apr–Aug 2026 schedules. Under the two-open/two-close rules the templates sit above 32/30 in slow and mid weeks, so a target-based bonus pays in about two months a year — the 'never pays' problem. Measuring against the template pays when the team runs the plan.")
    for lab, txt in [("Fairness.", "Options 1–3 leave 1899 and NMWA out in practice. Option 4 gives every cafe the same test — did you run the plan — with the amount scaled to the cafe's size."),
                     ("Payroll mechanics.", "A non-discretionary bonus has to be folded into the regular rate for overtime (FLSA); paying it as a percentage of core wages does that automatically. Pay it with the first payroll after the month closes, from the weekly report's labor block."),
                     ("Annual share.", f"Keep the monthly piece small: on v1.2 templates the three cafes clear the support entities by roughly {k(INC['profit_share_v12']['consolidated_ni_year_v12'])} a year, so a 5–10% profit share to retirement is {k(INC['profit_share_v12']['share_5pct'])}–{k(INC['profit_share_v12']['share_10pct'])} — real, but only if the monthly piece stays under $7K."),
                     ("Comparables.", "Peers with hourly bonuses (Compass, Blue Bottle) pay $0.50–1.50 an hour tied to store metrics; per-shift attendance bonuses are common at fast-casual. Option 4 at $300 a month for a 3-person-a-day cafe is ~$0.30 an hour, so the $1–1.50 an hour you want to advertise needs the beat share on top of the fixed amount, or a bigger fixed amount once COGS headroom exists.")]:
        add_bullet(doc, txt, bold_lead=lab)

    # ================= 4. AFFORDABILITY =================
    EC = AF["slipstream_economics_12mo"]; SC = EC["scenarios"]; LN = EC["lines"]; FS = AF["pay_scenarios_12mo"]["founder_salary_year"]
    SHORT = {"planning": "Templates at today's rates", "parity_floors": "Parity floors (SL $22.50 / ALL $25.50)", "full_curve": "Midpoint curve ($20.50 / $23.50 / $27)", "curve_a": "Band-top curve ($21 / $24.50 / $28)",
             "peer_avg_curve": "Peer-average curve ($21.50 / $26 / $29.50)", "peer_avg_curve_plus5": "Peer-average curve + 5%"}
    ROWS4 = ("planning", "parity_floors", "full_curve", "curve_a", "peer_avg_curve", "peer_avg_curve_plus5")
    add_heading(doc, "4. What we can afford — Slipstream's own twelve months", level=1)
    add_para(doc, "What is in the number: Downtown and Studio on the templates (the 12-month backtest, catering kept, core COGS re-run at each scenario); 1899 counted only as the management income Slipstream keeps ($5K a month plus 5% of sales — its NOI passes to the owner, so its template saving is the owner's, not ours); "
                  "NMWA at its July–August run-rate, which includes the $5,041 monthly fee, plus the NMWA Nights events at your figure (8 × $4,500 net); Commissary, SRP and Central at their 2026 average; and a $120K founder salary the financials do not carry today.", after=6)
    rows = [["Downtown + Studio NOI on templates (12 months, actual COGS)", m(LN["downtown_studio_noi_on_templates_12mo"])], ["1899 management income ($5K + 5% of sales)", m(LN["mgmt_income_1899_year"])],
            ["NMWA cafe at Jul–Aug run-rate (includes $5,041/mo fee)", m(LN["nmwa_cafe_run_rate_year_incl_fee"])], ["NMWA Nights, 8 × $4,500 net", m(LN["nmwa_nights_year"])],
            ["Commissary", m(LN["commissary_year"])], ["SRP (roasting)", m(LN["srp_year"])], ["Central", m(LN["central_year"])], ["Founder salary", m(LN["founder_salary_year"])],
            ["= Slipstream, templates at today's rates and today's COGS", m(SC["planning"]["ladder"]["actual_cogs"]["after_everything"])]]
    table(doc, ["Line, per year", "$"], rows, widths=[4.6, 1.2], total_rows=[8], size=9)
    add_caveat(doc, f"Not in the number: 1899's own NOI on templates ({m(LN['excluded_1899_noi_on_templates_12mo'])}, passed through) and the incentive ({m(sim['opt4']['p50'])}). NMWA's run-rate rests on two months (Aug: −$129 cafe NOI, +$4.4K after the fee). Support entities are 2026 Jan–Aug × 12.")
    rows = []
    for kk in ROWS4:
        v = SC[kk]; rows.append([SHORT[kk], f"${v['roster_avg_rate']:.2f}", p(v["labor_pct_of_core"]), ("−" + k(v["extra_pl_labor_year"])) if v["extra_pl_labor_year"] else "—"] + [k(v["ladder"][sc]["after_everything"]) for sc in ("actual_cogs", "30", "28.5", "27")])
    table(doc, ["Pay scenario (Downtown + Studio)", "Avg rate", "Labor % of core", "Extra labor / yr", f"COGS {p(EC['own_cafes']['cogs_pct'])} (actual)", "COGS 30%", "COGS 28.5%", "COGS 27%"], rows, widths=[1.9, 0.6, 0.7, 0.8, 0.85, 0.65, 0.65, 0.65], size=8)
    add_caveat(doc, f"Slipstream profit after everything above, per year. Avg rate = the August roster's hours-weighted guaranteed rate under that curve (trainee hours are {AF['hours_mix_aug']['trainee']*100:.0f}% of the mix today). Labor % = P&L labor on templates ÷ Downtown + Studio core sales. "
                    f"Each pay row moves every roster member up to the scenario rate for their band — nobody down; trainees stay at $18.50 in every row; raises at 1899 are the owner's cost and are not here. Peer average = ${AF['peer_avg_all_in']:.2f}, a DC peer's hours-weighted take-home with tips (baristas ~$24, new hires ~$23, shift leads ~$25.50).")
    add_bullet(doc, f"Templates at today's rates and COGS leave {k(SC['planning']['ladder']['actual_cogs']['after_everything'])} after the salary; {k(SC['planning']['ladder']['28.5']['after_everything'])} at 28.5%. Parity floors ({k(SC['parity_floors']['extra_pl_labor_year'])}), the midpoint curve ({k(SC['full_curve']['extra_pl_labor_year'])}) and the band-top curve ({k(SC['curve_a']['extra_pl_labor_year'])}) all fit at today's COGS.", bold_lead="Today's ladder.")
    add_bullet(doc, f"The goal is the peer AVERAGE with a steeper shape — entry below peers, the top half above. On today's hours mix that takes certified $21.50, Shift Lead $26 and ALL $29.50 to average ${SC['peer_avg_curve']['roster_avg_rate']:.2f}: a {SC['peer_avg_curve']['uplift_pct']}% raise on the wage bill, {k(SC['peer_avg_curve']['extra_pl_labor_year'])} a year, labor at {p(SC['peer_avg_curve']['labor_pct_of_core'])} of core sales. "
                    f"It is {k(SC['peer_avg_curve']['ladder']['actual_cogs']['after_everything'])} at today's COGS, {k(SC['peer_avg_curve']['ladder']['30']['after_everything'])} at 30%, {k(SC['peer_avg_curve']['ladder']['28.5']['after_everything'])} at 28.5%. Five percent above the peer average (certified $22.58 / SL $27.30 / ALL $30.98) is {k(SC['peer_avg_curve_plus5']['extra_pl_labor_year'])} a year: "
                    f"{k(SC['peer_avg_curve_plus5']['ladder']['actual_cogs']['after_everything'])} today, {k(SC['peer_avg_curve_plus5']['ladder']['28.5']['after_everything'])} at 28.5%.", bold_lead="Peer-average curve.")
    add_bullet(doc, f"The average is what costs money; the shape is free. Matching the peer average costs about {k(SC['peer_avg_curve']['extra_pl_labor_year'])} a year however it is distributed (a flat peer curve of $23 / $24 / $25.50 / $27 is {k(SC['peer']['extra_pl_labor_year'])} — more, because it lifts trainees). "
                    f"Keeping trainees at $18.50 is worth ~$1.50 an hour on the average only while trainee hours are a quarter of the mix; as people certify on the 60–90-day clock the average — and the cost — rises toward the curve's midpoint without any rate change. Budget the peer-average curve at $70–80K a year, not $67K.", bold_lead="What drives the cost.")
    add_bullet(doc, f"The management income ({k(LN['mgmt_income_1899_year'])}) and NMWA with its fee and Nights ({k(LN['nmwa_cafe_run_rate_year_incl_fee'] + LN['nmwa_nights_year'])}) are what make the peer-average curve reachable — together they are more than Downtown's or Studio's cafe profit. Central ({k(-LN['central_year'])}) is the largest single line on the page and the one to look at next; every $10K there is worth as much as a point of COGS.", bold_lead="What carries it.")
    add_bullet(doc, "It is not the tip structure. Because tips are applied to the guarantee, a $24 guarantee costs the P&L about $21.30 an hour (1.12 × $24 − $5.60 of tips) against roughly $20.60 for a peer paying $18.40 plus tax — near parity per hour. "
                    "The gap is sales per labor hour: the templates need $60–68 of core sales per hour to hit 32/30 at today's rates, and Downtown and Studio ran $45–50 in August. The peer-average curve at Slipstream's volumes is a 35% labor line; at a busy peer's volumes it would be under 30%. Price (1% ≈ $10K a year on Downtown + Studio before any volume response) and volume are what turn it from thin to comfortable.", bold_lead="Why.")
    add_heading(doc, "Levers, per year", level=3)
    rows = [["Templates vs what was run (Downtown + Studio)", k(AF['template_saving']['by_cafe']['Downtown']['12mo'] + AF['template_saving']['by_cafe']['Studio']['12mo']), k(round((AF['template_saving']['by_cafe']['Downtown']['jan_aug_2026'] + AF['template_saving']['by_cafe']['Studio']['jan_aug_2026']) * 1.5))],
            ["COGS to 30% (Downtown + Studio)", "", k(LN['cogs_adjustment']['30'])], ["COGS to 28.5%", "", k(LN['cogs_adjustment']['28.5'])], ["COGS to 27%", "", k(LN['cogs_adjustment']['27'])],
            ["Parity floors", "", "−" + k(SC['parity_floors']['extra_pl_labor_year'])], ["Midpoint curve", "", "−" + k(SC['full_curve']['extra_pl_labor_year'])], ["Band-top curve", "", "−" + k(SC['curve_a']['extra_pl_labor_year'])], ["Peer-average curve", "", "−" + k(SC['peer_avg_curve']['extra_pl_labor_year'])], ["Peer-average curve + 5%", "", "−" + k(SC['peer_avg_curve_plus5']['extra_pl_labor_year'])],
            ["Incentive, Option 4 (median)", "", "−" + k(sim['opt4']['p50'])], ["Founder salary", "", "−" + k(FS)]]
    table(doc, ["Lever", "Last 12 months", "Per year going forward"], rows, widths=[3.0, 1.3, 1.6], size=9)
    add_caveat(doc, "Pay costs are on the 12-month template wage base for Downtown + Studio (core + admin/training, loaded ×1.12); Section 2 quotes the same steps on August's actual hours across all cafes, which were higher. The flat peer curves and a three-cafe view that (wrongly) keeps 1899's NOI are on the workbook's Affordability sheet for reference.")

    # ================= 5. SEQUENCE =================
    add_heading(doc, "5. Suggested sequence", level=1)
    for lab, txt in [("October.", "Templates live in Square Teams; parity floors for Shift Leads and ALLs; trainee rate confirmed at $18.40 on every card. Announce the incentive rule and what 'on template' means."),
                     ("November.", "First incentive payout on October's results (if on template). Weekly report labor block becomes the scorecard: gross wages ÷ projected core sales, by cafe, against the tier's template."),
                     ("January.", "Review COGS by cafe. If the three cafes have held under 30% for two months, move certified baristas to $20.50 and SLs to the $23.50 midpoint; ALLs to $27 with the Cafe Leader hire."),
                     ("Every July.", "DC minimum resets (CPI); re-base the trainee rate and check the certified floor still clears it by $2.")]:
        add_bullet(doc, txt, bold_lead=lab)
    _ensure_zoom_percent(doc); doc.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    build()
