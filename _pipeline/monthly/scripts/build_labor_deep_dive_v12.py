#!/usr/bin/env python3
"""Build 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.2.docx' — templates rebuilt to Ryan's
Sep 5 rules (2 open / 2 close at Downtown & Studio, 6-hour minimum shifts, Studio 3 mid-day, admin days, roster planning
rates, slow-month ceiling +2), projections reality check, and the incentive bench test (four options)."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, add_callout, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, p, h, sd, sm

DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.2.docx")
J = lambda n: json.load(open(os.path.join(DD, n)))
S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
T = J("templates_v12.json"); NV = J("need_v12.json"); TIERS = J("tiers.json"); CURVES = J("curves.json"); STAFF = J("staffing.json"); CAP = J("capacity.json")
INC = J("incentive_v12.json"); I1 = J("incentive_v11.json"); RATE = J("rate_lever.json"); PC = J("projections_check_v12.json")
AUG = S["august_2026"]; BR = S["pl_bridge_validation"]; PRICE = S["price_of_rules"]
CAFES = ["Downtown", "Studio", "1899 Penn"]; DOW5 = ["Mon", "Tue", "Wed", "Thu", "Fri"]; DOW7 = DOW5 + ["Sat", "Sun"]
DOWS = {"Downtown": DOW5, "Studio": DOW7, "1899 Penn": DOW5}
WPM = 52 / 12; TGT = T["targets"]; ADMIN = NV["admin_days"]
ADMIN_WHO = {"Downtown": "Gee (Mon / Thu / Fri)", "Studio": "Colin (Mon / Wed), Claire (a weekend day — Saturday assumed)", "1899 Penn": "Eli (Tue)"}
def W(c, t): return T["cafes"][c][t]["week"]
def WN(c, t): return T["cafes"][c][t]["week_no_admin"]
def sel(d): return "admin" if d["admin_day"] else "std"
def mo(x): return x * WPM

def build():
    doc = Document(); sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.85)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Labor Deep Dive — Schedule Templates by Demand Tier", 22, True, COLOR_ACCENT, 0), ("September 2026  |  v1.2", 14, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "Prepared September 5, 2026 for Ryan Fleming. v1.2 rebuilds the templates to Ryan's operating rules of Sep 5 (two openers and two closers at Downtown and Studio, kitchen in 45 minutes and barista 30 minutes before open, "
                  "6-hour minimum shifts, a third person mid-day at Studio, admin days as the buffer, 1899/NMWA may stagger or go solo in slow months), prices them at the roster's actual rates, checks the rest-of-year projections "
                  "against 2025, and bench-tests four incentive structures against 2026 actuals and a simulated 2027. v1.1 (lean, 4-hour shifts) is kept alongside as the 'what the rules cost' reference. "
                  "Sources unchanged: Square transactions Jan–Aug 2026 (walk-in), 856 cafe timecards across 14 weeks, Consolidated Model v17, QBO P&L, Projections v1.8, the 10/11 first-cut schedule.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ================= EXEC SUMMARY =================
    add_heading(doc, "Executive Summary", level=1)
    dt, st, pn = AUG["Downtown"], AUG["Studio"], AUG["1899 Penn"]
    add_callout(doc, "The answer to the original question has not changed; what the rules cost has been added.",
                f"High-$40s sales per labor hour is a 40–47% wage bill at Slipstream's rates; the cafes need $60–68 per core hour to hit 32/30/35 and August delivered $44 / $52 / $43. On the wage basis core labor ran "
                f"{p(dt['core_wage_pct_of_core_sales'])} / {p(st['core_wage_pct_of_core_sales'])} / {p(pn['core_wage_pct_of_core_sales'])} — {m(dt['headroom_core_wages_vs_target']+st['headroom_core_wages_vs_target']+pn['headroom_core_wages_vs_target'])} over target in one month — and the P&L understates it because tips fund part of the guaranteed wage "
                "(booked labor ≈ 1.12 × gross − tips). Under the Sep 5 rules the achievable floor moves up: two people open-to-close with 6-hour shifts is 22–24 hours a day at Downtown whatever the sales, so the 32% target is "
                "reachable only in busy weeks (31.6%), a mid week lands at 34.5% (inside the 36% band) and a slow week at 37.2% (inside the new 38% slow-month buffer). Studio lands at 29.2 / 31.8 / 33.8% (busy / mid / slow, "
                "the slow figure using Colin's and Claire's admin days as the third body). 1899 reaches 36.5% only in busy weeks. Against the August roster that is still $1,000–3,300 a month less at each of Downtown and Studio "
                "than what was actually run — the rules cost roughly $900–1,700 a month per cafe at Downtown and Studio against the leaner v1.1 shapes, which is the price of redundancy and accountability and is Ryan's call.", body_size=10, title_size=10)
    add_para(doc, "What changed in v1.2 and why:", after=2, before=6)
    for line in S["changes_from_v1_1"]: add_bullet(doc, line)
    add_para(doc, "Also answered: current roster rates by role (Aug 2026 core hours) are Shift Lead $22.17 Downtown / $22.16 Studio / $21.07 1899, Production Line $19.86 / $19.41, Production Barista $19.98 / $19.99 / $20.98, "
                  "Barista $20.54 / $18.82 / $20.25, all-cafe $20.65; six people at $18.50 carry ~30% of core hours and the two ALLs ($24.75–25.25) ~200 core hours a month. The templates now price at $22.25 / $19.75 / $20.25 / $20.00 "
                  "(SL / PL / PB / barista-closer) — the August averages rounded up a quarter for development drift. On whether the targets are too aggressive: on the P&L they are not (32% wages books ~26% labor after tips); on the "
                  "schedule they are, once two-open/two-close and 6-hour shifts are fixed — so the targets stay, the slow-month ceiling moves to 38 / 36 / 42, and the incentive is measured against the template rather than the target (Section 8).", after=6)

    add_heading(doc, "What the v1.2 templates deliver — weekly core wages by tier (planning rates)", level=3)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier); wn = WN(c, tier); tg = TGT[c]["target"] * 100; ce = w["ceiling"] * 100
            status = "OK" if w["labor_pct"] <= tg + 0.05 else ("BAND" if w["labor_pct"] <= ce + 0.05 else "OVER")
            rows.append([f"{c} — {tier}", m(w["sales_typical"]), f"{h(w['actual_sampled_hours'],0)} h / {p(w['actual_sampled_labor_pct'])}", f"{h(w['v11_hours'],0)} h / {p(w['v11_labor_pct'])}",
                         h(w["hours"], 1), m(w["wages"]), p(w["labor_pct"]), p(w["labor_pct_std_rates"]), f"{h(wn['hours'],1)} h / {p(wn['labor_pct'])}", f"{tg:.0f} / {ce:.0f}", status, sm(w["wages"] - w["actual_sampled_wages"])])
    table(doc, ["Cafe / tier", "Typical wk sales", "Actual (h / %)", "v1.1 lean (h / %)", "v1.2 hours", "v1.2 wages", "v1.2 %", "% at $21/$19", "No-admin week", "Target / ceiling", "Status", "Δ$/wk vs actual"],
          rows, widths=[1.0, 0.6, 0.75, 0.75, 0.5, 0.55, 0.5, 0.55, 0.8, 0.6, 0.45, 0.6], key_rows=[0, 3, 6], size=7)
    add_caveat(doc, "v1.2 wages at planning rates ($22.25 SL / $19.75 PL / $20.25 PB / $20.00 barista & closer); '% at $21/$19' re-prices the same hours at the v1.0/v1.1 standard rates for comparison. 'No-admin week' = every day staffed without an admin person on site "
                    "(the fallback when Gee/Colin/Eli/Claire are elsewhere). Ceiling is 36/34/40, +2 in the slow tier per Ryan. Status is against the target (OK), the ceiling (BAND) or above it (OVER).")

    add_heading(doc, "The price of the Sep 5 rules — v1.1 lean shapes vs v1.2, weekly and monthly", level=3)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            v = PRICE[c][tier]
            rows.append([f"{c} — {tier}", f"{h(v['v11_hours'],1)} h / {p(v['v11_pct_std'])}", f"{h(v['v12_hours'],1)} h / {p(v['v12_pct_std'])}", sd(v["delta_hours"]), sm(v["delta_wages_std_basis"]), sm(v["delta_month_std_basis"])])
    table(doc, ["Cafe / tier", "v1.1 lean (h / % at $21/$19)", "v1.2 rules (h / % at $21/$19)", "Δ hours / wk", "Δ wages / wk", "Δ wages / month"], rows, widths=[1.4, 1.35, 1.35, 0.8, 0.85, 0.95], key_rows=[0, 3, 6], size=8)
    add_caveat(doc, "Like-for-like at the same standard rates. The rules add 10–24 hours a week at Downtown and Studio (two closers to 4:30/5:30 and a 6-hour floor on every shift) and take hours away at 1899 in slow weeks (solo Mondays and Fridays). "
                    "Annualised on the 2027 tier calendar the rules cost about $57K of open-cafe operating income against the lean shapes — Section 8d.")

    add_para(doc, "The three biggest changes from what was actually run (slow-tier weeks; mid in brackets):", after=4, before=6)
    dts, dtm, sts, stm, pnm, pnb = W("Downtown", "slow"), W("Downtown", "mid"), W("Studio", "slow"), W("Studio", "mid"), W("1899 Penn", "mid"), W("1899 Penn", "busy")
    add_bullet(doc, f"Downtown — three people, not four: Production Line 6:15–12:15, Shift Lead 6:30–4:30 (the 10-hour open-to-close lead is now the cheapest way to have two closers), Closer 8:30–4:30 Tue–Thu and 10:30–4:30 Mon/Fri; "
                    f"rush baristas (6 h, 7:45/8:15 start) only in mid and busy weeks. Slow week {h(dts['hours'],0)} h vs {h(dts['actual_sampled_hours'],0)} actual: −{m(mo(dts['actual_sampled_wages']-dts['wages']))}/mo [mid {h(dtm['hours'],0)} h, −{m(mo(dtm['actual_sampled_wages']-dtm['wages']))}/mo].", bold_lead="1.")
    add_bullet(doc, f"Studio Mon–Thu — exactly Ryan's shape: Production Line 6:15–12:15, Shift Lead 6:30–2:30, third person 8:00–2:30 (7:00 on Monday for the delivery); in slow weeks the third person is Colin on Mon/Wed. Weekend: one lead 7:30–5:30 "
                    f"plus a 6-hour barista and a closer instead of five people. Slow week {h(sts['hours'],0)} h vs {h(sts['actual_sampled_hours'],0)} actual: −{m(mo(sts['actual_sampled_wages']-sts['wages']))}/mo [mid {h(stm['hours'],0)} h, −{m(mo(stm['actual_sampled_wages']-stm['wages']))}/mo].", bold_lead="2.")
    add_bullet(doc, f"1899 — lead 6:30–3:30 plus a 6-hour second person 8:30–2:30; solo Mondays and Fridays in slow weeks (Ryan's stagger rule). Mid week {h(pnm['hours'],0)} h at {p(pnm['labor_pct'])} vs {h(pnm['actual_sampled_hours'],0)} actual: −{m(mo(pnm['actual_sampled_wages']-pnm['wages']))}/mo "
                    f"[busy {h(pnb['hours'],0)} h at {p(pnb['labor_pct'])}]. The mid template still breaches the 40% ceiling; a solo mid-tier Friday ($456 / 61 transactions) would bring it to ~39.6%.", bold_lead="3.")

    pc = PC["tier_by_month_sep_dec"]
    add_para(doc, "Projections reality check (Section 2c): Projections v1.8 carries Downtown +4–10% and Studio +1–10% over 2025 for Sep–Dec, but Jul–Aug ran flat at Downtown (−0.5% / +0.5%) and −14–17% at Studio. On that run-rate Downtown is a mid-tier cafe "
                  f"through December ({pc['Downtown']['2026-10']['runrate_dominant']} in October where the projection says busy) and Studio is slow in September, mid in October and slow again from mid-November — two tiers below the projection in some weeks. "
                  "1899's +40–78% over 2025 is consistent with its +88–90% Jul–Aug run-rate. Recommendation: re-base the Sep–Dec projections to 2025 × (1.00 Downtown, 0.86 Studio) before the weekly forward block uses them, and build the 10/11 week on MID at Downtown and SLOW at Studio.", before=6, after=6)
    bs = INC["backtest_summary"]; sim = INC["simulation_2027"]["options"]
    add_para(doc, f"Incentive (Section 8): under the new rules a target-based Beat Share never pays — none of Apr–Aug 2026 would have paid, and the 2027 simulation pays in 2 months of 12 — because the templates themselves sit above the 32/30 line in slow and mid weeks. "
                  f"Measured against the template instead, the plan pays when the team runs the plan: Option 4 (fixed $300/$300/$150/$75 per cafe-month at or under template, plus half of any beat, pooled, split as a % of wages) pays in ~11 months of 12, "
                  f"{m(sim['opt4']['p10'])}–{m(sim['opt4']['p90'])} a year, about {m(INC['per_person_30h_year_p50']['opt4'])} for a 30-hour person, and includes 1899 and NMWA on metrics they can actually hit. It is small by design — the annual profit share is where the real money should sit.", after=8)

    # ================= 1. BRIDGE (unchanged) =================
    add_heading(doc, "1. How labor reaches the P&L — the bridge (unchanged from v1.1)", level=1)
    add_para(doc, "P&L labor ≈ 1.12 × gross wages (core + admin + training) − credit-card tips, validated against QBO Apr–Aug within 0.4% (Downtown), 0.6% (Studio) and 2.3% (1899). A cafe exactly on its wage-basis target books ~26% labor at Downtown/Studio and ~32% at 1899. "
                  "Consequence for Ryan's question on the targets: the wage-basis targets are not aggressive on the P&L — they are what produce a 55–62% prime cost — but with two-open/two-close they are only reachable in busy weeks, which is why the slow-month ceiling moves to 38/36/42 and the incentive is measured against the template.", after=6)
    rows = []
    for c in CAFES:
        b = BR[c]; rows.append([f"{c} Apr–Aug", m(b["pl_est_sum"]), m(b["pl_actual_sum"]), f"{b['error_pct']:+.1f}%", m(sum(x["headroom"] for x in b["by_month"].values()))])
    table(doc, ["Cafe", "Bridge estimate", "QBO booked", "Δ", "Core wages over target, Apr–Aug"], rows, widths=[1.5, 1.2, 1.2, 0.8, 1.8], size=9)
    add_bullet(doc, "DC context: the District's minimum wage rose to $18.40 on July 1, 2026 (tipped 56% = $10.30). Four sampled timecards carried an $18.00 rate — worth a payroll check.")

    # ================= 2. TIERS + PROJECTIONS =================
    add_heading(doc, "2. Demand tiers and the rest-of-year outlook", level=1)
    rows = []
    for c in CAFES:
        b = S["tiers"][c]["bands_weekly_walkin"]; rows.append([c, f"< {m(b['slow'][1])}", f"{m(b['mid'][0])} – {m(b['mid'][1])}", f"≥ {m(b['busy'][0])}"])
    table(doc, ["Cafe", "Slow", "Mid", "Busy"], rows, widths=[1.4, 1.6, 1.8, 1.6])
    add_caveat(doc, "Weekly walk-in sales bands, unchanged from v1.0/v1.1 (validated against k-means and Jenks breaks in v1.1). 2026 membership: Studio slow Jan, Feb, Jul, Aug and busy Apr–Jun; Downtown slow Aug, mid Feb–Mar/Jul, busy Apr–Jun; 1899 mid most months, busy Jun–Jul.")
    add_heading(doc, "2c. Projections v1.8 vs 2025 actuals vs the Jul–Aug run-rate", level=2)
    rows = []
    for c in ["Downtown", "Studio", "1899"]:
        for r in PC["monthly"][c]:
            if r["month"] < "2026-07": continue
            rows.append([f"{c} {r['month'][5:]}", m(r["sales_2025"]), m(r["proj_v18"]), p(r["proj_vs_2025_pct"]) if r["proj_vs_2025_pct"] is not None else "—", m(r["actual_2026"]) if r["actual_2026"] else "—",
                         p(r["yoy_actual_pct"]) if r["yoy_actual_pct"] is not None else "—", m(r["runrate_case"]) if r["runrate_case"] else "—", p(r["runrate_vs_proj_pct"]) if r["runrate_vs_proj_pct"] is not None else "—"])
    table(doc, ["Cafe / month", "2025 actual", "Proj v1.8", "Proj vs 2025", "2026 actual", "2026 vs 2025", "Run-rate case", "Run-rate vs proj"], rows, widths=[1.1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.85, 0.85], key_rows=[0, 6, 12], size=8)
    add_caveat(doc, "Run-rate case = 2025 actual × 1.00 (Downtown; Jul–Aug 2026 ran −0.5% / +0.5% YoY after +13–32% in Jan–Jun) and × 0.86 (Studio; −16.5% / −14.3%). 1899 has no comparable 2025 base (opened mid-2025); its projection is consistent with its +88–90% run-rate. "
                    "Ryan's note that the projections need checking is confirmed for Studio (projection 13–22% too high Sep–Dec) and mildly for Downtown (4–9% too high).")
    rows = []
    for c in CAFES:
        row = [c]
        for mth in ["2026-09", "2026-10", "2026-11", "2026-12"]:
            r = pc[c].get(mth, {}); row.append(f"{r.get('proj_dominant','—')} → {r.get('runrate_dominant') or 'proj'}")
        rows.append(row)
    table(doc, ["Cafe", "Sep", "Oct", "Nov", "Dec"], rows, widths=[1.4, 1.3, 1.3, 1.3, 1.3], size=9)
    add_caveat(doc, "Dominant weekly tier per month: projection → run-rate case (weeks scaled to full weeks; Thanksgiving and Christmas weeks are slow at every cafe). Ryan's instruction stands for the holiday weeks (Dec 21–Jan 3): slow template.")

    cafe_sections(doc)
    add_heading(doc, "6. NMWA (for completeness)", level=1)
    nm = TIERS["NMWA"]["typical_day"]["slow"]
    add_para(doc, f"Unchanged: FOH 9:30–2:30 Thu–Sun (5 h; 6-hour minimum does not apply to a 4-hour window — Ryan to confirm) plus a second FOH 10:00–2:00 on the first Sunday and second Wednesday. Typical days Thu {m(nm['Thu']['net'])}, Fri {m(nm['Fri']['net'])}, Sat {m(nm['Sat']['net'])}, "
                  f"Sun {m(nm['Sun']['net'])}; community Wednesdays {m(nm['Wed']['net'])}. The 10/11 first cut is missing the 10/14 community-day shift. NMWA hours share in the incentive pool; its own metric is the staffing rule (Section 8).", after=6)
    sensitivity_section(doc)
    incentive_section(doc)
    data_notes(doc)
    _ensure_zoom_percent(doc); doc.save(OUT); print("Saved →", OUT)


def cafe_sections(doc):
    secno = {"Downtown": 3, "Studio": 4, "1899 Penn": 5}
    HOURS_SHOW = {"Downtown": list(range(7, 16)), "Studio": list(range(7, 17)), "1899 Penn": list(range(7, 15))}
    intro = {"Downtown": "Office-district cafe, 7a–4p Mon–Fri. Rules applied: kitchen 6:15, lead 6:30, two people until 4:30, 6-hour shifts. Admin on site: Gee Mon / Thu / Fri.",
             "Studio": "Neighbourhood cafe with a food program, 7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun. Rules applied: kitchen 6:15 (7:15 weekends), lead 6:30 (7:30), two until 2:30 / 5:30, a third person 8:00–12:30 Mon–Thu (Colin on Mon / Wed in slow weeks), Monday three at open for the delivery, capacity 12 per person-hour. Admin on site: Colin Mon / Wed, Claire one weekend day (Saturday assumed).",
             "1899 Penn": "Two-person office-building cafe, 7a–3p Mon–Fri. Rules applied: lead 6:30–3:30 every day; second person 8:30–2:30 (6 h) in mid and busy weeks; stagger / solo allowed in slow weeks (solo Monday and Friday). Admin on site: Eli Tue."}
    for c in CAFES:
        n = secno[c]
        add_heading(doc, f"{n}. {c}", level=1); add_para(doc, intro[c], after=6)
        add_heading(doc, f"{n}a. Demand curve — walk-in transactions per hour, by tier", level=2)
        hrs = HOURS_SHOW[c]; show_days = ["Tue", "Fri"] if c != "Studio" else ["Tue", "Fri", "Sat"]; rows = []
        for d in show_days:
            for tier in ["slow", "mid", "busy"]:
                cv = CURVES[c]["by_tier"][tier][d]; g = lambda hh: cv['hourly'].get(str(hh), {}).get('txns', 0)
                rows.append([f"{d} — {tier}"] + [(f"{g(hh):.0f}" if str(hh) in cv['hourly'] else "·") for hh in hrs] + [f"{sum(g(hh) for hh in hrs):.0f}"])
        table(doc, ["Day / tier"] + [f"{hh if hh <= 12 else hh - 12}{'a' if hh < 12 else 'p'}" for hh in hrs] + ["Day"], rows, widths=[1.1] + [0.42] * len(hrs) + [0.5], key_rows=[0, 3, 6], size=8)
        add_heading(doc, f"{n}b. Actual staffing vs need under the Sep 5 rules — hours per week", level=2)
        rows = []
        for tier in ["slow", "mid", "busy"]:
            dows = NV["cafes"][c]["tiers"][tier]
            for d in DOWS[c]:
                r = dows[d]["sets"]["std"]; ra = dows[d]["sets"]["admin"]
                rows.append([f"{tier} {d}{' *' if dows[d]['admin_day'] else ''}", h(r["need_hours"]), h(ra["need_hours"]) if dows[d]["admin_day"] else "—", h(r["actual_hours"]), h(r["over_hours"]), m(r["over_dollars"]), f"{r['peak_need']} / {r['peak_actual']:.1f}"])
            rows.append([f"{tier} — week", h(sum(x["sets"]["std"]["need_hours"] for x in dows.values())), h(sum((x["sets"]["admin"]["need_hours"] if x["admin_day"] else x["sets"]["std"]["need_hours"]) for x in dows.values())),
                         h(sum(x["sets"]["std"]["actual_hours"] for x in dows.values())), h(sum(x["sets"]["std"]["over_hours"] for x in dows.values())), m(sum(x["sets"]["std"]["over_dollars"] for x in dows.values())), ""])
        nd = len(DOWS[c])
        table(doc, ["Tier / day", "Need (h)", "Need, admin on site", "Actual (h)", "Over (h)", "Over ($/day)", "Peak need / actual"], rows, widths=[1.2, 0.8, 1.0, 0.85, 0.8, 0.9, 1.1], total_rows=[nd, 2 * nd + 1, 3 * nd + 2], size=9)
        add_caveat(doc, f"Need = the Sep 5 floors plus the cafe's busy-hour capacity ({NV['cap'][c]:.0f} per person-hour) applied to the 30-minute curve; '* ' marks admin days ({ADMIN_WHO[c]}), where the slow-tier need counts the admin person as a body. "
                        f"Actual = average core headcount in the sampled weeks. Dollars at the observed average core rate (${NV['cafes'][c]['avg_core_rate']:.2f}).")
        add_heading(doc, f"{n}c. The two levers under the rules", level=2)
        rl = RATE[c]
        over = {t: sum(x["sets"]["std"]["over_dollars"] for x in NV["cafes"][c]["tiers"][t].values()) for t in ["slow", "mid", "busy"]}
        add_bullet(doc, f"Hours placement: even with two-open/two-close the sampled weeks paid {m(over['slow'])} / {m(over['mid'])} / {m(over['busy'])} a week above need in slow / mid / busy weeks (≈ {m(mo(over['slow']))} / {m(mo(over['mid']))} / {m(mo(over['busy']))} a month). "
                        f"The v1.2 templates recover {m(mo(W(c,'slow')['actual_sampled_wages']-W(c,'slow')['wages']))} / {m(mo(W(c,'mid')['actual_sampled_wages']-W(c,'mid')['wages']))} / {m(mo(W(c,'busy')['actual_sampled_wages']-W(c,'busy')['wages']))} of it a month at planning rates.", bold_lead="Lever 1 —")
        add_bullet(doc, f"Cost per hour: observed average core rate ${rl['avg_rate']:.2f}; Shift Lead hours {rl['sl_share_hours']:.0f}% of core, two leads overlapping {rl['double_SL_hours_per_week']:.1f} h/wk, manager-rate people {rl['manager_rate_hours_per_week']:.1f} core h/wk. "
                        f"The templates now carry these rates, so the lever shows as the gap between the planning-rate % and the '% at $21/$19' column (about 1.5 points): the way to close it is who fills the Closer and rush slots, not the template.", bold_lead="Lever 2 —")
        add_heading(doc, f"{n}d. Schedule templates (v1.2 — Sep 5 rules)", level=2)
        tg = TGT[c]
        add_para(doc, f"Target {tg['target']*100:.0f}% (ceiling {tg['ceiling']*100:.0f}%, {tg['ceiling_slow']*100:.0f}% in slow weeks). Planning rates $22.25 Shift Lead, $19.75 Production Line, $20.25 Production Barista, $20.00 barista / closer. Each day is the cheapest 6-hour-minimum shift set that meets the floors "
                      "and the need curve; on admin days the slow-tier day is shown with the admin person on site (the no-admin alternative is in the workbook).", after=4)
        for tier in ["slow", "mid", "busy"]:
            o = T["cafes"][c][tier]; w = o["week"]; wn = o["week_no_admin"]
            add_heading(doc, f"{tier.capitalize()} template — {m(w['sales_typical'])}/wk typical, {w['hours']:.1f} core hours, {m(w['wages'])} = {p(w['labor_pct'])} (SPLH {m(w['splh'],0)})", level=3)
            rows = []
            for d, dd in o["days"].items():
                s_ = dd["sets"][sel(dd)]; tmpl = s_["template"]; act = dd.get("actual_sampled", {})
                for i, sh in enumerate(s_["shifts"]):
                    rows.append([(d + (" *" if dd["admin_day"] else "")) if i == 0 else "", sh["role"], sh["start"], sh["end"], h(sh["hours"], 2), f"${sh['rate']:.2f}", m(sh["wages"]), ""])
                notes = []
                if tmpl["solo_windows"]: notes.append("solo " + ", ".join(tmpl["solo_windows"]) + (" (admin on site)" if dd["admin_day"] else ""))
                if tmpl["admin_cover_windows"]: notes.append("Admin covers " + ", ".join(tmpl["admin_cover_windows"]))
                if dd["admin_day"] and tier == "slow" and abs(dd["sets"]["std"]["template"]["hours"] - tmpl["hours"]) > 0.1:
                    notes.append(f"no-admin day: {dd['sets']['std']['template']['hours']:.1f} h / {p(dd['sets']['std']['template']['labor_pct'])}")
                rows.append([f"{d} total", f"{len(s_['shifts'])} shifts", "", "", h(tmpl["hours"], 2), "", m(tmpl["wages"]), f"{p(tmpl['labor_pct'])} of {m(dd['sales_typical'])}; peak {tmpl['peak_staff']}; actual {act.get('core_hours', '—')} h / {p(act.get('labor_pct'))}" + ("; " + "; ".join(notes) if notes else "")])
            total_idx = [i for i, r in enumerate(rows) if str(r[0]).endswith("total")]
            table(doc, ["Day", "Role", "Start", "End", "Hours", "Rate", "Wages", "Check / notes"], rows, widths=[0.6, 1.45, 0.45, 0.45, 0.45, 0.5, 0.55, 2.25], total_rows=total_idx, size=8)
            add_caveat(doc, f"Week: {w['hours']:.1f} h, {m(w['wages'])} = {p(w['labor_pct'])} of {m(w['sales_typical'])} (target {p(tg['target']*100,0)}, ceiling {p(w['ceiling']*100,0)}); at $21/$19 standard rates {p(w['labor_pct_std_rates'])}. At −10% sales: {p(w['labor_pct_minus10'])}; at +10%: {p(w['labor_pct_plus10'])}. "
                            f"Without admin on site: {wn['hours']:.1f} h = {p(wn['labor_pct'])}. Sampled actual weeks: {w['actual_sampled_hours']:.1f} h = {p(w['actual_sampled_labor_pct'])}; v1.1 lean {w['v11_hours']:.1f} h = {p(w['v11_labor_pct'])}; first cut {w['ryan_first_cut_hours']:.1f} h = {p(w['ryan_first_cut_labor_pct'])}. "
                            + (f"Drop to the next-slower template below {m(w['ceiling_trigger_sales'])} of weekly sales (ceiling breach)." if tier != "slow" else f"Breaches the {p(w['ceiling']*100,0)} slow-week ceiling below {m(w['ceiling_trigger_sales'])} of weekly sales."))
        add_heading(doc, f"{n}e. Specific changes from the 10/11 first-cut template", level=2)
        for line in FIRST_CUT[c]: add_bullet(doc, line)
        add_heading(doc, f"{n}f. What moved from v1.1, and what the rules cost here", level=2)
        rows = []
        for tier in ["slow", "mid", "busy"]:
            v = PRICE[c][tier]; w = W(c, tier)
            rows.append([tier, f"{h(v['v11_hours'],1)} h / {p(v['v11_pct_std'])}", f"{h(w['hours'],1)} h / {p(w['labor_pct'])} ({p(v['v12_pct_std'])} at $21/$19)", sd(v["delta_hours"]), sm(v["delta_wages_std_basis"]), sm(v["delta_month_std_basis"])])
        table(doc, ["Tier", "v1.1 lean", "v1.2 rules (planning rates; at $21/$19)", "Δ h/wk", "Δ $/wk", "Δ $/mo"], rows, widths=[0.7, 1.3, 2.2, 0.7, 0.8, 0.8], size=9)
        for line in V11_CHANGES[c]: add_bullet(doc, line)


FIRST_CUT = {
    "Downtown": [
        "Monday (first cut: Production Barista 6:30–1, Production Line 6:12–4:30, Shift Lead 9–4:30 = 24.3 h). v1.2: Production Line 6:15–12:15, Shift Lead 6:30–4:30, Closer 9:00–4:30 (mid) or 10:30–4:30 (slow) = 23.5 / 22 h — the lead opens and closes, the Production Line stops at 12:15, the Production Barista shift goes.",
        "Tuesday (first cut: 4 people, 32.75 h). Mid: Production Line 6:15–12:15, Shift Lead 6:30–4:30, Barista 8:15–2:15, Closer 8:30–4:30 = 30 h (−2.75); slow: no barista, Production Line to 1:15, 25 h (−7.75).",
        "Wednesday / Thursday (first cut 31.75 h each). Mid: Production Line 6:15–12:15 (Wed to 2:15), Shift Lead 6:30–4:30, Barista 7:45–1:45, Closer 8:30–4:30 = 30–32 h (−1.75 to −0.25 — the rules keep most of these hours); slow: 24 h (−7.75) with no barista.",
        "Friday (first cut 23.75 h). Every tier below busy: Production Line 6:15–12:15, Shift Lead 6:30–4:30, Closer 10:30–4:30 = 22 h (−1.75). Friday is structurally 49–60% labor under two-open/two-close ($773–940 of sales); Gee's admin day is the practical relief, not the template.",
        "Net: first cut 144.3 h / $3,058 → mid 137.5 h / $2,854 (34.5%) → slow 117 h / $2,445 (37.2%). The first cut was closer to the rules than v1.1 gave it credit for; its problem was four-deep Tue–Thu and $22.75–25.25 people on barista shifts.",
    ],
    "Studio": [
        "Mon–Thu (first cut 22.25–22.75 h each). v1.2 is Ryan's stated shape: Production Line 6:15–12:15, Shift Lead 6:30–2:30, third person 8:00–2:30 (Monday 7:00–2:30 for the delivery) = 20.5–21.5 h (−1 to −2.25). Slow weeks: Colin's Mon/Wed admin days replace the third person (Mon 20 h with a 6:45–12:45 barista for the delivery; Wed 14 h, two core + Colin).",
        "Friday (first cut 30 h). Production Line 6:15–12:15, Shift Lead 6:30–2:30, Barista 7:45–1:45, then two closers Barista 11:15–5:15 + Closer 11:30–5:30 = 32 h (+2) — the two-closer rule on a 5:30 close costs Friday two hours versus the first cut.",
        "Saturday (first cut 34.75 h). Production Line 7:15–1:15, Shift Lead 7:30–5:30, Barista 8:45–2:45, Closer 9:30/10:30–5:30 = 29–30 h (−5 to −6): one lead through the day, four at the peak, two closers. Sunday similar at 28.5–30.5 h.",
        "Net: first cut 183.5 h / $3,718 → mid 174.5 h / $3,614 (31.8%) → slow 158.5 h / $3,290 (33.8% with admin days; 36.7% without). The first cut is a mid-tier schedule; Studio's Sep–Nov run-rate says slow.",
    ],
    "1899 Penn": [
        "Mon–Thu (first cut both people 6:30–3:30, 16–18 h). Lead 6:30–3:30 plus Production Barista 8:30–2:30 = 15 h (−1 to −3); busy Tue/Wed 7:30–3:00 (16.5 h) with the 1pm spike as Admin cover.",
        "Friday (first cut 16 h). Mid/busy: 15 h; slow: lead alone (9 h) — Ryan's stagger rule. Monday in slow weeks likewise solo (9 h).",
        "Net: first cut 85.5 h / $1,756 → mid 75 h / $1,609 (42.8%) → slow 57 h / $1,244 (43.3%) → busy 78.5 h (36.5%). The 6-hour minimum removes the 4-hour rush shifts v1.1 used, so 1899 sits 3–4 points above its ceiling in mid weeks unless Friday goes solo.",
    ],
}
V11_CHANGES = {
    "Downtown": ["Two closers to 4:30 and the 6-hour minimum add 10–12 h/wk against v1.1; the 10-hour open-to-close lead is now structural (a 9-hour lead needs a fourth 6-hour shift, +$90/day).",
                 "Admin days change nothing at Downtown: the closer must still be a 6-hour shift ending 4:30, so Gee's presence is the buffer, not a saving."],
    "Studio": ["The third person mid-day (Ryan's shape) and two closers add 19–23 h/wk against v1.1 — except in slow weeks, where Colin's and Claire's admin days give back 14 h/wk.",
               "Friday and the weekend are the expensive days under the rules: five people on Friday in every tier, and 43% on a slow Friday."],
    "1899 Penn": ["Slow weeks are cheaper than v1.1 (solo Mon/Fri, −13.5 h); mid and busy weeks are 1.5–2.5 h/wk dearer because the second person is 6 h rather than 4.",
                  "The mid template breaches the 40% ceiling (42.8%); options are a solo mid-tier Friday (39.6%) or accepting the band until sales clear ~$4,000/wk."],
}


def sensitivity_section(doc):
    add_heading(doc, "7. Sensitivity and tier triggers (v1.2)", level=1)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier); wn = WN(c, tier)
            rows.append([f"{c} — {tier}", m(w["sales_typical"]), m(w["wages"]), p(w["labor_pct"]), p(wn["labor_pct"]), p(w["labor_pct_minus10"]), p(w["labor_pct_plus10"]), p(w["ceiling"] * 100, 0), m(w["ceiling_trigger_sales"]), m(w["target_trigger_sales"])])
    table(doc, ["Template", "Typical wk sales", "Wages", "%", "% no admin", "At −10% sales", "At +10% sales", "Ceiling", "Ceiling trigger (wk sales)", "On-target at (wk sales)"], rows,
          widths=[1.3, 0.75, 0.65, 0.5, 0.65, 0.7, 0.7, 0.55, 0.9, 0.9], key_rows=[0, 3, 6], size=8)
    add_caveat(doc, "Because hours are now fixed by the rules rather than by demand, a 10% sales miss costs 3.5–4.5 points at every cafe. Downtown's mid template breaches 36% below ~$7,900/wk and its slow template breaches 38% below ~$6,400/wk — the Sep–Dec run-rate ($7,300–8,900) sits right on the mid trigger, "
                    "so the weekly forward block should show the trigger distance. Studio's slow template holds to ~$9,100/wk with admin days; its mid template to ~$10,600. 1899's mid template is above its ceiling at typical mid sales.")


def incentive_section(doc):
    add_heading(doc, "8. Incentive — four options, bench-tested", level=1)
    add_para(doc, "Ryan's constraints: 1899 and NMWA are managed locations that will not hit profit-style targets for a while; a plan that never pays is worse than no plan; the monthly piece must stay small so the annual profit share into retirement can be real. "
                  "All options keep the v1.1 mechanics — pooled, paid as a percentage of each person's core wages (FLSA-clean), guardrail = published schedules met the template floors and sales ≥ 90% of projection, kill-switch on COGS drift.", after=6)
    add_heading(doc, "8a. Why a target-based plan fails under the Sep 5 rules", level=2)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier); tg = TGT[c]["target"]
            rows.append([f"{c} — {tier}", p(w["labor_pct"]), p(tg * 100, 0), sm(round(tg * w["sales_typical"] - w["wages"])), "pays" if tg * w["sales_typical"] - w["wages"] > 0 else "never pays"])
    table(doc, ["Cafe / tier", "v1.2 template %", "Target", "Beat vs target at template ($/wk)", "Target-based Beat Share"], rows, widths=[1.5, 1.0, 0.7, 1.5, 1.4], key_rows=[0, 3, 6], size=9)
    add_caveat(doc, "If the team runs the template exactly, only Downtown-busy, Studio-busy and (barely) Studio-mid weeks sit under the target line; 1899 never does. A beat share measured against 32/30/35 therefore pays in about two months a year — the 'never pays out' failure Ryan named.")
    add_heading(doc, "8b. The four options", level=2)
    for k, txt in [("Option 1 — Open-cafe pool.", "Beat Share 50% of (target − actual core wages) at Downtown + Studio only, cap 1.5% of their sales; every core hour, including 1899 and NMWA, shares the payout. Fair to floaters, simple, but tied to the target — so it rarely pays under the rules, and 1899's own team has no lever."),
                   ("Option 2 — Own-metric pool.", "Option 1 plus a fixed $200 into the pool in any month 1899's NI is ≥ 5% of sales (the test that also earns Slipstream's incentive fee) and $100 in any month NMWA runs at or under its staffing-rule hours. Managed cafes get a lever they can pull; the pool is still mostly target-bound."),
                   ("Option 3 — Template beat.", "50% of Σ over all four cafes of (template core wages for the month's tier − actual core wages), cap 1.5% of all-cafe sales. Every cafe's yardstick is achievable by construction and improves as the templates improve. Pays only for beating the plan, not for running it."),
                   ("Option 4 — On-template + beat (recommended).", "A fixed amount for each cafe-month run at or under its template (1% tolerance): $300 Downtown, $300 Studio, $150 1899, $75 NMWA — plus 50% of any beat below template, cap 1.5% of sales. Pays for executing the plan, pays more for beating it, and the fixed amounts let 1899 and NMWA contribute on their own terms.")]:
        add_bullet(doc, txt, bold_lead=k)
    add_heading(doc, "8c. Bench test — what would have paid in 2026, what pays in a simulated 2027", level=2)
    bt = INC["backtest_2026"]; rows = []
    for b in bt:
        rows.append([b["month"]] + [f"{r['tier'][0].upper()} {p(r['core_pct'])} vs tmpl {m(r['tmpl_wages'])} / act {m(r['wages'])}" for r in b["cafes"]] + [m(b["opt1"]), m(b["opt2"]), m(b["opt3"]), m(b["opt4"])])
    table(doc, ["2026", "Downtown (tier, core %, template vs actual)", "Studio", "1899", "Opt 1", "Opt 2", "Opt 3", "Opt 4"], rows, widths=[0.45, 1.65, 1.65, 1.65, 0.45, 0.45, 0.45, 0.45], size=7)
    bs = INC["backtest_summary"]
    add_caveat(doc, f"Backtest on actual Apr–Aug 2026 core wages and sales (NMWA excluded — no template before July). Months paying: Option 1 {bs['opt1']['months_paying']}, Option 2 {bs['opt2']['months_paying']} (1899's NI ≥ 5% in Apr–Jun), Option 3 {bs['opt3']['months_paying']}, Option 4 {bs['opt4']['months_paying']} (Downtown May ran under the busy template). "
                    "Read it the right way round: none of the cafes ran to a v1.2 template in 2026 — that is the point of the templates, not a flaw in the plans.")
    sim = INC["simulation_2027"]; sv = sim["labor_saving_vs_aprAug_run_rate"]; pp = INC["per_person_30h_year_p50"]
    rows = []
    NAMES = {"opt1": "1. Open-cafe pool (target)", "opt2": "2. Own-metric pool", "opt3": "3. Template beat", "opt4": "4. On-template + beat"}
    for k in ["opt1", "opt2", "opt3", "opt4"]:
        v = sim["options"][k]; rows.append([NAMES[k], m(v["p10"]), m(v["p50"]), m(v["p90"]), f"{v['months_paying_p50']} of 12", p(v["share_of_saving_p50"]), f"{m(pp[k])} / yr", m(v["p50"] * 1.12)])
    table(doc, ["Option", "Annual payout p10", "p50", "p90", "Months paying (p50)", "Share of labor saved", "Per 30 h/wk person (p50)", "Employer cost incl. 12% load"], rows, widths=[1.7, 0.8, 0.7, 0.7, 0.9, 0.9, 1.0, 1.0], size=8)
    add_caveat(doc, f"2,000 runs of the 2027 tier calendar (Studio slow Sep–Nov, Downtown mid), sales ±6%, realised core % = template % + 0.5 pt ± 1.5 pts, 1899 NI ≥ 5% in 60% of busy months, NMWA within its rule about half the time. "
                    f"The v1.2 templates save {m(sv['p10'])}–{m(sv['p90'])} a year of core wages against the Apr–Aug 2026 run-rate (v1.1's lean shapes saved ~$65K — the difference is the rules). Option 4 returns about a quarter of the saving to the team and pays almost every month; Options 1–3 return 3–8% and pay in 2–7 months.")
    add_heading(doc, "8d. Resulting margins and the annual profit share under the v1.2 templates", level=2)
    mv = INC["margins_v12"]; rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            r = mv[c][tier]
            rows.append([f"{c} — {tier}", m(r["core_sales_month"]), p(r["actual_sampled"]["labor_pct"]), p(r["actual_sampled"]["noi_pct"]), p(r["v12_template"]["labor_pct"]), p(r["v12_template"]["prime_pct"]), m(r["v12_template"]["noi"]), p(r["v12_template"]["noi_pct"]), p(r["v12_template_opt4_month"]["noi_pct"]), p(r["at_target"]["noi_pct"])])
    table(doc, ["Cafe / tier", "Core sales / mo", "Actual P&L labor %", "Actual NOI %", "v1.2 P&L labor %", "v1.2 prime %", "v1.2 NOI", "v1.2 NOI %", "NOI % with Opt 4 paid", "NOI % at target"], rows,
          widths=[1.3, 0.75, 0.7, 0.6, 0.7, 0.65, 0.7, 0.6, 0.75, 0.65], key_rows=[0, 3, 6], size=8)
    ps = INC["profit_share_v12"]; p1 = I1["profit_share_sizing"]
    add_caveat(doc, f"Same economics as v1.1 (Apr–Aug COGS %, average opex, tips, 1.12 load). Open-cafe operating income on the 2027 calendar: {m(ps['open_cafe_noi_year_v12_templates'])} under the v1.2 templates vs {m(p1['open_cafe_noi_year_at_target_calendar'])} at target; after the support entities "
                    f"({m(abs(ps['support_burden_year']))} a year) consolidated income is ≈ {m(ps['consolidated_ni_year_v12'])} vs {m(p1['consolidated_ni_year_at_target'])}. A 5–10% retirement share is therefore {m(ps['share_5pct'])}–{m(ps['share_10pct'])} a year under the rules — Option 4's ~$6.5K monthly plan is of the same order, "
                    "which argues for keeping the fixed amounts modest and for revisiting the two-closer rule at Downtown on Fridays.")
    add_heading(doc, "8e. Recommendation", level=2)
    for line in ["Adopt Option 4. Metric per cafe = core wages vs the published template for the tier in force (from the weekly forward block); tolerance 1%; fixed $300 / $300 / $150 / $75 per cafe-month on-template plus 50% of any beat; pool capped at 1.5% of all-cafe core sales; paid as a % of each person's core wages at any location.",
                 "Why it fits Ryan's constraints: it pays when people do the thing the templates ask (most months), pays more when they find a better way, includes 1899 and NMWA on yardsticks they control, and the fixed amounts are small enough ($825/month maximum) that the annual share stays the main event.",
                 "1899 and NMWA fairness: their template is their own operating rule (lead + 6 h second person; staffing rule), not a profit target; NI ≥ 5% at 1899 can be an additional $100 kicker once its COGS move lands.",
                 "Guardrail unchanged: schedules met the floors (two open / two close, no unplanned solo), core sales ≥ 90% of the re-based projection, admin hours tracked alongside; the template is the floor, so 'beating' it means better placement, not fewer bodies.",
                 "Timing: shadow in the Sep–Oct weekly reports; first payout after two consecutive on-template months company-wide; December or January realistically. Re-set the fixed amounts after two payouts and again when the retirement plan is stood up so the two together stay ≤ 2.5% of core sales.",
                 "Payroll: separate earnings code; percentage-of-wages allocation (non-discretionary bonus in the regular rate is then automatic); accrue in the month earned so the monthly bridge ties."]:
        add_bullet(doc, line)
    add_para(doc, "Comparable plans, retirement-vehicle notes (SEP vs SIMPLE vs 401(k) profit sharing) and the DC minimum-wage context are unchanged from v1.1 §8f–8h.", after=6)


def data_notes(doc):
    add_heading(doc, "9. Data notes and method (v1.2 deltas)", level=1)
    for line in [
        "Rules encoded (need_v12): Downtown/Studio floor 2 from kitchen-in (6:15 / 7:15) to close+30; Studio Mon–Thu 3 from 8:00 to 12:30 (Mon 3 from open), slow weeks 2 core + admin on Colin's/Claire's days; 1899 lead all day + second person 8:30–2:30 in mid/busy, capacity-only in slow (solo allowed); admin window 9:00–2:00 (Studio to close) in slow weeks counts the admin as a body.",
        "Optimiser (ldd_15): 6-hour minimum on every shift; Production Line 6–8 h from kitchen-in; Shift Lead 8–9 h (10 h allowed where the day is 9+ hours so two closers can be met with three people); Closer 6–8 h ending close+30; up to three 6–8 h baristas; 1899 lead 6:30–3:30 plus 0–2 Production Barista shifts of 6–8 h; planning rates $22.25 / $19.75 / $20.25 / $20.00; +$12 per shift tie-break.",
        "Projections check: Projections v1.8 daily 2025 sales by 2026 date × run-rate factor (Downtown 1.00, Studio 0.86 from Jul–Aug 2026 YoY); weekly tiers scaled to full weeks; 1899 not re-based (no 2025 comparable).",
        "Incentive bench test (ldd_16): backtest on monthly_wage_basis Apr–Aug 2026 (core wages, sales, tier of month, 1899 NI from the model); 2027 simulation 2,000 runs with sales ±6% and execution noise; NMWA template 24 h/wk at $21.25; margins reuse the v1.1 cafe economics.",
        "Not re-pulled: timecards (v1.0 harvest, 856 cards, 14 weeks), transactions, financials. Not modified: Consolidated Model, monthly report. Reference JSON: _pipeline/monthly/data/labor_deep_dive_2026-09.json (v1.2; v1.0 and v1.1 copies alongside).",
        "Open questions for Ryan: (1) is the 10-hour open-to-close lead acceptable at Downtown every day (the alternative is a fourth 6-hour shift, +$90/day)? (2) NMWA 5-hour shifts vs the 6-hour minimum; (3) solo mid-tier Fridays at 1899; (4) Claire's weekend admin day — Saturday assumed.",
    ]:
        add_bullet(doc, line)


if __name__ == "__main__":
    build()
