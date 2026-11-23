#!/usr/bin/env python3
"""Build 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.1.docx'.
Revision of the v1.0 document (same day): optimised templates (need_v11 / templates_v11), cafe-specific capacity,
rate-lever exposure, variant pricing, and the rebuilt incentive section (incentive_v11). Uses report_style.py.
"""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, add_callout, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, p, h, sd, sm       # helpers (table = investor table with fixed layout)

DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.1.docx")
J = lambda n: json.load(open(os.path.join(DD, n)))
S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
T = J("templates_v11.json"); NV = J("need_v11.json"); TIERS = J("tiers.json"); CURVES = J("curves.json")
STAFF = J("staffing.json"); CAP = J("capacity.json"); INC = J("incentive_v11.json"); RATE = J("rate_lever.json"); T0 = J("templates.json")
AUG = S["august_2026"]; BR = S["pl_bridge_validation"]
CAFES = ["Downtown", "Studio", "1899 Penn"]
DOW5 = ["Mon", "Tue", "Wed", "Thu", "Fri"]; DOW7 = DOW5 + ["Sat", "Sun"]
DOWS = {"Downtown": DOW5, "Studio": DOW7, "1899 Penn": DOW5}
MAIN = "no10h"
WPM = 52 / 12
TGT = S["targets"]

def W(cafe, tier, variant=MAIN):
    key = tier if variant == "target" else f"{tier}_{variant}"
    return T["cafes"][cafe][key]["week"]
def D(cafe, tier, variant=MAIN):
    key = tier if variant == "target" else f"{tier}_{variant}"
    return T["cafes"][cafe][key]["days"]
def need_week(cafe, tier, k):
    return sum(r[k] for r in NV["variants"]["target"][cafe]["tiers"][tier].values())
def mo(x): return x * WPM

def build():
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.85)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Labor Deep Dive — Schedule Templates by Demand Tier", 22, True, COLOR_ACCENT, 0),
                                          ("September 2026  |  v1.1", 14, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "Prepared September 4, 2026 for Ryan Fleming. v1.1 supersedes the v1.0 issued earlier the same day: the templates are now minimum-cost shift sets solved against the "
                  "demand curves (rather than hand-built), capacity is set per cafe, every template is priced at standard rates and at the rates of the people actually on the roster, "
                  "and the incentive section has been rebuilt with nine structures, an annual simulation, margins by tier and comparable plans. Sources: Square payment-level transactions "
                  "Jan–Aug 2026 (walk-in only), 856 Square cafe timecards across 14 sampled weeks (every week of August), Consolidated Model v17, QBO P&L by location (Coefficient, Sep 3), "
                  "Projections v1.8, and the Square Teams first-cut template for the week of 10/11–10/17. Conventions match the monthly report: core sales = ex-catering walk-in sales; "
                  "core labor = Square timecards excluding 'Manager' (admin) and 'Training'; 1899 carries the same targets and framing as every cafe; NMWA is kept separate.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ================= EXECUTIVE SUMMARY =================
    add_heading(doc, "Executive Summary", level=1)
    dt, st, pn = AUG["Downtown"], AUG["Studio"], AUG["1899 Penn"]
    add_callout(doc, "Ryan's question: Downtown and 1899 ran high-$40s sales per labor hour Mon–Thu and only Friday looked bad — so why did both lose money while Studio, with two soft days, was very profitable?",
                f"Because high-$40s is not a healthy number at Slipstream's pay — it is a 40–47% wage bill. At guaranteed rates of ${dt['gross_rate_core']:.2f}–${pn['gross_rate_core']:.2f} gross per core hour, "
                f"a cafe needs ${pn['splh_needed_for_target']:.0f}–${st['splh_needed_for_target']:.0f} of walk-in sales per core labor hour to land on its target; August delivered ${dt['splh_core_walkin']:.0f} at Downtown, "
                f"${st['splh_core_walkin']:.0f} at Studio and ${pn['splh_core_walkin']:.0f} at 1899. On the wage basis the targets are written on, core labor ran {p(dt['core_wage_pct_of_core_sales'])} of core sales at Downtown "
                f"(target 32), {p(st['core_wage_pct_of_core_sales'])} at Studio (target 30) and {p(pn['core_wage_pct_of_core_sales'])} at 1899 (target 35) — "
                f"{m(dt['headroom_core_wages_vs_target'] + st['headroom_core_wages_vs_target'] + pn['headroom_core_wages_vs_target'])} of core wages above target in one month. "
                "Studio was 7 points over and still made money because its fixed costs sit on a bigger base and its COGS are 8–9 points lower; Downtown and 1899 were 12–14 points over AND carried "
                "37–38% COGS, so the same miss turned negative. The P&L understates the miss: credit-card tips (12% of sales) fund part of the guaranteed wage, so booked labor ≈ 1.12 × gross wages − tips. "
                "Sales per labor hour and the tip offset are the same lever — a cafe that sells more per hour also collects more tips per hour — which is most of why Studio's 'loaded cost per hour' "
                "looked like $18 against $22 at Downtown. Gross rates at the three cafes are within $0.70 of each other. The fix is hours placement first (3–6× the size of the rate lever), "
                "then who fills which shift.", body_size=10, title_size=10)

    add_para(doc, "What is new in v1.1 (and why the numbers moved):", after=2, before=6)
    for line in S["changes_from_v1_0"]:
        add_bullet(doc, line)

    over = {c: {t: (need_week(c, t, "over_hours"), need_week(c, t, "over_dollars")) for t in ["slow", "mid", "busy"]} for c in CAFES}
    add_para(doc, "Two levers, sized (need = the cafe's own busy-hour throughput applied to its demand curve, with the floors in Section 9):", after=4, before=6)
    add_bullet(doc, f"Hours placement. In the sampled weeks Downtown paid for {h(over['Downtown']['slow'][0])} person-hours a week above need in slow weeks ({m(over['Downtown']['slow'][1])}/wk), "
                    f"Studio {h(over['Studio']['slow'][0])} ({m(over['Studio']['slow'][1])}/wk) and 1899 {h(over['1899 Penn']['slow'][0])}–{h(over['1899 Penn']['busy'][0])} ({m(over['1899 Penn']['slow'][1])}–{m(over['1899 Penn']['busy'][1])}/wk). "
                    f"Downtown's core hours barely move with demand ({h(W('Downtown','slow')['actual_sampled_hours'],0)} / {h(W('Downtown','mid')['actual_sampled_hours'],0)} / {h(W('Downtown','busy')['actual_sampled_hours'],0)} h/wk in slow / mid / busy weeks), "
                    "so the ratio is set entirely by sales. The pattern is a 4th person from 8:30a to 3:30p and two Shift Leads overlapping for six hours, on days that need 3 people for 90 minutes and 2 the rest of the time.", bold_lead="1.")
    add_bullet(doc, f"Cost per hour. Senior people on barista shifts and overlapping leads cost ${RATE['Downtown']['premium_per_hour']:.2f}/core hour at Downtown (${RATE['Downtown']['premium_per_week']:,.0f}/wk), "
                    f"${RATE['Studio']['premium_per_hour']:.2f} at Studio (${RATE['Studio']['premium_per_week']:,.0f}/wk) and ${RATE['1899 Penn']['premium_per_hour']:.2f} at 1899 (${RATE['1899 Penn']['premium_per_week']:,.0f}/wk) against a standard mix of "
                    "$21 lead / $19 barista — about $1,600 a month company-wide. It also shows up inside the templates: the same shift list costs 1.2–1.4 points more at the roster's current average rates than at standard rates "
                    "(Section 7), which is why Downtown's slow template lands at 32.0% at standard rates and 33.3% with the people who worked August.", bold_lead="2.")

    rows = []
    for c in CAFES:
        a = AUG[c]
        rows.append([c, m(a["core_sales_model"]), h(a["core_hours"], 0), m(a["splh_core_walkin"], 0), f"${a['gross_rate_core']:.2f}", p(a["core_wage_pct_of_core_sales"]), p(a["target_pct"], 0),
                     m(a["headroom_core_wages_vs_target"]), m(a["splh_needed_for_target"], 0), p(a["pl_labor_pct_booked"]), p(a["pl_labor_pct_bridge"])])
    rows.append(["Open Cafes", m(sum(AUG[c]["core_sales_model"] for c in CAFES)), h(sum(AUG[c]["core_hours"] for c in CAFES), 0), "", "", "", "", m(sum(AUG[c]["headroom_core_wages_vs_target"] for c in CAFES)), "", "", ""])
    add_heading(doc, "August 2026 on the wage basis — the basis the targets are written on", level=3)
    table(doc, ["Cafe", "Core sales", "Core hrs", "SPLH", "Gross $/hr", "Core wage %", "Target", "Over target $", "SPLH needed", "P&L labor % booked", "P&L % bridge"], rows,
          widths=[1.0, 0.7, 0.5, 0.45, 0.6, 0.6, 0.45, 0.7, 0.55, 0.6, 0.55], total_rows=[3], size=8)
    add_caveat(doc, "SPLH = walk-in net sales ÷ core labor hours (Square timecards, Aug 1–31, missed clock-outs capped). Core wage % = gross guaranteed wages ÷ model core sales. 'P&L % bridge' = (1.12 × gross wages incl. admin & training − CC tips) ÷ core sales; "
                    "the booked figure differs month to month because payroll posts on pay dates. Over Apr–Aug the bridge ties to QBO within 0.4% (Downtown), 0.6% (Studio) and 2.3% (1899) — Section 1.")

    add_heading(doc, "What the templates deliver — weekly core wages by tier", level=3)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier); tg = TGT[c]["target"] * 100
            rows.append([f"{c} — {tier}", m(w["sales_typical"]), f"{h(w['actual_sampled_hours'], 0)} h / {p(w['actual_sampled_labor_pct'])}", f"{h(w['ryan_first_cut_hours'], 0)} h / {p(w['ryan_first_cut_labor_pct'])}",
                         f"{h(w['v10_hours'], 0)} h / {p(w['v10_labor_pct'])}", h(w["hours"], 1), m(w["wages"]), p(w["labor_pct"]), p(w["labor_pct_obs_rate"]), p(tg, 0),
                         sm(w["wages"] - w["actual_sampled_wages"]), sm(w["wages"] - w["ryan_first_cut_wages"])])
    table(doc, ["Cafe / tier", "Typical wk sales", "Actual (h / %)", "First cut (h / %)", "v1.0 (h / %)", "v1.1 hours", "v1.1 wages", "v1.1 % std", "v1.1 % roster", "Target", "Δ$/wk vs actual", "Δ$/wk vs 1st cut"],
          rows, widths=[1.0, 0.62, 0.78, 0.78, 0.72, 0.5, 0.55, 0.5, 0.55, 0.45, 0.6, 0.6], key_rows=[0, 3, 6], size=7)
    add_caveat(doc, "Weekly core wages at standard rates ($21 Shift Lead, $19 barista/production, $19.50 closer); 'roster' re-prices the same hours at the cafe's current average core rate ($20.73 / $20.20 / $20.84). "
                    "'Actual' = the sampled weeks of that tier at the rates paid; 'first cut' = Ryan's 10/11 week at the rates of the people scheduled (one schedule, compared to each tier's sales). "
                    "Downtown and Studio land at or under target in every tier at standard rates; 1899 reaches its 35% only at busy-tier sales — Section 5 explains why and what the options are.")

    dts, dtm = W("Downtown", "slow"), W("Downtown", "mid"); sts, stm = W("Studio", "slow"), W("Studio", "mid"); pnm, pnb = W("1899 Penn", "mid"), W("1899 Penn", "busy")
    add_para(doc, "The three biggest schedule changes, by monthly dollars (slow-tier weeks are what September–October look like on the current run-rate; mid-tier in brackets):", after=4, before=6)
    add_bullet(doc, f"Downtown — one lead, one barista-rate Closer, the Production Line goes home when prep is done. Shift Lead 6:30–2:30/3:30 with a Closer 8:30–4:30 (barista rate) instead of a second lead; "
                    f"Production Line 6:00–12:00/1:30 instead of 6:15–3:00; rush baristas (4–6 h) only in mid/busy weeks; Monday and Friday are 2- to 3-person days with the afternoon single-covered. "
                    f"Slow week {h(dts['hours'],0)} h vs {h(dts['actual_sampled_hours'],0)} actual: −{m(dts['actual_sampled_wages']-dts['wages'])}/wk ≈ −{m(mo(dts['actual_sampled_wages']-dts['wages']))}/mo "
                    f"[mid {h(dtm['hours'],0)} h, −{m(mo(dtm['actual_sampled_wages']-dtm['wages']))}/mo]; vs the first cut −{dts['ryan_first_cut_hours']-dts['hours']:.0f} h/wk.", bold_lead="1.")
    add_bullet(doc, f"Studio — Tuesday is a 2-person day, Wed/Thu carry one 4–5.5 h rush barista, and the weekend runs four people at the 10a–1p peak instead of five. Production Line 6:00–12:00/1:00 (7:00–1:00 weekends); "
                    f"Shift Lead 6:30–2:45 (7:30–4:30 weekends) with a Closer only Fri–Sun. Slow week {h(sts['hours'],0)} h vs {h(sts['actual_sampled_hours'],0)} actual: −{m(mo(sts['actual_sampled_wages']-sts['wages']))}/mo "
                    f"[mid {h(stm['hours'],0)} h, −{m(mo(stm['actual_sampled_wages']-stm['wages']))}/mo]; vs the first cut −{sts['ryan_first_cut_hours']-sts['hours']:.0f} h/wk. The 'under-staffed weekend' in the monthly report was a share-of-hours artifact — by throughput the weekend was over-covered.", bold_lead="2.")
    add_bullet(doc, f"1899 — the second person works 8:30–2:30 (both peaks) instead of 6:30–3:30, and 4 hours on Monday/Friday. Mid week {h(pnm['hours'],0)} h ({p(pnm['labor_pct'])}) vs {h(pnm['actual_sampled_hours'],0)} actual: −{m(mo(pnm['actual_sampled_wages']-pnm['wages']))}/mo "
                    f"[busy {h(pnb['hours'],0)} h at {p(pnb['labor_pct'])}, −{m(mo(pnb['actual_sampled_wages']-pnb['wages']))}/mo]. Below ~$4,300/week 1899 cannot reach 35% with a lead on the floor all day and a second person for the peaks — "
                    "that is a sales and COGS problem (38% COGS on a 92% coffee mix), not a schedule fix.", bold_lead="3.")
    add_bullet(doc, "Also: stop filling barista shifts with $24–25 team members and stop overlapping two Shift Leads (≈ $1,600/mo company-wide); the 10-hour open-to-close lead shift Ryan flagged is worth only ~$35/wk at Downtown and nothing at Studio — "
                    "not worth the conversation; add the second-Wednesday community-day shift at NMWA (missing from the 10/11 first cut).", bold_lead="4.")
    add_para(doc, "Tiers, in one line each (weekly walk-in sales; Section 2): Downtown slow < $7,500, mid $7,500–9,000, busy ≥ $9,000. Studio slow < $10,500, mid $10,500–12,500, busy ≥ $12,500. 1899 slow < $3,300, mid $3,300–4,200, busy ≥ $4,200. "
                  "The bands sit within $100–500 of the statistical breaks (k-means / Jenks) in the 2026 weekly data. Ryan's hypothesis holds for Studio (Jan, Feb, Jul, Aug slow; Apr–Jun busy) and for Downtown's August; Downtown's February was mid, "
                  "and 1899's August was a mid month by sales. Projections put October in the busy tier for Downtown and Studio, but Jul–Aug ran 4% (Downtown) and 14% (Studio) below 2025 — build the 10/11 week on the MID templates "
                  "and step up only when the trailing two weeks clear the band floor.", before=6, after=6)
    a = INC["annual_simulation"]["plans"]["A"]; sv = INC["annual_simulation"]["labor_saving_vs_aprAug_run_rate"]
    add_para(doc, f"Incentive (Section 8): keep the 'Beat Share' — 50% of core wages saved below the target line, capped at 1.5% of core sales, pooled across the open cafes, allocated as a percentage of each person's core wages "
                  f"(which keeps the FLSA overtime math clean), paid only when the schedules honoured the template floors. Simulated over a 2027 tier calendar it pays {m(a['p10'])}–{m(a['p90'])} a year ({m(a['p50'])} median, "
                  f"{a['share_of_saving_p50']:.0f}% of the ~{m(sv['p50'])} of core wages the templates save against the Apr–Aug run-rate), against $10–13K for a dollar-an-hour or per-shift plan that pays for merely hitting target. "
                  f"Consolidated net income at target on the same calendar is ≈ {m(INC['profit_share_sizing']['consolidated_ni_year_at_target'])} a year, so the 5–10% annual retirement share is "
                  f"{m(INC['profit_share_sizing']['share_5pct_of_consolidated'])}–{m(INC['profit_share_sizing']['share_10pct_of_consolidated'])} — the monthly piece has to stay small for that to be affordable, and the cap does that.", after=8)

    # ================= 1. BRIDGE =================
    add_heading(doc, "1. How labor reaches the P&L — the bridge between the schedule and the statement", level=1)
    add_para(doc, "Slipstream pays guaranteed rates; credit-card tips are applied toward that rate and the company tops up the difference (ADP: $12 tipped base + CREDTIPP + OTH top-up). QBO's FOH Labor line is therefore gross wages NET of tips, "
                  "and Payroll Tax is allocated at 12% of gross. That gives a simple, testable bridge: P&L labor ≈ 1.12 × gross wages (core + admin + training) − credit-card tips. The table tests it against QBO for Apr–Aug.", after=6)
    rows = []
    for c in CAFES:
        b = BR[c]
        for mon in b["months"]:
            r = b["by_month"][mon]; rows.append([f"{c} {mon}", p(r["core_wage_pct"]), m(r["pl_est"]), m(r["pl_actual"]), sm(r["pl_est"] - r["pl_actual"]), m(r["headroom"])])
        rows.append([f"{c} Apr–Aug", "", m(b["pl_est_sum"]), m(b["pl_actual_sum"]), f"{b['error_pct']:+.1f}%", m(sum(x["headroom"] for x in b["by_month"].values()))])
    table(doc, ["Cafe / month", "Core wage % (gross)", "Bridge estimate", "QBO booked", "Δ", "Core wages over target"], rows, widths=[1.3, 1.0, 1.0, 1.0, 0.8, 1.2], total_rows=[5, 11, 17], size=9)
    add_caveat(doc, "Months are lumpy because payroll posts on pay dates with a single-line accrual; the five-month totals tie. Core wage % for Apr and Aug is scaled from partial timecard coverage to the full month. 'Over target' = gross core wages − target % × model core sales.")
    add_heading(doc, "What the bridge means for the targets", level=3)
    add_bullet(doc, "A cafe exactly on its core target (30/32/35 wage basis) with ~2–4% admin and training shows about 25–27% total labor on the P&L, not 33–36%: 1.12 × (30 + 4) − 12 = 26 at Studio; 1.12 × (32 + 2) − 12 = 26 at Downtown; 1.12 × (35 + 3) − 10.4 = 32 at 1899 (lower tip rate). "
                    "With COGS at 28–30%, prime cost lands in the 55–60% healthy band at Studio and Downtown — that is the prize. Section 8e prices it by tier.")
    add_bullet(doc, f"Loaded cost per hour reconciles the same way. August booked ${dt['loaded_rate_booked_per_hour']:.2f}/hr at Downtown, ${st['loaded_rate_booked_per_hour']:.2f} at Studio and ${pn['loaded_rate_booked_per_hour']:.2f} at 1899; the bridge gives "
                    f"${dt['loaded_rate_bridge_per_hour']:.2f}, ${st['loaded_rate_bridge_per_hour']:.2f} and ${pn['loaded_rate_bridge_per_hour']:.2f}. Gross rates are ${dt['gross_rate_core']:.2f} / ${st['gross_rate_core']:.2f} / ${pn['gross_rate_core']:.2f}; tips per core hour "
                    f"${dt['tips_per_core_hour']:.2f} / ${st['tips_per_core_hour']:.2f} / ${pn['tips_per_core_hour']:.2f}. Studio's lower loaded rate is its higher tips per hour plus accrual timing; it is not lower pay.")
    add_bullet(doc, "Implication for the weekly check: keep scoring schedules on gross wages ÷ projected core sales (the 30/32/35 basis) and never compare that number to the monthly P&L labor % — the P&L figure is ~7–9 points lower at the same staffing. The monthly report should show both, with the bridge.")
    add_bullet(doc, "DC context: the District's minimum wage rose to $18.40 on July 1, 2026 and the tipped minimum to 56% of it ($10.30). Slipstream's guaranteed rates already clear the full minimum, but four sampled timecards carried an $18.00 rate — worth a payroll check.")

    # ================= 2. TIERS =================
    add_heading(doc, "2. Demand tiers", level=1)
    add_para(doc, "Tiers are defined on weekly walk-in sales per cafe (Mon–Sun; holiday-shortened weeks scaled to a full week; one weather week — Downtown Jan 26, $2,365 — excluded). Each week of 2026 is assigned to a band by what it actually sold, so every tier's demand curve "
                  "is built from days that really were that busy. The template for a week is chosen by the projected weekly sales in the weekly report's forward block, not by the calendar.", after=6)
    rows = []
    for c in CAFES:
        b = S["tiers"][c]["bands_weekly_walkin"]; v = S["tier_band_validation"][c]
        rows.append([c, f"< {m(b['slow'][1])}", f"{m(b['mid'][0])} – {m(b['mid'][1])}", f"≥ {m(b['busy'][0])}", f"{m(v['kmeans'][0])} / {m(v['kmeans'][1])}", f"{m(v['jenks'][0])} / {m(v['jenks'][1])}"])
    table(doc, ["Cafe", "Slow", "Mid", "Busy", "k-means breaks", "Jenks breaks"], rows, widths=[1.1, 1.0, 1.3, 1.0, 1.2, 1.2], size=9)
    add_caveat(doc, "New in v1.1: the bands were tested against two one-dimensional clusterings of the 33–34 weekly observations per cafe. Downtown and Studio bands sit within $100–500 of both; 1899's statistical breaks are higher ($3,500–3,600 / $4,400–4,700) "
                    "because the cafe is still growing ~90% year-over-year and the early-2026 weeks are ramp, not season. The 1899 bands are kept at $3,300 / $4,200 because the Sep–Dec projections ($3,830–4,190/wk) fall in the mid band either way.")
    add_heading(doc, "Month membership — 2026 actuals (Jan–Aug), projection (Sep–Dec)", level=3)
    MONS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    rows = []
    for c in CAFES:
        cal = S["tiers"][c]["calendar"]; row = [c]
        for mon in MONS:
            r = cal.get(mon)
            if not r: row.append("—"); continue
            t = r["tier"]
            if "weeks_by_tier" in r:
                w = r["weeks_by_tier"]; row.append(f"{t} ({w['slow']}/{w['mid']}/{w['busy']})")
            else:
                extra = f"; run-rate {r['yoy_view_tier']}" if "yoy_view_tier" in r else ""
                if mon == "Dec": extra += "; Ryan: slow"
                row.append(f"{t}{extra}")
        rows.append(row); rows.append(["  avg weekly $"] + [m(cal[mon]["avg_weekly"]) if mon in cal else "—" for mon in MONS])
    table(doc, ["Cafe"] + MONS, rows, widths=[0.9] + [0.5] * 12, key_rows=[0, 2, 4], size=7)
    add_caveat(doc, "Jan–Aug cells show the dominant tier and the count of weeks by slow/mid/busy. Sep–Dec use Projections v1.8; 'run-rate' applies the Jul–Aug 2026 year-over-year change (Downtown −4%, Studio −14%) to the 2025 actuals — 1899 has no comparable 2025 base. "
                    "December is projected mid at Downtown and 1899, but Ryan's instruction and the 2025 pattern say plan the holiday weeks (Dec 21–Jan 3) on the slow template.")
    add_heading(doc, "Typical day by tier (walk-in $ / transactions per open day)", level=3)
    for c in CAFES:
        dows = DOWS[c]; rows = []
        for tier in ["slow", "mid", "busy"]:
            td = S["tiers"][c]["typical_day"][tier]
            rows.append([f"{c} — {tier}"] + [f"{m(td[d]['net'])} / {td[d]['txns']:.0f}" if d in td else "—" for d in dows] + [m(sum(td[d]["net"] for d in dows if d in td))])
        table(doc, ["Cafe / tier"] + dows + ["Week"], rows, widths=[1.2] + [0.72] * len(dows) + [0.7], size=8)
    add_caveat(doc, "Ryan's hypothesis tested: Studio — Jan, Feb, Jul and Aug are slow, Apr–Jun busy: confirmed, with July added. Downtown — August is slow but February 2026 was mid and January itself was mixed; 2025 shows the same Aug/Dec/Jan trough. "
                    "1899 — August was a mid month ($3,763/wk vs $4,129 in July); its Jan–Feb 'slow' weeks were the ramp of a growing cafe, so the slow tier may not recur at that level outside the holiday weeks.")

    cafe_sections(doc)
    nmwa_section(doc)
    sensitivity_section(doc)
    incentive_section(doc)
    data_notes(doc)
    appendix(doc)
    _ensure_zoom_percent(doc)
    doc.save(OUT); print("Saved →", OUT)


def cafe_sections(doc):
    secno = {"Downtown": 3, "Studio": 4, "1899 Penn": 5}
    HOURS_SHOW = {"Downtown": list(range(7, 16)), "Studio": list(range(7, 17)), "1899 Penn": list(range(7, 15))}
    intro = {"Downtown": "An office-district cafe (7a–4p Mon–Fri): the 8:30–10:00 rush is a third of the day, Monday runs 60–70% of a mid-week day and Friday 40–50%. Two Shift Leads overlapped for six hours on most sampled days.",
             "Studio": "A neighbourhood cafe (7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun) with a real food program: the weekend is 40–45% of the week, weekday peaks are 10a–12p rather than the rush, and Tuesday–Thursday are $900–1,350 days that were staffed with three people from 6:00 to 2:30. "
                       "Weekend tickets are bigger ($10.70 vs $9.20) and slower — the busy-hour throughput the team sustains is 12 per person-hour against 14 at Downtown, which is why v1.1 gives Studio its own capacity.",
             "1899 Penn": "A two-person office-building cafe (7a–3p Mon–Fri) with a bimodal day — a 9am rush and a 1pm lunch peak of equal size — and a Friday that is half a Wednesday. Both people worked 6:30–3:30 every sampled day."}
    shape = {"Downtown": "What the shape says: Downtown needs 3 people from 8:30 to 10:00 and 2 for the rest of the morning in slow and mid weeks (4 at the 9am peak in busy weeks), and after 1:00p the curve is 8–28 transactions an hour — 1 to 2 people. "
                         "It ran 4 people from 9:00 to 3:30 in every tier; the 4th person costs ~$130/day and adds no measurable throughput. Monday and Friday are 2- to 3-person days at any tier. The last hour (3–4p) is 3–12 transactions; Ryan's caution on judging it alone stands — the templates keep the 4p close and let the Closer run it alone.",
             "Studio": "What the shape says: Tuesday–Thursday need 2 people plus a third for the 8:30–1:00 stretch on Wed/Thu in mid and busy weeks; the cafe ran 3 all day. Friday and the weekend need 3–4 at the 10a–1p peak (5 on a busy Saturday) and were staffed at 3.8–5.0. "
                       "Monday's third person at open is the catering delivery; where the ALL's Admin block is on the floor at 7:00 (as in the first cut) it counts toward the three.",
             "1899 Penn": "What the shape says: two people 6:30–3:30 is 18 hours a day; the day needs a lead 6:30–3:30 and a second person from the rush through lunch (8:30–2:30) — 15 hours — and only 4 hours of the second person on Monday and Friday. "
                          "At $750–950 a day (mid tier) the 35% target buys 13–16 hours, so the lean shape gets close (39%) and the busy tier lands on target (34%). 1899's lever is smaller in dollars than Downtown's or Studio's, but it is the difference between the incentive fee and a loss for the owner."}
    for c in CAFES:
        n = secno[c]
        add_heading(doc, f"{n}. {c}", level=1)
        add_para(doc, intro[c], after=6)
        # a. curve
        add_heading(doc, f"{n}a. Demand curve — walk-in transactions per hour, by tier", level=2)
        hrs = HOURS_SHOW[c]; show_days = ["Tue", "Fri"] if c != "Studio" else ["Tue", "Fri", "Sat"]; rows = []
        for d in show_days:
            for tier in ["slow", "mid", "busy"]:
                cv = CURVES[c]["by_tier"][tier][d]; g = lambda hh: cv['hourly'].get(str(hh), {}).get('txns', 0)
                rows.append([f"{d} — {tier}"] + [(f"{g(hh):.0f}" if str(hh) in cv['hourly'] else "·") for hh in hrs] + [f"{sum(g(hh) for hh in hrs):.0f}"])
        table(doc, ["Day / tier"] + [f"{hh if hh <= 12 else hh - 12}{'a' if hh < 12 else 'p'}" for hh in hrs] + ["Day"], rows, widths=[1.1] + [0.42] * len(hrs) + [0.5], key_rows=[0, 3, 6], size=8)
        add_caveat(doc, "Average transactions in the hour beginning at the time shown, per open day of that tier and weekday. Full 30-minute curves for every day are in the workbook ('Need vs Actual' sheet).")
        # b. need vs actual (v1.1 rules)
        add_heading(doc, f"{n}b. Actual staffing vs need — hours and dollars per week (v1.1 need rules)", level=2)
        rows = []
        nvc = NV["variants"]["target"][c]
        for tier in ["slow", "mid", "busy"]:
            dows = nvc["tiers"][tier]
            for d in DOWS[c]:
                r = dows.get(d)
                if not r: continue
                rows.append([f"{tier} {d}", h(r["need_hours"]), h(r["actual_hours"]), h(r["over_hours"]), m(r["over_dollars"]), h(r["under_hours"]), f"{r['peak_need']} / {r['peak_actual']:.1f}"])
            rows.append([f"{tier} — week", h(sum(x["need_hours"] for x in dows.values())), h(sum(x["actual_hours"] for x in dows.values())), h(sum(x["over_hours"] for x in dows.values())),
                         m(sum(x["over_dollars"] for x in dows.values())), h(sum(x["under_hours"] for x in dows.values())), ""])
        nd = len(DOWS[c])
        table(doc, ["Tier / day", "Need (h)", "Actual (h)", "Over (h)", "Over ($/day)", "Under (h)", "Peak need / actual"], rows, widths=[1.2, 0.8, 0.85, 0.8, 0.9, 0.8, 1.2], total_rows=[nd, 2 * nd + 1, 3 * nd + 2], size=9)
        add_caveat(doc, f"Need = the throughput rule in Section 9 (capacity {nvc['cap']:.0f} transactions per person-hour at this cafe) applied to the tier's 30-minute curve; actual = average core headcount on the clock in the sampled weeks. "
                        f"Dollars at the cafe's observed average core rate (${nvc['avg_core_rate']:.2f}). 'Over' is paid time above need; 'under' is need above staffing (mostly prep half-hours and single-person closes).")
        # c. levers
        add_heading(doc, f"{n}c. The two levers, quantified", level=2)
        rl = RATE[c]; hb = INC["headroom_by_tier"][c]
        tot_over = {t: sum(x["over_dollars"] for x in nvc["tiers"][t].values()) for t in ["slow", "mid", "busy"]}
        add_bullet(doc, f"Hours placement: {m(tot_over['slow'])} / {m(tot_over['mid'])} / {m(tot_over['busy'])} per week of paid time above need in slow / mid / busy weeks (≈ {m(mo(tot_over['slow']))} / {m(mo(tot_over['mid']))} / {m(mo(tot_over['busy']))} a month). "
                        f"Against the target line the monthly headroom in the sampled weeks is {m(hb['slow']['headroom_actual_vs_target'])} / {m(hb['mid']['headroom_actual_vs_target'])} / {m(hb['busy']['headroom_actual_vs_target'])}; the v1.1 templates recover "
                        f"{m(hb['slow']['template_saving_vs_actual'])} / {m(hb['mid']['template_saving_vs_actual'])} / {m(hb['busy']['template_saving_vs_actual'])} of it a month at standard rates.", bold_lead="Lever 1 —")
        add_bullet(doc, f"Cost per hour: observed average core rate ${rl['avg_rate']:.2f} vs ${rl['std_mix_rate']:.2f} at standard rates for the same role mix — a premium of ${rl['premium_per_hour']:.2f}/hr, {m(rl['premium_per_week'])}/wk, ≈ {m(mo(rl['premium_per_week']))}/mo. "
                        f"Manager-rate ($24–25) team members worked {rl['manager_rate_hours_per_week']:.1f} core hours a week here ({rl['manager_rate_share']:.0f}% of core hours); Shift Lead hours are {rl['sl_share_hours']:.0f}% of core and two leads were on the floor together {rl['double_SL_hours_per_week']:.1f} hours a week. "
                        f"Hours by rate band: under $19 {rl['rate_bands_share']['0-19']:.0f}%, $19–20.49 {rl['rate_bands_share']['19-20.5']:.0f}%, $20.50–22.49 {rl['rate_bands_share']['20.5-22.5']:.0f}%, $22.50+ {rl['rate_bands_share']['22.5-30']:.0f}%. "
                        f"On the v1.1 templates this is the gap between the 'std' and 'roster' columns: {p(W(c,'mid')['labor_pct_obs_rate'] - W(c,'mid')['labor_pct'])} of sales in a mid week.", bold_lead="Lever 2 —")
        add_para(doc, shape[c], after=6)
        # d. templates
        add_heading(doc, f"{n}d. Schedule templates (v1.1 — minimum-cost shift sets)", level=2)
        tg = TGT[c]
        add_para(doc, f"Targets: {tg['target']*100:.0f}% core wages of core sales (ceiling {tg['ceiling']*100:.0f}%). Standard rates $21 Shift Lead, $19 barista/production, $19.50 closer. Each day's shift list is the cheapest set that covers the tier's half-hour need curve "
                      "under the shape rules in Section 9 (no shift over 9 h; 4-hour minimum; a keyholder on the floor from 30 minutes before open to close). 'Solo' marks the windows one person runs the floor; 'Admin covers' marks the half-hours where the need exceeds the "
                      "template and the ALL's block is the intended cover.", after=4)
        for tier in ["slow", "mid", "busy"]:
            o = T["cafes"][c][f"{tier}_{MAIN}"]; w = o["week"]; label = tier.capitalize()
            add_heading(doc, f"{label} template — {m(w['sales_typical'])}/wk typical, {w['hours']:.1f} core hours, {m(w['wages'])} = {p(w['labor_pct'])} std / {p(w['labor_pct_obs_rate'])} roster (SPLH {m(w['splh'],0)})", level=3)
            rows = []
            for d, dd in o["days"].items():
                for i, s in enumerate(dd["shifts"]):
                    rows.append([d if i == 0 else "", s["role"], s["start"], s["end"], h(s["hours"], 2), f"${s['rate']:.2f}", m(s["wages"]), ""])
                tmpl = dd["template"]; act = dd.get("actual_sampled", {})
                notes = []
                if tmpl["solo_windows"]: notes.append("solo " + ", ".join(tmpl["solo_windows"]))
                if tmpl["admin_cover_windows"]: notes.append("Admin covers " + ", ".join(tmpl["admin_cover_windows"]))
                rows.append([f"{d} total", f"{len(dd['shifts'])} shifts", "", "", h(tmpl["hours"], 2), "", m(tmpl["wages"]),
                             f"{p(tmpl['labor_pct'])} of {m(dd['sales_typical'])} (roster {p(tmpl['labor_pct_obs_rate'])}); peak {tmpl['peak_staff']}; actual {act.get('core_hours', '—')} h / {p(act.get('labor_pct'))}" + ("; " + "; ".join(notes) if notes else "")])
            total_idx = [i for i, r in enumerate(rows) if str(r[0]).endswith("total")]
            table(doc, ["Day", "Role", "Start", "End", "Hours", "Rate", "Wages", "Check / notes"], rows, widths=[0.6, 1.45, 0.45, 0.45, 0.45, 0.5, 0.55, 2.25], total_rows=total_idx, size=8)
            v = S["templates"][c][tier]["variants_week"]
            add_caveat(doc, f"Week: {w['hours']:.1f} h, {m(w['wages'])} at standard rates = {p(w['labor_pct'])} of {m(w['sales_typical'])} (target {p(tg['target']*100,0)}); at the roster's current rates {m(w['wages_obs_rate'])} = {p(w['labor_pct_obs_rate'])}. "
                            f"At −10% sales: {p(w['labor_pct_minus10'])}. Sampled actual weeks of this tier: {w['actual_sampled_hours']:.1f} h, {m(w['actual_sampled_wages'])} = {p(w['actual_sampled_labor_pct'])}; v1.0 template {w['v10_hours']:.1f} h = {p(w['v10_labor_pct'])}; "
                            f"Ryan's first cut {w['ryan_first_cut_hours']:.1f} h = {p(w['ryan_first_cut_labor_pct'])} at this tier's sales. Variants (weekly wages): 10-hour lead option {m(v['target']['wages'])} ({p(v['target']['labor_pct'])}); comfort {m(v['comfort']['wages'])} ({p(v['comfort']['labor_pct'])}); "
                            f"6-hour minimum shifts {m(v['min6h']['wages'])} ({p(v['min6h']['labor_pct'])}); 8-hour Production Line {m(v['pl8']['wages'])} ({p(v['pl8']['labor_pct'])}). "
                            + (f"Drop to the next-slower template when projected weekly sales fall below {m(w['ceiling_trigger_sales'])} (standard rates) / {m(w['ceiling_trigger_sales_obs_rate'])} (roster rates)." if tier != "slow"
                               else f"This template breaches the {p(tg['ceiling']*100,0)} ceiling when weekly sales fall below {m(w['ceiling_trigger_sales'])} (standard) / {m(w['ceiling_trigger_sales_obs_rate'])} (roster); there is no slower template — below that, the floors and Admin cover are the only lever."))
        # e. changes vs first cut
        add_heading(doc, f"{n}e. Specific changes from the 10/11 first-cut template", level=2)
        for line in FIRST_CUT_CHANGES[c]: add_bullet(doc, line)
        # f. changes vs v1.0
        add_heading(doc, f"{n}f. What moved from the v1.0 templates", level=2)
        rows = []
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier)
            rows.append([tier, h(w["v10_hours"], 1), p(w["v10_labor_pct"]), h(w["hours"], 1), m(w["wages"]), p(w["labor_pct"]), p(w["labor_pct_obs_rate"]), sd(w["hours"] - w["v10_hours"]), sm(w["wages"] - w["v10_wages"])])
        table(doc, ["Tier", "v1.0 hrs", "v1.0 %", "v1.1 hrs", "v1.1 wages", "v1.1 % std", "v1.1 % roster", "Δ hrs/wk", "Δ $/wk"], rows, widths=[0.7, 0.7, 0.7, 0.7, 0.8, 0.75, 0.8, 0.7, 0.75], size=9)
        for line in V10_CHANGES[c]: add_bullet(doc, line)


FIRST_CUT_CHANGES = {
    "Downtown": [
        "Monday (first cut: Production Barista 6:30–1, Production Line 6:12–4:30, Shift Lead 9–4:30 = 24.3 h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–2:30, Closer 9:00–4:30 = 21.5 h (−2.8) — the lead opens instead of arriving at 9, the Production Line stops at noon, a barista-rate Closer takes the afternoon. "
        "Slow: Production Line 6:00–12:30, Shift Lead 6:30–2:30, Closer 12:30–4:30 = 18.5 h (−5.8), single-covered after 2:30.",
        "Tuesday (first cut: 4 people — Production Line 6:15–2, Production Barista 7:30–4:30, Shift Lead 6:30–3, Production Barista 9–4:30 = 32.75 h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–3:30, Barista 8:30–2:30, Closer 8:30–4:30 = 29 h (−3.75) — the 9-hour Production Barista becomes a 6-hour rush-and-lunch barista and the second 9–4:30 becomes the Closer. "
        "Slow: Production Line 6:00–1:30, Shift Lead 6:30–2:30, Closer 8:30–4:30 = 23.5 h (−9.25): three people, no barista.",
        "Wednesday (first cut: Shift Lead 8–4:30, Production Barista 6:30–1:30, Production Line 6:15–3, Barista 9–4:30 = 31.75 h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–3:30, Barista 8:00–12:00, Barista 8:30–2:30, Closer 12:30–4:30 = 29 h (−2.75) — the 8–4:30 second lead is replaced by a 4-hour rush barista plus a 4-hour Closer at barista rate. "
        "Slow: Production Line 6:00–12:00, Shift Lead 6:30–3:00, Closer 8:30–4:30 = 22.5 h (−9.25).",
        "Thursday (first cut as Wednesday, 31.75 h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–3:30, Barista 7:30–11:30, Closer 8:30–4:30 = 27 h (−4.75). Slow: 22.5 h (−9.25), same shape as Wednesday.",
        "Friday (first cut: Production Barista 9:30–4:30, Shift Lead 6:30–1, Production Line 6:15–4:30 = 23.75 h). Friday is 40–50% of a Wednesday: Production Line 6:00–12:30, Shift Lead 6:30–2:30, Closer 12:30–4:30 = 18.5 h (−5.25) in every tier below busy; the afternoon is one person at 2–9 transactions an hour. "
        "In the first cut the Production Line closes a 10.25-hour Friday — in the template the Closer does.",
        "Rate: the first cut has $22.75–25.25 team members on ~40 hours of barista/production shifts (Tue/Wed/Thu 9–4:30; Mon/Tue Production Line). At $19 those hours save ~$180/wk; if they must be scheduled, put them in the Shift Lead slot.",
        "Net: first cut 144.3 h / $3,058 at the people's rates (37.0% of a $8,266 mid week; 46.6% of a $6,566 slow week) → mid template 125 h / $2,477 (30.0% std, 31.3% roster) → slow template 105.5 h / $2,102 (32.0% / 33.3%). Columbus Day (10/12): Downtown closed, unschedule Monday.",
    ],
    "Studio": [
        "Sunday (first cut: Barista 7:30–2, Shift Lead 8–3, Production Line 7:15–3:30, Production Barista 10:30–5:30 = 28.75 h). Mid: Production Line 7:00–1:00, Shift Lead 7:30–4:00, Barista 9:00–2:30, Closer 10:00–5:30 = 27.5 h (−1.25) — the lead opens at 7:30 and stays to 4:00 so the Closer is the only person 4–5:30. "
        "Slow: Production Line 7–1, Shift Lead 7:30–3:30, Barista 9:30–2, Closer 11–5:30 = 25 h (−3.75).",
        "Monday (first cut: Shift Lead 6:30–2:30, Production Barista 6:30–1, Production Line 6:15–2:30 + Manager 7–2 = 22.75 core h). Mid: Production Line 6:00–1:30, Shift Lead 6:30–2:45, Barista 7:00–11:00 = 19.75 h (−3). Because the Manager block is on the floor at 7:00, the 7–11 barista can be dropped that week (15.75 h) — the delivery is covered by two core plus Admin.",
        "Tuesday (first cut: Production Barista 6:30–1, Shift Lead 7–2:30, Production Line 6:15–2:30 = 22.25 h). Studio's weakest day ($900–1,040): Production Line 6:00–1:00 + Shift Lead 6:30–2:45 = 15.25 h (−7) in every tier; the peak is 20–23/hr, two people at 12 each. Admin covers 8–10 when scheduled.",
        "Wednesday (first cut: Barista 6:30–1, Shift Lead 7–2:30, Production Line 6:15–2:30 = 22.25 h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–2:45, Barista 8:00–1:30 = 19.75 h (−2.5); slow: Barista 9:00–1:00 = 18.25 h (−4).",
        "Thursday (first cut: Shift Lead 7–2:30, Production Line 6:15–2:30, Production Barista 6:30–1:30 + Manager 9–2 = 22.75 core h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–2:45, Barista 8:00–1:00 = 19.25 h (−3.5); slow: Barista 8:30–1:00 = 18.75 h. With the Manager 9–2 on the floor the barista can be 8–12 or dropped.",
        "Friday (first cut: Shift Lead 6:30–3, Production Line 6:15–2:30, second Shift Lead 10–5:15, Barista 7–1 = 30 h). Mid: Production Line 6:00–12:00, Shift Lead 6:30–3:30, Barista 8:30–2:00, Closer 1:30–5:30 = 24.5 h (−5.5) — the second lead becomes a 4-hour Closer at barista rate and the barista starts at 8:30 not 7:00.",
        "Saturday (first cut: Production Line 7:15–3, Shift Lead 10–5:30, Shift Lead 7:30–1, Hospitality Register 8–3, Barista 9–4 = 34.75 h). Mid: Production Line 7:00–1:30, Shift Lead 7:30–4:30, Barista 9:00–2:30, Closer 9:30–5:30 = 29 h (−5.75): one lead through the peak, four people 9:30–1:30, the Closer alone 4:30–5:30. "
        "Slow: Production Line 7–1, Shift Lead 7:30–4:30, Barista 9–2, Closer 10:30–5:30 = 27 h (−7.75).",
        "Net: first cut 183.5 h / $3,718 at the people's rates (32.7% of a $11,360 mid week; 38.2% of a $9,734 slow week) → mid template 155 h / $3,074 (27.1% std, 27.6% roster) → slow template 147.5 h / $2,929 (30.1% / 30.6%). Ryan's own scoring of the first cut (30.8% on a $12,124 projected week) is consistent: "
        "the first cut is a busy-tier schedule and runs 37–38% if September–October land where August did.",
    ],
    "1899 Penn": [
        "Monday (first cut: Shift Lead 6:30–3:30 + Production Barista 6:30–1:30 + Manager 8–2 = 16 core h). Mid: Production Barista 8:30–2:00 = 14.5 h (−1.5); slow: 7:30–11:30 = 13 h (−3). With the Admin block on 8–2 the second person is a 4-hour rush shift in any tier.",
        "Tuesday (first cut: Shift Lead 6:30–3:30 + Production Barista 6:30–3:30 + Manager 6:30–1:30 = 18 core h + 7 admin). Three people on a $850–1,100 morning. Production Barista 8:30–2:30 = 15 h (−3) — the lead opens alone 6:30–8:30 (2–8 transactions an hour) and closes alone after 2:30.",
        "Wednesday and Thursday (both 6:30–3:30 = 18 h; Thu 17.5 h). Production Barista 8:30–2:30 = 15 h each (−3 / −2.5). Both people at 6:30 for a 7:00 open with 5–7 transactions in the first hour is the visible overlap.",
        "Friday (first cut: Shift Lead 6:30–3:30 + Production Barista 6:30–1:30 = 16 h). Friday is 52–68 transactions a day: Production Barista 7:30–11:30 = 13 h (−3) in slow and mid weeks (8:30–12:30 in busy weeks); the lead is alone from 11:30 (4–10 an hour).",
        "Busy weeks: the second person works 7:30–3:00 on Tue/Wed (16.5 h) and the 1pm peak (30–37/hr) is 'Admin covers 1:00–2:00' — a third person is not staffed for a one-hour spike.",
        "Rate: 1899's second person is often a $22.75 team member (Production Barista); a third of 1899's core hours are paid at $22.50+. At $19 that is ~$80/wk.",
        "Net: first cut 85.5 h / $1,756 at the people's rates (46.7% of a $3,759 mid week) → mid template 72.5 h / $1,468 (39.0% std, 40.2% roster) → busy template 77 h / $1,553 (33.8%). 35% at mid-tier sales needs either ~$4,300/wk of sales or the ALL covering one peak a day. Columbus Day (10/12): closed.",
    ],
}
V10_CHANGES = {
    "Downtown": [
        "The Production Line shift is 6–7.5 h (prep + rush + lunch) instead of 8 h; an 8-hour Production Line adds $105–150/wk (variant priced above). Confirm with Mo that prep fits in the shorter window.",
        "The Shift Lead closes at 2:30–3:30 with a barista-rate Closer finishing the day; Mid Friday drops from 3 people all day (23 h) to 18.5 h with the afternoon single-covered; the 2-person floor now ends at 1:00p rather than 2:30p.",
        "Rush baristas are 4–6 h and placed by the curve (8:00/8:30 starts) rather than 7:30; busy Tue/Wed run two of them plus the Closer for a 5-person 8:30–10:00 peak. The 10-hour open-to-close lead is no longer needed anywhere (+$35/wk to avoid it).",
    ],
    "Studio": [
        "Weekend Production Line 7:00–1:00/1:30 (6–6.5 h) and one lead 7:30–4:30; the Hospitality Register 8-hour shift and the fifth person are gone — four people 9:30–1:30 cover a 39–46/hr peak at 12 per person-hour.",
        "Wed/Thu get a 4–5.5 h rush barista in the slow tier too (v1.0 ran them as 2-person days); Tuesday stays a 2-person day in every tier; Monday's third person is 7:00–11:00 (open) rather than 6:30–10:30.",
        "Friday: Barista 8:30–2:00 and a 4-hour Closer 1:30–5:30 instead of a Closer from 9:00 or 10:30. Net −6 / −21 / −26 h per week against v1.0 in slow / mid / busy.",
    ],
    "1899 Penn": [
        "The second person's day is anchored on the two peaks — 8:30–2:30 (6 h) instead of 7:30–2:00/2:30 — and the lead opens alone 6:30–8:30; Monday and Friday carry a 4-hour second shift (7:30–11:30) in slow weeks and on mid Fridays; a mid Monday runs 8:30–2:00.",
        "The busy tier keeps two people 7:30–3:00 on Tue/Wed and flags the 1pm spike as Admin cover instead of staffing a third person; net −2 / −3.5 / −5 h per week against v1.0.",
        "The 'lean' variants of v1.0 are now the main template; the 'comfort' variant (two people 7:00–1:30) is priced in Section 7 at +$115–200/wk.",
    ],
}


def nmwa_section(doc):
    add_heading(doc, "6. NMWA (for completeness)", level=1)
    nm = TIERS["NMWA"]["typical_day"]["slow"]
    add_para(doc, f"NMWA is a 10a–2p museum cafe with a 1-person staffing rule (2 on community days). Typical days: Thu {m(nm['Thu']['net'])}, Fri {m(nm['Fri']['net'])}, Sat {m(nm['Sat']['net'])}, Sun {m(nm['Sun']['net'])}; community Wednesdays {m(nm['Wed']['net'])} (4 observed). "
                  "The sampled shifts run 9:00–3:00 (6 hours for a 4-hour window) and three of the five 16-hour-plus missed clock-outs in the whole sample were NMWA FOH punches in the last week of August. Template (unchanged from v1.0): FOH 9:30–2:30 Thu–Sun (5 h), plus a second FOH 10:00–2:00 on the "
                  "first Sunday and second Wednesday. Ryan's 10/11 first cut has no Wednesday 10/14 shift — that is the October community day — and the shifts are 9–3; the change is +1 community shift and −1 hour on each regular shift (net ≈ +$100/wk on a $1,000 week). "
                  "Labor % is not a useful lens here: at 5 hours a day the wage bill is 56% of typical sales; the lever is sales. NMWA hours count in the incentive split (same people) but NMWA is not in the metric.", after=6)


def sensitivity_section(doc):
    add_heading(doc, "7. Sensitivity, variants and tier triggers", level=1)
    add_para(doc, "August's failure mode was a template built for one tier meeting sales from a lower one. Each template is scored at 10% under its tier's typical sales, at the roster's current rates, and a trigger is given: the weekly projected "
                  "(or trailing two-week actual) sales below which the template breaches the ceiling and the next-slower template should be used. Because hours do not scale with sales inside a tier, a 10% sales miss costs 3–4 points of labor at every cafe.", after=6)
    rows = []
    for c in CAFES:
        tg = TGT[c]
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier)
            rows.append([f"{c} — {tier}", m(w["sales_typical"]), m(w["wages"]), p(w["labor_pct"]), p(w["labor_pct_obs_rate"]), p(w["labor_pct_minus10"]), p(w["labor_pct_plus10"]), p(tg["ceiling"] * 100, 0), m(w["ceiling_trigger_sales"]), m(w["ceiling_trigger_sales_obs_rate"]), m(w["target_trigger_sales"])])
    table(doc, ["Template", "Typical wk sales", "Wages (std)", "% std", "% roster", "At −10% sales", "At +10% sales", "Ceiling", "Ceiling trigger (std)", "Ceiling trigger (roster)", "On-target at (wk sales)"], rows,
          widths=[1.25, 0.7, 0.65, 0.5, 0.55, 0.62, 0.62, 0.5, 0.75, 0.8, 0.75], key_rows=[0, 3, 6], size=8)
    add_caveat(doc, "Reading the triggers: run the Downtown mid template while projected weekly sales stay above ~$6,900–7,000; below that it breaches 36% and the slow template applies (which itself breaches below ~$5,800–6,000). Studio's mid template holds to ~$9,000–9,100 and the busy template to ~$9,800. "
                    "1899's mid template breaches 40% below ~$3,700 (i.e. in the bottom half of its mid band) — the lean shape is already the template, so below that the lever is the ALL covering a peak. Weekly report: add a line per cafe showing 'tier in force' and the trigger distance.")
    add_heading(doc, "Variants priced — what each relaxation of the rules costs per week", level=3)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            w = W(c, tier); v = S["templates"][c][tier]["variants_week"]
            rows.append([f"{c} — {tier}", f"{h(w['hours'],1)} h / {m(w['wages'])} / {p(w['labor_pct'])}",
                         f"{sm(v['target']['wages']-w['wages'])} → {p(v['target']['labor_pct'])}", f"{sm(v['comfort']['wages']-w['wages'])} → {p(v['comfort']['labor_pct'])}",
                         f"{sm(v['min6h']['wages']-w['wages'])} → {p(v['min6h']['labor_pct'])}", f"{sm(v['pl8']['wages']-w['wages'])} → {p(v['pl8']['labor_pct'])}"])
    table(doc, ["Template", "v1.1 (no shift > 9 h)", "10-hour lead option", "Comfort (cap 12, longer 2-person floors)", "6-hour minimum shifts", "8-hour Production Line"], rows, widths=[1.25, 1.35, 1.0, 1.35, 1.0, 1.0], key_rows=[0, 3, 6], size=8)
    add_caveat(doc, "'Comfort' = capacity 12 at every cafe, Downtown's 2-person floor to 2:30p, Studio's to 60 minutes before close, 1899 two people 7:00–1:30p — the v1.0-style rules. At Studio it costs almost nothing because the capacity is already 12; at Downtown it is $190–370/wk and at 1899 $115–200/wk. "
                    "The 6-hour-minimum variant shows what 4-hour rush shifts are worth ($0–215/wk); the 8-hour Production Line variant is the price of keeping prep shifts at 8 hours ($0–170/wk).")
    add_heading(doc, "Capacity sensitivity — need hours per week at 12 / 14 / 16 transactions per person-hour", level=3)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            v = S["need_hours_by_capacity"][c][tier]
            rows.append([f"{c} — {tier}", h(float(v["12"])), h(float(v["14"])), h(float(v["16"])), f"{NV['cap']['target'][c]:.0f}"])
    table(doc, ["Cafe / tier", "Need at 12", "Need at 14", "Need at 16", "Used"], rows, widths=[1.5, 1.0, 1.0, 1.0, 0.8], key_rows=[0, 3, 6], size=9)
    add_caveat(doc, "Need hours are the area under the people-needed curve before shift shapes are applied; the floors dominate in slow weeks (need barely moves with capacity) and the peaks dominate in busy weeks (8–12 h/wk between 12 and 16). "
                    "If the teams report the templates feel thin at the peaks, the honest lever is capacity 12 at Downtown and 1899 (+$100–230/wk), not an extra all-day person.")


def incentive_section(doc):
    add_heading(doc, "8. Incentive — structures, payouts, margins, comparables, implementation", level=1)
    add_para(doc, "Purpose: give the team a visible, self-funding share of the upside from running to the templates — the 'shared upside' that tips used to provide — without crowding out the annual profit share into retirement accounts that Ryan wants in 1–2 years. "
                  "Everything is on the wage basis (gross core wages ÷ core sales), the number the team sees and controls. Payouts are wages (taxed and loaded at 12% like any other pay). Nine structures are compared on the same scenarios, then simulated over a full year.", after=6)
    hb = INC["headroom_by_tier"]
    add_heading(doc, "8a. Headroom — what running to target is worth, by tier", level=2)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            v = hb[c][tier]
            rows.append([f"{c} — {tier}", m(v["core_sales_month"]), m(v["actual_wages_month"]), p(v["actual_pct"]), m(v["target_wages_month"]), m(v["headroom_actual_vs_target"]), m(v["template_wages_month"]), p(v["template_pct"]), p(v["template_pct_obs_rate"]), m(v["beat_at_template_obs"])])
    tots = {t: sum(hb[c][t]["headroom_actual_vs_target"] for c in CAFES) for t in ["slow", "mid", "busy"]}
    rows.append(["Open Cafes — slow / mid / busy", "", "", "", "", f"{m(tots['slow'])} / {m(tots['mid'])} / {m(tots['busy'])}", "", "", "", ""])
    table(doc, ["Cafe / tier", "Core sales / mo", "Actual core wages", "Actual %", "Target wages", "Headroom vs target", "v1.1 template wages", "Template % std", "Template % roster", "Beat at template (roster)"], rows,
          widths=[1.45, 0.7, 0.75, 0.55, 0.7, 0.75, 0.75, 0.6, 0.65, 0.75], total_rows=[9], size=8)
    add_caveat(doc, f"Monthly = weekly × 52/12. 'Actual' is the sampled weeks of each tier at the rates paid. August's realised headroom was {m(INC['headroom_aug_total'])} (three cafes, core only); the Apr–Jun average was {m(INC['headroom_aprjun_avg_total'])} a month. "
                    "'Beat at template' is what the pool would be computed on if the roster ran the template exactly: positive at Downtown/Studio in mid and busy months, negative (no payout) in slow months and at 1899 below busy — "
                    "so the plan pays in the months where the money is, and the slow-month templates protect the P&L without a bonus.")

    add_heading(doc, "8b. Nine structures on the same scenarios (mid month, three open cafes pooled)", level=2)
    comp = INC["plan_comparison"]; oc = comp["scenarios"]; PL = INC["assumptions"]["plans"]
    add_para(doc, f"Open-cafe core sales {m(comp['mid_month']['open_cafe_sales'])} and {comp['mid_month']['open_cafe_core_hours']:,} core hours in a mid month on the v1.1 templates. Each cell: company payout for the month (as % of core sales) · what a 30 h/wk team member receives.", after=4)
    rows = []
    for key, label in PL.items():
        r = [f"{key}. {label}"]
        for scen in ["AT target", "1 pt under", "2 pts under", "3 pts under", "4 pts under"]:
            v = oc[scen][key]; r.append(f"{m(v['payout'])} ({v['pct_of_sales']:.1f}%) · ${v['per_person_month_30h']}")
        rows.append(r)
    table(doc, ["Structure", "At target", "1 pt under", "2 pts under", "3 pts under", "4 pts under"], rows, widths=[2.5, 0.84, 0.84, 0.84, 0.84, 0.84], size=7)
    mx = comp["mixed_example"]
    add_caveat(doc, f"Nothing pays when the pool is over target under any structure. Mixed month ({mx['description']}): A pays {m(mx['A'])}, per-cafe pools (F) {m(mx['F'])} — Downtown's team earns nothing under F for a month in which they hit their number, and 1899's over-run is not netted; "
                    f"B pays {m(mx['B'])} because Downtown 'hit target'. Plan G is A settled quarterly: same money in a steady year, less in a lumpy one (a bad month nets against two good ones).")
    add_para(doc, "How they compare on the criteria that matter here:", after=2)
    crit = [
        ("Self-funding", "A, A2, A3, C, E, F, G pay only out of dollars actually saved and Slipstream keeps at least half; B and D pay 1.5–1.6% of sales ($1,500–1,600 a month) for merely hitting target — new cost with no saving to fund it."),
        ("Simplicity", "B is the shortest sentence ('hit the number, everyone gets a dollar an hour'). A is one sentence longer ('we split what we save below the line 50/50, up to about a dollar an hour') and it shows the team the same number Ryan watches. E adds COGS, which the floor team only partly controls (waste yes, pricing and vendor cost no). C is too small to notice."),
        ("Cliffs and gaming", "A2's 2-point gate is a dead zone: a 1.9-point beat pays nothing and a 2.0-point beat pays ~$1,000, an invitation to cut a shift on the 28th. B's steps ($1.00 → $1.50 → $2.00) have the same problem three times. A, A3, C, E, F, G pay from the first dollar and cap, so there is nothing to game except the guardrail."),
        ("Controllability", "Hourly staff control hours far more than sales. Labor-only metrics (A-family, B, D) are the most controllable; E rewards a COGS beat the team may not have caused (or punishes a vendor price rise they did not). Do not pay on SPLH or sales alone."),
        ("Fairness across cafes", "23 of the 25 people with core hours work at more than one cafe, and 1899's target is structurally out of reach below busy-tier sales. Pooled plans (A, A3, C, E, G) treat a Downtown shift and a 1899 shift the same; per-cafe pools (F, A2, B, D) create haves and have-nots on the same roster."),
        ("Service risk", "Every plan that pays on labor needs the same guardrail — template floors honoured, no skeleton rush, sales ≥ 90% of projection — or it rewards under-staffing. Plans that pay on prime cost (E) carry the same risk plus a COGS-cutting one (portioning)."),
        ("Payroll and FLSA", "A non-discretionary bonus must be included in the regular rate for overtime; paying the pool as a percentage of each person's core wages for the month (including any overtime premium) satisfies the rule without recomputation; a flat per-hour split (B, D) needs an overtime true-up each month. Details in 8h."),
        ("Room for the retirement plan", "At the cap A costs at most 1.5% of core sales (~$18K/yr if every month capped, unrealistic; $3–6K realistic). B/D at ~1.6% of sales in every on-target month ($10–13K/yr) plus a 5–10% profit share would push total variable pay to 2.5–3% of sales before the templates have proven out."),
    ]
    for k, v in crit: add_bullet(doc, v, bold_lead=k + ".")

    add_heading(doc, "8c. A full year, simulated — what each plan costs against what the templates save", level=2)
    sim = INC["annual_simulation"]; sv = sim["labor_saving_vs_aprAug_run_rate"]
    add_para(doc, f"2,000 runs of a 2027 tier calendar (Section 2; December slow, Sep–Nov mid). Each month every cafe sells its tier-typical sales ±6% and runs the v1.1 template at the roster's current rates plus a 0.5-point drift and ±1.5 points of noise — "
                  f"i.e. the templates are followed but not perfectly. Against the Apr–Aug 2026 run-rate ({INC['base_actual_core_pct_aprAug']['Downtown']}% / {INC['base_actual_core_pct_aprAug']['Studio']}% / {INC['base_actual_core_pct_aprAug']['1899 Penn']}% core wages at Downtown / Studio / 1899) the templates save "
                  f"{m(sv['p10'])}–{m(sv['p90'])} of core wages a year (median {m(sv['p50'])}). What each plan pays out of that:", after=4)
    SHORT = {"A": "A. Beat Share 50%, cap 1.5%, pooled", "A2": "A2. Beat Share 50%, 2-pt gate, no cap", "A3": "A3. Beat Share 40%, cap 1.0%, pooled", "B": "B. Dollar-an-hour ($1 / $1.50 / $2)",
             "C": "C. Gainshare 20%, no gate, no cap", "D": "D. $8 per shift in on-target weeks", "E": "E. Prime-cost gainshare 50%, cap 2%", "F": "F. Beat Share 50%, per-cafe pools", "G": "G. Beat Share 50%, settled quarterly"}
    rows = []
    for key in PL:
        v = sim["plans"][key]; per30 = v["p50"] / INC["team"]["total_core_hours_per_week"] * 30
        rows.append([SHORT[key], m(v["p10"]), m(v["p50"]), m(v["p90"]), p(v["share_of_saving_p50"]), m(v["p50"] * 1.12), f"{m(per30)} / yr ({m(per30 / 12)} / mo)"])
    table(doc, ["Plan", "Annual payout p10", "p50", "p90", "Share of saving", "Employer cost incl. 12% load (p50)", "Per 30 h/wk person (p50)"], rows, widths=[2.2, 0.8, 0.7, 0.7, 0.75, 1.0, 1.1], size=8)
    add_caveat(doc, "Payouts are pooled across the open cafes and split by core hours (447 core hours a week across 25 people in the sample; a 30 h/wk person is ~7% of the pool). B and D pay in every on-target month regardless of the size of the beat, "
                    "which is why they cost 2–3× A for the same behaviour. G (quarterly) pays ~20% less than A because bad months net against good ones — a feature for the P&L, a bug for motivation. "
                    "Note the drift assumption: if the roster runs 1.5 points above the templates instead of 0.5, A's payout roughly halves and B/D barely move — the beat-share plans are the ones that scale with performance.")

    add_heading(doc, "8d. Recommendation — 'Beat Share', pooled, capped, guarded (unchanged in substance from v1.0, tightened in mechanics)", level=2)
    for line in [
        "Metric: monthly core wages ÷ monthly core sales per cafe, on the Square-timecard basis of the weekly report (Manager and Training excluded; missed clock-outs capped by the past-close rule). Published weekly, settled monthly.",
        "Pool: 50% × Σ over Downtown, Studio and 1899 of (target% × core sales − actual core wages), where a cafe over its target contributes a negative amount. If the sum is negative, no payout. Cap the pool at 1.5% of open-cafe core sales (≈ $1,500 in a mid month — about $1.00 per core hour on the templates, $1.15–1.20 in a busy month because sales per hour are higher).",
        "Split: as a percentage of each person's core wages earned in the month at any open cafe or NMWA (payout ÷ total core wages = the bonus %; each person receives that % of their own core wages, including any overtime premium). This is the FLSA-clean allocation; it also means an SL earns ~10% more bonus per hour than a barista, which is defensible. If Ryan prefers a flat per-hour split, payroll must add the overtime true-up.",
        "Guardrail (the service test): the pool pays only if the month's published schedules met the template floors — no single-person coverage inside a rush block (8:30–10:00 at Downtown/1899, 9:30–1:00 at Studio weekends) on more than two days, no week flagged for under-coverage in the weekly report, and core sales ≥ 90% of the month's projection (so the plan does not pay in a month the business shrank). Admin cover of a rush counts as coverage.",
        "Timing: shadow-mode in the September and October weekly reports (show what the pool WOULD have paid); first real payout after two consecutive months at or under target company-wide (Ryan's rule), realistically December or January; pay in the first payroll of the following month as its own earnings code ('Beat Share').",
        f"Size discipline: at the cap the plan is ~$18K a year if every month capped, which would need the templates beaten by 3+ points all year; the simulation says {m(sim['plans']['A']['p50'])} (p10–p90 {m(sim['plans']['A']['p10'])}–{m(sim['plans']['A']['p90'])}). "
        f"Together with a 5–10% annual profit share on consolidated income at target ({m(INC['profit_share_sizing']['share_5pct_of_consolidated'])}–{m(INC['profit_share_sizing']['share_10pct_of_consolidated'])}) total variable pay stays at ~1–1.5% of core sales.",
        "Kill-switch: if the guardrail is breached in two months of any four, or COGS runs 2+ points above plan for two consecutive months at a cafe (a sign of under-staffed production or waste), pause the plan at that cafe and review the templates before resuming.",
    ]:
        add_bullet(doc, line)

    add_heading(doc, "8e. Resulting margins — P&L labor, prime cost and NOI by tier under each path", level=2)
    add_para(doc, "Monthly P&L per cafe at each tier's typical sales, with COGS at the cafe's Apr–Aug average (Downtown 36.8%, Studio 29.2%, 1899 38.2%), operating expenses at the Apr–Aug monthly average, admin and training at August levels, and labor through the bridge (1.12 × gross − tips). "
                  "The last row of each block uses Ryan's planning COGS (28% / 28% / 30%) to show what the two programs together are worth.", after=4)
    mv = INC["margins_by_tier"]
    keep = ["Actual (sampled weeks of this tier)", "v1.1 template at current people's rates", "Exactly at target", "2 pts under target, Plan A (50% beat, cap)", "2 pts under target, Plan B ($1.50/hr)", "At target, Plan D ($8/shift)", "2 pts under target, Plan A, planning COGS"]
    for c in CAFES:
        rows = []
        for tier in ["slow", "mid", "busy"]:
            for i, k in enumerate(keep):
                r = mv[c][tier]["rows"][k]
                rows.append([f"{tier} — {k}" if i == 0 else f"   {k}", m(mv[c][tier]["core_sales_month"]) if i == 0 else "", p(r["core_wage_pct"]), p(r["labor_pct"]), p(r["prime_pct"]), m(r["noi"]), p(r["noi_pct"]), m(r["incentive"])])
        add_heading(doc, f"{c}", level=3)
        table(doc, ["Case", "Core sales / mo", "Core wage %", "P&L labor %", "Prime %", "NOI", "NOI %", "Incentive"], rows, widths=[2.6, 0.75, 0.65, 0.65, 0.6, 0.7, 0.55, 0.6], key_rows=[0, 7, 14], size=7)
    add_caveat(doc, "Reading it: at Downtown the template alone moves a slow month from −10% to +6% NOI and a mid month from 5% to 14%; Plan A at a 2-point beat costs 1 point of NOI and B costs 2–2.5. 1899 is the outlier — even at target its prime cost is 71% because COGS is 38%; "
                    "the planning-COGS row shows that the COGS move is worth as much as the labor move there. The consolidated picture nets the support entities: "
                    f"Commissary + SRP + Cafe Management + Central averaged {m(INC['support_burden_month_normalized'])} a month Apr–Jul (August was {m(INC['support_burden_aug'])} with an $8.1K Central expense spike).")

    add_heading(doc, "8f. Comparable plans", level=2)
    rows = [
        ["Restaurant prime-cost gainshare (industry guidance: RestaurantOwner, Toast)", "Prime-cost ceiling (e.g. 60%); 50% of savings below it to managers / key staff; monthly or quarterly", "Same 50% split as Plan A; usually manager-only, Slipstream's is team-wide; usually uncapped"],
        ["Zingerman's Community of Businesses (Ann Arbor)", "Open-book management since 2002; gainshare = a share of net operating profit gain, split equally among all staff when targets are hit; deferred by staff vote in 2008", "Profit-based and equal-split — closer to Ryan's annual plan than the monthly one; the weekly 'huddle' number is the model for the weekly report line"],
        ["USHG 'Hospitality Included' (New York, 2015–2020)", "Tips eliminated, prices raised, revenue share allocated by hours × level (100/200/300); unwound July 2020 back to tipping with a kitchen revenue share kept", "Cautionary: revenue share without a cost anchor was hard to sustain; Slipstream's guaranteed-rate base already exists, the incentive is upside only"],
        ["No-tip independents (e.g. Juliet, Somerville MA)", "Living-wage base + profit sharing + open books", "Small-operator precedent for base + share; annual profit share is the durable part"],
        ["Restaurant manager bonus norms", "10–20% of salary on prime cost and sales, quarterly", "Relevant for the ALL/LM layer (outside this plan); the team plan should stay at 1–2% of wages"],
        ["Retirement profit share (small employers)", "401(k) with discretionary profit-sharing (flexible allocation incl. hours-weighted points, vesting schedule, testing, most admin); SEP IRA (uniform % of pay for all eligible, immediate vesting, simplest); SIMPLE IRA (2% nonelective or 3% match, $17,000 deferral limit in 2026)", "For 'weighted by hours or comp' with vesting → 401(k) profit sharing; for 'simplest possible' → SEP at a uniform % of pay. A TPA should confirm before anything is promised"],
    ]
    table(doc, ["Plan", "Mechanics", "Read-across to Slipstream"], rows, widths=[1.7, 2.6, 2.4], size=8)
    add_caveat(doc, "Directional, from public descriptions (Toast, RestaurantOwner, ZingTrain / Zingerman's, Harvard Business School case on USHG, Restaurant Dive on the 2020 reversal, 2026 plan-limit guides). Slipstream's proposal — 50% beat share, 1.5% cap, plus a later 5–10%-of-profit retirement contribution — sits inside these ranges "
                    "and is on the generous side for the monthly piece relative to the size of the business, which is appropriate given the removal of tip upside, provided the cap holds.")

    add_heading(doc, "8g. Sizing the annual retirement share", level=2)
    ps = INC["profit_share_sizing"]; tm = INC["team"]
    add_para(doc, f"On the 2027 tier calendar with every cafe exactly at target, open-cafe operating income is ≈ {m(ps['open_cafe_noi_year_at_target_calendar'])} a year ({m(ps['open_cafe_noi_year_template_current_rates'])} on the templates at current rates); after the support entities "
                  f"(−{m(abs(ps['support_burden_year_normalized']))} a year) consolidated net income is ≈ {m(ps['consolidated_ni_year_at_target'])}. A 5–10% share of consolidated income is {m(ps['share_5pct_of_consolidated'])}–{m(ps['share_10pct_of_consolidated'])} a year — "
                  f"about {m(ps['per_person_at_10pct_consolidated_30h'] / 2)}–{m(ps['per_person_at_10pct_consolidated_30h'])} for a 30 h/wk team member, or 1.3–2.6% of pay. Measured on cafe operating income instead it would be {m(ps['share_5pct_of_cafe_noi'])}–{m(ps['share_10pct_of_cafe_noi'])}. "
                  f"For scale, 2026 consolidated net income is +$82K through August (positive in five of eight months, with March carrying $46K of it; January, February and August were losses). "
                  f"The team today: {tm['people_with_core_hours']} people with core hours, median {tm['weekly_core_hours_p50']} h/wk (quartiles {tm['weekly_core_hours_p25']}–{tm['weekly_core_hours_p75']}), {tm['multi_cafe_people']} of them at more than one cafe. "
                  "Two design consequences: (1) the annual share is affordable in a year like 2026 only if the August-type months are removed — which is exactly what the slow-tier templates do — so make 2027 the first plan year and set the percentage after Q1; (2) at these sizes the plan's administration cost matters — a SEP at a uniform percentage of pay is the cheapest way to deliver $6–12K, "
                  "a 401(k) profit-sharing plan is the way to weight by hours and vest.", after=6)

    add_heading(doc, "8h. Implementation plan and compliance notes", level=2)
    for line in [
        "September 2026: publish the tier bands and the 'tier in force' per cafe in the weekly report; build the 10/11 and later weeks on the Mid templates; add the shadow line ('what the Beat Share would have paid') and the guardrail checks (floor coverage, sales vs projection) to the weekly report.",
        "October: templates live for all three cafes; first monthly review of core % vs target on the wage basis; fix the data-hygiene items that distort the metric (adopt the past-close missed-clock-out cap in aggregate_labor.py; impute zero-rate timecards; check the $18.00 rates against the $18.40 DC minimum).",
        "November: if September and October both landed at or under target company-wide, announce the Beat Share with the guardrail and the cap; otherwise extend shadow mode one month. Team communication: one page — the formula, the guardrail, the cap, an example month, and the basis note (core wages 30% on the schedule = ~25% on the P&L).",
        "Payroll mechanics: a separate ADP earnings code; bonus = pool ÷ total core wages of participants × each person's core wages for the month; book it in the month earned (accrue in the close as a JE to the cafe's FOH Labor lines so the monthly report's bridge still ties) and pay it on the first pay date of the following month. Because it is non-discretionary, it is part of the regular rate — the percentage-of-wages method handles overtime automatically.",
        "December–January: first payout. Review the cap and the split after two payouts; decide whether 1899 stays in the pool or moves to its own guardrail (NI ≥ 5% of sales, which is also what earns Slipstream its incentive fee from the owner).",
        "2027: once consolidated income has been positive for four consecutive quarters, stand up the retirement vehicle (TPA/CPA to confirm 401(k) profit-sharing vs SEP, the allocation formula, vesting and testing), with the Beat Share cap re-set so the two together stay ≤ 2.5% of core sales.",
    ]:
        add_bullet(doc, line)
    add_heading(doc, "8i. Things to decide, and thoughts on the design", level=3)
    for line in [
        "Pooled vs per-cafe: pooled is recommended (floaters, 1899's structural gap). If per-cafe, give 1899 a different metric — NI ≥ 5% of sales — or its team will never earn.",
        "Percentage-of-wages vs per-hour split: percentage-of-wages is recommended (FLSA-clean, rewards seniority mildly). If per-hour, cap the individual amount at $1.25/hr and add the overtime true-up.",
        "Basis honesty: the team will see 'core wages 30%' while the P&L says 25%. Publish the bridge once so nobody thinks the number is being managed.",
        "Keep admin outside the metric but inside the guardrail: an ALL who covers the rush from an Admin block is doing exactly what the plan wants; an ALL who moves core hours into Admin to game the ratio is not — track admin hours per week alongside.",
        "The templates are the floor, not the ceiling: the plan pays for beating target, so the guardrail (floors honoured) is the only thing stopping a race to skeleton crews. Make the floors explicit in Square Teams as the minimum shift set per day.",
        "The cap is the promise-keeper for the retirement plan. Write it into the one-pager as a fixed maximum (1.5% of core sales, and never more than $1.25 per core hour) so it cannot drift.",
    ]:
        add_bullet(doc, line)


def data_notes(doc):
    add_heading(doc, "9. Data notes and method", level=1)
    dn = S["data_notes"]
    add_bullet(doc, f"Timecards: {dn['timecards_cafes']} cafe timecards ({dn['timecards_harvested']} harvested incl. 5 commissary/roasting cards) via labor.searchTimecards for 14 weeks — Jan 5, Feb 2, Feb 23, Mar 9, Apr 13, May 11, Jun 8, Jul 13, Jul 27, and every week Aug 3–Sep 6 (Sep partial). "
                    f"Every tier has at least 3 sampled weeks per cafe. {dn['dropped_zero_length']} zero-length punches dropped; {dn['zero_rate_imputed']} zero-rate cards given the team member's usual rate. The Square connector was not re-queried for v1.1; the v1.0 harvest was re-used unchanged.")
    add_bullet(doc, f"Missed clock-outs: {dn['capped_ruleA_16h']} punches ≥16 h were capped to the median shift for that cafe and role (task rule); a second rule capped {dn['capped_ruleB_past_close']} punches that ran more than 2¼ hours past close (Studio Aug 24/25/27 at 15.3, 15.2 and 10.1 h; 1899 Apr 17 at 10.9 h). "
                    f"Phantom wages removed: {m(dn['capped_wages_avoided'])}. The Studio cases (a 6:30a punch closed at 9:47p on a 2p-close day) pass a 16-hour test — the weekly pipeline should adopt the past-close rule too.")
    add_bullet(doc, f"Transactions: {dn['transactions_rows_kept']:,} walk-in payments Jan–Aug 2026 (Complete only, net > 0, invoice-sourced and Management/Equipment Fee rows excluded). Hour of day from the Eastern 'Time' field. Curves are per open day, averaged within cafe × tier × day-of-week (5–17 days per cell).")
    add_bullet(doc, f"Capacity (new in v1.1, per cafe): transactions per person-hour in the busiest fifth of observed hours — median {CAP['Downtown']['busy_hours']['tph_p50']} / {CAP['Studio']['busy_hours']['tph_p50']} / {CAP['1899 Penn']['busy_hours']['tph_p50']} (Downtown / Studio / 1899), "
                    f"90th percentile {CAP['Downtown']['busy_hours']['tph_p90']} / {CAP['Studio']['busy_hours']['tph_p90']} / {CAP['1899 Penn']['busy_hours']['tph_p90']}, maxima 21 / 18.7 / 21. Planning capacity = the p90: 14 / 12 / 14. Studio weekends run 9.3 median with 4–5 people on the clock and $10.70 tickets; "
                    "adding a 4th person at Downtown or a 5th at Studio did not raise busy-hour throughput, which is the statistical signature of over-coverage.")
    add_bullet(doc, S["definitions"]["need_rule"] + ". Prep starts 60 min before open (30 at 1899); close runs 30–45 min after.")
    add_bullet(doc, "Optimiser (new in v1.1): " + S["definitions"]["optimizer"] + "; +$12 per shift so ties go to fewer people; barista shifts placed on the busiest half-hours. Solved with PuLP/CBC in under a second per day; every solution was re-checked half-hour by half-hour against the need curve (zero under-coverage except the flagged 1899 busy-day spikes).")
    add_bullet(doc, "Standard rates in the templates are inputs (blue in the workbook): $21 Shift Lead, $19 barista/production/register, $19.50 closer. 'Roster' figures re-price the same hours at the cafe's observed average core rate ($20.73 / $20.20 / $20.84) — the gap is the cost-per-hour lever, reported separately rather than baked into the templates.")
    add_bullet(doc, "Incentive model: margins use Apr–Aug average COGS % and monthly operating expenses per cafe, August admin/training levels, the 1.12 load and cafe tip rates; the annual simulation is 2,000 runs of the 2027 tier calendar with ±6% sales noise and template % + 0.5 ± 1.5 points of execution noise. "
                    "Support burden = Apr–Jul average of Commissary, SRP, Cafe Management and Central net income.")
    add_bullet(doc, f"Projections: {dn['projections_file']}. Financials: Consolidated Model v17 extract (fin_2026_v17.json) and QBO P&L by location (Coefficient Output 2026.xlsx, Sep 3). Neither the Consolidated Model nor the monthly report was modified. Reference JSON: _pipeline/monthly/data/labor_deep_dive_2026-09.json (v1.1; v1.0 kept alongside). "
                    "Scripts: ldd_10_need_v11.py, ldd_11_optimize_templates.py, ldd_12_incentive_v11.py, ldd_13_summary_v11.py, build_labor_deep_dive_v11.py, build_schedule_templates_xlsx_v11.py.")
    add_bullet(doc, "Limits: templates are built on tier-average curves; individual days vary ±25% around them, which the Admin flex, the whole-person rounding and the 'comfort' variant absorb. Role names follow Square job titles; pay is per person, so the rate lever depends on who is scheduled. "
                    "Single-person windows are flagged, not judged — break coverage and deliveries on solo shifts are an operating decision. No item-level or catering labor is modelled.")


def appendix(doc):
    add_heading(doc, "Appendix A — Need vs actual, 30-minute detail (Tuesday, all tiers; v1.1 need rules)", level=1)
    add_para(doc, "Transactions per hour at each half-hour (tier average), the people needed under the Section 9 rule, and the average core headcount actually on the clock in the sampled weeks. Full detail for every day is in the workbook.", after=4)
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            rows = NV["variants"]["target"][c]["tiers"][tier]["Tue"]["by_halfhour"]
            def tlabel(t):
                hh = int(t); mm = "30" if t % 1 else "00"; return f"{hh if hh <= 12 else hh - 12}:{mm}"
            hdr = ["Tue " + tier] + [tlabel(x["t"]) for x in rows]
            body = [["txns/hr"] + [f"{x['txns_hr']:.0f}" for x in rows], ["need"] + [str(x["need"]) for x in rows], ["actual"] + [f"{x['actual']:.1f}" for x in rows]]
            add_heading(doc, f"{c} — Tuesday, {tier} tier", level=3)
            table(doc, hdr, body, widths=[0.6] + [0.27] * len(rows), size=6)


if __name__ == "__main__":
    build()
