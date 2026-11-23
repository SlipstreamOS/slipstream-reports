#!/usr/bin/env python3
"""Build 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.4.docx' — the readable, implementation-first
version: schedule grids per cafe × tier, short rules, tier calendar, pay review, condensed incentive. No prior-version comparisons."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, add_callout, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, p, h, sm

DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.4.docx")
J = lambda n: json.load(open(os.path.join(DD, n)))
T = J("templates_v13.json"); NV = J("need_v13.json"); TIERS = J("tiers.json"); BT = J("backtest_v13.json"); AF = J("affordability_v13.json")
CAL19 = J("tiers_v19_sep_dec.json"); S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
AUG = S["august_2026"]; BR = S["pl_bridge_validation"]
CAFES = ["Downtown", "Studio", "1899 Penn"]; DOW5 = ["Mon", "Tue", "Wed", "Thu", "Fri"]; DOW7 = DOW5 + ["Sat", "Sun"]
DOWS = {"Downtown": DOW5, "Studio": DOW7, "1899 Penn": DOW5}
TGT = T["targets"]; WPM = 52 / 12
CODE = {"PL": "Line", "SL": "Lead", "PB": "Prod Barista", "B": "Barista", "CL": "Closer", "FOH": "FOH"}
def W(c, t): return T["cafes"][c][t]["week"]
def D(c, t): return T["cafes"][c][t]["days"]

def grid_table(doc, cafe, tier):
    days = DOWS[cafe]; dd = D(cafe, tier); w = W(cafe, tier)
    nmax = max(len(dd[d]["shifts"]) for d in days)
    rows = []
    for i in range(nmax):
        row = [f"Shift {i + 1}"]
        for d in days:
            sh = dd[d]["shifts"]
            if i < len(sh):
                lab = CODE[sh[i]["code"]]
                if cafe == "Studio" and d in DOW5[:4] and sh[i]["code"] == "CL": lab = "Barista (3rd)"
                row.append(f"{lab} {sh[i]['start']}–{sh[i]['end']}")
            else: row.append("—")
        rows.append(row)
    rows.append(["Hours"] + [h(dd[d]["template"]["hours"], 1) for d in days])
    rows.append(["Typical sales"] + [m(dd[d]["sales_typical"]) for d in days])
    rows.append(["Labor %"] + [p(dd[d]["template"]["labor_pct"]) for d in days])
    widths = [0.9] + [(5.8 / len(days))] * len(days)
    table(doc, [f"{tier.capitalize()} week"] + days, rows, widths=widths, total_rows=[nmax, nmax + 1, nmax + 2], size=8)
    tg = TGT[cafe]
    status = "on target" if w["labor_pct"] <= tg["target"] * 100 + 0.05 else ("inside the band" if w["labor_pct"] <= w["ceiling"] * 100 + 0.05 else "over the ceiling")
    add_caveat(doc, f"Week: {w['hours']:.1f} core hours, {m(w['wages'])} = {p(w['labor_pct'])} of {m(w['sales_typical'])} — {status} (target {p(tg['target']*100,0)}, ceiling {p(w['ceiling']*100,0)}). "
                    f"If sales come in 10% under: {p(w['labor_pct_minus10'])}. " + (f"Switch to the next-slower template when the projected week falls below {m(w['ceiling_trigger_sales'])}." if tier != "slow" else f"There is no slower template: below {m(w['ceiling_trigger_sales'])} of projected weekly sales the week breaches the slow ceiling and admin days are the lever."))

def admin_notes(cafe, tier):
    """Where an admin shift is worth most: (a) replacing the cheapest non-opener/closer shift on the highest-labor-% days
    (slow weeks), (b) the peak windows where need exceeds the cap (busy weeks)."""
    dd = D(cafe, tier); w = W(cafe, tier); out = []
    cands = []
    for d in DOWS[cafe]:
        sh = [x for x in dd[d]["shifts"] if x["code"] in ("B", "CL", "PB")]
        if len(dd[d]["shifts"]) >= 3 and sh:
            x = min(sh, key=lambda z: z["hours"])
            cands.append((dd[d]["template"]["labor_pct"], d, x, x["wages"] / w["sales_typical"] * 100))
    cands.sort(reverse=True)
    if cands and w["labor_pct"] > w["ceiling"] * 100:
        parts = [f"{d} ({CODE[x['code']]} {x['start']}–{x['end']}, −{pts:.1f} pts on the week)" for _, d, x, pts in cands[:3]]
        out.append("swap an admin shift for " + "; ".join(parts))
    peaks = [f"{d} {', '.join(dd[d]['template']['admin_cover_windows'])}" for d in DOWS[cafe] if dd[d]["template"]["admin_cover_windows"]]
    if peaks: out.append("cover the peak above the cap on " + "; ".join(peaks))
    return out

def build():
    doc = Document(); sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.85)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Schedule Templates by Demand Tier", 22, True, COLOR_ACCENT, 0), ("Labor Deep Dive  |  September 2026  |  v1.4", 14, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "Prepared September 10, 2026 for Ryan Fleming. Templates are built to the operating rules agreed Sep 5–10 and priced at the roster's current rates. "
                  "Sources: Square transactions Jan–Aug 2026, 856 Square timecards (14 weeks), Consolidated Model v17 (2025 and 2026), QBO P&L, Projections v1.9. "
                  "The workbook 'Slipstream Schedule Templates - Sep 2026 v1.4.xlsx' has every shift as an editable row, the half-hour demand curves and the month-by-month backtest. "
                  "Pay and incentive are in the separate 'Slipstream Pay & Incentive Review - Sep 2026 v1.0'.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ================= 1. SUMMARY =================
    add_heading(doc, "1. Summary", level=1)
    dt, st, pn = AUG["Downtown"], AUG["Studio"], AUG["1899 Penn"]
    add_bullet(doc, f"Why August lost money at Downtown and 1899: high-$40s sales per labor hour is a 40–47% wage bill at Slipstream's guaranteed rates. The cafes need $60–68 of sales per core labor hour to hit 32/30/35; "
                    f"August delivered ${dt['splh_core_walkin']:.0f} / ${st['splh_core_walkin']:.0f} / ${pn['splh_core_walkin']:.0f}. Studio survived a similar miss because its fixed costs sit on a bigger base and its COGS are 8–9 points lower.", bold_lead="The problem.")
    add_bullet(doc, "The P&L understates the labor miss (tips fund part of the guaranteed wage, so booked labor ≈ 1.12 × gross wages − tips). Manage to the schedule number — gross wages ÷ projected core sales — and let the P&L follow. Detail in Appendix A.", bold_lead="Basis.")
    add_bullet(doc, "Two openers and two closers at Downtown and Studio, 6-hour minimum shifts, a third person mid-day at Studio, a third by 8:00 on Tue–Thu at Downtown, at most four people at Downtown and on Studio Fridays (five on busy Studio weekends), "
                    "two shifts every day at 1899 (stagger on slow days, both close together), NMWA 9:00–3:00. Admin days are NOT in the templates — they are the scheduler's tool for slow weeks; each template says where they help most.", bold_lead="Rules built in.")
    rows = []
    for c in CAFES:
        rows.append([c] + [f"{h(W(c, t)['hours'], 0)} h  ·  {p(W(c, t)['labor_pct'])}" for t in ["slow", "mid", "busy"]] + [f"{TGT[c]['target']*100:.0f} / {TGT[c]['ceiling']*100:.0f} (slow {TGT[c]['ceiling_slow']*100:.0f})"])
    rows.append(["NMWA", "24–30 h  ·  staffing rule", "—", "—", "1 person / day, 2 on community days"])
    table(doc, ["Cafe", "Slow week", "Mid week", "Busy week", "Target / ceiling %"], rows, widths=[1.1, 1.35, 1.35, 1.35, 1.55], size=9)
    add_caveat(doc, "Weekly core hours and core wages as % of the tier's typical sales, at planning rates ($22.25 lead / $19.75 line / $20.25 production barista / $20.00 barista & closer).")
    add_bullet(doc, "With the rules as written the targets are met in busy weeks, the band is met in mid weeks, and slow weeks sit at or just above the slow-month ceiling (38 / 36 / 42). That is what the admin-day tool is for: one admin shift replacing a 6-hour barista shift is worth 2–3 points in a slow week.", bold_lead="What that means.")
    add_bullet(doc, "Downtown: three people most days (four Tue–Thu in mid/busy weeks) instead of four all day. Studio: three people Mon–Thu, four on Friday and slow weekends, five only on busy weekends. 1899: a 12-hour stagger day when sales are under ~$650, a 15-hour day otherwise.", bold_lead="Biggest changes from what was run.")
    add_bullet(doc, "Projections v1.8 were high (Studio ran 13% under 2025 in Jul–Aug; Downtown flat). Projections v1.9 re-bases Sep–Dec; on it Downtown is a mid-tier cafe through December and Studio is slow except October to mid-November. Tier is chosen week by week — Section 2 has the calendar.", bold_lead="Projections.")
    bt3 = BT["three_cafes"]["totals"]; sv = AF["template_saving"]["total"]
    add_bullet(doc, f"Run over the last twelve months (Sep 2025–Aug 2026) with each week's actual sales choosing the tier, the templates would have cost {m(sv['12mo'])} less labor than was paid — "
                    f"{m(sv['jan_aug_2026'])} of it in Jan–Aug 2026 ({m(sv['annualized_2026'])} a year at that pace). Downtown and 1899 are where the money is; Studio's templates cost about what Studio already spends. "
                    f"Section 7 has the month-by-month; the bigger lever is COGS — every point is {m(AF['cogs']['per_point_year'])} a year across the three cafes.", bold_lead="Backtest.")

    # ================= 2. HOW TO USE =================
    add_heading(doc, "2. How to use the templates", level=1)
    for line in [
        "Each week, take the projected core (walk-in) sales for the cafe from the weekly report's forward block (Projections v1.9) and pick the tier from the bands below. Copy that tier's grid into Square Teams.",
        "Months are not uniform — November starts mid and ends slow; December is slow from the week before Christmas. Expect to change tier inside a month; the trigger under each grid says when.",
        "If the week is slow and the labor % is over the ceiling, use an admin day to replace the shift flagged in the 'admin helps most' note rather than cutting an opener or closer.",
        "Roles are names for shifts, not people: 'Closer' is a keyholder barista; 'Lead' opens; the Line person closes the line at Studio. Both baristas on a Studio weekday should be shift-lead capable so the scheduler can swap.",
        "Federal-holiday Mondays: Downtown and 1899 closed. Studio Monday keeps three at open for the standing delivery.",
    ]: add_bullet(doc, line)
    add_heading(doc, "Tier bands — projected weekly core sales", level=3)
    rows = []
    for c in CAFES:
        b = S["tiers"][c]["bands_weekly_walkin"]; rows.append([c, f"under {m(b['slow'][1])}", f"{m(b['mid'][0])} – {m(b['mid'][1])}", f"{m(b['busy'][0])} and up"])
    table(doc, ["Cafe", "Slow", "Mid", "Busy"], rows, widths=[1.4, 1.6, 1.8, 1.6], size=9)
    add_heading(doc, "Tier calendar, Sep–Dec 2026 (Projections v1.9)", level=3)
    weeks = [x["week"] for x in CAL19["Downtown"]]
    rows = []
    for c in CAFES:
        by = {x["week"]: x for x in CAL19[c]}
        rows.append([c] + [(by[w]["tier"][0].upper() if w in by else "—") for w in weeks])
    hdr = ["Week of"] + [w[5:].replace("-", "/") for w in weeks]
    table(doc, hdr, rows, widths=[0.9] + [0.36] * len(weeks), size=7)
    add_caveat(doc, "S = slow, M = mid, B = busy; weeks scaled to full weeks (Thanksgiving and Christmas weeks are slow everywhere). Projections v1.9 = 2025 same-day sales × (Downtown 1.00, Studio 0.87) for Sep–Dec — the Jul 6–Aug 30 run-rate; 1899 unchanged from v1.8. The band floors sit in the workbook 'Inputs' sheet.")

    # ================= 3–5 CAFES =================
    secno = {"Downtown": 3, "Studio": 4, "1899 Penn": 5}
    intro = {"Downtown": "7a–4p Mon–Fri. Line in at 6:15, Lead at 6:30, two people to 4:30. Third person by 8:00 Tue–Thu; four people at most. Tue/Wed/Thu use one template.",
             "Studio": "7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun. Line in at 6:15 (7:15 weekends) and on the line all day Mon–Thu; Lead at 6:30 (7:30); third person 8:00–2:30 Mon–Thu (7:00 Monday for the delivery); two closers to 2:30 / 5:30. Four people at most on Friday and slow weekends, five on mid/busy weekends.",
             "1899 Penn": "7a–3p Mon–Fri. Two shifts every day — never a solo day. Full day = Lead 6:30–2:30/3:30 with a second person 8:30–3:30; stagger day (sales under ~$650) = Lead 6:30–12:30 and second person 9:30–3:30. Anyone still on at 2:00 closes with the team."}
    for c in CAFES:
        n = secno[c]
        add_heading(doc, f"{n}. {c}", level=1); add_para(doc, intro[c], after=6)
        for tier in ["slow", "mid", "busy"]:
            grid_table(doc, c, tier)
            notes = admin_notes(c, tier)
            if notes: add_para(doc, "Admin shift helps most: " + ". ".join(notes) + ".", after=6)
        add_heading(doc, "Changes from your 10/11 template", level=3)
        for line in FIRST_CUT[c]: add_bullet(doc, line)

    # ================= 6. NMWA =================
    add_heading(doc, "6. NMWA", level=1)
    nm = TIERS["NMWA"]["typical_day"]["slow"]
    add_para(doc, f"One FOH shift 9:00–3:00 on Thu, Fri, Sat, Sun (a full hour to set up and to close). Second Wednesday and first Sunday of the month (community days): two FOH shifts 9:00–3:00. "
                  f"Typical days Thu {m(nm['Thu']['net'])}, Fri {m(nm['Fri']['net'])}, Sat {m(nm['Sat']['net'])}, Sun {m(nm['Sun']['net'])}; community Wednesdays {m(nm['Wed']['net'])}. Labor % is not the lens here — the staffing rule is; the lever is sales. "
                  "Your 10/11 template is missing the 10/14 community-day shift.", after=6)

    # ================= 7. BACKTEST =================
    add_heading(doc, "7. Backtest — the templates over the last twelve months", level=1)
    add_para(doc, "Method: for every open day from September 2025 to August 2026 the week's tier is set by that week's actual walk-in sales (same bands as Section 2), the template for that "
                  "cafe, tier and weekday is costed at today's planning rates, admin and training wages are added at their recent level, and the total is converted to P&L labor with the bridge in Appendix A. "
                  "Sales, COGS and operating expenses are the actuals from the Consolidated Model. So the only thing that changes is labor — the question is simply 'what would the P&L have looked like with these schedules'.", after=6)
    rows = []
    for c in CAFES:
        t = BT["cafes"][c]["totals"]; tw = t["tier_weeks"]
        rows.append([c, m(t["core_sales"]), f"{tw.get('slow',0)} / {tw.get('mid',0)} / {tw.get('busy',0)}", f"{m(t['labor_actual'])}  ({p(t['labor_pct_actual'])})", f"{m(t['labor_template'])}  ({p(t['labor_pct_template'])})", sm(t["labor_saving"]), f"{m(t['noi_reported'])} → {m(t['noi_template'])}"])
    t3 = BT["three_cafes"]["totals"]
    rows.append(["Three cafes", m(t3["core_sales"]), "", f"{m(t3['labor_actual'])}  ({p(t3['labor_actual']/t3['core_sales']*100)})", f"{m(t3['labor_template'])}  ({p(t3['labor_template']/t3['core_sales']*100)})", sm(t3["labor_saving"]), f"{m(t3['noi_reported'])} → {m(t3['noi_template'])}"])
    table(doc, ["Sep 25 – Aug 26", "Core sales", "Weeks S / M / B", "Labor paid (P&L)", "Labor on templates", "Saving", "Cafe NOI: reported → on templates"], rows, widths=[0.95, 0.8, 0.85, 1.1, 1.1, 0.7, 1.4], total_rows=[3], size=8)
    add_caveat(doc, "NOI as reported in the Consolidated Model (catering included), plus the labor saving. The templates are priced at August 2026 rates while 2025 labor was paid at lower rates, so the Sep–Dec 2025 months flatter what was run and understate the templates.")
    findings = {
        "Downtown": "Fall 2025 was already run close to the template (three people most days); the saving is all in June–August 2026, when four to five people were on the floor against slow-tier sales — August alone is $5.1K. On templates Downtown books 22–31% labor from February on (35% in a slow January) and August turns positive.",
        "Studio": "The templates cost about what Studio spends. They save $1–2.3K in slow months (Jan, Jun–Aug 2026) and cost $0.3–2.4K in mid and busy months — most of all against fall 2025, when Studio ran two people through the middle of the day at 24–26% labor. That is the price of the three-person rule; it is not a savings lever.",
        "1899 Penn": "The stagger day is worth $8.9K over the year; from February on 1899 books 31–48% labor on templates against the 33–53% paid. It still loses money in slow months — two people on a $300–450 day cannot be scheduled to profit — so 1899 is a sales problem with a schedule floor, not a schedule problem.",
    }
    for c in CAFES:
        add_heading(doc, c, level=3)
        rows = []
        for x in BT["cafes"][c]["months"]:
            tw = x["tier_weeks"]; rows.append([x["month"], f"{tw.get('slow',0)} / {tw.get('mid',0)} / {tw.get('busy',0)}", m(x["core_sales"]), p(x["labor_pct_actual"]), p(x["labor_pct_template"]), sm(x["labor_saving"]), m(x["noi_reported"]), m(x["noi_template"])])
        t = BT["cafes"][c]["totals"]; rows.append(["12 months", "", m(t["core_sales"]), p(t["labor_pct_actual"]), p(t["labor_pct_template"]), sm(t["labor_saving"]), m(t["noi_reported"]), m(t["noi_template"])])
        table(doc, ["Month", "Weeks S / M / B", "Core sales", "Labor % paid", "Labor % on templates", "Saving", "NOI reported", "NOI on templates"], rows, widths=[0.7, 0.9, 0.85, 0.8, 0.95, 0.75, 0.85, 0.9], total_rows=[12], size=8)
        add_para(doc, findings[c], after=6)
    add_heading(doc, "With COGS at 27–30%", level=3)
    add_para(doc, f"The three cafes ran {p(AF['cogs']['blended_2026_jan_aug_pct'])} blended core COGS in Jan–Aug 2026 (Downtown {p(AF['cogs']['by_cafe_2026']['Downtown'])}, Studio {p(AF['cogs']['by_cafe_2026']['Studio'])}, 1899 {p(AF['cogs']['by_cafe_2026']['1899'])}). "
                  f"On {m(AF['cogs']['core_sales_12mo'])} of core sales every point of COGS is {m(AF['cogs']['per_point_year'])} a year. The table re-runs the twelve months on a core basis (catering excluded on both sides) with COGS at each scenario, before the support entities.", after=6)
    cb = t3["core_basis"]; rows = []
    for k, lab in [("actual_cogs", "Actual COGS"), ("30", "30%"), ("28.5", "28.5%"), ("27", "27%")]:
        rows.append([lab] + [m(BT["cafes"][c]["totals"]["core_basis"][k]["template_labor"]) for c in CAFES] + [m(cb[k]["actual_labor"]), m(cb[k]["template_labor"])])
    table(doc, ["COGS scenario", "Downtown", "Studio", "1899", "Three cafes — labor as paid", "Three cafes — on templates"], rows, widths=[1.0, 0.9, 0.9, 0.9, 1.5, 1.5], size=8)
    add_caveat(doc, f"Core-basis profit = core sales − COGS − labor − operating expenses, twelve months. Support entities (Commissary, SRP, Cafe Mgmt, Central) net to about {m(-BT['three_cafes']['support_burden_2026_avg_month'])} a month in 2026, which is what the cafe profit has to carry; "
                    f"the 'what we can afford' ladder for pay and incentive is in the companion Pay & Incentive Review.")
    add_bullet(doc, f"Templates alone: about {m(sv['annualized_2026'])} a year at the 2026 pace. COGS from 33% to 30%: {m(AF['cogs']['gain_vs_2026_blend']['30'])}. To 28.5%: {m(AF['cogs']['gain_vs_2026_blend']['28.5'])}. The schedule fixes August-type months; COGS decides whether the year is comfortably profitable.", bold_lead="What matters most.")

    # ================= APPENDIX =================
    add_heading(doc, "Appendix A — the numbers behind the summary", level=1)
    rows = []
    for c in CAFES:
        a = AUG[c]; rows.append([c, m(a["core_sales_model"]), h(a["core_hours"], 0), m(a["splh_core_walkin"], 0), f"${a['gross_rate_core']:.2f}", p(a["core_wage_pct_of_core_sales"]), p(a["target_pct"], 0), m(a["headroom_core_wages_vs_target"]), p(a["pl_labor_pct_booked"])])
    table(doc, ["August 2026", "Core sales", "Core hrs", "Sales / hr", "Gross $/hr", "Core wage %", "Target", "Over target", "P&L labor %"], rows, widths=[1.0, 0.8, 0.6, 0.65, 0.65, 0.75, 0.55, 0.75, 0.8], size=8)
    add_caveat(doc, "Core wage % = gross guaranteed wages ÷ core (ex-catering) sales — the schedule basis. P&L labor % is lower because credit-card tips (10–12% of sales) fund part of the guaranteed wage: booked labor ≈ 1.12 × gross wages − tips, "
                    f"which ties to QBO within {abs(BR['Downtown']['error_pct'])}% / {abs(BR['Studio']['error_pct'])}% / {abs(BR['1899 Penn']['error_pct'])}% over Apr–Aug. A cafe exactly on its wage target books ~26% labor (32% at 1899).")
    add_heading(doc, "Appendix B — how the templates were built", level=1)
    for line in [
        "Demand: 89,756 walk-in payments Jan–Aug 2026, averaged by cafe × tier × day-of-week into half-hour curves. Tiers = weekly walk-in sales bands (natural breaks in the 2026 weeks).",
        "People needed each half hour = the rule floors above, or transactions ÷ capacity if higher. Capacity = the busiest-hour throughput each team already sustains: 14 transactions per person-hour at Downtown and 1899, 12 at Studio (bigger food tickets).",
        "Each day's shift list is the cheapest set of shifts (6-hour minimum, the shapes in the rules) that covers the need — solved as a small integer program per cafe × tier × day; Downtown Tue–Thu solved together on the busiest of the three curves.",
        "Rates: August 2026 roster averages by role, rounded up a quarter. Labor % = template wages ÷ the tier's typical sales for that day. 'Admin helps most' marks half-hours where the template runs at 10+ transactions per person-hour or where need exceeds the cap.",
        "Not modelled: catering labor, item mix, training hours (excluded from core), individual availability. Templates are the floor for a typical day of that tier; the scheduler adds or swaps shifts for known events.",
        "Backtest (Section 7): 2025 sales and P&L from the Consolidated Model's 2025 columns (recalculated copy; the model itself was not touched), 2026 from the monthly close files; weekly tiers from actual walk-in sales scaled to a normal week; admin + training at the Apr–Aug 2026 monthly average per cafe; CC tips at each cafe's Apr–Aug share of core sales.",
    ]: add_bullet(doc, line)
    _ensure_zoom_percent(doc); doc.save(OUT); print("Saved →", OUT)


FIRST_CUT = {
    "Downtown": [
        "Mon–Fri: the Lead opens at 6:30 and closes at 4:30 (your cut had the Lead arriving at 9 on Monday and leaving at 1 on Friday); the Line person is 6:15–12:15 or 1:15 instead of open-to-close.",
        "Tue/Wed/Thu are one template: Line 6:15–12:15, Lead 6:30–4:30 (2:30 in mid/busy), Closer 8:00–4:30 as the third person by 8:00; mid and busy weeks add a Barista 8:30–4:30 (four people, no fifth). Your cut ran four to five people 9:00–4:30 every day.",
        "Mon and Fri are three-person days: Line, Lead, and a Closer from 10:30 (9:00 in mid/busy). Friday is still 44–60% labor at its sales — the admin day is the relief.",
        "The $22.75–25.25 team members on barista shifts in your cut cost ~$180 a week; put them in the Lead or Closer slot.",
    ],
    "Studio": [
        "Mon–Thu: exactly your shape — Line 6:15–2:30, Lead 6:30–2:30, third person 8:00–2:30 (7:00 on Monday). Your cut had the third person 6:30–1:00 and, on some days, a fourth.",
        "Friday: Line 6:15–2:15, Lead 6:30–2:30, Closer 8:30–5:30, Barista 11:30–5:30 — four people, two closers (your cut: second Lead 10–5:15 and a Barista 7–1).",
        "Saturday/Sunday: Line 7:15–3:15, Lead 7:30–5:30, Barista 8:00–2:00, Closer 10:00/10:30–5:30 — four people (five only on busy Saturdays with a 9:30–5:30 barista). Your Saturday had five: two Leads, Register, Barista, Line.",
    ],
    "1899 Penn": [
        "Full days (Tue–Thu; Mon in busy weeks): Lead 6:30–2:30 and second person 8:30–3:30 — the second person closes; in busy weeks both run 6:30/7:30–3:30. Your cut had both people 6:30–3:30 every day.",
        "Stagger days (Mon and Fri in slow/mid weeks): Lead 6:30–12:30 and second person 9:30–3:30 — solo open and solo close, two people through the rush. Friday no longer copies Thursday.",
        "Nobody's shift ends 30 minutes before close: a second person who is on at 2:00 stays to 3:30.",
    ],
}

if __name__ == "__main__":
    build()
