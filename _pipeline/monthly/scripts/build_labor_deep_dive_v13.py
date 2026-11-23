#!/usr/bin/env python3
"""Build 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.3.docx' — the readable, implementation-first
version: schedule grids per cafe × tier, short rules, tier calendar, pay review, condensed incentive. No prior-version comparisons."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, add_callout, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, p, h, sm

DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.3.docx")
J = lambda n: json.load(open(os.path.join(DD, n)))
T = J("templates_v13.json"); NV = J("need_v13.json"); TIERS = J("tiers.json"); PAY = J("pay_review_v13.json"); INC = J("incentive_v12.json")
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
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Schedule Templates by Demand Tier", 22, True, COLOR_ACCENT, 0), ("Labor Deep Dive  |  September 2026  |  v1.3", 14, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "Prepared September 10, 2026 for Ryan Fleming. Templates are built to the operating rules agreed Sep 5–10 and priced at the roster's current rates. "
                  "Sources: Square transactions Jan–Aug 2026, 856 Square timecards (14 weeks), Consolidated Model v17, QBO P&L, Projections v1.9 (issued with this report). "
                  "The workbook 'Slipstream Schedule Templates - Sep 2026 v1.3.xlsx' has every shift as an editable row, the half-hour demand curves, and the incentive calculator.",
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
    add_bullet(doc, "Slipstream's guaranteed rates sit $2–5 an hour below what a barista at a busy DC peer takes home with tips; Shift Leads are at parity with a peer barista, ALLs slightly above. A steeper curve (SL $22.50–24.50, ALL $26–28) costs ~$2.2K a month for parity floors, ~$4.3K for the full curve. Section 7.", bold_lead="Pay.")
    add_bullet(doc, "Target-based bonuses do not pay under these rules; an on-template bonus does. Section 8 (for your later review).", bold_lead="Incentive.")

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

    # ================= 7. PAY =================
    add_heading(doc, "7. Pay review — where the guaranteed ladder sits against DC peers", level=1)
    add_para(doc, "Peers pay DC minimum (or a little more) plus tips; Slipstream pays a guaranteed rate with tips applied toward it. So the honest comparison is guaranteed rate vs base + tips. Slipstream's own tip pool is "
                  f"${PAY['slipstream_now']['tips']['Downtown']['tips_per_core_hour']:.2f} / ${PAY['slipstream_now']['tips']['Studio']['tips_per_core_hour']:.2f} / ${PAY['slipstream_now']['tips']['1899 Penn']['tips_per_core_hour']:.2f} per core hour (Downtown / Studio / 1899; 10–12% of sales) — "
                  "the same money a peer barista would keep on top of $18.40.", after=6)
    rows = [[x["who"], x["base"], x["tips"], x["all_in"], x["sl"]] for x in PAY["peers"]]
    table(doc, ["Peer", "Base", "Tips / hr", "Barista all-in / hr", "Shift lead"], rows, widths=[1.9, 1.4, 0.8, 1.2, 1.4], size=8)
    add_caveat(doc, "Public postings and aggregators, Sep 2026 (Indeed, Glassdoor, ZipRecruiter); directional. Ryan's rule of thumb — minimum + tips at a busy cafe ≈ $24–25, shift leads +$1–2 — is consistent with Compass and with Slipstream's own tip rate.")
    add_heading(doc, "Slipstream today (August core hours)", level=3)
    rows = [[b["band"], str(b["people"]), f"{b['share_hours']}%", f"${b['avg_rate']:.2f}"] for b in PAY["slipstream_now"]["bands"]]
    table(doc, ["Band", "People", "Share of core hours", "Avg rate"], rows, widths=[2.4, 0.8, 1.4, 1.0], size=9)
    add_caveat(doc, f"Average guaranteed rate ${PAY['slipstream_now']['avg_core_rate_aug']:.2f}. Read against the peer table: a new or certified barista here is $2–5 an hour below a peer barista with tips; a Shift Lead at $20.50–22.50 is at parity with a peer barista, below a peer shift lead; ALLs at $23–25 match a busy-cafe barista and only just exceed a peer shift lead. Your instinct is right.")
    add_heading(doc, "A steeper, more predictable curve", level=3)
    rows = [[x["level"], f"${x['low']:.2f}" + (f" – ${x['high']:.2f}" if x["high"] != x["low"] else ""), x["note"]] for x in PAY["ladder_proposed"]]
    table(doc, ["Level", "Proposed rate", "Why"], rows, widths=[2.0, 1.2, 3.5], size=8)
    s1, sf = PAY["cost_step1"], PAY["cost_of_proposed"]
    add_bullet(doc, f"Step 1 — parity floors now: every Shift Lead to at least $22.50 and every ALL to at least $25.50. Cost on the August roster ≈ {m(s1['month'])} a month ({m(s1['year'])} a year, {s1['pct_of_core_sales_aug']} points of core sales, ${s1['per_core_hour']:.2f} per core hour).", bold_lead="Cost.")
    add_bullet(doc, f"Full curve (certified $20.50, SL $23.50 midpoint, ALL $27): ≈ {m(sf['month'])} a month ({m(sf['year'])} a year, {sf['pct_of_core_sales_aug']} points of core sales, ${sf['per_core_hour']:.2f} per core hour). The templates save $6–10K a month against what was run, so the curve is fundable — but it lands on top of templates that already sit at the ceiling in slow weeks, so sequence it: parity floors with the templates, the full curve once two months run on-template.", bold_lead="")
    add_bullet(doc, "The trainee rate has to track the DC minimum ($18.40 since July 1, 2026, CPI-indexed each July); four sampled timecards were at $18.00. Keep it explicitly time-boxed (60–90 days to Essentials) so 'below peers' is a stage, not a status.", bold_lead="Floor.")
    add_bullet(doc, "The incentive ($0.50–1.50 an hour in on-template months) is what puts a Shift Lead above a peer shift lead; without it parity is the most the base curve buys at a sustainable labor %.", bold_lead="Incentive.")

    # ================= 8. INCENTIVE (condensed) =================
    add_heading(doc, "8. Incentive — summary for later review", level=1)
    add_para(doc, "Constraints: 1899 and NMWA cannot hit profit-style targets for a while; a plan that never pays is worse than none; the monthly piece must stay small so the annual retirement share is real. All options are pooled across cafes, paid as a percentage of each person's core wages (which keeps overtime math clean), with a guardrail (schedules met the floors; sales ≥ 90% of projection).", after=6)
    sim = INC["simulation_2027"]["options"]; pp = INC["per_person_30h_year_p50"]
    NAMES = {"opt1": "1. Beat share vs the 32/30/35 targets (Downtown + Studio pool; everyone shares)", "opt2": "2. As 1, plus fixed amounts when 1899 hits NI ≥ 5% and NMWA hits its staffing rule",
             "opt3": "3. Beat share vs the TEMPLATE (all four cafes)", "opt4": "4. Fixed amount for every cafe-month run on template ($300 / $300 / $150 / $75) plus half of any beat — recommended"}
    rows = [[NAMES[k], f"{sim[k]['months_paying_p50']} of 12", f"{m(sim[k]['p10'])} – {m(sim[k]['p90'])}", m(pp[k])] for k in ["opt1", "opt2", "opt3", "opt4"]]
    table(doc, ["Option", "Months that pay", "Annual cost (range)", "Per 30 h/wk person / yr"], rows, widths=[3.6, 1.0, 1.2, 0.9], size=8)
    add_caveat(doc, "Simulated 2027 on the tier calendar with the v1.2 rule-based templates (2,000 runs). Under the two-open/two-close rules the templates sit above 32/30 in slow and mid weeks, so a target-based bonus pays in about two months a year — the 'never pays' problem. Measuring against the template pays when the team runs the plan. Full mechanics, backtest on Apr–Aug 2026 and comparables are in the workbook 'Incentive' sheet and will be updated once you have commented.")

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
