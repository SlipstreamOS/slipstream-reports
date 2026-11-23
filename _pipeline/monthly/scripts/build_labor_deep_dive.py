#!/usr/bin/env python3
"""Build 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.0.docx'.
Reads the labor-deep-dive JSONs (data/labor_deep_dive/*.json + data/labor_deep_dive_2026-09.json)
and renders an investor-style document with report_style.py (Calibri 10, no cell gridlines).
"""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from report_style import (add_heading as _add_heading, add_para, add_bullet, add_caveat, add_callout, build_investor_table,
                          _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME)

def add_heading(doc, text, level=1):
    p_ = _add_heading(doc, text, level)
    p_.paragraph_format.keep_with_next = True
    return p_

DATA = os.path.join(os.path.dirname(HERE), "data")
DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))          # workspace folder
OUT = os.path.join(WS, "Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.0.docx")
J = lambda n: json.load(open(os.path.join(DD, n)))
S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
T = J("templates.json"); NV = J("need_vs_actual.json"); TIERS = J("tiers.json"); CURVES = J("curves.json")
STAFF = J("staffing.json"); CAP = J("capacity.json"); INC = J("incentive.json"); RATE = J("rate_lever.json")
FIN = json.load(open(os.path.join(DATA, "fin_2026_v17.json")))
AUG = S["august_2026"]; BR = S["pl_bridge_validation"]
CAFES = ["Downtown", "Studio", "1899 Penn"]
DOW5 = ["Mon", "Tue", "Wed", "Thu", "Fri"]; DOW7 = DOW5 + ["Sat", "Sun"]
FK = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}

def m(x, dec=0):
    if x is None: return "—"
    s = f"{abs(x):,.{dec}f}"
    return f"(${s})" if x < 0 else f"${s}"
def p(x, dec=1):
    return "—" if x is None else f"{x:.{dec}f}%"
def h(x, dec=1):
    return "—" if x is None else f"{x:.{dec}f}"
def sd(x):   # signed delta hours
    return "—" if x is None else f"{x:+.1f}"
def sm(x):
    return "—" if x is None else (f"+${x:,.0f}" if x >= 0 else f"−${abs(x):,.0f}")

def shrink(tbl, size=9):
    for row in tbl.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(size)

def fix_layout(tbl, widths):
    """Fixed layout + explicit grid so LibreOffice/Word honour the column widths; repeat header row; keep short tables on one page."""
    from docx.oxml import OxmlElement
    tblPr = tbl._tbl.tblPr
    lay = tblPr.find(qn("w:tblLayout"))
    if lay is None:
        lay = OxmlElement("w:tblLayout"); tblPr.append(lay)
    lay.set(qn("w:type"), "fixed")
    grid = tbl._tbl.find(qn("w:tblGrid"))
    if grid is not None and widths:
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
            gc.set(qn("w:w"), str(int(w * 1440)))
    # header row repeats on page breaks
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    # rows never split; short tables keep together
    n = len(tbl.rows)
    for i, row in enumerate(tbl.rows):
        rp = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); cs.set(qn("w:val"), "true"); rp.append(cs)
        if n <= 16 and i < n - 1:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.keep_with_next = True

def table(doc, headers, rows, widths=None, key_rows=None, total_rows=None, size=10, aligns=None):
    if widths:
        tot = sum(widths)
        if tot > 6.7:                      # page is 8.5in with 0.9in margins
            widths = [w * 6.7 / tot for w in widths]
    tbl = build_investor_table(doc, headers, rows, col_widths=[Inches(w) for w in widths] if widths else None,
                               key_rows=key_rows, total_rows=total_rows, col_aligns=aligns)
    if widths:
        fix_layout(tbl, widths)
    if size != 10:
        shrink(tbl, size)
    return tbl

def week_line(cafe, tier):
    return T["cafes"][cafe][tier]["week"]

def build():
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9)
    sec.top_margin = sec.bottom_margin = Inches(0.85)
    doc.styles["Normal"].font.name = FONT_NAME
    doc.styles["Normal"].font.size = Pt(10)

    # ---------------- masthead
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0),
                                          ("Labor Deep Dive — Schedule Templates by Demand Tier", 22, True, COLOR_ACCENT, 0),
                                          ("September 2026  |  v1.0", 14, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt)
        _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "Prepared September 4, 2026 for Ryan Fleming. One-time analysis. Sources: Square payment-level transactions Jan–Aug 2026 "
                  "(walk-in only), 856 Square cafe timecards across 14 sampled weeks (Jan–Sep 2026, every week of August), Consolidated Model v17 financials, "
                  "QBO P&L by location (Coefficient, Sep 3), Projections v1.8, and the Square Teams first-cut template for the week of 10/11–10/17. "
                  "Conventions match the monthly report: core sales = ex-catering walk-in sales; core labor = Square timecards excluding 'Manager' (admin) and "
                  "'Training'; 1899 carries the same targets and framing as every cafe; NMWA is kept separate.", italic=True, size=9, color=COLOR_MUTED, after=12)

    # ================= EXECUTIVE SUMMARY =================
    add_heading(doc, "Executive Summary", level=1)
    dt, st, pn = AUG["Downtown"], AUG["Studio"], AUG["1899 Penn"]
    add_callout(doc, "Ryan's question: Downtown and 1899 ran high-$40s sales per labor hour Mon–Thu and only Friday looked bad — so why did both lose money while Studio, with two soft days, was very profitable?",
                f"Because high-$40s is not a healthy number here — it is a 40–47% wage bill. At Slipstream's guaranteed rates (${dt['gross_rate_core']:.2f}–"
                f"${pn['gross_rate_core']:.2f} gross per core hour in August) a cafe needs ${pn['splh_needed_for_target']:.0f}–${st['splh_needed_for_target']:.0f} of walk-in sales per core labor hour to land on its "
                f"target; August delivered ${dt['splh_core_walkin']:.0f} at Downtown, ${st['splh_core_walkin']:.0f} at Studio and ${pn['splh_core_walkin']:.0f} at 1899. On the wage basis the targets are written on, "
                f"core labor ran {p(dt['core_wage_pct_of_core_sales'])} of core sales at Downtown (target 32), {p(st['core_wage_pct_of_core_sales'])} at Studio (target 30) and "
                f"{p(pn['core_wage_pct_of_core_sales'])} at 1899 (target 35) — {m(dt['headroom_core_wages_vs_target']+st['headroom_core_wages_vs_target']+pn['headroom_core_wages_vs_target'])} of core wages above target in one month. "
                "Studio was 7 points over its target and still made money because its fixed costs are a smaller share of a bigger sales base and its COGS were 8–9 points lower; "
                "Downtown and 1899 were 12–14 points over on labor AND carried higher COGS (37–38% vs 29%), so the same miss turned negative. The P&L hides how bad the labor miss is: "
                "credit-card tips (12% of sales) fund part of the guaranteed wage, so booked labor ≈ 1.12 × gross wages − tips. Sales per labor hour and the tip offset are the SAME lever — "
                "a cafe that sells more per hour also collects more tips per hour, which is most of why Studio's 'loaded cost per hour' looked like $18 against $22 at Downtown. "
                "Gross wage rates at the three cafes are within $0.70 of each other.", body_size=10, title_size=10)

    add_para(doc, "Two levers, sized. The hours-placement lever is 3–6 times bigger than the cost-per-hour lever, and both are real:", after=4)
    add_bullet(doc, f"Hours placement. Against a throughput-based need curve (14 transactions per person-hour — the busy-hour p90 the teams already sustain — with a 2-person floor "
                    f"where food is produced), the sampled weeks show {h(sum(r['over_hours'] for r in NV['cafes']['Downtown']['tiers']['slow'].values()))} over-covered person-hours a week at Downtown in slow weeks "
                    f"({m(sum(r['over_dollars'] for r in NV['cafes']['Downtown']['tiers']['slow'].values()))}/wk), {h(sum(r['over_hours'] for r in NV['cafes']['Studio']['tiers']['slow'].values()))} at Studio "
                    f"({m(sum(r['over_dollars'] for r in NV['cafes']['Studio']['tiers']['slow'].values()))}/wk) and {h(sum(r['over_hours'] for r in NV['cafes']['1899 Penn']['tiers']['mid'].values()))}–"
                    f"{h(sum(r['over_hours'] for r in NV['cafes']['1899 Penn']['tiers']['busy'].values()))} at 1899. Downtown's core hours barely move with demand — "
                    f"{h(week_line('Downtown','slow')['actual_sampled_hours'],0)} h/wk in slow weeks, {h(week_line('Downtown','mid')['actual_sampled_hours'],0)} in mid, {h(week_line('Downtown','busy')['actual_sampled_hours'],0)} in busy — "
                    "so the ratio is set entirely by sales. The pattern is a 4th person from 8:30a to 3:30p and two Shift Leads overlapping for 6 hours, on days that need 3 people for 90 minutes and 2 the rest of the time.",
               bold_lead="1.")
    add_bullet(doc, f"Cost per hour. Senior people on barista shifts and overlapping leads cost ${RATE['Downtown']['premium_per_hour']:.2f}/core hour at Downtown "
                    f"(${RATE['Downtown']['premium_per_week']:,.0f}/wk), ${RATE['Studio']['premium_per_hour']:.2f} at Studio (${RATE['Studio']['premium_per_week']:,.0f}/wk) and "
                    f"${RATE['1899 Penn']['premium_per_hour']:.2f} at 1899 (${RATE['1899 Penn']['premium_per_week']:,.0f}/wk) against a standard mix of $21 lead / $19 barista — about $1,600 a month "
                    f"company-wide. Manager-rate ($24–25) team members worked {RATE['Downtown']['manager_rate_hours_per_week']:.0f} core hours a week at Downtown and {RATE['Studio']['manager_rate_hours_per_week']:.0f} at Studio, "
                    f"mostly as the second Shift Lead. Shift-lead density is 30% of core hours at Downtown and Studio and 46% at 1899 (structural for a 2-person cafe).",
               bold_lead="2.")

    # headline table
    rows = []
    for c in CAFES:
        a = AUG[c]
        rows.append([c, m(a["core_sales_model"]), h(a["core_hours"], 0), m(a["splh_core_walkin"], 0), f"${a['gross_rate_core']:.2f}", p(a["core_wage_pct_of_core_sales"]),
                     p(a["target_pct"], 0), m(a["headroom_core_wages_vs_target"]), m(a["splh_needed_for_target"], 0), p(a["pl_labor_pct_booked"]), p(a["pl_labor_pct_bridge"])])
    tot_head = sum(AUG[c]["headroom_core_wages_vs_target"] for c in CAFES)
    rows.append(["Open Cafes", m(sum(AUG[c]["core_sales_model"] for c in CAFES)), h(sum(AUG[c]["core_hours"] for c in CAFES), 0), "", "", "", "", m(tot_head), "", "", ""])
    add_heading(doc, "August 2026 on the wage basis — the basis the targets are written on", level=3)
    table(doc, ["Cafe", "Core sales", "Core hrs", "SPLH", "Gross $/hr", "Core wage %", "Target", "Over target $", "SPLH needed", "P&L labor % booked", "P&L % bridge"],
          rows, widths=[1.0, 0.7, 0.5, 0.45, 0.6, 0.6, 0.45, 0.7, 0.55, 0.6, 0.55], total_rows=[3], size=8)
    add_caveat(doc, "SPLH = walk-in net sales ÷ core labor hours (Square timecards, Aug 1–31, missed clock-outs capped). Core wage % = gross guaranteed wages ÷ model core sales. "
                    "'P&L % bridge' = (1.12 × gross wages incl. admin & training − CC tips) ÷ core sales; the booked figure differs month to month because payroll posts on pay dates "
                    "(Downtown's July was booked at 29% against a bridge of 35%, August at 48% against 40%). Over Apr–Aug the bridge ties to QBO within 0.4% (Downtown), 0.6% (Studio) and 2.3% (1899).")

    add_heading(doc, "What the templates deliver", level=3)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            w = week_line(c, tier); tg = S["targets"][c]["target"] * 100
            rows.append([f"{c} — {tier}", m(w["sales_typical"]), h(w["actual_sampled_hours"], 0), p(w["actual_sampled_labor_pct"]), h(w["ryan_first_cut_hours"], 0), p(w["ryan_first_cut_labor_pct"]),
                         h(w["hours"], 0), m(w["wages"]), p(w["labor_pct"]), p(tg, 0), sm(w["wages"] - w["actual_sampled_wages"]), sm(w["wages"] - w["ryan_first_cut_wages"])])
    table(doc, ["Cafe / tier", "Typical wk sales", "Actual hrs", "Actual %", "First-cut hrs", "First-cut %", "Template hrs", "Template $", "Template %", "Target", "Δ$ vs actual", "Δ$ vs first cut"],
          rows, widths=[1.2, 0.62, 0.48, 0.48, 0.52, 0.52, 0.52, 0.55, 0.52, 0.42, 0.6, 0.6], key_rows=[0, 3, 6], size=7)
    add_caveat(doc, "Weekly core wages at standard rates ($21 Shift Lead, $19 barista/production, $19.50 closer). 'Actual' = the sampled weeks of that tier (2–6 per cafe-tier), at the rates actually paid. "
                    "'First cut' = Ryan's 10/11 week at the rates of the people scheduled, core shifts only; it is one schedule, so it is compared to each tier's typical sales. "
                    "The 1899 slow and mid tiers cannot reach 35% with two people on the floor all day — see Section 5 for the lean variants (36% mid, 41% slow).")

    add_para(doc, "The three biggest schedule changes, by monthly dollars (slow-tier weeks, which is what September–October look like on the current run-rate):", after=4, before=6)
    dts, dtm = week_line("Downtown", "slow"), week_line("Downtown", "mid")
    sts, stm = week_line("Studio", "slow"), week_line("Studio", "mid")
    pnm, pnl = week_line("1899 Penn", "mid"), week_line("1899 Penn", "mid_lean")
    add_bullet(doc, f"Downtown — one lead, one closer, no mid-day 4th. Open with a Shift Lead and the Production Line at 6:00–6:30, add a 4-hour rush barista 7:30–11:30 only in mid/busy weeks, and replace the second "
                    f"Shift Lead (8:00–4:30) with a barista-rate Closer (8:30/9:00–4:30). Friday becomes a 2-person day. Slow-tier week: {h(dts['hours'],0)} core hours vs {h(dts['actual_sampled_hours'],0)} actual "
                    f"(−${dts['actual_sampled_wages']-dts['wages']:,.0f}/wk ≈ −${(dts['actual_sampled_wages']-dts['wages'])*52/12:,.0f}/mo); vs the first cut −{dts['ryan_first_cut_hours']-dts['hours']:.0f} h/wk. "
                    f"Mid-tier week: {h(dtm['hours'],0)} h (−${(dtm['actual_sampled_wages']-dtm['wages'])*52/12:,.0f}/mo vs actual).", bold_lead="1.")
    add_bullet(doc, f"Studio Tue–Thu — a 2-person core day (Production Line 6:00–2:00 + Shift Lead 6:30–2:45) with the ALL's Admin block covering the 8–10 rush on the days it is scheduled; the third "
                    f"person returns only in busy weeks (a 4-hour rush shift). Weekend and Friday keep 4–5 heads at the 10a–1p peak. Slow-tier week: {h(sts['hours'],0)} h vs {h(sts['actual_sampled_hours'],0)} actual "
                    f"(−${(sts['actual_sampled_wages']-sts['wages'])*52/12:,.0f}/mo); vs the first cut −{sts['ryan_first_cut_hours']-sts['hours']:.0f} h/wk.", bold_lead="2.")
    add_bullet(doc, f"1899 — the second person leaves after lunch. Shift Lead 6:30–3:30 all week; Production Barista 7:30–2:30 Tue–Thu, 7:30–1:30 Monday, rush-only (8–12) Friday. Mid-tier week: "
                    f"{h(pnm['hours'],0)} h ({p(pnm['labor_pct'])}) vs {h(pnm['actual_sampled_hours'],0)} actual — −${(pnm['actual_sampled_wages']-pnm['wages'])*52/12:,.0f}/mo; the lean variant "
                    f"({h(pnl['hours'],0)} h, {p(pnl['labor_pct'])}) needs the ALL to cover lunch on Friday/Monday. Below ~$3,300/week 1899 is structurally a 1.5-person cafe; the 35% target is only reachable at busy-tier sales or with Admin covering the lunch peak.",
               bold_lead="3.")
    add_bullet(doc, "Also: stop filling barista shifts with $24–25 team members and stop overlapping two Shift Leads (≈ $1,600/mo company-wide); add the second-Wednesday community-day shift at NMWA "
                    "(missing from the 10/11 first cut — 10/14 is the 2nd Wednesday) and shorten NMWA shifts to 9:30–2:30.", bold_lead="4.")

    add_para(doc, "Tiers, in one line each (weekly walk-in sales; full definitions in Section 2): Downtown slow < $7,500, mid $7,500–9,000, busy ≥ $9,000. Studio slow < $10,500, mid $10,500–12,500, "
                  "busy ≥ $12,500. 1899 slow < $3,300, mid $3,300–4,200, busy ≥ $4,200. Ryan's hypothesis holds for Studio (Jan, Feb, Jul, Aug slow; Apr–Jun busy) and for Downtown's August; "
                  "Downtown's February was mid, not slow, and 1899's August was a mid month by sales — its August loss was labor placement and COGS, not a demand collapse. "
                  "The projection workbook puts October in the busy tier for Downtown and Studio, but Jul–Aug ran 4% (Downtown) and 14% (Studio) below 2025; at that run-rate September–November are mid-to-slow. "
                  "Recommendation: build the 10/11 week on the MID templates and step up only when the trailing two weeks clear the band floor.", before=6, after=6)
    add_para(doc, "Incentive (Section 8): the recommended structure is a monthly 'Beat Share' — 50% of core wages saved below the target line, capped at 1.5% of core sales, pooled across the open cafes "
                  "and paid per core hour worked, with the template floors as the service guardrail. At a 2-point beat in a mid month it pays ~$1,000 (1.0% of sales, ~$0.60 per core hour, ~$80 for a 30-hour week) "
                  "and Slipstream keeps the other $1,000; at the cap it is ~$1,500. That size is self-funding and leaves the 5–10% annual profit share ($11–22K a year at target profitability on August sales) intact.",
             after=8)

    # ================= 1. BRIDGE =================
    add_heading(doc, "1. How labor reaches the P&L — the bridge between the schedule and the statement", level=1)
    add_para(doc, "Slipstream pays guaranteed rates; credit-card tips are applied toward that rate and the company tops up the difference (ADP: $12 tipped base + CREDTIPP + OTH top-up). "
                  "QBO's FOH Labor line is therefore gross wages NET of tips, and Payroll Tax is allocated at 12% of gross. That gives a simple, testable bridge: "
                  "P&L labor ≈ 1.12 × gross wages (core + admin + training) − credit-card tips. The table tests it against QBO for Apr–Aug.", after=6)
    rows = []
    for c in CAFES:
        b = BR[c]
        for mon in b["months"]:
            r = b["by_month"][mon]
            rows.append([f"{c} {mon}", p(r["core_wage_pct"]), m(r["pl_est"]), m(r["pl_actual"]), sm(r["pl_est"] - r["pl_actual"]), m(r["headroom"])])
        rows.append([f"{c} Apr–Aug", "", m(b["pl_est_sum"]), m(b["pl_actual_sum"]), f"{b['error_pct']:+.1f}%", m(sum(x["headroom"] for x in b["by_month"].values()))])
    table(doc, ["Cafe / month", "Core wage % (gross)", "Bridge estimate", "QBO booked", "Δ", "Core wages over target"], rows,
          widths=[1.3, 1.0, 1.0, 1.0, 0.8, 1.2], total_rows=[5, 11, 17], size=9)
    add_caveat(doc, "Months are lumpy because payroll posts on pay dates with a single-line accrual; the two-month pairs and the five-month totals tie. Core wage % for Apr and Aug is scaled from partial "
                    "timecard coverage (Apr 12 onward; Aug 1–29 in the weekly store) to the full month. 'Over target' = gross core wages − target % × model core sales.")
    add_heading(doc, "What the bridge means for the targets", level=3)
    add_bullet(doc, "A cafe exactly on its core target (30/32/35 wage basis) with ~2% admin and ~1% training shows about 25–27% total labor on the P&L, not 33–36%: "
                    "1.12 × (30 + 3) − 12 = 25.0 at Studio; 1.12 × (32 + 2) − 12 = 26.1 at Downtown; 1.12 × (35 + 3) − 10.4 = 32.2 at 1899 (its tip rate is lower). "
                    "With COGS at 29–30%, prime cost lands in the 55–60% healthy band at Studio and Downtown — that is the prize.")
    add_bullet(doc, f"Loaded cost per hour reconciles the same way. August booked ${dt['loaded_rate_booked_per_hour']:.2f}/hr at Downtown, ${st['loaded_rate_booked_per_hour']:.2f} at Studio and "
                    f"${pn['loaded_rate_booked_per_hour']:.2f} at 1899; the bridge gives ${dt['loaded_rate_bridge_per_hour']:.2f}, ${st['loaded_rate_bridge_per_hour']:.2f} and ${pn['loaded_rate_bridge_per_hour']:.2f}. "
                    f"Gross rates are ${dt['gross_rate_core']:.2f} / ${st['gross_rate_core']:.2f} / ${pn['gross_rate_core']:.2f}; tips per core hour are ${dt['tips_per_core_hour']:.2f} / ${st['tips_per_core_hour']:.2f} / "
                    f"${pn['tips_per_core_hour']:.2f}. Studio's lower loaded rate is its higher tips per hour (more sales per hour) plus the accrual timing that made Downtown's August heavy; it is not lower pay.")
    add_bullet(doc, "Implication for the weekly check: keep scoring schedules on gross wages ÷ projected core sales (the 30/32/35 basis). Do not compare that number to the monthly P&L labor % — "
                    "they are different bases and the P&L one is ~7–9 points lower at the same staffing. The monthly report should show both, with the bridge.")

    # ================= 2. TIERS =================
    add_heading(doc, "2. Demand tiers", level=1)
    add_para(doc, "Tiers are defined on weekly walk-in sales per cafe (Mon–Sun; holiday-shortened weeks scaled to a full week; one weather week — Downtown Jan 26, $2,365 — excluded). "
                  "Each of the 35 weeks of 2026 is assigned to a band by what it actually sold, so every tier's demand curve is built from days that really were that busy. "
                  "Bands sit at the natural breaks between the month clusters; they differ by cafe because the cafes are different sizes. The template for a week is chosen by "
                  "the projected weekly sales in the weekly report's forward block, not by the calendar.", after=6)
    rows = []
    for c in CAFES:
        b = S["tiers"][c]["bands_weekly_walkin"]
        rows.append([c, f"< {m(b['slow'][1])}", f"{m(b['mid'][0])} – {m(b['mid'][1])}", f"≥ {m(b['busy'][0])}"])
    table(doc, ["Cafe", "Slow", "Mid", "Busy"], rows, widths=[1.4, 1.6, 1.8, 1.6])
    add_heading(doc, "Month membership — 2026 actuals (Jan–Aug), projection (Sep–Dec)", level=3)
    MONS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    rows = []
    for c in CAFES:
        cal = S["tiers"][c]["calendar"]
        row = [c]
        for mon in MONS:
            r = cal.get(mon)
            if not r: row.append("—"); continue
            t = r["tier"]
            if "weeks_by_tier" in r:
                w = r["weeks_by_tier"]; mix = "/".join(str(w[k]) for k in ["slow", "mid", "busy"])
                row.append(f"{t} ({mix})")
            else:
                extra = f"; run-rate {r['yoy_view_tier']}" if "yoy_view_tier" in r else ""
                if mon == "Dec": extra += "; Ryan: slow"
                row.append(f"{t}{extra}")
        rows.append(row)
        rows.append(["  avg weekly $"] + [m(cal[mon]["avg_weekly"]) if mon in cal else "—" for mon in MONS])
    table(doc, ["Cafe"] + MONS, rows, widths=[0.9] + [0.5] * 12, key_rows=[0, 2, 4], size=7)
    add_caveat(doc, "Jan–Aug cells show the dominant tier and the count of weeks by slow/mid/busy. Sep–Dec use Projections v1.8 (identical to v1.7 for these tabs); 'run-rate' applies the Jul–Aug 2026 year-over-year "
                    "change (Downtown −4%, Studio −14%) to the 2025 actuals in the projections workbook — 1899 has no comparable 2025 base. December is projected mid at Downtown and 1899, but Ryan's instruction "
                    "and the 2025 pattern (Dec 2025 = 83% of Oct at Downtown, 71% at Studio) say plan the holiday weeks (Dec 21–Jan 3) on the slow template.")
    add_heading(doc, "Typical day by tier (walk-in $ / transactions per open day)", level=3)
    for c in CAFES:
        dows = DOW7 if c == "Studio" else DOW5
        rows = []
        for tier in ["slow", "mid", "busy"]:
            td = S["tiers"][c]["typical_day"][tier]
            rows.append([f"{c} — {tier}"] + [f"{m(td[d]['net'])} / {td[d]['txns']:.0f}" if d in td else "—" for d in dows] + [m(sum(td[d]["net"] for d in dows if d in td))])
        table(doc, ["Cafe / tier"] + dows + ["Week"], rows, widths=[1.2] + [0.72] * len(dows) + [0.7], size=8)
    add_caveat(doc, "Ryan's hypothesis tested: Studio — Jan, Feb, Jul and Aug are slow (Feb 3 of 5 weeks), Apr–Jun busy: confirmed, with July added. Downtown — August is slow but February 2026 was mid (+23% on January's "
                    "per-day sales) and January itself was mixed; 2025 shows the same Aug/Dec/Jan trough. 1899 — August was a mid month ($3,763/wk vs $4,129 in July); its Jan–Feb 'slow' weeks were the ramp of a cafe that is still growing "
                    "~90% year-over-year, so the slow tier may not recur at that level outside the holiday weeks.")

    # ================= per-cafe sections =================
    cafe_sections(doc)

    # ================= NMWA =================
    add_heading(doc, "6. NMWA (for completeness)", level=1)
    td = S["tiers"]["1899 Penn"]  # placeholder to keep structure; NMWA data from TIERS
    nm = TIERS["NMWA"]["typical_day"]["slow"]
    add_para(doc, f"NMWA is a 10a–2p museum cafe with a 1-person staffing rule (2 on community days). Typical days: Thu {m(nm['Thu']['net'])}, Fri {m(nm['Fri']['net'])}, Sat {m(nm['Sat']['net'])}, "
                  f"Sun {m(nm['Sun']['net'])}; community Wednesdays {m(nm['Wed']['net'])} (4 observed). The sampled shifts run 9:00–3:00 (6 hours for a 4-hour window) and three of the five "
                  "16-hour-plus missed clock-outs in the whole sample were NMWA FOH punches in the last week of August. Template: FOH 9:30–2:30 Thu–Sun (5 h), plus a second FOH 10:00–2:00 on the "
                  "first Sunday and second Wednesday. Ryan's 10/11 first cut has no Wednesday 10/14 shift — that is the October community day — and the shifts are 9–3; the change is +1 community "
                  "shift and −1 hour on each regular shift (net ≈ +$100/wk on a $1,000 week). Labor % is not a useful lens here: at 5 hours a day the wage bill is 56% of typical sales; the lever is sales.", after=6)

    # ================= 7. SENSITIVITY =================
    add_heading(doc, "7. Sensitivity and tier triggers", level=1)
    add_para(doc, "August's failure mode was a template built for one tier meeting sales from a lower one. Each template is scored at 10% under its tier's typical sales, and a trigger is given: the weekly "
                  "projected (or trailing two-week actual) sales below which the template breaches the ceiling and the next-slower template should be used. Because hours do not scale with sales inside a tier, "
                  "a 10% sales miss costs 3–4 points of labor at every cafe.", after=6)
    rows = []
    for c in CAFES:
        tg = S["targets"][c]
        for tier in [t for t in T["cafes"][c] if t in ("slow", "slow_lean", "mid", "mid_lean", "busy")]:
            w = week_line(c, tier)
            rows.append([f"{c} — {tier.replace('_', ' ')}", m(w["sales_typical"]), m(w["wages"]), p(w["labor_pct"]), p(w["labor_pct_minus10"]), p(tg["ceiling"] * 100, 0), m(w.get("ceiling_trigger_sales")), m(w.get("target_trigger_sales"))])
    table(doc, ["Template", "Typical wk sales", "Template wages", "Labor %", "At −10% sales", "Ceiling", "Ceiling trigger (wk sales)", "On-target at (wk sales)"], rows,
          widths=[1.45, 0.85, 0.85, 0.65, 0.8, 0.6, 1.0, 1.0], size=9)
    add_caveat(doc, "Reading the triggers: run the Downtown mid template only while projected weekly sales stay above ~$7,300; below that it breaches 36% and the slow template applies. The Studio mid template holds to ~$10,200, "
                    "the busy template to ~$11,200. 1899's two-person mid template breaches 40% below ~$3,800 — i.e. in most mid weeks — which is why the lean variant exists. Weekly report: add a line per cafe showing "
                    "'tier in force' and the trigger distance.")

    # ================= 8. INCENTIVE =================
    incentive_section(doc)

    # ================= 9. DATA NOTES =================
    add_heading(doc, "9. Data notes and method", level=1)
    dn = S["data_notes"]
    add_bullet(doc, f"Timecards: {dn['timecards_cafes']} cafe timecards ({dn['timecards_harvested']} harvested incl. 5 commissary/roasting cards in the stored Aug 23–29 pull) via labor.searchTimecards for 14 weeks — Jan 5, Feb 2, Feb 23, Mar 9, Apr 13, May 11, Jun 8, Jul 13, Jul 27, and every week Aug 3–Sep 6 "
                    f"(Sep partial) — harvested from the session transcript, not re-typed. Every tier has at least 3 sampled weeks per cafe. {dn['dropped_zero_length']} zero-length punches dropped; "
                    f"{dn['zero_rate_imputed']} zero-rate cards given the team member's usual rate.")
    capd = dn["capped_detail"]
    add_bullet(doc, f"Missed clock-outs: {dn['capped_ruleA_16h']} punches ≥16 h were capped to the median shift for that cafe and role (task rule) — NMWA FOH ×3 (Aug 27–29), 1899 Production Barista (Aug 31), Studio Production Barista (Aug 31); "
                    f"a second rule capped {dn['capped_ruleB_past_close']} punches that ran more than 2¼ hours past close (Studio Aug 24/25/27 at 15.3, 15.2 and 10.1 h; 1899 Apr 17 at 10.9 h). Phantom wages removed: {m(dn['capped_wages_avoided'])}. "
                    "The Studio Monday/Tuesday cases (a 6:30a punch closed at 9:47p on a day the cafe shuts at 2p) would have passed a 16-hour test — recommend the weekly pipeline adopt the past-close rule too.")
    add_bullet(doc, f"Transactions: {dn['transactions_rows_kept']:,} walk-in payments Jan–Aug 2026 (Complete only, net > 0, invoice-sourced and Management/Equipment Fee rows excluded). Hour of day from the Eastern 'Time' field. "
                    "Curves are per open day, averaged within cafe × tier × day-of-week (5–17 days per cell).")
    add_bullet(doc, f"Capacity: transactions per person-hour in the busiest fifth of observed hours — median {CAP['Downtown']['busy_hours']['tph_p50']} / {CAP['Studio']['busy_hours']['tph_p50']} / {CAP['1899 Penn']['busy_hours']['tph_p50']} "
                    f"(Downtown / Studio / 1899), 90th percentile {CAP['Downtown']['busy_hours']['tph_p90']} / {CAP['Studio']['busy_hours']['tph_p90']} / {CAP['1899 Penn']['busy_hours']['tph_p90']}, maxima 21 / 18.7 / 21. "
                    "Planning capacity 14 per person-hour. Adding a 4th person at Downtown or Studio did not raise busy-hour throughput (median 38 transactions with 3 people vs 37–39 with 4–5), which is the statistical signature of over-coverage.")
    add_bullet(doc, "Need rule: need(h) = max(floor, ceil(transactions per hour ÷ 14)). Floors: 2 from prep through lunch where food is produced (Downtown, Studio), 1 in the final 90 minutes; Studio Monday 3 at open; "
                    "1899 2 through 1:30p then 1; NMWA 1 (2 on community days). Prep starts 60 min before open (30 at 1899); close runs 30–45 min after.")
    add_bullet(doc, "Standard rates in the templates are inputs (blue in the workbook): $21 Shift Lead, $19 barista/production/register, $19.50 closer. Observed cafe averages are $20.20–21.05 because senior people fill core shifts; "
                    "that gap is reported separately as the cost-per-hour lever rather than baked into the templates.")
    add_bullet(doc, f"Projections: {dn['projections_file']}. Financials: Consolidated Model v17 extract (fin_2026_v17.json) and QBO P&L by location (Coefficient Output 2026.xlsx, Sep 3). "
                    "Neither the Consolidated Model nor the monthly report was modified. Reference JSON for the monthly report: _pipeline/monthly/data/labor_deep_dive_2026-09.json.")
    add_bullet(doc, "Limits: templates are built on tier-average curves; individual days vary ±25% around them, which is what the Admin flex and the whole-person rounding absorb. Role names follow Square job titles; pay is per person, "
                    "not per role, so the rate lever depends on who is scheduled, not on renaming shifts. No item-level or catering labor is modelled. Studio's Monday catering delivery and 1899's owner-relations context are taken as given.")

    appendix(doc)
    _ensure_zoom_percent(doc)
    doc.save(OUT)
    print("Saved →", OUT)


def cafe_sections(doc):
    secno = {"Downtown": 3, "Studio": 4, "1899 Penn": 5}
    HOURS_SHOW = {"Downtown": list(range(7, 16)), "Studio": list(range(7, 17)), "1899 Penn": list(range(7, 15))}
    DOWS = {"Downtown": DOW5, "Studio": DOW7, "1899 Penn": DOW5}
    intro = {
        "Downtown": "An office-district cafe (7a–4p Mon–Fri): the 8:30–10:00 rush is a third of the day, Monday runs 60–70% of a mid-week day and Friday 40–50%. Two Shift Leads overlap for six hours on most days.",
        "Studio": "A neighborhood cafe (7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun) with a real food program: the weekend is the business (40–45% of the week's sales), weekday peaks are 10–12p rather than the rush, "
                  "and Tuesday–Thursday are $900–1,350 days staffed with three people from 6:00 to 2:30.",
        "1899 Penn": "A two-person office-building cafe (7a–3p Mon–Fri) with a bimodal day — a 9am rush and a 1pm lunch peak of equal size — and a Friday that is half a Wednesday. Both people work 6:30–3:30 every day.",
    }
    for c in CAFES:
        n = secno[c]
        add_heading(doc, f"{n}. {c}", level=1)
        add_para(doc, intro[c], after=6)
        # ---- a. demand curve
        add_heading(doc, f"{n}a. Demand curve — walk-in transactions per hour, by tier", level=2)
        hrs = HOURS_SHOW[c]
        show_days = ["Tue", "Fri"] if c != "Studio" else ["Tue", "Fri", "Sat"]
        rows = []
        for d in show_days:
            for tier in ["slow", "mid", "busy"]:
                cv = CURVES[c]["by_tier"][tier][d]
                g = lambda hh: cv['hourly'].get(str(hh), {}).get('txns', 0)
                rows.append([f"{d} — {tier}"] + [(f"{g(hh):.0f}" if str(hh) in cv['hourly'] else "·") for hh in hrs] + [f"{sum(g(hh) for hh in hrs):.0f}"])
        table(doc, ["Day / tier"] + [f"{hh if hh <= 12 else hh - 12}{'a' if hh < 12 else 'p'}" for hh in hrs] + ["Day"], rows,
              widths=[1.1] + [0.42] * len(hrs) + [0.5], key_rows=[0, 3, 6] if c != "Studio" else [0, 3, 6], size=8)
        add_caveat(doc, "Average transactions in the hour beginning at the time shown, per open day of that tier and weekday (Tuesday and Friday shown; Studio adds Saturday). Full 30-minute curves for every day are in the workbook.")
        # ---- b. actual vs needed
        add_heading(doc, f"{n}b. Actual staffing vs need — hours and dollars per week", level=2)
        rows = []
        for tier in ["slow", "mid", "busy"]:
            dows = NV["cafes"][c]["tiers"][tier]
            for d in DOWS[c]:
                r = dows.get(d)
                if not r: continue
                rows.append([f"{tier} {d}", h(r["need_hours"]), h(r["actual_hours"]), h(r["over_hours"]), m(r["over_dollars"]), h(r["under_hours"]), f"{r['peak_need']} / {r['peak_actual']:.1f}"])
            rows.append([f"{tier} — week", h(sum(x["need_hours"] for x in dows.values())), h(sum(x["actual_hours"] for x in dows.values())), h(sum(x["over_hours"] for x in dows.values())),
                         m(sum(x["over_dollars"] for x in dows.values())), h(sum(x["under_hours"] for x in dows.values())), ""])
        nd = len(DOWS[c])
        table(doc, ["Tier / day", "Need (h)", "Actual (h)", "Over (h)", "Over ($/day)", "Under (h)", "Peak need / actual"], rows,
              widths=[1.2, 0.8, 0.85, 0.8, 0.9, 0.8, 1.2], total_rows=[nd, 2 * nd + 1, 3 * nd + 2], size=9)
        rate = NV["cafes"][c]["avg_core_rate"]
        add_caveat(doc, f"Need = the throughput rule in Section 9 applied to the tier's 30-minute curve; actual = average core headcount on the clock in the sampled weeks. Dollars at the cafe's observed average core rate (${rate:.2f}). "
                        "'Over' is paid time above need; 'under' is need above staffing (mostly prep half-hours and single-person closes).")
        # ---- c. levers
        add_heading(doc, f"{n}c. The two levers, quantified", level=2)
        rl = RATE[c]
        tot_over = {t: sum(x["over_dollars"] for x in NV["cafes"][c]["tiers"][t].values()) for t in ["slow", "mid", "busy"]}
        add_bullet(doc, f"Hours placement: {m(tot_over['slow'])} / {m(tot_over['mid'])} / {m(tot_over['busy'])} per week of paid time above need in slow / mid / busy weeks "
                        f"(≈ {m(tot_over['slow']*52/12)} / {m(tot_over['mid']*52/12)} / {m(tot_over['busy']*52/12)} a month). Against the target line the monthly headroom is "
                        f"{m(INC['headroom_by_tier'][c]['slow']['headroom_actual_vs_target'])} / {m(INC['headroom_by_tier'][c]['mid']['headroom_actual_vs_target'])} / "
                        f"{m(INC['headroom_by_tier'][c]['busy']['headroom_actual_vs_target'])}.", bold_lead="Lever 1 —")
        add_bullet(doc, f"Cost per hour: observed average core rate ${rl['avg_rate']:.2f} vs ${rl['std_mix_rate']:.2f} at standard rates for the same role mix — a premium of ${rl['premium_per_hour']:.2f}/hr, "
                        f"{m(rl['premium_per_week'])}/wk, ≈ {m(rl['premium_per_week']*52/12)}/mo. Manager-rate ($24–25) team members worked {rl['manager_rate_hours_per_week']:.1f} core hours a week here "
                        f"({rl['manager_rate_share']:.0f}% of core hours); Shift Lead hours are {rl['sl_share_hours']:.0f}% of core and two leads were on the floor together {rl['double_SL_hours_per_week']:.1f} hours a week. "
                        f"Hours by rate band: under $19 {rl['rate_bands_share']['0-19']:.0f}%, $19–20.49 {rl['rate_bands_share']['19-20.5']:.0f}%, $20.50–22.49 {rl['rate_bands_share']['20.5-22.5']:.0f}%, $22.50+ {rl['rate_bands_share']['22.5-30']:.0f}%.",
                   bold_lead="Lever 2 —")
        if c == "Downtown":
            add_para(doc, "What the shape says: Downtown needs 3 people from 8:30 to 10:00 and 2 for the rest of the day in slow and mid weeks (4 at the 9am peak in busy weeks), yet it ran 4 people from 9:00 to 3:30 "
                          "in every tier. The 4th person costs ~$130/day and adds no measurable throughput. Monday and Friday are 2-person days at any tier. The last hour (3–4p) is 3–12 transactions; "
                          "Ryan's caution on judging it alone stands — the templates keep the 4p close and let the Closer run it alone.", after=6)
        if c == "Studio":
            add_para(doc, "What the shape says: Tuesday–Thursday need 2 people (a third for 30–60 minutes at 8:30–9:30 in mid/busy weeks); the cafe ran 3 all day. Friday and the weekend need 3–4 at the 10a–1p peak and "
                          "were staffed at 3.8–5.0; by throughput the weekend is adequately covered, so the monthly report's 'under-staffed weekend' is a share-of-hours artifact, not a service gap. Monday's third person at open "
                          "is the catering delivery; where the ALL's Admin block is on the floor at 7:00 (as in the first cut) it counts toward the three.", after=6)
        if c == "1899 Penn":
            add_para(doc, "What the shape says: two people 6:30–3:30 is 18 hours a day; the day needs 2 from 7:00 to 1:30 and 1 after that. At $750–950 a day (mid tier) the 35% target buys 13–16 hours, so the "
                          "second person must leave after lunch and Friday/Monday must be 1.5-person days. 1899's lever is smaller in dollars than Downtown's or Studio's, but it is the difference between the incentive fee "
                          "and a loss for the owner.", after=6)
        # ---- d. templates
        add_heading(doc, f"{n}d. Schedule templates", level=2)
        tg = S["targets"][c]
        add_para(doc, f"Targets: {tg['target']*100:.0f}% core wages of core sales (ceiling {tg['ceiling']*100:.0f}%). Standard rates $21 Shift Lead, $19 barista/production, $19.50 closer. "
                      "Every template is checked half-hour by half-hour against the tier's need curve; 'Admin covers' notes are the deliberate flex — staff 2 core where a day needs 2.5 and let the ALL's block absorb the rush.", after=4)
        for tier in [t for t in T["cafes"][c] if t.split("_")[0] in ("slow", "mid", "busy")]:
            o = T["cafes"][c][tier]; w = o["week"]
            label = tier.replace("_lean", " — LEAN variant (target)").replace("slow", "Slow").replace("mid", "Mid").replace("busy", "Busy")
            add_heading(doc, f"{label} template — {m(w['sales_typical'])}/wk typical, {w['hours']:.1f} core hours, {m(w['wages'])} = {p(w['labor_pct'])} (SPLH {m(w['splh'],0)})", level=3)
            rows = []
            for d, dd in o["days"].items():
                for i, s in enumerate(dd["shifts"]):
                    rows.append([d if i == 0 else "", s["role"], s["start"], s["end"], h(s["hours"], 2), f"${s['rate']:.2f}", m(s["wages"]), s["note"]])
                tmpl = dd["template"]; act = dd.get("actual_sampled", {})
                rows.append([f"{d} total", f"{len(dd['shifts'])} shifts", "", "", h(tmpl["hours"], 2), "", m(tmpl["wages"]),
                             f"{p(tmpl['labor_pct'])} of {m(dd['sales_typical'])}; peak {tmpl['peak_staff']}; actual {act.get('core_hours', '—')} h / {p(act.get('labor_pct'))}"])
            total_idx = [i for i, r in enumerate(rows) if str(r[0]).endswith("total")]
            table(doc, ["Day", "Role", "Start", "End", "Hours", "Rate", "Wages", "Notes / check"], rows, widths=[0.6, 1.45, 0.45, 0.45, 0.45, 0.5, 0.55, 2.25], total_rows=total_idx, size=8)
            add_caveat(doc, f"Week: {w['hours']:.1f} h, {m(w['wages'])} at standard rates = {p(w['labor_pct'])} of {m(w['sales_typical'])} (target {p(tg['target']*100,0)}). At −10% sales: {p(w['labor_pct_minus10'])}. "
                            f"Sampled actual weeks of this tier: {w['actual_sampled_hours']:.1f} h, {m(w['actual_sampled_wages'])} = {p(w['actual_sampled_labor_pct'])}; Ryan's first cut: {w['ryan_first_cut_hours']:.1f} h, {m(w['ryan_first_cut_wages'])} = {p(w['ryan_first_cut_labor_pct'])} at this tier's sales. "
                            + (f"Drop to the next-slower template when projected weekly sales fall below {m(w.get('ceiling_trigger_sales'))}." if not tier.startswith("slow")
                               else f"This template breaches the {p(tg['ceiling']*100,0)} ceiling when weekly sales fall below {m(w.get('ceiling_trigger_sales'))}; there is no slower template — below that, the day-level floors (and Admin cover) are the only lever."))
        # ---- e. changes vs first cut
        add_heading(doc, f"{n}e. Specific changes from the 10/11 first-cut template", level=2)
        for line in FIRST_CUT_CHANGES[c]:
            add_bullet(doc, line)


FIRST_CUT_CHANGES = {
    "Downtown": [
        "Monday (first cut: Production Barista 6:30–1, Production Line 6:12–4:30, Shift Lead 9–4:30 = 24.3 h). Open with the Shift Lead (6:30) instead of a 9:00 lead; Production Line 6:00–2:00 (not a 10.3-hour open-to-close); "
        "Closer 9:00–4:30. Mid: 24 h (same hours, better shape and no OT exposure). Slow: 18 h — Production Line 6–2 plus one open-to-close lead 6:30–4:30, the 10-hour option Ryan flagged for discussion.",
        "Tuesday (first cut: 4 people — Production Line 6:15–2, Production Barista 7:30–4:30, Shift Lead 6:30–3, Production Barista 9–4:30 = 32.75 h). This is the 'Tuesday runs 4-deep' cut from the memo: the 7:30–4:30 "
        "Production Barista becomes a 4-hour rush barista (7:30–11:30) and the 9–4:30 Production Barista becomes the Closer at 8:30–4:30. Mid: 28.5 h (−4.25). Slow: 24 h (drop the rush barista; Closer 8:30–4:30 is the 3rd through the rush).",
        "Wednesday and Thursday (first cut: Shift Lead 8–4:30, Production Barista/Barista 6:30–1:30, Production Line 6:15–3, Barista/Production Barista 9–4:30 = 31.75 h). Replace the 8–4:30 second Shift Lead with a barista-rate "
        "Closer 8:30–4:30 (the SL premium is $2/hr); the 6:30–1:30 shift becomes a 4–4.5 h rush barista (7:30–12:00); Production Line 6:00–2:00 not 6:15–3:00; Shift Lead opens 6:30–3:30. Mid: 29 h (−2.75/day). Slow: 24.5 h (no rush barista).",
        "Friday (first cut: Production Barista 9:30–4:30, Shift Lead 6:30–1, Production Line 6:15–4:30 = 23.75 h). Friday is 40–50% of a Wednesday: mid template Production Line 6–1, Shift Lead 6:30–3, Closer 9–4:30 = 23 h; slow template "
        "Production Line 6–12 + Shift Lead 6:30–4:30 (10-hour option) = 16 h, with Admin covering 8–10 when scheduled. In the first cut the Production Line closes a 10.25-hour Friday — in the template the lead does.",
        "Rate: the first cut has $22.75–25.25 team members on ~40 hours of barista/production shifts at Downtown (Tue/Wed/Thu 9–4:30 Production Barista/Barista; Mon/Tue Production Line). At $19 those hours save ~$180/wk; "
        "if they must be scheduled, put them in the Shift Lead slot.",
        "Net: first cut 144.3 h / $3,058 at the people's rates (37.0% of a $8,266 mid week; 46.6% of a $6,566 slow week) → mid template 133 h / $2,632 (31.8%) → slow template 108 h / $2,154 (32.8%). Columbus Day (10/12): Downtown closed, unschedule Monday.",
    ],
    "Studio": [
        "Sunday (first cut: Barista 7:30–2, Shift Lead 8–3, Production Line 7:15–3:30, Production Barista 10:30–5:30 = 28.75 h) — keep as is; it matches the slow-tier template (28.5 h, 28%) almost exactly.",
        "Monday (first cut: Shift Lead 6:30–2:30, Production Barista 6:30–1, Production Line 6:15–2:30 + Manager 7–2 = 22.75 core h). The Manager block is on the floor at open, so the catering delivery is covered by 2 core + Admin: "
        "Production Line 6–2, Shift Lead 6:30–2:45, rush barista 8:00–11:30 (not 6:30–1). Mid: 19.75 h (−3). Slow: 16.25 h with Admin (−6.5).",
        "Tuesday (first cut: Production Barista 6:30–1, Shift Lead 7–2:30, Production Line 6:15–2:30 = 22.25 h). Studio's weakest day ($900–1,040): 2-person core day — Production Line 6–2, Shift Lead 6:30–2:45 = 16.25 h (−6); "
        "peak is 20–23/hr, two people at 14 each. Admin covers 8–10 when scheduled.",
        "Wednesday (first cut: Barista 6:30–1, Shift Lead 7–2:30, Production Line 6:15–2:30 = 22.25 h). Mid: Barista becomes an 8–12 rush shift = 20.25 h (−2); slow: 2-person day, 16.25 h (−6).",
        "Thursday (first cut: Shift Lead 7–2:30, Production Line 6:15–2:30, Production Barista 6:30–1:30 + Manager 9–2 = 22.75 core h). The Manager block 9–2 covers the rush tail and lunch: mid 20.25 h with the barista at 8–12, "
        "or 16.25 h (−6.5) when the Admin block is on the floor.",
        "Friday (first cut: Shift Lead 6:30–3, Production Line 6:15–2:30, second Shift Lead 10–5:15, Barista 7–1 = 30 h). Second lead becomes a Closer 10:30–5:30 at barista rate; Barista 8–1 not 7–1; mid template 28.5 h (−1.5) at 30.7%.",
        "Saturday (first cut: Production Line 7:15–3, Shift Lead 10–5:30, Shift Lead 7:30–1, Hospitality Register 8–3, Barista 9–4 = 34.75 h). Shape is right for a mid Saturday (36 h template, 25%); make the 10–5:30 lead a Closer and "
        "move Hospitality Register to 8:30–4:30 so the 4p–5p hour is not a single person.",
        "Net: first cut 183.5 h / $3,718 at the people's rates (32.7% of a $11,360 mid week; 38.2% of a $9,734 slow week) → mid template 176 h / $3,470 (30.5%) → slow template 153.5 h / $3,043 (31.3%). Ryan's own scoring of the first cut (30.8% on a $12,124 projected week) is consistent with this: "
        "the first cut is a busy/mid-tier schedule and will run 37% if Studio's September–October land where August did.",
    ],
    "1899 Penn": [
        "Monday (first cut: Shift Lead 6:30–3:30 + Production Barista 6:30–1:30 + Manager 8–2 = 16 core h). With the Admin block on 8–2, the Production Barista is a 4-hour rush shift 8–12 (13 h); without it, 7:30–1:30 (15 h).",
        "Tuesday (first cut: Shift Lead 6:30–3:30 + Production Barista 6:30–3:30 + Manager 6:30–1:30 = 18 core h + 7 admin). Three people on a $850–1,100 morning. Production Barista 7:30–2:30 (16 h, −2); the Admin block, if kept, covers the 1pm peak "
        "and the Production Barista can leave at 1:30 (15 h).",
        "Wednesday (both 6:30–3:30 = 18 h) and Thursday (Shift Lead 6:30–3:30 + Barista 6:30–3 = 17.5 h): second person 7:30–2:30 → 16 h each (−2 / −1.5). Both people at 6:30 for a 7:00 open with 5–7 transactions in the first hour is the visible overlap.",
        "Friday (first cut: Shift Lead 6:30–3:30 + Production Barista 6:30–1:30 = 16 h). Friday is 61 transactions a day in mid weeks: Production Barista 8–12 (13 h, −3) or, lean, the lead alone with the ALL covering 8:30–10 and 12:30–1:30 (9 h, −7).",
        "Rate: 1899's second person is often a $22.75 team member (Production Barista); a third of 1899's core hours are paid at $22.50+. At $19 that is ~$80/wk.",
        "Net: first cut 85.5 h / $1,756 at the people's rates (46.7% of a $3,759 mid week) → mid template 76 h / $1,534 (40.8%, the ceiling) → lean 67 h / $1,363 (36.3%). The memo's 'second person goes half-day on slow days → ~$1,500/wk' lands at 40%; "
        "35% at mid-tier sales needs the lean shape. Columbus Day (10/12): closed.",
    ],
}


def incentive_section(doc):
    add_heading(doc, "8. Incentive — sizing, structures, margins, implementation", level=1)
    hb = INC["headroom_by_tier"]
    add_para(doc, "The purpose of the monthly incentive is to give the team a visible share of the upside from running to the templates, without crowding out the annual profit share into retirement accounts. "
                  "Everything below is on the wage basis (gross core wages ÷ core sales) so that the number the team sees is the number they control. Payouts would be wages (taxed and loaded at 12% like any other pay).", after=6)
    add_heading(doc, "8a. Headroom — what is available at target, by tier", level=2)
    rows = []
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            v = hb[c][tier]
            rows.append([f"{c} — {tier}", m(v["core_sales_month"]), m(v["actual_wages_month"]), p(v["actual_pct"]), m(v["target_wages_month"]), m(v["headroom_actual_vs_target"]), m(v["template_wages_month"]), p(v["template_pct"])])
    tot_slow = sum(hb[c]["slow"]["headroom_actual_vs_target"] for c in CAFES); tot_mid = sum(hb[c]["mid"]["headroom_actual_vs_target"] for c in CAFES); tot_busy = sum(hb[c]["busy"]["headroom_actual_vs_target"] for c in CAFES)
    rows.append(["Open Cafes — slow / mid / busy", "", "", "", "", f"{m(tot_slow)} / {m(tot_mid)} / {m(tot_busy)}", "", ""])
    table(doc, ["Cafe / tier", "Core sales / mo", "Actual core wages", "Actual %", "Target wages", "Headroom vs target", "Template wages", "Template %"], rows,
          widths=[1.5, 0.85, 0.9, 0.6, 0.85, 0.95, 0.85, 0.7], total_rows=[9], size=9)
    add_caveat(doc, f"Monthly = weekly × 52/12. 'Actual' is the sampled weeks of each tier at the rates paid. August's realized headroom was {m(INC['headroom_aug_total'])} (all three cafes, core only); the Apr–Jun average was {m(INC['headroom_aprjun_avg_total'])} a month. "
                    "Headroom is what running at target saves Slipstream; the incentive is a share of what is saved BELOW target, so the incentive never spends the headroom itself.")

    add_heading(doc, "8b. Structures compared", level=2)
    add_para(doc, "Five structures were modelled on a mid-tier month for the three open cafes combined ($101K core sales, ~1,670 core hours). Payout, cost as % of sales, and what a 30-hour-a-week team member would see:", after=4)
    oc = INC["plan_comparison_mid_tier"]["Open Cafes"]["scenarios"]
    names = {"A_beat_share_50_cap1.5": "A. Beat Share — 50% of wages saved below target, cap 1.5% of sales",
             "A2_ryan_2pt_gate": "A2. Ryan's draft — 50% of the beat, only when ≥2 pts under target",
             "B_dollar_per_hour": "B. Dollar-an-hour — $1.00/$1.50/$2.00 per core hour at ≤ target / −2 / −4",
             "C_prime_gainshare_20": "C. Gainshare 20% of the beat (no gate, no cap)",
             "D_shift_bonus_8": "D. $8 per shift in weeks that hit target"}
    rows = []
    for key, label in names.items():
        r = [label]
        for scen in ["AT target", "1 pt under", "2 pts under", "4 pts under"]:
            v = oc[scen][key]
            r.append(f"{m(v['payout'])} ({v['pct_of_sales']:.1f}%) · ${v['per_person_month_30h']}")
        rows.append(r)
    table(doc, ["Structure", "At target", "1 pt under", "2 pts under", "4 pts under"], rows, widths=[2.6, 1.05, 1.05, 1.05, 1.05], size=8)
    add_caveat(doc, "Each cell: company payout for the month (as % of core sales) · what a 30 h/wk team member receives. Nothing pays when a cafe is over target under any structure.")
    add_para(doc, "How they compare on Ryan's criteria:", after=2)
    add_bullet(doc, "Self-funding. A, A2 and C pay only out of dollars actually saved; Slipstream keeps at least half. B and D pay ~1.7% of sales just for hitting target — at exactly-target that is $1,700 a month of new cost with no saving to fund it.", bold_lead="")
    add_bullet(doc, "Simplicity and transparency. B is the easiest sentence ('hit the number, everyone gets a dollar an hour'); A is one sentence longer ('we split the savings below the line 50/50, up to $1.50 an hour') and it is the one that shows the team the same number Ryan watches. C is too small to notice. D fires weekly, which is motivating but noisy.", bold_lead="")
    add_bullet(doc, "Cliffs and gaming. A2's 2-point gate creates a dead zone where a 1.5-point beat pays nothing and a 2.0-point beat pays $1,000 — an invitation to cut a shift on the 28th. A pays from the first dollar and caps, so there is no cliff. Every structure needs the same guardrail against under-staffing (below).", bold_lead="")
    add_bullet(doc, "Fairness across cafes. 23 of the 26 team members work at more than one cafe, and 1899's target is structurally hard in slow and mid weeks. A pooled 'Open Cafes' calculation (sum of the three cafes' beats, split by core hours worked anywhere) is fairer and simpler to pay than three separate pools; the per-cafe scoreboard stays visible.", bold_lead="")
    add_heading(doc, "8c. Recommendation — 'Beat Share', pooled, capped, guarded", level=2)
    for line in [
        "Metric: monthly core wages ÷ monthly core sales per cafe, on the same Square-timecard basis as the weekly report (Manager and Training excluded; missed clock-outs capped). Published weekly, settled monthly.",
        "Pool: 50% × Σ over open cafes of (target% × core sales − actual core wages), where a cafe over its target contributes a negative amount. If the sum is negative, no payout. Cap the pool at 1.5% of open-cafe core sales (≈ $1,500 in a mid month).",
        "Split: by core hours worked in the month at any open cafe. A 30-hour week earns ~$40–120 a month depending on the beat; the maximum is about $1.20 an hour at the cap. NMWA hours are included in the split (its team is the same people) but NMWA is not in the metric.",
        "Guardrail (the service test): the pool pays only if the month's schedules stayed at or above the template floors — no single-person coverage during a rush block (8:30–10:00 at Downtown/1899, 9:30–1:00 at Studio weekends) on more than two days, "
        "and no week with under-coverage flagged by the weekly report. This is the one rule that stops the plan from rewarding skeleton crews.",
        "Timing: shadow-mode in the September and October weekly reports (show what the pool WOULD have paid); first real payout after two consecutive months at or under target company-wide (Ryan's rule), realistically December or January; pay in the first payroll of the following month as a separate earnings line.",
        "Size discipline: at 1.5% of core sales the cap is ~$18K a year across the company if every month hit it — which would only happen if the templates were beaten by 3+ points all year. Realistic: $6–10K a year. The annual profit share at 5–10% of open-cafe operating income at target ($218K annualized on August sales) is $11–22K; the two together are 1.5–3% of core sales.",
    ]:
        add_bullet(doc, line)
    add_heading(doc, "8d. Margins — what each path does to the P&L (August sales)", level=2)
    mv = INC["margin_view_august"]
    rows = []
    for c in CAFES:
        for label, r in mv[c]["rows"].items():
            rows.append([f"{c} — {label}", p(r["labor_pct"]), p(r["prime_pct"]), m(r["noi"]), p(r["noi_pct"]), m(r["incentive_paid"])])
    table(doc, ["Cafe / case", "P&L labor %", "Prime cost %", "NOI", "NOI %", "Incentive paid"], rows, widths=[3.0, 0.8, 0.8, 0.85, 0.65, 0.8], key_rows=[2, 7, 12], size=8)
    add_caveat(doc, "P&L labor via the bridge (1.12 × gross − tips) at August COGS and operating expenses; admin and training held at August levels. 'Plan B' costs $1.50 per core hour at a 2-point beat — about three times Plan A — and gives back a third of the beat. "
                    "Even with the incentive paid, every cafe at target is 8–10 prime-cost points better than August; the incentive moves prime cost by ~1 point at Plan A and ~2.5 at Plan B.")
    add_heading(doc, "8e. Comparable plans (directional, not audited)", level=2)
    add_para(doc, "Hourly team incentives in independent cafes and restaurants typically run 1–3% of wages and are tied to a single controllable number (labor % or prime cost) with a service guardrail; manager bonuses run 10–20% of salary on prime cost and sales. "
                  "Gainsharing programs (Zingerman's-style, or the 'MiniGame' pattern from open-book management) share a fixed fraction — commonly 25–50% — of a measured improvement over a baseline, pay monthly or quarterly, and cap the pool. "
                  "Profit-sharing contributions to retirement plans, when small employers make them, are usually 2–5% of pay in a good year, allocated by compensation or by hours under a plan document with a vesting schedule. "
                  "Slipstream's proposal (50% beat share, 1.5% cap, plus a later 5–10%-of-profit retirement contribution) sits inside those ranges and is on the generous side for the monthly piece relative to the size of the business — "
                  "which is appropriate given the removal of tip upside, provided the cap holds. These ranges are operator benchmarks, not statistics; a CPA/TPA should confirm the retirement-plan mechanics (401(k) with discretionary profit-sharing, "
                  "allocation formula, vesting, nondiscrimination testing) before anything is promised.", after=6)
    add_heading(doc, "8f. Implementation plan", level=2)
    for line in [
        "September: publish the tier bands and the 'tier in force' per cafe in the weekly report; build the 10/11 and later weeks on the Mid templates; start the shadow pool line ('what the Beat Share would have paid').",
        "October: templates live for all three cafes; first monthly review of core % vs target on the wage basis; fix the two data hygiene items that distort the metric (missed clock-outs — adopt the past-close cap in aggregate_labor.py; zero-rate timecards).",
        "November: if September and October both landed at or under target company-wide, announce the Beat Share with the guardrail and the cap; otherwise extend shadow mode one month. Team communication: one page — the formula, the guardrail, the cap, an example month.",
        "December–January: first payout. Review the cap and the split after two payouts; decide whether 1899 stays in the pool or moves to its own guardrail (NI ≥ 5% of sales).",
        "2027: once four consecutive quarters are profitable at the consolidated level, stand up the retirement plan (401(k) with discretionary profit sharing), with the monthly Beat Share cap re-set so the two together stay ≤ 3% of core sales.",
    ]:
        add_bullet(doc, line)
    add_heading(doc, "8g. Things to decide, and thoughts on the design", level=3)
    for line in [
        "Pooled vs per-cafe: pooled is recommended (floaters, 1899's structural gap). If per-cafe, give 1899 a different metric — NI ≥ 5% of sales — or its team will never earn.",
        "Basis honesty: the team will see 'core wages 30%' while the P&L says 25%. Publish the bridge once so nobody thinks the number is being managed.",
        "Do not pay on sales alone (SPLH) — hourly staff control hours more than sales; but a sales floor (payout requires the month's core sales ≥ 90% of projection) protects against the plan paying in a month the business shrank.",
        "Keep admin outside the metric but inside the guardrail: an ALL who covers the rush from an Admin block is doing exactly what the plan wants; an ALL who moves core hours into Admin to game the ratio is not — track admin hours per week alongside.",
        "The cap is the promise-keeper for the retirement plan. Write it into the one-pager as a fixed dollar-per-hour maximum ($1.20) as well as a % of sales.",
    ]:
        add_bullet(doc, line)


def appendix(doc):
    add_heading(doc, "Appendix A — Need vs actual, 30-minute detail (Tuesday, all tiers)", level=1)
    add_para(doc, "Transactions per hour at each half-hour (tier average), the people needed under the Section 9 rule, and the average core headcount actually on the clock in the sampled weeks. Full detail for every day is in the workbook.", after=4)
    for c in CAFES:
        for tier in ["slow", "mid", "busy"]:
            r = NV["cafes"][c]["tiers"][tier]["Tue"]
            rows = r["by_halfhour"]
            def tlabel(t):
                hh = int(t); mm = "30" if t % 1 else "00"; return f"{hh if hh <= 12 else hh - 12}:{mm}"
            hdr = ["Tue " + tier] + [tlabel(x["t"]) for x in rows]
            body = [["txns/hr"] + [f"{x['txns_hr']:.0f}" for x in rows], ["need"] + [str(x["need"]) for x in rows], ["actual"] + [f"{x['actual']:.1f}" for x in rows]]
            add_heading(doc, f"{c} — Tuesday, {tier} tier", level=3)
            table(doc, hdr, body, widths=[0.6] + [0.27] * len(rows), size=6)


if __name__ == "__main__":
    build()
