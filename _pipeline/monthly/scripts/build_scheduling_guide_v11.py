#!/usr/bin/env python3
"""Scheduling Guide v1.1 — Ryan's Sep 11 notes: labor bands first and prominent; two templates per cafe (BASE and FULL) with the
sales range for slow / mid / busy; slow = base + admin support + trim; Square Teams tip (deselect Manager/Training to see core)."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, add_callout, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, h
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Scheduling Guide - Sep 2026 v1.1.docx")
T = json.load(open(os.path.join(DD, "templates_v13.json")))["cafes"]
W = lambda c, t: T[c][t]["week"]

def build():
    doc = Document(); sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.8)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Scheduling Guide", 22, True, COLOR_ACCENT, 0), ("Labor bands, and the two schedules that hit them  |  September 2026  |  v1.1", 13, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "For whoever writes the schedule. The labor band is the target. Everything else here is support: two schedules per cafe (Base and Full), the weekly sales that decide which one, and what to do in a slow week.",
             italic=True, size=9, color=COLOR_MUTED, after=8)

    # ================= 1. LABOR BANDS =================
    add_heading(doc, "1. Labor bands — the target", level=1)
    add_callout(doc, "Core labor as % of core sales, for the week",
                "Downtown 32–36%, max 38%.   Studio 30–34%, max 36%.   1899 Penn 35–40%, max 42%.   NMWA: staffing rule, no %.\n"
                "Inside the band = on target. Between the band and the max = OK in a slow week, use an admin day. Over the max = Ryan's approval before the schedule is published.\n"
                "Core = floor shifts only. Admin (Square title 'Manager') and Training shifts are not in the number — in Square Teams, deselect Manager and Training when you review the week and the total you see is the core total.",
                body_size=10, title_size=10)
    add_para(doc, "Which schedule, and what it costs, by the week's projected core sales (Section 5 of the weekly report):", after=4)
    rows = [["Downtown", "under $7,500", "BASE + admin day\n120.5 h · $2,514", "$7,500 – 9,000", "FULL\n137 h · $2,832", "$9,000 +", "FULL, Mon/Fri 3rd at 9:00\n139 h · $2,872"],
            ["Studio", "under $10,500", "BASE + admin day\n183 h · $3,776", "$10,500 – 12,500", "BASE, Sat 9:30 / Sun to 5:30\n185 h · $3,821", "$12,500 +", "FULL\n193 h · $3,976"],
            ["1899 Penn", "under $3,300", "BASE + admin day\n67.5 h · $1,438", "$3,300 – 4,200", "FULL\n71 h · $1,513", "$4,200 +", "FULL, Tue/Wed to 3:30\n76.5 h · $1,629"],
            ["NMWA", "any", "1 person 9–3 (2 on community days)\n24 h · $510", "", "", "", ""]]
    table(doc, ["Cafe", "SLOW week", "schedule · hours", "MID week", "schedule · hours", "BUSY week", "schedule · hours"], rows, widths=[0.85, 0.85, 1.35, 0.95, 1.35, 0.7, 1.35], size=8)
    add_caveat(doc, "Dollars at planning rates (Lead $22.25 / Line $19.75 / Production Barista $20.25 / Barista & Closer $20.00); your roster will differ a little, so hold the HOURS. "
                    "In a slow week the Base schedule alone lands above the max (it needs about $6,600 / $10,500 / $3,400 of weekly sales to be under it) — that is a slow week, not a scheduling error: use the admin day, trim the shifts marked below, publish, and flag it.")

    # ================= 2. HOW TO USE =================
    add_heading(doc, "2. Every week", level=1)
    for lab, txt in [("Pick the column.", "Projected core sales for the coming Sun–Sat week from the weekly report. Core = walk-in, no catering or invoices."),
                     ("Copy the schedule.", "Base or Full, from Section 3. Do not go under Base: two openers and two closers at Downtown and Studio, six-hour minimum shifts, three people mid-day at Studio, two shifts every day at 1899."),
                     ("Slow week?", "Base schedule, plus: put the admin day where it covers the floor and drop the day's trim shift (marked ✂ below); look for another hour or two where the rush is thin. Never cut an opener, a closer or the Studio mid-day third."),
                     ("Check core hours, publish.", "Deselect Manager and Training in Square Teams; the total should be at or under the hours in the column you picked. Over the max % → ask Ryan first.")]:
        add_bullet(doc, txt, bold_lead=lab)

    # ================= 3. SCHEDULES =================
    add_heading(doc, "3. The two schedules per cafe", level=1)
    add_heading(doc, "Downtown — 7a–4p Mon–Fri", level=2)
    rows = [["Line (kitchen)", "6:15–12:15", "6:15–1:15", "6:15–1:15", "6:15–1:15", "6:15–12:15"], ["Lead (opens)", "6:30–4:30", "6:30–4:30", "6:30–4:30", "6:30–4:30", "6:30–4:30"],
            ["Third", "Barista 10:30–4:30 ✂", "Closer 8:00–4:30", "Closer 8:00–4:30", "Closer 8:00–4:30", "Barista 10:30–4:30 ✂"], ["Hours", "22", "25.5", "25.5", "25.5", "22"]]
    table(doc, ["BASE · 120.5 h", "Mon", "Tue", "Wed", "Thu", "Fri"], rows, widths=[1.15, 1.15, 1.1, 1.1, 1.1, 1.15], total_rows=[3], size=9)
    rows = [["Line (kitchen)", "6:15–12:15", "6:15–12:15", "6:15–12:15", "6:15–12:15", "6:15–12:15"], ["Lead (opens)", "6:30–4:30", "6:30–2:30", "6:30–2:30", "6:30–2:30", "6:30–4:30"],
            ["Third", "Closer 9:00–4:30", "Closer 8:00–4:30", "Closer 8:00–4:30", "Closer 8:00–4:30", "Closer 9:00–4:30"], ["Fourth", "—", "Barista 8:30–4:30", "Barista 8:30–4:30", "Barista 8:30–4:30", "—"], ["Hours", "23.5", "30.5", "30.5", "30.5", "23.5"]]
    table(doc, ["FULL · 137 h", "Mon", "Tue", "Wed", "Thu", "Fri"], rows, widths=[1.15, 1.15, 1.1, 1.1, 1.1, 1.15], total_rows=[4], size=9)
    add_caveat(doc, "Three people Tue–Thu under $7,500 a week; four from $7,500 — never five. In the $7,500–9,000 range keep Mon/Fri third at 10:30 (137 h). Friday is the weak day and reads 45–60% on its own; Gee's Friday admin day covers it. Federal-holiday Mondays: closed. Tue, Wed and Thu are always the same schedule.")

    add_heading(doc, "Studio — 7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun", level=2)
    rows = [["Line (kitchen)", "6:15–2:30", "6:15–2:30", "6:15–2:15", "7:15–3:15", "7:15–3:15"], ["Lead (opens)", "6:30–2:30", "6:30–2:30", "6:30–2:30", "7:30–5:30", "7:30–3:30"],
            ["Third", "Barista 7:00–1:00", "Closer 8:00–2:30 ✂", "Closer 8:30–5:30", "Barista 8:00–2:00", "Barista 9:30–5:30"], ["Fourth", "—", "—", "Barista 11:30–5:30", "Barista 10:30–5:30 ✂", "Closer 11:00–5:30"], ["Hours", "22.25", "22.75 each", "31", "31", "30.5"]]
    table(doc, ["BASE · 183 h", "Mon", "Tue–Thu", "Fri", "Sat", "Sun"], rows, widths=[1.1, 1.15, 1.15, 1.15, 1.15, 1.15], total_rows=[4], size=9)
    rows = [["Line (kitchen)", "6:15–2:30", "6:15–2:30", "6:15–2:15", "7:15–3:15", "7:15–3:15"], ["Lead (opens)", "6:30–2:30", "6:30–2:30", "6:30–2:30", "7:30–3:30", "7:30–5:30"],
            ["Third", "Barista 7:00–1:00", "Closer 8:00–2:30", "Closer 8:30–5:30", "Barista 8:00–2:00", "Barista 8:00–2:00"], ["Fourth", "—", "—", "Barista 9:30–5:30", "Barista 9:30–5:30", "Barista 9:30–5:30"], ["Fifth", "—", "—", "—", "Closer 10:00–5:30", "—"], ["Hours", "22.25", "22.75 each", "33", "37.5", "32"]]
    table(doc, ["FULL · 193 h", "Mon", "Tue–Thu", "Fri", "Sat", "Sun"], rows, widths=[1.1, 1.15, 1.15, 1.15, 1.15, 1.15], total_rows=[5], size=9)
    add_caveat(doc, "Mon–Thu never change: Line on all day and closing the line, Lead, and a third person 8:00–2:30 (7:00–1:00 Monday for the delivery). Friday is four people at most in any week. Five only on Saturday and only from $12,500. Mid weeks ($10,500–12,500): Base with the Saturday second barista at 9:30 and the Sunday Lead to 5:30 (185 h). Colin's Mon/Wed admin days cover the ✂ Closer shift in slow weeks; Claire's weekend day covers the Saturday ✂.")

    add_heading(doc, "1899 Penn — 7a–3p Mon–Fri", level=2)
    rows = [["Lead (opens)", "6:30–12:30", "6:30–2:00", "6:30–2:30", "6:30–2:30", "6:30–12:30"], ["Production Barista", "9:30–3:30", "9:00–3:30", "8:30–3:30", "9:00–3:30", "9:30–3:30"], ["Hours", "12", "14", "15", "14.5", "12"]]
    table(doc, ["BASE · 67.5 h", "Mon", "Tue", "Wed", "Thu", "Fri"], rows, widths=[1.4, 1.05, 1.05, 1.05, 1.05, 1.05], total_rows=[2], size=9)
    rows = [["Lead (opens)", "6:30–2:00", "6:30–2:30", "6:30–2:30", "6:30–2:30", "6:30–12:30"], ["Production Barista", "9:00–3:30", "8:30–3:30", "8:30–3:30", "8:30–3:30", "9:30–3:30"], ["Hours", "14", "15", "15", "15", "12"]]
    table(doc, ["FULL · 71 h", "Mon", "Tue", "Wed", "Thu", "Fri"], rows, widths=[1.4, 1.05, 1.05, 1.05, 1.05, 1.05], total_rows=[2], size=9)
    add_caveat(doc, "Two shifts every day — never a solo day; on stagger days the Lead opens alone and the Production Barista closes alone, both on through the rush. Whoever is on at 2:00 closes. Busy weeks ($4,200 and up): Tue and Wed both people to 3:30 with the Production Barista from 7:30, Thu from 8:00 (76.5 h). Eli's Tuesday admin day is the slow-week cover. Federal-holiday Mondays: closed.")

    add_heading(doc, "NMWA — Thu–Sun", level=2)
    add_para(doc, "One FOH shift 9:00–3:00 each open day (an hour to set up, an hour to close). Two shifts 9:00–3:00 on community days — the second Wednesday and the first Sunday. No labor % while sales ramp; a second person on any other day is an ask to Ryan.", after=6)

    # ================= 4. RULES =================
    add_heading(doc, "4. Rules that do not bend", level=1)
    for txt in ["Two openers (Line in 45 minutes before open, Lead 30) and two closers (out within 30 minutes of close) at Downtown and Studio. Six-hour minimum on every shift.",
                "Downtown: third person on by 8:00 Tue–Thu; four people at most. Studio: Line on all day Mon–Thu; three people mid-day; four max on Friday; five only on Full-week Saturdays. 1899: two shifts every day.",
                "ALLs and Shift Leads go in the Lead and Closer slots, not the Barista slots — a $25 person on a Barista shift costs the week about $180.",
                "Admin (Manager) and Training are scheduled on top and reported separately — set the Square job title correctly or the shift books as floor labor and the week reads over.",
                "Known event or delivery day: add the hours, note it on the schedule. Federal-holiday Mondays: Downtown and 1899 closed; Studio runs its Monday base."]:
        add_bullet(doc, txt)
    add_caveat(doc, "Source: Slipstream Labor Deep Dive — Schedule Templates by Tier v1.4 (Sep 2026). Questions on the numbers → Ryan.")
    _ensure_zoom_percent(doc); doc.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    build()
