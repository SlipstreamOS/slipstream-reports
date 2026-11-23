#!/usr/bin/env python3
"""Build 'Slipstream Scheduling Guide - Sep 2026 v1.0.docx' — the short, shareable version of the templates for the people who
write the schedules (Gee, Mo, Eli, Colin, Claire): one base schedule per cafe, add-ons keyed to projected weekly sales, the
weekly hours/$ budget, the labor-% band, and the rules. No pay or incentive content. Numbers from templates_v13.json."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m, p, h
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Scheduling Guide - Sep 2026 v1.0.docx")
T = json.load(open(os.path.join(DD, "templates_v13.json")))["cafes"]
W = lambda c, t: T[c][t]["week"]
BAND = {"Downtown": (32, 36, 38), "Studio": (30, 34, 36), "1899 Penn": (35, 40, 42)}
def need(c, t, pct): return W(c, t)["wages"] / (pct / 100)

def build():
    doc = Document(); sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.8)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Scheduling Guide", 22, True, COLOR_ACCENT, 0), ("How to build a week that hits the labor target  |  September 2026  |  v1.0", 13, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "For the people who write the schedules. One base schedule per cafe, a short list of add-ons that switch on when the week is projected to be busier, and the weekly hours and dollars each one costs. "
                  "Built from Jan–Aug 2026 sales by half hour and the operating rules agreed with Ryan in September. Hours are the thing to manage; the labor % follows from hours and sales.", italic=True, size=9, color=COLOR_MUTED, after=10)

    add_heading(doc, "1. The four steps, every week", level=1)
    for lab, txt in [("Get the number.", "Projected core sales for the cafe for the coming Sun–Sat week — Section 5 of the weekly report (Projections v1.9). Core = walk-in sales, no catering or invoices."),
                     ("Start from the base schedule.", "It is the floor: two openers and two closers at Downtown and Studio, six-hour minimum shifts, three people through the middle of the day at Studio, two shifts every day at 1899, one person 9:00–3:00 at NMWA. Never go under it."),
                     ("Switch on add-ons only when the projection crosses the line.", "Each cafe has one or two. If the week is projected under the line, leave them off — the base schedule is the plan."),
                     ("Check the budget, then publish.", "Total core hours for the week against the table in Section 2. Under the base + add-ons in force = on plan. Over the ceiling = ask Ryan before publishing. Admin and training shifts are budgeted on top and do not count here.")]:
        add_bullet(doc, txt, bold_lead=lab)

    add_heading(doc, "2. Weekly budget and labor band", level=1)
    rows = []
    for c in ["Downtown", "Studio", "1899 Penn"]:
        tg, ce, mx = BAND[c]
        rows.append([c, f"{h(W(c,'slow')['hours'],1)} h · {m(W(c,'slow')['wages'])}", f"{h(W(c,'mid')['hours'],1)} h · {m(W(c,'mid')['wages'])}", f"{h(W(c,'busy')['hours'],1)} h · {m(W(c,'busy')['wages'])}", f"{tg}–{ce}%  (max {mx}%)"])
    rows.append(["NMWA", "24 h · $510 (+6 h each community day)", "—", "—", "staffing rule"])
    table(doc, ["Cafe", "Base week", "+ Add-on 1", "+ Add-on 2", "Labor % band"], rows, widths=[1.0, 1.7, 1.5, 1.5, 1.3], size=9)
    add_caveat(doc, "Dollars at planning rates (Lead $22.25, Line $19.75, Production Barista $20.25, Barista/Closer $20.00) — your actual roster will differ a little, which is why hours are the number to hold. Labor % = core wages ÷ core sales for the week: green inside the band, amber up to the max, red above it (Ryan's approval needed).")
    rows = []
    for c in ["Downtown", "Studio", "1899 Penn"]:
        tg, ce, mx = BAND[c]; rows.append([c, m(need(c, "slow", tg)), m(need(c, "slow", ce)), m(need(c, "slow", mx))])
    table(doc, ["Base schedule sits at…", f"target", "ceiling", "max"], rows, widths=[1.8, 1.4, 1.4, 1.4], size=9)
    add_caveat(doc, "Weekly core sales the BASE schedule needs to land at each line. Below the 'max' figure the base schedule alone is over the max — that is not a scheduling error, it is a slow week. Do not cut an opener or closer to fix it: use an admin day (Section 4) and tell Ryan.")

    # ---------------- Downtown ----------------
    add_heading(doc, "3. Base schedules and add-ons", level=1)
    add_heading(doc, "Downtown — 7a–4p Mon–Fri  (Gee)", level=2)
    rows = [["Line (kitchen)", "6:15–12:15", "6:15–1:15", "6:15–1:15", "6:15–1:15", "6:15–12:15"], ["Lead (opens)", "6:30–4:30", "6:30–4:30", "6:30–4:30", "6:30–4:30", "6:30–4:30"],
            ["Third person", "Barista 10:30–4:30", "Closer 8:00–4:30", "Closer 8:00–4:30", "Closer 8:00–4:30", "Barista 10:30–4:30"], ["Hours", "22", "25.5", "25.5", "25.5", "22"]]
    table(doc, ["BASE · 120.5 h/wk", "Mon", "Tue", "Wed", "Thu", "Fri"], rows, widths=[1.3, 1.1, 1.1, 1.1, 1.1, 1.1], total_rows=[3], size=9)
    add_bullet(doc, "Tue/Wed/Thu add a Barista 8:30–4:30; the Line goes 6:15–12:15 and the Lead 6:30–2:30 (Closer keeps 8:00–4:30). Four people, +15 h → 137 h/wk. Never a fifth.", bold_lead="Add-on 1 — week projected $7,500 or more:")
    add_bullet(doc, "Mon and Fri third person becomes a Closer 8:30/9:00–4:30 instead of 10:30. +2 h → 139 h/wk.", bold_lead="Add-on 2 — $9,000 or more:")
    add_bullet(doc, "Friday is the weak day (sales ~$800–950): 22 hours is already the floor, so Friday alone will read 45–60% — that is expected. Gee's Friday admin day covers it. Federal-holiday Mondays: closed, no shifts.", bold_lead="Notes.")

    # ---------------- Studio ----------------
    add_heading(doc, "Studio — 7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun  (Mo)", level=2)
    rows = [["Line (kitchen)", "6:15–2:30", "6:15–2:30", "6:15–2:15", "7:15–3:15", "7:15–3:15"], ["Lead (opens)", "6:30–2:30", "6:30–2:30", "6:30–2:30", "7:30–5:30", "7:30–3:30"],
            ["Third", "Barista 7:00–1:00", "Closer 8:00–2:30", "Closer 8:30–5:30", "Barista 8:00–2:00", "Barista 9:30–5:30"], ["Fourth", "—", "—", "Barista 11:30–5:30", "Barista 10:30–5:30", "Closer 11:00–5:30"], ["Hours", "22.25", "22.75 each", "31", "31", "30.5"]]
    table(doc, ["BASE · 183 h/wk", "Mon", "Tue–Thu", "Fri", "Sat", "Sun"], rows, widths=[1.0, 1.2, 1.2, 1.25, 1.25, 1.25], total_rows=[4], size=9)
    add_bullet(doc, "Sat second Barista starts 9:30 (not 10:30); Sun Lead stays to 5:30 with Baristas 8:00–2:00 and 10:00–5:30. +2 h → 185 h/wk.", bold_lead="Add-on 1 — week projected $10,500 or more:")
    add_bullet(doc, "Fri Barista starts 9:30 (not 11:30); Sat adds a fifth person, Closer 10:00–5:30, with the Lead out at 3:30; Sun second Barista 9:30. +8 h → 193 h/wk. Five people only on Sat/Sun and only in these weeks; Friday is four at most.", bold_lead="Add-on 2 — $12,500 or more:")
    add_bullet(doc, "The Line person is on all day Mon–Thu and closes the line; the Baristas take the shorter shifts. Both weekday Baristas should be shift-lead capable so you can swap. Monday keeps three at open for the delivery.", bold_lead="Notes.")

    # ---------------- 1899 ----------------
    add_heading(doc, "1899 Penn — 7a–3p Mon–Fri  (Eli)", level=2)
    rows = [["Lead (opens)", "6:30–12:30", "6:30–2:00", "6:30–2:30", "6:30–2:30", "6:30–12:30"], ["Production Barista", "9:30–3:30", "9:00–3:30", "8:30–3:30", "9:00–3:30", "9:30–3:30"], ["Hours", "12", "14", "15", "14.5", "12"]]
    table(doc, ["BASE · 67.5 h/wk", "Mon", "Tue", "Wed", "Thu", "Fri"], rows, widths=[1.4, 1.1, 1.1, 1.1, 1.1, 1.1], total_rows=[2], size=9)
    add_bullet(doc, "Monday becomes a full day (Lead to 2:00, Production Barista from 9:00); Tue and Thu Production Barista from 8:30. +3.5 h → 71 h/wk.", bold_lead="Add-on 1 — week projected $3,300 or more:")
    add_bullet(doc, "Tue and Wed both people to 3:30 with the Production Barista from 7:30; Thu from 8:00; Fri from 9:00. +5.5 h → 76.5 h/wk.", bold_lead="Add-on 2 — $4,200 or more:")
    add_bullet(doc, "Two shifts every day — never a solo day. Mon and Fri are stagger days (solo open, solo close, two people through the rush). Whoever is on at 2:00 closes with the team. Federal-holiday Mondays: closed.", bold_lead="Notes.")

    # ---------------- NMWA ----------------
    add_heading(doc, "NMWA — Thu–Sun  (Gee / Claire)", level=2)
    add_para(doc, "One FOH shift 9:00–3:00 each open day (a full hour to set up and to close). Two FOH shifts 9:00–3:00 on community days — the second Wednesday and the first Sunday of the month. That is the whole rule; there is no labor % target while sales ramp. A second person on any other day is an ask to Ryan.", after=6)

    add_heading(doc, "4. Slow weeks and admin days", level=1)
    for lab, txt in [("What a slow week looks like.", "Projected sales under the 'max' figure in Section 2 (Downtown under ~$6,600, Studio under ~$10,500, 1899 under ~$3,400). The base schedule alone is over the max. Expected in January, July–August, the week before Christmas and Thanksgiving week."),
                     ("What to do.", "Do not cut an opener, a closer or the Studio mid-day third. Use the admin day: when the admin person is in the building (Gee Mon/Thu/Fri at Downtown, Colin Mon/Wed at Studio, Eli Tue at 1899, Claire one weekend day), drop the day's shortest non-opener/closer shift — Downtown's 10:30–4:30 Barista on Mon/Fri, Studio's 8:00–2:30 Closer on Mon–Thu, Studio's 10:30 Sat Barista on Claire's day. Each swap is worth 2–3 points on the week."),
                     ("Tell Ryan.", "If the week is still over the max after the swaps, publish it and flag it — a slow week is a sales problem, not a schedule problem, and the report will show it that way."),
                     ("Known events.", "Add hours for a booked event or a big delivery day on top of the plan and note it; the report tracks catering labor separately.")]:
        add_bullet(doc, txt, bold_lead=lab)

    add_heading(doc, "5. Rules that do not bend", level=1)
    for txt in ["Two openers (Line in 45 minutes before open, Lead 30 minutes) and two closers (out no more than 30 minutes after close) at Downtown and Studio.",
                "Six-hour minimum on every shift.", "Downtown: third person on by 8:00 Tue–Thu; four people at most, ever. Tue, Wed and Thu are the same schedule.",
                "Studio: Line person on all day Mon–Thu; three people mid-day; four at most on Friday and in base-week weekends; five only on add-on-2 weekends.",
                "1899: two shifts every day; the second person on at close closes.", "Put ALLs and Shift Leads in the Lead and Closer slots, not the Barista slots — a $25 person on a Barista shift costs the week about $180.",
                "Core hours only count against the budget. Admin (Square title 'Manager') and Training are scheduled on top and reported separately — set the Square job title correctly or the shift books as floor labor.",
                "Federal-holiday Mondays: Downtown and 1899 closed. Studio open with the Monday base."]:
        add_bullet(doc, txt)
    add_caveat(doc, "Source: Slipstream Labor Deep Dive — Schedule Templates by Tier v1.4 (Sep 2026); workbook 'Slipstream Schedule Templates - Sep 2026 v1.4.xlsx' has every shift as an editable row. Questions on the numbers → Ryan.")
    _ensure_zoom_percent(doc); doc.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    build()
