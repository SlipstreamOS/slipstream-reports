#!/usr/bin/env python3
"""Scheduling Guide v1.2 — Ryan's Sep 11 shift edits applied (Studio: third out by 1:30 Mon–Thu, Friday two leads and first out
1:00, weekends one closer / solo after 3:30 / first out by 2:00, Sunday full = 5 with a short 7:30–12:30 lead; 1899: second person
7:30 Tue–Thu in Base and every day in Full, first out 2:30 Base / 3:00 Full). Schedules are data here and also written to
labor_deep_dive/guide_schedules_v12.json for the weekly report. Downtown unchanged from templates_v13."""
import os, json, sys
from docx import Document
from docx.shared import Pt, Inches
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from report_style import add_para, add_bullet, add_caveat, add_callout, _ensure_zoom_percent, _set_run_font, _set_para_spacing, COLOR_MUTED, COLOR_ACCENT, FONT_NAME
from build_labor_deep_dive import table, add_heading, m
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Scheduling Guide - Sep 2026 v1.2.docx")
RATE = {"Line": 19.75, "Lead": 22.25, "Closing Lead": 22.25, "Closer": 20.00, "Barista": 20.00, "Production Barista": 20.25, "FOH": 21.25}
BANDS = {"Downtown": (7500, 9000), "Studio": (10500, 12500), "1899 Penn": (3300, 4200)}
def t2h(s):
    hh, mm = s.split(":"); hh = int(hh); mm = int(mm)
    return hh + mm / 60
def hours(sh):  # "6:15–14:30" in 24h
    a, b = sh.split("–"); return t2h(b) - t2h(a)
def fmt(sh):    # 24h -> 12h without am/pm
    a, b = sh.split("–")
    f = lambda t: f"{(int(t.split(':')[0]) - 1) % 12 + 1}:{t.split(':')[1]}"
    return f"{f(a)}–{f(b)}"

# role -> day -> shift (24h). None = no shift.
S = {
 "Downtown": {
  "BASE": {"days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "rows": [
     ("Line", ["6:15–12:15", "6:15–13:15", "6:15–13:15", "6:15–13:15", "6:15–12:15"]),
     ("Lead", ["6:30–16:30"] * 5),
     ("Third", [("Barista", "10:30–16:30", "✂"), ("Closer", "8:00–16:30", ""), ("Closer", "8:00–16:30", ""), ("Closer", "8:00–16:30", ""), ("Barista", "10:30–16:30", "✂")])]},
  "FULL": {"days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "rows": [
     ("Line", ["6:15–12:15"] * 5),
     ("Lead", ["6:30–16:30", "6:30–14:30", "6:30–14:30", "6:30–14:30", "6:30–16:30"]),
     ("Third", [("Closer", "9:00–16:30", ""), ("Closer", "8:00–16:30", ""), ("Closer", "8:00–16:30", ""), ("Closer", "8:00–16:30", ""), ("Closer", "9:00–16:30", "")]),
     ("Fourth", [None, ("Barista", "8:30–16:30", ""), ("Barista", "8:30–16:30", ""), ("Barista", "8:30–16:30", ""), None])]}},
 "Studio": {
  "BASE": {"days": ["Mon", "Tue–Thu", "Fri", "Sat", "Sun"], "mult": [1, 3, 1, 1, 1], "rows": [
     ("Line", ["6:15–14:30", "6:15–14:30", "6:15–14:15", "7:15–15:15", "7:15–15:15"]),
     ("Lead", ["6:30–14:30", "6:30–14:30", "6:30–13:00", "7:30–14:00", "7:30–14:00"]),
     ("Third", [("Barista", "7:00–13:00", ""), ("Barista", "7:30–13:30", "✂"), ("Barista", "9:00–15:00", "✂"), ("Barista", "8:00–14:00", "✂"), ("Barista", "8:00–14:00", "✂")]),
     ("Fourth", [None, None, ("Closing Lead", "10:30–17:30", ""), ("Closing Lead", "10:30–17:30", ""), ("Closing Lead", "10:30–17:30", "")])]},
  "FULL": {"days": ["Mon", "Tue–Thu", "Fri", "Sat", "Sun"], "mult": [1, 3, 1, 1, 1], "rows": [
     ("Line", ["6:15–14:30", "6:15–14:30", "6:15–14:15", "7:15–15:15", "7:15–15:15"]),
     ("Lead", ["6:30–14:30", "6:30–14:30", "6:30–13:00", "7:30–14:00", "7:30–12:30"]),
     ("Third", [("Barista", "7:00–13:00", ""), ("Barista", "7:30–13:30", ""), ("Barista", "8:30–15:00", ""), ("Barista", "8:00–14:00", ""), ("Barista", "8:00–14:00", "")]),
     ("Fourth", [None, None, ("Closing Lead", "10:00–17:30", ""), ("Barista", "10:00–16:00", ""), ("Barista", "10:00–16:00", "")]),
     ("Fifth", [None, None, None, ("Closing Lead", "10:30–17:30", ""), ("Closing Lead", "10:30–17:30", "")])]}},
 "1899 Penn": {
  "BASE": {"days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "rows": [
     ("Lead", ["6:30–12:30", "6:30–14:30", "6:30–14:30", "6:30–14:30", "6:30–12:30"]),
     ("Production Barista", ["9:30–15:30", "7:30–15:30", "7:30–15:30", "7:30–15:30", "9:30–15:30"])]},
  "FULL": {"days": ["Mon", "Tue", "Wed", "Thu", "Fri"], "rows": [
     ("Lead", ["6:30–14:30", "6:30–15:00", "6:30–15:00", "6:30–15:00", "6:30–14:30"]),
     ("Production Barista", ["7:30–15:30"] * 5)]}},
}

def cell(role, v):
    if v is None: return None
    if isinstance(v, tuple): return v
    return (role, v, "")

def week_totals(cafe, key):
    spec = S[cafe][key]; mult = spec.get("mult", [1] * len(spec["days"])); H = 0.0; Wg = 0.0; per_day = []
    for j, d in enumerate(spec["days"]):
        dh = 0.0; dw = 0.0
        for role, vals in spec["rows"]:
            c = cell(role, vals[j])
            if c: r, sh, _ = c; hh = hours(sh); dh += hh; dw += hh * RATE[r]
        per_day.append((d, dh, dw, mult[j])); H += dh * mult[j]; Wg += dw * mult[j]
    return H, Wg, per_day

def grid(doc, cafe, key, label):
    spec = S[cafe][key]; H, Wg, per_day = week_totals(cafe, key)
    rows = []
    for role, vals in spec["rows"]:
        row = [role]
        for j in range(len(spec["days"])):
            c = cell(role, vals[j])
            if c is None: row.append("—")
            else:
                r, sh, mark = c; row.append((f"{r} " if role in ("Third", "Fourth", "Fifth") else "") + fmt(sh) + (f" {mark}" if mark else ""))
        rows.append(row)
    rows.append(["Hours"] + [(f"{dh:g}" + (" each" if mu > 1 else "")) for d, dh, dw, mu in per_day])
    n = len(spec["days"]); table(doc, [f"{label} · {H:g} h · {m(Wg)}"] + spec["days"], rows, widths=[1.35] + [(6.7 - 1.35) / n] * n, total_rows=[len(rows) - 1], size=9)
    return H, Wg

def build():
    doc = Document(); sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.9); sec.top_margin = sec.bottom_margin = Inches(0.8)
    doc.styles["Normal"].font.name = FONT_NAME; doc.styles["Normal"].font.size = Pt(10)
    for txt, size, bold, color, after in [("SLIPSTREAM COFFEE", 11, True, COLOR_MUTED, 0), ("Scheduling Guide", 22, True, COLOR_ACCENT, 0), ("Labor bands, and the two schedules that hit them  |  September 2026  |  v1.2", 13, True, None, 2)]:
        para = doc.add_paragraph(); r = para.add_run(txt); _set_run_font(r, size=size, bold=bold, color=color); _set_para_spacing(para, before=0, after=after)
    add_para(doc, "For whoever writes the schedule. The labor band is the target. Everything else here is support: two schedules per cafe (Base and Full), the weekly sales that decide which one, and what to do in a slow week.",
             italic=True, size=9, color=COLOR_MUTED, after=8)
    tot = {c: {k: week_totals(c, k)[:2] for k in ("BASE", "FULL")} for c in S}

    add_heading(doc, "1. Labor bands — the target", level=1)
    add_callout(doc, "Core labor as % of core sales, for the week",
                "Downtown 32–36%, max 38%.   Studio 30–34%, max 36%.   1899 Penn 35–40%, max 42%.   NMWA: staffing rule, no %.\n"
                "Inside the band = on target. Between the band and the max = OK in a slow week, use an admin day. Over the max = Ryan's approval before the schedule is published.\n"
                "Core = floor shifts only. Admin (Square title 'Manager') and Training shifts are not in the number — in Square Teams, deselect Manager and Training when you review the week and the total you see is the core total.",
                body_size=10, title_size=10)
    add_para(doc, "Which schedule, and what it costs, by the week's projected core sales (Section 5 of the weekly report):", after=4)
    def hw(c, k): H, Wg = tot[c][k]; return f"{H:g} h · {m(Wg)}"
    rows = [["Downtown", "under $7,500", f"BASE + admin day\n{hw('Downtown','BASE')}", "$7,500 – 9,000", f"FULL\n{hw('Downtown','FULL')}", "$9,000 +", f"FULL\n{hw('Downtown','FULL')}"],
            ["Studio", "under $10,500", f"BASE + admin day\n{hw('Studio','BASE')}", "$10,500 – 12,500", f"BASE\n{hw('Studio','BASE')}", "$12,500 +", f"FULL\n{hw('Studio','FULL')}"],
            ["1899 Penn", "under $3,300", f"BASE + admin day\n{hw('1899 Penn','BASE')}", "$3,300 – 4,200", f"BASE\n{hw('1899 Penn','BASE')}", "$4,200 +", f"FULL\n{hw('1899 Penn','FULL')}"],
            ["NMWA", "any", "1 person 9–3 (2 on community days)\n24 h · $510", "", "", "", ""]]
    table(doc, ["Cafe", "SLOW week", "schedule · hours", "MID week", "schedule · hours", "BUSY week", "schedule · hours"], rows, widths=[0.85, 0.85, 1.35, 0.95, 1.35, 0.7, 1.35], size=8)
    need = lambda c, k, pct: tot[c][k][1] / (pct / 100)
    add_caveat(doc, "Dollars at planning rates (Lead / Closing Lead $22.25, Line $19.75, Production Barista $20.25, Barista & Closer $20.00); your roster will differ a little, so hold the HOURS.")
    add_para(doc, "Where each schedule lands against the band — labor % at the edges of each sales range:", after=4)
    pc = lambda c, k, sales: f"{tot[c][k][1] / sales * 100:.0f}%"
    rows = []
    MIDSCHED = {"Downtown": "FULL", "Studio": "BASE", "1899 Penn": "BASE"}
    for c, (lo, hi) in BANDS.items():
        slow_typ = {"Downtown": 6566, "Studio": 9734, "1899 Penn": 2872}[c]; ms = MIDSCHED[c]
        rows.append([c, f"BASE at a typical slow week ({m(slow_typ)}): {pc(c,'BASE',slow_typ)}", f"{ms} at {m(lo)}: {pc(c,ms,lo)}", f"{ms} at {m(hi)}: {pc(c,ms,hi)}", f"FULL at {m(hi)}: {pc(c,'FULL',hi)}",
                     f"under max from {m(need(c,'BASE',{'Downtown':38,'Studio':36,'1899 Penn':42}[c]))}; in band from {m(need(c,'BASE',{'Downtown':36,'Studio':34,'1899 Penn':40}[c]))}"])
    table(doc, ["Cafe", "Slow week", "Bottom of mid", "Top of mid", "Bottom of busy", "BASE is…"], rows, widths=[0.8, 1.55, 1.05, 1.05, 1.05, 1.2], size=8)
    add_caveat(doc, "A slow week lands above the max on the Base schedule by design — that is a slow week, not a scheduling error: use the admin day, drop the ✂ shift that day, trim where the rush is thin, publish, and flag it.")

    add_heading(doc, "2. Every week", level=1)
    for lab, txt in [("Pick the column.", "Projected core sales for the coming Sun–Sat week from the weekly report. Core = walk-in, no catering or invoices."),
                     ("Copy the schedule.", "Base or Full, from Section 3. Do not go under Base: two openers at Downtown and Studio, two closers at Downtown and on Studio Mon–Thu, three people through the Studio morning, two shifts every day at 1899."),
                     ("Slow week?", "Base schedule, plus: put the admin day where it covers the floor and drop that day's ✂ shift; look for another hour or two where the rush is thin. Never cut an opener or the Studio mid-day third."),
                     ("Check core hours, publish.", "Deselect Manager and Training in Square Teams; the total should be at or under the hours in the column you picked. Over the max % → ask Ryan first.")]:
        add_bullet(doc, txt, bold_lead=lab)

    add_heading(doc, "3. The two schedules per cafe", level=1)
    add_heading(doc, "Downtown — 7a–4p Mon–Fri", level=2)
    grid(doc, "Downtown", "BASE", "BASE"); grid(doc, "Downtown", "FULL", "FULL")
    add_caveat(doc, "Three people Tue–Thu under $7,500 a week; four from $7,500 — never five. Tue, Wed and Thu are always the same schedule. Friday is the weak day and reads 45–60% on its own; the Friday admin day covers it. Federal-holiday Mondays: closed.")

    add_heading(doc, "Studio — 7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun", level=2)
    grid(doc, "Studio", "BASE", "BASE"); grid(doc, "Studio", "FULL", "FULL")
    add_caveat(doc, "Mon–Thu: Line on all day and closing the line, Lead, and a third person out by 1:30 (7:00–1:00 Monday for the delivery). Friday: two leads — the opener out at 1:00, the Closing Lead closes alone from 3:00 (second-to-last out at 3:00 in Base and Full); four people at most in any week. "
                    "Sat/Sun: one closer — the Closing Lead 10:30–5:30 closes alone; first out at 2:00; second-to-last out at 3:15 in Base (the Line) and 4:00 in Full. Sunday Full is five with the opening Lead on a short 7:30–12:30. Admin days (Mon/Wed, one weekend day) cover the ✂ shifts in slow weeks.")

    add_heading(doc, "1899 Penn — 7a–3p Mon–Fri", level=2)
    grid(doc, "1899 Penn", "BASE", "BASE"); grid(doc, "1899 Penn", "FULL", "FULL")
    add_caveat(doc, "Second person in at 7:30 Tue–Thu in both schedules, every day in Full. First out 2:30 in Base, 3:00 in Full. Mon and Fri stagger in Base (Lead opens alone 6:30–12:30, Production Barista 9:30–3:30 closes alone). Two shifts every day — never a solo day; whoever is on at 2:00 closes. The Tuesday admin day is the slow-week cover. Federal-holiday Mondays: closed.")

    add_heading(doc, "NMWA — Thu–Sun", level=2)
    add_para(doc, "One FOH shift 9:00–3:00 each open day (an hour to set up, an hour to close). Two shifts 9:00–3:00 on community days — the second Wednesday and the first Sunday. No labor % while sales ramp; a second person on any other day is an ask to Ryan.", after=6)

    add_heading(doc, "4. Rules that do not bend", level=1)
    for txt in ["Two openers (Line in 45 minutes before open, Lead 30) at Downtown and Studio; two closers at Downtown and on Studio Mon–Thu; Studio Fri–Sun the Closing Lead closes alone; closers out within 30 minutes of close.",
                "Six-hour minimum on a shift, with the two exceptions written into the grids (Studio weekday third 7:30–1:30 is six; the Sunday Full opening Lead 7:30–12:30 is five by design).",
                "Downtown: third person on by 8:00 Tue–Thu; four people at most. Studio: Line on all day Mon–Thu; three people through the morning; four max on Friday; five only on Full-week weekends. 1899: two shifts every day.",
                "ALLs and Shift Leads go in the Lead and Closing Lead slots, not the Barista slots — a $25 person on a Barista shift costs the week about $180.",
                "Admin (Manager) and Training are scheduled on top and reported separately — set the Square job title correctly or the shift books as floor labor and the week reads over.",
                "Known event or delivery day: add the hours, note it on the schedule. Federal-holiday Mondays: Downtown and 1899 closed; Studio runs its Monday base."]:
        add_bullet(doc, txt)
    add_caveat(doc, "Source: Slipstream Labor Deep Dive — Schedule Templates by Tier v1.4 (Sep 2026), with Ryan's shift edits of Sep 11. Questions on the numbers → Ryan.")
    _ensure_zoom_percent(doc); doc.save(OUT); print("Saved →", OUT)
    # data for the weekly report
    out = {"version": "guide v1.2", "rates": RATE, "bands_weekly_sales": BANDS, "labor_bands_pct": {"Downtown": [32, 36, 38], "Studio": [30, 34, 36], "1899 Penn": [35, 40, 42]},
           "which": {"Downtown": {"slow": "BASE", "mid": "FULL", "busy": "FULL"}, "Studio": {"slow": "BASE", "mid": "BASE", "busy": "FULL"}, "1899 Penn": {"slow": "BASE", "mid": "BASE", "busy": "FULL"}}, "cafes": {}}
    for c in S:
        out["cafes"][c] = {}
        for k in ("BASE", "FULL"):
            spec = S[c][k]; H, Wg, per_day = week_totals(c, k); days = {}
            for j, d in enumerate(spec["days"]):
                shifts = []
                for role, vals in spec["rows"]:
                    cc = cell(role, vals[j])
                    if cc: r, sh, mark = cc; shifts.append({"role": r, "shift": sh, "hours": hours(sh), "wages": round(hours(sh) * RATE[r], 2), "trim": mark == "✂"})
                for dd in (["Tue", "Wed", "Thu"] if d == "Tue–Thu" else [d]): days[dd] = {"hours": per_day[j][1], "wages": round(per_day[j][2], 2), "shifts": shifts}
            out["cafes"][c][k] = {"week_hours": H, "week_wages": round(Wg, 2), "days": days}
    out["cafes"]["NMWA"] = {"BASE": {"week_hours": 24, "week_wages": 510, "rule": "1 FOH 9:00–15:00 per open day; 2 on community days"}}
    json.dump(out, open(os.path.join(DD, "guide_schedules_v12.json"), "w"), indent=1)
    for c in S:
        for k in ("BASE", "FULL"): H, Wg, pd_ = week_totals(c, k); print(c, k, H, round(Wg), [(d, dh) for d, dh, dw, mu in pd_])

if __name__ == "__main__":
    build()
