#!/usr/bin/env python3
"""Build 'Slipstream Schedule Templates - Sep 2026 v1.0.xlsx' from templates.json.
Slipstream color protocol: blue = input (rates, targets, capacity, typical sales), green = pulled data
(need/actual curves, first-cut transcription), black = formulas/text. Hours, wages, labor %, SPLH, status,
sensitivity and trigger cells are live formulas.
"""
import os, json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Schedule Templates - Sep 2026 v1.0.xlsx")
T = json.load(open(os.path.join(DD, "templates.json"))); NV = json.load(open(os.path.join(DD, "need_vs_actual.json")))
S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))

BLUE = Font(name="Calibri", size=10, color="1F3864"); BLUE_FILL = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
YELLOW = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GREEN = Font(name="Calibri", size=10, color="006100"); BLACK = Font(name="Calibri", size=10, color="000000")
BOLD = Font(name="Calibri", size=10, bold=True); HDR = Font(name="Calibri", size=10, bold=True, color="1F3A5F")
HDR_FILL = PatternFill("solid", start_color="E8ECEF", end_color="E8ECEF"); TOT_FILL = PatternFill("solid", start_color="D9E2EC", end_color="D9E2EC")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F3A5F"); NOTE = Font(name="Calibri", size=9, italic=True, color="606060")
THIN = Side(style="thin", color="1F3A5F")
CAFES = ["Downtown", "Studio", "1899 Penn", "NMWA"]
DOW7 = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
ROLE_CODES = [("SL", "Shift Lead"), ("PL", "Production Line"), ("PB", "Production Barista"), ("B", "Barista"), ("HR", "Hospitality Register"), ("CL", "Closer (keyholder barista)"), ("FOH", "FOH")]
TIER_LABEL = {"slow": "Slow", "slow_lean": "Slow — lean", "mid": "Mid", "mid_lean": "Mid — lean", "busy": "Busy"}

def t2x(h):   # decimal hours -> Excel time fraction
    return h / 24.0

def hdr(ws, row, headers, widths=None):
    for j, t in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=t); c.font = HDR; c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = Border(bottom=THIN)
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[L(j)].width = w

def inp(c, fmt=None, key=False):
    c.font = BLUE; c.fill = YELLOW if key else BLUE_FILL
    if fmt: c.number_format = fmt

def pull(c, fmt=None):
    c.font = GREEN
    if fmt: c.number_format = fmt

def fx(c, fmt=None, bold=False):
    c.font = BOLD if bold else BLACK
    if fmt: c.number_format = fmt

def main():
    wb = Workbook()
    # ---------------- README ----------------
    ws = wb.active; ws.title = "README"
    ws["A1"] = "Slipstream Schedule Templates by Demand Tier — September 2026 v1.0"; ws["A1"].font = TITLE
    lines = [
        "Companion to 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.0.docx'. Built Sep 4 2026 from Square transactions (Jan–Aug 2026) and 861 Square timecards (14 sampled weeks).",
        "",
        "COLOR PROTOCOL — blue text on light blue = input you can change (rates, targets, capacity, typical sales; yellow = key assumptions); green = pulled from the analysis data (need/actual curves, first-cut transcription); black = formula or label.",
        "",
        "HOW TO USE",
        "1. Pick the tier for the week from the projected weekly core sales in the weekly report's forward block (bands on the Inputs sheet). Below the band floor → use the next slower template.",
        "2. Open the cafe sheet, copy the shifts for that tier into Square Teams. Notes flag where the ALL's Admin block is meant to cover a rush (staff 2 core, Admin covers) — drop the marked barista shift on those days.",
        "3. The day and week summaries recompute labor % from the Inputs rates and typical sales. Change a start/end time and everything flows. Status: OK = at/under target, BAND = between target and ceiling, OVER = above ceiling (ask Ryan before publishing).",
        "4. Summary sheet: every cafe × tier × day, weekly rollups, −10% sales sensitivity and the trigger (weekly sales below which the template breaches the ceiling).",
        "5. First Cut 10-11 sheet: Ryan's Square Teams week transcribed and scored on the same basis, with the specific change to each shift.",
        "",
        "BASIS — labor % = gross guaranteed wages ÷ core (ex-catering walk-in) sales, the basis of the 30/32/35 targets. The P&L shows ≈ 1.12 × gross − CC tips (see docx §1). Times are Eastern; shifts are open-to-clock-out, prep starts 30–60 min before open.",
        "",
        "CONSTRAINTS honoured — Downtown 7a–4p Mon–Fri (10-hour open-to-close lead shifts flagged as an option); Studio 7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun, 3 people at open on Monday (catering); 1899 7a–3p Mon–Fri; NMWA 10a–2p Thu–Sun + 2nd Wednesday, 1 person (2 on community days). Federal-holiday Mondays: Downtown and 1899 closed.",
    ]
    for i, t in enumerate(lines, 3):
        ws.cell(row=i, column=1, value=t).font = NOTE if t.startswith(("Companion", "COLOR")) else BLACK
        ws.cell(row=i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 150

    # ---------------- Inputs ----------------
    wi = wb.create_sheet("Inputs")
    wi["A1"] = "Inputs — rates, targets, capacity, tier bands, typical sales"; wi["A1"].font = TITLE
    wi["A3"] = "Standard role rates ($/hr)"; wi["A3"].font = BOLD
    hdr(wi, 4, ["Code", "Role", "Rate"], [10, 30, 12])
    rate_rows = {}
    for i, (code, name) in enumerate(ROLE_CODES):
        r = 5 + i
        wi.cell(row=r, column=1, value=code).font = BLACK; wi.cell(row=r, column=2, value=name).font = BLACK
        c = wi.cell(row=r, column=3, value=T["rates"][code]); inp(c, '"$"#,##0.00', key=True); rate_rows[code] = r
    RATE_RANGE = f"Inputs!$A$5:$C${4 + len(ROLE_CODES)}"
    r0 = 5 + len(ROLE_CODES) + 1
    wi.cell(row=r0, column=1, value="Load factor on gross wages (payroll tax/benefit allocation)").font = BLACK
    c = wi.cell(row=r0, column=3, value=1.12); inp(c, "0.00", key=True); LOAD_CELL = f"Inputs!$C${r0}"
    wi.cell(row=r0 + 1, column=1, value="Planning capacity (transactions per person-hour)").font = BLACK
    c = wi.cell(row=r0 + 1, column=3, value=T["cap"]); inp(c, "0.0", key=True)
    wi.cell(row=r0 + 2, column=1, value="Weeks per month").font = BLACK
    c = wi.cell(row=r0 + 2, column=3, value=52 / 12); inp(c, "0.000")
    WPM_CELL = f"Inputs!$C${r0 + 2}"
    # targets
    rt = r0 + 4
    wi.cell(row=rt, column=1, value="Core labor targets (% of core sales)").font = BOLD
    hdr(wi, rt + 1, ["Cafe", "Target", "Ceiling"], None)
    tgt_rows = {}
    for i, cafe in enumerate(["Downtown", "Studio", "1899 Penn"]):
        r = rt + 2 + i
        wi.cell(row=r, column=1, value=cafe).font = BLACK
        c = wi.cell(row=r, column=2, value=S["targets"][cafe]["target"]); inp(c, "0%", key=True)
        c = wi.cell(row=r, column=3, value=S["targets"][cafe]["ceiling"]); inp(c, "0%", key=True)
        tgt_rows[cafe] = r
    wi.cell(row=rt + 5, column=1, value="NMWA").font = BLACK; wi.cell(row=rt + 5, column=2, value="1 person/day; 2 on community days (no %)").font = BLACK
    TGT_RANGE = f"Inputs!$A${rt + 2}:$C${rt + 4}"
    # tier bands
    rb = rt + 8
    wi.cell(row=rb, column=1, value="Tier bands — weekly core (walk-in) sales").font = BOLD
    hdr(wi, rb + 1, ["Cafe", "Slow below", "Mid from", "Busy from"], None)
    for i, cafe in enumerate(["Downtown", "Studio", "1899 Penn"]):
        b = S["tiers"][cafe]["bands_weekly_walkin"]; r = rb + 2 + i
        wi.cell(row=r, column=1, value=cafe).font = BLACK
        inp(wi.cell(row=r, column=2, value=b["slow"][1]), '"$"#,##0'); inp(wi.cell(row=r, column=3, value=b["mid"][0]), '"$"#,##0'); inp(wi.cell(row=r, column=4, value=b["busy"][0]), '"$"#,##0')
    # typical sales per cafe × tier × dow
    rs = rb + 7
    wi.cell(row=rs, column=1, value="Typical daily core sales by tier (walk-in $, per open day; from 2026 transactions) — change these to re-score every template").font = BOLD
    hdr(wi, rs + 1, ["Cafe", "Tier"] + DOW7 + ["Week"], None)
    sales_row = {}
    r = rs + 2
    for cafe in CAFES:
        tiers = ["slow", "mid", "busy"] if cafe != "NMWA" else ["slow"]
        for tier in tiers:
            wi.cell(row=r, column=1, value=cafe).font = BLACK; wi.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK
            for j, d in enumerate(DOW7):
                v = None
                if cafe == "NMWA":
                    v = S["tiers"].get("1899 Penn") and None
                    td = json.load(open(os.path.join(DD, "tiers.json")))["NMWA"]["typical_day"]["slow"]
                    v = td[d]["net"] if d in td else None
                else:
                    td = S["tiers"][cafe]["typical_day"][tier]
                    v = td[d]["net"] if d in td else None
                c = wi.cell(row=r, column=3 + j, value=v)
                if v is not None: inp(c, '"$"#,##0')
            c = wi.cell(row=r, column=10, value=f"=SUM(C{r}:I{r})"); fx(c, '"$"#,##0', bold=True)
            sales_row[(cafe, tier)] = r
            r += 1
    for col, w in zip("ABCDEFGHIJ", [44, 12, 11, 11, 11, 11, 11, 11, 11, 12]):
        wi.column_dimensions[col].width = w

    # ---------------- cafe sheets ----------------
    summary_refs = []     # (cafe, tier, dow, hours_cell, wages_cell, sales_ref)
    for cafe in CAFES:
        w = wb.create_sheet(cafe)
        w["A1"] = f"{cafe} — schedule templates by tier"; w["A1"].font = TITLE
        w["A2"] = ("Rows = shifts. Start/End are inputs (blue); Hours, Rate (from Inputs), Wages and Loaded cost are formulas. Day summaries below each tier. "
                   "'Admin covers' notes mark the deliberate flex: staff 2 core and let the ALL's Admin block absorb the rush."); w["A2"].font = NOTE
        hdr(w, 4, ["Tier", "Day", "Role", "Code", "Start", "End", "Hours", "Rate", "Wages", "Loaded cost", "Notes"], [12, 6, 26, 6, 8, 8, 8, 9, 10, 11, 70])
        r = 5
        tiers = [t for t in T["cafes"][cafe] if t in TIER_LABEL]
        for tier in tiers:
            o = T["cafes"][cafe][tier]
            first_row = r
            day_rows = {}
            for d, dd in o["days"].items():
                d_start = r
                for s in dd["shifts"]:
                    w.cell(row=r, column=1, value=TIER_LABEL[tier]).font = BLACK; w.cell(row=r, column=2, value=d).font = BLACK
                    w.cell(row=r, column=3, value=s["role"]).font = BLACK; w.cell(row=r, column=4, value=s["code"]).font = BLACK
                    hh, mm = map(int, s["start"].split(":")); inp(w.cell(row=r, column=5, value=dt.time(hh, mm)), "h:mm")
                    hh, mm = map(int, s["end"].split(":")); inp(w.cell(row=r, column=6, value=dt.time(hh, mm)), "h:mm")
                    fx(w.cell(row=r, column=7, value=f"=(F{r}-E{r})*24"), "0.00")
                    fx(w.cell(row=r, column=8, value=f"=VLOOKUP(D{r},{RATE_RANGE},3,FALSE)"), '"$"#,##0.00')
                    fx(w.cell(row=r, column=9, value=f"=G{r}*H{r}"), '"$"#,##0')
                    fx(w.cell(row=r, column=10, value=f"=I{r}*{LOAD_CELL}"), '"$"#,##0')
                    w.cell(row=r, column=11, value=s["note"]).font = BLACK
                    r += 1
                day_rows[d] = (d_start, r - 1)
            # day summary block for this tier
            r += 1
            w.cell(row=r, column=1, value=f"{TIER_LABEL[tier]} — day summary").font = BOLD; r += 1
            hdr(w, r, ["Tier", "Day", "Shifts", "", "", "", "Hours", "", "Wages", "Loaded", "Typical sales | Labor % | SPLH | Target | Status"]); r += 1
            sum_start = r
            base = tier.split("_")[0]
            srow = sales_row.get((cafe, base), sales_row.get((cafe, "slow")))
            for d in o["days"].keys():
                a, b = day_rows[d]
                w.cell(row=r, column=1, value=TIER_LABEL[tier]).font = BLACK; w.cell(row=r, column=2, value=d).font = BLACK
                fx(w.cell(row=r, column=3, value=f"=COUNTA(C{a}:C{b})"), "0")
                fx(w.cell(row=r, column=7, value=f"=SUM(G{a}:G{b})"), "0.00", bold=True)
                fx(w.cell(row=r, column=9, value=f"=SUM(I{a}:I{b})"), '"$"#,##0', bold=True)
                fx(w.cell(row=r, column=10, value=f"=SUM(J{a}:J{b})"), '"$"#,##0')
                dcol = L(3 + DOW7.index(d))
                sales_ref = f"Inputs!${dcol}${srow}"
                fx(w.cell(row=r, column=11, value=f"={sales_ref}"), '"$"#,##0')
                fx(w.cell(row=r, column=12, value=f"=IF(K{r}>0,I{r}/K{r},\"\")"), "0.0%", bold=True)
                fx(w.cell(row=r, column=13, value=f"=IF(G{r}>0,K{r}/G{r},\"\")"), '"$"#,##0')
                if cafe != "NMWA":
                    fx(w.cell(row=r, column=14, value=f"=VLOOKUP(\"{cafe}\",{TGT_RANGE},2,FALSE)"), "0%")
                    fx(w.cell(row=r, column=15, value=f"=IF(L{r}=\"\",\"\",IF(L{r}<=N{r},\"OK\",IF(L{r}<=VLOOKUP(\"{cafe}\",{TGT_RANGE},3,FALSE),\"BAND\",\"OVER\")))"))
                summary_refs.append((cafe, tier, d, f"'{cafe}'!G{r}", f"'{cafe}'!I{r}", sales_ref, f"'{cafe}'!C{r}"))
                r += 1
            # week total
            w.cell(row=r, column=1, value=TIER_LABEL[tier]).font = BOLD; w.cell(row=r, column=2, value="WEEK").font = BOLD
            fx(w.cell(row=r, column=3, value=f"=SUM(C{sum_start}:C{r-1})"), "0", bold=True)
            fx(w.cell(row=r, column=7, value=f"=SUM(G{sum_start}:G{r-1})"), "0.00", bold=True)
            fx(w.cell(row=r, column=9, value=f"=SUM(I{sum_start}:I{r-1})"), '"$"#,##0', bold=True)
            fx(w.cell(row=r, column=10, value=f"=SUM(J{sum_start}:J{r-1})"), '"$"#,##0', bold=True)
            fx(w.cell(row=r, column=11, value=f"=SUM(K{sum_start}:K{r-1})"), '"$"#,##0', bold=True)
            fx(w.cell(row=r, column=12, value=f"=IF(K{r}>0,I{r}/K{r},\"\")"), "0.0%", bold=True)
            fx(w.cell(row=r, column=13, value=f"=IF(G{r}>0,K{r}/G{r},\"\")"), '"$"#,##0', bold=True)
            if cafe != "NMWA":
                fx(w.cell(row=r, column=14, value=f"=VLOOKUP(\"{cafe}\",{TGT_RANGE},2,FALSE)"), "0%")
                fx(w.cell(row=r, column=15, value=f"=IF(L{r}=\"\",\"\",IF(L{r}<=N{r},\"OK\",IF(L{r}<=VLOOKUP(\"{cafe}\",{TGT_RANGE},3,FALSE),\"BAND\",\"OVER\")))"), bold=True)
            for col in range(1, 16):
                w.cell(row=r, column=col).fill = TOT_FILL
            summary_refs.append((cafe, tier, "WEEK", f"'{cafe}'!G{r}", f"'{cafe}'!I{r}", f"'{cafe}'!K{r}", f"'{cafe}'!C{r}"))
            r += 3
        for col, wd in zip("KLMNO", [12, 10, 9, 9, 9]):
            w.column_dimensions[col].width = max(w.column_dimensions[col].width or 0, wd)
        w.column_dimensions["K"].width = 70
        w.freeze_panes = "A5"

    # ---------------- Summary ----------------
    ws = wb.create_sheet("Summary", 1)
    ws["A1"] = "Summary — projected core labor % and SPLH per cafe × tier × day vs target"; ws["A1"].font = TITLE
    ws["A2"] = "All values are formulas off the cafe sheets and Inputs. Sensitivity: labor % if sales land 10% under the tier's typical level with the template unchanged. Trigger: weekly sales below which the week's template breaches the ceiling — drop to the next slower template."; ws["A2"].font = NOTE
    hdr(ws, 4, ["Cafe", "Tier", "Day", "Shifts", "Core hours", "Wages", "Typical sales", "Labor %", "SPLH", "Target", "Ceiling", "Status", "Labor % at −10% sales", "Ceiling trigger (wk sales)", "On-target at (wk sales)"],
        [12, 12, 7, 7, 10, 10, 12, 9, 8, 8, 8, 8, 12, 14, 14])
    r = 5
    for (cafe, tier, d, hcell, wcell, scell, ccell) in summary_refs:
        ws.cell(row=r, column=1, value=cafe).font = BLACK; ws.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK; ws.cell(row=r, column=3, value=d).font = BLACK
        fx(ws.cell(row=r, column=4, value=f"={ccell}"), "0")
        fx(ws.cell(row=r, column=5, value=f"={hcell}"), "0.00"); fx(ws.cell(row=r, column=6, value=f"={wcell}"), '"$"#,##0'); fx(ws.cell(row=r, column=7, value=f"={scell}"), '"$"#,##0')
        fx(ws.cell(row=r, column=8, value=f"=IF(G{r}>0,F{r}/G{r},\"\")"), "0.0%", bold=True); fx(ws.cell(row=r, column=9, value=f"=IF(E{r}>0,G{r}/E{r},\"\")"), '"$"#,##0')
        if cafe != "NMWA":
            fx(ws.cell(row=r, column=10, value=f"=VLOOKUP(A{r},{TGT_RANGE},2,FALSE)"), "0%"); fx(ws.cell(row=r, column=11, value=f"=VLOOKUP(A{r},{TGT_RANGE},3,FALSE)"), "0%")
            fx(ws.cell(row=r, column=12, value=f"=IF(H{r}=\"\",\"\",IF(H{r}<=J{r},\"OK\",IF(H{r}<=K{r},\"BAND\",\"OVER\")))"))
            fx(ws.cell(row=r, column=13, value=f"=IF(G{r}>0,F{r}/(G{r}*0.9),\"\")"), "0.0%")
            if d == "WEEK":
                fx(ws.cell(row=r, column=14, value=f"=F{r}/K{r}"), '"$"#,##0'); fx(ws.cell(row=r, column=15, value=f"=F{r}/J{r}"), '"$"#,##0')
        if d == "WEEK":
            for col in range(1, 16): ws.cell(row=r, column=col).fill = TOT_FILL
            ws.cell(row=r, column=3).font = BOLD
        r += 1
    ws.conditional_formatting.add(f"L5:L{r}", CellIsRule(operator="equal", formula=['"OVER"'], font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(f"L5:L{r}", CellIsRule(operator="equal", formula=['"OK"'], font=Font(color="006100", bold=True)))
    ws.freeze_panes = "A5"

    # ---------------- First cut ----------------
    wf = wb.create_sheet("First Cut 10-11")
    wf["A1"] = "Ryan's first-cut template — Square Teams week of Sun 10/11 – Sat 10/17 (core shifts; Manager/Commissary/Roasting excluded), scored on the same basis"; wf["A1"].font = TITLE
    wf["A2"] = ("Green = transcribed from the screenshots (team member names omitted; rates from the Square Teams $ totals). Hours/wages are formulas. 'Template change' is the specific edit that turns this shift into the Mid template; "
                "Slow-tier changes in the docx §3e/4e/5e. 10/12 is Columbus Day: Downtown and 1899 closed — unschedule Monday that week."); wf["A2"].font = NOTE
    hdr(wf, 4, ["Cafe", "Day", "Role", "Start", "End", "Hours", "Rate (as staffed)", "Wages", "Template change (Mid)"], [12, 6, 22, 8, 8, 8, 12, 10, 95])
    FC = {
        "Downtown": [
            ("Mon", "Production Barista", 6.5, 13.0, 19.00, "→ Production Line 6:00–2:00 opens prep; this shift is not needed on Monday (2-person day + closer)"),
            ("Mon", "Production Line", 6.2, 16.5, 22.75, "10.3 h open-to-close → 6:00–2:00; the Closer (9:00–4:30, barista rate) closes"),
            ("Mon", "Shift Lead", 9.0, 16.5, 21.75, "Lead should open: Shift Lead 6:30–3:00; a Closer 9:00–4:30 replaces the 9–4:30 lead"),
            ("Tue", "Production Line", 6.25, 14.0, 18.50, "→ 6:00–2:00"),
            ("Tue", "Production Barista", 7.5, 16.5, 22.75, "9 h → 4 h rush barista 7:30–11:30 (drop entirely when Admin covers 8–10)"),
            ("Tue", "Shift Lead", 6.5, 15.0, 21.75, "→ 6:30–3:30 (lead stays through the 2–3:30p afternoon)"),
            ("Tue", "Production Barista", 9.0, 16.5, 25.25, "→ Closer 8:30–4:30 at barista rate ($19.50), not a $25.25 team member"),
            ("Wed", "Shift Lead", 8.0, 16.5, 19.00, "Second lead → Closer 8:30–4:30 (barista rate)"),
            ("Wed", "Production Barista", 6.5, 13.5, 21.75, "→ Shift Lead opens 6:30–3:30 (this person is the lead)"),
            ("Wed", "Production Line", 6.25, 15.0, 19.25, "→ 6:00–2:00"),
            ("Wed", "Barista", 9.0, 16.5, 25.25, "→ rush barista 7:30–12:00 (4.5 h) at $19; the 9–4:30 slot is the Closer"),
            ("Thu", "Shift Lead", 8.0, 16.5, 19.00, "Second lead → Closer 8:30–4:30"),
            ("Thu", "Barista", 6.5, 13.5, 21.75, "→ Shift Lead opens 6:30–3:30"),
            ("Thu", "Production Line", 6.25, 15.0, 19.25, "→ 6:00–2:00"),
            ("Thu", "Production Barista", 9.0, 16.5, 25.25, "→ rush barista 7:30–11:30 (4 h) at $19"),
            ("Fri", "Production Barista", 9.5, 16.5, 19.00, "→ Closer 9:00–4:30"),
            ("Fri", "Shift Lead", 6.5, 13.0, 21.75, "→ 6:30–3:00 (lead covers the afternoon; slow tier: 6:30–4:30 open-to-close, 2-person day)"),
            ("Fri", "Production Line", 6.25, 16.5, 19.25, "10.25 h → 6:00–1:00; the lead/closer close"),
        ],
        "Studio": [
            ("Sun", "Barista", 7.5, 14.0, 18.50, "keep"), ("Sun", "Shift Lead", 8.0, 15.0, 24.75, "keep (7:30–3:30 in template)"), ("Sun", "Production Line", 7.25, 15.5, 18.50, "keep (7:00–3:00)"), ("Sun", "Production Barista", 10.5, 17.5, 18.50, "keep as Closer 10:00–5:30"),
            ("Mon", "Shift Lead", 6.5, 14.5, 20.75, "→ 6:30–2:45"), ("Mon", "Production Barista", 6.5, 13.0, 19.00, "Admin (Manager 7–2) is on at open → rush barista 8:00–11:30 instead of 6:30–1:00"), ("Mon", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–2:00"),
            ("Tue", "Production Barista", 6.5, 13.0, 20.25, "DROP — Tuesday is a 2-person core day (Admin covers 8–10 when scheduled)"), ("Tue", "Shift Lead", 7.0, 14.5, 24.75, "→ 6:30–2:45 (lead opens)"), ("Tue", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–2:00"),
            ("Wed", "Barista", 6.5, 13.0, 19.00, "→ rush barista 8:00–12:00 (mid); DROP in slow weeks"), ("Wed", "Shift Lead", 7.0, 14.5, 22.75, "→ 6:30–2:45"), ("Wed", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–2:00"),
            ("Thu", "Shift Lead", 7.0, 14.5, 20.25, "→ 6:30–2:45"), ("Thu", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–2:00"), ("Thu", "Production Barista", 6.5, 13.5, 24.75, "Admin (Manager 9–2) covers → DROP; otherwise rush barista 8:00–12:00 at $19"),
            ("Fri", "Shift Lead", 6.5, 15.0, 20.25, "→ 6:30–3:00"), ("Fri", "Production Line", 6.25, 14.5, 22.75, "→ 6:00–2:00 at $19 (this is a $22.75 team member on a PL shift)"), ("Fri", "Shift Lead", 10.0, 17.25, 20.75, "Second lead → Closer 10:30–5:30 at $19.50"), ("Fri", "Barista", 7.0, 13.0, 18.50, "→ 8:00–1:00"),
            ("Sat", "Production Line", 7.25, 15.0, 18.50, "→ 7:00–3:00"), ("Sat", "Shift Lead", 10.0, 17.5, 18.50, "→ Closer 10:30–5:30"), ("Sat", "Shift Lead", 7.5, 13.0, 20.75, "→ Shift Lead 7:30–3:30 (one lead through the peak)"), ("Sat", "Hospitality Register", 8.0, 15.0, 22.75, "→ 8:30–4:30 so 4–5p is not a single person"), ("Sat", "Barista", 9.0, 16.0, 18.50, "→ 9:30–2:30 peak 4th"),
        ],
        "1899 Penn": [
            ("Mon", "Shift Lead", 6.5, 15.5, 20.50, "keep"), ("Mon", "Production Barista", 6.5, 13.5, 22.75, "Admin (Manager 8–2) on → rush only 8:00–12:00; else 7:30–1:30"),
            ("Tue", "Shift Lead", 6.5, 15.5, 20.50, "keep"), ("Tue", "Production Barista", 6.5, 15.5, 18.50, "→ 7:30–2:30 (Admin 6:30–1:30 makes three people on the morning — if kept, PB 7:30–1:30)"),
            ("Wed", "Shift Lead", 6.5, 15.5, 20.75, "keep"), ("Wed", "Production Barista", 6.5, 15.5, 18.50, "→ 7:30–2:30"),
            ("Thu", "Shift Lead", 6.5, 15.5, 22.75, "keep"), ("Thu", "Barista", 6.5, 15.0, 18.50, "→ 7:30–2:30"),
            ("Fri", "Shift Lead", 6.5, 15.5, 20.75, "keep"), ("Fri", "Production Barista", 6.5, 13.5, 22.75, "→ 8:00–12:00 rush only (lean: none; ALL covers 8:30–10 and 12:30–1:30)"),
        ],
        "NMWA": [
            ("Sun", "FOH", 9.0, 15.0, 20.25, "→ 9:30–2:30 (open 10–2); add 2nd FOH 10–2 on the FIRST Sunday of the month"),
            ("Wed", "(none scheduled)", 0, 0, 0, "10/14 is the 2nd Wednesday = community day: ADD FOH 9:30–2:30 + FOH 10:00–2:00"),
            ("Thu", "FOH", 9.0, 15.0, 20.50, "→ 9:30–2:30"), ("Fri", "FOH", 9.0, 15.0, 22.75, "→ 9:30–2:30 (Manager 10–2 block that day is admin, outside core)"), ("Sat", "FOH", 9.0, 15.0, 20.50, "→ 9:30–2:30"),
        ],
    }
    r = 5
    fc_rows = {}
    for cafe, shifts in FC.items():
        start = r
        for (d, role, s, e, rate, chg) in shifts:
            wf.cell(row=r, column=1, value=cafe).font = BLACK; pull(wf.cell(row=r, column=2, value=d)); pull(wf.cell(row=r, column=3, value=role))
            if e > 0:
                pull(wf.cell(row=r, column=4, value=dt.time(int(s), int(round((s % 1) * 60)))), "h:mm"); pull(wf.cell(row=r, column=5, value=dt.time(int(e), int(round((e % 1) * 60)))), "h:mm")
                fx(wf.cell(row=r, column=6, value=f"=(E{r}-D{r})*24"), "0.00"); pull(wf.cell(row=r, column=7, value=rate), '"$"#,##0.00'); fx(wf.cell(row=r, column=8, value=f"=F{r}*G{r}"), '"$"#,##0')
            wf.cell(row=r, column=9, value=chg).font = BLACK
            r += 1
        wf.cell(row=r, column=1, value=cafe).font = BOLD; wf.cell(row=r, column=2, value="WEEK").font = BOLD
        fx(wf.cell(row=r, column=6, value=f"=SUM(F{start}:F{r-1})"), "0.00", bold=True); fx(wf.cell(row=r, column=8, value=f"=SUM(H{start}:H{r-1})"), '"$"#,##0', bold=True)
        for col in range(1, 10): wf.cell(row=r, column=col).fill = TOT_FILL
        fc_rows[cafe] = r
        r += 2
    # comparison block
    r += 1
    wf.cell(row=r, column=1, value="First cut vs templates — weekly core hours and wages").font = BOLD; r += 1
    hdr(wf, r, ["Cafe", "First-cut hours", "First-cut wages", "Mid template hours", "Mid template wages", "Slow template hours", "Slow template wages", "Δ hours vs Mid", "Δ wages vs Mid"]); r += 1
    for cafe in ["Downtown", "Studio", "1899 Penn", "NMWA"]:
        mid_week = [x for x in summary_refs if x[0] == cafe and x[1] == ("mid" if cafe != "NMWA" else "slow") and x[2] == "WEEK"][0]
        slow_week = [x for x in summary_refs if x[0] == cafe and x[1] == "slow" and x[2] == "WEEK"][0]
        wf.cell(row=r, column=1, value=cafe).font = BLACK
        fx(wf.cell(row=r, column=2, value=f"=F{fc_rows[cafe]}"), "0.0"); fx(wf.cell(row=r, column=3, value=f"=H{fc_rows[cafe]}"), '"$"#,##0')
        fx(wf.cell(row=r, column=4, value=f"={mid_week[3]}"), "0.0"); fx(wf.cell(row=r, column=5, value=f"={mid_week[4]}"), '"$"#,##0')
        fx(wf.cell(row=r, column=6, value=f"={slow_week[3]}"), "0.0"); fx(wf.cell(row=r, column=7, value=f"={slow_week[4]}"), '"$"#,##0')
        fx(wf.cell(row=r, column=8, value=f"=D{r}-B{r}"), "+0.0;-0.0"); fx(wf.cell(row=r, column=9, value=f"=E{r}-C{r}"), '+"$"#,##0;-"$"#,##0')
        r += 1
    wf.cell(row=r + 1, column=1, value="Note: first-cut wages use the rates of the people scheduled ($18.50–25.25); template wages use standard rates ($21 lead / $19 barista / $19.50 closer). The rate gap is the cost-per-hour lever (docx §3c–5c).").font = NOTE
    wf.freeze_panes = "A5"

    # ---------------- Need vs Actual ----------------
    wn = wb.create_sheet("Need vs Actual")
    wn["A1"] = "Half-hour demand, people needed and actual core headcount — per cafe × tier × day (sampled weeks)"; wn["A1"].font = TITLE
    wn["A2"] = f"need = max(floor, ceil(txns/hr ÷ {T['cap']})); actual = average core headcount on the clock in the sampled weeks of that tier. Green = analysis data."; wn["A2"].font = NOTE
    grid = [x["t"] for x in NV["cafes"]["Downtown"]["tiers"]["slow"]["Tue"]["by_halfhour"]]
    allt = sorted({x["t"] for c in NV["cafes"].values() for tt in c["tiers"].values() for dd in tt.values() for x in dd["by_halfhour"]})
    def tl(t):
        hh = int(t); return f"{hh}:{'30' if t % 1 else '00'}"
    hdr(wn, 4, ["Cafe", "Tier", "Day", "Series"] + [tl(t) for t in allt] + ["Day total"], [12, 8, 6, 9] + [6] * len(allt) + [9])
    r = 5
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        for tier in ["slow", "mid", "busy"]:
            for d, rec in NV["cafes"][cafe]["tiers"][tier].items():
                by = {x["t"]: x for x in rec["by_halfhour"]}
                for series, key, fmt in [("txns/hr", "txns_hr", "0"), ("need", "need", "0"), ("actual", "actual", "0.0")]:
                    wn.cell(row=r, column=1, value=cafe).font = BLACK; wn.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK
                    wn.cell(row=r, column=3, value=d).font = BLACK; wn.cell(row=r, column=4, value=series).font = BLACK
                    for j, t in enumerate(allt):
                        v = by.get(t, {}).get(key)
                        if v is not None: pull(wn.cell(row=r, column=5 + j, value=v), fmt)
                    if series != "txns/hr":
                        fx(wn.cell(row=r, column=5 + len(allt), value=f"=SUM({L(5)}{r}:{L(4+len(allt))}{r})*0.5"), "0.0")
                    r += 1
    wn.freeze_panes = "E5"

    wb.save(OUT)
    print("Saved →", OUT)

if __name__ == "__main__":
    main()
