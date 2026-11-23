#!/usr/bin/env python3
"""Build 'Slipstream Schedule Templates - Sep 2026 v1.1.xlsx' from templates_v11.json (primary = no-10h variant).
Slipstream color protocol: blue = input (rates, targets, capacity, typical sales, start/end times), green = pulled data
(need/actual curves, first-cut transcription, simulation results), black = formulas/labels.
Sheets: README, Summary, Inputs, Downtown, Studio, 1899 Penn, NMWA, First Cut 10-11, v1.0 vs v1.1, Variants, Need vs Actual, Incentive.
"""
import os, json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Schedule Templates - Sep 2026 v1.1.xlsx")
T = json.load(open(os.path.join(DD, "templates_v11.json"))); T0 = json.load(open(os.path.join(DD, "templates.json")))
NV = json.load(open(os.path.join(DD, "need_v11.json"))); S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
INC = json.load(open(os.path.join(DD, "incentive_v11.json"))); TIERS = json.load(open(os.path.join(DD, "tiers.json"))); STAFF = json.load(open(os.path.join(DD, "staffing.json")))
MAIN = "no10h"

BLUE = Font(name="Calibri", size=10, color="1F3864"); BLUE_FILL = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
YELLOW = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GREEN = Font(name="Calibri", size=10, color="006100"); BLACK = Font(name="Calibri", size=10, color="000000")
BOLD = Font(name="Calibri", size=10, bold=True); HDR = Font(name="Calibri", size=10, bold=True, color="1F3A5F")
HDR_FILL = PatternFill("solid", start_color="E8ECEF", end_color="E8ECEF"); TOT_FILL = PatternFill("solid", start_color="D9E2EC", end_color="D9E2EC")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F3A5F"); NOTE = Font(name="Calibri", size=9, italic=True, color="606060")
THIN = Side(style="thin", color="1F3A5F")
CAFES = ["Downtown", "Studio", "1899 Penn", "NMWA"]; OPEN3 = ["Downtown", "Studio", "1899 Penn"]
DOW7 = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
ROLE_CODES = [("SL", "Shift Lead"), ("PL", "Production Line"), ("PB", "Production Barista"), ("B", "Barista"), ("HR", "Hospitality Register"), ("CL", "Closer (keyholder barista)"), ("FOH", "FOH")]
TIER_LABEL = {"slow": "Slow", "mid": "Mid", "busy": "Busy"}
RATES = dict(T["rates"]); RATES["FOH"] = 20.50

def hdr(ws, row, headers, widths=None):
    for j, t in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=t); c.font = HDR; c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = Border(bottom=THIN)
    if widths:
        for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w
def inp(c, fmt=None, key=False):
    c.font = BLUE; c.fill = YELLOW if key else BLUE_FILL
    if fmt: c.number_format = fmt
def pull(c, fmt=None):
    c.font = GREEN
    if fmt: c.number_format = fmt
def fx(c, fmt=None, bold=False):
    c.font = BOLD if bold else BLACK
    if fmt: c.number_format = fmt
def tm(s):
    hh, mm = map(int, s.split(":")); return dt.time(hh, mm)
def dec2time(x):
    return dt.time(int(x), int(round((x % 1) * 60)))

def main():
    wb = Workbook()
    # ---------------- README ----------------
    ws = wb.active; ws.title = "README"
    ws["A1"] = "Slipstream Schedule Templates by Demand Tier — September 2026 v1.1"; ws["A1"].font = TITLE
    lines = [
        "Companion to 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.1.docx'. v1.1 supersedes v1.0 (same day): templates are minimum-cost shift sets solved against the demand curves; capacity is per cafe (Downtown 14 / Studio 12 / 1899 14 transactions per person-hour); every template is priced at standard rates and at the roster's current average rate.",
        "",
        "COLOR PROTOCOL — blue text on light blue = input you can change (rates, targets, capacity, typical sales, shift start/end; yellow = key assumptions); green = pulled from the analysis data (need/actual curves, first-cut transcription, simulation results); black = formula or label.",
        "",
        "HOW TO USE",
        "1. Pick the tier for the week from the projected weekly core sales in the weekly report's forward block (bands on the Inputs sheet). Below the band floor → use the next slower template. Build the 10/11 week on MID.",
        "2. Open the cafe sheet, copy the shifts for that tier into Square Teams. 'Solo' notes mark the windows one person runs the floor; 'Admin covers' marks half-hours where the need exceeds the template and the ALL's block is the intended cover.",
        "3. The day and week summaries recompute labor % from the Inputs rates and typical sales. Change a start/end time and everything flows. Status: OK = at/under target, BAND = between target and ceiling, OVER = above ceiling (ask Ryan before publishing). The 'roster' % re-prices the same hours at the cafe's current average core rate (Inputs).",
        "4. Summary sheet: every cafe × tier × day, weekly rollups, −10% sales sensitivity and the ceiling trigger (weekly sales below which the template breaches the ceiling).",
        "5. First Cut 10-11: Ryan's Square Teams week transcribed and scored on the same basis, with the specific change to each shift (Mid template). v1.0 vs v1.1: what moved between the two issues. Variants: the price of each rule relaxation (10-hour lead, comfort floors, 6-hour minimum shifts, 8-hour Production Line).",
        "6. Incentive: plan comparison, the annual simulation, and a live Beat Share calculator (enter a month's sales and core wages per cafe).",
        "",
        "BASIS — labor % = gross guaranteed wages ÷ core (ex-catering walk-in) sales, the basis of the 30/32/35 targets. The P&L shows ≈ 1.12 × gross − CC tips (docx §1). Times are Eastern; shifts are open-to-clock-out; prep starts 30–60 min before open.",
        "",
        "CONSTRAINTS honoured — Downtown 7a–4p Mon–Fri; Studio 7a–2p Mon–Thu, 7a–5p Fri, 8a–5p Sat–Sun, 3 people at open on Monday (catering); 1899 7a–3p Mon–Fri; NMWA 10a–2p Thu–Sun + 2nd Wednesday, 1 person (2 on community days). No shift over 9 h; 4-hour minimum; a keyholder on the floor open−30 min to close. Federal-holiday Mondays: Downtown and 1899 closed.",
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
    for i, (code, name) in enumerate(ROLE_CODES):
        r = 5 + i
        wi.cell(row=r, column=1, value=code).font = BLACK; wi.cell(row=r, column=2, value=name).font = BLACK
        inp(wi.cell(row=r, column=3, value=RATES[code]), '"$"#,##0.00', key=True)
    RATE_RANGE = f"Inputs!$A$5:$C${4 + len(ROLE_CODES)}"
    r0 = 5 + len(ROLE_CODES) + 1
    wi.cell(row=r0, column=1, value="Load factor on gross wages (payroll tax/benefit allocation)").font = BLACK
    inp(wi.cell(row=r0, column=3, value=1.12), "0.00", key=True); LOAD_CELL = f"Inputs!$C${r0}"
    wi.cell(row=r0 + 1, column=1, value="Weeks per month").font = BLACK; inp(wi.cell(row=r0 + 1, column=3, value=52 / 12), "0.000")
    # roster rates + capacity + targets per cafe
    rt = r0 + 3
    wi.cell(row=rt, column=1, value="Per-cafe assumptions").font = BOLD
    hdr(wi, rt + 1, ["Cafe", "Target", "Ceiling", "Roster avg core rate ($/hr)", "Planning capacity (txns per person-hour)"], None)
    tgt_rows = {}
    for i, cafe in enumerate(OPEN3):
        r = rt + 2 + i
        wi.cell(row=r, column=1, value=cafe).font = BLACK
        inp(wi.cell(row=r, column=2, value=S["targets"][cafe]["target"]), "0%", key=True); inp(wi.cell(row=r, column=3, value=S["targets"][cafe]["ceiling"]), "0%", key=True)
        inp(wi.cell(row=r, column=4, value=STAFF["rolemix"][cafe]["avg_core_rate"]), '"$"#,##0.00'); inp(wi.cell(row=r, column=5, value=NV["cap"]["target"][cafe]), "0.0", key=True)
        tgt_rows[cafe] = r
    wi.cell(row=rt + 5, column=1, value="NMWA").font = BLACK; wi.cell(row=rt + 5, column=2, value="1 person/day; 2 on community days (no %)").font = BLACK
    inp(wi.cell(row=rt + 5, column=4, value=STAFF["rolemix"]["NMWA"]["avg_core_rate"]), '"$"#,##0.00')
    tgt_rows["NMWA"] = rt + 5
    TGT_RANGE = f"Inputs!$A${rt + 2}:$E${rt + 5}"
    rb = rt + 8
    wi.cell(row=rb, column=1, value="Tier bands — weekly core (walk-in) sales").font = BOLD
    hdr(wi, rb + 1, ["Cafe", "Slow below", "Mid from", "Busy from"], None)
    for i, cafe in enumerate(OPEN3):
        b = S["tiers"][cafe]["bands_weekly_walkin"]; r = rb + 2 + i
        wi.cell(row=r, column=1, value=cafe).font = BLACK
        inp(wi.cell(row=r, column=2, value=b["slow"][1]), '"$"#,##0'); inp(wi.cell(row=r, column=3, value=b["mid"][0]), '"$"#,##0'); inp(wi.cell(row=r, column=4, value=b["busy"][0]), '"$"#,##0')
    rs = rb + 7
    wi.cell(row=rs, column=1, value="Typical daily core sales by tier (walk-in $, per open day; from 2026 transactions) — change these to re-score every template").font = BOLD
    hdr(wi, rs + 1, ["Cafe", "Tier"] + DOW7 + ["Week"], None)
    sales_row = {}; r = rs + 2
    for cafe in CAFES:
        for tier in (["slow", "mid", "busy"] if cafe != "NMWA" else ["slow"]):
            wi.cell(row=r, column=1, value=cafe).font = BLACK; wi.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK
            td = TIERS[cafe]["typical_day"][tier]
            for j, d in enumerate(DOW7):
                v = td[d]["net"] if d in td else None
                c = wi.cell(row=r, column=3 + j, value=v)
                if v is not None: inp(c, '"$"#,##0')
            fx(wi.cell(row=r, column=10, value=f"=SUM(C{r}:I{r})"), '"$"#,##0', bold=True)
            sales_row[(cafe, tier)] = r; r += 1
    for col, w in zip("ABCDEFGHIJ", [44, 12, 12, 12, 12, 12, 12, 12, 12, 12]): wi.column_dimensions[col].width = w

    # ---------------- cafe sheets ----------------
    summary_refs = []
    for cafe in CAFES:
        w = wb.create_sheet(cafe)
        w["A1"] = f"{cafe} — schedule templates by tier (v1.1)"; w["A1"].font = TITLE
        w["A2"] = ("Rows = shifts. Start/End are inputs (blue); Hours, Rate (from Inputs), Wages and Loaded cost are formulas. Day summaries below each tier show labor % at standard rates and at the roster's current average rate. "
                   "'Solo' = one person on the floor; 'Admin covers' = need above the template that the ALL's block should absorb."); w["A2"].font = NOTE
        hdr(w, 4, ["Tier", "Day", "Role", "Code", "Start", "End", "Hours", "Rate", "Wages", "Loaded cost", "Notes"], [12, 6, 26, 6, 8, 8, 8, 9, 10, 11, 70])
        r = 5
        if cafe == "NMWA":
            tiers_data = {"slow": {"days": {d: {"shifts": [{"role": s["role"], "code": s["code"], "start": s["start"], "end": s["end"], "note": s["note"]} for s in dd["shifts"]], "notes": ""} for d, dd in T0["cafes"]["NMWA"]["slow"]["days"].items()}}}
        else:
            tiers_data = {}
            for tier in ["slow", "mid", "busy"]:
                o = T["cafes"][cafe][f"{tier}_{MAIN}"]
                tiers_data[tier] = {"days": {d: {"shifts": [{"role": s["role"], "code": s["code"], "start": s["start"], "end": s["end"], "note": ""} for s in dd["shifts"]],
                                                 "notes": ("solo " + ", ".join(dd["template"]["solo_windows"]) if dd["template"]["solo_windows"] else "") + ("; Admin covers " + ", ".join(dd["template"]["admin_cover_windows"]) if dd["template"]["admin_cover_windows"] else "")}
                                             for d, dd in o["days"].items()}}
        for tier, o in tiers_data.items():
            day_rows = {}
            for d, dd in o["days"].items():
                d_start = r
                for i, s in enumerate(dd["shifts"]):
                    w.cell(row=r, column=1, value=TIER_LABEL[tier]).font = BLACK; w.cell(row=r, column=2, value=d).font = BLACK
                    w.cell(row=r, column=3, value=s["role"]).font = BLACK; w.cell(row=r, column=4, value=s["code"]).font = BLACK
                    inp(w.cell(row=r, column=5, value=tm(s["start"])), "h:mm"); inp(w.cell(row=r, column=6, value=tm(s["end"])), "h:mm")
                    fx(w.cell(row=r, column=7, value=f"=(F{r}-E{r})*24"), "0.00")
                    fx(w.cell(row=r, column=8, value=f"=VLOOKUP(D{r},{RATE_RANGE},3,FALSE)"), '"$"#,##0.00')
                    fx(w.cell(row=r, column=9, value=f"=G{r}*H{r}"), '"$"#,##0'); fx(w.cell(row=r, column=10, value=f"=I{r}*{LOAD_CELL}"), '"$"#,##0')
                    w.cell(row=r, column=11, value=(s["note"] if s["note"] else (dd["notes"] if i == 0 else ""))).font = BLACK
                    r += 1
                day_rows[d] = (d_start, r - 1)
            r += 1
            w.cell(row=r, column=1, value=f"{TIER_LABEL[tier]} — day summary").font = BOLD; r += 1
            hdr(w, r, ["Tier", "Day", "Shifts", "", "", "", "Hours", "", "Wages", "Loaded", "Typical sales", "Labor % std", "SPLH", "Target", "Status", "Labor % roster"]); r += 1
            sum_start = r
            srow = sales_row.get((cafe, tier), sales_row.get((cafe, "slow")))
            for d in o["days"].keys():
                a, b = day_rows[d]
                w.cell(row=r, column=1, value=TIER_LABEL[tier]).font = BLACK; w.cell(row=r, column=2, value=d).font = BLACK
                fx(w.cell(row=r, column=3, value=f"=COUNTA(C{a}:C{b})"), "0")
                fx(w.cell(row=r, column=7, value=f"=SUM(G{a}:G{b})"), "0.00", bold=True); fx(w.cell(row=r, column=9, value=f"=SUM(I{a}:I{b})"), '"$"#,##0', bold=True); fx(w.cell(row=r, column=10, value=f"=SUM(J{a}:J{b})"), '"$"#,##0')
                dcol = L(3 + DOW7.index(d)); sales_ref = f"Inputs!${dcol}${srow}"
                fx(w.cell(row=r, column=11, value=f"={sales_ref}"), '"$"#,##0')
                fx(w.cell(row=r, column=12, value=f"=IF(K{r}>0,I{r}/K{r},\"\")"), "0.0%", bold=True); fx(w.cell(row=r, column=13, value=f"=IF(G{r}>0,K{r}/G{r},\"\")"), '"$"#,##0')
                if cafe != "NMWA":
                    fx(w.cell(row=r, column=14, value=f"=VLOOKUP(\"{cafe}\",{TGT_RANGE},2,FALSE)"), "0%")
                    fx(w.cell(row=r, column=15, value=f"=IF(L{r}=\"\",\"\",IF(L{r}<=N{r},\"OK\",IF(L{r}<=VLOOKUP(\"{cafe}\",{TGT_RANGE},3,FALSE),\"BAND\",\"OVER\")))"))
                fx(w.cell(row=r, column=16, value=f"=IF(K{r}>0,G{r}*VLOOKUP(\"{cafe}\",{TGT_RANGE},4,FALSE)/K{r},\"\")"), "0.0%")
                summary_refs.append((cafe, tier, d, f"'{cafe}'!G{r}", f"'{cafe}'!I{r}", sales_ref, f"'{cafe}'!C{r}"))
                r += 1
            w.cell(row=r, column=1, value=TIER_LABEL[tier]).font = BOLD; w.cell(row=r, column=2, value="WEEK").font = BOLD
            fx(w.cell(row=r, column=3, value=f"=SUM(C{sum_start}:C{r-1})"), "0", bold=True); fx(w.cell(row=r, column=7, value=f"=SUM(G{sum_start}:G{r-1})"), "0.00", bold=True)
            fx(w.cell(row=r, column=9, value=f"=SUM(I{sum_start}:I{r-1})"), '"$"#,##0', bold=True); fx(w.cell(row=r, column=10, value=f"=SUM(J{sum_start}:J{r-1})"), '"$"#,##0', bold=True)
            fx(w.cell(row=r, column=11, value=f"=SUM(K{sum_start}:K{r-1})"), '"$"#,##0', bold=True)
            fx(w.cell(row=r, column=12, value=f"=IF(K{r}>0,I{r}/K{r},\"\")"), "0.0%", bold=True); fx(w.cell(row=r, column=13, value=f"=IF(G{r}>0,K{r}/G{r},\"\")"), '"$"#,##0', bold=True)
            if cafe != "NMWA":
                fx(w.cell(row=r, column=14, value=f"=VLOOKUP(\"{cafe}\",{TGT_RANGE},2,FALSE)"), "0%")
                fx(w.cell(row=r, column=15, value=f"=IF(L{r}=\"\",\"\",IF(L{r}<=N{r},\"OK\",IF(L{r}<=VLOOKUP(\"{cafe}\",{TGT_RANGE},3,FALSE),\"BAND\",\"OVER\")))"), bold=True)
            fx(w.cell(row=r, column=16, value=f"=IF(K{r}>0,G{r}*VLOOKUP(\"{cafe}\",{TGT_RANGE},4,FALSE)/K{r},\"\")"), "0.0%", bold=True)
            for col in range(1, 17): w.cell(row=r, column=col).fill = TOT_FILL
            summary_refs.append((cafe, tier, "WEEK", f"'{cafe}'!G{r}", f"'{cafe}'!I{r}", f"'{cafe}'!K{r}", f"'{cafe}'!C{r}"))
            r += 3
        for col, wd in zip("KLMNOP", [12, 11, 9, 9, 9, 12]):
            w.column_dimensions[col].width = max(w.column_dimensions[col].width or 0, wd)
        w.column_dimensions["K"].width = 60
        w.freeze_panes = "A5"

    # ---------------- Summary ----------------
    ws = wb.create_sheet("Summary", 1)
    ws["A1"] = "Summary — projected core labor % and SPLH per cafe × tier × day vs target (v1.1 templates)"; ws["A1"].font = TITLE
    ws["A2"] = "All values are formulas off the cafe sheets and Inputs. 'Roster %' re-prices the hours at the cafe's current average core rate. Sensitivity: labor % if sales land 10% under the tier's typical level with the template unchanged. Trigger: weekly sales below which the week's template breaches the ceiling — drop to the next slower template."; ws["A2"].font = NOTE
    hdr(ws, 4, ["Cafe", "Tier", "Day", "Shifts", "Core hours", "Wages", "Typical sales", "Labor %", "SPLH", "Target", "Ceiling", "Status", "Labor % at −10% sales", "Ceiling trigger (wk sales)", "On-target at (wk sales)", "Roster %", "Ceiling trigger (roster)"],
        [12, 8, 7, 7, 10, 10, 12, 9, 8, 8, 8, 8, 12, 14, 14, 9, 14])
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
            fx(ws.cell(row=r, column=16, value=f"=IF(G{r}>0,E{r}*VLOOKUP(A{r},{TGT_RANGE},4,FALSE)/G{r},\"\")"), "0.0%")
            if d == "WEEK":
                fx(ws.cell(row=r, column=14, value=f"=F{r}/K{r}"), '"$"#,##0'); fx(ws.cell(row=r, column=15, value=f"=F{r}/J{r}"), '"$"#,##0')
                fx(ws.cell(row=r, column=17, value=f"=E{r}*VLOOKUP(A{r},{TGT_RANGE},4,FALSE)/K{r}"), '"$"#,##0')
        if d == "WEEK":
            for col in range(1, 18): ws.cell(row=r, column=col).fill = TOT_FILL
            ws.cell(row=r, column=3).font = BOLD
        r += 1
    ws.conditional_formatting.add(f"L5:L{r}", CellIsRule(operator="equal", formula=['"OVER"'], font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(f"L5:L{r}", CellIsRule(operator="equal", formula=['"OK"'], font=Font(color="006100", bold=True)))
    ws.freeze_panes = "A5"

    # ---------------- First cut ----------------
    wf = wb.create_sheet("First Cut 10-11")
    wf["A1"] = "Ryan's first-cut template — Square Teams week of Sun 10/11 – Sat 10/17 (core shifts; Manager/Commissary/Roasting excluded), scored on the same basis"; wf["A1"].font = TITLE
    wf["A2"] = ("Green = transcribed from the screenshots (team member names omitted; rates from the Square Teams $ totals). Hours/wages are formulas. 'Template change (Mid)' is the specific edit that turns this shift into the v1.1 Mid template; "
                "Slow-tier shapes are in the docx §3e/4e/5e. 10/12 is Columbus Day: Downtown and 1899 closed — unschedule Monday that week."); wf["A2"].font = NOTE
    hdr(wf, 4, ["Cafe", "Day", "Role", "Start", "End", "Hours", "Rate (as staffed)", "Wages", "Template change (v1.1 Mid)"], [12, 6, 22, 8, 8, 8, 12, 10, 100])
    FC = {
        "Downtown": [
            ("Mon", "Production Barista", 6.5, 13.0, 19.00, "DROP — Monday is a 3-shift day: Production Line 6:00–12:00, Shift Lead 6:30–2:30, Closer 9:00–4:30 (21.5 h)"),
            ("Mon", "Production Line", 6.2, 16.5, 22.75, "10.3 h open-to-close → 6:00–12:00 (prep + rush); the Closer takes the afternoon"),
            ("Mon", "Shift Lead", 9.0, 16.5, 21.75, "Lead opens instead: Shift Lead 6:30–2:30; this 9–4:30 slot becomes the Closer 9:00–4:30 at barista rate"),
            ("Tue", "Production Line", 6.25, 14.0, 18.50, "→ 6:00–12:00"),
            ("Tue", "Production Barista", 7.5, 16.5, 22.75, "9 h → Barista 8:30–2:30 (rush + lunch, 6 h) at $19"),
            ("Tue", "Shift Lead", 6.5, 15.0, 21.75, "→ 6:30–3:30 (lead stays through the 2–3:30p afternoon)"),
            ("Tue", "Production Barista", 9.0, 16.5, 25.25, "→ Closer 8:30–4:30 at barista rate ($19.50), not a $25.25 team member"),
            ("Wed", "Shift Lead", 8.0, 16.5, 19.00, "Second lead → Closer 12:30–4:30 (4 h, barista rate)"),
            ("Wed", "Production Barista", 6.5, 13.5, 21.75, "→ Shift Lead opens 6:30–3:30 (this person is the lead)"),
            ("Wed", "Production Line", 6.25, 15.0, 19.25, "→ 6:00–12:00"),
            ("Wed", "Barista", 9.0, 16.5, 25.25, "→ two rush baristas at $19: 8:00–12:00 and 8:30–2:30 (the 8:30–10 peak needs 4)"),
            ("Thu", "Shift Lead", 8.0, 16.5, 19.00, "Second lead → Closer 8:30–4:30 (barista rate)"),
            ("Thu", "Barista", 6.5, 13.5, 21.75, "→ Shift Lead opens 6:30–3:30"),
            ("Thu", "Production Line", 6.25, 15.0, 19.25, "→ 6:00–12:00"),
            ("Thu", "Production Barista", 9.0, 16.5, 25.25, "→ rush barista 7:30–11:30 (4 h) at $19"),
            ("Fri", "Production Barista", 9.5, 16.5, 19.00, "→ Closer 12:30–4:30 (4 h); Friday afternoon is single-covered"),
            ("Fri", "Shift Lead", 6.5, 13.0, 21.75, "→ 6:30–2:30"),
            ("Fri", "Production Line", 6.25, 16.5, 19.25, "10.25 h → 6:00–12:30; the Closer closes"),
        ],
        "Studio": [
            ("Sun", "Barista", 7.5, 14.0, 18.50, "→ 9:00–2:30"), ("Sun", "Shift Lead", 8.0, 15.0, 24.75, "→ 7:30–4:00 (lead opens; alone 4–5:30 is the Closer)"), ("Sun", "Production Line", 7.25, 15.5, 18.50, "→ 7:00–1:00"), ("Sun", "Production Barista", 10.5, 17.5, 18.50, "→ Closer 10:00–5:30"),
            ("Mon", "Shift Lead", 6.5, 14.5, 20.75, "→ 6:30–2:45"), ("Mon", "Production Barista", 6.5, 13.0, 19.00, "→ Barista 7:00–11:00 (3rd at open for the delivery); DROP when the Manager 7–2 block is on the floor"), ("Mon", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–1:30"),
            ("Tue", "Production Barista", 6.5, 13.0, 20.25, "DROP — Tuesday is a 2-person core day in every tier (Admin covers 8–10 when scheduled)"), ("Tue", "Shift Lead", 7.0, 14.5, 24.75, "→ 6:30–2:45 (lead opens)"), ("Tue", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–1:00"),
            ("Wed", "Barista", 6.5, 13.0, 19.00, "→ 8:00–1:30 (mid) / 9:00–1:00 (slow)"), ("Wed", "Shift Lead", 7.0, 14.5, 22.75, "→ 6:30–2:45"), ("Wed", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–12:00"),
            ("Thu", "Shift Lead", 7.0, 14.5, 20.25, "→ 6:30–2:45"), ("Thu", "Production Line", 6.25, 14.5, 18.50, "→ 6:00–12:00"), ("Thu", "Production Barista", 6.5, 13.5, 24.75, "→ Barista 8:00–1:00 at $19; with the Manager 9–2 block on the floor, 8–12 or DROP"),
            ("Fri", "Shift Lead", 6.5, 15.0, 20.25, "→ 6:30–3:30"), ("Fri", "Production Line", 6.25, 14.5, 22.75, "→ 6:00–12:00 at $19 (this is a $22.75 team member on a PL shift)"), ("Fri", "Shift Lead", 10.0, 17.25, 20.75, "Second lead → Closer 1:30–5:30 (4 h) at $19.50"), ("Fri", "Barista", 7.0, 13.0, 18.50, "→ 8:30–2:00"),
            ("Sat", "Production Line", 7.25, 15.0, 18.50, "→ 7:00–1:30"), ("Sat", "Shift Lead", 10.0, 17.5, 18.50, "→ Closer 9:30–5:30"), ("Sat", "Shift Lead", 7.5, 13.0, 20.75, "→ Shift Lead 7:30–4:30 (one lead through the peak)"), ("Sat", "Hospitality Register", 8.0, 15.0, 22.75, "DROP — four people 9:30–1:30 cover the peak at 12 per person-hour"), ("Sat", "Barista", 9.0, 16.0, 18.50, "→ 9:00–2:30"),
        ],
        "1899 Penn": [
            ("Mon", "Shift Lead", 6.5, 15.5, 20.50, "keep"), ("Mon", "Production Barista", 6.5, 13.5, 22.75, "→ 8:30–2:00 (mid) / 7:30–11:30 (slow); with the Manager 8–2 block on, 4 h rush only"),
            ("Tue", "Shift Lead", 6.5, 15.5, 20.50, "keep"), ("Tue", "Production Barista", 6.5, 15.5, 18.50, "→ 8:30–2:30 (both peaks); the Manager 6:30–1:30 block makes three people on the morning — if kept, PB 8:30–2:30 still"),
            ("Wed", "Shift Lead", 6.5, 15.5, 20.75, "keep"), ("Wed", "Production Barista", 6.5, 15.5, 18.50, "→ 8:30–2:30"),
            ("Thu", "Shift Lead", 6.5, 15.5, 22.75, "keep"), ("Thu", "Barista", 6.5, 15.0, 18.50, "→ 8:30–2:30"),
            ("Fri", "Shift Lead", 6.5, 15.5, 20.75, "keep"), ("Fri", "Production Barista", 6.5, 13.5, 22.75, "→ 7:30–11:30 (4 h rush); the lead is alone from 11:30"),
        ],
        "NMWA": [
            ("Sun", "FOH", 9.0, 15.0, 20.25, "→ 9:30–2:30 (open 10–2); add 2nd FOH 10–2 on the FIRST Sunday of the month"),
            ("Wed", "(none scheduled)", 0, 0, 0, "10/14 is the 2nd Wednesday = community day: ADD FOH 9:30–2:30 + FOH 10:00–2:00"),
            ("Thu", "FOH", 9.0, 15.0, 20.50, "→ 9:30–2:30"), ("Fri", "FOH", 9.0, 15.0, 22.75, "→ 9:30–2:30 (Manager 10–2 block that day is admin, outside core)"), ("Sat", "FOH", 9.0, 15.0, 20.50, "→ 9:30–2:30"),
        ],
    }
    r = 5; fc_rows = {}
    for cafe, shifts in FC.items():
        start = r
        for (d, role, s, e, rate, chg) in shifts:
            wf.cell(row=r, column=1, value=cafe).font = BLACK; pull(wf.cell(row=r, column=2, value=d)); pull(wf.cell(row=r, column=3, value=role))
            if e > 0:
                pull(wf.cell(row=r, column=4, value=dec2time(s)), "h:mm"); pull(wf.cell(row=r, column=5, value=dec2time(e)), "h:mm")
                fx(wf.cell(row=r, column=6, value=f"=(E{r}-D{r})*24"), "0.00"); pull(wf.cell(row=r, column=7, value=rate), '"$"#,##0.00'); fx(wf.cell(row=r, column=8, value=f"=F{r}*G{r}"), '"$"#,##0')
            wf.cell(row=r, column=9, value=chg).font = BLACK
            r += 1
        wf.cell(row=r, column=1, value=cafe).font = BOLD; wf.cell(row=r, column=2, value="WEEK").font = BOLD
        fx(wf.cell(row=r, column=6, value=f"=SUM(F{start}:F{r-1})"), "0.00", bold=True); fx(wf.cell(row=r, column=8, value=f"=SUM(H{start}:H{r-1})"), '"$"#,##0', bold=True)
        for col in range(1, 10): wf.cell(row=r, column=col).fill = TOT_FILL
        fc_rows[cafe] = r; r += 2
    r += 1
    wf.cell(row=r, column=1, value="First cut vs v1.1 templates — weekly core hours and wages").font = BOLD; r += 1
    hdr(wf, r, ["Cafe", "First-cut hours", "First-cut wages", "Mid template hours", "Mid template wages", "Slow template hours", "Slow template wages", "Δ hours vs Mid", "Δ wages vs Mid"]); r += 1
    for cafe in CAFES:
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

    # ---------------- v1.0 vs v1.1 ----------------
    wv = wb.create_sheet("v1.0 vs v1.1")
    wv["A1"] = "What moved between the v1.0 (hand-built) and v1.1 (optimised) templates"; wv["A1"].font = TITLE
    wv["A2"] = "Green = analysis data. Weekly totals at standard rates; per-day shift lists side by side."; wv["A2"].font = NOTE
    hdr(wv, 4, ["Cafe", "Tier", "Day", "v1.0 shifts", "v1.0 hours", "v1.0 %", "v1.1 shifts", "v1.1 hours", "v1.1 %", "v1.1 % roster", "Δ hours", "Δ wages"], [12, 8, 7, 52, 9, 8, 60, 9, 8, 10, 9, 10])
    r = 5
    for cafe in OPEN3:
        for tier in ["slow", "mid", "busy"]:
            o = T["cafes"][cafe][f"{tier}_{MAIN}"]
            for d, dd in o["days"].items():
                v10 = dd["v10_template"]; t = dd["template"]
                wv.cell(row=r, column=1, value=cafe).font = BLACK; wv.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK; wv.cell(row=r, column=3, value=d).font = BLACK
                pull(wv.cell(row=r, column=4, value="; ".join(f"{s['code']} {s['start']}–{s['end']}" for s in v10["shifts"]))); pull(wv.cell(row=r, column=5, value=v10["hours"]), "0.00"); pull(wv.cell(row=r, column=6, value=v10["labor_pct"] / 100), "0.0%")
                pull(wv.cell(row=r, column=7, value="; ".join(f"{s['code']} {s['start']}–{s['end']}" for s in dd["shifts"]))); pull(wv.cell(row=r, column=8, value=t["hours"]), "0.00"); pull(wv.cell(row=r, column=9, value=t["labor_pct"] / 100), "0.0%")
                pull(wv.cell(row=r, column=10, value=t["labor_pct_obs_rate"] / 100), "0.0%")
                fx(wv.cell(row=r, column=11, value=f"=H{r}-E{r}"), "+0.00;-0.00"); pull(wv.cell(row=r, column=12, value=round(t["wages"] - v10["wages"])), '+"$"#,##0;-"$"#,##0')
                r += 1
            w = o["week"]
            wv.cell(row=r, column=1, value=cafe).font = BOLD; wv.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BOLD; wv.cell(row=r, column=3, value="WEEK").font = BOLD
            pull(wv.cell(row=r, column=5, value=w["v10_hours"]), "0.0"); pull(wv.cell(row=r, column=6, value=w["v10_labor_pct"] / 100), "0.0%")
            pull(wv.cell(row=r, column=8, value=w["hours"]), "0.0"); pull(wv.cell(row=r, column=9, value=w["labor_pct"] / 100), "0.0%"); pull(wv.cell(row=r, column=10, value=w["labor_pct_obs_rate"] / 100), "0.0%")
            fx(wv.cell(row=r, column=11, value=f"=H{r}-E{r}"), "+0.0;-0.0"); pull(wv.cell(row=r, column=12, value=round(w["wages"] - w["v10_wages"])), '+"$"#,##0;-"$"#,##0')
            for col in range(1, 13): wv.cell(row=r, column=col).fill = TOT_FILL
            r += 2
    wv.freeze_panes = "A5"

    # ---------------- Variants ----------------
    wvar = wb.create_sheet("Variants")
    wvar["A1"] = "Variants — what each relaxation of the rules costs per week (standard rates)"; wvar["A1"].font = TITLE
    wvar["A2"] = "v1.1 primary = no shift over 9 h. 10-hour lead = Shift Lead may work open-to-close; Comfort = capacity 12 everywhere + longer 2-person floors (v1.0-style); 6-hour minimum = no 4–5.5 h shifts; 8-hour Production Line = prep shift held at 8 h. Green = analysis data; Δ columns are formulas."; wvar["A2"].font = NOTE
    hdr(wvar, 4, ["Cafe", "Tier", "v1.1 hours", "v1.1 wages", "v1.1 %", "10h-lead wages", "Δ", "Comfort wages", "Δ", "6h-min wages", "Δ", "PL-8h wages", "Δ", "Typical wk sales"], [12, 8, 10, 11, 8, 12, 9, 12, 9, 12, 9, 12, 9, 12])
    r = 5
    for cafe in OPEN3:
        for tier in ["slow", "mid", "busy"]:
            w = T["cafes"][cafe][f"{tier}_{MAIN}"]["week"]; v = S["templates"][cafe][tier]["variants_week"]
            wvar.cell(row=r, column=1, value=cafe).font = BLACK; wvar.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK
            pull(wvar.cell(row=r, column=3, value=w["hours"]), "0.0"); pull(wvar.cell(row=r, column=4, value=w["wages"]), '"$"#,##0'); pull(wvar.cell(row=r, column=5, value=w["labor_pct"] / 100), "0.0%")
            for j, key in enumerate(["target", "comfort", "min6h", "pl8"]):
                pull(wvar.cell(row=r, column=6 + 2 * j, value=v[key]["wages"]), '"$"#,##0'); fx(wvar.cell(row=r, column=7 + 2 * j, value=f"={L(6 + 2 * j)}{r}-D{r}"), '+"$"#,##0;-"$"#,##0')
            pull(wvar.cell(row=r, column=14, value=w["sales_typical"]), '"$"#,##0')
            r += 1
    wvar.freeze_panes = "A5"

    # ---------------- Need vs Actual ----------------
    wn = wb.create_sheet("Need vs Actual")
    wn["A1"] = "Half-hour demand, people needed (v1.1 rules) and actual core headcount — per cafe × tier × day (sampled weeks)"; wn["A1"].font = TITLE
    wn["A2"] = "need = max(floor, ceil(txns/hr ÷ cafe capacity)); capacity Downtown 14 / Studio 12 / 1899 14; actual = average core headcount on the clock in the sampled weeks of that tier. Green = analysis data."; wn["A2"].font = NOTE
    allt = sorted({x["t"] for c in NV["variants"]["target"].values() for tt in c["tiers"].values() for dd in tt.values() for x in dd["by_halfhour"]})
    def tl(t):
        hh = int(t); return f"{hh}:{'30' if t % 1 else '00'}"
    hdr(wn, 4, ["Cafe", "Tier", "Day", "Series"] + [tl(t) for t in allt] + ["Day total"], [12, 8, 6, 9] + [6] * len(allt) + [9])
    r = 5
    for cafe in OPEN3:
        for tier in ["slow", "mid", "busy"]:
            for d, rec in NV["variants"]["target"][cafe]["tiers"][tier].items():
                by = {x["t"]: x for x in rec["by_halfhour"]}
                for series, key, fmt in [("txns/hr", "txns_hr", "0"), ("floor", "floor", "0"), ("need", "need", "0"), ("actual", "actual", "0.0")]:
                    wn.cell(row=r, column=1, value=cafe).font = BLACK; wn.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK
                    wn.cell(row=r, column=3, value=d).font = BLACK; wn.cell(row=r, column=4, value=series).font = BLACK
                    for j, t in enumerate(allt):
                        v = by.get(t, {}).get(key)
                        if v is not None: pull(wn.cell(row=r, column=5 + j, value=v), fmt)
                    if series in ("need", "actual"):
                        fx(wn.cell(row=r, column=5 + len(allt), value=f"=SUM({L(5)}{r}:{L(4+len(allt))}{r})*0.5"), "0.0")
                    r += 1
    wn.freeze_panes = "E5"

    # ---------------- Incentive ----------------
    wc = wb.create_sheet("Incentive")
    wc["A1"] = "Incentive — Beat Share calculator, plan comparison and the annual simulation"; wc["A1"].font = TITLE
    wc["A2"] = "Calculator: enter a month's core sales and core wages per cafe (blue); the pool, cap and payout are formulas. Below: the nine structures on the same scenarios and the 2,000-run 2027 simulation (green = analysis data)."; wc["A2"].font = NOTE
    wc["A4"] = "Beat Share calculator (Plan A)"; wc["A4"].font = BOLD
    hdr(wc, 5, ["Cafe", "Core sales (month)", "Core wages (month)", "Target", "Core wage %", "Beat $ (target × sales − wages)", "Core hours (month)"], [22, 16, 16, 9, 12, 22, 14])
    for i, cafe in enumerate(OPEN3):
        r = 6 + i; wk = T["cafes"][cafe][f"mid_{MAIN}"]["week"]
        wc.cell(row=r, column=1, value=cafe).font = BLACK
        inp(wc.cell(row=r, column=2, value=round(wk["sales_typical"] * 52 / 12)), '"$"#,##0', key=True)
        inp(wc.cell(row=r, column=3, value=round((S["targets"][cafe]["target"] - 0.02) * wk["sales_typical"] * 52 / 12)), '"$"#,##0', key=True)
        fx(wc.cell(row=r, column=4, value=f"=VLOOKUP(A{r},{TGT_RANGE},2,FALSE)"), "0%")
        fx(wc.cell(row=r, column=5, value=f"=IF(B{r}>0,C{r}/B{r},\"\")"), "0.0%"); fx(wc.cell(row=r, column=6, value=f"=D{r}*B{r}-C{r}"), '"$"#,##0;("$"#,##0)')
        inp(wc.cell(row=r, column=7, value=round(wk["hours"] * 52 / 12)), "#,##0")
    wc.cell(row=9, column=1, value="Open cafes").font = BOLD
    fx(wc.cell(row=9, column=2, value="=SUM(B6:B8)"), '"$"#,##0', bold=True); fx(wc.cell(row=9, column=3, value="=SUM(C6:C8)"), '"$"#,##0', bold=True)
    fx(wc.cell(row=9, column=5, value="=IF(B9>0,C9/B9,\"\")"), "0.0%", bold=True); fx(wc.cell(row=9, column=6, value="=SUM(F6:F8)"), '"$"#,##0;("$"#,##0)', bold=True); fx(wc.cell(row=9, column=7, value="=SUM(G6:G8)"), "#,##0", bold=True)
    for col in range(1, 8): wc.cell(row=9, column=col).fill = TOT_FILL
    labels = [("Share of the beat", 0.5, "0%"), ("Cap (% of open-cafe core sales)", 0.015, "0.0%"), ("Pool before cap = share × beat (0 if negative)", "=MAX(0,B11*F9)", '"$"#,##0'),
              ("Cap $", "=B12*B9", '"$"#,##0'), ("Payout (pool, capped)", "=MIN(B13,B14)", '"$"#,##0'), ("Payout as % of open-cafe core sales", "=IF(B9>0,B15/B9,0)", "0.00%"),
              ("Payout per core hour", "=IF(G9>0,B15/G9,0)", '"$"#,##0.00'), ("Bonus % of core wages (percentage-of-wages split)", "=IF(C9>0,B15/C9,0)", "0.00%"),
              ("Employer cost incl. 12% load", f"=B15*{LOAD_CELL}", '"$"#,##0'), ("Slipstream keeps (beat − payout)", "=MAX(0,F9)-B15", '"$"#,##0')]
    for i, (lab, val, fmt) in enumerate(labels):
        r = 11 + i; wc.cell(row=r, column=1, value=lab).font = BLACK
        c = wc.cell(row=r, column=2, value=val)
        if isinstance(val, (int, float)): inp(c, fmt, key=True)
        else: fx(c, fmt, bold=(lab.startswith("Payout (pool")))
    wc.cell(row=22, column=1, value="Guardrail (not modelled): pays only if the month's schedules met the template floors, no week was flagged for under-coverage, and core sales ≥ 90% of projection.").font = NOTE
    # plan comparison
    r = 24
    wc.cell(row=r, column=1, value="Nine structures on the same scenarios — mid month, three open cafes pooled (company payout $ / % of sales / per 30 h-per-week person)").font = BOLD; r += 1
    scen = ["AT target", "1 pt under", "2 pts under", "3 pts under", "4 pts under"]
    hdr(wc, r, ["Structure"] + [f"{s} — payout" for s in scen] + [f"{s} — per 30h person" for s in scen]); r += 1
    PLANS = INC["assumptions"]["plans"]; comp = INC["plan_comparison"]["scenarios"]
    for key, label in PLANS.items():
        wc.cell(row=r, column=1, value=f"{key}. {label}").font = BLACK
        for j, s in enumerate(scen):
            pull(wc.cell(row=r, column=2 + j, value=comp[s][key]["payout"]), '"$"#,##0'); pull(wc.cell(row=r, column=7 + j, value=comp[s][key]["per_person_month_30h"]), '"$"#,##0')
        r += 1
    r += 1
    sim = INC["annual_simulation"]
    wc.cell(row=r, column=1, value=f"Annual simulation — {sim['runs']} runs; {sim['model']}").font = BOLD; r += 1
    hdr(wc, r, ["Plan", "Annual payout p10", "p50", "p90", "Mean", "Share of labor saving (p50)", "Labor saving vs Apr–Aug run-rate p10 / p50 / p90"]); r += 1
    sv = sim["labor_saving_vs_aprAug_run_rate"]
    for key, label in PLANS.items():
        v = sim["plans"][key]
        wc.cell(row=r, column=1, value=f"{key}. {label}").font = BLACK
        pull(wc.cell(row=r, column=2, value=v["p10"]), '"$"#,##0'); pull(wc.cell(row=r, column=3, value=v["p50"]), '"$"#,##0'); pull(wc.cell(row=r, column=4, value=v["p90"]), '"$"#,##0'); pull(wc.cell(row=r, column=5, value=v["mean"]), '"$"#,##0')
        pull(wc.cell(row=r, column=6, value=v["share_of_saving_p50"] / 100), "0.0%")
        if key == "A": pull(wc.cell(row=r, column=7, value=f"${sv['p10']:,} / ${sv['p50']:,} / ${sv['p90']:,}"))
        r += 1
    r += 1
    ps = INC["profit_share_sizing"]
    wc.cell(row=r, column=1, value="Retirement profit-share sizing (2027 tier calendar, every cafe exactly at target)").font = BOLD; r += 1
    for lab, val, fmt in [("Open-cafe operating income, year", ps["open_cafe_noi_year_at_target_calendar"], '"$"#,##0'), ("Support entities (Commissary + SRP + Cafe Mgmt + Central), year", ps["support_burden_year_normalized"], '"$"#,##0;("$"#,##0)'),
                          ("Consolidated net income at target, year", ps["consolidated_ni_year_at_target"], '"$"#,##0'), ("5% share", ps["share_5pct_of_consolidated"], '"$"#,##0'), ("10% share", ps["share_10pct_of_consolidated"], '"$"#,##0'),
                          ("Per 30 h/wk team member at 10%", ps["per_person_at_10pct_consolidated_30h"], '"$"#,##0')]:
        wc.cell(row=r, column=1, value=lab).font = BLACK; pull(wc.cell(row=r, column=2, value=val), fmt); r += 1
    wc.column_dimensions["A"].width = 70
    for col in "BCDEFGHIJK": wc.column_dimensions[col].width = max(wc.column_dimensions[col].width or 0, 16)
    wc.freeze_panes = "A5"

    wb.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    main()
