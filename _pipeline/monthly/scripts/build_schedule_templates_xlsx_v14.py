#!/usr/bin/env python3
"""Build 'Slipstream Schedule Templates - Sep 2026 v1.4.xlsx' from templates_v13.json (Sep 5 + Sep 10 rules; no admin sets).
Slipstream color protocol: blue = input (rates, targets, capacity, typical sales, start/end times), green = pulled data
(need/actual curves, first-cut transcription, simulation results), black = formulas/labels.
Sheets: README, Summary, Inputs, Downtown, Studio, 1899 Penn, NMWA, First Cut 10-11, Tier Calendar, Backtest, Need vs Actual. Pay Review and Incentive moved to 'Slipstream Pay & Incentive Review - Sep 2026 v1.0.xlsx'.
"""
import os, json, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Schedule Templates - Sep 2026 v1.4.xlsx")
T = json.load(open(os.path.join(DD, "templates_v13.json"))); T0 = json.load(open(os.path.join(DD, "templates.json"))); PAY = json.load(open(os.path.join(DD, "pay_review_v13.json"))); CAL19 = json.load(open(os.path.join(DD, "tiers_v19_sep_dec.json")))
NV = json.load(open(os.path.join(DD, "need_v13.json"))); S = json.load(open(os.path.join(DATA, "labor_deep_dive_2026-09.json")))
INC = json.load(open(os.path.join(DD, "incentive_v12.json"))); PC = json.load(open(os.path.join(DD, "projections_check_v12.json"))); TIERS = json.load(open(os.path.join(DD, "tiers.json"))); STAFF = json.load(open(os.path.join(DD, "staffing.json")))
MAIN = "std"

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
RATES = dict(T["rates"])

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
    ws["A1"] = "Slipstream Schedule Templates by Demand Tier — September 2026 v1.4"; ws["A1"].font = TITLE
    lines = [
        "Companion to 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.4.docx'. Rules: two openers (kitchen 6:15, lead 6:30) and two closers (to close+30) at Downtown & Studio; 6-hour minimum shifts; Downtown third person by 8:00 Tue–Thu, max 4, Tue/Wed/Thu one template; Studio line person on all day Mon–Thu, third person 8:00–2:30, max 4 on Friday and slow weekends, 5 on mid/busy weekends; 1899 two shifts every day (stagger when slow, both close together); NMWA 9:00–3:00. Admin days are NOT in the templates (scheduler's tool). Planning rates = Aug 2026 roster averages by role. Slow-month ceiling +2 pts.",
        "",
        "COLOR PROTOCOL — blue text on light blue = input you can change (rates, targets, capacity, typical sales, shift start/end; yellow = key assumptions); green = pulled from the analysis data (need/actual curves, first-cut transcription, simulation results); black = formula or label.",
        "",
        "HOW TO USE",
        "1. Pick the tier for the week from the projected weekly core sales in the weekly report's forward block (bands on the Inputs sheet). Below the band floor → use the next slower template. Build the 10/11 week on MID.",
        "2. Open the cafe sheet, copy the shifts for that tier into Square Teams. 'Solo' notes mark the windows one person runs the floor; 'Admin covers' marks half-hours where the need exceeds the template and the ALL's block is the intended cover.",
        "3. The day and week summaries recompute labor % from the Inputs rates and typical sales. Change a start/end time and everything flows. Status: OK = at/under target, BAND = between target and ceiling, OVER = above ceiling (ask Ryan before publishing). The 'roster' % re-prices the same hours at the cafe's current average core rate (Inputs).",
        "4. Summary sheet: every cafe × tier × day, weekly rollups, −10% sales sensitivity and the ceiling trigger (weekly sales below which the template breaches the ceiling).",
        "5. First Cut 10-11: Ryan's Square Teams week transcribed and scored on the same basis, with the specific change to each shift. Tier Calendar: the Sep–Dec weekly tiers on Projections v1.9.",
        "6. Backtest: Sep 2025–Aug 2026 month by month per cafe — actual sales, COGS, labor and NOI from the Consolidated Model against the templates costed on each week's actual tier. Change the bridge factor or the COGS scenario (blue) and the profit columns recompute. Pay and incentive are in the separate Pay & Incentive Review workbook.",
        "",
        "BASIS — labor % = gross guaranteed wages ÷ core (ex-catering walk-in) sales, the basis of the 30/32/35 targets. The P&L shows ≈ 1.12 × gross − CC tips (docx §1). Times are Eastern; shifts are open-to-clock-out; prep starts 30–60 min before open.",
        "",
        "Notes column: 'Admin covers' = the peak above the cap; 'solo' = one person on the floor (1899 stagger days only). Federal-holiday Mondays: Downtown and 1899 closed.",
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
    hdr(wi, rt + 1, ["Cafe", "Target", "Ceiling", "Roster avg core rate ($/hr)", "Planning capacity (txns per person-hour)", "Ceiling, slow months"], None)
    tgt_rows = {}
    for i, cafe in enumerate(OPEN3):
        r = rt + 2 + i
        wi.cell(row=r, column=1, value=cafe).font = BLACK
        inp(wi.cell(row=r, column=2, value=T["targets"][cafe]["target"]), "0%", key=True); inp(wi.cell(row=r, column=3, value=T["targets"][cafe]["ceiling"]), "0%", key=True)
        inp(wi.cell(row=r, column=4, value=STAFF["rolemix"][cafe]["avg_core_rate"]), '"$"#,##0.00'); inp(wi.cell(row=r, column=5, value=NV["cap"][cafe]), "0.0", key=True); inp(wi.cell(row=r, column=6, value=T["targets"][cafe]["ceiling"] + 0.02), "0%", key=True)
        tgt_rows[cafe] = r
    wi.cell(row=rt + 5, column=1, value="NMWA").font = BLACK; wi.cell(row=rt + 5, column=2, value="1 person/day; 2 on community days (no %)").font = BLACK
    inp(wi.cell(row=rt + 5, column=4, value=STAFF["rolemix"]["NMWA"]["avg_core_rate"]), '"$"#,##0.00')
    tgt_rows["NMWA"] = rt + 5
    TGT_RANGE = f"Inputs!$A${rt + 2}:$F${rt + 5}"
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
            tiers_data = {"slow": {"days": {d: {"shifts": [{"role": x["role"], "code": x["code"], "start": x["start"], "end": x["end"], "note": ""} for x in dd["shifts"]], "notes": dd.get("note", "")} for d, dd in T["cafes"]["NMWA"]["slow"]["days"].items()}}}
        else:
            tiers_data = {}
            for tier in ["slow", "mid", "busy"]:
                o = T["cafes"][cafe][tier]; days = {}
                for d, dd in o["days"].items():
                    t = dd["template"]
                    note = ("solo " + ", ".join(t["solo_windows"]) if t["solo_windows"] else "")
                    if t["admin_cover_windows"]: note += ("; " if note else "") + "Admin covers " + ", ".join(t["admin_cover_windows"])
                    days[d] = {"shifts": [{"role": ("Barista (3rd)" if (cafe == "Studio" and d in DOW7[:4] and x["code"] == "CL") else x["role"]), "code": x["code"], "start": x["start"], "end": x["end"], "note": ""} for x in dd["shifts"]], "notes": note}
                tiers_data[tier] = {"days": days}
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
                    cc = 6 if tier == "slow" else 3
                    fx(w.cell(row=r, column=15, value=f"=IF(L{r}=\"\",\"\",IF(L{r}<=N{r},\"OK\",IF(L{r}<=VLOOKUP(\"{cafe}\",{TGT_RANGE},{cc},FALSE),\"BAND\",\"OVER\")))"))
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
                cc = 6 if tier == "slow" else 3
                fx(w.cell(row=r, column=15, value=f"=IF(L{r}=\"\",\"\",IF(L{r}<=N{r},\"OK\",IF(L{r}<=VLOOKUP(\"{cafe}\",{TGT_RANGE},{cc},FALSE),\"BAND\",\"OVER\")))"), bold=True)
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
    ws["A2"] = "All values are formulas off the cafe sheets and Inputs. Ceiling = the slow-month ceiling (+2) for Slow rows. 'Roster %' re-prices the hours at the cafe's current average core rate. Sensitivity: labor % if sales land 10% under the tier's typical level with the template unchanged. Trigger: weekly sales below which the week's template breaches the ceiling — drop to the next slower template."; ws["A2"].font = NOTE
    hdr(ws, 4, ["Cafe", "Tier", "Day", "Shifts", "Core hours", "Wages", "Typical sales", "Labor %", "SPLH", "Target", "Ceiling", "Status", "Labor % at −10% sales", "Ceiling trigger (wk sales)", "On-target at (wk sales)", "Roster %", "Ceiling trigger (roster)"],
        [12, 8, 7, 7, 10, 10, 12, 9, 8, 8, 8, 8, 12, 14, 14, 9, 14])
    r = 5
    for (cafe, tier, d, hcell, wcell, scell, ccell) in summary_refs:
        ws.cell(row=r, column=1, value=cafe).font = BLACK; ws.cell(row=r, column=2, value=TIER_LABEL[tier]).font = BLACK; ws.cell(row=r, column=3, value=d).font = BLACK
        fx(ws.cell(row=r, column=4, value=f"={ccell}"), "0")
        fx(ws.cell(row=r, column=5, value=f"={hcell}"), "0.00"); fx(ws.cell(row=r, column=6, value=f"={wcell}"), '"$"#,##0'); fx(ws.cell(row=r, column=7, value=f"={scell}"), '"$"#,##0')
        fx(ws.cell(row=r, column=8, value=f"=IF(G{r}>0,F{r}/G{r},\"\")"), "0.0%", bold=True); fx(ws.cell(row=r, column=9, value=f"=IF(E{r}>0,G{r}/E{r},\"\")"), '"$"#,##0')
        if cafe != "NMWA":
            fx(ws.cell(row=r, column=10, value=f"=VLOOKUP(A{r},{TGT_RANGE},2,FALSE)"), "0%"); fx(ws.cell(row=r, column=11, value=f"=VLOOKUP(A{r},{TGT_RANGE},{6 if tier == 'slow' else 3},FALSE)"), "0%")
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
    wf["A2"] = ("Green = transcribed from the screenshots (team member names omitted; rates from the Square Teams $ totals). Hours/wages are formulas. 'Template change (Mid)' is the specific edit that turns this shift into the Mid template (docx §3–5); "
                "Slow-tier shapes are in the docx §3e/4e/5e. 10/12 is Columbus Day: Downtown and 1899 closed — unschedule Monday that week."); wf["A2"].font = NOTE
    hdr(wf, 4, ["Cafe", "Day", "Role", "Start", "End", "Hours", "Rate (as staffed)", "Wages", "Template change (Mid template)"], [12, 6, 22, 8, 8, 8, 12, 10, 100])
    FC = {
        "Downtown": [
            ("Mon", "Production Barista", 6.5, 13.0, 19.00, "DROP — Monday is three people: Line 6:15–12:15, Lead 6:30–4:30, Closer 9:00–4:30 (mid) / 10:30–4:30 (slow)"),
            ("Mon", "Production Line", 6.2, 16.5, 22.75, "10.3 h → 6:15–12:15"),
            ("Mon", "Shift Lead", 9.0, 16.5, 21.75, "Lead opens and closes 6:30–4:30; this 9–4:30 slot becomes the Closer at barista rate"),
            ("Tue", "Production Line", 6.25, 14.0, 18.50, "→ 6:15–12:15 (mid/busy) or 6:15–1:15 (slow)"),
            ("Tue", "Production Barista", 7.5, 16.5, 22.75, "→ Barista 8:30–4:30 (mid/busy) at $20; DROP in slow weeks"),
            ("Tue", "Shift Lead", 6.5, 15.0, 21.75, "→ 6:30–2:30 (mid/busy) or 6:30–4:30 (slow)"),
            ("Tue", "Production Barista", 9.0, 16.5, 25.25, "→ Closer 8:00–4:30 (third person by 8:00) at $20, not a $25.25 team member"),
            ("Wed", "Shift Lead", 8.0, 16.5, 19.00, "Second lead → Closer 8:00–4:30 (barista rate) — same template as Tue/Thu"),
            ("Wed", "Production Barista", 6.5, 13.5, 21.75, "→ Shift Lead opens 6:30–2:30 (this person is the lead)"),
            ("Wed", "Production Line", 6.25, 15.0, 19.25, "→ 6:15–12:15"),
            ("Wed", "Barista", 9.0, 16.5, 25.25, "→ Barista 8:30–4:30 at $20 (mid/busy); DROP in slow weeks"),
            ("Thu", "Shift Lead", 8.0, 16.5, 19.00, "Second lead → Closer 8:00–4:30 (barista rate)"),
            ("Thu", "Barista", 6.5, 13.5, 21.75, "→ Shift Lead opens 6:30–2:30"),
            ("Thu", "Production Line", 6.25, 15.0, 19.25, "→ 6:15–12:15"),
            ("Thu", "Production Barista", 9.0, 16.5, 25.25, "→ Barista 8:30–4:30 at $20 (mid/busy); DROP in slow weeks"),
            ("Fri", "Production Barista", 9.5, 16.5, 19.00, "→ Closer/Barista 10:30–4:30 (9:00–4:30 in busy weeks)"),
            ("Fri", "Shift Lead", 6.5, 13.0, 21.75, "→ 6:30–4:30"),
            ("Fri", "Production Line", 6.25, 16.5, 19.25, "10.25 h → 6:15–12:15"),
        ],
        "Studio": [
            ("Sun", "Barista", 7.5, 14.0, 18.50, "→ 8:00–2:00 (mid/busy) or 9:30–5:30 (slow)"), ("Sun", "Shift Lead", 8.0, 15.0, 24.75, "→ 7:30–5:30 (one lead, opens and closes; 7:30–3:30 in slow weeks)"), ("Sun", "Production Line", 7.25, 15.5, 18.50, "→ 7:15–3:15"), ("Sun", "Production Barista", 10.5, 17.5, 18.50, "→ Closer 10:00–5:30 (11:00 in slow weeks)"),
            ("Mon", "Shift Lead", 6.5, 14.5, 20.75, "→ 6:30–2:30"), ("Mon", "Production Barista", 6.5, 13.0, 19.00, "→ Barista 7:00–1:00 (three at open for the delivery)"), ("Mon", "Production Line", 6.25, 14.5, 18.50, "→ 6:15–2:30 (line person all day, closes the line)"),
            ("Tue", "Production Barista", 6.5, 13.0, 20.25, "→ third person 8:00–2:30 (closes with the team)"), ("Tue", "Shift Lead", 7.0, 14.5, 24.75, "→ 6:30–2:30 (lead opens)"), ("Tue", "Production Line", 6.25, 14.5, 18.50, "→ 6:15–2:30"),
            ("Wed", "Barista", 6.5, 13.0, 19.00, "→ 8:00–2:30"), ("Wed", "Shift Lead", 7.0, 14.5, 22.75, "→ 6:30–2:30"), ("Wed", "Production Line", 6.25, 14.5, 18.50, "→ 6:15–2:30"),
            ("Thu", "Shift Lead", 7.0, 14.5, 20.25, "→ 6:30–2:30"), ("Thu", "Production Line", 6.25, 14.5, 18.50, "→ 6:15–2:30"), ("Thu", "Production Barista", 6.5, 13.5, 24.75, "→ third person 8:00–2:30 at $20"),
            ("Fri", "Shift Lead", 6.5, 15.0, 20.25, "→ 6:30–2:30"), ("Fri", "Production Line", 6.25, 14.5, 22.75, "→ 6:15–2:15 at $19.75 (this is a $22.75 team member on a line shift)"), ("Fri", "Shift Lead", 10.0, 17.25, 20.75, "Second lead → Closer 8:30–5:30 (barista rate)"), ("Fri", "Barista", 7.0, 13.0, 18.50, "→ Barista 11:30–5:30 (second closer); 9:30–5:30 in busy weeks"),
            ("Sat", "Production Line", 7.25, 15.0, 18.50, "→ 7:15–3:15"), ("Sat", "Shift Lead", 10.0, 17.5, 18.50, "→ Barista 9:30/10:30–5:30 (closer)"), ("Sat", "Shift Lead", 7.5, 13.0, 20.75, "→ Shift Lead 7:30–5:30 (one lead, opens and closes; 7:30–3:30 in busy weeks with a Closer 10:00–5:30)"), ("Sat", "Hospitality Register", 8.0, 15.0, 22.75, "DROP in slow/mid weeks — four people cover the peak; fifth person (Closer 10:00–5:30) only in busy weeks"), ("Sat", "Barista", 9.0, 16.0, 18.50, "→ 8:00–2:00"),
        ],
        "1899 Penn": [
            ("Mon", "Shift Lead", 6.5, 15.5, 20.50, "→ 6:30–2:00 (mid) / 6:30–12:30 (slow stagger) / 6:30–2:30 (busy)"), ("Mon", "Production Barista", 6.5, 13.5, 22.75, "→ 9:00–3:30 (closes; 9:30–3:30 on stagger days)"),
            ("Tue", "Shift Lead", 6.5, 15.5, 20.50, "→ 6:30–2:30 (6:30–3:30 in busy weeks)"), ("Tue", "Production Barista", 6.5, 15.5, 18.50, "→ 8:30–3:30 (7:30–3:30 in busy weeks)"),
            ("Wed", "Shift Lead", 6.5, 15.5, 20.75, "→ 6:30–2:30 (6:30–3:30 in busy weeks)"), ("Wed", "Production Barista", 6.5, 15.5, 18.50, "→ 8:30–3:30 (7:30–3:30 in busy weeks)"),
            ("Thu", "Shift Lead", 6.5, 15.5, 22.75, "→ 6:30–2:30"), ("Thu", "Barista", 6.5, 15.0, 18.50, "→ 8:30–3:30 (8:00–3:30 in busy weeks)"),
            ("Fri", "Shift Lead", 6.5, 15.5, 20.75, "→ 6:30–12:30 (stagger: solo open)"), ("Fri", "Production Barista", 6.5, 13.5, 22.75, "→ 9:30–3:30 (stagger: solo close; 9:00–3:30 in busy weeks)"),
        ],
        "NMWA": [
            ("Sun", "FOH", 9.0, 15.0, 20.25, "keep 9:00–3:00; add a second FOH 9:00–3:00 on the FIRST Sunday of the month"),
            ("Wed", "(none scheduled)", 0, 0, 0, "10/14 is the 2nd Wednesday = community day: ADD 2 × FOH 9:00–3:00"),
            ("Thu", "FOH", 9.0, 15.0, 20.50, "keep"), ("Fri", "FOH", 9.0, 15.0, 22.75, "keep"), ("Sat", "FOH", 9.0, 15.0, 20.50, "keep"),
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
    wf.cell(row=r + 1, column=1, value="Note: first-cut wages use the rates of the people scheduled ($18.50–25.25); template wages use planning rates ($22.25 lead / $19.75 line / $20.25 production barista / $20 barista & closer).").font = NOTE
    wf.freeze_panes = "A5"

    # ---------------- Tier Calendar ----------------
    wp = wb.create_sheet("Tier Calendar")
    wp["A1"] = "Sep–Dec 2026 weekly tiers on Projections v1.9 (2025 same-day sales × Downtown 1.00 / Studio 0.87 for Sep–Dec; 1899 unchanged)"; wp["A1"].font = TITLE
    wp["A2"] = "Green = pulled from the recalculated v1.9 workbook; weeks scaled to full weeks. Pick the tier from the projected week, not the calendar month — November starts mid and ends slow."; wp["A2"].font = NOTE
    hdr(wp, 4, ["Cafe", "Week of", "Projected weekly core sales", "Tier"], [12, 12, 20, 8])
    r = 5
    for cafe, rows in CAL19.items():
        for x in rows:
            wp.cell(row=r, column=1, value=cafe).font = BLACK; pull(wp.cell(row=r, column=2, value=x["week"])); pull(wp.cell(row=r, column=3, value=x["proj"]), '"$"#,##0'); pull(wp.cell(row=r, column=4, value=x["tier"]))
            r += 1
        r += 1
    wp.freeze_panes = "A5"



    # ---------------- Backtest ----------------
    BT = json.load(open(os.path.join(DD, "backtest_v13.json")))
    wb_ = wb.create_sheet("Backtest")
    wb_["A1"] = "Backtest — the v1.3 templates over Sep 2025 – Aug 2026, per cafe (labor is the only thing that changes)"; wb_["A1"].font = TITLE
    wb_["A2"] = "Green = actuals (Consolidated Model 2025/2026, transactions) and the template wages for each month's actual tier weeks; blue = assumptions; black = formulas. Template P&L labor = bridge × (template core wages + admin/training) − CC tips."; wb_["A2"].font = NOTE
    wb_["A4"] = "Bridge factor (P&L labor = factor × gross wages − CC tips)"; wb_["A4"].font = BLACK; inp(wb_["D4"], "0.00", key=True); wb_["D4"] = 1.12; inp(wb_["D4"], "0.00", key=True)
    wb_["A5"] = "COGS scenario (% of core sales) for the 'core-basis profit' columns"; wb_["A5"].font = BLACK; wb_["D5"] = 0.285; inp(wb_["D5"], "0.0%", key=True)
    cols = ["Cafe", "Month", "Open days", "Weeks slow", "Weeks mid", "Weeks busy", "Core sales", "Walk-in sales", "COGS actual", "COGS %", "Opex", "Labor paid (P&L)", "Labor % paid",
            "Template core hours", "Template core wages", "Template wage %", "Admin + training", "Tips (est.)", "Labor on templates", "Labor % templates", "Saving", "NOI reported", "NOI on templates",
            "Core profit @ scenario, labor paid", "Core profit @ scenario, templates"]
    hdr(wb_, 7, cols, [12, 8, 6, 6, 6, 6, 11, 11, 11, 7, 10, 12, 8, 9, 11, 8, 10, 10, 12, 8, 10, 11, 11, 14, 14])
    r = 8; first = r
    for cafe in OPEN3:
        c0 = r
        for x in BT["cafes"][cafe]["months"]:
            tw = x["tier_weeks"]
            wb_.cell(row=r, column=1, value=cafe).font = BLACK; wb_.cell(row=r, column=2, value=x["month"]).font = BLACK
            for j, v in [(3, x["open_days"]), (4, tw.get("slow", 0)), (5, tw.get("mid", 0)), (6, tw.get("busy", 0))]: pull(wb_.cell(row=r, column=j, value=v), "0")
            for j, kk in [(7, "core_sales"), (8, "walkin_sales"), (9, "cogs_actual"), (11, "opex"), (12, "labor_actual"), (15, "template_core_wages"), (17, "admin_training"), (18, "tips_est"), (22, "noi_reported")]:
                pull(wb_.cell(row=r, column=j, value=x[kk]), '"$"#,##0;("$"#,##0)')
            pull(wb_.cell(row=r, column=14, value=x["template_core_hours"]), "#,##0")
            fx(wb_.cell(row=r, column=10, value=f"=IF(G{r}>0,I{r}/G{r},0)"), "0.0%"); fx(wb_.cell(row=r, column=13, value=f"=IF(G{r}>0,L{r}/G{r},0)"), "0.0%")
            fx(wb_.cell(row=r, column=16, value=f"=IF(H{r}>0,O{r}/H{r},0)"), "0.0%")
            fx(wb_.cell(row=r, column=19, value=f"=$D$4*(O{r}+Q{r})-R{r}"), '"$"#,##0'); fx(wb_.cell(row=r, column=20, value=f"=IF(G{r}>0,S{r}/G{r},0)"), "0.0%")
            fx(wb_.cell(row=r, column=21, value=f"=L{r}-S{r}"), '"$"#,##0;("$"#,##0)'); fx(wb_.cell(row=r, column=23, value=f"=V{r}+U{r}"), '"$"#,##0;("$"#,##0)')
            fx(wb_.cell(row=r, column=24, value=f"=G{r}*(1-$D$5)-L{r}-K{r}"), '"$"#,##0;("$"#,##0)'); fx(wb_.cell(row=r, column=25, value=f"=G{r}*(1-$D$5)-S{r}-K{r}"), '"$"#,##0;("$"#,##0)')
            r += 1
        wb_.cell(row=r, column=1, value=cafe).font = BOLD; wb_.cell(row=r, column=2, value="12 months").font = BOLD
        for j in [3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 17, 18, 19, 21, 22, 23, 24, 25]:
            fx(wb_.cell(row=r, column=j, value=f"=SUM({L(j)}{c0}:{L(j)}{r-1})"), ("0" if j <= 6 else "#,##0" if j == 14 else '"$"#,##0;("$"#,##0)'), bold=True)
        fx(wb_.cell(row=r, column=10, value=f"=IF(G{r}>0,I{r}/G{r},0)"), "0.0%", bold=True); fx(wb_.cell(row=r, column=13, value=f"=IF(G{r}>0,L{r}/G{r},0)"), "0.0%", bold=True)
        fx(wb_.cell(row=r, column=16, value=f"=IF(H{r}>0,O{r}/H{r},0)"), "0.0%", bold=True); fx(wb_.cell(row=r, column=20, value=f"=IF(G{r}>0,S{r}/G{r},0)"), "0.0%", bold=True)
        for col in range(1, 26): wb_.cell(row=r, column=col).fill = TOT_FILL
        r += 2
    wb_.cell(row=r, column=1, value=f"Three cafes, 12 months — support entities (Commissary, SRP, Cafe Mgmt, Central) ran {BT['three_cafes']['support_burden_2026_avg_month']:,} a month on average in Jan–Aug 2026; the Consolidated Model carries no 2025 support figures.").font = NOTE
    r += 1
    wb_.cell(row=r, column=1, value="Admin + training = each cafe's Apr–Aug 2026 monthly average; tips = the cafe's Apr–Aug share of core sales; template wages at Aug 2026 planning rates (2025 actual labor was paid at lower rates). NOI reported includes catering; the core-profit columns exclude it (core sales − COGS scenario − labor − opex).").font = NOTE
    wb_.freeze_panes = "C8"

    # ---------------- Need vs Actual ----------------
    wn = wb.create_sheet("Need vs Actual")
    wn["A1"] = "Half-hour demand, people needed (rules + capacity) and actual core headcount in the sampled weeks — per cafe × tier × day"; wn["A1"].font = TITLE
    wn["A2"] = "need = max(floor, ceil(txns/hr ÷ cafe capacity)); capacity Downtown 14 / Studio 12 / 1899 14; actual = average core headcount on the clock in the sampled weeks of that tier. Green = analysis data."; wn["A2"].font = NOTE
    allt = sorted({x["t"] for c in NV["cafes"].values() for tt in c["tiers"].values() for dd in tt.values() for x in dd["by_halfhour"]})
    def tl(t):
        hh = int(t); return f"{hh}:{'30' if t % 1 else '00'}"
    hdr(wn, 4, ["Cafe", "Tier", "Day", "Series"] + [tl(t) for t in allt] + ["Day total"], [12, 8, 6, 9] + [6] * len(allt) + [9])
    r = 5
    for cafe in OPEN3:
        for tier in ["slow", "mid", "busy"]:
            for d, rec in NV["cafes"][cafe]["tiers"][tier].items():
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


    wb.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    main()
