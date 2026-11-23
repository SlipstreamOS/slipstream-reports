#!/usr/bin/env python3
"""Build 'Slipstream Schedule Templates - Sep 2026 v1.3.xlsx' from templates_v13.json (Sep 5 + Sep 10 rules; no admin sets).
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
OUT = os.path.join(WS, "Slipstream Schedule Templates - Sep 2026 v1.3.xlsx")
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
    ws["A1"] = "Slipstream Schedule Templates by Demand Tier — September 2026 v1.3"; ws["A1"].font = TITLE
    lines = [
        "Companion to 'Slipstream Labor Deep Dive - Schedule Templates by Tier - Sep 2026 v1.3.docx'. Rules: two openers (kitchen 6:15, lead 6:30) and two closers (to close+30) at Downtown & Studio; 6-hour minimum shifts; Downtown third person by 8:00 Tue–Thu, max 4, Tue/Wed/Thu one template; Studio line person on all day Mon–Thu, third person 8:00–2:30, max 4 on Friday and slow weekends, 5 on mid/busy weekends; 1899 two shifts every day (stagger when slow, both close together); NMWA 9:00–3:00. Admin days are NOT in the templates (scheduler's tool). Planning rates = Aug 2026 roster averages by role. Slow-month ceiling +2 pts.",
        "",
        "COLOR PROTOCOL — blue text on light blue = input you can change (rates, targets, capacity, typical sales, shift start/end; yellow = key assumptions); green = pulled from the analysis data (need/actual curves, first-cut transcription, simulation results); black = formula or label.",
        "",
        "HOW TO USE",
        "1. Pick the tier for the week from the projected weekly core sales in the weekly report's forward block (bands on the Inputs sheet). Below the band floor → use the next slower template. Build the 10/11 week on MID.",
        "2. Open the cafe sheet, copy the shifts for that tier into Square Teams. 'Solo' notes mark the windows one person runs the floor; 'Admin covers' marks half-hours where the need exceeds the template and the ALL's block is the intended cover.",
        "3. The day and week summaries recompute labor % from the Inputs rates and typical sales. Change a start/end time and everything flows. Status: OK = at/under target, BAND = between target and ceiling, OVER = above ceiling (ask Ryan before publishing). The 'roster' % re-prices the same hours at the cafe's current average core rate (Inputs).",
        "4. Summary sheet: every cafe × tier × day, weekly rollups, −10% sales sensitivity and the ceiling trigger (weekly sales below which the template breaches the ceiling).",
        "5. First Cut 10-11: Ryan's Square Teams week transcribed and scored on the same basis, with the specific change to each shift. Tier Calendar: the Sep–Dec weekly tiers on Projections v1.9. Pay Review: the peer comparison and the cost of the proposed curve.",
        "6. Incentive: the four options bench-tested (Apr–Aug 2026 backtest + 2027 simulation) and a live Option 4 calculator (enter a month's sales, core wages and template wages per cafe).",
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
        inp(wi.cell(row=r, column=2, value=S["targets"][cafe]["target"]), "0%", key=True); inp(wi.cell(row=r, column=3, value=S["targets"][cafe]["ceiling"]), "0%", key=True)
        inp(wi.cell(row=r, column=4, value=STAFF["rolemix"][cafe]["avg_core_rate"]), '"$"#,##0.00'); inp(wi.cell(row=r, column=5, value=NV["cap"][cafe]), "0.0", key=True); inp(wi.cell(row=r, column=6, value=S["targets"][cafe]["ceiling"] + 0.02), "0%", key=True)
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

    # ---------------- Pay Review ----------------
    wpay = wb.create_sheet("Pay Review")
    wpay["A1"] = "Pay review — Slipstream's guaranteed ladder vs DC peers (base + tips), and the cost of a steeper curve"; wpay["A1"].font = TITLE
    wpay["A2"] = "Green = pulled (Aug 2026 timecards, transactions, public postings); blue = the proposed ladder (edit to re-cost); black = formulas."; wpay["A2"].font = NOTE
    hdr(wpay, 4, ["Peer", "Base", "Tips / hr", "Barista all-in", "Shift lead", "Source"], [34, 30, 12, 16, 18, 44])
    r = 5
    for x in PAY["peers"]:
        for j, k in enumerate(["who", "base", "tips", "all_in", "sl", "src"]): pull(wpay.cell(row=r, column=1 + j, value=x[k]))
        r += 1
    r += 1
    wpay.cell(row=r, column=1, value="Slipstream today (Aug 2026 core hours)").font = BOLD; r += 1
    hdr(wpay, r, ["Band", "People", "Share of core hours", "Avg rate", "Hours / month"]); r += 1
    for b in PAY["slipstream_now"]["bands"]:
        pull(wpay.cell(row=r, column=1, value=b["band"])); pull(wpay.cell(row=r, column=2, value=b["people"])); pull(wpay.cell(row=r, column=3, value=b["share_hours"] / 100), "0%"); pull(wpay.cell(row=r, column=4, value=b["avg_rate"]), '"$"#,##0.00'); pull(wpay.cell(row=r, column=5, value=b["hours_month"]), "#,##0")
        r += 1
    r += 1
    wpay.cell(row=r, column=1, value="Tips per core hour (Aug 2026) — applied toward the guaranteed rate").font = BOLD; r += 1
    for cafe, t in PAY["slipstream_now"]["tips"].items():
        wpay.cell(row=r, column=1, value=cafe).font = BLACK; pull(wpay.cell(row=r, column=2, value=t["tips_per_core_hour"]), '"$"#,##0.00'); pull(wpay.cell(row=r, column=3, value=t["tip_pct_of_net"] / 100), "0.0%"); r += 1
    r += 1
    wpay.cell(row=r, column=1, value="Proposed ladder (edit the blue cells) and its cost on the August roster").font = BOLD; r += 1
    hdr(wpay, r, ["Level", "Proposed low", "Proposed high", "Why"]); r += 1
    for x in PAY["ladder_proposed"]:
        wpay.cell(row=r, column=1, value=x["level"]).font = BLACK; inp(wpay.cell(row=r, column=2, value=x["low"]), '"$"#,##0.00', key=True); inp(wpay.cell(row=r, column=3, value=x["high"]), '"$"#,##0.00', key=True); wpay.cell(row=r, column=4, value=x["note"]).font = BLACK; r += 1
    r += 1
    hdr(wpay, r, ["Band today", "People", "Avg rate now", "Proposed (full curve)", "Hours / month", "Cost / month (full curve)"]); r += 1
    tot_row0 = r
    for b in PAY["slipstream_now"]["bands"]:
        bb = PAY["cost_of_proposed"]["by_band"][b["band"]]
        pull(wpay.cell(row=r, column=1, value=b["band"])); pull(wpay.cell(row=r, column=2, value=bb["people"])); pull(wpay.cell(row=r, column=3, value=bb["avg_now"]), '"$"#,##0.00')
        inp(wpay.cell(row=r, column=4, value=bb["proposed"]), '"$"#,##0.00', key=True); pull(wpay.cell(row=r, column=5, value=b["hours_month"]), "#,##0"); fx(wpay.cell(row=r, column=6, value=f"=(D{r}-C{r})*E{r}"), '"$"#,##0')
        r += 1
    wpay.cell(row=r, column=1, value="Total").font = BOLD; fx(wpay.cell(row=r, column=6, value=f"=SUM(F{tot_row0}:F{r-1})"), '"$"#,##0', bold=True)
    for col in range(1, 7): wpay.cell(row=r, column=col).fill = TOT_FILL
    r += 1
    wpay.cell(row=r, column=1, value=f"Step 1 (parity floors only: SL ≥ $22.50, ALL ≥ $25.50): ${PAY['cost_step1']['month']:,} / month, {PAY['cost_step1']['pct_of_core_sales_aug']} points of core sales. Full curve: ${PAY['cost_of_proposed']['month']:,} / month, {PAY['cost_of_proposed']['pct_of_core_sales_aug']} points.").font = NOTE
    wpay.freeze_panes = "A5"

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

    # ---------------- Incentive ----------------
    wc = wb.create_sheet("Incentive")
    wc["A1"] = "Incentive — Option 4 calculator, the four options bench-tested (2026 backtest, 2027 simulation)"; wc["A1"].font = TITLE
    wc["A2"] = "Calculator: enter a month's core sales, core wages and template wages per cafe (blue); on-template fixed amounts, the beat share, cap and payout are formulas. Below: backtest and simulation results (green)."; wc["A2"].font = NOTE
    wc["A4"] = "Option 4 calculator — on-template fixed $ + 50% of the beat below template, cap 1.5% of sales"; wc["A4"].font = BOLD
    hdr(wc, 5, ["Cafe", "Core sales (month)", "Core wages (month)", "Template wages (month)", "On-template $", "Tolerance", "Qualifies?", "Fixed earned", "Beat $ (template − actual)"], [22, 16, 16, 18, 13, 10, 10, 12, 20])
    FIXED = {"Downtown": 300, "Studio": 300, "1899 Penn": 150, "NMWA": 75}
    for i, cafe in enumerate(["Downtown", "Studio", "1899 Penn", "NMWA"]):
        r = 6 + i
        wc.cell(row=r, column=1, value=cafe).font = BLACK
        if cafe != "NMWA":
            wk = T["cafes"][cafe]["mid"]["week"]; inp(wc.cell(row=r, column=2, value=round(wk["sales_typical"] * 52 / 12)), '"$"#,##0', key=True)
            inp(wc.cell(row=r, column=3, value=round(wk["wages"] * 52 / 12 * 0.98)), '"$"#,##0', key=True); inp(wc.cell(row=r, column=4, value=round(wk["wages"] * 52 / 12)), '"$"#,##0', key=True)
        else:
            inp(wc.cell(row=r, column=2, value=4593), '"$"#,##0', key=True); inp(wc.cell(row=r, column=3, value=2250), '"$"#,##0', key=True); inp(wc.cell(row=r, column=4, value=2210), '"$"#,##0', key=True)
        inp(wc.cell(row=r, column=5, value=FIXED[cafe]), '"$"#,##0', key=True); inp(wc.cell(row=r, column=6, value=0.01), "0%")
        fx(wc.cell(row=r, column=7, value=f'=IF(C{r}<=D{r}*(1+F{r}),"yes","no")')); fx(wc.cell(row=r, column=8, value=f'=IF(G{r}="yes",E{r},0)'), '"$"#,##0'); fx(wc.cell(row=r, column=9, value=f"=D{r}-C{r}"), '"$"#,##0;("$"#,##0)')
    wc.cell(row=10, column=1, value="All cafes").font = BOLD
    for col in (2, 3, 4, 8, 9): fx(wc.cell(row=10, column=col, value=f"=SUM({L(col)}6:{L(col)}9)"), '"$"#,##0;("$"#,##0)', bold=True)
    for col in range(1, 10): wc.cell(row=10, column=col).fill = TOT_FILL
    for i, (lab, val, fmt) in enumerate([("Share of the beat", 0.5, "0%"), ("Cap (% of all-cafe core sales)", 0.015, "0.0%"), ("Pool before cap = fixed earned + share × max(0, beat)", "=H10+B12*MAX(0,I10)", '"$"#,##0'),
                                          ("Cap $", "=B13*B10", '"$"#,##0'), ("Payout (capped)", "=MIN(B14,B15)", '"$"#,##0'), ("Bonus % of core wages (percentage-of-wages split)", "=IF(C10>0,B16/C10,0)", "0.00%"),
                                          ("Employer cost incl. 12% load", f"=B16*{LOAD_CELL}", '"$"#,##0')]):
        r = 12 + i; wc.cell(row=r, column=1, value=lab).font = BLACK; c = wc.cell(row=r, column=2, value=val)
        if isinstance(val, (int, float)): inp(c, fmt, key=True)
        else: fx(c, fmt, bold=lab.startswith("Payout"))
    wc.cell(row=20, column=1, value="Guardrail (not modelled): schedules met the floors (two open / two close, no unplanned solo), core sales ≥ 90% of the re-based projection, admin hours tracked alongside.").font = NOTE
    r = 22
    wc.cell(row=r, column=1, value="Backtest — Apr–Aug 2026 actual months (what each option would have paid)").font = BOLD; r += 1
    hdr(wc, r, ["Month", "Downtown tier / core % / template $ / actual $", "Studio", "1899 (NI %)", "Opt 1 Open-cafe pool", "Opt 2 Own-metric", "Opt 3 Template beat", "Opt 4 On-template + beat"]); r += 1
    for b in INC["backtest_2026"]:
        wc.cell(row=r, column=1, value=b["month"]).font = BLACK
        for j, x in enumerate(b["cafes"]):
            pull(wc.cell(row=r, column=2 + j, value=f"{x['tier']} / {x['core_pct']}% / ${x['tmpl_wages']:,} / ${x['wages']:,}" + (f" / NI {x['ni_pct']}%" if x["cafe"] == "1899 Penn" else "")))
        for j, k in enumerate(["opt1", "opt2", "opt3", "opt4"]): pull(wc.cell(row=r, column=5 + j, value=b[k]), '"$"#,##0')
        r += 1
    r += 1
    sim = INC["simulation_2027"]
    wc.cell(row=r, column=1, value=f"2027 simulation — {sim['runs']} runs; labor saved vs Apr–Aug run-rate p50 ${sim['labor_saving_vs_aprAug_run_rate']['p50']:,}").font = BOLD; r += 1
    hdr(wc, r, ["Option", "Annual payout p10", "p50", "p90", "Months paying (p50)", "Share of labor saved", "Per 30 h/wk person (p50, year)"]); r += 1
    for k in ["opt1", "opt2", "opt3", "opt4"]:
        v = sim["options"][k]; wc.cell(row=r, column=1, value=INC["definitions"][k]).font = BLACK
        pull(wc.cell(row=r, column=2, value=v["p10"]), '"$"#,##0'); pull(wc.cell(row=r, column=3, value=v["p50"]), '"$"#,##0'); pull(wc.cell(row=r, column=4, value=v["p90"]), '"$"#,##0')
        pull(wc.cell(row=r, column=5, value=v["months_paying_p50"]), "0"); pull(wc.cell(row=r, column=6, value=v["share_of_saving_p50"] / 100), "0.0%"); pull(wc.cell(row=r, column=7, value=INC["per_person_30h_year_p50"][k]), '"$"#,##0')
        r += 1
    wc.column_dimensions["A"].width = 70
    for col in "BCDEFGHI": wc.column_dimensions[col].width = max(wc.column_dimensions[col].width or 0, 18)
    wc.freeze_panes = "A5"

    wb.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    main()
