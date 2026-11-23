#!/usr/bin/env python3
"""Slipstream Scale Model v1.0 — 'how many locations until the Location Task model pays'.
Inputs (blue): store archetypes (sales, COGS %, labor % on templates, opex %, other income), leadership team roster by stage,
commissary/SRP net by stage, founder salary, pay-curve uplift, LT standard % of sales, profit-share %.
Path (formulas): system sales, store contribution, LT cost and % of sales, profit after founder salary, LT headroom vs standard,
sales required to justify the LT, profit-share pool. All archetype numbers come from the Sep 2026 backtest (12 months) and the
Consolidated Model v17 (Central / Commissary / SRP Jan–Aug 2026)."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
HERE = os.path.dirname(os.path.abspath(__file__)); WS = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
OUT = os.path.join(WS, "Slipstream Scale Model - Sep 2026 v1.0.xlsx")
BLUE = Font(name="Calibri", size=10, color="1F3864"); BLUE_FILL = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1"); YELLOW = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GREEN = Font(name="Calibri", size=10, color="006100"); BLACK = Font(name="Calibri", size=10); BOLD = Font(name="Calibri", size=10, bold=True)
HDR = Font(name="Calibri", size=10, bold=True, color="1F3A5F"); HDR_FILL = PatternFill("solid", start_color="E8ECEF", end_color="E8ECEF"); TOT_FILL = PatternFill("solid", start_color="D9E2EC", end_color="D9E2EC")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F3A5F"); NOTE = Font(name="Calibri", size=9, italic=True, color="606060"); THIN = Side(style="thin", color="1F3A5F")
USD = '"$"#,##0;("$"#,##0)'; PCT = "0.0%"
def hdr(ws, row, headers, widths=None):
    for j, t in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=t); c.font = HDR; c.fill = HDR_FILL; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = Border(bottom=THIN)
    if widths:
        for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w
def inp(c, fmt=None, key=False): c.font = BLUE; c.fill = YELLOW if key else BLUE_FILL; c.number_format = fmt or c.number_format
def pull(c, fmt=None): c.font = GREEN; c.number_format = fmt or c.number_format
def fx(c, fmt=None, bold=False): c.font = BOLD if bold else BLACK; c.number_format = fmt or c.number_format

# ---------------- archetypes (from the 12-month backtest, templates at today's rates, COGS at the 30% scenario) ----------------
ARCH = [  # name, sales, cogs%, labor% (P&L on templates), opex%, other NOI $ (catering/events/fees), note
    ("Full cafe (Studio-class)", 586000, 0.30, 0.292, 0.205, 0, "Studio 12-mo: core $586K, labor on templates 29.2%, opex 20.5%; 7 days"),
    ("Weekday cafe (Downtown-class)", 407000, 0.30, 0.278, 0.235, 40000, "Downtown 12-mo: core $407K, labor 27.8%, opex 23.5%; other = catering margin (~$70K last 12 mo, $40K assumed)"),
    ("Managed cafe (1899-type)", 158000, 0.0, 0.0, 0.0, 60000, "NOI passes to the owner; Slipstream keeps $5K/mo fee + 5% of sales (formula adds the 5%)"),
    ("Museum cafe (NMWA-type)", 45000, 0.30, 0.60, 0.20, 96500, "Cafe at Jul–Aug run-rate; other = $5,041/mo fee ($60.5K) + NMWA Nights 8 × $4.5K ($36K)"),
    ("Sprouts (small format)", 200000, 0.30, 0.30, 0.12, 0, "ASSUMPTION — Ryan to confirm: $200K sales, 1 person most of the day + 2–3 h overlap ≈ 55–60 h/wk at ~$23 ≈ 30% labor after tips; opex 12% (low-rent partner space?)"),
]
STAGES = ["Now (Sep 2026)", "2027", "2028", "2029"]
COUNTS = {  # archetype -> counts by stage (blue)
    "Full cafe (Studio-class)": [1, 1, 1, 1], "Weekday cafe (Downtown-class)": [1, 1, 2, 2], "Managed cafe (1899-type)": [1, 1, 1, 1], "Museum cafe (NMWA-type)": [1, 1, 1, 1], "Sprouts (small format)": [0, 2, 5, 8]}
LT = [  # role, cost by stage (blue), note
    ("Ryan (founder) — salary not in the books today", [120000, 120000, 120000, 120000], "≥ $10K/month"),
    ("Mo + Ines (shared management labor, Central)", [118000, 118000, 118000, 118000], "Central 60130 Jan–Aug 2026 $78.9K → $118K/yr"),
    ("Gee — hourly LM (premium over ALL rate)", [0, 12000, 0, 0], "hourly LM: ~$3/h premium on ~40 h/wk, admin hours already in cafe labor"),
    ("Salaried LM #1 (Gee) — 4–5 cafes", [0, 0, 85000, 85000], "replaces the hourly-LM line; her cafe hours come out of cafe labor (not modelled)"),
    ("Food manager over Ines (Sprouts ≥ 3)", [0, 0, 70000, 70000], "trigger: 3+ Sprouts"),
    ("Salaried LM #2 — Sprouts (6–8)", [0, 0, 0, 85000], "trigger: 6+ Sprouts"),
    ("Other central (car, accounting, software, fees)", [34000, 36000, 40000, 45000], "Central non-labor Jan–Aug 2026 $22.8K → $34K/yr"),
]
SUPPORT = [("Commissary net", [-50000, -35000, -25000, -15000], "−$4.1K/mo Jan–Aug 2026; internal pricing + scale should shrink it (Ryan)"), ("SRP net", [-8000, 0, 0, 0], "target breakeven; −$7K Jan–Aug 2026")]

def main():
    wb = Workbook(); wi = wb.active; wi.title = "Inputs"
    wi["A1"] = "Slipstream Scale Model — inputs (blue = change me; yellow = key assumptions)"; wi["A1"].font = TITLE
    wi["A2"] = "Archetype economics from the Sep 2026 12-month backtest (templates at today's rates) and Consolidated Model v17. Every number is per location per year."; wi["A2"].font = NOTE
    hdr(wi, 4, ["Archetype", "Core sales / yr", "COGS %", "Labor % (templates, P&L basis)", "Opex %", "Other NOI $ (catering / fees / events)", "Contribution / yr", "Contribution %", "Note"], [34, 14, 9, 16, 9, 18, 15, 12, 90])
    for i, (n, s, cg, lb, ox, oth, note) in enumerate(ARCH):
        r = 5 + i; wi.cell(row=r, column=1, value=n).font = BLACK
        inp(wi.cell(row=r, column=2, value=s), USD, key=(n.startswith("Sprouts"))); inp(wi.cell(row=r, column=3, value=cg), PCT); inp(wi.cell(row=r, column=4, value=lb), PCT, key=(n.startswith("Sprouts"))); inp(wi.cell(row=r, column=5, value=ox), PCT); inp(wi.cell(row=r, column=6, value=oth), USD)
        if n.startswith("Managed"): fx(wi.cell(row=r, column=7, value=f"=F{r}+0.05*B{r}"), USD)
        else: fx(wi.cell(row=r, column=7, value=f"=B{r}*(1-C{r}-D{r}*(1+$B$13)-E{r})+F{r}"), USD)
        fx(wi.cell(row=r, column=8, value=f"=IF(B{r}>0,G{r}/B{r},0)"), PCT); wi.cell(row=r, column=9, value=note).font = NOTE
    wi["A11"] = "Global assumptions"; wi["A11"].font = BOLD
    wi["A12"] = "Founder salary (per year)"; wi["B12"] = 120000; inp(wi["B12"], USD, key=True)
    wi["A13"] = "Pay-curve uplift on cafe labor (0% = today's rates; 16.5% = peer-average curve; 22.4% = +5%)"; wi["B13"] = 0.0; inp(wi["B13"], PCT, key=True)
    wi["A14"] = "Leadership + central standard — max % of system sales"; wi["B14"] = 0.10; inp(wi["B14"], PCT, key=True)
    wi["A15"] = "Profit share to retirement — % of profit after founder salary"; wi["B15"] = 0.10; inp(wi["B15"], PCT)
    wi["A16"] = "Managed-cafe incentive fee — % of sales"; wi["B16"] = 0.05; inp(wi["B16"], PCT)
    wi["A18"] = "How to read: Contribution = sales × (1 − COGS − labor × (1 + uplift) − opex) + other NOI. Managed cafes contribute only the fee + incentive. Change counts and the leadership roster on the Path sheet."; wi["A18"].font = NOTE
    for r in range(12, 17): wi.cell(row=r, column=1).font = BLACK

    wp = wb.create_sheet("Path")
    wp["A1"] = "Path — locations, leadership team, and what is left after the founder salary, by stage"; wp["A1"].font = TITLE
    wp["A2"] = "Counts and roster are blue (illustrative, Ryan's Sep 2026 sketch); everything else is formulas off the Inputs sheet."; wp["A2"].font = NOTE
    hdr(wp, 4, ["", *STAGES, "Note"], [52, 16, 16, 16, 16, 70])
    r = 5; wp.cell(row=r, column=1, value="LOCATIONS (count)").font = BOLD; r += 1
    cnt_rows = {}
    for i, (n, *_ ) in enumerate(ARCH):
        wp.cell(row=r, column=1, value=n).font = BLACK
        for j in range(4): inp(wp.cell(row=r, column=2 + j, value=COUNTS[n][j]), "0", key=True)
        cnt_rows[n] = r; r += 1
    wp.cell(row=r, column=1, value="Total locations").font = BOLD
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=SUM({L(2+j)}6:{L(2+j)}{r-1})"), "0", bold=True)
    tot_loc = r; r += 2
    wp.cell(row=r, column=1, value="SYSTEM ($ per year)").font = BOLD; r += 1
    wp.cell(row=r, column=1, value="System sales (all locations incl. managed)").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value="=" + "+".join(f"{L(2+j)}{cnt_rows[n]}*Inputs!$B${5+i}" for i, (n, *_ ) in enumerate(ARCH))), USD)
    sales_r = r; r += 1
    wp.cell(row=r, column=1, value="Store contribution (after labor, COGS, opex; fees for managed)").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value="=" + "+".join(f"{L(2+j)}{cnt_rows[n]}*Inputs!$G${5+i}" for i, (n, *_ ) in enumerate(ARCH))), USD)
    contrib_r = r; r += 1
    wp.cell(row=r, column=1, value="Store contribution % of system sales").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=IF({L(2+j)}{sales_r}>0,{L(2+j)}{contrib_r}/{L(2+j)}{sales_r},0)"), PCT)
    r += 2
    wp.cell(row=r, column=1, value="LEADERSHIP TEAM + CENTRAL ($ per year)").font = BOLD; r += 1
    lt0 = r
    for role, costs, note in LT:
        wp.cell(row=r, column=1, value=role).font = BLACK
        for j in range(4): inp(wp.cell(row=r, column=2 + j, value=costs[j]), USD, key=(role.startswith("Ryan")))
        wp.cell(row=r, column=6, value=note).font = NOTE; r += 1
    wp.cell(row=r, column=1, value="Leadership + central, total").font = BOLD
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=SUM({L(2+j)}{lt0}:{L(2+j)}{r-1})"), USD, bold=True)
    for col in range(1, 7): wp.cell(row=r, column=col).fill = TOT_FILL
    lt_r = r; r += 1
    wp.cell(row=r, column=1, value="   as % of system sales").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=IF({L(2+j)}{sales_r}>0,{L(2+j)}{lt_r}/{L(2+j)}{sales_r},0)"), PCT)
    ltpct_r = r; r += 1
    wp.cell(row=r, column=1, value="   sales needed to hold the standard (LT ÷ standard %)").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"={L(2+j)}{lt_r}/Inputs!$B$14"), USD)
    need_r = r; r += 1
    wp.cell(row=r, column=1, value="   sales gap vs standard (negative = short)").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"={L(2+j)}{sales_r}-{L(2+j)}{need_r}"), USD)
    r += 1
    wp.cell(row=r, column=1, value="   LT headroom at the standard (standard % × sales − LT; positive = room to add a seat)").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=Inputs!$B$14*{L(2+j)}{sales_r}-{L(2+j)}{lt_r}"), USD)
    r += 2
    wp.cell(row=r, column=1, value="SUPPORT ENTITIES ($ per year, negative = cost)").font = BOLD; r += 1
    s0 = r
    for role, costs, note in SUPPORT:
        wp.cell(row=r, column=1, value=role).font = BLACK
        for j in range(4): inp(wp.cell(row=r, column=2 + j, value=costs[j]), USD)
        wp.cell(row=r, column=6, value=note).font = NOTE; r += 1
    sup_r = r; wp.cell(row=r, column=1, value="Support, total").font = BOLD
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=SUM({L(2+j)}{s0}:{L(2+j)}{r-1})"), USD, bold=True)
    r += 2
    wp.cell(row=r, column=1, value="RESULT ($ per year)").font = BOLD; r += 1
    wp.cell(row=r, column=1, value="Profit after leadership, central, support and the founder salary").font = BOLD
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"={L(2+j)}{contrib_r}-{L(2+j)}{lt_r}+{L(2+j)}{sup_r}"), USD, bold=True)
    for col in range(1, 7): wp.cell(row=r, column=col).fill = TOT_FILL
    prof_r = r; r += 1
    wp.cell(row=r, column=1, value="   as % of system sales").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=IF({L(2+j)}{sales_r}>0,{L(2+j)}{prof_r}/{L(2+j)}{sales_r},0)"), PCT)
    r += 1
    wp.cell(row=r, column=1, value="   profit-share pool (Inputs B15 × profit, if positive)").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=MAX(0,{L(2+j)}{prof_r})*Inputs!$B$15"), USD)
    r += 1
    wp.cell(row=r, column=1, value="   profit per location").font = BLACK
    for j in range(4): fx(wp.cell(row=r, column=2 + j, value=f"=IF({L(2+j)}{tot_loc}>0,{L(2+j)}{prof_r}/{L(2+j)}{tot_loc},0)"), USD)
    r += 2
    wp.cell(row=r, column=1, value="Set Inputs!B13 to 16.5% to see every stage with the peer-average pay curve; 22.4% for +5%. The Sprouts row is the least certain input — labor % and opex % are guesses until the first one opens.").font = NOTE
    wp.freeze_panes = "B5"
    wb.save(OUT); print("Saved →", OUT)

if __name__ == "__main__":
    main()
