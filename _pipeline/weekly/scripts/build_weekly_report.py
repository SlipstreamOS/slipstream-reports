#!/usr/bin/env python3
"""Generate the weekly business review .docx from the weekly data workbook.

Audience: ALL team + Leadership.
Produces a concise review doc with "The Formula" header (reinforces Finance 101),
a Targets Scorecard (Prime/Labor vs handout bands), wins/watch callouts, daily
Budget-vs-Actual (Core basis), weekly sales by category (cafes as columns),
labor & COGS, ops metrics (cafes as columns), forward 4-week targets.

Core-denominator rule: all %s use CORE Sales (ex-catering/invoice payments).

Versions:
  v0.1 — initial build, dummy projections.
  v1.0 — projections wired; forward 4-week targets added.
  v1.1 — language aligned with Cafe Finance 101 handout v4 (Core Sales / Labor /
         COGS); The Formula header; Targets Scorecard; cross-sectional tables
         flipped to metrics-as-rows × cafes-as-cols to match Monthly Report v1.3;
         Forward Targets restructured to one-table-per-cafe; Historical Context
         section populated from transactions store.
  v1.2 — Labor scope expanded to ALL roles (FOH + BOH combined) per Apr 28
         directive — matches Square Teams labor cost reports. Tip / transaction-
         count / avg-ticket source switched from Orders API → Payments API,
         which matches Square's Sales Summary (Orders API misses post-close
         tip adjustments).
"""
import json, os
from datetime import datetime, date, timedelta
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from _paths import DATA_DIR, WORKSPACE    # noqa — side effect: puts shared/scripts on sys.path
from projections import (load_projections, load_forward_week_targets, load_labor_targets,
                         CORE_LABOR_TARGETS, NMWA_STAFF_RULE, TRAINING_PLAN_PCT,
                         TRAINING_AMBER_PCT, is_holiday_closed)
import transactions_store as txn_store

# --- Week config --------------------------------------------------------------
WEEK_LABEL = "nov15-nov21"
WEEK_TITLE = "Week of Sunday Nov 15 – Saturday Nov 21, 2026"
WEEK_FILE_TAG = "Nov 15-Nov 21 2026"
VERSION = "v1.2"
SHOW_OUTSIDE_CORE = False   # Admin/Training table under the scorecard

SALES_JSON    = os.path.join(DATA_DIR, f"week_{WEEK_LABEL}.json")
LABOR_JSON    = os.path.join(DATA_DIR, f"labor_{WEEK_LABEL}.json")
PAYMENTS_JSON = os.path.join(DATA_DIR, f"payments_{WEEK_LABEL}.json")  # Square Payments-derived
OUT = os.path.join(WORKSPACE, f"Slipstream Weekly Report - {WEEK_FILE_TAG} {VERSION}.docx")

CAFES = ["Downtown", "Studio", "1899 Penn", "NMWA"]
OPEN_CAFES = ["Downtown", "Studio", "1899 Penn"]   # "Open Cafes" totals column — excludes pre-open NMWA
DAYS = ["2026-11-15","2026-11-16","2026-11-17","2026-11-18",
        "2026-11-19","2026-11-20","2026-11-21"]
DAY_LABELS = ["Sun 11/15","Mon 11/16","Tue 11/17","Wed 11/18",
              "Thu 11/19","Fri 11/20","Sat 11/21"]

COGS_PCT = {"coffee_bev":0.22,"food":0.32,"retail_coffee":0.55,"catering":0.30,"other":0.30}

# Projections for the reporting week + the next 4 forward weeks + labor targets.
PROJECTIONS, PROJECTIONS_SOURCE = load_projections(DAYS[0], DAYS[-1], CAFES)
LABOR_TARGETS, _ = load_labor_targets(CAFES)

def _next_sunday_iso(last_day_iso):
    return (date.fromisoformat(last_day_iso) + timedelta(days=1)).isoformat()

FORWARD_WEEKS, _ = load_forward_week_targets(_next_sunday_iso(DAYS[-1]), num_weeks=4, cafes=CAFES)

# --- Formatting helpers ------------------------------------------------------

def money(cents_or_dollars, is_cents=True):
    v = cents_or_dollars/100 if is_cents else cents_or_dollars
    if v == 0: return "$0"
    return f"${v:,.0f}" if abs(v) >= 1 else f"${v:,.2f}"

def pct(x):
    return f"{x*100:.1f}%" if x else "0.0%"

def set_cell_shading(cell, hex_color):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def bold_cell(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            r.bold = True

def center_cell(cell):
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def set_font(run, name="Calibri", size=10, bold=False, color=None):
    run.font.name = name; run.font.size = Pt(size); run.bold = bold
    if color:
        run.font.color.rgb = color

# --- Data layer --------------------------------------------------------------

def load():
    sales = json.load(open(SALES_JSON))
    labor = json.load(open(LABOR_JSON))
    payments = json.load(open(PAYMENTS_JSON)) if os.path.exists(PAYMENTS_JSON) else None
    return sales, labor, payments

def day_info(sales, cafe, d):
    return sales["per_cafe_per_day"].get(cafe,{}).get(d)

def day_labor(labor, cafe, d):
    """Total cafe labor (FOH + BOH combined per Apr 2026 directive).

    Tolerates legacy JSONs that only have foh_cents.
    """
    raw = labor["per_cafe_per_day"].get(cafe, {}).get(d, {})
    if "labor_cents" in raw:
        return raw
    return {"labor_cents": raw.get("foh_cents", 0),
            "labor_hours": raw.get("foh_hours", 0.0),
            "shifts":      raw.get("shifts", 0)}

def day_payments(payments, cafe, d):
    if not payments: return None
    return payments["per_cafe_per_day"].get(cafe, {}).get(d)

def day_labor_buckets(labor, cafe, d):
    """Per-day (core, admin, training) labor cents. Tolerates legacy JSONs
    that predate the Aug-2026 bucket split by treating everything as core."""
    raw = labor["per_cafe_per_day"].get(cafe, {}).get(d, {})
    if "core_cents" in raw:
        return (raw.get("core_cents", 0), raw.get("admin_cents", 0), raw.get("training_cents", 0))
    return (raw.get("labor_cents", raw.get("foh_cents", 0)), 0, 0)


def week_labor_buckets(labor, cafe):
    core = admin = train = 0
    for d in DAYS:
        c, a, t = day_labor_buckets(labor, cafe, d)
        core += c; admin += a; train += t
    return {"core": core, "admin": admin, "training": train}


def week_totals(sales, labor, cafe, payments=None):
    tx_orders = tip_orders = 0
    tx_pay = tip_pay = 0
    net = cat = cbev = food = retail = other = tax = disc = labor_total = 0
    for d in DAYS:
        info = day_info(sales, cafe, d)
        if info:
            tx_orders += info["orders"]
            net += info["net_sales_cents"]
            cat += info["by_category"].get("catering",0)
            cbev += info["by_category"].get("coffee_bev",0)
            food += info["by_category"].get("food",0)
            retail += info["by_category"].get("retail_coffee",0)
            other += info["by_category"].get("other",0)
            tip_orders += info["tip_cents"]
            tax += info["tax_cents"]; disc += info["discount_cents"]
        pay = day_payments(payments, cafe, d)
        if pay:
            tx_pay  += pay["register_payments"]
            tip_pay += pay["register_tip_cents"]
        labor_total += day_labor(labor, cafe, d)["labor_cents"]
    # Prefer Payments-derived counts/tips (matches Square Sales Summary).
    tx   = tx_pay  if payments else tx_orders
    tips = tip_pay if payments else tip_orders
    core = net - cat
    coffee_cogs = cbev * COGS_PCT["coffee_bev"] + retail * COGS_PCT["retail_coffee"]
    food_cogs   = food * COGS_PCT["food"]
    other_cogs  = cat * COGS_PCT["catering"] + other * COGS_PCT["other"]
    # For Core-only COGS (what ties to Prime % of Core Sales), we exclude catering COGS.
    cogs_core   = cbev * COGS_PCT["coffee_bev"] + food * COGS_PCT["food"] + retail * COGS_PCT["retail_coffee"] + other * COGS_PCT["other"]
    bk = week_labor_buckets(labor, cafe)
    return {"core_labor":bk["core"],"admin_labor":bk["admin"],"training_labor":bk["training"],
            "tx":tx,"net":net,"core":core,"catering":cat,"cbev":cbev,"food":food,
            "retail":retail,"other":other,"tips":tips,"tax":tax,"disc":disc,
            "labor":labor_total,"foh":labor_total,  # 'foh' alias kept temporarily for downstream access
            "coffee_cogs":coffee_cogs,"food_cogs":food_cogs,
            "cogs":cogs_core,"prime":bk["core"]+cogs_core}

def sum_totals(rows, keys=None):
    """Combine week_totals dicts across cafes."""
    agg = {}
    for r in rows:
        for k, v in r.items():
            agg[k] = agg.get(k, 0) + v
    return agg

# --- Table builders ----------------------------------------------------------

def add_table(doc, rows, cols, headers=None, header_fill="1F4E79"):
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Light Grid Accent 1"
    if headers:
        for i, h in enumerate(headers):
            c = table.rows[0].cells[i]; c.text = h
            bold_cell(c); center_cell(c); set_cell_shading(c, header_fill)
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF); r.font.size = Pt(9)
    return table


# Colors used by the financial-statement style (sections 1-3) — match the
# Slipstream Income Statement template Ryan referenced.
FIN_TOTAL_REVENUE = "D6EAD3"   # light green   (Total Sales)
FIN_TOTAL_COGS    = "FCE5B4"   # light amber   (Total COGS)
FIN_TOTAL_LABOR   = "F4C7A4"   # peach         (Total Labor)
FIN_TOTAL_PRIME   = "DDE5F1"   # light blue    (Prime Cost / Margin)
FIN_HEADER_BORDER = "9BAFCB"   # mid blue-grey (just under col headers)
FIN_GRID_GREY     = "BFBFBF"


def _set_cell_borders(cell, *, top=None, bottom=None, left=None, right=None):
    """Set explicit per-side borders on a cell. None = leave alone; 'none' = remove; 'thin'/'thick' = set."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders"); tcPr.append(tcBorders)
    for side, spec in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        if spec is None:
            continue
        existing = tcBorders.find(qn(f"w:{side}"))
        if existing is not None:
            tcBorders.remove(existing)
        b = OxmlElement(f"w:{side}")
        if spec == "none":
            b.set(qn("w:val"), "nil")
        else:
            sz = "12" if spec == "thick" else "4"
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), sz)
            b.set(qn("w:color"), FIN_GRID_GREY)
        tcBorders.append(b)


def _strip_table_borders(table):
    """Remove ALL borders from every cell of a table — gives the financial-statement clean look."""
    for row in table.rows:
        for c in row.cells:
            _set_cell_borders(c, top="none", bottom="none", left="none", right="none")
    # also clear table-level borders
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is not None:
        existing = tblPr.find(qn("w:tblBorders"))
        if existing is not None:
            tblPr.remove(existing)
        tblBorders = OxmlElement("w:tblBorders")
        for side in ("top","left","bottom","right","insideH","insideV"):
            b = OxmlElement(f"w:{side}")
            b.set(qn("w:val"), "nil")
            tblBorders.append(b)
        tblPr.append(tblBorders)


def fin_table(doc, n_rows, n_cols, *, col_headers, label_col_width_in=2.4, val_col_width_in=0.95):
    """Create a borderless 'financial statement' style table.

    col_headers: list of column header strings. First column is the row label
    (left-aligned, bold for sections); remaining columns are numeric.

    Headers are bold black centered, with a thin bottom rule under the header
    row (mimicking the Income Statement screenshot). All other borders removed.
    """
    table = doc.add_table(rows=n_rows, cols=n_cols)
    # No table style — strip everything explicitly
    table.autofit = False
    _strip_table_borders(table)

    # Column widths
    from docx.shared import Inches
    for i, col in enumerate(table.columns):
        for cell in col.cells:
            cell.width = Inches(label_col_width_in if i == 0 else val_col_width_in)

    # Header row
    for j, h in enumerate(col_headers):
        c = table.rows[0].cells[j]
        c.text = h
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            for r in p.runs:
                r.font.size = Pt(10)
                r.bold = True
        # Thin bottom rule under the header row
        _set_cell_borders(c, bottom="thin")
    return table


def fin_write_row(table, row_idx, label, values, *, kind="detail", indent=0, fill=None,
                  pct_left=None, pct_color=None):
    """Write a row in the financial-statement style.

    kind:
      'section' — bold black label, no fill, slightly larger (10pt). For 'Sales', 'Cost of Goods Sold', 'Labor', 'Prime Cost'.
      'detail'  — regular weight, indented label.
      'total'   — bold label + bold values, full-row colored fill (Income-Statement-style highlight).

    indent: number of leading spaces on the label.
    pct_left: optional small % stamp drawn before the label (like '36.5%' on the Payroll row).
    """
    cells = table.rows[row_idx].cells
    label_text = (" " * (indent * 4)) + label
    label_cell = cells[0]
    label_cell.text = ""
    p = label_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if pct_left:
        r = p.add_run(pct_left + "  ")
        r.font.size = Pt(9); r.font.bold = True
        r.font.color.rgb = pct_color or RGBColor(0x2E, 0x7D, 0x32)
    r = p.add_run(label_text)
    r.font.size = Pt(10)
    if kind == "section":
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    elif kind == "total":
        r.font.bold = True

    # Values
    for j, v in enumerate(values, start=1):
        c = cells[j]
        c.text = str(v) if v is not None else ""
        for pp in c.paragraphs:
            pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for rr in pp.runs:
                rr.font.size = Pt(10)
                if kind == "total":
                    rr.font.bold = True

    # Fill
    if fill:
        for c in cells:
            set_cell_shading(c, fill)

def write_row(table, row_idx, values, *, bold=False, section=False, indent=False,
              shade=None, align_right_from=1, size=9):
    """Write a data row in a cross-sectional (metrics-as-rows) table."""
    for j, v in enumerate(values):
        c = table.rows[row_idx].cells[j]
        c.text = str(v)
        if shade:
            set_cell_shading(c, shade)
        for p in c.paragraphs:
            if j >= align_right_from:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for r in p.runs:
                r.font.size = Pt(size)
                if bold:
                    r.bold = True
                if indent and j == 0:
                    # leading spaces in label already handle visual indent
                    pass
                if section:
                    r.bold = True
    if section:
        for c in table.rows[row_idx].cells:
            set_cell_shading(c, "EAEFF5")

def formula_box(doc):
    """Render 'The Formula' header box — matches Monthly Report v1.3."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Light Grid Accent 1"
    c = t.rows[0].cells[0]
    set_cell_shading(c, "F3F5F7")
    c.text = ""
    p1 = c.paragraphs[0]
    set_font(p1.add_run("The Formula"), size=11, bold=True, color=RGBColor(0x1F,0x3A,0x5F))
    p2 = c.add_paragraph()
    set_font(p2.add_run(
        "Control COGS  +  Protect Labor Hours  +  Grow Sales  =  Better Business + Better Pay. "
        "Every headline below maps back to one of these three levers."
    ), size=10)

def core_labor_rag(core_pct, cafe):
    """Four-state read on CORE wages as % of core walk-in sales.

    Ryan, Sep 2026: each cafe has an operating RANGE and a hard MAX.
      GREEN  at/under target  — core_pct <= range low
      GREEN  in range         — between low and high (the planned zone)
      AMBER  over range       — between high and max (lead's discretion)
      RED    over max         — above max; needs Ryan's pre-approval before
                                the schedule is published
    NMWA is excluded by design — staffing rule, not a %.
    """
    GREEN = RGBColor(0x2E,0x7D,0x32)
    AMBER = RGBColor(0xB7,0x79,0x0C)
    RED   = RGBColor(0xC6,0x28,0x28)
    GREY  = RGBColor(0x88,0x88,0x88)
    band = CORE_LABOR_TARGETS.get(cafe)
    if not band or not core_pct:
        return ("\u2014", GREY, None)
    lo, hi, mx = band["low"], band["high"], band["max"]
    if core_pct <= lo:
        return ("\u25cf GREEN \u2014 at/under target", GREEN, band)
    if core_pct <= hi:
        return ("\u25cf GREEN \u2014 in range", GREEN, band)
    if core_pct <= mx:
        return ("\u25cf AMBER \u2014 over range, under max", AMBER, band)
    return ("\u25cf RED \u2014 over max", RED, band)


# Per-cafe Prime bands (healthy up to, watch up to). Default = Cafe Finance 101 (60 / 65).
# 1899 Penn (Ryan, Nov 17 2026): labor range top 40% + ~25% COGS -> healthy to 65%; labor max 42% + COGS -> watch to 67%.
PRIME_BANDS = {"1899 Penn": (0.65, 0.67)}

def scorecard_tag(value_pct, target_pct, *, kind):
    """Return (tag_char, hex_color) for a ✓/△/✗ scorecard cell.

    Bands (per Apr 28 2026 directive — same threshold for both Labor and Prime):
      ≤ target           → ✓ Healthy / On target     (green)
      target … target+3pp → △ Watch  (within 3pp over — single-week variability OK)
      > target+3pp        → ✗ Problem (red)

    kind:
      'prime' — target is 60% (Cafe Finance 101 'Healthy' threshold).
      'labor' — target_pct is the per-cafe Projections workbook value.
    """
    if not value_pct:
        return ("—", RGBColor(0x88,0x88,0x88))
    GREEN = RGBColor(0x2E,0x7D,0x32)
    AMBER = RGBColor(0xB7,0x79,0x0C)
    RED   = RGBColor(0xC6,0x28,0x28)
    if kind == "prime":
        # Per Apr 28 directive: tolerance buffer is wider for Prime than Labor.
        # ≤60 Healthy · 60–65 Watch · >65 Problem (matches handout bands).
        hi, mx = PRIME_BANDS.get(target_pct, (0.60, 0.65)) if isinstance(target_pct, str) else (0.60, 0.65)
        if value_pct <= hi:  return ("✓ Healthy",  GREEN)
        if value_pct <= mx:  return ("△ Watch",    AMBER)
        return                        ("✗ Problem",  RED)
    if kind == "labor":
        if not target_pct:
            return ("—", RGBColor(0x88,0x88,0x88))
        diff = value_pct - target_pct
        if diff < -0.01:     return ("✓ Under target", GREEN)
        if diff <= 0.01:     return ("✓ On target",    GREEN)
        if diff <= 0.03:     return ("△ Watch",        AMBER)
        return                      ("✗ Problem",      RED)
    return ("—", RGBColor(0x88,0x88,0x88))

# --- Main --------------------------------------------------------------------

def main():
    sales, labor, payments = load()
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.6); s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.7); s.right_margin = Inches(0.7)

    # ---- Title + subtitle ----
    title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(title.add_run("Slipstream Weekly Business Review"), size=18, bold=True)
    sub = doc.add_paragraph(); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(sub.add_run(WEEK_TITLE), size=12, bold=True, color=RGBColor(0x1F,0x4E,0x79))
    v = doc.add_paragraph(); v.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(v.add_run(
        f"{VERSION} · audience: ALL team + Leadership · aligned with the Cafe Finance 101 handout · "
        "all %s are of Core Sales (ex-catering/invoice payments)"
    ), size=9, color=RGBColor(0x66,0x66,0x66))

    # ---- The Formula box ----
    formula_box(doc)

    # ---- Per-cafe totals (used everywhere below) ----
    totals_by_cafe = {c: week_totals(sales, labor, c, payments) for c in CAFES}
    open_totals = sum_totals([totals_by_cafe[c] for c in OPEN_CAFES])

    # ==== Targets Scorecard =================================================
    doc.add_heading("Targets Scorecard", level=1)
    p = doc.add_paragraph()
    set_font(p.add_run(
        "CORE labor only — Admin (Square job title \u201cManager\u201d) and Training are excluded. "
        "Basis is core wages \u00f7 core walk-in sales (invoice/catering payments excluded). "
        "Each cafe has an operating RANGE and a hard MAX: green = inside the range (or better), amber = "
        "over the range but under the max (lead\u2019s discretion), red = over the max (needs Ryan\u2019s "
        "pre-approval before the schedule is published). NMWA runs on a staffing rule, not a labor %. "
        "Prime bands per Cafe Finance 101 \u2014 "
        "55\u201360% healthy, 60\u201365% watch, above 65% problem; 1899 Penn healthy to 65%, watch to 67% (in line with its labor range)."
    ), size=9, color=RGBColor(0x44,0x44,0x44))

    hdr = ["Cafe", "Core Sales", "Core Labor $", "Core Labor %", "Range", "Max", "Status", "Prime %"]
    scored = [c for c in CAFES if c in CORE_LABOR_TARGETS]
    t = add_table(doc, rows=len(scored)+1, cols=len(hdr), headers=hdr)
    for i, cafe in enumerate(scored, start=1):
        w = totals_by_cafe[cafe]
        bk = week_labor_buckets(labor, cafe)
        core_v = w["core"]
        core_lab_pct = bk["core"]/core_v if core_v else 0
        prime_pct = (w["core_labor"]+w["cogs"])/core_v if core_v else 0
        tag, color, band = core_labor_rag(core_lab_pct, cafe)
        prime_tag, prime_color = scorecard_tag(prime_pct, cafe, kind="prime")

        vals = [cafe,
                money(core_v) if core_v else "—",
                money(bk["core"]),
                pct(core_lab_pct) if core_v else "—",
                f"{pct(band['low'])}–{pct(band['high'])}" if band else "—",
                pct(band["max"]) if band else "—",
                tag,
                f"{pct(prime_pct)}  {prime_tag}" if core_v else "—"]
        for j, val in enumerate(vals):
            c = t.rows[i].cells[j]; c.text = val
            for pp in c.paragraphs:
                if j >= 1: pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in pp.runs: r.font.size = Pt(9)
        for pp in t.rows[i].cells[6].paragraphs:
            for r in pp.runs: r.font.color.rgb = color; r.bold = True
        for pp in t.rows[i].cells[7].paragraphs:
            for r in pp.runs: r.font.color.rgb = prime_color

    # ---- NMWA: staffing rule, not a labor % ----
    nm = totals_by_cafe.get("NMWA")
    if nm:
        nm_bk = week_labor_buckets(labor, "NMWA")
        # The 1-person/day rule governs CORE floor staffing. Training and admin
        # shifts sit outside core and do not consume the floor slot, so they are
        # not counted here (e.g. a trainee shadowing the scheduled barista).
        staffed_days = [(d, labor["per_cafe_per_day"].get("NMWA", {}).get(d, {}))
                        for d in DAYS]
        staffed = [(d, b) for d, b in staffed_days if b.get("core_shifts", b.get("shifts", 0))]
        # Event-titled shifts (e.g. NMWA Nights) staff a billed event, not the
        # register, so they don't count against the floor rule either.
        def _floor(b): return b.get("core_shifts", b.get("shifts", 0)) - b.get("role_breakdown", {}).get("Event", {}).get("shifts", 0)
        over = [d for d, b in staffed if _floor(b) > NMWA_STAFF_RULE["people_per_day"]]
        nm_train = sum(labor["per_cafe_per_day"].get("NMWA", {}).get(d, {}).get("training_cents", 0) for d in DAYS)
        nm_tx = nm["tx"]
        p2 = doc.add_paragraph()
        set_font(p2.add_run("NMWA — staffing rule (no labor % target): "), size=9, bold=True)
        set_font(p2.add_run(
            f"{len(staffed)} of 7 days staffed, {NMWA_STAFF_RULE['people_per_day']} person/day rule; "
            f"{'no days over rule' if not over else str(len(over)) + ' day(s) over rule (' + ', '.join(over) + ')'} "
            f"\u2014 {NMWA_STAFF_RULE['note']}. Sales ramp: {money(nm['core'])} core on {nm_tx:,} transactions, "
            f"core labor {money(nm_bk['core'])}"
            + (f", plus {money(nm_train)} training (outside the staffing rule)" if nm_train else "")
            + ". Watch the transactions/day slope, not labor %."
        ), size=9, color=RGBColor(0x44,0x44,0x44))

    # Outside-core table hidden for now (Ryan, Nov 8 2026). Set SHOW_OUTSIDE_CORE = True to restore.
    if SHOW_OUTSIDE_CORE:
        # ---- Outside core: Admin + Training ----
        doc.add_paragraph()
        ph = doc.add_paragraph()
        set_font(ph.add_run("Outside core (tracked, not targeted)"), size=10, bold=True)
        hdr2 = ["Cafe", "Admin $", "Admin shifts", "Training $", "Training % of Core", "vs plan"]
        t2 = add_table(doc, rows=len(CAFES)+1, cols=len(hdr2), headers=hdr2)
        for i, cafe in enumerate(CAFES, start=1):
            bk = week_labor_buckets(labor, cafe)
            core_v = totals_by_cafe[cafe]["core"]
            tr_pct = bk["training"]/core_v if core_v else 0
            adm_shifts = sum(labor["per_cafe_per_day"].get(cafe, {}).get(d, {}).get("admin_shifts", 0) for d in DAYS)
            if cafe not in CORE_LABOR_TARGETS:
                # NMWA and any other staffing-rule cafe: no % bands apply.
                tr_tag, tr_col = "n/a (staffing rule)", RGBColor(0x88,0x88,0x88)
            elif not tr_pct:
                tr_tag, tr_col = "—", RGBColor(0x88,0x88,0x88)
            elif tr_pct <= TRAINING_PLAN_PCT:
                tr_tag, tr_col = "\u25cf at/under plan", RGBColor(0x2E,0x7D,0x32)
            elif tr_pct <= TRAINING_AMBER_PCT:
                tr_tag, tr_col = "\u25cf amber", RGBColor(0xB7,0x79,0x0C)
            else:
                tr_tag, tr_col = "\u25cf over amber", RGBColor(0xC6,0x28,0x28)
            show_pct = (pct(tr_pct) if (tr_pct and cafe in CORE_LABOR_TARGETS) else "—")
            vals = [cafe, money(bk["admin"]), str(adm_shifts) if adm_shifts else "—",
                    money(bk["training"]), show_pct, tr_tag]
            for j, val in enumerate(vals):
                c = t2.rows[i].cells[j]; c.text = val
                for pp in c.paragraphs:
                    if j >= 1: pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    for r in pp.runs: r.font.size = Pt(9)
            for pp in t2.rows[i].cells[5].paragraphs:
                for r in pp.runs: r.font.color.rgb = tr_col; r.bold = True
        pn = doc.add_paragraph()
        set_font(pn.add_run(
            f"Admin = Square job title \u201cManager\u201d; no target or cap for now (Ryan, Aug 2026) \u2014 tracked for "
            f"visibility. Training = Square \u201cTraining\u201d role, plan {pct(TRAINING_PLAN_PCT)} of core sales, "
            f"amber to {pct(TRAINING_AMBER_PCT)}."
        ), size=8, color=RGBColor(0x66,0x66,0x66))

    p = doc.add_paragraph()
    set_font(p.add_run(
        "ALL THREE GREEN, a first for this report. Downtown 33.3%, Studio 27.1% and 1899 Penn 37.8%, and Prime Cost is Healthy at all three. "
        "Open-cafe core labor 31.7%, $22 under target in total (was $914 over). Studio was closed Mon and Tue for a Studio Theatre event."
    ), size=8, color=RGBColor(0x66,0x66,0x66))

    # ==== Headlines (Income Statement read; Open Cafes basis) ================
    doc.add_heading("Headlines — Income Statement Read (Open Cafes)", level=1)
    o = open_totals  # Open Cafes aggregate — matches Monthly Report totals column
    cogs_pct_o = o["cogs"]/o["core"] if o["core"] else 0
    lab_pct_o  = o["core_labor"]/o["core"] if o["core"] else 0
    prime_pct_o = (o["core_labor"]+o["cogs"])/o["core"] if o["core"] else 0
    tip_pct_o  = o["tips"]/o["core"] if o["core"] else 0

    bullets = [
        f"Core Sales (Revenue): {money(o['core'])} across {o['tx']:,} transactions. Catering layered on top: {money(o['catering'])} (reported separately, not in the ratios below).",
        f"COGS (applied): {money(o['cogs'])} → {pct(cogs_pct_o)} of Core Sales, using March 2026 category margins.",
        f"Gross Profit: {money(o['core'] - o['cogs'])} → {pct((o['core']-o['cogs'])/o['core'] if o['core'] else 0)} of Core Sales.",
        f"Core Labor: {money(o['core_labor'])} → {pct(lab_pct_o)} of Core Sales.",
        f"Prime Cost: {money(o['core_labor'] + o['cogs'])} → {pct(prime_pct_o)} of Core Sales. (Target 55–60%.)",
        f"Tips: {money(o['tips'])} ({pct(tip_pct_o)} of Core) — proxy for service quality.",
    ]
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        set_font(p.add_run(b), size=10)

    # ==== Wins & Watch-outs =================================================
    # TODO (future): auto-generate from totals_by_cafe + per-day data using rules
    # (best avg-ticket day, tip% YoY direction, prime-% bands, labor-vs-target gap, etc.).
    doc.add_heading("Wins & Watch-outs", level=1)
    wins = [
        "ALL THREE CAFES GREEN for the first time: Downtown 33.3%, Studio 27.1% and 1899 Penn 37.8%. Prime Cost is Healthy at all three (58.7%, 53.1% and 62.4%).",
        "DOWNTOWN GREEN at 33.3% on $10,793 of core sales, $564 over plan and up $2,742 on last week. Tue 11/17 did $3,226 at 23% labor and Wed 11/18 ran 26%.",
        "1899 PENN GREEN at 37.8% on $4,955 of core sales, $705 over plan. Tue 11/17 did $1,305.",
        "STUDIO at 27.1%, under its target, on five open days. Sat 11/21 ran 21% and Sun 11/15 24%.",
        "FILTER COFFEE: Studio is at 28% beyond Cruise Control, its second week near 30% (it was about 20% through October).",
    ]
    for w in wins:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run("WIN — "); set_font(r, size=10, bold=True, color=RGBColor(0x2E,0x7D,0x32))
        set_font(p.add_run(w), size=10)
    watches = [
        "STUDIO CLOSED Mon 11/16 and Tue 11/17 for a Studio Theatre event: $52 rang across the two days against a $2,068 plan, so Studio's week is $1,622 under plan. Wed to Sat ran $372 to $479 a day over plan except Saturday.",
        "DOWNTOWN FRIDAY: Fri 11/20 did $824 against a $1,261 plan at 61% labor, and Thu 11/19 ran 43%. Downtown also shows a 5-hour shift on Sat 11/21 ($105, Jay Singleton) when the cafe is closed; it is counted in Downtown's labor.",
        "1899 PENN FRIDAY ran 56% ($357 of labor on $640).",
        "NMWA EVENT Wed 11/18 (evening): four people on for about five hours, $510 of labor against $161 of register sales. Event revenue should come in by invoice; it is not in this week's numbers.",
        "THANKSGIVING WEEK NEXT: Downtown and 1899 Penn are closed Thu and Fri, and Studio is closed Thursday. Schedule to the week targets in section 6.",
        "TIMECARDS TO CHECK: (1) Christina Jie, Studio, Wed 11/18 never clocked out (Square closed it at 24h; capped at 7.2h). It is her third missed clock-out in three weeks. (2) Opal Payne, 1899 Penn, Tue 11/17 shows 6:53 am to 6:11 pm (11.3h) at a cafe that closes at 3 pm; likely a late clock-out, not capped, about $70 over a normal shift. (3) Jay Singleton, Downtown, Sat 11/21, 12:49 to 5:46 pm on a closed day ($105): event or wrong location? (4) NMWA had 31 sales on Thu 11/19 but no timecard.",
    ]
    for w in watches:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run("WATCH — "); set_font(r, size=10, bold=True, color=RGBColor(0xC6,0x28,0x28))
        set_font(p.add_run(w), size=10)

    # ==== 1. Cafe Financials — Income-Statement style (no cell borders) ====
    doc.add_heading("1. Cafe Financials — Week Summary", level=1)
    p = doc.add_paragraph()
    set_font(p.add_run(
        "All %s are of Core Sales (ex-catering/invoice payments). Catering reported separately."
    ), size=9, color=RGBColor(0x44,0x44,0x44))

    cafe_cols = CAFES + ["Open Cafes"]
    col_hdrs = ["($, except as noted)"] + cafe_cols

    # Each entry: (kind, label, fn_or_None, fill_or_None, indent)
    # kinds: 'section' (bold label, no values), 'detail', 'total', 'pct' (italic dim)
    rows_spec = [
        ("section", "Revenue", None, None, 0),
        ("detail",  "Transactions (count)",    lambda w: f"{w['tx']:,}" if w["tx"] else "—", None, 1),
        ("detail",  "Coffee — Beverage",       lambda w: money(w["cbev"]),   None, 1),
        ("detail",  "Coffee — Retail (bean)",  lambda w: money(w["retail"]), None, 1),
        ("detail",  "Food",                    lambda w: money(w["food"]),   None, 1),
        ("detail",  "Other",                   lambda w: money(w["other"]),  None, 1),
        ("total",   "Total Core Sales",        lambda w: money(w["core"]),   FIN_TOTAL_REVENUE, 0),
        ("detail",  "Avg Ticket (Core)",       lambda w: f"${(w['core']/w['tx'])/100:,.2f}" if w["tx"] else "—", None, 1),
        ("detail",  "Catering (reported separately)", lambda w: money(w["catering"]), None, 1),
        ("blank",   "", None, None, 0),
        ("section", "Cost of Goods Sold", None, None, 0),
        ("detail",  "Coffee COGS",             lambda w: money(w["coffee_cogs"]), None, 1),
        ("detail",  "Food COGS",               lambda w: money(w["food_cogs"]),   None, 1),
        ("total",   "Total COGS",              lambda w: money(w["cogs"]),        FIN_TOTAL_COGS, 0),
        ("pct",     "COGS % of Core Sales",    lambda w: pct(w["cogs"]/w["core"]) if w["core"] else "—", None, 1),
        ("total",   "Gross Profit (Core − COGS)", lambda w: money(w["core"] - w["cogs"]), None, 0),
        ("pct",     "Gross Profit % of Core Sales", lambda w: pct((w["core"]-w["cogs"])/w["core"]) if w["core"] else "—", None, 1),
        ("blank",   "", None, None, 0),
        ("section", "Labor — CORE (schedulable floor labor; targeted)", None, None, 0),
        ("detail",  "Core Labor $",            lambda w: money(w["core_labor"]), None, 1),
        ("pct",     "Core Labor % of Core Sales", lambda w: pct(w["core_labor"]/w["core"]) if w["core"] else "—", None, 1),
        ("detail",  "  memo: Admin (Manager)", lambda w: money(w["admin_labor"]), None, 1),
        ("detail",  "  memo: Training",        lambda w: money(w["training_labor"]), None, 1),
        ("total",   "Total Labor (all buckets)", lambda w: money(w["foh"]), FIN_TOTAL_LABOR, 0),
        ("blank",   "", None, None, 0),
        ("section", "Prime Cost", None, None, 0),
        ("total",   "Prime Cost (Core Labor + COGS)", lambda w: money(w["core_labor"] + w["cogs"]), FIN_TOTAL_PRIME, 0),
        ("pct",     "Prime Cost % of Core Sales", lambda w: pct((w["core_labor"]+w["cogs"])/w["core"]) if w["core"] else "—", None, 1),
        ("total",   "Prime Margin (Core − Prime)", lambda w: money(w["core"] - w["core_labor"] - w["cogs"]), FIN_TOTAL_PRIME, 0),
        ("pct",     "Prime Margin % of Core Sales", lambda w: pct((w["core"]-w["core_labor"]-w["cogs"])/w["core"]) if w["core"] else "—", None, 1),
        ("blank",   "", None, None, 0),
        ("section", "Tips (tracked separately, not in Prime)", None, None, 0),
        ("detail",  "Tips",                    lambda w: money(w["tips"]), None, 1),
        ("pct",     "Tip % of Core Sales",     lambda w: pct(w["tips"]/w["core"]) if w["core"] else "—", None, 1),
    ]

    t = fin_table(doc, n_rows=len(rows_spec)+1, n_cols=len(col_hdrs), col_headers=col_hdrs,
                  label_col_width_in=2.6, val_col_width_in=0.85)
    for i, (kind, label, fn, fill, indent) in enumerate(rows_spec, start=1):
        if kind == "blank":
            # leave row empty (visual gap)
            continue
        if kind == "section":
            fin_write_row(t, i, label, [None]*len(cafe_cols), kind="section", indent=indent)
            continue
        # data rows (detail / pct / total)
        values = []
        for c in CAFES:
            w = totals_by_cafe[c]
            if w["core"] == 0 and w["tx"] == 0 and c == "NMWA":
                values.append("—")
            else:
                values.append(fn(w))
        values.append(fn(open_totals))
        # Render with appropriate kind
        render_kind = kind if kind in ("total","detail") else "detail"
        fin_write_row(t, i, label, values, kind=render_kind, indent=indent, fill=fill)
        # Tone down the % rows (italicize, dim color)
        if kind == "pct":
            for c in t.rows[i].cells:
                for pp in c.paragraphs:
                    for rr in pp.runs:
                        rr.italic = True
                        rr.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
                        rr.font.size = Pt(9)

    # ==== 2. Daily Budget vs Actual (time-series; financial-statement style) ====
    doc.add_heading("2. Daily Budget vs Actual — by Cafe (Core Sales basis)", level=1)
    _p2 = doc.add_paragraph()
    set_font(_p2.add_run(
        "Labor rows are CORE labor only. Labor Plan = that day\u2019s projected sales \u00d7 the cafe\u2019s core "
        "target %. Federal-holiday Mondays show no plan and no variance at Downtown and 1899 Penn \u2014 "
        "those days are closed by rule."
    ), size=9, color=RGBColor(0x44,0x44,0x44))
    for cafe in CAFES:
        doc.add_heading(cafe, level=2)
        col_hdrs = [""] + DAY_LABELS + ["Week"]
        proj_net = [PROJECTIONS[cafe][d]["sales"] for d in DAYS]
        act_net = []
        for d in DAYS:
            info = day_info(sales, cafe, d)
            if info:
                cat_day = info["by_category"].get("catering", 0)
                act_net.append((info["net_sales_cents"] - cat_day) / 100)
            else:
                act_net.append(0)
        # Labor plan is DERIVED from the finalized core target (the workbook's
        # Labor Proj column is stale — Ryan, Aug 2026), and actuals are CORE
        # labor only so plan and actual sit on the same basis.
        _band = CORE_LABOR_TARGETS.get(cafe)
        proj_lab = [(PROJECTIONS[cafe][d]["sales"] * _band["target"]) if _band else 0.0 for d in DAYS]
        act_lab = [day_labor_buckets(labor, cafe, d)[0]/100 for d in DAYS]

        # Federal-holiday Mondays: DT and 1899 are closed. $0 is expected, so
        # blank the plan and suppress variance rather than showing a big miss.
        holiday_idx = [i for i, d in enumerate(DAYS) if is_holiday_closed(cafe, d)]
        for i in holiday_idx:
            proj_net[i] = 0.0
            proj_lab[i] = 0.0

        def fmt_dollar(v):
            if v == 0 or v is None: return "—"
            return f"${v:,.0f}" if v >= 0 else f"(${abs(v):,.0f})"

        # Per-day Labor % (Labor Actual / Core Sales Actual) — shows daily efficiency.
        labor_pct_per_day = []
        for _i, (s_, l_) in enumerate(zip(act_net, act_lab)):
            labor_pct_per_day.append(None if _i in holiday_idx else ((l_/s_) if s_ > 0 else None))
        labor_pct_week = (sum(act_lab)/sum(act_net)) if sum(act_net) > 0 else None
        # Per-cafe target for color-coding (from projections workbook)
        target_pct = (_band["target"] if _band else 0.0)

        # Sales variance: + is good (more sales). Labor variance: - is good (under-spend).
        rows_data = [
            ("section",  "Core Sales", None, None, None),
            ("detail",   "Plan",        proj_net, sum(proj_net), None),
            ("total",    "Actual",      act_net,  sum(act_net),  None),
            ("variance", "Δ",           [0 if i in holiday_idx else a-p for i,(a,p) in enumerate(zip(act_net, proj_net))], sum(act_net)-sum(proj_net), "neg_bad"),
            ("blank",    "", None, None, None),
            ("section",  "Core Labor (ex-Admin, ex-Training)", None, None, None),
            ("detail",   "Plan",        proj_lab, sum(proj_lab), None),
            ("total",    "Actual",      act_lab,  sum(act_lab),  None),
            ("variance", "Δ",           [0 if i in holiday_idx else a-p for i,(a,p) in enumerate(zip(act_lab, proj_lab))], sum(act_lab)-sum(proj_lab), "neg_good"),
            ("labor_pct","Core Labor %", labor_pct_per_day, labor_pct_week, target_pct),
        ]
        # Use slightly slimmer day cols for daily layouts
        t = fin_table(doc, n_rows=len(rows_data)+1, n_cols=len(col_hdrs), col_headers=col_hdrs,
                      label_col_width_in=1.6, val_col_width_in=0.7)

        for ri, row in enumerate(rows_data, start=1):
            kind = row[0]
            if kind == "blank":
                continue
            label = row[1]
            if kind == "section":
                fin_write_row(t, ri, label, [None]*(len(col_hdrs)-1), kind="section")
                continue
            vals, tot = row[2], row[3]
            if kind == "labor_pct":
                # Format pct values, color vs target (≤target green, ≤target+3pp amber, >target+3pp red)
                target = row[4]
                def _fmt_pct(v):
                    return f"{v*100:.1f}%" if v is not None and v > 0 else "—"
                display_vals = [_fmt_pct(v) for v in vals] + [_fmt_pct(tot)]
                fin_write_row(t, ri, label, display_vals, kind="detail", indent=1)
                for cidx, v in enumerate(list(vals) + [tot], start=1):
                    if v is None or target is None or target == 0 or v == 0:
                        color = RGBColor(0x88,0x88,0x88)
                    elif v <= target:
                        color = RGBColor(0x2E,0x7D,0x32)  # green: at/under target
                    elif v <= target + 0.03:
                        color = RGBColor(0xB7,0x79,0x0C)  # amber: within 3pp over
                    else:
                        color = RGBColor(0xC6,0x28,0x28)  # red: >3pp over
                    c = t.rows[ri].cells[cidx]
                    for pp in c.paragraphs:
                        for rr in pp.runs:
                            rr.font.color.rgb = color
                            rr.font.size = Pt(9)
                            rr.italic = True
                continue
            display_vals = [fmt_dollar(v) for v in vals] + [fmt_dollar(tot)]
            if kind == "variance":
                neg_is_good = (row[4] == "neg_good")
                fin_write_row(t, ri, label, display_vals, kind="detail", indent=1)
                for cidx, v in enumerate(list(vals) + [tot], start=1):
                    if v is None or v == 0:
                        color = RGBColor(0x88,0x88,0x88)
                    else:
                        good = (v < 0) if neg_is_good else (v > 0)
                        color = RGBColor(0x2E,0x7D,0x32) if good else RGBColor(0xC6,0x28,0x28)
                    c = t.rows[ri].cells[cidx]
                    for pp in c.paragraphs:
                        for rr in pp.runs:
                            rr.font.color.rgb = color
                            rr.font.size = Pt(9)
            elif kind == "total":
                fin_write_row(t, ri, label, display_vals, kind="total", indent=1, fill="EEF6E8")
            else:
                fin_write_row(t, ri, label, display_vals, kind="detail", indent=1)

    # ==== 3. Weekly Sales by Category — financial-statement style ===========
    doc.add_heading("3. Weekly Sales by Category", level=1)
    col_hdrs = [""] + cafe_cols
    cat_rows = [
        ("detail", "Coffee — Beverage",      "cbev",    None),
        ("detail", "Coffee — Retail (bean)", "retail",  None),
        ("detail", "Food",                   "food",    None),
        ("detail", "Other",                  "other",   None),
        ("total",  "Total Core Sales",       "core",    FIN_TOTAL_REVENUE),
        ("blank",  "", None, None),
        ("detail", "Catering (separate line)", "catering", None),
    ]
    t = fin_table(doc, n_rows=len(cat_rows)+1, n_cols=len(col_hdrs), col_headers=col_hdrs,
                  label_col_width_in=2.4, val_col_width_in=0.95)
    for i, (kind, label, key, fill) in enumerate(cat_rows, start=1):
        if kind == "blank":
            continue
        values = [money(totals_by_cafe[c][key]) for c in CAFES] + [money(open_totals[key])]
        fin_write_row(t, i, label, values, kind=kind, indent=(0 if kind == "total" else 1), fill=fill)

    # ==== 4. Ops Metrics — flipped =========================================
    doc.add_heading("4. Ops Metrics", level=1)
    p = doc.add_paragraph()
    set_font(p.add_run(
        "Per-cafe indicators that don't sum cleanly across locations — read each cafe on its own."
    ), size=9, color=RGBColor(0x44,0x44,0x44))

    # Build per-cafe scalar ops metrics
    ops = {}
    for cafe in CAFES:
        w = totals_by_cafe[cafe]
        open_days = sum(1 for d in DAYS if day_info(sales, cafe, d))
        avg_tx = w["tx"]/open_days if open_days else 0
        avg_ticket = w["core"]/w["tx"] if w["tx"] else 0  # Core-basis avg ticket
        tip_pct = w["tips"]/w["core"] if w["core"] else 0
        fulf = {}
        for d in DAYS:
            info = day_info(sales, cafe, d)
            if info:
                for k, v in info.get("by_fulfillment", {}).items():
                    fulf[k] = fulf.get(k, 0) + v
        total_f = sum(fulf.values())
        mix = ", ".join(f"{k}: {v/total_f*100:.0f}%" for k, v in sorted(fulf.items(), key=lambda x:-x[1])) if total_f else "—"
        ops[cafe] = {"days": open_days, "avg_tx": avg_tx, "avg_ticket": avg_ticket, "tip_pct": tip_pct, "mix": mix}

    ops_rows = [
        ("Open days this week",        lambda o: f"{o['days']}" if o['days'] else "—"),
        ("Avg Daily Transactions",     lambda o: f"{o['avg_tx']:.0f}" if o['avg_tx'] else "—"),
        ("Avg Ticket (Core)",          lambda o: f"${o['avg_ticket']/100:,.2f}" if o['avg_ticket'] else "—"),
        ("Tip % of Core Sales",        lambda o: pct(o['tip_pct']) if o['tip_pct'] else "—"),
        ("Dine-in / Takeout Mix",      lambda o: o['mix']),
    ]
    col_hdrs = ["Metric"] + CAFES  # no Open Cafes col here — these are per-cafe indicators
    t = add_table(doc, rows=len(ops_rows)+1, cols=len(col_hdrs), headers=col_hdrs)
    for i, (label, fn) in enumerate(ops_rows, start=1):
        values = [label] + [fn(ops[c]) for c in CAFES]
        write_row(t, i, values)

    # ==== 5. Next 4 Weeks — scheduling targets ==============================
    # Ryan committed this block to the team (Aug 2026): ALLs schedule directly
    # against these numbers and never compute a target themselves.
    # Target core labor $ = projected sales x that cafe's CORE target %.
    # Labor $ is NOT read from the projections workbook — those columns are
    # stale (Ryan, Aug 2026) — it is derived here from the finalized targets.
    doc.add_heading("5. Next 4 Weeks — Scheduling Targets", level=1)
    p = doc.add_paragraph()
    set_font(p.add_run(
        "Schedule against these numbers directly \u2014 no calculation needed. For each of the next four "
        "Sun\u2013Sat weeks: projected sales, the core labor dollars that hit target, the top of the "
        "operating range, and the hard MAX. Schedule to TARGET; the range is your working room; the MAX "
        "is the most you may schedule without Ryan\u2019s prior approval. Core labor excludes Admin (\u201cManager\u201d) "
        "and Training shifts \u2014 budget those on top. Rule of thumb: one 8-hour shift per $500\u2013600 of "
        "projected sales; balance across the WEEK, not each day."
    ), size=10, color=RGBColor(0x44,0x44,0x44))

    def _fmt_week_hdr(ws_iso, we_iso):
        ws = date.fromisoformat(ws_iso); we = date.fromisoformat(we_iso)
        end_s = we.strftime('%-d') if ws.month == we.month else we.strftime('%b %-d')
        return f"{ws.strftime('%b %-d')}\u2013{end_s}"

    week_hdrs = [_fmt_week_hdr(w["week_start"], w["week_end"]) for w in FORWARD_WEEKS]

    def _cell(cell, text, *, bold=False, right=True, color=None, size=9):
        cell.text = text
        for pg in cell.paragraphs:
            if right: pg.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for r in pg.runs:
                r.font.size = Pt(size); r.bold = bold
                if color: r.font.color.rgb = color

    # Which forward weeks contain a federal-holiday Monday closure?
    def _holiday_days(cafe, w):
        ws = date.fromisoformat(w["week_start"])
        out = []
        for i in range(7):
            d = (ws + timedelta(days=i)).isoformat()
            if is_holiday_closed(cafe, d):
                out.append(d)
        return out

    for cafe in CAFES:
        band = CORE_LABOR_TARGETS.get(cafe)
        if band:
            header_label = (f"{cafe}  \u2014 core range {pct(band['low'])}\u2013{pct(band['high'])} "
                            f"\u00b7 max {pct(band['max'])}")
        else:
            header_label = f"{cafe}  \u2014 staffing rule: {NMWA_STAFF_RULE['people_per_day']} person/day "\
                           f"({NMWA_STAFF_RULE['note']})"
        h = doc.add_paragraph()
        set_font(h.add_run(header_label), size=11, bold=True)

        sales = [w["cafes"][cafe]["sales"] for w in FORWARD_WEEKS]
        sales_total = sum(sales)

        if band:
            tgt_lab  = [v * band["low"] for v in sales]
            hi_lab   = [v * band["high"] for v in sales]
            max_lab  = [v * band["max"] for v in sales]
            rows_n = 5
        else:
            tgt_lab = hi_lab = max_lab = None
            rows_n = 2

        hdr = ["Metric"] + week_hdrs + ["4-Wk Total"]
        t = add_table(doc, rows=rows_n, cols=len(hdr), headers=hdr)

        _cell(t.rows[1].cells[0], "Projected Sales", right=False)
        for j, v in enumerate(sales):
            _cell(t.rows[1].cells[j+1], f"${v:,.0f}" if v else "\u2014")
        c = t.rows[1].cells[-1]; _cell(c, f"${sales_total:,.0f}", bold=True); set_cell_shading(c, "FFF2CC")

        if band:
            _cell(t.rows[2].cells[0], "TARGET Core Labor $", right=False, bold=True)
            for j, v in enumerate(tgt_lab):
                _cell(t.rows[2].cells[j+1], f"${v:,.0f}" if v else "\u2014", bold=True,
                      color=RGBColor(0x2E,0x7D,0x32))
            c = t.rows[2].cells[-1]
            _cell(c, f"${sum(tgt_lab):,.0f}", bold=True, color=RGBColor(0x2E,0x7D,0x32))
            set_cell_shading(c, "FFF2CC")

            _cell(t.rows[3].cells[0], "Top of range $", right=False)
            for j, v in enumerate(hi_lab):
                _cell(t.rows[3].cells[j+1], f"${v:,.0f}" if v else "\u2014",
                      color=RGBColor(0xB7,0x79,0x0C))
            c = t.rows[3].cells[-1]
            _cell(c, f"${sum(hi_lab):,.0f}", color=RGBColor(0xB7,0x79,0x0C))
            set_cell_shading(c, "FFF2CC")

            _cell(t.rows[4].cells[0], "MAX $ (ask Ryan above)", right=False)
            for j, v in enumerate(max_lab):
                _cell(t.rows[4].cells[j+1], f"${v:,.0f}" if v else "\u2014",
                      color=RGBColor(0xC6,0x28,0x28))
            c = t.rows[4].cells[-1]
            _cell(c, f"${sum(max_lab):,.0f}", color=RGBColor(0xC6,0x28,0x28))
            set_cell_shading(c, "FFF2CC")

        # Holiday-closure callout for this cafe, if any
        hol = []
        for w in FORWARD_WEEKS:
            hol.extend(_holiday_days(cafe, w))
        if hol:
            hp = doc.add_paragraph()
            nice = ", ".join(date.fromisoformat(d).strftime("%a %b %-d") for d in hol)
            set_font(hp.add_run(f"Closed (federal holiday): {nice} \u2014 zero staffing; "
                                f"the projection above already excludes these days."),
                     size=8, color=RGBColor(0xC6,0x28,0x28))
        doc.add_paragraph()

    note = doc.add_paragraph()
    set_font(note.add_run(
        f"Source: {PROJECTIONS_SOURCE} (Sales Proj column, aggregated Sun\u2013Sat). Target, range and max "
        f"dollars are computed here from the finalized core targets \u2014 the workbook\u2019s Labor Proj column "
        f"is not used. NMWA has no labor % target: staff one person per day and treat a second as an ask. "
        f"Federal-holiday Mondays: Downtown and 1899 Penn are closed, so $0 sales on those days is expected."
    ), size=8, color=RGBColor(0x66,0x66,0x66))

    # ==== 5b. Timecard Exceptions — fix in Square before payroll ============
    # Ryan, Sep 2026: surface everything that needed an adjustment so the
    # timesheets can be locked down and the labor numbers stand on their own.
    ex_capped   = labor.get("capped_shifts", []) or []
    ex_zero     = labor.get("skipped_zero_rate", []) or []
    ex_short    = labor.get("short_punches", []) or []
    ex_unmapped = labor.get("unmapped_location", []) or []
    ex_override = labor.get("bucket_overrides_applied", []) or []
    # A cafe that rang real walk-in sales with no timecard at all means someone
    # worked without clocking in (or clocked in at the wrong site). The $50
    # online eGift-card case is filtered out by requiring >= 5 transactions.
    ex_nostaff = []
    _sales_all = json.load(open(SALES_JSON))   # `sales` is reused as a list in section 5
    for _c in CAFES:
        for _d in DAYS:
            _si = _sales_all["per_cafe_per_day"].get(_c, {}).get(_d)
            if not _si: continue
            _core = _si["net_sales_cents"] - _si.get("by_category", {}).get("catering", 0)
            _lab = labor["per_cafe_per_day"].get(_c, {}).get(_d, {}).get("shifts", 0)
            if _si.get("orders", 0) >= 5 and _core > 0 and not _lab:
                ex_nostaff.append({"cafe": _c, "date": _d, "tx": _si["orders"], "sales": _core/100})
    ex_total = len(ex_capped) + len(ex_zero) + len(ex_short) + len(ex_unmapped) + len(ex_nostaff)

    doc.add_heading("5b. Timecard Exceptions — fix in Square before payroll", level=1)
    p = doc.add_paragraph()
    set_font(p.add_run(
        "Every timecard this week that needed an adjustment to keep the labor numbers honest. "
        "The report\u2019s figures already reflect these corrections \u2014 but the underlying punches are still "
        "wrong in Square, so payroll will pay the raw values unless they are fixed. Correct these, then the "
        "numbers above and payroll agree."
    ), size=9, color=RGBColor(0x44,0x44,0x44))

    if ex_total == 0 and not ex_override and not labor.get("location_overrides_applied"):
        pc = doc.add_paragraph()
        set_font(pc.add_run("\u2713  No timecard exceptions this week \u2014 nothing to fix."),
                 size=10, bold=True, color=RGBColor(0x2E,0x7D,0x32))
    else:
        hdrx = ["Issue", "Who", "Cafe", "Date", "Role", "What Square shows", "What the report used", "$ impact"]
        rowsx = []
        for c in ex_capped:
            rowsx.append(("Missed clock-out", c.get("who","?"), c["cafe"], c["date"], c.get("role",""),
                          f"{c['original_hours']:.2f}h", f"{c['capped_to_hours']:.2f}h",
                          f"-${c['cost_avoided']:,.2f}", RGBColor(0xC6,0x28,0x28)))
        for z in ex_zero:
            rowsx.append(("No wage rate set", z.get("who","?"), z["cafe"], z["date"], z.get("role",""),
                          f"{z['hours']:.2f}h @ $0/hr", "excluded", "understated",
                          RGBColor(0xC6,0x28,0x28)))
        for u in ex_unmapped:
            if u.get("kind") == "cafe_role_at_non_cafe":
                rowsx.append(("Cafe role at non-cafe site", u.get("who","?"), u.get("location_name","?"), u["date"],
                              u.get("role",""), f"{u['hours']:.2f}h", "excluded from all cafes",
                              f"${u['cost']:,.2f} — check location", RGBColor(0xB7,0x79,0x0C)))
            else:
                rowsx.append(("Unknown location", u.get("who","?"), f"loc {u['location_id']}", u["date"], u.get("role",""),
                              f"{u['hours']:.2f}h", "excluded from all cafes", f"${u['cost']:,.2f} unassigned",
                              RGBColor(0xB7,0x79,0x0C)))
        for ns in ex_nostaff:
            rowsx.append(("Sales, no timecard", "\u2014", ns["cafe"], ns["date"], "\u2014",
                          f"{ns['tx']} sales / ${ns['sales']:,.0f}", "0 labor booked", "labor missing",
                          RGBColor(0xC6,0x28,0x28)))
        for sp in ex_short:
            rowsx.append(("Mis-punch (too short)", sp.get("who","?"), sp["cafe"], sp["date"], sp.get("role",""),
                          f"{sp['minutes']:.0f} min", "kept (no cost)", "~$0",
                          RGBColor(0xB7,0x79,0x0C)))
        for lo in (labor.get("location_overrides_applied", []) or []):
            rowsx.append(("Wrong site (moved)", lo.get("who","?"), lo["to_cafe"], lo["date"], lo.get("role",""),
                          f"clocked at {lo['from_cafe']}", f"moved to {lo['to_cafe']}", f"{lo['hours']:.2f}h moved",
                          RGBColor(0x1F,0x4E,0x79)))
        for o in ex_override:
            rowsx.append(("Reclassified", o.get("who","?"), o["cafe"], o["date"], o.get("role",""),
                          f"{o['from_bucket']} (job title)", f"{o['to_bucket']}", f"${o['cost']:,.2f} moved",
                          RGBColor(0x1F,0x4E,0x79)))

        tx_ = add_table(doc, rows=len(rowsx)+1, cols=len(hdrx), headers=hdrx)
        for i, r in enumerate(rowsx, start=1):
            for j, val in enumerate(r[:-1]):
                c = tx_.rows[i].cells[j]; c.text = str(val)
                for pp in c.paragraphs:
                    if j >= 5: pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    for rr in pp.runs: rr.font.size = Pt(8)
            for pp in tx_.rows[i].cells[0].paragraphs:
                for rr in pp.runs: rr.font.color.rgb = r[-1]; rr.bold = True

        pf = doc.add_paragraph()
        set_font(pf.add_run(
            f"{ex_total} punch(es) need correcting in Square Teams"
            + (f"; {len(ex_override)} shift(s) reclassified by policy (no Square change needed if the job "
               f"title is updated)." if ex_override else ".")
        ), size=9, bold=True)
        ph = doc.add_paragraph()
        set_font(ph.add_run(
            "How each is handled: a missed clock-out (Square auto-closes at ~24h) is capped to the median "
            "shift for that cafe and role. A punch with no wage rate contributes $0, so it is excluded and "
            "that cafe\u2019s labor is understated until the rate is set. An unknown location is excluded from "
            "every cafe, so those hours disappear from the numbers entirely. Sub-5-minute punches are left in "
            "\u2014 they cost nothing but inflate shift counts."
        ), size=8, color=RGBColor(0x66,0x66,0x66))

    doc.add_paragraph()

    # ==== 6. Historical Average (from transactions store) ==================
    doc.add_heading("6. Historical Context — Monthly Averages vs This Week", level=1)
    p = doc.add_paragraph()
    set_font(p.add_run(
        "Pulled from the payment-level transactions store (Square Transactions export, Jan–Mar 2026 backfill). "
        "Gives a quick 'is this week normal?' read for avg ticket, avg daily transactions, and tip % "
        "(tips / net sales, gross basis — approximates the Core tip % shown above)."
    ), size=9, color=RGBColor(0x44,0x44,0x44))

    # Month ranges we compare against
    MONTHS = [
        ("Jan 2026", "2026-01-01", "2026-01-31"),
        ("Feb 2026", "2026-02-01", "2026-02-28"),
        ("Mar 2026", "2026-03-01", "2026-03-31"),
    ]

    def hist_avg_ticket(cafe, s, e):
        c = txn_store.avg_ticket_cents(cafe, s, e)
        return f"${c/100:,.2f}" if c else "—"
    def hist_avg_daily(cafe, s, e):
        v = txn_store.avg_daily_transactions(cafe, s, e)
        return f"{v:,.0f}" if v else "—"
    def hist_tip_pct(cafe, s, e):
        n, net, tips, _ = txn_store.counts_and_totals(cafe, s, e)
        return pct(tips/net) if net else "—"

    # Build compact per-open-cafe tables: metrics × (Jan, Feb, Mar, This Wk)
    _wk_short = " ".join(WEEK_FILE_TAG.split()[:2])  # "Apr 12-18" (drop year)
    col_hdrs = ["Metric"] + [m[0] for m in MONTHS] + [f"This Wk ({_wk_short})"]

    for cafe in OPEN_CAFES:
        h = doc.add_paragraph()
        set_font(h.add_run(cafe), size=11, bold=True)

        # This-week values (from weekly JSON, Ops section semantics)
        w = totals_by_cafe[cafe]
        # Avg ticket = Core Sales $ ÷ transactions; show cents (money() rounds to $0 above $1).
        this_wk_avg_ticket = f"${(w['core']/w['tx'])/100:,.2f}" if w["tx"] else "—"
        this_wk_avg_daily  = f"{ops[cafe]['avg_tx']:.0f}" if ops[cafe]['avg_tx'] else "—"
        this_wk_tip_pct    = pct(w["tips"]/w["core"]) if w["core"] else "—"

        rows_data = [
            ("Avg Ticket",
             [hist_avg_ticket(cafe, s, e) for _, s, e in MONTHS],
             this_wk_avg_ticket),
            ("Avg Daily Transactions",
             [hist_avg_daily(cafe, s, e) for _, s, e in MONTHS],
             this_wk_avg_daily),
            ("Tip % (tips ÷ net)",
             [hist_tip_pct(cafe, s, e) for _, s, e in MONTHS],
             this_wk_tip_pct),
        ]

        t = add_table(doc, rows=len(rows_data)+1, cols=len(col_hdrs), headers=col_hdrs)
        for i, (label, month_vals, week_val) in enumerate(rows_data, start=1):
            values = [label] + list(month_vals) + [week_val]
            write_row(t, i, values)
            # Shade the This Week col
            c = t.rows[i].cells[-1]; bold_cell(c); set_cell_shading(c, "FFF2CC")

    p = doc.add_paragraph()
    set_font(p.add_run(
        "Store coverage: Jan–Mar 2026. NMWA excluded (no register transactions in the backfilled period). "
        "Add April data by dropping the Square Transactions export for April into "
        "`_pipeline/shared/transactions/2026-04.csv`."
    ), size=8, color=RGBColor(0x66,0x66,0x66))

    # ==== 7. Caveats ========================================================
    doc.add_heading("7. Caveats & Notes", level=1)
    caveats = [
        "DATA: the Sunday automatic pull did not save this week, so sales were pulled straight from Square on Nov 22 (3,152 orders; payments built from each order's tenders; 3,111 transaction rows; 22 cafe-days). Loyalty visits were not pulled this week. Labor was pulled live from Square on Nov 22 (61 timecards, all closed).",
        "LABOR RULES: Square's new admin job title is counted as admin, like Manager. One Commissary shift punched at Studio (Sat 11/21, 8.2 hours, $158) is commissary work and is left out of Studio's labor.",
        "FILTER COFFEE now comes from Square's item sales report, which counts every cup on a ticket. Earlier weeks are restated on the same basis; last week was Downtown 23%, Studio 29% and 1899 Penn 34% (posted as 19%, 31% and 36%).",
        "Targets: core wages divided by core walk-in sales. Downtown 32% to 36% (max 38%), Studio 30% to 34% (max 36%), 1899 Penn 35% to 40% (max 42%). NMWA: staffing rule, no %. Green = inside range or better; amber = over range, under max; red = over max (needs Ryan's pre-approval).",
        "Core excludes Admin (Manager or admin job title: Downtown $319, Studio $144, 1899 Penn $79) and Training; salaried pay is not in Square timecards.",
        "Prime Cost uses CORE labor, the same basis as the labor scorecard. 1899 Penn's Prime band runs to 65% healthy, in line with its labor range.",
        "PROJECTIONS: plans and section 6 read 2026 Projections v1.11 (Nov 10), reviewed day by day through year end for Thanksgiving, Christmas and New Year.",
        "Section 6 dollars = projected sales times the cafe's low/high/max %. The workbook's Labor Proj column is not used.",
    ]
    for c in caveats:
        p = doc.add_paragraph(style="List Bullet")
        set_font(p.add_run(c), size=9, color=RGBColor(0x44,0x44,0x44))

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run(f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M ET')} · Source: Slipstream Weekly Data.xlsx"), size=8, color=RGBColor(0x88,0x88,0x88))

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print(f"Saved → {OUT}")

if __name__ == "__main__":
    main()
