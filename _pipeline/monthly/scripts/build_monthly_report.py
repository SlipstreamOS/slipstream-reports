#!/usr/bin/env python3
"""Build the Slipstream Monthly Business Review (v0.1 — July 2026).

Source: Consolidated Model v14 — a REBUILT model (vs v7 used through May).
Key structural changes handled here:
  - Formula-driven; 2026 columns carry no cached values, so the workbook must be
    recalculated (LibreOffice) before reading. Jul = col Y on every IS tab.
  - Ops metrics now split weekday / weekend (avg daily transactions + avg ticket),
    surfaced on a dedicated 'Ops Metrics' tab. Section 2 restructured to match.
  - Consolidated tab now has a correct 'TOTAL NET INCOME (all entities)' rollup
    (r106). The three v6/v7 formula bugs are FIXED — Appendix B no longer needs
    a manual restatement.
  - 'Total Core Location Sales' (r35) INCLUDES catering; Core ex-catering is
    derived as r35 - r30.

July status:
  - NMWA is OPEN and trading (31.4 weekday txns/day) but deep in ramp. Per Ryan,
    it stays OUT of the 'Open Cafes' blend and gets its own ramp section (1f).
  - Section 3 (Items) and Section 5 (Loyalty) PENDING July inputs.
  - Section 4 (Discounts) still on hold per Ryan.
"""
import json
import os
from copy import deepcopy

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PIPELINE_DIR = os.path.dirname(SCRIPT_DIR)
WORKSPACE = os.path.dirname(os.path.dirname(PIPELINE_DIR))
DATA_DIR = os.path.join(PIPELINE_DIR, "data")

MONTH_LABEL = "July 2026"
VERSION = "v0.1"
OUT = os.path.join(WORKSPACE, f"Slipstream Monthly Report - {MONTH_LABEL} {VERSION}.docx")

ITEM_SLICES_JSON = os.path.join(DATA_DIR, "jul26_item_slices.json")

# ---------- investor-report formatting helpers ----------

FONT_NAME = "Calibri"
COLOR_BLACK = RGBColor(0x00, 0x00, 0x00)
COLOR_MUTED = RGBColor(0x60, 0x60, 0x60)
COLOR_ACCENT = RGBColor(0x1F, 0x3A, 0x5F)  # Slipstream-ish navy for section headers
SHADE_HEADER = "E8ECEF"   # light gray-blue for table header rows
SHADE_KEYROW = "F3F5F7"   # subtle shade for key metric rows
SHADE_TOTAL = "D9E2EC"    # slightly darker for bold totals


def _set_cell_border(cell, **kwargs):
    """Set individual cell borders. kwargs keys: top/left/bottom/right (values = dict with sz/color/val).

    OOXML requires tcBorders children in canonical order: top, left, bottom, right,
    insideH, insideV, tl2br, tr2bl. We rebuild the element fresh each call, merging
    the new edges with any already-set edges, so order is always correct.
    """
    CANONICAL = ("top", "left", "bottom", "right", "insideH", "insideV")
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:tcBorders"))
    current = {}
    if existing is not None:
        for edge in CANONICAL:
            el = existing.find(qn(f"w:{edge}"))
            if el is not None:
                current[edge] = el
        tc_pr.remove(existing)
    # apply updates
    for edge, spec in kwargs.items():
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), spec.get("val", "single"))
        el.set(qn("w:sz"), str(spec.get("sz", 4)))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), spec.get("color", "000000"))
        current[edge] = el
    # write back in canonical order, inserted before any shd/tcMar/vAlign siblings
    new_borders = OxmlElement("w:tcBorders")
    for edge in CANONICAL:
        if edge in current:
            new_borders.append(current[edge])
    _insert_in_tcpr(tc_pr, new_borders, "w:tcBorders")


_TCPR_ORDER = (
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
    "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
    "w:tcFitText", "w:vAlign", "w:hideMark",
)


def _insert_in_tcpr(tc_pr, new_el, tag):
    """Insert new_el into tc_pr in canonical OOXML position, replacing any existing."""
    # remove any existing
    for existing in tc_pr.findall(qn(tag)):
        tc_pr.remove(existing)
    # find the first later-ordered sibling and insert before it
    order = list(_TCPR_ORDER)
    try:
        target_idx = order.index(tag)
    except ValueError:
        tc_pr.append(new_el)
        return
    later = order[target_idx + 1:]
    for child in tc_pr:
        child_tag_local = child.tag.split("}")[-1]
        child_tag = f"w:{child_tag_local}"
        if child_tag in later:
            child.addprevious(new_el)
            return
    tc_pr.append(new_el)


def _set_cell_shading(cell, fill_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    _insert_in_tcpr(tc_pr, shd, "w:shd")


def _clear_all_borders(cell):
    """Reset cell borders to nil on all edges, preserving canonical child order."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        tc_borders.append(el)
    _insert_in_tcpr(tc_pr, tc_borders, "w:tcBorders")


def _set_run_font(run, size=10, bold=False, italic=False, color=None, font=FONT_NAME):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    # east-asia font fallback so Word doesn't swap us to Calibri Light on some systems
    r = run._element
    rPr = r.find(qn("w:rPr")) or r.makeelement(qn("w:rPr"), {})
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)


def _set_para_spacing(para, before=0, after=2, line=1.15):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def add_heading(doc, text, level=1):
    sizes = {1: 18, 2: 13, 3: 11}
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=sizes.get(level, 11), bold=True,
                  color=COLOR_ACCENT if level <= 2 else COLOR_BLACK)
    _set_para_spacing(p, before=14 if level == 1 else 10, after=4, line=1.1)
    return p


def add_para(doc, text, *, italic=False, size=10, color=None, bold=False, before=0, after=4):
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=size, bold=bold, italic=italic, color=color)
    _set_para_spacing(p, before=before, after=after, line=1.2)
    return p


def add_bullet(doc, text, bold_lead=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_lead:
        r1 = p.add_run(bold_lead)
        _set_run_font(r1, size=10, bold=True)
        r2 = p.add_run(" " + text)
        _set_run_font(r2, size=10)
    else:
        r = p.add_run(text)
        _set_run_font(r, size=10)
    _set_para_spacing(p, before=0, after=2, line=1.2)
    return p


# ---------- investor-style table builder ----------

def build_investor_table(doc, headers, rows, *, col_aligns=None, col_widths=None,
                         key_rows=None, total_rows=None,
                         header_shaded=True, label_col=0):
    """Render a clean financial table.

    - No per-cell grid.
    - Thin rule under header; thin rule above the last (total) row; a light
      bottom rule below the final row.
    - Rows listed in `key_rows` get a very subtle fill; rows in `total_rows`
      get a heavier fill + bold and a rule on top.
    - Numeric columns are right-aligned.
    """
    key_rows = set(key_rows or [])
    total_rows = set(total_rows or [])
    ncols = len(headers)
    if col_aligns is None:
        col_aligns = ["left"] + ["right"] * (ncols - 1)

    tbl = doc.add_table(rows=1 + len(rows), cols=ncols)
    tbl.autofit = False
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in tbl.rows:
                row.cells[i].width = w

    # header
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
        }[col_aligns[j]]
        r = p.add_run(str(h))
        _set_run_font(r, size=10, bold=True, color=COLOR_ACCENT)
        _set_para_spacing(p, before=1, after=1, line=1.1)
        _clear_all_borders(cell)
        if header_shaded:
            _set_cell_shading(cell, SHADE_HEADER)
        # thin top + thicker bottom rule on header row
        _set_cell_border(cell, top={"sz": 12, "color": "1F3A5F", "val": "single"},
                                bottom={"sz": 8, "color": "1F3A5F", "val": "single"})
        cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    # data rows
    last_idx = len(rows) - 1
    for i, row in enumerate(rows):
        is_key = i in key_rows
        is_total = i in total_rows
        for j, val in enumerate(row):
            cell = tbl.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
            }[col_aligns[j]]
            r = p.add_run(str(val))
            _set_run_font(r, size=10, bold=(is_total or (is_key and j == label_col)))
            _set_para_spacing(p, before=1, after=1, line=1.15)
            _clear_all_borders(cell)
            if is_total:
                _set_cell_shading(cell, SHADE_TOTAL)
                _set_cell_border(cell, top={"sz": 6, "color": "1F3A5F", "val": "single"})
            elif is_key:
                _set_cell_shading(cell, SHADE_KEYROW)
            if i == last_idx:
                _set_cell_border(cell, bottom={"sz": 6, "color": "1F3A5F", "val": "single"})
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # little spacer paragraph after the table
    spacer = doc.add_paragraph()
    _set_para_spacing(spacer, before=0, after=4, line=1.0)
    return tbl


def add_caveat(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=9, italic=True, color=COLOR_MUTED)
    _set_para_spacing(p, before=0, after=10, line=1.15)
    return p


def add_callout(doc, title, body_text, *, accent=None, fill="F3F5F7",
                body_size=10, title_size=10):
    """Render a single-cell navy-bordered callout (used for 'The Formula' and
    target-band notes). Matches the handout's callout-box visual feel.
    """
    accent = accent or COLOR_ACCENT
    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = True
    cell = tbl.rows[0].cells[0]
    # clear default cell text
    cell.text = ""
    # title paragraph
    p_title = cell.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_title.add_run(title)
    _set_run_font(r, size=title_size, bold=True, color=accent)
    _set_para_spacing(p_title, before=2, after=2, line=1.15)
    # body paragraph
    p_body = cell.add_paragraph()
    p_body.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_body.add_run(body_text)
    _set_run_font(r, size=body_size, bold=False)
    _set_para_spacing(p_body, before=0, after=2, line=1.2)
    # borders + fill — navy thin rule on all edges
    _clear_all_borders(cell)
    for edge in ("top", "left", "bottom", "right"):
        _set_cell_border(cell, **{edge: {"sz": 8, "color": "1F3A5F", "val": "single"}})
    _set_cell_shading(cell, fill)
    # spacer after
    spacer = doc.add_paragraph()
    _set_para_spacing(spacer, before=0, after=6, line=1.0)
    return tbl


# ---------- document assembly ----------

def build():
    doc = Document()

    # page + default font
    section = doc.sections[0]
    section.left_margin = Inches(0.9)
    section.right_margin = Inches(0.9)
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)

    # tweak normal style
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = Pt(10)

    # ---------- masthead ----------
    p = doc.add_paragraph()
    r = p.add_run("SLIPSTREAM COFFEE")
    _set_run_font(r, size=11, bold=True, color=COLOR_MUTED)
    _set_para_spacing(p, before=0, after=0)

    p = doc.add_paragraph()
    r = p.add_run("Monthly Business Review")
    _set_run_font(r, size=22, bold=True, color=COLOR_ACCENT)
    _set_para_spacing(p, before=0, after=0)

    p = doc.add_paragraph()
    r = p.add_run(MONTH_LABEL)
    _set_run_font(r, size=14, bold=True)
    _set_para_spacing(p, before=0, after=2)

    add_para(doc,
             f"Prepared August 12, 2026  |  {VERSION} — First July pass, built from the REBUILT "
             "Consolidated Model v14. Section 2 restructured to the model's new weekday / weekend "
             "split. New Section 1f covers NMWA's opening ramp (kept out of the Open Cafes blend "
             "per direction). Appendix B no longer needs a manual restatement — v14 fixes the three "
             "consolidation bugs carried in v6/v7. Sections 3 (Items) and 5 (Loyalty) are PENDING "
             "July inputs; Section 4 (Discounts) remains on hold.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------- Executive summary ----------
    add_heading(doc, "Executive Summary", level=1)

    add_callout(doc,
                "The Formula",
                "Control COGS  +  Protect Labor Hours  +  Grow Sales  =  Better Business + Better Pay. "
                "Every headline below maps back to one of these three levers.",
                body_size=10, title_size=11)

    add_para(doc,
             "July core cafe sales (ex-catering, three established cafes): $101,892, down from "
             "$113,259 in June (-10.0% MoM). Blended core prime cost rose to 68.5% — above the "
             "65% 'real problem' threshold for the first time since February, with both levers "
             "moving the wrong way at once: COGS 33.8% → 34.6% and labor 30.1% → 33.9%. Total "
             "Net Income across all entities was $11,142 (June $14,068). Two bright spots: SRP "
             "posted its best month of the year (+$4,178) and NMWA opened for real trading — "
             "31 weekday transactions/day and $2,907 of core sales in its first full month.",
             after=8)

    add_bullet(doc, "68.5% blended, vs 63.9% June and a 56.7% low in March. Downtown 69.5%, 1899 76.8%, "
                    "Studio 64.2% — every cafe is now at or above the 65% problem line except Studio, "
                    "which is in the warning band. This is the headline for the ALL meeting.",
               bold_lead="Prime cost crossed into the problem zone.")
    add_bullet(doc, "Studio labor jumped 29.2% → 35.7% of core on a 15.9% sales decline ($53,083 → $44,656) — "
                    "the single largest swing in the business this month. 1899 also worsened (32.8% → 39.7%). "
                    "Downtown was the only cafe to hold, easing slightly from 30.0% to 29.2%. Blended "
                    "30.1% → 33.9%.",
               bold_lead="Labor didn't scale down with July volume.")
    add_bullet(doc, "Studio -15.9% ($53,083 → $44,656), Downtown -4.3% ($41,137 → $39,356), 1899 -6.1% "
                    "($19,039 → $17,881). Studio's drop is the outlier and drives most of the company "
                    "decline. Weekday transactions there fell 158.7 → 133.4/day and weekend 258.4 → 196.6/day.",
               bold_lead="Sales fell at all three established cafes.")
    add_bullet(doc, "First full month of real trading: $2,907 core sales ex-catering, 31.4 weekday and 15.3 "
                    "weekend transactions/day, $6.49 / $8.32 tickets. Still deeply ramp-negative "
                    "(-$1,904 operating) as expected. Detail in Section 1f; excluded from blended ratios.",
               bold_lead="NMWA is open and trading.")
    add_bullet(doc, "+$4,178 in July, best month of 2026 and a second positive month in three. Concentrate "
                    "sales $5,745 and COGS down sharply to $2,200 (from $4,353 June). The roasting "
                    "path-to-profitability story is finally showing signal.",
               bold_lead="SRP had its strongest month of the year.")
    add_bullet(doc, "Coffee COGS 29.0% blended (target 22–27%, above); food COGS 38.9% (target 28–35%, "
                    "above); prime cost 68.5% (>65% problem). All three outside band — July is a "
                    "'control COGS + protect labor' month on both levers at once.",
               bold_lead="All three target bands breached.")

    # ========== 1. Financials ==========
    add_heading(doc, "1. Cafe Financials — July 2026", level=1)
    add_para(doc,
             "Source: Consolidated Model v14. Core rows reflect cafe operations net of catering. "
             "Prime Cost = COGS + Labor; Prime Margin = Sales − Prime Cost. Labor is not allocated "
             "between core and catering in the current chart of accounts; all labor sits under Core. "
             "'Open Cafes' = Downtown + Studio + 1899. NMWA opened during Q2 and is still in ramp — "
             "it is shown as its own column and in Section 1f, but excluded from blended ratios so "
             "the trend stays comparable.",
             italic=True, size=9, color=COLOR_MUTED, after=10)

    # ----- 1a. Core P&L -----
    add_heading(doc, "1a. Core Cafe P&L — July 2026", level=2)
    add_para(doc,
             "Sales shown by revenue category; COGS by category (Coffee, Food) to make category-level "
             "cost efficiency visible (coffee COGS % flags waste; food COGS % flags pricing/spec drift).",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    headers_1a = ["Metric (Core, ex-catering)", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes"]
    rows_1a = [
        ["Sales", "", "", "", "", ""],                                                                       # 0
        ["  Coffee — Beverage",         "$26,779", "$28,931", "$15,560", "$2,763", "$71,270"],               # 1
        ["  Coffee — Retail (bean)",    "$696",    "$1,125",  "$820",    "$59",    "$2,642"],                # 2
        ["  Food",                       "$13,006", "$16,128", "$2,353",  "$711",   "$31,487"],              # 3
        ["  Less: Discounts / Voids / Refunds", "($1,125)", "($1,529)", "($853)", "($626)", "($3,507)"],     # 4
        ["Total Core Sales (Revenue)",   "$39,356", "$44,656", "$17,881", "$2,907", "$101,892"],             # 5 TOTAL
        ["  Coffee % of Core Sales",    "69.8%",   "67.3%",   "91.6%",   "97.1%",  "72.5%"],                 # 6
        ["  Food % of Core Sales",      "33.0%",   "36.1%",   "13.2%",   "24.5%",  "30.9%"],                 # 7
        ["Cost of Goods Sold",           "", "", "", "", ""],                                                # 8
        ["  Coffee COGS",                "$9,645",  "$6,945",  "$4,858",  "$1,957", "$21,448"],              # 9
        ["  Coffee COGS % of Coffee",   "35.1%",   "23.1%",   "29.7%",   "69.3%",  "29.0%"],                 # 10
        ["  Food COGS",                  "$4,688",  "$5,766",  "$1,788",  "$1,076", "$12,241"],              # 11
        ["  Food COGS % of Food",       "36.0%",   "35.8%",   "76.0%",   "151.3%", "38.9%"],                 # 12
        ["  Merchandise / Retail Purchases", "$1,537", "$0",   "$0",      "$0",     "$1,537"],               # 13
        ["Total Core COGS",              "$15,869", "$12,711", "$6,646",  "$3,033", "$35,226"],              # 14 TOTAL
        ["  COGS % of Core Sales",      "40.3%",   "28.5%",   "37.2%",   "104.3%", "34.6%"],                 # 15
        ["Gross Profit (Revenue − COGS)","$23,487", "$31,944", "$11,235", "($126)", "$66,666"],              # 16 TOTAL
        ["  Gross Profit % of Core Sales","59.7%", "71.5%",   "62.8%",   "(4.3%)", "65.4%"],                 # 17 KEY
        ["Labor",                        "", "", "", "", ""],                                                # 18
        ["  Total Labor",                "$11,496", "$15,949", "$7,094",  "$2,547", "$34,539"],              # 19
        ["  Labor % of Core Sales",     "29.2%",   "35.7%",   "39.7%",   "87.6%",  "33.9%"],                 # 20
        ["Core Prime Cost",              "$27,365", "$28,660", "$13,740", "$5,580", "$69,765"],              # 21 TOTAL
        ["  Prime Cost % of Core Sales","69.5%",   "64.2%",   "76.8%",   "192.0%", "68.5%"],                 # 22
        ["Core Prime Margin",            "$11,991", "$15,996", "$4,141",  "($2,673)","$32,127"],             # 23 TOTAL
        ["  Prime Margin % of Core Sales","30.5%", "35.8%",   "23.2%",   "(92.0%)","31.5%"],                 # 24
    ]
    col_widths_1a = [Inches(2.55), Inches(1.0), Inches(1.0), Inches(1.0), Inches(0.85), Inches(1.1)]
    build_investor_table(doc, headers_1a, rows_1a,
                         col_widths=col_widths_1a,
                         key_rows=[0, 8, 17, 18],
                         total_rows=[5, 14, 16, 21, 23])
    add_caveat(doc,
               "NMWA ratios are shown for completeness but are not meaningful at current volume — it "
               "is in its first full month of trading (see 1f). It is excluded from the Open Cafes "
               "column. Downtown's $1,537 Merchandise / Retail purchase is a lumpy retail-bean "
               "inventory buy; net of it Downtown COGS would be 36.4% of core rather than 40.3%. "
               "Bar, Market, and Unclassified categories omitted — all cafes report zero for July.")

    add_callout(doc,
                "How to read this against the Cafe Finance 101 targets",
                "Coffee COGS target 22–27% of coffee revenue → Studio 23.1% (in band), 1899 29.7% "
                "(above), Downtown 35.1% (well above). Food COGS target 28–35% → Downtown 36.0% and "
                "Studio 35.8% (both just above), 1899 76.0% (far above, on a small $2,353 base). "
                "Prime Cost target 55–60% healthy / 60–65% warning / >65% problem → Studio 64.2% "
                "(warning), Downtown 69.5% and 1899 76.8% (both problem). Blended 68.5% is the first "
                "reading above the 65% problem line since February. Unlike April (a COGS-only spike) "
                "and May (a labor-only drift), July has both levers moving against us at once.",
                body_size=9, title_size=10)

    # ----- 1b. Catering -----
    add_heading(doc, "1b. Catering Sales — Jan → Jul 2026", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "All Cafes"],
        [
            ["Jan 2026", "$8,252",  "$63",    "$1,525", "$9,700",  "$19,540"],
            ["Feb 2026", "$1,560",  "$219",   "$1,035", "$0",      "$2,814"],
            ["Mar 2026", "$11,944", "$63",    "$419",   "$29,100", "$41,526"],
            ["Apr 2026", "$3,301",  "$28",    "$1,042", "$0",      "$4,371"],
            ["May 2026", "$3,033",  "$1,555", "$127",   "$335",    "$5,049"],
            ["Jun 2026", "$737",    "$34",    "$557",   "$0",      "$1,327"],
            ["Jul 2026", "$2,265",  "$42",    "$132",   "$1,508",  "$3,947"],
        ],
        col_widths=[Inches(1.1)] + [Inches(1.2)] * 5,
        total_rows=[6])
    add_caveat(doc,
               "July catering recovered to $3,947 from June's $1,327 low, led by Downtown ($2,265) "
               "and NMWA's first meaningful catering ($1,508). Catering COGS and margin are tracked "
               "in the model but not broken out here.")

    # ----- 1c. Prime Cost & Component Trends -----
    add_heading(doc, "1c. Prime Cost & Component Trends — Jan → Jul 2026 (ex-catering)", level=2)
    add_para(doc,
             "Prime Cost = COGS + Labor, all as % of Core Sales. Splitting the components shows which "
             "lever is moving. Through June the two moved in opposition and largely cancelled; in July "
             "both moved against us.",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    add_heading(doc, "Prime Cost % of Core Sales", level=3)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "Open Cafes Blended"],
        [
            ["Jan 2026", "70.2%", "71.5%", "112.0%", "75.5%"],
            ["Feb 2026", "59.8%", "64.2%", "92.6%",  "66.3%"],
            ["Mar 2026", "54.9%", "51.5%", "81.6%",  "56.7%"],
            ["Apr 2026", "62.8%", "53.4%", "81.3%",  "61.3%"],
            ["May 2026", "64.7%", "52.4%", "82.2%",  "60.9%"],
            ["Jun 2026", "61.1%", "61.3%", "77.1%",  "63.9%"],
            ["Jul 2026", "69.5%", "64.2%", "76.8%",  "68.5%"],
        ],
        col_widths=[Inches(1.15)] + [Inches(1.3)] * 3 + [Inches(1.7)],
        total_rows=[6])

    add_heading(doc, "COGS % of Core Sales", level=3)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "Open Cafes Blended"],
        [
            ["Jan 2026", "35.8%", "31.0%", "43.0%", "33.9%"],
            ["Feb 2026", "34.5%", "32.4%", "43.2%", "34.6%"],
            ["Mar 2026", "30.3%", "26.7%", "34.3%", "29.0%"],
            ["Apr 2026", "39.6%", "30.8%", "35.0%", "34.7%"],
            ["May 2026", "35.5%", "25.8%", "37.2%", "30.8%"],
            ["Jun 2026", "31.2%", "32.0%", "44.2%", "33.8%"],
            ["Jul 2026", "40.3%", "28.5%", "37.2%", "34.6%"],
        ],
        col_widths=[Inches(1.15)] + [Inches(1.3)] * 3 + [Inches(1.7)],
        total_rows=[6])

    add_heading(doc, "Labor % of Core Sales", level=3)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "Open Cafes Blended"],
        [
            ["Jan 2026", "34.4%", "40.5%", "69.0%", "41.6%"],
            ["Feb 2026", "25.3%", "31.8%", "49.4%", "31.6%"],
            ["Mar 2026", "24.6%", "24.9%", "47.3%", "27.7%"],
            ["Apr 2026", "23.2%", "22.6%", "46.3%", "26.7%"],
            ["May 2026", "29.2%", "26.7%", "45.0%", "30.1%"],
            ["Jun 2026", "30.0%", "29.2%", "32.8%", "30.1%"],
            ["Jul 2026", "29.2%", "35.7%", "39.7%", "33.9%"],
        ],
        col_widths=[Inches(1.15)] + [Inches(1.3)] * 3 + [Inches(1.7)],
        total_rows=[6])

    add_caveat(doc,
               "Target bands (Cafe Finance 101): Prime Cost 55–60% healthy / 60–65% warning / >65% "
               "problem. Reading the components for July: Studio is the mover — labor 29.2% → 35.7% "
               "on a 15.9% sales decline, while its COGS actually improved (32.0% → 28.5%). Downtown "
               "is the reverse — labor eased slightly but COGS jumped 31.2% → 40.3% (partly the "
               "$1,537 retail-bean buy). 1899's COGS improved 7 points but labor gave back most of "
               "June's gain.")

    # ----- 1d. Core Sales Trend -----
    add_heading(doc, "1d. Core Sales Trend — Jan → Jul 2026 (ex-catering)", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes"],
        [
            ["Jan 2026", "$24,669", "$41,813", "$8,324",  "($970)",  "$74,807"],
            ["Feb 2026", "$32,323", "$41,578", "$11,340", "$0",      "$85,241"],
            ["Mar 2026", "$39,027", "$55,009", "$14,100", "($2,910)","$108,136"],
            ["Apr 2026", "$43,482", "$56,232", "$19,291", "$68",     "$119,005"],
            ["May 2026", "$37,010", "$54,283", "$15,106", "$34",     "$106,399"],
            ["Jun 2026", "$41,137", "$53,083", "$19,039", "$20",     "$113,259"],
            ["Jul 2026", "$39,356", "$44,656", "$17,881", "$2,907",  "$101,892"],
        ],
        col_widths=[Inches(1.35)] + [Inches(1.15)] * 5,
        total_rows=[6])
    add_caveat(doc,
               "Open Cafes excludes NMWA. July's $101,892 is the lowest since February and -10.0% vs "
               "June. Studio accounts for $8,427 of the $11,367 company decline. NMWA's negative "
               "Jan/Mar figures are accounting offsets from the pre-opening period, not trading losses.")

    # ----- 1e. Actual vs Projected -----
    add_heading(doc, "1e. Actual vs. Projected — July 2026", level=2)
    add_para(doc,
             "Source: 2026 Projections v1.7 (FY26 budget). Compared to July actual Core Sales "
             "(ex-catering) and Total Labor. NMWA carries a $0 July projection — the budget assumed "
             "an August open, so it is running ahead of plan on timing.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Cafe", "Sales Actual", "Sales Proj", "Sales Δ ($)", "Sales Δ (%)", "Labor Actual", "Labor Proj", "Labor Δ ($)"],
        [
            ["Downtown",   "$39,356",  "$39,361",  "($5)",      "(0.0%)",  "$11,496", "$14,170", "($2,674)"],
            ["Studio",     "$44,656",  "$53,525",  "($8,869)",  "(16.6%)", "$15,949", "$19,269", "($3,320)"],
            ["1899 Penn",  "$17,881",  "$17,425",  "+$456",     "+2.6%",   "$7,094",  "$8,713",  "($1,619)"],
            ["NMWA",       "$2,907",   "$0",       "+$2,907",   "—",        "$2,547",  "$0",      "+$2,547"],
            ["Open Cafes", "$101,892", "$110,311", "($8,419)",  "(7.6%)",  "$34,539", "$42,152", "($7,613)"],
        ],
        col_aligns=["left"] + ["right"] * 7,
        col_widths=[Inches(1.2), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.85), Inches(0.95), Inches(0.95), Inches(0.95)],
        total_rows=[4])
    add_caveat(doc,
               "Open Cafes row excludes NMWA. First month of 2026 where open-cafe sales came in below "
               "plan (-7.6%), driven almost entirely by Studio (-16.6%). Downtown landed within $5 of "
               "plan. Labor remains under budget in dollars at every cafe — but that is a function of "
               "lower sales, not tighter scheduling: as a % of actual sales, labor is above target.")

    # ----- 1f. NMWA opening ramp (NEW) -----
    add_heading(doc, "1f. NMWA — Opening Ramp", level=2)
    add_para(doc,
             "NMWA began trading in Q2 and July is its first full month of normal operation. It is "
             "excluded from all blended ratios above so the Jan→Jul trend stays comparable. Ramp "
             "economics are expected to be negative; the purpose of this section is to track the "
             "slope, not to judge the ratios.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Metric", "Apr 2026", "May 2026", "Jun 2026", "Jul 2026"],
        [
            ["Core Sales (ex-catering)",     "$68",     "$34",     "$20",     "$2,907"],
            ["Catering & Events",            "$0",      "$335",    "$0",      "$1,508"],
            ["Weekday avg transactions/day", "0.4",     "2.2",     "13.0",    "31.4"],
            ["Weekend avg transactions/day", "0.2",     "0.8",     "19.5",    "15.3"],
            ["Weekday avg ticket",           "$4.96",   "$7.34",   "$0.50",   "$6.49"],
            ["Weekend avg ticket",           "$14.12",  "$6.56",   "$0.00",   "$8.32"],
            ["Total Labor",                  "$2,662",  "$755",    "$2,528",  "$2,547"],
            ["Net Operating Income",         "($4,573)","($1,527)","($2,911)","($1,904)"],
            ["Net Income (incl. mgmt fee)",  "($129)",  "$2,953",  "$1,604",  "$3,375"],
        ],
        col_aligns=["left", "right", "right", "right", "right"],
        col_widths=[Inches(2.5), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.2)],
        key_rows=[2, 3], total_rows=[8])
    add_caveat(doc,
               "Weekday transaction count has more than doubled each month since May (2.2 → 13.0 → "
               "31.4/day) — the ramp is real and accelerating. Operating loss narrowed from ($2,911) "
               "in June to ($1,904) in July. Net Income is positive because of the $5,841 monthly "
               "management fee credit booked on the NMWA entity; the operating line is the one to "
               "watch. June's $0.50 weekday ticket reflects opening-period comps and test transactions "
               "rather than normal trading.")

    # ========== 2. Operating Indicators ==========
    add_heading(doc, "2. Operating Indicators", level=1)
    add_para(doc,
             "Source: Consolidated Model v14 'Ops Metrics' tab. The rebuilt model now separates "
             "weekday and weekend performance, so this section is presented that way from July "
             "onward. Downtown and 1899 are weekday-only operations; Studio and NMWA trade both.",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    add_heading(doc, "2a. Weekday Average Ticket ($) — Jan → Jul 2026", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA"],
        [
            ["Jan 2026", "$8.44",  "$9.00", "$7.43", "—"],
            ["Feb 2026", "$8.75",  "$8.79", "$7.05", "—"],
            ["Mar 2026", "$8.56",  "$9.00", "$7.21", "—"],
            ["Apr 2026", "$11.52", "$9.14", "$8.00", "$4.96"],
            ["May 2026", "$14.11", "$9.37", "$7.41", "$7.34"],
            ["Jun 2026", "$8.65",  "$9.21", "$7.40", "$0.50"],
            ["Jul 2026", "$8.38",  "$8.87", "$7.25", "$6.49"],
        ],
        col_widths=[Inches(1.2)] + [Inches(1.4)] * 4,
        total_rows=[6])
    add_caveat(doc,
               "Downtown's April–May readings ($11.52 / $14.11) are the eat-in anomaly flagged in "
               "prior reports — a handful of unusually large dine-in tickets. June and July are back "
               "on trend at $8.65 and $8.38. Studio slipped to $8.87, its lowest of the year.")

    add_heading(doc, "2b. Weekday Average Daily Transactions — Jan → Jul 2026", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "Total/Day"],
        [
            ["Jan 2026", "133", "111", "53",  "—",  "297"],
            ["Feb 2026", "179", "115", "81",  "—",  "375"],
            ["Mar 2026", "203", "134", "91",  "—",  "428"],
            ["Apr 2026", "219", "157", "116", "0",  "492"],
            ["May 2026", "203", "159", "98",  "2",  "460"],
            ["Jun 2026", "218", "159", "125", "13", "515"],
            ["Jul 2026", "193", "133", "113", "31", "470"],
        ],
        col_widths=[Inches(1.15)] + [Inches(1.2)] * 5,
        total_rows=[6])

    add_heading(doc, "2c. Weekend Metrics — Studio & NMWA", level=2)
    add_para(doc,
             "Downtown and 1899 do not trade weekends; only Studio and NMWA are shown.",
             italic=True, size=9, color=COLOR_MUTED, after=4)
    build_investor_table(doc,
        ["Period", "Studio — Txns/Day", "Studio — Avg Ticket", "NMWA — Txns/Day", "NMWA — Avg Ticket"],
        [
            ["Jan 2026", "203", "$10.06", "—",  "—"],
            ["Feb 2026", "244", "$10.23", "—",  "—"],
            ["Mar 2026", "294", "$10.05", "—",  "—"],
            ["Apr 2026", "289", "$10.71", "0",  "$14.12"],
            ["May 2026", "209", "$10.66", "1",  "$6.56"],
            ["Jun 2026", "258", "$10.13", "20", "$0.00"],
            ["Jul 2026", "197", "$9.73",  "15", "$8.32"],
        ],
        col_widths=[Inches(1.1), Inches(1.4), Inches(1.4), Inches(1.35), Inches(1.35)],
        total_rows=[6])
    add_caveat(doc,
               "Studio's weekend is the soft spot in July: transactions fell from 258 to 197/day "
               "(-24%) and ticket from $10.13 to $9.73 — together roughly $8K of the month's sales "
               "decline. Weekend has historically been Studio's strongest daypart, so this is worth "
               "a targeted look before August.")

    # ========== 3. Item-Level Category Insights (PENDING) ==========
    add_heading(doc, "3. Item-Level Category Insights — July 2026", level=1)
    add_callout(doc,
                "Pending July inputs",
                "This section is built from the per-cafe Square Item Sales exports, which have not "
                "yet been provided for July. Once the Downtown / Studio / 1899 CSVs land (plus NMWA, "
                "now that it is trading), this section returns in full: 3a Coffee Roll-up top & "
                "bottom 10, 3b Food Roll-up top & bottom 10, 3c Pastry (all items >3 units), 3d "
                "Filter Coffee & Retail units by location with the Premium Hot Filter share trend. "
                "Everything else in this report is complete.",
                body_size=9, title_size=10)

    # ========== 4. Discount Report (on hold) ==========
    add_heading(doc, "4. Discount Report — July 2026", level=1)
    add_para(doc,
             "On hold per Ryan. The March 2026 baseline (~3.66% discount-to-gross blended, Downtown "
             "highest at 4.42%) remains the reference point. For context, discounts as booked in the "
             "P&L ran $3,507 across the three established cafes in July (3.3% of core sales) — "
             "Downtown $1,125, Studio $1,529, 1899 $853. A by-type breakdown (staff / promo / "
             "loyalty redemption) needs a Square Orders export with discount name populated.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ========== 5. Loyalty Report (PENDING) ==========
    add_heading(doc, "5. Loyalty Report — July 2026", level=1)
    add_callout(doc,
                "Pending July inputs",
                "This section is built from the per-cafe Square Loyalty screens, which have not yet "
                "been provided for July. Once they land, this section returns in full: 5a headline "
                "insights, 5b Loyalty Member Value in dollars (using Square's average-spend-per-"
                "customer metric), and 5c–5f per-cafe detail with first-time / repeat visits, "
                "average visit frequency, and visit-bucket distribution. For reference, June "
                "partial-month data showed Downtown 2.6 / Studio 2.6 / 1899 3.0 loyalty avg visits "
                "vs 1.2–1.4 non-loyalty.",
                body_size=9, title_size=10)

    # ========== 6. Notes ==========
    add_heading(doc, "6. Notes & Watch Items", level=1)
    add_heading(doc, "Data caveats", level=3)
    add_bullet(doc, "Financials come from the REBUILT Consolidated Model v14. The model is fully "
                    "formula-driven and ships without cached values for 2026, so it must be "
                    "recalculated before reading. All figures here are post-recalculation "
                    "(17,760 formulas, zero errors).")
    add_bullet(doc, "'Total Core Location Sales' in v14 includes catering. Core ex-catering — the basis "
                    "for every ratio in this report — is derived by subtracting the Catering & Events line.")
    add_bullet(doc, "NMWA is trading but in ramp. It is excluded from all blended ratios and shown "
                    "separately in Section 1f. Its Net Income is positive only because of the monthly "
                    "management fee credit booked on that entity; the operating line is negative.")
    add_bullet(doc, "Downtown July COGS includes a $1,537 Merchandise / Retail purchase — a lumpy "
                    "retail-bean inventory buy, not a recurring cost of sales. Net of it, Downtown "
                    "COGS is 36.4% of core rather than 40.3%.")
    add_bullet(doc, "Sections 3 (Items) and 5 (Loyalty) await July Square exports. Section 4 "
                    "(Discounts) remains on hold per Ryan.")
    add_bullet(doc, "Some Jan–Jun figures differ slightly from prior reports — v14 restated a handful "
                    "of line items during the rebuild. All trends in this report are computed from "
                    "v14 end-to-end, so they are internally consistent.")

    add_heading(doc, "Consolidated Model — the v6/v7 bugs are fixed", level=3)
    add_bullet(doc, "The three formula bugs flagged in the April and May reports (Total Sales missing "
                    "Commissary + SRP, SRP Net Income mismatch, Mgmt Net Income missing from the "
                    "consolidated rollup) are all resolved in v14. The Consolidated tab now carries a "
                    "correct 'TOTAL NET INCOME (all entities)' line that ties to the sum of the entity "
                    "tabs. Appendix B no longer needs a manual restatement — a real improvement in "
                    "the close process.",
               bold_lead="Resolved in the rebuild.")

    add_heading(doc, "Watch items for August", level=3)
    add_bullet(doc, "Weekend transactions fell 258 → 197/day (-24%) and weekend ticket $10.13 → $9.73. "
                    "Weekend is Studio's strongest daypart, so this is the highest-value diagnostic in "
                    "the business right now. Was it staffing, weather, a local calendar effect, or "
                    "genuine demand softening?",
               bold_lead="Studio weekend is the single biggest driver of the July decline.")
    add_bullet(doc, "Labor was under budget in dollars at all three established cafes but above target "
                    "as a % of sales. That is the signature of hours not flexing down with volume. "
                    "Studio 35.7% and 1899 39.7% are the two to address first.",
               bold_lead="Labor scheduling did not flex with the volume drop.")
    add_bullet(doc, "40.3% of core (36.4% net of the retail-bean buy) with coffee COGS at 35.1% — well "
                    "above the 22–27% band and a repeat of the April pattern. Two spikes in four months "
                    "at the same cafe suggests something structural rather than a timing artifact.",
               bold_lead="Downtown COGS spiked again.")
    add_bullet(doc, "76.0% of food sales on a small $2,353 base. The percentage is volatile at this "
                    "volume, but it has now been above 55% for seven straight months. Worth costing "
                    "the actual food menu at 1899 rather than tracking the ratio.",
               bold_lead="1899 food COGS remains structurally high.")
    add_bullet(doc, "Weekday transactions 13 → 31/day and operating loss narrowing ($2,911) → ($1,904). "
                    "Track whether August sustains the slope; the FY26 budget assumed an August start "
                    "at $2,555, which NMWA has already exceeded.",
               bold_lead="NMWA ramp is tracking well — keep watching the slope.")
    add_bullet(doc, "+$4,178, best month of 2026, with COGS down to $2,200 from $4,353. Worth confirming "
                    "whether the COGS improvement is a green-coffee timing effect or a durable change "
                    "before treating the trajectory as established.",
               bold_lead="SRP finally showing real profitability.")

    # ========== Appendices ==========
    doc.add_paragraph().add_run().add_break()
    p = doc.add_paragraph()
    r = p.add_run("LEADERSHIP TEAM ONLY — APPENDICES")
    _set_run_font(r, size=12, bold=True, color=COLOR_MUTED)
    _set_para_spacing(p, before=10, after=6)

    add_heading(doc, "A. Support Entities — July 2026", level=2)
    add_para(doc,
             "Source: Consolidated Model v14, per-entity Income Statement tabs.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Metric", "Commissary", "Central Admin", "Cafe Mgmt", "SRP (Roasting)", "Support Total"],
        [
            ["Revenue",      "$10,175",  "$0",        "$5,424", "$11,826",  "$27,424"],
            ["COGS",         "$6,340",   "$0",        "$0",     "$2,200",   "$8,540"],
            ["Total Labor",  "$4,552",   "$9,361",    "$0",     "$2,402",   "$16,315"],
            ["Expenses",     "$2,634",   "$2,002",    "$0",     "$3,045",   "$7,681"],
            ["Net Income",   "($3,351)", "($11,362)", "$5,424", "$4,178",   "($5,112)"],
        ],
        col_widths=[Inches(1.5), Inches(1.15), Inches(1.2), Inches(1.05), Inches(1.25), Inches(1.15)],
        total_rows=[4])

    add_heading(doc, "A1. Support Entities — Net Income Trend", level=3)
    build_investor_table(doc,
        ["Entity", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "MoM Δ"],
        [
            ["Commissary",     "($5,509)",  "($3,826)",  "($4,199)",  "($4,321)",  "($1,048)", "($3,351)",  "($2,303)"],
            ["Central Admin",  "($10,244)", "($13,794)", "($11,739)", "($10,214)", "($8,877)", "($11,362)", "($2,485)"],
            ["Cafe Mgmt",      "$5,234",    "$5,693",    "$6,017",    "$5,762",    "$5,980",   "$5,424",    "($556)"],
            ["SRP (Roasting)", "($2,543)",  "($1,697)",  "$3,123",    "($4,794)",  "$406",     "$4,178",    "+$3,772"],
        ],
        col_widths=[Inches(1.35), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.95)])
    add_caveat(doc,
               "SRP posted its strongest month of 2026 at +$4,178 — COGS fell to $2,200 from $4,353 "
               "in June while concentrate sales held at $5,745. Commissary gave back June's "
               "improvement, and Central Admin costs rose $2,485 on higher labor ($7,084 → $9,361).")

    add_heading(doc, "B. Consolidated P&L — July 2026", level=2)
    add_para(doc,
             "Source: Consolidated Model v14 'Consolidated' tab. Unlike April and May, no manual "
             "restatement is required — v14's 'TOTAL NET INCOME (all entities)' line correctly rolls "
             "up the four cafes plus Commissary, SRP, Cafe Mgmt and Central Admin. Verified: "
             "$16,253 cafes − $3,351 Commissary + $4,178 SRP + $5,424 Cafe Mgmt − $11,362 Central "
             "= $11,142.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Metric", "Mar", "Apr", "May", "Jun", "Jul", "MoM Δ"],
        [
            ["Total Core Location Sales", "$146,752", "$123,444", "$111,482", "$114,605", "$108,746", "(5.1%)"],
            ["Total COGS",                "$33,141",  "$43,246",  "$33,547",  "$38,486",  "$38,259",  "(0.6%)"],
            ["Gross Profit",              "$113,611", "$80,198",  "$77,935",  "$76,120",  "$70,486",  "(7.4%)"],
            ["Total Labor",               "$30,493",  "$34,388",  "$32,831",  "$36,632",  "$37,086",  "+1.2%"],
            ["Total Expenses",            "$23,377",  "$28,134",  "$18,284",  "$24,138",  "$21,213",  "(12.1%)"],
            ["Net Operating Income",      "$59,742",  "$17,675",  "$26,820",  "$15,350",  "$12,188",  "(20.6%)"],
            ["Cafes — Net Income",        "$60,452",  "$19,412",  "$29,615",  "$17,607",  "$16,253",  "(7.7%)"],
            ["TOTAL NET INCOME (all entities)", "$46,827", "$12,613", "$16,048", "$14,068", "$11,142", "(20.8%)"],
        ],
        col_widths=[Inches(2.2), Inches(1.0), Inches(1.0), Inches(1.0), Inches(1.0), Inches(1.0), Inches(0.95)],
        key_rows=[5, 6], total_rows=[7])
    add_caveat(doc,
               "Total Net Income of $11,142 is the third consecutive monthly decline (from May's "
               "$16,048), though it remains comfortably positive. The March figure reflects NMWA's "
               "$29,100 of pre-opening catering and is not a steady-state comparison. Gross profit "
               "fell 7.4% on the sales decline while labor rose 1.2% in absolute dollars — the "
               "squeeze visible in Section 1c, at the consolidated level.")

    add_heading(doc, "C. Methodology & Data Sources", level=2)
    add_heading(doc, "Data sources", level=3)
    add_bullet(doc, "Financials: Slipstream Consolidated Model 2026 v14.xlsx — per-entity Income "
                    "Statement tabs. July 2026 = column Y on every IS tab and on Consolidated; date "
                    "header on row 14. The workbook must be recalculated before reading (it ships "
                    "without cached 2026 values).")
    add_bullet(doc, "Operating indicators: v14 'Ops Metrics' tab — weekday / weekend average daily "
                    "transactions and average ticket, per cafe, per month.")
    add_bullet(doc, "Projections (Section 1e): 2026 Projections v1.7.xlsx — per-cafe monthly summary, "
                    "row 13 (July), column R (Sales Proj) and T (Labor Proj). Also requires "
                    "recalculation. Sales projections tie exactly to v14's own Projections tab.")
    add_bullet(doc, "Items (Section 3) and Loyalty (Section 5): Square exports — pending for July.")

    add_heading(doc, "Definitions", level=3)
    add_bullet(doc, "Core Sales = Total Core Location Sales − Catering & Events Sales.")
    add_bullet(doc, "Core COGS = Total COGS − Catering & Events COGS (includes merchandise purchases).")
    add_bullet(doc, "Core Prime Cost = Core COGS + Total Labor. Labor is not yet allocated between "
                    "core and catering.")
    add_bullet(doc, "Core Prime Margin = Core Sales − Core Prime Cost.")
    add_bullet(doc, "Open Cafes = Downtown + Studio + 1899 Penn. NMWA is excluded while in ramp.")

    add_heading(doc, "Refresh process", level=3)
    add_para(doc,
             f"Recalculate the Consolidated Model and Projections workbooks, then run "
             f"build_monthly_report.py. Output: Slipstream Monthly Report - [Month] [Year] {VERSION}.docx.",
             size=10, after=6)

    # patch settings.xml zoom element to include w:percent (python-docx quirk)
    _ensure_zoom_percent(doc)

    doc.save(OUT)
    print(f"Saved → {OUT}")


def _ensure_zoom_percent(doc):
    """python-docx emits <w:zoom /> without the required w:percent attribute.
    Word is lenient but strict validators reject it. Set a default percent.
    """
    settings_part = doc.settings.element
    zoom = settings_part.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")


if __name__ == "__main__":
    build()
