#!/usr/bin/env python3
"""Build the Slipstream Monthly Business Review (v1.0 — May 2026).

Built from Consolidated Model v7 (financials), 2026 Projections v1.2 (FY26 plan),
Square Item Sales exports (May), and Square Loyalty screenshots (May).

May story (vs April):
  - Open-cafe Core Sales $106,399, down $12,606 (-10.6%) vs April.
  - Coffee COGS recovered to 25.3% blended (was 32.3%) — the Downtown spike
    receded, Studio dropped to 18.4% (excellent), 1899 28.7% (close to band).
  - Labor crept up to 30.5% of Core (was 25.9%) — drove prime cost roughly flat
    at 61.3% (was 60.6%).
  - SRP turned negative again (-$4,794) after April's +$3,123 — worth dedicated
    review.
  - Consolidated Net Income (corrected): $15,750. Spreadsheet r294 still shows
    $7,909 — same v6 bugs in v7 (documented in Notes).
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

MONTH_LABEL = "May 2026"
VERSION = "v1.1"
OUT = os.path.join(WORKSPACE, f"Slipstream Monthly Report - {MONTH_LABEL} {VERSION}.docx")

ITEM_SLICES_JSON = os.path.join(DATA_DIR, "may26_item_slices.json")

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
             f"Prepared June 5, 2026  |  {VERSION} — Built from Consolidated Model v7. Carries "
             "forward all April v1.2 structure (Formula callout, Gross Profit row, target-band "
             "overlays, Core-only sales trend, Premium Hot Filter share, by-month visit buckets, "
             "corrected Consolidated NI). v1.1: 1899 May loyalty visit detail now populated.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ---------- Executive summary ----------
    add_heading(doc, "Executive Summary", level=1)

    # The Formula callout — opens the summary with the same frame the team heard at today's ALL Meeting
    add_callout(doc,
                "The Formula",
                "Control COGS  +  Protect Labor Hours  +  Grow Sales  =  Better Business + Better Pay. "
                "Every headline below maps back to one of these three levers.",
                body_size=10, title_size=11)

    add_para(doc,
             "May core cafe sales (ex-catering, open cafes): $106,399, down from $119,005 "
             "in April (-10.6% MoM — grow sales lever moved the wrong way). Catering held "
             "roughly flat at $5,049 (vs. April $4,371). Blended core prime cost held at "
             "61.3% — coffee COGS recovered strongly (32.3% → 25.3% blended, back in target "
             "band) but labor crept up (25.9% → 30.5%), offsetting the COGS win. Corrected "
             "consolidated net income across all entities was $15,750 (the v6 'Consolidated "
             "Net Income' bugs still present in v7 — spreadsheet shows $7,909). Per-cafe "
             "Net Income strong (Downtown $9,582, Studio $16,782) despite the sales decline "
             "— driven by lower expenses ($23,145 vs April's $34,575). SRP turned negative "
             "again (-$4,794 vs +$3,123 April) — a watch item.",
             after=8)

    add_bullet(doc, "Open-cafe Core down 10.6% MoM. Downtown -14.9% ($43,482 → $37,010), 1899 -21.7% "
                    "($19,291 → $15,106), Studio -3.5% ($56,232 → $54,283). Worth understanding the drivers — "
                    "seasonality? Specific event/closure days? May normally a softer month than April but the "
                    "magnitude here is notable.",
               bold_lead="Sales pulled back across all three open cafes.")
    add_bullet(doc, "Blended coffee COGS dropped 7 points to 25.3% — Studio 18.4% (well below target band), 1899 "
                    "28.7% (just above), Downtown 33.1% (still above target but down from 40.3%). Whatever drove "
                    "the April spike has eased.",
               bold_lead="Coffee COGS recovered — back in target zone.")
    add_bullet(doc, "Labor % rose at every cafe: Downtown 23.9% → 30.3%, Studio 22.6% → 26.6%, 1899 40.0% → 45.0%. "
                    "On lower sales, labor hours weren't fully scaled back — protect-labor-hours lever needs "
                    "attention in June.",
               bold_lead="Labor % drift — needs eyes on it.")
    add_bullet(doc, "+$5,023 March → +$3,123 April → -$4,794 May. Coffee sales held but expense / write-down? "
                    "Worth a dedicated review before June actuals.",
               bold_lead="SRP swung sharply negative.")
    add_bullet(doc, "Per-cafe Net Income up despite sales drop. Downtown NI $5,476 → $9,582, Studio $13,956 → "
                    "$16,782, NMWA -$129 → $2,953. Operating expenses dropped ~$11K consolidated, more than "
                    "offsetting the gross-profit hit from lower sales.",
               bold_lead="Expense discipline carried the bottom line.")
    add_bullet(doc, "Coffee COGS 25.3% blended (in 22–27% target band — good); food COGS 35.6% (just above target "
                    "28–35%); prime cost 61.3% (just inside warning band 60–65%). 1899 prime cost still at 82% "
                    "(problem zone).",
               bold_lead="Target bands: one back in, two on the edge.")

    # ========== 1. Financials ==========
    add_heading(doc, "1. Cafe Financials — May 2026", level=1)
    add_para(doc,
             "Source: Consolidated Model v7. Core rows reflect cafe operations (net of catering). Catering "
             "rows shown separately to prime margin. Prime Cost = COGS + Labor; Prime Margin = Sales − "
             "Prime Cost. Labor is not allocated between core and catering in the current chart of "
             "accounts; all labor is shown under Core.",
             italic=True, size=9, color=COLOR_MUTED, after=10)

    # ----- 1a. Core P&L (investor style) -----
    add_heading(doc, "1a. Core Cafe P&L — May 2026", level=2)
    add_para(doc,
             "Sales shown by revenue category; COGS by category (Coffee, Food) to make category-level cost "
             "efficiency visible (coffee COGS % flags waste; food COGS % flags pricing/spec drift).",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    headers_1a = ["Metric (Core, ex-catering)", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes"]
    rows_1a = [
        ["Sales", "", "", "", "", ""],                                                                          # 0 section
        ["  Coffee — Beverage",         "$24,982", "$34,971", "$12,900", "$678",    "$72,853"],                 # 1
        ["  Coffee — Retail (bean)",    "$828",    "$991",    "$288",    "$0",      "$2,107"],                  # 2
        ["  Food",                       "$12,335", "$20,012", "$2,292",  "$0",      "$34,640"],                # 3
        ["  Less: Discounts / Voids / Refunds", "($1,135)", "($1,690)", "($375)",  "($644)", "($3,200)"],       # 4
        ["Total Core Sales (Revenue)",   "$37,010", "$54,283", "$15,106", "($300)",  "$106,399"],               # 5 TOTAL
        ["  Coffee % of Core Sales",    "69.7%",   "66.2%",   "87.3%",   "—",       "70.5%"],                   # 6
        ["  Food % of Core Sales",      "33.3%",   "36.9%",   "15.2%",   "—",       "32.6%"],                   # 7
        ["Cost of Goods Sold",           "", "", "", "", ""],                                                   # 8 section
        ["  Coffee COGS",                "$8,541",  "$6,634",  "$3,783",  "$0",      "$18,958"],                # 9
        ["  Coffee COGS % of Coffee",   "33.1%",   "18.4%",   "28.7%",   "—",       "25.3%"],                   # 10
        ["  Food COGS",                  "$4,076",  "$6,772",  "$1,493",  "$0",      "$12,341"],                # 11
        ["  Food COGS % of Food",       "33.0%",   "33.8%",   "65.1%",   "—",       "35.6%"],                   # 12
        ["Total Core COGS",              "$13,189", "$13,978", "$5,620",  "$229",    "$32,787"],                # 13 TOTAL
        ["  COGS % of Core Sales",      "35.6%",   "25.8%",   "37.2%",   "—",       "30.8%"],                   # 14
        ["Gross Profit (Revenue − COGS)","$23,821", "$40,305", "$9,486",  "($529)",  "$73,612"],                # 15 TOTAL
        ["  Gross Profit % of Core Sales","64.4%", "74.2%",   "62.8%",   "—",       "69.2%"],                   # 16 KEY
        ["Labor",                        "", "", "", "", ""],                                                   # 17 section
        ["  Total Labor",                "$11,206", "$14,467", "$6,804",  "$755",    "$32,477"],                # 18
        ["  Labor % of Core Sales",     "30.3%",   "26.6%",   "45.0%",   "—",       "30.5%"],                   # 19
        ["Core Prime Cost",              "$24,395", "$28,445", "$12,424", "$984",    "$65,263"],                # 20 TOTAL
        ["  Prime Cost % of Core Sales","65.9%",   "52.4%",   "82.2%",   "—",       "61.3%"],                   # 21
        ["Core Prime Margin",            "$12,615", "$25,838", "$2,682",  "($1,284)","$41,136"],                # 22 TOTAL
        ["  Prime Margin % of Core Sales","34.1%", "47.6%",   "17.8%",   "—",       "38.7%"],                   # 23
    ]
    # section headers: rows 0, 8, 17. Totals (shaded+bold): 5, 13, 15, 20, 22. Key (shaded): 16.
    col_widths_1a = [Inches(2.55), Inches(1.0), Inches(1.0), Inches(1.0), Inches(0.85), Inches(1.1)]
    build_investor_table(doc, headers_1a, rows_1a,
                         col_widths=col_widths_1a,
                         key_rows=[0, 8, 16, 17],
                         total_rows=[5, 13, 15, 20, 22])
    add_caveat(doc,
               "NMWA pre-opening — Core operations not yet active; May Core ex-catering shows -$300 "
               "(catering offsetting limited Square activity). NMWA had a few test days only. Bar, "
               "Market, and Unclassified categories omitted because all cafes report zero in those "
               "lines for May 2026.")

    # Target-band callout — same Cafe Finance 101 vocabulary, calibrated to May actuals.
    add_callout(doc,
                "How to read this against the Cafe Finance 101 targets",
                "Coffee COGS target: 22–27% of coffee revenue → Studio 18.4% (excellent — below "
                "target band), 1899 28.7% (just above), Downtown 33.1% (still above but down from "
                "40.3% April). Food COGS target: 28–35% of food revenue → Downtown 33.0% and "
                "Studio 33.8% (both in band), 1899 65.1% (well above — got worse vs April's 57.7%). "
                "Prime Cost target: 55–60% healthy, 60–65% warning, >65% problem → Studio 52.4% "
                "healthy, Downtown 65.9% (just inside problem zone), 1899 82.2% (problem). "
                "Blended open-cafe prime cost 61.3% — held vs. 60.6% April; the coffee-COGS "
                "recovery was offset by labor % drift.",
                body_size=9, title_size=10)

    # ----- 1b. Catering trend -----
    add_heading(doc, "1b. Catering Sales — Jan → May 2026", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "All Cafes"],
        [
            ["Jan 2026", "$8,252", "$63",   "$1,525", "$9,700",  "$19,540"],
            ["Feb 2026", "$1,560", "$219",  "$1,035", "$0",      "$2,814"],
            ["Mar 2026", "$11,944","$63",   "$419",   "$29,100", "$41,526"],
            ["Apr 2026", "$3,301", "$28",   "$1,042", "$0",      "$4,371"],
            ["May 2026", "$3,033", "$1,555","$126",   "$334",    "$5,049"],
        ],
        col_widths=[Inches(1.1)] + [Inches(1.2)] * 5,
        total_rows=[4])
    add_caveat(doc,
               "May catering ticked up at Studio ($1,555 vs $28 April) and NMWA had its first "
               "small catering since the March pre-opening events. Downtown held flat. Catering "
               "COGS and margin are tracked in the Consolidated Model but not broken out here.")

    # ----- 1c. Prime Cost & Component Trends -----
    add_heading(doc, "1c. Prime Cost & Component Trends — Jan → May 2026 (ex-catering)", level=2)
    add_para(doc,
             "Prime Cost = COGS + Labor, all expressed as % of Core Sales. Trends below show how "
             "each lever is moving independently. May story: COGS came back into band, Labor "
             "drifted up — net prime cost held roughly flat at 61%.",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    add_heading(doc, "Prime Cost % of Core Sales", level=3)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes Blended"],
        [
            ["Jan 2026", "66.1%",  "66.2%", "100.6%", "—", "70.0%"],
            ["Feb 2026", "62.1%",  "64.2%", "92.6%",  "—", "67.2%"],
            ["Mar 2026", "55.7%",  "51.5%", "81.6%",  "—", "56.9%"],
            ["Apr 2026", "63.6%",  "53.4%", "75.2%",  "—", "60.6%"],
            ["May 2026", "65.9%",  "52.4%", "82.2%",  "—", "61.3%"],
        ],
        col_widths=[Inches(1.1)] + [Inches(1.1)] * 4 + [Inches(1.5)],
        total_rows=[4])

    add_heading(doc, "COGS % of Core Sales", level=3)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes Blended"],
        [
            ["Jan 2026", "30.7%", "25.7%", "31.6%", "—", "28.0%"],
            ["Feb 2026", "34.7%", "32.4%", "43.2%", "—", "34.7%"],
            ["Mar 2026", "30.4%", "26.7%", "34.3%", "—", "29.0%"],
            ["Apr 2026", "39.7%", "30.8%", "35.1%", "—", "34.8%"],
            ["May 2026", "35.6%", "25.8%", "37.2%", "—", "30.8%"],
        ],
        col_widths=[Inches(1.1)] + [Inches(1.1)] * 4 + [Inches(1.5)],
        total_rows=[4])

    add_heading(doc, "Labor % of Core Sales", level=3)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes Blended"],
        [
            ["Jan 2026", "35.4%", "40.5%", "69.0%", "—", "42.0%"],
            ["Feb 2026", "27.4%", "31.8%", "49.4%", "—", "32.5%"],
            ["Mar 2026", "25.3%", "24.9%", "47.3%", "—", "28.0%"],
            ["Apr 2026", "23.9%", "22.6%", "40.0%", "—", "25.9%"],
            ["May 2026", "30.3%", "26.6%", "45.0%", "—", "30.5%"],
        ],
        col_widths=[Inches(1.1)] + [Inches(1.1)] * 4 + [Inches(1.5)],
        total_rows=[4])

    add_caveat(doc,
               "Target bands (Cafe Finance 101): Prime Cost 55–60% healthy / 60–65% warning / "
               ">65% problem. Coffee COGS 22–27% of coffee revenue (see 1a). May reversal "
               "directions: COGS % stepped back down (the April spike receded) but Labor % rose "
               "at every cafe (Downtown +6.4 pts, Studio +4.0 pts, 1899 +5.0 pts) — on softer "
               "May sales, hours weren't fully scaled back. That's the next lever to pull.")

    # ----- 1d. Core Sales Trend (Core-only) -----
    add_heading(doc, "1d. Core Sales Trend — Jan → May 2026 (ex-catering)", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "NMWA", "Open Cafes"],
        [
            ["Jan 2026", "$24,669", "$41,813", "$8,324",  "($970)", "$74,807"],
            ["Feb 2026", "$32,323", "$41,578", "$11,340", "$0",     "$85,241"],
            ["Mar 2026", "$39,027", "$55,009", "$14,100", "($35)",  "$108,136"],
            ["Apr 2026", "$43,482", "$56,232", "$19,291", "$68",    "$119,005"],
            ["May 2026", "$37,010", "$54,283", "$15,106", "($300)", "$106,399"],
        ],
        col_widths=[Inches(1.35)] + [Inches(1.15)] * 5,
        total_rows=[4])
    add_caveat(doc,
               "Core Sales = cafe operations net of catering. Open Cafes column excludes NMWA "
               "(pre-opening). May Open-cafe Core = $37,010 + $54,283 + $15,106 = $106,399, "
               "down $12,606 (-10.6%) vs April. Downtown -14.9%, 1899 -21.7%, Studio -3.5%.")

    # ----- 1e. Actual vs Projected (May) -----
    add_heading(doc, "1e. Actual vs. Projected — May 2026", level=2)
    add_para(doc,
             "Source: 2026 Projections v1.2 (FY26 budget). Projections are daily Net Sales and Labor $ "
             "rolled to monthly. Compared to May actual Core Sales (ex-catering) and Total Labor. "
             "NMWA May projection is $0 because it remains pre-opening.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Cafe", "Sales Actual", "Sales Proj", "Sales Δ ($)", "Sales Δ (%)", "Labor Actual", "Labor Proj", "Labor Δ ($)"],
        [
            ["Downtown",      "$37,010",  "$35,259",  "+$1,751",  "+5.0%",  "$11,206", "$12,693", "($1,487)"],
            ["Studio",        "$54,283",  "$53,850",  "+$433",    "+0.8%",  "$14,467", "$19,386", "($4,919)"],
            ["1899 Penn",     "$15,106",  "$14,625",  "+$481",    "+3.3%",  "$6,804",  "$7,313",  "($509)"],
            ["NMWA",          "($300)",   "$0",       "($300)",   "—",       "$755",    "$0",      "+$755"],
            ["Open Cafes",    "$106,399", "$103,734", "+$2,665",  "+2.6%",  "$32,477", "$39,392", "($6,915)"],
        ],
        col_aligns=["left"] + ["right"] * 7,
        col_widths=[Inches(1.2), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.85), Inches(0.95), Inches(0.95), Inches(0.95)],
        total_rows=[4])
    add_caveat(doc,
               "Open Cafes row excludes NMWA. May story vs April: the +17.4% sales-over-plan margin "
               "April had collapsed to +2.6% — sales softened materially closer to plan. Labor still "
               "under plan ($6,915 favorable) but the gap narrowed (April was $6,624 with much higher "
               "sales). 1899 absolute labor $ also came in under plan for the first time.")


    # ========== 2. Operating Indicators ==========
    add_heading(doc, "2. Operating Indicators", level=1)
    add_para(doc,
             "Source: Consolidated Model v6. NMWA excluded — pre-opening, no meaningful ticket/transaction averages.",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    add_heading(doc, "2a. Average Ticket ($) — Jan → May 2026", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "Blended (ex-NMWA)"],
        [
            ["Jan 2026", "$8.41", "$9.46", "$7.43", "$8.43"],
            ["Feb 2026", "$8.76", "$9.45", "$7.05", "$8.42"],
            ["Mar 2026", "$8.56", "$9.50", "$7.21", "$8.42"],
            ["Apr 2026", "$8.70", "$9.76", "$8.01", "$8.82"],
            ["May 2026", "$8.76", "$8.84", "$7.17", "$8.46"],
        ],
        col_widths=[Inches(1.2)] + [Inches(1.4)] * 4,
        total_rows=[4])
    add_caveat(doc,
               "Per-transaction Square averages (consistent with the April adjustment). Studio's May "
               "drop from $9.76 to $8.84 is notable — the April figure was elevated by the eat-in "
               "side; May's read is closer to the Jan–Mar baseline. 1899 also slipped (-$0.84) on a "
               "smaller sales base. Blended $8.46 is back near the Q1 trend.")

    add_heading(doc, "2b. Average Daily Transactions — Jan → May 2026", level=2)
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "Total/Day"],
        [
            ["Jan 2026", "194", "138", "73",  "404"],
            ["Feb 2026", "274", "152", "116", "541"],
            ["Mar 2026", "297", "181", "125", "603"],
            ["Apr 2026", "318", "192", "159", "669"],
            ["May 2026", "304", "183", "137", "624"],
        ],
        col_widths=[Inches(1.2)] + [Inches(1.4)] * 4,
        total_rows=[4])

    # ========== 3. Item-Level Category Insights ==========
    add_heading(doc, "3. Item-Level Category Insights — May 2026", level=1)
    add_para(doc,
             "Source: Square Item Sales export (May 2026), three open locations combined (Downtown, "
             "Studio, 1899). NMWA excluded — pre-opening. Quantities are units sold across locations; "
             "gross sales are pre-discount, pre-tax. Filter Coffee and Retail (whole bean) categories "
             "are reported separately in 3d.",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    # Other Categories callout (per standing rule — never silently fold into Food)
    add_callout(doc,
                "Other Categories rollup (Downtown only)",
                "Square's 'Other Categories' rollup at Downtown for May totals $20,046 across 3 items: "
                "Management Fee $18,950 (4 transactions), Equipment Fee $1,091 (1), Iced Latte $5 (1). "
                "These are intercompany / accounting transactions running through the Downtown POS — "
                "they are NOT folded into Food or Coffee for the Section 3 rollups below, and they "
                "explain why Downtown's item-CSV gross ($61,059) sits above v7's Total Sales ($40,043). "
                "Net of those Other items, the CSV reconciles to v7.",
                body_size=9, title_size=10)

    slices = json.load(open(ITEM_SLICES_JSON))

    # ----- 3a. Coffee Rollup (ex Filter & Retail) -----
    add_heading(doc, "3a. Coffee Roll-up — Top & Bottom 10 (ex Filter & Retail, >5 units, combined)", level=2)
    add_para(doc, "Top 10 by units sold.", size=9, color=COLOR_MUTED, after=4)
    rows = [[i + 1, r["Item Name"], f"{int(r['units']):,}", f"${r['gross']:,.0f}"]
            for i, r in enumerate(slices["coffee_top10"])]
    build_investor_table(doc,
        ["#", "Item", "Units", "Gross Sales"],
        rows,
        col_aligns=["right", "left", "right", "right"],
        col_widths=[Inches(0.45), Inches(3.5), Inches(1.0), Inches(1.5)])

    add_para(doc, "Bottom 10 by units sold (filtered to items with >5 units).", size=9, color=COLOR_MUTED, after=4)
    rows = [[i + 1, r["Item Name"], f"{int(r['units']):,}", f"${r['gross']:,.0f}"]
            for i, r in enumerate(slices["coffee_bot10"])]
    build_investor_table(doc,
        ["#", "Item", "Units", "Gross Sales"],
        rows,
        col_aligns=["right", "left", "right", "right"],
        col_widths=[Inches(0.45), Inches(3.5), Inches(1.0), Inches(1.5)])

    # ----- 3b. Food Rollup (ex Pastry) -----
    add_heading(doc, "3b. Food Roll-up — Top & Bottom 10 (ex Pastry, >5 units, combined)", level=2)
    add_para(doc, "Top 10 by units sold.", size=9, color=COLOR_MUTED, after=4)
    rows = [[i + 1, r["Item Name"], f"{int(r['units']):,}", f"${r['gross']:,.0f}"]
            for i, r in enumerate(slices["food_top10"])]
    build_investor_table(doc,
        ["#", "Item", "Units", "Gross Sales"],
        rows,
        col_aligns=["right", "left", "right", "right"],
        col_widths=[Inches(0.45), Inches(3.5), Inches(1.0), Inches(1.5)])

    add_para(doc, "Bottom 10 by units sold (filtered to items with >5 units).", size=9, color=COLOR_MUTED, after=4)
    rows = [[i + 1, r["Item Name"], f"{int(r['units']):,}", f"${r['gross']:,.0f}"]
            for i, r in enumerate(slices["food_bot10"])]
    build_investor_table(doc,
        ["#", "Item", "Units", "Gross Sales"],
        rows,
        col_aligns=["right", "left", "right", "right"],
        col_widths=[Inches(0.45), Inches(3.5), Inches(1.0), Inches(1.5)])

    # ----- 3c. Pastry (all >3u) -----
    add_heading(doc, "3c. Pastry — All Items with >3 Units Sold (combined)", level=2)
    rows = [[i + 1, r["Item Name"], f"{int(r['units']):,}", f"${r['gross']:,.0f}"]
            for i, r in enumerate(slices["pastry_over3"])]
    build_investor_table(doc,
        ["#", "Item", "Units", "Gross Sales"],
        rows,
        col_aligns=["right", "left", "right", "right"],
        col_widths=[Inches(0.45), Inches(3.5), Inches(1.0), Inches(1.5)])

    # ----- 3d. Filter Coffee + Retail by location -----
    add_heading(doc, "3d. Filter Coffee & Retail — Units by Location", level=2)
    add_para(doc,
             "Filter Coffee units grouped into Comfort (Cruise Control), Daily Selection, Nitro Iced "
             "Coffee, and Other Filter (single-origin: Bright / Decaf / Floral / Fruit). Total Retail is "
             "the sum of all whole-bean + boxed retail SKUs. Percentages in the four Filter rows show "
             "each category's share of that location's Total Filter Coffee.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    f = slices["filter_by_loc"]
    rt = slices["retail_total_by_loc"]
    filter_total = {loc: sum(int(f[k][loc]) for k in f) for loc in ["Downtown", "Studio", "1899"]}
    filter_total_combined = sum(filter_total.values())

    def _row_with_pct(label, key):
        """Main filter-category rows: show units and (% of that column's Total Filter Coffee)."""
        def fmt(loc, combined=False):
            units = int(f[key][loc]) if not combined else int(f[key]['Downtown']) + int(f[key]['Studio']) + int(f[key]['1899'])
            denom = filter_total_combined if combined else filter_total[loc]
            pct = (units / denom * 100) if denom else 0
            return f"{units:,} ({pct:.1f}%)"
        return [label, fmt('Downtown'), fmt('Studio'), fmt('1899'), fmt(None, combined=True)]

    rows_3d = [
        _row_with_pct("Comfort (Cruise Control)", "Comfort (Cruise Control)"),
        _row_with_pct("Daily Selection", "Daily Selection"),
        _row_with_pct("Nitro Iced Coffee", "Nitro"),
        _row_with_pct("Other Filter (single-origin)", "Other Filter"),
        ["Total Filter Coffee",
         f"{filter_total['Downtown']:,}",
         f"{filter_total['Studio']:,}",
         f"{filter_total['1899']:,}",
         f"{filter_total_combined:,}"],
        ["Total Retail (whole bean + boxed)",
         f"{int(rt['Downtown']):,}",
         f"{int(rt['Studio']):,}",
         f"{int(rt['1899']):,}",
         f"{int(rt['Downtown'])+int(rt['Studio'])+int(rt['1899']):,}"],
    ]
    build_investor_table(doc,
        ["Category", "Downtown", "Studio", "1899 Penn", "Combined"],
        rows_3d,
        col_aligns=["left", "right", "right", "right", "right"],
        col_widths=[Inches(2.6), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.2)],
        total_rows=[4, 5])

    # Premium Hot Filter share trend (Mar / Apr / May)
    add_heading(doc, "Premium Hot Filter Share — (Daily Selection + Single Origin) / Total Hot Filter", level=3)
    mar_slices_path = os.path.join(DATA_DIR, "mar26_item_slices.json")
    apr_slices_path = os.path.join(DATA_DIR, "apr26_item_slices.json")
    historical_slices = []
    if os.path.exists(mar_slices_path):
        historical_slices.append(("Mar 2026", json.load(open(mar_slices_path))["filter_by_loc"]))
    if os.path.exists(apr_slices_path):
        historical_slices.append(("Apr 2026", json.load(open(apr_slices_path))["filter_by_loc"]))
    # May is the current month — from `f` (current slices' filter_by_loc)
    historical_slices.append(("May 2026", f))

    def _premium_row(period, ff):
        def cell(loc):
            ds = int(ff["Daily Selection"][loc])
            so = int(ff["Other Filter"][loc])
            cf = int(ff["Comfort (Cruise Control)"][loc])
            hot_total = ds + so + cf
            premium = ds + so
            pct = (premium / hot_total * 100) if hot_total else 0
            return f"{premium:,} / {hot_total:,} ({pct:.1f}%)"
        ds_c = sum(int(ff["Daily Selection"][l]) for l in ["Downtown","Studio","1899"])
        so_c = sum(int(ff["Other Filter"][l]) for l in ["Downtown","Studio","1899"])
        cf_c = sum(int(ff["Comfort (Cruise Control)"][l]) for l in ["Downtown","Studio","1899"])
        hot_c = ds_c + so_c + cf_c
        pct_c = ((ds_c+so_c)/hot_c*100) if hot_c else 0
        return [period, cell("Downtown"), cell("Studio"), cell("1899"),
                f"{ds_c+so_c:,} / {hot_c:,} ({pct_c:.1f}%)"]

    premium_rows = [_premium_row(label, ff) for label, ff in historical_slices]
    build_investor_table(doc,
        ["Period", "Downtown", "Studio", "1899 Penn", "Combined"],
        premium_rows,
        col_aligns=["left", "right", "right", "right", "right"],
        col_widths=[Inches(1.0), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)],
        total_rows=[len(premium_rows) - 1])
    add_caveat(doc,
               "Premium Hot Filter Share = (Daily Selection + Single Origin) units divided by "
               "hot-filter total (Daily Selection + Single Origin + Comfort). Nitro is iced so "
               "excluded from the hot-filter denominator. May combined held at 21.2% — flat vs "
               "April's 22.3%, slightly down. 1899's premium share continues to ease (32.7% → "
               "31.5% → 28.2%) — worth watching whether guests are shifting back to Comfort or "
               "to Nitro.")

    # ========== 4. Discount Report (deferred this month) ==========
    add_heading(doc, "4. Discount Report — May 2026", level=1)
    add_para(doc,
             "Held this month per Ryan. The March 2026 baseline (~3.66% discount-to-gross blended, "
             "Downtown highest at 4.42%) remains the reference point. To resume, we need a Square "
             "Orders re-pull (or a one-off Discounts export) with discount name/type populated.",
             italic=True, size=9, color=COLOR_MUTED, after=12)

    # ========== 5. Loyalty Report ==========
    add_heading(doc, "5. Loyalty Report — May 2026", level=1)
    add_para(doc,
             "Source: Square Loyalty (per-cafe screenshots). 5b Loyalty Member Value now uses Square's "
             "actual 'average spend per customer' metric directly, replacing the visits × avg ticket "
             "estimate used in April. Visit-bucket schema changed AGAIN in May — buckets are now "
             "1, 2, 3, 4+ (simpler than April's 1/2/3-4/5-7-or-8/8+-or-9+).",
             italic=True, size=9, color=COLOR_MUTED, after=6)

    add_heading(doc, "5a. Headline Insights", level=2)
    add_bullet(doc, "Loyalty avg visits Downtown 2.6, Studio 2.6, 1899 3.0 — vs non-loyalty 1.3–1.4. "
                    "Downtown and Studio softened slightly vs April (2.8 / 2.7); 1899 held its 3.0 lead "
                    "and remains the deepest per-member engagement in the system.",
               bold_lead="Loyalty member frequency held at 1899, eased at DT/Studio.")
    add_bullet(doc, "39 → 73 → 91 → 101 → 90. First decline since January. Repeat visits also fell "
                    "(1,759 → 1,539). Worth understanding — May seasonality, staff push fatigue, or a "
                    "specific calendar event affecting visitor mix.",
               bold_lead="Downtown first-time signups pulled back.")
    add_bullet(doc, "118 → 101 → 118 → 120 → 120. Acquisition still at its modest steady rate. Repeat "
                    "visits ticked down slightly (2,029 → 1,940). Studio still the largest active loyalty "
                    "base in the system at 2,060 total monthly visits.",
               bold_lead="Studio held flat at 120 first-time signups.")
    add_bullet(doc, "29 → 52 → 64 → 59 → 53. Second consecutive decline after three months of strong "
                    "growth. Repeat visits also softened (785 → 724). Avg visits/member held at 3.0. "
                    "Two-month signup softness at the smallest cafe — worth a check on the in-store "
                    "signup prompt + barista script.",
               bold_lead="1899 first-time signups slipped again.")
    add_bullet(doc, "Per prior note, NMWA is not enrolled in the loyalty program. Enroll once it opens "
                    "to capture early adopters from day one.",
               bold_lead="NMWA not yet enrolled.")

    # ----- 5b. Member LTV in dollars (methodology change — uses Square avg spend directly) -----
    add_heading(doc, "5b. Loyalty Member Value — In Dollars (May 2026)", level=2)
    add_para(doc,
             "Methodology change vs April: Square Loyalty now provides 'average spend per customer' "
             "directly per cafe. May figures use that metric × 12 for annual instead of "
             "visits × avg ticket. The Square metric is higher (it captures multi-transaction days "
             "and stronger spenders) — annual loyalty value at Downtown jumped from $292 (April est.) "
             "to $404 (May actual) on Square's read. 5-yr value at current rate, no decay assumption.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Cafe", "Avg Spend / Customer / Mo", "Loyalty ($/yr)", "Non-Loyalty ($/yr)", "Δ per member / yr", "Loyalty 5-yr value"],
        [
            ["Downtown",       "L $33.67 / NL $25.03", "$404", "$300", "+$104", "$2,020"],
            ["Studio Theatre", "L $29.08 / NL $16.82", "$349", "$202", "+$147", "$1,745"],
            ["1899 Penn",      "L $24.73 / NL $13.55", "$297", "$163", "+$134", "$1,484"],
            ["Blended (ex-NMWA)","L $29.16 / NL $18.47","$350", "$222", "+$128", "$1,750"],
        ],
        col_aligns=["left", "left", "right", "right", "right", "right"],
        col_widths=[Inches(1.5), Inches(1.5), Inches(1.0), Inches(1.1), Inches(1.1), Inches(1.1)],
        total_rows=[3])
    add_caveat(doc,
               "Source: Square Loyalty 'Average spend per customer' (monthly figure). Multiplied by 12 "
               "for annual. Blended = simple average across the 3 open cafes. Downtown loyalty members "
               "are worth ~$2,020 / 5 yrs vs $1,500 for non-loyalty — that's an extra ~$520 per member "
               "per 5-year arc. Studio's premium is the largest at $147/yr ($735 / 5yr).")

    loyalty_cafes = [
        ("5c. Downtown — Loyalty Detail",
         [["First-time loyalty visits", "39",    "73",    "91",    "101",   "90"],
          ["Repeat loyalty visits",     "1,164", "1,363", "1,661", "1,759", "1,539"]],
         [["Loyalty avg visits / member",         "2.5", "2.6", "2.8", "2.8", "2.6"],
          ["Non-loyalty avg visits / customer",   "1.4", "1.4", "1.4", "1.3", "1.4"]],
         # Three schemas: Jan-Mar old (1,2,3,4-6,7+), Apr (1,2,3-4,5-7,8+), May (1,2,3,4+)
         [["1 visit",         "242", "284", "311", "322", "200"],
          ["2 visits",        "83",  "110", "117", "127", "51"],
          ["3 visits",        "55",  "59",  "113", "—",   "20"],
          ["4–6 visits",      "55",  "66",  "60",  "—",   "—"],
          ["7+ visits",       "46",  "49",  "39",  "—",   "—"],
          ["3–4 visits",      "—",   "—",   "—",   "116", "—"],
          ["5–7 visits",      "—",   "—",   "—",   "66",  "—"],
          ["8+ visits",       "—",   "—",   "—",   "41",  "—"],
          ["4+ visits",       "—",   "—",   "—",   "—",   "12"]]),
        ("5d. Studio Theatre — Loyalty Detail",
         [["First-time loyalty visits", "118",   "101",   "118",   "120",   "120"],
          ["Repeat loyalty visits",     "1,448", "1,531", "2,003", "2,029", "1,940"]],
         [["Loyalty avg visits / member",         "2.3", "2.4", "2.7", "2.7", "2.6"],
          ["Non-loyalty avg visits / customer",   "1.3", "1.3", "1.3", "1.3", "1.3"]],
         [["1 visit",         "408", "384", "460", "454", "114"],
          ["2 visits",        "128", "135", "139", "146", "20"],
          ["3 visits",        "55",  "84",  "72",  "—",   "16"],
          ["4–6 visits",      "69",  "37",  "74",  "—",   "—"],
          ["7+ visits",       "38",  "36",  "58",  "—",   "—"],
          ["3–4 visits",      "—",   "—",   "—",   "113", "—"],
          ["5–8 visits",      "—",   "—",   "—",   "58",  "—"],
          ["9+ visits",       "—",   "—",   "—",   "45",  "—"],
          ["4+ visits",       "—",   "—",   "—",   "—",   "10"]]),
        ("5e. 1899 Penn — Loyalty Detail",
         [["First-time loyalty visits", "29",  "52",  "64",  "59",  "53"],
          ["Repeat loyalty visits",     "349", "504", "620", "785", "724"]],
         [["Loyalty avg visits / member",         "2.7", "2.8", "3.0", "3.0", "3.0"],
          ["Non-loyalty avg visits / customer",   "1.4", "1.4", "1.4", "1.4", "1.4"]],
         [["1 visit",         "79",  "110", "137", "134", "62"],
          ["2 visits",        "25",  "34",  "34",  "42",  "19"],
          ["3 visits",        "20",  "28",  "27",  "—",   "11"],
          ["4–6 visits",      "10",  "17",  "22",  "—",   "—"],
          ["7+ visits",       "10",  "17",  "22",  "—",   "—"],
          ["3–4 visits",      "—",   "—",   "—",   "50",  "—"],
          ["5–8 visits",      "—",   "—",   "—",   "27",  "—"],
          ["9+ visits",       "—",   "—",   "—",   "19",  "—"],
          ["4+ visits",       "—",   "—",   "—",   "—",   "5"]]),
    ]
    for head, visits, freq, buckets in loyalty_cafes:
        add_heading(doc, head, level=2)
        build_investor_table(doc,
            ["Metric", "Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "May 2026"], visits,
            col_widths=[Inches(2.5), Inches(1.05), Inches(1.05), Inches(1.05), Inches(1.05), Inches(1.05)])
        build_investor_table(doc,
            ["Metric", "Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "May 2026"], freq,
            col_widths=[Inches(2.5), Inches(1.05), Inches(1.05), Inches(1.05), Inches(1.05), Inches(1.05)])
        build_investor_table(doc,
            ["Visit Bucket", "Jan 2026", "Feb 2026", "Mar 2026", "Apr 2026", "May 2026"], buckets,
            col_widths=[Inches(1.8), Inches(1.15), Inches(1.15), Inches(1.15), Inches(1.15), Inches(1.15)],
            key_rows=[0, 1])
        add_caveat(doc,
                   "Visit bucket schema changed twice. Jan–Mar use 1 / 2 / 3 / 4–6 / 7+; April uses "
                   "1 / 2 / 3–4 / 5–7 (DT) or 5–8 (S/1899) / 8+ (DT) or 9+ (S/1899); May simplifies "
                   "to 1 / 2 / 3 / 4+. 1-visit and 2-visit buckets stable across all schemas.")

    add_heading(doc, "5f. NMWA", level=2)
    add_para(doc,
             "Not enrolled in the loyalty program. Recommend enrolling at launch to capture signups "
             "from day one.",
             italic=True, color=COLOR_MUTED, after=10)

    # ========== 6. Notes ==========
    add_heading(doc, "6. Notes & Watch Items", level=1)
    add_heading(doc, "Data caveats", level=3)
    add_bullet(doc, "Core vs Catering split is pulled from the Consolidated Model v7 (rows 20 and 24 of each "
                    "cafe's Income Statement tab). Catering labor is not separated in the current chart of "
                    "accounts; all labor is shown under Core.")
    add_bullet(doc, "NMWA remains pre-opening as of the close of May (a few test days only per Ryan). "
                    "The Management Fee Credit on NMWA's tab is an intentional accounting offset and is "
                    "included in NMWA's Total Sales but not in blended ratios.")
    add_bullet(doc, "Loyalty data sourced directly from Square Loyalty per-cafe screenshots. May visit-bucket "
                    "schema simplified to 1 / 2 / 3 / 4+ (a third schema in five months — April was 1/2/3-4/"
                    "5-7-or-8/8+-or-9+, Jan-Mar was 1/2/3/4-6/7+).")
    add_bullet(doc, "Section 3 Items: Downtown Square Items CSV gross ($61K) exceeds v7 Total Sales ($40K) "
                    "because of the Management Fee + Equipment Fee transactions running through the Downtown "
                    "POS (Other Categories rollup, $20K). Net of those, the CSV reconciles. Flagged at top "
                    "of Section 3.")
    add_bullet(doc, "Discount Report (Section 4) deferred per Ryan; no discount-by-type re-pull from Square "
                    "this month.")

    add_heading(doc, "Spreadsheet issues found in Consolidated Model v7 (carried over from v6)", level=3)
    add_bullet(doc, "v7 still does NOT include Commissary sales ($9,732) or SRP sales ($9,301) in the "
                    "Consolidated 'Total Sales' line. Same bug pattern as v6. Fix: extend r44 formula to "
                    "include Commissary r254 and SRP r274.",
               bold_lead="Total Sales formula still missing Commissary + SRP.")
    add_bullet(doc, "v7 r291 'SRP Net Income' still mismatches SRP tab's own r37. The Consolidated section "
                    "computes SRP NI from a different expense rollup than the SRP IS tab uses. Fix: re-link "
                    "r291 to SRP tab r37 directly.",
               bold_lead="SRP Net Income mismatch persists.")
    add_bullet(doc, "v7 r294 'Consolidated Net Income' still does NOT include Mgmt entity Net Income "
                    "($5,762 in May). Fix: extend r294 formula to include Mgmt tab NI.",
               bold_lead="Mgmt Net Income still missing from r294.")
    add_bullet(doc, "May spreadsheet r294 = $7,909 vs. corrected $15,750 (sum of all entity NIs from their "
                    "own tabs + Consulting/Other adjustment). The corrected trend appears in Appendix B. "
                    "Restated April corrected NI per v7 is $12,505 — matches what we shipped last month.",
               bold_lead="Net May delta: ~$7,841 understated.")

    add_heading(doc, "Watch items for June", level=3)
    add_bullet(doc, "Labor % rose at every open cafe (Downtown +6.4 pts, Studio +4.0 pts, 1899 +5.0 pts). "
                    "On softer May sales, hours weren't scaled back to match. Plan-vs-actual labor is still "
                    "favorable in dollars but the gap is closing. First priority for June.",
               bold_lead="Labor scaling — the next lever.")
    add_bullet(doc, "Open-cafe Core down -10.6%, Downtown -14.9% and 1899 -21.7%. Worth a same-week look at "
                    "the daily sales curve — was there a single soft week, a holiday effect, or genuine "
                    "demand softening that needs an operational response in June?",
               bold_lead="Sales decline — diagnostic needed.")
    add_bullet(doc, "+$5,023 March → +$3,123 April → -$4,794 May. Coffee sales held ($8,108 → $5,801) but "
                    "concentrate dropped ($4,410 → $3,500). Worth understanding whether this is one-off "
                    "(write-down, timing) or a real demand shift.",
               bold_lead="SRP swung sharply negative.")
    add_bullet(doc, "Food COGS at 65.1% (Apr 57.7%, Mar 73.1%) — direction worse on a small absolute base "
                    "($2,292 food sales). The volatility makes the % unreliable as a target measure here, "
                    "but worth digging into actual food cost per item.",
               bold_lead="1899 food COGS got worse.")
    add_bullet(doc, "Catering ticked up at Studio ($1,555 vs $28 April) and NMWA had its first small "
                    "catering ($334). Worth tracking whether Studio is converting a new sales channel.",
               bold_lead="Catering pickup at Studio + NMWA.")
    add_bullet(doc, "NMWA: confirm opening date and finalize loyalty-program enrollment for launch.")

    # ========== Appendices ==========
    doc.add_paragraph().add_run().add_break()
    p = doc.add_paragraph()
    r = p.add_run("LEADERSHIP TEAM ONLY — APPENDICES")
    _set_run_font(r, size=12, bold=True, color=COLOR_MUTED)
    _set_para_spacing(p, before=10, after=6)

    add_heading(doc, "A. Support Entities — May 2026", level=2)
    add_para(doc,
             "Source: Consolidated Model v7. Full P&L by entity. Catering is included in totals (not broken "
             "out for support entities).",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Metric", "Commissary", "Central Admin", "Mgmt Fees", "SRP (Roasting)", "Support Total"],
        [
            ["Sales / Revenue", "$9,732",   "$0",        "$5,762", "$9,301",   "$24,795"],
            ["COGS",            "$11,585",  "$0",        "$0",     "$6,852",   "$18,437"],
            ["Total Labor",     "$0",       "$8,436",    "$0",     "$0",       "$8,436"],
            ["Expenses",        "$2,468",   "$1,778",    "$0",     "$7,243",   "$11,489"],
            ["Net Income",      "($4,321)", "($10,214)", "$5,762", "($4,794)", "($13,567)"],
        ],
        col_widths=[Inches(1.6), Inches(1.1), Inches(1.15), Inches(1.0), Inches(1.25), Inches(1.15)],
        total_rows=[4])

    add_heading(doc, "A1. Support Entities — Net Income Trend", level=3)
    build_investor_table(doc,
        ["Entity", "Jan", "Feb", "Mar", "Apr", "May", "MoM Δ (Apr→May)"],
        [
            ["Commissary",     "($6,287)",  "($5,509)",  "($3,843)",  "($4,199)",  "($4,321)",  "($122)"],
            ["Central Admin",  "($13,519)", "($10,244)", "($13,794)", "($11,739)", "($10,214)", "+$1,525"],
            ["Mgmt Fees",      "$5,422",    "$5,234",    "$5,693",    "$6,017",    "$5,762",    "($255)"],
            ["SRP (Roasting)", "($3,894)",  "($2,543)",  "($1,697)",  "$3,123",    "($4,794)",  "($7,917)"],
        ],
        col_widths=[Inches(1.6), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95), Inches(1.35)])
    add_caveat(doc,
               "SRP swung sharply negative ($3,123 → -$4,794, a -$7,917 MoM move). Coffee sales held "
               "($8,108 → $5,801), concentrate dropped ($4,410 → $3,500), and roasting labor ticked "
               "up ($1,680 → $2,176). Worth a dedicated review — is this a write-down/timing issue or "
               "a real shift in roasting demand? Central Admin improved by $1,525 (lower legal/consulting "
               "spend).")

    add_heading(doc, "B. Consolidated P&L — May 2026 (Corrected)", level=2)
    add_para(doc,
             "Source: Consolidated Model v7 'Income Statement - Consolidated' tab for sales/COGS/labor/"
             "expense lines, BUT the Consol. Net Income line is restated below as a sum of per-entity "
             "Net Income from each entity's own tab (Downtown / Studio / 1899 / NMWA / Commissary / "
             "Central Admin / Mgmt / SRP) plus the Consulting and Other adjustment. The Consolidated "
             "tab's r294 'Consolidated Net Income' still has the same three formula bugs in v7 (see "
             "Notes section, 'Spreadsheet issues found'). May corrected NI = $15,750.",
             italic=True, size=9, color=COLOR_MUTED, after=6)
    build_investor_table(doc,
        ["Metric", "Jan", "Feb", "Mar", "Apr", "May", "MoM Δ"],
        [
            ["Total Sales (all entities)",  "$99,769",  "$93,289",  "$158,230", "$134,506", "$122,929", "(8.6%)"],
            ["Total COGS",                  "$25,364",  "$29,833",  "$31,603",  "$41,728",  "$32,787",  "(21.4%)"],
            ["Total Labor",                 "$49,653",  "$41,355",  "$49,864",  "$49,521",  "$50,955",  "+2.9%"],
            ["Total Expenses",              "$20,448",  "$17,128",  "$28,623",  "$34,575",  "$23,145",  "(33.1%)"],
            ["NOI (ex Depreciation, sum of entity NIs)", "$235",  "($2,086)", "$46,730",  "$13,757",  "$16,668",  "+21.2%"],
            ["Consulting and Other adjustment",          "($422)","($234)",   "($693)",   "($1,253)", "($918)",   "(26.7%)"],
            ["Consolidated Net Income (corrected)",      "($187)","($2,320)", "$46,037",  "$12,505",  "$15,750",  "+25.9%"],
            ["  (Spreadsheet r294 — uncorrected)",       "($3,844)","($1,193)","$43,807", "$8,252",   "$7,909",   "(4.2%)"],
        ],
        col_widths=[Inches(2.1), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95), Inches(0.95)],
        key_rows=[4, 5], total_rows=[6])
    add_caveat(doc,
               "May corrected NI of $15,750 represents the true bottom line — driven by a $11,430 drop "
               "in Total Expenses despite the 8.6% sales decline. The spreadsheet r294 figure of $7,909 "
               "understates by $7,841 (Mgmt $5,762 missing + SRP overstatement). Note v7 restated some "
               "Jan-Apr line items (Jan COGS up ~$4K, restated entity NIs). The Mar peak still reflects "
               "NMWA's pre-opening catering event(s) — not a steady-state.")

    add_heading(doc, "C. Methodology & Data Sources", level=2)
    add_heading(doc, "Data sources", level=3)
    add_bullet(doc, "Financials & Core vs Catering split: Slipstream Consolidated Model 2026 v7.xlsx — "
                    "per-entity Income Statement tabs. May 2026 columns: Downtown BS, Studio AM, 1899 AA, "
                    "NMWA O; Consolidated BK; Support entities (Comm/Central/Mgmt/SRP) col AA.")
    add_bullet(doc, "Operating indicators (avg ticket, avg daily transactions): same spreadsheet, "
                    "per-transaction 'Ave Ticket' (row 51/54/56) and 'Total - Ave Daily Transactions' "
                    "rows on each cafe tab.")
    add_bullet(doc, "Projections (Section 1e): 2026 Projections v1.2.xlsx — per-cafe monthly summary "
                    "row 11 (May), columns R (Sales Proj) and T (Labor Proj).")
    add_bullet(doc, "Items: Square Item Sales export (May 2026), three open locations. Other Categories "
                    "rollup flagged separately (Management Fee + Equipment Fee at Downtown).")
    add_bullet(doc, "Loyalty: per-cafe Square Loyalty screenshots (sales comparison + loyalty visits + "
                    "visit frequency buckets) — 1899 visits screenshot pending.")

    add_heading(doc, "Definitions", level=3)
    add_bullet(doc, "Core Sales = Total Core Location Sales − Total Catering and Events Sales.")
    add_bullet(doc, "Core COGS = Total Core Location COGS − Total Catering and Events COGS.")
    add_bullet(doc, "Core Prime Cost = Core COGS + Total Labor. (Labor is not yet allocated between core and catering.)")
    add_bullet(doc, "Core Prime Margin = Core Sales − Core Prime Cost.")
    add_bullet(doc, "Catering Prime Margin = Catering Sales − Catering COGS. (Labor excluded; see above.)")
    add_bullet(doc, "Blended / Open Cafes Total = Downtown + Studio Theatre + 1899 Penn (excludes NMWA while pre-opening).")

    add_heading(doc, "Refresh process", level=3)
    add_para(doc,
             f"Monthly report is built by running build_monthly_report.py against updated copies of the "
             f"Consolidated Model and Loyalty Summary workbooks plus refreshed Square item-sales exports. "
             f"Output: Slipstream Monthly Report - [Month] [Year] {VERSION}.docx.",
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
