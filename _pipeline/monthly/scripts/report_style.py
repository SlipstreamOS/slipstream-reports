"""Shared docx formatting helpers for the Slipstream monthly report (investor style)."""
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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



def _ensure_zoom_percent(doc):
    """python-docx emits <w:zoom /> without the required w:percent attribute.
    Word is lenient but strict validators reject it. Set a default percent.
    """
    settings_part = doc.settings.element
    zoom = settings_part.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")


