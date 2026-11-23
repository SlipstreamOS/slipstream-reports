#!/usr/bin/env python3
"""Render the weekly report docx (source of numbers + narrative) as a clean, readable PDF.
    python3 build_weekly_pdf.py "<path to Slipstream Weekly Report - ... .docx>"
Writes the .pdf next to the docx."""
import sys, re, os, json
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame, Paragraph as P, Table as T,
                                TableStyle, Spacer, KeepTogether, PageBreak)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

FD = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("F", FD + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("FB", FD + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("FC", FD + "DejaVuSansCondensed.ttf"))
pdfmetrics.registerFont(TTFont("FCB", FD + "DejaVuSansCondensed-Bold.ttf"))

INK = colors.HexColor("#1F2933"); MUTE = colors.HexColor("#5F6B76"); LINE = colors.HexColor("#D9DEE3")
BAND = colors.HexColor("#F3F5F7"); HEAD = colors.HexColor("#2F3E4C"); ACC = colors.HexColor("#E8EEF3")
GREEN = colors.HexColor("#2E7D32"); GREEN_BG = colors.HexColor("#E6F2E7")
AMBER = colors.HexColor("#B7791F"); AMBER_BG = colors.HexColor("#FBF1DC")
RED = colors.HexColor("#C62828"); RED_BG = colors.HexColor("#FBE3E3")
BANDS = {"Downtown": (.32, .36, .38), "Studio": (.30, .34, .36), "1899 Penn": (.35, .40, .42)}

def clean(s):
    s = s.replace(" ", " ")
    s = re.sub(r"[✓△✗●]", "", s)
    s = s.replace(" — ", ": ").replace("—", "–").replace("Δ", "Difference")
    return re.sub(r"\s+", " ", s).strip()

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def table_rows(t):
    rows = []
    for r in t.rows:
        cells, prev = [], None
        for c in r.cells:
            if c._tc is prev: continue
            prev = c._tc; cells.append(clean(c.text))
        rows.append(cells)
    return rows

def parse(path):
    d = Document(path)
    out = {"tables": [], "paras": [], "bullets": [], "title": "", "week": ""}
    head = ""; sub = ""
    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            p = Paragraph(el, d); t = p.text.strip()
            if not t: continue
            st = p.style.name
            if st.startswith("Heading 1"): head, sub = clean(t), ""; continue
            if st.startswith("Heading"): sub = clean(t); continue
            if "List" in st: out["bullets"].append((head, t)); continue
            out["paras"].append((head, sub, t))
        elif tag == "tbl":
            out["tables"].append((head, sub, table_rows(Table(el, d))))
    out["week"] = next(t for h, s, t in out["paras"] if t.startswith("Week of"))
    return out

def S(name, **kw):
    base = dict(fontName="F", fontSize=8.6, leading=11.4, textColor=INK)
    base.update(kw); return ParagraphStyle(name, **base)
body = S("b"); small = S("s", fontSize=7.4, leading=9.6, textColor=MUTE)
h1 = S("h1", fontName="FB", fontSize=12.5, leading=15, spaceBefore=8, spaceAfter=5, textColor=HEAD)
h2 = S("h2", fontName="FB", fontSize=9.6, leading=12, spaceBefore=6, spaceAfter=3, textColor=HEAD)
cell = S("c", fontName="FC", fontSize=7.8, leading=9.4)
cellb = S("cb", fontName="FCB", fontSize=7.8, leading=9.4)

def money_to_num(s):
    s = s.replace("$", "").replace(",", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try: v = float(s.rstrip("%"))
    except ValueError: return None
    return -v if neg else v

def std_table(rows, widths, header_rows=1, first_bold=False, num_from=1, section_rows=(), bold_rows=(), bold_cols=(), extra=()):
    st = [("FONT", (0, 0), (-1, -1), "FC", 7.8), ("TEXTCOLOR", (0, 0), (-1, -1), INK),
          ("FONT", (0, 0), (-1, header_rows - 1), "FCB", 7.6),
          ("BACKGROUND", (0, 0), (-1, header_rows - 1), HEAD), ("TEXTCOLOR", (0, 0), (-1, header_rows - 1), colors.white),
          ("ALIGN", (num_from, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 2.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.6),
          ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
          ("LINEBELOW", (0, header_rows), (-1, -1), 0.3, LINE)]
    for i in range(header_rows, len(rows)):
        if (i - header_rows) % 2 == 1: st.append(("BACKGROUND", (0, i), (-1, i), BAND))
    for i in section_rows:
        st += [("BACKGROUND", (0, i), (-1, i), ACC), ("FONT", (0, i), (-1, i), "FCB", 7.6), ("SPAN", (0, i), (-1, i)),
               ("ALIGN", (0, i), (-1, i), "LEFT")]
    for i in bold_rows: st += [("FONT", (0, i), (-1, i), "FCB", 7.8), ("LINEABOVE", (0, i), (-1, i), 0.6, INK)]
    for c in bold_cols: st.append(("FONT", (c, header_rows), (c, -1), "FCB", 7.8))
    if first_bold: st.append(("FONT", (0, header_rows), (0, -1), "FCB", 7.8))
    st += list(extra)
    t = T(rows, colWidths=widths, repeatRows=header_rows); t.setStyle(TableStyle(st)); return t

def chip(text, kind):
    fg, bg = {"g": (GREEN, GREEN_BG), "a": (AMBER, AMBER_BG), "r": (RED, RED_BG)}[kind]
    t = T([[text]], colWidths=None)
    t.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "FCB", 7.4), ("TEXTCOLOR", (0, 0), (-1, -1), fg),
                           ("BACKGROUND", (0, 0), (-1, -1), bg), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                           ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
                           ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5)]))
    return t


import csv, glob, datetime as _dt
TX_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "shared", "transactions")
LOCMAP = {"Slipstream Downtown": "Downtown", "Slipstream Studio": "Studio", "Slipstream 1899 Penn": "1899 Penn", "Slipstream NMWA": "NMWA"}
def _filter_class(name):
    n = name.lower()
    if "whole bean" in n or "nitro" in n: return None          # retail beans and iced nitro are not hot filter
    if "cruise control" in n: return "cc"
    if "daily selection" in n: return "ds"
    if re.match(r"\[(bright|fruit|floral|decaf)\]", n): return "so"
    return None
def filter_mix(start, weeks_back=6):
    """Hot filter units by cafe for the week starting `start` and the weeks before it."""
    first = start - _dt.timedelta(7 * (weeks_back - 1)); last = start + _dt.timedelta(6)
    out = {}
    for f in sorted(glob.glob(os.path.join(TX_DIR, "*.csv"))):
        ym = os.path.basename(f)[:7]
        if ym < first.strftime("%Y-%m") or ym > last.strftime("%Y-%m"): continue
        for r in csv.DictReader(open(f, newline="")):
            if r.get("Event Type") != "Payment": continue
            try: d = _dt.date.fromisoformat(r["Date"])
            except Exception: continue
            if not (first <= d <= last): continue
            cafe = LOCMAP.get(r.get("Location"))
            if not cafe: continue
            wk = first + _dt.timedelta(7 * ((d - first).days // 7))
            for part in (r.get("Description") or "").split(","):
                part = part.strip(); m = re.match(r"^(\d+) x (.*)$", part)
                q, item = (int(m.group(1)), m.group(2)) if m else (1, part)
                c = _filter_class(re.sub(r"\s*\(.*\)$", "", item))
                if c:
                    k = out.setdefault(wk, {}).setdefault(cafe, {"cc": 0, "ds": 0, "so": 0}); k[c] += q
    return out

def is_tc(t):
    return bool(re.match(r"^WATCH\s*[—:-]*\s*TIMECARD", t.strip(), re.I))

def write_timecards(path, week, rows, notes):
    """Ryan-only: timecards to fix in Square before payroll (kept out of the shared report)."""
    out = os.path.join(os.path.dirname(path), f"Timecards to Fix - {week}.pdf")
    doc = SimpleDocTemplate(out, pagesize=letter, leftMargin=.6 * inch, rightMargin=.6 * inch, topMargin=.6 * inch, bottomMargin=.6 * inch)
    W = letter[0] - 1.2 * inch
    el = [P(f"Timecards to Fix Before Payroll: {esc(week)}", h1),
          P("For Ryan only; not in the shared weekly report. The weekly report already uses the corrected hours, but Square still holds "
            "the raw punches, so payroll pays them unless they are edited in Square Teams.", small), Spacer(1, 6)]
    if rows:
        rows = [x for x in rows if any(x)]
        el.append(std_table(rows, [1.0*inch, 1.05*inch, .75*inch, .8*inch, 1.15*inch, .8*inch, .85*inch, .9*inch], num_from=5))
    for n_ in notes:
        el += [Spacer(1, 5), P(esc(re.sub(r"^WATCH\s*:?\s*", "", n_)), body)]
    if not rows and not notes: el.append(P("None this week.", body))
    doc.build(el)

def build(path):
    R = parse(path)
    tabs = R["tables"]
    def tab(head, sub=None):
        return [r for h, s, r in tabs if h.startswith(head) and (sub is None or s == sub)]
    paras = R["paras"]
    out = os.path.splitext(path)[0] + ".pdf"
    week = R["week"].replace("Week of Sunday ", "").replace(" – Saturday ", " to ")

    def on_page(c, doc):
        c.saveState(); c.setFont("F", 7); c.setFillColor(MUTE)
        c.drawString(0.6 * inch, 0.42 * inch, f"Slipstream Weekly Business Review  ·  {week}")
        c.drawRightString(letter[0] - 0.6 * inch, 0.42 * inch, f"Page {doc.page}")
        c.restoreState()
    doc = BaseDocTemplate(out, pagesize=letter, leftMargin=.6 * inch, rightMargin=.6 * inch,
                          topMargin=.55 * inch, bottomMargin=.65 * inch,
                          title=f"Slipstream Weekly Report {week}", author="Slipstream")
    W = letter[0] - 1.2 * inch
    doc.addPageTemplates([PageTemplate(frames=[Frame(doc.leftMargin, doc.bottomMargin, W, doc.height, id="f")], onPage=on_page)])
    el = []

    # ---------- Title band
    title = T([[P("<font name='FB' size=17 color='white'>Slipstream Weekly Business Review</font>", body)],
               [P(f"<font size=10 color='#D7E0E8'>{esc(week)}  ·  ALL team + Leadership</font>", body)]], colWidths=[W])
    title.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), HEAD), ("LEFTPADDING", (0, 0), (-1, -1), 12),
                               ("TOPPADDING", (0, 0), (0, 0), 10), ("BOTTOMPADDING", (0, -1), (-1, -1), 9)]))
    el += [title, Spacer(1, 6),
           P("<b>The Formula:</b> Control COGS + Protect Labor Hours + Grow Sales = Better Business + Better Pay. "
             "All percentages are of Core Sales (catering and invoice payments excluded).", small), Spacer(1, 8)]

    # ---------- KPI strip (Open Cafes column of Cafe Financials)
    fin = tab("1. Cafe Financials")[0]
    fr = {r[0]: r for r in fin}
    def oc(label): return fr[label][-1] if label in fr else "–"
    kpis = [("Core Sales", oc("Total Core Sales"), f"{oc('Transactions (count)')} transactions"),
            ("Core Labor", oc("Core Labor % of Core Sales"), oc("Core Labor $") + " core wages"),
            ("Prime Cost", oc("Prime Cost % of Core Sales"), "labor + COGS; healthy 55–60%"),
            ("Tips", oc("Tip % of Core Sales"), oc("Tips") + " total")]
    kc = [[P(f"<font size=7.4 color='#5F6B76'>{k.upper()}  ·  OPEN CAFES</font><br/>"
             f"<font name='FB' size=16>{esc(v)}</font><br/><font size=7.2 color='#5F6B76'>{esc(n)}</font>", body)
           for k, v, n in kpis]]
    kt = T(kc, colWidths=[W / 4] * 4)
    kt.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, LINE), ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE),
                            ("BACKGROUND", (0, 0), (-1, -1), BAND), ("TOPPADDING", (0, 0), (-1, -1), 6),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("LEFTPADDING", (0, 0), (-1, -1), 8)]))
    el += [kt, Spacer(1, 4)]

    # ---------- Scorecard
    el.append(P("Targets Scorecard", h1))
    sc = tab("Targets Scorecard")[0]
    rows = [["Cafe", "Core Sales", "Core Labor $", "Core Labor %", "Target range", "Max", "Labor status", "Prime Cost"]]
    for r in sc[1:]:
        cafe, sales, lab, pct, rng, mx, status, prime = r[:8]
        k = "g" if "GREEN" in status else ("a" if "AMBER" in status else "r")
        lbl = {"g": "GREEN  in range", "a": "AMBER  over range", "r": "RED  over max"}[k]
        pv = money_to_num(prime.split("%")[0] + "%") if prime else None
        pk = "g" if "Healthy" in prime else ("a" if "Watch" in prime else ("r" if "Problem" in prime else
             ("g" if pv is not None and pv <= 60 else ("a" if pv is not None and pv <= 65 else "r"))))
        plbl = prime.split("%")[0] + "%  " + {"g": "Healthy", "a": "Watch", "r": "Problem"}[pk]
        rows.append([cafe, sales, lab, pct, rng.replace("–", " – "), mx, chip(lbl, k), chip(plbl, pk)])
    el.append(std_table(rows, [0.95*inch, .85*inch, .9*inch, .85*inch, 1.05*inch, .55*inch, 1.15*inch, 1.0*inch],
                        first_bold=True, extra=[("ALIGN", (6, 0), (-1, -1), "CENTER"), ("FONT", (3, 1), (3, -1), "FCB", 8.4)]))
    el.append(Spacer(1, 4))
    el.append(P("Core labor only: Manager (admin) and Training shifts are excluded. Amber = the lead's discretion; "
                "above MAX needs Ryan's OK before the schedule is published. Prime Cost (core labor + COGS): 60% or less healthy, to 65% watch, above 65% problem; 1899 Penn healthy to 65%, watch to 67%, in line with its labor range.", small))
    nm = next((t for h, s, t in paras if h.startswith("Targets") and t.startswith("NMWA")), None)
    if nm: el += [Spacer(1, 3), P(esc(clean(nm)), small)]
    hl = next((t for h, s, t in paras if h.startswith("Targets") and not t.startswith(("NMWA", "CORE"))), None)
    if hl:
        hl = clean(hl); first, _, rest = hl.partition(". ")
        box = T([[P(f"<font name='FB'>{esc(first)}.</font> {esc(rest)}", S("hl", fontSize=9, leading=12.4))]], colWidths=[W])
        box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), ACC), ("LEFTPADDING", (0, 0), (-1, -1), 9),
                                 ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                                 ("LINEBEFORE", (0, 0), (0, -1), 3, HEAD)]))
        el += [Spacer(1, 7), box]

    # ---------- Wins & watch-outs
    wins = [clean(t)[len("WIN"):].lstrip(" :–") for h, t in R["bullets"] if h.startswith("Wins") and t.startswith("WIN")]
    wats = [clean(t)[len("WATCH"):].lstrip(" :–") for h, t in R["bullets"] if h.startswith("Wins") and t.startswith("WATCH") and not is_tc(t)]
    PROPER = {'studio':'Studio','downtown':'Downtown','penn':'Penn','nmwa':'NMWA','cogs':'COGS','monday':'Monday','tuesday':'Tuesday',
              'wednesday':'Wednesday','thursday':'Thursday','friday':'Friday','saturday':'Saturday','sunday':'Sunday','open-cafe':'Open-cafe'}
    def sentence(x):
        w = x.split(); o = []
        for i, t in enumerate(w):
            lt = t.lower(); key = lt.strip(':,'); o.append((PROPER[key] + lt[len(key):]) if key in PROPER else (t if any(ch.isdigit() for ch in t) else (lt.capitalize() if i == 0 else lt)))
        return ' '.join(o)
    def blist(items, color):
        rows = []
        for it in items:
            toks = it.split(" "); n = 0
            for t in toks:
                core = t.strip(":,.()")
                if core and (core.isupper() or not any(ch.isalpha() for ch in core)) and not core.startswith("$"): n += 1
                else: break
            lead = toks[:n]
            if any(w.strip(":,") not in ("NMWA", "COGS") and any(ch.isalpha() for ch in w) for w in lead):
                txt = f"<font name='FB'>{esc(sentence(' '.join(lead)))}</font> {esc(' '.join(toks[n:]))}"
            else: txt = esc(it)
            rows.append(["", P(txt, body)])
        t = T(rows, colWidths=[4, W - 4])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), color), ("LEFTPADDING", (1, 0), (1, -1), 8),
                               ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                               ("LINEBELOW", (1, 0), (1, -2), 0.3, LINE), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                               ("LEFTPADDING", (0, 0), (0, -1), 0), ("RIGHTPADDING", (0, 0), (0, -1), 0)]))
        return t
    el += [P("Wins", h1), blist(wins, GREEN), P("Watch-outs", h1), blist(wats, AMBER)]

    # ---------- Cafe financials
    el += [P("1. Cafe Financials: Week Summary", h1),
           P("Income-statement order. All percentages are of Core Sales; catering is shown on its own line and kept out of the ratios. "
             "Open Cafes = Downtown, Studio and 1899 Penn (NMWA excluded).", small), Spacer(1, 4)]
    rows, secs, bolds = [], [], []
    for i, r in enumerate(fin):
        if not any(x for x in r[1:]) and not r[0]: continue
        if i > 0 and not any(x for x in r[1:]):
            secs.append(len(rows)); rows.append([r[0]] + [""] * (len(r) - 1)); continue
        if r[0].startswith(("Total Core Sales", "Gross Profit (", "Prime Cost (", "Total Labor")): bolds.append(len(rows))
        rows.append(r)
    rows[0][0] = ""
    el.append(std_table(rows, [2.3*inch] + [(W - 2.3*inch) / 5] * 5, section_rows=secs, bold_rows=bolds, bold_cols=[5]))

    cat = tab("3. Weekly Sales by Category")
    ops = tab("4. Ops Metrics")
    if cat:
        rows = [r for r in cat[0] if any(r)]; rows[0][0] = "Sales by category"
        b = [i for i, r in enumerate(rows) if r[0].startswith("Total")]
        el += [P("2. Weekly Sales by Category", h1), std_table(rows, [2.3*inch] + [(W - 2.3*inch) / 5] * 5, bold_rows=b, bold_cols=[5])]
    if ops:
        rows = [r for r in ops[0] if any(r) and not all(x in ("–", "") for x in r[1:])]
        el += [P("3. Ops Metrics", h1), P("Per-cafe indicators that do not add up across cafes; read each cafe on its own.", small),
               Spacer(1, 3), std_table(rows, [2.3*inch] + [(W - 2.3*inch) / 4] * 4)]

    # ---------- Filter coffee distribution
    try:
        wstart = _dt.datetime.strptime(R["week"].split(" – ")[0].replace("Week of Sunday ", "") + " " + R["week"][-4:], "%b %d %Y").date()
        # Preferred source (Nov 22 2026): Square item-sales report saved as data/filter_mix_weekly.json
        # {"YYYY-MM-DD" (Sunday): {cafe: {"cc","ds","so"}}}. Falls back to the transactions store.
        fj = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "filter_mix_weekly.json")
        FM = {}
        if os.path.exists(fj):
            FM = {_dt.date.fromisoformat(k): v for k, v in json.load(open(fj)).items()}
            FM = {k: v for k, v in FM.items() if k <= wstart}
        if wstart not in FM:
            FM = filter_mix(wstart)
    except Exception as e:
        FM = {}
    if FM.get(wstart):
        cafes = [c for c in ["Downtown", "Studio", "1899 Penn", "NMWA"] if any(c in FM[w] for w in FM)]
        cur = FM[wstart]
        def g(c, k): return cur.get(c, {}).get(k, 0)
        tot = {c: g(c, "cc") + g(c, "ds") + g(c, "so") for c in cafes}
        oc = [c for c in cafes if c != "NMWA"]
        def pctc(n, d): return f"{n / d:.0%}" if d else "–"
        rows = [["Hot filter, cups"] + cafes + ["Open Cafes"]]
        for lab, k in [("Cruise Control (house blend)", "cc"), ("Daily Selection", "ds"), ("Single origin (Bright, Fruit, Floral, Decaf)", "so")]:
            rows.append([lab] + [f"{g(c, k):,}  ({pctc(g(c, k), tot[c])})" for c in cafes] +
                        [f"{sum(g(c, k) for c in oc):,}  ({pctc(sum(g(c, k) for c in oc), sum(tot[c] for c in oc))})"])
        rows.append(["Total hot filter"] + [f"{tot[c]:,}" for c in cafes] + [f"{sum(tot[c] for c in oc):,}"])
        ncc = lambda c: (g(c, "ds") + g(c, "so"), tot[c])
        rows.append(["Beyond Cruise Control %"] + [pctc(*ncc(c)) for c in cafes] +
                    [pctc(sum(ncc(c)[0] for c in oc), sum(tot[c] for c in oc))])
        n = len(rows)
        extra = [("BACKGROUND", (1, n - 1), (-1, n - 1), GREEN_BG), ("FONT", (0, n - 1), (-1, n - 1), "FCB", 8.2)]
        wc = (W - 2.45 * inch) / (len(cafes) + 1)
        el += [P("4. Filter Coffee Distribution", h1),
               P("Goal: find the right cup for each guest. The KPI is the share of hot filter cups that go beyond our house blend "
                 "(Cruise Control) to the Daily Selection or a single origin. If nearly every guest gets Cruise Control, we are not "
                 "doing that job. Cups counted from Square's item sales report; Nitro (iced) and whole-bean or bagged retail are left out.", small),
               Spacer(1, 3),
               std_table(rows, [2.45 * inch] + [wc] * (len(cafes) + 1), bold_rows=[n - 2], bold_cols=[len(cafes) + 1], extra=extra)]
        weeks = sorted(FM)[-6:]
        trow = [["Beyond Cruise Control %"] + [f"{w.strftime('%b')} {w.day}" for w in weeks]]
        for c in oc:
            r_ = [c]
            for w in weeks:
                v = FM[w].get(c, {}); t = sum(v.values()) if v else 0
                r_.append(pctc(v.get("ds", 0) + v.get("so", 0), t) if t else "–")
            trow.append(r_)
        el += [Spacer(1, 4), P("Trend, last 6 weeks (each cafe on its own line)", h2),
               std_table(trow, [2.45 * inch] + [(W - 2.45 * inch) / len(weeks)] * len(weeks),
                         extra=[("FONT", (-1, 1), (-1, -1), "FCB", 7.8), ("BACKGROUND", (-1, 1), (-1, -1), GREEN_BG)])]

    # ---------- Daily budget vs actual
    el += [PageBreak(), P("5. Daily Plan vs Actual", h1),
           P("Core labor only. Labor plan = the day's projected sales × the cafe's target %. Labor % is colored against the cafe's band: "
             "green inside the range or better, amber over the range but under the max, red over the max. Balance to the week, not each day.", small)]
    for h, s, r in tabs:
        if not h.startswith("2. Daily"): continue
        rows, secs = [], []
        for x in r:
            if not any(x): continue
            if not any(v for v in x[1:]): secs.append(len(rows)); rows.append([x[0]] + [""] * (len(x) - 1)); continue
            rows.append(x)
        rows[0][0] = ""
        extra = []
        lo, hi, mx = BANDS.get(s, (None, None, None))
        for i, x in enumerate(rows):
            if x[0] == "Difference":
                for j, v in enumerate(x[1:], 1):
                    if v.startswith("("): extra.append(("TEXTCOLOR", (j, i), (j, i), RED))
            if x[0].startswith("Core Labor %") and lo:
                for j, v in enumerate(x[1:], 1):
                    n = money_to_num(v)
                    if n is None: continue
                    n /= 100
                    col, bg = (GREEN, GREEN_BG) if n <= hi else ((AMBER, AMBER_BG) if n <= mx else (RED, RED_BG))
                    extra += [("TEXTCOLOR", (j, i), (j, i), col), ("BACKGROUND", (j, i), (j, i), bg), ("FONT", (j, i), (j, i), "FCB", 7.8)]
        cap = s + (f"   (target {lo:.0%}, range to {hi:.0%}, max {mx:.0%})" if lo else "   (staffing rule: 1 person a day, 2 on community days)")
        el.append(KeepTogether([P(cap, h2), std_table(rows, [1.55*inch] + [(W - 1.55*inch) / 8] * 8, section_rows=secs,
                                                     bold_cols=[8], extra=extra)]))

    # ---------- Next 4 weeks + timecards
    el += [P("6. Next 4 Weeks: Scheduling Targets", h1),
           P("Schedule each week to TARGET; the range is working room; above MAX needs Ryan's OK first. Core labor excludes Manager (admin) "
             "and Training shifts, so budget those on top. Balance across the week, not each day.", small)]
    for h, s, r in tabs:
        if not h.startswith("5. Next"): continue
        lab = next((t for hh, ss, t in paras if hh.startswith("5. Next") and t.split()[0] in ("Downtown", "Studio", "1899", "NMWA")
                    and r and True and t not in [x for x in []]), None)
    labels = [clean(t) for hh, ss, t in paras if hh.startswith("5. Next") and t.split()[0] in ("Downtown", "Studio", "1899", "NMWA")]
    fwd = tab("5. Next")
    for lbl, r in zip(labels, fwd):
        rows = [x for x in r if any(x)]
        extra = []
        for i, x in enumerate(rows):
            if x[0].startswith("TARGET"): extra += [("BACKGROUND", (1, i), (-1, i), GREEN_BG), ("FONT", (0, i), (-1, i), "FCB", 7.8)]
        lbl = lbl.replace(" : core range", ": range").replace(": core range", ": range").replace(" · ", ", ")
        el.append(KeepTogether([P(lbl, h2), std_table(rows, [1.9*inch] + [(W - 1.9*inch) / 5] * 5, bold_cols=[5], extra=extra)]))
    src = next((t for hh, ss, t in paras if hh.startswith("5. Next") and t.startswith("Source")), None)
    if src: el += [Spacer(1, 3), P(esc(clean(src)), small)]

    tc = tab("5b. Timecard")
    tc_notes = [clean(t) for h, t in R["bullets"] if h.startswith("Wins") and is_tc(t)]
    write_timecards(path, week, tc[0] if tc else None, tc_notes)

    # ---------- Historical + caveats
    hist = [(s or "", r) for h, s, r in tabs if h.startswith("6. Historical")]
    hlabels = [t for hh, ss, t in paras if hh.startswith("6. Historical") and t in ("Downtown", "Studio", "1899 Penn")]
    if hist:
        el += [P("7. Is This Week Normal? Monthly Averages vs This Week", h1),
               P("From the transactions store (Jan to Mar 2026 backfill). NMWA excluded.", small), Spacer(1, 3)]
        for lbl, (_, r) in zip(hlabels, hist):
            rows = [x for x in r if any(x)]
            rows[0] = [lbl] + [c.replace(" 2026", "").split(" (")[0].replace("This Wk", "This week") for c in rows[0][1:]]
            rows = [[x[0].replace("Tip % (tips ÷ net)", "Tip % (tips ÷ net sales)")] + x[1:] for x in rows]
            el.append(KeepTogether([std_table(rows, [2.45 * inch] + [(W - 2.45 * inch) / 4] * 4,
                                              extra=[("FONT", (4, 1), (4, -1), "FCB", 7.8), ("BACKGROUND", (4, 1), (4, -1), GREEN_BG)]),
                                    Spacer(1, 6)]))
    cav = [clean(t) for h, t in R["bullets"] if h.startswith("7. Caveats")]
    if cav:
        el.append(P("Notes", h1))
        for c in cav: el.append(P("•  " + esc(c), S("cv", fontSize=7.6, leading=10, textColor=MUTE, leftIndent=8, firstLineIndent=-8, spaceAfter=2)))
    doc.build(el)
    return out

if __name__ == "__main__":
    print(build(sys.argv[1]))
