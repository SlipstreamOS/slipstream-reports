"""Populate the payroll workbook from a Square shifts CSV + per-location daily CC tips.

Maps shift rows into the 5 location Schedule Hours tabs, populates Tip Paste B14:B27
per location, sets Schedule to Tips Calc!G2 = period start, and saves a copy.
"""
from __future__ import annotations
import csv, copy, datetime as dt, shutil, sys
from collections import defaultdict
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
# v2 template includes the SRP (Slipstream Roasting) section.
TEMPLATE = ROOT / "template_with_srp.xlsx" if (ROOT / "template_with_srp.xlsx").exists() else ROOT / "template.xlsx"

# Ryan's process: paste the shifts CSV into ALL 5 location Schedule Hours tabs
# (the per-location filtering happens via cross-tab formulas in Schedule to Tips
# Calc and Tip Pay Summary (ADP)).
SCHEDULE_HOURS_TABS = [
    "CR - Schedule Hours",
    "SRP - Schedule Hours",   # Slipstream Roasting (added v2)
    "LC - Schedule Hours",
    "DT - Schedule Hours",
    "1899 - Schedule Hours",
    "NMWA - Schedule Hours",
]

# Tip Paste tab → Square location_name (matches the CSV "Location" column).
# Verified by reading $A$2 of each Schedule Hours tab in the template.
# Yes — "LC - Tip Paste" is for Studio (the cafe is in the Logan Circle area),
# and "CR - Tip Paste" is for CR Commissary (kitchen, typically $0 tips).
TIP_PASTE_TABS = {
    "CR - Tip Paste":   "CR Commissary (Food)",
    "LC - Tip Paste":   "Slipstream Studio",
    "DT - Tip Paste":   "Slipstream Downtown",
    "1899 - Tip Paste": "Slipstream 1899 Penn",
    "NMWA - Tip Paste": "Slipstream NMWA",
}

# Template's Schedule Hours data area is cols A-P (16 cols). The 26-col Square
# CSV is mapped down to these. Cols Q+ are pre-built formulas the template
# uses to compute derived values per row — DO NOT touch them.
CSV_TO_TEMPLATE_COL = [
    # (Template column letter, CSV header name)
    ("A", "Employee number"),
    ("B", "First name"),
    ("C", "Last name"),
    ("D", "Job title"),
    ("E", "Location"),
    ("F", "Clockin date"),
    ("G", "Clockin time"),
    ("H", "Clockout date"),
    ("I", "Clockout time"),
    ("J", "Unpaid break"),
    ("K", "Paid break"),
    ("L", "Regular hours"),
    ("M", "Overtime hours"),
    ("N", "Doubletime hours"),
    ("O", "Total paid hours"),
    ("P", "Hourly wage"),
]
DATA_COL_LETTERS = [c for c, _ in CSV_TO_TEMPLATE_COL]
DATA_LAST_COL = openpyxl.utils.column_index_from_string(DATA_COL_LETTERS[-1])  # 16

# Numeric columns — convert "$1.23" → 1.23, "8.23" → 8.23, etc.
NUMERIC_COLS = {
    "Unpaid break","Paid break",
    "Regular hours","Overtime hours","Doubletime hours","Total paid hours",
    "Hourly wage",
}
# Date columns — coerce "M/D/YY" strings to datetime objects so Excel can
# compare them to date-typed cells (e.g., Schedule to Tips Calc!J8).
DATE_COLS = {"Clockin date", "Clockout date", "Break start date", "Break end date"}

# Default range — actual paste anchor is auto-detected per tab in populate()
# because tabs vary: CR starts row 26, LC row 28, DT/1899/NMWA row 27.
DATA_END_ROW = 351

def _find_data_start(sh) -> int:
    """Return the 1-indexed row immediately below the 'Employee number' header."""
    for r in range(1, 35):
        if sh.cell(row=r, column=1).value == "Employee number":
            return r + 1
    raise RuntimeError(f"Couldn't locate 'Employee number' header in {sh.title}")

def _coerce(col, val):
    if val is None or val == "":
        return None
    if col in NUMERIC_COLS:
        s = str(val).replace("$", "").replace(",", "").strip()
        try:
            return float(s) if s else None
        except ValueError:
            return None
    if col in DATE_COLS:
        # Parse "M/D/YY" → datetime so Excel matches against date-typed refs.
        try:
            s = str(val).strip()
            if not s: return None
            m, d, y = s.split("/")
            year = int(y) + 2000 if int(y) < 100 else int(y)
            return dt.datetime(year, int(m), int(d))
        except Exception:
            return str(val)
    return str(val)

def populate(csv_path: Path, tips_by_loc_day: dict, period_start: dt.date, out_path: Path):
    """Render workbook for the period.

    Args:
      csv_path: Square shifts export CSV (with header row).
      tips_by_loc_day: {location_name: {date_str: float}} — daily CC tip totals.
      period_start: First day of the 14-day pay period (Sunday).
      out_path: Output xlsx path.
    """
    # Copy template → out_path so we never mutate the master template.
    shutil.copy(TEMPLATE, out_path)
    wb = openpyxl.load_workbook(out_path)

    # 1. Read CSV
    rows = []
    with open(csv_path, newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows.append(r)
    # Capacity check is now per-tab (LC has the smallest capacity since it starts
    # at row 28). Validated inside the per-tab loop below.

    # 2. Populate the five Schedule Hours tabs (full CSV in each).
    # Each tab has a different paste anchor (CR=26, LC=28, DT/1899/NMWA=27).
    for tab in SCHEDULE_HOURS_TABS:
        sh = wb[tab]
        data_start = _find_data_start(sh)
        # Clear ONLY data columns A-P for the data range; keep formula cols Q-AD.
        for r in range(data_start, DATA_END_ROW + 1):
            for c in range(1, DATA_LAST_COL + 1):
                sh.cell(row=r, column=c).value = None
        # Write each CSV row into A..P starting at this tab's data anchor.
        for i, row in enumerate(rows):
            r = data_start + i
            if r > DATA_END_ROW:
                raise ValueError(f"{tab}: CSV has more rows than tab's capacity at this anchor")
            for col_letter, csv_col in CSV_TO_TEMPLATE_COL:
                col_idx = openpyxl.utils.column_index_from_string(col_letter)
                sh.cell(row=r, column=col_idx).value = _coerce(csv_col, row.get(csv_col, ""))

    # 3. Set Schedule to Tips Calc!G2 = period_start
    wb["Schedule to Tips Calc"]["G2"].value = period_start

    # 4. Populate Tip Paste tabs B14:B27 with daily CC tips
    # The 14 days are period_start..period_start+13. Format the CSV uses: M/D/YY (e.g. "4/5/26")
    days = [period_start + dt.timedelta(days=i) for i in range(14)]
    for tab, loc_name in TIP_PASTE_TABS.items():
        sh = wb[tab]
        loc_tips = tips_by_loc_day.get(loc_name, {})
        for i, d in enumerate(days):
            row = 14 + i  # B14..B27
            csv_date = f"{d.month}/{d.day}/{d.year % 100}"
            tip = loc_tips.get(csv_date, 0.0)
            sh.cell(row=row, column=2).value = float(tip) if tip else 0

    # 5. Save
    wb.save(out_path)
    return out_path


def aggregate_tips_from_csv(csv_path: Path) -> dict:
    """Build {location_name: {date_str: cc_tips}} from a CSV's
    'Direct and pooled by transaction tips' column."""
    out = defaultdict(lambda: defaultdict(float))
    with open(csv_path, newline="") as f:
        for r in csv.DictReader(f):
            loc = r["Location"]
            date = r["Clockin date"]
            cc = r.get("Direct and pooled by transaction tips", "").replace("$","").replace(",","").strip()
            try:
                out[loc][date] += float(cc) if cc else 0
            except ValueError:
                pass
    return {k: dict(v) for k, v in out.items()}


def validate(out_path: Path) -> dict:
    """Re-open the populated workbook with computed values and check Tip Pay
    Summary (ADP) validation cells.

    Returns dict with keys: hours_zero_check (5 cells must all = 0),
    sub_min_wage (rows in any of N/AD/AS/BH/BW with value < 18.0)."""
    # openpyxl doesn't compute formulas — need a real Excel engine. Use
    # libreoffice headless to recalc and save.
    import subprocess
    recalc = out_path.parent / f"{out_path.stem}_recalc.xlsx"
    shutil.copy(out_path, recalc)
    # libreoffice --convert-to recomputes formulas and writes a fresh xlsx.
    # Output to a different directory to avoid overwriting input.
    recalc_dir = Path("/tmp/payroll_recalc")
    recalc_dir.mkdir(exist_ok=True)
    for f in recalc_dir.glob("*"):
        try: f.unlink()
        except Exception: pass
    subprocess.run(
        ["libreoffice", "--headless", "--calc",
         "--convert-to", "xlsx", "--outdir", str(recalc_dir),
         str(out_path)],
        check=True, capture_output=True, timeout=90,
    )
    recalc_file = recalc_dir / out_path.name
    wb = openpyxl.load_workbook(recalc_file, data_only=True)
    sh = wb["Tip Pay Summary (ADP)"]

    # Validation cells (per-section "expected sum X minus per-employee aggregate")
    # Layout (after SRP insert): CR=E, SRP=R, LC=AE, DT=AU, 1899=BJ, NMWA=BY
    hours = {c: sh[c].value for c in ["E361","R361","AE361","AU361","BJ361","BY361"]}
    # Min wage check: per-employee Total Hourly across each location's section.
    # 6 sections after SRP insert: CR=N, SRP=AA, LC=AR, DT=BG, 1899=BV, NMWA=CK.
    # Each employee has 5 rows (data rows 14, 19, 24, ... — every 5th row).
    sub_min = []
    LOC_HEADERS = {"N":"CR","AA":"SRP","AR":"LC (Studio)","BG":"DT","BV":"1899","CK":"NMWA"}
    for col_letter in list(LOC_HEADERS.keys()):
        col_idx = openpyxl.utils.column_index_from_string(col_letter)
        # Iterate every 5th row from 14 onwards (employee header row)
        for r in range(14, 361, 5):
            v = sh.cell(row=r, column=col_idx).value
            name = sh.cell(row=r, column=1).value  # col A: Staff Name
            if isinstance(v, (int, float)) and 0 < v < 18.0 and name not in (None, "[New]"):
                sub_min.append({
                    "row": r, "section": LOC_HEADERS[col_letter],
                    "name": name,
                    "value": round(float(v), 2),
                })
    return {"hours_zero": hours, "sub_min_wage": sub_min}


if __name__ == "__main__":
    csv_path = Path(sys.argv[1])
    period_start = dt.date.fromisoformat(sys.argv[2])  # YYYY-MM-DD
    out_path = Path(sys.argv[3])
    tips = aggregate_tips_from_csv(csv_path)
    populate(csv_path, tips, period_start, out_path)
    print(f"Populated → {out_path}")
    result = validate(out_path)
    print(f"Hours-zero check: {result['hours_zero']}")
    if result["sub_min_wage"]:
        print(f"Sub-minimum-wage entries: {len(result['sub_min_wage'])}")
        for e in result["sub_min_wage"][:10]:
            print(f"  row {e['row']} col {e['col']}: ${e['value']}")
    else:
        print("Sub-minimum-wage check: PASS (no entries < $18)")
