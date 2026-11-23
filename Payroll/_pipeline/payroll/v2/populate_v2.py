"""Populate the v2 payroll template from a Square shifts CSV + tips JSON.

Writes to the Inputs tab only:
  - B2 = period start
  - C9:F22 = daily CC tips per location (Studio, Downtown, 1899, NMWA)
  - A41:P700 = shifts data (16 cols A-P; helper Q/R are formulas)
"""
from __future__ import annotations
import csv, datetime as dt, json, shutil, sys
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "template_v2.xlsx"

CSV_TO_TEMPLATE_COL = [
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

NUMERIC_COLS = {"Unpaid break","Paid break","Regular hours","Overtime hours",
                "Doubletime hours","Total paid hours","Hourly wage"}
DATE_COLS = {"Clockin date","Clockout date"}

# Tipped locations in the order columns C-F on Inputs
TIPPED_LOC_ORDER = [
    "Slipstream Studio",
    "Slipstream Downtown",
    "Slipstream 1899 Penn",
    "Slipstream NMWA",
]


def _coerce(col, val):
    if val is None or val == "":
        return None
    if col in NUMERIC_COLS:
        s = str(val).replace("$","").replace(",","").strip()
        try:
            return float(s) if s else None
        except ValueError:
            return None
    if col in DATE_COLS:
        try:
            s = str(val).strip()
            if not s: return None
            m, d, y = s.split("/")
            year = int(y) + 2000 if int(y) < 100 else int(y)
            return dt.datetime(year, int(m), int(d))
        except Exception:
            return str(val)
    return str(val)


def apply_weekly_ot(ws, period_start: dt.date) -> bool:
    """Split each shift's Total Paid Hours (col O) into Regular (col L) and
    Overtime (col M) using a Sunday-Saturday workweek: any hours over 40 in a
    workweek are OT. Shifts are walked chronologically within each week.

    Only runs when the OT column is empty/zero across all rows — i.e. the data
    was synthesized from raw timecards. A real Square export that already has
    OT split out is left untouched. Returns True if OT was applied.
    """
    import re
    from collections import defaultdict
    rows = []
    for r in range(41, 701):
        last = ws.cell(row=r, column=3).value
        if not last:
            continue
        tot = ws.cell(row=r, column=15).value  # O — total paid hours
        if tot in (None, ""):
            continue
        rows.append((r, ws.cell(row=r, column=2).value, last,
                     ws.cell(row=r, column=6).value,   # F clockin date
                     ws.cell(row=r, column=7).value,   # G clockin time
                     float(tot)))
    if not rows:
        return False
    # Bail if OT is already populated (real Square export — trust it).
    if any((ws.cell(row=r, column=13).value or 0) for r, *_ in rows):
        return False

    def tsec(s):
        m = re.match(r"\s*(\d+):(\d+):?(\d*)\s*([AaPp][Mm])", str(s or ""))
        if not m:
            return 0
        h, mi, se = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        ap = m.group(4).upper()
        if ap == "PM" and h != 12:
            h += 12
        if ap == "AM" and h == 12:
            h = 0
        return h * 3600 + mi * 60 + se

    def week_of(cd):
        d = cd.date() if hasattr(cd, "date") else cd
        return (d - period_start).days // 7

    groups = defaultdict(list)
    for rec in rows:
        groups[(rec[1], rec[2], week_of(rec[3]))].append(rec)
    for recs in groups.values():
        recs.sort(key=lambda x: ((x[3].date() if hasattr(x[3], "date") else x[3]),
                                  tsec(x[4])))
        cum = 0.0
        for r, first, last, cd, ct, tot in recs:
            if cum >= 40:
                ot = tot
            elif cum + tot <= 40:
                ot = 0.0
            else:
                ot = cum + tot - 40
            cum += tot
            ot = round(ot, 2)
            ws.cell(row=r, column=12).value = round(tot - ot, 2)  # L regular
            ws.cell(row=r, column=13).value = ot                  # M overtime
    return True


def populate(csv_path: Path, tips: dict, period_start: dt.date,
             out_path: Path,
             exclude_employees: set | None = None):
    # Salaried employees — paid via the Salary tab, never via hourly shifts.
    exclude_employees = exclude_employees or {"Ines Chavez", "Mo Koolphanich"}
    shutil.copy(TEMPLATE, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb["Inputs"]

    # 1. Period start
    ws["B2"].value = period_start

    # 2. Daily CC tips (rows 9-22, cols C-F)
    days = [period_start + dt.timedelta(days=i) for i in range(14)]
    for i, d in enumerate(days):
        date_str = f"{d.month}/{d.day}/{d.year % 100}"
        for col_idx, loc in enumerate(TIPPED_LOC_ORDER, start=3):  # C=3
            loc_tips = tips.get(loc, {})
            tip_val = loc_tips.get(date_str, 0)
            ws.cell(row=9 + i, column=col_idx).value = float(tip_val) if tip_val else 0

    # 3. Shifts CSV — paste at row 41
    rows = []
    with open(csv_path, newline="") as f:
        for r in csv.DictReader(f):
            full = f"{r.get('First name','').strip()} {r.get('Last name','').strip()}".strip()
            if full in exclude_employees:
                continue
            rows.append(r)
    if len(rows) > 660:  # capacity = 700 - 41 + 1
        raise ValueError(f"CSV has {len(rows)} rows, capacity is 660")
    for i, row in enumerate(rows):
        r = 41 + i
        for col_letter, csv_col in CSV_TO_TEMPLATE_COL:
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            ws.cell(row=r, column=col_idx).value = _coerce(csv_col, row.get(csv_col, ""))

    # 4. Split Regular/Overtime hours by Sun-Sat workweek if OT wasn't
    #    already provided (synthesized data has all hours in the Reg column).
    apply_weekly_ot(ws, period_start)

    wb.save(out_path)
    return out_path


def aggregate_tips_from_json(path: Path) -> dict:
    return json.loads(path.read_text())


if __name__ == "__main__":
    csv_path = Path(sys.argv[1])
    tips_path = Path(sys.argv[2])  # cc_tips.json
    period_start = dt.date.fromisoformat(sys.argv[3])
    out_path = Path(sys.argv[4])
    tips = aggregate_tips_from_json(tips_path)
    populate(csv_path, tips, period_start, out_path)
    print(f"Populated → {out_path}")
