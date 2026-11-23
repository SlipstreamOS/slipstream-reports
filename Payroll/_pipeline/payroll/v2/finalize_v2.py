"""Finalize a populated v2 workbook: generate the ADP Timesheet Import.

Reads the recalculated Calc tab, builds ##GENERIC## V2.0 import detail rows,
writes them into the workbook's "ADP Import" tab (static values), and exports
a ready-to-import .csv.

Usage:
  python3 finalize_v2.py <payroll_v2.xlsx> <period_start YYYY-MM-DD> <out_csv>
"""
from __future__ import annotations
import csv, datetime as dt, shutil, subprocess, sys
from pathlib import Path
import openpyxl

# ADP import constants (match build_template_v2.py)
ADP_IID       = "26642723"   # confirmed working IID (21684235 was rejected)
ADP_PAY_FREQ  = "B"
ADP_DEPT_FOH  = "001"
ADP_RATE_BASE  = "BASE"      # Rate 1 = $12 tipped minimum
ADP_RATE_REG   = "RATE_2"    # Rate 2 = employee's regular hourly wage
ADP_RATE_EVENT = "RATE_3"    # Rate 3 = Wage 3 (Event work) — confirm in ADP
SEPARATE_CHK  = "0"
JOB_CODE      = ""   # blank — FOH is captured by department

# Earnings codes paid at the REGULAR rate (Rate 2). Everything else uses BASE.
REGULAR_RATE_CODES = {"REG", "OVT", "VAC", "HOL", "SPY", "PERS"}

# Calc role → earnings codes. Each tuple: (calc_col, earnings_code, kind, rate_code)
# kind "H" = hours (Pay Hours), "D" = dollars (Dollars).
TIPPED_COMPONENTS = [
    ("G", "TIPHOURS", "H", ADP_RATE_BASE),   # Tipped Hrs  → $12
    ("H", "OTTIP",    "H", ADP_RATE_BASE),   # Tip OT Hrs  → $12
    ("I", "CREDTIPP", "D", ADP_RATE_BASE),   # CC Tips     (dollar — rate n/a)
    ("J", "OTH",      "D", ADP_RATE_BASE),   # Misc/Extra  (dollar — rate n/a)
]
NONTIPPED_COMPONENTS = [
    ("E", "REG", "H", ADP_RATE_REG),         # Reg Hrs → Rate 2 (real wage)
    ("F", "OVT", "H", ADP_RATE_REG),         # OT Hrs  → Rate 2
]
# Event role → Wage 3 (RATE_3 in ADP). Per Ryan, Wage 3 IS imported via the
# normal import; just surface which employees have Event hours so Ryan can
# confirm Wage 3 is configured for each of them in ADP before uploading.
EVENT_COMPONENTS = [
    ("E", "REG", "H", ADP_RATE_EVENT),       # Event reg hrs → Wage 3
    ("F", "OVT", "H", ADP_RATE_EVENT),       # Event OT hrs  → Wage 3
]

COL = {c: i for i, c in enumerate("ABCDEFGHIJK", start=1)}


def recalc(path: Path) -> Path:
    """Recalculate the workbook via LibreOffice; return path to recalc'd copy."""
    recalc_dir = Path("/tmp/v2_finalize")
    recalc_dir.mkdir(exist_ok=True)
    for f in recalc_dir.glob("*"):
        try: f.unlink()
        except Exception: pass
    subprocess.run(
        ["libreoffice", "--headless", "--calc", "--convert-to", "xlsx",
         "--outdir", str(recalc_dir), str(path)],
        check=True, capture_output=True, timeout=90,
    )
    return recalc_dir / path.name


def build_import_rows(calc_ws, roster_ws, period_start: dt.date):
    """Return list of 12-field import detail rows from the Calc tab."""
    # Roster: ADP Name (col C) → Employee Code (col G), Rate 2 flag (col H)
    code_by_name = {}
    rate2_by_name = {}
    for r in range(4, roster_ws.max_row + 1):
        name = roster_ws.cell(row=r, column=3).value
        code = roster_ws.cell(row=r, column=7).value
        rate2 = roster_ws.cell(row=r, column=8).value
        if name:
            code_by_name[str(name).strip()] = str(code).strip() if code else ""
            rate2_by_name[str(name).strip()] = str(rate2).strip() if rate2 else "?"

    start = period_start.strftime("%m/%d/%Y")
    end = (period_start + dt.timedelta(days=13)).strftime("%m/%d/%Y")

    rows = []
    missing_codes = []
    event_hours = []  # (name, hrs) — surface so Ryan can confirm Wage 3 is
                      #              configured in ADP for each before upload
    rate2_risk = []   # (name, ecode, hours_or_dollars) — RATE_2 rows for
                      #              employees whose Roster flag isn't "Y"
    for r in range(4, calc_ws.max_row + 1):
        name = calc_ws.cell(row=r, column=COL["A"]).value
        if not name:
            continue
        role = calc_ws.cell(row=r, column=COL["B"]).value
        tipped = calc_ws.cell(row=r, column=COL["C"]).value
        emp_code = code_by_name.get(str(name).strip(), "")
        if not emp_code:
            missing_codes.append(str(name))

        # Pick the right earnings-code/rate-code mapping per role.
        if role == "Event":
            components = EVENT_COMPONENTS
            try:
                hrs = float(calc_ws.cell(row=r, column=COL["E"]).value or 0) \
                    + float(calc_ws.cell(row=r, column=COL["F"]).value or 0)
            except (TypeError, ValueError):
                hrs = 0.0
            if hrs > 0.005:
                event_hours.append((str(name), round(hrs, 2)))
        else:
            components = TIPPED_COMPONENTS if tipped == "Y" else NONTIPPED_COMPONENTS
        for col_letter, ecode, kind, rate_code in components:
            val = calc_ws.cell(row=r, column=COL[col_letter]).value
            try:
                num = float(val) if val not in (None, "") else 0.0
            except (TypeError, ValueError):
                num = 0.0
            if abs(num) < 0.005:
                continue  # skip zero rows
            pay_hours = f"{num:.2f}" if kind == "H" else ""
            dollars   = f"{num:.2f}" if kind == "D" else ""
            rows.append([
                ADP_IID, ADP_PAY_FREQ, start, end, emp_code, ecode,
                pay_hours, dollars, SEPARATE_CHK, ADP_DEPT_FOH,
                rate_code, JOB_CODE,
            ])
            # RATE_2 safety: flag if this row uses RATE_2 and the employee
            # isn't confirmed set up in ADP. ADP silently drops these.
            if rate_code == ADP_RATE_REG and rate2_by_name.get(str(name).strip()) != "Y":
                rate2_risk.append((str(name), ecode, pay_hours or f"${dollars}"))
    return rows, sorted(set(missing_codes)), event_hours, rate2_risk


def build_salary_rows(salary_ws, period_start: dt.date):
    """Salary tab → REG import rows carrying both hours and dollars."""
    start = period_start.strftime("%m/%d/%Y")
    end = (period_start + dt.timedelta(days=13)).strftime("%m/%d/%Y")
    rows = []
    for r in range(4, salary_ws.max_row + 1):
        code = salary_ws.cell(row=r, column=2).value
        if not code:
            continue
        dept = salary_ws.cell(row=r, column=3).value or ADP_DEPT_FOH
        hrs  = salary_ws.cell(row=r, column=5).value
        pay  = salary_ws.cell(row=r, column=6).value  # computed Period Pay
        try:
            hrs_f = float(hrs) if hrs not in (None, "") else 0.0
            pay_f = float(pay) if pay not in (None, "") else 0.0
        except (TypeError, ValueError):
            continue
        # Salary rows carry an explicit Dollars amount, so the rate code is
        # irrelevant — use BASE (salaried employees have no Rate 2).
        rows.append([
            ADP_IID, ADP_PAY_FREQ, start, end, str(code).strip(), "REG",
            f"{hrs_f:.2f}", f"{pay_f:.2f}", SEPARATE_CHK, str(dept).strip(),
            ADP_RATE_BASE, JOB_CODE,
        ])
    return rows


def build_adjustment_rows(adj_ws, roster_ws, period_start: dt.date):
    """Adjustments tab → one-off import rows (PTO, extra misc, bonuses)."""
    start = period_start.strftime("%m/%d/%Y")
    end = (period_start + dt.timedelta(days=13)).strftime("%m/%d/%Y")
    # Roster ADP Name → Employee Code, Rate 2 flag
    code_by_name = {}
    rate2_by_name = {}
    for r in range(4, roster_ws.max_row + 1):
        name = roster_ws.cell(row=r, column=3).value
        code = roster_ws.cell(row=r, column=7).value
        rate2 = roster_ws.cell(row=r, column=8).value
        if name:
            code_by_name[str(name).strip()] = str(code).strip() if code else ""
            rate2_by_name[str(name).strip()] = str(rate2).strip() if rate2 else "?"

    rows, missing, rate2_risk = [], [], []
    for r in range(5, adj_ws.max_row + 1):
        name = adj_ws.cell(row=r, column=1).value
        ecode = adj_ws.cell(row=r, column=2).value
        if not name or not ecode:
            continue
        hrs = adj_ws.cell(row=r, column=3).value
        dol = adj_ws.cell(row=r, column=4).value
        dept = adj_ws.cell(row=r, column=5).value or ADP_DEPT_FOH
        emp_code = code_by_name.get(str(name).strip(), "")
        if not emp_code:
            missing.append(str(name))
        pay_hours = f"{float(hrs):.2f}" if hrs not in (None, "", 0) else ""
        dollars   = f"{float(dol):.2f}" if dol not in (None, "", 0) else ""
        ec = str(ecode).strip()
        rate_code = ADP_RATE_REG if ec.upper() in REGULAR_RATE_CODES else ADP_RATE_BASE
        rows.append([
            ADP_IID, ADP_PAY_FREQ, start, end, emp_code, ec,
            pay_hours, dollars, SEPARATE_CHK, str(dept).strip(),
            rate_code, JOB_CODE,
        ])
        # RATE_2 safety: flag if Adjustments row lands on RATE_2 and employee
        # isn't confirmed set up.
        if rate_code == ADP_RATE_REG and rate2_by_name.get(str(name).strip()) != "Y":
            rate2_risk.append((str(name), ec, pay_hours or f"${dollars}"))
    return rows, sorted(set(missing)), rate2_risk


def main(wb_path: Path, period_start: dt.date, out_csv: Path):
    recalc_path = recalc(wb_path)
    rc = openpyxl.load_workbook(recalc_path, data_only=True)
    rows, missing, event_hours, rate2_risk = build_import_rows(rc["Calc"], rc["Roster"], period_start)
    # Append salaried workers and one-off adjustments
    salary_rows = build_salary_rows(rc["Salary"], period_start) if "Salary" in rc.sheetnames else []
    if "Adjustments" in rc.sheetnames:
        adj_rows, adj_missing, adj_rate2_risk = build_adjustment_rows(
            rc["Adjustments"], rc["Roster"], period_start)
    else:
        adj_rows, adj_missing, adj_rate2_risk = [], [], []
    rows = rows + salary_rows + adj_rows
    missing = sorted(set(missing) | set(adj_missing))
    rate2_risk = rate2_risk + adj_rate2_risk

    # 1. Write detail rows into the workbook's ADP Import tab (static values)
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["ADP Import"]
    # unmerge everything below the header rows (the A4 note is merged)
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 3:
            ws.unmerge_cells(str(mr))
    # clear any prior detail rows / the note
    for r in range(3, ws.max_row + 1):
        for c in range(1, 13):
            ws.cell(row=r, column=c).value = None
    for i, row in enumerate(rows):
        for c, val in enumerate(row, start=1):
            ws.cell(row=3 + i, column=c).value = val
    wb.save(wb_path)

    # 2. Export standalone .csv (##GENERIC## V2.0 + header + detail)
    # Build the CSV to match the ADP template byte-for-byte:
    #   - row 1 = "##GENERIC## V2.0" + 11 empty fields (12 columns)
    #   - CRLF line endings BETWEEN rows
    #   - NO trailing newline after the final row
    import io
    all_lines = [["##GENERIC## V2.0"] + [""] * 11,
                 ["IID","Pay Frequency","Pay Period Start","Pay Period End",
                  "Employee Id","Earnings Code","Pay Hours","Dollars",
                  "Separate Check","Department Number","Rate Code","Job Code"]]
    all_lines.extend(rows)
    buf = io.StringIO()
    writer = csv.writer(buf)
    for ln in all_lines:
        writer.writerow(ln)
    text = buf.getvalue().rstrip("\r\n")  # drop the trailing CRLF
    with open(out_csv, "w", newline="") as f:
        f.write(text)

    print(f"Wrote {len(rows)} import rows")
    print(f"  → ADP Import tab in {wb_path.name}")
    print(f"  → {out_csv}")
    if missing:
        print(f"  WARNING: no employee code for: {', '.join(missing)}")
    if event_hours:
        total = round(sum(h for _, h in event_hours), 2)
        print(f"  FLAG: {total} Event hours INCLUDED in import at rate code {ADP_RATE_EVENT} (Wage 3).")
        print(f"  Confirm Wage 3 is configured in ADP for each of these employees before uploading:")
        for nm, hrs in event_hours:
            print(f"    - {nm}: {hrs} hrs")
    if rate2_risk:
        # De-dupe (name, ecode) pairs while preserving order
        seen = set(); dedup = []
        for entry in rate2_risk:
            key = (entry[0], entry[1])
            if key in seen: continue
            seen.add(key); dedup.append(entry)
        print(f"  FLAG: {len(dedup)} {ADP_RATE_REG} row(s) for employees NOT confirmed set up in ADP.")
        print(f"  ADP will silently DROP these rows unless Rate 2 is configured. Set up Rate 2 in")
        print(f"  ADP for each, then flip Roster col H to 'Y'. Affected rows:")
        for nm, ec, amt in dedup:
            print(f"    - {nm}: {ec} {amt}")


if __name__ == "__main__":
    wb_path = Path(sys.argv[1])
    period_start = dt.date.fromisoformat(sys.argv[2])
    out_csv = Path(sys.argv[3])
    main(wb_path, period_start, out_csv)
