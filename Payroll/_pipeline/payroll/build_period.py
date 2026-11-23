"""End-to-end orchestration: from cached Square + Asana data, produce the
period's deliverables (wage audit, timecard review, populated workbook).

Usage:
  python3 build_period.py <period_id> [<input_csv_or_blank>]

If no CSV is passed, synthesizes one from cached Square timecards (live mode).
"""
from __future__ import annotations
import csv, datetime as dt, json, sys
from pathlib import Path

import run_pull
import populator

ROOT = Path(__file__).resolve().parent
CACHE = ROOT.parent / "cache"

LOC_NAME = {
    "LXM7V1QWDHKZJ": "Slipstream Downtown",
    "LGVKCKX9B9DZT": "Slipstream Studio",
    "L4ZBMGCWJ5C7R": "Slipstream 1899 Penn",
    "LV2D18H0DD9ES": "Slipstream NMWA",
    "LZCA3EY4GB7BM": "CR Commissary (Food)",
    "LW5RCZWPJEDWD": "LC Commissary (Beverage)",
    "LFXQF0X8CKYF5": "Slipstream Roasting",
}

def synthesize_csv(period_id: str, out_csv: Path, wage_overrides: dict | None = None,
                   exclude_employees: set | None = None):
    """Render cached Square timecards as a CSV matching Ryan's Square Shifts export.

    Output has the full 26-column shape (Employee number..Declared cash tips), but
    cols Y/Z (per-shift tip allocation) are left blank — the Tip Paste tabs use
    daily totals from Square Sales Summary, not per-shift allocation.

    Args:
      wage_overrides: {full_name: hourly_rate_dollars} — overrides Square's recorded
        wage on the timecard. Used when rates were updated mid-period or to apply
        Ryan's "what should they be paid" decisions.
      exclude_employees: set of full names to drop from the CSV entirely (e.g.,
        salaried employees who don't clock in).
    """
    wage_overrides = wage_overrides or {}
    exclude_employees = exclude_employees or set()
    timecards, _ = run_pull.load_period(period_id)
    sq_by_id, _ = run_pull.load_square_team()
    jobs = run_pull.load_square_jobs()

    headers = [
        "Employee number","First name","Last name","Job title","Location",
        "Clockin date","Clockin time","Clockout date","Clockout time",
        "Break start date","Break start time","Break end date","Break end time",
        "Unpaid break","Paid break",
        "Regular hours","Overtime hours","Doubletime hours","Total paid hours",
        "Hourly wage",
        "Regular labor cost","Overtime labor cost","Doubletime labor cost","Total labor cost",
        "Direct and pooled by transaction tips","Declared cash tips",
    ]

    def fmt_date(d): return f"{d.month}/{d.day}/{d.year % 100}"
    def fmt_time(d): return d.strftime("%-I:%M:%S %p ") + d.tzname().replace("EST", "EST").replace("EDT","EDT")

    rows = []
    for tc in timecards:
        if not tc.get("start_at") or not tc.get("end_at"):
            continue
        start = run_pull.parse_iso(tc["start_at"])
        end = run_pull.parse_iso(tc["end_at"])
        # Calendar tz should be ET. Add a manual tz tag matching Square export style:
        tz = "EDT" if start.utcoffset() and start.utcoffset().total_seconds() == -4*3600 else "EST"
        tm = sq_by_id.get(tc.get("team_member_id"), {})
        full_name = tm.get("full", "")
        if full_name in exclude_employees:
            continue
        wage = tc.get("wage", {}) or {}
        hr_cents = (wage.get("hourly_rate") or {}).get("amount") or 0
        hr = hr_cents / 100.0
        # If the timecard's wage is $0 (e.g., Manager job has pay_type=NONE in
        # Square), fall back to the employee's max hourly rate from any other
        # job assignment. This catches Colin's Manager shifts which Square
        # records as $0 but should be paid at his standard rate.
        if hr == 0:
            other_rates = [
                (ja.get("hourly_rate") or {}).get("amount", 0) / 100.0
                for ja in tm.get("wage_setting", {}).get("job_assignments", [])
                if (ja.get("hourly_rate") or {}).get("amount")
            ]
            if other_rates:
                hr = max(other_rates)
        # Apply wage overrides — Ryan-decided rate for the period
        if full_name in wage_overrides:
            hr = float(wage_overrides[full_name])
        total_hrs = round((end - start).total_seconds() / 3600, 2)
        # Square doesn't return reg/OT split in Timecard directly — approximate as all-regular
        # for sub-40hr weeks. (For payroll math the Total paid hours column is what feeds the
        # workbook's Hours.)
        reg_hrs = total_hrs
        ot_hrs = 0
        dt_hrs = 0
        loc_disp = LOC_NAME.get(tc.get("location_id"), tc.get("location_id"))
        # Job title from wage.title, falls back to jobs map
        title = wage.get("title") or jobs.get(wage.get("job_id"), "")
        cash_tip = ((tc.get("declared_cash_tip_money") or {}).get("amount") or 0) / 100.0

        rows.append({
            "Employee number": "",
            "First name": tm.get("first", ""),
            "Last name":  tm.get("last", ""),
            "Job title":  title,
            "Location":   loc_disp,
            "Clockin date":  fmt_date(start),
            "Clockin time":  start.strftime(f"%-I:%M:%S %p {tz}"),
            "Clockout date": fmt_date(end),
            "Clockout time": end.strftime(f"%-I:%M:%S %p {tz}"),
            "Break start date": "",
            "Break start time": "",
            "Break end date":   "",
            "Break end time":   "",
            "Unpaid break": 0,
            "Paid break":   0,
            "Regular hours":    reg_hrs,
            "Overtime hours":   ot_hrs,
            "Doubletime hours": dt_hrs,
            "Total paid hours": total_hrs,
            "Hourly wage":     f"${hr:.2f}",
            "Regular labor cost":    f"${reg_hrs * hr:.2f}",
            "Overtime labor cost":   "$0.00",
            "Doubletime labor cost": "$0.00",
            "Total labor cost":      f"${total_hrs * hr:.2f}",
            "Direct and pooled by transaction tips": "",  # not exposed via Labor API
            "Declared cash tips": f"${cash_tip:.2f}",
        })
    # Sort by (Last, First, Clockin date, Clockin time) — matches Ryan's CSV order
    def date_key(d):
        m, day, y = (int(p) for p in d.split("/"))
        return (y, m, day)
    rows.sort(key=lambda r: (r["Last name"], r["First name"], date_key(r["Clockin date"]), r["Clockin time"]))
    with open(out_csv, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, quoting=csv.QUOTE_ALL)
        w.writeheader()
        for r in rows:
            w.writerow(r)
    return out_csv, len(rows)


# Per-period wage overrides — Ryan's source-of-truth for what to actually pay
# people this period (overrides the rate Square recorded on each timecard).
# Keys are full names as Square reports them (given_name + " " + family_name).
PERIOD_WAGE_OVERRIDES = {
    "2026-04-19_2026-05-02": {
        "Amy Romero":              20.75,
        "Kathy Guevara Orellana":  18.50,
        "Sophia Brodnax":          19.00,
        "Delmi Menjivar":          18.50,
        "Josias Michalko":         19.00,
        "Opal Payne":              20.50,
        "Aryana Sandroni":         18.00,
        "Marleni Nolasco":         22.50,
    },
}
EXCLUDE_FROM_CSV = {"Ines Chavez"}  # salary, doesn't clock


def main(period_id: str, input_csv: Path | None):
    out = ROOT / period_id
    out.mkdir(parents=True, exist_ok=True)

    # 1. Wage audit + timecard review (both periods, regardless of CSV)
    print(f"[1/3] Running wage audit + timecard review for {period_id}...")
    run_pull.main(period_id)

    # 2. Build CSV (synthesize if not provided)
    if input_csv is None or not input_csv.exists():
        synth_path = out / f"shifts_synthesized.csv"
        synthesize_csv(period_id, synth_path,
                       wage_overrides=PERIOD_WAGE_OVERRIDES.get(period_id, {}),
                       exclude_employees=EXCLUDE_FROM_CSV)
        input_csv = synth_path
        print(f"[2/3] Synthesized CSV from Square timecards → {synth_path.name}")
    else:
        print(f"[2/3] Using provided CSV: {input_csv.name}")

    # 3. Populate workbook with Square-derived tips
    tips_path = CACHE / "periods" / period_id / "cc_tips.json"
    if tips_path.exists():
        tips = json.loads(tips_path.read_text())
    else:
        tips = populator.aggregate_tips_from_csv(input_csv)
        print(f"  (no cc_tips.json; falling back to CSV-derived tips)")
    period_start = dt.date.fromisoformat(period_id.split("_")[0])
    out_wb = out / f"payroll_{period_id}.xlsx"
    populator.populate(input_csv, tips, period_start, out_wb)
    print(f"[3/3] Populated workbook → {out_wb.name}")

    # 4. Validate
    result = populator.validate(out_wb)
    print(f"\nValidation:")
    print(f"  Hours-zero check: {result['hours_zero']}")
    if result["sub_min_wage"]:
        print(f"  Sub-minimum-wage entries: {len(result['sub_min_wage'])}")
        for e in result["sub_min_wage"][:5]:
            print(f"    {e['name']:25s} {e['section']:14s} ${e['value']:.2f}")
    else:
        print(f"  Sub-minimum-wage check: PASS")


if __name__ == "__main__":
    period_id = sys.argv[1]
    input_csv = Path(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2] else None
    main(period_id, input_csv)
