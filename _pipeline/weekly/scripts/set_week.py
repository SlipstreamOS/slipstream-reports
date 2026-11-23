#!/usr/bin/env python3
"""Point every weekly-pipeline script at one Sun–Sat week in a single command.

    python3 set_week.py 2026-09-27            # Sunday that starts the week

Rewrites WEEK_START / WEEK_END / WEEK_LABEL / WEEK_TITLE / WEEK_FILE_TAG /
DAYS / DAY_LABELS in the five scripts below, replacing the old manual
find-and-replace step. Idempotent; prints what it changed.
"""
import re, sys
from datetime import date, timedelta

SCRIPTS = ["build_week_from_daily_metrics.py", "aggregate_labor.py",
           "patch_week_from_txstore.py", "build_weekly_data.py", "build_weekly_report.py"]

def main():
    start = date.fromisoformat(sys.argv[1])
    if start.weekday() != 6:
        sys.exit(f"{start} is not a Sunday")
    days = [start + timedelta(days=i) for i in range(7)]
    end = days[-1]
    label = f"{start.strftime('%b').lower()}{start.day}-{end.strftime('%b').lower()}{end.day}"
    title = (f"Week of Sunday {start.strftime('%b')} {start.day} – "
             f"Saturday {end.strftime('%b')} {end.day}, {end.year}")
    tag = f"{start.strftime('%b')} {start.day}-{end.strftime('%b')} {end.day} {end.year}"
    days_lit = ('["' + '","'.join(d.isoformat() for d in days[:4]) + '",\n        "'
                + '","'.join(d.isoformat() for d in days[4:]) + '"]')
    lab = [f"{d.strftime('%a')} {d.month}/{d.day}" for d in days]
    labels_lit = ('["' + '","'.join(lab[:4]) + '",\n              "' + '","'.join(lab[4:]) + '"]')

    subs = [
        (r'^WEEK_START\s*=\s*"[^"]*"', f'WEEK_START = "{start.isoformat()}"'),
        (r'^WEEK_END(\s*)=\s*"[^"]*"', lambda m: f'WEEK_END{m.group(1)}= "{end.isoformat()}"'),
        (r'^WEEK_LABEL\s*=\s*"[^"]*"', f'WEEK_LABEL = "{label}"'),
        (r'^WEEK_TITLE\s*=\s*"[^"]*"', f'WEEK_TITLE = "{title}"'),
        (r'^WEEK_FILE_TAG\s*=\s*"[^"]*"', f'WEEK_FILE_TAG = "{tag}"'),
        (r'^DAYS\s*=\s*\[[^\]]*\]', f'DAYS = {days_lit}'),
        (r'^DAY_LABELS\s*=\s*\[[^\]]*\]', f'DAY_LABELS = {labels_lit}'),
    ]
    for fn in SCRIPTS:
        s = open(fn).read()
        n_total = 0
        for pat, rep in subs:
            s, n = re.subn(pat, rep, s, count=1, flags=re.M)
            n_total += n
        open(fn, "w").write(s)
        print(f"{fn:<34} {n_total} constant(s) set")
    print(f"\nWeek set to {start} .. {end}  (label {label})")

if __name__ == "__main__":
    main()
