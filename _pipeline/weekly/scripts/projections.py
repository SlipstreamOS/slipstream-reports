"""Load per-cafe weekly sales & labor $ projections from the Projections workbook.

Single source of truth: `Weekly and Monthly Reports/Projections/2026 Projections v1.x.xlsx`.
Auto-picks the highest v-number in the folder so future drops (v1.2, v1.3, ...)
are picked up without pipeline edits.

Workbook layout (verified v1.1, 2026-04-22):
  Per-cafe tabs:  "Downtown 2026", "Studio 2026", "1899 2026", "NMWA 2026"
  Row 5 header row. Row 6+ daily rows.
  Col F (6) = 2026 date (datetime)
  Col G (7) = Sales projection ($)
  Col H (8) = Labor $ projection

Cells can be None (closed days, pre-open cafes, un-filled rows) — treated as 0.
"""
import os
import re
import glob
from datetime import datetime, date, timedelta
from openpyxl import load_workbook

from _paths import WORKSPACE

PROJ_DIR = os.path.join(WORKSPACE, "Projections")

# Map internal cafe name → sheet name in the projections workbook.
CAFE_TO_SHEET = {
    "Downtown":  "Downtown 2026",
    "Studio":    "Studio 2026",
    "1899 Penn": "1899 2026",
    "NMWA":      "NMWA 2026",
}

COL_DATE = 6   # F
COL_SALES = 7  # G
COL_LABOR = 8  # H
DATA_START_ROW = 6


def find_latest_projections_file():
    """Return the absolute path to the highest-versioned 2026 Projections v*.xlsx
    in the Projections/ folder. Ignores lock files and temp files."""
    pattern = os.path.join(PROJ_DIR, "2026 Projections v*.xlsx")
    candidates = [
        p for p in glob.glob(pattern)
        if not os.path.basename(p).startswith(("~$", ".~"))
    ]
    if not candidates:
        return None

    def ver_key(path):
        m = re.search(r"v(\d+)(?:\.(\d+))?", os.path.basename(path))
        if not m:
            return (0, 0)
        return (int(m.group(1)), int(m.group(2) or 0))

    candidates.sort(key=ver_key)
    return candidates[-1]



# ---------------------------------------------------------------------------
# FINALIZED LABOR TARGETS (Ryan, Aug 2026). Core wages as % of core walk-in
# (ex-invoice) sales. TARGET = plan; CEILING = requires Ryan's pre-approval
# before publishing the schedule. Between the two = lead's discretion.
# NMWA is deliberately excluded: it is governed by a staffing rule
# (1 person/day; a 2nd person is an ask), not a labor %.
# These are hard-coded here, NOT read from the workbook, because the
# workbook's labor columns are stale (Ryan, Aug 2026).
# ---------------------------------------------------------------------------
# Structure (Ryan, Sep 2026): each cafe has an operating RANGE (low-high) and a
# hard MAX. `target` and `ceiling` are kept as aliases for low/high so older
# call sites keep working.
#   <= low        -> GREEN, at or better than target
#   low .. high   -> GREEN, inside the planned operating range
#   high .. max   -> AMBER, over range but under max (lead's discretion)
#   > max         -> RED, over max; needs Ryan's pre-approval before publishing
CORE_LABOR_TARGETS = {
    "Studio":    {"low": 0.30, "high": 0.34, "max": 0.36},
    "Downtown":  {"low": 0.32, "high": 0.36, "max": 0.38},
    "1899 Penn": {"low": 0.35, "high": 0.40, "max": 0.42},
}
for _b in CORE_LABOR_TARGETS.values():
    _b["target"] = _b["low"]      # alias
    _b["ceiling"] = _b["high"]    # alias

NMWA_STAFF_RULE = {"people_per_day": 1, "note": "2nd person = ask Ryan"}

# Training plan bands (% of core sales), tracked separately from core.
TRAINING_PLAN_PCT = 0.010
TRAINING_AMBER_PCT = 0.015

# US federal holidays 2026 falling on a Monday. On these days Downtown and
# 1899 Penn are CLOSED (zero staffing) — $0 sales is expected, not an anomaly.
FEDERAL_HOLIDAY_MONDAYS_2026 = {
    "2026-01-19",  # MLK Day
    "2026-02-16",  # Presidents Day
    "2026-05-25",  # Memorial Day
    "2026-09-07",  # Labor Day
    "2026-10-12",  # Columbus Day
}
HOLIDAY_CLOSED_CAFES = ("Downtown", "1899 Penn")


def is_holiday_closed(cafe, iso_date):
    from datetime import date as _d
    if _d.fromisoformat(iso_date).weekday() != 0:   # rule covers federal-holiday MONDAYS only
        return False
    """True if `cafe` is expected to be closed on `iso_date` per the
    federal-holiday-Monday rule (Ryan, Aug 2026)."""
    return cafe in HOLIDAY_CLOSED_CAFES and iso_date in FEDERAL_HOLIDAY_MONDAYS_2026


_RECALC_DIR = os.path.join(os.path.dirname(PROJ_DIR), "_pipeline", ".cache", "projections_recalc")


def _ensure_recalced(path):
    """Return a path to a copy of `path` whose formula results are cached.

    Why: the Sales Proj column is a formula (`=IF(D..="","",D..*(1+$C$2))`)
    at cafes that have not been manually overridden — Downtown, notably.
    openpyxl's data_only=True reads only CACHED results, and the workbook is
    saved without a formula cache, so those cells read as None and the cafe
    silently projects $0. Recalculating headlessly with LibreOffice populates
    the cache. Result is memoised on disk and refreshed when the source
    workbook changes. Falls back to the original path if LibreOffice is
    unavailable or fails, so callers always get a usable workbook.
    """
    import shutil, subprocess
    try:
        os.makedirs(_RECALC_DIR, exist_ok=True)
        dst = os.path.join(_RECALC_DIR, os.path.basename(path))
        if os.path.exists(dst) and os.path.getmtime(dst) >= os.path.getmtime(path):
            return dst
        if shutil.which("soffice") is None:
            return path
        subprocess.run(
            ["soffice", "--headless", "--norestore", "--convert-to", "xlsx",
             "--outdir", _RECALC_DIR, path],
            check=True, capture_output=True, timeout=300,
        )
        return dst if os.path.exists(dst) else path
    except Exception:
        return path


def _as_date(d):
    if isinstance(d, date) and not isinstance(d, datetime):
        return d
    if isinstance(d, datetime):
        return d.date()
    if isinstance(d, str):
        return date.fromisoformat(d)
    raise TypeError(f"Unsupported date type: {type(d)}")


def load_projections(start_date, end_date, cafes=None):
    """Load per-cafe daily projections for [start_date, end_date] inclusive.

    Returns (projections, source_filename) where projections is:
        {cafe: {iso_date: {"sales": float, "labor": float}}}

    Missing days (closed cafe, pre-open, unfilled rows) → 0.0.
    """
    if cafes is None:
        cafes = list(CAFE_TO_SHEET.keys())
    start_date = _as_date(start_date)
    end_date = _as_date(end_date)

    path = find_latest_projections_file()
    if not path:
        raise FileNotFoundError(f"No '2026 Projections v*.xlsx' in {PROJ_DIR}")
    wb = load_workbook(_ensure_recalced(path), data_only=True)

    out = {c: {} for c in cafes}
    for cafe in cafes:
        sheet = CAFE_TO_SHEET.get(cafe)
        if sheet is None or sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for r in range(DATA_START_ROW, ws.max_row + 1):
            d = ws.cell(r, COL_DATE).value
            if not isinstance(d, datetime):
                continue
            day = d.date()
            if start_date <= day <= end_date:
                s = ws.cell(r, COL_SALES).value
                l = ws.cell(r, COL_LABOR).value
                out[cafe][day.isoformat()] = {
                    "sales": float(s) if isinstance(s, (int, float)) else 0.0,
                    "labor": float(l) if isinstance(l, (int, float)) else 0.0,
                }

    # Fill any missing day with zeros so callers can iterate cleanly.
    day = start_date
    while day <= end_date:
        iso = day.isoformat()
        for cafe in cafes:
            out[cafe].setdefault(iso, {"sales": 0.0, "labor": 0.0})
        day += timedelta(days=1)

    return out, os.path.basename(path)


def load_labor_targets(cafes=None):
    """Read per-cafe labor target % (Row 3, Col C) from the projections workbook.

    Returns (targets, source_filename) where targets is {cafe: float_pct}
    (e.g. {"Downtown": 0.38, "Studio": 0.38, "1899 Penn": 0.50, "NMWA": 0.28}).
    Returns 0.0 for any cafe whose tab is missing or whose cell is blank.
    """
    if cafes is None:
        cafes = list(CAFE_TO_SHEET.keys())
    path = find_latest_projections_file()
    if not path:
        raise FileNotFoundError(f"No '2026 Projections v*.xlsx' in {PROJ_DIR}")
    wb = load_workbook(_ensure_recalced(path), data_only=True)
    out = {}
    for cafe in cafes:
        sheet = CAFE_TO_SHEET.get(cafe)
        if sheet is None or sheet not in wb.sheetnames:
            out[cafe] = 0.0
            continue
        v = wb[sheet]["C3"].value
        out[cafe] = float(v) if isinstance(v, (int, float)) else 0.0
    return out, os.path.basename(path)


def load_forward_week_targets(first_week_start, num_weeks=4, cafes=None):
    """Aggregate per-cafe sales + labor $ targets for the next `num_weeks`
    starting from `first_week_start` (a Sunday).

    Returns (weeks, source_filename) where weeks is a list of:
        {
          "week_start": iso,
          "week_end":   iso,
          "cafes": {cafe: {"sales": float, "labor": float}}
        }
    """
    if cafes is None:
        cafes = list(CAFE_TO_SHEET.keys())
    first_week_start = _as_date(first_week_start)
    last_day = first_week_start + timedelta(days=7 * num_weeks - 1)
    proj, fname = load_projections(first_week_start, last_day, cafes)

    weeks = []
    for w in range(num_weeks):
        ws_start = first_week_start + timedelta(days=7 * w)
        ws_end = ws_start + timedelta(days=6)
        entry = {
            "week_start": ws_start.isoformat(),
            "week_end":   ws_end.isoformat(),
            "cafes":      {},
        }
        for cafe in cafes:
            s_tot = l_tot = 0.0
            d = ws_start
            while d <= ws_end:
                v = proj[cafe].get(d.isoformat(), {"sales": 0.0, "labor": 0.0})
                s_tot += v["sales"]
                l_tot += v["labor"]
                d += timedelta(days=1)
            entry["cafes"][cafe] = {"sales": s_tot, "labor": l_tot}
        weeks.append(entry)

    return weeks, fname


if __name__ == "__main__":
    # Smoke-test: print Apr 12-18 and the next 4 forward weeks.
    from datetime import date as _d
    proj, src = load_projections(_d(2026, 4, 12), _d(2026, 4, 18))
    print(f"Projections source: {src}")
    print(f"\n-- Week of Apr 12-18 per-day --")
    for cafe, days in proj.items():
        tot_s = sum(v["sales"] for v in days.values())
        tot_l = sum(v["labor"] for v in days.values())
        print(f"  {cafe}: sales ${tot_s:,.0f}  labor ${tot_l:,.0f}")
        for d in sorted(days):
            v = days[d]
            if v["sales"] or v["labor"]:
                print(f"    {d}: sales ${v['sales']:,.0f}  labor ${v['labor']:,.0f}")

    weeks, _ = load_forward_week_targets(_d(2026, 4, 19), num_weeks=4)
    print(f"\n-- Forward 4 weeks starting Apr 19 --")
    for w in weeks:
        print(f"  {w['week_start']} → {w['week_end']}:")
        for cafe, v in w["cafes"].items():
            print(f"    {cafe:<10}  sales ${v['sales']:>7,.0f}   labor ${v['labor']:>7,.0f}")
