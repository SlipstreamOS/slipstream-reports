"""Payroll pipeline — pulls Square shifts/schedule + saves cached state.

This module does NOT call Square or Asana directly — it consumes JSON
caches that the LLM produces via MCP calls. The LLM writes raw responses
into cache/ and this script joins them into the structures we need:

  - synthetic_csv: a row-per-shift CSV mirroring the format of Ryan's
    Square Shifts export. Used to validate Square pull = export.
  - timecard_review: shifts on the schedule with no clock-in, plus
    actuals where clock-out > scheduled end + threshold_minutes.
  - wage_audit: per-employee comparison of Asana Pay vs Square
    wage_setting hourly rate.

Outputs:
  outputs/payroll/<period>/synthetic_shifts.csv
  outputs/payroll/<period>/timecard_review.txt
  outputs/payroll/<period>/wage_audit.csv
"""
from __future__ import annotations
import csv, json, os, re, sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CACHE = ROOT.parent / "cache"

# -- location & display mapping ------------------------------------------------

# Square location_id -> short code & display name. Display matches the column
# layout used in the monthly/weekly reports (Downtown -> Studio -> 1899 -> NMWA).
LOCATIONS = {
    "LXM7V1QWDHKZJ": ("DT",   "Slipstream Downtown"),
    "LGVKCKX9B9DZT": ("CR",   "Slipstream Studio"),       # Studio = "CR" tab
    "L4ZBMGCWJ5C7R": ("1899", "Slipstream 1899 Penn"),
    "LV2D18H0DD9ES": ("NMWA", "Slipstream NMWA"),
    "LZCA3EY4GB7BM": ("LC",   "Slipstream Commissary (Food)"),  # Edgewood
    "LW5RCZWPJEDWD": ("BL",   "Slipstream Commissary (Bev)"),   # 14th St
    "LFXQF0X8CKYF5": ("RST",  "Slipstream Roasting"),
}

# Asana Location enum -> Square location_id (primary). NMWA folks share
# Studio/Downtown roster, so they have no separate Asana location.
ASANA_LOC_TO_SQUARE = {
    "Downtown":        "LXM7V1QWDHKZJ",
    "Studio Theater ": "LGVKCKX9B9DZT",
    "Studio Theater":  "LGVKCKX9B9DZT",
    "1899":            "L4ZBMGCWJ5C7R",
}

# -- helpers -------------------------------------------------------------------

def cents_to_dollars(c):
    return float(c) / 100.0 if c is not None else None

def parse_iso(s):
    # Square returns RFC3339 like "2026-04-06T07:01:00-04:00"
    return datetime.fromisoformat(s)

def name_key(s):
    """Normalize names for matching. Drops middle initials, trims, lowers."""
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def first_last(name):
    parts = name_key(name).split()
    if not parts:
        return ("", "")
    return (parts[0], parts[-1])

# -- loaders -------------------------------------------------------------------

def load_square_team():
    """Returns (by_id, by_name_key) lookups."""
    data = json.loads((CACHE / "square_team.json").read_text())
    members = data.get("team_members", [])
    by_id = {}
    by_name = {}
    for m in members:
        first = (m.get("given_name") or "").strip()
        last = (m.get("family_name") or "").strip()
        full = f"{first} {last}".strip()
        rec = {
            "id": m["id"],
            "first": first,
            "last": last,
            "full": full,
            "status": m.get("status"),
            "wage_setting": m.get("wage_setting", {}),
            "locations": m.get("assigned_locations", {}).get("location_ids", []),
        }
        by_id[m["id"]] = rec
        by_name[name_key(full)] = rec
    return by_id, by_name

def load_square_jobs():
    data = json.loads((CACHE / "square_jobs.json").read_text())
    return {j["id"]: j["title"] for j in data.get("jobs", [])}

def load_asana_team():
    """Asana Team Summary -> [{name, pay, location, ...}] for non-archive entries."""
    data = json.loads((CACHE / "asana_team.json").read_text())
    out = []
    for item in data.get("data", []):
        cf = {f["name"]: f for f in item.get("custom_fields", [])}
        loc = (cf.get("Location", {}) or {}).get("display_value")
        if loc in (None, "ARCHIVE"):
            continue
        if item["name"].strip().lower() == "development tracker":
            continue
        out.append({
            "name": item["name"].strip(),
            "pay":  (cf.get("Pay", {}) or {}).get("number_value"),
            "asana_location": loc,
            "asana_gid": item["gid"],
        })
    return out

# -- person matcher ------------------------------------------------------------

import difflib

# Manual overrides — Asana name (key) -> Square name (preferred match).
# Use when names diverge in ways the fuzzy matcher can't reliably bridge.
NAME_OVERRIDES = {
    "griselda romero": "gricelda hernandez",
    "katherine sung":  "kat sung",
    "elijah r schug":  "eli schug",
}

def build_person_match(asana_team, square_by_name):
    """For each Asana entry, find the matching Square team_member.

    Strategy:
      1. Manual override
      2. Exact normalized full-name match
      3. Last-name match where first letter also matches
      4. First-name unique match (handles last-name changes)
      5. Fuzzy (difflib) similarity > 0.80
      6. Otherwise: no_match (surfaced to user)
    """
    by_first = defaultdict(list)
    by_last  = defaultdict(list)
    for k, rec in square_by_name.items():
        f, l = first_last(rec["full"])
        by_first[f].append(rec)
        by_last[l].append(rec)

    sq_keys = list(square_by_name.keys())

    matches = []
    for a in asana_team:
        key = name_key(a["name"])
        f, l = first_last(a["name"])

        # 1. manual override
        if key in NAME_OVERRIDES:
            rec = square_by_name.get(NAME_OVERRIDES[key])
            matches.append({"asana": a, "square": rec, "method": "override"}); continue

        # 2. exact normalized full match
        if key in square_by_name:
            matches.append({"asana": a, "square": square_by_name[key], "method": "exact"}); continue

        # 3. last name + first letter (avoids Griselda Romero → Amy Romero)
        cands = by_last.get(l, [])
        good = [c for c in cands if c["first"][:1].lower() == f[:1].lower()]
        if len(good) == 1:
            matches.append({"asana": a, "square": good[0], "method": "last+first1"}); continue

        # 4. first name unique match (Delmi Aleman → Delmi Menjivar)
        cands = by_first.get(f, [])
        if len(cands) == 1:
            matches.append({"asana": a, "square": cands[0], "method": "first_unique"}); continue

        # 5. difflib fuzzy similarity on full normalized name
        best = difflib.get_close_matches(key, sq_keys, n=1, cutoff=0.80)
        if best:
            matches.append({"asana": a, "square": square_by_name[best[0]], "method": f"fuzzy({difflib.SequenceMatcher(None, key, best[0]).ratio():.2f})"}); continue

        matches.append({"asana": a, "square": None, "method": "no_match"})
    return matches

# -- wage audit ----------------------------------------------------------------

def wage_audit(matches):
    """Compare Asana Pay vs Square current wage_setting (max across job_assignments)."""
    rows = []
    for m in matches:
        a = m["asana"]
        s = m["square"]
        asana_pay = a["pay"]
        if s is None:
            rows.append({
                "name_asana": a["name"], "name_square": "",
                "asana_pay": asana_pay, "square_pay": None,
                "delta": None, "status": "MISSING_IN_SQUARE",
                "asana_location": a["asana_location"],
            })
            continue
        wages = []
        for ja in s["wage_setting"].get("job_assignments", []):
            r = ja.get("hourly_rate") or {}
            if r.get("amount") is not None:
                wages.append(cents_to_dollars(r["amount"]))
        sq_pay = max(wages) if wages else None
        delta = None if (asana_pay is None or sq_pay is None) else round(asana_pay - sq_pay, 2)
        if delta is None:
            status = "NO_PAY_IN_SQUARE" if sq_pay is None else "NO_PAY_IN_ASANA"
        elif abs(delta) < 0.01:
            status = "OK"
        else:
            status = "MISMATCH"
        rows.append({
            "name_asana": a["name"],
            "name_square": s["full"],
            "asana_pay": asana_pay,
            "square_pay": sq_pay,
            "delta": delta,
            "status": status,
            "asana_location": a["asana_location"],
        })
    return rows

# -- shift processing ----------------------------------------------------------

def load_period(period_id):
    """Loads cached timecards/scheduled shifts for a given period directory."""
    pdir = CACHE / "periods" / period_id
    timecards = []
    scheduled = []
    for f in sorted(pdir.glob("timecards_*.json")):
        timecards.extend(json.loads(f.read_text()).get("timecards", []))
    for f in sorted(pdir.glob("scheduled_*.json")):
        scheduled.extend(json.loads(f.read_text()).get("scheduled_shifts", []))
    return timecards, scheduled

def synth_shift_rows(timecards, sq_by_id, jobs):
    """Render timecards as rows matching Ryan's Square shifts export."""
    rows = []
    for tc in timecards:
        loc_id = tc.get("location_id")
        loc_disp = LOCATIONS.get(loc_id, ("?", loc_id))[1]
        tm = sq_by_id.get(tc.get("team_member_id"), {})
        wage = tc.get("wage", {}) or {}
        hr = (wage.get("hourly_rate") or {}).get("amount")
        start = parse_iso(tc["start_at"]) if tc.get("start_at") else None
        end   = parse_iso(tc["end_at"])   if tc.get("end_at")   else None
        total_hours = round((end - start).total_seconds() / 3600, 2) if (start and end) else 0
        cash_tip = ((tc.get("declared_cash_tip_money") or {}).get("amount") or 0) / 100.0
        rows.append({
            "first": tm.get("first", ""),
            "last":  tm.get("last", ""),
            "title": wage.get("title", "") or jobs.get(wage.get("job_id"), ""),
            "location": loc_disp,
            "clockin_date":  start.strftime("%-m/%-d/%y") if start else "",
            "clockin_time":  start.strftime("%-I:%M:%S %p %Z") if start else "",
            "clockout_date": end.strftime("%-m/%-d/%y") if end else "",
            "clockout_time": end.strftime("%-I:%M:%S %p %Z") if end else "",
            "total_paid_hours": total_hours,
            "hourly_wage": cents_to_dollars(hr) if hr is not None else None,
            "labor_cost": round(total_hours * (cents_to_dollars(hr) or 0), 2),
            "declared_cash_tips": cash_tip,
            "status": tc.get("status"),
        })
    return rows

def timecard_review(timecards, scheduled, sq_by_id, threshold_min=30, salary_employees=None):
    """Identify (a) scheduled shifts with no clock-in,
              (b) timecards with clock-out > scheduled end + threshold,
              (c) timecards with no scheduled match (likely unscheduled extra shift).

    salary_employees: set of full_name strings — these are filtered out of the
    missing-clockin list (they don't clock in by design)."""
    salary_employees = salary_employees or set()
    # Index actual timecards: (team_member_id, location_id, workday)
    def workday(dt):
        return dt.date().isoformat()

    actuals_by_key = defaultdict(list)
    for tc in timecards:
        if not tc.get("start_at"): continue
        key = (tc["team_member_id"], tc["location_id"], workday(parse_iso(tc["start_at"])))
        actuals_by_key[key].append(tc)

    sched_by_key = defaultdict(list)
    for ss in scheduled:
        d = ss.get("published_shift_details") or ss.get("draft_shift_details") or {}
        if not d.get("team_member_id") or d.get("is_deleted"):
            continue
        key = (d["team_member_id"], d["location_id"], workday(parse_iso(d["start_at"])))
        sched_by_key[key].append(d)

    missing = []          # scheduled but no clock-in
    late_clockouts = []   # clock-out exceeds scheduled end by threshold
    unscheduled = []      # actual shift with no match in schedule

    for key, scheds in sched_by_key.items():
        actuals = actuals_by_key.get(key, [])
        if not actuals:
            for d in scheds:
                tm = sq_by_id.get(d["team_member_id"], {})
                full_name = tm.get("full", d["team_member_id"])
                if full_name in salary_employees:
                    continue
                missing.append({
                    "name": full_name,
                    "location": LOCATIONS.get(d["location_id"], ("?", d["location_id"]))[1],
                    "scheduled_start": parse_iso(d["start_at"]).strftime("%a %m/%d %-I:%M %p"),
                    "scheduled_end":   parse_iso(d["end_at"]).strftime("%-I:%M %p"),
                })
        else:
            # pair earliest scheduled to earliest actual; check clock-out drift
            scheds_sorted = sorted(scheds, key=lambda d: d["start_at"])
            actuals_sorted = sorted(actuals, key=lambda t: t["start_at"])
            for d, a in zip(scheds_sorted, actuals_sorted):
                if not a.get("end_at"): continue
                sched_end = parse_iso(d["end_at"])
                actual_end = parse_iso(a["end_at"])
                drift_min = (actual_end - sched_end).total_seconds() / 60
                if drift_min >= threshold_min:
                    tm = sq_by_id.get(d["team_member_id"], {})
                    late_clockouts.append({
                        "name": tm.get("full", d["team_member_id"]),
                        "location": LOCATIONS.get(d["location_id"], ("?", d["location_id"]))[1],
                        "date": parse_iso(d["start_at"]).strftime("%a %m/%d"),
                        "scheduled_end": sched_end.strftime("%-I:%M %p"),
                        "actual_end":    actual_end.strftime("%-I:%M %p"),
                        "drift_min": int(round(drift_min)),
                    })

    # actuals not in schedule
    for key, actuals in actuals_by_key.items():
        if key not in sched_by_key:
            for a in actuals:
                tm = sq_by_id.get(a["team_member_id"], {})
                start = parse_iso(a["start_at"])
                end = parse_iso(a["end_at"]) if a.get("end_at") else None
                unscheduled.append({
                    "name": tm.get("full", a["team_member_id"]),
                    "location": LOCATIONS.get(a["location_id"], ("?", a["location_id"]))[1],
                    "date":  start.strftime("%a %m/%d"),
                    "actual_start": start.strftime("%-I:%M %p"),
                    "actual_end":   end.strftime("%-I:%M %p") if end else "OPEN",
                    "hours": round(((end - start).total_seconds() / 3600) if end else 0, 2),
                })

    return {"missing_clockins": missing, "late_clockouts": late_clockouts, "unscheduled": unscheduled}


# -- entry point ---------------------------------------------------------------

SALARY_EMPLOYEES = {"Ines Chavez"}  # Don't clock in; filter from missing-clockin lists.

def main(period_id):
    sq_by_id, sq_by_name = load_square_team()
    jobs = load_square_jobs()
    asana = load_asana_team()
    matches = build_person_match(asana, sq_by_name)
    audit = wage_audit(matches)

    # Save wage audit
    out = ROOT / period_id
    out.mkdir(parents=True, exist_ok=True)
    with open(out / "wage_audit.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["name_asana","name_square","asana_location","asana_pay","square_pay","delta","status"])
        w.writeheader()
        for r in audit:
            w.writerow(r)

    # Process shifts
    timecards, scheduled = load_period(period_id)
    synth = synth_shift_rows(timecards, sq_by_id, jobs)
    with open(out / "synthetic_shifts.csv", "w", newline="") as f:
        flds = ["first","last","title","location","clockin_date","clockin_time","clockout_date","clockout_time","total_paid_hours","hourly_wage","labor_cost","declared_cash_tips","status"]
        w = csv.DictWriter(f, fieldnames=flds)
        w.writeheader()
        for r in synth:
            w.writerow(r)

    review = timecard_review(timecards, scheduled, sq_by_id, threshold_min=30,
                             salary_employees=SALARY_EMPLOYEES)
    lines = []
    lines.append(f"# Timecard review — period {period_id}")
    lines.append(f"# threshold for late clock-out: 30 minutes past scheduled end\n")
    lines.append(f"## Missing clock-ins ({len(review['missing_clockins'])})")
    for r in review["missing_clockins"]:
        lines.append(f"  - {r['name']:25s}  {r['location']:32s}  {r['scheduled_start']} → {r['scheduled_end']}")
    lines.append(f"\n## Late clock-outs >30min ({len(review['late_clockouts'])})")
    for r in sorted(review["late_clockouts"], key=lambda r: -r["drift_min"]):
        lines.append(f"  - {r['name']:25s}  {r['location']:32s}  {r['date']}  scheduled {r['scheduled_end']} → actual {r['actual_end']}  (+{r['drift_min']} min)")
    lines.append(f"\n## Unscheduled shifts (actual without a matching schedule entry) ({len(review['unscheduled'])})")
    for r in sorted(review["unscheduled"], key=lambda r: r["date"]):
        lines.append(f"  - {r['name']:25s}  {r['location']:32s}  {r['date']}  {r['actual_start']} → {r['actual_end']}  ({r['hours']}h)")
    (out / "timecard_review.txt").write_text("\n".join(lines))

    print(f"Wrote {out}/wage_audit.csv  ({sum(1 for r in audit if r['status']=='MISMATCH')} mismatches)")
    print(f"Wrote {out}/synthetic_shifts.csv  ({len(synth)} shifts)")
    print(f"Wrote {out}/timecard_review.txt  (missing={len(review['missing_clockins'])}, late>{30}m={len(review['late_clockouts'])}, unscheduled={len(review['unscheduled'])})")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "2026-04-05_2026-04-18")
