#!/usr/bin/env python3
"""Labor Deep Dive step 4 — clean the harvested timecards into shifts and build:
  * per cafe × date: core/admin/training hours & wages (shift-level)
  * per cafe × tier × DOW: ACTUAL headcount on the clock by hour (core only)
  * role mix and wage rates by cafe / role / daypart
  * August 2026 wage totals vs P&L labor -> loaded-cost factor per cafe
Outputs: data/labor_deep_dive/shifts.csv, staffing.json
"""
import os, json, statistics
from collections import defaultdict, Counter
from datetime import datetime, timedelta
import pandas as pd, numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
LOC = {"LXM7V1QWDHKZJ": "Downtown", "LGVKCKX9B9DZT": "Studio", "L4ZBMGCWJ5C7R": "1899 Penn", "LV2D18H0DD9ES": "NMWA"}
DOWN = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
# close hour by cafe × dow (used only to detect missed clock-outs)
CLOSE = {"Downtown": {d: 16 for d in range(5)},
         "Studio": {0: 14, 1: 14, 2: 14, 3: 14, 4: 17, 5: 17, 6: 17},
         "1899 Penn": {d: 15 for d in range(5)},
         "NMWA": {d: 14 for d in range(7)}}
MAX_PLAUSIBLE = 16.0
LATE_GRACE_H = 2.25      # a punch ending later than close + grace is treated as a missed clock-out

def bucket(title):
    t = (title or "").strip()
    if t == "Manager": return "admin"
    if "Training" in t: return "training"
    if t == "Event": return "event"
    return "core"

def main():
    raw = json.load(open(os.path.join(DD, "timecards_raw.json")))["timecards"]
    rows = []
    for tc in raw:
        cafe = LOC.get(tc.get("location_id"))
        if not cafe: continue
        s = datetime.fromisoformat(tc["start_at"]).replace(tzinfo=None); e = datetime.fromisoformat(tc["end_at"]).replace(tzinfo=None)  # Eastern wall-clock
        rows.append({"id": tc["id"], "cafe": cafe, "tm": tc.get("team_member_id"), "role": tc.get("wage", {}).get("title", "(none)"),
                     "rate": tc.get("wage", {}).get("hourly_rate", {}).get("amount", 0) / 100.0,
                     "start": s, "end": e, "date": s.date().isoformat(), "dow": s.weekday()})
    df = pd.DataFrame(rows); df["start"] = pd.to_datetime(df.start); df["end"] = pd.to_datetime(df.end)
    df["bucket"] = df.role.map(bucket)
    df["hours_raw"] = (df.end - df.start).dt.total_seconds() / 3600.0
    notes = {}
    # --- zero-length punches
    zero = df[df.hours_raw < 0.25]
    notes["dropped_zero_length"] = int(len(zero))
    df = df[df.hours_raw >= 0.25].copy()
    # --- zero-rate cards: impute the team member's modal rate (else cafe×role median)
    modal = df[df.rate > 0].groupby("tm").rate.agg(lambda s: s.mode().iloc[0])
    role_med = df[df.rate > 0].groupby(["cafe", "role"]).rate.median()
    imputed = 0
    for i, r in df[df.rate == 0].iterrows():
        v = modal.get(r.tm)
        if v is None or (isinstance(v, float) and np.isnan(v)):
            v = role_med.get((r.cafe, r.role), df[df.rate > 0].rate.median())
        df.at[i, "rate"] = float(v); imputed += 1
    notes["zero_rate_imputed"] = imputed
    # --- missed clock-outs. Rule A (task spec): >= 16h -> median shift length for cafe+role.
    #     Rule B: ends later than close + grace -> cap end at that cafe×role median END time (same dow class), keep start.
    clean = df[(df.hours_raw < MAX_PLAUSIBLE)].copy()
    clean["end_h"] = clean.end.dt.hour + clean.end.dt.minute / 60
    close_of = clean.apply(lambda r: CLOSE[r.cafe].get(r.dow, 16), axis=1)
    clean_ok = clean[clean.end_h <= close_of + LATE_GRACE_H]
    med_len = clean_ok.groupby(["cafe", "role"]).hours_raw.median()
    med_len_cafe = clean_ok.groupby("cafe").hours_raw.median()
    med_end = clean_ok.groupby(["cafe", "role"]).end_h.median()
    med_end_cafe = clean_ok.groupby("cafe").end_h.median()
    capped = []
    df["end_adj"] = df.end
    for i, r in df.iterrows():
        close = CLOSE[r.cafe].get(r.dow, 16)
        end_h = r.end.hour + r.end.minute / 60 + (24 if r.end.date() > r.start.date() else 0)
        if r.hours_raw >= MAX_PLAUSIBLE:
            L = med_len.get((r.cafe, r.role), med_len_cafe.get(r.cafe, 8.0))
            new_end = r.start + timedelta(hours=float(L)); rule = "A_16h"
        elif end_h > close + LATE_GRACE_H and r.cafe != "NMWA" and r.bucket != "event":
            eh = float(med_end.get((r.cafe, r.role), med_end_cafe.get(r.cafe, close)))
            new_end = r.start.replace(hour=int(eh), minute=int(round((eh % 1) * 60)) % 60, second=0)
            if new_end <= r.start + timedelta(hours=1): new_end = r.start + timedelta(hours=float(med_len.get((r.cafe, r.role), 8.0)))
            if new_end >= r.end:      # never extend a punch
                continue
            rule = "B_past_close"
        else:
            continue
        capped.append({"id": r.id, "cafe": r.cafe, "date": r.date, "role": r.role, "orig_hours": round(r.hours_raw, 2),
                       "capped_hours": round((new_end - r.start).total_seconds() / 3600, 2), "rule": rule,
                       "wages_avoided": round((r.hours_raw - (new_end - r.start).total_seconds() / 3600) * r.rate, 2)})
        df.at[i, "end_adj"] = new_end
    df["hours"] = (df.end_adj - df.start).dt.total_seconds() / 3600.0
    df["wages"] = df.hours * df.rate
    notes["capped"] = capped
    notes["capped_count_ruleA"] = sum(1 for c in capped if c["rule"] == "A_16h")
    notes["capped_count_ruleB"] = sum(1 for c in capped if c["rule"] == "B_past_close")
    notes["capped_wages_avoided"] = round(sum(c["wages_avoided"] for c in capped), 2)
    df.to_csv(os.path.join(DD, "shifts.csv"), index=False)

    # --- tier per cafe × date from tiers.json weeks
    tiers = json.load(open(os.path.join(DD, "tiers.json")))
    week_tier = {c: {w["week"]: w["tier"] for w in t["weeks"]} for c, t in tiers.items()}
    def tier_for(cafe, date):
        d = datetime.fromisoformat(date).date(); wk = (d - timedelta(days=d.weekday())).isoformat()
        return week_tier.get(cafe, {}).get(wk, "slow" if cafe == "NMWA" else None)
    df["tier"] = [tier_for(c, d) for c, d in zip(df.cafe, df.date)]

    # --- per cafe × date totals
    daily = df.groupby(["cafe", "date", "bucket"]).agg(hours=("hours", "sum"), wages=("wages", "sum"), shifts=("id", "size")).reset_index()
    daily.to_csv(os.path.join(DD, "shifts_daily.csv"), index=False)

    # --- hourly headcount (core), per cafe × tier × dow, averaged over sampled days; quarter-hour resolution
    core = df[df.bucket == "core"].copy()
    def coverage(sub, lo=5, hi=19):
        grid = np.arange(lo, hi, 0.25)
        cov = np.zeros(len(grid))
        for _, r in sub.iterrows():
            s = r.start.hour + r.start.minute / 60; e = r.end_adj.hour + r.end_adj.minute / 60 + (24 if r.end_adj.date() > r.start.date() else 0)
            cov += ((grid + 0.125) >= s) & ((grid + 0.125) < e)
        return grid, cov
    staffing = {}
    for cafe in ["Downtown", "Studio", "1899 Penn", "NMWA"]:
        staffing[cafe] = {}
        sub_c = core[core.cafe == cafe]
        for tier in ["slow", "mid", "busy"]:
            sub_t = sub_c[sub_c.tier == tier]
            if sub_t.empty: continue
            staffing[cafe][tier] = {}
            for dow in range(7):
                sub_d = sub_t[sub_t.dow == dow]
                if sub_d.empty: continue
                days = sorted(sub_d.date.unique())
                grid, cov = coverage(sub_d)
                cov = cov / len(days)
                hourly = {str(h): round(float(cov[(grid >= h) & (grid < h + 1)].mean()), 2) for h in range(5, 19)}
                by_role = sub_d.groupby("role").agg(hours=("hours", "sum"), wages=("wages", "sum"), shifts=("id", "size"))
                staffing[cafe][tier][DOWN[dow]] = {
                    "n_days": len(days), "days": days,
                    "core_hours_per_day": round(float(sub_d.hours.sum()) / len(days), 2),
                    "core_wages_per_day": round(float(sub_d.wages.sum()) / len(days), 2),
                    "avg_rate": round(float(sub_d.wages.sum() / sub_d.hours.sum()), 2),
                    "headcount_by_hour": hourly,
                    "headcount_quarter": {f"{g:.2f}": round(float(v), 2) for g, v in zip(grid, cov)},
                    "shifts_per_day": round(len(sub_d) / len(days), 2),
                    "role_hours_per_day": {r: round(v / len(days), 2) for r, v in by_role.hours.items()},
                    "first_start": str(sub_d.start.dt.time.min())[:5], "median_start_h": round(float(sub_d.start.apply(lambda t: t.hour + t.minute/60).median()), 2),
                    "median_end_h": round(float(sub_d.end_adj.apply(lambda t: t.hour + t.minute/60).median()), 2),
                    "shifts": [{"date": r.date, "role": r.role, "start": r.start.strftime("%H:%M"), "end": r.end_adj.strftime("%H:%M"), "hours": round(r.hours, 2), "rate": r.rate} for _, r in sub_d.sort_values(["date", "start"]).iterrows()],
                }
    # --- role mix & rates by cafe (core), by daypart of shift midpoint
    rolemix = {}
    for cafe in ["Downtown", "Studio", "1899 Penn", "NMWA"]:
        sub = core[core.cafe == cafe]
        tot_h = sub.hours.sum()
        rm = sub.groupby("role").agg(hours=("hours", "sum"), wages=("wages", "sum"), shifts=("id", "size"))
        rolemix[cafe] = {"core_hours": round(float(tot_h), 1), "avg_core_rate": round(float(sub.wages.sum() / tot_h), 2),
                         "roles": {r: {"share_hours": round(float(v.hours / tot_h * 100), 1), "avg_rate": round(float(v.wages / v.hours), 2), "shifts": int(v.shifts)} for r, v in rm.iterrows()},
                         "rate_distribution": {str(k): int(v) for k, v in Counter(sub.rate.round(2)).most_common()},
                         "people_rate": {tm: round(float(v), 2) for tm, v in sub.groupby("tm").rate.median().items()}}
        # admin + training
        for b in ["admin", "training"]:
            sb = df[(df.cafe == cafe) & (df.bucket == b)]
            rolemix[cafe][b] = {"hours": round(float(sb.hours.sum()), 1), "wages": round(float(sb.wages.sum()), 2), "avg_rate": round(float(sb.wages.sum() / sb.hours.sum()), 2) if sb.hours.sum() else None}
    # --- August 2026 wage totals (full month) for loaded-factor calibration
    aug = df[(df.date >= "2026-08-01") & (df.date <= "2026-08-31")]
    aug_tot = aug.groupby(["cafe", "bucket"]).agg(hours=("hours", "sum"), wages=("wages", "sum")).reset_index()
    aug_dates = aug.groupby("cafe").date.nunique()
    august = {c: {"days_with_timecards": int(aug_dates.get(c, 0)),
                  "by_bucket": {r.bucket: {"hours": round(r.hours, 1), "wages": round(r.wages, 2)} for _, r in aug_tot[aug_tot.cafe == c].iterrows()}} for c in ["Downtown", "Studio", "1899 Penn", "NMWA"]}
    out = {"notes": notes, "staffing": staffing, "rolemix": rolemix, "august_wages": august,
           "weeks_sampled": sorted(set((datetime.fromisoformat(d).date() - timedelta(days=datetime.fromisoformat(d).weekday())).isoformat() for d in df.date))}
    json.dump(out, open(os.path.join(DD, "staffing.json"), "w"), indent=1, default=str)
    # ---- console
    print("NOTES:", {k: v for k, v in notes.items() if k != "capped"})
    for c in notes["capped"]:
        print("   capped", c)
    print("\nAUGUST wages by cafe/bucket:")
    for c, v in august.items():
        print(" ", c, v)
    print("\nROLE MIX (core):")
    for c, v in rolemix.items():
        print(f"  {c}: core hrs {v['core_hours']} avg rate ${v['avg_core_rate']}  admin {v['admin']}  training {v['training']}")
        for r, x in sorted(v["roles"].items(), key=lambda kv: -kv[1]["share_hours"]):
            print(f"      {r:28} {x['share_hours']:5.1f}%  ${x['avg_rate']:.2f}  n={x['shifts']}")

if __name__ == "__main__":
    main()
