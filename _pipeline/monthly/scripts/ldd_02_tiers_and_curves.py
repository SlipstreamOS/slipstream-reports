#!/usr/bin/env python3
"""Labor Deep Dive step 2 — demand tiers + hourly demand curves.

Tier assignment is by ACTUAL weekly walk-in sales (Mon–Sun), scaled to a full
week when a federal holiday shortens it, so each tier's curve is built from days
that really were that busy. Bands (weekly walk-in $) per cafe are set below and
reported with month membership (2026 actuals Jan–Aug, projections Sep–Dec).
Outputs: data/labor_deep_dive/tiers.json, curves.json
"""
import os, json
import pandas as pd, numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

# Weekly walk-in sales bands (full week). Set from the 2026 weekly distribution
# (see ldd notebook output): natural breaks at the month clusters.
BANDS = {
    "Downtown":  {"slow": (0, 7500),   "mid": (7500, 9000),   "busy": (9000, 1e9)},
    "Studio":    {"slow": (0, 10500),  "mid": (10500, 12500), "busy": (12500, 1e9)},
    "1899 Penn": {"slow": (0, 3300),   "mid": (3300, 4200),   "busy": (4200, 1e9)},
    "NMWA":      {"slow": (0, 1e9)},   # staffing rule, single tier
}
OPEN_DAYS = {"Downtown": [0,1,2,3,4], "Studio": [0,1,2,3,4,5,6], "1899 Penn": [0,1,2,3,4], "NMWA": [2,3,4,5,6]}
# weeks below this (scaled to a full week) are weather/closure anomalies, not demand tiers (e.g. DT week of Jan 26: $2,365)
MIN_WEEKLY = {"Downtown": 4500, "Studio": 7000, "1899 Penn": 1800, "NMWA": 0}
MIN_DAY_NET = {"Downtown": 100, "Studio": 100, "1899 Penn": 100, "NMWA": 20}
# open hours (start hour inclusive, end hour exclusive) by cafe × dow for the curve window
HOURS = {
    "Downtown": {d: (7, 16) for d in range(5)},
    "Studio": {0: (7,14), 1: (7,14), 2: (7,14), 3: (7,14), 4: (7,17), 5: (8,17), 6: (8,17)},
    "1899 Penn": {d: (7, 15) for d in range(5)},
    "NMWA": {d: (10, 14) for d in range(7)},   # community days (2nd Wed / 1st Sun) can run later
}

def tier_of(cafe, weekly):
    for t, (lo, hi) in BANDS[cafe].items():
        if lo <= weekly < hi:
            return t
    return "busy"

def main():
    df = pd.read_csv(os.path.join(DD, "txns_2026.csv"))
    df["date"] = pd.to_datetime(df["date"])
    df["halfhour"] = df["hour"] + (df["minute"] >= 30) * 0.5
    out_tiers, out_curves = {}, {}
    for cafe in ["Downtown", "Studio", "1899 Penn", "NMWA"]:
        sub = df[df.cafe == cafe].copy()
        daily = sub.groupby("date").agg(net=("net", "sum"), txns=("net", "size")).reset_index()
        daily = daily[daily.net > MIN_DAY_NET[cafe]].copy()   # drop test/closed days
        daily["dow"] = daily.date.dt.dayofweek
        daily = daily[daily.dow.isin(OPEN_DAYS[cafe])]
        daily["week"] = daily.date.dt.to_period("W-SUN").apply(lambda p: p.start_time.normalize())
        wk = daily.groupby("week").agg(net=("net", "sum"), days=("net", "size")).reset_index()
        full = len(OPEN_DAYS[cafe]) if cafe != "NMWA" else 4
        wk["scaled"] = wk.net * full / wk.days.clip(lower=1)   # holiday-shortened weeks scaled to a full week
        wk = wk[(wk.days >= full - 1) & (wk.scaled >= MIN_WEEKLY[cafe])] if cafe != "NMWA" else wk   # full-1 open days & not an anomaly week
        wk["tier"] = wk.scaled.map(lambda v: tier_of(cafe, v))
        daily = daily.merge(wk[["week", "tier", "scaled"]], on="week", how="inner")
        # month membership (share of weeks by tier)
        daily["month"] = daily.date.dt.month
        mm = {}
        for m, g in daily.groupby("month"):
            w = g.drop_duplicates("week")
            vc = w.tier.value_counts()
            mm[int(m)] = {"weeks": int(len(w)), "tiers": {t: int(vc.get(t, 0)) for t in BANDS[cafe]},
                          "avg_weekly_scaled": round(float(w.scaled.mean()), 0),
                          "dominant": vc.idxmax()}
        # typical daily sales per tier × dow
        tier_dow = daily.groupby(["tier", "dow"]).agg(net=("net", "mean"), txns=("txns", "mean"), n=("net", "size")).reset_index()
        td = {t: {DOWN[int(r.dow)]: {"net": round(r.net, 0), "txns": round(r.txns, 1), "n_days": int(r.n)}
                  for _, r in tier_dow[tier_dow.tier == t].iterrows()} for t in BANDS[cafe]}
        out_tiers[cafe] = {"bands_weekly": {t: [lo, (hi if hi < 1e8 else None)] for t, (lo, hi) in BANDS[cafe].items()},
                           "month_membership": mm, "typical_day": td,
                           "weeks": [{"week": str(r.week.date()), "net": round(r.net, 0), "days": int(r.days), "scaled": round(r.scaled, 0), "tier": r.tier} for _, r in wk.iterrows()]}
        # hourly curves: per tier × dow × hour, averaged per open day
        sub = sub.merge(daily[["date", "tier"]], on="date", how="inner")
        curves = {}
        for (t, d), g in sub.groupby(["tier", "dow"]):
            ndays = g.date.nunique()
            lo, hi = HOURS[cafe][d]
            hrs = list(range(lo - 1, hi + 1))   # one hour either side of open hours
            hh = g.groupby("hour").agg(net=("net", "sum"), txns=("net", "size"))
            h30 = g.groupby("halfhour").agg(net=("net", "sum"), txns=("net", "size"))
            curves.setdefault(t, {})[DOWN[d]] = {
                "n_days": int(ndays),
                "hourly": {str(h): {"txns": round(float(hh.txns.get(h, 0)) / ndays, 1), "net": round(float(hh.net.get(h, 0)) / ndays, 0)} for h in hrs},
                "halfhour": {str(h): {"txns": round(float(h30.txns.get(h, 0)) / ndays, 1), "net": round(float(h30.net.get(h, 0)) / ndays, 0)} for h in np.arange(lo - 1, hi + 1, 0.5)},
                "outside_hours_txns": round(float(g[(g.hour < lo - 1) | (g.hour >= hi + 1)].shape[0]) / ndays, 1),
            }
        # whole-2026 curve per dow
        all_c = {}
        for d, g in sub.groupby("dow"):
            ndays = g.date.nunique(); lo, hi = HOURS[cafe][d]
            hh = g.groupby("hour").agg(net=("net", "sum"), txns=("net", "size"))
            all_c[DOWN[d]] = {"n_days": int(ndays), "hourly": {str(h): {"txns": round(float(hh.txns.get(h, 0)) / ndays, 1), "net": round(float(hh.net.get(h, 0)) / ndays, 0)} for h in range(lo - 1, hi + 1)}}
        out_curves[cafe] = {"by_tier": curves, "all_2026": all_c}
    json.dump(out_tiers, open(os.path.join(DD, "tiers.json"), "w"), indent=1)
    json.dump(out_curves, open(os.path.join(DD, "curves.json"), "w"), indent=1)
    # ---- console report ----
    for cafe, t in out_tiers.items():
        print(f"\n=== {cafe} bands (weekly walk-in $): {t['bands_weekly']}")
        print("month  weeks  slow/mid/busy   avg weekly (scaled)  dominant")
        for m, v in t["month_membership"].items():
            print(f"  {m:>2}     {v['weeks']}     {v['tiers']}   {v['avg_weekly_scaled']:>8,.0f}   {v['dominant']}")
        print("typical day by tier × DOW ($ / txns / n):")
        for tier, dows in t["typical_day"].items():
            print(f"  {tier:5}", "  ".join(f"{d}: ${v['net']:,.0f}/{v['txns']:.0f}/{v['n_days']}" for d, v in dows.items()))

if __name__ == "__main__":
    main()
