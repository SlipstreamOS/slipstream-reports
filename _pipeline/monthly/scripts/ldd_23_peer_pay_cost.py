#!/usr/bin/env python3
"""Labor Deep Dive step 23 — uplift % of moving the August roster to each pay scenario (parity floors, full curve, DC peer
all-in, peer + 5%), same method as ldd_20 (each person's current band → scenario rate, × their August core hours, × 12).
Peer all-in (base + tips): new hire ~$23.00 (DC min $18.40 + ~$4.50 tips), barista ~$24.00, shift lead ~$25.50, ALL ~$27.00.
Output: labor_deep_dive/peer_pay_cost_v13.json"""
import os, json
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__)); DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
SCEN = {"parity_floors": {"trainee": 18.50, "certified": 0, "sl": 22.50, "all": 25.50},        # floors only; 0 = unchanged
        "full_curve": {"trainee": 18.50, "certified": 20.50, "sl": 23.50, "all": 27.00},
        "peer": {"trainee": 23.00, "certified": 24.00, "sl": 25.50, "all": 27.00}}
SCEN["peer_plus5"] = {k: round(v * 1.05, 2) for k, v in SCEN["peer"].items()}
# Ryan's framing (Sep 10): do not compete on entry pay; match the PEER AVERAGE per hour with a steeper curve so the top half earns more.
PEER_AVG = 24.25   # hours-weighted peer all-in: mostly baristas at ~$24 (new hires ~$23, shift leads ~$25.50)
SCEN["curve_a"] = {"trainee": 18.50, "certified": 21.00, "sl": 24.50, "all": 28.00}          # tops of the proposed bands
SCEN["peer_avg_curve"] = {"trainee": 18.50, "certified": 21.50, "sl": 26.00, "all": 29.50}   # same average as peers on the Aug hours mix
SCEN["peer_avg_curve_plus5"] = {k: (v if k == "trainee" else round(v * 1.05, 2)) for k, v in SCEN["peer_avg_curve"].items()}   # trainee stays at the minimum
def band(rate):
    return "trainee" if rate < 19.0 else "certified" if rate < 20.5 else "sl" if rate < 22.5 else "all"
def main():
    sh = pd.read_csv(os.path.join(DD, "shifts.csv")); a = sh[(sh.date >= "2026-08-01") & (sh.bucket.isin(["core", "admin", "training"]))].copy()
    people = a.groupby("tm").agg(rate=("rate", "median"), hrs=("hours", "sum")).reset_index(); people["band"] = people.rate.apply(band)
    core = a[a.bucket == "core"]; tot_h = float(core.hours.sum()); tot_w = float(core.wages.sum())
    mix = (people.groupby("band").hrs.sum() / people.hrs.sum()).round(3).to_dict()
    out = {"basis": "August 2026 roster (core + admin + training hours), each person's band → scenario rate; no one is cut", "avg_rate_now": round(tot_w / tot_h, 2), "core_hours_aug": round(tot_h), "peer_avg_all_in": PEER_AVG, "hours_mix_aug": mix,
           "roster_hours_aug_all_buckets": round(float(people.hrs.sum())), "scenarios": {}}
    for name, rates in SCEN.items():
        people[name] = people.band.map(rates); people[name] = people[[name, "rate"]].max(axis=1)   # nobody goes down
        cost_m = float(((people[name] - people.rate) * people.hrs).sum())
        w_new = float((people[name] * people.hrs).sum()); w_now = float((people.rate * people.hrs).sum())
        out["scenarios"][name] = {"rates": rates, "cost_month": round(cost_m), "cost_year": round(cost_m * 12), "avg_rate": round(w_new / float(people.hrs.sum()), 2), "uplift_pct": round((w_new / w_now - 1) * 100, 1),
                                  "by_band": {b: {"people": int((people.band == b).sum()), "avg_now": round(float(people[people.band == b].rate.mean()), 2), "to": rates[b]} for b in ("trainee", "certified", "sl", "all")}}
    json.dump(out, open(os.path.join(DD, "peer_pay_cost_v13.json"), "w"), indent=1); print(json.dumps(out, indent=1))
if __name__ == "__main__":
    main()
