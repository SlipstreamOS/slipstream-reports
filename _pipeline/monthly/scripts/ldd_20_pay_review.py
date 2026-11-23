#!/usr/bin/env python3
"""Labor Deep Dive v1.3 step 20 — pay review: Slipstream's guaranteed ladder vs DC peers (base + tips), and the cost of a
steeper curve. Peer figures from public postings/aggregators (Sep 2026 searches) — directional.
Output: data/labor_deep_dive/pay_review_v13.json
"""
import os, json
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__)); DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DC_MIN = 18.40
PEERS = [
    {"who": "DC minimum + tips (Ryan's rule of thumb)", "base": "$18.40", "tips": "$4–6", "all_in": "$22.50–24.50", "sl": "+$1–2 → $24–26", "src": "DC DOES; tips = Slipstream's own $4.50–5.60 per core hour"},
    {"who": "Compass Coffee (DC)", "base": "$17.75–19.75", "tips": "~$6", "all_in": "$24–26", "sl": "n/a", "src": "Glassdoor job postings / hourly pay, 2025–26"},
    {"who": "Tatte (DC)", "base": "$15–16.75 (tipped wage)", "tips": "$5–8", "all_in": "$20–24", "sl": "n/a", "src": "ZipRecruiter / Glassdoor postings, 2026"},
    {"who": "Blue Bottle (DC)", "base": "$20.20 listed for Shift Lead", "tips": "$2–4", "all_in": "$22–25", "sl": "$23 avg (Glassdoor)", "src": "Glassdoor, 2026"},
    {"who": "DC barista, all employers", "base": "$19.43 avg (Indeed) / $17.62 (ZipRecruiter)", "tips": "$3–6 ($20/day reported)", "all_in": "$22–25", "sl": "+$1–3", "src": "Indeed / ZipRecruiter / Glassdoor, 2026"},
]
LADDER_NOW = [("Trainee / new (< 90 days)", 18.50, 18.50), ("Certified barista", 19.00, 20.50), ("Shift Lead (SL-1 → SL-3)", 20.50, 22.50), ("Assistant Location Lead", 22.75, 25.25)]
LADDER_PROPOSED = [("Trainee / new (< 90 days)", 18.50, 18.50, "DC minimum $18.40 — below peers by design; short and time-boxed"),
                   ("Certified barista (Essentials + 2 primaries)", 20.00, 21.00, "~$2–3 below a peer barista's all-in; guaranteed, no tip volatility"),
                   ("Shift Lead SL-1 / SL-2 / SL-3", 22.50, 24.50, "parity with a peer barista at a busy cafe; SL-3 above"),
                   ("Assistant Location Lead", 26.00, 28.00, "above any peer hourly role; the 'take on more, make more' promise"),
                   ("+ incentive in on-template months", 0.50, 1.50, "the upside tips used to provide")]

def main():
    sh = pd.read_csv(os.path.join(DD, "shifts.csv")); c = sh[(sh.bucket == "core") & (sh.date >= "2026-08-01")].copy()
    tx = pd.read_csv(os.path.join(DD, "txns_2026.csv")); tx = tx[tx.date >= "2026-08-01"]
    tot_h = float(c.hours.sum()); tot_w = float(c.wages.sum())
    people = c.groupby("tm").agg(rate=("rate", "median"), hrs=("hours", "sum"), role=("role", lambda s: s.mode().iloc[0])).sort_values("rate")
    bands = [("$18.50 (trainee / base)", 0, 18.99), ("$19–20.49 (certified barista)", 19.0, 20.49), ("$20.50–22.49 (Shift Lead)", 20.5, 22.49), ("$22.75–25.25 (ALL)", 22.5, 30)]
    band_rows = []
    for lab, lo, hi in bands:
        q = c[(c.rate >= lo) & (c.rate <= hi)]; pp = people[(people.rate >= lo) & (people.rate <= hi)]
        band_rows.append({"band": lab, "people": int(len(pp)), "share_hours": round(q.hours.sum() / tot_h * 100), "avg_rate": round(q.wages.sum() / max(q.hours.sum(), 1), 2), "hours_month": round(q.hours.sum())})
    tips = {cafe: {"tips": round(float(t.tip.sum())), "tips_per_core_hour": round(float(t.tip.sum()) / max(float(c[c.cafe == cafe].hours.sum()), 1), 2), "tip_pct_of_net": round(float(t.tip.sum()) / float(t.net.sum()) * 100, 1)}
            for cafe, t in tx.groupby("cafe")}
    # cost of the proposed curve: map each person's current band to the proposed midpoint
    def proposed(rate):
        if rate < 19.0: return 18.50
        if rate < 20.5: return 20.50
        if rate < 22.5: return 23.50
        return 27.00
    people["proposed"] = people.rate.apply(proposed); people["delta"] = people.proposed - people.rate
    cost_month = float((people.delta * people.hrs).sum())
    by_band = {}
    for lab, lo, hi in bands:
        pp = people[(people.rate >= lo) & (people.rate <= hi)]
        by_band[lab] = {"people": int(len(pp)), "avg_now": round(float(pp.rate.mean()), 2) if len(pp) else None, "proposed": round(float(pp.proposed.mean()), 2) if len(pp) else None, "cost_month": round(float((pp.delta * pp.hrs).sum()))}
    def step1(rate):
        if rate < 19.0: return 18.50
        if rate < 20.5: return rate
        if rate < 22.5: return max(rate, 22.50)
        return max(rate, 25.50)
    people["step1"] = people.rate.apply(step1); step1_cost = float(((people.step1 - people.rate) * people.hrs).sum())
    core_sales_aug = 93558 + 3606
    out = {"dc_minimum": DC_MIN, "peers": PEERS, "slipstream_now": {"avg_core_rate_aug": round(tot_w / tot_h, 2), "core_hours_aug": round(tot_h), "bands": band_rows, "tips": tips,
                                                                    "note": "Tips are applied toward the guaranteed rate (ADP: $12 tipped base + CC tips + top-up), so a team member's effective pay is the guaranteed rate; Slipstream's own tip pool is $4.50–5.60 per core hour"},
           "ladder_now": [{"level": l, "low": lo, "high": hi} for l, lo, hi in LADDER_NOW],
           "ladder_proposed": [{"level": l, "low": lo, "high": hi, "note": n} for l, lo, hi, n in LADDER_PROPOSED],
           "cost_step1": {"month": round(step1_cost), "year": round(step1_cost * 12), "pct_of_core_sales_aug": round(step1_cost / core_sales_aug * 100, 1), "per_core_hour": round(step1_cost / tot_h, 2),
                          "note": "Step 1 = parity floors only: every Shift Lead to at least $22.50, every ALL to at least $25.50; baristas unchanged"},
           "cost_of_proposed": {"month": round(cost_month), "year": round(cost_month * 12), "pct_of_core_sales_aug": round(cost_month / core_sales_aug * 100, 1), "per_core_hour": round(cost_month / tot_h, 2), "by_band": by_band,
                                "note": "Applied to the August roster at each person's current band midpoint → proposed midpoint; everyone at $18.50 stays (trainee rate); certified $19–20.49 → $20.50; SL $20.50–22.49 → $23.50; ALL → $27.00"}}
    json.dump(out, open(os.path.join(DD, "pay_review_v13.json"), "w"), indent=1)
    print(json.dumps(out["slipstream_now"], indent=1)); print(json.dumps(out["cost_of_proposed"], indent=1))

if __name__ == "__main__":
    main()
