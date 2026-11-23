#!/usr/bin/env python3
"""Labor Deep Dive step 5 — throughput capacity from the data.
Joins hourly transactions with hourly core headcount for every sampled day and
looks at transactions per person-hour in the busiest sustained hours.
Output: data/labor_deep_dive/capacity.json (+ console)
"""
import os, json
import pandas as pd, numpy as np
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")

def main():
    tx = pd.read_csv(os.path.join(DD, "txns_2026.csv"))
    sh = pd.read_csv(os.path.join(DD, "shifts.csv"), parse_dates=["start", "end_adj"])
    sh = sh[sh.bucket == "core"]
    sampled_days = set(zip(sh.cafe, sh.date))
    tx = tx[[ (c, d) in sampled_days for c, d in zip(tx.cafe, tx.date)]]
    hourly_tx = tx.groupby(["cafe", "date", "hour"]).agg(txns=("net", "size"), net=("net", "sum")).reset_index()
    # headcount per (cafe, date, hour): fraction of the hour covered summed over shifts
    recs = []
    for (cafe, date), g in sh.groupby(["cafe", "date"]):
        for h in range(5, 19):
            hs = pd.Timestamp(date) + pd.Timedelta(hours=h); he = hs + pd.Timedelta(hours=1)
            cov = 0.0
            for _, r in g.iterrows():
                ov = (min(r.end_adj, he) - max(r.start, hs)).total_seconds() / 3600
                cov += max(0.0, ov)
            recs.append({"cafe": cafe, "date": date, "hour": h, "heads": cov})
    hc = pd.DataFrame(recs)
    j = hourly_tx.merge(hc, on=["cafe", "date", "hour"], how="right").fillna({"txns": 0, "net": 0})
    j = j[j.heads >= 0.9]                        # hours with at least ~1 person on the clock
    j["tph"] = j.txns / j.heads                  # transactions per person-hour
    j["sph"] = j.net / j.heads
    out = {}
    for cafe, g in j.groupby("cafe"):
        g = g[g.txns > 0]
        busy = g[g.txns >= g.txns.quantile(0.80)]          # busiest sustained hours (top quintile of hours)
        q = lambda s, p: float(np.percentile(s, p))
        out[cafe] = {
            "hours_observed": int(len(g)),
            "tph_all_hours": {"p50": round(q(g.tph, 50), 1), "p75": round(q(g.tph, 75), 1), "p90": round(q(g.tph, 90), 1)},
            "busy_hours": {"n": int(len(busy)), "txns_threshold": float(g.txns.quantile(0.80)),
                            "tph_p50": round(q(busy.tph, 50), 1), "tph_p75": round(q(busy.tph, 75), 1), "tph_p90": round(q(busy.tph, 90), 1),
                            "tph_max": round(float(busy.tph.max()), 1),
                            "heads_p50": round(q(busy.heads, 50), 2), "sph_p50": round(q(busy.sph, 50), 0)},
            "by_heads": {},
        }
        # observed throughput by headcount level (rounded) in busy hours
        for hlev, gg in busy.assign(hl=busy.heads.round()).groupby("hl"):
            out[cafe]["by_heads"][str(int(hlev))] = {"n_hours": int(len(gg)), "txns_p50": round(float(gg.txns.median()), 1), "txns_p90": round(float(gg.txns.quantile(0.9)), 1), "tph_p50": round(float(gg.tph.median()), 1)}
        # top-20 single hours by txns to see what the team has actually handled
        top = g.sort_values("txns", ascending=False).head(12)
        out[cafe]["top_hours"] = [{"date": r.date, "hour": int(r.hour), "txns": int(r.txns), "heads": round(r.heads, 2), "tph": round(r.tph, 1)} for _, r in top.iterrows()]
    json.dump(out, open(os.path.join(DD, "capacity.json"), "w"), indent=1)
    for c, v in out.items():
        print(f"\n=== {c} === hours observed {v['hours_observed']}")
        print("  all hours tph p50/p75/p90:", v["tph_all_hours"])
        print("  busy hours:", v["busy_hours"])
        print("  by headcount:", v["by_heads"])
        print("  top hours:", [(t['date'], t['hour'], t['txns'], t['heads'], t['tph']) for t in v["top_hours"][:8]])

if __name__ == "__main__":
    main()
