#!/usr/bin/env python3
"""Labor Deep Dive step 6 — people needed by hour vs actual headcount, per cafe × tier × DOW.
Need(h) = max(floor(h), ceil(txns(h) / CAP)).  CAP is observed (capacity.json, busy-hour p90 ≈ 14/person-hour,
used as the planning capacity). Floors: 2 from open through the lunch daypart where a food/production load exists
(Downtown, Studio), 1 in the final 90 minutes; Studio Monday 3 at open (catering); 1899 floor 2 through 1:30p then 1;
NMWA 1 (2 on community days). Over/under coverage in person-hours and $ at the cafe's average core wage.
Output: data/labor_deep_dive/need_vs_actual.json
"""
import os, json, math
import numpy as np
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
CAP = 14.0
OPEN = {"Downtown": {d: (7, 16) for d in DOWN[:5]},
        "Studio": {"Mon": (7, 14), "Tue": (7, 14), "Wed": (7, 14), "Thu": (7, 14), "Fri": (7, 17), "Sat": (8, 17), "Sun": (8, 17)},
        "1899 Penn": {d: (7, 15) for d in DOWN[:5]},
        "NMWA": {d: (10, 14) for d in DOWN}}
PREP_LEAD = {"Downtown": 1.0, "Studio": 1.0, "1899 Penn": 0.5, "NMWA": 1.0}   # hours before open with someone on the clock
CLOSE_TAIL = {"Downtown": 0.5, "Studio": 0.75, "1899 Penn": 0.5, "NMWA": 0.5}

def floor_curve(cafe, dow, h):
    o, c = OPEN[cafe][dow]
    if h < o - PREP_LEAD[cafe] or h >= c + CLOSE_TAIL[cafe]:
        return 0
    if cafe == "NMWA":
        return 1
    if cafe == "1899 Penn":
        return 2 if h < 13.5 else 1
    # Downtown / Studio: food/production load -> 2 at open; last 90 min 1
    if h < o - PREP_LEAD[cafe] + 0.5:      # first half hour of prep: 1 (production line) 
        return 1 if not (cafe == "Studio" and dow == "Mon") else 2
    if h >= c - 1.5:
        return 1
    if cafe == "Studio" and dow == "Mon" and h < 10:
        return 3
    return 2

def main():
    curves = json.load(open(os.path.join(DD, "curves.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    rolemix = staffing["rolemix"]
    out = {"cap": CAP, "cafes": {}}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        rate = rolemix[cafe]["avg_core_rate"]
        out["cafes"][cafe] = {"avg_core_rate": rate, "tiers": {}}
        for tier, dows in curves[cafe]["by_tier"].items():
            out["cafes"][cafe]["tiers"][tier] = {}
            for dow, cv in dows.items():
                st = staffing["staffing"][cafe].get(tier, {}).get(dow)
                if not st: continue
                o, c = OPEN[cafe][dow]
                grid = np.arange(5, 19, 0.5)
                need, act, tx = [], [], []
                for g in grid:
                    hh = cv["halfhour"].get(f"{g}", cv["halfhour"].get(str(g), {"txns": 0}))
                    t = hh["txns"] * 2      # half-hour txns -> hourly rate
                    n_cap = math.ceil(t / CAP) if t > 0 else 0
                    n = max(floor_curve(cafe, dow, g), n_cap) if floor_curve(cafe, dow, g) > 0 else 0
                    # actual headcount at this half hour (average of two quarter slots)
                    a = (st["headcount_quarter"].get(f"{g:.2f}", 0) + st["headcount_quarter"].get(f"{g+0.25:.2f}", 0)) / 2
                    need.append(n); act.append(a); tx.append(t)
                need = np.array(need); act = np.array(act); tx = np.array(tx)
                over = np.clip(act - need, 0, None).sum() * 0.5
                under = np.clip(need - act, 0, None).sum() * 0.5
                out["cafes"][cafe]["tiers"][tier][dow] = {
                    "n_days_curve": cv["n_days"], "n_days_staff": st["n_days"],
                    "need_hours": round(float(need.sum() * 0.5), 1), "actual_hours": round(float(act.sum() * 0.5), 1),
                    "over_hours": round(float(over), 1), "under_hours": round(float(under), 1),
                    "over_dollars": round(float(over * rate), 0), "under_dollars": round(float(under * rate), 0),
                    "by_halfhour": [{"t": float(g), "txns_hr": round(float(t), 1), "need": int(n), "actual": round(float(a), 2)} for g, t, n, a in zip(grid, tx, need, act) if (g >= o - 1.5 and g <= c + 0.5)],
                    "peak_need": int(need.max()), "peak_actual": round(float(act.max()), 2),
                }
    json.dump(out, open(os.path.join(DD, "need_vs_actual.json"), "w"), indent=1)
    # console summary
    for cafe, v in out["cafes"].items():
        print(f"\n=== {cafe} (rate ${v['avg_core_rate']}) ===")
        print(f"{'tier':5} {'dow':3} {'need_h':>7} {'act_h':>6} {'over_h':>6} {'under_h':>7} {'over_$':>7} {'under_$':>7}  peak need/act")
        for tier, dows in v["tiers"].items():
            for dow, r in dows.items():
                print(f"{tier:5} {dow:3} {r['need_hours']:7.1f} {r['actual_hours']:6.1f} {r['over_hours']:6.1f} {r['under_hours']:7.1f} {r['over_dollars']:7.0f} {r['under_dollars']:7.0f}  {r['peak_need']}/{r['peak_actual']}")
    # print one detailed curve as a sanity check
    r = out["cafes"]["Downtown"]["tiers"]["slow"]["Tue"]
    print("\nDowntown slow Tue half-hour detail (t, txns/hr, need, actual):")
    print(" ".join(f"{x['t']:.1f}:{x['txns_hr']:.0f}/{x['need']}/{x['actual']:.1f}" for x in r["by_halfhour"]))

if __name__ == "__main__":
    main()
