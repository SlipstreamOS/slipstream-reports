#!/usr/bin/env python3
"""Labor Deep Dive v1.3 step 18 — people-needed curves under Ryan's Sep 5 + Sep 10 rules. No admin sets: admin is a
scheduler's tool, not part of the template (the optimiser output marks where an admin shift would help most).

Floors:
  Downtown  2 from kitchen-in (6:15) to close+30 (16:30); Tue–Thu 3 from 8:00 to 10:00 in every tier (Ryan: third
            person by 8am); Mon/Fri per the curve. Capacity 14/person-hour. Tue/Wed/Thu share one template (solved on the
            max of their curves).
  Studio    2 from 6:15 (7:15 weekends) to close+30; Mon–Thu 3 from 8:00 to 12:30 (Mon from open — delivery). Capacity 12.
  1899      1 from 6:30 to 15:30 (never a solo DAY — the optimiser requires two shifts); capacity 14 adds the second person.
  NMWA      1 from 9:00 to 15:00 (a full hour to set up and close); 2 on community days.
Output: data/labor_deep_dive/need_v13.json
"""
import os, json, math
import numpy as np
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
CAP = {"Downtown": 14.0, "Studio": 12.0, "1899 Penn": 14.0}
OPEN = {"Downtown": {d: (7, 16) for d in DOWN[:5]},
        "Studio": {"Mon": (7, 14), "Tue": (7, 14), "Wed": (7, 14), "Thu": (7, 14), "Fri": (7, 17), "Sat": (8, 17), "Sun": (8, 17)},
        "1899 Penn": {d: (7, 15) for d in DOWN[:5]}}
PREP = {"Downtown": 0.75, "Studio": 0.75, "1899 Penn": 0.5}
TAIL = {"Downtown": 0.5, "Studio": 0.5, "1899 Penn": 0.5}

def floor_curve(cafe, dow, tier, h):
    o, c = OPEN[cafe][dow]; prep = o - PREP[cafe]; close = c + TAIL[cafe]
    if h < prep or h >= close: return 0
    if cafe == "1899 Penn": return 1
    if h < prep + 0.25: return 1
    if cafe == "Downtown":
        return 3 if (dow in ("Tue", "Wed", "Thu") and 8.0 <= h < 10.0) else 2
    if dow in DOWN[:4]:
        if dow == "Mon" and o <= h < 12.5: return 3
        if 8.0 <= h < 12.5: return 3
    return 2

def main():
    curves = json.load(open(os.path.join(DD, "curves.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    out = {"cap": CAP, "open": OPEN, "prep": PREP, "tail": TAIL, "cafes": {}}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        rate = staffing["rolemix"][cafe]["avg_core_rate"]
        out["cafes"][cafe] = {"avg_core_rate": rate, "cap": CAP[cafe], "tiers": {}}
        for tier, dows in curves[cafe]["by_tier"].items():
            out["cafes"][cafe]["tiers"][tier] = {}
            for dow, cv in dows.items():
                st = staffing["staffing"][cafe].get(tier, {}).get(dow)
                o, c = OPEN[cafe][dow]; rows = []
                for g in np.arange(5, 19, 0.5):
                    hh = cv["halfhour"].get(f"{g}", cv["halfhour"].get(str(g), {"txns": 0})); t = hh["txns"] * 2
                    fl = floor_curve(cafe, dow, tier, g)
                    n = max(fl, math.ceil(t / CAP[cafe])) if fl > 0 else 0
                    a = 0.0
                    if st: a = (st["headcount_quarter"].get(f"{g:.2f}", 0) + st["headcount_quarter"].get(f"{g+0.25:.2f}", 0)) / 2
                    rows.append({"t": float(g), "txns_hr": round(float(t), 1), "need": int(n), "floor": int(fl), "actual": round(float(a), 2)})
                need = np.array([r["need"] for r in rows]); act = np.array([r["actual"] for r in rows])
                out["cafes"][cafe]["tiers"][tier][dow] = {"n_days_curve": cv["n_days"], "need_hours": round(float(need.sum() * 0.5), 1), "actual_hours": round(float(act.sum() * 0.5), 1),
                                                          "over_hours": round(float(np.clip(act - need, 0, None).sum() * 0.5), 1), "over_dollars": round(float(np.clip(act - need, 0, None).sum() * 0.5 * rate)),
                                                          "peak_need": int(need.max()), "peak_actual": round(float(act.max()), 2),
                                                          "by_halfhour": [r for r in rows if (r["t"] >= o - 1.5 and r["t"] <= c + 0.75)]}
    json.dump(out, open(os.path.join(DD, "need_v13.json"), "w"), indent=1)
    for cafe, v in out["cafes"].items():
        for tier, dows in v["tiers"].items():
            print(f"{cafe:10} {tier:5} need {sum(r['need_hours'] for r in dows.values()):6.1f}h actual {sum(r['actual_hours'] for r in dows.values()):6.1f}h | " + " ".join(f"{d}:{r['need_hours']:.1f}" for d, r in dows.items()))

if __name__ == "__main__":
    main()
