#!/usr/bin/env python3
"""Labor Deep Dive v1.1 step 10 — people-needed curves with CAFE-SPECIFIC planning capacity and
explicit floor rules, in two variants, plus a capacity sensitivity (12 / 14 / 16 per person-hour).

TARGET variant (the templates): capacity = busy-hour p90 observed at each cafe (Downtown 14, Studio 12,
  1899 14). Floors: Downtown 1 at 6:00, 2 from 6:30 to 1:00p, then 1 (capacity adds people); Studio
  unchanged from v1.0 (1 at 6:00, 2 to 90 min before close, then 1; Monday 3 at open until 10:00);
  1899 1 (the lead) throughout, 2 in the two peaks (8–10a, 12:30–2p) only when the curve is ≥ 12/hr.
COMFORT variant (the buffer, for comparison): capacity 12 everywhere; 2-person floor to 2:30p at Downtown,
  to 60 min before close at Studio, and 7:00–1:30p at 1899 (the v1.0 rule).
Output: data/labor_deep_dive/need_v11.json  (per cafe × variant × tier × DOW half-hour rows)
"""
import os, json, math
import numpy as np
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

CAP = {"target": {"Downtown": 14.0, "Studio": 12.0, "1899 Penn": 14.0},
       "comfort": {"Downtown": 12.0, "Studio": 12.0, "1899 Penn": 12.0}}
OPEN = {"Downtown": {d: (7, 16) for d in DOWN[:5]},
        "Studio": {"Mon": (7, 14), "Tue": (7, 14), "Wed": (7, 14), "Thu": (7, 14), "Fri": (7, 17), "Sat": (8, 17), "Sun": (8, 17)},
        "1899 Penn": {d: (7, 15) for d in DOWN[:5]}}
PREP_LEAD = {"Downtown": 1.0, "Studio": 1.0, "1899 Penn": 0.5}
TAIL = {"Downtown": {d: 0.5 for d in DOWN}, "Studio": {"Mon": 0.75, "Tue": 0.75, "Wed": 0.75, "Thu": 0.75, "Fri": 0.5, "Sat": 0.5, "Sun": 0.5}, "1899 Penn": {d: 0.5 for d in DOWN}}

def floor_curve(variant, cafe, dow, h, txns_hr):
    o, c = OPEN[cafe][dow]; prep = o - PREP_LEAD[cafe]; close = c + TAIL[cafe][dow]
    if h < prep or h >= close:
        return 0
    if cafe == "1899 Penn":
        if variant == "comfort":
            return 1 if h < o else (2 if h < 13.5 else 1)
        in_peak = (8.0 <= h < 10.0) or (12.5 <= h < 14.0)
        return 2 if (in_peak and txns_hr >= 12) else 1
    if cafe == "Downtown":
        if h < prep + 0.5: return 1
        two_until = 13.0 if variant == "target" else 14.5
        return 2 if h < two_until else 1
    # Studio
    if h < prep + 0.5:
        return 1
    if dow == "Mon" and o <= h < 10: return 3          # 3 people at open (catering delivery) until 10:00
    one_from = c - 1.5 if variant == "target" else c - 1.0
    return 1 if h >= one_from else 2

def main():
    curves = json.load(open(os.path.join(DD, "curves.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    rolemix = staffing["rolemix"]
    out = {"cap": CAP, "open": OPEN, "prep_lead": PREP_LEAD, "tail": TAIL, "variants": {}, "capacity_sensitivity": {}}
    for variant in ["target", "comfort"]:
        out["variants"][variant] = {}
        for cafe in ["Downtown", "Studio", "1899 Penn"]:
            rate = rolemix[cafe]["avg_core_rate"]; cap = CAP[variant][cafe]
            out["variants"][variant][cafe] = {"avg_core_rate": rate, "cap": cap, "tiers": {}}
            sens = {}
            for tier, dows in curves[cafe]["by_tier"].items():
                out["variants"][variant][cafe]["tiers"][tier] = {}
                for dow, cv in dows.items():
                    st = staffing["staffing"][cafe].get(tier, {}).get(dow)
                    o, c = OPEN[cafe][dow]
                    grid = np.arange(5, 19, 0.5)
                    rows = []; need_h = {12.0: 0.0, 14.0: 0.0, 16.0: 0.0}
                    for g in grid:
                        hh = cv["halfhour"].get(f"{g}", cv["halfhour"].get(str(g), {"txns": 0}))
                        t = hh["txns"] * 2
                        fl = floor_curve(variant, cafe, dow, g, t)
                        n = max(fl, math.ceil(t / cap)) if fl > 0 else 0
                        for capx in need_h:
                            need_h[capx] += 0.5 * (max(fl, math.ceil(t / capx)) if fl > 0 else 0)
                        a = 0.0
                        if st:
                            a = (st["headcount_quarter"].get(f"{g:.2f}", 0) + st["headcount_quarter"].get(f"{g+0.25:.2f}", 0)) / 2
                        rows.append({"t": float(g), "txns_hr": round(float(t), 1), "need": int(n), "floor": int(fl), "actual": round(float(a), 2)})
                    need = np.array([r["need"] for r in rows]); act = np.array([r["actual"] for r in rows])
                    over = np.clip(act - need, 0, None).sum() * 0.5; under = np.clip(need - act, 0, None).sum() * 0.5
                    out["variants"][variant][cafe]["tiers"][tier][dow] = {
                        "n_days_curve": cv["n_days"], "n_days_staff": st["n_days"] if st else 0,
                        "need_hours": round(float(need.sum() * 0.5), 1), "actual_hours": round(float(act.sum() * 0.5), 1),
                        "over_hours": round(float(over), 1), "under_hours": round(float(under), 1),
                        "over_dollars": round(float(over * rate), 0), "under_dollars": round(float(under * rate), 0),
                        "peak_need": int(need.max()), "peak_actual": round(float(act.max()), 2),
                        "need_hours_by_cap": {str(int(k)): round(v, 1) for k, v in need_h.items()},
                        "by_halfhour": [r for r in rows if (r["t"] >= o - PREP_LEAD[cafe] - 0.5 and r["t"] <= c + 0.75)],
                    }
                    sens.setdefault(tier, {k: 0.0 for k in need_h})
                    for k, v in need_h.items(): sens[tier][k] += v
            if variant == "target":
                out["capacity_sensitivity"][cafe] = {t: {str(int(k)): round(v, 1) for k, v in v2.items()} for t, v2 in sens.items()}
    json.dump(out, open(os.path.join(DD, "need_v11.json"), "w"), indent=1)
    for variant in ["target", "comfort"]:
        print(f"\n######## {variant.upper()} variant")
        for cafe, v in out["variants"][variant].items():
            print(f"=== {cafe} cap {v['cap']} rate ${v['avg_core_rate']} ===")
            for tier, dows in v["tiers"].items():
                nh = sum(r["need_hours"] for r in dows.values()); ah = sum(r["actual_hours"] for r in dows.values())
                oh = sum(r["over_hours"] for r in dows.values()); uh = sum(r["under_hours"] for r in dows.values())
                extra = f"   need@12/14/16: {out['capacity_sensitivity'][cafe][tier]}" if variant == "target" else ""
                print(f"  {tier:5} need {nh:6.1f}h  actual {ah:6.1f}h  over {oh:5.1f}h  under {uh:4.1f}h{extra}")
                print("        " + "  ".join(f"{d}:{r['need_hours']:.1f}/{r['actual_hours']:.1f}" for d, r in dows.items()))

if __name__ == "__main__":
    main()
