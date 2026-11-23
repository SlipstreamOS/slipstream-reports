#!/usr/bin/env python3
"""Labor Deep Dive v1.2 step 14 — people-needed curves under Ryan's Sep 5 2026 rules.

Rules (Ryan, Sep 5): Downtown and Studio run 2 people from open to close — kitchen/Production Line in 45 min before
open, barista/lead 30 min before, two closers to at most 30 min after close. Studio Mon–Thu really needs 3 in the
middle of the day (opener leaves 12:30, third person arrives ~8:00); only in slow months can it run 2 core + the
admin person on admin days. 1899 and NMWA may stagger / go solo in slow months. Admin days on site: Gee Mon/Thu/Fri
(Downtown), Colin Mon/Wed (Studio), Eli Tue (1899), Claire one weekend day (Studio/NMWA).
Capacity per cafe unchanged (Downtown 14, Studio 12, 1899 14). Two need sets per cafe×tier×DOW:
  'std'   — the day without admin on site (the floors above, capacity adds people at the peaks)
  'admin' — the same day when the admin person is on site 9:00–14:00 (Downtown / Studio slow tier only: floor 1 core
            inside that window; Studio Mon–Thu: the mid-day third person is the admin)
Output: data/labor_deep_dive/need_v12.json
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
PREP = {"Downtown": 0.75, "Studio": 0.75, "1899 Penn": 0.5}       # kitchen 45 min before open (Downtown/Studio); lead 30 min at 1899
TAIL = {"Downtown": 0.5, "Studio": 0.5, "1899 Penn": 0.5}           # max 30 min after close
ADMIN_DAYS = {"Downtown": ["Mon", "Thu", "Fri"], "Studio": ["Mon", "Wed", "Sat"], "1899 Penn": ["Tue"]}
ADMIN_WINDOW = (9.0, 14.0)

def floor_curve(cafe, dow, tier, h, txns_hr, admin_on):
    o, c = OPEN[cafe][dow]; prep = o - PREP[cafe]; close = c + TAIL[cafe]
    if h < prep or h >= close: return 0
    aw_end = close if (cafe == "Studio" and dow in DOWN[:4]) else ADMIN_WINDOW[1]     # Studio Mon–Thu: admin stays to close (2nd closer)
    in_admin = admin_on and ADMIN_WINDOW[0] <= h < aw_end
    if cafe == "1899 Penn":
        if tier == "slow":
            return 1                                   # stagger / solo allowed in slow months: capacity alone adds the second person
        # mid/busy: lead all day + second person through the rush-to-lunch stretch (8:30–2:30)
        if 8.5 <= h < 14.5: return 1 if in_admin else 2
        return 1
    # Downtown / Studio: 2 open-to-close (kitchen alone for the first 15 min of prep)
    if h < prep + 0.25: return 1
    base = 2
    if cafe == "Studio" and dow in DOWN[:4]:
        if dow == "Mon" and o <= h < 10: base = 3                    # catering delivery: 3 at open
        elif 8.0 <= h < 12.5 and tier != "slow": base = 3           # third person mid-day in mid/busy
        elif 8.0 <= h < 12.5 and tier == "slow" and not admin_on: base = 3   # slow: third person unless admin is on site
    if tier == "slow" and in_admin and base == 2: base = 1           # slow month + admin on site: 1 core + admin
    return base

def main():
    curves = json.load(open(os.path.join(DD, "curves.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    out = {"cap": CAP, "open": OPEN, "prep": PREP, "tail": TAIL, "admin_days": ADMIN_DAYS, "admin_window": ADMIN_WINDOW, "cafes": {}}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        rate = staffing["rolemix"][cafe]["avg_core_rate"]
        out["cafes"][cafe] = {"avg_core_rate": rate, "cap": CAP[cafe], "tiers": {}}
        for tier, dows in curves[cafe]["by_tier"].items():
            out["cafes"][cafe]["tiers"][tier] = {}
            for dow, cv in dows.items():
                st = staffing["staffing"][cafe].get(tier, {}).get(dow)
                o, c = OPEN[cafe][dow]
                rec = {"n_days_curve": cv["n_days"], "sets": {}}
                for setname, admin_on in [("std", False), ("admin", True)]:
                    rows = []
                    for g in np.arange(5, 19, 0.5):
                        hh = cv["halfhour"].get(f"{g}", cv["halfhour"].get(str(g), {"txns": 0})); t = hh["txns"] * 2
                        fl = floor_curve(cafe, dow, tier, g, t, admin_on)
                        aw_end = (c + TAIL[cafe]) if (cafe == "Studio" and dow in DOWN[:4]) else ADMIN_WINDOW[1]
                        adm_body = 1 if (admin_on and tier == "slow" and ADMIN_WINDOW[0] <= g < aw_end) else 0     # admin on the floor counts as a body for throughput in slow months
                        n = max(fl, math.ceil(t / CAP[cafe]) - adm_body) if fl > 0 else 0
                        a = 0.0
                        if st: a = (st["headcount_quarter"].get(f"{g:.2f}", 0) + st["headcount_quarter"].get(f"{g+0.25:.2f}", 0)) / 2
                        rows.append({"t": float(g), "txns_hr": round(float(t), 1), "need": int(n), "floor": int(fl), "actual": round(float(a), 2)})
                    need = np.array([r["need"] for r in rows]); act = np.array([r["actual"] for r in rows])
                    rec["sets"][setname] = {"need_hours": round(float(need.sum() * 0.5), 1), "actual_hours": round(float(act.sum() * 0.5), 1),
                                            "over_hours": round(float(np.clip(act - need, 0, None).sum() * 0.5), 1), "under_hours": round(float(np.clip(need - act, 0, None).sum() * 0.5), 1),
                                            "over_dollars": round(float(np.clip(act - need, 0, None).sum() * 0.5 * rate)), "peak_need": int(need.max()), "peak_actual": round(float(act.max()), 2),
                                            "by_halfhour": [r for r in rows if (r["t"] >= o - 1.5 and r["t"] <= c + 0.75)]}
                rec["admin_day"] = dow in ADMIN_DAYS[cafe]
                out["cafes"][cafe]["tiers"][tier][dow] = rec
    json.dump(out, open(os.path.join(DD, "need_v12.json"), "w"), indent=1)
    for cafe, v in out["cafes"].items():
        print(f"=== {cafe} ===")
        for tier, dows in v["tiers"].items():
            print(f"  {tier:5} std need {sum(r['sets']['std']['need_hours'] for r in dows.values()):6.1f}h  admin-set need {sum(r['sets']['admin']['need_hours'] for r in dows.values()):6.1f}h  actual {sum(r['sets']['std']['actual_hours'] for r in dows.values()):6.1f}h | "
                  + " ".join(f"{d}:{r['sets']['std']['need_hours']:.1f}{'*' if r['admin_day'] else ''}" for d, r in dows.items()))

if __name__ == "__main__":
    main()
