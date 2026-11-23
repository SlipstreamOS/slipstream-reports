#!/usr/bin/env python3
"""Labor Deep Dive v1.1 step 11 — minimum-cost shift sets per cafe × tier × day-of-week (integer program, CBC).

For every cafe × tier × DOW the half-hour need curve (need_v11.json) is covered by shift 'families' with
practical shapes:
  Downtown  : Production Line opens 6:00 (6–8 h) · Shift Lead opens 6:30 (8–10 h) · optional Closer
              (keyholder barista, ends 4:30p, 4–8 h) · up to two 4–6 h Barista rush shifts.
  Studio    : Production Line opens 6:00 (7:00 Sat–Sun) · Shift Lead opens 6:30 (7:30 Sat–Sun), 8–10 h ·
              optional Closer (ends close + tail) · up to two (three Fri–Sun) 4–8 h Barista / Register shifts.
  1899 Penn : Shift Lead 6:30–3:30 fixed · one or two 4–8 h Production Barista shifts (at least one — nobody
              works the cafe alone all day) · never more than 2 on the floor (a 3rd person at the 1p peak is
              'Admin covers').
A keyholder (Shift Lead or Closer) is on the floor every half hour from open−30 min to close+tail.
Objective: minimise wages at standard rates, +$3 per shift (fewer shifts win ties), +$1 per half hour a
barista starts before open (rush shifts start at/after open unless coverage needs them earlier).
Variants: 'target' (v1.1 floors, cafe capacity) = the templates; 'comfort' (v1.0-style floors, capacity 12);
'no10h' = target with Shift Lead ≤ 9 h (so the 10-hour open-to-close option's value is visible).
Output: data/labor_deep_dive/templates_v11.json
"""
import os, json
import pulp
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
RATES = {"SL": 21.00, "PL": 19.00, "PB": 19.00, "B": 19.00, "HR": 19.00, "CL": 19.50}
ROLE_NAMES = {"SL": "Shift Lead", "PL": "Production Line", "PB": "Production Barista", "B": "Barista", "HR": "Hospitality Register", "CL": "Closer (keyholder barista)"}
TARGET = {"Downtown": (0.32, 0.36), "Studio": (0.30, 0.34), "1899 Penn": (0.35, 0.40)}
SHIFT_FIXED_COST = 12.0
PREOPEN_PENALTY = 2.0     # $ per hour of barista time before open
MAX_STAFF = {"Downtown": 5, "Studio": 6, "1899 Penn": 2}

def frange(a, b, step=0.5):
    x = a; out = []
    while x <= b + 1e-9:
        out.append(round(x, 2)); x += step
    return out

def candidates(cafe, dow, cfg, sl_max, min_len=4.0, pl_min=6.0):
    o, c = cfg["open"][cafe][dow]; prep = o - cfg["prep_lead"][cafe]; close = c + cfg["tail"][cafe][dow]
    cands = []
    if cafe == "1899 Penn":
        cands.append(("SL", "SL", prep, close))
        for s in frange(prep, close - min_len):
            for L in frange(min_len, 8):
                if s + L <= close + 1e-9: cands.append(("PB", "PB", s, s + L))
        return cands
    for L in frange(pl_min, 8):
        cands.append(("PL", "PL", prep, prep + L))
    sl_start = o - 0.5
    ends = {close} if close - sl_start <= sl_max + 1e-9 else set()
    for L in frange(8, sl_max):
        if sl_start + L <= close + 1e-9: ends.add(round(sl_start + L, 2))
    for e in sorted(ends):
        cands.append(("SL", "SL", sl_start, e))
    for L in frange(min_len, 8):
        if close - L >= o - 0.5 - 1e-9: cands.append(("CL", "CL", round(close - L, 2), close))
    maxL = 8 if (cafe == "Studio" and dow in ("Fri", "Sat", "Sun")) else max(6, min_len)
    for s in frange(prep, close - min_len):
        for L in frange(min_len, maxL):
            if s + L <= close + 1e-9: cands.append(("B", "B", s, s + L))
    return cands

def solve(cafe, dow, need_rows, cfg, sl_max=10.0, min_len=4.0, pl_min=6.0):
    o, c = cfg["open"][cafe][dow]; close = c + cfg["tail"][cafe][dow]
    grid = [r["t"] for r in need_rows]; need = {r["t"]: r["need"] for r in need_rows}
    cands = candidates(cafe, dow, cfg, sl_max, min_len, pl_min)
    prob = pulp.LpProblem("shifts", pulp.LpMinimize)
    x = [pulp.LpVariable(f"x{i}", lowBound=0, upBound=2, cat="Integer") for i in range(len(cands))]
    txn = {r["t"]: r["txns_hr"] for r in need_rows}
    def cost_of(fam, role, s, e):
        c_ = (e - s) * RATES[role] + SHIFT_FIXED_COST
        if fam == "B" and s < o: c_ += PREOPEN_PENALTY * (o - s)
        if fam in ("B", "PB", "CL"): c_ -= 0.002 * sum(txn.get(g, 0) for g in grid if s <= g < e)   # tie-break: cover the busiest half-hours
        return c_
    prob += pulp.lpSum(x[i] * cost_of(*cands[i]) for i in range(len(cands)))
    # coverage (need above MAX_STAFF is left to Admin cover)
    for g in grid:
        n = min(need[g], MAX_STAFF[cafe])
        if n > 0:
            prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if s <= g < e) >= n
        prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if s <= g < e) <= MAX_STAFF[cafe]
    fams = {}
    for i, (fam, role, s, e) in enumerate(cands): fams.setdefault(fam, []).append(i)
    if cafe != "1899 Penn":
        prob += pulp.lpSum(x[i] for i in fams["PL"]) == 1
        prob += pulp.lpSum(x[i] for i in fams["CL"]) <= 1
        prob += pulp.lpSum(x[i] for i in fams["B"]) <= 3
    else:
        prob += pulp.lpSum(x[i] for i in fams["PB"]) >= 1
        prob += pulp.lpSum(x[i] for i in fams["PB"]) <= 2
    prob += pulp.lpSum(x[i] for i in fams["SL"]) == 1
    for g in grid:
        if o - 0.5 <= g < close and need[g] > 0:
            prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if fam in ("SL", "CL") and s <= g < e) >= 1
    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    assert pulp.LpStatus[prob.status] == "Optimal", (cafe, dow, pulp.LpStatus[prob.status])
    shifts = []
    for i, (fam, role, s, e) in enumerate(cands):
        for _ in range(int(round(x[i].value() or 0))): shifts.append({"role": role, "start": s, "end": e})
    order = {"PL": 0, "SL": 1, "PB": 2, "B": 3, "HR": 3, "CL": 4}
    shifts.sort(key=lambda s: (s["start"], order[s["role"]]))
    return shifts

def fmt(t):
    h = int(t); m = int(round((t - h) * 60)); return f"{h}:{m:02d}"

def score(shifts, need_rows, sales, obs_rate, o, close):
    grid = [r["t"] for r in need_rows]; need = [r["need"] for r in need_rows]; tx = [r["txns_hr"] for r in need_rows]
    staffed = [sum(1 for s in shifts if s["start"] <= g < s["end"]) for g in grid]
    hours = sum(s["end"] - s["start"] for s in shifts); wages = sum((s["end"] - s["start"]) * RATES[s["role"]] for s in shifts)
    over = sum(max(0, a - n) for a, n in zip(staffed, need)) * 0.5
    under = sum(max(0, n - a) for a, n in zip(staffed, need)) * 0.5
    under_slots = [g for g, a, n in zip(grid, staffed, need) if n > a]
    single = sum(0.5 for g, a in zip(grid, staffed) if a == 1 and o <= g < close - 0.5)
    tph = max([t / a for a, t in zip(staffed, tx) if a > 0] + [0])
    def windows(flags):
        w = []; cur = None
        for g, f in zip(grid, flags):
            if f and cur is None: cur = [g, g + 0.5]
            elif f: cur[1] = g + 0.5
            elif cur is not None: w.append(cur); cur = None
        if cur is not None: w.append(cur)
        return [f"{fmt(a)}–{fmt(b)}" for a, b in w]
    solo_windows = windows([a == 1 and o <= g < close - 0.5 for g, a in zip(grid, staffed)])
    admin_windows = windows([n > a for a, n in zip(staffed, need)])
    return {"solo_windows": solo_windows, "admin_cover_windows": admin_windows,"hours": round(hours, 2), "wages": round(wages, 2), "wages_obs_rate": round(hours * obs_rate, 2),
            "labor_pct": round(wages / sales * 100, 1) if sales else None, "labor_pct_obs_rate": round(hours * obs_rate / sales * 100, 1) if sales else None,
            "splh": round(sales / hours, 1) if hours else None, "peak_staff": max(staffed), "over_person_hours": round(over, 1),
            "under_person_hours": round(under, 1), "under_slots": [fmt(g) for g in under_slots], "single_cover_hours": round(single, 1), "max_tph": round(tph, 1),
            "long_shifts": sum(1 for s in shifts if s["end"] - s["start"] >= 10 - 1e-9)}

def main():
    NV = json.load(open(os.path.join(DD, "need_v11.json")))
    cfg = {"open": NV["open"], "prep_lead": NV["prep_lead"], "tail": NV["tail"]}
    tiers = json.load(open(os.path.join(DD, "tiers.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    T0 = json.load(open(os.path.join(DD, "templates.json")))
    out = {"rates": RATES, "cap": NV["cap"], "targets": {k: {"target": v[0], "ceiling": v[1]} for k, v in TARGET.items()}, "cafes": {}}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        obs_rate = staffing["rolemix"][cafe]["avg_core_rate"]
        out["cafes"][cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            for variant, need_variant, sl_max, min_len, pl_min in [("target", "target", 10.0, 4.0, 6.0), ("no10h", "target", 9.0, 4.0, 6.0), ("comfort", "comfort", 9.0, 4.0, 6.0), ("min6h", "target", 9.0, 6.0, 6.0), ("pl8", "target", 9.0, 4.0, 8.0)]:
                key = tier if variant == "target" else f"{tier}_{variant}"
                days_out = {}; wk = {k: 0.0 for k in ["sales", "hours", "wages", "wages_obs", "act_h", "act_w", "v10_h", "v10_w", "ry_h", "ry_w", "need_h", "single", "under"]}
                for dow in (DOWN if cafe == "Studio" else DOWN[:5]):
                    nrec = NV["variants"][need_variant][cafe]["tiers"][tier][dow]
                    need_rows = nrec["by_halfhour"]
                    o, c = cfg["open"][cafe][dow]; close = c + cfg["tail"][cafe][dow]
                    shifts = solve(cafe, dow, need_rows, cfg, sl_max, min_len, pl_min)
                    sales = tiers[cafe]["typical_day"][tier][dow]["net"]
                    sc = score(shifts, need_rows, sales, obs_rate, o, close)
                    rec = {"sales_typical": sales, "need_hours": nrec["need_hours"], "peak_need": nrec["peak_need"],
                           "shifts": [{"role": ROLE_NAMES[s["role"]], "code": s["role"], "start": fmt(s["start"]), "end": fmt(s["end"]), "start_h": s["start"], "end_h": s["end"],
                                       "hours": round(s["end"] - s["start"], 2), "rate": RATES[s["role"]], "wages": round((s["end"] - s["start"]) * RATES[s["role"]], 2)} for s in shifts],
                           "template": sc, "sensitivity_minus10": round(sc["wages"] / (sales * 0.9) * 100, 1)}
                    act = staffing["staffing"][cafe].get(tier, {}).get(dow)
                    if act:
                        rec["actual_sampled"] = {"core_hours": act["core_hours_per_day"], "core_wages": act["core_wages_per_day"], "avg_rate": act["avg_rate"], "labor_pct": round(act["core_wages_per_day"] / sales * 100, 1), "n_days": act["n_days"]}
                        wk["act_h"] += act["core_hours_per_day"]; wk["act_w"] += act["core_wages_per_day"]
                    v10 = T0["cafes"][cafe][tier]["days"][dow]
                    rec["v10_template"] = {"hours": v10["template"]["hours"], "wages": v10["template"]["wages"], "labor_pct": v10["template"]["labor_pct"],
                                           "shifts": [{"code": s["code"], "start": s["start"], "end": s["end"], "hours": s["hours"], "note": s["note"]} for s in v10["shifts"]]}
                    wk["v10_h"] += v10["template"]["hours"]; wk["v10_w"] += v10["template"]["wages"]
                    if "ryan_first_cut" in v10:
                        rec["ryan_first_cut"] = v10["ryan_first_cut"]; wk["ry_h"] += v10["ryan_first_cut"]["hours"]; wk["ry_w"] += v10["ryan_first_cut"]["wages"]
                    days_out[dow] = rec
                    wk["sales"] += sales; wk["hours"] += sc["hours"]; wk["wages"] += sc["wages"]; wk["wages_obs"] += sc["wages_obs_rate"]; wk["need_h"] += nrec["need_hours"]
                    wk["single"] += sc["single_cover_hours"]; wk["under"] += sc["under_person_hours"]
                tg = TARGET[cafe]
                week = {"sales_typical": round(wk["sales"]), "hours": round(wk["hours"], 2), "wages": round(wk["wages"]), "labor_pct": round(wk["wages"] / wk["sales"] * 100, 1),
                        "wages_obs_rate": round(wk["wages_obs"]), "labor_pct_obs_rate": round(wk["wages_obs"] / wk["sales"] * 100, 1),
                        "need_hours": round(wk["need_h"], 1), "single_cover_hours": round(wk["single"], 1), "under_person_hours": round(wk["under"], 1),
                        "splh": round(wk["sales"] / wk["hours"], 1), "labor_pct_minus10": round(wk["wages"] / (wk["sales"] * 0.9) * 100, 1), "labor_pct_plus10": round(wk["wages"] / (wk["sales"] * 1.1) * 100, 1),
                        "actual_sampled_hours": round(wk["act_h"], 1), "actual_sampled_wages": round(wk["act_w"]), "actual_sampled_labor_pct": round(wk["act_w"] / wk["sales"] * 100, 1) if wk["act_w"] else None,
                        "v10_hours": round(wk["v10_h"], 1), "v10_wages": round(wk["v10_w"]), "v10_labor_pct": round(wk["v10_w"] / wk["sales"] * 100, 1),
                        "ryan_first_cut_hours": round(wk["ry_h"], 1), "ryan_first_cut_wages": round(wk["ry_w"]), "ryan_first_cut_labor_pct": round(wk["ry_w"] / wk["sales"] * 100, 1) if wk["ry_w"] else None,
                        "target_wages": round(wk["sales"] * tg[0]), "ceiling_wages": round(wk["sales"] * tg[1]),
                        "ceiling_trigger_sales": round(wk["wages"] / tg[1]), "target_trigger_sales": round(wk["wages"] / tg[0]),
                        "ceiling_trigger_sales_obs_rate": round(wk["wages_obs"] / tg[1])}
                out["cafes"][cafe][key] = {"days": days_out, "week": week, "variant": variant, "tier": tier}
    json.dump(out, open(os.path.join(DD, "templates_v11.json"), "w"), indent=1)
    for cafe, tiers_o in out["cafes"].items():
        print(f"\n=== {cafe} (obs rate ${staffing['rolemix'][cafe]['avg_core_rate']}) ===")
        for key, o in tiers_o.items():
            w = o["week"]
            print(f" {key:13} sales ${w['sales_typical']:,} need {w['need_hours']}h -> {w['hours']:6.2f}h ${w['wages']:,} = {w['labor_pct']}% (obs-rate {w['labor_pct_obs_rate']}%; -10% {w['labor_pct_minus10']}%) SPLH ${w['splh']} single {w['single_cover_hours']}h under {w['under_person_hours']}"
                  f" | v1.0 {w['v10_hours']}h {w['v10_labor_pct']}% | actual {w['actual_sampled_hours']}h {w['actual_sampled_labor_pct']}% | 1st cut {w['ryan_first_cut_hours']}h {w['ryan_first_cut_labor_pct']}%")
            if o["variant"] == "target":
                for dow, d in o["days"].items():
                    t = d["template"]
                    print(f"    {dow}: ${d['sales_typical']:,.0f} need {d['need_hours']}h/pk{d['peak_need']} -> {t['hours']:5.2f}h ${t['wages']:5.0f} {t['labor_pct']:5.1f}% pk {t['peak_staff']} over {t['over_person_hours']} under {t['under_person_hours']} {t['under_slots'][:4]} single {t['single_cover_hours']}h tph {t['max_tph']} | "
                          + "; ".join(f"{s['code']} {s['start']}-{s['end']}" for s in d["shifts"]) + f" | v1.0 {d['v10_template']['hours']}h")

if __name__ == "__main__":
    main()
