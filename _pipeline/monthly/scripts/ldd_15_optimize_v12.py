#!/usr/bin/env python3
"""Labor Deep Dive v1.2 step 15 (adapted from v1.1 step 11) — minimum-cost shift sets per cafe × tier × day-of-week (integer program, CBC).

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
v1.2 rules: 6-hour minimum shifts, 2 open / 2 close at Downtown & Studio (need_v12 floors), planning rates = Aug 2026 roster
averages by role, per-day "std" and "admin-day" sets, Shift Lead <= 9 h. Output: data/labor_deep_dive/templates_v12.json
"""
import os, json
import pulp
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
RATES = {"SL": 22.25, "PL": 19.75, "PB": 20.25, "B": 20.00, "HR": 20.25, "CL": 20.00}      # planning rates: Aug 2026 roster averages by role, rounded up ~$0.25
STD_RATES = {"SL": 21.00, "PL": 19.00, "PB": 19.00, "B": 19.00, "HR": 19.00, "CL": 19.50}
ROLE_NAMES = {"SL": "Shift Lead", "PL": "Production Line", "PB": "Production Barista", "B": "Barista", "HR": "Hospitality Register", "CL": "Closer (keyholder barista)"}
TARGET = {"Downtown": (0.32, 0.36), "Studio": (0.30, 0.34), "1899 Penn": (0.35, 0.40)}
SHIFT_FIXED_COST = 12.0
PREOPEN_PENALTY = 2.0     # $ per hour of barista time before open
MAX_STAFF = {"Downtown": 6, "Studio": 6, "1899 Penn": 2}

def frange(a, b, step=0.5):
    x = a; out = []
    while x <= b + 1e-9:
        out.append(round(x, 2)); x += step
    return out

def candidates(cafe, dow, cfg, sl_max, min_len=6.0, pl_min=6.0):
    o, c = cfg["open"][cafe][dow]; prep = o - cfg["prep"][cafe]; close = c + cfg["tail"][cafe]
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

def solve(cafe, dow, need_rows, cfg, sl_max=9.0, min_len=6.0, pl_min=6.0, pb_min_count=1):
    o, c = cfg["open"][cafe][dow]; close = c + cfg["tail"][cafe]
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
        prob += pulp.lpSum(x[i] for i in fams["PB"]) >= pb_min_count
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
    NV = json.load(open(os.path.join(DD, "need_v12.json")))
    cfg = {"open": NV["open"], "prep": NV["prep"], "tail": NV["tail"]}
    tiers = json.load(open(os.path.join(DD, "tiers.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    T1 = json.load(open(os.path.join(DD, "templates_v11.json")))
    out = {"rates": RATES, "std_rates": STD_RATES, "cap": NV["cap"], "admin_days": NV["admin_days"], "targets": {k: {"target": v[0], "ceiling": v[1], "ceiling_slow": v[1] + 0.02} for k, v in TARGET.items()}, "cafes": {}}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        obs_rate = staffing["rolemix"][cafe]["avg_core_rate"]
        out["cafes"][cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            days_out = {}
            for dow in (DOWN if cafe == "Studio" else DOWN[:5]):
                nrec = NV["cafes"][cafe]["tiers"][tier][dow]
                o, c = cfg["open"][cafe][dow]; close = c + cfg["tail"][cafe]
                sales = tiers[cafe]["typical_day"][tier][dow]["net"]
                rec = {"sales_typical": sales, "admin_day": nrec["admin_day"], "sets": {}}
                for setname in ["std", "admin"]:
                    need_rows = nrec["sets"][setname]["by_halfhour"]
                    pb_min = 0 if (cafe == "1899 Penn" and tier == "slow") else 1
                    slmax = 10.0 if (cafe == "Downtown" or (cafe == "Studio" and dow in ("Fri", "Sat", "Sun"))) else 9.0     # long days: open-to-close lead allowed (Ryan: on the table)
                    shifts = solve(cafe, dow, need_rows, cfg, slmax, 6.0, 6.0, pb_min)
                    sc = score(shifts, need_rows, sales, obs_rate, o, close)
                    sc["wages_std_rates"] = round(sum((s["end"] - s["start"]) * STD_RATES[s["role"]] for s in shifts), 2)
                    rec["sets"][setname] = {"need_hours": nrec["sets"][setname]["need_hours"], "peak_need": nrec["sets"][setname]["peak_need"],
                        "shifts": [{"role": ROLE_NAMES[s["role"]], "code": s["role"], "start": fmt(s["start"]), "end": fmt(s["end"]), "start_h": s["start"], "end_h": s["end"],
                                    "hours": round(s["end"] - s["start"], 2), "rate": RATES[s["role"]], "wages": round((s["end"] - s["start"]) * RATES[s["role"]], 2)} for s in shifts],
                        "template": sc}
                act = staffing["staffing"][cafe].get(tier, {}).get(dow)
                if act:
                    rec["actual_sampled"] = {"core_hours": act["core_hours_per_day"], "core_wages": act["core_wages_per_day"], "avg_rate": act["avg_rate"], "labor_pct": round(act["core_wages_per_day"] / sales * 100, 1), "n_days": act["n_days"]}
                v11 = T1["cafes"][cafe][f"{tier}_no10h"]["days"][dow]
                rec["v11_template"] = {"hours": v11["template"]["hours"], "wages": v11["template"]["wages"], "labor_pct": v11["template"]["labor_pct"], "shifts": [f"{s['code']} {s['start']}–{s['end']}" for s in v11["shifts"]]}
                if "ryan_first_cut" in v11: rec["ryan_first_cut"] = v11["ryan_first_cut"]
                days_out[dow] = rec
            # weekly rollups: primary = admin set on admin days (slow tier only matters), std elsewhere; also all-std week
            def roll(select):
                W = {"sales": 0.0, "hours": 0.0, "wages": 0.0, "wages_obs": 0.0, "wages_std": 0.0, "single": 0.0, "under": 0.0, "act_h": 0.0, "act_w": 0.0, "v11_h": 0.0, "v11_w": 0.0, "ry_h": 0.0, "ry_w": 0.0}
                for dow, d in days_out.items():
                    st = d["sets"][select(d)]["template"]
                    W["sales"] += d["sales_typical"]; W["hours"] += st["hours"]; W["wages"] += st["wages"]; W["wages_obs"] += st["wages_obs_rate"]; W["wages_std"] += st["wages_std_rates"]
                    W["single"] += st["single_cover_hours"]; W["under"] += st["under_person_hours"]
                    if "actual_sampled" in d: W["act_h"] += d["actual_sampled"]["core_hours"]; W["act_w"] += d["actual_sampled"]["core_wages"]
                    W["v11_h"] += d["v11_template"]["hours"]; W["v11_w"] += d["v11_template"]["wages"]
                    if "ryan_first_cut" in d: W["ry_h"] += d["ryan_first_cut"]["hours"]; W["ry_w"] += d["ryan_first_cut"]["wages"]
                tg = TARGET[cafe]; ceil = tg[1] + (0.02 if tier == "slow" else 0.0)
                return {"sales_typical": round(W["sales"]), "hours": round(W["hours"], 2), "wages": round(W["wages"]), "labor_pct": round(W["wages"] / W["sales"] * 100, 1),
                        "wages_std_rates": round(W["wages_std"]), "labor_pct_std_rates": round(W["wages_std"] / W["sales"] * 100, 1),
                        "wages_obs_rate": round(W["wages_obs"]), "labor_pct_obs_rate": round(W["wages_obs"] / W["sales"] * 100, 1),
                        "single_cover_hours": round(W["single"], 1), "under_person_hours": round(W["under"], 1),
                        "splh": round(W["sales"] / W["hours"], 1), "labor_pct_minus10": round(W["wages"] / (W["sales"] * 0.9) * 100, 1), "labor_pct_plus10": round(W["wages"] / (W["sales"] * 1.1) * 100, 1),
                        "actual_sampled_hours": round(W["act_h"], 1), "actual_sampled_wages": round(W["act_w"]), "actual_sampled_labor_pct": round(W["act_w"] / W["sales"] * 100, 1) if W["act_w"] else None,
                        "v11_hours": round(W["v11_h"], 1), "v11_wages": round(W["v11_w"]), "v11_labor_pct": round(W["v11_w"] / W["sales"] * 100, 1),
                        "ryan_first_cut_hours": round(W["ry_h"], 1), "ryan_first_cut_wages": round(W["ry_w"]), "ryan_first_cut_labor_pct": round(W["ry_w"] / W["sales"] * 100, 1) if W["ry_w"] else None,
                        "target_wages": round(W["sales"] * tg[0]), "ceiling": ceil, "ceiling_wages": round(W["sales"] * ceil),
                        "ceiling_trigger_sales": round(W["wages"] / ceil), "target_trigger_sales": round(W["wages"] / tg[0])}
            out["cafes"][cafe][tier] = {"days": days_out, "week": roll(lambda d: "admin" if d["admin_day"] else "std"), "week_no_admin": roll(lambda d: "std"), "tier": tier}
    json.dump(out, open(os.path.join(DD, "templates_v12.json"), "w"), indent=1)
    for cafe, tiers_o in out["cafes"].items():
        print(f"\n=== {cafe} (obs rate ${staffing['rolemix'][cafe]['avg_core_rate']}) ===")
        for tier, o in tiers_o.items():
            w = o["week"]; wn = o["week_no_admin"]
            print(f" {tier:5} sales ${w['sales_typical']:,}  {w['hours']:6.2f}h ${w['wages']:,} = {w['labor_pct']}% (std-rates {w['labor_pct_std_rates']}%, -10% {w['labor_pct_minus10']}%) SPLH ${w['splh']} solo {w['single_cover_hours']}h | no-admin week {wn['hours']}h {wn['labor_pct']}% | v1.1 {w['v11_hours']}h {w['v11_labor_pct']}% | actual {w['actual_sampled_hours']}h {w['actual_sampled_labor_pct']}% | 1st cut {w['ryan_first_cut_hours']}h {w['ryan_first_cut_labor_pct']}% | ceil {w['ceiling']:.2f} trig ${w['ceiling_trigger_sales']:,}")
            for dow, d in o["days"].items():
                sel = "admin" if d["admin_day"] else "std"; t = d["sets"][sel]["template"]
                print(f"    {dow}{'*' if d['admin_day'] else ' '} ${d['sales_typical']:,.0f}: {t['hours']:5.2f}h ${t['wages']:5.0f} {t['labor_pct']:5.1f}% pk {t['peak_staff']} solo {t['solo_windows']} admin-cover {t['admin_cover_windows']} | " + "; ".join(f"{s['code']} {s['start']}-{s['end']}" for s in d["sets"][sel]["shifts"]))

if __name__ == "__main__":
    main()
