#!/usr/bin/env python3
"""Labor Deep Dive v1.3 step 19 — minimum-cost shift sets under the Sep 5 + Sep 10 rules (integer program, CBC).

Shift families
  Downtown  : Production Line 6:15 (6–8 h) · Shift Lead 6:30 (8–10 h) · Closer ends 4:30 (6–8 h) · Baristas 6–8 h.
              Max 4 people on the floor (Ryan: four handle anything seen so far). Fewer, longer shifts preferred
              ($40 per shift). Tue/Wed/Thu share one template.
  Studio    : Mon–Thu Production Line 6:15–2:30 all day (Ryan: the line person closes the line); Shift Lead 6:30 (8 h);
              baristas are the shorter shifts (6–6.5 h); max 3 (4 if the curve demands).
              Fri: Production Line 6:15 (8 h), Shift Lead 6:30 (8–10 h), closers to 5:30; max 4 in every tier.
              Sat/Sun: Production Line 7:15 (8 h), Shift Lead 7:30 (8–10 h), closers to 5:30; max 4 slow, 5 mid/busy; 6 h shifts fine.
  1899 Penn : two shifts every day (never a solo day). Lead 6:30 start, 6–9 h; Production Barista 6–8.5 h. A second person
              still on at 2:00 stays to close (both close together); on a stagger day (lead 6:30–12:30, second 9:30–3:30)
              the second person closes alone. Max 2 on the floor (a third at a spike = Admin covers).
  NMWA      : FOH 9:00–3:00 (full hour to set up and close); second FOH 9:00–3:00 on community days.
Planning rates: Aug 2026 roster averages by role, rounded up ~$0.25. Output: data/labor_deep_dive/templates_v13.json
"""
import os, json
import pulp
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
RATES = {"SL": 22.25, "PL": 19.75, "PB": 20.25, "B": 20.00, "HR": 20.25, "CL": 20.00, "FOH": 21.25}
ROLE_NAMES = {"SL": "Shift Lead", "PL": "Production Line", "PB": "Production Barista", "B": "Barista", "HR": "Hospitality Register", "CL": "Closer (keyholder barista)", "FOH": "FOH"}
TARGET = {"Downtown": (0.32, 0.36), "Studio": (0.30, 0.34), "1899 Penn": (0.35, 0.40)}
MAX_STAFF = {"Downtown": {d: 4 for d in DOWN}, "Studio": {"Mon": 4, "Tue": 4, "Wed": 4, "Thu": 4, "Fri": 4, "Sat": 5, "Sun": 5}, "1899 Penn": {d: 2 for d in DOWN}}
SHIFT_COST = {"Downtown": 40.0, "Studio": 15.0, "1899 Penn": 15.0}

def frange(a, b, step=0.5):
    x = a; out = []
    while x <= b + 1e-9: out.append(round(x, 2)); x += step
    return out

def end_ok(e, close):
    """Ryan: anyone still on the floor in the last hour closes with the team — a shift ends by close−60 min or at close+30."""
    return e <= close - 1.0 + 1e-9 or abs(e - close) < 1e-9

def candidates(cafe, dow, tier, cfg):
    o, c = cfg["open"][cafe][dow]; prep = o - cfg["prep"][cafe]; close = c + cfg["tail"][cafe]
    cands = []
    if cafe == "1899 Penn":
        for L in frange(6, 9): cands.append(("SL", "SL", prep, prep + L))
        for s in frange(prep, close - 6):
            for L in frange(6, 8.5):
                e = round(s + L, 2)
                if e <= close + 1e-9 and (e <= 14.0 + 1e-9 or abs(e - close) < 1e-9): cands.append(("PB", "PB", s, e))
        return [c_ for c_ in cands if end_ok(c_[3], close)]
    weekday_studio = cafe == "Studio" and dow in DOWN[:4]
    if weekday_studio:
        cands.append(("PL", "PL", prep, close))                                        # line person all day, closes the line
    elif cafe == "Studio":
        cands.append(("PL", "PL", prep, prep + 8))                                     # Fri–Sun: line person 8 h, baristas take the short shifts
    else:
        for L in frange(6, 8): cands.append(("PL", "PL", prep, prep + L))
    sl_start = o - 0.5; sl_max = 8.0 if weekday_studio else 10.0
    ends = {close} if close - sl_start <= sl_max + 1e-9 else set()
    for L in frange(8, sl_max):
        if sl_start + L <= close + 1e-9: ends.add(round(sl_start + L, 2))
    for e in sorted(ends): cands.append(("SL", "SL", sl_start, e))
    cl_max = 9 if (cafe == "Studio" and not weekday_studio) else 8.5
    for L in frange(6, cl_max):
        if close - L >= o - 0.5 - 1e-9: cands.append(("CL", "CL", round(close - L, 2), close))
    maxL = 6.5 if weekday_studio else 8
    b_min_start = 8.0 if (weekday_studio and dow != "Mon") else prep          # Studio Tue–Thu: the third person comes in at 8:00 (Ryan)
    for s in sorted(set(frange(prep, close - 6) + frange(o - 0.5, close - 6))):
        if s < b_min_start - 1e-9: continue
        for L in frange(6, maxL):
            if s + L <= close + 1e-9: cands.append(("B", "B", s, round(s + L, 2)))
    if weekday_studio:
        cands.append(("B", "B", 8.0, close))                                                                                  # third person 8:00–2:30 closes with the team
        if dow == "Mon": cands.append(("B", "B", 7.0, close)); cands.append(("B", "B", 6.75, close))
    return [c_ for c_ in cands if end_ok(c_[3], close)]

def solve(cafe, dow, tier, need, cfg, floor):
    o, c = cfg["open"][cafe][dow]; close = c + cfg["tail"][cafe]
    grid = sorted(need); cands = candidates(cafe, dow, tier, cfg)
    prob = pulp.LpProblem("shifts", pulp.LpMinimize)
    x = [pulp.LpVariable(f"x{i}", lowBound=0, upBound=2, cat="Integer") for i in range(len(cands))]
    mx = MAX_STAFF[cafe][dow]
    u = {g: pulp.LpVariable(f"u{int(g*2)}", lowBound=0) for g in grid}          # uncovered capacity need above the floor (soft; reported as 'Admin covers')
    def cost_of(fam, role, s, e):
        c_ = (e - s) * RATES[role] + SHIFT_COST[cafe]
        if fam == "B" and s < o: c_ += 2.0 * (o - s)                                # baristas are the later, shorter shifts unless the floor needs them at open
        return c_
    prob += pulp.lpSum(x[i] * cost_of(*cands[i]) for i in range(len(cands))) + 150.0 * pulp.lpSum(u.values())
    for g in grid:
        n = min(need[g], mx); fl = min(floor[g], mx)
        if fl > 0: prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if s <= g < e) >= fl          # floors are hard (2 open / 2 close, third person)
        if n > fl: prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if s <= g < e) + u[g] >= n
        prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if s <= g < e) <= mx
    fams = {}
    for i, (fam, role, s, e) in enumerate(cands): fams.setdefault(fam, []).append(i)
    if cafe == "1899 Penn":
        prob += pulp.lpSum(x[i] for i in fams["SL"]) == 1
        prob += pulp.lpSum(x[i] for i in fams["PB"]) >= 1
        # someone on the floor open-30 to close+30 (need already >= 1); if the lead leaves early, a PB must run to close
        prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if abs(e - close) < 1e-9) >= 1
    else:
        prob += pulp.lpSum(x[i] for i in fams["PL"]) == 1
        prob += pulp.lpSum(x[i] for i in fams["SL"]) == 1
        prob += pulp.lpSum(x[i] for i in fams["CL"]) <= 1
        for g in grid:
            if o - 0.5 <= g < close and need[g] > 0:
                prob += pulp.lpSum(x[i] for i, (fam, role, s, e) in enumerate(cands) if fam in ("SL", "CL") and s <= g < e) >= 1
    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    assert pulp.LpStatus[prob.status] == "Optimal", (cafe, dow, tier, pulp.LpStatus[prob.status])
    shifts = []
    for i, (fam, role, s, e) in enumerate(cands):
        for _ in range(int(round(x[i].value() or 0))): shifts.append({"role": role, "start": s, "end": e})
    order = {"PL": 0, "SL": 1, "PB": 2, "B": 3, "CL": 4}
    shifts.sort(key=lambda s: (s["start"], order[s["role"]]))
    return shifts

def fmt(t):
    h = int(t); mm = int(round((t - h) * 60)); return f"{h}:{mm:02d}"

def windows(grid, flags):
    w = []; cur = None
    for g, f in zip(grid, flags):
        if f and cur is None: cur = [g, g + 0.5]
        elif f: cur[1] = g + 0.5
        elif cur is not None: w.append(cur); cur = None
    if cur is not None: w.append(cur)
    return [f"{fmt(a)}–{fmt(b)}" for a, b in w]

def score(shifts, rows, sales, o, close):
    grid = [r["t"] for r in rows]; need = [r["need"] for r in rows]; tx = [r["txns_hr"] for r in rows]
    staffed = [sum(1 for s in shifts if s["start"] <= g < s["end"]) for g in grid]
    hours = sum(s["end"] - s["start"] for s in shifts); wages = sum((s["end"] - s["start"]) * RATES[s["role"]] for s in shifts)
    under = sum(max(0, n - a) for a, n in zip(staffed, need)) * 0.5
    load = [t / a if a > 0 else 0 for a, t in zip(staffed, tx)]
    return {"hours": round(hours, 2), "wages": round(wages, 2), "labor_pct": round(wages / sales * 100, 1) if sales else None, "splh": round(sales / hours, 1) if hours else None,
            "peak_staff": max(staffed), "under_person_hours": round(under, 1), "max_tph": round(max(load), 1),
            "solo_windows": windows(grid, [a == 1 and o <= g < close - 0.5 for g, a in zip(grid, staffed)]),
            "admin_cover_windows": windows(grid, [n > a for a, n in zip(staffed, need)]),
            "admin_helpful_windows": windows(grid, [(l >= 10.0 and o <= g < close) for g, l in zip(grid, load)]),
            "staffed": staffed}

def main():
    NV = json.load(open(os.path.join(DD, "need_v13.json")))
    cfg = {"open": NV["open"], "prep": NV["prep"], "tail": NV["tail"]}
    tiers = json.load(open(os.path.join(DD, "tiers.json"))); staffing = json.load(open(os.path.join(DD, "staffing.json")))
    T1 = json.load(open(os.path.join(DD, "templates_v11.json")))       # for Ryan's first-cut scores only
    out = {"rates": RATES, "cap": NV["cap"], "targets": {k: {"target": v[0], "ceiling": v[1], "ceiling_slow": v[1] + 0.02} for k, v in TARGET.items()}, "max_staff": MAX_STAFF, "cafes": {}}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        out["cafes"][cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            days_out = {}; dows = DOWN if cafe == "Studio" else DOWN[:5]
            needs = {d: {r["t"]: r["need"] for r in NV["cafes"][cafe]["tiers"][tier][d]["by_halfhour"]} for d in dows}
            floors = {d: {r["t"]: r["floor"] for r in NV["cafes"][cafe]["tiers"][tier][d]["by_halfhour"]} for d in dows}
            shared = None
            if cafe == "Downtown":      # Tue/Wed/Thu share one template, solved on the max of the three curves
                keys = sorted(set().union(*[needs[d].keys() for d in ("Tue", "Wed", "Thu")]))
                shared = solve(cafe, "Wed", tier, {g: max(needs[d].get(g, 0) for d in ("Tue", "Wed", "Thu")) for g in keys}, cfg, {g: max(floors[d].get(g, 0) for d in ("Tue", "Wed", "Thu")) for g in keys})
            for dow in dows:
                rows = NV["cafes"][cafe]["tiers"][tier][dow]["by_halfhour"]; o, c = cfg["open"][cafe][dow]; close = c + cfg["tail"][cafe]
                shifts = shared if (cafe == "Downtown" and dow in ("Tue", "Wed", "Thu")) else solve(cafe, dow, tier, needs[dow], cfg, floors[dow])
                sales = tiers[cafe]["typical_day"][tier][dow]["net"]; sc = score(shifts, rows, sales, o, close)
                rec = {"sales_typical": sales, "need_hours": NV["cafes"][cafe]["tiers"][tier][dow]["need_hours"], "peak_need": NV["cafes"][cafe]["tiers"][tier][dow]["peak_need"],
                       "shifts": [{"role": ROLE_NAMES[s["role"]], "code": s["role"], "start": fmt(s["start"]), "end": fmt(s["end"]), "start_h": s["start"], "end_h": s["end"],
                                   "hours": round(s["end"] - s["start"], 2), "rate": RATES[s["role"]], "wages": round((s["end"] - s["start"]) * RATES[s["role"]], 2)} for s in shifts],
                       "template": sc}
                act = staffing["staffing"][cafe].get(tier, {}).get(dow)
                if act: rec["actual_sampled"] = {"core_hours": act["core_hours_per_day"], "core_wages": act["core_wages_per_day"], "labor_pct": round(act["core_wages_per_day"] / sales * 100, 1), "n_days": act["n_days"]}
                v11 = T1["cafes"][cafe][f"{tier}_no10h"]["days"][dow]
                if "ryan_first_cut" in v11: rec["ryan_first_cut"] = v11["ryan_first_cut"]
                days_out[dow] = rec
            W = {"sales": 0.0, "hours": 0.0, "wages": 0.0, "act_h": 0.0, "act_w": 0.0, "ry_h": 0.0, "ry_w": 0.0, "n_act": 0}
            for d in days_out.values():
                W["sales"] += d["sales_typical"]; W["hours"] += d["template"]["hours"]; W["wages"] += d["template"]["wages"]
                if "actual_sampled" in d: W["act_h"] += d["actual_sampled"]["core_hours"]; W["act_w"] += d["actual_sampled"]["core_wages"]; W["n_act"] += 1
                if "ryan_first_cut" in d: W["ry_h"] += d["ryan_first_cut"]["hours"]; W["ry_w"] += d["ryan_first_cut"]["wages"]
            tg = TARGET[cafe]; ceil = tg[1] + (0.02 if tier == "slow" else 0.0)
            out["cafes"][cafe][tier] = {"days": days_out, "week": {
                "sales_typical": round(W["sales"]), "hours": round(W["hours"], 2), "wages": round(W["wages"]), "labor_pct": round(W["wages"] / W["sales"] * 100, 1), "splh": round(W["sales"] / W["hours"], 1),
                "labor_pct_minus10": round(W["wages"] / (W["sales"] * 0.9) * 100, 1), "labor_pct_plus10": round(W["wages"] / (W["sales"] * 1.1) * 100, 1),
                "actual_sampled_hours": round(W["act_h"], 1), "actual_sampled_wages": round(W["act_w"]), "actual_sampled_labor_pct": round(W["act_w"] / W["sales"] * 100, 1) if W["act_w"] else None,
                "ryan_first_cut_hours": round(W["ry_h"], 1), "ryan_first_cut_wages": round(W["ry_w"]), "ryan_first_cut_labor_pct": round(W["ry_w"] / W["sales"] * 100, 1) if W["ry_w"] else None,
                "target": tg[0], "ceiling": ceil, "target_wages": round(W["sales"] * tg[0]), "ceiling_wages": round(W["sales"] * ceil),
                "ceiling_trigger_sales": round(W["wages"] / ceil), "target_trigger_sales": round(W["wages"] / tg[0])}}
    # NMWA
    nm = tiers["NMWA"]["typical_day"]["slow"]; nm_days = {}
    for d in ["Thu", "Fri", "Sat", "Sun", "Wed"]:
        sh = [{"role": "FOH", "code": "FOH", "start": "9:00", "end": "15:00", "start_h": 9.0, "end_h": 15.0, "hours": 6.0, "rate": RATES["FOH"], "wages": 6 * RATES["FOH"]}]
        if d == "Wed": sh.append({**sh[0]})
        sales = nm.get(d, {}).get("net", 0)
        nm_days[d] = {"sales_typical": sales, "shifts": sh, "template": {"hours": 6.0 * len(sh), "wages": round(6 * RATES["FOH"] * len(sh), 2), "labor_pct": round(6 * RATES["FOH"] * len(sh) / sales * 100, 1) if sales else None},
                      "note": "second Wednesday of the month (community day): 2 × FOH 9:00–3:00" if d == "Wed" else ("first Sunday: add a second FOH 9:00–3:00" if d == "Sun" else "")}
    out["cafes"]["NMWA"] = {"slow": {"days": nm_days, "week": {"hours": 24.0, "wages": round(24 * RATES["FOH"]), "note": "4 × 6 h regular days; +6 h on each community day (≈ +3 h/wk averaged)"}}}
    json.dump(out, open(os.path.join(DD, "templates_v13.json"), "w"), indent=1)
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        print(f"\n=== {cafe} ===")
        for tier in ["slow", "mid", "busy"]:
            o = out["cafes"][cafe][tier]; w = o["week"]
            print(f" {tier:5} ${w['sales_typical']:,} {w['hours']:6.1f}h ${w['wages']:,} = {w['labor_pct']}% (ceil {w['ceiling']:.2f}; -10% {w['labor_pct_minus10']}%) | actual {w['actual_sampled_hours']}h {w['actual_sampled_labor_pct']}% | first cut {w['ryan_first_cut_hours']}h {w['ryan_first_cut_labor_pct']}%")
            for dow, d in o["days"].items():
                t = d["template"]
                print(f"    {dow} ${d['sales_typical']:,.0f}: {t['hours']:5.2f}h {t['labor_pct']:5.1f}% pk {t['peak_staff']} | " + "; ".join(f"{s['code']} {s['start']}-{s['end']}" for s in d["shifts"]) + f" | solo {t['solo_windows']} cover {t['admin_cover_windows']} adminhelps {t['admin_helpful_windows']}")

if __name__ == "__main__":
    main()
