#!/usr/bin/env python3
"""Labor Deep Dive v1.2 step 17 — projections reality check (v1.8 vs 2025 actuals vs Jul–Aug run-rate), Sep–Dec tier
calendar on the run-rate case, the price of the Sep 5 rules (v1.1 lean → v1.2), and the v1.2 reference JSON.
Writes labor_deep_dive_2026-09.json (v1.2; keeps v1.1 copy) and data/labor_deep_dive/projections_check_v12.json.
"""
import os, json, shutil
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
J = lambda n: json.load(open(os.path.join(DD, n)))

def main():
    ref = os.path.join(DATA, "labor_deep_dive_2026-09.json"); cur = json.load(open(ref))
    v11p = os.path.join(DATA, "labor_deep_dive_2026-09_v1.1.json")
    if cur.get("version") == "1.1" and not os.path.exists(v11p):
        shutil.copy(ref, v11p)
    cur = json.load(open(v11p))
    T = J("templates_v12.json"); T1 = J("templates_v11.json"); INC = J("incentive_v12.json"); NV = J("need_v12.json"); tiers = J("tiers.json")
    # ---------- projections check ----------
    p = pd.read_csv(os.path.join(DD, "projections_daily_v18.csv")); p["m"] = p.d26.str[:7]
    p["d26"] = pd.to_datetime(p.d26)
    g = p.groupby(["cafe", "m"]).agg(s25=("s25", "sum"), proj=("proj26", "sum"), act=("act26", "sum")).reset_index()
    RUNRATE = {"Downtown": 1.00, "Studio": 0.86, "1899": None}        # Jul–Aug 2026 vs Jul–Aug 2025 (Downtown −1%/+1%, Studio −17%/−14%); 1899 has no comparable base
    check = {}
    for cafe in ["Downtown", "Studio", "1899"]:
        rows = []
        for _, r in g[g.cafe == cafe].iterrows():
            mon = r.m[5:]
            yoy = (r.act / r.s25 - 1) * 100 if (r.s25 and r.act) else None
            pv = (r.proj / r.s25 - 1) * 100 if r.s25 else None
            rr = r.s25 * RUNRATE[cafe] if (RUNRATE[cafe] and r.s25) else None
            rows.append({"month": r.m, "sales_2025": round(r.s25), "proj_v18": round(r.proj), "actual_2026": round(r.act) if r.act else None,
                         "yoy_actual_pct": round(yoy, 1) if yoy is not None else None, "proj_vs_2025_pct": round(pv, 1) if pv is not None else None,
                         "runrate_case": round(rr) if rr else None, "runrate_vs_proj_pct": round((rr / r.proj - 1) * 100, 1) if (rr and r.proj) else None})
        check[cafe] = rows
    # Sep–Dec weekly tiers on the run-rate case (2025 daily sales by 2026 date × factor)
    bands = {c: tiers[c]["bands_weekly"] for c in ["Downtown", "Studio", "1899 Penn"]}
    def tier_of(cafe, w):
        for t, (lo, hi) in bands[cafe].items():
            if lo <= w < (hi if hi and hi < 1e8 else 1e12): return t
        return "busy"
    cal = {}
    CK = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        q = p[(p.cafe == CK[cafe]) & (p.d26 >= "2026-09-01")].copy()
        q["wk"] = q.d26.dt.to_period("W-SUN").apply(lambda x: x.start_time.normalize())
        fac = RUNRATE[CK[cafe]]
        wk = q.groupby("wk").agg(s25=("s25", "sum"), proj=("proj26", "sum"), days=("s25", lambda s: (s > 0).sum())).reset_index()
        rows = []
        for _, r in wk.iterrows():
            if r.days < 3: continue
            full = 5 if cafe != "Studio" else 7
            scale = full / max(r.days, 1)
            rr = (r.s25 * fac * scale) if fac else None; pj = r.proj * scale
            rows.append({"week": str(r.wk.date()), "proj_weekly": round(pj), "proj_tier": tier_of(cafe, pj), "runrate_weekly": round(rr) if rr else None, "runrate_tier": tier_of(cafe, rr) if rr else None})
        cal[cafe] = rows
    # month-level tier recommendation Sep–Dec
    rec = {}
    for cafe, rows in cal.items():
        by_m = {}
        for r in rows:
            m = r["week"][:7]; by_m.setdefault(m, []).append(r)
        rec[cafe] = {m: {"proj_dominant": max(set(x["proj_tier"] for x in v), key=[x["proj_tier"] for x in v].count),
                         "runrate_dominant": (max(set(x["runrate_tier"] for x in v), key=[x["runrate_tier"] for x in v].count) if v[0]["runrate_tier"] else None),
                         "weeks": len(v)} for m, v in by_m.items()}
    json.dump({"monthly": check, "weekly_sep_dec": cal, "tier_by_month_sep_dec": rec, "runrate_factor": RUNRATE}, open(os.path.join(DD, "projections_check_v12.json"), "w"), indent=1)
    # ---------- price of the rules ----------
    price = {}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        price[cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            w2 = T["cafes"][cafe][tier]["week"]; w1 = T1["cafes"][cafe][f"{tier}_no10h"]["week"]
            # v1.1 hours re-priced at v1.2 planning rates for a like-for-like: use wages_obs_rate proxies -> simpler: compare hours and wages at each version's basis
            price[cafe][tier] = {"v11_hours": w1["hours"], "v11_wages_std": w1["wages"], "v11_pct_std": w1["labor_pct"], "v11_pct_roster": w1["labor_pct_obs_rate"],
                                 "v12_hours": w2["hours"], "v12_wages_plan": w2["wages"], "v12_pct_plan": w2["labor_pct"], "v12_pct_std": w2["labor_pct_std_rates"],
                                 "delta_hours": round(w2["hours"] - w1["hours"], 1), "delta_wages_std_basis": round(w2["wages_std_rates"] - w1["wages"]),
                                 "delta_month_std_basis": round((w2["wages_std_rates"] - w1["wages"]) * 52 / 12)}
    out = {"generated": "2026-09-05", "version": "1.2", "supersedes": "1.1 (Sep 5) and 1.0 (Sep 4)",
           "scope": cur["scope"],
           "rules_sep5": {"open_close": "Downtown & Studio: 2 people open-to-close — kitchen 45 min before open, barista/lead 30 min before, two closers to 30 min after close",
                          "studio_midday": "Studio Mon–Thu: 3 people 8:00–12:30 (opener leaves 12:30, third arrives 8:00); slow months: 2 core + admin on admin days",
                          "managed": "1899 and NMWA may stagger / go solo in slow months", "min_shift": "6 hours", "sl_max": "9 h (10 h open-to-close lead allowed on 9-hour+ days as the cheapest way to meet two closers)",
                          "admin_days": NV["admin_days"], "admin_window": "9:00–14:00 (Studio Mon–Thu: to close)",
                          "rates": T["rates"], "targets": T["targets"], "slow_month_ceiling": "+2 pts (Downtown 38 / Studio 36 / 1899 42)"},
           "changes_from_v1_1": ["Floors rebuilt to the Sep 5 rules; 6-hour minimum shifts; planning rates = Aug 2026 roster averages by role ($22.25 SL / $19.75 PL / $20.25 PB / $20.00 barista & closer)",
                                 "Admin-day sets: on Gee/Colin/Eli/Claire days the slow-tier template drops one core body inside the admin window",
                                 "Projections v1.8 checked against 2025 actuals and the Jul–Aug run-rate; Sep–Dec tiers re-stated on the run-rate case",
                                 "Incentive: four options bench-tested (backtest Apr–Aug 2026 + 2027 simulation); target-based beat shares do not pay under the new rules — template-based ones do"],
           "definitions": cur["definitions"], "tiers": cur["tiers"], "tier_band_validation": cur["tier_band_validation"], "capacity_planning": NV["cap"],
           "projections_check": {"monthly": check, "tier_by_month_sep_dec": rec, "runrate_factor": RUNRATE},
           "price_of_rules": price, "august_2026": cur["august_2026"], "pl_bridge_validation": cur["pl_bridge_validation"], "targets": T["targets"], "standard_rates": T["rates"],
           "templates": {c: {t: {"week": T["cafes"][c][t]["week"], "week_no_admin": T["cafes"][c][t]["week_no_admin"],
                                 "days": {d: {"sales_typical": dd["sales_typical"], "admin_day": dd["admin_day"],
                                              "sets": {s: {"hours": v["template"]["hours"], "wages": v["template"]["wages"], "labor_pct": v["template"]["labor_pct"], "peak_staff": v["template"]["peak_staff"],
                                                           "solo_windows": v["template"]["solo_windows"], "admin_cover_windows": v["template"]["admin_cover_windows"],
                                                           "shifts": [{k: sh[k] for k in ("role", "code", "start", "end", "hours", "rate", "wages")} for sh in v["shifts"]]} for s, v in dd["sets"].items()}}
                                          for d, dd in T["cafes"][c][t]["days"].items()}} for t in ["slow", "mid", "busy"]} for c in T["cafes"]},
           "incentive_v12": INC, "incentive_v11_reference": cur["incentive"], "data_notes": cur["data_notes"]}
    json.dump(out, open(ref, "w"), indent=1)
    print("projection check (monthly):")
    for cafe, rows in check.items():
        for r in rows:
            if r["month"] >= "2026-07": print(f"  {cafe:9} {r['month']} 2025 ${r['sales_2025']:,} proj ${r['proj_v18']:,} ({r['proj_vs_2025_pct']}%) actual {r['actual_2026']} ({r['yoy_actual_pct']}%) run-rate {r['runrate_case']} ({r['runrate_vs_proj_pct']}% vs proj)")
    print("tiers Sep–Dec:", json.dumps(rec, indent=0))
    print("price of rules:", json.dumps(price, indent=0)[:1500])

if __name__ == "__main__":
    main()
