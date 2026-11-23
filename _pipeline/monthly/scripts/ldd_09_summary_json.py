#!/usr/bin/env python3
"""Labor Deep Dive step 9 — assemble the reference JSON the monthly report can cite:
tier definitions + month membership, capacity assumptions, per-template labor %, the P&L bridge,
levers, incentive sizing. Output: _pipeline/monthly/data/labor_deep_dive_2026-09.json
"""
import os, json
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data")
DD = os.path.join(DATA, "labor_deep_dive")
J = lambda n: json.load(open(os.path.join(DD, n)))

def main():
    tiers, curves, staffing, cap, nv, T, inc, rate = (J("tiers.json"), J("curves.json"), J("staffing.json"), J("capacity.json"),
                                                      J("need_vs_actual.json"), J("templates.json"), J("incentive.json"), J("rate_lever.json"))
    mwb = J("monthly_wage_basis.json")
    fin = json.load(open(os.path.join(DATA, "fin_2026_v17.json")))
    tx = pd.read_csv(os.path.join(DD, "txns_2026.csv")); tx["m"] = tx.date.str[:7]
    sh = pd.read_csv(os.path.join(DD, "shifts.csv"))
    aug_tx = tx[tx.m == "2026-08"].groupby("cafe").agg(net=("net", "sum"), tips=("tip", "sum"), txns=("net", "size"))
    aug_sh = sh[(sh.date >= "2026-08-01") & (sh.date <= "2026-08-31")]
    FK = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}
    TGT = T["targets"]
    august = {}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        c = aug_sh[aug_sh.cafe == cafe]; core = c[c.bucket == "core"]
        core_h = core.hours.sum(); core_w = core.wages.sum(); all_h = c[c.bucket != "event"].hours.sum(); all_w = c[c.bucket != "event"].wages.sum()
        sales_model = fin["cafes"][FK[cafe]]["Aug"]["core"]; walk = float(aug_tx.loc[cafe, "net"]); tips = float(aug_tx.loc[cafe, "tips"])
        pl = fin["cafes"][FK[cafe]]["Aug"]["labor"]
        rate_g = core_w / core_h
        august[cafe] = {
            "core_sales_model": round(sales_model), "walkin_net": round(walk), "tips": round(tips), "tip_rate": round(tips / walk * 100, 1),
            "core_hours": round(core_h, 1), "core_wages": round(core_w), "admin_wages": round(c[c.bucket == "admin"].wages.sum()), "training_wages": round(c[c.bucket == "training"].wages.sum()),
            "core_wage_pct_of_core_sales": round(core_w / sales_model * 100, 1), "target_pct": TGT[cafe]["target"] * 100,
            "headroom_core_wages_vs_target": round(core_w - TGT[cafe]["target"] * sales_model),
            "splh_core_walkin": round(walk / core_h, 1), "gross_rate_core": round(rate_g, 2), "tips_per_core_hour": round(tips / core_h, 2),
            "splh_needed_for_target": round(rate_g / TGT[cafe]["target"], 1),
            "pl_labor_booked": round(pl), "pl_labor_pct_booked": round(pl / sales_model * 100, 1),
            "pl_labor_bridge_est": round(1.12 * all_w - tips), "pl_labor_pct_bridge": round((1.12 * all_w - tips) / sales_model * 100, 1),
            "loaded_rate_booked_per_hour": round(pl / all_h, 2), "loaded_rate_bridge_per_hour": round((1.12 * all_w - tips) / all_h, 2),
        }
    # bridge validation Apr–Aug
    bridge = {}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        rows = [r for r in mwb if r["cafe"] == cafe]
        est = sum(r["pl_labor_est"] for r in rows); act = sum(r["pl_labor_actual"] for r in rows)
        bridge[cafe] = {"months": [r["month"] for r in rows], "pl_est_sum": est, "pl_actual_sum": act, "error_pct": round((est - act) / act * 100, 1),
                        "by_month": {r["month"]: {"core_wage_pct": r["core_wage_pct"], "pl_est": r["pl_labor_est"], "pl_actual": r["pl_labor_actual"], "headroom": r["headroom_vs_target"]} for r in rows}}
    # tier calendar
    PROJ = {"Downtown": {"Sep": 36805 / 21 * 5, "Oct": 38329 / 21 * 5, "Nov": 32001 / 19 * 5, "Dec": 26637 / 16 * 5},
            "Studio": {"Sep": 49650 / 30 * 7, "Oct": 56000 / 31 * 7, "Nov": 47925 / 29 * 7, "Dec": 40875 / 30 * 7},
            "1899 Penn": {"Sep": 16375 / 21 * 5, "Oct": 17600 / 21 * 5, "Nov": 14550 / 19 * 5, "Dec": 12400 / 16 * 5}}
    YOY = {"Downtown": 0.96, "Studio": 0.86, "1899 Penn": 1.0}     # Jul–Aug 2026 YoY run-rate applied to 2025 base (1899 has no comparable 2025)
    BASE25 = {"Downtown": {"Sep": 1597 * 5, "Oct": 1680 * 5, "Nov": 1561 * 5, "Dec": 1396 * 5},
              "Studio": {"Sep": 1293 * 5 + 2565 * 2, "Oct": 1430 * 5 + 2784 * 2, "Nov": 1191 * 5 + 2517 * 2, "Dec": 990 * 5 + 2031 * 2}}
    def tier_of(cafe, w):
        for t, (lo, hi) in tiers[cafe]["bands_weekly"].items():
            if lo <= w < (hi if hi else 1e12): return t
        return "busy"
    calendar = {}
    MN = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug"}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        cal = {}
        for m, v in tiers[cafe]["month_membership"].items():
            cal[MN[int(m)]] = {"basis": "actual 2026", "avg_weekly": v["avg_weekly_scaled"], "weeks_by_tier": v["tiers"], "tier": v["dominant"]}
        for mon in ["Sep", "Oct", "Nov", "Dec"]:
            w = PROJ[cafe][mon]
            rec = {"basis": "projection v1.8", "avg_weekly": round(w), "tier": tier_of(cafe, w)}
            if cafe in BASE25:
                w2 = BASE25[cafe][mon] * YOY[cafe]
                rec["yoy_view_weekly"] = round(w2); rec["yoy_view_tier"] = tier_of(cafe, w2)
            if mon == "Dec":
                rec["ryan_override"] = "slow (holiday weeks Dec 21–Jan 3 slow; Dec 1–18 may run mid)"
            cal[mon] = rec
        calendar[cafe] = cal
    out = {
        "generated": "2026-09-04", "version": "1.0",
        "scope": "Slipstream Labor Deep Dive — schedule templates by demand tier (Downtown, Studio, 1899 Penn; NMWA for completeness)",
        "definitions": {"core_labor": "Square timecards excluding job titles 'Manager' (admin) and 'Training'; 'Event' excluded",
                        "wage_basis": "gross guaranteed wages (Square hourly_rate × hours) as % of core (ex-catering) sales — the basis of the 30/32/35 targets",
                        "pl_bridge": "P&L labor ≈ 1.12 × gross wages (core+admin+training) − credit-card tips; validated Apr–Aug 2026 within 0.4–2.3% per cafe",
                        "capacity": f"planning capacity {T['cap']} transactions per person-hour (observed p90 of busy-hour throughput; teams have sustained 16–21)",
                        "need_rule": "need(h) = max(floor, ceil(txns/h ÷ 14)); floors: 2 at open through lunch where a food/production load exists (DT, Studio), 1 in the final 90 min; Studio Mon 3 at open; 1899 2 through 1:30p then 1; NMWA 1 (2 on community days)"},
        "tiers": {c: {"bands_weekly_walkin": tiers[c]["bands_weekly"], "typical_day": tiers[c]["typical_day"], "calendar": calendar[c]} for c in ["Downtown", "Studio", "1899 Penn"]},
        "capacity_observed": {c: cap[c]["busy_hours"] for c in cap},
        "august_2026": august, "pl_bridge_validation": bridge,
        "levers": {"hours_placement_monthly_headroom_by_tier": inc["headroom_by_tier"], "rate_mix": rate},
        "templates": {c: {t: {"week": T["cafes"][c][t]["week"], "days": {d: {"hours": dd["template"]["hours"], "wages": dd["template"]["wages"], "labor_pct": dd["template"]["labor_pct"],
                                                                              "sales_typical": dd["sales_typical"], "shifts": dd["shifts"]} for d, dd in T["cafes"][c][t]["days"].items()}}
                          for t in T["cafes"][c]} for c in T["cafes"]},
        "standard_rates": T["rates"], "targets": TGT,
        "incentive": {"headroom_aug_total": inc["headroom_aug_total"], "headroom_aprjun_avg_total": inc["headroom_aprjun_avg_total"],
                      "plan_comparison_mid_tier": inc["plan_comparison_mid_tier"], "margin_view_august": inc["margin_view_august"],
                      "annual_profit_share_sizing": inc["annual_profit_share_sizing"]},
        "data_notes": {"timecards_harvested": len(json.load(open(os.path.join(DD, "timecards_raw.json")))["timecards"]), "timecards_cafes": int(len(sh)) + staffing["notes"]["dropped_zero_length"], "timecards_used": int(len(sh)), "weeks_sampled": staffing["weeks_sampled"],
                       "dropped_zero_length": staffing["notes"]["dropped_zero_length"], "zero_rate_imputed": staffing["notes"]["zero_rate_imputed"],
                       "capped_ruleA_16h": staffing["notes"]["capped_count_ruleA"], "capped_ruleB_past_close": staffing["notes"]["capped_count_ruleB"],
                       "capped_wages_avoided": staffing["notes"]["capped_wages_avoided"], "capped_detail": staffing["notes"]["capped"],
                       "transactions_rows_kept": int(len(tx)), "projections_file": "Projections/2026 Projections v1.8.xlsx (monthly summary identical to v1.7)"},
    }
    json.dump(out, open(os.path.join(DATA, "labor_deep_dive_2026-09.json"), "w"), indent=1)
    print(json.dumps(august, indent=1)); print(json.dumps(bridge, indent=1)[:1500]); print(json.dumps(calendar, indent=1))

if __name__ == "__main__":
    main()
