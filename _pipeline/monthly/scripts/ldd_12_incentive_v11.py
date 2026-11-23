#!/usr/bin/env python3
"""Labor Deep Dive v1.1 step 12 — incentive deep dive: structures, payouts, margins by tier, annual simulation,
retirement profit-share sizing, per-person view. All on the WAGE basis (core wages ÷ core sales).
Inputs: templates_v11.json (primary = no10h variant), monthly_wage_basis.json, fin_2026_v17.json, staffing.json, shifts.csv, tiers.json
Output: data/labor_deep_dive/incentive_v11.json
"""
import os, json, random
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WPM = 52 / 12
TARGET = {"Downtown": 0.32, "Studio": 0.30, "1899 Penn": 0.35}
CEIL = {"Downtown": 0.36, "Studio": 0.34, "1899 Penn": 0.40}
TIP_RATE = {"Downtown": 0.121, "Studio": 0.120, "1899 Penn": 0.104}
LOAD = 1.12
FK = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}
CAFES = ["Downtown", "Studio", "1899 Penn"]
PLAN_COGS = {"Downtown": 0.28, "Studio": 0.28, "1899 Penn": 0.30}      # Ryan's planning COGS (labor targets memo)

# ---------------- plan definitions ----------------
# payout(plan, cafe, sales, wages, hours) -> monthly payout $ for that cafe (pooling handled by caller)
def beat(cafe, sales, wages):
    return TARGET[cafe] * sales - wages          # $ saved below target (negative = over)

PLANS = {
    "A":  "Beat Share 50%, cap 1.5% of sales, pooled (recommended)",
    "A2": "Beat Share 50%, paid only when ≥ 2 pts under target (Ryan's draft gate)",
    "A3": "Beat Share 40%, cap 1.0% of sales, pooled (leaner sibling of A)",
    "B":  "Dollar-an-hour: $1.00 / $1.50 / $2.00 per core hour at ≤ target / −2 / −4 pts",
    "C":  "Gainshare 20% of the beat, no gate, no cap",
    "D":  "$8 per shift in every WEEK that hits target (weekly, no cap)",
    "E":  "Prime-cost gainshare: 50% of (prime target − actual prime) × sales, cap 2% of sales",
    "F":  "Beat Share 50%, cap 1.5%, PER-CAFE pools (no pooling across cafes)",
    "G":  "Quarterly Beat Share 50%, cap 1.5%, pooled — settled every 3 months",
}

def plan_payout_pooled(plan, rows):
    """rows: list of dicts per cafe {cafe, sales, wages, hours, prime_beat_dollars}. Returns total payout $ (month)."""
    S = sum(r["sales"] for r in rows); H = sum(r["hours"] for r in rows)
    B = sum(beat(r["cafe"], r["sales"], r["wages"]) for r in rows)
    if plan == "A":  return min(max(0.0, 0.5 * B), 0.015 * S)
    if plan == "A3": return min(max(0.0, 0.4 * B), 0.010 * S)
    if plan == "G":  return min(max(0.0, 0.5 * B), 0.015 * S)      # same math, settled quarterly (netting handled by caller)
    if plan == "C":  return max(0.0, 0.2 * B)
    if plan == "E":  return min(max(0.0, 0.5 * sum(r["prime_beat_dollars"] for r in rows)), 0.02 * S)
    if plan in ("A2", "B", "D", "F"):
        tot = 0.0
        for r in rows:
            b = beat(r["cafe"], r["sales"], r["wages"]); d = b / r["sales"]      # fraction under target (positive = beat)
            if plan == "A2": tot += 0.5 * b if d >= 0.02 - 1e-9 else 0.0
            if plan == "F":  tot += min(max(0.0, 0.5 * b), 0.015 * r["sales"])
            if plan == "B":
                if d >= 0.04 - 1e-9: tot += 2.0 * r["hours"]
                elif d >= 0.02 - 1e-9: tot += 1.5 * r["hours"]
                elif d >= -1e-9: tot += 1.0 * r["hours"]
            if plan == "D":  tot += 8.0 * (r["hours"] / 7.5) if d >= -1e-9 else 0.0     # ≈ every week at target (month-level proxy)
        return tot
    return 0.0

def main():
    T = json.load(open(os.path.join(DD, "templates_v11.json")))
    mwb = json.load(open(os.path.join(DD, "monthly_wage_basis.json")))
    fin = json.load(open(os.path.join(DATA, "fin_2026_v17.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    tiers = json.load(open(os.path.join(DD, "tiers.json")))
    sh = pd.read_csv(os.path.join(DD, "shifts.csv"))
    out = {"assumptions": {"weeks_per_month": WPM, "load_factor": LOAD, "tip_rate": TIP_RATE, "targets": TARGET, "ceilings": CEIL, "plan_cogs": PLAN_COGS, "plans": PLANS}}

    # ---------- cafe economics (Apr–Aug averages) ----------
    econ = {}
    for cafe in CAFES:
        f = fin["cafes"][FK[cafe]]
        months = ["Apr", "May", "Jun", "Jul", "Aug"]
        cogs_pct = sum(f[m]["cogs"] for m in months) / sum(f[m]["core"] for m in months)
        opex_avg = sum(f[m]["texp"] for m in months) / len(months)
        aug = [r for r in mwb if r["cafe"] == cafe and r["month"] == "Aug"][0]
        at_pct = (aug["admin_wages_full"] + aug["training_wages_full"]) / aug["core_sales"]
        econ[cafe] = {"cogs_pct_aprAug": round(cogs_pct * 100, 1), "cogs_pct_aug": round(f["Aug"]["cogs"] / f["Aug"]["core"] * 100, 1), "opex_month_avg": round(opex_avg),
                      "opex_month_aug": round(f["Aug"]["texp"]), "admin_train_pct": round(at_pct * 100, 2), "core_sales_aug": round(f["Aug"]["core"]),
                      "labor_booked_aug": round(f["Aug"]["labor"]), "noi_aug": round(f["Aug"]["noi"])}
    out["cafe_economics"] = econ
    # support burden (Commissary + SRP + Cafe Mgmt + Central), normalized = Apr–Jul avg (Aug Central had an $8.1K opex spike)
    sup = fin["support"]; months = ["Apr", "May", "Jun", "Jul"]
    burden = sum(sum(sup[k][m]["ni"] for k in sup) for m in months) / len(months)
    out["support_burden_month_normalized"] = round(burden); out["support_burden_aug"] = round(sum(sup[k]["Aug"]["ni"] for k in sup))

    # ---------- margin model by cafe × tier ----------
    def pl(cafe, sales, core_wages, incentive=0.0, cogs_pct=None):
        e = econ[cafe]; cp = (e["cogs_pct_aprAug"] / 100) if cogs_pct is None else cogs_pct
        gross = core_wages + e["admin_train_pct"] / 100 * sales + incentive
        labor = LOAD * gross - TIP_RATE[cafe] * sales
        cogs = cp * sales; opex = e["opex_month_avg"]
        noi = sales - cogs - labor - opex
        return {"labor_pl": round(labor), "labor_pct": round(labor / sales * 100, 1), "prime_pct": round((cogs + labor) / sales * 100, 1), "noi": round(noi), "noi_pct": round(noi / sales * 100, 1),
                "core_wage_pct": round(core_wages / sales * 100, 1), "incentive": round(incentive)}
    margins = {}
    for cafe in CAFES:
        margins[cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            w = T["cafes"][cafe][f"{tier}_no10h"]["week"]
            S = w["sales_typical"] * WPM; H = w["hours"] * WPM
            act_w = w["actual_sampled_wages"] * WPM; tmpl_w = w["wages"] * WPM; tmpl_obs = w["wages_obs_rate"] * WPM; tgt_w = TARGET[cafe] * S
            rows = {
                "Actual (sampled weeks of this tier)": pl(cafe, S, act_w),
                "v1.1 template at standard rates": pl(cafe, S, tmpl_w),
                "v1.1 template at current people's rates": pl(cafe, S, tmpl_obs),
                "Exactly at target": pl(cafe, S, tgt_w),
                "2 pts under target, Plan A (50% beat, cap)": pl(cafe, S, tgt_w - 0.02 * S, min(0.5 * 0.02 * S, 0.015 * S)),
                "2 pts under target, Plan B ($1.50/hr)": pl(cafe, S, tgt_w - 0.02 * S, 1.5 * H),
                "2 pts under target, Plan C (20% gainshare)": pl(cafe, S, tgt_w - 0.02 * S, 0.2 * 0.02 * S),
                "At target, Plan B ($1.00/hr)": pl(cafe, S, tgt_w, 1.0 * H),
                "At target, Plan D ($8/shift)": pl(cafe, S, tgt_w, 8.0 * H / 7.5),
                "2 pts under target, Plan A, planning COGS": pl(cafe, S, tgt_w - 0.02 * S, min(0.5 * 0.02 * S, 0.015 * S), cogs_pct=PLAN_COGS[cafe]),
            }
            margins[cafe][tier] = {"core_sales_month": round(S), "core_hours_month": round(H), "rows": rows}
    out["margins_by_tier"] = margins

    # ---------- scenario payouts (mid month, open cafes pooled) ----------
    scen = [("2 pts OVER target", 0.02), ("AT target", 0.0), ("1 pt under", -0.01), ("2 pts under", -0.02), ("3 pts under", -0.03), ("4 pts under", -0.04)]
    def rows_for(d, tier="mid", per_cafe_d=None):
        rows = []
        for cafe in CAFES:
            w = T["cafes"][cafe][f"{tier}_no10h"]["week"]; S = w["sales_typical"] * WPM; H = w["hours"] * WPM
            dd = d if per_cafe_d is None else per_cafe_d[cafe]
            wages = (TARGET[cafe] + dd) * S
            # prime beat: COGS held at plan; prime target = plan COGS + (target + admin/train) labor -> beat equals labor beat when COGS at plan
            rows.append({"cafe": cafe, "sales": S, "wages": wages, "hours": H, "prime_beat_dollars": -dd * S})
        return rows
    comp = {}
    S_mid = sum(r["sales"] for r in rows_for(0)); H_mid = sum(r["hours"] for r in rows_for(0))
    comp["mid_month"] = {"open_cafe_sales": round(S_mid), "open_cafe_core_hours": round(H_mid)}
    comp["scenarios"] = {}
    for label, d in scen:
        rowset = rows_for(d)
        comp["scenarios"][label] = {}
        for p in PLANS:
            pay = plan_payout_pooled(p, rowset)
            comp["scenarios"][label][p] = {"payout": round(pay), "pct_of_sales": round(pay / S_mid * 100, 2), "per_core_hour": round(pay / H_mid, 2),
                                           "per_person_month_30h": round(pay / H_mid * 30 * WPM), "per_person_month_20h": round(pay / H_mid * 20 * WPM),
                                           "employer_cost_incl_load": round(pay * LOAD), "slipstream_keeps": round(max(0.0, -d) * S_mid - pay)}
    # mixed-cafe scenario: Studio 3 under, Downtown at target, 1899 2 over (pooling vs per-cafe)
    mixed = rows_for(0, per_cafe_d={"Downtown": 0.0, "Studio": -0.03, "1899 Penn": 0.02})
    comp["mixed_example"] = {"description": "Studio 3 pts under, Downtown at target, 1899 2 pts over (mid month)",
                             **{p: round(plan_payout_pooled(p, mixed)) for p in PLANS}}
    out["plan_comparison"] = comp

    # ---------- annual simulation 2027 ----------
    # tier calendar per cafe (2026 dominant tiers Jan–Aug; Sep–Dec: projections tempered by the run-rate view; Dec slow per Ryan)
    CAL = {"Downtown": ["slow", "mid", "mid", "busy", "busy", "busy", "mid", "slow", "mid", "mid", "mid", "slow"],
           "Studio":   ["slow", "slow", "mid", "busy", "busy", "busy", "slow", "slow", "mid", "mid", "mid", "slow"],
           "1899 Penn": ["slow", "slow", "mid", "mid", "mid", "busy", "busy", "mid", "mid", "mid", "mid", "slow"]}
    out["annual_calendar_2027"] = CAL
    random.seed(7)
    # achievement model: realised core % = template % (at current people's rates) + noise; noise sd 1.5 pts, plus a 0.5-pt drift above template (real life)
    N = 2000
    ann = {p: [] for p in PLANS}; ann_saving = []; ann_wages = []
    base_actual_pct = {}   # Apr–Aug actual core wage % per cafe (what 'no change' would cost)
    for cafe in CAFES:
        rows = [r for r in mwb if r["cafe"] == cafe]
        base_actual_pct[cafe] = sum(r["core_wages_full"] for r in rows) / sum(r["core_sales"] for r in rows)
    out["base_actual_core_pct_aprAug"] = {c: round(v * 100, 1) for c, v in base_actual_pct.items()}
    months_detail = []
    for it in range(N):
        tot = {p: 0.0 for p in PLANS}; saving = 0.0; wages_year = 0.0; q_acc = []
        for mi in range(12):
            rowset = []
            for cafe in CAFES:
                tier = CAL[cafe][mi]; w = T["cafes"][cafe][f"{tier}_no10h"]["week"]
                S = w["sales_typical"] * WPM * random.gauss(1.0, 0.06)                  # sales noise ±6%
                tmpl_pct = w["labor_pct_obs_rate"] / 100                                  # at current people's rates (conservative)
                real_pct = tmpl_pct + 0.005 + random.gauss(0, 0.015)                      # +0.5 pt drift, sd 1.5 pts
                wages = real_pct * S; H = w["hours"] * WPM
                rowset.append({"cafe": cafe, "sales": S, "wages": wages, "hours": H, "prime_beat_dollars": beat(cafe, S, wages)})
                saving += base_actual_pct[cafe] * S - wages; wages_year += wages
            for p in PLANS:
                if p == "G":
                    q_acc.append(rowset)
                    if len(q_acc) == 3:
                        agg = [{"cafe": c, "sales": sum(r["sales"] for rs in q_acc for r in rs if r["cafe"] == c), "wages": sum(r["wages"] for rs in q_acc for r in rs if r["cafe"] == c),
                                "hours": sum(r["hours"] for rs in q_acc for r in rs if r["cafe"] == c), "prime_beat_dollars": 0} for c in CAFES]
                        tot[p] += plan_payout_pooled("G", agg); q_acc = []
                else:
                    tot[p] += plan_payout_pooled(p, rowset)
        for p in PLANS: ann[p].append(tot[p])
        ann_saving.append(saving); ann_wages.append(wages_year)
    def q(v, k): v = sorted(v); return v[int(k * (len(v) - 1))]
    out["annual_simulation"] = {"runs": N, "model": "12-month 2027 calendar by tier; sales noise ±6%; realised core % = template % at current rates + 0.5 pt drift ± 1.5 pts (sd)",
                                "labor_saving_vs_aprAug_run_rate": {"p10": round(q(ann_saving, .1)), "p50": round(q(ann_saving, .5)), "p90": round(q(ann_saving, .9))},
                                "core_wages_year_p50": round(q(ann_wages, .5)),
                                "plans": {p: {"p10": round(q(ann[p], .1)), "p50": round(q(ann[p], .5)), "p90": round(q(ann[p], .9)), "mean": round(sum(ann[p]) / N),
                                              "share_of_saving_p50": round(q(ann[p], .5) / max(1, q(ann_saving, .5)) * 100, 1)} for p in PLANS}}

    # ---------- headroom by tier (v1.1 templates) ----------
    head = {}
    for cafe in CAFES:
        head[cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            w = T["cafes"][cafe][f"{tier}_no10h"]["week"]; S = w["sales_typical"] * WPM
            head[cafe][tier] = {"core_sales_month": round(S), "actual_wages_month": round(w["actual_sampled_wages"] * WPM), "actual_pct": w["actual_sampled_labor_pct"],
                                "target_wages_month": round(TARGET[cafe] * S), "headroom_actual_vs_target": round(w["actual_sampled_wages"] * WPM - TARGET[cafe] * S),
                                "template_wages_month": round(w["wages"] * WPM), "template_pct": w["labor_pct"], "template_pct_obs_rate": w["labor_pct_obs_rate"],
                                "template_saving_vs_actual": round((w["actual_sampled_wages"] - w["wages"]) * WPM),
                                "beat_at_template_std": round(TARGET[cafe] * S - w["wages"] * WPM), "beat_at_template_obs": round(TARGET[cafe] * S - w["wages_obs_rate"] * WPM)}
    out["headroom_by_tier"] = head
    out["headroom_aug_total"] = round(sum(r["headroom_vs_target"] for r in mwb if r["month"] == "Aug"))
    out["headroom_aprjun_avg_total"] = round(sum(r["headroom_vs_target"] for r in mwb if r["month"] in ("Apr", "May", "Jun")) / 3)

    # ---------- retirement profit-share sizing ----------
    noi_target_month = sum(margins[c][t]["rows"]["Exactly at target"]["noi"] for c in CAFES for t in ["mid"])
    noi_target_year_cal = sum(margins[c][CAL[c][mi]]["rows"]["Exactly at target"]["noi"] for c in CAFES for mi in range(12))
    noi_tmpl_year_cal = sum(margins[c][CAL[c][mi]]["rows"]["v1.1 template at current people's rates"]["noi"] for c in CAFES for mi in range(12))
    cons_target_year = noi_target_year_cal + 12 * burden
    cons_tmpl_year = noi_tmpl_year_cal + 12 * burden
    # people & hours
    core = sh[sh.bucket == "core"]
    wk_hours = core.groupby(["tm"]).hours.sum() / len(staffing["weeks_sampled"])
    n_people = int((wk_hours > 2).sum())
    out["team"] = {"people_with_core_hours": n_people, "weekly_core_hours_p25": round(float(wk_hours.quantile(.25)), 1), "weekly_core_hours_p50": round(float(wk_hours.quantile(.5)), 1),
                   "weekly_core_hours_p75": round(float(wk_hours.quantile(.75)), 1), "total_core_hours_per_week": round(float(wk_hours.sum()), 1),
                   "multi_cafe_people": int((core.groupby("tm").cafe.nunique() > 1).sum())}
    out["profit_share_sizing"] = {"open_cafe_noi_month_at_target_mid": round(noi_target_month), "open_cafe_noi_year_at_target_calendar": round(noi_target_year_cal),
                                  "open_cafe_noi_year_template_current_rates": round(noi_tmpl_year_cal),
                                  "support_burden_year_normalized": round(12 * burden), "consolidated_ni_year_at_target": round(cons_target_year),
                                  "consolidated_ni_year_template_current_rates": round(cons_tmpl_year),
                                  "share_5pct_of_consolidated": round(0.05 * cons_target_year), "share_10pct_of_consolidated": round(0.10 * cons_target_year),
                                  "share_5pct_of_cafe_noi": round(0.05 * noi_target_year_cal), "share_10pct_of_cafe_noi": round(0.10 * noi_target_year_cal),
                                  "per_person_at_10pct_consolidated_30h": round(0.10 * cons_target_year / max(1.0, float(wk_hours.sum())) * 30),
                                  "note": "Calendar = 2027 tier calendar above; cafe NOI at exactly-target core wages with Apr–Aug COGS and opex; support burden = Apr–Jul average of Commissary+SRP+Cafe Mgmt+Central. NMWA excluded."}
    json.dump(out, open(os.path.join(DD, "incentive_v11.json"), "w"), indent=1)
    # ---- console
    print("ECON:", json.dumps(econ, indent=0)[:900]); print("support burden/mo normalized:", round(burden), " Aug:", out["support_burden_aug"])
    print("\nSCENARIOS (mid month, open cafes pooled, sales $%s, hours %s):" % (f"{S_mid:,.0f}", f"{H_mid:,.0f}"))
    for label, row in comp["scenarios"].items():
        print(f"  {label:18} " + " | ".join(f"{p}: ${v['payout']:>5,} ({v['pct_of_sales']:.2f}%, ${v['per_person_month_30h']}/mo@30h)" for p, v in row.items()))
    print("  mixed:", comp["mixed_example"])
    print("\nANNUAL SIM:", json.dumps(out["annual_simulation"], indent=0)[:1500])
    print("\nMARGINS (mid tier):")
    for c in CAFES:
        print(" ", c, {k: (v["labor_pct"], v["prime_pct"], v["noi_pct"]) for k, v in margins[c]["mid"]["rows"].items()})
    print("\nPROFIT SHARE:", json.dumps(out["profit_share_sizing"], indent=0)); print("TEAM:", out["team"])

if __name__ == "__main__":
    main()
