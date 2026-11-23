#!/usr/bin/env python3
"""Labor Deep Dive v1.2 step 16 — three incentive options worked up and bench-tested.

Option 1  'Open-cafe pool'      : Beat Share 50% of (target − actual core wages) at Downtown + Studio only, cap 1.5% of their
                                  core sales; EVERY core hour (incl. 1899 / NMWA) shares the payout as a % of core wages.
Option 2  'Own-metric pool'     : Downtown + Studio as Option 1; 1899 adds a fixed $200 to the pool in any month its NI ≥ 5% of
                                  sales (the owner-relations / incentive-fee test); NMWA adds $100 in any month it runs at or
                                  under its staffing-rule hours. Same split.
Option 3  'Template beat'       : pool = 50% × Σ over all four cafes of (template core wages for the month's tier − actual core
                                  wages), cap 1.5% of all-cafe core sales. Every cafe's template is achievable by construction.
Option 4  'On-template + beat'  : fixed $ for every cafe-month run at or under its template (+1% tolerance): $300 Downtown, $300 Studio,
                                  $150 1899, $75 NMWA — plus 50% of any beat below template, cap 1.5% of all-cafe sales. Pays for executing the
                                  plan, pays more for beating it.
Bench tests: (a) backtest on Apr–Aug 2026 actual months (wages / sales from the timecard store; 1899 NI from the model);
(b) 2,000-run 2027 simulation on the tier calendar with execution noise. Output: data/labor_deep_dive/incentive_v12.json
"""
import os, json, random
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
WPM = 52 / 12
TARGET = {"Downtown": 0.32, "Studio": 0.30, "1899 Penn": 0.35}
CAFES3 = ["Downtown", "Studio", "1899 Penn"]
FK = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}
CAP_PCT = 0.015; SHARE = 0.5; FIX_1899 = 200.0; FIX_NMWA = 100.0
ON_TMPL = {"Downtown": 300.0, "Studio": 300.0, "1899 Penn": 150.0, "NMWA": 75.0}; TOL = 1.01
NMWA_TEMPLATE_WEEK = {"hours": 24.0, "wages": 24.0 * 21.25}     # 4 × 5 h regular days + ~4 h community-day share per week at the FOH rate
NMWA_SALES_WEEK = 1060.0

def payouts(rows, nmwa=None):
    """rows: list per cafe {cafe, sales, wages, tmpl_wages, ni_ok}. Returns dict option -> payout $ (month)."""
    dt_st = [r for r in rows if r["cafe"] in ("Downtown", "Studio")]
    S2 = sum(r["sales"] for r in dt_st)
    beat2 = sum(TARGET[r["cafe"]] * r["sales"] - r["wages"] for r in dt_st)
    o1 = min(max(0.0, SHARE * beat2), CAP_PCT * S2)
    o2 = o1 + (FIX_1899 if any(r["cafe"] == "1899 Penn" and r.get("ni_ok") for r in rows) else 0.0) + (FIX_NMWA if (nmwa and nmwa.get("rule_ok")) else 0.0)
    S_all = sum(r["sales"] for r in rows) + (nmwa["sales"] if nmwa else 0.0)
    beat_t = sum(r["tmpl_wages"] - r["wages"] for r in rows) + ((nmwa["tmpl_wages"] - nmwa["wages"]) if nmwa else 0.0)
    o3 = min(max(0.0, SHARE * beat_t), CAP_PCT * S_all)
    fixed = sum(ON_TMPL[r["cafe"]] for r in rows if r["wages"] <= r["tmpl_wages"] * TOL) + (ON_TMPL["NMWA"] if (nmwa and nmwa["wages"] <= nmwa["tmpl_wages"] * TOL) else 0.0)
    o4 = min(fixed + max(0.0, SHARE * beat_t), CAP_PCT * S_all)
    return {"opt1": o1, "opt2": o2, "opt3": o3, "opt4": o4, "beat_dt_studio": beat2, "beat_template": beat_t}

def main():
    T = json.load(open(os.path.join(DD, "templates_v12.json")))
    mwb = json.load(open(os.path.join(DD, "monthly_wage_basis.json")))
    fin = json.load(open(os.path.join(DATA, "fin_2026_v17.json")))
    tiers = json.load(open(os.path.join(DD, "tiers.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    MN = {4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug"}
    def tier_of_month(cafe, mon_num): return tiers[cafe]["month_membership"][str(mon_num)]["dominant"]
    def tmpl_month(cafe, tier, weeks): return T["cafes"][cafe][tier]["week"]["wages"] * weeks
    out = {"definitions": {"opt1": "Open-cafe pool — DT+Studio beat share 50%, cap 1.5% of DT+Studio sales; all core hours share",
                           "opt2": "Own-metric pool — Option 1 + 1899 adds $200 when NI ≥ 5% of sales + NMWA adds $100 when at/under staffing-rule hours",
                           "opt3": "Template beat — 50% of Σ(template wages − actual core wages) over all cafes, cap 1.5% of all-cafe core sales",
                           "opt4": "On-template + beat — $300/$300/$150/$75 per cafe-month at or under template (1% tolerance) + 50% of the beat, cap 1.5%",
                           "split": "payout ÷ total core wages of all participants = bonus %; each person receives that % of own core wages"}}
    # ---------- backtest Apr–Aug 2026 ----------
    bt = []
    for mnum, mon in MN.items():
        rows = []
        for cafe in CAFES3:
            r = [x for x in mwb if x["cafe"] == cafe and x["month"] == mon][0]
            tier = tier_of_month(cafe, mnum); days = fin["cafes"][FK[cafe]][mon]["days"]
            weeks = days / 7.0
            f = fin["cafes"][FK[cafe]][mon]
            rows.append({"cafe": cafe, "tier": tier, "sales": r["core_sales"], "wages": r["core_wages_full"], "core_pct": r["core_wage_pct"],
                         "tmpl_wages": round(tmpl_month(cafe, tier, weeks)), "ni_ok": (f["noi"] / f["core"]) >= 0.05 if f["core"] else False, "ni_pct": round(f["noi"] / f["core"] * 100, 1) if f["core"] else None})
        p = payouts(rows)
        wages_all = sum(r["wages"] for r in rows)
        bt.append({"month": mon, "cafes": rows, **{k: round(v) for k, v in p.items()},
                   "bonus_pct_of_wages": {k: round(p[k] / wages_all * 100, 2) for k in ("opt1", "opt2", "opt3", "opt4")}})
    out["backtest_2026"] = bt
    out["backtest_summary"] = {k: {"months_paying": sum(1 for b in bt if b[k] > 0), "total": round(sum(b[k] for b in bt)), "avg_month": round(sum(b[k] for b in bt) / len(bt))} for k in ("opt1", "opt2", "opt3", "opt4")}
    # ---------- 2027 simulation ----------
    CAL = {"Downtown": ["slow", "mid", "mid", "busy", "busy", "busy", "mid", "slow", "mid", "mid", "slow", "slow"],
           "Studio":   ["slow", "slow", "mid", "busy", "busy", "busy", "slow", "slow", "slow", "mid", "slow", "slow"],
           "1899 Penn": ["slow", "slow", "mid", "mid", "mid", "busy", "busy", "mid", "mid", "mid", "mid", "slow"]}
    out["calendar_2027"] = CAL
    random.seed(11); N = 2000
    res = {k: [] for k in ("opt1", "opt2", "opt3", "opt4")}; months_pay = {k: [] for k in res}; saving = []
    base_pct = {c: sum(r["core_wages_full"] for r in mwb if r["cafe"] == c) / sum(r["core_sales"] for r in mwb if r["cafe"] == c) for c in CAFES3}
    for it in range(N):
        tot = {k: 0.0 for k in res}; mp = {k: 0 for k in res}; sav = 0.0
        for mi in range(12):
            rows = []
            for cafe in CAFES3:
                tier = CAL[cafe][mi]; w = T["cafes"][cafe][tier]["week"]
                S = w["sales_typical"] * WPM * random.gauss(1.0, 0.06)
                real_pct = w["labor_pct"] / 100 + 0.005 + random.gauss(0, 0.015)
                wages = real_pct * S
                ni_ok = (cafe == "1899 Penn") and (tier == "busy") and random.random() < 0.6
                rows.append({"cafe": cafe, "sales": S, "wages": wages, "tmpl_wages": w["wages"] * WPM, "ni_ok": ni_ok})
                sav += base_pct[cafe] * S - wages
            nm_h = NMWA_TEMPLATE_WEEK["hours"] * WPM * (1 + random.gauss(0.03, 0.08))
            nmwa = {"sales": NMWA_SALES_WEEK * WPM, "wages": nm_h * 21.25, "tmpl_wages": NMWA_TEMPLATE_WEEK["wages"] * WPM, "rule_ok": nm_h <= NMWA_TEMPLATE_WEEK["hours"] * WPM * 1.02}
            p = payouts(rows, nmwa)
            for k in res:
                tot[k] += p[k]; mp[k] += 1 if p[k] > 0 else 0
        for k in res: res[k].append(tot[k]); months_pay[k].append(mp[k])
        saving.append(sav)
    q = lambda v, k: sorted(v)[int(k * (len(v) - 1))]
    out["simulation_2027"] = {"runs": N, "labor_saving_vs_aprAug_run_rate": {"p10": round(q(saving, .1)), "p50": round(q(saving, .5)), "p90": round(q(saving, .9))},
                              "options": {k: {"p10": round(q(res[k], .1)), "p50": round(q(res[k], .5)), "p90": round(q(res[k], .9)), "months_paying_p50": q(months_pay[k], .5),
                                              "share_of_saving_p50": round(q(res[k], .5) / max(1, q(saving, .5)) * 100, 1)} for k in res}}
    # per-person view (30 h/wk, $20.65 avg) at the p50 payout, using total core wages ≈ 447 h/wk × $20.65 × 52
    total_core_wages_yr = staffing_total = 447.2 * 20.65 * 52
    out["per_person_30h_year_p50"] = {k: round(q(res[k], .5) / total_core_wages_yr * 30 * 20.65 * 52) for k in res}
    # ---------- margins by tier under the v1.2 templates (same economics as v1.1: Apr–Aug COGS %, opex, tips, 1.12 load) ----------
    I1 = json.load(open(os.path.join(DD, "incentive_v11.json"))); econ = I1["cafe_economics"]; TIP = I1["assumptions"]["tip_rate"]; LOAD = 1.12
    def pl(cafe, S, wages, inc=0.0):
        e = econ[cafe]; gross = wages + e["admin_train_pct"] / 100 * S + inc; labor = LOAD * gross - TIP[cafe] * S; cogs = e["cogs_pct_aprAug"] / 100 * S; noi = S - cogs - labor - e["opex_month_avg"]
        return {"core_wage_pct": round(wages / S * 100, 1), "labor_pct": round(labor / S * 100, 1), "prime_pct": round((cogs + labor) / S * 100, 1), "noi": round(noi), "noi_pct": round(noi / S * 100, 1)}
    margins = {}
    for cafe in CAFES3:
        margins[cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            w = T["cafes"][cafe][tier]["week"]; S = w["sales_typical"] * WPM
            margins[cafe][tier] = {"core_sales_month": round(S),
                                   "actual_sampled": pl(cafe, S, w["actual_sampled_wages"] * WPM), "v12_template": pl(cafe, S, w["wages"] * WPM),
                                   "v12_template_opt4_month": pl(cafe, S, w["wages"] * WPM * 0.99, ON_TMPL[cafe] + 0.5 * 0.01 * w["wages"] * WPM),
                                   "at_target": pl(cafe, S, TARGET[cafe] * S)}
    out["margins_v12"] = margins
    noi_year = sum(margins[c][CAL[c][mi]]["v12_template"]["noi"] for c in CAFES3 for mi in range(12))
    out["profit_share_v12"] = {"open_cafe_noi_year_v12_templates": round(noi_year), "support_burden_year": I1["profit_share_sizing"]["support_burden_year_normalized"],
                               "consolidated_ni_year_v12": round(noi_year + I1["profit_share_sizing"]["support_burden_year_normalized"]),
                               "share_5pct": round(0.05 * (noi_year + I1["profit_share_sizing"]["support_burden_year_normalized"])), "share_10pct": round(0.10 * (noi_year + I1["profit_share_sizing"]["support_burden_year_normalized"]))}
    json.dump(out, open(os.path.join(DD, "incentive_v12.json"), "w"), indent=1)
    print("BACKTEST Apr–Aug 2026:")
    for b in bt:
        print(f"  {b['month']}: " + " | ".join(f"{r['cafe'][:3]} {r['tier'][:1]} {r['core_pct']}% tmpl ${r['tmpl_wages']:,} act ${r['wages']:,} NI {r['ni_pct']}%" for r in b["cafes"]) + f" -> opt1 ${b['opt1']:,} opt2 ${b['opt2']:,} opt3 ${b['opt3']:,} opt4 ${b['opt4']:,} (beat DT+St ${b['beat_dt_studio']:,}, beat tmpl ${b['beat_template']:,})")
    print(out["backtest_summary"]); print("SIM 2027:", json.dumps(out["simulation_2027"], indent=0)); print("per person:", out["per_person_30h_year_p50"])

if __name__ == "__main__":
    main()
