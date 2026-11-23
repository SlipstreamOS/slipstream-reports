#!/usr/bin/env python3
"""Labor Deep Dive v1.3 step 22 — what the templates free up, and what the pay curve / incentive would cost against it.
Inputs: backtest_v13.json, pay_review_v13.json, incentive_v12.json, fin_2026_v17.json. Output: affordability_v13.json"""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__)); DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
J = lambda n: json.load(open(os.path.join(DD, n)))

def main():
    BT = J("backtest_v13.json"); PAY = J("pay_review_v13.json"); INC = J("incentive_v12.json"); fin = json.load(open(os.path.join(DATA, "fin_2026_v17.json"))); PP = J("peer_pay_cost_v13.json")
    FOUNDER_SALARY = 120000
    cafes = BT["cafes"]; out = {}
    # template effect
    sav = {c: {"12mo": cafes[c]["totals"]["labor_saving"], "sep_dec_2025": sum(x["labor_saving"] for x in cafes[c]["months"] if x["ym"] < "2026"),
               "jan_aug_2026": sum(x["labor_saving"] for x in cafes[c]["months"] if x["ym"] >= "2026")} for c in cafes}
    tot = {k: sum(sav[c][k] for c in sav) for k in ("12mo", "sep_dec_2025", "jan_aug_2026")}; tot["annualized_2026"] = round(tot["jan_aug_2026"] * 12 / 8)
    out["template_saving"] = {"by_cafe": sav, "total": tot}
    # COGS lever
    core12 = BT["three_cafes"]["totals"]["core_sales"]; core26 = sum(fin["cafes"][c][m]["core"] for c in ("Downtown", "Studio", "1899") for m in fin["cafes"]["Downtown"])
    cogs26 = sum(fin["cafes"][c][m]["cogs"] for c in ("Downtown", "Studio", "1899") for m in fin["cafes"]["Downtown"])
    by_cafe_26 = {c: round(sum(fin["cafes"][c][m]["cogs"] for m in fin["cafes"][c]) / sum(fin["cafes"][c][m]["core"] for m in fin["cafes"][c]) * 100, 1) for c in ("Downtown", "Studio", "1899")}
    out["cogs"] = {"core_sales_12mo": core12, "per_point_year": round(core12 / 100), "blended_2026_jan_aug_pct": round(cogs26 / core26 * 100, 1), "by_cafe_2026": by_cafe_26,
                   "gain_vs_2026_blend": {f"{s:g}": round((cogs26 / core26 - s / 100) * core12) for s in (27, 28.5, 30)},
                   "note": "2026 core COGS excludes catering COGS; Jan–Aug blended for the three cafes. Gain = (blended − scenario) × 12-month core sales"}
    # cost side
    s1, sf = PAY["cost_step1"], PAY["cost_of_proposed"]; sim = INC["simulation_2027"]["options"]["opt4"]
    out["costs"] = {"pay_step1_year": s1["year"], "pay_full_year": sf["year"], "incentive_opt4_p50": round(sim["p50"]), "incentive_opt4_p90": round(sim["p90"]),
                    "incentive_per_hour_equiv": INC.get("per_person_30h_year_p50", {}).get("opt4")}
    # headroom ladder: 12-month, three cafes, templates, COGS scenario, before/after support
    base = BT["three_cafes"]["totals"]["core_basis"]; sup = BT["three_cafes"]["support_burden_2026_avg_month"] * 12
    ladder = {}
    for sc in ("27", "28.5", "30", "actual_cogs"):
        b = base[sc]["template_labor"]; a = base[sc]["actual_labor"]
        ladder[sc] = {"actual_labor_before_support": a, "template_before_support": b, "template_after_support": b + sup,
                      "after_step1": b + sup - s1["year"], "after_step1_incentive": b + sup - s1["year"] - round(sim["p50"]),
                      "after_full_curve": b + sup - sf["year"], "after_full_curve_incentive": b + sup - sf["year"] - round(sim["p50"])}
    out["headroom_12mo"] = {"support_burden_year": sup, "ladder": ladder,
                            "note": "Three cafes only (NMWA opened Jul 2026 and is excluded); core basis = core sales − COGS at the scenario % − labor − operating expenses; support = Commissary + SRP + Cafe Mgmt + Central at the 2026 Jan–Aug average"}
    # ---- pay scenarios on the template wage base (12 months, three cafes): uplift % from the August roster × (template core + admin/training wages) × load
    BR = 1.12; t3 = BT["three_cafes"]["totals"]
    wage_base = sum(BT["cafes"][c]["totals"]["template_core_wages"] + BT["cafes"][c]["totals"]["admin_training"] for c in BT["cafes"])
    core_w = sum(BT["cafes"][c]["totals"]["template_core_wages"] for c in BT["cafes"]); hrs = sum(BT["cafes"][c]["totals"]["template_core_hours"] for c in BT["cafes"])
    scen = {"planning": {"label": "Templates at today's rates", "uplift_pct": 0.0, "avg_rate": round(core_w / hrs, 2)}}
    for kk, lab in [("parity_floors", "Parity floors (SL ≥ $22.50, ALL ≥ $25.50)"), ("full_curve", "Midpoint curve ($18.50 / $20.50 / $23.50 / $27)"), ("curve_a", "Steeper curve, band tops ($18.50 / $21 / $24.50 / $28)"),
                    ("peer_avg_curve", "Peer-average curve ($18.50 / $21.50 / $26 / $29.50)"), ("peer_avg_curve_plus5", "Peer-average curve + 5%"), ("peer", "Flat peer all-in ($23 / $24 / $25.50 / $27)"), ("peer_plus5", "Flat peer + 5%")]:
        u = PP["scenarios"][kk]["uplift_pct"]; scen[kk] = {"label": lab, "uplift_pct": u, "avg_rate": round(core_w / hrs * (1 + u / 100), 2), "roster_avg_rate": PP["scenarios"][kk]["avg_rate"], "rates": PP["scenarios"][kk]["rates"], "roster_cost_year_aug_hours": PP["scenarios"][kk]["cost_year"]}
    scen["planning"]["roster_avg_rate"] = PP["avg_rate_now"]; out["peer_avg_all_in"] = PP["peer_avg_all_in"]; out["hours_mix_aug"] = PP["hours_mix_aug"]
    # reported basis: NOI as reported on templates (catering kept), core COGS re-run at the scenario; core basis kept alongside
    cogs_adj = {"actual_cogs": 0, "30": t3["cogs_actual"] - round(0.30 * t3["core_sales"]), "28.5": t3["cogs_actual"] - round(0.285 * t3["core_sales"]), "27": t3["cogs_actual"] - round(0.27 * t3["core_sales"])}
    for kk, v in scen.items():
        extra = round(wage_base * v["uplift_pct"] / 100 * BR); v["extra_pl_labor_year"] = extra
        v["labor_template_year"] = t3["labor_template"] + extra; v["labor_pct_of_core"] = round(v["labor_template_year"] / t3["core_sales"] * 100, 1)
        v["ladder"] = {}; v["ladder_core_basis"] = {}
        for sc in ("actual_cogs", "30", "28.5", "27"):
            b = t3["noi_template"] + cogs_adj[sc] - extra
            v["ladder"][sc] = {"cafes": b, "after_support": b + sup, "after_support_salary": b + sup - FOUNDER_SALARY, "after_support_salary_incentive": b + sup - FOUNDER_SALARY - round(sim["p50"]),
                               "gap_pct_of_core": round((b + sup - FOUNDER_SALARY) / t3["core_sales"] * 100, 1)}
            bc = base[sc]["template_labor"] - extra
            v["ladder_core_basis"][sc] = {"cafes": bc, "after_support_salary": bc + sup - FOUNDER_SALARY}
    out["pay_scenarios_12mo"] = {"basis": "reported NOI on templates (catering kept) with core COGS re-run at the scenario; three cafes; support at 2026 avg; founder salary $120K", "cogs_adjustment_year": cogs_adj}
    out["pay_scenarios_12mo"] |= {"founder_salary_year": FOUNDER_SALARY, "wage_base_template_12mo": wage_base, "load": BR, "scenarios": scen, "noi_template_reported_12mo": t3["noi_template"],
                                 "price_lever": {"pct_1_of_core_sales": round(t3["core_sales"] / 100), "note": "a 1% price move ≈ $11.5K/yr before any volume response (merchant fees ~3% of the increase)"},
                                 "note": "Uplift % = each August roster member moved from their current band to the scenario rate (nobody down), hours-weighted, incl. admin/training; applied to the 12-month template wage base (core + admin/training) and loaded ×1.12. Peer all-in = DC minimum + tips as observed in postings; ALL at $27 assumes peer shift lead + $1.50."}
    # ---- Slipstream economics: 1899 is a managed location (NI passes to the owner) → count only the management income;
    #      NMWA counted at its Jul–Aug 2026 run-rate incl. the $5,041 fee, plus the NMWA Nights events (Ryan: 8 × ~$4,500 net)
    OWN = ("Downtown", "Studio"); con = fin["consolidated"]; nmon = len(con)
    avg = lambda key: sum(con[mth][key] for mth in con) / nmon
    mgmt_1899 = round(avg("mgmt") * 12); comm = round(avg("comm") * 12); srp = round(avg("srp") * 12); central = round(avg("central") * 12)
    nm = fin["cafes"]["NMWA"]; nmwa_run = round(sum(nm[mth]["ni"] for mth in ("Jul", "Aug")) / 2 * 12); nmwa_fee = 5041 * 12
    NIGHTS_N, NIGHTS_NET = 8, 4500; nights = NIGHTS_N * NIGHTS_NET
    own_noi = sum(BT["cafes"][c]["totals"]["noi_template"] for c in OWN); own_core = sum(BT["cafes"][c]["totals"]["core_sales"] for c in OWN); own_cogs = sum(BT["cafes"][c]["totals"]["cogs_actual"] for c in OWN)
    own_base = sum(BT["cafes"][c]["totals"]["template_core_wages"] + BT["cafes"][c]["totals"]["admin_training"] for c in OWN); own_labor = sum(BT["cafes"][c]["totals"]["labor_template"] for c in OWN)
    own_adj = {"actual_cogs": 0, "30": own_cogs - round(0.30 * own_core), "28.5": own_cogs - round(0.285 * own_core), "27": own_cogs - round(0.27 * own_core)}
    fixed = mgmt_1899 + nmwa_run + nights + comm + srp + central - FOUNDER_SALARY
    econ = {"lines": {"downtown_studio_noi_on_templates_12mo": own_noi, "cogs_adjustment": own_adj, "mgmt_income_1899_year": mgmt_1899, "nmwa_cafe_run_rate_year_incl_fee": nmwa_run, "nmwa_fee_in_run_rate": nmwa_fee,
                      "nmwa_nights_year": nights, "nmwa_nights_assumption": f"{NIGHTS_N} × ${NIGHTS_NET:,} net", "commissary_year": comm, "srp_year": srp, "central_year": central, "founder_salary_year": -FOUNDER_SALARY,
                      "excluded_1899_noi_on_templates_12mo": BT["cafes"]["1899 Penn"]["totals"]["noi_template"], "excluded_1899_labor_saving_12mo": BT["cafes"]["1899 Penn"]["totals"]["labor_saving"]},
            "own_cafes": {"core_sales": own_core, "cogs_actual": own_cogs, "cogs_pct": round(own_cogs / own_core * 100, 1), "wage_base": own_base, "labor_template": own_labor}, "scenarios": {}}
    for kk, v in scen.items():
        extra = round(own_base * v["uplift_pct"] / 100 * BR)
        e = {"label": v["label"], "uplift_pct": v["uplift_pct"], "avg_rate": v["avg_rate"], "roster_avg_rate": v["roster_avg_rate"], "rates": v.get("rates"), "extra_pl_labor_year": extra, "labor_pct_of_core": round((own_labor + extra) / own_core * 100, 1), "ladder": {}}
        for sc in ("actual_cogs", "30", "28.5", "27"):
            val = own_noi + own_adj[sc] - extra + fixed
            e["ladder"][sc] = {"after_everything": val, "after_everything_incentive": val - round(sim["p50"]), "pct_of_own_core": round(val / own_core * 100, 1)}
        econ["scenarios"][kk] = e
    econ["note"] = ("Slipstream's own economics: Downtown + Studio NOI on templates (12-month backtest, catering kept, core COGS re-run) + 1899 management income ($5K + 5% of sales, 2026 average) "
                    "+ NMWA at its Jul–Aug 2026 run-rate (includes the $5,041/mo fee) + NMWA Nights (8 × $4,500) − Commissary, SRP, Central (2026 average × 12) − founder salary. 1899's own NOI and its template saving are excluded (passed to the owner); "
                    "pay scenarios apply to Downtown + Studio only (1899 raises are the owner's cost). NMWA run-rate rests on two months.")
    out["slipstream_economics_12mo"] = econ
    json.dump(out, open(os.path.join(DD, "affordability_v13.json"), "w"), indent=1)
    print(json.dumps(econ["lines"], indent=1))
    for kk, e in econ["scenarios"].items(): print(f"{e['label'][:44]:44} extra {e['extra_pl_labor_year']:>8,} labor% {e['labor_pct_of_core']}", {sc: x["after_everything"] for sc, x in e["ladder"].items()})

if __name__ == "__main__":
    main()
