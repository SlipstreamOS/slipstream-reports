#!/usr/bin/env python3
"""Labor Deep Dive step 8 — incentive sizing and structure comparison.
Inputs: templates.json (tier sales, template wages, sampled actual wages), monthly_wage_basis.json,
fin_2026_v17.json (COGS / opex by cafe). Output: data/labor_deep_dive/incentive.json
All incentive math is on the WAGE basis (core wages / core sales) — the basis the team can see and control.
"""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
WEEKS_PER_MONTH = 52 / 12
TARGET = {"Downtown": 0.32, "Studio": 0.30, "1899 Penn": 0.35}
TIP_RATE = {"Downtown": 0.121, "Studio": 0.120, "1899 Penn": 0.104}
LOAD = 1.12   # payroll tax/benefit allocation on gross wages (Payroll Detail C2)
FK = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}

def main():
    T = json.load(open(os.path.join(DD, "templates.json")))
    mwb = json.load(open(os.path.join(DD, "monthly_wage_basis.json")))
    fin = json.load(open(os.path.join(DD, "..", "fin_2026_v17.json")))
    out = {"assumptions": {"weeks_per_month": WEEKS_PER_MONTH, "load_factor": LOAD, "tip_rate": TIP_RATE, "targets": TARGET}}
    # ---- 1. headroom by tier (monthly): actual sampled wages vs target wages at tier-typical sales
    head = {}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        head[cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            w = T["cafes"][cafe][tier]["week"]
            sales_m = w["sales_typical"] * WEEKS_PER_MONTH
            act_m = w["actual_sampled_wages"] * WEEKS_PER_MONTH
            tmpl_m = w["wages"] * WEEKS_PER_MONTH
            tgt_m = sales_m * TARGET[cafe]
            head[cafe][tier] = {"core_sales_month": round(sales_m), "actual_wages_month": round(act_m), "actual_pct": round(act_m / sales_m * 100, 1),
                                "target_wages_month": round(tgt_m), "headroom_actual_vs_target": round(act_m - tgt_m),
                                "template_wages_month": round(tmpl_m), "template_pct": round(tmpl_m / sales_m * 100, 1),
                                "template_saving_vs_actual": round(act_m - tmpl_m), "core_hours_month_template": round(w["hours"] * WEEKS_PER_MONTH)}
    out["headroom_by_tier"] = head
    # Apr–Aug actual headroom from monthly_wage_basis
    out["headroom_actual_months"] = {r["cafe"] + " " + r["month"]: r["headroom_vs_target"] for r in mwb}
    out["headroom_aug_total"] = round(sum(r["headroom_vs_target"] for r in mwb if r["month"] == "Aug"))
    out["headroom_aprjun_avg_total"] = round(sum(r["headroom_vs_target"] for r in mwb if r["month"] in ("Apr", "May", "Jun")) / 3)

    # ---- 2. Plan comparison at a MID-tier month, per cafe and company-wide
    plans = {}
    def payout(plan, cafe, sales, hours, pct_vs_target):
        """pct_vs_target: core wage % minus target (negative = beat). Returns monthly payout $."""
        beat_dollars = max(0.0, -pct_vs_target) * sales
        if plan == "A_beat_share_50_cap1.5":
            return min(0.5 * beat_dollars, 0.015 * sales)
        if plan == "A2_ryan_2pt_gate":
            return 0.5 * beat_dollars if pct_vs_target <= -0.02 else 0.0
        if plan == "B_dollar_per_hour":
            if pct_vs_target <= -0.04: return 2.0 * hours
            if pct_vs_target <= -0.02: return 1.5 * hours
            if pct_vs_target <= 0.0: return 1.0 * hours
            return 0.0
        if plan == "D_shift_bonus_8":      # $8 per shift when the week hits target (≈ $1/hr on an 8h shift); shifts ≈ hours/7.5
            return 8.0 * (hours / 7.5) if pct_vs_target <= 0.0 else 0.0
        if plan == "C_prime_gainshare_20":  # 20% of labor beat, no gate, no cap (COGS component held at plan)
            return 0.20 * beat_dollars
        return 0.0
    scen = [("2 pts OVER target", 0.02), ("AT target", 0.0), ("1 pt under", -0.01), ("2 pts under", -0.02), ("4 pts under", -0.04)]
    plan_list = ["A_beat_share_50_cap1.5", "A2_ryan_2pt_gate", "B_dollar_per_hour", "C_prime_gainshare_20", "D_shift_bonus_8"]
    comp = {}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        w = T["cafes"][cafe]["mid"]["week"]
        sales = w["sales_typical"] * WEEKS_PER_MONTH
        hours = w["hours"] * WEEKS_PER_MONTH
        comp[cafe] = {"mid_month_sales": round(sales), "mid_month_core_hours": round(hours), "scenarios": {}}
        for label, d in scen:
            row = {}
            for p in plan_list:
                pay = payout(p, cafe, sales, hours, d)
                row[p] = {"payout": round(pay), "pct_of_sales": round(pay / sales * 100, 2), "per_core_hour": round(pay / hours, 2) if hours else 0,
                          "employer_cost_incl_load": round(pay * LOAD), "net_labor_pct_after": round(((TARGET[cafe] + d) * sales + pay) / sales * 100, 1),
                          "slipstream_keeps_of_beat": round(max(0.0, -d) * sales - pay)}
            comp[cafe]["scenarios"][label] = row
    # company-wide (open 3) at mid tier
    sales_all = sum(comp[c]["mid_month_sales"] for c in comp); hours_all = sum(comp[c]["mid_month_core_hours"] for c in comp)
    comp["Open Cafes"] = {"mid_month_sales": round(sales_all), "mid_month_core_hours": round(hours_all), "scenarios": {}}
    for label, d in scen:
        row = {}
        for p in plan_list:
            pay = sum(payout(p, c, comp[c]["mid_month_sales"], comp[c]["mid_month_core_hours"], d) for c in ["Downtown", "Studio", "1899 Penn"])
            row[p] = {"payout": round(pay), "pct_of_sales": round(pay / sales_all * 100, 2), "per_core_hour": round(pay / hours_all, 2),
                      "employer_cost_incl_load": round(pay * LOAD), "per_person_month_30h": round(pay / hours_all * 30 * WEEKS_PER_MONTH)}
        comp["Open Cafes"]["scenarios"][label] = row
    out["plan_comparison_mid_tier"] = comp

    # ---- 3. Margin view: P&L at August actual vs at target (with/without incentive), per cafe, August sales
    margins = {}
    for cafe in ["Downtown", "Studio", "1899 Penn"]:
        f = fin["cafes"][FK[cafe]]["Aug"]
        sales = f["core"]; cogs = f["cogs"]; opex = f["texp"]; lab_act = f["labor"]
        aug = [r for r in mwb if r["cafe"] == cafe and r["month"] == "Aug"][0]
        admin_train_pct = (aug["admin_wages_full"] + aug["training_wages_full"]) / sales
        def pl_labor(core_pct, incentive=0.0):
            gross = (core_pct + admin_train_pct) * sales + incentive
            return LOAD * gross - TIP_RATE[cafe] * sales
        rows = {}
        for label, core_pct, inc in [("August actual (booked)", None, 0), ("August at wage-basis est.", aug["core_wage_pct"] / 100, 0),
                                     ("At target, no incentive", TARGET[cafe], 0),
                                     ("2 pts under target, Plan A (50% beat share)", TARGET[cafe] - 0.02, 0.5 * 0.02 * sales),
                                     ("2 pts under target, Plan B ($1.50/hr)", TARGET[cafe] - 0.02, 1.5 * aug["core_hours_full"] * (1 - 0.02 / max(aug["core_wage_pct"]/100, 0.01)))]:
            lab = lab_act if core_pct is None else pl_labor(core_pct, inc)
            noi = sales - cogs - lab - opex
            rows[label] = {"labor_pl": round(lab), "labor_pct": round(lab / sales * 100, 1), "prime_pct": round((cogs + lab) / sales * 100, 1),
                           "noi": round(noi), "noi_pct": round(noi / sales * 100, 1), "incentive_paid": round(inc)}
        margins[cafe] = {"aug_core_sales": round(sales), "cogs_pct": round(cogs / sales * 100, 1), "opex_pct": round(opex / sales * 100, 1),
                         "admin_train_pct_of_sales": round(admin_train_pct * 100, 1), "rows": rows}
    out["margin_view_august"] = margins
    # ---- 4. annual profit-share sizing at target
    cafe_noi_target = sum(margins[c]["rows"]["At target, no incentive"]["noi"] for c in margins) * 12
    out["annual_profit_share_sizing"] = {"open_cafe_noi_at_target_annualized_on_aug_sales": round(cafe_noi_target),
                                          "5pct": round(0.05 * cafe_noi_target), "10pct": round(0.10 * cafe_noi_target),
                                          "note": "Annualized from August (slowest month) sales — a floor. Consolidated NI is lower after Commissary/SRP/Central (~$14–16K/mo normalized)."}
    json.dump(out, open(os.path.join(DD, "incentive.json"), "w"), indent=1)
    # console
    print("HEADROOM by tier (monthly, actual sampled vs target):")
    for c, t in head.items():
        for tier, v in t.items():
            print(f"  {c:10} {tier:5} sales ${v['core_sales_month']:>7,} actual {v['actual_pct']}% -> headroom ${v['headroom_actual_vs_target']:>6,}  template {v['template_pct']}% (saves ${v['template_saving_vs_actual']:,} vs actual)")
    print("Aug total headroom:", out["headroom_aug_total"], " Apr–Jun avg:", out["headroom_aprjun_avg_total"])
    print("\nPLAN COMPARISON — Open Cafes, mid tier month:")
    oc = comp["Open Cafes"]
    print(f"  sales ${oc['mid_month_sales']:,}  core hours {oc['mid_month_core_hours']:,}")
    for label, row in oc["scenarios"].items():
        print(f"  {label:18} " + " | ".join(f"{p[:14]} ${v['payout']:>5,} ({v['pct_of_sales']:.2f}%, ${v['per_core_hour']:.2f}/h, ${v['per_person_month_30h']}/mo)" for p, v in row.items()))
    print("\nMARGIN VIEW (August sales):")
    for c, m in margins.items():
        print(f"  {c}: COGS {m['cogs_pct']}% opex {m['opex_pct']}% admin+train {m['admin_train_pct_of_sales']}%")
        for label, r in m["rows"].items():
            print(f"     {label:45} labor {r['labor_pct']:5.1f}%  prime {r['prime_pct']:5.1f}%  NOI ${r['noi']:>7,} ({r['noi_pct']}%)  incentive ${r['incentive_paid']}")
    print("\nAnnual profit share sizing:", out["annual_profit_share_sizing"])

if __name__ == "__main__":
    main()
