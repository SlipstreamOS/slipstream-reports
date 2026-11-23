#!/usr/bin/env python3
"""Labor Deep Dive v1.3 step 21 — 12-month backtest (Sep 2025 – Aug 2026) of the v1.3 schedule templates per cafe.

For every open day the week's demand tier is set from ACTUAL weekly walk-in sales (same bands as tiers.json), the template
for that cafe × tier × weekday is costed at planning rates, admin + training wages are added at their recent actual level,
and the result is converted to P&L labor with the validated bridge (P&L labor ≈ 1.12 × gross wages − CC tips).
Everything else (sales, COGS, operating expenses) is held at actual, then COGS is re-run at 27 / 28.5 / 30 % of core sales.

Inputs : fin_2025_v17.json (recalculated copy of the Consolidated Model, 2025 columns), fin_2026_v17.json,
         labor_deep_dive/{txns_2026.csv, projections_daily_v18.csv (s25 = 2025 daily walk-in), templates_v13.json,
         tiers.json, monthly_wage_basis.json}
Output : labor_deep_dive/backtest_v13.json
"""
import os, json
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__)); DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
CAFES = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899"}          # template name -> fin/model name
NORMAL_DAYS = {"Downtown": 5, "Studio": 7, "1899 Penn": 5}
BRIDGE = 1.12
COGS_SCEN = [0.27, 0.285, 0.30]
# Alternative rate sets for the pay scenarios (applied to template hours by shift code). Peer = DC peer all-in (base + tips):
# barista/line/production ~$24.00 (DC min $18.40 + ~$5.60 tips; Compass $24–26; Indeed avg + tips), shift lead ~$25.50 (peer barista + $1–2).
RATE_SETS = {"planning": None,
             "full_curve": {"SL": 23.50, "PL": 20.50, "PB": 20.50, "B": 20.50, "CL": 20.50, "HR": 20.50, "FOH": 20.50},
             "peer": {"SL": 25.50, "PL": 24.00, "PB": 24.00, "B": 24.00, "CL": 24.00, "HR": 24.00, "FOH": 24.00},
             "peer_plus5": {"SL": 26.78, "PL": 25.20, "PB": 25.20, "B": 25.20, "CL": 25.20, "HR": 25.20, "FOH": 25.20}}
MONTHS = [f"2025-{m:02d}" for m in (9, 10, 11, 12)] + [f"2026-{m:02d}" for m in range(1, 9)]
MON_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def label(ym): y, m = ym.split("-"); return f"{MON_ABBR[int(m)-1]} {y[2:]}"

def main():
    fin25 = json.load(open(os.path.join(DATA, "fin_2025_v17.json"))); fin26 = json.load(open(os.path.join(DATA, "fin_2026_v17.json")))
    tpl = json.load(open(os.path.join(DD, "templates_v13.json")))["cafes"]; tiers = json.load(open(os.path.join(DD, "tiers.json")))
    mwb = pd.DataFrame(json.load(open(os.path.join(DD, "monthly_wage_basis.json"))))
    # daily walk-in sales: 2026 from transactions, Sep–Dec 2025 from the 2025 daily series
    tx = pd.read_csv(os.path.join(DD, "txns_2026.csv")); d26 = tx.groupby(["cafe", "date"]).net.sum().reset_index()
    p = pd.read_csv(os.path.join(DD, "projections_daily_v18.csv")); p = p[(p.d25 >= "2025-09-01") & (p.d25 <= "2025-12-31")].dropna(subset=["s25"])
    inv = {v: k for k, v in CAFES.items()}; p["cafe"] = p.cafe.map(inv); d25 = p.rename(columns={"d25": "date", "s25": "net"})[["cafe", "date", "net"]]
    daily = pd.concat([d25, d26[d26.cafe.isin(CAFES)]]); daily["date"] = pd.to_datetime(daily.date); daily = daily[daily.net > 0]
    daily["dow"] = daily.date.dt.day_name().str[:3]; daily["week"] = daily.date - pd.to_timedelta(daily.date.dt.weekday, unit="D"); daily["ym"] = daily.date.dt.strftime("%Y-%m")
    out = {"method": {"window": "Sep 2025 – Aug 2026", "bridge": "P&L labor = 1.12 × gross wages − CC tips", "cogs_scenarios": COGS_SCEN,
                      "tiering": "week tier from actual weekly walk-in sales scaled to a normal week (tiers.json bands); template wages per open day at planning rates",
                      "admin_training": "added at the cafe's Apr–Aug 2026 monthly average (full-month basis)",
                      "tips": "CC tips as the cafe's Apr–Aug 2026 average share of core sales",
                      "caveats": ["template rates are Aug 2026 planning rates; 2025 actual labor was paid at lower rates, so the 2025 comparison understates the template's saving",
                                  "reported NOI includes catering; the core-basis rows exclude catering revenue and COGS entirely",
                                  "2025 core COGS % includes catering COGS in the Consolidated Model (catering-heavy months look worse than they are)"]},
           "cafes": {}}
    for cafe, fname in CAFES.items():
        bands = tiers[cafe]["bands_weekly"]; T = tpl[cafe]; nd = NORMAL_DAYS[cafe]
        m = mwb[mwb.cafe == cafe]; admin_tr = float((m.admin_wages_full + m.training_wages_full).mean()); tip_share = float(m.tips.sum() / m.core_sales.sum())
        dc = daily[daily.cafe == cafe].copy()
        wk = dc.groupby("week").agg(net=("net", "sum"), days=("date", "count")).reset_index(); wk["scaled"] = wk.net * nd / wk.days.clip(upper=nd)
        def tier_of(s):
            if s < bands["slow"][1]: return "slow"
            if s < bands["mid"][1]: return "mid"
            return "busy"
        wk["tier"] = wk.scaled.apply(tier_of); dc = dc.merge(wk[["week", "tier", "scaled"]], on="week")
        def tpl_wages(r):
            day = T[r.tier]["days"].get(r.dow)
            return day["template"]["wages"] if day else None
        def tpl_hours(r):
            day = T[r.tier]["days"].get(r.dow)
            return day["template"]["hours"] if day else None
        dc["tpl_wages"] = dc.apply(tpl_wages, axis=1); dc["tpl_hours"] = dc.apply(tpl_hours, axis=1)
        def tpl_wages_at(r, rates):
            day = T[r.tier]["days"].get(r.dow)
            return sum(sh["hours"] * rates[sh["code"]] for sh in day["shifts"]) if day else None
        for rs, rates in RATE_SETS.items():
            if rates: dc[f"tpl_wages_{rs}"] = dc.apply(lambda r: tpl_wages_at(r, rates), axis=1)
        untemplated = dc[dc.tpl_wages.isna()].groupby("dow").net.agg(["count", "sum"])
        months = []
        for ym in MONTHS:
            y, mm = ym.split("-"); mon = MON_ABBR[int(mm) - 1]
            if y == "2025":
                r = fin25[fname][ym]; core = r["core"]; cogs = r["tcogs"]; labor = r["labor"]; texp = r["texp"]; noi = r["noi"]; cat = r["cat"]
            else:
                r = fin26["cafes"][fname][mon]; core = r["core"]; cogs = r["cogs"]; labor = r["labor"]; texp = r["texp"]; noi = r["noi"]; cat = r["cat"]
            dm = dc[dc.ym == ym]; wkm = dm.groupby("week").tier.first().value_counts().to_dict()
            gross_core = float(dm.tpl_wages.fillna(0).sum()); hours = float(dm.tpl_hours.fillna(0).sum()); open_days = int(len(dm)); walkin = float(dm.net.sum())
            tips = tip_share * core; tpl_pl = BRIDGE * (gross_core + admin_tr) - tips
            delta = labor - tpl_pl
            rec = {"month": label(ym), "ym": ym, "open_days": open_days, "walkin_sales": round(walkin), "tier_weeks": {k: int(v) for k, v in wkm.items()},
                   "core_sales": round(core), "catering": round(cat), "cogs_actual": round(cogs), "cogs_pct_actual": round(cogs / core * 100, 1) if core else None,
                   "opex": round(texp), "labor_actual": round(labor), "labor_pct_actual": round(labor / core * 100, 1) if core else None,
                   "template_core_hours": round(hours), "template_core_wages": round(gross_core), "template_wage_pct": round(gross_core / walkin * 100, 1) if walkin else None,
                   "admin_training": round(admin_tr), "tips_est": round(tips), "labor_template": round(tpl_pl), "labor_pct_template": round(tpl_pl / core * 100, 1) if core else None,
                   "labor_saving": round(delta), "noi_reported": round(noi), "noi_template": round(noi + delta),
                   "core_basis": {}}
            for sc in COGS_SCEN:
                k = f"{sc*100:g}"
                rec["core_basis"][k] = {"actual_labor": round(core * (1 - sc) - labor - texp), "template_labor": round(core * (1 - sc) - tpl_pl - texp)}
            rec["core_basis"]["actual_cogs"] = {"actual_labor": round(core - cogs - labor - texp), "template_labor": round(core - cogs - tpl_pl - texp)}
            rec["pay_scenarios"] = {}
            for rs, rates in RATE_SETS.items():
                if not rates: continue
                gw = float(dm[f"tpl_wages_{rs}"].fillna(0).sum()); pl = BRIDGE * (gw + admin_tr) - tips
                rec["pay_scenarios"][rs] = {"template_core_wages": round(gw), "labor_template": round(pl), "extra_vs_planning": round(pl - tpl_pl),
                                            "core_basis": {f"{sc*100:g}": round(core * (1 - sc) - pl - texp) for sc in COGS_SCEN} | {"actual_cogs": round(core - cogs - pl - texp)}}
            months.append(rec)
        tot = lambda k: round(sum(x[k] for x in months))
        totals = {k: tot(k) for k in ("open_days", "walkin_sales", "core_sales", "catering", "cogs_actual", "opex", "labor_actual", "template_core_hours", "template_core_wages", "admin_training", "tips_est", "labor_template", "labor_saving", "noi_reported", "noi_template")}
        totals["cogs_pct_actual"] = round(totals["cogs_actual"] / totals["core_sales"] * 100, 1); totals["labor_pct_actual"] = round(totals["labor_actual"] / totals["core_sales"] * 100, 1)
        totals["labor_pct_template"] = round(totals["labor_template"] / totals["core_sales"] * 100, 1); totals["template_wage_pct"] = round(totals["template_core_wages"] / totals["walkin_sales"] * 100, 1)
        totals["core_basis"] = {k: {s: round(sum(x["core_basis"][k][s] for x in months)) for s in ("actual_labor", "template_labor")} for k in months[0]["core_basis"]}
        totals["pay_scenarios"] = {rs: {"template_core_wages": sum(x["pay_scenarios"][rs]["template_core_wages"] for x in months), "labor_template": sum(x["pay_scenarios"][rs]["labor_template"] for x in months),
                                        "extra_vs_planning": sum(x["pay_scenarios"][rs]["extra_vs_planning"] for x in months),
                                        "core_basis": {kk: sum(x["pay_scenarios"][rs]["core_basis"][kk] for x in months) for kk in months[0]["pay_scenarios"][rs]["core_basis"]}} for rs in months[0]["pay_scenarios"]}
        tw = pd.Series(sum((list(x["tier_weeks"].items()) for x in months), [])); totals["tier_weeks"] = dc.groupby("week").tier.first().value_counts().to_dict()
        months_neg_actual = sum(1 for x in months if x["core_basis"]["28.5"]["actual_labor"] < 0); months_neg_tpl = sum(1 for x in months if x["core_basis"]["28.5"]["template_labor"] < 0)
        out["cafes"][cafe] = {"admin_training_month": round(admin_tr), "tip_share_of_core": round(tip_share * 100, 1), "bands_weekly": bands, "months": months, "totals": totals,
                              "months_losing_at_28_5_actual": months_neg_actual, "months_losing_at_28_5_template": months_neg_tpl,
                              "untemplated_days": {k: {"days": int(v["count"]), "sales": round(v["sum"])} for k, v in untemplated.iterrows()}}
        print(f"{cafe}: 12-mo core {totals['core_sales']:,} | labor actual {totals['labor_actual']:,} ({totals['labor_pct_actual']}%) → template {totals['labor_template']:,} ({totals['labor_pct_template']}%) | saving {totals['labor_saving']:,} | NOI reported {totals['noi_reported']:,} → {totals['noi_template']:,} | core-basis @28.5%: {totals['core_basis']['28.5']} | tiers {totals['tier_weeks']} | untemplated {out['cafes'][cafe]['untemplated_days']}")
        for x in months: print(f"   {x['month']} core {x['core_sales']:>7,} tiers {x['tier_weeks']} labor {x['labor_actual']:>6,} ({x['labor_pct_actual']}%) → {x['labor_template']:>6,} ({x['labor_pct_template']}%) Δ {x['labor_saving']:>6,} | NOI {x['noi_reported']:>7,} → {x['noi_template']:>7,} | core@28.5 {x['core_basis']['28.5']['actual_labor']:>7,} → {x['core_basis']['28.5']['template_labor']:>7,} cogs% {x['cogs_pct_actual']}")
    # cafe total + support burden view
    allm = []
    for i, ym in enumerate(MONTHS):
        rec = {"month": label(ym), "ym": ym}
        for k in ("core_sales", "labor_actual", "labor_template", "labor_saving", "noi_reported", "noi_template", "opex", "cogs_actual"):
            rec[k] = sum(out["cafes"][c]["months"][i][k] for c in CAFES)
        rec["core_basis"] = {k: {s: sum(out["cafes"][c]["months"][i]["core_basis"][k][s] for c in CAFES) for s in ("actual_labor", "template_labor")} for k in out["cafes"]["Downtown"]["months"][i]["core_basis"]}
        rec["pay_scenarios"] = {rs: {"labor_template": sum(out["cafes"][c]["months"][i]["pay_scenarios"][rs]["labor_template"] for c in CAFES), "extra_vs_planning": sum(out["cafes"][c]["months"][i]["pay_scenarios"][rs]["extra_vs_planning"] for c in CAFES),
                                     "core_basis": {kk: sum(out["cafes"][c]["months"][i]["pay_scenarios"][rs]["core_basis"][kk] for c in CAFES) for kk in ("27", "28.5", "30", "actual_cogs")}} for rs in ("full_curve", "peer", "peer_plus5")}
        y, mm = ym.split("-")
        if y == "2026":
            cm = fin26["consolidated"][MON_ABBR[int(mm) - 1]]; rec["support_ni"] = round(cm["comm"] + cm["srp"] + cm["mgmt"] + cm["central"]); rec["nmwa_noi"] = round(fin26["cafes"]["NMWA"][MON_ABBR[int(mm) - 1]]["noi"]); rec["total_ni_reported"] = round(cm["total_ni"])
        else:
            rec["support_ni"] = None; rec["nmwa_noi"] = 0; rec["total_ni_reported"] = None
        allm.append(rec)
    sup26 = [x["support_ni"] for x in allm if x["support_ni"] is not None]; sup_avg = round(sum(sup26) / len(sup26))
    tot3 = {k: sum(x[k] for x in allm) for k in ("core_sales", "labor_actual", "labor_template", "labor_saving", "noi_reported", "noi_template", "opex", "cogs_actual")}
    tot3["core_basis"] = {k: {s: sum(x["core_basis"][k][s] for x in allm) for s in ("actual_labor", "template_labor")} for k in allm[0]["core_basis"]}
    tot3["pay_scenarios"] = {rs: {"labor_template": sum(x["pay_scenarios"][rs]["labor_template"] for x in allm), "extra_vs_planning": sum(x["pay_scenarios"][rs]["extra_vs_planning"] for x in allm),
                                  "core_basis": {kk: sum(x["pay_scenarios"][rs]["core_basis"][kk] for x in allm) for kk in ("27", "28.5", "30", "actual_cogs")}} for rs in ("full_curve", "peer", "peer_plus5")}
    out["method"]["rate_sets"] = RATE_SETS
    out["three_cafes"] = {"months": allm, "totals": tot3, "support_burden_2026_avg_month": sup_avg, "support_burden_2026_ytd": round(sum(sup26)),
                          "note": "Support entities (Commissary, SRP, Cafe Mgmt, Central) net to a monthly burden in 2026; the Consolidated Model carries no 2025 support figures, so the 12-month 'after support' view uses the 2026 Jan–Aug average × 12"}
    out["three_cafes"]["after_support_12mo"] = {"reported": tot3["noi_reported"] + sup_avg * 12, "template": tot3["noi_template"] + sup_avg * 12,
                                                "core_basis_28_5_template": tot3["core_basis"]["28.5"]["template_labor"] + sup_avg * 12, "core_basis_30_template": tot3["core_basis"]["30"]["template_labor"] + sup_avg * 12, "core_basis_27_template": tot3["core_basis"]["27"]["template_labor"] + sup_avg * 12}
    print("3 cafes:", {k: (v if not isinstance(v, dict) else v) for k, v in tot3.items()}); print("support avg/mo", sup_avg, "after support 12mo", out["three_cafes"]["after_support_12mo"])
    json.dump(out, open(os.path.join(DD, "backtest_v13.json"), "w"), indent=1)

if __name__ == "__main__":
    main()
