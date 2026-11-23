#!/usr/bin/env python3
"""Labor Deep Dive v1.1 step 13 — reference JSON for the monthly report (supersedes the v1.0 ldd_09 output).
Keeps the v1.0 file as labor_deep_dive_2026-09_v1.0.json and writes v1.1 to labor_deep_dive_2026-09.json.
"""
import os, json, shutil
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data"); DD = os.path.join(DATA, "labor_deep_dive")
J = lambda n: json.load(open(os.path.join(DD, n)))

def main():
    v10_path = os.path.join(DATA, "labor_deep_dive_2026-09.json")
    v10 = json.load(open(v10_path))
    if v10.get("version") == "1.0" and not os.path.exists(os.path.join(DATA, "labor_deep_dive_2026-09_v1.0.json")):
        shutil.copy(v10_path, os.path.join(DATA, "labor_deep_dive_2026-09_v1.0.json"))
    tiers, cap, NV, T, inc, staffing = J("tiers.json"), J("capacity.json"), J("need_v11.json"), J("templates_v11.json"), J("incentive_v11.json"), J("staffing.json")
    MAIN = "no10h"      # primary templates: no shift over 9 h; the 10-hour open-to-close lead is reported as an option
    templates = {}
    for cafe, tiers_o in T["cafes"].items():
        templates[cafe] = {}
        for tier in ["slow", "mid", "busy"]:
            main = tiers_o[f"{tier}_{MAIN}"]
            templates[cafe][tier] = {
                "week": main["week"],
                "days": {d: {"sales_typical": dd["sales_typical"], "need_hours": dd["need_hours"], "hours": dd["template"]["hours"], "wages": dd["template"]["wages"],
                             "labor_pct": dd["template"]["labor_pct"], "labor_pct_obs_rate": dd["template"]["labor_pct_obs_rate"], "peak_staff": dd["template"]["peak_staff"],
                             "solo_windows": dd["template"]["solo_windows"], "admin_cover_windows": dd["template"]["admin_cover_windows"],
                             "shifts": [{k: s[k] for k in ("role", "code", "start", "end", "hours", "rate", "wages")} for s in dd["shifts"]]} for d, dd in main["days"].items()},
                "variants_week": {v: {"hours": tiers_o[f"{tier}_{v}" if v != "target" else tier]["week"]["hours"], "wages": tiers_o[f"{tier}_{v}" if v != "target" else tier]["week"]["wages"],
                                      "labor_pct": tiers_o[f"{tier}_{v}" if v != "target" else tier]["week"]["labor_pct"]} for v in ["target", "comfort", "min6h", "pl8"]},
            }
    out = {
        "generated": "2026-09-04", "version": "1.1", "supersedes": "1.0 (same day)",
        "scope": v10["scope"],
        "changes_from_v1_0": [
            "Planning capacity is cafe-specific (busy-hour p90 per cafe): Downtown 14, Studio 12, 1899 14 transactions per person-hour (v1.0 used 14 everywhere; Studio's weekend food tickets run slower).",
            "Floors: Downtown 2-person floor ends at 1:00p (was 2:30p); 1899 second person required only in the two peaks when the curve is >= 12/hr (was 2 people to 1:30p); Studio unchanged (Mon 3 at open until 10:00).",
            "Templates are minimum-cost shift sets from an integer program (CBC) over practical shift families, not hand-built; no shift over 9 h in the primary templates (the 10-hour lead option is worth only ~$35/wk at Downtown).",
            "Every template is scored at standard rates AND at the cafe's current average core rate (the rate-lever exposure).",
            "Variants priced: comfort (capacity 12 + longer 2-person floors), 6-hour minimum shifts, 8-hour Production Line.",
            "Incentive section rebuilt: nine structures, mixed-cafe example, 2,000-run annual simulation of 2027 payouts vs labor saved, margins by tier per plan, retirement profit-share sized on consolidated NI, FLSA regular-rate and DC minimum-wage notes.",
        ],
        "definitions": {**v10["definitions"],
                        "capacity": "planning capacity per cafe = busy-hour p90 throughput observed: Downtown 14, Studio 12, 1899 14 transactions per person-hour",
                        "need_rule": "need(h) = max(floor, ceil(txns/h ÷ capacity)); floors — Downtown: 1 at 6:00, 2 from 6:30 to 1:00p, then 1; Studio: 1 at 6:00 (7:00 weekend), 2 to 90 min before close, then 1, Monday 3 from open to 10:00; 1899: 1 (lead) throughout, 2 in the 8–10a and 12:30–2p peaks when the curve is >= 12/hr; NMWA 1 (2 on community days)",
                        "optimizer": "per cafe × tier × DOW integer program: Production Line opens (6–8 h), Shift Lead opens (8–9 h), optional Closer (4–8 h, keyholder), up to three 4–8 h barista shifts; 1899: lead 6:30–3:30 + one or two 4–8 h Production Barista shifts, max 2 on the floor; keyholder on the floor open−30 min to close+tail; minimise wages at standard rates"},
        "tiers": v10["tiers"],
        "tier_band_validation": {"note": "k-means (k=3) and Jenks breaks on the 33–34 weekly 2026 observations per cafe",
                                 "Downtown": {"v1_bands": [7500, 9000], "kmeans": [7548, 9348], "jenks": [7092, 9160]},
                                 "Studio": {"v1_bands": [10500, 12500], "kmeans": [10384, 12194], "jenks": [10488, 12194]},
                                 "1899 Penn": {"v1_bands": [3300, 4200], "kmeans": [3534, 4388], "jenks": [3632, 4688], "comment": "1899 is still growing ~90% YoY so absolute bands mix ramp with season; kept at 3,300/4,200 because Sep–Dec projections fall in the mid band"}},
        "capacity_observed": {c: cap[c]["busy_hours"] for c in cap},
        "capacity_planning": NV["cap"]["target"],
        "need_hours_by_capacity": NV["capacity_sensitivity"],
        "august_2026": v10["august_2026"], "pl_bridge_validation": v10["pl_bridge_validation"],
        "levers": {"hours_placement_monthly_headroom_by_tier": inc["headroom_by_tier"], "rate_mix": v10["levers"]["rate_mix"]},
        "templates": templates,
        "standard_rates": T["rates"], "targets": T["targets"],
        "incentive": {k: inc[k] for k in ["assumptions", "cafe_economics", "support_burden_month_normalized", "plan_comparison", "annual_simulation", "margins_by_tier", "profit_share_sizing", "team", "headroom_aug_total", "headroom_aprjun_avg_total"]},
        "data_notes": {**v10["data_notes"], "optimizer": "PuLP 3.3 / CBC; scripts ldd_10_need_v11.py, ldd_11_optimize_templates.py, ldd_12_incentive_v11.py"},
    }
    json.dump(out, open(v10_path, "w"), indent=1)
    print("wrote", v10_path)
    for cafe, t in templates.items():
        for tier, v in t.items():
            w = v["week"]; print(f"  {cafe:10} {tier:5} {w['hours']:6.1f}h ${w['wages']:,} {w['labor_pct']}% (obs {w['labor_pct_obs_rate']}%) variants {v['variants_week']}")

if __name__ == "__main__":
    main()
