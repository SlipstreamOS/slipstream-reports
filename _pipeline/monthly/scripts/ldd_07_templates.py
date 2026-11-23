#!/usr/bin/env python3
"""Labor Deep Dive step 7 — schedule TEMPLATES per cafe × tier × DOW.
Shift lists are designed against the half-hour need curves (need_vs_actual.json) and scored:
  hours, wages at STANDARD role rates (blue inputs), core labor % of the tier's typical daily/weekly
  sales, SPLH, coverage check vs need, delta vs sampled actual weeks, and -10% sales sensitivity.
Ryan's first-cut template week (Square Teams, 10/11–10/17) is scored on the same basis for the diff.
Output: data/labor_deep_dive/templates.json
"""
import os, json, math
import numpy as np
HERE = os.path.dirname(os.path.abspath(__file__))
DD = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
DOWN = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

# ---------------- INPUT ASSUMPTIONS (blue in the workbook) ----------------
RATES = {"SL": 21.00, "PL": 19.00, "PB": 19.00, "B": 19.00, "HR": 19.00, "CL": 19.50, "FOH": 20.50}
ROLE_NAMES = {"SL": "Shift Lead", "PL": "Production Line", "PB": "Production Barista", "B": "Barista",
              "HR": "Hospitality Register", "CL": "Closer (keyholder barista)", "FOH": "FOH"}
TARGET = {"Downtown": (0.32, 0.36), "Studio": (0.30, 0.34), "1899 Penn": (0.35, 0.40)}
CAP = 14.0

def S(role, start, end, note=""):
    return {"role": role, "start": start, "end": end, "note": note}

# ---------------- TEMPLATES ----------------
# Times are decimal hours (6.5 = 6:30). Notes flag Admin cover and options.
T = {
 "Downtown": {
  "slow": {
   "Mon": [S("PL", 6.0, 14.0, "food prep, then bar support"), S("SL", 6.5, 16.5, "open-to-close lead (10h option; alt: SL 6:30–14:30 + CL 12:30–16:30)")],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.5), S("CL", 8.5, 16.5, "rush 8:30–10 as 3rd, closes")],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 8.5, 16.5, "rush as 3rd, closes")],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 8.5, 16.5, "rush as 3rd, closes")],
   "Fri": [S("PL", 6.0, 13.0, "prep + rush; leaves after lunch"), S("SL", 6.5, 16.5, "open-to-close lead (10h option); Admin covers 8–10 if scheduled")],
  },
  "mid": {
   "Mon": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 9.0, 16.5, "3rd through rush/lunch, closes")],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.5, 11.5, "rush barista (4h); DROP when Admin covers 8–10"), S("CL", 8.5, 16.5, "4th through rush, closes")],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.5, 12.0, "rush barista; DROP when Admin covers 8–10"), S("CL", 8.5, 16.5)],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.5, 11.5, "rush barista (4h); DROP when Admin covers 8–10"), S("CL", 8.5, 16.5)],
   "Fri": [S("PL", 6.0, 13.0), S("SL", 6.5, 15.0), S("CL", 9.0, 16.5)],
  },
  "slow_lean": {
   "Mon": [S("PL", 6.0, 14.0), S("SL", 6.5, 16.5, "open-to-close lead (10h option)")],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 10.0, 16.5, "2 people through the 8:30–10 rush; Admin covers rush when scheduled")],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 10.0, 16.5, "Admin covers rush when scheduled")],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 10.0, 16.5, "Admin covers rush when scheduled")],
   "Fri": [S("PL", 6.0, 12.0, "prep + rush only"), S("SL", 6.5, 16.5, "open-to-close lead (10h option)")],
  },
  "busy": {
   "Mon": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.5, 11.5, "rush barista"), S("CL", 10.0, 16.5)],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.5, 12.0, "rush barista"), S("PB", 8.0, 15.5, "4th: rush + lunch + 2–3:30p (26–29/hr)"), S("CL", 10.0, 16.5)],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.0, 12.0, "rush barista"), S("PB", 8.0, 15.5, "4th: rush + lunch + afternoon"), S("CL", 10.0, 16.5)],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.0, 12.0, "rush barista"), S("PB", 8.0, 15.5, "4th: rush + lunch + afternoon"), S("CL", 10.0, 16.5)],
   "Fri": [S("PL", 6.0, 13.0), S("SL", 6.5, 15.0), S("CL", 9.0, 16.5, "3 people 9–1 covers the 30/hr Friday peak")],
  },
 },
 "Studio": {
  "slow": {
   "Mon": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 6.5, 10.5, "3rd at open for catering delivery; DROP when Admin (Manager) is on the floor at open")],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75, "2-person day; Admin covers 8–10 rush when scheduled")],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75, "2-person day; Admin covers 8–10 rush when scheduled")],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75, "2-person day; Admin covers 8–10 rush when scheduled")],
   "Fri": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("CL", 9.0, 17.5)],
   "Sat": [S("PL", 7.0, 15.0), S("SL", 7.5, 15.5), S("HR", 8.5, 16.5), S("CL", 10.5, 17.5)],
   "Sun": [S("PL", 7.0, 15.0), S("SL", 7.5, 15.5), S("B", 9.0, 14.0, "mid-morning/lunch 3rd"), S("CL", 10.0, 17.5)],
  },
  "mid": {
   "Mon": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 6.5, 11.5, "3rd at open (catering) + rush; DROP to 8–11:30 when Admin is on at open")],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75, "2-person day; Admin covers 8–10 rush when scheduled")],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 8.0, 12.0, "rush barista; DROP when Admin covers 8–10")],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 8.0, 12.0, "rush barista; DROP when Admin covers 8–10")],
   "Fri": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 8.0, 13.0, "rush + lunch"), S("CL", 10.5, 17.5)],
   "Sat": [S("PL", 7.0, 15.0), S("SL", 7.5, 15.5), S("HR", 8.5, 16.5), S("B", 9.5, 14.5, "peak 4th"), S("CL", 10.5, 17.5)],
   "Sun": [S("PL", 7.0, 15.0), S("SL", 7.5, 15.5), S("HR", 8.5, 13.5), S("B", 9.5, 14.5, "peak 4th"), S("CL", 10.0, 17.5)],
  },
  "busy": {
   "Mon": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 6.5, 12.5, "3rd at open (catering) + rush/lunch; 8–12:30 when Admin is on at open")],
   "Tue": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 7.5, 11.5, "rush barista")],
   "Wed": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 7.5, 11.5, "rush barista")],
   "Thu": [S("PL", 6.0, 14.0), S("SL", 6.5, 14.75), S("B", 7.5, 11.5, "rush barista")],
   "Fri": [S("PL", 6.0, 14.0), S("SL", 6.5, 15.0), S("B", 7.5, 12.5, "rush barista"), S("PB", 8.5, 13.5, "peak 4th"), S("CL", 10.5, 17.5)],
   "Sat": [S("PL", 7.0, 15.0), S("SL", 7.5, 15.5), S("HR", 8.0, 16.0), S("B", 9.0, 14.0), S("PB", 9.5, 14.5, "peak 5th"), S("CL", 10.5, 17.5)],
   "Sun": [S("PL", 7.0, 15.0), S("SL", 7.5, 15.5), S("HR", 8.5, 16.5), S("B", 9.5, 14.5), S("CL", 10.0, 17.5)],
  },
 },
 "1899 Penn": {
  "slow": {
   "Mon": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only (4h); Admin covers lunch 12:30–2 when scheduled")],
   "Tue": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.0)],
   "Wed": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.0)],
   "Thu": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.0)],
   "Fri": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only (4h)")],
  },
  "slow_lean": {
   "Mon": [S("SL", 6.5, 15.5, "1-person day; Admin covers 8:30–10 and 12:30–2 when scheduled")],
   "Tue": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only")],
   "Wed": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only")],
   "Thu": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only")],
   "Fri": [S("SL", 6.5, 15.5, "1-person day; Admin covers rush/lunch when scheduled")],
  },
  "mid": {
   "Mon": [S("SL", 6.5, 15.5), S("PB", 7.5, 13.5)],
   "Tue": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.5)],
   "Wed": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.5)],
   "Thu": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.5)],
   "Fri": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only (4h); Admin covers lunch when scheduled")],
  },
  "mid_lean": {
   "Mon": [S("SL", 6.5, 15.5), S("PB", 8.0, 12.0, "rush only; Admin covers lunch when scheduled")],
   "Tue": [S("SL", 6.5, 15.5), S("PB", 7.5, 13.5, "rush + lunch; SL alone 1:30–3:30 (13–22/hr)")],
   "Wed": [S("SL", 6.5, 15.5), S("PB", 7.5, 13.5)],
   "Thu": [S("SL", 6.5, 15.5), S("PB", 7.5, 13.5)],
   "Fri": [S("SL", 6.5, 15.5, "1-person Friday (61 txns/day); Admin covers 8:30–10 & 12:30–1:30 when scheduled")],
  },
  "busy": {
   "Mon": [S("SL", 6.5, 15.5), S("PB", 7.5, 14.0)],
   "Tue": [S("SL", 6.5, 15.5), S("PB", 6.5, 15.0, "1p peak 37/hr is at 2-person capacity; Admin covers 12:30–2 when scheduled")],
   "Wed": [S("SL", 6.5, 15.5), S("PB", 6.5, 15.0, "1p peak 30/hr; Admin covers 12:30–2 when scheduled")],
   "Thu": [S("SL", 6.5, 15.5), S("PB", 6.5, 15.0)],
   "Fri": [S("SL", 6.5, 15.5), S("PB", 7.5, 12.5)],
  },
 },
 "NMWA": {
  "slow": {
   "Wed": [S("FOH", 9.5, 14.5, "2nd Wednesday community day only: add 2nd FOH 10–14"), S("FOH", 10.0, 14.0, "community-day 2nd person")],
   "Thu": [S("FOH", 9.5, 14.5)], "Fri": [S("FOH", 9.5, 14.5)], "Sat": [S("FOH", 9.5, 14.5)],
   "Sun": [S("FOH", 9.5, 14.5, "1st Sunday community day: add 2nd FOH 10–14")],
  }
 }
}

# Ryan's first-cut template (Square Teams, week of 10/11–10/17), core shifts only (Manager/Commissary/Roasting excluded)
def R(role, start, end, rate):
    return {"role": role, "start": start, "end": end, "note": "", "rate": rate}
RYAN = {
 "Downtown": {
  "Mon": [R("PB", 6.5, 13.0, 19.00), R("PL", 6.2, 16.5, 22.75), R("SL", 9.0, 16.5, 21.75)],
  "Tue": [R("PL", 6.25, 14.0, 18.50), R("PB", 7.5, 16.5, 22.75), R("SL", 6.5, 15.0, 21.75), R("PB", 9.0, 16.5, 25.25)],
  "Wed": [R("SL", 8.0, 16.5, 19.00), R("PB", 6.5, 13.5, 21.75), R("PL", 6.25, 15.0, 19.25), R("B", 9.0, 16.5, 25.25)],
  "Thu": [R("SL", 8.0, 16.5, 19.00), R("B", 6.5, 13.5, 21.75), R("PL", 6.25, 15.0, 19.25), R("PB", 9.0, 16.5, 25.25)],
  "Fri": [R("PB", 9.5, 16.5, 19.00), R("SL", 6.5, 13.0, 21.75), R("PL", 6.25, 16.5, 19.25)],
 },
 "Studio": {
  "Sun": [R("B", 7.5, 14.0, 18.50), R("SL", 8.0, 15.0, 24.75), R("PL", 7.25, 15.5, 18.50), R("PB", 10.5, 17.5, 18.50)],
  "Mon": [R("SL", 6.5, 14.5, 20.75), R("PB", 6.5, 13.0, 19.00), R("PL", 6.25, 14.5, 18.50)],        # + Manager 7–2 (admin)
  "Tue": [R("PB", 6.5, 13.0, 20.25), R("SL", 7.0, 14.5, 24.75), R("PL", 6.25, 14.5, 18.50)],
  "Wed": [R("B", 6.5, 13.0, 19.00), R("SL", 7.0, 14.5, 22.75), R("PL", 6.25, 14.5, 18.50)],
  "Thu": [R("SL", 7.0, 14.5, 20.25), R("PL", 6.25, 14.5, 18.50), R("PB", 6.5, 13.5, 24.75)],       # + Manager 9–2 (admin)
  "Fri": [R("SL", 6.5, 15.0, 20.25), R("PL", 6.25, 14.5, 22.75), R("SL", 10.0, 17.25, 20.75), R("B", 7.0, 13.0, 18.50)],
  "Sat": [R("PL", 7.25, 15.0, 18.50), R("SL", 10.0, 17.5, 18.50), R("SL", 7.5, 13.0, 20.75), R("HR", 8.0, 15.0, 22.75), R("B", 9.0, 16.0, 18.50)],
 },
 "1899 Penn": {
  "Mon": [R("SL", 6.5, 15.5, 20.50), R("PB", 6.5, 13.5, 22.75)],                                 # + Manager 8–2 (admin)
  "Tue": [R("SL", 6.5, 15.5, 20.50), R("PB", 6.5, 15.5, 18.50)],                                 # + Manager 6:30–1:30 (admin)
  "Wed": [R("SL", 6.5, 15.5, 20.75), R("PB", 6.5, 15.5, 18.50)],
  "Thu": [R("SL", 6.5, 15.5, 22.75), R("B", 6.5, 15.0, 18.50)],
  "Fri": [R("SL", 6.5, 15.5, 20.75), R("PB", 6.5, 13.5, 22.75)],
 },
 "NMWA": {"Sun": [R("FOH", 9.0, 15.0, 20.25)], "Thu": [R("FOH", 9.0, 15.0, 20.50)], "Fri": [R("FOH", 9.0, 15.0, 22.75)], "Sat": [R("FOH", 9.0, 15.0, 20.50)]},
}
RYAN_ADMIN = {"Studio": {"Mon": 7.0, "Thu": 5.0}, "1899 Penn": {"Mon": 6.0, "Tue": 7.0}, "NMWA": {"Fri": 4.0}, "Downtown": {}}
# Ryan's actual wage rates for the people on those shifts (from the Square Teams $ totals) — for the "as-staffed" cost
RYAN_RATE_AVG = {"Downtown": 20.72, "Studio": 19.78, "1899 Penn": 20.89, "NMWA": 20.42}

def fmt(t):
    h = int(t); m = int(round((t - h) * 60)); return f"{h}:{m:02d}"

def score(shifts, need_rows, sales, rate_override=None):
    grid = [r["t"] for r in need_rows]
    staffed = np.zeros(len(grid))
    hours = 0.0; wages = 0.0
    for s in shifts:
        h = s["end"] - s["start"]; hours += h
        r = s.get("rate") or rate_override or RATES[s["role"]]
        wages += h * r
        for i, g in enumerate(grid):
            if s["start"] <= g < s["end"]:
                staffed[i] += 1
    need = np.array([r["need"] for r in need_rows]); tx = np.array([r["txns_hr"] for r in need_rows])
    under = np.clip(need - staffed, 0, None)
    # pressure: max txns per staffed person-hour during open hours
    with np.errstate(divide="ignore", invalid="ignore"):
        tph = np.where(staffed > 0, tx / staffed, 0)
    return {"hours": round(hours, 2), "wages": round(wages, 2), "labor_pct": round(wages / sales * 100, 1) if sales else None,
            "splh": round(sales / hours, 1) if hours else None, "peak_staff": int(staffed.max()),
            "under_halfhours": int((under > 0).sum()), "under_person_hours": round(float(under.sum() * 0.5), 1),
            "max_tph": round(float(tph.max()), 1), "over_person_hours": round(float(np.clip(staffed - need, 0, None).sum() * 0.5), 1)}

def main():
    nv = json.load(open(os.path.join(DD, "need_vs_actual.json")))
    tiers = json.load(open(os.path.join(DD, "tiers.json")))
    staffing = json.load(open(os.path.join(DD, "staffing.json")))
    out = {"rates": RATES, "cap": CAP, "targets": {k: {"target": v[0], "ceiling": v[1]} for k, v in TARGET.items()}, "cafes": {}}
    for cafe, tiers_t in T.items():
        out["cafes"][cafe] = {}
        for tier, days in tiers_t.items():
            wk_sales = 0; wk_wages = 0; wk_hours = 0; wk_act_wages = 0; wk_act_hours = 0; wk_ryan_w = 0; wk_ryan_h = 0
            days_out = {}
            base = tier.split("_")[0]
            for dow, shifts in days.items():
                if cafe == "NMWA":
                    sales = tiers[cafe]["typical_day"]["slow"].get(dow, {}).get("net", 0)
                    need_rows = [{"t": float(t), "need": (2 if dow == "Wed" else 1) if 10 <= t < 14 else 0, "txns_hr": 0} for t in np.arange(9, 15.5, 0.5)]
                    act = staffing["staffing"][cafe].get("slow", {}).get(dow)
                else:
                    sales = tiers[cafe]["typical_day"][base][dow]["net"]
                    need_rows = nv["cafes"][cafe]["tiers"][base][dow]["by_halfhour"]
                    act = staffing["staffing"][cafe][base][dow]
                sc = score(shifts, need_rows, sales)
                rec = {"sales_typical": sales, "shifts": [{"role": ROLE_NAMES[s["role"]], "code": s["role"], "start": fmt(s["start"]), "end": fmt(s["end"]),
                                                            "hours": round(s["end"] - s["start"], 2), "rate": RATES[s["role"]], "wages": round((s["end"] - s["start"]) * RATES[s["role"]], 2), "note": s["note"]} for s in shifts],
                       "template": sc}
                if act:
                    rec["actual_sampled"] = {"core_hours": act["core_hours_per_day"], "core_wages": act["core_wages_per_day"], "avg_rate": act["avg_rate"],
                                             "labor_pct": round(act["core_wages_per_day"] / sales * 100, 1), "n_days": act["n_days"]}
                    rec["delta_vs_actual"] = {"hours": round(sc["hours"] - act["core_hours_per_day"], 1), "wages": round(sc["wages"] - act["core_wages_per_day"], 0)}
                    wk_act_wages += act["core_wages_per_day"]; wk_act_hours += act["core_hours_per_day"]
                if dow in RYAN.get(cafe, {}):
                    ry = score(RYAN[cafe][dow], need_rows, sales, rate_override=RYAN_RATE_AVG[cafe])
                    ry["admin_hours"] = RYAN_ADMIN.get(cafe, {}).get(dow, 0)
                    rec["ryan_first_cut"] = ry
                    rec["delta_vs_ryan"] = {"hours": round(sc["hours"] - ry["hours"], 2), "wages": round(sc["wages"] - ry["wages"], 0)}
                    wk_ryan_w += ry["wages"]; wk_ryan_h += ry["hours"]
                # sensitivity: sales -10%
                rec["sensitivity_minus10"] = round(sc["wages"] / (sales * 0.9) * 100, 1) if sales else None
                days_out[dow] = rec
                wk_sales += sales; wk_wages += sc["wages"]; wk_hours += sc["hours"]
            tgt = TARGET.get(cafe, (None, None))
            week = {"sales_typical": round(wk_sales, 0), "hours": round(wk_hours, 1), "wages": round(wk_wages, 0),
                    "labor_pct": round(wk_wages / wk_sales * 100, 1) if wk_sales else None,
                    "splh": round(wk_sales / wk_hours, 1) if wk_hours else None,
                    "labor_pct_minus10": round(wk_wages / (wk_sales * 0.9) * 100, 1) if wk_sales else None,
                    "actual_sampled_hours": round(wk_act_hours, 1), "actual_sampled_wages": round(wk_act_wages, 0),
                    "actual_sampled_labor_pct": round(wk_act_wages / wk_sales * 100, 1) if wk_sales and wk_act_wages else None,
                    "ryan_first_cut_hours": round(wk_ryan_h, 1), "ryan_first_cut_wages": round(wk_ryan_w, 0),
                    "ryan_first_cut_labor_pct": round(wk_ryan_w / wk_sales * 100, 1) if wk_sales and wk_ryan_w else None}
            if tgt[0]:
                week["target_wages"] = round(wk_sales * tgt[0], 0); week["ceiling_wages"] = round(wk_sales * tgt[1], 0)
                # trigger: weekly sales below which this template breaches the ceiling -> drop to next slower template
                week["ceiling_trigger_sales"] = round(wk_wages / tgt[1], 0)
                week["target_trigger_sales"] = round(wk_wages / tgt[0], 0)
            out["cafes"][cafe][tier] = {"days": days_out, "week": week}
    json.dump(out, open(os.path.join(DD, "templates.json"), "w"), indent=1)
    # console
    for cafe, tiers_o in out["cafes"].items():
        print(f"\n=== {cafe} ===")
        for tier, o in tiers_o.items():
            w = o["week"]
            print(f" {tier:5} week: sales ${w['sales_typical']:,.0f}  tmpl {w['hours']}h ${w['wages']:,.0f} = {w['labor_pct']}%  (-10% sales -> {w['labor_pct_minus10']}%)  SPLH ${w['splh']}"
                  f" | actual {w['actual_sampled_hours']}h ${w['actual_sampled_wages']:,.0f} = {w['actual_sampled_labor_pct']}% | Ryan {w['ryan_first_cut_hours']}h ${w['ryan_first_cut_wages']:,.0f} = {w['ryan_first_cut_labor_pct']}%"
                  + (f" | ceiling trigger ${w['ceiling_trigger_sales']:,.0f}" if 'ceiling_trigger_sales' in w else ""))
            for dow, d in o["days"].items():
                t = d["template"]
                a = d.get("actual_sampled", {})
                print(f"    {dow}: ${d['sales_typical']:,.0f}  {t['hours']:5.2f}h ${t['wages']:6.0f} {t['labor_pct'] or 0:5.1f}%  peak {t['peak_staff']} under {t['under_person_hours']}ph maxtph {t['max_tph']} | act {a.get('core_hours','-')}h {a.get('labor_pct','-')}%")

if __name__ == "__main__":
    main()
