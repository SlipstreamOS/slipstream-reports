#!/usr/bin/env python3
"""Build the weekly per-cafe per-day sales JSON from the lean daily_metrics.csv store.

This replaces aggregate_week.py in the "post-scheduled-task" world (July 2026+).
The scheduled task `slipstream-weekly-data-pull` runs Sunday morning and refreshes
_pipeline/shared/daily_metrics.csv from the Square Orders + Payments pull. This
script converts those daily rows into the `week_<label>.json` schema that
build_weekly_data.py + build_weekly_report.py expect.

Downstream, patch_week_from_txstore.py still runs on top of this file to fill in
tax/discount and catering vs core split (from _pipeline/shared/transactions/YYYY-MM.csv),
so we only need to seed transactions, net_sales, tips, and by-category from
daily_metrics.csv here.

Run:
  python3 build_week_from_daily_metrics.py

Reads WEEK_START/WEEK_END below (Sun–Sat), writes data/week_<label>.json.
"""
import os, csv, json
from datetime import datetime, timedelta
from collections import defaultdict

from _paths import DATA_DIR

# ---- Week config ----------
WEEK_START = "2026-11-15"
WEEK_END   = "2026-11-21"
WEEK_LABEL = "nov15-nov21"
# ---------------------------

HERE = os.path.dirname(os.path.abspath(__file__))
DAILY_CSV = os.path.abspath(os.path.join(HERE, "..", "..", "shared", "daily_metrics.csv"))
OUT = os.path.join(DATA_DIR, f"week_{WEEK_LABEL}.json")

CAFES = ["Downtown", "Studio", "1899 Penn", "NMWA"]


def _f(v):
    try: return float(v or 0)
    except: return 0.0


def to_cents(v):
    return int(round(_f(v) * 100))


def main():
    if not os.path.exists(DAILY_CSV):
        print(f"ERROR: {DAILY_CSV} not found. Did the weekly data pull run?")
        return 1

    per = defaultdict(dict)
    total_in_window = 0
    for r in csv.DictReader(open(DAILY_CSV)):
        d = r["date"]
        if d < WEEK_START or d > WEEK_END:
            continue
        cafe = r["cafe"]
        if cafe not in CAFES:
            continue
        tx = int(r["transactions"]) if r["transactions"] else 0
        net_c = to_cents(r["net_sales"])
        tip_c = to_cents(r["tips"])
        cbev_c = to_cents(r["coffee_bev"])
        food_c = to_cents(r["food"])
        ret_c = to_cents(r["retail_coffee"])
        oth_c = to_cents(r["other"])
        per[cafe][d] = {
            "orders": tx,
            "net_sales_cents": net_c,
            "discount_cents": 0,      # patch_week_from_txstore will populate
            "tax_cents": 0,           # patch_week_from_txstore will populate
            "tip_cents": tip_c,
            "grand_total_cents": net_c + tip_c,
            "by_category": {
                "coffee_bev": cbev_c,
                "food": food_c,
                "retail_coffee": ret_c,
                "other": oth_c,
                "catering": 0,         # patch_week_from_txstore will populate
            },
            "by_fulfillment": {},
            "items": {},
            "avg_ticket_cents": (net_c / tx) if tx else 0,
            "tip_pct_of_net": (tip_c / net_c * 100) if net_c else 0,
        }
        total_in_window += tx

    os.makedirs(DATA_DIR, exist_ok=True)
    payload = {
        "week_start": WEEK_START,
        "week_end": WEEK_END,
        "generated_at": datetime.now().isoformat(),
        "per_cafe_per_day": per,
        "summary": {
            "orders_in_week_total": total_in_window,
            "source": "daily_metrics.csv (built by scheduled task slipstream-weekly-data-pull)",
        },
    }
    json.dump(payload, open(OUT, "w"), indent=2)
    print(f"Saved → {OUT}")
    print(f"Orders in window: {total_in_window}")
    for cafe in CAFES:
        n = sum(d["orders"] for d in per.get(cafe, {}).values())
        net = sum(d["net_sales_cents"] for d in per.get(cafe, {}).values())
        tips = sum(d["tip_cents"] for d in per.get(cafe, {}).values())
        days = len(per.get(cafe, {}))
        print(f"  {cafe:<11}  days={days}  n={n:>5}  net=${net/100:>10,.2f}  tips=${tips/100:>7,.2f}")


if __name__ == "__main__":
    main()
