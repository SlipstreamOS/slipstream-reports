#!/usr/bin/env python3
"""Aggregate a Sun–Sat week of Square payments into per-cafe per-day register totals.

Reads raw/payments_week.json (compact extract from paginated payments.list calls)
and writes data/payments_<week>.json for the report builder.

Why payments and not orders? Square's Sales Summary report (and the daily
transaction count Ryan sees on the iPad) is keyed off Payments — tips can be
adjusted post-order-close on the Payment record and the Order's tip field
won't reflect that. Reading from Payments matches what Ryan sees in Square.

Output structure:
{
  "week_start": "...", "week_end": "...",
  "per_cafe_per_day": {
    "Downtown": {
      "2026-04-13": {
        "register_payments": 324,         # COMPLETED payments excluding EXTERNAL catering
        "register_amount_cents": 289640,  # sum of amount_money (pre-tip; subtotal)
        "register_tip_cents": 33312,      # sum of tip_money
        "register_total_cents": 322952,   # amount + tip
        "catering_payments": 0,           # EXTERNAL source (e.g. invoice payments)
        "catering_amount_cents": 0,
      }, ...
    }, ...
  }
}

Register vs catering split mirrors the orders pipeline's Core/Catering split,
so downstream `tip_cents` / `tx_count` lines always line up with Core Sales.

EXTERNAL source = Square Invoices / off-site payments (invoice paid outside
the register). Per Apr 28 2026 directive, only these get pulled out of Core.
Anything rung at the cafe register (CARD/CASH source, regardless of item
name) counts as Core Sales.
"""
import os, json
from datetime import datetime, timezone, timedelta
from collections import defaultdict

from _paths import DATA_DIR, RAW_DIR

EDT = timezone(timedelta(hours=-4))
WEEK_START = "2026-06-28"
WEEK_END = "2026-07-04"
WEEK_LABEL = "jun28-jul4"

IN  = os.path.join(RAW_DIR, "payments_week.json")
OUT = os.path.join(DATA_DIR, f"payments_{WEEK_LABEL}.json")

LOC_TO_CAFE = {
    "LXM7V1QWDHKZJ": "Downtown",
    "LGVKCKX9B9DZT": "Studio",
    "L4ZBMGCWJ5C7R": "1899 Penn",
    "LV2D18H0DD9ES": "NMWA",
}


def main():
    if not os.path.exists(IN):
        print(f"ERROR: {IN} not found.")
        print("Pull payments via Square MCP payments.list (paginated by location, ASC by created_at).")
        return
    raw = json.load(open(IN))
    payments = raw.get("payments", [])

    # Dedupe by id defensively (paginated pulls can overlap if cursors are reused).
    seen = set()
    unique = []
    for p in payments:
        pid = p.get("id")
        if pid and pid not in seen:
            seen.add(pid)
            unique.append(p)
    if len(unique) != len(payments):
        print(f"Note: deduped {len(payments) - len(unique)} duplicate payment records.")

    per = defaultdict(lambda: defaultdict(lambda: {
        "register_payments": 0,
        "register_amount_cents": 0,
        "register_tip_cents": 0,
        "register_total_cents": 0,
        "catering_payments": 0,
        "catering_amount_cents": 0,
    }))

    for p in unique:
        if p.get("status") != "COMPLETED":
            continue
        cafe = LOC_TO_CAFE.get(p.get("location_id"))
        if not cafe:
            continue
        # Convert created_at (UTC RFC3339) to local EDT date
        dt = datetime.fromisoformat(p["created_at"].replace("Z","+00:00"))
        date = dt.astimezone(EDT).date().isoformat()
        if date < WEEK_START or date > WEEK_END:
            continue
        b = per[cafe][date]
        amt = p.get("amount_cents", 0)
        tip = p.get("tip_cents", 0)
        tot = p.get("total_cents", 0)
        if p.get("source_type") == "EXTERNAL":
            b["catering_payments"] += 1
            b["catering_amount_cents"] += amt
        else:
            b["register_payments"] += 1
            b["register_amount_cents"] += amt
            b["register_tip_cents"] += tip
            b["register_total_cents"] += tot

    out = {c: dict(d) for c, d in per.items()}
    payload = {
        "week_start": WEEK_START,
        "week_end":   WEEK_END,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "per_cafe_per_day": out,
        "payments_processed": len(unique),
    }
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(OUT, "w") as fh:
        json.dump(payload, fh, indent=2)
    print(f"Saved → {OUT}")
    print(f"Payments processed: {len(unique)}")

    # Sanity print
    print(f"\n{'Cafe':<12} {'Reg #':>5}  {'Reg $':>10}  {'Reg Tip':>9}  {'Avg Tkt':>7} | {'Cat #':>5}  {'Cat $':>10}")
    for cafe in ["Downtown","Studio","1899 Penn","NMWA"]:
        days = out.get(cafe, {})
        reg_n  = sum(b["register_payments"]      for b in days.values())
        reg_a  = sum(b["register_amount_cents"]  for b in days.values())
        reg_t  = sum(b["register_tip_cents"]     for b in days.values())
        cat_n  = sum(b["catering_payments"]      for b in days.values())
        cat_a  = sum(b["catering_amount_cents"]  for b in days.values())
        avg = (reg_a/reg_n/100) if reg_n else 0
        print(f"{cafe:<12} {reg_n:>5}  ${reg_a/100:>8,.2f}  ${reg_t/100:>7,.2f}  ${avg:>5,.2f} | {cat_n:>5}  ${cat_a/100:>8,.2f}")


if __name__ == "__main__":
    main()
