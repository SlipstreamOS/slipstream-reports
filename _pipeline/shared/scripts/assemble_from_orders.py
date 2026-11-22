#!/usr/bin/env python3
"""Assemble the transactions CSV layer from spilled Square orders.search results.

Orders-only flow (payments.list pages are too small to spill and would flood
context — see memory project_weekly_pull_spill_dir_gap). Full order objects
carry tenders[] with payment_id, amount_money (INCLUDES tip — verified vs
net_amounts.total_money on live data 2026-08-02), tip_money, card_details,
type. We synthesize one payment per tender (skipping NO_SALE drawer-opens)
and reuse build_transactions_csv.build_csv via monkeypatched load_api_data,
so output rows match the Square Transactions CSV export shape exactly.

Cross-month windows are split so each calendar month's rows land in the
right transactions/YYYY-MM.csv file.

Usage: python3 assemble_from_orders.py START END   (YYYY-MM-DD, inclusive)
Then:  python3 build_daily_metrics.py START END
"""
import os, sys, json, glob
from datetime import date, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))   # shared/scripts
sys.path.insert(0, HERE)
import build_transactions_csv as B

# Walk up from this file to the session root containing mnt/.claude
# (same technique as build_transactions_csv).
_p = HERE
SESSION_ROOT = None
while _p != "/":
    if os.path.isdir(os.path.join(_p, "mnt", ".claude")):
        SESSION_ROOT = _p
        break
    _p = os.path.dirname(_p)
SPILL_GLOB = os.path.join(SESSION_ROOT or "/", "mnt", ".claude", "projects",
                          "*", "*", "tool-results", "mcp-*make_api_request*.txt")

def load_orders():
    orders = {}
    # Spilled tool-result pages + any inline-page rescue files we wrote to raw/.
    # (A final page small enough to fit under the token cap is returned INLINE
    #  and never spilled; the agent extracts it from the session transcript and
    #  drops it in shared/raw/inline_orders_*.json so it isn't silently lost.)
    files = sorted(glob.glob(SPILL_GLOB)) + sorted(
        glob.glob(os.path.join(HERE, "..", "raw", "inline_orders_*.json")))
    for f in files:
        try:
            raw = open(f).read()
        except Exception:
            continue
        if not raw.strip().startswith("{") or '"orders"' not in raw[:200]:
            continue
        try:
            j = json.loads(raw)
        except Exception:
            print(f"  ! JSON parse failed: {f}")
            continue
        for o in j.get("orders", []):
            if o.get("id"):
                orders[o["id"]] = o
    print(f"Read {len(files)} spilled files -> {len(orders)} unique orders")
    return orders

def synthesize_payments(orders):
    payments = {}
    skipped_no_sale = 0
    for o in orders.values():
        for t in o.get("tenders", []) or []:
            if (t.get("type") or "").upper() == "NO_SALE":
                skipped_no_sale += 1
                continue
            pid = t.get("payment_id") or t.get("id")
            if not pid:
                continue
            total = (t.get("amount_money") or {}).get("amount", 0)   # includes tip
            tip = (t.get("tip_money") or {}).get("amount", 0)
            payments[pid] = {
                "id": pid,
                "order_id": o.get("id"),
                "location_id": t.get("location_id") or o.get("location_id"),
                "created_at": t.get("created_at") or o.get("closed_at") or o.get("created_at"),
                "status": "COMPLETED",
                "amount_money": {"amount": total - tip, "currency": "USD"},
                "tip_money": {"amount": tip, "currency": "USD"},
                "total_money": {"amount": total, "currency": "USD"},
                "source_type": (t.get("type") or "").upper(),
                "card_details": t.get("card_details") or {},
                "customer_id": o.get("customer_id", ""),
                "team_member_id": "",
            }
    print(f"Synthesized {len(payments)} payments (skipped {skipped_no_sale} NO_SALE)")
    return payments

def month_windows(start, end):
    """Split [start, end] into per-calendar-month sub-windows."""
    s = date.fromisoformat(start); e = date.fromisoformat(end)
    out = []
    cur = s
    while cur <= e:
        if cur.month == 12:
            month_end = date(cur.year, 12, 31)
        else:
            month_end = date(cur.year, cur.month + 1, 1) - timedelta(days=1)
        w_end = min(month_end, e)
        out.append((cur.isoformat(), w_end.isoformat()))
        cur = w_end + timedelta(days=1)
    return out

def main(start, end):
    orders = load_orders()
    by_loc = {}
    for o in orders.values():
        by_loc[o.get("location_id")] = by_loc.get(o.get("location_id"), 0) + 1
    for loc, n in sorted(by_loc.items()):
        print(f"  {B.LOC_TO_CAFE.get(loc, loc)}: {n} orders")
    payments = synthesize_payments(orders)
    B.load_api_data = lambda: (payments, orders)
    for ws, we in month_windows(start, end):
        print(f"--- build_csv {ws}..{we}")
        B.build_csv(ws, we)

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
