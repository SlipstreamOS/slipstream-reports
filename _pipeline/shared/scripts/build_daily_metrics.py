#!/usr/bin/env python3
"""Roll payment-level transactions up into the LEAN DAILY STORE we keep long-term.

We do NOT keep every transaction. We keep one row per (date, cafe) with just the
daily numbers the model + reports need. Source can be either the transactions
store (historical monthly CSVs) or a freshly-pulled week; either way the output
is the same compact daily_metrics.csv.

Output: _pipeline/shared/daily_metrics.csv  (idempotent by date+cafe)
Columns: date, cafe, transactions, net_sales, tips, coffee_bev, food,
         retail_coffee, other, avg_ticket

Usage:  python3 build_daily_metrics.py [start YYYY-MM-DD] [end YYYY-MM-DD]
        (no args = rebuild the whole store from all available transactions)
"""
import os, sys, csv
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))                 # shared/scripts
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "weekly", "scripts"))
import transactions_store as ts
from aggregate_week import classify_item   # reuse the exact category classifier

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "daily_metrics.csv")
COLS = ["date","cafe","transactions","net_sales","tips","coffee_bev","food","retail_coffee","other",
        "avg_ticket","walkin_transactions","walkin_net_sales"]

# Ryan rule (Aug 2026): avg ticket and traffic metrics EXCLUDE invoice-sourced
# transactions (billed catering/events/museum fees are not walk-in tickets).
# Revenue totals KEEP them. So `transactions`/`net_sales` stay all-in for
# back-compat, `walkin_*` carry the ex-invoice slice, and `avg_ticket` is
# derived from walk-in only.
def _is_invoice(row):
    return "invoice" in str(row.get("source", "")).lower()

def build(start=None, end=None):
    df = ts.load()
    if start: df = df[df["date"] >= start]
    if end:   df = df[df["date"] <= end]
    rows = {}
    # group by (date, cafe)
    for (d, cafe), g in df.groupby(["date","cafe"]):
        tx = g["transaction_id"].nunique()
        net = g["net_sales_cents"].sum() / 100.0
        tips = g["tip_cents"].sum() / 100.0
        # Walk-in slice (ex-invoice) drives avg ticket + traffic metrics.
        if "source" in g.columns:
            gw = g[~g["source"].astype(str).str.lower().str.contains("invoice", na=False)]
        else:
            gw = g
        w_tx = gw["transaction_id"].nunique()
        w_net = gw["net_sales_cents"].sum() / 100.0
        cat = {"coffee_bev":0.0,"food":0.0,"retail_coffee":0.0,"other":0.0}
        for _, r in g.iterrows():
            c = classify_item(str(r.get("description",""))) or "other"
            if c not in cat: c = "other"
            cat[c] += r["net_sales_cents"] / 100.0
        rows[(d,cafe)] = {"date":d,"cafe":cafe,"transactions":int(tx),
            "net_sales":round(net,2),"tips":round(tips,2),
            "coffee_bev":round(cat["coffee_bev"],2),"food":round(cat["food"],2),
            "retail_coffee":round(cat["retail_coffee"],2),"other":round(cat["other"],2),
            "avg_ticket":round(w_net/w_tx,2) if w_tx else 0,
            "walkin_transactions":int(w_tx),
            "walkin_net_sales":round(w_net,2)}
    # UPSERT into existing store
    existing = {}
    if os.path.exists(OUT):
        with open(OUT) as f:
            for row in csv.DictReader(f):
                existing[(row["date"],row["cafe"])] = row
    existing.update(rows)
    out = sorted(existing.values(), key=lambda r:(str(r["date"]), str(r["cafe"])))
    with open(OUT,"w",newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS); w.writeheader()
        for r in out: w.writerow({k:r.get(k,"") for k in COLS})
    return len(rows), len(out)

if __name__ == "__main__":
    a = sys.argv
    n_new, n_total = build(a[1] if len(a)>1 else None, a[2] if len(a)>2 else None)
    print(f"daily_metrics.csv: upserted {n_new} (date,cafe) rows; {n_total} total. → {os.path.abspath(OUT)}")
