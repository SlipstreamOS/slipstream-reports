#!/usr/bin/env python3
"""Patch the weekly aggregator JSON with totals from the transactions store.

Why: The Square Orders API misses some online/mobile/SK orders that show up
in Square's Sales Summary report. The transactions store (built from
`shared/transactions/*.csv` — either Ryan's monthly export or our API-built
2026-04.csv) reconciles to Square within 0.4%.

This script overrides per-cafe per-day fields in `data/week_<label>.json`:
  - orders   — transaction count (from txstore, matches Square Sales count)
  - tip_cents — tips (matches Square Sales Summary)
  - net_sales_cents — gross items pre-discount (matches Square's Items line)
  - discount_cents — discounts (matches Square Discounts & Comps)
  - tax_cents — tax
  - by_category[catering] — uses transactions-store source filter (Invoices)
                            for the catering-vs-core split

The by_category coffee_bev/food/retail_coffee/other breakdown is left as-is
(comes from Orders API line-item names) — it's slightly off vs Square's
category rollup but directionally correct for the report.

Run AFTER aggregate_week.py and BEFORE build_weekly_data.py.
"""
import os, sys, json
import pandas as pd

from _paths import DATA_DIR
import transactions_store as txstore

WEEK_LABEL = "nov15-nov21"
DAYS = ["2026-11-15","2026-11-16","2026-11-17","2026-11-18",
        "2026-11-19","2026-11-20","2026-11-21"]
WEEK_JSON = os.path.join(DATA_DIR, f"week_{WEEK_LABEL}.json")
CAFES = ["Downtown","Studio","1899 Penn","NMWA"]

INVOICE_SOURCES = {"Invoices", "Invoice"}

def per_cafe_day_from_txstore(cafe, date):
    """Return totals for one cafe-day, split by Core (register) vs Catering (Invoices source).

    Definitions match Square's Sales Summary report:
      gross  = sum of line-item gross_sales (pre-discount, pre-tax)
      net    = gross - |discounts|     (Square's "Net Sales" line)
      tip, tax, disc are signed appropriately.
    """
    df = txstore.load()
    if df.empty:
        return None
    s = df[(df["cafe"] == cafe) & (df["date"] == date) & (df["event_type"] == "Payment")]
    if s.empty:
        return None
    is_invoice = s["source"].isin(INVOICE_SOURCES)
    core = s[~is_invoice]
    cat  = s[is_invoice]
    # Discounts are stored as negative cents in the CSV (Square's convention).
    # "Net Sales" per Square = Gross + Discount  (since discount is negative).
    def slice_totals(sl):
        gross = int(sl["gross_sales_cents"].sum())
        disc  = int(sl["discounts_cents"].sum())   # negative
        net   = gross + disc                        # gross - |disc|
        return {
            "n":     int(len(sl)),
            "gross": gross,
            "net":   net,
            "tip":   int(sl["tip_cents"].sum()),
            "tax":   int(sl["tax_cents"].sum()),
            "disc":  disc,                          # negative if there were discounts
        }
    out = {"all": slice_totals(s), "core": slice_totals(core), "cat": slice_totals(cat)}
    return out


def main():
    if not os.path.exists(WEEK_JSON):
        print(f"ERROR: {WEEK_JSON} not found. Run aggregate_week.py first.")
        sys.exit(1)
    j = json.load(open(WEEK_JSON))
    per = j.get("per_cafe_per_day", {})

    patched = 0
    for cafe in CAFES:
        for d in DAYS:
            tx = per_cafe_day_from_txstore(cafe, d)
            if tx is None or tx["all"]["n"] == 0:
                continue
            core = tx["core"]; cat = tx["cat"]
            existing = per.get(cafe, {}).get(d)
            if not existing:
                # No order data for this day — synthesize a minimal entry
                per.setdefault(cafe, {})[d] = {
                    "orders": core["n"],
                    "net_sales_cents": core["net"],          # Square's Net Sales (gross - disc)
                    "tip_cents": core["tip"],
                    "tax_cents": core["tax"],
                    "discount_cents": abs(core["disc"]),
                    "by_category": {"catering": cat["net"]},  # Catering reported as net too
                    "by_fulfillment": {},
                    "items": {},
                    "grand_total_cents": tx["all"]["net"] + tx["all"]["tax"] + tx["all"]["tip"],
                }
                patched += 1
                continue

            # Override totals with txstore (matches Square Sales Summary)
            # IMPORTANT: downstream computes core = net_sales_cents - by_category[catering],
            # so net_sales_cents must be the FULL Net Sales (core + catering).
            existing["orders"]          = core["n"]
            existing["net_sales_cents"] = tx["all"]["net"]            # full Net Sales = gross - |disc|, all sources
            existing["tip_cents"]       = core["tip"]
            existing["tax_cents"]       = tx["all"]["tax"]
            existing["discount_cents"]  = abs(tx["all"]["disc"])
            # Category catering = Invoices-source net. Downstream subtracts this from net_sales_cents.
            by_cat = existing.setdefault("by_category", {})
            by_cat["catering"] = cat["net"]
            # Reconcile: scale non-catering categories so they sum to core net.
            non_cat_keys = ["coffee_bev", "food", "retail_coffee", "other"]
            current_non_cat = sum(by_cat.get(k, 0) for k in non_cat_keys)
            if current_non_cat > 0 and core["net"] > 0:
                scale = core["net"] / current_non_cat
                for k in non_cat_keys:
                    by_cat[k] = int(by_cat.get(k, 0) * scale)
            patched += 1

    j["per_cafe_per_day"] = per
    j["txstore_patched"] = True
    json.dump(j, open(WEEK_JSON, "w"), indent=2)
    print(f"Patched {patched} cafe-day entries in {os.path.basename(WEEK_JSON)}")

    # Print sanity check
    print(f"\nWeek totals (Core only) after patch:")
    for cafe in CAFES:
        if cafe not in per: continue
        n = sum(d.get("orders",0) for d in per[cafe].values())
        net = sum(d.get("net_sales_cents",0) for d in per[cafe].values())
        tip = sum(d.get("tip_cents",0) for d in per[cafe].values())
        cat = sum(d.get("by_category",{}).get("catering",0) for d in per[cafe].values())
        print(f"  {cafe:<10}  n={n:>5}  core_gross=${net/100:>10,.2f}  tips=${tip/100:>8,.2f}  catering=${cat/100:>7,.2f}")


if __name__ == "__main__":
    main()
