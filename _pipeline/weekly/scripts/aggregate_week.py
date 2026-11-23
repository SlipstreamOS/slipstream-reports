#!/usr/bin/env python3
"""Aggregate a Sun–Sat week of Square order pulls into a clean weekly data file.

Reads all MCP tool-result files from this chat's results folder, filters to the
configured Sun–Sat week, classifies orders by local date, and produces per-cafe
per-day aggregates for the weekly report.

To run a different week: edit the WEEK_START constant (Sunday midnight EDT) and
the default min-mtime filter if you want to scope to files from a specific day.
"""
import json, os, glob, re
from datetime import datetime, timezone, timedelta
from collections import defaultdict

from _paths import DATA_DIR, find_tool_results_folders  # shared path helper

# Which week to process (Sun–Sat, EDT). Change these to re-run for another week.
EDT = timezone(timedelta(hours=-4))
WEEK_START = datetime(2026, 6, 28, 0, 0, 0, tzinfo=EDT)
WEEK_END   = datetime(2026, 7, 5, 0, 0, 0, tzinfo=EDT)
WEEK_LABEL = f"{WEEK_START.date().strftime('%b%d').lower()}-{(WEEK_END - timedelta(days=1)).date().strftime('%d')}"  # e.g. "apr12-18"
OUT = os.path.join(DATA_DIR, f"week_{WEEK_LABEL}.json")

# Square location_id -> cafe display name
LOC_TO_CAFE = {
    "LXM7V1QWDHKZJ": "Downtown",
    "LGVKCKX9B9DZT": "Studio",
    "L4ZBMGCWJ5C7R": "1899 Penn",
    "LV2D18H0DD9ES": "NMWA",
}

# Catering / invoice-payment items to exclude from core metrics
CATERING_ITEMS = {
    "Open Event","Management Fee","Loss Pass-through",
    "Continental Breakfast Service (per person)",
    "Elegant Cold Breakfast service (per person)",
    "Dinner Salad Station (per person)",
    "Equipment Fee","Iced Coffee Service Monthly Maintenance",
    "Keg of Iced Coffee",
}

def classify_item(name):
    """Rough classification into sales categories — matches monthly report convention.
    Returns one of: coffee_bev, retail_coffee, food, other.

    Note (Apr 28 2026): catering classification is now SOURCE-based (orders with
    source.name == "Invoices" are tagged catering at the order level; see
    `is_invoice_order`). The old name-matching CATERING_ITEMS list is retained
    only for backward reference — it is no longer consulted here. Catering
    rung at the cafe register goes into Coffee/Food/Other like any other sale.
    """
    n = (name or "").strip()
    low = n.lower()
    if "whole bean" in low or re.match(r"^\[(bright|comfort|fruit|decaf|featured)\]", low, re.I):
        if "whole bean" in low or "bag" in low or re.search(r"\b(8 ?oz|12 ?oz|1 ?lb|5 ?lb)\b", low):
            return "retail_coffee"
        return "coffee_bev"
    food_kw = ("sandwich","toast","bread","salad","breakfast","scone","muffin","bagel",
               "cookie","yogurt","oatmeal","wrap","quiche","granola","pastry","croissant",
               "parfait","fruit","side ","bar ","biscotti","brownie","donut","banana bread",
               "ham","egg","avocado","cheese")
    if any(k in low for k in food_kw):
        return "food"
    bev_kw = ("latte","cappuccino","espresso","americano","macchiato","cortado","mocha",
              "matcha","chai","tea","coffee","nitro","cold brew","lemonade","drip",
              "pour over","cruise control","vanilla","maple","caramel","breakfast latte",
              "desvelado","song red","red tea")
    if any(k in low for k in bev_kw):
        return "coffee_bev"
    return "other"

def to_local_date(iso_utc):
    dt = datetime.fromisoformat(iso_utc.replace("Z","+00:00"))
    return dt.astimezone(EDT).date().isoformat()

def in_week(iso_utc):
    dt = datetime.fromisoformat(iso_utc.replace("Z","+00:00"))
    return WEEK_START <= dt.astimezone(EDT) < WEEK_END

def fresh_bucket():
    return {
        "orders": 0,
        "net_sales_cents": 0,
        "discount_cents": 0,
        "tax_cents": 0,
        "tip_cents": 0,
        "grand_total_cents": 0,
        "by_category": defaultdict(int),
        "by_fulfillment": defaultdict(int),
        "items": defaultdict(lambda: {"units":0,"gross_cents":0}),
    }

INVOICE_SOURCE_NAMES = {"Invoices", "Invoice"}  # Square Invoices ⇒ true catering/off-site

def is_invoice_order(order):
    """An order is 'catering' (excluded from Core Sales) iff it was placed via
    Square Invoices — i.e. NOT rung up at the register. Anything else (POS,
    Restaurants, gift card, no-source) counts as register sales and lives in
    Core, regardless of what items are on the ticket.
    """
    src = (order.get("source") or {}).get("name", "")
    return src in INVOICE_SOURCE_NAMES


def process_order(order, agg):
    loc = order.get("location_id")
    cafe = LOC_TO_CAFE.get(loc)
    if not cafe:
        return
    closed = order.get("closed_at") or order.get("created_at")
    if not closed or not in_week(closed):
        return
    date = to_local_date(closed)
    bucket = agg[cafe][date]
    bucket["orders"] += 1
    na = order.get("net_amounts", {})
    bucket["grand_total_cents"] += na.get("total_money", {}).get("amount", 0)
    bucket["tax_cents"]          += na.get("tax_money", {}).get("amount", 0)
    bucket["tip_cents"]          += na.get("tip_money", {}).get("amount", 0)
    bucket["discount_cents"]     += na.get("discount_money", {}).get("amount", 0)
    invoice = is_invoice_order(order)
    for li in order.get("line_items", []):
        gross = li.get("gross_sales_money", {}).get("amount", 0)
        bucket["net_sales_cents"] += gross
        if invoice:
            cat = "catering"
        else:
            cat = classify_item(li.get("name",""))
        bucket["by_category"][cat] += gross
        name = li.get("name","Unknown")
        qty = int(li.get("quantity","1"))
        bucket["items"][name]["units"] += qty
        bucket["items"][name]["gross_cents"] += gross
    for f in order.get("fulfillments", []):
        t = f.get("type", "UNKNOWN")
        bucket["by_fulfillment"][t] += 1

def main():
    agg = defaultdict(lambda: defaultdict(fresh_bucket))
    # Gather tool-result files from all visible projects folders (current chat).
    folders = find_tool_results_folders()
    files = []
    for folder in folders:
        files.extend(glob.glob(os.path.join(folder, "mcp-*make_api_request*.txt")))
    files = sorted(set(files), key=os.path.getmtime)
    if not files:
        print("WARNING: no tool-result files found. Are you running after the Square pulls completed?")
        print(f"Searched folders: {folders}")
        return
    total_orders_in_week = 0
    total_orders_seen = 0
    seen_ids = set()
    for f in files:
        try:
            data = json.loads(open(f).read())
        except Exception as e:
            print(f"skip {f}: {e}")
            continue
        orders = data.get("orders", [])
        total_orders_seen += len(orders)
        for o in orders:
            oid = o.get("id")
            if oid and oid in seen_ids:
                continue
            if oid:
                seen_ids.add(oid)
            process_order(o, agg)
    out = {}
    for cafe, days in agg.items():
        out[cafe] = {}
        for date, b in days.items():
            total_orders_in_week += b["orders"]
            b["by_category"] = dict(b["by_category"])
            b["by_fulfillment"] = dict(b["by_fulfillment"])
            items_sorted = sorted(b["items"].items(), key=lambda kv: -kv[1]["gross_cents"])
            b["items"] = {k: v for k, v in items_sorted[:20]}
            b["avg_ticket_cents"] = b["net_sales_cents"] / b["orders"] if b["orders"] else 0
            b["tip_pct_of_net"] = (b["tip_cents"] / b["net_sales_cents"] * 100) if b["net_sales_cents"] else 0
            out[cafe][date] = b
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(OUT, "w") as fh:
        json.dump({
            "week_start": WEEK_START.date().isoformat(),
            "week_end":   (WEEK_END - timedelta(days=1)).date().isoformat(),
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "per_cafe_per_day": out,
            "summary": {
                "orders_parsed_total": total_orders_seen,
                "orders_in_week_total": total_orders_in_week,
                "files_processed": len(files),
            }
        }, fh, indent=2)
    print(f"Saved → {OUT}")
    print(f"Orders parsed across files: {total_orders_seen}")
    print(f"Orders in window: {total_orders_in_week}")
    for cafe in ["Downtown","Studio","1899 Penn","NMWA"]:
        print(f"\n{cafe}:")
        for d in sorted(out.get(cafe,{}).keys()):
            b = out[cafe][d]
            print(f"  {d}: {b['orders']:4d} orders  net=${b['net_sales_cents']/100:>8,.0f}  "
                  f"tips=${b['tip_cents']/100:>6,.0f}  avg=${b['avg_ticket_cents']/100:>5.2f}")

if __name__ == "__main__":
    main()
