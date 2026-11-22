#!/usr/bin/env python3
"""Slipstream CLOUD weekly delta builder (stdlib only, no pandas, no mounted folder needed).

Reads the Square results this session already holds (spilled tool-result files +
the session transcript), and writes ONE small delta bundle for the week:

    delta_<START>_<END>.tar.xz.b64   (text; safe to store as a Project doc)
        transactions.csv        payment rows in Square-export column shape (same as transactions/YYYY-MM.csv)
        daily_metrics.csv       one row per (date, cafe), same columns/rounding as daily_metrics.csv
        loyalty_events.jsonl    loyalty events in the window, deduped on id
        manifest.json           counts per cafe for the two-line summary

The big stores (daily_metrics.csv, transactions/YYYY-MM.csv, loyalty_events/YYYY-MM.jsonl)
are NOT touched here. merge_deltas.py folds bundles into them later (idempotent).

Usage:  python3 slipstream_cloud_delta.py START END [--out DIR] [--root DIR ...]
  START/END inclusive, America/New_York (real zone: EDT UTC-4 in summer, EST UTC-5 in winter).
  Before 2026-11-08 this used a fixed UTC-4, so winter Time values were 1 hour late (dates were unaffected).
  --root: extra dirs to search (default: ~/.claude/projects). Finds tool-results/*.txt and *.jsonl transcripts.
"""
import os, sys, re, io, csv, json, glob, base64, tarfile, argparse, hashlib
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from collections import defaultdict, Counter

ET = ZoneInfo("America/New_York")  # was a fixed UTC-4 until 2026-11-08
LOC_TO_CAFE = {
    "LXM7V1QWDHKZJ": "Slipstream Downtown",
    "LGVKCKX9B9DZT": "Slipstream Studio",
    "L4ZBMGCWJ5C7R": "Slipstream 1899 Penn",
    "LV2D18H0DD9ES": "Slipstream NMWA",
}
LOCATION_MAP = {  # transactions_store.LOCATION_MAP
    "Slipstream Downtown": "Downtown", "Slipstream Studio": "Studio",
    "Slipstream 1899 Penn": "1899 Penn", "Slipstream NMWA": "NMWA", "Slipstream 1899": "1899 Penn",
}
LOYALTY_LOC = {"LXM7V1QWDHKZJ": "Downtown", "LGVKCKX9B9DZT": "Studio", "L4ZBMGCWJ5C7R": "1899", "LV2D18H0DD9ES": "NMWA"}

CSV_COLUMNS = [
    "Date","Time","Time Zone","Gross Sales","Discounts","Service Charges","Net Sales","Gift Card Sales","Tax","Tip",
    "Partial Refunds","Total Collected","Source","Card","Card Entry Methods","Cash","Square Gift Card","Other Tender",
    "Other Tender Type","Tender Note","Fees","Net Total","Transaction ID","Payment ID","Card Brand","PAN Suffix",
    "Device Name","Staff Name","Staff ID","Details","Description","Event Type","Location","Dining Option","Customer ID",
    "Customer Name","Customer Reference ID","Device Nickname","Third Party Fees","Deposit ID","Deposit Date",
    "Deposit Details","Fee Percentage Rate","Fee Fixed Rate","Refund Reason","Discount Name","Transaction Status",
    "Cash App","Order Reference ID","Fulfillment Note","Free Processing Applied","Channel","Unattributed Tips",
    "Table Info","International Fee",
]
MONEY_DEFAULT_COLS = {"Gross Sales","Discounts","Service Charges","Net Sales","Gift Card Sales","Tax","Tip","Partial Refunds",
    "Total Collected","Card","Cash","Square Gift Card","Other Tender","Fees","Net Total","Cash App",
    "Free Processing Applied","Third Party Fees","Unattributed Tips","International Fee"}
DM_COLS = ["date","cafe","transactions","net_sales","tips","coffee_bev","food","retail_coffee","other",
           "avg_ticket","walkin_transactions","walkin_net_sales"]


# ---------------------------------------------------------------- helpers (ported verbatim in behaviour)
def classify_item(name):  # aggregate_week.classify_item
    n = (name or "").strip(); low = n.lower()
    if "whole bean" in low or re.match(r"^\[(bright|comfort|fruit|decaf|featured)\]", low, re.I):
        if "whole bean" in low or "bag" in low or re.search(r"\b(8 ?oz|12 ?oz|1 ?lb|5 ?lb)\b", low):
            return "retail_coffee"
        return "coffee_bev"
    food_kw = ("sandwich","toast","bread","salad","breakfast","scone","muffin","bagel","cookie","yogurt","oatmeal","wrap",
               "quiche","granola","pastry","croissant","parfait","fruit","side ","bar ","biscotti","brownie","donut",
               "banana bread","ham","egg","avocado","cheese")
    if any(k in low for k in food_kw): return "food"
    bev_kw = ("latte","cappuccino","espresso","americano","macchiato","cortado","mocha","matcha","chai","tea","coffee",
              "nitro","cold brew","lemonade","drip","pour over","cruise control","vanilla","maple","caramel",
              "breakfast latte","desvelado","song red","red tea")
    if any(k in low for k in bev_kw): return "coffee_bev"
    return "other"

def cents_to_csv(c):
    if c == 0: return "$0.00"
    if c < 0: return f"-${abs(c)/100:,.2f}"
    return f"${c/100:,.2f}"

def parse_money(s):  # transactions_store._parse_money
    if s is None: return 0
    s = str(s).strip()
    if not s or s.lower() == "nan": return 0
    neg = False
    if s.startswith("(") and s.endswith(")"): neg = True; s = s[1:-1]
    if s.startswith("-"): neg = True; s = s[1:]
    s = s.replace("$", "").replace(",", "").strip()
    if not s: return 0
    try: c = int(round(float(s) * 100))
    except ValueError: return 0
    return -c if neg else c

def r2(x):
    """Match legacy numpy rounding (np.round(x,2) == rint(x*100)/100), not Python's exact-decimal round."""
    return round(x * 100) / 100.0

def local_dt(iso_utc):
    return datetime.fromisoformat(iso_utc.replace("Z", "+00:00")).astimezone(ET)

def local_date_time(iso_utc):
    d = local_dt(iso_utc); return d.strftime("%Y-%m-%d"), d.strftime("%H:%M:%S")

def sg(o, *path, default=0):
    cur = o
    for k in path:
        if cur is None: return default
        cur = cur.get(k) if isinstance(cur, dict) else None
    return cur if cur is not None else default

def allocate(order_amount, payment_amount, order_total_paid):
    if order_total_paid <= 0: return order_amount
    if payment_amount == order_total_paid: return order_amount
    return int(round(order_amount * (payment_amount / order_total_paid)))


# ---------------------------------------------------------------- discovery
def result_texts_from_transcript(path):
    try: fh = open(path, errors="replace")
    except OSError: return
    with fh:
        for line in fh:
            if '"tool_result"' not in line: continue
            try: obj = json.loads(line)
            except Exception: continue
            msg = obj.get("message", {})
            content = msg.get("content") if isinstance(msg, dict) else None
            for c in (content or []):
                if not isinstance(c, dict) or c.get("type") != "tool_result": continue
                cc = c.get("content")
                if isinstance(cc, list): yield "".join(t.get("text", "") for t in cc if isinstance(t, dict))
                elif isinstance(cc, str): yield cc

def all_result_texts(roots):
    for root in roots:
        for f in sorted(glob.glob(os.path.join(root, "**", "*.jsonl"), recursive=True)):
            yield from result_texts_from_transcript(f)
        # NOTE: *make_api_request* (no trailing '-') so BOTH make_api_request and make_api_request_read spills match.
        for f in sorted(glob.glob(os.path.join(root, "**", "tool-results", "mcp-*make_api_request*.txt"), recursive=True)):
            try: yield open(f, errors="replace").read()
            except OSError: pass

def harvest(roots):
    payments, orders, events = {}, {}, {}
    for raw in all_result_texts(roots):
        if not raw or not raw.lstrip().startswith("{"): continue
        head = raw[:200]
        if '"payments"' in head:
            kind = "payments"
        elif '"orders"' in head:
            kind = "orders"
        elif '"events"' in head:
            kind = "events"
        else: continue
        try: j = json.loads(raw)
        except Exception: continue
        if kind == "payments":
            for p in j.get("payments", []):
                if p.get("id"): payments[p["id"]] = p
        elif kind == "orders":
            for o in j.get("orders", []):
                if o.get("id"): orders[o["id"]] = o
        else:
            for e in j.get("events", []):
                if e.get("id"): events[e["id"]] = e
    return payments, orders, events


# ---------------------------------------------------------------- row builder (port of build_transactions_csv.build_row)
def build_row(payment, order, siblings):
    loc_name = LOC_TO_CAFE.get(payment.get("location_id", ""), "")
    date, t = local_date_time(payment["created_at"])
    lis = (order or {}).get("line_items", []) or []
    o_gross = sum(sg(li, "gross_sales_money", "amount") for li in lis)
    o_discount = sg(order, "net_amounts", "discount_money", "amount")
    o_tax = sg(order, "net_amounts", "tax_money", "amount")
    o_service = sg(order, "net_amounts", "service_charge_money", "amount")
    order_total_paid = sum(p.get("total_money", {}).get("amount", 0) for p in siblings)
    if order_total_paid == 0:
        order_total_paid = sg(order, "net_amounts", "total_money", "amount")
    p_tip = sg(payment, "tip_money", "amount")
    p_total = sg(payment, "total_money", "amount")
    p_fees = sum(sg(f, "amount_money", "amount") for f in (payment.get("processing_fee") or []))
    gross = allocate(o_gross, p_total, order_total_paid)
    discount = allocate(o_discount, p_total, order_total_paid)
    tax_alloc = allocate(o_tax, p_total, order_total_paid)
    service = allocate(o_service, p_total, order_total_paid)
    net_sales = gross - discount
    src_type = (payment.get("source_type") or "").upper()
    card_amt = p_total if src_type == "CARD" else 0
    cash_amt = p_total if src_type == "CASH" else 0
    sgc_amt = p_total if src_type == "SQUARE_GIFT_CARD" else 0
    other_amt = p_total if src_type not in ("CARD", "CASH", "SQUARE_GIFT_CARD") else 0
    src_name = sg(order, "source", "name", default="") or "Register"
    fulfs = (order or {}).get("fulfillments", []) or []
    dining = (fulfs[0].get("type", "") or "") if fulfs else ""
    def _nm(li):
        n = li.get("name", "") or ""; v = li.get("variation_name", "") or ""
        q = li.get("quantity", "1") or "1"
        try: q = int(float(q))
        except ValueError: q = 1
        n = f"{n} ({v})" if v else n
        return f"{q} x {n}" if q != 1 else n
    desc = ", ".join(_nm(li) for li in lis)  # all items (Nov 22 2026); was first item only
    status = {"COMPLETED": "Complete", "CANCELED": "Canceled", "FAILED": "Failed", "APPROVED": "Approved"}.get(
        payment.get("status", ""), payment.get("status", ""))
    card = (payment.get("card_details", {}) or {}).get("card", {}) or {}
    brand = card.get("card_brand", "") or ""
    entry = (payment.get("card_details", {}) or {}).get("entry_method", "") or ""
    return {
        "Date": date, "Time": t, "Time Zone": "Eastern Time (US & Canada)",
        "Gross Sales": cents_to_csv(gross),
        "Discounts": "$0.00" if discount == 0 else f"-${discount/100:,.2f}",
        "Service Charges": cents_to_csv(service), "Net Sales": cents_to_csv(net_sales),
        "Gift Card Sales": "$0.00", "Tax": cents_to_csv(tax_alloc), "Tip": cents_to_csv(p_tip),
        "Partial Refunds": "$0.00", "Total Collected": cents_to_csv(p_total), "Source": src_name,
        "Card": cents_to_csv(card_amt), "Card Entry Methods": entry, "Cash": cents_to_csv(cash_amt),
        "Square Gift Card": cents_to_csv(sgc_amt), "Other Tender": cents_to_csv(other_amt),
        "Other Tender Type": src_type if other_amt else "", "Tender Note": "",
        "Fees": "$0.00" if p_fees == 0 else f"-${p_fees/100:,.2f}",
        "Net Total": cents_to_csv(p_total - p_fees),
        "Transaction ID": payment.get("order_id", "") or "", "Payment ID": payment["id"],
        "Card Brand": {"VISA": "Visa", "MASTERCARD": "Mastercard", "AMERICAN_EXPRESS": "American Express",
                       "DISCOVER": "Discover", "JCB": "JCB", "CHINA_UNIONPAY": "UnionPay"}.get(brand, brand),
        "PAN Suffix": card.get("last_4", "") or "",
        "Device Name": sg(payment, "device_details", "device_name", default="") or "",
        "Staff Name": "", "Staff ID": payment.get("team_member_id", "") or "", "Details": "",
        "Description": desc, "Event Type": "Payment", "Location": loc_name, "Dining Option": dining,
        "Customer ID": payment.get("customer_id", "") or "", "Customer Name": "", "Customer Reference ID": "",
        "Device Nickname": "", "Third Party Fees": "$0.00", "Deposit ID": "", "Deposit Date": "", "Deposit Details": "",
        "Fee Percentage Rate": "", "Fee Fixed Rate": "", "Refund Reason": "", "Discount Name": "",
        "Transaction Status": status, "Cash App": "$0.00", "Order Reference ID": "", "Fulfillment Note": "",
        "Free Processing Applied": "$0.00", "Channel": loc_name, "Unattributed Tips": "$0.00", "Table Info": "",
        "International Fee": "$0.00",
    }


def build_transactions(payments, orders, start, end):
    by_order = defaultdict(list)
    for p in payments.values():
        if p.get("order_id"): by_order[p["order_id"]].append(p)
    rows, skipped = [], Counter()
    for p in payments.values():
        if (p.get("status") or "") != "COMPLETED": skipped["non_completed"] += 1; continue
        if p.get("location_id") not in LOC_TO_CAFE: skipped["no_location"] += 1; continue
        try: d, _ = local_date_time(p["created_at"])
        except Exception: continue
        if d < start or d > end: skipped["out_of_window"] += 1; continue
        oid = p.get("order_id"); order = orders.get(oid)
        if not order: skipped["no_order"] += 1
        rows.append(build_row(p, order, by_order.get(oid, [p]) if oid else [p]))
    rows.sort(key=lambda r: (r["Date"], r["Time"], r["Transaction ID"]))
    return rows, skipped


# ---------------------------------------------------------------- daily metrics (port of build_daily_metrics.build + store load rules)
def build_daily(rows):
    # store.load(): drop unknown cafe / empty transaction_id, dedupe on Transaction ID keeping LAST (file order = Date,Time,TxnID)
    kept = {}
    for r in rows:
        cafe = LOCATION_MAP.get(r["Location"].strip())
        tid = r["Transaction ID"].strip()
        if not cafe or not tid: continue
        kept[tid] = (cafe, r)
    groups = defaultdict(list)
    for tid, (cafe, r) in kept.items():
        groups[(r["Date"].strip(), cafe)].append(r)
    out = []
    for (d, cafe) in sorted(groups):
        g = groups[(d, cafe)]
        tx = len({x["Transaction ID"].strip() for x in g})
        net = sum(parse_money(x["Net Sales"]) for x in g) / 100.0
        tips = sum(parse_money(x["Tip"]) for x in g) / 100.0
        gw = [x for x in g if "invoice" not in x["Source"].strip().lower()]
        w_tx = len({x["Transaction ID"].strip() for x in gw})
        w_net = sum(parse_money(x["Net Sales"]) for x in gw) / 100.0
        cat = {"coffee_bev": 0.0, "food": 0.0, "retail_coffee": 0.0, "other": 0.0}
        for x in g:
            first = re.sub(r"^\d+ x ", "", x["Description"].split(", ")[0].strip())
            first = re.sub(r"\s*\([^()]*\)$", "", first)  # drop "(variation)"; category from the first item as before
            c = classify_item(first) or "other"
            if c not in cat: c = "other"
            cat[c] += parse_money(x["Net Sales"]) / 100.0
        out.append({"date": d, "cafe": cafe, "transactions": int(tx), "net_sales": r2(net), "tips": r2(tips),
                    "coffee_bev": r2(cat["coffee_bev"]), "food": r2(cat["food"]),
                    "retail_coffee": r2(cat["retail_coffee"]), "other": r2(cat["other"]),
                    "avg_ticket": r2(w_net / w_tx) if w_tx else 0,
                    "walkin_transactions": int(w_tx), "walkin_net_sales": r2(w_net)})
    return out


def payments_from_tenders(orders):
    """One payment per tender (skips NO_SALE drawer opens and failed/voided cards). tender.amount_money includes the tip."""
    out = {}
    for o in orders.values():
        if o.get("state") not in ("COMPLETED", "OPEN"): continue
        for t in o.get("tenders", []) or []:
            ty = (t.get("type") or "").upper()
            if ty == "NO_SALE": continue
            cd = t.get("card_details") or {}
            if cd and (cd.get("status") or "CAPTURED") not in ("CAPTURED", "AUTHORIZED"): continue
            pid = t.get("payment_id") or t.get("id")
            out[pid] = {"id": pid, "order_id": o["id"], "location_id": t.get("location_id") or o.get("location_id"),
                        "created_at": t.get("created_at") or o.get("created_at"), "status": "COMPLETED",
                        "total_money": t.get("amount_money") or {"amount": 0}, "tip_money": t.get("tip_money") or {"amount": 0},
                        "source_type": "CARD" if ty == "CARD" else ("CASH" if ty == "CASH" else ("SQUARE_GIFT_CARD" if ty == "SQUARE_GIFT_CARD" else ty)),
                        "card_details": cd, "customer_id": o.get("customer_id", "")}
    return out


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("start"); ap.add_argument("end")
    ap.add_argument("--out", default="."); ap.add_argument("--root", action="append", default=[])
    ap.add_argument("--from-tenders", action="store_true",
                    help="build payments from each order's tenders (no payments.list needed; pull OPEN orders with tenders too)")
    a = ap.parse_args()
    roots = a.root or [os.path.expanduser("~/.claude/projects")]
    os.makedirs(a.out, exist_ok=True)

    payments, orders, events_all = harvest(roots)
    if a.from_tenders:
        payments = payments_from_tenders(orders)
    tx_rows, skipped = build_transactions(payments, orders, a.start, a.end)
    daily = build_daily(tx_rows)
    # Paid orders that orders.search(COMPLETED) misses (Mobile / Square Online / kiosk orders stay OPEN): list them so the
    # caller can fetch them with orders.batchGet (<=100 ids per call) and re-run.
    missing = sorted({p['order_id'] for p in payments.values()
                      if p.get('status') == 'COMPLETED' and p.get('location_id') in LOC_TO_CAFE and p.get('order_id')
                      and p['order_id'] not in orders and a.start <= local_date_time(p['created_at'])[0] <= a.end})
    json.dump(missing, open(os.path.join(a.out, 'missing_order_ids.json'), 'w'))

    lo = datetime.fromisoformat(a.start + "T00:00:00").replace(tzinfo=ET)
    hi = (datetime.fromisoformat(a.end + "T00:00:00") + timedelta(days=1)).replace(tzinfo=ET)
    events = [e for e in events_all.values()
              if e.get("location_id") in LOYALTY_LOC and lo <= local_dt(e["created_at"]) < hi]
    events.sort(key=lambda e: e["created_at"])

    # per-cafe summary
    by_cafe_tx = Counter();
    for d in daily: by_cafe_tx[d["cafe"]] += d["transactions"]
    acc = Counter(); tot = Counter()
    for e in events:
        c = LOYALTY_LOC[e["location_id"]]; tot[c] += 1
        if e.get("type") == "ACCUMULATE_POINTS": acc[c] += 1
    manifest = {
        "start": a.start, "end": a.end, "built_at": datetime.now(timezone.utc).isoformat(),
        "payments_seen": len(payments), "orders_seen": len(orders), "txn_rows": len(tx_rows),
        "skipped": dict(skipped), "missing_orders": len(missing), "daily_rows": len(daily),
        "transactions_by_cafe": dict(by_cafe_tx),
        "loyalty_events": len(events), "loyalty_accumulate_by_cafe": dict(acc), "loyalty_events_by_cafe": dict(tot),
    }
    # warnings: a zero at Downtown / Studio / 1899 is a pull failure (NMWA loyalty has been zero since it opened)
    warn = []
    for cafe in ("Downtown", "Studio", "1899 Penn", "NMWA"):
        if by_cafe_tx[cafe] == 0: warn.append(f"ZERO transactions at {cafe}")
    for c in ("Downtown", "Studio", "1899"):
        if acc[c] == 0: warn.append(f"ZERO loyalty visits at {c} (pull failure?)")
    if missing: warn.append(f"{len(missing)} paid orders missing (fetch with orders.batchGet and re-run)")
    manifest["timezone"] = "America/New_York"
    manifest["days_by_cafe"] = dict(Counter(d["cafe"] for d in daily))
    manifest["warnings"] = warn

    def csv_bytes(cols, rows):
        b = io.StringIO(); w = csv.DictWriter(b, fieldnames=cols, extrasaction="ignore"); w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "$0.00" if c in MONEY_DEFAULT_COLS else "") for c in cols})
        return b.getvalue().encode()

    files = {
        "transactions.csv": csv_bytes(CSV_COLUMNS, tx_rows),
        "daily_metrics.csv": csv_bytes(DM_COLS, daily),
        "loyalty_events.jsonl": "".join(json.dumps(e) + "\n" for e in events).encode(),
    }
    manifest["sha256"] = {k: hashlib.sha256(v).hexdigest() for k, v in files.items()}
    files["manifest.json"] = json.dumps(manifest, indent=1).encode()

    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w:xz") as tar:
        for name, data in files.items():
            ti = tarfile.TarInfo(name); ti.size = len(data); ti.mtime = 0; tar.addfile(ti, io.BytesIO(data))
    b64 = base64.b64encode(buf.getvalue()).decode()
    bundle = os.path.join(a.out, f"delta_{a.start}_{a.end}.tar.xz.b64")
    open(bundle, "w").write(b64)
    open(os.path.join(a.out, f"manifest_{a.start}_{a.end}.json"), "w").write(json.dumps(manifest, indent=1))
    print(json.dumps(manifest, indent=1)); print("bundle:", bundle, f"({len(b64):,} chars)")

if __name__ == "__main__":
    main()
