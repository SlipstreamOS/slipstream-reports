#!/usr/bin/env python3
"""Assemble a Square Transactions CSV from cached API responses.

This is the post-processing half of the "pull all transactions" pipeline.
The other half is the agent loop that calls Square Payments+Orders MCPs
and lets the responses spill to disk.

Inputs (auto-discovered from tool-results dirs):
  • mcp-*-make_api_request-*.txt files containing either
    `{"payments": [...]}` or `{"orders": [...]}` responses.

For each completed Payment in the requested date range:
  - Look up its Order via order_id (from a payments-only response we'd lack
    line items, sources, dining options, and discount/tax breakdowns)
  - Allocate the order's gross/discount/tax/service-charge proportionally
    across its payments (split-tender orders typically have N payments
    summing to the order's total). For most coffee orders this is 1-1.
  - Output one row per Payment, in the same column shape as Square's
    monthly Transactions CSV export.

Output: one file per month-of-data covered by the date range, e.g.
  `_pipeline/shared/transactions/2026-04.csv`
This appends/overwrites only the rows for the selected date window so
weekly re-runs are safe.

Run:
  python3 build_transactions_csv.py 2026-04-01 2026-04-25
"""
import os, sys, json, csv, glob, math
from datetime import datetime, timezone, timedelta
from collections import defaultdict

EDT = timezone(timedelta(hours=-4))

# Resolve directories from this file's location (works across chats).
HERE = os.path.dirname(os.path.abspath(__file__))                 # shared/scripts
SHARED_DIR = os.path.dirname(HERE)                                 # shared
PIPELINE_DIR = os.path.dirname(SHARED_DIR)                         # _pipeline
WORKSPACE = os.path.dirname(PIPELINE_DIR)                          # workspace root
TXN_DIR = os.path.join(SHARED_DIR, "transactions")
SESSION_ROOT = WORKSPACE
# Walk up to find session root containing mnt/.claude
p = HERE
while p != "/":
    if os.path.isdir(os.path.join(p, "mnt", ".claude")):
        SESSION_ROOT = p
        break
    p = os.path.dirname(p)

LOC_TO_CAFE = {
    "LXM7V1QWDHKZJ": "Slipstream Downtown",
    "LGVKCKX9B9DZT": "Slipstream Studio",
    "L4ZBMGCWJ5C7R": "Slipstream 1899 Penn",
    "LV2D18H0DD9ES": "Slipstream NMWA",
}

# Output column order — matches Square's Transactions CSV export so
# transactions_store.py reads it transparently alongside monthly exports.
CSV_COLUMNS = [
    "Date","Time","Time Zone",
    "Gross Sales","Discounts","Service Charges","Net Sales",
    "Gift Card Sales","Tax","Tip","Partial Refunds","Total Collected",
    "Source",
    "Card","Card Entry Methods","Cash","Square Gift Card",
    "Other Tender","Other Tender Type","Tender Note",
    "Fees","Net Total",
    "Transaction ID","Payment ID",
    "Card Brand","PAN Suffix",
    "Device Name","Staff Name","Staff ID",
    "Details","Description","Event Type",
    "Location","Dining Option",
    "Customer ID","Customer Name","Customer Reference ID","Device Nickname",
    "Third Party Fees","Deposit ID","Deposit Date","Deposit Details",
    "Fee Percentage Rate","Fee Fixed Rate",
    "Refund Reason","Discount Name",
    "Transaction Status","Cash App","Order Reference ID","Fulfillment Note",
    "Free Processing Applied","Channel","Unattributed Tips",
    "Table Info","International Fee",
]

# ---------------------------------------------------------------------------
# Money helpers — Square money fields are integer cents in API; CSV uses $X.XX

def cents_to_csv(c, *, neg_in_parens=True):
    if c == 0: return "$0.00"
    if c < 0:
        if neg_in_parens:
            return f"-${abs(c)/100:,.2f}"
        return f"-${abs(c)/100:,.2f}"
    return f"${c/100:,.2f}"


# ---------------------------------------------------------------------------
# Load all payments + orders from tool-result files

def _tool_result_dirs():
    pat = os.path.join(SESSION_ROOT, "mnt", ".claude", "projects", "*", "*", "tool-results")
    return sorted(glob.glob(pat))

def _ingest_api_text(raw, payments, orders):
    """Merge one `{"payments":[...]}` / `{"orders":[...]}` response body."""
    if not raw or not raw.lstrip().startswith("{"):
        return
    head = raw[:200]
    if '"payments"' in head:
        try: j = json.loads(raw)
        except Exception: return
        for p in j.get("payments", []):
            pid = p.get("id")
            if pid: payments[pid] = p
    elif '"orders"' in head:
        try: j = json.loads(raw)
        except Exception: return
        for o in j.get("orders", []):
            oid = o.get("id")
            if oid: orders[oid] = o

def _transcript_files():
    """Session transcript JSONLs (newest first). Small pages that fit in the
    reply are NOT spilled to tool-results/ — they only exist here."""
    pat = os.path.join(SESSION_ROOT, "mnt", ".claude", "projects", "*", "*.jsonl")
    fs = glob.glob(pat)
    fs.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return fs

def _ingest_transcript(path, payments, orders):
    """Harvest inline (non-spilled) make_api_request results from a transcript.
    Added Sep 2026: the last page of a paginated pull is often small enough
    to come back inline, and was silently dropped before this."""
    try:
        fh = open(path, errors="replace")
    except OSError:
        return
    with fh:
        for line in fh:
            # cheap prefilter — quotes are escaped inside the JSONL line
            if 'payments' not in line and 'orders' not in line:
                continue
            try:
                obj = json.loads(line)
            except Exception:
                continue
            for c in (obj.get("message", {}).get("content") or []):
                if not isinstance(c, dict) or c.get("type") != "tool_result":
                    continue
                content = c.get("content")
                if isinstance(content, list):
                    text = "".join(t.get("text", "") for t in content if isinstance(t, dict))
                elif isinstance(content, str):
                    text = content
                else:
                    continue
                _ingest_api_text(text, payments, orders)

def load_api_data():
    """Return (payments_by_id, orders_by_id) — both dicts keyed by id."""
    payments = {}
    orders = {}
    # 1) inline results from the session transcript(s)
    for t in _transcript_files():
        _ingest_transcript(t, payments, orders)
    # 2) spilled results (large pages) — these take precedence
    for d in _tool_result_dirs():
        for f in sorted(glob.glob(os.path.join(d, "mcp-*make_api_request*.txt"))):
            try:
                raw = open(f).read()
            except Exception:
                continue
            _ingest_api_text(raw, payments, orders)
    return payments, orders


# ---------------------------------------------------------------------------
# Convert a Payment+Order pair to a CSV row

def _local_date_time(iso_utc):
    dt = datetime.fromisoformat(iso_utc.replace("Z","+00:00"))
    local = dt.astimezone(EDT)
    return local.strftime("%Y-%m-%d"), local.strftime("%H:%M:%S")

def _safe_get(o, *path, default=0):
    cur = o
    for k in path:
        if cur is None: return default
        cur = cur.get(k) if isinstance(cur, dict) else None
    return cur if cur is not None else default

def _allocate(order_amount, payment_amount, order_total_paid):
    """Allocate an order-level amount to a single payment proportionally.
    For single-payment orders this returns the full amount.
    """
    if order_total_paid <= 0:
        return order_amount
    if payment_amount == order_total_paid:
        return order_amount
    return int(round(order_amount * (payment_amount / order_total_paid)))

def payments_for_order(order_id, payments_by_id):
    """Find all payments belonging to an order (by order_id)."""
    return [p for p in payments_by_id.values() if p.get("order_id") == order_id]

def build_row(payment, order, all_payments_for_this_order):
    """Build one CSV row for a single payment, using its parent order for
    item-level / discount / tax detail."""
    pid = payment["id"]
    loc_id = payment.get("location_id","")
    loc_name = LOC_TO_CAFE.get(loc_id, "")

    # Time/date in EDT
    date, t = _local_date_time(payment["created_at"])

    # Order-level money totals (gross BEFORE discounts; after discounts = net)
    o_gross_pre_disc  = sum(_safe_get(li,"gross_sales_money","amount",default=0) for li in (order or {}).get("line_items",[]))
    o_discount        = _safe_get(order, "net_amounts","discount_money","amount",default=0)
    o_tax             = _safe_get(order, "net_amounts","tax_money","amount",default=0)
    o_tip             = _safe_get(order, "net_amounts","tip_money","amount",default=0)
    o_service         = _safe_get(order, "net_amounts","service_charge_money","amount",default=0)
    o_returned_total  = _safe_get(order, "returned_amounts","total_money","amount",default=0)
    # Order's total paid across all its payments (for allocation)
    order_total_paid  = sum(p.get("total_money",{}).get("amount",0) for p in all_payments_for_this_order)
    if order_total_paid == 0:
        order_total_paid = _safe_get(order, "net_amounts","total_money","amount",default=0)

    # This payment's amount/tip/total
    p_amount = _safe_get(payment, "amount_money","amount", default=0)
    p_tip    = _safe_get(payment, "tip_money","amount", default=0)
    p_total  = _safe_get(payment, "total_money","amount", default=0)
    p_fees   = sum(_safe_get(f,"amount_money","amount",default=0) for f in payment.get("processing_fee",[]) or [])

    # Allocate order-level amounts to this payment
    gross    = _allocate(o_gross_pre_disc, p_total, order_total_paid)
    discount = _allocate(o_discount,       p_total, order_total_paid)
    tax_alloc= _allocate(o_tax,            p_total, order_total_paid)
    service  = _allocate(o_service,        p_total, order_total_paid)
    net_sales= gross - discount  # Square: Net Sales = Gross - Discounts

    # Tender split (CARD/CASH/SQUARE_GIFT_CARD/EXTERNAL etc.)
    src_type = (payment.get("source_type") or "").upper()
    card_amt = p_total if src_type == "CARD" else 0
    cash_amt = p_total if src_type == "CASH" else 0
    sgc_amt  = p_total if src_type == "SQUARE_GIFT_CARD" else 0
    other_amt = p_total if src_type not in ("CARD","CASH","SQUARE_GIFT_CARD") else 0
    other_type = src_type if other_amt else ""

    # Source.name on order — Register / Mobile Order / Square Online / SK / Invoices / etc.
    # (Falls back to Square's default of "Register" for register POS).
    src_name = _safe_get(order, "source","name", default="") or "Register"

    # Dining option (often blank for register, "To Go"/"For Here" for restaurants)
    dining = ""
    fulfs = (order or {}).get("fulfillments", []) or []
    if fulfs:
        dining = fulfs[0].get("type","") or ""

    # Description = primary line-item name (Square usually echoes the first item)
    desc = ""
    lis = (order or {}).get("line_items", []) or []
    if lis:
        desc = lis[0].get("name","") or ""

    # Event type — "Payment" for register sales, "Refund" if this came from a refund
    event_type = "Payment"
    # Status mapping
    status_map = {"COMPLETED":"Complete","CANCELED":"Canceled","FAILED":"Failed","APPROVED":"Approved"}
    status = status_map.get(payment.get("status",""), payment.get("status",""))

    # Card details
    card = payment.get("card_details", {}).get("card", {}) or {}
    card_brand = card.get("card_brand","") or ""
    card_last4 = card.get("last_4","") or ""
    entry = (payment.get("card_details", {}) or {}).get("entry_method","") or ""

    row = {
        "Date": date, "Time": t, "Time Zone": "Eastern Time (US & Canada)",
        "Gross Sales": cents_to_csv(gross),
        "Discounts": "$0.00" if discount == 0 else f"-${discount/100:,.2f}",
        "Service Charges": cents_to_csv(service),
        "Net Sales": cents_to_csv(net_sales),
        "Gift Card Sales": "$0.00",  # gift card sales = sales of gift cards as items, derived elsewhere if needed
        "Tax": cents_to_csv(tax_alloc),
        "Tip": cents_to_csv(p_tip),
        "Partial Refunds": "$0.00",
        "Total Collected": cents_to_csv(p_total),
        "Source": src_name,
        "Card": cents_to_csv(card_amt),
        "Card Entry Methods": entry,
        "Cash": cents_to_csv(cash_amt),
        "Square Gift Card": cents_to_csv(sgc_amt),
        "Other Tender": cents_to_csv(other_amt),
        "Other Tender Type": other_type if other_amt else "",
        "Tender Note": "",
        "Fees": "$0.00" if p_fees == 0 else f"-${p_fees/100:,.2f}",
        "Net Total": cents_to_csv(p_total - p_fees),
        "Transaction ID": payment.get("order_id","") or "",
        "Payment ID": pid,
        "Card Brand": {"VISA":"Visa","MASTERCARD":"Mastercard","AMERICAN_EXPRESS":"American Express",
                       "DISCOVER":"Discover","JCB":"JCB","CHINA_UNIONPAY":"UnionPay"}.get(card_brand, card_brand),
        "PAN Suffix": card_last4,
        "Device Name": _safe_get(payment, "device_details","device_name", default="") or "",
        "Staff Name": "",  # Would require team_member name lookup
        "Staff ID": payment.get("team_member_id","") or "",
        "Details": "",
        "Description": desc,
        "Event Type": event_type,
        "Location": loc_name,
        "Dining Option": dining,
        "Customer ID": payment.get("customer_id","") or "",
        "Customer Name": "",
        "Customer Reference ID": "",
        "Device Nickname": "",
        "Third Party Fees": "$0.00",
        "Deposit ID": "",
        "Deposit Date": "",
        "Deposit Details": "",
        "Fee Percentage Rate": "",
        "Fee Fixed Rate": "",
        "Refund Reason": "",
        "Discount Name": "",
        "Transaction Status": status,
        "Cash App": "$0.00",
        "Order Reference ID": "",
        "Fulfillment Note": "",
        "Free Processing Applied": "$0.00",
        "Channel": loc_name,
        "Unattributed Tips": "$0.00",
        "Table Info": "",
        "International Fee": "$0.00",
    }
    return row


def build_csv(start_date, end_date, *, out_path=None, verbose=True):
    """Build a Transactions CSV covering [start_date, end_date] inclusive.
    Reads ALL cached API responses, filters by created_at in EDT.
    """
    if out_path is None:
        # Default: <month>.csv based on start_date's month
        ym = start_date[:7]
        out_path = os.path.join(TXN_DIR, f"{ym}.csv")

    payments_by_id, orders_by_id = load_api_data()
    if verbose:
        print(f"Loaded {len(payments_by_id):,} payments + {len(orders_by_id):,} orders from tool-results.")

    # Group payments by order_id for split-tender allocation
    by_order = defaultdict(list)
    for p in payments_by_id.values():
        oid = p.get("order_id")
        if oid: by_order[oid].append(p)

    rows = []
    skipped = {"out_of_window": 0, "non_completed": 0, "no_order": 0, "no_location": 0}
    for pid, p in payments_by_id.items():
        if (p.get("status") or "") != "COMPLETED":
            skipped["non_completed"] += 1; continue
        if p.get("location_id") not in LOC_TO_CAFE:
            skipped["no_location"] += 1; continue
        # Date filter (EDT local)
        try:
            local_date, _ = _local_date_time(p["created_at"])
        except Exception:
            continue
        if local_date < start_date or local_date > end_date:
            skipped["out_of_window"] += 1; continue
        oid = p.get("order_id")
        order = orders_by_id.get(oid)
        if not order:
            # We can still build a row from payment alone; tax/discount/category will be missing.
            skipped["no_order"] += 1
        siblings = by_order.get(oid, [p]) if oid else [p]
        rows.append(build_row(p, order, siblings))

    # Sort by date+time, then by transaction_id (stable)
    rows.sort(key=lambda r: (r["Date"], r["Time"], r["Transaction ID"]))

    if verbose:
        print(f"Building CSV: {len(rows):,} rows in window {start_date}..{end_date}")
        print(f"Skipped: {skipped}")

    # Merge with existing file: keep rows OUTSIDE our window, replace IN-window
    existing_rows = []
    if os.path.exists(out_path):
        with open(out_path, newline="") as f:
            for r in csv.DictReader(f):
                d = r.get("Date","")
                if d < start_date or d > end_date:
                    existing_rows.append(r)
    if verbose and existing_rows:
        print(f"Preserving {len(existing_rows):,} existing rows outside window.")

    all_rows = existing_rows + rows
    all_rows.sort(key=lambda r: (r.get("Date",""), r.get("Time",""), r.get("Transaction ID","")))

    os.makedirs(TXN_DIR, exist_ok=True)
    with open(out_path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        w.writeheader()
        for r in all_rows:
            # Fill missing columns with default-empty
            for c in CSV_COLUMNS:
                if c not in r:
                    r[c] = "$0.00" if c in ("Gross Sales","Discounts","Service Charges","Net Sales","Gift Card Sales","Tax","Tip","Partial Refunds","Total Collected","Card","Cash","Square Gift Card","Other Tender","Fees","Net Total","Cash App","Free Processing Applied","Third Party Fees","Unattributed Tips","International Fee") else ""
            w.writerow(r)
    print(f"Saved → {out_path}  ({len(all_rows):,} total rows)")


if __name__ == "__main__":
    if len(sys.argv) >= 3:
        start, end = sys.argv[1], sys.argv[2]
    else:
        # Default: current month, through today
        today = datetime.now(EDT).date()
        start = today.replace(day=1).isoformat()
        end = today.isoformat()
    build_csv(start, end)
