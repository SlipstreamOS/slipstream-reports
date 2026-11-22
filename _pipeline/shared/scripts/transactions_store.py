"""Slipstream transactions store — canonical payment-level history.

Backed by Square "Transactions" CSV exports living in
`_pipeline/shared/transactions/`. One CSV per month (e.g. 2026-01.csv).
Files are auto-discovered.

Why CSV, not sqlite? The workspace folder is a FUSE mount that doesn't
support fcntl locking (sqlite fails with "disk I/O error"). CSV is portable
across chats, human-readable, and fast enough at current volume (~25K
rows/month per cafe, trivial for pandas).

Public API (all queries take iso date strings 'YYYY-MM-DD'):
    load()                                — pandas DataFrame of all transactions
    counts_and_totals(cafe, start, end)   — (n, net_cents, tips_cents, gross_cents)
    avg_ticket_cents(cafe, start, end)    — None if no txns
    avg_daily_transactions(cafe, start, end) — per-open-day average
    dining_option_mix(cafe, start, end)   — {option: count}
    available_date_range()                — (min_date, max_date, total_rows)
    per_cafe_date_coverage()              — list of (cafe, min, max, n)

Location normalization: "Slipstream Downtown" → "Downtown", etc. See
LOCATION_MAP.

All money in the returned DataFrame is integer CENTS (columns suffixed
`_cents`).
"""
import os
import glob
from functools import lru_cache

import pandas as pd

# Self-locate the transactions folder so this module works whether it's imported
# from shared/scripts/ (where _paths.TXN_DIR exists) or from any other pipeline
# (weekly/scripts/_paths.py doesn't define TXN_DIR). Keeps us free of module-name
# collisions on `_paths`.
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))        # shared/scripts
_SHARED_DIR = os.path.dirname(_THIS_DIR)                       # shared
TXN_DIR = os.path.join(_SHARED_DIR, "transactions")

LOCATION_MAP = {
    "Slipstream Downtown":  "Downtown",
    "Slipstream Studio":    "Studio",
    "Slipstream 1899 Penn": "1899 Penn",
    "Slipstream NMWA":      "NMWA",
    "Slipstream 1899":      "1899 Penn",
}

CAFES = ["Downtown", "Studio", "1899 Penn", "NMWA"]
OPEN_CAFES = ["Downtown", "Studio", "1899 Penn"]

MONEY_COLS = [
    "Gross Sales", "Discounts", "Service Charges", "Net Sales",
    "Gift Card Sales", "Tax", "Tip", "Partial Refunds",
    "Total Collected", "Fees", "Net Total",
    "Card", "Cash", "Square Gift Card", "Other Tender",
]


def _parse_money(s):
    """'$5.50' / '-$13.00' / '($2.10)' / '' → integer cents."""
    if s is None:
        return 0
    s = str(s).strip()
    if not s or s.lower() == "nan":
        return 0
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True; s = s[1:-1]
    if s.startswith("-"):
        negative = True; s = s[1:]
    s = s.replace("$", "").replace(",", "").strip()
    if not s:
        return 0
    try:
        cents = int(round(float(s) * 100))
    except ValueError:
        return 0
    return -cents if negative else cents


def _csv_paths():
    """Sorted list of discovered transaction CSV paths."""
    return sorted(glob.glob(os.path.join(TXN_DIR, "*.csv")))


@lru_cache(maxsize=1)
def load():
    """Load all transaction CSVs, normalize, return a single pandas DataFrame.

    Columns of interest (money as int CENTS):
      transaction_id, payment_id, date, time, cafe, location_raw, source,
      event_type, transaction_status, dining_option,
      gross_sales_cents, discounts_cents, net_sales_cents, tip_cents,
      tax_cents, partial_refunds_cents, total_collected_cents, tender_*_cents,
      card_brand, customer_id, staff_name, description, source_csv
    """
    paths = _csv_paths()
    if not paths:
        return pd.DataFrame()

    frames = []
    for p in paths:
        df = pd.read_csv(p, dtype=str, keep_default_na=False, low_memory=False)
        df["source_csv"] = os.path.basename(p)
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True)

    # Trim the columns we use, then normalize.
    out = pd.DataFrame({
        "transaction_id":      raw["Transaction ID"].str.strip(),
        "payment_id":          raw.get("Payment ID", "").str.strip(),
        "date":                raw["Date"].str.strip(),
        "time":                raw.get("Time", "").str.strip(),
        "timezone":            raw.get("Time Zone", "").str.strip(),
        "location_raw":        raw["Location"].str.strip(),
        "source":              raw.get("Source", "").str.strip(),
        "event_type":          raw.get("Event Type", "").str.strip(),
        "transaction_status":  raw.get("Transaction Status", "").str.strip(),
        "dining_option":       raw.get("Dining Option", "").str.strip(),
        "customer_id":         raw.get("Customer ID", "").str.strip(),
        "staff_name":          raw.get("Staff Name", "").str.strip(),
        "card_brand":          raw.get("Card Brand", "").str.strip(),
        "description":         raw.get("Description", "").str.strip(),
        "source_csv":          raw["source_csv"],
    })
    out["cafe"] = out["location_raw"].map(LOCATION_MAP)

    # Money columns → integer cents
    col_rename = {
        "Gross Sales":       "gross_sales_cents",
        "Discounts":         "discounts_cents",
        "Service Charges":   "service_charges_cents",
        "Net Sales":         "net_sales_cents",
        "Gift Card Sales":   "gift_card_sales_cents",
        "Tax":               "tax_cents",
        "Tip":               "tip_cents",
        "Partial Refunds":   "partial_refunds_cents",
        "Total Collected":   "total_collected_cents",
        "Fees":              "fees_cents",
        "Net Total":         "net_total_cents",
        "Card":              "card_cents",
        "Cash":              "cash_cents",
        "Square Gift Card":  "square_gift_card_cents",
        "Other Tender":      "other_tender_cents",
    }
    for src, dst in col_rename.items():
        out[dst] = raw.get(src, "").map(_parse_money)

    # Drop rows without a recognized cafe (unknown location strings).
    out = out[out["cafe"].notna() & out["transaction_id"].ne("")]

    # Deduplicate on transaction_id (keeps the last occurrence → newer CSV wins).
    out = out.drop_duplicates(subset=["transaction_id"], keep="last").reset_index(drop=True)

    return out


def _slice(df, cafe, start_date, end_date, *, include_refunds=False):
    mask = (df["cafe"] == cafe) & (df["date"] >= start_date) & (df["date"] <= end_date)
    if not include_refunds:
        mask &= (df["event_type"] == "Payment")
    return df[mask]


def counts_and_totals(cafe, start_date, end_date, *, include_refunds=False):
    df = load()
    if df.empty:
        return (0, 0, 0, 0)
    s = _slice(df, cafe, start_date, end_date, include_refunds=include_refunds)
    if s.empty:
        return (0, 0, 0, 0)
    return (
        int(len(s)),
        int(s["net_sales_cents"].sum()),
        int(s["tip_cents"].sum()),
        int(s["gross_sales_cents"].sum()),
    )


def avg_ticket_cents(cafe, start_date, end_date):
    """Avg ticket EXCLUDING invoice-sourced transactions (Ryan rule, Aug 2026):
    invoices are billed catering/events, not walk-in tickets."""
    df = load()
    if df.empty: return None
    s = _slice(df, cafe, start_date, end_date)
    if "source" in s.columns:
        s = s[~s["source"].astype(str).str.contains("Invoice", na=False)]
    if s.empty: return None
    return float(s["net_sales_cents"].sum()) / len(s)


def avg_daily_transactions(cafe, start_date, end_date):
    df = load()
    if df.empty:
        return None
    s = _slice(df, cafe, start_date, end_date)
    if s.empty:
        return None
    daily = s.groupby("date").size()
    return float(daily.mean())


def dining_option_mix(cafe, start_date, end_date):
    df = load()
    if df.empty:
        return {}
    s = _slice(df, cafe, start_date, end_date)
    if s.empty:
        return {}
    d = s["dining_option"].replace("", "(unspecified)").value_counts()
    return d.to_dict()


def available_date_range():
    df = load()
    if df.empty:
        return (None, None, 0)
    return (df["date"].min(), df["date"].max(), int(len(df)))


def per_cafe_date_coverage():
    df = load()
    if df.empty:
        return []
    g = df.groupby("cafe")["date"].agg(["min", "max", "count"]).reset_index()
    return [(r["cafe"], r["min"], r["max"], int(r["count"])) for _, r in g.iterrows()]


if __name__ == "__main__":
    lo, hi, n = available_date_range()
    print(f"Transactions store ({TXN_DIR})")
    print(f"Rows: {n:,}   Date range: {lo} → {hi}")
    for cafe, lo, hi, n in per_cafe_date_coverage():
        print(f"  {cafe:<10}  {lo} → {hi}   n={n:,}")

    # Sanity: Jan totals per cafe
    print("\nJan 2026 sanity check:")
    for cafe in CAFES:
        n, net, tips, gross = counts_and_totals(cafe, "2026-01-01", "2026-01-31")
        at = avg_ticket_cents(cafe, "2026-01-01", "2026-01-31")
        adt = avg_daily_transactions(cafe, "2026-01-01", "2026-01-31")
        print(f"  {cafe:<10}  n={n:>5}  gross=${gross/100:>9,.0f}  net=${net/100:>9,.0f}  "
              f"avgTicket=${(at or 0)/100:>5,.2f}  avgDailyTx={adt or 0:>5,.0f}")
