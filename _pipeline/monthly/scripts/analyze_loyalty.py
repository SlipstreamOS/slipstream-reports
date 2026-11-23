#!/usr/bin/env python3
"""Loyalty analysis for the Monthly Business Review — replaces screenshot input.

Reads the loyalty-events store (Square Loyalty API ledger) and produces, per
cafe and blended, the same figures the report's Section 6 has always carried:
  - loyalty visits (ACCUMULATE_POINTS events), unique members, avg visits/member
  - first-time vs repeat  (first-time = account has no earlier event in the store;
    needs >= 1 prior month in the store to be exact — otherwise flagged)
  - visit-frequency buckets (1 / 2 / 3-4 / 5-7 / 8+), stable schema every month
  - loyalty-attributed pre-tax spend (points / 7) and avg spend per member
  - reward redemptions, and points-per-visit as a proxy for loyalty ticket
  - loyalty share of walk-in transactions (visits / walk-in tickets from the
    transactions store)

Usage:  python3 analyze_loyalty.py 2026-08
Output: data/<yymm>_loyalty.json
"""
import os, sys, json
from collections import defaultdict, Counter
HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(PIPE, "shared", "scripts"))
import loyalty_store as ls  # noqa: E402

CAFES = ["Downtown", "Studio", "1899", "NMWA"]
OPEN3 = ["Downtown", "Studio", "1899"]
BUCKETS = [("1 visit", 1, 1), ("2 visits", 2, 2), ("3–4 visits", 3, 4), ("5–7 visits", 5, 7), ("8+ visits", 8, 10**9)]


def bucket(n):
    for name, lo, hi in BUCKETS:
        if lo <= n <= hi:
            return name
    return "8+ visits"


def walkin_tickets(ym):
    """Walk-in transaction counts per cafe from the transactions store (for loyalty share)."""
    try:
        import pandas as pd
        y_, m_ = int(ym[:4]), int(ym[5:7])   # month's first days can sit in the prior month's file (Sun-Sat bundles)
        keys = [f"{y_ - (m_ == 1)}-{(m_ - 2) % 12 + 1:02d}", ym, f"{y_ + (m_ == 12)}-{m_ % 12 + 1:02d}"]
        fs = [os.path.join(PIPE, "shared", "transactions", f"{k}.csv") for k in keys]
        df = pd.concat([pd.read_csv(f, dtype=str, low_memory=False) for f in fs if os.path.exists(f)], ignore_index=True)
        df = df[df["Date"].fillna("").str.startswith(ym)]
        if "Transaction ID" in df.columns:
            df = df.drop_duplicates(subset=["Transaction ID", "Payment ID"] if "Payment ID" in df.columns else ["Transaction ID"])
        df = df[df["Transaction Status"].fillna("").str.lower() == "complete"]
        net = df["Net Sales"].str.replace(r"[$,]", "", regex=True).astype(float)
        df = df[net > 0]
        m = {"Slipstream Downtown": "Downtown", "Slipstream Studio": "Studio", "Slipstream 1899 Penn": "1899", "Slipstream NMWA": "NMWA"}
        return df["Location"].map(m).value_counts().to_dict()
    except Exception:
        return {}


def analyze(ym):
    cur = [e for e in ls.load(ym)]
    prior_all = ls.load_range("2000-01", f"{ym[:4]}-{int(ym[5:]) - 1:02d}" if int(ym[5:]) > 1 else f"{int(ym[:4]) - 1}-12")
    seen_before = {e["loyalty_account_id"] for e in prior_all}
    prior_months = len({e["created_at"][:7] for e in prior_all})
    tix = walkin_tickets(ym)

    def block(evs, cafe_label):
        acc = [e for e in evs if e["type"] == "ACCUMULATE_POINTS"]
        red = [e for e in evs if e["type"] == "REDEEM_REWARD"]
        per_acct = Counter(e["loyalty_account_id"] for e in acc)
        pts = sum(e["accumulate_points"]["points"] for e in acc)
        members = len(per_acct)
        first_time = sum(1 for a in per_acct if a not in seen_before)
        buckets = Counter(bucket(n) for n in per_acct.values())
        visits = len(acc)
        spend = pts / ls.POINTS_PER_DOLLAR
        return {
            "visits": visits, "members": members,
            "avg_visits_per_member": round(visits / members, 2) if members else None,
            "first_time_members": first_time, "repeat_members": members - first_time,
            "first_time_visits": sum(n for a, n in per_acct.items() if a not in seen_before),
            "repeat_visits": sum(n for a, n in per_acct.items() if a in seen_before),
            "buckets": {name: buckets.get(name, 0) for name, _, _ in BUCKETS},
            "points": pts, "loyalty_spend_pretax": round(spend, 2),
            "avg_spend_per_member": round(spend / members, 2) if members else None,
            "avg_spend_per_visit": round(spend / visits, 2) if visits else None,
            "redemptions": len(red),
            "walkin_tickets": tix.get(cafe_label),
            "loyalty_share_of_tickets": round(visits / tix[cafe_label] * 100, 1) if tix.get(cafe_label) else None,
            "days_with_events": len({e["created_at"][:10] for e in acc}),
        }

    out = {"month": ym, "prior_months_in_store": prior_months,
           "first_time_note": ("first-time = no prior event in store" if prior_months else
                               "NO prior months in store — first-time counts are upper bounds"),
           "cafes": {}, "open3": None, "all": None}
    for c in CAFES:
        out["cafes"][c] = block([e for e in cur if ls.LOCATION_MAP.get(e["location_id"]) == c], c)
    o3 = [e for e in cur if ls.LOCATION_MAP.get(e["location_id"]) in OPEN3]
    b = block(o3, "__o3__"); b["walkin_tickets"] = sum(tix.get(c, 0) for c in OPEN3) or None
    b["loyalty_share_of_tickets"] = round(b["visits"] / b["walkin_tickets"] * 100, 1) if b["walkin_tickets"] else None
    out["open3"] = b
    out["coverage"] = ls.coverage().get(ym)
    path = os.path.join(os.path.dirname(HERE), "data", f"{ym.replace('-', '')[2:]}_loyalty.json")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    json.dump(out, open(path, "w"), indent=1)
    return path, out


if __name__ == "__main__":
    ym = sys.argv[1] if len(sys.argv) > 1 else "2026-08"
    path, out = analyze(ym)
    print("Wrote", path)
    print("Coverage:", out["coverage"], "|", out["first_time_note"])
    for c in CAFES + ["open3"]:
        b = out["cafes"][c] if c != "open3" else out["open3"]
        print(f"{c:9s} visits={b['visits']:5d} members={b['members']:4d} avg={b['avg_visits_per_member']} "
              f"first={b['first_time_members']} spend=${b['loyalty_spend_pretax']:,.0f} $/visit={b['avg_spend_per_visit']} "
              f"redeem={b['redemptions']} share={b['loyalty_share_of_tickets']}% buckets={b['buckets']}")
