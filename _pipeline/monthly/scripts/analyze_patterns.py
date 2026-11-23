#!/usr/bin/env python3
"""Sales-pattern analysis for the Monthly Business Review.

Reads the payment-level transactions store (Square Transactions CSV exports in
_pipeline/shared/transactions/) and the weekly labor JSONs, and produces the
JSON consumed by the report builder's "Sales Patterns & Insights" section.

What it computes (per cafe + Open-3 blend, for the target month, with the prior
month for MoM comparison):
  - hourly & daypart sales curve (net $ and transactions, share of day)
  - day-of-week revenue mix, per-open-day averages, avg ticket by DOW
  - CORE labor hours (admin + training excluded) by DOW overlaid on sales ->
    sales per labor hour, loaded cost per hour, txns per
    labor hour, and a "labor share vs sales share" alignment gap
  - ticket-size distribution and an attach-rate PROXY (share of tickets at or
    above a threshold calibrated on April, the last month with full multi-item
    descriptions)
  - dining-option mix (pickup vs in-store) with avg ticket per option
  - tip rate by cafe and daypart
  - lead-item velocity by daypart (first item on ticket -- descriptions are
    truncated to one item in Jul/Aug exports; the Square Item Sales CSV remains
    the source of truth for unit counts)

Usage:  python3 analyze_patterns.py 2026-08   [prior month auto = 2026-07]
Output: _pipeline/monthly/data/<yyyymm>_patterns.json
"""
import os, sys, json, glob, re
from collections import defaultdict
import pandas as pd
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))          # _pipeline
SHARED = os.path.join(PIPE, "shared")
TXN_DIR = os.path.join(SHARED, "transactions")
LABOR_DIR = os.path.join(PIPE, "weekly", "data")
DATA_OUT = os.path.join(os.path.dirname(HERE), "data")
sys.path.insert(0, os.path.join(PIPE, "weekly", "scripts"))
from aggregate_week import classify_item  # noqa: E402  (shared category rules)

LOCATION_MAP = {
    "Slipstream Downtown": "Downtown", "Slipstream Studio": "Studio",
    "Slipstream 1899 Penn": "1899", "Slipstream 1899": "1899", "Slipstream NMWA": "NMWA",
}
LABOR_CAFE_MAP = {"Downtown": "Downtown", "Studio": "Studio", "1899 Penn": "1899", "NMWA": "NMWA"}
OPEN3 = ["Downtown", "Studio", "1899"]
DOW = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

# Daypart buckets (hour boundaries, inclusive start / exclusive end)
DAYPARTS = [
    ("Early (6-8a)",      6,  8),
    ("Rush (8-10a)",      8, 10),
    ("Mid-morning (10-12p)", 10, 12),
    ("Lunch (12-2p)",    12, 14),
    ("Afternoon (2-5p)", 14, 17),
    ("Late (5p+)",       17, 24),
]

# Aliases for items the shared classifier misses (seasonal / renamed SKUs)
ALIASES = {
    "cruise control": "coffee_bev", "coco americano": "coffee_bev",
    "peachy keen matcha": "coffee_bev", "peach theory": "coffee_bev",
    "stellar collisions": "coffee_bev", "las esmeraldas": "coffee_bev",
    "earl grey lavender loaf": "food", "orange poppy": "food", "poppyseed": "food",
    "cornbread": "food", "tortilla de patatas": "food", "loaf": "food", "muffin": "food",
    "sparkling water": "other", "bottled water": "other", "soda": "other",
    "management fee": "invoice", "equipment fee": "invoice",
}


def money(s):
    return pd.to_numeric(s.astype(str).str.replace(r"[$,]", "", regex=True).str.replace(r"^\((.*)\)$", r"-\1", regex=True), errors="coerce").fillna(0.0)


def load_month(ym):
    # The weekly stores file a Sun-Sat bundle under its start month, so the first days of a month can sit
    # in the prior month's CSV (spillover). Read this month's file plus any neighbours and keep only <ym> dates.
    y_, m_ = int(ym[:4]), int(ym[5:7])
    prev_ = f"{y_ - (m_ == 1)}-{(m_ - 2) % 12 + 1:02d}"; next_ = f"{y_ + (m_ == 12)}-{m_ % 12 + 1:02d}"
    parts = [pd.read_csv(os.path.join(TXN_DIR, f"{k}.csv"), dtype=str, low_memory=False)
             for k in (prev_, ym, next_) if os.path.exists(os.path.join(TXN_DIR, f"{k}.csv"))]
    df = pd.concat(parts, ignore_index=True)
    df = df[df["Date"].fillna("").str.startswith(ym)]
    if "Transaction ID" in df.columns:
        df = df.drop_duplicates(subset=["Transaction ID", "Payment ID"] if "Payment ID" in df.columns else ["Transaction ID"])
    df = df[df["Transaction Status"].fillna("").str.lower() == "complete"].copy()
    df["cafe"] = df["Location"].map(LOCATION_MAP)
    df = df[df["cafe"].notna()].copy()
    for src, dst in [("Net Sales", "net"), ("Gross Sales", "gross"), ("Tip", "tip"), ("Discounts", "disc")]:
        df[dst] = money(df[src])
    df["dt"] = pd.to_datetime(df["Date"] + " " + df["Time"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
    df = df[df["dt"].notna()].copy()
    df["hour"] = df["dt"].dt.hour
    df["dow"] = df["dt"].dt.dayofweek          # 0=Mon
    df["date"] = df["dt"].dt.date.astype(str)
    df["dining"] = df["Dining Option"].fillna("").str.upper().replace({"": "UNKNOWN", "FOR HERE": "IN_STORE", "TO GO": "PICKUP"})
    df["desc"] = df["Description"].fillna("").str.strip()
    # invoice / non-walk-in detection (Source column == Invoices, or fee items)
    src = df["Source"].fillna("").str.lower()
    df["is_invoice"] = src.str.contains("invoice") | df["desc"].str.lower().str.contains("management fee|equipment fee")
    df["category"] = df["desc"].map(lambda d: classify_lead(d))
    return df


def classify_lead(desc):
    low = desc.lower()
    for k, v in ALIASES.items():
        if k in low:
            return v
    if not low:
        return "other"
    return classify_item(desc)


def daypart_of(h):
    for name, lo, hi in DAYPARTS:
        if lo <= h < hi:
            return name
    return "Late (5p+)"


def load_labor(ym):
    """Return {cafe: {date: {hours, cents}}} for every day covered by weekly labor JSONs."""
    out = defaultdict(dict)
    for f in glob.glob(os.path.join(LABOR_DIR, "labor_*.json")):
        try:
            j = json.load(open(f))
        except Exception:
            continue
        for cafe_raw, days in (j.get("per_cafe_per_day") or {}).items():
            cafe = LABOR_CAFE_MAP.get(cafe_raw, cafe_raw)
            for d, rec in days.items():
                if d.startswith(ym):
                    # CORE labor only (Ryan, Sep 2026): exclude admin (title "Manager") and training.
                    out[cafe][d] = {"hours": float(rec.get("core_hours", rec.get("labor_hours", 0.0))),
                                    "cents": int(rec.get("core_cents", rec.get("labor_cents", 0))),
                                    "admin_hours": float(rec.get("admin_hours", 0.0)),
                                    "training_hours": float(rec.get("training_hours", 0.0))}
    return out


def pct(n, d):
    return round(n / d * 100, 1) if d else None


def analyze(ym, prior_ym):
    cur = load_month(ym)
    pri = load_month(prior_ym)
    labor = load_labor(ym)
    walk = cur[(~cur["is_invoice"]) & (cur["net"] > 0)].copy()      # walk-in tickets only for pattern work
    walk_p = pri[(~pri["is_invoice"]) & (pri["net"] > 0)].copy()

    # ---- attach-rate threshold calibration on April (full descriptions) ----
    try:
        apr = load_month("2026-04")
        apr = apr[(~apr["is_invoice"]) & (apr["net"] > 0)]
        apr_multi = apr["desc"].str.count(",") > 0
        # pick the threshold t in [7..12] that best matches actual multi-item rate
        best_t, best_err = 9.0, 9e9
        for t in np.arange(7.0, 12.5, 0.25):
            err = abs((apr["net"] >= t).mean() - apr_multi.mean())
            if err < best_err:
                best_t, best_err = float(t), err
        attach_cal = {"april_multi_item_rate": round(apr_multi.mean() * 100, 1), "threshold": best_t,
                      "april_rate_at_threshold": round((apr["net"] >= best_t).mean() * 100, 1)}
    except Exception:
        best_t = 9.0
        attach_cal = {"threshold": 9.0, "note": "April calibration unavailable"}

    out = {"month": ym, "prior": prior_ym, "attach_calibration": attach_cal, "cafes": {}, "open3": {}}

    def cafe_block(df, dfp, lab):
        b = {}
        n = len(df); net = df["net"].sum()
        b["txns"] = int(n); b["net"] = round(net, 2)
        b["avg_ticket"] = round(net / n, 2) if n else None
        b["open_days"] = int(df["date"].nunique())
        # hourly
        hr = df.groupby("hour").agg(net=("net", "sum"), txns=("net", "size"))
        b["hourly"] = {int(h): {"net": round(r.net, 2), "txns": int(r.txns), "share": pct(r.net, net)} for h, r in hr.iterrows()}
        # daypart
        df["dp"] = df["hour"].map(daypart_of)
        dp = df.groupby("dp").agg(net=("net", "sum"), txns=("net", "size"), tip=("tip", "sum"))
        b["daypart"] = {}
        for name, _, _ in DAYPARTS:
            if name in dp.index:
                r = dp.loc[name]
                b["daypart"][name] = {"net": round(r.net, 2), "txns": int(r.txns), "share": pct(r.net, net),
                                      "avg_ticket": round(r.net / r.txns, 2) if r.txns else None,
                                      "tip_rate": pct(r.tip, r.net)}
            else:
                b["daypart"][name] = {"net": 0.0, "txns": 0, "share": 0.0, "avg_ticket": None, "tip_rate": None}
        # prior-month daypart share for delta
        dfp = dfp.copy(); dfp["dp"] = dfp["hour"].map(daypart_of)
        dpp = dfp.groupby("dp")["net"].sum(); netp = dfp["net"].sum()
        for name in b["daypart"]:
            b["daypart"][name]["share_prior"] = pct(dpp.get(name, 0.0), netp)
        # day of week
        dow = df.groupby("dow").agg(net=("net", "sum"), txns=("net", "size"), days=("date", "nunique"))
        dow_m = df[df["date"].isin(set(lab.keys()))].groupby("dow").agg(net=("net", "sum"), txns=("net", "size"), days=("date", "nunique")) if lab else dow
        b["dow"] = {}
        for i, d in enumerate(DOW):
            if i in dow.index:
                r = dow.loc[i]
                lh = sum(v["hours"] for dt, v in lab.items() if pd.Timestamp(dt).dayofweek == i)
                lc = sum(v["cents"] for dt, v in lab.items() if pd.Timestamp(dt).dayofweek == i) / 100.0
                ldays = sum(1 for dt in lab if pd.Timestamp(dt).dayofweek == i)
                b["dow"][d] = {"net": round(r.net, 2), "txns": int(r.txns), "days": int(r.days),
                               "net_per_day": round(r.net / r.days, 2) if r.days else None,
                               "txns_per_day": round(r.txns / r.days, 1) if r.days else None,
                               "avg_ticket": round(r.net / r.txns, 2) if r.txns else None,
                               "share": pct(r.net, net),
                               "labor_hours": round(lh, 1), "labor_cost": round(lc, 2), "labor_days": ldays,
                               "labor_hours_per_day": round(lh / ldays, 1) if ldays else None,
                               "sales_per_labor_hour": round(dow_m.loc[i].net / lh, 2) if (lh and i in dow_m.index) else None,
                               "txns_per_labor_hour": round(dow_m.loc[i].txns / lh, 1) if (lh and i in dow_m.index) else None,
                               "labor_pct_of_sales": pct(lc, dow_m.loc[i].net) if (lh and i in dow_m.index) else None,
                               "cost_per_hour": round(lc / lh, 2) if lh else None}
        # labor totals — SPLH / cost ratios computed on MATCHED days only (days present in labor data),
        # so partial labor coverage never inflates sales-per-labor-hour.
        tot_h = sum(v["hours"] for v in lab.values()); tot_c = sum(v["cents"] for v in lab.values()) / 100.0
        matched = df[df["date"].isin(set(lab.keys()))]
        m_net, m_n = matched["net"].sum(), len(matched)
        b["labor"] = {"hours": round(tot_h, 1), "cost": round(tot_c, 2), "days_covered": len(lab),
                      "open_days_in_month": int(df["date"].nunique()),
                      "matched_days_net": round(m_net, 2), "matched_days_txns": int(m_n),
                      "admin_hours": round(sum(v.get("admin_hours", 0) for v in lab.values()), 1),
                      "training_hours": round(sum(v.get("training_hours", 0) for v in lab.values()), 1),
                      "cost_per_hour": round(tot_c / tot_h, 2) if tot_h else None,
                      "labor_pct_of_sales": pct(tot_c, m_net) if tot_h else None,
                      "sales_per_labor_hour": round(m_net / tot_h, 2) if tot_h else None,
                      "txns_per_labor_hour": round(m_n / tot_h, 1) if tot_h else None}
        # labor share vs sales share by DOW (alignment gap)
        if tot_h:
            mnet_tot = dow_m["net"].sum()
            for i, d in enumerate(DOW):
                if d not in b["dow"]:
                    continue
                ls = pct(b["dow"][d]["labor_hours"], tot_h)
                ss = pct(dow_m.loc[i].net, mnet_tot) if i in dow_m.index else None
                b["dow"][d]["labor_share"] = ls
                b["dow"][d]["sales_share_matched"] = ss
                b["dow"][d]["alignment_gap"] = round((ls or 0) - (ss or 0), 1) if ss is not None else None
        # ticket distribution
        bins = [0, 5, 8, 10, 15, 25, 1e9]; labels = ["<$5", "$5-8", "$8-10", "$10-15", "$15-25", "$25+"]
        cut = pd.cut(df["net"], bins=bins, labels=labels, right=False)
        vc = cut.value_counts().reindex(labels).fillna(0)
        b["ticket_dist"] = {k: {"txns": int(v), "share": pct(v, n)} for k, v in vc.items()}
        b["median_ticket"] = round(float(df["net"].median()), 2) if n else None
        # attach proxy
        b["attach_proxy"] = pct((df["net"] >= best_t).sum(), n)
        b["attach_proxy_prior"] = pct((dfp["net"] >= best_t).sum(), len(dfp))
        # attach proxy by daypart
        b["attach_by_daypart"] = {name: pct((df[df["dp"] == name]["net"] >= best_t).sum(), (df["dp"] == name).sum()) for name, _, _ in DAYPARTS}
        # dining option
        dm = df.groupby("dining").agg(net=("net", "sum"), txns=("net", "size"))
        b["dining"] = {k: {"txns": int(r.txns), "share": pct(r.txns, n), "avg_ticket": round(r.net / r.txns, 2)} for k, r in dm.iterrows()}
        # tips
        b["tip_rate"] = pct(df["tip"].sum(), net)
        b["tip_rate_prior"] = pct(dfp["tip"].sum(), dfp["net"].sum())
        # lead item by daypart (top 5 per daypart) + category share by daypart
        li = {}
        for name, _, _ in DAYPARTS:
            sub = df[df["dp"] == name]
            top = sub["desc"].str.replace(r"\s*\(.*?\)\s*$", "", regex=True).value_counts().head(5)
            li[name] = [{"item": k, "txns": int(v)} for k, v in top.items()]
        b["lead_item_by_daypart"] = li
        cat = df.groupby(["dp", "category"]).size().unstack(fill_value=0)
        b["lead_category_by_daypart"] = {dp: {c: pct(cat.loc[dp, c], cat.loc[dp].sum()) for c in cat.columns} for dp in cat.index}
        # top lead items overall + category totals
        top_all = df["desc"].str.replace(r"\s*\(.*?\)\s*$", "", regex=True).value_counts().head(15)
        b["top_lead_items"] = [{"item": k, "txns": int(v)} for k, v in top_all.items()]
        b["lead_category_mix"] = {c: pct(v, n) for c, v in df["category"].value_counts().items()}
        # prior-month comparators
        b["prior"] = {"txns": int(len(dfp)), "net": round(dfp["net"].sum(), 2),
                      "avg_ticket": round(dfp["net"].sum() / len(dfp), 2) if len(dfp) else None,
                      "open_days": int(dfp["date"].nunique())}
        return b

    for cafe in ["Downtown", "Studio", "1899", "NMWA"]:
        out["cafes"][cafe] = cafe_block(walk[walk["cafe"] == cafe].copy(), walk_p[walk_p["cafe"] == cafe].copy(), labor.get(cafe, {}))
    lab3 = {}
    for c in OPEN3:
        for d, v in labor.get(c, {}).items():
            lab3.setdefault(d, {"hours": 0.0, "cents": 0})
            lab3[d]["hours"] += v["hours"]; lab3[d]["cents"] += v["cents"]
    out["open3"] = cafe_block(walk[walk["cafe"].isin(OPEN3)].copy(), walk_p[walk_p["cafe"].isin(OPEN3)].copy(), lab3)

    # walk-in avg ticket by month, all months in the store (for the Ops Indicators trend)
    wt = {c: {} for c in ["Downtown", "Studio", "1899", "NMWA"]}
    for f_ in sorted(glob.glob(os.path.join(TXN_DIR, "2026-*.csv"))):
        ym_ = os.path.basename(f_)[:7]
        if ym_ > ym:
            continue
        try:
            d_ = load_month(ym_)
        except Exception:
            continue
        d_ = d_[(~d_["is_invoice"]) & (d_["net"] > 0)]
        mon = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][int(ym_[5:]) - 1]
        for c in wt:
            sub = d_[d_["cafe"] == c]
            wt[c][mon] = round(sub["net"].mean(), 2) if len(sub) >= 50 else None
    out["walkin_ticket_by_month"] = wt

    # invoice/non-walk-in summary (so the report can explain reconciliation to the model)
    inv = cur[cur["is_invoice"]]
    out["invoice_summary"] = {c: {"txns": int((inv["cafe"] == c).sum()), "net": round(inv[inv["cafe"] == c]["net"].sum(), 2)} for c in ["Downtown", "Studio", "1899", "NMWA"]}
    out["zero_dollar_txns"] = int((cur["net"] <= 0).sum())

    os.makedirs(DATA_OUT, exist_ok=True)
    path = os.path.join(DATA_OUT, f"{ym.replace('-', '')[2:]}_patterns.json")
    json.dump(out, open(path, "w"), indent=1)
    return path, out


if __name__ == "__main__":
    ym = sys.argv[1] if len(sys.argv) > 1 else "2026-08"
    y, m = map(int, ym.split("-"))
    prior = f"{y if m > 1 else y-1}-{(m-1) if m > 1 else 12:02d}"
    path, out = analyze(ym, prior)
    print("Wrote", path)
    print("Attach calibration:", out["attach_calibration"])
    for c in ["Downtown", "Studio", "1899", "NMWA"]:
        b = out["cafes"][c]
        print(f"\n{c}: txns={b['txns']} net=${b['net']:,.0f} avg=${b['avg_ticket']} days={b['open_days']} "
              f"tip={b['tip_rate']}% attach~{b['attach_proxy']}% laborHrs={b['labor']['hours']} SPLH=${b['labor']['sales_per_labor_hour']}")
        print("  daypart share:", {k: v['share'] for k, v in b['daypart'].items()})
        print("  dow $/day:", {k: v['net_per_day'] for k, v in b['dow'].items()})
        print("  dow align gap (labor% - sales%):", {k: v.get('alignment_gap') for k, v in b['dow'].items()})
        print("  dining:", {k: (v['share'], v['avg_ticket']) for k, v in b['dining'].items()})
    b = out["open3"]
    print(f"\nOPEN3: txns={b['txns']} net=${b['net']:,.0f} avg=${b['avg_ticket']} tip={b['tip_rate']}% attach~{b['attach_proxy']}% SPLH=${b['labor']['sales_per_labor_hour']}")
    print("  daypart share:", {k: v['share'] for k, v in b['daypart'].items()})
    print("  invoice:", out["invoice_summary"])
