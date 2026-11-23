#!/usr/bin/env python3
"""Labor Deep Dive step 1 — build a compact walk-in transaction table from the
Square Transactions CSV store (Jan–Aug 2026), applying the standard exclusions:
Complete only, net > 0, not invoice-sourced (Source contains 'invoice' or a
Management/Equipment Fee description). Output: data/labor_deep_dive/txns_2026.csv
"""
import os, glob, sys
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))
TXN_DIR = os.path.join(PIPE, "shared", "transactions")
OUT_DIR = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive")
os.makedirs(OUT_DIR, exist_ok=True)

LOCATION_MAP = {
    "Slipstream Downtown": "Downtown", "Slipstream Studio": "Studio",
    "Slipstream 1899 Penn": "1899 Penn", "Slipstream 1899": "1899 Penn", "Slipstream NMWA": "NMWA",
}

def money(s):
    return pd.to_numeric(s.astype(str).str.replace(r"[$,]", "", regex=True)
                         .str.replace(r"^\((.*)\)$", r"-\1", regex=True), errors="coerce").fillna(0.0)

frames = []
stats = []
for path in sorted(glob.glob(os.path.join(TXN_DIR, "2026-*.csv"))):
    ym = os.path.basename(path)[:7]
    df = pd.read_csv(path, dtype=str, low_memory=False)
    n0 = len(df)
    status_ok = df["Transaction Status"].fillna("").str.lower() == "complete"
    df = df[status_ok].copy()
    n1 = len(df)
    df["cafe"] = df["Location"].map(LOCATION_MAP)
    unmapped = df["Location"][df["cafe"].isna()].value_counts().to_dict()
    df = df[df["cafe"].notna()].copy()
    df["net"] = money(df["Net Sales"]); df["gross"] = money(df["Gross Sales"]); df["tip"] = money(df["Tip"])
    src = df["Source"].fillna("").str.lower()
    desc = df["Description"].fillna("").str.lower()
    df["is_invoice"] = src.str.contains("invoice") | desc.str.contains("management fee|equipment fee")
    n_inv = int(df["is_invoice"].sum()); inv_net = float(df.loc[df["is_invoice"], "net"].sum())
    n_nonpos = int((df["net"] <= 0).sum())
    df = df[(~df["is_invoice"]) & (df["net"] > 0)].copy()
    df["dt"] = pd.to_datetime(df["Date"] + " " + df["Time"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
    df = df[df["dt"].notna()].copy()
    tz = df["Time Zone"].fillna("").str.contains("Eastern").mean()
    out = pd.DataFrame({
        "date": df["dt"].dt.date.astype(str), "hour": df["dt"].dt.hour, "minute": df["dt"].dt.minute,
        "dow": df["dt"].dt.dayofweek, "cafe": df["cafe"], "net": df["net"].round(2),
        "gross": df["gross"].round(2), "tip": df["tip"].round(2),
        "dining": df["Dining Option"].fillna("").str.upper(),
        "source": df["Source"].fillna(""),
    })
    frames.append(out)
    stats.append((ym, n0, n1, n_inv, round(inv_net), n_nonpos, len(out), round(tz, 3), unmapped))

all_ = pd.concat(frames, ignore_index=True)
all_.to_csv(os.path.join(OUT_DIR, "txns_2026.csv"), index=False)
print("month   rows  complete  invoice(n,$)     net<=0  kept   eastern_share  unmapped_locations")
for s in stats:
    print(f"{s[0]}  {s[1]:6d}  {s[2]:7d}  {s[3]:5d} ${s[4]:>8,}  {s[5]:6d}  {s[6]:6d}  {s[7]:.3f}  {s[8]}")
print("\nKept rows:", len(all_))
print(all_.groupby("cafe")["net"].agg(["count", "sum"]).round(0))
