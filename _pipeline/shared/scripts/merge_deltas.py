#!/usr/bin/env python3
"""Fold weekly cloud delta bundles into the big Slipstream stores. Stdlib only. Idempotent.

    python3 merge_deltas.py BUNDLE_OR_DIR [more ...] [--shared PATH_TO__pipeline/shared] [--dry-run]

Bundles come from slipstream_cloud_delta.py (delta_<START>_<END>.tar.xz.b64).
For each bundle it updates, inside --shared:
  transactions/YYYY-MM.csv   rows in [start,end] are REPLACED by the delta's rows (rest untouched)
  daily_metrics.csv          upsert on (date, cafe)
  loyalty_events/YYYY-MM.jsonl   add events by id (dedupe), month = created_at[:7], sorted by created_at
and appends a line to deltas_merged.log in --shared. Re-running a bundle changes nothing.
--shared defaults to the parent of this script's folder (i.e. run it from _pipeline/shared/scripts/).
"""
import os, sys, io, csv, json, glob, base64, tarfile, argparse, hashlib
from collections import defaultdict
from datetime import datetime, timezone

DM_COLS = ["date","cafe","transactions","net_sales","tips","coffee_bev","food","retail_coffee","other",
           "avg_ticket","walkin_transactions","walkin_net_sales"]

def read_bundle(path):
    raw = base64.b64decode(open(path).read())
    t = tarfile.open(fileobj=io.BytesIO(raw), mode="r:*")
    g = lambda n: t.extractfile(n).read().decode()
    return {"manifest": json.loads(g("manifest.json")), "transactions": g("transactions.csv"),
            "daily": g("daily_metrics.csv"), "loyalty": g("loyalty_events.jsonl"), "sha": hashlib.sha256(raw).hexdigest()[:12]}

def verify(b):
    for name, key in (("transactions.csv", "transactions"), ("daily_metrics.csv", "daily"), ("loyalty_events.jsonl", "loyalty")):
        want = b["manifest"]["sha256"][name]
        if hashlib.sha256(b[key].encode()).hexdigest() != want:
            raise SystemExit(f"CHECKSUM MISMATCH in bundle for {name} - bundle corrupt, not merging")

def merge_transactions(shared, b, dry):
    start, end = b["manifest"]["start"], b["manifest"]["end"]
    rd = csv.DictReader(io.StringIO(b["transactions"])); cols = rd.fieldnames
    new = defaultdict(list)
    for r in rd: new[r["Date"][:7]].append(r)
    n = 0
    for ym, rows in sorted(new.items()):
        p = os.path.join(shared, "transactions", f"{ym}.csv")
        keep = []
        if os.path.exists(p):
            with open(p, newline="") as f:
                for r in csv.DictReader(f):
                    if not (start <= r.get("Date", "") <= end): keep.append(r)
        allr = keep + rows
        allr.sort(key=lambda r: (r.get("Date", ""), r.get("Time", ""), r.get("Transaction ID", "")))
        n += len(rows)
        if not dry:
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", newline="") as f:
                w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore"); w.writeheader(); w.writerows(allr)
    return n

def merge_daily(shared, b, dry):
    p = os.path.join(shared, "daily_metrics.csv")
    ex = {}
    if os.path.exists(p):
        with open(p, newline="") as f:
            for r in csv.DictReader(f): ex[(r["date"], r["cafe"])] = r
    new = list(csv.DictReader(io.StringIO(b["daily"])))
    for r in new: ex[(r["date"], r["cafe"])] = r
    if not dry:
        with open(p, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=DM_COLS); w.writeheader()
            for r in sorted(ex.values(), key=lambda r: (str(r["date"]), str(r["cafe"]))):
                w.writerow({k: r.get(k, "") for k in DM_COLS})
    return len(new)

def merge_loyalty(shared, b, dry):
    evs = [json.loads(l) for l in b["loyalty"].splitlines() if l.strip()]
    by = defaultdict(list)
    for e in evs: by[e["created_at"][:7]].append(e)
    added = 0
    for ym, es in sorted(by.items()):
        p = os.path.join(shared, "loyalty_events", f"{ym}.jsonl")
        ex = {}
        if os.path.exists(p):
            for l in open(p):
                if l.strip(): j = json.loads(l); ex[j["id"]] = j
        for e in es:
            if e["id"] not in ex: ex[e["id"]] = e; added += 1
        if not dry:
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w") as f:
                for e in sorted(ex.values(), key=lambda x: x["created_at"]): f.write(json.dumps(e) + "\n")
    return added

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+"); ap.add_argument("--shared"); ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    shared = a.shared or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files = []
    for p in a.paths:
        files += sorted(glob.glob(os.path.join(p, "delta_*.tar.xz.b64"))) if os.path.isdir(p) else [p]
    for f in files:
        b = read_bundle(f); verify(b); m = b["manifest"]
        t = merge_transactions(shared, b, a.dry_run); d = merge_daily(shared, b, a.dry_run); l = merge_loyalty(shared, b, a.dry_run)
        msg = f"{m['start']}..{m['end']}: txn rows {t}, daily rows {d}, loyalty events new {l} / {m['loyalty_events']} in bundle [{b['sha']}]"
        print(("DRY-RUN " if a.dry_run else "") + msg)
        if not a.dry_run:
            open(os.path.join(shared, "deltas_merged.log"), "a").write(f"{datetime.now(timezone.utc).isoformat()} {msg}\n")

if __name__ == "__main__":
    main()
