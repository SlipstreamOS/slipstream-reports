"""Slipstream loyalty-events store — Square Loyalty API ledger, one JSONL per month.

Location: _pipeline/shared/loyalty_events/YYYY-MM.jsonl   (one event per line)
Idempotent on event `id`.

Why: the Square Loyalty dashboard screenshots are lossy (unlabeled bar charts).
The Loyalty API `searchEvents` ledger gives every visit (ACCUMULATE_POINTS),
redemption (REDEEM_REWARD), account ID, timestamp, location and points —
enough to rebuild the whole monthly loyalty section deterministically.

Pull pattern (30 events/page, cursor pagination):
  loyalty.searchEvents  {query:{filter:{date_time_filter:{created_at:{start_at,end_at}},
                                          location_filter:{location_ids:[...]}}}, limit:30, cursor}
  ~3 pages per cafe-day at Downtown; ~35 pages/week for all four cafes.
  Intended to run inside the Sunday weekly pull; the aggregator reads the store.

Public API:
  ingest(events: list[dict]) -> int         # add/dedupe, returns new count
  ingest_response_text(text: str) -> int    # accepts the raw make_api_request JSON text
  ingest_dir(dirpath) -> int                # ingest every *.json page file in a folder (weekly job)
  load(ym: str) -> list[dict]               # all events for a month
  load_range(start_ym, end_ym) -> list[dict]
  coverage() -> {ym: {"events": n, "min": iso, "max": iso}}
"""
import os, json, glob
from collections import defaultdict

_THIS = os.path.dirname(os.path.abspath(__file__))
STORE_DIR = os.path.join(os.path.dirname(_THIS), "loyalty_events")
LOCATION_MAP = {
    "LXM7V1QWDHKZJ": "Downtown", "LGVKCKX9B9DZT": "Studio",
    "L4ZBMGCWJ5C7R": "1899", "LV2D18H0DD9ES": "NMWA",
}
POINTS_PER_DOLLAR = 7.0   # earn rule: 7 pts per $1 pre-tax


def _path(ym):
    return os.path.join(STORE_DIR, f"{ym}.jsonl")


def _ym_of(ev):
    return ev["created_at"][:7]


def ingest(events):
    os.makedirs(STORE_DIR, exist_ok=True)
    by_month = defaultdict(list)
    for e in events:
        by_month[_ym_of(e)].append(e)
    added = 0
    for ym, evs in by_month.items():
        existing = {}
        if os.path.exists(_path(ym)):
            for line in open(_path(ym)):
                if line.strip():
                    j = json.loads(line); existing[j["id"]] = j
        for e in evs:
            if e["id"] not in existing:
                existing[e["id"]] = e; added += 1
        with open(_path(ym), "w") as fh:
            for e in sorted(existing.values(), key=lambda x: x["created_at"]):
                fh.write(json.dumps(e) + "\n")
    return added


def ingest_response_text(text):
    j = json.loads(text) if isinstance(text, str) else text
    return ingest(j.get("events", []))


def load(ym):
    if not os.path.exists(_path(ym)):
        return []
    return [json.loads(l) for l in open(_path(ym)) if l.strip()]


def load_range(start_ym, end_ym):
    out = []
    for f in sorted(glob.glob(os.path.join(STORE_DIR, "*.jsonl"))):
        ym = os.path.basename(f)[:7]
        if start_ym <= ym <= end_ym:
            out.extend(load(ym))
    return out


def ingest_dir(dirpath, delete_after=False):
    """Ingest every *.json file in dirpath (raw make_api_request responses). Idempotent — re-ingesting is safe.
    The workspace mount does not permit deletes, so write raw pages to /tmp and leave delete_after=False."""
    total = 0
    for f in sorted(glob.glob(os.path.join(dirpath, "*.json"))):
        try:
            total += ingest_response_text(open(f).read())
            if delete_after:
                os.remove(f)
        except Exception as ex:
            print("skip", f, ex)
    return total


def coverage():
    out = {}
    for f in sorted(glob.glob(os.path.join(STORE_DIR, "*.jsonl"))):
        ym = os.path.basename(f)[:7]
        evs = load(ym)
        if evs:
            out[ym] = {"events": len(evs), "min": min(e["created_at"] for e in evs), "max": max(e["created_at"] for e in evs)}
    return out


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "ingest":
        n = ingest_response_text(sys.stdin.read()); print("ingested", n)
    elif len(sys.argv) > 2 and sys.argv[1] == "ingest-dir":
        n = ingest_dir(sys.argv[2]); print("ingested", n, "new events from", sys.argv[2])
        print(json.dumps(coverage(), indent=1))
    else:
        print(json.dumps(coverage(), indent=1))
