"""Harvest Square Loyalty events straight out of the chat transcript.

WHY THIS EXISTS
---------------
`loyalty.searchEvents` is capped at 30 events/page and there is no bulk export,
so every event has to pass through the model's context on its way in. The old
weekly pattern then had the model *re-type* each page verbatim into
/tmp/loyalty_pages/*.json before `loyalty_store.py ingest-dir` could read it —
which doubles the token cost of the pull and is the slow, error-prone part.

The harness already writes every tool result verbatim into the session
transcript JSONL on disk. So the pages are ALREADY saved — we can just read
them back out. The model only has to paginate; nothing gets transcribed by hand.

Discovered Sep 2026 during the Aug-2026 backfill (~4,500 events, 152 pages).

USAGE (weekly job)
------------------
    # after paginating searchEvents for every cafe:
    python3 _pipeline/shared/scripts/loyalty_from_transcript.py

    python3 ... loyalty_from_transcript.py --dry-run        # report, don't ingest
    python3 ... loyalty_from_transcript.py --since 2026-08-30 --until 2026-09-06
    python3 ... loyalty_from_transcript.py --transcript /path/to/session.jsonl
    python3 ... loyalty_from_transcript.py --also-dir /tmp/loyalty_pages

Idempotent: `loyalty_store` dedupes on event id, so re-running is always safe.
"""
import os, sys, json, glob
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import loyalty_store
from _paths import WORKSPACE

LOCATION_MAP = loyalty_store.LOCATION_MAP
# workspace = <mount>/Weekly and Monthly Reports  →  transcripts at <mount>/.claude/projects
MOUNT = os.path.dirname(WORKSPACE)


def find_transcript(explicit=None):
    """Newest *.jsonl under the mount's .claude/projects (the live session log)."""
    if explicit:
        return explicit
    roots = [
        os.path.join(MOUNT, ".claude", "projects"),
        os.path.expanduser("~/.claude/projects"),
    ]
    best, best_mtime = None, -1
    for root in roots:
        for f in glob.glob(os.path.join(root, "**", "*.jsonl"), recursive=True):
            try:
                m = os.path.getmtime(f)
            except OSError:
                continue
            if m > best_mtime:
                best, best_mtime = f, m
    return best


def _result_texts(line):
    """Yield the text payload of every tool_result in a transcript line."""
    try:
        obj = json.loads(line)
    except Exception:
        return
    for c in (obj.get("message", {}).get("content") or []):
        if not isinstance(c, dict) or c.get("type") != "tool_result":
            continue
        content = c.get("content")
        if isinstance(content, list):
            yield "".join(t.get("text", "") for t in content if isinstance(t, dict))
        elif isinstance(content, str):
            yield content


def extract(transcript):
    """All distinct loyalty events found in the transcript, keyed by event id."""
    events = {}
    with open(transcript, errors="replace") as fh:
        for line in fh:
            # cheap prefilter — transcript lines are large and mostly irrelevant
            if ("ACCUMULATE_POINTS" not in line
                    and "REDEEM_REWARD" not in line
                    and '"OTHER"' not in line):
                continue
            for text in _result_texts(line):
                if '"events"' not in text:
                    continue
                try:
                    data = json.loads(text)
                except Exception:
                    continue
                for ev in data.get("events", []):
                    if isinstance(ev, dict) and "id" in ev and "location_id" in ev:
                        events[ev["id"]] = ev
    return list(events.values())


def summarize(events, label="found"):
    if not events:
        print(f"{label}: 0 events")
        return
    by_loc = defaultdict(list)
    for e in events:
        by_loc[e["location_id"]].append(e)
    print(f"{label}: {len(events)} distinct events")
    for loc, evs in sorted(by_loc.items(), key=lambda kv: LOCATION_MAP.get(kv[0], kv[0])):
        name = LOCATION_MAP.get(loc, loc)
        ts = [e["created_at"] for e in evs]
        kinds = Counter(e.get("type") for e in evs)
        print(f"  {name:9s} {len(evs):5d} events  "
              f"visits={kinds.get('ACCUMULATE_POINTS', 0):5d} "
              f"redeem={kinds.get('REDEEM_REWARD', 0):4d}  "
              f"{min(ts)} → {max(ts)}")
    missing = [n for lid, n in LOCATION_MAP.items() if lid not in by_loc]
    if missing:
        print(f"  !! no events at all for: {', '.join(sorted(missing))}")


def main():
    args = sys.argv[1:]
    dry = "--dry-run" in args
    def opt(flag):
        return args[args.index(flag) + 1] if flag in args else None

    transcript = find_transcript(opt("--transcript"))
    if not transcript or not os.path.exists(transcript):
        print("ERROR: could not locate a session transcript. "
              "Pass --transcript /path/to/session.jsonl, or fall back to "
              "`loyalty_store.py ingest-dir /tmp/loyalty_pages`.")
        return 1

    print(f"transcript: {transcript}")
    events = extract(transcript)

    since, until = opt("--since"), opt("--until")
    if since or until:
        lo = (since or "0000") + "T00:00:00Z"
        hi = (until or "9999") + "T23:59:59Z"
        before = len(events)
        events = [e for e in events if lo <= e["created_at"] <= hi]
        print(f"window {since or '-'}..{until or '-'}: kept {len(events)} of {before}")

    summarize(events, "extracted")

    also = opt("--also-dir")
    if also and os.path.isdir(also):
        n = loyalty_store.ingest_dir(also)
        print(f"also ingested {n} new events from {also}")

    if dry:
        print("\n--dry-run: nothing written to the store")
        return 0

    new = loyalty_store.ingest(events)
    print(f"\ningested {new} NEW events into the loyalty store")
    print("coverage:", json.dumps(loyalty_store.coverage(), indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
