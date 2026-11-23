#!/usr/bin/env python3
"""Labor Deep Dive step 3 — harvest Square `labor.searchTimecards` responses
straight out of the session transcript (same trick as loyalty_from_transcript.py:
the harness writes every tool result verbatim to the JSONL, so nothing is re-typed).
Dedupes on timecard id. Output: data/labor_deep_dive/timecards_raw.json
"""
import os, sys, json, glob
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.dirname(os.path.dirname(HERE))
WORKSPACE = os.path.dirname(PIPE)
MOUNT = os.path.dirname(WORKSPACE)
OUT = os.path.join(os.path.dirname(HERE), "data", "labor_deep_dive", "timecards_raw.json")

def find_transcripts():
    roots = [os.path.join(MOUNT, ".claude", "projects"), os.path.expanduser("~/.claude/projects")]
    files = []
    for root in roots:
        files += glob.glob(os.path.join(root, "**", "*.jsonl"), recursive=True)
    return sorted(set(files), key=os.path.getmtime)

def result_texts(line):
    try:
        obj = json.loads(line)
    except Exception:
        return
    for c in (obj.get("message", {}).get("content") or []):
        if isinstance(c, dict) and c.get("type") == "tool_result":
            content = c.get("content")
            if isinstance(content, list):
                yield "".join(t.get("text", "") for t in content if isinstance(t, dict))
            elif isinstance(content, str):
                yield content

def main():
    cards = {}
    if os.path.exists(OUT):
        for tc in json.load(open(OUT)).get("timecards", []):
            cards[tc["id"]] = tc
    n_before = len(cards)
    for f in find_transcripts():
        with open(f, errors="replace") as fh:
            for line in fh:
                if 'timecards' not in line:
                    continue
                for text in result_texts(line):
                    if '"timecards"' not in text:
                        continue
                    try:
                        data = json.loads(text)
                    except Exception:
                        continue
                    for tc in data.get("timecards", []):
                        if isinstance(tc, dict) and "id" in tc and "start_at" in tc:
                            cards[tc["id"]] = tc
    # also fold in the pipeline's stored raw week (Aug 23–29) for completeness
    raw_week = os.path.join(PIPE, "weekly", "raw", "labor_week.json")
    if os.path.exists(raw_week):
        for tc in json.load(open(raw_week)).get("timecards", []):
            cards.setdefault(tc["id"], tc)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump({"timecards": list(cards.values())}, open(OUT, "w"))
    print(f"timecards: {n_before} existing -> {len(cards)} total")
    by_day = defaultdict(int)
    for tc in cards.values():
        by_day[tc["start_at"][:10]] += 1
    days = sorted(by_day)
    print(f"date range {days[0]} .. {days[-1]}; {len(days)} distinct days")
    # weeks covered
    import datetime as dt
    wk = Counter()
    for d in days:
        D = dt.date.fromisoformat(d); wk[(D - dt.timedelta(days=D.weekday())).isoformat()] += by_day[d]
    for w, n in sorted(wk.items()):
        print(f"  week of {w}: {n} timecards")

if __name__ == "__main__":
    main()
