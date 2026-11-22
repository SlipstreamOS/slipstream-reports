#!/usr/bin/env python3
"""Collect this session's Square results for the weekly build (stdlib only).

    python3 run/harvest.py START END

Looks through every tool result this session holds (transcripts and spilled
tool-results files under ~/.claude/projects, same search as
slipstream_cloud_delta.py) and writes:

  _pipeline/weekly/raw/labor_week.json     {"timecards": [...]}, deduped by id,
                                           start_at local date within START..END
  _pipeline/weekly/raw/filter_result.json  {"data": [...]} rows of the ItemSales
                                           reporting load (for build_filter_mix.py)

Prints counts, and exits 2 if either is empty, so the run stops before building
a report from missing data.
"""
import os, sys, json
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, "_pipeline", "shared", "scripts"))
from slipstream_cloud_delta import all_result_texts  # noqa: E402

RAW = os.path.join(REPO, "_pipeline", "weekly", "raw")
CAFES = {"LXM7V1QWDHKZJ": "Downtown", "LGVKCKX9B9DZT": "Studio", "L4ZBMGCWJ5C7R": "1899 Penn", "LV2D18H0DD9ES": "NMWA"}


def main():
    start, end = sys.argv[1], sys.argv[2]
    roots = [os.path.expanduser("~/.claude/projects")]
    timecards, rows, seen_rows = {}, [], set()
    for raw in all_result_texts(roots):
        s = (raw or "").lstrip()
        if not s.startswith("{"):
            continue
        head = s[:300]
        if '"timecards"' not in head and '"ItemSales.' not in s[:3000]:
            continue
        try:
            j = json.loads(s)
        except Exception:
            continue
        for t in j.get("timecards", []) or []:
            d = (t.get("start_at") or "")[:10]
            if t.get("id") and start <= d <= end:
                timecards[t["id"]] = t
        data = j.get("data")
        if data is None and isinstance(j.get("results"), list) and j["results"]:
            data = j["results"][0].get("data")
        for r in data or []:
            if isinstance(r, dict) and "ItemSales.item_name" in r:
                key = json.dumps(r, sort_keys=True)
                if key not in seen_rows:
                    seen_rows.add(key); rows.append(r)
    os.makedirs(RAW, exist_ok=True)
    tc = sorted(timecards.values(), key=lambda t: t.get("start_at", ""))
    json.dump({"timecards": tc}, open(os.path.join(RAW, "labor_week.json"), "w"))
    json.dump({"data": rows}, open(os.path.join(RAW, "filter_result.json"), "w"))
    by = Counter(CAFES.get(t.get("location_id"), t.get("location_id")) for t in tc)
    print(f"timecards {len(tc)} {dict(by)}; open (no end_at): {sum(1 for t in tc if not t.get('end_at'))}")
    print(f"filter rows {len(rows)}")
    # Projections: Drive download_file_content spills the xlsx as a blob file; newest one wins.
    import glob, shutil
    blobs = sorted(glob.glob(os.path.join(roots[0], "**", "tool-results", "mcp-Google_Drive-blob-*.xlsx"), recursive=True),
                   key=os.path.getmtime)
    if blobs:
        dst = os.path.join(REPO, "Projections", "2026 Projections.xlsx")
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(blobs[-1], dst)
        print(f"projections <- {os.path.basename(blobs[-1])}")
    else:
        print("projections: no Drive download found in this session")
    if not tc or not rows:
        sys.exit(2)


if __name__ == "__main__":
    main()
