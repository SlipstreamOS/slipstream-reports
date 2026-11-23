#!/usr/bin/env python3
"""Hot filter cups by cafe and Sun-Sat week, from a Square Reporting `load` result, merged into ../data/filter_mix_weekly.json.

    python3 build_filter_mix.py <reporting_result.json>

The result must come from this query (America/New_York, granularity day, any date range covering whole weeks):
  measures ["ItemSales.net_quantity"], dimensions ["ItemSales.location_id","ItemSales.item_name"],
  timeDimensions [{"dimension":"ItemSales.local_reporting_timestamp","dateRange":[START,END],"granularity":"day"}],
  segments ["ItemSales.non_gift_cards"], filters [{"or":[{"member":"ItemSales.item_name","operator":"contains",
  "values":["Cruise Control","Daily Selection","[Bright]","[Fruit]","[Floral]","[Decaf]","[BRIGHT]","[FRUIT]","[FLORAL]","[DECAF]"]}]}]
Classes: cc = Cruise Control (house), ds = Daily Selection, so = single origin ([Bright]/[Fruit]/[Floral]/[Decaf]).
Left out: Nitro (iced) and retail (whole bean, bags, 'RETAIL', 5#). Weeks present in the input replace the stored ones.
"""
import sys, os, json, re, datetime as dt, collections
LOC = {"LXM7V1QWDHKZJ": "Downtown", "LGVKCKX9B9DZT": "Studio", "L4ZBMGCWJ5C7R": "1899 Penn", "LV2D18H0DD9ES": "NMWA"}
def cls(n):
    n = n.lower()
    if any(k in n for k in ("whole bean", "nitro", "bag", "retail", "5#")): return None
    if "cruise control" in n: return "cc"
    if "daily selection" in n: return "ds"
    if re.match(r"\[(bright|fruit|floral|decaf)\]", n): return "so"
j = json.load(open(sys.argv[1]))
data = j.get("data") or j["results"][0]["data"]
out = collections.defaultdict(lambda: collections.defaultdict(lambda: {"cc": 0.0, "ds": 0.0, "so": 0.0}))
for r in data:
    if r.get("ItemSales.location_id") not in LOC: continue
    day = r.get("ItemSales.local_reporting_timestamp.day") or r.get("ItemSales.local_reporting_timestamp")
    d = dt.date.fromisoformat(day[:10]); wk = (d - dt.timedelta((d.weekday() + 1) % 7)).isoformat()
    c = cls(r["ItemSales.item_name"])
    if c: out[wk][LOC[r["ItemSales.location_id"]]][c] += float(r.get("ItemSales.net_quantity") or 0)
p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "filter_mix_weekly.json")
store = json.load(open(p)) if os.path.exists(p) else {}
for w, cafes in out.items():
    store[w] = {c: {k: int(round(x)) for k, x in v.items()} for c, v in cafes.items()}
json.dump(dict(sorted(store.items())), open(p, "w"), indent=1)
for w in sorted(out):
    print(w, {c: f"{(v['ds'] + v['so']) / max(1, sum(v.values())):.0%} of {int(sum(v.values()))}" for c, v in out[w].items()})
