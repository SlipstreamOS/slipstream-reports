import json, glob, datetime, collections, os
BASE = "/sessions/festive-kind-thompson/mnt/outputs/cache/periods/2026-05-03_2026-05-16"
WORK = os.path.join(BASE, "_raw")
LOC = {
  "LXM7V1QWDHKZJ": "Slipstream Downtown",
  "LGVKCKX9B9DZT": "Slipstream Studio",
  "L4ZBMGCWJ5C7R": "Slipstream 1899 Penn",
  "LV2D18H0DD9ES": "Slipstream NMWA",
  "LZCA3EY4GB7BM": "CR Commissary (Food)",
}
def et_date(rfc):
    dt = datetime.datetime.strptime(rfc[:19], "%Y-%m-%dT%H:%M:%S")
    et = dt - datetime.timedelta(hours=4)  # May = EDT, UTC-4
    return f"{et.month}/{et.day}/{str(et.year)[2:]}"
agg = collections.defaultdict(lambda: collections.defaultdict(int))
total_processed = 0
completed_with_tip = 0
def consume(payments):
    global total_processed, completed_with_tip
    for p in payments:
        total_processed += 1
        if p.get('status') != 'COMPLETED':
            continue
        tip = (p.get('tip_money') or {}).get('amount', 0) or 0
        if tip > 0:
            completed_with_tip += 1
            lid = p['location_id']
            agg[LOC[lid]][et_date(p['created_at'])] += tip
for lid in ["LXM7V1QWDHKZJ","LGVKCKX9B9DZT","L4ZBMGCWJ5C7R"]:
    files = sorted(glob.glob(f"{WORK}/{lid}_p*.json"),
                   key=lambda x:int(x.split('_p')[-1].split('.')[0]))
    for f in files:
        consume(json.load(open(f)).get('payments',[]))
for extra in ["LV2D18H0DD9ES_inline.json","LZCA3EY4GB7BM_inline.json"]:
    consume(json.load(open(os.path.join(WORK,extra))).get('payments',[]))
def datekey(s):
    m,d,y = s.split('/'); return (int(y),int(m),int(d))
out = {}
for loc in ["Slipstream Downtown","Slipstream Studio","Slipstream 1899 Penn","Slipstream NMWA","CR Commissary (Food)"]:
    days = agg.get(loc,{})
    out[loc] = {dt: round(days[dt]/100,2) for dt in sorted(days,key=datekey)}
with open(os.path.join(BASE,"cc_tips.json"),"w") as fh:
    json.dump(out,fh,indent=2)
print("Total records processed:",total_processed)
print("COMPLETED w/ tip>0:",completed_with_tip)
print()
for loc,days in out.items():
    print(f"{loc}: ${round(sum(days.values()),2):,.2f}  ({len(days)} days)")
    for dt,v in days.items():
        print(f"   {dt}: ${v:,.2f}")
