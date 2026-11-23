#!/usr/bin/env python3
"""Compute Ops Metrics (weekday/weekend avg daily transactions + avg ticket) for a
month from the LEAN DAILY STORE (_pipeline/shared/daily_metrics.csv) and write them
into the Consolidated Model's "Ops Metrics" tab. Daily-resolution, so month-end-
mid-week and weekday/weekend splits are exact. No Square call; reads only the
compact daily store the weekly job maintains. Never skips a cafe.

Usage:  python3 ops_metrics_from_txstore.py <year> <month> "<model .xlsx path>"
"""
import sys, os, csv, datetime
from openpyxl import load_workbook
from openpyxl.styles import Font

STORE = os.path.join(os.path.dirname(__file__), "..", "..", "shared", "daily_metrics.csv")
# model cafe name -> daily_metrics cafe name
CAFE_MAP = {"Studio": "Studio", "Downtown": "Downtown", "1899": "1899 Penn", "NMWA": "NMWA"}
GREEN = Font(color="2E7D32")

def month_wdwe(store_cafe, year, month):
    """WD/WE avg daily txns + avg ticket from the FULL monthly transactions CSV.
    EXCLUDES invoice-sourced transactions (billed catering/events — projected via the
    Catering row, and they massively skew ticket size). Ryan rule, Aug 2026."""
    import calendar
    TXDIR = os.path.join(os.path.dirname(__file__), "..", "..", "shared", "transactions")
    path = os.path.join(TXDIR, f"{year}-{month:02d}.csv")
    FULLNAME = {"Studio":"Slipstream Studio","Downtown":"Slipstream Downtown",
                "1899 Penn":"Slipstream 1899 Penn","NMWA":"Slipstream NMWA"}
    loc = FULLNAME.get(store_cafe, store_cafe)
    def money(x):
        x=str(x).replace("$","").replace(",","").strip()
        if x in ("","-"): return 0.0
        neg=x.startswith("-"); x=x.lstrip("-")
        try: v=float(x)
        except: return 0.0
        return -v if neg else v
    wd_days={}; we_days={}; wd_net=we_net=0.0
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r.get("Location")!=loc: continue
            if r.get("Event Type")!="Payment": continue
            if "Invoice" in (r.get("Source") or ""): continue
            d=datetime.date.fromisoformat(r["Date"]); net=money(r["Net Sales"])
            bucket = we_days if d.weekday()>=5 else wd_days
            bucket[d]=bucket.get(d,0)+1
            if d.weekday()>=5: we_net+=net
            else: wd_net+=net
    wd_tx=sum(wd_days.values()); we_tx=sum(we_days.values())
    wd_avg=round(wd_tx/len(wd_days),1) if wd_days else 0
    we_avg=round(we_tx/len(we_days),1) if we_days else 0
    wd_tkt=round(wd_net/wd_tx,2) if wd_tx else 0
    we_tkt=round(we_net/we_tx,2) if we_tx else 0
    return wd_avg, we_avg, wd_tkt, we_tkt

def write_model(path, year, month):
    wb = load_workbook(path); OM = wb["Ops Metrics"]; cur=None; wrote=[]
    for r in range(1, OM.max_row+1):
        b = OM.cell(row=r, column=2).value
        if isinstance(b, str) and b.strip() in CAFE_MAP: cur=b.strip(); continue
        if isinstance(b, datetime.datetime) and cur and b.year==year and b.month==month:
            vals = month_wdwe(CAFE_MAP[cur], year, month)
            for col,val in zip((3,4,5,6), vals):
                c=OM.cell(row=r, column=col, value=val); c.font=GREEN
            OM.cell(row=r, column=7, value=None)
            wrote.append((cur,r)+vals); cur=None
    wb.save(path); return wrote

if __name__ == "__main__":
    y,m,model = int(sys.argv[1]), int(sys.argv[2]), sys.argv[3]
    for cur,r,wd,we,wdt,wet in write_model(model,y,m):
        print(f"  {cur:10} row{r}: WD/day={wd} WE/day={we} WD tkt=${wdt} WE tkt=${wet}")
    print("Ops Metrics updated from daily_metrics.csv (lean daily store).")
