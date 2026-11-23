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
    """IN-PLACE XML PATCH (Oct 2026, Ryan rule: edit the existing model, never recreate).
    Reads the tab with openpyxl (never saved) to find the rows, then rewrites only the 4 metric
    cells per cafe inside the Ops Metrics sheet XML. Everything else in the file stays byte-identical."""
    import zipfile, re, shutil, tempfile
    from lxml import etree
    NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    Q=lambda t:'{%s}%s'%(NS,t)
    wb = load_workbook(path, read_only=True); OM = wb["Ops Metrics"]; cur=None; wrote=[]
    for r,row in enumerate(OM.iter_rows(min_col=2,max_col=2,values_only=True),1):
        b=row[0]
        if isinstance(b, str) and b.strip() in CAFE_MAP: cur=b.strip(); continue
        if isinstance(b, datetime.datetime) and cur and b.year==year and b.month==month:
            wrote.append((cur,r)+month_wdwe(CAFE_MAP[cur], year, month)); cur=None
    wb.close()
    zin=zipfile.ZipFile(path); members={n:zin.read(n) for n in zin.namelist()}; infos={i.filename:i for i in zin.infolist()}
    wbx=etree.fromstring(members['xl/workbook.xml']); rels=etree.fromstring(members['xl/_rels/workbook.xml.rels'])
    rid={x.get('Id'):x.get('Target') for x in rels}
    sf=[ 'xl/'+rid[s.get('{%s}id'%RNS)].lstrip('/').replace('xl/','') for s in wbx.find(Q('sheets')) if s.get('name')=='Ops Metrics'][0]
    root=etree.fromstring(members[sf]); sd=root.find(Q('sheetData')); rows={int(x.get('r')):x for x in sd}
    order=lambda c:sum((ord(ch)-64)*26**i for i,ch in enumerate(reversed(re.match(r'[A-Z]+',c.get('r')).group(0))))
    for cur,r,*vals in wrote:
        row=rows[r]; prev=rows.get(r-1)
        for col,val in zip('CDEF',vals):
            ref=f'{col}{r}'; c=next((x for x in row if x.get('r')==ref),None)
            if c is None: c=etree.SubElement(row,Q('c')); c.set('r',ref)
            for ch in list(c): c.remove(ch)
            if 't' in c.attrib: del c.attrib['t']
            ps=next((x.get('s') for x in (prev if prev is not None else []) if x.get('r')==f'{col}{r-1}'),None)
            if ps: c.set('s',ps)                      # same number format / green font as the prior month's cell
            etree.SubElement(c,Q('v')).text=repr(val)
        g=next((x for x in row if x.get('r')==f'G{r}'),None)
        if g is not None:
            for ch in list(g): g.remove(ch)
        cs=sorted(row,key=order)
        for c in cs: row.remove(c)
        for c in cs: row.append(c)
    members[sf]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    tmp=path+'.tmp'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
        for n in zin.namelist():
            zi=zipfile.ZipInfo(n,date_time=infos[n].date_time); zi.compress_type=zipfile.ZIP_DEFLATED; zi.external_attr=infos[n].external_attr
            zo.writestr(zi,members[n])
    z2=zipfile.ZipFile(tmp); diff=[n for n in zin.namelist() if z2.read(n)!=zin.read(n)]
    assert diff in ([sf],[]), diff
    zin.close(); z2.close(); os.replace(tmp,path); return wrote

if __name__ == "__main__":
    y,m,model = int(sys.argv[1]), int(sys.argv[2]), sys.argv[3]
    for cur,r,wd,we,wdt,wet in write_model(model,y,m):
        print(f"  {cur:10} row{r}: WD/day={wd} WE/day={we} WD tkt=${wdt} WE tkt=${wet}")
    print("Ops Metrics updated from daily_metrics.csv (lean daily store).")
