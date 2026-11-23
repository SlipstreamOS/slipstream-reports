#!/usr/bin/env python3
"""Extract Jan..<month> 2026 financials from the RECALCULATED model copy (~/rc) + Projections into data/fin_2026_v<N>.json.
Usage: python3 extract_fin.py <model_ver> <YYYY-MM> "<projections recalculated xlsx>"
(recalc the model first: recalc_copy.py <ver>; the projections copy: soffice --convert-to xlsx of the latest Projections file)."""
import openpyxl,warnings,os,sys,json,datetime; warnings.simplefilter('ignore')
V,YM,PROJ=sys.argv[1],sys.argv[2],sys.argv[3]; y,mo=int(YM[:4]),int(YM[5:7])
H=os.path.expanduser("~"); wb=openpyxl.load_workbook(f"{H}/rc/Slipstream Consolidated Model 2026 v{V}.xlsx",data_only=True)
MN=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][:mo]
COL={m:19+i for i,m in enumerate(MN)}            # 2026 Jan = col S on IS + Consolidated tabs
CAFE_ROWS=dict(cbev=25,cret=26,coffee=27,food=29,cat=30,disc=33,ref=34,tcls=35,ccogs=37,fcogs=39,catcogs=41,merch=42,tcogs=43,
               labor=53,texp=88,noi=89,ni=96,wd_txn=20,we_txn=21,wd_tkt=22,we_tkt=23,days=15,weekdays=16,weekend=17)
SUP={'Commissary':dict(rev=25,cogs=30,labor=37,opex=43,ni=45),'SRP':dict(rev=24,cogs=27,labor=34,opex=45,ni=47),
     'Cafe Mgmt':dict(rev=21,ni=22),'Central':dict(labor=24,opex=34,ni=36)}
CONS=dict(coffee=27,food=29,cat=30,tcls=35,tcogs=43,gp=45,labor=53,texp=88,noi=89,cafes_ni=96,comm=100,srp=101,mgmt=102,central=103,total_ni=106)
num=lambda v: float(v) if isinstance(v,(int,float)) else 0.0
out={'cafes':{},'support':{},'consolidated':{},'proj_sales':{},'model':f"v{V}",'months':MN}
ws0=wb['Income Statement - Studio']
for m in MN:
    d=ws0.cell(14,COL[m]).value; assert isinstance(d,datetime.datetime) and d.month==MN.index(m)+1, (m,d)
for c in ['Downtown','Studio','1899','NMWA']:
    ws=wb[f'Income Statement - {c}']; out['cafes'][c]={}
    for m in MN:
        r={k:num(ws.cell(row,COL[m]).value) for k,row in CAFE_ROWS.items()}
        r['core']=r['tcls']-r['cat']; r['cogs']=r['tcogs']-r['catcogs']; r['prime']=r['cogs']+r['labor']
        r['gp']=r['core']-r['cogs']; r['pm']=r['core']-r['prime']; r['disc_ref']=r['disc']+r['ref']
        out['cafes'][c][m]=r
for s,rows in SUP.items():
    ws=wb[f'Income Statement - {s}']; out['support'][s]={m:{k:num(ws.cell(row,COL[m]).value) for k,row in rows.items()} for m in MN}
ws=wb['Consolidated']; out['consolidated']={m:{k:num(ws.cell(row,COL[m]).value) for k,row in CONS.items()} for m in MN}
pw=openpyxl.load_workbook(PROJ,data_only=True)
for c,t in {'Studio':'Studio 2026','Downtown':'Downtown 2026','1899':'1899 2026','NMWA':'NMWA 2026'}.items():
    ws=pw[t]; tot={m:0.0 for m in MN}
    for r in range(6,372):
        d=ws.cell(r,6).value
        if isinstance(d,datetime.datetime) and d.year==2026 and d.month<=mo: tot[MN[d.month-1]]+=num(ws.cell(r,7).value)
    out['proj_sales'][c]={m:round(v,2) for m,v in tot.items()}
p=os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","data",f"fin_2026_v{V}.json"); json.dump(out,open(p,'w'),indent=1)
print("wrote",p); print({c:round(out['cafes'][c][MN[-1]]['ni'],2) for c in out['cafes']}, out['consolidated'][MN[-1]]['total_ni'])
print('proj',{c:out['proj_sales'][c] for c in out['proj_sales']})
