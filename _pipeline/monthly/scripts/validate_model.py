#!/usr/bin/env python3
"""Read back a recalc'd model: per-location NI for a month + error scan + key checks.
Usage: python3 validate_model.py <version> <YYYY-MM> [recalc_dir=/tmp/rc]
Checks: per-loc NI, 1899 NI == 0, NMWA/CafeMgmt fee lines, any #ERR cells anywhere.
"""
import openpyxl, sys, os
V,YM=sys.argv[1],sys.argv[2]; RD=sys.argv[3] if len(sys.argv)>3 else "/tmp/rc"
y,mo=int(YM[:4]),int(YM[5:7])
p=os.path.join(RD,f"Slipstream Consolidated Model 2026 v{V}.xlsx")
wb=openpyxl.load_workbook(p,data_only=True)
def mcol(ws):
    best=max(range(1,16),key=lambda r:sum(1 for c in range(1,ws.max_column+1) if hasattr(ws.cell(r,c).value,'year')))
    for c in range(1,ws.max_column+1):
        v=ws.cell(best,c).value
        if hasattr(v,'year') and v.year==y and v.month==mo: return c
def nirow(ws):
    for r in range(ws.max_row,0,-1):
        if str(ws.cell(r,2).value or '').strip().lower()=='net income': return r
print(f"=== Per-location {YM} NI (v{V}) ===")
vals={}
for t in ['Studio','Downtown','1899','NMWA','Commissary','SRP','Cafe Mgmt','Central']:
    ws=wb[f'Income Statement - {t}']; c=mcol(ws); r=nirow(ws)
    vals[t]=ws.cell(r,c).value; print(f"  {t:11} {vals[t]}")
c=wb['Consolidated']; jc=mcol(c)
print(f"Consolidated: CAFES NI={c.cell(96,jc).value}  TOTAL NI={c.cell(106,jc).value}")
assert abs(vals['1899'] or 0)<0.01, "1899 NI should be 0 post-profit-pass!"
errs=0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value,str) and cell.value in ('#REF!','#VALUE!','#DIV/0!','#NAME?','#N/A','#NUM!','#NULL!'):
                errs+=1; print(f"  ERR {ws.title}!{cell.coordinate}={cell.value}")
print("error cells:",errs)
print("MANUAL CHECKS: NMWA fee line r98 = 5041?  CafeMgmt r19=5000 r20=incentive?  Spot-check avg ticket/txns.")
