#!/usr/bin/env python3
"""Compute per-location payroll inputs (gross + tips) from a period's Payroll Workbook v2.
Validated June 2026: reproduces Location Totals exactly when populated.
Usage: python3 payroll_inputs.py 2026-06-14_to_2026-06-27 [more periods...]
Output feeds Ryan's Month Close Sheet -> Payroll Detail tab (Location Payroll Report block).
ADP side (Overall/Mgmt gross, ER taxes, medical, CC tips) comes from the ADP PDFs, not here.
"""
import openpyxl, os, sys
BASE=os.path.join(os.path.dirname(__file__),"..","..","..","Payroll")
# canonical keys; match Inputs location strings by substring so Square renames don't zero a location
LOCS=["Slipstream Studio","Slipstream Downtown","Slipstream 1899 Penn","Slipstream NMWA","Commissary","Roasting"]
def canon(loc):
    s=str(loc or "")
    for k in LOCS:
        if k in s: return k
    return None
DAYS={'Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'}
def compute(period):
    p=os.path.join(BASE,period,"Payroll Workbook v2.xlsx")
    if not os.path.exists(p): return None
    wb=openpyxl.load_workbook(p,data_only=True); inp=wb['Inputs']
    lt=wb['Location Totals']
    # prefer populated Location Totals (cached values)
    if any(isinstance(lt.cell(r,5).value,(int,float)) and lt.cell(r,5).value for r in range(4,10)):
        out={}
        for r in range(4,10):
            out[lt.cell(r,1).value]=(lt.cell(r,5).value or 0, lt.cell(r,6).value or 0)
        return out
    cost={l:0.0 for l in LOCS}
    for r in range(41,inp.max_row+1):
        if not inp.cell(r,2).value: continue
        loc=canon(inp.cell(r,5).value)
        if loc is None: continue
        rh=inp.cell(r,12).value or 0; oh=inp.cell(r,13).value or 0; w=inp.cell(r,16).value or 0
        cost[loc]+=rh*w+oh*w*1.5
    tips={3:0.0,4:0.0,5:0.0,6:0.0}
    for r in range(1,24):
        a=inp.cell(r,1).value
        if a and str(a).strip() in DAYS:
            for c in (3,4,5,6):
                v=inp.cell(r,c).value
                if isinstance(v,(int,float)): tips[c]+=v
    tipmap={"Slipstream Studio":tips[3],"Slipstream Downtown":tips[4],"Slipstream 1899 Penn":tips[5],"Slipstream NMWA":tips[6]}
    return {l:(round(cost[l],2),round(tipmap.get(l,0.0),2)) for l in LOCS}
if __name__=="__main__":
    for period in sys.argv[1:]:
        r=compute(period)
        print(f"\n### {period}")
        if not r: print("  workbook missing"); continue
        for l,(g,t) in r.items(): print(f"  {l:26} gross {g:>10.2f}  tips {t:>9.2f}")
