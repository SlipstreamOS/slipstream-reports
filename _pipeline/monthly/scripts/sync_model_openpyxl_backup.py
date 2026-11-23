#!/usr/bin/env python3
"""Sync Consolidated Model from Coefficient Output + bump last-closed-month. June-2026-validated.
Usage: python3 sync_model.py <in_version> <out_version> <YYYY-MM>  e.g. python3 sync_model.py 10 11 2026-07
- FULL-syncs all 8 'QBO - X' tabs (rows 1-320, cols A-I + col P acct# helper) from 'Coefficient Output 2026.xlsx'
- Bumps Assumptions!C5 to EOM of <YYYY-MM>; sets fullCalcOnLoad
- Saves as new version; NEVER overwrites the input. Then recalc: soffice --headless --convert-to xlsx
PRE-REQ: Ryan refreshed Coefficient Output IN EXCEL first (account numbers must be present in col A).
NEVER openpyxl-save the Coefficient Output file itself (kills the add-in connection).
"""
import openpyxl, re, sys, datetime, calendar, os
V_IN,V_OUT,YM=sys.argv[1],sys.argv[2],sys.argv[3]
y,mo=int(YM[:4]),int(YM[5:7]); eom=datetime.datetime(y,mo,calendar.monthrange(y,mo)[1])
D=os.path.join(os.path.dirname(__file__),"..","..","..")
BASE=os.path.join(D,f"Slipstream Consolidated Model 2026 v{V_IN}.xlsx")
CO=os.path.join(D,"Coefficient Output 2026.xlsx")
OUT=os.path.join(D,f"Slipstream Consolidated Model 2026 v{V_OUT}.xlsx")
TABMAP={'QBO - Studio':'QBO - Studio','QBO - Downtown':'QBO - DT','QBO - 1899':'QBO - 1899',
 'QBO - NMWA':'QBO - NMWA','QBO - Commissary':'QBO - Commissary','QBO - SRP':'QBO - SRP',
 'QBO - Cafe Mgmt':'QBO - Café Mgmt','QBO - Central':'QBO - Central Admin'}
m=openpyxl.load_workbook(BASE,data_only=False); co=openpyxl.load_workbook(CO,data_only=True)
lead=lambda s:(re.match(r'\s*(\d{5})\b',str(s or '')) or [None]) and (re.match(r'\s*(\d{5})\b',str(s or '')).group(1) if re.match(r'\s*(\d{5})\b',str(s or '')) else None)
for mt,ot in TABMAP.items():
    mws,ows=m[mt],co[ot]
    for r in range(1,321):
        for c in range(1,17): mws.cell(r,c).value=None
    nums=0
    for r in range(1,ows.max_row+1):
        for c in range(1,10): mws.cell(r,c).value=ows.cell(r,c).value
        if r>=6:
            n=lead(ows.cell(r,1).value)
            if n: mws.cell(r,16).value=int(n); nums+=1
    mws.cell(5,16).value='acct# (helper)'
    assert nums>3, f"{ot}: no account numbers found — Ryan must re-add numbers in Coefficient settings"
    print(f"  {mt}: synced, {nums} numbered accounts")
m['Assumptions']['C5'].value=eom
try: m.calculation.fullCalcOnLoad=True
except Exception:
    from openpyxl.workbook.properties import CalcProperties; m.calculation=CalcProperties(fullCalcOnLoad=True)
m.save(OUT); print("Saved",OUT,"| C5 ->",eom.date())
print("NEXT: recalc + validate:  soffice --headless --calc --convert-to xlsx --outdir /tmp/rc '"+OUT+"'  then python3 validate_model.py <ver> <YYYY-MM> /tmp/rc")
