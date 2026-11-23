#!/usr/bin/env python3
"""Aug 2026 close: add missing pull rows to the Consolidated Model IS tabs (v15 -> v16).
Cafe tabs: 60130 Shared Mgmt Labor (admin), 60145 Training Labor (both into Total Labor),
82100 Licenses & Fees (into Total Other OpEx), 50100 parent Coffee COGS (into Total COGS).
SRP: 62140. Central: 62140/62150/62151/82100. Consolidated: mirror rows. Labor drivers -> final targets.
Usage: python3 add_model_rows_v16.py <in_ver> <out_ver>"""
import openpyxl, re, sys, copy, os
from openpyxl.styles import Font
VI,VO=sys.argv[1],sys.argv[2]
D=os.path.join(os.path.dirname(__file__),"..","..","..")
wb=openpyxl.load_workbook(os.path.join(D,f"Slipstream Consolidated Model 2026 v{VI}.xlsx"))
L=openpyxl.utils.get_column_letter
BLUE=Font(color="FF0070C0"); GREEN=Font(color="FF2E7D32")
def rowsub(f,t,n): return re.sub(r'(\$?[A-Z]{1,3}\$?)'+str(t)+r'(?![0-9])', lambda m:m.group(1)+str(n), f) if isinstance(f,str) and f.startswith('=') else None
def clone_row(ws,t,n,label,acct,drv=None,eval_=None,sheet_from=None,sheet_to=None):
    for c in range(1,ws.max_column+1):
        src=ws.cell(t,c); dst=ws.cell(n,c)
        f=rowsub(src.value,t,n)
        if f and sheet_from: f=f.replace(sheet_from,sheet_to)
        dst.value=f if f else (None if c>5 else None)
        dst.number_format=src.number_format; dst.font=copy.copy(src.font); dst.alignment=copy.copy(src.alignment)
    ws.cell(n,2).value=label; ws.cell(n,3).value=acct
    ws.cell(n,4).value=drv; ws.cell(n,5).value=eval_
    if drv is not None: ws.cell(n,4).font=BLUE
    if eval_ is not None: ws.cell(n,5).font=BLUE
def add_to_total(ws,total_row,anchor_row,new_rows):
    for c in range(1,ws.max_column+1):
        f=ws.cell(total_row,c).value; col=L(c)
        if isinstance(f,str) and f.startswith('=') and re.search(r'(?<![A-Z$])'+col+str(anchor_row)+r'(?![0-9])',f):
            ws.cell(total_row,c).value=f+''.join(f'+{col}{r}' for r in new_rows)
TARGET={'Studio':(0.30,0.03,0.01),'Downtown':(0.32,0.03,0.01),'1899':(0.35,0.02,0.01),'NMWA':(None,0.0,0.0)}
QT={'Studio':'QBO - Studio','Downtown':'QBO - Downtown','1899':'QBO - 1899','NMWA':'QBO - NMWA'}
newrows={}
t1899=wb['Income Statement - 1899']; EXP_T=99   # 1899 row 99 = Printing and postage (fixed $ w/ inflation forecast)
for t in ['Studio','Downtown','1899','NMWA']:
    ws=wb[f'Income Statement - {t}']; n=ws.max_row+1
    core,adm,trn=TARGET[t]
    if core is not None: ws['D48'].value=core
    clone_row(ws,48,n,  '   Shared Mgmt Labor (admin, 60130)',60130,drv=adm)
    clone_row(ws,48,n+1,'   Training Labor (60145)',60145,drv=trn)
    # licenses: copy 1899 row 99 pattern into this tab (fixed $ + inflation)
    for c in range(1,ws.max_column+1):
        src=t1899.cell(EXP_T,c); dst=ws.cell(n+2,c)
        f=rowsub(src.value,EXP_T,n+2)
        if f: f=f.replace("'QBO - 1899'",f"'{QT[t]}'")
        dst.value=f; dst.number_format=src.number_format; dst.font=copy.copy(src.font)
    ws.cell(n+2,2).value='   Licenses & Fees (82100)'; ws.cell(n+2,3).value=82100; ws.cell(n+2,4).value=None; ws.cell(n+2,5).value=0; ws.cell(n+2,5).font=BLUE
    clone_row(ws,48,n+3,'   Coffee COGS adj (50100 parent — rent-JE concentrate)',50100,drv=0.0)
    add_to_total(ws,53,48,[n,n+1]); add_to_total(ws,80,71,[n+2]); add_to_total(ws,43,37,[n+3])
    newrows[t]=(n,n+1,n+2,n+3); print(t,'rows',newrows[t],'D48->',ws['D48'].value)
# Consolidated mirror rows
c=wb['Consolidated']; base=c.max_row+1
labels=['   Shared Mgmt Labor (admin, 60130)','   Training Labor (60145)','   Licenses & Fees (82100)','   Coffee COGS adj (50100 parent)']
for i,lab in enumerate(labels):
    r=base+i; c.cell(r,2).value=lab
    for col in range(1,c.max_column+1):
        f=c.cell(25,col).value
        if isinstance(f,str) and f.startswith('=$D$5'):
            c.cell(r,col).value=(f"=$D$5*'Income Statement - Studio'!{L(col)}{newrows['Studio'][i]}+$D$6*'Income Statement - Downtown'!{L(col)}{newrows['Downtown'][i]}"
                                 f"+$D$7*'Income Statement - 1899'!{L(col)}{newrows['1899'][i]}+$D$8*'Income Statement - NMWA'!{L(col)}{newrows['NMWA'][i]}")
            c.cell(r,col).number_format=c.cell(25,col).number_format
        elif isinstance(f,str) and f.startswith('=SUM('):
            c.cell(r,col).value=rowsub(f,25,r); c.cell(r,col).number_format=c.cell(25,col).number_format
print('Consolidated rows',base,'-',base+3)
# SRP
s=wb['Income Statement - SRP']; n=s.max_row+1
clone_row(s,38,n,'   62140 Marketing Expense',62140); s.cell(n,4).value=None; s.cell(n,5).value=s.cell(38,5).value
add_to_total(s,45,37,[n]); print('SRP row',n)
# Central
ce=wb['Income Statement - Central']; n=ce.max_row+1; add=[]
for i,(lab,a) in enumerate([('   62140 Marketing Expense',62140),('   62150 Legal & Professional Fee',62150),('   62151 Accounting and Bookkeeping',62151),('   82100 Licenses and Fees',82100)]):
    clone_row(ce,39,n+i,lab,a); ce.cell(n+i,4).value=None; ce.cell(n+i,5).value=ce.cell(39,5).value; add.append(n+i)
add_to_total(ce,34,27,add); print('Central rows',add)
try: wb.calculation.fullCalcOnLoad=True
except Exception:
    from openpyxl.workbook.properties import CalcProperties; wb.calculation=CalcProperties(fullCalcOnLoad=True)
out=os.path.join(D,f"Slipstream Consolidated Model 2026 v{VO}.xlsx"); wb.save(out); print('saved',out)
