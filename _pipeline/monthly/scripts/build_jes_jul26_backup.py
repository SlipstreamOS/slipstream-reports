#!/usr/bin/env python3
"""Slipstream month-close JE builder. CONFIG block below is the only thing to edit each month.
Validated June+July 2026. Reads Square sales CSVs from uploads/, writes per-JE xlsx + combined
import CSV (QBO format: UTF-8 BOM, CRLF, date/USD on first JE line, zero rows filtered)."""
import csv, openpyxl, os, sys
from openpyxl.styles import Font

### ================= CONFIG (JULY 2026) =================
MON="Jul26"; MONNAME="July"; EOM="7/31/26"; REV_DATE="8/1/26"
UP="uploads/"
OUT=f"Weekly and Monthly Reports/{MONNAME} Sales JEs/"
CSVDIR=f"Weekly and Monthly Reports/{MONNAME} Sales JEs CSV/"
SALES={  # loc tag -> (csv, opts)
 'GHC5 - Studio':   (f"Sales Report {MONNAME} Studio.csv", dict(oc_coffee=12.00)),
 'GHC6 - 1899':     (f"Sales Report {MONNAME} 1899.csv",   dict(oc_food=47.00)),
 'NMWA':            (f"Sales Report {MONNAME} NMWA.csv",   dict(oc_coffee=186.70, oc_food=195.30, oc_extra=5841.00,
                      other_cash=6786.00, other_cash_desc='Museum ACH (fees + invoices)',
                      extra=[('Sales:Management Income:Management Fees',0,5041.00,'Mgmt + Equip Fee (NMWA)','NMWA'),
                             ('Sales:Management Income:Management Fees',0,800.00,'Cafe closure payment','NMWA')])),
 'GHC3 - Downtown': (f"Sales Report {MONNAME} DT.csv",     dict(oc_coffee=103.00, oc_food=24.00, oc_extra=10000.00,
                      svc_to_merchant=200.00, other_check=731.50,
                      extra=[('Management Income Clearing',0,10000.00,'Mgmt Fee (1899) 2x5000','Cafe Management'),
                             ('Other Operating Expenses:Merchant Fees',0,200.00,'Mgmt Fee CC offset (1899)','GHC3 - Downtown')])),
}
TRANSFERS=dict(comm_sales_food=9332.39, comm_sales_coffee=842.31,
 comm_purch={'GHC3 - Downtown':(3284.93,368.34),'GHC5 - Studio':(4171.63,318.57),'GHC6 - 1899':(1462.33,155.40),'NMWA':(413.50,0.0)},
 srp_roasted=6080.50, srp_concentrate=5745.00,
 srp_purch={'GHC3 - Downtown':4099.00,'GHC5 - Studio':3878.00,'GHC6 - 1899':3175.50,'NMWA':673.00})
ADJ=dict(tech={'GHC3 - Downtown':-700.00,'GHC5 - Studio':450.00,'GHC6 - 1899':250.00},
 insr={'GHC3 - Downtown':-650.00,'GHC5 - Studio':325.00,'GHC6 - 1899':325.00},
 linen_dt_to_1899=150.00, catering_dt_to_1899=52.78)
PAYROLL=[ # (journal suffix MMDD, date, dict)
 ('0703','7/3/26', dict(L3=4376.31,L5=6129.75,L6=2725.85,L7=1001.89,Comm=1730.91,CommDrv=300,SRP=844.63,SRPDrv=180,Mgmt=3657.43,
   Tips=6027.69,TL3=785.56,TL5=1092.04,TL6=433.55,TL7=120.23,TComm=120,TSRP=101.36,MgmtRes=-337.05,HC=418.97,C1899=4046.49)),
 ('0717','7/17/26',dict(L3=4412.26,L5=6187.64,L6=2837.53,L7=1361.28,Comm=2101.01,CommDrv=300,SRP=873.25,SRPDrv=180,Mgmt=4453.60,
   Tips=5389.54,TL3=762.05,TL5=1033.53,TL6=439.81,TL7=187.20,TComm=120,TSRP=104.79,MgmtRes=-241.42,HC=418.97,C1899=4104.92)),
 ('0731','7/31/26',dict(L3=4341.70,L5=6026.58,L6=2745.09,L7=662.58,Comm=2098.69,CommDrv=300,SRP=893.79,SRPDrv=180,Mgmt=4347.14,
   Tips=6154.40,TL3=817.20,TL5=1032.81,TL6=438.69,TL7=102.94,TComm=120,TSRP=107.25,MgmtRes=-225.11,HC=0.0,C1899=4094.47)),
]
ACCR=dict(L3=1860.73,L5=2582.82,L6=1176.47,L7=283.96,Comm=899.44,CommDrv=128.57,SRP=383.05,SRPDrv=77.14,Mgmt=1863.06,
 TL3=350.23,TL5=442.63,TL6=188.01,TL7=44.12,TComm=51.43,TSRP=45.97,MgmtRes=-96.47)
RECUR=dict(mgmt_fee=5000.00, srp_green_coffee=2200.17, dt_rent=2112.09, srp_rent=1900.00,
 st_rent=dict(core=5396.85, linen=191.60, repairs=0.0, total_due=5483.36, general_supplies=105.09, freeze_dried=0.0))
### ================= END CONFIG =================

def money(s):
    s=str(s).replace('$','').replace(',','').strip()
    if s in ('','-'): return 0.0
    neg=s.startswith('-') or (s.startswith('(') and s.endswith(')'))
    s=s.strip('()').lstrip('-')
    try: v=float(s)
    except: return 0.0
    return -v if neg else v
def parse_sales(path):
    rows=list(csv.reader(open(path,newline='',encoding='utf-8-sig')))
    def sv(label):
        for r in rows:
            if r and r[0].strip()==label: return money(r[1]) if len(r)>1 else 0.0
        return 0.0
    cats={}; i=0
    while i<len(rows):
        if rows[i] and rows[i][0].strip()=='Category Rollup':
            j=i+1
            while j<len(rows) and rows[j] and rows[j][0].strip() and rows[j][0].strip()!='Category Rollup':
                r=rows[j]; nm=r[0].strip()
                if nm in ('Catering and Events','Coffee','Food','Retail','Other Categories','Bar'):
                    cats[nm]={'gross':money(r[2]),'refund':money(r[4])}
                j+=1
            break
        i+=1
    return sv,cats
def sales_je(loc, path, o):
    sv,cats=parse_sales(path)
    g=lambda c:cats.get(c,{}).get('gross',0.0); ref=lambda c:cats.get(c,{}).get('refund',0.0)
    oc=g('Other Categories'); oc_c=o.get('oc_coffee',0.0); oc_f=o.get('oc_food',0.0)
    extra=o.get('extra',[]); oc_extra=o.get('oc_extra',0.0)
    assert abs(oc-(oc_c+oc_f+oc_extra))<0.01, f"{loc}: OtherCat {oc} != mapped {oc_c+oc_f+oc_extra}"
    svc=sv('Service Charges')-o.get('svc_to_merchant',0.0)
    L=[]; A=L.append
    A(['Sales:Core Location Sales:Food Sales',0,g('Food')+oc_f,'Food','',loc])
    A(['Sales:Core Location Sales:Catering and Event Sales',0,g('Catering and Events')+svc,'Catering and Other','',loc])
    A(['Sales:Core Location Sales:Coffee Sales:Retail Coffee Sales',0,g('Retail'),'Retail','',loc])
    A(['Sales:Core Location Sales:Coffee Sales:Coffee Beverage Sales',0,g('Coffee')+oc_c,'Coffee','',loc])
    for acct,d,c,desc,l in extra: A([acct,d,c,desc,'',l])
    A(['Sales Tax Payable',0,sv('Tax'),'Sales Tax','',loc])
    A(['Employee Tips Payable',0,sv('Tip'),'Tips','',loc])
    A(['Gift Card Liability',0,sv('Gift Card Sales'),'gift cards sold','',loc])
    A(['Sales:Core Location Sales:Discounts',abs(sv('Discounts & Comps')),0,'Discounts','',loc])
    A(['Sales:Core Location Sales:Coffee Sales:Coffee Beverage Sales',abs(ref('Coffee'))+abs(ref('Other Categories')),0,'Returns - Coffee','',loc])
    A(['Sales:Core Location Sales:Coffee Sales:Retail Coffee Sales',abs(ref('Retail')),0,'Returns - Retail','',loc])
    A(['Sales:Core Location Sales:Food Sales',abs(ref('Food')),0,'Returns - Food','',loc])
    A(['Sales:Core Location Sales:Catering and Event Sales',abs(ref('Catering and Events')),0,'Returns - Catering and events','',loc])
    A(['Other Operating Expenses:Merchant Fees',abs(sv('Fees')),0,'merchant fees','',loc])
    A(['Gift Card Liability',sv('Gift Card'),0,'gift card liability claimed','',loc])
    A(['GHC3 Cash on hand',sv('Cash'),0,'Cash','',loc])
    if o.get('other_check'): A(['GHC3 Cash on hand',o['other_check'],0,'check','',loc])
    if o.get('other_cash'): A(['GHC3 Cash on hand',o['other_cash'],0,o.get('other_cash_desc','ACH'),'',loc])
    oe=sv('Other')-o.get('other_check',0.0)-o.get('other_cash',0.0)
    A(['GHC3 Credit Card Receivables',max(oe,0),0,'Other','',loc])
    A(['Sales:Core Location Sales:Refunds and Voids',abs(sv('Refunds by Amount')),0,'Refunds','',loc])
    cr=sum(l[2] for l in L); dr=sum(l[1] for l in L)
    L.append(['GHC3 Credit Card Receivables',round(cr-dr,2),0,'CC Receivables','',loc])
    return L
FOODP='Cost of Goods Sold:Core Location COGS:Location Food COGS:Food Purchases'
COFP='Cost of Goods Sold:Core Location COGS:Location Coffee COGS:Coffee Purchases'
APACCT={'GHC3 - Downtown':'Inter-company AP - GHC3','GHC5 - Studio':'Inter-company AP - GHC5','GHC6 - 1899':'Inter-company AP - GHC6','NMWA':'Inter-company AP - NMWA'}
def transfers_je(t):
    L=[]
    L.append(['Sales:Commissary Sales:Commissary Food Sales',0,t['comm_sales_food'],'Commissary Food Sales','','Commissary'])
    L.append(['Sales:Commissary Sales:Commissary Coffee Sales',0,t['comm_sales_coffee'],'Commissary Coffee Sales','','Commissary'])
    comm_ar=0.0
    for loc,(f,c) in t['comm_purch'].items():
        tag=loc.split(' - ')[0].replace('GHC3','L3').replace('GHC5','L5').replace('GHC6','L6') if loc!='NMWA' else 'L7'
        if f: L.append([FOODP,f,0,f'{tag} Food Purchases from Com','',loc])
        if c: L.append([COFP,c,0,f'{tag} Coffee Purchases from Com','',loc])
        comm_ar+=f+c
    L.append(['Sales:Roasting Sales:Roasted Coffee Sales',0,t['srp_roasted'],'SRP Sales - Roasted Coffee','','SRP'])
    L.append(['Sales:Roasting Sales:Concentrate Sales',0,t['srp_concentrate'],'SRP Sales - Concentrate','','SRP'])
    srp_ar=0.0
    for loc,v in t['srp_purch'].items():
        tag='L3' if 'GHC3' in loc else 'L5' if 'GHC5' in loc else 'L6' if 'GHC6' in loc else 'L7'
        if v: L.append([COFP,v,0,f'{tag} SRP Purchases','',loc]); srp_ar+=v
    L.append(['Inter-company AR - COMM',round(comm_ar,2),0,'AR/AP Adj','','Commissary'])
    for loc,(f,c) in t['comm_purch'].items():
        if f+c: L.append([APACCT[loc],0,round(f+c,2),'AR/AP Adj','',loc])
    L.append(['Inter-company AR - SRP',round(srp_ar,2),0,'AR/AP Adj','','SRP'])
    for loc,v in t['srp_purch'].items():
        if v: L.append([APACCT[loc],0,v,'AR/AP Adj','',loc])
    return L
def adjustment_je(a):
    L=[]
    for acct,d in [('Other Operating Expenses:POS & Technology',a['tech']),('Other Operating Expenses:Insurance - Liability and Commercial',a['insr'])]:
        for loc,v in d.items():
            if v<0: L.append([acct,0,-v,acct.split(':')[-1]+' Adj','',loc])
            elif v: L.append([acct,v,0,acct.split(':')[-1]+' Adj','',loc])
    if a['linen_dt_to_1899']:
        L.append(['Operating Expenses:Operating Supplies:Linen and Uniforms',0,a['linen_dt_to_1899'],'Linen Adj','','GHC3 - Downtown'])
        L.append(['Operating Expenses:Operating Supplies:Linen and Uniforms',a['linen_dt_to_1899'],0,'Linen Adj','','GHC6 - 1899'])
    if a['catering_dt_to_1899']:
        L.append(['Cost of Goods Sold:Core Location COGS:Location Catering and Event COGS',0,a['catering_dt_to_1899'],'Catering COGS Adj','','GHC3 - Downtown'])
        L.append(['Cost of Goods Sold:Core Location COGS:Location Catering and Event COGS',a['catering_dt_to_1899'],0,'Catering COGS Adj','','GHC6 - 1899'])
    share=a['tech'].get('GHC6 - 1899',0)+a['insr'].get('GHC6 - 1899',0)+a['linen_dt_to_1899']+a['catering_dt_to_1899']
    L.append(['Inter-company AR - GHC3',round(share,2),0,'1899 share','','GHC3 - Downtown'])
    L.append(['Inter-company AP - GHC6',0,round(share,2),'1899 share','','GHC6 - 1899'])
    return L
FOH='Employee Labor:Salaries & Wages:FOH Labor'; TAX='Employee Labor:Payroll Tax'
ADMIN='Employee Labor:Salaries & Wages:Management Labor:Shared Management Labor'
TRAIN='Employee Labor:Salaries & Wages:Training Labor'   # NEW account (Aug 2026) — confirm exact QBO name/path before import
LOC={'L3':'GHC3 - Downtown','L5':'GHC5 - Studio','L6':'GHC6 - 1899','L7':'NMWA'}
def payroll_reg(d):
    # Core/Admin/Training split (Aug 2026+): d['SPLIT'] = output of labor_split.py --json for the pay period
    # {'L3':{'admin':x,'training':y},...}. FOH = ADP location gross - admin - training; admin -> 60130 @ location; training -> Training Labor @ location.
    L=[]; A=L.append; SP=d.get('SPLIT',{})
    for k in ('L3','L5','L6','L7'):
        adm=round(SP.get(k,{}).get('admin',0),2); trn=round(SP.get(k,{}).get('training',0),2)
        core=round(d[k]-adm-trn,2); assert core>=0, f"{k}: admin+training {adm+trn} exceeds ADP gross {d[k]}"
        A([FOH,core,0,f'{k} FOH (core)' if SP else f'{k} Adj FOH','',LOC[k]])
        if adm: A([ADMIN,adm,0,f'{k} Admin shifts (Manager)','',LOC[k]])
        if trn: A([TRAIN,trn,0,f'{k} Training shifts','',LOC[k]])
    A(['Employee Labor:Salaries & Wages:Commissary Labor',d['Comm'],0,'Comm Labor (incl Ines)','','Commissary'])
    A(['Operating Expenses:Commissary Delivery',d['CommDrv'],0,'Commissary Delivery','','Commissary'])
    A(['Employee Labor:Salaries & Wages:Roasting Labor',d['SRP'],0,'SRP Labor (Mo + Rstg)','','SRP'])
    A(['Operating Expenses:Roasting Expenses:Roasting Shipping and Delivery',d['SRPDrv'],0,'SRP Delivery','','SRP'])
    A(['Employee Labor:Salaries & Wages:Management Labor:Shared Management Labor',d['Mgmt'],0,'Management + Adj','','Central Admin'])
    A(['Employee Tips Payable',d['Tips'],0,'Tips (Location)','','GHC3 - Downtown'])
    for k,lab,loc in [('TL3','L3','GHC3 - Downtown'),('TL5','L5','GHC5 - Studio'),('TL6','L6','GHC6 - 1899'),('TL7','L7','NMWA'),('TComm','Comm','Commissary'),('TSRP','SRP','SRP')]:
        A([TAX,d[k],0,f'{lab} Payroll Tax','',loc])
    A([TAX,0,abs(d['MgmtRes']),'Mgmt Tax Residual','','Central Admin'])
    if d.get('HC'): A(['Employee Labor:Health Insurance Expense',0,d['HC'],'Healthcare Adj','','GHC3 - Downtown'])
    A(['Inter-company AR - GHC3',d['C1899'],0,'1899 Payroll Clearing','','GHC3 - Downtown'])
    A(['Penn 1899 Payroll Clearing',0,d['C1899'],'1899 Payroll Clearing','','GHC6 - 1899'])
    dr=sum(l[1] for l in L); cr=sum(l[2] for l in L)
    pos=max(i for i,l in enumerate(L) if l[0]==TAX)+1   # clearing plug right after the tax lines (was fixed index 17)
    L.insert(pos,['Payroll Clearing Account',0,round(dr-cr,2),'Clearing','','GHC3 - Downtown'])
    return L
def payroll_accr(d,rev=False):
    base=[(FOH,'L3','GHC3 - Downtown','L3 Adj FOH'),(FOH,'L5','GHC5 - Studio','L5 Adj FOH'),(FOH,'L6','GHC6 - 1899','L6 Adj FOH'),(FOH,'L7','NMWA','L7 Adj FOH'),
     ('Employee Labor:Salaries & Wages:Commissary Labor','Comm','Commissary','Comm Labor (incl Ines)'),
     ('Operating Expenses:Commissary Delivery','CommDrv','Commissary','Commissary Delivery'),
     ('Employee Labor:Salaries & Wages:Roasting Labor','SRP','SRP','SRP Labor (Mo)'),
     ('Operating Expenses:Roasting Expenses:Roasting Shipping and Delivery','SRPDrv','SRP','SRP Delivery'),
     ('Employee Labor:Salaries & Wages:Management Labor:Shared Management Labor','Mgmt','Central Admin','Management + Adj'),
     (TAX,'TL3','GHC3 - Downtown','L3 Payroll Tax'),(TAX,'TL5','GHC5 - Studio','L5 Payroll Tax'),(TAX,'TL6','GHC6 - 1899','L6 Payroll Tax'),
     (TAX,'TL7','NMWA','L7 Payroll Tax'),(TAX,'TComm','Commissary','Comm Payroll Tax'),(TAX,'TSRP','SRP','SRP Payroll Tax')]
    L=[[acct,0 if rev else d[k], d[k] if rev else 0, desc,'',loc] for acct,k,loc,desc in base]
    m=abs(d['MgmtRes']); L.append([TAX,m if rev else 0,0 if rev else m,'Mgmt Tax Residual','','Central Admin'])
    dr=sum(l[1] for l in L); cr=sum(l[2] for l in L); clr=round(abs(dr-cr),2)
    L.append(['Payroll Clearing Account',clr if rev else 0,0 if rev else clr,'Clearing','','GHC3 - Downtown'])
    return L
def recurring(R):
    out=[]
    out.append((f'MgmtFee-{MON}',EOM,[['Management Income Clearing',R['mgmt_fee'],0,'Mgmt Fee (1899)','','Cafe Management'],
        ['Sales:Management Income:Management Fees',0,R['mgmt_fee'],'Mgmt Fee (1899)','','Cafe Management']]))
    out.append((f'SRP-GreenCoffee-{MON}',EOM,[['Cost of Goods Sold:Roasting COGS:Green Coffee COGS',R['srp_green_coffee'],0,f'Implied {MONNAME} Green Coffee','','SRP'],
        ['Green Coffee Inventory',0,R['srp_green_coffee'],f'Implied {MONNAME} Green Coffee','','SRP']]))
    st=R['st_rent']
    stl=[['Occupancy:Rent',st['core'],0,'Rent Amount','','GHC5 - Studio']]
    if st['linen']: stl.append(['Operating Expenses:Operating Supplies:Linen and Uniforms',st['linen'],0,'Linen Share','','GHC5 - Studio'])
    stl.append(['Operating Expenses:Repairs & Maintenance',st['repairs'],0,'Misc Repairs / Etc','','GHC5 - Studio'])
    stl.append(['Rent Liability',0,st['total_due'],'Rent Payment','','GHC5 - Studio'])
    if st['freeze_dried']: stl.append(['Cost of Goods Sold:Core Location COGS:Location Coffee COGS',0,st['freeze_dried'],'Coffee Concentrate','','GHC5 - Studio'])
    if st['general_supplies']: stl.append(['Operating Expenses:Operating Supplies:General Supplies',0,st['general_supplies'],'Dry Goods Share','','GHC5 - Studio'])
    out.append((f'ST-Rent-{MON}',EOM,stl))
    out.append((f'DT-Rent-{MON}',EOM,[['Occupancy:Rent',R['dt_rent'],0,'Rent','','GHC3 - Downtown'],['Rent Liability',0,R['dt_rent'],'Rent','','GHC3 - Downtown']]))
    out.append((f'SRP-Rent-{MON}',EOM,[['Occupancy:Rent',R['srp_rent'],0,'Rent','','SRP'],['Rent Liability',0,R['srp_rent'],'Rent','','SRP']]))
    return out
# ---------- assemble ----------
os.makedirs(OUT,exist_ok=True); os.makedirs(CSVDIR,exist_ok=True)
JEs=[]
name={'GHC5 - Studio':'Studio','GHC6 - 1899':'1899','NMWA':'NMWA','GHC3 - Downtown':'DT'}
fname={'GHC5 - Studio':'Studio Sales JE','GHC6 - 1899':'1899 Penn Sales JE','NMWA':'NMWA Sales JE','GHC3 - Downtown':'Downtown Sales JE'}
for loc,(f,o) in SALES.items():
    JEs.append((f'Sales-{name[loc]}-{MON}',EOM,sales_je(loc,UP+f,o),fname[loc]))
JEs.append((f'Transfers-{MON}',EOM,transfers_je(TRANSFERS),'Transfers JE'))
JEs.append((f'Adjustment-{MON}',EOM,adjustment_je(ADJ),'Adjustment JE'))
for suf,dt,d in PAYROLL:
    JEs.append((f'Payroll-{suf}-26',dt,payroll_reg(d),f'Payroll JE {dt.replace("/","-")}'))
JEs.append((f'Payroll-Accr-{MON}',EOM,payroll_accr(ACCR,False),'Payroll JE Accr'))
JEs.append((f'Payroll-Rev-Aug26',REV_DATE,payroll_accr(ACCR,True),'Payroll JE Reversal'))
for jn,dt,lines in recurring(RECUR):
    JEs.append((jn,dt,lines,jn.replace(f'-{MON}','')+' JE'))
print("JE BALANCE CHECK"); ok=True
for jn,dt,lines,_ in JEs:
    dr=round(sum(l[1] for l in lines),2); cr=round(sum(l[2] for l in lines),2)
    good=abs(dr-cr)<0.005; ok&=good
    print(f"  {jn:22} DR {dr:>12,.2f}  CR {cr:>12,.2f}  Δ {dr-cr:+.2f}  {'OK' if good else 'FAIL'}")
print("ALL BALANCED" if ok else "!!! FAILURES"); assert ok
BOLD=Font(bold=True)
def fmt(x):
    x=round(x,2); return '' if x==0 else (str(int(x)) if x==int(x) else f"{x:.2f}")
for jn,dt,lines,fn in JEs:
    wb=openpyxl.Workbook(); ws=wb.active
    ws.append(['ACCOUNT','DEBITS','CREDITS','DESCRIPTION','NAME','LOCATION'])
    for c in ws[1]: c.font=BOLD
    for a,d,c,desc,nm,loc in lines: ws.append([a,round(d,2) or None,round(c,2) or None,desc,nm,loc])
    ws.column_dimensions['A'].width=58
    for col in 'DEF': ws.column_dimensions[col].width=24
    wb.save(OUT+fn+".xlsx")
rows=[['*JournalNo','*JournalDate','*AccountName','*Debits','*Credits','Description','Name','Currency','Location','Class']]
for jn,dt,lines,_ in JEs:
    first=True
    for a,d,c,desc,nm,loc in lines:
        if round(d,2)==0 and round(c,2)==0: continue
        rows.append([jn,dt if first else '',a,fmt(d),fmt(c),desc,'','USD' if first else '',loc,''])
        first=False
with open(CSVDIR+f"All {MONNAME} JEs (names).csv",'w',newline='',encoding='utf-8-sig') as f:
    csv.writer(f,lineterminator='\r\n').writerows(rows)
print(f"\nWrote {len(JEs)} JEs -> {CSVDIR}All {MONNAME} JEs (names).csv ({len(rows)-1} rows)")
print("1899 shares:", ", ".join(f"{s}={d['C1899']}" for s,_,d in PAYROLL), "| month-end clear =", round(sum(d['C1899'] for _,_,d in PAYROLL),2))
