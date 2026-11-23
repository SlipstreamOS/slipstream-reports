#!/usr/bin/env python3
"""Slipstream month-close JE builder. CONFIG block below is the only thing to edit each month.
Validated June+July 2026; August 2026 adds the Core/Admin/Training payroll split (SPLIT per pay date),
`oc_ref_food` (route Other-Categories refunds to Returns - Food), and PENDING inputs (set TRANSFERS /
RECUR['srp_green_coffee'] to None to skip those JEs until Ryan's sheets are filled -> DRAFT CSV).
Reads Square sales CSVs from UP, writes per-JE xlsx + combined import CSV (QBO format: UTF-8 BOM, CRLF,
date/USD on first JE line, zero rows filtered).
Run on-device:  cd _pipeline/monthly/scripts && CLOSE_BASE="$HOME/mnt" CLOSE_UP="../data/sales/2026-08/" python3 build_jes.py"""
import csv, openpyxl, os, sys
from openpyxl.styles import Font

### ================= CONFIG (AUGUST 2026) =================
MON="Aug26"; MONNAME="August"; EOM="8/31/26"; REV_DATE="9/1/26"; REV_MON="Sep26"
BASE=os.environ.get("CLOSE_BASE",".")            # dir containing uploads/ and "Weekly and Monthly Reports/"
UP=os.environ.get("CLOSE_UP", os.path.join(BASE,"uploads/"))   # sales CSVs; on-device copy: _pipeline/monthly/data/sales/2026-08/
OUT=os.path.join(BASE,f"Weekly and Monthly Reports/{MONNAME} Sales JEs/")
CSVDIR=os.path.join(BASE,f"Weekly and Monthly Reports/{MONNAME} Sales JEs CSV/")
SALES={  # loc tag -> (csv, opts)   oc_* = mapping of Square "Other Categories" rollup (must sum to the rollup gross)
 'GHC5 - Studio':   (f"Sales Report {MONNAME} Studio.csv", dict(oc_coffee=302.00, oc_food=5.73, oc_ref_food=True)),   # Coffee Drinks 252 + Uncategorized 50 -> Coffee; Commissary-Food Prep 5.73 (refunded 5.73) -> Food
 'GHC6 - 1899':     (f"Sales Report {MONNAME} 1899.csv",   dict(oc_coffee=150.00, oc_food=24.00)),                    # Coffee Drinks 90 + Uncategorized(invoice) 60 -> Coffee; Grab&Go 24 -> Food; svc chg 12 -> Catering
 'NMWA':            (f"Sales Report {MONNAME} NMWA.csv",   dict(oc_coffee=158.25, oc_food=243.00, oc_extra=5041.00,   # Uncategorized 128.25 + Coffee Drinks 30 -> Coffee; Grab&Go 243 -> Food
                      other_cash=5041.00, other_cash_desc='Museum ACH (Mgmt + Equip fee)',
                      extra=[('Sales:Management Income:Management Fees',0,3950.00,'Mgmt Fee (NMWA)','NMWA'),
                             ('Sales:Management Income:Management Fees',0,1091.00,'Equip Fee (NMWA)','NMWA')])),
 'GHC3 - Downtown': (f"Sales Report {MONNAME} DT.csv",     dict(oc_coffee=216.00, oc_food=105.50, oc_extra=5000.00,   # Coffee Drinks 198 + Uncategorized 18 -> Coffee; Grab&Go 102 + Comm Food Prep 3.50 -> Food
                      svc_to_merchant=100.00,                                                                          # 1 x $5,100 1899 mgmt-fee payment via DT POS (invoice $5,000 + $100 svc chg)
                      extra=[('Management Income Clearing',0,5000.00,'Mgmt Fee (1899) 1x5000','Cafe Management'),
                             ('Other Operating Expenses:Merchant Fees',0,100.00,'Mgmt Fee CC offset (1899)','GHC3 - Downtown')])),
}
TRANSFERS=dict(comm_sales_food=8248.40, comm_sales_coffee=801.99,   # Month Close Sheet -> Transfers tab col Z (8/31/26); built by Ryan from Square invoices
 comm_purch={'GHC3 - Downtown':(2922.03,257.88),'GHC5 - Studio':(4236.80,388.92),'GHC6 - 1899':(915.71,155.19),'NMWA':(173.86,0.0)},
 srp_roasted=5014.00, srp_concentrate=5895.00,
 srp_purch={'GHC3 - Downtown':3686.00,'GHC5 - Studio':4610.00,'GHC6 - 1899':2247.00,'NMWA':366.00})
ADJ=dict(tech={'GHC3 - Downtown':-700.00,'GHC5 - Studio':450.00,'GHC6 - 1899':250.00},    # Transfers tab col Z rows 25-32 (Studio = 350+100 / 275+50), same as July
 insr={'GHC3 - Downtown':-650.00,'GHC5 - Studio':325.00,'GHC6 - 1899':325.00},
 linen_dt_to_1899=150.00, catering_dt_to_1899=22.40)   # Catering COGS = 40% x 1899 catering booked (44.00 + 12.00 svc chg = 56.00)
PAYROLL=[ # (journal suffix MMDD, date, dict) — values replicate Month Close Sheet Payroll Detail row formulas from ADP PDFs + Payroll Workbook Location Totals
 ('0814','8/14/26', dict(L3=4920.51,L5=6417.61,L6=2956.88,L7=1070.27,Comm=2123.88,CommDrv=300,SRP=940.07,SRPDrv=180,Mgmt=5538.00,
   Tips=5378.93,TL3=827.50,TL5=1070.32,TL6=439.25,TL7=152.24,TComm=120,TSRP=112.81,MgmtRes=-186.15,HC=418.97,C1899=4099.69,
   SPLIT={'L3':{'admin':484.81,'training':0.0},'L5':{'admin':173.25,'training':327.50},'L6':{'admin':101.62,'training':0.0},'L7':{'admin':0.0,'training':0.0}})),   # labor_split.py 2026-07-26 2026-08-08
 ('0828','8/28/26', dict(L3=5527.47,L5=6527.72,L6=3310.82,L7=760.53,Comm=2456.73,CommDrv=300,SRP=1066.80,SRPDrv=180,Mgmt=5353.72,
   Tips=4883.78,TL3=841.01,TL5=1065.07,TL6=490.53,TL7=124.63,TComm=120,TSRP=128.02,MgmtRes=-179.20,HC=418.97,C1899=4578.27,
   SPLIT={'L3':{'admin':197.79,'training':0.0},'L5':{'admin':504.03,'training':471.87},'L6':{'admin':166.08,'training':328.98},'L7':{'admin':0.0,'training':0.0}})),   # labor_split.py 2026-08-09 2026-08-22
]
ACCR=dict(L3=3553.37,L5=4196.39,L6=2128.38,L7=488.91,Comm=1579.33,CommDrv=192.86,SRP=685.80,SRPDrv=115.71,Mgmt=3441.68,   # 9/14 of 8/28 row (8/23-8/31)
 TL3=540.65,TL5=684.69,TL6=315.34,TL7=80.12,TComm=77.14,TSRP=82.30,MgmtRes=-115.20)
RECUR=dict(mgmt_fee=5000.00, srp_green_coffee=5691.17,   # SRP - MONTHLY INVENTORY 2026 -> SEPTEMBER 2026 tab E41 (28,116.37 + 8,376.39 - 30,801.59; label still reads 'July')
 dt_rent=2112.09, srp_rent=1900.00,
 st_rent=dict(core=5769.61, linen=191.60, repairs=0.0, total_due=5784.90, general_supplies=96.31, freeze_dried=80.00))   # Studio Rent Sheet col Aug 2026: Net Sales 45,251.85
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
    ocr=abs(ref('Other Categories')); ocr_food=ocr if o.get('oc_ref_food') else 0.0; ocr_coffee=ocr-ocr_food
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
    A(['Sales:Core Location Sales:Coffee Sales:Coffee Beverage Sales',abs(ref('Coffee'))+ocr_coffee,0,'Returns - Coffee','',loc])
    A(['Sales:Core Location Sales:Coffee Sales:Retail Coffee Sales',abs(ref('Retail')),0,'Returns - Retail','',loc])
    A(['Sales:Core Location Sales:Food Sales',abs(ref('Food'))+ocr_food,0,'Returns - Food','',loc])
    A(['Sales:Core Location Sales:Catering and Event Sales',abs(ref('Catering and Events')),0,'Returns - Catering and events','',loc])
    A(['Other Operating Expenses:Merchant Fees',abs(sv('Fees')),0,'merchant fees','',loc])
    A(['Gift Card Liability',sv('Gift Card'),0,'gift card liability claimed','',loc])
    A(['GHC3 Cash on hand',sv('Cash'),0,'Cash','',loc])
    if o.get('other_check'): A(['GHC3 Cash on hand',o['other_check'],0,'check','',loc])
    if o.get('other_cash'): A(['GHC3 Cash on hand',o['other_cash'],0,o.get('other_cash_desc','ACH'),'',loc])
    oe=sv('Other')-o.get('other_check',0.0)-o.get('other_cash',0.0)
    assert oe>-0.005, f"{loc}: Other tender {sv('Other')} < mapped check/ACH"
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
TRAIN='Employee Labor:Salaries & Wages:Training Labor'   # 60145 Training Labor, created by Ryan Aug 2026 (parent assumed Salaries & Wages — confirm path)
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
    pos=max(i for i,l in enumerate(L) if l[0]==TAX)+1   # clearing plug right after the tax lines
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
    if R.get('srp_green_coffee') is not None:
        out.append((f'SRP-GreenCoffee-{MON}',EOM,[['Cost of Goods Sold:Roasting COGS:Green Coffee COGS',R['srp_green_coffee'],0,f'Implied {MONNAME} Green Coffee','','SRP'],
            ['Green Coffee Inventory',0,R['srp_green_coffee'],f'Implied {MONNAME} Green Coffee','','SRP']]))
    else: print("  ! SRP-GreenCoffee JE SKIPPED (srp_green_coffee pending)")
    st=R['st_rent']
    stl=[['Occupancy:Rent',st['core'],0,'Rent Amount','','GHC5 - Studio']]
    if st['linen']: stl.append(['Operating Expenses:Operating Supplies:Linen and Uniforms',st['linen'],0,'Linen Share','','GHC5 - Studio'])
    stl.append(['Operating Expenses:Repairs & Maintenance',st['repairs'],0,'Misc Repairs / Etc','','GHC5 - Studio'])
    stl.append(['Rent Liability',0,st['total_due'],'Rent Payment','','GHC5 - Studio'])
    if st['freeze_dried']: stl.append(['Cost of Goods Sold:Core Location COGS:Location Coffee COGS:Coffee Purchases',0,st['freeze_dried'],'Coffee Concentrate','','GHC5 - Studio'])   # 50110, not the 50100 parent (fixed Sept 2026)
    if st['general_supplies']: stl.append(['Operating Expenses:Operating Supplies:General Supplies',0,st['general_supplies'],'Dry Goods Share','','GHC5 - Studio'])
    out.append((f'ST-Rent-{MON}',EOM,stl))
    out.append((f'DT-Rent-{MON}',EOM,[['Occupancy:Rent',R['dt_rent'],0,'Rent','','GHC3 - Downtown'],['Rent Liability',0,R['dt_rent'],'Rent','','GHC3 - Downtown']]))
    out.append((f'SRP-Rent-{MON}',EOM,[['Occupancy:Rent',R['srp_rent'],0,'Rent','','SRP'],['Rent Liability',0,R['srp_rent'],'Rent','','SRP']]))
    return out
# ---------- assemble ----------
os.makedirs(OUT,exist_ok=True); os.makedirs(CSVDIR,exist_ok=True)
JEs=[]; pending=[]
name={'GHC5 - Studio':'Studio','GHC6 - 1899':'1899','NMWA':'NMWA','GHC3 - Downtown':'DT'}
fname={'GHC5 - Studio':'Studio Sales JE','GHC6 - 1899':'1899 Penn Sales JE','NMWA':'NMWA Sales JE','GHC3 - Downtown':'Downtown Sales JE'}
for loc,(f,o) in SALES.items():
    JEs.append((f'Sales-{name[loc]}-{MON}',EOM,sales_je(loc,UP+f,o),fname[loc]))
if TRANSFERS: JEs.append((f'Transfers-{MON}',EOM,transfers_je(TRANSFERS),'Transfers JE'))
else: pending.append('Transfers'); print("  ! Transfers JE SKIPPED (TRANSFERS pending)")
JEs.append((f'Adjustment-{MON}',EOM,adjustment_je(ADJ),'Adjustment JE'))
for suf,dt,d in PAYROLL:
    JEs.append((f'Payroll-{suf}-26',dt,payroll_reg(d),f'Payroll JE {dt.replace("/","-")}'))
JEs.append((f'Payroll-Accr-{MON}',EOM,payroll_accr(ACCR,False),'Payroll JE Accr'))
JEs.append((f'Payroll-Rev-{REV_MON}',REV_DATE,payroll_accr(ACCR,True),'Payroll JE Reversal'))
for jn,dt,lines in recurring(RECUR):
    JEs.append((jn,dt,lines,jn.replace(f'-{MON}','')+' JE'))
if RECUR.get('srp_green_coffee') is None: pending.append('SRP-GreenCoffee')
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
tag=" DRAFT" if pending else ""
csvname=f"All {MONNAME} JEs (names){tag}.csv"
with open(CSVDIR+csvname,'w',newline='',encoding='utf-8-sig') as f:
    csv.writer(f,lineterminator='\r\n').writerows(rows)
print(f"\nWrote {len(JEs)} JEs -> {CSVDIR}{csvname} ({len(rows)-1} rows)"+(f"  [PENDING: {', '.join(pending)}]" if pending else ""))
print("1899 shares:", ", ".join(f"{s}={d['C1899']}" for s,_,d in PAYROLL), "| month-end clear =", round(sum(d['C1899'] for _,_,d in PAYROLL),2))
