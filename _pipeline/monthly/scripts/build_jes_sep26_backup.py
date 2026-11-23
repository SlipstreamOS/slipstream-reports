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

### ================= CONFIG (SEPTEMBER 2026) =================
MON="Sep26"; MONNAME="September"; EOM="9/30/26"; REV_DATE="10/1/26"; REV_MON="Oct26"
BASE=os.environ.get("CLOSE_BASE",".")            # dir containing uploads/ and "Weekly and Monthly Reports/"
UP=os.environ.get("CLOSE_UP", os.path.join(BASE,"uploads/"))   # sales CSVs; on-device copy: _pipeline/monthly/data/sales/2026-09/
OUT=os.path.join(BASE,f"Weekly and Monthly Reports/{MONNAME} Sales JEs/")
CSVDIR=os.path.join(BASE,f"Weekly and Monthly Reports/{MONNAME} Sales JEs CSV/")
SALES={  # loc tag -> (csv, opts)   oc_* = mapping of Square "Other Categories" rollup (must sum to the rollup gross)
 'GHC5 - Studio':   (f"Sales Report {MONNAME} Studio.csv", dict(oc_coffee=235.00, oc_food=0.0)),                      # Coffee Drinks 198 + Uncategorized 37 -> Coffee
 'GHC6 - 1899':     (f"Sales Report {MONNAME} 1899.csv",   dict(oc_coffee=84.00, oc_food=48.00)),                     # Coffee Drinks 84 -> Coffee; Grab&Go 48 -> Food
 'NMWA':            (f"Sales Report {MONNAME} NMWA.csv",   dict(oc_coffee=123.75, oc_food=234.00, oc_extra=10882.00,  # Uncategorized 117.75 + Coffee Drinks 6 -> Coffee; Grab&Go 234 -> Food; Services 10,882 = 5 invoice lines
                      other_cash=19242.00, other_cash_desc='Museum ACH (fees #6 + #7, 9/12 closure, NMWA Night inv 000013)',   # Other tender = 5,041 + 5,041 + 800 + 8,360 (Square invoices 000007/000010/000009/000013)
                      extra=[('Management Income Clearing',0,5041.00,'NMWA fee payment - inv #6 (Aug), paid 9/15','NMWA'),      # Ryan 9/30: fee PAYMENTS go to clearing; income is recognized monthly by NMWA-MgmtEquipFee JE
                             ('Management Income Clearing',0,5041.00,'NMWA fee payment - inv #7 (Sept), paid 9/15','NMWA'),
                             ('Sales:Management Income:Management Fees',0,800.00,'Cafe closure payment 9/12','NMWA')])),   # July precedent: closure -> Management Fees @ NMWA
 'GHC3 - Downtown': (f"Sales Report {MONNAME} DT.csv",     dict(oc_coffee=240.00, oc_food=174.75)),   # Coffee Drinks 240 -> Coffee; Grab&Go 174.75 -> Food. No $5,100 1899 mgmt-fee item this month; svc chg 661 -> Catering
}
TRANSFERS=dict(comm_sales_food=7739.82, comm_sales_coffee=677.19,   # Month Close Sheet -> Transfers tab col AA (9/30/26); built by Ryan from Square invoices
 comm_purch={'GHC3 - Downtown':(2807.67,209.00),'GHC5 - Studio':(3967.54,290.48),'GHC6 - 1899':(767.03,177.71),'NMWA':(197.58,0.0)},
 srp_roasted=4984.00, srp_concentrate=3415.00,
 srp_purch={'GHC3 - Downtown':2637.00,'GHC5 - Studio':3569.00,'GHC6 - 1899':1993.00,'NMWA':200.00})
ADJ=dict(tech={'GHC3 - Downtown':-700.00,'GHC5 - Studio':450.00,'GHC6 - 1899':250.00},    # Transfers col AA rows 25-32 (Studio = 350+100 / 275+50), same as Aug
 insr={'GHC3 - Downtown':-650.00,'GHC5 - Studio':325.00,'GHC6 - 1899':325.00},
 linen_dt_to_1899=150.00, catering_dt_to_1899=110.40)   # Catering COGS = 40% x 1899 catering booked (276.00, no svc chg)
PAYROLL=[ # (journal suffix MMDD, date, dict) — values replicate Month Close Sheet Payroll Detail row formulas from ADP PDFs + Payroll Workbook Location Totals
 ('0911','9/11/26', dict(L3=4070.64,L5=6162.66,L6=2675.20,L7=928.51,Comm=2482.12,CommDrv=300,SRP=913.34,SRPDrv=180,Mgmt=4298.18,
   Tips=5375.02,TL3=698.99,TL5=1047.82,TL6=433.77,TL7=124.86,TComm=120,TSRP=109.60,MgmtRes=-246.36,HC=418.97,C1899=4048.56,
   SPLIT={'L3':{'admin':0.0,'training':0.0},'L5':{'admin':742.94,'training':279.23},'L6':{'admin':142.19,'training':0.0},'L7':{'admin':0.0,'training':0.0}})),   # labor_split.py 2026-08-23 2026-09-05
 ('0925','9/25/26', dict(L3=4034.29,L5=4806.83,L6=2328.10,L7=1638.37,Comm=1669.66,CommDrv=300,SRP=906.66,SRPDrv=180,Mgmt=4795.11,
   Tips=6171.38,TL3=753.07,TL5=904.72,TL6=393.93,TL7=225.76,TComm=120,TSRP=108.80,MgmtRes=-282.14,HC=418.97,C1899=3676.65,
   SPLIT={'L3':{'admin':559.29,'training':0.0},'L5':{'admin':266.28,'training':279.78},'L6':{'admin':268.77,'training':0.0},'L7':{'admin':136.12,'training':0.0}})),   # labor_split.py 2026-09-06 2026-09-19
]
ACCR=dict(L3=3169.80,L5=3776.79,L6=1829.22,L7=1287.29,Comm=1311.88,CommDrv=235.71,SRP=712.38,SRPDrv=141.43,Mgmt=3767.59,   # 11/14 of 9/25 row (9/20-9/30)
 TL3=591.70,TL5=710.85,TL6=309.51,TL7=177.38,TComm=94.29,TSRP=85.49,MgmtRes=-221.68)
NN_COGS=434.70   # 35% x $1,242 Member Preview breakfast (9/16, inv 000012) - Ryan's estimate 10/1; food was billed to Downtown, so credit DT Food Purchases (NMWA Night $799.88 was charged to the GHC3 card and categorized direct to NMWA)
RECUR=dict(mgmt_fee=5000.00, nmwa_fee=5041.00, srp_green_coffee=3866.91,   # SRP - MONTHLY INVENTORY 2026 -> OCTOBER 2026 tab E41 (31,591.55 + 0.00 purchases - 27,724.64)
 dt_rent=2112.09, srp_rent=1900.00,
 st_rent=dict(core=5857.50, linen=191.60, repairs=0.0, total_due=5951.10, general_supplies=98.00, freeze_dried=0.0))   # Studio Rent Sheet col Sep 2026: Net Sales 45,941.14
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
    if st is None:
        print("  ! ST-Rent JE SKIPPED (st_rent pending)")
        st=dict(core=0,linen=0,repairs=0,total_due=0,freeze_dried=0,general_supplies=0)
    stl=[['Occupancy:Rent',st['core'],0,'Rent Amount','','GHC5 - Studio']]
    if st['linen']: stl.append(['Operating Expenses:Operating Supplies:Linen and Uniforms',st['linen'],0,'Linen Share','','GHC5 - Studio'])
    stl.append(['Operating Expenses:Repairs & Maintenance',st['repairs'],0,'Misc Repairs / Etc','','GHC5 - Studio'])
    stl.append(['Rent Liability',0,st['total_due'],'Rent Payment','','GHC5 - Studio'])
    if st['freeze_dried']: stl.append(['Cost of Goods Sold:Core Location COGS:Location Coffee COGS:Coffee Purchases',0,st['freeze_dried'],'Coffee Concentrate','','GHC5 - Studio'])   # 50110, not the 50100 parent (fixed Sept 2026)
    if st['general_supplies']: stl.append(['Operating Expenses:Operating Supplies:General Supplies',0,st['general_supplies'],'Dry Goods Share','','GHC5 - Studio'])
    if R['st_rent'] is not None: out.append((f'ST-Rent-{MON}',EOM,stl))
    out.append((f'NMWA-MgmtFee-{MON}',EOM,[['Management Income Clearing',R['nmwa_fee'],0,'NMWA Mgmt + Equip fee','','NMWA'],
        ['Sales:Management Income:Management Fees',0,3950.00,'Mgmt Fee (NMWA)','','NMWA'],['Sales:Management Income:Management Fees',0,1091.00,'Equip Fee (NMWA)','','NMWA']]))
    out.append((f'DT-Rent-{MON}',EOM,[['Occupancy:Rent',R['dt_rent'],0,'Rent','','GHC3 - Downtown'],['Rent Liability',0,R['dt_rent'],'Rent','','GHC3 - Downtown']]))
    out.append((f'SRP-Rent-{MON}',EOM,[['Occupancy:Rent',R['srp_rent'],0,'Rent','','SRP'],['Rent Liability',0,R['srp_rent'],'Rent','','SRP']]))
    return out
# ---------- assemble ----------
os.makedirs(OUT,exist_ok=True); os.makedirs(CSVDIR,exist_ok=True)
JEs=[]; pending=[]
name={'GHC5 - Studio':'Studio','GHC6 - 1899':'1899','NMWA':'NMWA','GHC3 - Downtown':'DT'}
fname={'GHC5 - Studio':'Studio Sales JE','GHC6 - 1899':'1899 Penn Sales JE','NMWA':'NMWA Sales JE','GHC3 - Downtown':'Downtown Sales JE'}
for loc,(f,o) in SALES.items():
    if not os.path.exists(UP+f): pending.append(f'Sales-{name[loc]}'); print(f"  ! Sales-{name[loc]} JE SKIPPED ({f} missing)"); continue
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
if RECUR.get('st_rent') is None: pending.append('ST-Rent')
CATCOGS='Cost of Goods Sold:Core Location COGS:Location Catering and Event COGS'
if NN_COGS: JEs.append((f'NN-COGS-Trans-{MON}',EOM,[[CATCOGS,NN_COGS,0,'NMWA Member Preview breakfast food cost (35% est)','','NMWA'],[FOODP,0,NN_COGS,'NMWA event food cost (billed to DT)','','GHC3 - Downtown']],'NN COGS Trans JE'))
else: print('  ! NN-COGS-Trans not included (Commissary food cost for NMWA events pending) - import separately')
for jn,_,_,_ in JEs: assert len(jn)<=21, f"JournalNo {jn} is {len(jn)} chars; QBO limit is 21"
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
