#!/usr/bin/env python3
"""Sync Consolidated Model from Coefficient Output + bump last-closed-month.  v2 (Oct 2026): IN-PLACE XML PATCH.
Usage: python3 sync_model.py <in_version> <out_version> <YYYY-MM>   e.g. python3 sync_model.py 17 18 2026-09

WHY v2: Ryan edits the model offline in Excel. The old version round-tripped the whole workbook through openpyxl,
which rebuilds the file and silently drops things (e.g. conditional-formatting extensions). Ryan's rule (Oct 2 2026):
"make edits to the existing, do not recreate."  So this script NEVER loads/saves the model with openpyxl. It copies
the CURRENT on-disk v<in> byte-for-byte and patches only:
  - the 8 'QBO - X' sheets' cell data (rows 1-320, cols A..<close month> + col P acct# helper)
  - Assumptions!C5 (last closed month end)
  - workbook calcPr fullCalcOnLoad=1 (Excel recalcs on open)
Every other sheet, style, format, formula and setting is the identical bytes Ryan saved. A verify step asserts that.
ALWAYS run against the latest file on disk (check its modified time; if a '~$' lock file exists Ryan has it open -
ask him to save and close first).  NEVER openpyxl-save the Coefficient Output file (kills the add-in connection).
"""
import openpyxl, re, sys, datetime, calendar, os, zipfile, copy, warnings
from lxml import etree
warnings.simplefilter('ignore')
V_IN,V_OUT,YM=sys.argv[1],sys.argv[2],sys.argv[3]
y,mo=int(YM[:4]),int(YM[5:7]); eom=datetime.datetime(y,mo,calendar.monthrange(y,mo)[1])
D=os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","..",".."))
BASE=os.path.join(D,f"Slipstream Consolidated Model 2026 v{V_IN}.xlsx")
CO=os.path.join(D,"Coefficient Output 2026.xlsx")
OUT=os.path.join(D,f"Slipstream Consolidated Model 2026 v{V_OUT}.xlsx")
assert not os.path.exists(os.path.join(D,f"~$Slipstream Consolidated Model 2026 v{V_IN}.xlsx")), "model is open in Excel - save and close first"
TABMAP={'QBO - Studio':'QBO - Studio','QBO - Downtown':'QBO - DT','QBO - 1899':'QBO - 1899',
 'QBO - NMWA':'QBO - NMWA','QBO - Commissary':'QBO - Commissary','QBO - SRP':'QBO - SRP',
 'QBO - Cafe Mgmt':'QBO - Café Mgmt','QBO - Central':'QBO - Central Admin'}
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
Q=lambda t:'{%s}%s'%(NS,t)
def col2n(s):
    n=0
    for ch in s: n=n*26+ord(ch)-64
    return n
def n2col(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
serial=lambda d:(d-datetime.datetime(1899,12,30)).total_seconds()/86400
co=openpyxl.load_workbook(CO,data_only=True)          # read-only use; never saved
zin=zipfile.ZipFile(BASE); members={n:zin.read(n) for n in zin.namelist()}; infos={i.filename:i for i in zin.infolist()}
wbx=etree.fromstring(members['xl/workbook.xml']); rels=etree.fromstring(members['xl/_rels/workbook.xml.rels'])
rid2t={r.get('Id'):r.get('Target') for r in rels}
sheetfile={s.get('name'):'xl/'+rid2t[s.get('{%s}id'%RNS)].lstrip('/').replace('xl/','') for s in wbx.find(Q('sheets'))}
changed=set()
for mt,ot in TABMAP.items():
    ows=co[ot]
    # month columns to bring over: header row 5 dates <= close month (plus col A)
    ncol=1
    for c in range(2,ows.max_column+1):
        h=ows.cell(5,c).value
        if isinstance(h,datetime.datetime) and (h.year,h.month)<=(y,mo): ncol=c
    assert ncol>=1+mo, f"{ot}: close month column not found in Coefficient Output"
    root=etree.fromstring(members[sheetfile[mt]]); sd=root.find(Q('sheetData'))
    rows={int(r.get('r')):r for r in sd}
    style={}
    for rn,r in rows.items():
        if rn>320: continue
        for c in list(r):
            m=re.match(r'([A-Z]+)(\d+)',c.get('r')); cn=col2n(m.group(1))
            if cn<=16:
                if c.get('s'): style[(rn,cn)]=c.get('s')
                r.remove(c)
    nums=0
    def put(rn,cn,val):
        s=style.get((rn,cn))
        if s is None and cn>2 and cn<=ncol: s=style.get((rn,cn-1)) if rn==5 else None   # new month header inherits date format
        if val is None and s is None: return
        if rn not in rows:
            r=etree.SubElement(sd,Q('row')); r.set('r',str(rn)); rows[rn]=r
        c=etree.SubElement(rows[rn],Q('c')); c.set('r',n2col(cn)+str(rn))
        if s: c.set('s',s)
        if val is None: return
        if isinstance(val,datetime.datetime): etree.SubElement(c,Q('v')).text=repr(serial(val)).rstrip('0').rstrip('.') if serial(val)%1 else str(int(serial(val)))
        elif isinstance(val,bool): c.set('t','b'); etree.SubElement(c,Q('v')).text='1' if val else '0'
        elif isinstance(val,(int,float)): etree.SubElement(c,Q('v')).text=repr(val)
        else:
            c.set('t','inlineStr'); t=etree.SubElement(etree.SubElement(c,Q('is')),Q('t')); t.text=str(val)
            t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    lead=lambda s:(lambda m:m.group(1) if m else None)(re.match(r'\s*(\d{5})\b',str(s or '')))
    for rn in range(1,max(ows.max_row,max([k for k in rows if k<=320] or [1]))+1):
        vals={c:(ows.cell(rn,c).value if rn<=ows.max_row and c<=ncol else None) for c in range(1,16)}
        vals[16]=None
        if rn==5: vals[16]='acct# (helper)'
        elif rn>=6 and rn<=ows.max_row:
            n=lead(ows.cell(rn,1).value)
            if n: vals[16]=int(n); nums+=1
        for cn in range(1,17): put(rn,cn,vals[cn])
    for r in rows.values():                                 # cells in column order, rows in row order
        cs=sorted(r,key=lambda c:col2n(re.match(r'[A-Z]+',c.get('r')).group(0)))
        for c in cs: r.remove(c)
        for c in cs: r.append(c)
        if 'spans' in r.attrib: del r.attrib['spans']
    ordered=sorted(sd,key=lambda r:int(r.get('r')))
    for r in ordered: sd.remove(r)
    for r in ordered: sd.append(r)
    dim=root.find(Q('dimension'))
    if dim is not None: dim.set('ref',f"A1:P{max(rows)}")
    assert nums>3, f"{ot}: no account numbers found - Ryan must re-add numbers in Coefficient settings"
    members[sheetfile[mt]]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True); changed.add(sheetfile[mt])
    print(f"  {mt}: synced cols A-{n2col(ncol)}, {nums} numbered accounts")
# Assumptions!C5
af=sheetfile['Assumptions']; root=etree.fromstring(members[af])
c5=[c for c in root.iter(Q('c')) if c.get('r')=='C5'][0]
assert c5.find(Q('f')) is None, "Assumptions!C5 is a formula now - check with Ryan"
c5.find(Q('v')).text=str(int(serial(eom)))
members[af]=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True); changed.add(af)
cp=wbx.find(Q('calcPr'))
if cp is None: cp=etree.SubElement(wbx,Q('calcPr'))
cp.set('fullCalcOnLoad','1')
members['xl/workbook.xml']=etree.tostring(wbx,xml_declaration=True,encoding='UTF-8',standalone=True); changed.add('xl/workbook.xml')
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as zo:
    for n in zin.namelist():
        zi=zipfile.ZipInfo(n,date_time=infos[n].date_time); zi.compress_type=zipfile.ZIP_DEFLATED; zi.external_attr=infos[n].external_attr
        zo.writestr(zi,members[n])
# verify: everything Ryan authored is byte-identical
z2=zipfile.ZipFile(OUT); diff=[n for n in zin.namelist() if z2.read(n)!=zin.read(n)]
assert set(diff)<=changed and z2.namelist()==zin.namelist(), ("unexpected differences",set(diff)-changed)
print(f"Saved {os.path.basename(OUT)} | C5 -> {eom.date()} | {len(zin.namelist())-len(diff)} of {len(zin.namelist())} parts byte-identical to v{V_IN}; changed: 8 QBO sheets, Assumptions (C5), workbook calc flag")
print("NEXT: recalc copy + validate:  soffice --headless --calc --convert-to xlsx --outdir $HOME/rc '<OUT>'  then python3 validate_model.py <ver> <YYYY-MM> $HOME/rc")
