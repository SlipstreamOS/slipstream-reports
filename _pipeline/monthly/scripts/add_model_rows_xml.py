#!/usr/bin/env python3
"""Add IS rows to the Consolidated Model by IN-PLACE XML PATCH (never openpyxl-save).
Each new row is a copy of an existing add-on row on the same tab (same INDEX/MATCH-on-acct# formula),
with its own label + account number, and is added to a total row's formulas (master/plain formulas only;
shared children follow their master). Also supports plain cell value patches on QBO tabs.
Usage: python3 add_model_rows_xml.py <version>     (edits CONFIG below; patches the file in place, keeps a .bak copy)
"""
import sys,os,re,zipfile,shutil
from lxml import etree
V=sys.argv[1]
D=os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","..",".."))
P=os.path.join(D,f"Slipstream Consolidated Model 2026 v{V}.xlsx")
# CONFIG (Oct 2026 close): (tab, template_row, new_row, label, acct, total_row, sign)
ROWS=[('Income Statement - Studio',104,105,'   Travel & Entertainment (61280)',61280,80,'+'),
      ('Income Statement - Central',43,44,'   62152 Legal & Professional Services',62152,34,'+'),
      ('Income Statement - Central',43,45,'   62180 Cash Over/Short',62180,34,'+'),
      ('Income Statement - Central',43,46,'   82200 Interest Expense (below the line)',82200,36,'-'),
      ('Income Statement - Commissary',41,48,'   62110 Bank Charges & Fees',62110,43,'+')]
# (tab, cell, value or None): QBO-tab data cells to mirror a QBO change made after the Coefficient refresh
CELLS=[('QBO - Downtown','K47',296.92),('QBO - Downtown','K88',None)]
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
Q=lambda t:'{%s}%s'%(NS,t)
def col2n(s):
    n=0
    for ch in s: n=n*26+ord(ch)-64
    return n
zin=zipfile.ZipFile(P); members={n:zin.read(n) for n in zin.namelist()}; infos={i.filename:i for i in zin.infolist()}
wbx=etree.fromstring(members['xl/workbook.xml']); rels=etree.fromstring(members['xl/_rels/workbook.xml.rels'])
rid={r.get('Id'):r.get('Target') for r in rels}
sf={s.get('name'):'xl/'+rid[s.get('{%s}id'%RNS)].lstrip('/').replace('xl/','') for s in wbx.find(Q('sheets'))}
roots={}; changed=set()
import openpyxl,warnings; warnings.simplefilter('ignore')
OPX=openpyxl.load_workbook(P,data_only=False)   # read-only use; never saved
def root(t):
    if t not in roots: roots[t]=etree.fromstring(members[sf[t]])
    return roots[t]
def rowmap(rt): return {int(r.get('r')):r for r in rt.find(Q('sheetData'))}
for tab,tr,nr,label,acct,tot,sign in ROWS:
    rt=root(tab); sd=rt.find(Q('sheetData')); rows=rowmap(rt)
    assert nr not in rows, f"{tab} row {nr} already exists"
    new=etree.fromstring(etree.tostring(rows[tr])); new.set('r',str(nr))
    if 'spans' in new.attrib: del new.attrib['spans']
    refre=re.compile(r'(?<![A-Za-z0-9_$])(\$?[A-Z]{1,3}\$?)%d(?![0-9])'%tr)
    for c in list(new):
        m=re.match(r'([A-Z]+)(\d+)',c.get('r')); cl=m.group(1); c.set('r',f"{cl}{nr}")
        f=c.find(Q('f')); v=c.find(Q('v'))
        if f is not None:
            if f.get('t')=='shared':   # expand: take this cell's own formula text (openpyxl read-only translates shared formulas)
                ftxt=OPX[tab].cell(tr,col2n(cl)).value; assert isinstance(ftxt,str) and ftxt.startswith('=')
                for k in list(f.attrib): del f.attrib[k]
                f.text=ftxt[1:]
            f.text=refre.sub(lambda mm:mm.group(1)+str(nr),f.text)
            if v is not None: c.remove(v)
        if cl=='B':
            for k in list(c): c.remove(k)
            c.set('t','inlineStr'); t=etree.SubElement(etree.SubElement(c,Q('is')),Q('t')); t.text=label
            t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        if cl=='C': c.find(Q('v')).text=str(acct)
    # insert in order
    after=max(k for k in rows if k<nr); idx=list(sd).index(rows[after]); sd.insert(idx+1,new)
    # add to total row
    n_upd=0
    for c in rowmap(rt)[tot]:
        f=c.find(Q('f'))
        if f is not None and f.text:
            cl=re.match(r'[A-Z]+',c.get('r')).group(0); f.text=f"{f.text}{sign}{cl}{nr}"; n_upd+=1
            v=c.find(Q('v'))
            if v is not None: c.remove(v)
    assert n_upd>0
    dim=rt.find(Q('dimension'))
    if dim is not None:
        a,b=dim.get('ref').split(':'); mb=re.match(r'([A-Z]+)(\d+)',b); dim.set('ref',f"{a}:{mb.group(1)}{max(int(mb.group(2)),nr)}")
    changed.add(tab); print(f"  {tab}: row {nr} '{label.strip()}' (acct {acct}) added; {n_upd} formulas in row {tot} now {sign}row {nr}")
for tab,ref,val in CELLS:
    rt=root(tab); rn=int(re.match(r'[A-Z]+(\d+)',ref).group(1)); r=rowmap(rt)[rn]
    c=[x for x in r if x.get('r')==ref]
    if not c: c=etree.SubElement(r,Q('c')); c.set('r',ref)
    else: c=c[0]
    for k in list(c): c.remove(k)
    if 't' in c.attrib: del c.attrib['t']
    if val is not None: etree.SubElement(c,Q('v')).text=repr(val)
    changed.add(tab); print(f"  {tab}!{ref} = {val}")
for t in changed: members[sf[t]]=etree.tostring(roots[t],xml_declaration=True,encoding='UTF-8',standalone=True)
bak=P+".pre_rows.bak"; shutil.copy2(P,bak)
tmp=P+".tmp"
with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
    for n in zin.namelist():
        zi=zipfile.ZipInfo(n,date_time=infos[n].date_time); zi.compress_type=zipfile.ZIP_DEFLATED; zi.external_attr=infos[n].external_attr
        zo.writestr(zi,members[n])
z2=zipfile.ZipFile(tmp); diff=[n for n in zin.namelist() if z2.read(n)!=zin.read(n)]
assert set(diff)<={sf[t] for t in changed}, set(diff)
os.replace(tmp,P)
print(f"patched {os.path.basename(P)}: changed parts {sorted(diff)}; all other parts byte-identical. backup {os.path.basename(bak)}")
