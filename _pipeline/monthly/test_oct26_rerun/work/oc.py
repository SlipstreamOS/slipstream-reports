import csv,sys,glob,collections
def m(s):
    s=s.replace('$','').replace(',','')
    try: return float(s)
    except: return 0.0
for f in sorted(glob.glob(sys.argv[1]+'/*.csv')):
    rows=list(csv.reader(open(f,encoding='utf-8-sig')))
    sec=None; cats=collections.defaultdict(float); oc=None; tender={}; svc=None
    for r in rows:
        if not r or not r[0]: sec=None; continue
        if r[0]=='Dining Option': sec='dine'; continue
        if r[0]=='Other Categories': oc=r
        if r[0]=='Services': svc=r
        if r[0] in ('Check','Other','Cash') and sec is None: tender[r[0]]=r[1]
        if sec=='dine' and len(r)>=4: cats[r[1]]+=m(r[3])
    known={'All-Day Food','Catering and Event','Espresso and Specialty','Filter Coffee','Pastry','Retail','Tea and Other Drinks','Services','Whole Bean','Merch'}
    print(f.split('/')[-1], '| OC gross', oc[2] if oc else None, 'net',oc[6] if oc else None,'| Services',svc[2] if svc else None, tender)
    for k,v in sorted(cats.items()):
        if k not in known: print('   ',k, round(v,2))
