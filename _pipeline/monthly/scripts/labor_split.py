#!/usr/bin/env python3
"""labor_split.py <start YYYY-MM-DD> <end YYYY-MM-DD> [--json]
Per-cafe Core / Admin / Training WAGES for a date range (inclusive), from the weekly Square
labor pulls (_pipeline/weekly/data/labor_*.json, role_breakdown by Square job title).
  Admin    = job title contains "Manager"
  Training = job title contains "Training"
  Core     = everything else
Used at month close to split the Payroll JE's FOH Labor line per location:
  FOH Labor (core) / Shared Management Labor @ location (admin) / Training Labor @ location (training)
The ADP gross stays the control total; admin+training $ come from Square, FOH = ADP gross - admin - training.
Prints a table; --json emits {'L3':{'admin':..,'training':..}, ...} for build_jes.py CONFIG.
"""
import sys, json, glob, os, datetime
HERE=os.path.dirname(os.path.abspath(__file__))
DATA=os.path.normpath(os.path.join(HERE,'..','..','weekly','data'))
KEY={'Downtown':'L3','Studio':'L5','1899 Penn':'L6','NMWA':'L7'}
def main():
    a=[x for x in sys.argv[1:] if not x.startswith('--')]
    if len(a)!=2: sys.exit(__doc__)
    s=datetime.date.fromisoformat(a[0]); e=datetime.date.fromisoformat(a[1])
    acc={k:{'core':0,'admin':0,'training':0,'days':set()} for k in KEY}
    for f in glob.glob(os.path.join(DATA,'labor_*.json')):
        d=json.load(open(f))
        for cafe,days in d['per_cafe_per_day'].items():
            if cafe not in KEY: continue
            for day,v in days.items():
                dd=datetime.date.fromisoformat(day)
                if not (s<=dd<=e): continue
                if day in acc[cafe]['days']: continue   # overlapping pulls: count each cafe-day once
                acc[cafe]['days'].add(day)
                for role,rb in v.get('role_breakdown',{}).items():
                    b='admin' if 'manager' in role.lower() else 'training' if 'training' in role.lower() else 'core'
                    acc[cafe][b]+=rb['cents']
    out={}; missing=[]
    ndays=(e-s).days+1
    print(f"Labor split {s} → {e} ({ndays} days)\n{'Cafe':10} {'Core':>10} {'Admin':>9} {'Train':>8} {'Total':>10}  days")
    for cafe,k in KEY.items():
        c=acc[cafe]; tot=c['core']+c['admin']+c['training']
        print(f"{cafe:10} {c['core']/100:10,.2f} {c['admin']/100:9,.2f} {c['training']/100:8,.2f} {tot/100:10,.2f}  {len(c['days'])}")
        out[k]={'admin':round(c['admin']/100,2),'training':round(c['training']/100,2),'square_total':round(tot/100,2),'days_found':len(c['days'])}
        expected = sum(1 for i in range(ndays) if (s+datetime.timedelta(i)).weekday()<5) if cafe in ('Downtown','1899 Penn') else ndays
        if cafe!='NMWA' and len(c['days'])<expected-1: missing.append(f"{cafe} ({len(c['days'])}/{expected})")
    if missing: print(f"⚠ coverage gaps (fewer cafe-days than range): {missing} — run the weekly pull for the missing weeks first")
    if '--json' in sys.argv: print(json.dumps(out,indent=1))
if __name__=='__main__': main()
