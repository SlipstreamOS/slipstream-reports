#!/usr/bin/env python3
"""accrual_from_timecards.py <month_end YYYY-MM-DD> <straddle_start YYYY-MM-DD> <prior_pay_JE.xlsx> <prior_start> <prior_end>
Payroll accrual for the four cafes, by work date (Ryan, Nov 7 2026).
Use ONLY when the payroll for the period that crosses month-end has not been processed by close.
If it has (e.g. October 2026: 10/18-10/31 paid 11/6, run 11/3), book the actual check instead.

Method (back-tested Jul-Oct 2026, cut estimate error ~87% vs the old proration):
  unpaid days        = straddle_start .. month_end
  base               = Square timecard wages for those days (weekly labor pulls) - card tips for those days (daily_metrics)
  k                  = prior full check's (FOH + admin + training + payroll tax) / (timecard wages - card tips) for its period
  accrual (FOH+tax)  = base x k, split into FOH and Tax using the prior check's tax share
Prints the L3/L5/L6/L7 and TL3/TL5/TL6/TL7 values for build_jes.py ACCR. Comm, SRP and Mgmt lines are unchanged.
"""
import sys, os, json, glob, csv, datetime as dt, openpyxl
HERE=os.path.dirname(os.path.abspath(__file__)); PIPE=os.path.dirname(os.path.dirname(HERE))
LOC={'GHC3 - Downtown':('Downtown','3'),'GHC5 - Studio':('Studio','5'),'GHC6 - 1899':('1899 Penn','6'),'NMWA':('NMWA','7')}
def je(path):
    ws=openpyxl.load_workbook(path,data_only=True).active; out={}
    for r in ws.iter_rows(min_row=2,values_only=True):
        acct,deb,cred,desc,name,loc=(list(r)+[None]*6)[:6]
        if loc not in LOC or not acct: continue
        c=LOC[loc][0]; o=out.setdefault(c,{'pay':0.0,'tax':0.0}); v=(deb or 0)-(cred or 0)
        if 'Payroll Tax' in acct: o['tax']+=v
        elif any(k in acct for k in ['FOH Labor','Shared Management Labor','Training Labor']): o['pay']+=v
    return out
def main():
    if len(sys.argv)!=6: sys.exit(__doc__)
    me,ss,pje,ps,pe=sys.argv[1:]
    days={}
    for f in glob.glob(os.path.join(PIPE,'weekly','data','labor_*.json')):
        for c,dd in json.load(open(f))['per_cafe_per_day'].items():
            for d,v in dd.items(): days.setdefault((c,d),v['labor_cents']/100)
    tips={(r['cafe'],r['date']):float(r['tips'] or 0) for r in csv.DictReader(open(os.path.join(PIPE,'shared','daily_metrics.csv')))}
    def span(c,s,e,src):
        s=dt.date.fromisoformat(s); e=dt.date.fromisoformat(e); n=(e-s).days+1
        vals=[src.get((c,(s+dt.timedelta(i)).isoformat())) for i in range(n)]
        return sum(v or 0 for v in vals), sum(1 for v in vals if v is not None)
    prior=je(pje); out={}
    print(f"Accrual {ss} -> {me} (prior check {ps} -> {pe})")
    for loc,(c,n) in LOC.items():
        w,wd=span(c,ps,pe,days); t,_=span(c,ps,pe,tips)
        p=prior.get(c,{'pay':0,'tax':0}); cost=p['pay']+p['tax']
        k=cost/(w-t) if w-t>0 else 0; taxshare=p['tax']/cost if cost else 0
        wi,wdi=span(c,ss,me,days); ti,_=span(c,ss,me,tips)
        acc=max(wi-ti,0)*k
        out[f'L{n}']=round(acc*(1-taxshare),2); out[f'TL{n}']=round(acc*taxshare,2)
        print(f"  {c:10} days w/ timecards {wdi:2}  wages {wi:9,.2f}  tips {ti:8,.2f}  k {k:.3f}  accrual {acc:9,.2f}  (FOH {out[f'L{n}']:,.2f} / tax {out[f'TL{n}']:,.2f})")
    print(json.dumps(out))
if __name__=='__main__': main()
