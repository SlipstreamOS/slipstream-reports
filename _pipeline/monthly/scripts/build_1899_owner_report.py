#!/usr/bin/env python3
"""1899 Penn owner (MRP) monthly report -> PowerPoint in Ryan's Canva layout.
Usage (on-device):  python3 build_1899_owner_report.py <model_version> <close YYYY-MM>
  e.g. python3 build_1899_owner_report.py 18 2026-09
Reads  Slipstream Consolidated Model 2026 v<N>.xlsx -> 'Income Statement - 1899' (cached values; recalc first if freshly synced).
Notes  1899 Owner Report/notes_<YYYY-MM>.json  {headline, model_notes[], outlook[], preview_title, preview[], next_steps[]}
       (nested bullets: use a list inside the list).  Missing file -> placeholder text.
Writes 1899 Owner Report/Slipstream x 1899 Penn Cafe Update <Issue Month YYYY>.pptx  (issue month = month after the close month)."""
import sys, os, json, calendar, datetime, warnings
warnings.filterwarnings("ignore")
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

VER=sys.argv[1]; CLOSE=sys.argv[2]; YEAR,MO=int(CLOSE[:4]),int(CLOSE[5:7])
ROOT=os.path.abspath(os.path.join(os.path.dirname(__file__),"..","..",".."))
OUTDIR=os.path.join(ROOT,"1899 Owner Report"); AS=os.path.join(OUTDIR,"assets")
MODEL=os.environ.get("MODEL_XLSX") or os.path.join(ROOT,f"Slipstream Consolidated Model 2026 v{VER}.xlsx")   # pass MODEL_XLSX=<recalculated copy> when the model was just synced (the real file has stale cached values until opened in Excel)
close_name=f"{calendar.month_name[MO]} {YEAR}"
iy,im=(YEAR+1,1) if MO==12 else (YEAR,MO+1); issue_name=f"{calendar.month_name[im]} {iy}"

# ---------- data ----------
ws=openpyxl.load_workbook(MODEL,data_only=True)["Income Statement - 1899"]
hdr={}
for c in ws[14]:
    if isinstance(c.value,datetime.datetime) and c.value.year==YEAR: hdr[c.value.month]=c.column
    if c.value==f"{YEAR} FY": FYCOL=c.column
assert len(hdr)==12, "could not find 12 month columns for the year on row 14"
lab={}
for r in range(15,ws.max_row+1):
    v=ws.cell(r,2).value
    if isinstance(v,str): lab.setdefault(v.strip(),r)
def row(label): return lab[label]
def val(r,m):
    v=ws.cell(r,hdr[m]).value
    return float(v) if isinstance(v,(int,float)) else 0.0
def series(label,add=()):
    rs=[row(label)]+[row(a) for a in add]
    return [sum(val(r,m) for r in rs) for m in range(1,13)]
S=dict(
 bev=series('Coffee Beverage Sales'), retail=series('Retail Coffee Sales'), coffee=series('Total Coffee Sales'),
 food=series('Food Sales'), cater=series('Catering & Event Sales'), disc=series('Discounts',('Refunds & Voids',)),
 sales=series('Total Core Location Sales'), cogs=series('Total COGS'), gp=series('Gross Profit'), labor=series('Total Labor'),
 var=series('Total Variable Operating Expenses'), fixed=series('Total Fixed Operating Expenses'),
 other=series('Total Other Operating Expenses',('Total Occupancy Expenses',)), exp=series('Total Expenses'), noi=series('Net Operating Income'),
 tx=series('Weekday avg daily transactions'), tk=series('Weekday avg ticket'))
A=list(range(MO))            # actual month indexes
# The model's Total Expenses / NOI rows do not include the add-on rows under Net Income (Printing and postage, Marketing Expense).
# Fold them in here so NOI ties to QuickBooks (March and April 2026 differ otherwise).
_extra=[sum(val(lab[l],m) for l in ('Printing and postage','Marketing Expense') if l in lab) for m in range(1,13)]
S['other']=[S['other'][i]+_extra[i] for i in range(12)]; S['exp']=[S['exp'][i]+_extra[i] for i in range(12)]; S['noi']=[S['noi'][i]-_extra[i] for i in range(12)]
def ytd(k): return sum(S[k][i] for i in A)
def fy(k): return sum(S[k])
def incentive(sales,noi): return max(0.0,min(0.05*sales,noi))
S['inc']=[incentive(S['sales'][i],S['noi'][i]) for i in range(12)]
S['own']=[S['noi'][i]-S['inc'][i] for i in range(12)]

# ---------- notes ----------
npath=os.path.join(OUTDIR,f"notes_{CLOSE}.json")
N=json.load(open(npath)) if os.path.exists(npath) else {}
PH="[Ryan to write]"
headline=N.get("headline",PH); model_notes=N.get("model_notes",[PH]); outlook=N.get("outlook",[PH])
preview_title=N.get("preview_title",f"{issue_name} Preview"); preview=N.get("preview",[PH]); next_steps=N.get("next_steps",[PH])

# ---------- deck ----------
BAR=RGBColor(0x98,0x93,0xA4); PURPLE=RGBColor(0x5F,0x41,0x70); GREY=RGBColor(0x59,0x59,0x59); SLATE=RGBColor(0x4A,0x6A,0x82)
INK=RGBColor(0x22,0x22,0x22); PROJ=RGBColor(0xEF,0xEE,0xF3); PROJTXT=RGBColor(0x55,0x55,0x66); RULE=RGBColor(0x40,0x40,0x40)
FONT="Museo 500"   # body (Ryan supplied Museo500-Regular.otf, kept in _pipeline/fonts); headings use Noyh Black
HEAD_FONT="Noyh Black"
prs=Presentation(); prs.slide_width=Inches(11); prs.slide_height=Inches(8.5)
blank=prs.slide_layouts[6]
def pic(s,name,x,y,w,h): return s.shapes.add_picture(os.path.join(AS,name),Inches(x),Inches(y),Inches(w),Inches(h))
def tb(s,x,y,w,h,text,size,color=INK,bold=False,align=PP_ALIGN.LEFT,spacing=None,italic=False):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=t.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    lines=text if isinstance(text,list) else [text]
    for i,l in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        r=p.add_run(); r.text=l; f=r.font; f.size=Pt(size); f.color.rgb=color; f.bold=bold; f.italic=italic; f.name=FONT
        if spacing: r._r.get_or_add_rPr().set('spc',str(spacing))
    return t
def bullets(s,x,y,w,h,items,size=13,color=INK):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=t.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    first=[True]
    def add(it,lvl):
        if isinstance(it,list):
            for j in it: add(j,lvl+1)
            return
        p=tf.paragraphs[0] if first[0] else tf.add_paragraph(); first[0]=False
        ind=0.25+0.28*lvl
        pPr=p._p.get_or_add_pPr(); pPr.set('marL',str(int(Inches(ind)))); pPr.set('indent',str(-int(Inches(0.2))))
        bu=pPr.makeelement(qn('a:buChar'),{'char':'•' if lvl==0 else '◦'}); pPr.append(bu)
        p.space_after=Pt(5)
        r=p.add_run(); r.text=it; r.font.size=Pt(size if lvl==0 else size-1); r.font.color.rgb=color; r.font.name=FONT
    for it in items: add(it,0)
    return t
def content_slide(label,title,subtitle,num):
    s=prs.slides.add_slide(blank)
    pic(s,"brick_soft.jpg",0,0,11,8.5)
    b=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(11),Inches(1.0)); b.fill.solid(); b.fill.fore_color.rgb=BAR; b.line.fill.background(); b.shadow.inherit=False
    ln=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(1.05),Inches(11),Emu(19050)); ln.fill.solid(); ln.fill.fore_color.rgb=PURPLE; ln.line.fill.background(); ln.shadow.inherit=False
    tb(s,0.62,0.52,6,0.25,label.upper(),10.5,PURPLE,bold=True,spacing=300)
    tb(s,0.5,1.12,10,0.65,title,34,GREY)
    if subtitle: tb(s,0.63,1.72,9.8,0.3,subtitle,14,GREY)
    tb(s,10.2,8.03,0.3,0.2,str(num),9,PURPLE,bold=True,align=PP_ALIGN.RIGHT)
    return s
def money(v):
    v=round(v); return "–" if v==0 else (f"({abs(v):,})" if v<0 else f"{v:,}")
def pct(n,d): return "" if not d else f"{n/d*100:.1f}%"
def set_cell(c,text,size=7.5,bold=False,italic=False,color=INK,align=PP_ALIGN.RIGHT,fill=None,top=False):
    c.margin_left=c.margin_right=Emu(27432); c.margin_top=c.margin_bottom=Emu(9144); c.vertical_anchor=MSO_ANCHOR.MIDDLE
    tf=c.text_frame; p=tf.paragraphs[0]; p.alignment=align
    r=p.add_run(); r.text=text if text else '\u00a0'; f=r.font; f.size=Pt(size); f.bold=bold; f.italic=italic; f.color.rgb=color; f.name=FONT
    e=p._p.makeelement(qn('a:endParaRPr'),{'sz':str(int(size*100))}); p._p.append(e)
    if fill is None: c.fill.background()
    else: c.fill.solid(); c.fill.fore_color.rgb=fill
    tcPr=c._tc.get_or_add_tcPr()
    for tag in ('a:lnL','a:lnR','a:lnT','a:lnB'):
        e=tcPr.makeelement(qn(tag),{'w':'9525' if (tag=='a:lnT' and top) else '0'})
        if tag=='a:lnT' and top:
            sf=e.makeelement(qn('a:solidFill'),{}); clr=sf.makeelement(qn('a:srgbClr'),{'val':'404040'}); sf.append(clr); e.append(sf)
        else: e.append(e.makeelement(qn('a:noFill'),{}))
        tcPr.append(e)
def table(s,x,y,w,rows,labw=2.25,rh=0.2):
    """rows: list of (label, [12 monthly strings], ytd, fy, style) ; style keys: bold, italic, top, gap, head"""
    ncol=15; t=s.shapes.add_table(len(rows),ncol,Inches(x),Inches(y),Inches(w),Inches(rh*len(rows))).table
    tbl=t._tbl; tblPr=tbl.tblPr
    for a in ('firstRow','bandRow'): tblPr.set(a,'0')
    sid=tblPr.find(qn('a:tableStyleId'))
    if sid is not None: sid.text='{2D5ABB26-0587-4C30-8999-92F81FD0307C}'   # No Style, No Grid
    cw=(w-labw)/14
    t.columns[0].width=Inches(labw)
    for i in range(1,ncol): t.columns[i].width=Inches(cw)
    for ri,(label,mvals,yv,fv,st) in enumerate(rows):
        gap=st.get('gap',False); sz=1 if gap else 7.5
        t.rows[ri].height=Inches(rh if not gap else rh*0.45)
        b=st.get('bold',False); it=st.get('italic',False); top=st.get('top',False); head=st.get('head',False)
        set_cell(t.cell(ri,0),label,size=sz,bold=b or head,italic=it,align=PP_ALIGN.LEFT,top=top)
        for m in range(12):
            proj=m>=MO
            set_cell(t.cell(ri,m+1),mvals[m] if mvals else "",size=sz,bold=b or head,italic=it or (proj and not head),
                     color=PROJTXT if proj else INK,fill=PROJ if proj else None,top=top)
        set_cell(t.cell(ri,13),yv,size=sz,bold=True if (b or head) else False,italic=it,top=top)
        set_cell(t.cell(ri,14),fv,size=sz,bold=b or head,italic=True if not head else False,color=PROJTXT,fill=PROJ,top=top)
    return t
mon=[calendar.month_abbr[m]+"-"+str(YEAR)[2:] for m in range(1,13)]
def R(label,key,**st): return (label,[money(v) for v in S[key]],money(ytd(key)),money(fy(key)),st)
def P(label,key,**st):
    st.setdefault('italic',True)
    return (label,[pct(S[key][i],S['sales'][i]) for i in range(12)],pct(ytd(key),ytd('sales')),pct(fy(key),fy('sales')),st)
GAP=("",None,"","",{'gap':True})
days=[calendar.monthrange(YEAR,m)[1] for m in range(1,13)]
txrow=("Avg Daily Transactions (weekday)",[f"{v:.0f}" if v else "" for v in S['tx']],"","",{})
tkrow=("   Avg Ticket",[f"{v:.2f}" if v else "" for v in S['tk']],"","",{'italic':True})
_oa_p=os.path.join(OUTDIR,"owner_account.json"); OA0=json.load(open(_oa_p)) if os.path.exists(_oa_p) else None
def _balrow():
    if not OA0: return None
    vals=[]; b=None
    for m in range(1,13):
        if m<=MO: b=sum(e["amount"] for e in OA0["entries"] if e["date"][:7]<=f"{YEAR}-{m:02d}")
        else: b=b+S['own'][m-1]
        vals.append(money(b))
    cur=sum(e["amount"] for e in OA0["entries"] if e["date"][:7]<=CLOSE)
    return ("MRP Account Balance (end of month)",vals,money(cur),vals[-1],{'bold':True})
fin=[("($, except as noted)",[("Proj " if m>=MO else "")+mon[m] if False else mon[m] for m in range(12)],"YTD Actual",f"{YEAR} Proj",{'head':True}),
 ("",["Actual" if m<MO else "Projected" for m in range(12)],f"Jan–{calendar.month_abbr[MO]}","Full year",{'italic':True}),
 ("Sales",None,"","",{'bold':True,'top':True}),
 R("   Coffee Beverage Sales",'bev'),R("   Retail Coffee Sales",'retail'),R("  Total Coffee Sales",'coffee',top=True),
 R("  Total Food Sales",'food'),R("  Total Catering and Events Sales",'cater'),R("  Discounts/Voids/Refunds",'disc'),
 R("Total Sales",'sales',bold=True,top=True),GAP,txrow,tkrow,GAP,
 R("Total COGS",'cogs',bold=True,top=True),P("   % of Sales",'cogs'),
 R("Gross Profit",'gp',bold=True,top=True),GAP,
 R("Total Labor",'labor',bold=True,top=True),P("   % of Sales",'labor'),GAP,
 R("  Variable Operating Expenses",'var'),R("  Fixed Operating Expenses",'fixed'),R("  Other and Occupancy Expenses",'other'),
 R("Total Expenses",'exp',bold=True,top=True),P("   % of Sales",'exp'),GAP,
 R("Net Operating Income (before mgmt fee)",'noi',bold=True,top=True),P("   % of Sales",'noi'),GAP,
 R("Share of Profit to MRP",'own',bold=True,top=True)]
if _balrow(): fin.append(_balrow())

# 1 cover
s=prs.slides.add_slide(blank)
pic(s,"brick_soft.jpg",0,0,11,8.5); pic(s,"cover_collage.jpg",0,0,3.96,8.5); pic(s,"logo_2026_color.png",5.6,0.45,3.3,1.452)
tb(s,4.2,2.4,6.1,0.75,["1899 Penn Cafe Update",issue_name],17,INK,bold=True,align=PP_ALIGN.CENTER,spacing=200)
# 2 financials
s=content_slide("Monthly Summary",f"Update through {close_name}",headline,2)
table(s,0.5,2.15,10.0,fin,rh=0.19)
_res=float(OA0.get("reserve",0)) if OA0 else 0
tb(s,0.5,7.88,9.5,0.35,"Shaded columns are projections from the operating model. Actuals are from the closed books. Share of Profit to MRP is net operating income less the incentive fee (5% of sales, capped at net operating income)."+(f" Under Section 5(a) of the Operating Agreement, the first ${_res:,.0f} of the MRP balance remains in the Operating Account to fund operations." if _res else ""),7.5,PROJTXT,italic=True)
# 3 owner summary + notes
s=content_slide("Notes","Model Notes and Outlook",f"Year to date through {close_name}",3)
own=[("($)",mon,"YTD",f"{YEAR} Proj",{'head':True}),
 R("Total Sales",'sales',top=True),R("Net Operating Income (before mgmt fee)",'noi',bold=True),
 R("  Slipstream Incentive Fee (5% of sales, capped at NOI)",'inc'),R("Owner Share of Profit / (Loss)",'own',bold=True,top=True)]
tb(s,0.6,2.4,4.6,0.3,"Model Notes",16,SLATE); bullets(s,0.6,2.8,4.6,4.5,model_notes,13,SLATE)
tb(s,5.75,2.4,4.6,0.3,"Outlook",16,SLATE); bullets(s,5.75,2.8,4.6,2.2,outlook,13,SLATE)
_ns=N.get("next_steps_p3",[])
if _ns: tb(s,5.75,5.0,4.6,0.3,"Next Steps and Observations",16,SLATE); bullets(s,5.75,5.4,4.6,2.4,_ns,13,SLATE)
# 4 owner account (QuickBooks 24700 Due to Owner), by month since inception
num_next=4
oa_path=os.path.join(OUTDIR,"owner_account.json")
if os.path.exists(oa_path) and json.load(open(oa_path)).get("show_account_page",False):   # Ryan Oct 5 2026: owner account page off by default
    OA=json.load(open(oa_path)); ents=[e for e in OA["entries"] if e["date"][:7]<=CLOSE]
    first=min(e["date"][:7] for e in ents); y,m=int(first[:4]),int(first[5:7]); rowsO=[]; bal=0.0; tot=dict(contribution=0.0,startup=0.0,share=0.0)
    while (y,m)<=(YEAR,MO):
        k=f"{y}-{m:02d}"; v=dict(contribution=0.0,startup=0.0,share=0.0)
        for e in ents:
            if e["date"][:7]==k: v[e["kind"]]+=e["amount"]; tot[e["kind"]]+=e["amount"]
        bal+=sum(v.values()); rowsO.append((f"{calendar.month_abbr[m]} {y}",v,bal))
        m+=1
        if m>12: y+=1; m=1
    def m2(x): return "–" if abs(x)<0.005 else (f"({abs(x):,.2f})" if x<0 else f"{x:,.2f}")
    s=content_slide("Owner Account","Owner Account by Month",f"From first activity through {close_name}",4); num_next=5
    hdrs=["Month","Owner contributions","Startup costs","Share of profit / (loss)","Balance owed to owner"]
    n=len(rowsO)+2; widths=[1.15,1.6,1.35,1.75,1.65]; rh=min(0.235,4.9/n)
    t=s.shapes.add_table(n,5,Inches(0.6),Inches(2.2),Inches(sum(widths)),Inches(rh*n)).table
    tblPr=t._tbl.tblPr
    for a_ in ('firstRow','bandRow'): tblPr.set(a_,'0')
    sid=tblPr.find(qn('a:tableStyleId'))
    if sid is not None: sid.text='{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    for i,w in enumerate(widths): t.columns[i].width=Inches(w)
    def rowset(ri,cells,bold=False,top=False):
        t.rows[ri].height=Inches(rh)
        for ci,tx in enumerate(cells): set_cell(t.cell(ri,ci),tx,size=8.5,bold=bold,align=PP_ALIGN.LEFT if ci==0 else PP_ALIGN.RIGHT,top=top)
    rowset(0,hdrs,bold=True)
    for ri,(lab_,v,b) in enumerate(rowsO,1): rowset(ri,[lab_,m2(v["contribution"]),m2(v["startup"]),m2(v["share"]),m2(b)],top=(ri==1))
    rowset(n-1,["Total",m2(tot["contribution"]),m2(tot["startup"]),m2(tot["share"]),m2(bal)],bold=True,top=True)
    res=float(OA.get("reserve",0)); avail=max(0.0,bal-res)
    kx=8.45; tb(s,kx,2.25,2.1,0.3,"Balance summary",14,SLATE)
    lines=[("Balance owed to MRP",f"${bal:,.2f}"),("Held in Operating Account (up to)",f"${res:,.2f}"+("" if OA.get("reserve_confirmed") else " *")),("Available to distribute",f"${avail:,.2f}")]
    yy=2.7
    for lab_,v_ in lines:
        tb(s,kx,yy,2.1,0.22,lab_,9.5,MUTED if False else PROJTXT); tb(s,kx,yy+0.2,2.1,0.3,v_,15,INK,bold=True); yy+=0.72
    if not OA.get("reserve_confirmed"): tb(s,kx,yy+0.05,2.0,0.6,"* Minimum balance to be confirmed against the management agreement.",8,PROJTXT,italic=True)
    else: tb(s,kx,yy+0.05,2.0,1.2,f"Per Section 5(a) of the Operating Agreement, the first ${res:,.0f} remains in the Operating Account to fund operations. The balance owed is below that amount, so no distribution is due this month.",8,PROJTXT,italic=True)
    tb(s,0.6,7.72,9.8,0.3,"Contributions are deposits and loss coverage payments received from the owner. Share of profit / (loss) is net operating income less the incentive fee.",8.5,PROJTXT,italic=True)
    print("owner account balance",round(bal,2),"totals",{k:round(v,2) for k,v in tot.items()})
# 4 next steps
if N.get("show_next_steps_page",False):   # Ryan Oct 5 2026: next steps now live on page 3
    s=content_slide("Next Steps",f"{calendar.month_name[im]} Plans and Observations",N.get("next_headline",PH),num_next)
    tb(s,0.85,2.35,9.5,0.3,preview_title,15,SLATE); bullets(s,0.85,2.72,9.4,1.6,preview,12.5)
    tb(s,0.85,4.45,9.5,0.3,"Observations and Proposed Next Steps",15,SLATE); bullets(s,0.85,4.82,9.4,3.0,next_steps,12.5)
# 5 thanks
s=prs.slides.add_slide(blank)
pic(s,"brick_soft.jpg",0,0,11,8.5)
b=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(11),Inches(1.0)); b.fill.solid(); b.fill.fore_color.rgb=BAR; b.line.fill.background(); b.shadow.inherit=False
tb(s,0.62,0.52,6,0.25,"SLIPSTREAM",10.5,PURPLE,bold=True,spacing=300)
tb(s,0.5,1.12,10,0.65,"Thank you.",34,GREY)
pic(s,"thanks_photos.jpg",0,1.99,11,6.51)
def apply_fonts(prs):
    """Headings (15pt+ bold, or 30pt+) -> Noyh Black. Everything else -> Museo 500 (bold/italic flags kept)."""
    def fix(r):
        f=r.font; sz=f.size.pt if f.size else 0; head=(sz>=15 and (bool(f.bold) or sz>=30)) or (bool(f.bold) and r._r.get_or_add_rPr().get('spc'))
        name=HEAD_FONT if head else FONT
        f.name=name
        if head: f.bold=False
        rPr=r._r.get_or_add_rPr()
        for tag in ('a:ea','a:cs'):
            e=rPr.find(qn(tag))
            if e is None: e=rPr.makeelement(qn(tag),{}); rPr.append(e)
            e.set('typeface',name)
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs: fix(r)
            if getattr(sh,'has_table',False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            for r in p.runs: fix(r)
apply_fonts(prs)
out=os.path.join(OUTDIR,f"Slipstream x 1899 Penn Cafe Update {issue_name}.pptx")
prs.save(out); print("wrote",out)
print("YTD sales",round(ytd('sales'),2),"NOI",round(ytd('noi'),2),"incentive",round(ytd('inc'),2),"owner",round(ytd('own'),2))
print("monthly NOI",[round(v) for v in S['noi']]); print("incentive",[round(v) for v in S['inc']]); print("owner",[round(v) for v in S['own']])
print("labor%",[round(S['labor'][i]/S['sales'][i]*100,1) for i in range(12)]); print("cogs%",[round(S['cogs'][i]/S['sales'][i]*100,1) for i in range(12)])
