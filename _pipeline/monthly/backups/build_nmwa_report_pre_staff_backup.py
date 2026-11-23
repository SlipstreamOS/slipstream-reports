#!/usr/bin/env python3
"""NMWA cafe monthly report -> PowerPoint, same look as the 1899 owner report (short version).
Usage (on-device):  python3 build_nmwa_report.py <close YYYY-MM>     e.g. python3 build_nmwa_report.py 2026-09
Data   _pipeline/shared/transactions/<YYYY-MM>.csv (and the two prior months) - walk-in only, invoices excluded.
Notes  NMWA Report/notes_<YYYY-MM>.json  {community_days[], open_from, closed_days{date:reason}, headline, notes[], events[], asks[]}
Guests NMWA Report/guest_counts_<YYYY-MM>.csv  (optional; columns: date,guests  with date as YYYY-MM-DD).
       When present the guest and conversion columns fill in; when absent they show as pending.
Writes NMWA Report/Slipstream x NMWA Cafe Update <Issue Month YYYY>.pptx  (issue month = month after the close month)."""
import sys, os, json, csv, calendar, datetime, collections
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
CLOSE=sys.argv[1]; YEAR,MO=int(CLOSE[:4]),int(CLOSE[5:7])
ROOT=os.path.abspath(os.path.join(os.path.dirname(__file__),"..","..",".."))
OUTDIR=os.path.join(ROOT,"NMWA Report"); AS=os.path.join(ROOT,"1899 Owner Report","assets"); TX=os.path.join(ROOT,"_pipeline","shared","transactions")
close_name=f"{calendar.month_name[MO]} {YEAR}"
iy,im=(YEAR+1,1) if MO==12 else (YEAR,MO+1); issue_name=f"{calendar.month_name[im]} {iy}"
N=json.load(open(os.path.join(OUTDIR,f"notes_{CLOSE}.json")))
COMM=set(N.get("community_days",[])); OPEN_FROM=N.get("open_from","2026-07-09"); CLOSED=N.get("closed_days",{})
PH="[Ryan to write]"
def money(x):
    x=str(x).replace("$","").replace(",","").strip()
    try: return float(x)
    except: return 0.0
def _rows(y,m):
    # Sun-Sat weekly bundles spill a month's first days into the prior month's file (and its last days into the next), so read all three and dedupe
    seen=set()
    for k in (-1,0,1):
        yy,mm=y,m+k
        if mm<1: yy,mm=yy-1,12
        if mm>12: yy,mm=yy+1,1
        p=os.path.join(TX,f"{yy}-{mm:02d}.csv")
        if not os.path.exists(p): continue
        for r in csv.DictReader(open(p,newline="",encoding="utf-8-sig")):
            key=(r.get("Transaction ID"),r.get("Payment ID"),r.get("Event Type"))
            if key in seen: continue
            seen.add(key); yield r
def load(y,m):
    d=collections.OrderedDict()
    for r in _rows(y,m):
        if r.get("Location")!="Slipstream NMWA" or r.get("Event Type")!="Payment": continue
        if "Invoice" in (r.get("Source") or ""): continue          # walk-in only (Ryan rule)
        if r["Date"]<OPEN_FROM or r["Date"][:7]!=f"{y}-{m:02d}": continue   # the month file can carry spillover days from the next month
        v=d.setdefault(r["Date"],[0,0.0]); v[0]+=1; v[1]+=money(r["Net Sales"])
    return collections.OrderedDict(sorted(d.items()))
def stats(d):
    reg={k:v for k,v in d.items() if k not in COMM}; com={k:v for k,v in d.items() if k in COMM}
    tx=sum(v[0] for v in d.values()); net=sum(v[1] for v in d.values())
    return dict(days=len(d),tx=tx,net=net,ticket=net/tx if tx else 0,
      regdays=len(reg),regnet=sum(v[1] for v in reg.values()),regtx=sum(v[0] for v in reg.values()),
      comdays=len(com),comnet=sum(v[1] for v in com.values()),comtx=sum(v[0] for v in com.values()))
months=[]
for k in (2,1,0):
    y,m=YEAR,MO-k
    while m<1: y-=1; m+=12
    d=load(y,m)
    if d:
        st=stats(d); G={}
        _gp=os.path.join(OUTDIR,f"guest_counts_{y}-{m:02d}.csv")
        if os.path.exists(_gp):
            for r in csv.DictReader(open(_gp,newline="",encoding="utf-8-sig")):
                try: G[r["date"].strip()]=int(float(str(r["guests"]).replace(",","")))
                except Exception: pass
        st["gopen"]=sum(G[k] for k in d if k in G) if G else None          # museum guests on days the cafe sold
        st["gtx"]=sum(d[k][0] for k in d if k in G) if G else None
        months.append((calendar.month_name[m],d,st))
cur=months[-1][1]; cs=months[-1][2]
guests={}; comps={}   # comps = cafe-only visitors NMWA let in without a museum ticket (optional cafe_comps column)
gp=os.path.join(OUTDIR,f"guest_counts_{CLOSE}.csv")
if os.path.exists(gp):
    for r in csv.DictReader(open(gp,newline="",encoding="utf-8-sig")):
        try: guests[r["date"].strip()]=int(float(str(r["guests"]).replace(",","")))
        except Exception: pass
        try:
            cc=str(r.get("cafe_comps") or "").replace(",","").strip()
            if cc: comps[r["date"].strip()]=int(float(cc))
        except Exception: pass
# ---------- deck helpers (same look as build_1899_owner_report.py) ----------
BAR=RGBColor(0x98,0x93,0xA4); PURPLE=RGBColor(0x5F,0x41,0x70); GREY=RGBColor(0x59,0x59,0x59); SLATE=RGBColor(0x4A,0x6A,0x82)
INK=RGBColor(0x22,0x22,0x22); SHADE=RGBColor(0xEF,0xEE,0xF3); MUTED=RGBColor(0x55,0x55,0x66)
FONT="Museo 500"; HEAD_FONT="Noyh Black"
prs=Presentation(); prs.slide_width=Inches(11); prs.slide_height=Inches(8.5); blank=prs.slide_layouts[6]
def pic(s,name,x,y,w,h): return s.shapes.add_picture(os.path.join(AS,name),Inches(x),Inches(y),Inches(w),Inches(h))
def tb(s,x,y,w,h,text,size,color=INK,bold=False,align=PP_ALIGN.LEFT,spacing=None,italic=False):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=t.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    for i,l in enumerate(text if isinstance(text,list) else [text]):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        r=p.add_run(); r.text=l; f=r.font; f.size=Pt(size); f.color.rgb=color; f.bold=bold; f.italic=italic; f.name=FONT
        if spacing: r._r.get_or_add_rPr().set('spc',str(spacing))
    return t
def bullets(s,x,y,w,h,items,size=13,color=INK):
    t=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=t.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0; first=[True]
    def add(it,lvl):
        if isinstance(it,list):
            for j in it: add(j,lvl+1)
            return
        p=tf.paragraphs[0] if first[0] else tf.add_paragraph(); first[0]=False
        pPr=p._p.get_or_add_pPr(); pPr.set('marL',str(int(Inches(0.25+0.28*lvl)))); pPr.set('indent',str(-int(Inches(0.2))))
        pPr.append(pPr.makeelement(qn('a:buChar'),{'char':'•' if lvl==0 else '◦'})); p.space_after=Pt(6)
        r=p.add_run(); r.text=it; r.font.size=Pt(size if lvl==0 else size-1); r.font.color.rgb=color; r.font.name=FONT
    for it in items: add(it,0)
def content_slide(label,title,subtitle,num):
    s=prs.slides.add_slide(blank); pic(s,"brick_soft.jpg",0,0,11,8.5)
    b=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(11),Inches(1.0)); b.fill.solid(); b.fill.fore_color.rgb=BAR; b.line.fill.background(); b.shadow.inherit=False
    ln=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(1.05),Inches(11),Emu(19050)); ln.fill.solid(); ln.fill.fore_color.rgb=PURPLE; ln.line.fill.background(); ln.shadow.inherit=False
    tb(s,0.62,0.52,6,0.25,label.upper(),10.5,PURPLE,bold=True,spacing=300)
    tb(s,0.5,1.12,10,0.65,title,34,GREY)
    if subtitle: tb(s,0.63,1.72,9.8,0.3,subtitle,14,GREY)
    tb(s,10.2,8.03,0.3,0.2,str(num),9,PURPLE,bold=True,align=PP_ALIGN.RIGHT)
    return s
def set_cell(c,text,size=9,bold=False,italic=False,color=INK,align=PP_ALIGN.RIGHT,fill=None,top=False):
    c.margin_left=c.margin_right=Emu(45720); c.margin_top=c.margin_bottom=Emu(9144); c.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=c.text_frame.paragraphs[0]; p.alignment=align
    r=p.add_run(); r.text=text if text else ' '; f=r.font; f.size=Pt(size); f.bold=bold; f.italic=italic; f.color.rgb=color; f.name=FONT
    p._p.append(p._p.makeelement(qn('a:endParaRPr'),{'sz':str(int(size*100))}))
    if fill is None: c.fill.background()
    else: c.fill.solid(); c.fill.fore_color.rgb=fill
    tcPr=c._tc.get_or_add_tcPr()
    for tag in ('a:lnL','a:lnR','a:lnT','a:lnB'):
        e=tcPr.makeelement(qn(tag),{'w':'9525' if (tag=='a:lnT' and top) else '0'})
        if tag=='a:lnT' and top:
            sf=e.makeelement(qn('a:solidFill'),{}); sf.append(sf.makeelement(qn('a:srgbClr'),{'val':'404040'})); e.append(sf)
        else: e.append(e.makeelement(qn('a:noFill'),{}))
        tcPr.append(e)
def table(s,x,y,widths,rows,rh=0.24,size=9):
    """rows: list of (cells[], style) ; style: head, bold, top, fill, italic, muted"""
    t=s.shapes.add_table(len(rows),len(widths),Inches(x),Inches(y),Inches(sum(widths)),Inches(rh*len(rows))).table
    tblPr=t._tbl.tblPr
    for a in ('firstRow','bandRow'): tblPr.set(a,'0')
    sid=tblPr.find(qn('a:tableStyleId'))
    if sid is not None: sid.text='{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    for i,w in enumerate(widths): t.columns[i].width=Inches(w)
    for ri,(cells,st) in enumerate(rows):
        t.rows[ri].height=Inches(rh)
        for ci,txt in enumerate(cells):
            set_cell(t.cell(ri,ci),txt,size=size,bold=st.get('bold',False) or st.get('head',False),italic=st.get('italic',False),
                     color=MUTED if st.get('muted') else INK,align=PP_ALIGN.LEFT if ci==0 else PP_ALIGN.RIGHT,fill=SHADE if st.get('fill') else None,top=st.get('top',False))
    return t
def apply_fonts(prs):
    def fix(r):
        f=r.font; sz=f.size.pt if f.size else 0; head=(sz>=15 and (bool(f.bold) or sz>=30)) or (bool(f.bold) and r._r.get_or_add_rPr().get('spc'))
        name=HEAD_FONT if head else FONT; f.name=name
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
D=lambda v:f"${v:,.0f}"
# 1 cover
s=prs.slides.add_slide(blank); pic(s,"brick_soft.jpg",0,0,11,8.5); pic(s,"cover_collage.jpg",0,0,3.96,8.5); pic(s,"logo_2026_color.png",5.6,0.45,3.3,1.452)
tb(s,4.2,2.4,6.1,0.75,["NMWA Cafe Update",issue_name],17,INK,bold=True,align=PP_ALIGN.CENTER,spacing=200)
# 2 month at a glance
s=content_slide("Month at a Glance",f"{close_name} at the Cafe",N.get("headline",PH),2)
hdr=([""]+[m[0] for m in months],{'head':True})
def line(label,fn,**st): return ([label]+[fn(m[2]) for m in months],st)
rows=[hdr,
 line("Walk-in sales",lambda x:D(x['net']),top=True,bold=True),
 line("Open days",lambda x:str(x['days'])),
 line("Transactions",lambda x:f"{x['tx']:,}"),
 line("Average ticket",lambda x:f"${x['ticket']:.2f}"),
 line("Regular days",lambda x:str(x['regdays']),top=True),
 line("   Sales per regular day",lambda x:D(x['regnet']/x['regdays']) if x['regdays'] else "–"),
 line("   Transactions per regular day",lambda x:f"{x['regtx']/x['regdays']:.0f}" if x['regdays'] else "–"),
 line("Community days",lambda x:str(x['comdays']),top=True),
 line("   Sales per community day",lambda x:D(x['comnet']/x['comdays']) if x['comdays'] else "–"),
 line("   Transactions per community day",lambda x:f"{x['comtx']/x['comdays']:.0f}" if x['comdays'] else "–")]
if any(m[2].get("gopen") for m in months):
    rows+=[line("Guests on cafe days",lambda x:f"{x['gopen']:,}" if x.get('gopen') else "–",top=True),
           line("   Cafe conversion",lambda x:f"{x['gtx']/x['gopen']*100:.1f}%" if x.get('gopen') else "–")]
w0=2.5; wn=(5.3-w0)/len(months)
table(s,0.6,2.45,[w0]+[wn]*len(months),rows,rh=0.3,size=10.5)
tb(s,0.6,2.45+0.3*len(rows)+0.4,4.9,0.5,f"Walk-in cafe sales only. Catered events and museum invoices are not included. {months[0][0]} is counted from the July 9 opening." if months[0][0]=="July" else "Walk-in cafe sales only. Catered events and museum invoices are not included.",8.5,MUTED,italic=True)
_imp=N.get("improvements",[])
tb(s,5.95,2.35,4.5,0.3,"Notes",16,SLATE); bullets(s,5.95,2.75,4.5,2.4 if _imp else 4.8,N.get("notes",[PH]),12 if _imp else 12.5,SLATE)
if _imp: tb(s,5.95,5.15,4.5,0.6,N.get("improvements_title","Slipstream Improvements"),16,SLATE); bullets(s,5.95,5.85,4.5,1.7,_imp,12,SLATE)
# 3 guests and conversion
have=bool(guests)
s=content_slide("Guests and the Cafe","Daily Guests and Cafe Conversion",
   f"{close_name}, by day" if have else f"{close_name}, by day. Guest counts pending from NMWA",3)
days=sorted(set(list(cur.keys())+[k for k in CLOSED if k.startswith(CLOSE)]+[k for k in guests if k.startswith(CLOSE)]))
rows=[(["Date","Day","Museum guests","Cafe transactions","Conversion","Walk-in sales","Note"],{'head':True})]
tg=tt=0; tn=0.0
for k in days:
    dt=datetime.date.fromisoformat(k); tx,net=cur.get(k,[0,0.0]); g=guests.get(k)
    note="Community day" if k in COMM else CLOSED.get(k,"")
    if comps.get(k): note=(note+"; " if note else "")+f"{comps[k]} cafe-only"
    conv=f"{tx/g*100:.0f}%" if (g and tx) else ("" if have else "pending")
    rows.append(([dt.strftime("%b %-d"),dt.strftime("%a"),f"{g:,}" if g is not None else ("" if have else "pending"),str(tx) if tx else "–",conv,D(net) if tx else "–",note],
                 {'fill':k in COMM,'muted':k in CLOSED,'top':len(rows)==1}))
    if g and tx: tg+=g; tt+=tx
    tn+=net
alltx=sum(v[0] for v in cur.values())
rows.append((["Total","",f"{sum(guests.get(k,0) for k in days):,}" if have else "pending",f"{alltx:,}",f"{tt/tg*100:.0f}%" if tg else ("" if have else "pending"),D(tn),""],{'bold':True,'top':True}))
t=table(s,0.6,2.3,[1.0,0.7,1.5,1.7,1.3,1.4,2.2],rows,rh=min(0.255,5.35/len(rows)),size=9.5 if len(rows)<=22 else 8)
for ri in range(len(rows)):     # note column left-aligned
    t.cell(ri,6).text_frame.paragraphs[0].alignment=PP_ALIGN.LEFT
tb(s,0.6,7.72,9.8,0.3,"Conversion = cafe transactions divided by museum guests on days the cafe was open. Shaded rows are free community days."+(f" Cafe-only = visitors NMWA admitted for the cafe without a museum ticket ({sum(v for k,v in comps.items() if k.startswith(CLOSE))} this month)." if comps else "")+("" if have else " Guest and conversion columns fill in once NMWA sends the daily counts."),8.5,MUTED,italic=True)
# 4 next month schedule and events (calendar view) - falls back to the old events/asks page if notes has no "schedule"
SCH=N.get("schedule")
if SCH:
    sy,sm=int(SCH["month"][:4]),int(SCH["month"][5:7])
    s=content_slide("Schedule",f"{calendar.month_name[sm]} Schedule and Events",SCH.get("subtitle"),4)
    C_OPEN=RGBColor(0xFF,0xFF,0xFF); C_COMM=RGBColor(0xDD,0xD0,0xE6); C_CLOSED=RGBColor(0xC4,0xC1,0xCB); C_OFF=RGBColor(0xF1,0xF0,0xF4)
    C_BD=RGBColor(0xC2,0x72,0x2A); C_EVE=SLATE; C_LINE=RGBColor(0xB9,0xB5,0xC2); WHITE=RGBColor(0xFF,0xFF,0xFF)
    def box(x,y,w,h,fill,line=C_LINE):
        b=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); b.shadow.inherit=False
        if fill is None: b.fill.background()
        else: b.fill.solid(); b.fill.fore_color.rgb=fill
        if line is None: b.line.fill.background()
        else: b.line.color.rgb=line; b.line.width=Pt(0.75)
        return b
    gx,gy,cw=0.5,2.2,10.0/7; hh=0.26
    cal=calendar.Calendar(firstweekday=6).monthdatescalendar(sy,sm); ch=min(1.0,(7.5-gy-hh)/len(cal))
    for i,dn in enumerate(["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]):
        box(gx+i*cw,gy,cw,hh,BAR,None); tb(s,gx+i*cw,gy+0.055,cw,0.2,dn.upper(),7.5,WHITE,bold=True,align=PP_ALIGN.CENTER,spacing=120)
    comm=set(SCH.get("community",[])); closed=SCH.get("closed",{}); bd=set(SCH.get("breakdown",[])); eve=SCH.get("evening",{}); ev=SCH.get("events",{})
    openwd=set(SCH.get("open_weekdays",[3,4,5,6])); cnt=dict(open=0,comm=0,closed=0,bd=0)
    for wi,week in enumerate(cal):
        for di,dt in enumerate(week):
            x,y=gx+di*cw,gy+hh+wi*ch; k=dt.isoformat()
            if dt.month!=sm: box(x,y,cw,ch,C_OFF); continue
            is_open=(dt.weekday() in openwd or k in comm) and k not in closed
            if k in closed: fill,status,sc=C_CLOSED,closed[k],INK; cnt['closed']+=1
            elif k in comm: fill,status,sc=C_COMM,"Community Day",PURPLE; cnt['comm']+=1; cnt['open']+=1
            elif is_open: fill,status,sc=C_OPEN,SCH.get("open_label","Open 10 to 2"),INK; cnt['open']+=1
            else: fill,status,sc=C_OFF,"",MUTED
            box(x,y,cw,ch,fill)
            tb(s,x+0.07,y+0.05,0.4,0.2,str(dt.day),11,INK if (is_open or k in closed) else MUTED,bold=True)
            if status: tb(s,x+0.38,y+0.075,cw-0.44,0.18,status,8,sc,bold=True,align=PP_ALIGN.RIGHT)
            if ev.get(k): tb(s,x+0.07,y+0.33,cw-0.14,0.36,ev[k],7.5,MUTED)
            strip=None
            if k in bd and k not in closed: strip=(C_BD,"BREAK DOWN AFTER SERVICE"); cnt['bd']+=1
            elif k in eve: strip=(C_EVE,eve[k].upper())
            if strip: box(x,y+ch-0.2,cw,0.2,strip[0],None); tb(s,x,y+ch-0.165,cw,0.15,strip[1],6.5,WHITE,bold=True,align=PP_ALIGN.CENTER,spacing=60)
    ly=gy+hh+len(cal)*ch+0.13; lx=gx
    for fill,lab,wd in [(C_OPEN,SCH.get("open_label","Open 10 to 2"),0.95),(C_COMM,"Community day (open)",1.45),(C_CLOSED,"Closed for event",1.15),(C_BD,"Open, cafe broken down after service",2.45),(C_EVE,"Evening service",1.1),(C_OFF,"Not a cafe day",1.0)]:
        box(lx,ly,0.16,0.16,fill); tb(s,lx+0.22,ly+0.01,wd,0.18,lab,8,MUTED); lx+=wd+0.22
    tb(s,gx,ly+0.27,10,0.2,SCH.get("footnote",f"{cnt['open']} open days including {cnt['comm']} community days. {cnt['closed']} closed for an event. {cnt['bd']} days require a full cafe break down after service."),8.5,MUTED,italic=True)
    print("schedule",cnt)
elif N.get("show_events_page",True):   # Ryan Nov 6 2026: off when events/asks are covered in the meeting instead
    s=content_slide("Events and Next Steps",f"Events and {calendar.month_name[im]} Asks",None,4)
    tb(s,0.85,2.1,9.5,0.3,f"{calendar.month_name[MO]} Events",15,SLATE); bullets(s,0.85,2.5,9.4,2.0,N.get("events",[PH]),12.5)
    tb(s,0.85,4.6,9.5,0.3,"Asks and Next Steps",15,SLATE); bullets(s,0.85,5.0,9.4,2.8,N.get("asks",[PH]),12.5)
# 5 thanks
s=prs.slides.add_slide(blank); pic(s,"brick_soft.jpg",0,0,11,8.5)
b=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(11),Inches(1.0)); b.fill.solid(); b.fill.fore_color.rgb=BAR; b.line.fill.background(); b.shadow.inherit=False
tb(s,0.62,0.52,6,0.25,"SLIPSTREAM",10.5,PURPLE,bold=True,spacing=300); tb(s,0.5,1.12,10,0.65,"Thank you.",34,GREY)
pic(s,"thanks_photos.jpg",0,1.99,11,6.51)
apply_fonts(prs)
out=os.path.join(OUTDIR,f"Slipstream x NMWA Cafe Update {issue_name}.pptx"); prs.save(out); print("wrote",out)
for nm,d,st in months: print(nm,{k:(round(v,2) if isinstance(v,float) else v) for k,v in st.items()})
print("guest counts loaded:",len(guests))
