import re
P='build_monthly_report_v2.py'
s=open(P).read()
def rep(old,new,count=1):
    global s
    n=s.count(old)
    assert n>=1, ("MISSING",old[:90])
    s=s.replace(old,new) if count=='all' else s.replace(old,new,1)

# ---------- config
rep('MONTH = "Sep"','MONTH = "Oct"')
rep('MONTH_LABEL = "September 2026"','MONTH_LABEL = "October 2026"')
rep('PREPARED = "October 2, 2026"','PREPARED = "November 5, 2026"')
rep('MODEL_VER = "v18"','MODEL_VER = "v20"')
rep('FIN_JSON = os.path.join(DATA_DIR, "fin_2026_v18.json")','FIN_JSON = os.path.join(DATA_DIR, "fin_2026_v20.json")')
rep('PAT_JSON = os.path.join(DATA_DIR, "2609_patterns.json")','PAT_JSON = os.path.join(DATA_DIR, "2610_patterns.json")')
rep('LOY_JSON = os.path.join(DATA_DIR, "2609_loyalty.json")','LOY_JSON = os.path.join(DATA_DIR, "2610_loyalty.json")')
rep('LOY_PRIOR_JSON = os.path.join(DATA_DIR, "2608_loyalty.json")','LOY_PRIOR_JSON = os.path.join(DATA_DIR, "2609_loyalty.json")')
rep('MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]',
    'MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"]\n'
    '_FULL = {"Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April", "May": "May", "Jun": "June", "Jul": "July",\n'
    '         "Aug": "August", "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December"}\n'
    'PM = MONTHS[MONTHS.index(MONTH) - 1]                    # prior month key\n'
    'MONTH_NAME, PM_NAME = _FULL[MONTH], _FULL[PM]\n'
    'NEXT_NAME = list(_FULL.values())[(list(_FULL).index(MONTH) + 1) % 12]\n'
    'RANGE = f"Jan → {MONTH} 2026"')

# ---------- dynamic headings / labels
for old in ['"1b. Catering Sales — Jan → Sep 2026"','"1c. Prime Cost & Component Trends — Jan → Sep 2026 (ex-catering)"',
            '"1d. Core Sales Trend — Jan → Sep 2026 (ex-catering)"','"2a. Average Ticket ($, walk-in) — Jan → Sep 2026"',
            '"2b. Weekday Average Daily Transactions — Jan → Sep 2026"']:
    rep(old,'f'+old.replace('Jan → Sep 2026','{RANGE}'))
rep('["Jun", "Jul", "Aug", "Sep"]','MONTHS[-4:]','all')
rep('["Metric", "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026"]','["Metric"] + [f"{mm} 2026" for mm in MONTHS[-4:]]')
rep('["Daypart", "Avg Ticket", "Attach (proxy)", "Tip Rate", "Share Aug", "Share Jul"]','["Daypart", "Avg Ticket", "Attach (proxy)", "Tip Rate", f"Share {MONTH}", f"Share {PM}"]')
rep('"6b. Loyalty Activity by Cafe — September vs August"','f"6b. Loyalty Activity by Cafe — {MONTH_NAME} vs {PM_NAME}"')
rep('lrow("First-time members (exact Sep; Aug = upper bound)"','lrow("First-time members (no prior event in ledger)"')
rep('["Metric  (Sep  (Aug))", "Downtown"','[f"Metric  ({MONTH}  ({PM}))", "Downtown"')
rep('"6d. Visit-Frequency Distribution — September (members by visits in month)"','f"6d. Visit-Frequency Distribution — {MONTH_NAME} (members by visits in month)"')
rep('add_heading(doc, "Watch items for September", level=3)','add_heading(doc, f"Watch items going into {NEXT_NAME}", level=3)')

# ---------- month-specific caveats -> NARRATIVE keys
def caveat_block(start, end_marker, key):
    """replace add_caveat(doc, <...>) starting at `start` text up to the closing ')' of that call"""
    global s
    i=s.index(start); j=s.index(end_marker,i)+len(end_marker)
    s=s[:i]+f'add_caveat(doc, NARRATIVE["{key}"])'+s[j:]
caveat_block('add_caveat(doc, "1899 Net Income is a pass-through','see 1f.")','cav_1a')
# 1a callout: data-driven band words
i=s.index('add_callout(doc, "How to read this against the Cafe Finance 101 targets",'); j=s.index('body_size=9, title_size=10)',i)+len('body_size=9, title_size=10)')
s=s[:i]+'''def band(v, lo, hi):
        return "in band" if v <= hi else ("above" if v <= hi + 8 else "well above")
    def pband(v):
        return "healthy" if v <= 60 else ("warning" if v <= 65 else "problem")
    def pv(c, n, d):
        return F[c][n] / F[c][d] * 100 if F[c][d] else 0
    add_callout(doc, "How to read this against the Cafe Finance 101 targets",
                "Coffee COGS target 22–27% → " + ", ".join(f"{CAFE_LABEL[c]} {pv(c, 'ccogs', 'coffee'):.1f}% ({band(pv(c, 'ccogs', 'coffee'), 22, 27)})" for c in OPEN3) + ". "
                "Food COGS target 28–35% → " + ", ".join(f"{CAFE_LABEL[c]} {pv(c, 'fcogs', 'food'):.1f}% ({band(pv(c, 'fcogs', 'food'), 28, 35)})" for c in OPEN3) + ". "
                "Prime Cost 55–60% healthy / 60–65% warning / >65% problem → " + ", ".join(f"{CAFE_LABEL[c]} {pv(c, 'prime', 'core'):.1f}% ({pband(pv(c, 'prime', 'core'))})" for c in OPEN3) + ". "
                f"Blended {fmt_pct(O['prime'], O['core'])}. " + NARRATIVE["callout_1a"], body_size=9, title_size=10)'''+s[j:]
caveat_block('add_caveat(doc, "September catering rebounded','core ratio above.")','cav_1b')
caveat_block('add_caveat(doc, "Labor % reversed hard:','in the P&L.")','cav_1c')
# 1d: keep computed lead, replace the trailing narrative sentence
rep('''f"1899 {delta_pct(F['1899']['core'], fin['cafes']['1899'][pm]['core'])}. The post-Labor-Day recovery landed at the two office-district cafes; Studio did not get it.")''',
    '''f"1899 {delta_pct(F['1899']['core'], fin['cafes']['1899'][pm]['core'])}. " + NARRATIVE["cav_1d"])''')
rep('''add_caveat(doc, f"Open Cafes {fmt_money(oa)} in September, {delta_pct(oa, oj)} vs August''','''add_caveat(doc, f"Open Cafes {fmt_money(oa)} in {MONTH_NAME}, {delta_pct(oa, oj)} vs {PM_NAME}''')
i=s.index('add_para(doc, "Source: 2026 Projections v1.7 (FY26 budget)'); j=s.index('after=6)',i)+len('after=6)')
s=s[:i]+'add_para(doc, NARRATIVE["src_1e"], italic=True, size=9, color=COLOR_MUTED, after=6)'+s[j:]
caveat_block('add_caveat(doc, "Downtown within $200 of plan;','covered it at the entity level.")','cav_1e')
caveat_block('add_caveat(doc, "A catering month: $11,070','October is the read that matters.")','cav_1f')
i=s.index('add_para(doc, "Source changed this month per Ryan: ticket is now'); j=s.index('after=4)',i)+len('after=4)')
s=s[:i]+'add_para(doc, "Ticket is the per-transaction walk-in average from the Square transactions store (invoice / catering excluded), not the model\'s Ops Metrics ticket.", italic=True, size=9, color=COLOR_MUTED, after=4)'+s[j:]
caveat_block('add_caveat(doc, "Tickets have been stable all year:','not ticket.")','cav_2a')
caveat_block('add_caveat(doc, "Weekday transactions recovered at Downtown (158','a traffic story.")','cav_2c')
rep('add_para(doc, "NEW. Source: payment-level Square transactions','add_para(doc, "Source: payment-level Square transactions')
rep('''are placed where the revenue is, and what the basket looks like. CORE labor hours only (admin + training excluded). Timecards cover Sep 1–19; "
                  "every sales-per-hour and alignment figure is computed on matched days only, so partial coverage cannot inflate it.",''',
    '''are placed where the revenue is, and what the basket looks like. CORE labor hours only (admin + training excluded). " + NARRATIVE["src_3_coverage"],''')
caveat_block('add_caveat(doc, "Share of the month\'s walk-in net sales by daypart.','neighborhood profile (35–45%).")','cav_3a')
rep('The proxy is applied to Jul/Aug where descriptions are truncated to the first item. "','The proxy is applied from July on, where descriptions are truncated to the first item. "')
rep('add_para(doc, "Source: transaction lead item (first item on the ticket). The Jul/Aug Square exports truncate','add_para(doc, "Source: transaction lead item (first item on the ticket). Square exports since July truncate')
caveat_block('add_caveat(doc, "Send the September Square Item Sales CSVs','prior months.")','cav_4')
rep('"event in the store (store starts Aug 1, so September is the first month this is exact). NMWA is not "','"event in the store (store starts Aug 1). NMWA is not "')
caveat_block('add_caveat(doc, "Loyalty share of walk-in tickets is ~41%','in prior screenshot months).")','cav_6b')
caveat_block('add_caveat(doc, "Measured directly from the ledger','over five years at current spend.")','cav_6c')
caveat_block('add_caveat(doc, "Fixed schema from here on,','despite the smallest base.")','cav_6d')
i=s.index('add_caveat(doc, "Reading across the rows: sales per core labor hour moved'); j=s.index('"Benchmark notes: "',i)
s=s[:i]+'add_caveat(doc, NARRATIVE["cav_7a"] + " " + '+s[j:]
# section 8 data caveats -> list
i=s.index('    add_bullet(doc, f"Financials from Consolidated Model {MODEL_VER} (recalculated before reading; 18,536'); j=s.index('    add_heading(doc, f"Watch items going into',i)
s=s[:i]+'    for b in NARRATIVE["data_caveats"]:\n        add_bullet(doc, b)\n'+s[j:]
caveat_block('add_caveat(doc, f"Central Admin {fmt_money(S[\'Central\'][pm][\'ni\'])}','{fmt_money(S[\'Commissary\'][m][\'ni\'])}.")','cav_A1')
caveat_block('add_caveat(doc, f"Total Net Income {fmt_money(C[m][\'total_ni\'])} — best month','as the August tax items rolled off.")','cav_B')
rep('add_bullet(doc, f"Financials: Consolidated Model {MODEL_VER} → data/fin_2026_v18.json (extract of Jan–Aug per-entity IS tabs, Consolidated, Projections).")',
    'add_bullet(doc, f"Financials: Consolidated Model {MODEL_VER} → {os.path.basename(FIN_JSON)} (extract_fin.py: Jan–{MONTH} per-entity IS tabs, Consolidated, Projections).")')
rep('add_bullet(doc, "Sales patterns: _pipeline/shared/transactions/2026-08.csv (Square Transactions export) + weekly labor JSONs → analyze_patterns.py → data/2609_patterns.json.")',
    'add_bullet(doc, f"Sales patterns: _pipeline/shared/transactions (Square Transactions export, {MONTH_NAME} rows from the month file plus spillover in the neighbouring files) + weekly labor JSONs → analyze_patterns.py → {os.path.basename(PAT_JSON)}.")')

# ---------- NARRATIVE (October)
i=s.index('NARRATIVE = dict('); j=s.index('\n)\n',i)+3
NAR=open('nar_2610.txt').read()
s=s[:i]+NAR+s[j:]
open(P,'w').write(s)
print("patched OK")
