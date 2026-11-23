#!/usr/bin/env python3
"""Make a recalculated throwaway copy of the model for validation/report builds (never touches the real file).
Usage: python3 recalc_copy.py <version>  -> $HOME/rc/Slipstream Consolidated Model 2026 v<version>.xlsx
Strips cached <v> from formula cells so LibreOffice must recalc, then converts."""
import sys,os,zipfile,re,subprocess
from lxml import etree
V=sys.argv[1]; H=os.path.expanduser("~")
D=os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","..",".."))
name=f"Slipstream Consolidated Model 2026 v{V}.xlsx"; src=os.path.join(D,name)
os.makedirs(H+"/rc_in",exist_ok=True); os.makedirs(H+"/rc",exist_ok=True); tmp=os.path.join(H,"rc_in",name)
NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
zin=zipfile.ZipFile(src)
with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
    for n in zin.namelist():
        b=zin.read(n)
        if re.match(r'xl/worksheets/sheet\d+\.xml$',n):
            t=etree.fromstring(b)
            for c in t.iter(NS+'c'):
                if c.find(NS+'f') is not None:
                    v=c.find(NS+'v')
                    if v is not None: c.remove(v)
                    if c.get('t') in ('str','e','b'): del c.attrib['t']
            b=etree.tostring(t,xml_declaration=True,encoding='UTF-8',standalone=True)
        zo.writestr(n,b)
out=os.path.join(H,"rc",name)
if os.path.exists(out): os.remove(out)
subprocess.run(["soffice","--headless","--calc","--convert-to","xlsx","--outdir",H+"/rc",tmp],check=True,capture_output=True)
print("recalculated copy:",out)
