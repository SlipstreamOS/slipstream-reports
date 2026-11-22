#!/usr/bin/env python3
"""Swap the weekly narrative in build_weekly_report.py.

    python3 set_narrative.py "<headline python string literal(s)>" wins.txt watches.txt caveats.txt

Each .txt holds python string-literal lines ending in a comma (one list item per line).
The headline argument is the body passed to set_font(p.add_run( ... )) under the scorecard.
"""
import sys
p = "build_weekly_report.py"; s = open(p).read()
i = s.index('    p = doc.add_paragraph()\n    set_font(p.add_run(\n        "', s.index('for pp in t.rows[i].cells[7].paragraphs'))
j = s.index('    ), size=8, color=RGBColor(0x66,0x66,0x66))\n\n    # ==== Headlines')
s = s[:i] + '    p = doc.add_paragraph()\n    set_font(p.add_run(\n' + sys.argv[1] + '\n' + s[j:]
for key, path in (("wins", sys.argv[2]), ("watches", sys.argv[3]), ("caveats", sys.argv[4])):
    a = s.index(f"    {key} = ["); b = s.index("    ]", a) + 6
    s = s[:a] + f"    {key} = [\n" + open(path).read() + "    ]\n" + s[b:]
open(p, "w").write(s); print("narrative set")
