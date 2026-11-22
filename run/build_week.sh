#!/usr/bin/env bash
# Weekly report build, cloud only. See CLAUDE.md for the full Saturday runbook.
#
#   run/build_week.sh data   START   # after the Square pulls: bundle + merge sales, harvest labor/filter, build data
#   run/build_week.sh report START   # after writing data/narrative_<label>/*.txt: docx -> PDF -> reports/START/
#
# START is the Sunday that begins the week (YYYY-MM-DD). The week ends the following Saturday.
set -euo pipefail
MODE=${1:?data|report}; START=${2:?week start Sunday YYYY-MM-DD}
REPO=$(cd "$(dirname "$0")/.." && pwd)
END=$(python3 -c "import datetime as d,sys;s=d.date.fromisoformat('$START');assert s.weekday()==6,'not a Sunday';print(s+d.timedelta(6))")
LABEL=$(python3 -c "import datetime as d;s=d.date.fromisoformat('$START');e=s+d.timedelta(6);f=lambda x:x.strftime('%b').lower()+str(x.day);print(f'{f(s)}-{f(e)}')")
W=$REPO/_pipeline/weekly/scripts
SH=$REPO/_pipeline/shared/scripts

if [ "$MODE" = data ]; then
  OUT=$(mktemp -d)
  python3 "$SH/slipstream_cloud_delta.py" "$START" "$END" --out "$OUT" --root "$HOME/.claude/projects" --from-tenders
  B=$(ls "$OUT"/delta_*.tar.xz.b64)
  mkdir -p "$REPO/deltas" && cp "$B" "$REPO/deltas/"
  python3 "$SH/merge_deltas.py" "$B" --shared "$REPO/_pipeline/shared"
  python3 "$REPO/run/harvest.py" "$START" "$END"
  cd "$W"
  python3 set_week.py "$START"
  python3 build_week_from_daily_metrics.py
  python3 aggregate_labor.py
  python3 patch_week_from_txstore.py
  python3 build_weekly_data.py
  python3 build_filter_mix.py "$REPO/_pipeline/weekly/raw/filter_result.json"
  # First draft with whatever narrative is in place, so the numbers can be read before writing this week's.
  python3 build_weekly_report.py >/dev/null
  DOCX=$(ls -t "$REPO"/Slipstream\ Weekly\ Report\ -\ *.docx | head -1)
  python3 - "$DOCX" <<'EOF'
import sys, docx
d = docx.Document(sys.argv[1])
for p in d.paragraphs:
    if p.text.strip(): print(p.text)
for t in d.tables:
    print("---")
    for r in t.rows: print(" | ".join(c.text.strip() for c in r.cells))
EOF
  echo
  echo "DATA READY for $LABEL. Labor flags: see _pipeline/weekly/data/labor_$LABEL.json (exceptions, non_cafe_title_shifts)."
  echo "Write _pipeline/weekly/data/narrative_$LABEL/{headline,wins,watches,caveats}.txt, then: run/build_week.sh report $START"
elif [ "$MODE" = report ]; then
  N=$REPO/_pipeline/weekly/data/narrative_$LABEL
  for f in headline wins watches caveats; do [ -s "$N/$f.txt" ] || { echo "missing $N/$f.txt"; exit 2; }; done
  cd "$W"
  python3 set_week.py "$START" >/dev/null
  python3 set_narrative.py "$(cat "$N/headline.txt")" "$N/wins.txt" "$N/watches.txt" "$N/caveats.txt"
  python3 -c "import ast;ast.parse(open('build_weekly_report.py').read())"
  python3 build_weekly_report.py
  DOCX=$(ls -t "$REPO"/Slipstream\ Weekly\ Report\ -\ *.docx | head -1)
  python3 build_weekly_pdf.py "$DOCX"
  mkdir -p "$REPO/reports/$START"
  TAG=$(python3 -c "import datetime as d;s=d.date.fromisoformat('$START');e=s+d.timedelta(6);print(f\"{s.strftime('%b')} {s.day} to {e.strftime('%b')} {e.day}, {e.year}\")")
  cp "${DOCX%.docx}.pdf" "$REPO/reports/$START/"
  cp "$REPO/Timecards to Fix - $TAG.pdf" "$REPO/reports/$START/" 2>/dev/null || echo "(no timecard file this week)"
  ls -la "$REPO/reports/$START"
else
  echo "mode must be data or report"; exit 1
fi
