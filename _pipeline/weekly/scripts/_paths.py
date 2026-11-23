"""Shared path resolution for the weekly-report pipeline.

Discovers paths relative to the script's own location so the pipeline works
from any chat in this project (each chat has a different session root, but
all chats share the same workspace folder).

Layout assumed:
    <session>/mnt/Weekly and Monthly Reports/
        _pipeline/weekly/
            scripts/      ← this file lives here
            data/         ← intermediate JSONs (week_*.json, labor_*.json)
            raw/          ← raw inputs (labor_week.json extracted from transcript)
        Slipstream Weekly Data.xlsx       ← final xlsx
        Slipstream Weekly Report*.docx    ← final docx
"""
import os
import sys
import glob

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PIPELINE_DIR = os.path.dirname(SCRIPT_DIR)                           # _pipeline/weekly
WORKSPACE = os.path.dirname(os.path.dirname(PIPELINE_DIR))           # "Weekly and Monthly Reports"
DATA_DIR = os.path.join(PIPELINE_DIR, "data")
RAW_DIR = os.path.join(PIPELINE_DIR, "raw")

# Shared-pipeline scripts (transactions store, etc.) — make importable.
SHARED_SCRIPTS = os.path.join(os.path.dirname(PIPELINE_DIR), "shared", "scripts")
if SHARED_SCRIPTS not in sys.path:
    sys.path.insert(0, SHARED_SCRIPTS)

# Session root = ancestor containing "mnt/". The current-chat tool-results live
# under <SESSION>/mnt/.claude/projects/<proj>/<session-uuid>/tool-results.
def _find_session_root():
    p = SCRIPT_DIR
    while p != "/":
        if os.path.isdir(os.path.join(p, "mnt", ".claude")):
            return p
        p = os.path.dirname(p)
    return None

SESSION_ROOT = _find_session_root()

def find_tool_results_folders():
    """Return all tool-results folders visible from the current session.
    New chats will see their own session's folder; old chats' folders may
    still be present if not pruned."""
    if not SESSION_ROOT:
        return []
    pattern = os.path.join(SESSION_ROOT, "mnt", ".claude", "projects", "*", "*", "tool-results")
    return sorted(glob.glob(pattern))
