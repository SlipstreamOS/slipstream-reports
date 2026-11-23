---
name: slipstream-spreadsheet-colors
description: Apply Slipstream's color-coding protocol to any Excel spreadsheet. Use whenever creating, editing, or reformatting an .xlsx file for Ryan / Slipstream — projections, weekly/monthly reports, financial close, budget models, anything tabular. Blue = input, Green = pull from another file/sheet, Black = formula or text. Apply even on existing sheets that don't yet follow the convention.
---

# Slipstream spreadsheet color protocol

Ryan's standard convention for every Slipstream spreadsheet. Apply it to every new file and retrofit existing files when editing them.

## The three cell roles

| Role | Font color | Background | When to use |
|------|-----------|------------|-------------|
| **Input** | Dark blue `#1F3864` | Light blue `#DCE6F1` | Any cell a human is meant to type into — assumptions, overrides, manual projections, toggles, targets. |
| **Pull** | Green `#006100` | (none) | Values that came from another file or another tab via import — prior-year actuals, YTD data loaded from CSVs, cross-sheet references. |
| **Formula / text** | Black `#000000` | (none) | Computed cells, labels, headers. The default. |

## Special cases

- **Key assumption cells** (growth rate, labor target %, etc.) get the blue input treatment PLUS a yellow highlight (`#FFEB9C`) to draw the eye.
- **Projection columns that look like formulas but are really inputs** — e.g., a sales projection column seeded with `=D6*(1+growth)` but meant to be overwritten day-by-day — count as inputs and get the blue font + light-blue fill.
- **Headers** stay black (and usually bold).
- **Date cells that auto-increment** (`=F7+1`) are formulas → black.

## Number formats

- Dollars: `"$"#,##0` (whole-dollar) or `"$"#,##0.00` (for avg ticket etc.)
- Percentages: `0.0%` unless whole percents are called for
- Counts: `#,##0`

## Retrofitting

When given a spreadsheet that doesn't yet follow the protocol, re-color only the cells you touch plus obvious assumption cells nearby. Mention the partial pass so Ryan can request a full sweep.

## Code snippet

```python
from openpyxl.styles import Font, PatternFill

BLUE_INPUT_FONT  = Font(name='Arial', size=10, color='1F3864')
BLUE_INPUT_FILL  = PatternFill('solid', start_color='DCE6F1', end_color='DCE6F1')
YELLOW_KEY_FILL  = PatternFill('solid', start_color='FFEB9C', end_color='FFEB9C')
GREEN_PULL_FONT  = Font(name='Arial', size=10, color='006100')
BLACK_FONT       = Font(name='Arial', size=10, color='000000')
```

## Why

A glance at any sheet should tell a manager or auditor what they can change (blue), what came from elsewhere and might be stale (green), and what's computed (black). Reduces handoff friction.
