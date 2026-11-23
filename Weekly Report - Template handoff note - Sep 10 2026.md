# Weekly Report — changes from the Labor Deep Dive (handoff note, Sep 10 2026)

Paste this into the weekly-report chat. Everything referenced is already in the shared pipeline folder — nothing needs to be re-shared.

## 1. Where the data lives

- Templates (per cafe × tier × day): `_pipeline/monthly/data/labor_deep_dive/templates_v13.json` → `cafes[cafe][tier]["week"]` (hours, wages, labor_pct, sales_typical, ceiling, ceiling_trigger_sales) and `["days"][dow]["template"]` (hours, wages, shifts). Cafe keys: `Downtown`, `Studio`, `1899 Penn`, `NMWA`.
- Tier bands (weekly core/walk-in sales): `_pipeline/monthly/data/labor_deep_dive/tiers.json` → `[cafe]["bands_weekly"]`. Downtown slow < $7,500 / mid $7,500–9,000 / busy ≥ $9,000; Studio < $10,500 / $10,500–12,500 / ≥ $12,500; 1899 < $3,300 / $3,300–4,200 / ≥ $4,200. Scale short weeks to a normal week (5 days DT/1899, 7 Studio) before banding.
- Everything else (rules, targets, planning rates, Sep–Dec weekly tiers on Projections v1.9): `_pipeline/monthly/data/labor_deep_dive_2026-09.json` (v1.4).
- Projections: `Projections/2026 Projections v1.9.xlsx` now exists; `projections.py` auto-picks the highest version, so Section 5 sales will move on the next run (Studio Sep–Dec re-based −13% vs 2025; Downtown flat; Jan–Aug, 1899 and NMWA untouched). Source line in Section 5 should read v1.9.

## 2. Template weeks (planning rates: Lead $22.25, Line $19.75, Prod Barista $20.25, Barista/Closer $20.00)

| Cafe | Slow | Mid | Busy | Target / ceiling (slow ceiling) |
|---|---|---|---|---|
| Downtown | 120.5 h · $2,514 · 38.3% | 137 h · $2,832 · 34.3% | 139 h · $2,872 · 29.5% | 32 / 36 (38) |
| Studio | 183 h · $3,776 · 38.8% | 185 h · $3,821 · 33.6% | 193 h · $3,976 · 30.1% | 30 / 34 (36) |
| 1899 Penn | 67.5 h · $1,438 · 50.1% | 71 h · $1,513 · 40.2% | 76.5 h · $1,629 · 35.4% | 35 / 40 (42) |
| NMWA | 24 h · $510 (4 × 6 h, 9:00–3:00; +6 h on each community day) | — | — | staffing rule |

Day shapes are in the JSON (e.g. Downtown slow: Mon 22 h, Tue–Thu 25.5 h, Fri 22 h; Studio slow: Mon–Thu 22.25–22.75 h, Fri 31 h, Sat 31 h, Sun 30.5 h; 1899 stagger days Mon/Fri 12 h).

## 3. What changes in the report

**Scorecard (Targets Scorecard + Section 1).** Keep 32/30/35 as the target and the ceilings as the ask-Ryan line, but score the week against the TEMPLATE for the tier in force:
- Tier in force = the band the week's actual core sales fell in (report it: "Slow / Mid / Busy").
- Add columns: Template hours, Template core $ (at planning rates), Actual hours, Δ hours, Δ $.
- Status: GREEN = at or under template hours (or $); AMBER = between template and the ceiling (slow-tier ceiling is +2: 38 / 36 / 42); RED = over the ceiling. Compare hours first — the template is priced at planning rates and the roster's actual rates differ.
- Under the rules agreed Sep 5–10 (two openers/two closers at DT & Studio, 6-hour minimum shifts, three people mid-day at Studio, two shifts every day at 1899) the 32/30/35 targets are met only in busy weeks; mid weeks land inside the band; slow weeks sit at the slow ceiling. That is expected, not a miss — the lever in slow weeks is an admin-day swap (worth 1–3 pts), not cutting an opener or closer.

**Section 2 (daily plan).** Labor Plan for each day = the template day's wages for the tier in force (from `days[dow]["template"]["wages"]`), not projected sales × target %. The template's daily shape is deliberate (Downtown Friday runs 44–60% at its sales; Studio weekends carry the week), so a flat % per day misreads it. Federal-holiday Mondays stay closed / no plan at Downtown and 1899.

**Section 5 (next 4 weeks).** For each forward week, choose the tier from the projected core sales (v1.9) and show: Tier · Template hours · Template core $ · Ceiling $ (ask Ryan above). Drop the "one 8-hour shift per $500–600" rule of thumb — the template replaces it. Add the trigger: "switch to the next-slower template when projected weekly sales fall below $X" (`ceiling_trigger_sales` per tier).

**NMWA.** Staffing rule unchanged (1 person/day, 2 on community days = 2nd Wednesday and 1st Sunday), but shifts are 9:00–3:00 (a full hour to set up and to close) → 6 h/day, 24 h/week baseline.

**Wording.** Admin days are not in the templates — they are the scheduler's tool for slow weeks; keep reporting Admin and Training outside core as now.

## 4. Last week (Aug 30–Sep 5) re-read against the templates

| Cafe | Core sales → tier | Template | Actual | Δ | Read |
|---|---|---|---|---|---|
| Downtown | $7,425 → Slow | 120.5 h · $2,514 (38.3%) | 147 h · $3,103 (41.8%) | +26.5 h · +$589 | RED — over template and over the slow ceiling (38%) |
| Studio | $9,907 → Slow | 183 h · $3,776 (38.8%) | 196 h · $4,003 (40.4%) | +13 h · +$227 | over template by 13 h; over the slow ceiling (36%) — but the gap is one 6-h barista shift plus rate, not the $1,031 the 30% target implies |
| 1899 Penn | $4,425 → Busy | 76.5 h · $1,629 (35.4%) | 82 h · $1,711 (38.7%) | +5.5 h · +$82 | AMBER — one short shift over the busy template |

The headline changes from "everyone red against 30–32%" to "Downtown is 26 hours over its slow-week template; Studio and 1899 are within a shift of theirs."

## 5. Coming later (do not build yet)

- Incentive (Option 4: fixed amount per cafe-month run on template + half of any beat) will need a "months on template" tally — the scorecard above is the input.
- If the pay curve moves (band-top or peer-average curve), template $ rise ~12–17% and the % columns shift 4–5.5 points; hours do not change. Another reason to lead with hours.
