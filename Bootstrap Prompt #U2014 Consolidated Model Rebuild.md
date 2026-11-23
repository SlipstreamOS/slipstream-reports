# Bootstrap Prompt — Consolidated Model Rebuild
## Use this to start a NEW dedicated chat for rebuilding the Consolidated Model

Copy everything below the `---` line into a fresh chat. This is a one-time refactor, not a monthly close. Plan for it to take multiple sessions if needed.

---

I'm Ryan, CEO of Slipstream Coffee. I want to **rebuild the Slipstream Consolidated Model from scratch**. The current version (`Slipstream Consolidated Model 2026 v7.xlsx`) has accumulated legacy structure from prior repurposing — inconsistent column layouts across tabs, mislabeled rows, dead tabs, and a Coefficient pull layer that adds friction. We're going to design and build a cleaner v8 (or however we version it) that's easier for me to maintain and easier for you to validate.

This is a SEPARATE workstream from the monthly close. The monthly close has its own process (see `/Documents/Claude/Projects/Weekly and Monthly Reports/Month Close Playbook.md`) that I run in a separate chat. Once we cut over to the rebuilt model, I'll update the close process to point at it.

## Before you do anything else — read these files in order

1. **`/Documents/Claude/Projects/Weekly and Monthly Reports/Month Close Playbook.md`** — full canonical month-close process. Gives you context on what the model is used for downstream and what conventions the close depends on.

2. **Memory files** about the current model:
   - `reference_consolidated_model.md` — current structure description
   - `project_consolidated_model_conventions.md` — Mgmt fee, NMWA, 1899 P/L, intercompany conventions that flow INTO the model
   - `feedback_consolidated_model_validation.md` — current NI tie-out procedure
   - `project_consolidated_model_rebuild_ryan_requirements.md` (if present) — Ryan's specific target design from the kickoff conversation

3. **The current model itself** — open `Slipstream Consolidated Model 2026 v7.xlsx` in the workspace folder. Spend time reading every tab. Don't write anything yet.

## Target design — Ryan's requirements

**Identity & granularity**:
- One tab per location/department: Studio, Downtown, 1899, NMWA, Commissary, SRP, Cafe Mgmt, Central.
- Plus a "New Location / Proforma" capability — but reworked (see Pro-forma section below).
- Consolidated tab with **toggles** so Ryan can select any subset of locations/support functions and see the rolled-up view. Should be able to include or exclude the New Location / Proforma in any selection.

**Per-location tab structure**:
- **Historical** columns — hardcoded after fiscal year close. Don't repull historical from QBO every refresh; it's frozen and doesn't change.
- **Current** columns — YTD monthly, pulled fresh from QBO each close.
- **Projections** columns:
  - Monthly through the next 6–18 months (e.g., if we're in June 2026, projections run monthly through end of 2027).
  - Annual totals for years 2–3 after that (e.g., 2028 and 2029 as single annual columns).
- Different locations may have **different historical depths** — don't force all tabs to line up. That's a feature, not a bug.

**Projection drivers** (Ryan is open to improvement suggestions):
- Top-line sales → driven by the Projections Sheet (currently top-line-only; needs product mix expansion).
- Product mix → estimate from **LTM average distribution** (last twelve months actual).
- Other expense lines → driven by one of:
  - % of sales (variable costs)
  - Inflation step-up (overhead, supplies)
  - Contractual (rent — fixed schedule with known steps)
- Assumptions typically live in **column E** of each location tab today. Open to restructuring this.

**Data refresh approach**:
- Replace Coefficient with direct QBO MCP pulls. Claude pulls P&L data via `qbo_accounting_get_profit_loss_generator` per location/class filter each close.
- Optional: leave Coefficient as a backup if Ryan wants ad-hoc refresh between closes without a chat session.

**Tabs to remove**:
- `2023 Capital Raise` — no longer needed.
- `Budget` and `Income Statement - New Location` — replaced by the new pro-forma template (see below).

**New Pro-forma template** (replaces Budget + New Location tabs):
Designed for evaluating new locations or new investments. Likely 3 tabs:
- **Inputs and Budget** — assumptions, drivers, operating costs.
- **Income Statement** — pro-forma P&L over the analysis horizon.
- **Returns** — sources & uses, debt waterfall, IRR, equity multiple, sensitivities.

Ryan has reference examples to share in this chat once we get to the pro-forma section.

## Label hygiene
The current model has a lot of legacy labels from when locations were combined on shared tabs (e.g., "Core Net Income (Excluding Comm+SRP)" — no longer needed since Comm/SRP are on separate tabs now). Clean these up as you go. Don't preserve labels that no longer carry meaning.

## Recommended sequence

1. **Audit v7 thoroughly** — read every tab, produce a structure map (tab name → row layout → column convention → what's a formula vs hardcoded → what's live vs estimated). Identify everything that looks legacy or unused.

2. **Surface design questions to Ryan** — anything you found that doesn't have a clear answer in this prompt or in the memory files. Ryan will weigh in.

3. **Sketch the target design** in markdown before building anything. Get Ryan's sign-off on:
   - Tab list and naming
   - Column layout strategy (historical | current | projections)
   - How toggles work on the Consolidated tab
   - Projection driver structure (assumptions location, formula approach)
   - Pro-forma template structure

4. **Build a proof-of-concept on Studio** end-to-end. Pull QBO P&L data directly. Apply the target structure. Add a sample projection through 2027. Show Ryan and iterate until the pattern is right.

5. **Validate the POC** by reproducing v7's Studio NI for a couple of closed months. They must match (or any difference must be explained — e.g., methodology change Ryan signed off on).

6. **Roll out to all 8 location/department tabs** using the validated pattern.

7. **Build the Consolidated tab** with toggles.

8. **Build the Pro-forma template** (after Ryan shares his reference examples).

9. **Test against a full closed month** — pick April or May 2026, run the rebuilt model, confirm consolidated NI matches QBO P&L within rounding.

10. **Cutover plan** — switch to v8 at the start of the next fresh month. Don't switch mid-close.

11. **Write a v8 Release Notes** markdown to the workspace folder summarizing what changed, what to watch for in the first close using v8, and any procedural differences from v7.

12. **Update the canonical references** so the monthly close process flows cleanly into v8:
    - Update `Month Close Playbook.md` (Section 4 Phase 12 + Section 13 Phase 13 + Section 2 Appendix).
    - Update memory files:
      - `reference_consolidated_model.md` → new structure
      - `project_consolidated_model_conventions.md` → preserved + new conventions
      - `feedback_consolidated_model_validation.md` → new validation procedure
    - Update `Bootstrap Prompt — Month Close.md` if any file paths or names changed.

## Things to actively decide (collect Ryan's input before building)

- **Projection horizon**: confirm 18 months monthly + 2 years annual is the right shape, or adjust.
- **Sales projection refinement**: how to bring product mix into the Projections Sheet — should the Projections Sheet itself be expanded, or should mix derivation live in each location tab using LTM data?
- **Inflation rate**: what default annual inflation % for overhead/supplies lines?
- **Toggles on Consolidated**: dropdown selection vs checkboxes vs named scenarios (e.g., "Operating", "Operating + Pipeline", "All including Proforma")?
- **Assumptions column**: stick with column E or move to a dedicated Assumptions tab that each location references?
- **Coefficient retention**: keep as backup or fully remove?

## Don't touch v7

Save the rebuild as `Slipstream Consolidated Model 2026 v8.xlsx` (or a name Ryan picks). v7 stays as the authoritative source until cutover. We may need to roll back if something breaks.

## End-state success criteria

When the rebuild is done:
- v8 produces the same consolidated NI as v7 for at least one fully closed historical month (within rounding).
- All 8 location tabs follow a consistent design pattern (same structure conceptually, even if column counts differ by historical depth).
- Consolidated tab toggles work and produce sensible roll-ups.
- Pro-forma template lets Ryan model a new location end-to-end.
- Projections run cleanly through the agreed horizon.
- Monthly close process (separate chat) can plug v8 into Phase 12 without friction.
- Canonical refs (system doc + memory) reflect v8 as the current state.

---

## Notes for me (Ryan)

**When to use this prompt**: when you're ready to start the rebuild — dedicated session, not during a monthly close.

**Before sending**: read it once, edit anything that's drifted, attach any pro-forma reference examples you want the new model to see upfront.

**Multi-session**: the rebuild may span multiple chats. Each session should end with a clear "where we left off" note saved to the workspace folder so the next session can pick up cleanly. The new model should propose this checkpoint structure at the end of each session.

**During the rebuild**: keep doing your monthly closes in the *other* chat against v7. Don't try to use the half-rebuilt v8 for an actual close.

**After cutover**: archive v7 (don't delete) and update the monthly close bootstrap prompt to reference v8.
