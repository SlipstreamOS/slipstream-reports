"""Add Slipstream Roasting (SRP) section to the payroll template.

Steps:
  1. Duplicate "CR - Schedule Hours" tab → "SRP - Schedule Hours" right after CR.
     Set $A$2 = "Slipstream Roasting" (so its X-column filter picks SRP shifts).
  2. In "Tip Pay Summary (ADP)":
     a. Insert 13 cols at position 15 (column O). All existing cols >= 15 shift +13.
     b. Update all existing formulas in the sheet to shift cell references for
        cols originally >= 15 by +13 (openpyxl's insert_cols moves cells but
        does NOT update formula text).
     c. Copy CR section (cols C-N) to new SRP section (cols P-AA), with two
        formula transforms: rename "CR - Schedule Hours" → "SRP - Schedule
        Hours" and shift within-section refs (C-N → P-AA, +13).
     d. Set P11 = "SRP" header.
     e. Hardcode 0 in SRP Tips column (col V, every 5th row from 14 to 358)
        since SRP has no tips.
     f. Update Totals row 14 formulas (originally BZ-CF; after shift CM-CS) to
        include the new SRP cells.
"""
from __future__ import annotations
import re, shutil, sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "template.xlsx"
OUT = ROOT / "template_with_srp.xlsx"

INSERT_AT = 15            # column O — insertion point
SHIFT = 13                # 13 new cols inserted
CR_FIRST, CR_LAST = 3, 14 # C-N (the CR section being copied)


# ---- formula transforms ------------------------------------------------------

# Two-stage strategy: stash other-sheet refs as placeholders so we never
# accidentally shift a column reference inside a different sheet's range.
SHEET_REF_RE = re.compile(
    r"((?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)"
)
CELL_REF_RE = re.compile(r"(\$?)([A-Z]+)(\$?)(\d+)")


CURRENT_SHEET = "Tip Pay Summary (ADP)"

def transform_same_sheet(formula: str, ref_fn) -> str:
    """Apply ref_fn to cell refs in formula. Stash ALL sheet-prefixed refs as
    placeholders so the bare-ref pass can't double-process them. For same-sheet
    refs we shift the cell part inside the placeholder before stashing."""
    placeholders: list[str] = []
    def stash(m):
        ref = m.group(1)
        sheet_part, cell_part = ref.split("!", 1)
        sheet_name = sheet_part.strip("'")
        if sheet_name == CURRENT_SHEET:
            shifted = CELL_REF_RE.sub(lambda m: ref_fn(*m.groups()), cell_part)
            placeholders.append(f"{sheet_part}!{shifted}")
        else:
            placeholders.append(ref)
        return f"~~ph_{len(placeholders)-1}_~~"
    formula = SHEET_REF_RE.sub(stash, formula)
    formula = CELL_REF_RE.sub(lambda m: ref_fn(*m.groups()), formula)
    for i, ph in enumerate(placeholders):
        formula = formula.replace(f"~~ph_{i}_~~", ph)
    return formula


def shift_ge_insert(col_dollar, col, row_dollar, row_num):
    """For cell refs with col index >= INSERT_AT and not column-absolute:
    shift col index by +SHIFT. Used after insert_cols to repair references."""
    if col_dollar == "$":
        return f"{col_dollar}{col}{row_dollar}{row_num}"
    idx = column_index_from_string(col)
    if idx >= INSERT_AT:
        return f"{col_dollar}{get_column_letter(idx + SHIFT)}{row_dollar}{row_num}"
    return f"{col_dollar}{col}{row_dollar}{row_num}"


def shift_cn_to_paa(col_dollar, col, row_dollar, row_num):
    """For cell refs with col index in CR section (3..14), shift to P-AA (16..27).
    Used when copying CR formulas into the new SRP section."""
    if col_dollar == "$":
        return f"{col_dollar}{col}{row_dollar}{row_num}"
    idx = column_index_from_string(col)
    if CR_FIRST <= idx <= CR_LAST:
        return f"{col_dollar}{get_column_letter(idx + SHIFT)}{row_dollar}{row_num}"
    return f"{col_dollar}{col}{row_dollar}{row_num}"


def cr_to_srp(formula: str) -> str:
    formula = formula.replace("CR - Schedule Hours", "SRP - Schedule Hours")
    return transform_same_sheet(formula, shift_cn_to_paa)


# ---- main --------------------------------------------------------------------

def main():
    shutil.copy(TEMPLATE, OUT)
    wb = openpyxl.load_workbook(OUT)

    # 1. Duplicate CR Schedule Hours -> SRP Schedule Hours
    src = wb["CR - Schedule Hours"]
    new_sh = wb.copy_worksheet(src)
    new_sh.title = "SRP - Schedule Hours"
    new_sh["A2"] = "Slipstream Roasting"
    # Replace SRP's role-lookup table (A3:C24) with LC's. Roasting staff are
    # the same people who work the cafes, so they punch FOH/Barista/etc., not
    # the kitchen-focused roles in CR's table.
    lc_src = wb["LC - Schedule Hours"]
    for r in range(3, 25):
        for c in range(1, 4):
            new_sh.cell(row=r, column=c).value = lc_src.cell(row=r, column=c).value
    # SRP-specific tweak: map every job title to "Admin" so SRP shifts use the
    # NORMAL wage rather than the tipped wage (no tips at the roastery).
    # Leave row 4 ("Kitchen Admin"→"KA") as-is and rows above the table.
    for r in range(5, 24):
        if new_sh.cell(row=r, column=1).value:  # only rows with a job title
            new_sh.cell(row=r, column=2).value = "Admin"
    # Move SRP to immediately after CR
    cr_idx = wb.sheetnames.index("CR - Schedule Hours")
    srp_idx = wb.sheetnames.index("SRP - Schedule Hours")
    wb.move_sheet("SRP - Schedule Hours", offset=cr_idx + 1 - srp_idx)

    # 2. Modify ADP tab
    adp = wb["Tip Pay Summary (ADP)"]

    # 2-pre. Fix existing range bug in CR section's E/F/G array formulas before
    # we copy CR to SRP. The Reg/OT/PTO Hours array formulas reference
    # 'CR - Schedule Hours'!$S$27:$S$328 which excludes row 26 (the first data
    # row in CR). The Min Guar (D) formula already uses the correct $26:$351,
    # so match that range. Apply ONLY to CR refs — LC/DT/1899/NMWA have their
    # own correct row ranges based on their own paste anchors.
    CR_RANGE_FIXES = [
        ("'CR - Schedule Hours'!$S$27:$S$328",   "'CR - Schedule Hours'!$S$26:$S$351"),
        ("'CR - Schedule Hours'!$V$27:$V$328",   "'CR - Schedule Hours'!$V$26:$V$351"),
        ("'CR - Schedule Hours'!$X$27:$X$328",   "'CR - Schedule Hours'!$X$26:$X$351"),
        ("'CR - Schedule Hours'!$Y$27:$Y$328",   "'CR - Schedule Hours'!$Y$26:$Y$351"),
        ("'CR - Schedule Hours'!$AB$27:$AB$328", "'CR - Schedule Hours'!$AB$26:$AB$351"),
        ("'CR - Schedule Hours'!$W$27:$W$328",   "'CR - Schedule Hours'!$W$26:$W$351"),
    ]
    range_fix_count = 0
    for row in adp.iter_rows():
        for cell in row:
            v = cell.value
            text = None
            if isinstance(v, str) and v.startswith("="):
                text = v
            elif isinstance(v, ArrayFormula):
                text = v.text
            if text is None:
                continue
            new_text = text
            for old_range, new_range in CR_RANGE_FIXES:
                new_text = new_text.replace(old_range, new_range)
            if new_text != text:
                if isinstance(v, ArrayFormula):
                    cell.value = ArrayFormula(v.ref, new_text)
                else:
                    cell.value = new_text
                range_fix_count += 1
    print(f"  Fixed {range_fix_count} CR array-formula ranges ($27:$328 → $26:$351)")

    # 2a. Insert 13 cols at position 15 (O) — physical shift only
    adp.insert_cols(INSERT_AT, SHIFT)

    # 2a.1. CRITICAL FIX: openpyxl's insert_cols moves cells but does NOT
    # update the `ref` attribute of ArrayFormula values. After insert, the
    # cell at (e.g.) AE14 still has ref='R14:R14' which makes Excel/LibreOffice
    # silently compute 0. Walk every ArrayFormula and rebuild its ref to match
    # the cell's actual coordinate.
    array_ref_fixes = 0
    for row in adp.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, ArrayFormula):
                expected_ref = f"{cell.coordinate}:{cell.coordinate}"
                # Most ArrayFormulas in this template are 1-cell. If the ref
                # doesn't match, rebuild as a 1-cell array at the new location.
                if v.ref != expected_ref and ":" in v.ref:
                    start, end = v.ref.split(":")
                    if start == end:
                        cell.value = ArrayFormula(expected_ref, v.text)
                        array_ref_fixes += 1
    print(f"  Repaired {array_ref_fixes} ArrayFormula refs after column insert")

    # 2b. Walk all formula cells; update text refs >= INSERT_AT by +SHIFT
    formula_updates = 0
    for row in adp.iter_rows():
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                new_v = transform_same_sheet(v, shift_ge_insert)
                if new_v != v:
                    cell.value = new_v
                    formula_updates += 1
            elif isinstance(v, ArrayFormula):
                old_text = v.text
                new_text = transform_same_sheet(old_text, shift_ge_insert)
                # also shift the ref's columns >= INSERT_AT by +SHIFT
                # (ref is like 'D14:D14'; for SRP-relevant refs we'll regenerate later)
                # For now just update text in place (single-cell array)
                if new_text != old_text:
                    cell.value = ArrayFormula(v.ref, new_text)
                    formula_updates += 1
    print(f"  Updated {formula_updates} formula references for col-shift")

    # 2c. Copy CR (cols 3..14, all rows) into new SRP location (cols 16..27)
    copied = 0
    for r in range(1, adp.max_row + 1):
        for src_col in range(CR_FIRST, CR_LAST + 1):
            dst_col = src_col + SHIFT  # 16..27 (P..AA)
            sc = adp.cell(row=r, column=src_col)
            dc = adp.cell(row=r, column=dst_col)
            v = sc.value
            if v is None:
                continue
            if isinstance(v, str):
                if v.startswith("="):
                    dc.value = cr_to_srp(v)
                else:
                    dc.value = v
            elif isinstance(v, ArrayFormula):
                # Preserve ArrayFormula — these need array evaluation. Build
                # a 1-cell array ref at the destination cell.
                dst_coord = dc.coordinate
                dc.value = ArrayFormula(f"{dst_coord}:{dst_coord}", cr_to_srp(v.text))
            else:
                dc.value = v
            # Style copy
            if sc.has_style:
                dc.font = copy(sc.font)
                dc.fill = copy(sc.fill)
                dc.border = copy(sc.border)
                dc.alignment = copy(sc.alignment)
                dc.number_format = sc.number_format
                dc.protection = copy(sc.protection)
            copied += 1
    print(f"  Copied {copied} cells from CR section to SRP section")

    # 2d. Set SRP header
    adp["P11"] = "SRP"

    # 2e. Hardcode 0 in SRP Tips col (V) for each employee block
    tips_col = column_index_from_string("V")
    for r in range(14, 359, 5):
        adp.cell(row=r, column=tips_col).value = 0

    # 2f. Update Totals formulas (now at CM-CS in row 14, originally BZ-CF) to
    #     include SRP cells. After step 2b the existing formulas already shifted
    #     their cross-section refs. Now we ADD the SRP refs.
    #
    # SRP cell mapping (CR col → SRP col):
    #   C (Wage)           -> P
    #   D (Min Guar)       -> Q
    #   E (Reg Hours)      -> R
    #   F (OT Hours)       -> S
    #   G (PTO Hours)      -> T
    #   I (Tips)           -> V (hardcoded 0)
    #   K (Misc)           -> X
    #   M (Gross Pay)      -> Z
    #
    # After bulk shift, totals at row 14 are:
    #   CM14 (was BZ): =MAX(AS14,AC14,C14,BH14)               # Wage
    #   CN14 (was CA): =E14+AE14+AU14+BJ14                    # Reg Hours
    #   CO14 (was CB): =AV19+AF19+F19+BK14                    # OT Hours
    #   CP14 (was CC): =AW19+AG19+G19+BL14                    # PTO Hours
    #   CQ14 (was CD): =AY14+BA14+AI14+AK14+I14+K14+BN14+BP14 # Tips+Misc
    #   CR14 (was CE): =BC14+AM14+M14+BR14+(CP14*CP15)        # Gross Pay
    #   CS14 (was CF): =MAX(0,MAX(AT14,AD14,D14,BI14)*SUM(CN19:CP19)-CR14)
    #
    # Inject SRP refs into EVERY row of the Totals area, not just row 14.
    # Each employee block has 5 data rows + 1 sum row. Most rows have at least
    # CM/CN formulas; the block-start row has the full set CM-CS. We walk all
    # rows and apply targeted string substitutions that are no-ops if SRP is
    # already present.
    totals_inject_count = 0
    CM = column_index_from_string("CM"); CN = column_index_from_string("CN")
    CO = column_index_from_string("CO"); CP = column_index_from_string("CP")
    CQ = column_index_from_string("CQ"); CR = column_index_from_string("CR")
    CS = column_index_from_string("CS")
    for r in range(14, 359):
        # CM (Wage MAX): "=MAX(AS{r},AC{r},C{r},BH{r})" → add P{r}
        cell = adp.cell(row=r, column=CM)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            old = cell.value
            new = old.replace(f",C{r},BH{r})", f",C{r},P{r},BH{r})")
            if new != old: cell.value = new; totals_inject_count += 1

        # CN (Reg Hours): "=E{r}+AE{r}+AU{r}+BJ{r}" → add +R{r}
        cell = adp.cell(row=r, column=CN)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            old = cell.value
            new = old.replace(f"=E{r}+AE{r}", f"=E{r}+R{r}+AE{r}")
            if new != old: cell.value = new; totals_inject_count += 1

        # CO (OT Hours, block-start only): existing has S{r+5}? add if not.
        cell = adp.cell(row=r, column=CO)
        if isinstance(cell.value, str) and cell.value.startswith("=") and "SUM(" not in cell.value:
            old = cell.value
            r5 = r + 5
            if f"S{r5}" not in old:
                # Standard pattern: "=AV{r5}+AF{r5}+F{r5}+BK{r}" or variants
                new = old.replace(f"=AV{r5}+", f"=AV{r5}+S{r5}+")
                if new == old:
                    new = old.replace(f"+AF{r5}+", f"+AF{r5}+S{r5}+")
                if new == old:
                    new = old.replace(f"+F{r5}+", f"+F{r5}+S{r5}+")
                if new != old: cell.value = new; totals_inject_count += 1

        # CP (PTO Hours): block-start "=AW{r5}+AG{r5}+G{r5}+BL{r}" → add T{r5}
        # Continuation cont1: "=MAX(AT{r-1},AD{r-1},D{r-1},BI{r-1})" → add Q{r-1}
        cell = adp.cell(row=r, column=CP)
        if isinstance(cell.value, str) and cell.value.startswith("=") and "SUM(" not in cell.value:
            old = cell.value
            r5 = r + 5; r1 = r - 1
            new = old
            if f"T{r5}" not in old and "+" in old:
                new = old.replace(f"=AW{r5}+", f"=AW{r5}+T{r5}+")
                if new == old:
                    new = old.replace(f"+G{r5}+", f"+G{r5}+T{r5}+")
            if "MAX(" in new and f"Q{r1}" not in new:
                new = new.replace(f",D{r1},BI{r1})", f",D{r1},Q{r1},BI{r1})")
            if new != old: cell.value = new; totals_inject_count += 1

        # CQ (Tips+Misc, block-start only): add +V{r}+X{r}
        cell = adp.cell(row=r, column=CQ)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            old = cell.value
            if f"V{r}" not in old:
                new = old.replace(f"+I{r}+K{r}+", f"+I{r}+K{r}+V{r}+X{r}+")
                if new == old:
                    new = old.replace(f"+I{r}+K{r})", f"+I{r}+K{r}+V{r}+X{r})")
                if new != old: cell.value = new; totals_inject_count += 1

        # CR (Gross Pay, block-start only): "=BC{r}+AM{r}+M{r}+BR{r}+(...)" → add +Z{r}
        cell = adp.cell(row=r, column=CR)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            old = cell.value
            if f"Z{r}" not in old:
                new = old.replace(f"+M{r}+", f"+M{r}+Z{r}+")
                if new != old: cell.value = new; totals_inject_count += 1

        # CS (Min Guar Extra): inner MAX needs Q{r}
        cell = adp.cell(row=r, column=CS)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            old = cell.value
            if f"Q{r}" not in old:
                new = old.replace(f",D{r},BI{r})", f",D{r},Q{r},BI{r})")
                if new != old: cell.value = new; totals_inject_count += 1

    print(f"  Injected SRP refs into {totals_inject_count} totals formulas")

    wb.save(OUT)
    print(f"Saved → {OUT}")
    print(f"\nVerification — open {OUT.name} and check:")
    print(f"  - SRP - Schedule Hours tab exists right after CR")
    print(f"  - Tip Pay Summary (ADP) P11 = 'SRP'")
    print(f"  - The totals section is at CM-CS (originally BZ-CF)")


if __name__ == "__main__":
    main()
