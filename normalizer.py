"""Turn confirmed sheet/column mappings into canonical pandas DataFrames.

Handles the one piece of real business logic that isn't a straight column
copy: the collateral (Sheet C-1) sheet's `opb` and `claim_amount` columns
don't exist in any Data Disk -- they're each borrower's Sheet-A OPB / claim
total allocated across that borrower's properties, weighted by each
property's share of the borrower's total appraisal amount. See schema.py
for how this was reverse-engineered from VF.xlsx's cell formulas.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List

import pandas as pd

from column_matcher import ColumnMatch, looks_like_data_row
from schema import CANONICAL_SCHEMAS

SheetGrid = List[List[object]]


@dataclass
class ConfirmedSheet:
    """One source sheet, with sheet-type and column mapping locked in
    (either accepted from the auto-matcher or overridden by the user)."""
    file_label: str
    sheet_name: str
    type_id: str
    header_row_index: int
    column_matches: List[ColumnMatch]
    grid: SheetGrid


def find_data_start_row(grid: SheetGrid, header_row_index: int) -> int:
    row = header_row_index + 1
    if row < len(grid) and not looks_like_data_row(grid[row]):
        row += 1  # skip a second (bilingual) header row
    return row


def extract_rows(sheet: ConfirmedSheet) -> List[dict]:
    key_by_col = {cm.source_col_index: cm.matched_key for cm in sheet.column_matches if cm.matched_key}
    data_start = find_data_start_row(sheet.grid, sheet.header_row_index)

    rows = []
    for r in range(data_start, len(sheet.grid)):
        raw_row = sheet.grid[r]
        record = {}
        for col_idx, key in key_by_col.items():
            record[key] = raw_row[col_idx] if col_idx < len(raw_row) else None
        if all(v is None or str(v).strip() == "" for v in record.values()):
            continue  # blank trailing row
        record["_source_file"] = sheet.file_label
        record["_source_sheet"] = sheet.sheet_name
        rows.append(record)
    return rows


def build_dataframe(type_id: str, rows: List[dict]) -> pd.DataFrame:
    schema = CANONICAL_SCHEMAS[type_id]
    columns = [c.key for c in schema.columns] + ["_source_file", "_source_sheet"]
    df = pd.DataFrame(rows, columns=columns)
    return df


NUMERIC_KEYS = {
    "opb_excl_prepaid", "prepaid_expense", "opb_incl_prepaid", "accrued_interest",
    "claim_total", "mortgage_amount_converted", "senior_mortgage_amount",
    "pre_cutoff_court_deposit", "appraisal_amount_total", "guarantee_ratio",
    "initial_guarantee_amount", "guarantee_balance", "guarantee_balance_converted",
}


def _coerce_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if col in NUMERIC_KEYS:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def allocate_collateral_amounts(collateral_df: pd.DataFrame, borrower_df: pd.DataFrame) -> pd.DataFrame:
    """Fill in Sheet C-1's derived `opb` / `claim_amount` columns.

    borrower_control_no is only unique WITHIN a pool -- once multiple Data
    Disk files (one per pool) are combined, the same control number can be
    reused by an unrelated borrower in a different pool. Joining on
    borrower_control_no alone then makes borrower_lookup's index non-unique,
    which crashes with "cannot reindex on an axis with duplicate labels" (or
    worse, silently mismatches borrowers across pools). Every lookup here is
    therefore keyed on (pool_type, borrower_control_no) together."""
    if collateral_df.empty:
        collateral_df["opb"] = pd.Series(dtype=float)
        collateral_df["claim_amount"] = pd.Series(dtype=float)
        return collateral_df

    key_cols = ["pool_type", "borrower_control_no"]

    appraisal_sum = collateral_df.groupby(key_cols)["appraisal_amount_total"].transform("sum")
    count_per_borrower = collateral_df.groupby(key_cols)["borrower_control_no"].transform("count")

    ratio = collateral_df["appraisal_amount_total"] / appraisal_sum
    # if a borrower's properties have no appraisal amount at all, split evenly
    ratio = ratio.where(appraisal_sum.fillna(0) > 0, 1.0 / count_per_borrower)

    # drop_duplicates as a defensive backstop against messy source data (a
    # genuine duplicate borrower row within one pool); merge() -- unlike
    # join() -- can't crash on a non-unique key, but a duplicate would still
    # silently double-count via row expansion without this.
    borrower_lookup = borrower_df.drop_duplicates(key_cols)[key_cols + ["opb_incl_prepaid", "claim_total"]]
    joined = collateral_df.merge(borrower_lookup, on=key_cols, how="left")

    collateral_df["opb"] = joined["opb_incl_prepaid"].to_numpy() * ratio.to_numpy()
    collateral_df["claim_amount"] = joined["claim_total"].to_numpy() * ratio.to_numpy()
    return collateral_df


def normalize_all(confirmed_sheets: List[ConfirmedSheet]) -> Dict[str, pd.DataFrame]:
    """Extract, concatenate (across possibly-multiple source files), and
    return one DataFrame per canonical sheet type."""
    rows_by_type: Dict[str, List[dict]] = {t: [] for t in CANONICAL_SCHEMAS}
    for sheet in confirmed_sheets:
        rows_by_type[sheet.type_id].extend(extract_rows(sheet))

    frames = {t: _coerce_numeric(build_dataframe(t, rows)) for t, rows in rows_by_type.items()}

    frames["collateral"] = allocate_collateral_amounts(frames["collateral"], frames["borrower"])
    return frames
