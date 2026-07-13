"""Fuzzy-match a source sheet's header columns onto our canonical schema
columns for whichever sheet type was already picked by sheet_matcher.

Headers in these Data Disks are often split across two rows (English on one
row, Korean on the next). We concatenate both rows per column before
matching so we're comparing against whichever language actually matches.
"""
from __future__ import annotations

import datetime
from dataclasses import dataclass
from typing import List, Optional

from schema import SheetSchema
from text_utils import normalize, similarity

MATCH_THRESHOLD = 0.35


@dataclass
class ColumnMatch:
    source_col_index: int
    source_header: str
    matched_key: Optional[str]
    confidence: float


def looks_like_data_row(row: List[object]) -> bool:
    """A row counts as data as soon as it has ANY numeric- or date-typed
    cell -- header/label rows are text-only, so even one real number or
    date is a strong signal. A proportional threshold (e.g. "half the
    cells") turned out to be fragile: which borrower happens to land in the
    first data row can shift the text/number mix a lot (e.g. one borrower
    with a long rehab case-number remark and multiple related-borrower IDs
    pushed a genuine data row's numeric share as low as 43%), so a ratio-
    based cutoff would misclassify some pools' first data row as a second
    header row and pull real values (account numbers, amounts, addresses)
    into what's supposed to be header text for matching."""
    non_empty = [c for c in row if c is not None and str(c).strip() != ""]
    if not non_empty:
        return False
    return any(isinstance(c, (int, float, datetime.date, datetime.datetime)) for c in non_empty)


def get_header_texts(grid: List[List[object]], header_row_index: int) -> List[str]:
    """Combine the header row (and a following bilingual row, if present)
    into one text string per source column."""
    rows_to_use = [header_row_index]
    next_row = header_row_index + 1
    if next_row < len(grid) and not looks_like_data_row(grid[next_row]):
        rows_to_use.append(next_row)

    width = max(len(grid[r]) for r in rows_to_use)
    texts = []
    for c in range(width):
        parts = []
        for r in rows_to_use:
            row = grid[r]
            if c < len(row) and row[c] is not None and str(row[c]).strip():
                parts.append(str(row[c]))
        texts.append(normalize(" ".join(parts)))
    return texts


def match_columns(header_texts: List[str], schema: SheetSchema) -> List[ColumnMatch]:
    """Greedy one-to-one matching: repeatedly pick the best remaining
    (source column, target column) pair until no pair scores above
    MATCH_THRESHOLD."""
    matchable_cols = [c for c in schema.columns if not c.derived]

    pairs = []
    for si, text in enumerate(header_texts):
        if not text:
            continue
        for col in matchable_cols:
            candidates = (col.header_kr, col.header_en) + col.aliases
            score = max(similarity(text, c) for c in candidates)
            if score >= MATCH_THRESHOLD:
                pairs.append((score, si, col.key))
    pairs.sort(key=lambda p: p[0], reverse=True)

    assigned_source = set()
    assigned_target = set()
    best_for_source = {}
    for score, si, key in pairs:
        if si in assigned_source or key in assigned_target:
            continue
        assigned_source.add(si)
        assigned_target.add(key)
        best_for_source[si] = (key, score)

    results = []
    for si, text in enumerate(header_texts):
        if not text:
            continue
        key, score = best_for_source.get(si, (None, 0.0))
        results.append(ColumnMatch(si, text, key, score))
    return results
