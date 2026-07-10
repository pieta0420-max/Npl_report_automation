"""Auto-detect which tab in a source workbook corresponds to which of our
4 canonical sheet types (borrower / collateral / guarantee / rehab), and
which row within that tab is the header row.

Detection is a weighted blend of:
  - title keyword overlap (the sheet's descriptive title, usually printed a
    few rows above the header, e.g. "차주일반정보 (Borrower Information)")
  - header token overlap (Jaccard similarity between the header row's words
    and the canonical schema's header words)

Anything below CONFIDENCE_THRESHOLD is left for the user to confirm/assign
in the GUI rather than guessed silently.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from schema import CANONICAL_SCHEMAS, SheetSchema
from text_utils import normalize, similarity

ANCHOR_TERMS = ("일련번호", "serial number")
MAX_TITLE_SCAN_ROWS = 8
MAX_HEADER_SCAN_ROWS = 15
CONFIDENCE_THRESHOLD = 0.45


@dataclass
class SheetMatch:
    sheet_name: str
    type_id: Optional[str]
    confidence: float
    header_row_index: Optional[int]  # 0-indexed row within the grid


def find_header_row(grid: List[List[object]]) -> Optional[int]:
    """Locate the row that contains an anchor term like '일련번호' / 'Serial Number'."""
    for r in range(min(MAX_HEADER_SCAN_ROWS, len(grid))):
        row = grid[r]
        for cell in row:
            norm = normalize(cell)
            if any(anchor in norm for anchor in ANCHOR_TERMS):
                return r
    return None


def _title_blob(grid: List[List[object]]) -> str:
    parts = []
    for r in range(min(MAX_TITLE_SCAN_ROWS, len(grid))):
        for cell in grid[r]:
            if cell is not None and str(cell).strip():
                parts.append(str(cell))
    return normalize(" ".join(parts))


def _header_tokens(grid: List[List[object]], header_row_index: int) -> set:
    tokens: set = set()
    for r in (header_row_index, header_row_index + 1):
        if r >= len(grid):
            continue
        for cell in grid[r]:
            norm = normalize(cell)
            if norm:
                tokens.update(norm.split())
    return tokens


def _schema_tokens(schema: SheetSchema) -> set:
    tokens: set = set()
    for col in schema.columns:
        tokens.update(normalize(col.header_kr).split())
        tokens.update(normalize(col.header_en).split())
    return tokens


def _jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


def score_sheet(grid: List[List[object]], schema: SheetSchema, header_row_index: Optional[int]) -> float:
    title_blob = _title_blob(grid)
    title_score = max(similarity(title_blob, kw) for kw in schema.title_keywords) if title_blob else 0.0
    # also reward exact substring hits, which are common and should score high
    kw_hits = sum(1 for kw in schema.title_keywords if normalize(kw) in title_blob)
    kw_score = kw_hits / len(schema.title_keywords)
    title_component = max(title_score, kw_score)

    header_component = 0.0
    if header_row_index is not None:
        header_component = _jaccard(_header_tokens(grid, header_row_index), _schema_tokens(schema))

    return 0.6 * title_component + 0.4 * header_component


def detect_sheet_type(sheet_name: str, grid: List[List[object]]) -> SheetMatch:
    header_row_index = find_header_row(grid)
    best_type: Optional[str] = None
    best_score = -1.0
    for type_id, schema in CANONICAL_SCHEMAS.items():
        score = score_sheet(grid, schema, header_row_index)
        if score > best_score:
            best_score = score
            best_type = type_id

    if best_score < CONFIDENCE_THRESHOLD:
        return SheetMatch(sheet_name, None, best_score, header_row_index)
    return SheetMatch(sheet_name, best_type, best_score, header_row_index)


def detect_all(workbook: Dict[str, List[List[object]]]) -> List[SheetMatch]:
    """Detect canonical type for every sheet in a loaded workbook."""
    results = []
    for sheet_name, grid in workbook.items():
        if not grid:
            continue
        results.append(detect_sheet_type(sheet_name, grid))
    return results
