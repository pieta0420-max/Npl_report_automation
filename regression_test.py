"""Regression check: run the full pipeline against a real Data Disk and
compare the 투심 aggregation output to numbers hand-verified against the
real WRB 2026-2 Program 투심보고서_VF.xlsx (opened via COM on 2026-07-09).

Usage:
    python regression_test.py "path\\to\\Data Disk.xlsx"

If no path is given, falls back to the WRB sample path used during
development. Exits non-zero if any check fails.
"""
from __future__ import annotations

import sys

from aggregator import compute_all, load_config
from column_matcher import get_header_texts, match_columns
from excel_io import load_workbook_any
from normalizer import ConfirmedSheet, normalize_all
from schema import CANONICAL_SCHEMAS
from sheet_matcher import detect_all

DEFAULT_PATH = (
    r"D:\2026\11.2026.2Q\15.Bid Pacakge\WRB 2026-2 Program_Bid_Package"
    r"\02. DataDisk\WRB 2026-2 Program_Data Disk_Ver III(Clean).xlsx"
)

# (table, pool, category, field, expected, tolerance)
EXPECTED = [
    ("overall_summary", "A", "일반담보부", "count", 198, 0),
    ("overall_summary", "A", "일반담보부", "opb_million", 125178, 1),
    ("overall_summary", "A", "일반담보부", "claim_million", 128168, 1),
    ("overall_summary", "A", "특별담보부", "count", 8, 0),
    ("overall_summary", "A", "특별담보부", "opb_million", 15699, 1),
    ("overall_summary", "B", "일반담보부", "count", 187, 0),
    ("overall_summary", "B", "일반담보부", "opb_million", 97403, 1),
    ("overall_summary", "B", "특별담보부", "count", 9, 0),
    # 100억 이상 is a dynamically-added tier (only appears when the data has
    # a qualifying borrower) -- this deal does, so 50억 이상 now means
    # [50억,100억) rather than VF's old un-split [50억,inf).
    ("amount_bucket", "A", "100억 이상", "count", 2, 0),
    ("amount_bucket", "A", "100억 이상", "opb_million", 28610, 0),
    ("amount_bucket", "A", "50억 이상", "count", 4, 0),
    ("amount_bucket", "A", "50억 이상", "opb_million", 29379, 0),
    ("amount_bucket", "B", "100억 이상", "count", 2, 0),
    ("amount_bucket", "B", "50억 이상", "count", 1, 0),
    ("foreclosure", "A", "경매 미개시", "count", 195, 1),
    ("rehab", "A", "개시신청", "count", 3, 0),
    ("rehab", "A", "개시신청", "opb_million", 14021, 1),
    ("rehab", "B", "개시결정", "count", 9, 0),
]

EXPECTED_GUARANTEE = [
    # pool, target_borrower_count, target_opb_million, guarantee_balance_million
    ("A", 4, 21441, 5313),
    ("B", 2, 8132, 979),
]


def run(dd_path: str) -> bool:
    wb = load_workbook_any(dd_path)
    matches = detect_all(wb)

    # Same disambiguation a careful user would do in the GUI: MCI-guarantee
    # (Sheet D) and credit-guarantee (Sheet E) sheets both look like
    # "guarantee" to the auto-matcher; keep only the higher-confidence one.
    best_per_type = {}
    for m in matches:
        if m.type_id is None:
            continue
        if m.type_id not in best_per_type or m.confidence > best_per_type[m.type_id].confidence:
            best_per_type[m.type_id] = m

    confirmed = []
    for type_id, m in best_per_type.items():
        schema = CANONICAL_SCHEMAS[type_id]
        grid = wb[m.sheet_name]
        header_texts = get_header_texts(grid, m.header_row_index)
        col_matches = match_columns(header_texts, schema)
        confirmed.append(ConfirmedSheet(dd_path, m.sheet_name, type_id, m.header_row_index, col_matches, grid))

    frames = normalize_all(confirmed)
    result = compute_all(frames, load_config())

    failures = []

    for table_name, pool, category, field, expected, tol in EXPECTED:
        df = result.tables[table_name]
        rows = df[(df["pool_type"] == pool) & (df["category"] == category)]
        if rows.empty:
            failures.append(f"{table_name}/{pool}/{category}: row not found")
            continue
        actual = rows.iloc[0][field]
        if abs(actual - expected) > tol:
            failures.append(f"{table_name}/{pool}/{category}/{field}: expected {expected}, got {actual}")

    g = result.tables["guarantee"]
    for pool, exp_n, exp_opb, exp_bal in EXPECTED_GUARANTEE:
        row = g[g["pool_type"] == pool]
        if row.empty:
            failures.append(f"guarantee/{pool}: row not found")
            continue
        r = row.iloc[0]
        if r["target_borrower_count"] != exp_n:
            failures.append(f"guarantee/{pool}/target_borrower_count: expected {exp_n}, got {r['target_borrower_count']}")
        if abs(r["target_opb_million"] - exp_opb) > 1:
            failures.append(f"guarantee/{pool}/target_opb_million: expected {exp_opb}, got {r['target_opb_million']}")
        if abs(r["guarantee_balance_million"] - exp_bal) > 1:
            failures.append(f"guarantee/{pool}/guarantee_balance_million: expected {exp_bal}, got {r['guarantee_balance_million']}")

    if failures:
        print(f"FAIL ({len(failures)} check(s) failed):")
        for f in failures:
            print("  -", f)
        return False

    print(f"PASS: all {len(EXPECTED) + len(EXPECTED_GUARANTEE) * 3} checks matched VF.xlsx values.")
    if result.unclassified_property_types:
        print("  (note) unclassified property types:", sorted(result.unclassified_property_types))
    return True


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PATH
    ok = run(path)
    sys.exit(0 if ok else 1)
