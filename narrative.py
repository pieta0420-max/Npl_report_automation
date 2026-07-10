"""Generates the 4 dynamic narrative sentences that appear in the Word
report's 자산 소개 section (pages 1-9). Patterns were reverse-engineered by
reading the real WRB 2026-2 investment committee report sentence-by-sentence
against its own tables on 2026-07-10:

1. 총괄표 intro: pool count / borrower count / total OPB in 억원.
2. 금액별 분류 (per pool): reports the 2nd of the top-2 (by threshold) amount
   tiers using its CUMULATIVE count/share through that tier, e.g. "Pool A는
   50억 이상 6개 차주가 전체 OPB의 41.16%를 차지하고 있습니다." -- matches
   the same 2 tiers the table highlights (100억 이상 + 50억 이상 combined).
3. 담보물 종류별 (per pool): names the top-2 collateral-type groups by OPB
   share, e.g. "상업용... 36.17%... 주거용... 29.94%".
4. 담보물 지역별 (per pool): same top-2 pattern for region groups.
"""
from __future__ import annotations

import pandas as pd


def fmt_pct(x: float) -> str:
    return f"{x * 100:.2f}%"


def fmt_int(n) -> str:
    return f"{int(round(n)):,}"


def overall_summary_sentence(bank_name: str, borrower_df: pd.DataFrame) -> str:
    n_pool = borrower_df["pool_type"].dropna().nunique()
    n_borrower = len(borrower_df)
    total_opb_eok = round(borrower_df["opb_incl_prepaid"].sum() / 100_000_000)
    return (
        f"금번 {bank_name} 매각 Program은 총 {n_pool}개 Pool, {n_borrower}개 차주, "
        f"OPB 약 {total_opb_eok:,}억원으로 세부적인 자산의 구성 내역은 다음과 같습니다."
    )


def amount_bucket_sentence(pool: str, pool_amount_table: pd.DataFrame) -> str:
    """pool_amount_table: rows for one pool from aggregator's amount_bucket
    table, row_kind=='data' only, in descending-threshold order (as produced
    by table_amount_bucket). Reports the 2nd of the top-2 (by threshold)
    tiers using its CUMULATIVE count/share through that tier -- e.g. "Pool A
    는 50억 이상 6개 차주가 전체 OPB의 41.16%를 차지하고 있습니다." If only
    1 tier exists, its own (== cumulative) count/share is used instead."""
    data_rows = pool_amount_table[pool_amount_table["row_kind"] == "data"]
    top2 = data_rows.head(2)
    if len(top2) == 0:
        return f"Pool {pool}의 금액별 분포는 아래와 같습니다."
    r = top2.iloc[-1]
    return (
        f"Pool {pool}는 {r['category']} {int(r['cum_count'])}개 차주가 "
        f"전체 OPB의 {fmt_pct(r['cum_opb_pct'])}를 차지하고 있습니다."
    )


def _top2_groups(pool_table: pd.DataFrame) -> list:
    """pool_table: rows for one pool from property_type/region table
    (row_kind=='data'), already ordered group-desc by aggregator.py. Returns
    up to 2 (group_name, group_opb_pct) tuples for distinct groups in order
    of first appearance (== descending share, per the sort in aggregator)."""
    seen = []
    for _, row in pool_table[pool_table["row_kind"] == "data"].iterrows():
        name = row["group"]
        if name not in [g for g, _ in seen]:
            seen.append((name, row["group_opb_pct"]))
        if len(seen) == 2:
            break
    return seen


def property_type_sentence(pool: str, pool_property_table: pd.DataFrame) -> str:
    top2 = _top2_groups(pool_property_table)
    if not top2:
        return f"Pool {pool}의 담보물건 종류별 분포는 아래와 같습니다."
    g1, p1 = top2[0]
    if len(top2) == 1:
        return f"Pool {pool} 담보물건은 {g1} 부동산이 전체 OPB의 {fmt_pct(p1)}로 가장 많은 비중을 차지하고 있습니다."
    g2, p2 = top2[1]
    return (
        f"Pool {pool} 담보물건은 {g1} 부동산이 전체 OPB의 {fmt_pct(p1)}로 가장 많은 비중을 차지하며, "
        f"그 외 {g2} 부동산이 {fmt_pct(p2)}를 차지하고 있습니다."
    )


def region_sentence(pool: str, pool_region_table: pd.DataFrame) -> str:
    top2 = _top2_groups(pool_region_table)
    if not top2:
        return f"Pool {pool}의 담보물건 지역별 분포는 아래와 같습니다."
    g1, p1 = top2[0]
    if len(top2) == 1:
        return f"Pool {pool} 담보물건은 {g1} 소재 물건이 전체 OPB 대비 약 {fmt_pct(p1)}를 차지하고 있습니다."
    g2, p2 = top2[1]
    return (
        f"Pool {pool} 담보물건은 {g1} 소재 물건이 전체 OPB 대비 약 {fmt_pct(p1)}를 차지하고 있으며, "
        f"그 외 {g2} 소재 물건이 {fmt_pct(p2)}를 차지하고 있습니다."
    )
