"""Compute the 8 투심 summary tables from the normalized DataFrames.

Every source field these tables need (Pool Type, Asset Type, 차주형태, OPB,
채권총액, 담보유형, 담보소재지, 경매상태, 회생단계, 신용보증서 존재여부) is a
plain column on the borrower/collateral/guarantee/rehab DataFrames produced
by normalizer.py -- confirmed by reading VF.xlsx's own SUMIFS/COUNTIFS
formulas by hand. This module computes the same numbers in pandas so the
row/label structure (which pools, which categories, which order) is known
up front and testable (see regression_test.py); writer.py then re-expresses
each cell as a live Excel formula (COUNTIFS/SUMIFS) against the sheets it
writes, using this module's output only for the row labels/ordering, not
the numbers themselves. classification_config.json holds the only genuinely
deal-specific judgment calls (amount buckets, collateral-type groupings,
region groupings) so they can be tuned without touching this code.
"""
from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Set

import pandas as pd

from classify import (
    active_amount_buckets, amount_bucket_label, apply_property_alias, asset_type_label,
    group_lookup, pool_list, property_group, region_group,
)

# When frozen into an .exe (PyInstaller), __file__ points inside the bundled
# temp extraction dir -- classification_config.json needs to live next to
# the .exe itself so users can edit it without rebuilding.
_BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
CONFIG_PATH = _BASE_DIR / "classification_config.json"


def load_config(path: Path = CONFIG_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


@dataclass
class AggregationResult:
    tables: Dict[str, pd.DataFrame] = field(default_factory=dict)
    unclassified_property_types: Set[str] = field(default_factory=set)
    unclassified_regions: Set[str] = field(default_factory=set)


def _to_million(series: pd.Series) -> pd.Series:
    return (series.fillna(0) / 1_000_000).round(0)


def _pct(series: pd.Series, total: float) -> pd.Series:
    if not total:
        return series * 0.0
    return series / total


def _pool_list(borrower_df: pd.DataFrame) -> List[str]:
    return pool_list(borrower_df)


def _asset_label(asset_type: str, config: dict) -> str:
    return asset_type_label(asset_type, config)


# ---------------------------------------------------------------------------
# Table 1: 총괄표 (Pool x Asset Type)
# ---------------------------------------------------------------------------

def table_overall_summary(borrower_df: pd.DataFrame, config: dict) -> pd.DataFrame:
    rows = []
    grand_n = grand_opb = grand_claim = 0.0
    for pool in _pool_list(borrower_df):
        pool_df = borrower_df[borrower_df["pool_type"] == pool]
        pool_n = len(pool_df)
        pool_opb = pool_df["opb_incl_prepaid"].sum()
        pool_claim = pool_df["claim_total"].sum()
        for asset_type, grp in pool_df.groupby("asset_type"):
            n = len(grp)
            opb = grp["opb_incl_prepaid"].sum()
            claim = grp["claim_total"].sum()
            rows.append({
                "row_kind": "data", "pool_type": pool, "category": _asset_label(asset_type, config),
                "count": n, "count_pct": n / pool_n if pool_n else 0,
                "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "claim_million": round(_to_million(pd.Series([claim])).iloc[0]),
                "claim_pct": claim / pool_claim if pool_claim else 0,
            })
        rows.append({
            "row_kind": "subtotal", "pool_type": pool, "category": "소계",
            "count": pool_n, "count_pct": 1.0,
            "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]), "opb_pct": 1.0,
            "claim_million": round(_to_million(pd.Series([pool_claim])).iloc[0]), "claim_pct": 1.0,
        })
        grand_n += pool_n
        grand_opb += pool_opb
        grand_claim += pool_claim

    rows.append({
        "row_kind": "total", "pool_type": "", "category": "합계",
        "count": grand_n, "count_pct": None,
        "opb_million": round(_to_million(pd.Series([grand_opb])).iloc[0]), "opb_pct": None,
        "claim_million": round(_to_million(pd.Series([grand_claim])).iloc[0]), "claim_pct": None,
    })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 2: 차주형태별 분류 (Pool x 법인/개인사업자/개인)
# ---------------------------------------------------------------------------

def table_borrower_type(borrower_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for pool in _pool_list(borrower_df):
        pool_df = borrower_df[borrower_df["pool_type"] == pool]
        pool_n = len(pool_df)
        pool_opb = pool_df["opb_incl_prepaid"].sum()
        pool_claim = pool_df["claim_total"].sum()
        for borrower_type, grp in pool_df.groupby("borrower_type"):
            n = len(grp)
            opb = grp["opb_incl_prepaid"].sum()
            claim = grp["claim_total"].sum()
            rows.append({
                "row_kind": "data", "pool_type": pool, "category": borrower_type,
                "count": n, "count_pct": n / pool_n if pool_n else 0,
                "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "claim_million": round(_to_million(pd.Series([claim])).iloc[0]),
                "claim_pct": claim / pool_claim if pool_claim else 0,
            })
        rows.append({
            "row_kind": "subtotal", "pool_type": pool, "category": "소계",
            "count": pool_n, "count_pct": 1.0,
            "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]), "opb_pct": 1.0,
            "claim_million": round(_to_million(pd.Series([pool_claim])).iloc[0]), "claim_pct": 1.0,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 3: 금액별 분류 (Pool x OPB 구간, 누적치)
# ---------------------------------------------------------------------------

def table_amount_bucket(borrower_df: pd.DataFrame, config: dict) -> pd.DataFrame:
    thresholds, labels = active_amount_buckets(borrower_df, config)
    rows = []
    for pool in _pool_list(borrower_df):
        pool_df = borrower_df[borrower_df["pool_type"] == pool].copy()
        pool_n = len(pool_df)
        pool_opb = pool_df["opb_incl_prepaid"].sum()
        pool_df["_bucket"] = pool_df["opb_incl_prepaid"].apply(
            lambda v: amount_bucket_label(v, thresholds, labels))

        cum_n = cum_opb = 0
        for label in labels:
            grp = pool_df[pool_df["_bucket"] == label]
            n = len(grp)
            opb = grp["opb_incl_prepaid"].sum()
            cum_n += n
            cum_opb += opb
            rows.append({
                "row_kind": "data", "pool_type": pool, "category": label,
                "count": n, "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "cum_count": cum_n, "cum_opb_million": round(_to_million(pd.Series([cum_opb])).iloc[0]),
                "cum_opb_pct": cum_opb / pool_opb if pool_opb else 0,
            })
        rows.append({
            "row_kind": "subtotal", "pool_type": pool, "category": "소계",
            "count": pool_n, "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]), "opb_pct": 1.0,
            "cum_count": None, "cum_opb_million": None, "cum_opb_pct": None,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 4: 담보물 종류별 분석 (Pool x 담보유형, +유사유형 그룹 롤업)
# ---------------------------------------------------------------------------

def table_property_type(collateral_df: pd.DataFrame, config: dict,
                         unclassified: Set[str]) -> pd.DataFrame:
    groups = config["property_type_groups"]
    fallback = config["fallback_property_group"]
    rows = []
    for pool in sorted(x for x in collateral_df["pool_type"].dropna().unique()):
        pool_df = collateral_df[collateral_df["pool_type"] == pool].copy()
        pool_n = len(pool_df)
        pool_opb = pool_df["opb"].sum()
        pool_df["property_type"] = pool_df["property_type"].apply(lambda v: apply_property_alias(v, config))
        pool_df["_group"] = pool_df["property_type"].apply(
            lambda v: group_lookup(v, groups, fallback, unclassified))

        group_n = pool_df.groupby("_group")["property_type"].count()
        group_opb = pool_df.groupby("_group")["opb"].sum()

        pool_rows = []
        for prop_type, grp in pool_df.groupby("property_type"):
            n = len(grp)
            opb = grp["opb"].sum()
            grp_name = grp["_group"].iloc[0]
            pool_rows.append({
                "row_kind": "data", "pool_type": pool, "property_type": prop_type,
                "count": n, "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "group": grp_name,
                "group_count": int(group_n.get(grp_name, 0)),
                "group_opb_million": round(_to_million(pd.Series([group_opb.get(grp_name, 0)])).iloc[0]),
                "group_opb_pct": group_opb.get(grp_name, 0) / pool_opb if pool_opb else 0,
            })
        # cluster rows by group, groups ordered by their total OPB share (desc);
        # within a group, biggest individual type first
        pool_rows.sort(key=lambda r: (-r["group_opb_pct"], -r["opb_pct"]))
        rows.extend(pool_rows)
        rows.append({
            "row_kind": "total", "pool_type": pool, "property_type": "합계",
            "count": pool_n, "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]),
            "opb_pct": 1.0, "group": "", "group_count": None,
            "group_opb_million": None, "group_opb_pct": None,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 5: 담보물 지역별 분류 (Pool x 지역, +유사지역 그룹, 지역x유형 매트릭스)
# ---------------------------------------------------------------------------

def table_region(collateral_df: pd.DataFrame, config: dict,
                  unclassified: Set[str]) -> pd.DataFrame:
    groups = config["region_groups"]
    fallback = config["fallback_region_group"]
    rows = []
    for pool in sorted(x for x in collateral_df["pool_type"].dropna().unique()):
        pool_df = collateral_df[collateral_df["pool_type"] == pool].copy()
        pool_n = len(pool_df)
        pool_opb = pool_df["opb"].sum()
        pool_df["_group"] = pool_df["property_addr1"].apply(
            lambda v: group_lookup(v, groups, fallback, unclassified))

        group_n = pool_df.groupby("_group")["property_addr1"].count()
        group_opb = pool_df.groupby("_group")["opb"].sum()

        pool_rows = []
        for region, grp in pool_df.groupby("property_addr1"):
            n = len(grp)
            opb = grp["opb"].sum()
            grp_name = grp["_group"].iloc[0]
            pool_rows.append({
                "row_kind": "data", "pool_type": pool, "region": region,
                "count": n, "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "group": grp_name,
                "group_count": int(group_n.get(grp_name, 0)),
                "group_opb_million": round(_to_million(pd.Series([group_opb.get(grp_name, 0)])).iloc[0]),
                "group_opb_pct": group_opb.get(grp_name, 0) / pool_opb if pool_opb else 0,
            })
        # cluster rows by group, groups ordered by their total OPB share (desc);
        # within a group, biggest individual region first
        pool_rows.sort(key=lambda r: (-r["group_opb_pct"], -r["opb_pct"]))
        rows.extend(pool_rows)
        rows.append({
            "row_kind": "total", "pool_type": pool, "region": "합계",
            "count": pool_n, "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]),
            "opb_pct": 1.0, "group": "", "group_count": None,
            "group_opb_million": None, "group_opb_pct": None,
        })
    return pd.DataFrame(rows)


def table_region_x_property_matrix(collateral_df: pd.DataFrame, config: dict,
                                     unclassified_p: Set[str], unclassified_r: Set[str]) -> pd.DataFrame:
    """(*) 지역별/물건별 OPB 비중 matrix, one block per pool."""
    p_groups = config["property_type_groups"]
    r_groups = config["region_groups"]

    rows = []
    for pool in sorted(x for x in collateral_df["pool_type"].dropna().unique()):
        pool_df = collateral_df[collateral_df["pool_type"] == pool].copy()
        pool_opb = pool_df["opb"].sum()
        pool_df["_pgroup"] = pool_df["property_type"].apply(
            lambda v: property_group(v, config, unclassified_p))
        pool_df["_rgroup"] = pool_df["property_addr1"].apply(
            lambda v: region_group(v, config, unclassified_r))

        pivot = pool_df.pivot_table(index="_rgroup", columns="_pgroup", values="opb",
                                     aggfunc="sum", fill_value=0)
        region_group_total = pool_df.groupby("_rgroup")["opb"].sum()
        ordered_rgroups = sorted(r_groups, key=lambda g: -region_group_total.get(g, 0))
        for rgroup in ordered_rgroups:
            row = {"pool_type": pool, "region_group": rgroup}
            for pgroup in p_groups:
                val = pivot.loc[rgroup, pgroup] if rgroup in pivot.index and pgroup in pivot.columns else 0
                row[pgroup] = val / pool_opb if pool_opb else 0
            rows.append(row)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 6: 경매절차에 따른 분류 (Pool x Filed/Not Filed)
# ---------------------------------------------------------------------------

def table_foreclosure(collateral_df: pd.DataFrame) -> pd.DataFrame:
    labels = {"Filed": "경매 개시", "Not Filed": "경매 미개시"}
    rows = []
    for pool in sorted(x for x in collateral_df["pool_type"].dropna().unique()):
        pool_df = collateral_df[collateral_df["pool_type"] == pool]
        pool_n = len(pool_df)
        pool_opb = pool_df["opb"].sum()
        pool_claim = pool_df["claim_amount"].sum()
        for status, grp in pool_df.groupby("foreclosure_status"):
            n = len(grp)
            opb = grp["opb"].sum()
            claim = grp["claim_amount"].sum()
            rows.append({
                "row_kind": "data", "pool_type": pool, "category": labels.get(status, status),
                "count": n, "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "claim_million": round(_to_million(pd.Series([claim])).iloc[0]),
                "claim_pct": claim / pool_claim if pool_claim else 0,
            })
        rows.append({
            "row_kind": "total", "pool_type": pool, "category": "합계",
            "count": pool_n, "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]), "opb_pct": 1.0,
            "claim_million": round(_to_million(pd.Series([pool_claim])).iloc[0]), "claim_pct": 1.0,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 7: 회생절차에 따른 분류 (Pool x 세부 진행단계) -- borrower-level
# ---------------------------------------------------------------------------

def table_rehab(borrower_df: pd.DataFrame, rehab_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if rehab_df.empty:
        return pd.DataFrame(rows)
    merged = rehab_df.merge(
        borrower_df[["borrower_control_no", "opb_incl_prepaid", "claim_total"]],
        on="borrower_control_no", how="left",
    )
    for pool in sorted(x for x in merged["pool_type"].dropna().unique()):
        pool_df = merged[merged["pool_type"] == pool]
        pool_n = len(pool_df)
        pool_opb = pool_df["opb_incl_prepaid"].sum()
        pool_claim = pool_df["claim_total"].sum()
        for state, grp in pool_df.groupby("state"):
            n = len(grp)
            opb = grp["opb_incl_prepaid"].sum()
            claim = grp["claim_total"].sum()
            rows.append({
                "row_kind": "data", "pool_type": pool, "category": state,
                "count": n, "opb_million": round(_to_million(pd.Series([opb])).iloc[0]),
                "opb_pct": opb / pool_opb if pool_opb else 0,
                "claim_million": round(_to_million(pd.Series([claim])).iloc[0]),
                "claim_pct": claim / pool_claim if pool_claim else 0,
            })
        rows.append({
            "row_kind": "total", "pool_type": pool, "category": "합계",
            "count": pool_n, "opb_million": round(_to_million(pd.Series([pool_opb])).iloc[0]), "opb_pct": 1.0,
            "claim_million": round(_to_million(pd.Series([pool_claim])).iloc[0]), "claim_pct": 1.0,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Table 8: 신용보증서 유무에 따른 분류 (Pool 단위)
# ---------------------------------------------------------------------------

def table_guarantee(borrower_df: pd.DataFrame, guarantee_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    guaranteed_borrowers = set(guarantee_df["borrower_control_no"].dropna().unique())
    # MCI-classified guarantees (guarantor name contains "MCI") don't count
    # toward the 신용보증서 유효잔액 total -- confirmed with user 2026-07-10.
    is_mci = guarantee_df["guarantor"].fillna("").str.contains("MCI", case=False)
    balance_eligible = guarantee_df[~is_mci]
    for pool in _pool_list(borrower_df):
        pool_df = borrower_df[borrower_df["pool_type"] == pool]
        total_opb = pool_df["opb_incl_prepaid"].sum()
        guaranteed_pool_df = pool_df[pool_df["borrower_control_no"].isin(guaranteed_borrowers)]
        target_n = len(guaranteed_pool_df)
        target_opb = guaranteed_pool_df["opb_incl_prepaid"].sum()
        pool_guarantee_balance = balance_eligible[balance_eligible["pool_type"] == pool]["guarantee_balance_converted"].sum()
        rows.append({
            "pool_type": pool,
            "total_opb_million": round(_to_million(pd.Series([total_opb])).iloc[0]),
            "target_borrower_count": target_n,
            "target_opb_million": round(_to_million(pd.Series([target_opb])).iloc[0]),
            "guarantee_balance_million": round(_to_million(pd.Series([pool_guarantee_balance])).iloc[0]),
            "ratio_to_total_opb": pool_guarantee_balance / total_opb if total_opb else 0,
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------

def compute_all(frames: Dict[str, pd.DataFrame], config: dict = None) -> AggregationResult:
    config = config or load_config()
    result = AggregationResult()
    borrower_df = frames["borrower"]
    collateral_df = frames["collateral"]
    guarantee_df = frames["guarantee"]
    rehab_df = frames["rehab"]

    result.tables["overall_summary"] = table_overall_summary(borrower_df, config)
    result.tables["borrower_type"] = table_borrower_type(borrower_df)
    result.tables["amount_bucket"] = table_amount_bucket(borrower_df, config)
    result.tables["property_type"] = table_property_type(
        collateral_df, config, result.unclassified_property_types)
    result.tables["region"] = table_region(
        collateral_df, config, result.unclassified_regions)
    result.tables["region_x_property"] = table_region_x_property_matrix(
        collateral_df, config, result.unclassified_property_types, result.unclassified_regions)
    result.tables["foreclosure"] = table_foreclosure(collateral_df)
    result.tables["rehab"] = table_rehab(borrower_df, rehab_df)
    result.tables["guarantee"] = table_guarantee(borrower_df, guarantee_df)
    return result
