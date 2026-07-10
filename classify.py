"""Shared classification helpers used by both aggregator.py (computes
ground-truth numbers for validation) and writer.py (writes the same
classification as Excel-visible helper columns + live formulas), so the
two can never drift apart.
"""
from __future__ import annotations

from typing import Dict, List, Set, Tuple

import pandas as pd


def pool_list(df: pd.DataFrame, col: str = "pool_type") -> List[str]:
    return sorted(x for x in df[col].dropna().unique())


def asset_type_label(asset_type: str, config: dict) -> str:
    return config["asset_type_labels"].get(asset_type, asset_type)


def apply_property_alias(value: str, config: dict) -> str:
    """Fold raw DD labels (e.g. '다가구', '오피스텔(주거)') onto the coarser
    canonical labels used in board reports (e.g. '다세대,연립,다가구주택')."""
    if value is None:
        return value
    return config.get("property_type_aliases", {}).get(value, value)


def group_lookup(value: str, groups: Dict[str, List[str]], fallback: str, unclassified: Set[str]) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        unclassified.add("(빈값)")
        return fallback
    for group_name, members in groups.items():
        if value in members:
            return group_name
    unclassified.add(str(value))
    return fallback


def property_group(value: str, config: dict, unclassified: Set[str]) -> str:
    canonical = apply_property_alias(value, config)
    return group_lookup(canonical, config["property_type_groups"], config["fallback_property_group"], unclassified)


def region_group(value: str, config: dict, unclassified: Set[str]) -> str:
    return group_lookup(value, config["region_groups"], config["fallback_region_group"], unclassified)


def active_amount_buckets(borrower_df: pd.DataFrame, config: dict) -> Tuple[List[int], List[str]]:
    """Only keep threshold tiers (e.g. '100억 이상') that at least one
    borrower in the *whole* dataset actually qualifies for, so a deal
    without any 100억+ borrower doesn't show an empty tier -- but a deal
    that does have one gets it automatically, no config edit needed. The
    catch-all bottom tier (last label, e.g. '10억 미만') is always kept.
    Computed once on the full borrower_df so every pool's table uses the
    same set of tiers."""
    all_thresholds = config["amount_buckets_krw"]
    all_labels = config["amount_bucket_labels"]
    opb = borrower_df["opb_incl_prepaid"]
    thresholds, labels = [], []
    for threshold, label in zip(all_thresholds, all_labels):
        if (opb >= threshold).any():
            thresholds.append(threshold)
            labels.append(label)
    labels.append(all_labels[-1])
    return thresholds, labels


def amount_bucket_label(opb: float, thresholds: List[int], labels: List[str]) -> str:
    if opb is None or (isinstance(opb, float) and pd.isna(opb)):
        return labels[-1]
    for threshold, label in zip(thresholds, labels):
        if opb >= threshold:
            return label
    return labels[-1]
