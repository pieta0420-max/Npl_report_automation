"""Shared text normalization helpers for fuzzy matching."""
from __future__ import annotations

import re


def normalize(text: object) -> str:
    if text is None:
        return ""
    s = str(text)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher

    a_n, b_n = normalize(a), normalize(b)
    if not a_n or not b_n:
        return 0.0
    return SequenceMatcher(None, a_n, b_n).ratio()
