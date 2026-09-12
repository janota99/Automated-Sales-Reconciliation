"""Fuzzy PO matching for rows that survive every exact pass in matching.py.

Exact matching requires normalized POs, invoices, and signed amounts to
agree precisely, and treats anything less as unresolved rather than risk an
incorrect match. Real-world PO fields are sometimes a buyer's name or an ad
hoc note rather than a clean PO number -- QuickBooks might record "Hopper"
where Infinium records "DAVID HOPPER 2.2" for the exact same transaction --
so a small number of otherwise-correct items never clear the exact passes.

This module adds one additional, narrowly-scoped pass for rows that remain
unresolved after every exact pass in ``matching.perform_matching``:

  * The signed-cent amount must still match exactly -- fuzziness never
    applies to the dollar amount, only to the PO text.
  * The PO comparison is whole-word containment, not a general similarity
    score: both sides are split into significant (3+ letter, alphabetic)
    word tokens, and every token on the shorter side must appear on the
    longer side. "HOPPER" matches "DAVID HOPPER 2.2" because HOPPER is a
    whole word on both sides; two POs that merely look similar are not
    accepted. Numbers and short fragments are dropped before comparing, so
    shared product codes (e.g. "24" from "24 CASE") can never be the sole
    basis for a match.
  * The match must be unique on both sides: if a row would fuzzy-match more
    than one row on the opposing side, none of those candidates are
    accepted -- the engine never guesses among ambiguous options.

Every accepted fuzzy match carries its own confidence ("Fuzzy") and method
name, distinct from every exact-match method, so it is always separately
identifiable in every report this engine produces.
"""

from __future__ import annotations

import re
from collections import Counter
from typing import Any

import pandas as pd

from duplicates import AMOUNT_CENTS, SOURCE_POS

__all__ = [
    "FUZZY_PO_CONFIDENCE",
    "FUZZY_PO_EXPLANATION",
    "FUZZY_PO_METHOD",
    "PO_TOKENS",
    "find_unique_fuzzy_po_matches",
    "is_fuzzy_po_match",
    "significant_po_tokens",
]

# Working-frame column name for a row's precomputed significant PO tokens.
# Populated once in matching.prepare_working_frame from the raw PO column,
# so perform_matching never needs the column-mapping dict itself.
PO_TOKENS = "__REC_PO_TOKENS"

FUZZY_PO_METHOD = "Fuzzy PO + Amount (Word Match, Unique)"
FUZZY_PO_CONFIDENCE = "Fuzzy"
FUZZY_PO_EXPLANATION = (
    "Applied only after every exact pass left this row unresolved. The signed amount "
    "agrees exactly; the PO comparison is whole-word containment -- every significant "
    "word (3+ letters) on the shorter side's PO text appears on the longer side's -- "
    "and this was the sole such candidate on both sides."
)

_RE_WORD = re.compile(r"[A-Z]+")
_MIN_TOKEN_LENGTH = 3


def significant_po_tokens(value: Any) -> frozenset[str]:
    """Return the significant alphabetic word tokens in a raw PO value.

    Numbers and short fragments (fewer than 3 letters) are dropped: they are
    common to both PO numbers and product codes and would otherwise create
    false-positive containment matches on a shared number or abbreviation
    rather than a shared name.
    """
    if value is None or pd.isna(value):
        return frozenset()
    text = str(value).strip().upper()
    return frozenset(token for token in _RE_WORD.findall(text) if len(token) >= _MIN_TOKEN_LENGTH)


def is_fuzzy_po_match(tokens_a: frozenset[str], tokens_b: frozenset[str]) -> bool:
    """True if the smaller non-empty token set is fully contained in the other.

    Both sides must have at least one significant token -- an empty token
    set (blank PO, or a PO with no word 3+ letters long) never matches
    anything, since it would otherwise be a trivial subset of every row.
    """
    if not tokens_a or not tokens_b:
        return False
    shorter, longer = (tokens_a, tokens_b) if len(tokens_a) <= len(tokens_b) else (tokens_b, tokens_a)
    return shorter.issubset(longer)


def find_unique_fuzzy_po_matches(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    remaining_q: set[int],
    remaining_i: set[int],
) -> list[tuple[int, int]]:
    """Return unique one-to-one (QB index, Infinium index) fuzzy PO pairs.

    Only rows with a valid, exactly-matching signed amount and a non-empty
    fuzzy PO match are considered as candidates, and a pair is accepted
    only when it is the sole candidate for both of its rows -- ambiguous
    candidates are left unresolved rather than guessed.
    """
    q_rows = sorted(remaining_q, key=lambda idx: qb.at[idx, SOURCE_POS])
    i_rows = sorted(remaining_i, key=lambda idx: inf.at[idx, SOURCE_POS])

    candidates: list[tuple[int, int]] = []
    for qidx in q_rows:
        qamount = qb.at[qidx, AMOUNT_CENTS]
        q_tokens = qb.at[qidx, PO_TOKENS]
        if qamount is None or pd.isna(qamount) or not q_tokens:
            continue
        for iidx in i_rows:
            iamount = inf.at[iidx, AMOUNT_CENTS]
            if iamount is None or pd.isna(iamount) or int(qamount) != int(iamount):
                continue
            if is_fuzzy_po_match(q_tokens, inf.at[iidx, PO_TOKENS]):
                candidates.append((qidx, iidx))

    q_occurrences = Counter(qidx for qidx, _ in candidates)
    i_occurrences = Counter(iidx for _, iidx in candidates)
    return [
        (qidx, iidx)
        for qidx, iidx in candidates
        if q_occurrences[qidx] == 1 and i_occurrences[iidx] == 1
    ]
