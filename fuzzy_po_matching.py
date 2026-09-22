"""Fuzzy PO matching for rows that survive every exact pass in matching.py.

This module provides a bounded, graph-based fuzzy matching pass for rows that
remain unresolved after every exact one-to-one and grouped pass. Real-world PO 
fields often contain ad-hoc notes or entity markers that break exact string 
equality (e.g., "DAVID HOPPER LLC" vs. "HOPPER").

Safeguards applied to circumstantial text matches:
  * The aggregate signed-cent amounts of the matched cluster must agree exactly.
  * Tokens must be alphanumeric. Pure numbers and financial stop-words are ignored.
  * Whole-word containment: EVERY significant word of the shorter reference must
    appear, exactly, in the longer one, in either direction (QuickBooks contained
    in Infinium or the reverse) -- e.g. {TIM, FERRIS} inside {PRIME, STAINLESS,
    TIM, FERRIS}. Always held for review.
  * Controlled typo (a SEPARATE, explicitly labeled method -- see
    find_controlled_typo_matches): exactly one word may differ by exactly one
    character inserted, deleted, or substituted (e.g. "ELIOT" vs "ELLIOT"),
    everything else must match exactly, and the match is accepted only when the
    exact signed amount agrees and there is exactly one qualifying candidate on
    each side. No similarity percentage or "closest match" is ever used.
  * Temporal anchor: If transaction dates are available, candidates must be
    within 30 days of each other.
  * Isolated clusters: Bipartite graph components are evaluated as a whole. If
    a linked cluster of rows ties out to a zero-variance aggregate sum, the entire
    cluster is cleared simultaneously (supporting 1:1, 1:M, and M:1 relationships).
  * Unique individual tie-out fallback: a recurring name (e.g. a dock-sale
    customer with several unrelated transactions that period) can pull more
    than one same-side row into a component even though only one is the real
    counterpart -- forcing all of them to sum together is the wrong question
    in that case. When one side of a component is a single row and the group
    sum doesn't tie, exactly one candidate on the other side individually
    matching that row's amount is accepted as the pair; more than one such
    candidate stays genuinely ambiguous and unresolved. See _resolve_components.
  * Exact-before-typo precedence: exact-token components are resolved first,
    and only rows left unclaimed afterward are considered for near-miss
    (typo-level) matching -- so an unrelated row elsewhere in the population
    that merely resembles a token (e.g. "SHOPPER" near-missing "HOPPER")
    can never drag a clean exact match into a larger, unbounded cluster and
    void it. See find_fuzzy_po_matches.
"""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any, Callable, Optional

import pandas as pd

from duplicates import AMOUNT_CENTS, SOURCE_POS

__all__ = [
    "FUZZY_PO_CONFIDENCE",
    "FUZZY_PO_EXPLANATION",
    "FUZZY_PO_METHOD",
    "PO_TOKENS",
    "TYPO_PO_CONFIDENCE",
    "TYPO_PO_METHOD",
    "controlled_typo_pair",
    "find_controlled_typo_matches",
    "find_fuzzy_po_matches",
    "is_fuzzy_po_match",
    "significant_po_tokens",
]

# Working-frame column name for a row's precomputed significant PO tokens.
PO_TOKENS = "__REC_PO_TOKENS"

FUZZY_PO_METHOD = "Fuzzy PO + Amount (Token Intersection & Aggregate)"
FUZZY_PO_CONFIDENCE = "Fuzzy"
FUZZY_PO_EXPLANATION = (
    "Applied only after exact passes left these rows unresolved. The aggregate signed "
    "amounts agree exactly; the PO comparison requires every significant whole word of the "
    "shorter reference to appear in the longer one (either direction; stop-words and pure "
    "numbers ignored). Matches require temporal proximity, a unique candidate, and avoid "
    "single-word generic false positives. Always held for review -- never posted automatically."
)

TYPO_PO_METHOD = "Controlled PO Typo + Exact Amount"
TYPO_PO_CONFIDENCE = "Typo"

_RE_WORD = re.compile(r"[A-Z]+|[0-9]+")
_MIN_TOKEN_LENGTH = 3
_MAX_GROUP_SIZE = 8

# A controlled typo needs a meaningful word: both spellings at least this long.
_MIN_TYPO_TOKEN_LENGTH = 5

# Globally filter lazy data entry and generic corporate entity markers
_STOP_WORDS = frozenset([
    "INC", "LLC", "LTD", "THE", "AND", "CORP", "COMPANY", "CO", 
    "MISC", "VOID", "NONE", "TBD", "NULL", "N/A"
])


def significant_po_tokens(value: Any) -> frozenset[str]:
    """Return the significant alphanumeric tokens in a raw PO value.

    Letters and digits split at their boundary even with no separator
    (e.g. "JWOOD7060" -> "JWOOD" + "7060", "ABOLT184225" -> "ABOLT" +
    "184225") -- a concatenated name-plus-reference-number code is common
    in these exports, and without the split the whole blob would never
    equal or resemble anything else again. Pure numbers and short
    fragments are then dropped to prevent false-positive containment
    matches on shared product codes or generic abbreviations.
    """
    if value is None or pd.isna(value):
        return frozenset()
    text = str(value).strip().upper()
    tokens = set()
    for token in _RE_WORD.findall(text):
        if token.isdigit():
            continue  
        if token in _STOP_WORDS:
            continue
        if len(token) >= _MIN_TOKEN_LENGTH:
            tokens.add(token)
    return frozenset(tokens)


def _typo_kind(a: str, b: str) -> Optional[str]:
    """"insertion" / "deletion" / "substitution" if the two words differ by exactly
    one character, else None. Deterministic edit-distance-one only -- no
    similarity ratio. Both words must be meaningful (>= 5 letters) and share
    their first letter: a typo rarely changes it, and requiring it rejects an
    unrelated word that merely adds a leading letter (SHOPPER vs HOPPER)."""
    if len(a) < _MIN_TYPO_TOKEN_LENGTH or len(b) < _MIN_TYPO_TOKEN_LENGTH or a == b or a[0] != b[0]:
        return None
    if len(a) == len(b):
        return "substitution" if sum(x != y for x, y in zip(a, b)) == 1 else None
    longer, shorter = (a, b) if len(a) > len(b) else (b, a)
    if len(longer) - len(shorter) != 1:
        return None
    for position in range(len(longer)):
        if longer[:position] + longer[position + 1:] == shorter:
            return "deletion" if longer is a else "insertion"
    return None


def controlled_typo_pair(
    tokens_a: frozenset[str], tokens_b: frozenset[str],
) -> Optional[tuple[str, str]]:
    """The one (word_a, word_b) pair by which two references differ, if -- and
    only if -- they are otherwise strongly consistent:

      * the shorter reference is contained in the longer EXCEPT for exactly one
        word, and exactly one word on the longer side is a one-character typo of
        it (two possible typo readings means no controlled typo);
      * a substitution (one letter swapped for another) additionally needs at
        least one other word that matches exactly, since a lone substituted
        word is as likely to be a different name (HOPPER vs HOOPER) as a typo;
      * a lone typo word must be at least six letters in one of its spellings.

    Returns None otherwise. Never a percentage, never "closest wins"."""
    if not tokens_a or not tokens_b:
        return None
    a_is_shorter = len(tokens_a) <= len(tokens_b)
    shorter, longer = (tokens_a, tokens_b) if a_is_shorter else (tokens_b, tokens_a)
    exact = shorter & longer
    unmatched_shorter = shorter - exact
    if len(unmatched_shorter) != 1:
        return None
    (short_word,) = unmatched_shorter
    readings = [word for word in longer - exact if _typo_kind(short_word, word)]
    if len(readings) != 1:
        return None
    long_word = readings[0]
    if _typo_kind(short_word, long_word) == "substitution" and not exact:
        return None
    if not exact and max(len(short_word), len(long_word)) < 6:
        return None
    return (short_word, long_word) if a_is_shorter else (long_word, short_word)


def is_fuzzy_po_match(
    tokens_a: frozenset[str], tokens_b: frozenset[str], *, allow_near_miss: bool = True
) -> bool:
    """True if the shorter reference is completely contained, word for word, in
    the longer -- or, with ``allow_near_miss``, is contained except for one
    controlled typo (see controlled_typo_pair).

    ``allow_near_miss=False`` restricts the comparison to exact whole-word
    containment -- the only rule the (held-for-review) fuzzy pass uses; the typo
    rule has its own explicitly-labeled pass."""
    if not tokens_a or not tokens_b:
        return False

    exact = tokens_a & tokens_b
    shorter_len = min(len(tokens_a), len(tokens_b))
    if exact and len(exact) >= shorter_len:
        # Prevent single-word generic false positives unless highly specific.
        if len(exact) == 1 and len(next(iter(exact))) < 6:
            return False
        return True
    return allow_near_miss and controlled_typo_pair(tokens_a, tokens_b) is not None


def _build_candidate_edges(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    q_rows: list[int],
    i_rows: list[int],
    qb_date_col: Optional[str],
    inf_date_col: Optional[str],
    max_days_variance: int,
    *,
    match_fn: Callable[[frozenset[str], frozenset[str]], bool],
) -> list[tuple[int, int]]:
    """Build candidate (qidx, iidx) edges that pass the temporal check and
    ``match_fn`` on the two rows' PO_TOKENS. Shared with the vendor-alias
    pass (see vendor_aliases.py) so every text-driven matching pass runs
    through the exact same graph-building and component-acceptance logic
    -- one implementation to keep correct, not two that could drift apart."""
    edges: list[tuple[int, int]] = []
    for qidx in q_rows:
        qamount = qb.at[qidx, AMOUNT_CENTS]
        q_tokens = qb.at[qidx, PO_TOKENS]
        if pd.isna(qamount) or not q_tokens:
            continue

        q_date = pd.to_datetime(qb.at[qidx, qb_date_col]) if qb_date_col else None

        for iidx in i_rows:
            iamount = inf.at[iidx, AMOUNT_CENTS]
            i_tokens = inf.at[iidx, PO_TOKENS]
            if pd.isna(iamount) or not i_tokens:
                continue

            # Temporal Anchor constraint
            if q_date is not None and inf_date_col is not None:
                i_date = pd.to_datetime(inf.at[iidx, inf_date_col])
                if pd.notna(q_date) and pd.notna(i_date):
                    if abs((q_date - i_date).days) > max_days_variance:
                        continue

            if match_fn(q_tokens, i_tokens):
                edges.append((qidx, iidx))

    return edges


def _resolve_components(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    edges: list[tuple[int, int]],
) -> list[tuple[tuple[int, ...], tuple[int, ...]]]:
    """Isolate connected components in the bipartite graph of edges and
    accept the ones that tie out to a bounded, aggregate-equal cluster."""
    q_adj: dict[int, set[int]] = defaultdict(set)
    i_adj: dict[int, set[int]] = defaultdict(set)
    for q, i in edges:
        q_adj[q].add(i)
        i_adj[i].add(q)

    visited_q: set[int] = set()
    accepted_groups: list[tuple[tuple[int, ...], tuple[int, ...]]] = []

    # Traverse isolated components and evaluate financial tie-out
    for q_start in list(q_adj.keys()):
        if q_start in visited_q:
            continue

        comp_q: set[int] = set()
        comp_i: set[int] = set()

        queue_q = [q_start]
        queue_i: list[int] = []

        # BFS traversal to isolate the full connected cluster
        while queue_q or queue_i:
            while queue_q:
                curr_q = queue_q.pop(0)
                if curr_q not in comp_q:
                    comp_q.add(curr_q)
                    visited_q.add(curr_q)
                    for nxt_i in q_adj[curr_q]:
                        if nxt_i not in comp_i:
                            queue_i.append(nxt_i)

            while queue_i:
                curr_i = queue_i.pop(0)
                if curr_i not in comp_i:
                    comp_i.add(curr_i)
                    for nxt_q in i_adj[curr_i]:
                        if nxt_q not in comp_q:
                            queue_q.append(nxt_q)

        # Enforce strict aggregate financial agreement on the isolated cluster.
        # Only bounded one-to-one, one-to-many, or many-to-one shapes are ever
        # posted downstream (see the grouped-matching control in matching.py);
        # a true many-to-many component is too coincidental to trust on an
        # aggregate sum alone and is left unresolved for manual review instead.
        q_sum = sum(int(qb.at[q, AMOUNT_CENTS]) for q in comp_q)
        i_sum = sum(int(inf.at[i, AMOUNT_CENTS]) for i in comp_i)
        bounded_shape = len(comp_q) == 1 or len(comp_i) == 1
        within_size_limit = len(comp_q) <= _MAX_GROUP_SIZE and len(comp_i) <= _MAX_GROUP_SIZE

        if q_sum == i_sum and bounded_shape and within_size_limit:
            accepted_groups.append((tuple(sorted(comp_q)), tuple(sorted(comp_i))))
            continue

        # Fallback: a real recurring name (e.g. a dock-sale customer with
        # several separate transactions that period) can pull more than one
        # same-side row into the text-match graph even though only one of
        # them is the genuine counterpart -- a group sum across all of them
        # is the wrong question to ask (they were never meant to net
        # together). When one side is a single row and the aggregate sum
        # doesn't tie, check whether exactly one candidate on the other
        # side individually matches that row's amount -- a strictly
        # stronger, more specific signal than a coincidental group sum.
        # If more than one candidate ties individually, the choice is
        # genuinely ambiguous and is correctly left unresolved, same as
        # today (see test_ambiguous_fuzzy_candidates_are_never_guessed).
        if within_size_limit and len(comp_q) == 1 and len(comp_i) > 1:
            single_q = next(iter(comp_q))
            single_amount = int(qb.at[single_q, AMOUNT_CENTS])
            ties = [i for i in comp_i if int(inf.at[i, AMOUNT_CENTS]) == single_amount]
            if len(ties) == 1:
                accepted_groups.append(((single_q,), (ties[0],)))
        elif within_size_limit and len(comp_i) == 1 and len(comp_q) > 1:
            single_i = next(iter(comp_i))
            single_amount = int(inf.at[single_i, AMOUNT_CENTS])
            ties = [q for q in comp_q if int(qb.at[q, AMOUNT_CENTS]) == single_amount]
            if len(ties) == 1:
                accepted_groups.append(((ties[0],), (single_i,)))

    return accepted_groups


def find_fuzzy_po_matches(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    remaining_q: set[int],
    remaining_i: set[int],
    qb_date_col: Optional[str] = None,
    inf_date_col: Optional[str] = None,
    max_days_variance: int = 30,
) -> list[tuple[tuple[int, ...], tuple[int, ...]]]:
    """Return grouped fuzzy PO pairs and aggregates using a bipartite graph.

    Builds edges between any rows that pass the temporal and fuzzy-text checks.
    Isolates connected components and clears them if the component's aggregate
    QuickBooks sum exactly equals its Infinium sum.

    Only exact whole-word containment is used here; a typo-level difference
    is handled by find_controlled_typo_matches, a separate, explicitly labeled
    pass, so a near-miss can never pull a self-contained exact pair into a
    larger cluster and void it.
    """
    q_rows = sorted(remaining_q, key=lambda idx: qb.at[idx, SOURCE_POS])
    i_rows = sorted(remaining_i, key=lambda idx: inf.at[idx, SOURCE_POS])

    exact_edges = _build_candidate_edges(
        qb, inf, q_rows, i_rows, qb_date_col, inf_date_col, max_days_variance,
        match_fn=lambda a, b: is_fuzzy_po_match(a, b, allow_near_miss=False),
    )
    accepted_groups = _resolve_components(qb, inf, exact_edges)

    return accepted_groups

def find_controlled_typo_matches(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    remaining_q: set[int],
    remaining_i: set[int],
) -> tuple[list[tuple[int, int, tuple[str, str]]], dict[int, list[int]]]:
    """Controlled PO-typo matches within the remaining unmatched population.

    A (QuickBooks, Infinium) pair qualifies only if the signed amounts agree
    EXACTLY row for row and the two PO references differ by exactly one
    controlled typo (see controlled_typo_pair). A pair is accepted only when each
    side has exactly one qualifying candidate; a row with more than one -- or
    whose single candidate is also claimed by another row -- is NOT matched, and
    is returned in the review mapping (QuickBooks index -> candidate Infinium
    indexes) so it can be sent to a reviewer instead of guessed.
    """
    q_candidates: dict[int, list[int]] = defaultdict(list)
    i_candidates: dict[int, list[int]] = defaultdict(list)
    typo_words: dict[tuple[int, int], tuple[str, str]] = {}
    q_rows = sorted(remaining_q, key=lambda idx: qb.at[idx, SOURCE_POS])
    i_rows = sorted(remaining_i, key=lambda idx: inf.at[idx, SOURCE_POS])
    for qidx in q_rows:
        q_amount, q_tokens = qb.at[qidx, AMOUNT_CENTS], qb.at[qidx, PO_TOKENS]
        if pd.isna(q_amount) or not q_tokens:
            continue
        for iidx in i_rows:
            i_amount, i_tokens = inf.at[iidx, AMOUNT_CENTS], inf.at[iidx, PO_TOKENS]
            if pd.isna(i_amount) or not i_tokens or int(q_amount) != int(i_amount):
                continue
            pair = controlled_typo_pair(q_tokens, i_tokens)
            if pair is not None:
                q_candidates[qidx].append(iidx)
                i_candidates[iidx].append(qidx)
                typo_words[(qidx, iidx)] = pair

    accepted: list[tuple[int, int, tuple[str, str]]] = []
    review: dict[int, list[int]] = {}
    for qidx, candidates in q_candidates.items():
        if len(candidates) == 1 and len(i_candidates[candidates[0]]) == 1:
            accepted.append((qidx, candidates[0], typo_words[(qidx, candidates[0])]))
        else:
            review[qidx] = list(candidates)
    return accepted, review
