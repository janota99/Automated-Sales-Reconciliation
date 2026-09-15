"""Fuzzy PO matching for rows that survive every exact pass in matching.py.

This module provides a bounded, graph-based fuzzy matching pass for rows that
remain unresolved after every exact one-to-one and grouped pass. Real-world PO 
fields often contain ad-hoc notes or entity markers that break exact string 
equality (e.g., "DAVID HOPPER LLC" vs. "HOPPER").

Safeguards applied to circumstantial text matches:
  * The aggregate signed-cent amounts of the matched cluster must agree exactly.
  * Tokens must be alphanumeric. Pure numbers and financial stop-words are ignored.
  * Intersection ratio: Shared words must account for >= 60% of the shorter string.
    A word counts as shared if it matches exactly, or is a near-miss typo of a
    word on the other side (e.g. "ELIOT" vs "ELLIOT") -- see is_fuzzy_po_match.
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

import difflib
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
    "amounts agree exactly; the PO comparison requires a strong intersection of significant, "
    "alphanumeric tokens (excluding stop-words and pure numbers). Matches require temporal "
    "proximity and avoid single-word generic false positives."
)

_RE_WORD = re.compile(r"[A-Z0-9]+")
_MIN_TOKEN_LENGTH = 3
_MAX_GROUP_SIZE = 8

# A near-miss token pair must clear this similarity ratio to count as a typo
# of the same word rather than a different word. Calibrated so a one-character
# insertion/substitution in a 5-7 letter word matches (e.g. HOPER/HOPPER,
# ELIOT/ELLIOT both score ~0.91) while distinct-but-similar-looking words stay
# separate (HOPPER/HOOPER scores 0.833; STONE/STORE and SMITH/SMYTH score 0.80).
_TYPO_SIMILARITY_THRESHOLD = 0.84

# Globally filter lazy data entry and generic corporate entity markers
_STOP_WORDS = frozenset([
    "INC", "LLC", "LTD", "THE", "AND", "CORP", "COMPANY", "CO", 
    "MISC", "VOID", "NONE", "TBD", "NULL", "N/A"
])


def significant_po_tokens(value: Any) -> frozenset[str]:
    """Return the significant alphanumeric tokens in a raw PO value.

    Pure numbers and short fragments are dropped to prevent false-positive 
    containment matches on shared product codes or generic abbreviations.
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


def _near_miss_pairs(remaining_a: set[str], remaining_b: set[str]) -> list[tuple[str, str]]:
    """Greedily pair leftover tokens that are typo-level similar (not exact).

    Candidate pairs are consumed highest-similarity-first so a marginal
    near-miss never "steals" a token that had a better match available on
    either side. Pairs must also share a leading character -- a typo rarely
    changes the first letter, and requiring it excludes an unrelated real
    word that merely happens to contain the target word (e.g. "SHOPPER" or
    "CHOPPER" vs "HOPPER" both score ~0.92 on ratio() alone, well past the
    threshold, despite being different words rather than a misspelling).
    """
    candidates = sorted(
        (
            (difflib.SequenceMatcher(None, a, b).ratio(), a, b)
            for a in remaining_a if len(a) >= _MIN_TOKEN_LENGTH
            for b in remaining_b if len(b) >= _MIN_TOKEN_LENGTH
            if a[0] == b[0]
        ),
        key=lambda item: item[0],
        reverse=True,
    )
    used_a: set[str] = set()
    used_b: set[str] = set()
    pairs: list[tuple[str, str]] = []
    for ratio, a, b in candidates:
        if ratio < _TYPO_SIMILARITY_THRESHOLD:
            break
        if a in used_a or b in used_b:
            continue
        used_a.add(a)
        used_b.add(b)
        pairs.append((a, b))
    return pairs


def is_fuzzy_po_match(
    tokens_a: frozenset[str], tokens_b: frozenset[str], *, allow_near_miss: bool = True
) -> bool:
    """True if the token sets intersect (exactly, or by typo) with a strong
    operational ratio.

    ``allow_near_miss=False`` restricts the comparison to exact token
    overlap only -- used to resolve the clean, unambiguous matches in a
    population before typo-level near-misses are considered at all (see
    ``find_fuzzy_po_matches``).
    """
    if not tokens_a or not tokens_b:
        return False

    exact = tokens_a & tokens_b
    near_misses = _near_miss_pairs(tokens_a - exact, tokens_b - exact) if allow_near_miss else []
    matched_count = len(exact) + len(near_misses)
    if not matched_count:
        return False

    # Prevent single-word generic false positives unless highly specific.
    if matched_count == 1:
        if exact:
            match_word = next(iter(exact))
            if len(match_word) < 6:
                return False
        else:
            word_a, word_b = near_misses[0]
            if max(len(word_a), len(word_b)) < 6:
                return False

    # Require the intersection to cover the majority of the shorter side, so a
    # short buyer-name note (e.g. "HOPPER") can still match a longer PO field
    # that fully contains it (e.g. "DAVID HOPPER LLC") -- this is the primary
    # case this module exists to catch. Matching against the longer side
    # instead would demand the short side subsume most of the long side too,
    # which no genuine buyer-name-vs-PO-field pair ever does.
    # e.g., {"DAVID", "SMITH", "LLC"} vs {"DAVID", "HOPPER"} still fails the ratio
    shorter_len = min(len(tokens_a), len(tokens_b))
    ratio = matched_count / shorter_len

    return ratio >= 0.60


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

    Resolved in two tiers so a typo-level near-miss can never contaminate a
    clean exact match: an unrelated row elsewhere in the population that
    merely *resembles* a real token (e.g. "SHOPPER" near-missing "HOPPER")
    would otherwise pull a valid, self-contained 1:1 pair into a larger,
    unbounded or non-tying component and void the whole cluster. Exact-token
    components are found and accepted first; only the rows still unclaimed
    afterward are considered for near-miss (typo-level) matching.
    """
    q_rows = sorted(remaining_q, key=lambda idx: qb.at[idx, SOURCE_POS])
    i_rows = sorted(remaining_i, key=lambda idx: inf.at[idx, SOURCE_POS])

    exact_edges = _build_candidate_edges(
        qb, inf, q_rows, i_rows, qb_date_col, inf_date_col, max_days_variance,
        match_fn=lambda a, b: is_fuzzy_po_match(a, b, allow_near_miss=False),
    )
    accepted_groups = _resolve_components(qb, inf, exact_edges)

    claimed_q = {q for q_group, _ in accepted_groups for q in q_group}
    claimed_i = {i for _, i_group in accepted_groups for i in i_group}
    leftover_q = [qidx for qidx in q_rows if qidx not in claimed_q]
    leftover_i = [iidx for iidx in i_rows if iidx not in claimed_i]

    near_miss_edges = _build_candidate_edges(
        qb, inf, leftover_q, leftover_i, qb_date_col, inf_date_col, max_days_variance,
        match_fn=lambda a, b: is_fuzzy_po_match(a, b, allow_near_miss=True),
    )
    accepted_groups.extend(_resolve_components(qb, inf, near_miss_edges))

    return accepted_groups