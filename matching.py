"""Controlled QuickBooks-to-Infinium normalization and matching engine.

Automatic matching first exhausts unique one-to-one relationships. A bounded,
unambiguous one-to-many or many-to-one pass may then match rows sharing an exact
normalized PO and/or invoice when their signed-cent totals agree exactly.

Duplicate detection, disposition, and reporting are handled by
``duplicates.py``. Strong groups retain one canonical primary row and exclude
only excess copies. Weaker candidates remain visible for review. Historical
rows overlapping their own primary dataset are prevented from clearing an
exception.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from difflib import get_close_matches
from functools import lru_cache
from itertools import zip_longest
from typing import Any, Optional

import pandas as pd

from duplicates import (
    AMOUNT_CENTS,
    DUPLICATE_RULE_VERSION,
    NORM_INV,
    NORM_PO,
    SOURCE_POS,
    combine_duplicate_reports,
    finalize_review_dispositions,
    screen_duplicates,
)
from fuzzy_po_matching import (
    FUZZY_PO_CONFIDENCE,
    FUZZY_PO_EXPLANATION,
    FUZZY_PO_METHOD,
    PO_TOKENS,
    find_fuzzy_po_matches,
    significant_po_tokens,
)

__all__ = [
    "AMOUNT_CENTS",
    "AMOUNT_VARIANCE_COLUMNS",
    "APP_VERSION",
    "INF_ID",
    "MATCHING_RULE_VERSION",
    "QB_ID",
    "ReconciliationResult",
    "build_fiscal_exception_summary",
    "build_reference_amount_variances",
    "build_reconciliation",
    "cents_or_zero",
    "cents_to_float",
    "clean_alphanumeric",
    "clean_po",
    "get_fuzzy_lexicon_match",
    "numeric_quantity_sum",
    "numeric_sum",
    "parse_amount_cents",
    "parse_fiscal_period",
    "valid_cents",
]


APP_VERSION = "2.16.3"
MATCHING_RULE_VERSION = "2026.09-STREAMLINED-DUPLICATE-OUTPUT"

# Grouped matching is intentionally bounded to keep reconciliation runs
# predictable. Larger or more complex reference pools remain unresolved for
# human review instead of risking a slow or arbitrary automatic allocation.
MAX_GROUP_POOL_ROWS = 20
MAX_GROUP_SIZE = 8

QB_ID = "__REC_QB_ID"
INF_ID = "__REC_INF_ID"
PRODUCT_STANDARD = "__REC_PRODUCT_STANDARD"
FISCAL_LABEL = "__REC_FISCAL_LABEL"
PRODUCT_LEXICON = {
    "Allsups 24 Case": ["ALLSUPS", "ALLSUPS 24", "ALLSUPS 24 CASE", "ALLSUP 24"],

    "Food Club 24 Case": [
        "FOOD CLUB", "FOOD CLUB 24", "FOOD CLUB 24 CASE", "FC 24",
        "FOODCLUB 24", "FOODCLUB24 CASE", "FC24 CASE",
    ],
    "Food Club 40 Case": [
        "FOOD CLUB 40", "FOOD CLUB 40 CASE", "FC 40", "FOODCLUB 40",
        "FOODCLUB40 CASE", "FC40 CASE",
    ],  
    "Food King 24 Case": [
        "FOOD KING 24", "FOOD KING 24 CASE", "FK 24", "FOODKING 24",
        "FOODKING24 CASE", "FK24 CASE", "KINGS 24 CASE", "KINGS 24",
    ],
    "Food King 40 Case": [
        "FOOD KING 40", "FOOD KING 40 CASE", "FK 40", "FOODKING 40",
        "FOODKING40 CASE", "FK40 CASE", "KINGS 40 CASE", "KINGS 40",
    ],
    "Juniors 24 Case": ["JUNIORS", "JUNIORS 24", "JUNIORS 24 CASE"],
    "Lowes 24 Case": ["LOWES", "LOWES 24", "LOWES 24 CASE", "LOWES24"],
    "Lowes 40 Case": ["LOWES 40", "LOWES 40 CASE"],
    "Panhandle Pure 24 Case": [
        "PPL24", "PP 24 CASE", "PPL 24 CASSE", "PP24",
        "PANHANDLE PURE 24 CASE", "PPL 24 CASE", "PPL 24",
    ],
    "Panhandle Pure 40 Case": [
        "PPL40", "PP 40 CASE", "PPL 40 CASSE", "PP40",
        "PANHANDLE PURE 40 CASE", "PPL 40 CASE", "PPL 40",
    ],
    "Plains 24 Case": ["PLAINS", "PLAINS 24", "PLAINS 24 CASE", "PLAINS24"],
    "Spring House 24 Case": [
        "SPRING HOUSE", "SPRING HOUSE 24", "SPRING HOUSE 24 CASE", "SH 24",
    ],
    "Toot N Totum 24 Case": [
        "TNT24", "TOOT N TOTUM", "TOOT 'N TOTUM 24", "TNT 24",
        "TOOT N TOTUM 24 CASE", "TOOTN TOTUM 24 CASE",
    ],
    "Spring House 24 Case": [
        "SPRING HOUSE 24", "SPRING HOUSE 24 CASE", "SH 24", "SPRINGHOUSE 24", "SPRINGHOUSE 24 CASE",
    ]
}

_RE_TRAILING_ZEROS = re.compile(r"^([0-9]+)\.0+$")
_RE_NON_ALNUM = re.compile(r"[^A-Z0-9]")
_RE_PO_PREFIX = re.compile(r"^P\.?\s*O\.?(?:\s*[#:\-]\s*|\s+|(?=\d))")
_RE_FISCAL_PERIOD = re.compile(r"(?:PD|P|PERIOD)?\s*(\d{1,2})(?:\.0+)?(?:\s*[-/]\s*(\d{4}|\d{2}))?")


@dataclass
class MatchGroup:
    qb_rows: list[int]
    inf_rows: list[int]
    method: str
    confidence: str
    explanation: str
    group_level: bool = False
    match_id: str = ""


@dataclass
class ReconciliationResult:
    run_id: str
    run_timestamp: datetime
    qb_raw: pd.DataFrame
    inf_raw: pd.DataFrame
    qb_work: pd.DataFrame
    inf_work: pd.DataFrame
    matches: list[MatchGroup]
    paired_rows: list[dict[str, Any]]
    candidates: pd.DataFrame
    normalization: pd.DataFrame
    assessments: pd.DataFrame
    method_summary: pd.DataFrame
    exception_analysis: pd.DataFrame
    amount_variance_analysis: pd.DataFrame
    duplicate_analysis: pd.DataFrame
    infinium_duplicate_analysis: pd.DataFrame
    product_summary: pd.DataFrame
    controls: pd.DataFrame
    metrics: dict[str, Any]
    rules: pd.DataFrame
    config: pd.DataFrame
    qb_mapping: dict[str, Optional[str]]
    inf_mapping: dict[str, Optional[str]]
    metadata: dict[str, Any]
    unmatched_qb: list[int] = field(default_factory=list)
    unmatched_inf: list[int] = field(default_factory=list)
    historical_clearances: pd.DataFrame = field(default_factory=pd.DataFrame)
    qb_secondary_raw: Optional[pd.DataFrame] = None
    inf_secondary_raw: Optional[pd.DataFrame] = None
    qb_secondary_work: Optional[pd.DataFrame] = None
    inf_secondary_work: Optional[pd.DataFrame] = None
    qb_secondary_mapping: Optional[dict[str, Optional[str]]] = None
    inf_secondary_mapping: Optional[dict[str, Optional[str]]] = None
    duplicate_qb_rows: list[int] = field(default_factory=list)
    duplicate_inf_rows: list[int] = field(default_factory=list)
    duplicate_qb_secondary_rows: list[int] = field(default_factory=list)
    duplicate_inf_secondary_rows: list[int] = field(default_factory=list)
    suspected_qb_rows: list[int] = field(default_factory=list)
    suspected_inf_rows: list[int] = field(default_factory=list)
    suspected_qb_secondary_rows: list[int] = field(default_factory=list)
    suspected_inf_secondary_rows: list[int] = field(default_factory=list)
    duplicate_review_hold_qb_rows: list[int] = field(default_factory=list)
    duplicate_review_hold_inf_rows: list[int] = field(default_factory=list)
    amount_variance_review_hold_qb_rows: list[int] = field(default_factory=list)
    amount_variance_review_hold_inf_rows: list[int] = field(default_factory=list)
    fuzzy_match_review_hold_analysis: pd.DataFrame = field(default_factory=pd.DataFrame)
    fuzzy_match_review_hold_qb_rows: list[int] = field(default_factory=list)
    fuzzy_match_review_hold_inf_rows: list[int] = field(default_factory=list)


@lru_cache(maxsize=4096)
def _cached_clean_alphanumeric(text: str) -> str:
    text = _RE_TRAILING_ZEROS.sub(r"\1", text)
    return _RE_NON_ALNUM.sub("", text)


def clean_alphanumeric(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return _cached_clean_alphanumeric(str(value).strip().upper())


@lru_cache(maxsize=4096)
def _cached_clean_po(text: str) -> str:
    text = _RE_TRAILING_ZEROS.sub(r"\1", text)
    text = _RE_PO_PREFIX.sub("", text)
    return _RE_NON_ALNUM.sub("", text)


def clean_po(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return _cached_clean_po(str(value).strip().upper())


def parse_amount_cents(value: Any) -> Optional[int]:
    """Return exact signed cents, or None when the source amount is invalid."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, bool):
        return None
    text = str(value).strip()
    if not text:
        return None
    negative_parentheses = text.startswith("(") and text.endswith(")")
    text = text.replace("$", "").replace(",", "").replace(" ", "")
    if negative_parentheses:
        text = "-" + text[1:-1]
    try:
        amount = Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if not amount.is_finite():
            return None
        return int(amount * 100)
    except (InvalidOperation, ValueError, TypeError):
        return None


def cents_to_float(value: Any) -> float:
    if value is None or pd.isna(value):
        return 0.0
    return float(Decimal(int(value)) / Decimal(100))


def cents_or_zero(value: Any) -> int:
    return 0 if value is None or pd.isna(value) else int(value)


def valid_cents(value: Any) -> bool:
    return value is not None and not pd.isna(value)


def _amount_total(frame: pd.DataFrame, indexes: Any) -> int:
    """Return an exact signed-cent total for the selected working rows."""
    return sum(
        cents_or_zero(frame.at[int(idx), AMOUNT_CENTS])
        for idx in indexes
    )


def _gross_amount_total(frame: pd.DataFrame, indexes: Any) -> int:
    """Return an exact absolute-cent total for the selected working rows."""
    return sum(
        abs(cents_or_zero(frame.at[int(idx), AMOUNT_CENTS]))
        for idx in indexes
    )


def _matched_row_indexes(
    matches: list[MatchGroup],
    dataset: str,
) -> list[int]:
    """Return unique primary row indexes consumed by accepted match groups."""
    attribute = "qb_rows" if dataset == "QB" else "inf_rows"
    return sorted(
        {
            int(idx)
            for group in matches
            for idx in getattr(group, attribute)
        }
    )


def _historical_row_indexes(
    clearances: pd.DataFrame,
    primary_dataset: str,
    index_column: str,
) -> list[int]:
    """Return unique populated indexes from one historical-clearance side."""
    if clearances.empty:
        return []
    values = clearances.loc[
        clearances["Primary Dataset"].eq(primary_dataset)
        & clearances[index_column].notna(),
        index_column,
    ]
    return sorted({int(value) for value in values})


def numeric_sum(series: pd.Series) -> float:
    return cents_to_float(sum(int(c) for c in series.map(parse_amount_cents) if valid_cents(c)))


def numeric_quantity_sum(series: pd.Series) -> float:
    return float(pd.to_numeric(series, errors="coerce").fillna(0).sum())


@lru_cache(maxsize=1)
def _product_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    for standard_name, variants in PRODUCT_LEXICON.items():
        lookup[clean_alphanumeric(standard_name)] = standard_name
        for variant in variants:
            lookup[clean_alphanumeric(variant)] = standard_name
    return lookup


@lru_cache(maxsize=1024)
def _cached_fuzzy_match(cleaned: str) -> Optional[str]:
    lookup = _product_lookup()
    if cleaned in lookup:
        return lookup[cleaned]
    close = get_close_matches(cleaned, lookup.keys(), n=1, cutoff=0.82)
    return lookup[close[0]] if close else None


def get_fuzzy_lexicon_match(value: Any) -> Optional[str]:
    """Classify products for reporting without affecting financial matches."""
    cleaned = clean_alphanumeric(value)
    if not cleaned:
        return None
    return _cached_fuzzy_match(cleaned)


def product_match(value: Any) -> Optional[str]:
    return get_fuzzy_lexicon_match(value)


@lru_cache(maxsize=1024)
def _cached_parse_fiscal_period(text: str, default_year: int) -> tuple[Optional[int], Optional[int]]:
    match = _RE_FISCAL_PERIOD.fullmatch(text)
    if not match:
        return None, None
    period = int(match.group(1))
    year_text = match.group(2)
    year = int(year_text) if year_text else int(default_year)
    if year_text and len(year_text) == 2:
        year += 2000
    if not 1 <= period <= 13 or not 1900 <= year <= 2199:
        return None, None
    return period, year


def parse_fiscal_period(value: Any, default_year: int) -> tuple[Optional[int], Optional[int]]:
    if value is None or pd.isna(value):
        return None, None
    return _cached_parse_fiscal_period(str(value).strip().upper(), default_year)


def fiscal_label(value: Any, default_year: int) -> str:
    period, year = parse_fiscal_period(value, default_year)
    return f"P{period:02d}-{year}" if period is not None else "Unspecified"


def prepare_working_frame(
    raw: pd.DataFrame,
    mapping: dict[str, Optional[str]],
    source: str,
    fiscal_year: int,
    id_prefix: Optional[str] = None,
) -> pd.DataFrame:
    frame = raw.copy().reset_index(drop=True)
    frame[SOURCE_POS] = range(len(frame))
    id_column = QB_ID if source == "QB" else INF_ID
    prefix = id_prefix or source
    frame[id_column] = [f"{prefix}-{row + 1}" for row in range(len(frame))]
    frame[NORM_PO] = frame[mapping["po"]].map(clean_po)
    frame[NORM_INV] = frame[mapping["invoice"]].map(clean_alphanumeric)
    frame[AMOUNT_CENTS] = frame[mapping["amount"]].map(parse_amount_cents)
    frame[PO_TOKENS] = frame[mapping["po"]].map(significant_po_tokens)
    if source == "QB":
        product_col = mapping.get("product")
        period_col = mapping.get("period")
        frame[PRODUCT_STANDARD] = frame[product_col].map(product_match) if product_col else None
        if period_col:
            frame[FISCAL_LABEL] = frame[period_col].map(lambda value: fiscal_label(value, fiscal_year))
        else:
            frame[FISCAL_LABEL] = "Unspecified"
    return frame


def _subset_solutions_for_target(
    frame: pd.DataFrame,
    indexes: list[int],
    target_cents: int,
    solution_limit: int = 2,
) -> list[tuple[int, ...]]:
    """Return up to ``solution_limit`` exact multi-row subsets.

    Retaining at most two paths per subtotal is sufficient because grouped
    matching only needs to distinguish a unique solution from an ambiguous
    one. Pools outside the documented bounds are left for manual review.
    """
    usable = [
        int(idx) for idx in indexes
        if valid_cents(frame.at[idx, AMOUNT_CENTS])
    ]
    usable.sort(key=lambda idx: frame.at[idx, SOURCE_POS])
    if len(usable) < 2 or len(usable) > MAX_GROUP_POOL_ROWS:
        return []

    states: dict[tuple[int, int], list[tuple[int, ...]]] = {(0, 0): [()]}
    for idx in usable:
        amount = int(frame.at[idx, AMOUNT_CENTS])
        additions: dict[tuple[int, int], list[tuple[int, ...]]] = defaultdict(list)
        for (subtotal, size), paths in list(states.items()):
            if size >= MAX_GROUP_SIZE:
                continue
            key = (subtotal + amount, size + 1)
            for path in paths:
                candidate = (*path, idx)
                if candidate not in additions[key]:
                    additions[key].append(candidate)
                if len(additions[key]) >= solution_limit:
                    break
        for key, paths in additions.items():
            existing = states.setdefault(key, [])
            for path in paths:
                if path not in existing:
                    existing.append(path)
                if len(existing) >= solution_limit:
                    break

    solutions: list[tuple[int, ...]] = []
    for size in range(2, MAX_GROUP_SIZE + 1):
        for path in states.get((int(target_cents), size), []):
            if path not in solutions:
                solutions.append(path)
            if len(solutions) >= solution_limit:
                return solutions
    return solutions


def _reference_groups(
    frame: pd.DataFrame,
    remaining: set[int],
    fields: tuple[str, ...],
) -> dict[tuple[str, ...], list[int]]:
    groups: dict[tuple[str, ...], list[int]] = defaultdict(list)
    for idx in sorted(remaining, key=lambda row: frame.at[row, SOURCE_POS]):
        if not valid_cents(frame.at[idx, AMOUNT_CENTS]):
            continue
        key = tuple(str(frame.at[idx, field]) for field in fields)
        if any(not value for value in key):
            continue
        groups[key].append(int(idx))
    return groups


def _group_candidates(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    remaining_q: set[int],
    remaining_i: set[int],
    fields: tuple[str, ...],
) -> list[tuple[tuple[int, ...], tuple[int, ...]]]:
    """Build exact one-to-many and many-to-one candidates for one rule pass."""
    q_groups = _reference_groups(qb, remaining_q, fields)
    i_groups = _reference_groups(inf, remaining_i, fields)
    candidates: set[tuple[tuple[int, ...], tuple[int, ...]]] = set()

    for key in sorted(set(q_groups).intersection(i_groups)):
        q_indexes = q_groups[key]
        i_indexes = i_groups[key]
        q_solutions_by_target: dict[int, list[tuple[int, ...]]] = {}
        i_solutions_by_target: dict[int, list[tuple[int, ...]]] = {}
        for iidx in i_indexes:
            target = int(inf.at[iidx, AMOUNT_CENTS])
            if target not in q_solutions_by_target:
                q_solutions_by_target[target] = _subset_solutions_for_target(
                    qb, q_indexes, target
                )
            for q_subset in q_solutions_by_target[target]:
                candidates.add((tuple(q_subset), (int(iidx),)))
        for qidx in q_indexes:
            target = int(qb.at[qidx, AMOUNT_CENTS])
            if target not in i_solutions_by_target:
                i_solutions_by_target[target] = _subset_solutions_for_target(
                    inf, i_indexes, target
                )
            for i_subset in i_solutions_by_target[target]:
                candidates.add(((int(qidx),), tuple(i_subset)))

    return sorted(
        candidates,
        key=lambda pair: (
            min(qb.at[idx, SOURCE_POS] for idx in pair[0]),
            min(inf.at[idx, SOURCE_POS] for idx in pair[1]),
            pair,
        ),
    )


def perform_matching(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    *,
    enable_fuzzy: bool = True,
) -> tuple[list[MatchGroup], list[int], list[int], pd.DataFrame]:
    """Match one-to-one first, then unique exact aggregate relationships."""
    remaining_q = set(qb.index)
    remaining_i = set(inf.index)
    matches: list[MatchGroup] = []

    passes = [
        ((NORM_PO, NORM_INV, AMOUNT_CENTS), "PO + Invoice + Amount", "Strong",
         "Unique normalized PO, invoice, and signed amount agree."),
        ((NORM_PO, AMOUNT_CENTS), "PO + Amount", "Moderate",
         "Unique normalized PO and signed amount agree."),
        ((NORM_INV, AMOUNT_CENTS), "Invoice + Amount", "Moderate",
         "Unique normalized invoice and signed amount agree."),
    ]

    def unique_rows(
        frame: pd.DataFrame,
        remaining: set[int],
        fields: tuple[str, ...],
        index_column: str,
    ) -> pd.DataFrame:
        if not remaining:
            return pd.DataFrame(columns=[*fields, index_column])
        subset = frame.loc[sorted(remaining), list(fields)].copy()
        valid = subset.notna().all(axis=1) & subset.ne("").all(axis=1)
        subset = subset.loc[valid]
        subset[index_column] = subset.index
        return subset.drop_duplicates(subset=list(fields), keep=False)

    for fields, method, confidence, explanation in passes:
        q_unique = unique_rows(qb, remaining_q, fields, "__QB_INDEX")
        i_unique = unique_rows(inf, remaining_i, fields, "__INF_INDEX")
        if q_unique.empty or i_unique.empty:
            continue
        pairs = q_unique.merge(
            i_unique,
            on=list(fields),
            how="inner",
            sort=False,
            validate="one_to_one",
        ).sort_values("__QB_INDEX")
        pair_indexes = list(
            pairs[["__QB_INDEX", "__INF_INDEX"]].itertuples(index=False, name=None)
        )
        matches.extend(
            MatchGroup([int(qidx)], [int(iidx)], method, confidence, explanation)
            for qidx, iidx in pair_indexes
        )
        remaining_q.difference_update(int(qidx) for qidx, _ in pair_indexes)
        remaining_i.difference_update(int(iidx) for _, iidx in pair_indexes)

    grouped_passes = [
        (
            (NORM_PO, NORM_INV),
            "PO + Invoice + Aggregate Amount (Grouped)",
            "Moderate",
            "After all one-to-one passes, one unique group shares the normalized "
            "PO and invoice and agrees to the opposing row's exact signed-cent total.",
        ),
        (
            (NORM_PO,),
            "PO + Aggregate Amount (Grouped)",
            "Moderate",
            "After all one-to-one passes, one unique group shares the normalized PO "
            "and agrees to the opposing row's exact signed-cent total.",
        ),
        (
            (NORM_INV,),
            "Invoice + Aggregate Amount (Grouped)",
            "Moderate",
            "After all one-to-one passes, one unique group shares the normalized "
            "invoice and agrees to the opposing row's exact signed-cent total.",
        ),
    ]

    for fields, method, confidence, explanation in grouped_passes:
        while remaining_q and remaining_i:
            possible = _group_candidates(
                qb, inf, remaining_q, remaining_i, fields
            )
            if not possible:
                break
            q_occurrences = Counter(idx for q_rows, _ in possible for idx in q_rows)
            i_occurrences = Counter(idx for _, i_rows in possible for idx in i_rows)
            accepted = [
                (q_rows, i_rows)
                for q_rows, i_rows in possible
                if all(q_occurrences[idx] == 1 for idx in q_rows)
                and all(i_occurrences[idx] == 1 for idx in i_rows)
            ]
            if not accepted:
                break
            for q_rows, i_rows in accepted:
                matches.append(
                    MatchGroup(
                        list(q_rows),
                        list(i_rows),
                        method,
                        confidence,
                        explanation,
                        group_level=True,
                    )
                )
                remaining_q.difference_update(q_rows)
                remaining_i.difference_update(i_rows)

    # After every exact one-to-one and grouped-aggregate pass, a small,
    # tightly-bounded fuzzy PO pass covers rows whose PO field is a buyer
    # name or ad hoc note rather than a clean PO number (e.g. QuickBooks
    # "Hopper" vs Infinium "DAVID HOPPER 2.2"). The signed amount must
    # still agree exactly, and only unique whole-word-containment matches
    # are accepted -- see fuzzy_po_matching.py for the full rule. Disabled
    # when matching historical/secondary data against the primary population
    # -- a text-similarity guess is a much bigger risk against a population
    # that isn't the accrual source of truth, so it's simply not attempted
    # there rather than held for review.
    if enable_fuzzy:
        fuzzy_groups = find_fuzzy_po_matches(qb, inf, remaining_q, remaining_i)
        for q_group, i_group in fuzzy_groups:
            matches.append(
                MatchGroup(
                    list(q_group), list(i_group), FUZZY_PO_METHOD, FUZZY_PO_CONFIDENCE,
                    FUZZY_PO_EXPLANATION, group_level=len(q_group) > 1 or len(i_group) > 1,
                )
            )
            remaining_q.difference_update(q_group)
            remaining_i.difference_update(i_group)

    matches.sort(
        key=lambda group: min(
            qb.at[idx, SOURCE_POS] for idx in group.qb_rows
        ) if group.qb_rows else 10**12
    )
    for number, group in enumerate(matches, 1):
        group.match_id = f"M-{number:06d}"

    candidate_columns = [
        "QuickBooks Row ID", "Exception Cause", "Cause Confidence",
        "Disposition", "Total Candidate Count",
        "Available Candidate Count", "Available Infinium Candidate IDs",
        "Already-Matched Candidate IDs", "Minimum Amount Difference",
    ]
    q_rows = sorted(remaining_q)
    if not q_rows:
        candidates = pd.DataFrame(columns=candidate_columns)
        return matches, q_rows, sorted(remaining_i), candidates

    inf_reference = inf[[INF_ID, NORM_PO, NORM_INV, AMOUNT_CENTS]].copy()
    inf_reference["__IIDX"] = inf_reference.index

    def reference_groups(field: str) -> dict[str, list[int]]:
        populated = inf_reference.loc[inf_reference[field].ne("")]
        if populated.empty:
            return {}
        return populated.groupby(field, sort=False)["__IIDX"].agg(list).to_dict()

    po_groups = reference_groups(NORM_PO)
    invoice_groups = reference_groups(NORM_INV)

    def format_candidate_ids(indexes: list[int]) -> str:
        display_limit = 25
        displayed = [str(inf.at[idx, INF_ID]) for idx in indexes[:display_limit]]
        if len(indexes) > display_limit:
            displayed.append(f"... ({len(indexes) - display_limit} more)")
        return "; ".join(displayed)

    candidate_records: list[dict[str, Any]] = []
    for qidx in q_rows:
        candidate_indexes = set(po_groups.get(qb.at[qidx, NORM_PO], [])) if qb.at[qidx, NORM_PO] else set()
        if qb.at[qidx, NORM_INV]:
            candidate_indexes.update(invoice_groups.get(qb.at[qidx, NORM_INV], []))
        available = sorted(candidate_indexes.intersection(remaining_i))
        used = sorted(candidate_indexes.difference(remaining_i))
        qamount = qb.at[qidx, AMOUNT_CENTS]
        differences = [
            abs(int(qamount) - int(inf.at[iidx, AMOUNT_CENTS]))
            for iidx in available
            if valid_cents(qamount) and valid_cents(inf.at[iidx, AMOUNT_CENTS])
        ]
        minimum_difference = min(differences) if differences else None
        if not valid_cents(qamount):
            reason = "Invalid or missing QuickBooks amount"
            exception_cause, cause_confidence = "Invalid QuickBooks amount", "High"
        elif len(available) > 1:
            reason = "Duplicate or ambiguous Infinium values - multiple candidates"
            exception_cause, cause_confidence = "Potential Duplicate - Multiple records share the PO and/or Invoice", "Review"
        elif len(available) == 1 and minimum_difference == 0:
            reason = "Duplicate or ambiguous values - exact candidate is not uniquely one-to-one"
            exception_cause, cause_confidence = "Potential Duplicate - Exact candidate is not uniquely available", "Review"
        elif len(available) == 1 and minimum_difference is not None:
            reason = "Potential amount variance - reference agrees but amount differs"
            exception_cause, cause_confidence = "Potential Typo - Matching PO and/or Invoice values have different amounts", "Review"
        elif len(available) == 1:
            reason = "Reference candidate has an invalid or missing amount"
            exception_cause, cause_confidence = "Invalid Infinium candidate amount", "High"
        elif used:
            reason = "Potential duplicate - referenced Infinium candidate is already matched"
            exception_cause, cause_confidence = "Potential Duplicate - Reference already used by another match", "Review"
        else:
            reason = "No matching Infinium records"
            exception_cause, cause_confidence = "No corresponding Infinium record", "High"
        candidate_records.append(
            {
                "QuickBooks Row ID": qb.at[qidx, QB_ID],
                "Exception Cause": exception_cause,
                "Cause Confidence": cause_confidence,
                "Disposition": reason,
                "Total Candidate Count": len(candidate_indexes),
                "Available Candidate Count": len(available),
                "Available Infinium Candidate IDs": format_candidate_ids(available),
                "Already-Matched Candidate IDs": format_candidate_ids(used),
                "Minimum Amount Difference": cents_to_float(minimum_difference)
                if minimum_difference is not None else None,
            }
        )
    candidates = pd.DataFrame(candidate_records, columns=candidate_columns)
    return matches, sorted(remaining_q), sorted(remaining_i), candidates


AMOUNT_VARIANCE_COLUMNS = [
    "Variance ID", "Classification", "Confidence", "Reference Evidence",
    "QuickBooks Row ID", "QuickBooks Row Index", "Infinium Row ID",
    "Infinium Row Index", "Normalized PO", "Normalized Invoice",
    "QuickBooks Amount", "Infinium Amount", "Potential Difference",
    "Absolute Difference", "Variance Direction", "Possible Sign Reversal",
    "Mutually Unique Reference", "Accrual Treatment", "Posting Disposition",
    "Explanation", "Manual Decision", "Reviewed By", "Review Timestamp",
    "Review Rationale",
]


def build_reference_amount_variances(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    unmatched_qb: list[int],
    unmatched_inf: list[int],
) -> tuple[pd.DataFrame, list[int], list[int]]:
    """Identify mutually unique reference matches whose signed amounts differ.

    This is a classification and posting-control pass, never an automatic
    correction. It runs only after primary and historical matching are final.
    A relationship is accepted into the variance-review population only when:

      * both rows remain unresolved;
      * at least one populated normalized PO or invoice agrees exactly;
      * the candidate is unique from QuickBooks to Infinium and Infinium to
        QuickBooks within the remaining populations; and
      * the signed-cent amounts are both valid and differ.

    The full QuickBooks amount and the calculated difference are displayed,
    but neither is automatically posted. Both source rows are moved out of the
    ordinary exception populations and the run remains non-postable until a
    documented disposition is completed and the reconciliation is rerun.
    """
    q_indexes = sorted(set(int(idx) for idx in unmatched_qb))
    i_indexes = sorted(set(int(idx) for idx in unmatched_inf))
    if not q_indexes or not i_indexes:
        return pd.DataFrame(columns=AMOUNT_VARIANCE_COLUMNS), [], []

    # Establish relationships in evidence-strength order. Rows assigned under
    # a stronger tier are removed before the next tier so a reused PO cannot
    # defeat an otherwise unique PO+invoice relationship.
    selected_pairs: list[tuple[int, int, str]] = []
    available_q = set(q_indexes)
    available_i = set(i_indexes)

    def add_unique_tier(fields: tuple[str, ...], tier: str) -> None:
        q_groups = _reference_groups(qb, available_q, fields)
        i_groups = _reference_groups(inf, available_i, fields)
        tier_pairs: list[tuple[int, int, str]] = []
        for key in sorted(set(q_groups).intersection(i_groups)):
            q_rows = q_groups[key]
            i_rows = i_groups[key]
            if len(q_rows) != 1 or len(i_rows) != 1:
                continue
            qidx, iidx = q_rows[0], i_rows[0]
            q_amount = qb.at[qidx, AMOUNT_CENTS]
            i_amount = inf.at[iidx, AMOUNT_CENTS]
            if (
                valid_cents(q_amount)
                and valid_cents(i_amount)
                and int(q_amount) != int(i_amount)
            ):
                tier_pairs.append((qidx, iidx, tier))
        for qidx, iidx, evidence_tier in tier_pairs:
            if qidx in available_q and iidx in available_i:
                selected_pairs.append((qidx, iidx, evidence_tier))
                available_q.remove(qidx)
                available_i.remove(iidx)

    add_unique_tier((NORM_PO, NORM_INV), "PO_AND_INVOICE")
    add_unique_tier((NORM_INV,), "INVOICE_ONLY")
    add_unique_tier((NORM_PO,), "PO_ONLY")

    records: list[dict[str, Any]] = []
    held_q: list[int] = []
    held_i: list[int] = []
    for qidx, iidx, evidence_tier in selected_pairs:
        q_po, i_po = qb.at[qidx, NORM_PO], inf.at[iidx, NORM_PO]
        q_inv, i_inv = qb.at[qidx, NORM_INV], inf.at[iidx, NORM_INV]
        po_agrees = bool(q_po and q_po == i_po)
        invoice_agrees = bool(q_inv and q_inv == i_inv)
        if evidence_tier == "PO_AND_INVOICE":
            classification = "High-likelihood amount variance"
            confidence = "High"
            evidence = "Exact PO and invoice; mutually unique unresolved rows"
        elif evidence_tier == "INVOICE_ONLY":
            classification = "Strong invoice-linked amount variance"
            confidence = "Strong review"
            evidence = "Exact invoice; mutually unique unresolved rows"
        else:
            classification = "PO-linked amount variance"
            confidence = "Moderate review"
            evidence = "Exact PO; mutually unique unresolved rows"
        q_cents = int(qb.at[qidx, AMOUNT_CENTS])
        i_cents = int(inf.at[iidx, AMOUNT_CENTS])
        difference = q_cents - i_cents
        sign_reversal = q_cents == -i_cents and q_cents != 0
        if sign_reversal:
            classification = "Critical possible sign reversal"
            confidence = "High"
        direction = (
            "QuickBooks exceeds Infinium" if difference > 0
            else "Infinium exceeds QuickBooks"
        )
        records.append({
            "Variance ID": f"VAR-{len(records) + 1:06d}",
            "Classification": classification,
            "Confidence": confidence,
            "Reference Evidence": evidence,
            "QuickBooks Row ID": qb.at[qidx, QB_ID],
            "QuickBooks Row Index": qidx,
            "Infinium Row ID": inf.at[iidx, INF_ID],
            "Infinium Row Index": iidx,
            "Normalized PO": q_po if po_agrees else "",
            "Normalized Invoice": q_inv if invoice_agrees else "",
            "QuickBooks Amount": cents_to_float(q_cents),
            "Infinium Amount": cents_to_float(i_cents),
            "Potential Difference": cents_to_float(difference),
            "Absolute Difference": cents_to_float(abs(difference)),
            "Variance Direction": direction,
            "Possible Sign Reversal": sign_reversal,
            "Mutually Unique Reference": True,
            "Accrual Treatment": (
                "Excluded from automatic JE; neither the full amount nor the "
                "difference is posted without documented review"
            ),
            "Posting Disposition": "REVIEW REQUIRED - DO NOT POST",
            "Explanation": (
                "The references indicate a likely common transaction, but signed "
                "amounts differ. The program does not infer which system is correct."
            ),
            "Manual Decision": None,
            "Reviewed By": None,
            "Review Timestamp": None,
            "Review Rationale": None,
        })
        held_q.append(qidx)
        held_i.append(iidx)

    return (
        pd.DataFrame(records, columns=AMOUNT_VARIANCE_COLUMNS),
        sorted(held_q),
        sorted(held_i),
    )


FUZZY_MATCH_REVIEW_COLUMNS = [
    "Fuzzy Match ID", "Classification", "Confidence", "Match Basis",
    "QuickBooks Row IDs", "QuickBooks Row Indexes", "Infinium Row IDs",
    "Infinium Row Indexes", "QuickBooks Row Count", "Infinium Row Count",
    "Matched Text Evidence", "QuickBooks Amount", "Infinium Amount",
    "Accrual Treatment", "Posting Disposition", "Explanation",
    "Manual Decision", "Reviewed By", "Review Timestamp", "Review Rationale",
]

FUZZY_MATCH_CLASSIFICATION_SINGLE = "Fuzzy Match - Possible Text Variant of PO or Invoice"
FUZZY_MATCH_CLASSIFICATION_GROUPED = "Fuzzy Match - Multiple Line Items Netting to One Total"
FUZZY_MATCH_CONFIDENCE = "Review"


def build_fuzzy_match_review_holds(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    fuzzy_groups: list[MatchGroup],
) -> pd.DataFrame:
    """Report every fuzzy PO/text match held for a documented human decision.

    A fuzzy match is a text-similarity guess, not a certain relationship --
    unlike an exact match it is never posted automatically. This produces the
    plain-English, auditor-facing report; the caller is responsible for
    removing these rows from ``matches`` and treating them as held rather
    than accepted (see ``build_reconciliation``).
    """
    records: list[dict[str, Any]] = []
    for number, group in enumerate(fuzzy_groups, 1):
        q_rows = sorted(group.qb_rows, key=lambda idx: qb.at[idx, SOURCE_POS])
        i_rows = sorted(group.inf_rows, key=lambda idx: inf.at[idx, SOURCE_POS])
        grouped = len(q_rows) > 1 or len(i_rows) > 1
        classification = (
            FUZZY_MATCH_CLASSIFICATION_GROUPED if grouped else FUZZY_MATCH_CLASSIFICATION_SINGLE
        )
        shared_tokens: set[str] = set()
        for qidx in q_rows:
            for iidx in i_rows:
                shared_tokens |= qb.at[qidx, PO_TOKENS] & inf.at[iidx, PO_TOKENS]
        q_total = _amount_total(qb, q_rows)
        i_total = _amount_total(inf, i_rows)
        records.append({
            "Fuzzy Match ID": f"FUZZY-{number:06d}",
            "Classification": classification,
            "Confidence": FUZZY_MATCH_CONFIDENCE,
            "Match Basis": (
                f"Shared PO/text tokens ({', '.join(sorted(shared_tokens)) or 'n/a'}) "
                "and exact aggregate amount tie-out"
            ),
            "QuickBooks Row IDs": "; ".join(str(qb.at[idx, QB_ID]) for idx in q_rows),
            "QuickBooks Row Indexes": "; ".join(str(idx) for idx in q_rows),
            "Infinium Row IDs": "; ".join(str(inf.at[idx, INF_ID]) for idx in i_rows),
            "Infinium Row Indexes": "; ".join(str(idx) for idx in i_rows),
            "QuickBooks Row Count": len(q_rows),
            "Infinium Row Count": len(i_rows),
            "Matched Text Evidence": ", ".join(sorted(shared_tokens)) or "n/a",
            "QuickBooks Amount": cents_to_float(q_total),
            "Infinium Amount": cents_to_float(i_total),
            "Accrual Treatment": (
                "Excluded from automatic JE; a text-similarity match is never posted "
                "without documented review"
            ),
            "Posting Disposition": "REVIEW REQUIRED - DO NOT POST",
            "Explanation": (
                f"{group.explanation} Neither amount is automatically posted pending "
                "documented review."
            ),
            "Manual Decision": None,
            "Reviewed By": None,
            "Review Timestamp": None,
            "Review Rationale": None,
        })
    return pd.DataFrame(records, columns=FUZZY_MATCH_REVIEW_COLUMNS)


HISTORICAL_CLEARANCE_COLUMNS = [
    "Clearance ID",
    "Group Sequence",
    "Primary Row Count",
    "Secondary Row Count",
    "Group-Level Match",
    "Primary Dataset",
    "Primary Row ID",
    "Primary Row Index",
    "Secondary Dataset",
    "Secondary Row ID",
    "Secondary Row Index",
    "Match Method",
    "Confidence",
    "Normalized PO",
    "Normalized Invoice",
    "Primary Amount",
    "Secondary Amount",
    "Amount Difference",
    "Disposition",
]


def build_historical_clearances(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    unmatched_qb: list[int],
    unmatched_inf: list[int],
    qb_secondary: Optional[pd.DataFrame],
    inf_secondary: Optional[pd.DataFrame],
) -> tuple[pd.DataFrame, list[int], list[int]]:
    """Clear opposing-primary exceptions with optional historical sources.

    Historical rows never become exceptions themselves. Accepted one-to-one
    and controlled grouped matches are retained as evidence; all unused
    historical rows are discarded from the reconciliation population.

    ``qb_secondary``/``inf_secondary`` must already be duplicate-screened by
    the caller (see ``duplicates.screen_duplicates``) -- a duplicated
    historical row is just as capable of improperly clearing a real primary
    exception as a duplicated primary row is of misstating the accrual, so
    neither may reach this function.
    """
    records: list[dict[str, Any]] = []
    cleared_qb: set[int] = set()
    cleared_inf: set[int] = set()
    clearance_number = 0

    def append_clearance(
        group: MatchGroup,
        primary_frame: pd.DataFrame,
        secondary_frame: pd.DataFrame,
        primary_indexes: list[int],
        secondary_indexes: list[int],
        primary_dataset: str,
        secondary_dataset: str,
        primary_id_column: str,
        secondary_id_column: str,
        method_prefix: str,
        disposition: str,
    ) -> None:
        nonlocal clearance_number
        clearance_number += 1
        clearance_id = f"H-{clearance_number:06d}"
        ordered_primary = sorted(
            (int(idx) for idx in primary_indexes),
            key=lambda idx: primary_frame.at[idx, SOURCE_POS],
        )
        ordered_secondary = sorted(
            (int(idx) for idx in secondary_indexes),
            key=lambda idx: secondary_frame.at[idx, SOURCE_POS],
        )
        primary_total = _amount_total(primary_frame, ordered_primary)
        secondary_total = _amount_total(secondary_frame, ordered_secondary)
        for sequence, (pidx, sidx) in enumerate(
            zip_longest(ordered_primary, ordered_secondary), 1
        ):
            primary_amount = (
                cents_or_zero(primary_frame.at[pidx, AMOUNT_CENTS])
                if pidx is not None else 0
            )
            secondary_amount = (
                cents_or_zero(secondary_frame.at[sidx, AMOUNT_CENTS])
                if sidx is not None else 0
            )
            reference_frame = primary_frame if pidx is not None else secondary_frame
            reference_index = pidx if pidx is not None else sidx
            records.append(
                {
                    "Clearance ID": clearance_id,
                    "Group Sequence": sequence,
                    "Primary Row Count": len(ordered_primary),
                    "Secondary Row Count": len(ordered_secondary),
                    "Group-Level Match": group.group_level,
                    "Primary Dataset": primary_dataset,
                    "Primary Row ID": (
                        primary_frame.at[pidx, primary_id_column]
                        if pidx is not None else ""
                    ),
                    "Primary Row Index": int(pidx) if pidx is not None else None,
                    "Secondary Dataset": secondary_dataset,
                    "Secondary Row ID": (
                        secondary_frame.at[sidx, secondary_id_column]
                        if sidx is not None else ""
                    ),
                    "Secondary Row Index": int(sidx) if sidx is not None else None,
                    "Match Method": f"{method_prefix} | {group.method}",
                    "Confidence": group.confidence,
                    "Normalized PO": reference_frame.at[reference_index, NORM_PO],
                    "Normalized Invoice": reference_frame.at[reference_index, NORM_INV],
                    "Primary Amount": cents_to_float(primary_amount),
                    "Secondary Amount": cents_to_float(secondary_amount),
                    "Amount Difference": cents_to_float(primary_amount - secondary_amount),
                    "Disposition": disposition,
                }
            )
        if primary_total != secondary_total:
            raise ValueError(
                "Historical clearance construction failure: aggregate amounts differ."
            )

    if inf_secondary is not None and len(inf_secondary) and unmatched_qb:
        secondary_matches, _, _, _ = perform_matching(
            qb.loc[unmatched_qb].copy(), inf_secondary, enable_fuzzy=False
        )
        for group in secondary_matches:
            cleared_qb.update(int(idx) for idx in group.qb_rows)
            append_clearance(
                group,
                qb,
                inf_secondary,
                group.qb_rows,
                group.inf_rows,
                "QuickBooks Primary",
                "Infinium Secondary (Historical)",
                QB_ID,
                INF_ID,
                "QuickBooks primary ↔ Infinium prior-period match",
                (
                    "QuickBooks primary item(s) cleared by the displayed Infinium "
                    "prior-period item(s)"
                ),
            )

    if qb_secondary is not None and len(qb_secondary) and unmatched_inf:
        secondary_matches, _, _, _ = perform_matching(
            qb_secondary, inf.loc[unmatched_inf].copy(), enable_fuzzy=False
        )
        for group in secondary_matches:
            cleared_inf.update(int(idx) for idx in group.inf_rows)
            append_clearance(
                group,
                inf,
                qb_secondary,
                group.inf_rows,
                group.qb_rows,
                "Infinium Primary",
                "QuickBooks Secondary (Historical)",
                INF_ID,
                QB_ID,
                "QuickBooks prior-period match ↔ Infinium primary",
                (
                    "Infinium primary item(s) cleared by the displayed QuickBooks "
                    "prior-period item(s)"
                ),
            )

    clearances = pd.DataFrame(records, columns=HISTORICAL_CLEARANCE_COLUMNS)
    remaining_qb = sorted(set(unmatched_qb).difference(cleared_qb))
    remaining_inf = sorted(set(unmatched_inf).difference(cleared_inf))
    return clearances, remaining_qb, remaining_inf


def _optional_index(value: Any) -> Optional[int]:
    return None if value is None or pd.isna(value) else int(value)


def build_paired_rows(
    matches: list[MatchGroup],
    historical_clearances: pd.DataFrame,
    unmatched_qb: list[int],
    unmatched_inf: list[int],
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    candidates: pd.DataFrame,
    duplicate_qb_rows: list[int],
    duplicate_inf_rows: list[int],
    duplicate_review_hold_qb_rows: list[int],
    duplicate_review_hold_inf_rows: list[int],
    amount_variance_analysis: pd.DataFrame,
    fuzzy_review_hold_groups: list[MatchGroup],
    qb_duplicate_report: Optional[pd.DataFrame] = None,
    inf_duplicate_report: Optional[pd.DataFrame] = None,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    candidate_reason = (
        candidates.set_index("QuickBooks Row ID")["Disposition"].to_dict()
        if not candidates.empty else {}
    )
    candidate_cause = (
        candidates.set_index("QuickBooks Row ID")["Exception Cause"].to_dict()
        if not candidates.empty else {}
    )
    for group in matches:
        ordered_q = sorted(group.qb_rows, key=lambda idx: qb.at[idx, SOURCE_POS])
        ordered_i = sorted(group.inf_rows, key=lambda idx: inf.at[idx, SOURCE_POS])
        for sequence, (qidx, iidx) in enumerate(zip_longest(ordered_q, ordered_i), 1):
            result_label = group.method
            if group.group_level:
                result_label += " [group-level; no line allocation]"
            rows.append(
                {
                    "Section": "01 Matched",
                    "Match ID": group.match_id,
                    "Match Result": result_label,
                    "QB Index": qidx,
                    "Infinium Index": iidx,
                    "QB Record Scope": "Primary",
                    "Infinium Record Scope": "Primary",
                    "Group Sequence": sequence,
                    "Confidence": group.confidence,
                    "Explanation": group.explanation,
                }
            )
    if not historical_clearances.empty:
        for clearance in historical_clearances.to_dict("records"):
            primary_is_qb = clearance["Primary Dataset"] == "QuickBooks Primary"
            primary_index = _optional_index(clearance["Primary Row Index"])
            secondary_index = _optional_index(clearance["Secondary Row Index"])
            qb_index = primary_index if primary_is_qb else secondary_index
            inf_index = secondary_index if primary_is_qb else primary_index
            rows.append(
                {
                    "Section": "01 Matched - Historical Clearance",
                    "Match ID": clearance["Clearance ID"],
                    "Match Result": clearance["Match Method"],
                    "QB Index": qb_index,
                    "Infinium Index": inf_index,
                    "QB Record Scope": (
                        ("Primary" if primary_is_qb else "Historical")
                        if qb_index is not None else None
                    ),
                    "Infinium Record Scope": (
                        ("Historical" if primary_is_qb else "Primary")
                        if inf_index is not None else None
                    ),
                    "Group Sequence": clearance.get("Group Sequence", 1),
                    "Confidence": clearance["Confidence"],
                    "Explanation": clearance["Disposition"],
                }
            )
    for qidx in unmatched_qb:
        qb_id = qb.at[qidx, QB_ID]
        rows.append(
            {
                "Section": "02 Unmatched QuickBooks",
                "Match ID": "",
                "Match Result": candidate_reason.get(qb_id, "Unmatched QuickBooks"),
                "QB Index": qidx,
                "Infinium Index": None,
                "QB Record Scope": "Primary",
                "Infinium Record Scope": None,
                "Group Sequence": None,
                "Confidence": "Review",
                "Explanation": "No unique automatic Infinium match was established.",
            }
        )
    for iidx in unmatched_inf:
        rows.append(
            {
                "Section": "03 Unmatched Infinium",
                "Match ID": "",
                "Match Result": "Unmatched Infinium",
                "QB Index": None,
                "Infinium Index": iidx,
                "QB Record Scope": None,
                "Infinium Record Scope": "Primary",
                "Group Sequence": None,
                "Confidence": "Review",
                "Explanation": "No unique automatic QuickBooks match was established.",
            }
        )
    for qidx in duplicate_qb_rows:
        rows.append(
            {
                "Section": "04 Duplicate QuickBooks",
                "Match ID": "",
                "Match Result": "Excluded excess QuickBooks copy",
                "QB Index": qidx,
                "Infinium Index": None,
                "QB Record Scope": "Primary",
                "Infinium Record Scope": None,
                "Group Sequence": None,
                "Confidence": "Duplicate",
                "Explanation": (
                    "An earlier canonical QuickBooks row with the same normalized PO, "
                    "invoice, and signed cents was retained. Only this excess copy was "
                    "excluded from matching and the proposed journal entry."
                ),
            }
        )
    for iidx in duplicate_inf_rows:
        rows.append(
            {
                "Section": "05 Duplicate Infinium",
                "Match ID": "",
                "Match Result": "Excluded excess Infinium copy",
                "QB Index": None,
                "Infinium Index": iidx,
                "QB Record Scope": None,
                "Infinium Record Scope": "Primary",
                "Group Sequence": None,
                "Confidence": "Duplicate",
                "Explanation": (
                    "An earlier canonical Infinium row with the same normalized PO, "
                    "invoice, and signed cents was retained. Only this excess copy was "
                    "excluded from matching and is reported separately."
                ),
            }
        )
    for qidx in duplicate_review_hold_qb_rows:
        rows.append(
            {
                "Section": "06 Duplicate Review Hold QuickBooks",
                "Match ID": "",
                "Match Result": "Weak-basis duplicate candidate - unresolved",
                "QB Index": qidx,
                "Infinium Index": None,
                "QB Record Scope": "Primary",
                "Infinium Record Scope": None,
                "Group Sequence": None,
                "Confidence": "Hold",
                "Explanation": (
                    "This row shares only a PO or only an invoice (never both) with another "
                    "QuickBooks row at the same signed amount, so it was never excluded "
                    "outright -- it remained active and eligible to match normally. It did "
                    "not match anything, so it is excluded from the proposed JE support total "
                    "and held here pending a documented human disposition."
                ),
            }
        )
    for iidx in duplicate_review_hold_inf_rows:
        rows.append(
            {
                "Section": "07 Duplicate Review Hold Infinium",
                "Match ID": "",
                "Match Result": "Weak-basis duplicate candidate - unresolved",
                "QB Index": None,
                "Infinium Index": iidx,
                "QB Record Scope": None,
                "Infinium Record Scope": "Primary",
                "Group Sequence": None,
                "Confidence": "Hold",
                "Explanation": (
                    "This row shares only a PO or only an invoice (never both) with another "
                    "Infinium row at the same signed amount. It remained active and eligible "
                    "to match normally, did not match anything, and is held here for review. "
                    "Infinium items never feed the QuickBooks accrual regardless."
                ),
            }
        )
    if not amount_variance_analysis.empty:
        for variance in amount_variance_analysis.to_dict("records"):
            rows.append(
                {
                    "Section": "08 Reference-Matched Amount Variance Review Hold",
                    "Match ID": variance["Variance ID"],
                    "Match Result": variance["Classification"],
                    "QB Index": int(variance["QuickBooks Row Index"]),
                    "Infinium Index": int(variance["Infinium Row Index"]),
                    "QB Record Scope": "Primary",
                    "Infinium Record Scope": "Primary",
                    "Group Sequence": 1,
                    "Confidence": variance["Confidence"],
                    "Explanation": (
                        f"{variance['Reference Evidence']}. QuickBooks "
                        f"{variance['QuickBooks Amount']}; Infinium "
                        f"{variance['Infinium Amount']}; potential difference "
                        f"{variance['Potential Difference']}. Neither amount is "
                        "automatically posted pending documented review."
                    ),
                }
            )
    for group in fuzzy_review_hold_groups:
        ordered_q = sorted(group.qb_rows, key=lambda idx: qb.at[idx, SOURCE_POS])
        ordered_i = sorted(group.inf_rows, key=lambda idx: inf.at[idx, SOURCE_POS])
        grouped = len(ordered_q) > 1 or len(ordered_i) > 1
        classification = (
            FUZZY_MATCH_CLASSIFICATION_GROUPED if grouped else FUZZY_MATCH_CLASSIFICATION_SINGLE
        )
        for sequence, (qidx, iidx) in enumerate(zip_longest(ordered_q, ordered_i), 1):
            rows.append(
                {
                    "Section": "09 Fuzzy Match Review Hold",
                    "Match ID": group.match_id,
                    "Match Result": classification,
                    "QB Index": qidx,
                    "Infinium Index": iidx,
                    "QB Record Scope": "Primary" if qidx is not None else None,
                    "Infinium Record Scope": "Primary" if iidx is not None else None,
                    "Group Sequence": sequence,
                    "Confidence": FUZZY_MATCH_CONFIDENCE,
                    "Explanation": (
                        f"{group.explanation} Neither amount is automatically posted "
                        "pending documented review."
                    ),
                }
            )

    def duplicate_lookup(report: Optional[pd.DataFrame]) -> dict[str, dict[str, Any]]:
        if report is None or report.empty:
            return {}
        primary = report.loc[report["Source Scope"].eq("Primary")].copy()
        if primary.empty:
            return {}
        primary["__STAGE_ORDER"] = primary["Screening Stage"].map(
            {"Same-file": 0, "Cross-scope": 1}
        ).fillna(9)
        primary = primary.sort_values(
            ["Source Row ID", "__STAGE_ORDER"], kind="stable"
        )
        return {
            str(source_id): group.iloc[0].to_dict()
            for source_id, group in primary.groupby("Source Row ID", sort=False)
        }

    qb_duplicate_detail = duplicate_lookup(qb_duplicate_report)
    inf_duplicate_detail = duplicate_lookup(inf_duplicate_report)
    variance_detail = (
        amount_variance_analysis.set_index("Variance ID").to_dict("index")
        if not amount_variance_analysis.empty else {}
    )
    unresolved_q_indexes = [
        int(row["QB Index"])
        for row in rows
        if row.get("Section") == "02 Unmatched QuickBooks"
        and row.get("QB Index") is not None
    ]
    unresolved_po_groups: dict[str, list[int]] = defaultdict(list)
    unresolved_invoice_groups: dict[str, list[int]] = defaultdict(list)
    for idx in unresolved_q_indexes:
        if qb.at[idx, NORM_PO]:
            unresolved_po_groups[qb.at[idx, NORM_PO]].append(idx)
        if qb.at[idx, NORM_INV]:
            unresolved_invoice_groups[qb.at[idx, NORM_INV]].append(idx)

    def group_lacks_matching_infinium_total(
        field: str,
        value: str,
        qb_indexes: list[int],
    ) -> bool:
        qb_total = _amount_total(qb, qb_indexes)
        inf_indexes = [
            int(idx) for idx in inf.index
            if inf.at[idx, field] == value and valid_cents(inf.at[idx, AMOUNT_CENTS])
        ]
        return not any(
            int(inf.at[idx, AMOUNT_CENTS]) == qb_total for idx in inf_indexes
        )

    for row in rows:
        section = str(row.get("Section", ""))
        row.update({
            "Exception Cause": "Not an exception",
            "Cause Confidence": row.get("Confidence", ""),
            "Financial Treatment": "No exception accrual treatment",
            "Related Source Row IDs": "",
            "Duplicate Basis": "",
            "Potential Duplicate Reason": "",
            "Differing Confirmation Fields": "",
            "Duplicate Values": "",
            "Duplicate Group ID": "",
            "Confirmed Copy Set ID": "",
            "Canonical Source Row ID": "",
            "Potential Amount Difference": None,
        })
        if section == "02 Unmatched QuickBooks":
            qidx = int(row["QB Index"])
            q_invoice = qb.at[qidx, NORM_INV]
            q_po = qb.at[qidx, NORM_PO]
            invoice_group = unresolved_invoice_groups.get(q_invoice, [])
            po_group = unresolved_po_groups.get(q_po, [])
            if (
                q_invoice
                and len(invoice_group) > 1
                and group_lacks_matching_infinium_total(
                    NORM_INV, q_invoice, invoice_group
                )
            ):
                row["Exception Cause"] = (
                    "Potential Duplicate - Repeated Invoice numbers do not net to a "
                    "matching Infinium value"
                )
            elif (
                q_po
                and len(po_group) > 1
                and group_lacks_matching_infinium_total(NORM_PO, q_po, po_group)
            ):
                row["Exception Cause"] = (
                    "Potential Duplicate - Repeated PO values do not net to a "
                    "matching Infinium value"
                )
            else:
                row["Exception Cause"] = candidate_cause.get(
                    qb.at[qidx, QB_ID], "Unresolved matching exception"
                )
            row["Cause Confidence"] = "Review"
            row["Financial Treatment"] = "Included in provisional QuickBooks accrual support"
        elif section == "03 Unmatched Infinium":
            row["Exception Cause"] = "No corresponding QuickBooks record"
            row["Cause Confidence"] = "Review"
            row["Financial Treatment"] = "Informational; Infinium rows do not feed the QuickBooks accrual"
        elif section in {"04 Duplicate QuickBooks", "05 Duplicate Infinium"}:
            row["Exception Cause"] = "Confirmed Duplicate - Identical transaction copy"
            row["Cause Confidence"] = "High"
            row["Financial Treatment"] = (
                "Excluded from matching and automatic JE; canonical row retained"
            )
        elif section in {
            "06 Duplicate Review Hold QuickBooks",
            "07 Duplicate Review Hold Infinium",
        }:
            row["Exception Cause"] = "Potential Duplicate - Insufficient evidence for automatic exclusion"
            row["Cause Confidence"] = "Hold"
            row["Financial Treatment"] = "Excluded pending documented duplicate disposition"
        elif section == "08 Reference-Matched Amount Variance Review Hold":
            detail = variance_detail.get(str(row.get("Match ID", "")), {})
            row["Exception Cause"] = (
                "Potential Typo - Matching PO and/or Invoice values have different amounts"
            )
            row["Cause Confidence"] = str(detail.get("Confidence", "Review"))
            row["Financial Treatment"] = str(detail.get("Accrual Treatment", "Review hold"))
            row["Related Source Row IDs"] = "; ".join(
                filter(None, [
                    str(detail.get("QuickBooks Row ID", "")),
                    str(detail.get("Infinium Row ID", "")),
                ])
            )
            row["Potential Amount Difference"] = detail.get("Potential Difference")
        elif section == "09 Fuzzy Match Review Hold":
            row["Exception Cause"] = "Potential Match - Pending Manual Confirmation"
            row["Cause Confidence"] = FUZZY_MATCH_CONFIDENCE
            row["Financial Treatment"] = "Excluded pending documented match confirmation"

        detail: dict[str, Any] = {}
        if row.get("QB Index") is not None and row.get("QB Record Scope") == "Primary":
            detail = qb_duplicate_detail.get(str(qb.at[row["QB Index"], QB_ID]), {})
        if not detail and row.get("Infinium Index") is not None and row.get("Infinium Record Scope") == "Primary":
            detail = inf_duplicate_detail.get(str(inf.at[row["Infinium Index"], INF_ID]), {})
        if detail:
            def clean_detail_value(value: Any) -> str:
                return "" if value is None or pd.isna(value) else str(value)

            related = [
                clean_detail_value(detail.get("Other Source Row IDs In Group")),
                clean_detail_value(detail.get("Reference Source Row IDs")),
            ]
            row["Related Source Row IDs"] = "; ".join(
                value for value in related if value
            )
            row["Duplicate Basis"] = clean_detail_value(detail.get("Duplicate Basis"))
            row["Potential Duplicate Reason"] = clean_detail_value(
                detail.get("Potential Duplicate Reason")
            )
            row["Differing Confirmation Fields"] = clean_detail_value(
                detail.get("Differing Confirmation Fields")
            )
            row["Duplicate Values"] = clean_detail_value(detail.get("Duplicate Values"))
            row["Duplicate Group ID"] = clean_detail_value(detail.get("Duplicate Group ID"))
            row["Confirmed Copy Set ID"] = clean_detail_value(detail.get("Confirmed Copy Set ID"))
            row["Canonical Source Row ID"] = clean_detail_value(
                detail.get("Canonical Source Row ID")
            )
            if section in {
                "06 Duplicate Review Hold QuickBooks",
                "07 Duplicate Review Hold Infinium",
            }:
                basis = row["Duplicate Basis"]
                if basis == "Invoice + Amount (PO blank on both rows)":
                    row["Exception Cause"] = (
                        "Potential Duplicate - Repeated Invoice numbers do not net to a "
                        "matching opposing-system value"
                    )
                elif basis == "PO + Amount (invoice blank on both rows)":
                    row["Exception Cause"] = (
                        "Potential Duplicate - Repeated PO values do not net to a "
                        "matching opposing-system value"
                    )
                elif basis == "PO + Invoice + Amount":
                    row["Exception Cause"] = (
                        "Potential Duplicate - PO, Invoice, and Amount match but "
                        "other source fields differ"
                    )
    return rows


def build_normalization_detail(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    qb_mapping: dict[str, Optional[str]],
    inf_mapping: dict[str, Optional[str]],
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for source, frame, mapping, id_col in (
        ("QuickBooks", qb, qb_mapping, QB_ID),
        ("Infinium", inf, inf_mapping, INF_ID),
    ):
        for idx, row in frame.iterrows():
            record = {
                "Dataset": source,
                "Source Row ID": row[id_col],
                "Original PO": row[mapping["po"]],
                "Normalized PO": row[NORM_PO],
                "Original Invoice": row[mapping["invoice"]],
                "Normalized Invoice": row[NORM_INV],
                "Original Amount": row[mapping["amount"]],
                "Normalized Amount": cents_to_float(row[AMOUNT_CENTS])
                if valid_cents(row[AMOUNT_CENTS]) else None,
                "Amount Valid": valid_cents(row[AMOUNT_CENTS]),
            }
            if source == "QuickBooks":
                product_col = mapping.get("product")
                period_col = mapping.get("period")
                record.update(
                    {
                        "Original Product Description": row[product_col] if product_col else None,
                        "Standardized Product": row[PRODUCT_STANDARD],
                        "Original Fiscal Period": row[period_col] if period_col else None,
                        "Fiscal Period Label": row[FISCAL_LABEL],
                    }
                )
            records.append(record)
    return pd.DataFrame(records)


def build_match_assessments(
    matches: list[MatchGroup],
    historical_clearances: pd.DataFrame,
    unmatched_qb: list[int],
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    candidates: pd.DataFrame,
    amount_variance_analysis: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    candidate_map = candidates.set_index("QuickBooks Row ID").to_dict("index") if not candidates.empty else {}
    for group in matches:
        q_po = {qb.at[idx, NORM_PO] for idx in group.qb_rows if qb.at[idx, NORM_PO]}
        i_po = {inf.at[idx, NORM_PO] for idx in group.inf_rows if inf.at[idx, NORM_PO]}
        q_inv = {qb.at[idx, NORM_INV] for idx in group.qb_rows if qb.at[idx, NORM_INV]}
        i_inv = {inf.at[idx, NORM_INV] for idx in group.inf_rows if inf.at[idx, NORM_INV]}
        q_total = _amount_total(qb, group.qb_rows)
        i_total = _amount_total(inf, group.inf_rows)
        records.append(
            {
                "Match ID": group.match_id,
                "Decision": "Matched",
                "Match Method": group.method,
                "Confidence": group.confidence,
                "QuickBooks Row Count": len(group.qb_rows),
                "Infinium Row Count": len(group.inf_rows),
                "QuickBooks Row IDs": "; ".join(qb.at[idx, QB_ID] for idx in group.qb_rows),
                "Infinium Row IDs": "; ".join(inf.at[idx, INF_ID] for idx in group.inf_rows),
                "PO Criterion": "Agree" if q_po and q_po == i_po else "Not used / differs",
                "Invoice Criterion": "Agree" if q_inv and q_inv == i_inv else "Not used / differs",
                "Signed Amount Criterion": "Agree" if q_total == i_total else "Differs",
                "QuickBooks Amount": cents_to_float(q_total),
                "Infinium Amount": cents_to_float(i_total),
                "Amount Difference": cents_to_float(q_total - i_total),
                "Group-Level Match": group.group_level,
                "Assessment Explanation": group.explanation,
            }
        )
    if not historical_clearances.empty:
        for _, clearance_group in historical_clearances.groupby(
            "Clearance ID", sort=False, dropna=False
        ):
            clearance = clearance_group.iloc[0]
            primary_is_qb = clearance["Primary Dataset"] == "QuickBooks Primary"
            primary_ids = [
                str(value) for value in clearance_group["Primary Row ID"] if value
            ]
            secondary_ids = [
                str(value) for value in clearance_group["Secondary Row ID"] if value
            ]
            primary_total = float(clearance_group["Primary Amount"].sum())
            secondary_total = float(clearance_group["Secondary Amount"].sum())
            records.append(
                {
                    "Match ID": clearance["Clearance ID"],
                    "Decision": "Matched - Historical Clearance",
                    "Match Method": clearance["Match Method"],
                    "Confidence": clearance["Confidence"],
                    "QuickBooks Row Count": (
                        len(primary_ids) if primary_is_qb else len(secondary_ids)
                    ),
                    "Infinium Row Count": (
                        len(secondary_ids) if primary_is_qb else len(primary_ids)
                    ),
                    "QuickBooks Row IDs": "; ".join(
                        primary_ids if primary_is_qb else secondary_ids
                    ),
                    "Infinium Row IDs": "; ".join(
                        secondary_ids if primary_is_qb else primary_ids
                    ),
                    "PO Criterion": "Applied under controlled historical pass",
                    "Invoice Criterion": "Applied under controlled historical pass",
                    "Signed Amount Criterion": "Agree",
                    "QuickBooks Amount": primary_total if primary_is_qb else secondary_total,
                    "Infinium Amount": secondary_total if primary_is_qb else primary_total,
                    "Amount Difference": float(clearance_group["Amount Difference"].sum()),
                    "Group-Level Match": bool(clearance["Group-Level Match"]),
                    "Assessment Explanation": clearance["Disposition"],
                }
            )
    for qidx in unmatched_qb:
        qid = qb.at[qidx, QB_ID]
        candidate = candidate_map.get(qid, {})
        records.append(
            {
                "Match ID": "",
                "Decision": "Unresolved",
                "Match Method": candidate.get("Disposition", "No match"),
                "Confidence": "Review",
                "QuickBooks Row Count": 1,
                "Infinium Row Count": candidate.get("Available Candidate Count", 0),
                "QuickBooks Row IDs": qid,
                "Infinium Row IDs": candidate.get("Available Infinium Candidate IDs", ""),
                "PO Criterion": "Candidate search performed" if qb.at[qidx, NORM_PO] else "Missing",
                "Invoice Criterion": "Candidate search performed" if qb.at[qidx, NORM_INV] else "Missing",
                "Signed Amount Criterion": "Valid" if valid_cents(qb.at[qidx, AMOUNT_CENTS]) else "Invalid / missing",
                "QuickBooks Amount": cents_to_float(qb.at[qidx, AMOUNT_CENTS]),
                "Infinium Amount": None,
                "Amount Difference": candidate.get("Minimum Amount Difference"),
                "Group-Level Match": False,
                "Assessment Explanation": "No unique automatic match was found.",
            }
        )
    if amount_variance_analysis is not None and not amount_variance_analysis.empty:
        for variance in amount_variance_analysis.to_dict("records"):
            records.append(
                {
                    "Match ID": variance["Variance ID"],
                    "Decision": "Reference-Matched Amount Variance Review Hold",
                    "Match Method": variance["Reference Evidence"],
                    "Confidence": variance["Confidence"],
                    "QuickBooks Row Count": 1,
                    "Infinium Row Count": 1,
                    "QuickBooks Row IDs": variance["QuickBooks Row ID"],
                    "Infinium Row IDs": variance["Infinium Row ID"],
                    "PO Criterion": (
                        "Agree" if variance["Normalized PO"] else "Not used / differs"
                    ),
                    "Invoice Criterion": (
                        "Agree" if variance["Normalized Invoice"] else "Not used / differs"
                    ),
                    "Signed Amount Criterion": "Differs - review required",
                    "QuickBooks Amount": variance["QuickBooks Amount"],
                    "Infinium Amount": variance["Infinium Amount"],
                    "Amount Difference": variance["Potential Difference"],
                    "Group-Level Match": False,
                    "Assessment Explanation": variance["Explanation"],
                }
            )
    return pd.DataFrame(records)


def build_method_summary(
    matches: list[MatchGroup],
    historical_clearances: pd.DataFrame,
    unmatched_qb: list[int],
    unmatched_inf: list[int],
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    duplicate_qb_rows: list[int],
    duplicate_inf_rows: list[int],
    duplicate_review_hold_qb_rows: list[int],
    duplicate_review_hold_inf_rows: list[int],
    amount_variance_review_hold_qb_rows: list[int],
    amount_variance_review_hold_inf_rows: list[int],
    fuzzy_match_review_hold_qb_rows: list[int],
    fuzzy_match_review_hold_inf_rows: list[int],
) -> pd.DataFrame:
    buckets: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"QB Rows": 0, "Infinium Rows": 0, "QB Cents": 0, "Infinium Cents": 0}
    )
    for group in matches:
        bucket = buckets[group.method]
        bucket["QB Rows"] += len(group.qb_rows)
        bucket["Infinium Rows"] += len(group.inf_rows)
        bucket["QB Cents"] += _amount_total(qb, group.qb_rows)
        bucket["Infinium Cents"] += _amount_total(inf, group.inf_rows)
    if not historical_clearances.empty:
        for clearance in historical_clearances.to_dict("records"):
            bucket = buckets[str(clearance["Match Method"])]
            primary_cents = int(Decimal(str(clearance["Primary Amount"])) * 100)
            if clearance["Primary Dataset"] == "QuickBooks Primary":
                if _optional_index(clearance["Primary Row Index"]) is not None:
                    bucket["QB Rows"] += 1
                    bucket["QB Cents"] += primary_cents
            else:
                if _optional_index(clearance["Primary Row Index"]) is not None:
                    bucket["Infinium Rows"] += 1
                    bucket["Infinium Cents"] += primary_cents
    if unmatched_qb:
        bucket = buckets["Unmatched QuickBooks"]
        bucket["QB Rows"] = len(unmatched_qb)
        bucket["QB Cents"] = _amount_total(qb, unmatched_qb)
    if unmatched_inf:
        bucket = buckets["Unmatched Infinium"]
        bucket["Infinium Rows"] = len(unmatched_inf)
        bucket["Infinium Cents"] = _amount_total(inf, unmatched_inf)
    if duplicate_qb_rows:
        bucket = buckets["Excess QuickBooks copies excluded from accrual"]
        bucket["QB Rows"] = len(duplicate_qb_rows)
        bucket["QB Cents"] = _amount_total(qb, duplicate_qb_rows)
    if duplicate_inf_rows:
        bucket = buckets["Excess Infinium copies excluded (separate worksheet)"]
        bucket["Infinium Rows"] = len(duplicate_inf_rows)
        bucket["Infinium Cents"] = _amount_total(inf, duplicate_inf_rows)
    if duplicate_review_hold_qb_rows:
        bucket = buckets["Duplicate Review Hold - QuickBooks (excluded from JE, pending disposition)"]
        bucket["QB Rows"] = len(duplicate_review_hold_qb_rows)
        bucket["QB Cents"] = _amount_total(qb, duplicate_review_hold_qb_rows)
    if duplicate_review_hold_inf_rows:
        bucket = buckets["Duplicate Review Hold - Infinium (pending disposition)"]
        bucket["Infinium Rows"] = len(duplicate_review_hold_inf_rows)
        bucket["Infinium Cents"] = _amount_total(inf, duplicate_review_hold_inf_rows)
    if amount_variance_review_hold_qb_rows or amount_variance_review_hold_inf_rows:
        bucket = buckets[
            "Reference-Matched Amount Variance Review Hold (excluded from automatic JE)"
        ]
        bucket["QB Rows"] = len(amount_variance_review_hold_qb_rows)
        bucket["Infinium Rows"] = len(amount_variance_review_hold_inf_rows)
        bucket["QB Cents"] = _amount_total(qb, amount_variance_review_hold_qb_rows)
        bucket["Infinium Cents"] = _amount_total(inf, amount_variance_review_hold_inf_rows)
    if fuzzy_match_review_hold_qb_rows or fuzzy_match_review_hold_inf_rows:
        bucket = buckets["Fuzzy Match Review Hold (excluded from JE, pending confirmation)"]
        bucket["QB Rows"] = len(fuzzy_match_review_hold_qb_rows)
        bucket["Infinium Rows"] = len(fuzzy_match_review_hold_inf_rows)
        bucket["QB Cents"] = _amount_total(qb, fuzzy_match_review_hold_qb_rows)
        bucket["Infinium Cents"] = _amount_total(inf, fuzzy_match_review_hold_inf_rows)
    records = []
    for method, values in buckets.items():
        records.append(
            {
                "Match Method": method,
                "QuickBooks Rows": values["QB Rows"],
                "Infinium Rows": values["Infinium Rows"],
                "QuickBooks Amount": cents_to_float(values["QB Cents"]),
                "Infinium Amount": cents_to_float(values["Infinium Cents"]),
                "Amount Difference": cents_to_float(values["QB Cents"] - values["Infinium Cents"]),
                "Share of QuickBooks Rows": values["QB Rows"] / len(qb) if len(qb) else 0,
            }
        )
    return pd.DataFrame(records).sort_values(
        ["QuickBooks Rows", "Infinium Rows", "Match Method"], ascending=[False, False, True]
    ).reset_index(drop=True)


def build_product_summary(
    qb: pd.DataFrame,
    mapping: dict[str, Optional[str]],
    current_fiscal_period: Optional[int] = None,
    fiscal_year: Optional[int] = None,
) -> pd.DataFrame:
    qty_col = mapping.get("quantity")
    amount_col = mapping.get("amount")
    if not qty_col or not amount_col:
        return pd.DataFrame(columns=["Product Name", "Product Quantity", "Product Value"])
    work = qb[qb[PRODUCT_STANDARD].notna()].copy()
    if current_fiscal_period is not None:
        expected_label = f"P{int(current_fiscal_period):02d}-{int(fiscal_year or 0)}"
        work = work.loc[work[FISCAL_LABEL].eq(expected_label)].copy()
    if work.empty:
        return pd.DataFrame(columns=["Product Name", "Product Quantity", "Product Value"])
    work["__QTY"] = pd.to_numeric(work[qty_col], errors="coerce").fillna(0)
    work["__AMOUNT"] = work[AMOUNT_CENTS].map(cents_to_float)
    return (
        work.groupby(PRODUCT_STANDARD, as_index=False)
        .agg(**{"Product Quantity": ("__QTY", "sum"), "Product Value": ("__AMOUNT", "sum")})
        .rename(columns={PRODUCT_STANDARD: "Product Name"})
        .sort_values("Product Name")
        .reset_index(drop=True)
    )


def _period_sort(label: str) -> tuple[int, int, str]:
    match = re.fullmatch(r"P(\d{2})-(\d{4})", str(label))
    if not match:
        return 9999, 99, str(label)
    return int(match.group(2)), int(match.group(1)), str(label)


def build_exception_analysis(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    unmatched_qb: list[int],
    unmatched_inf: list[int],
    candidates: pd.DataFrame,
    amount_variance_analysis: Optional[pd.DataFrame] = None,
    paired_rows: Optional[list[dict[str, Any]]] = None,
    fuzzy_match_review_hold_analysis: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    candidate_reason = candidates.set_index("QuickBooks Row ID")["Exception Cause"].to_dict() if not candidates.empty else {}
    paired_cause = {
        int(row["QB Index"]): str(row["Exception Cause"])
        for row in (paired_rows or [])
        if row.get("Section") == "02 Unmatched QuickBooks"
        and row.get("QB Index") is not None
    }
    period_values = qb.loc[unmatched_qb, FISCAL_LABEL] if unmatched_qb else pd.Series(dtype=object)
    has_period = not period_values.empty and period_values.ne("Unspecified").any()
    period_groups: dict[str, list[int]] = defaultdict(list)
    for idx in unmatched_qb:
        label = qb.at[idx, FISCAL_LABEL] if has_period else "All QuickBooks transactions"
        period_groups[label].append(idx)
    for label in sorted(period_groups, key=_period_sort):
        rows = period_groups[label]
        records.append(
            {
                "Analysis Type": "QuickBooks exceptions by source period",
                "Dimension": label,
                "Transaction Count": len(rows),
                "Amount": cents_to_float(_amount_total(qb, rows)),
            }
        )
    reason_groups: dict[str, list[int]] = defaultdict(list)
    for idx in unmatched_qb:
        reason_groups[
            paired_cause.get(
                idx,
                candidate_reason.get(qb.at[idx, QB_ID], "No match"),
            )
        ].append(idx)
    for reason, rows in sorted(reason_groups.items()):
        records.append(
            {
                "Analysis Type": "QuickBooks exceptions by reason",
                "Dimension": reason,
                "Transaction Count": len(rows),
                "Amount": cents_to_float(_amount_total(qb, rows)),
            }
        )
    if unmatched_inf:
        records.append(
            {
                "Analysis Type": "Infinium exceptions",
                "Dimension": "Unmatched Infinium",
                "Transaction Count": len(unmatched_inf),
                "Amount": cents_to_float(_amount_total(inf, unmatched_inf)),
            }
        )
    if amount_variance_analysis is not None and not amount_variance_analysis.empty:
        for classification, group in amount_variance_analysis.groupby(
            "Classification", sort=False
        ):
            records.append(
                {
                    "Analysis Type": "Reference-matched amount variances on review hold",
                    "Dimension": classification,
                    "Transaction Count": len(group),
                    "Amount": float(group["QuickBooks Amount"].sum()),
                }
            )
    if fuzzy_match_review_hold_analysis is not None and not fuzzy_match_review_hold_analysis.empty:
        for classification, group in fuzzy_match_review_hold_analysis.groupby(
            "Classification", sort=False
        ):
            records.append(
                {
                    "Analysis Type": "Fuzzy matches on review hold",
                    "Dimension": classification,
                    "Transaction Count": len(group),
                    "Amount": float(group["QuickBooks Amount"].sum()),
                }
            )
    return pd.DataFrame(records)


def build_rules_and_config(
    qb_mapping: dict[str, Optional[str]],
    inf_mapping: dict[str, Optional[str]],
    metadata: dict[str, Any],
    qb_secondary_mapping: Optional[dict[str, Optional[str]]] = None,
    inf_secondary_mapping: Optional[dict[str, Optional[str]]] = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    rules = pd.DataFrame(
        [
            {"Priority": 1, "Rule": "PO + Invoice + Amount", "Automatic": "Yes",
             "Requirement": "One unique row per dataset; normalized references and exact signed cents agree."},
            {"Priority": 2, "Rule": "PO + Amount", "Automatic": "Yes",
             "Requirement": "One unique remaining row per dataset; normalized PO and exact signed cents agree."},
            {"Priority": 3, "Rule": "Invoice + Amount", "Automatic": "Yes",
             "Requirement": "One unique remaining row per dataset; normalized invoice and exact signed cents agree."},
            {"Priority": 4, "Rule": "PO + Invoice + Aggregate Amount", "Automatic": "Yes, after one-to-one",
             "Requirement": "One unique remaining one-to-many or many-to-one relationship; every grouped row shares the normalized PO and invoice and exact signed-cent totals agree."},
            {"Priority": 5, "Rule": "PO + Aggregate Amount", "Automatic": "Yes, after one-to-one",
             "Requirement": "One unique remaining one-to-many or many-to-one relationship; every grouped row shares the normalized PO and exact signed-cent totals agree."},
            {"Priority": 6, "Rule": "Invoice + Aggregate Amount", "Automatic": "Yes, after one-to-one",
             "Requirement": "One unique remaining one-to-many or many-to-one relationship; every grouped row shares the normalized invoice and exact signed-cent totals agree."},
            {"Priority": 7, "Rule": "Fuzzy PO + Amount (word match, unique) (see fuzzy_po_matching.py)",
             "Automatic": "Yes, after every exact pass",
             "Requirement": (
                 "Applies only to rows still unresolved after every exact one-to-one and grouped-aggregate "
                 "pass. The signed amount must still match exactly; the PO comparison is whole-word "
                 "containment -- every significant word (3+ letters) on the shorter side's PO text must "
                 "appear on the longer side's -- not a general similarity score. Numbers and short "
                 "fragments are ignored so a shared product code can never be the sole basis for a match. "
                 "Accepted only when the match is unique on both sides; ambiguous candidates are left "
                 "unresolved. Tagged with its own 'Fuzzy' confidence, distinct from every exact-match method."
             )},
            {"Priority": 8, "Rule": "Ambiguous or many-to-many groups", "Automatic": "No",
             "Requirement": "Overlapping combinations, many-to-many relationships, groups over the safety limits, and nonunique solutions remain unresolved."},
            {"Priority": 9, "Rule": "Amount variance", "Automatic": "No",
             "Requirement": "Any nonzero cent difference is flagged as an exception; no tolerance is applied."},
            {"Priority": 10, "Rule": "Fuzzy product classification", "Automatic": "No financial effect",
             "Requirement": "Used only for Product Aggregate Summary; never determines transaction matching."},
            {"Priority": 11, "Rule": "Secondary historical clearance", "Automatic": "Yes, second pass",
             "Requirement": "After primary matching, historical rows may clear unresolved rows from the opposing primary dataset using the same one-to-one-then-controlled-grouped sequence, including the fuzzy PO pass (Priority 7)."},
            {"Priority": 12, "Rule": "Unused secondary rows", "Automatic": "Excluded",
             "Requirement": "Unmatched historical rows remain background data and never become exceptions or reconciliation items."},
            {"Priority": 13, "Rule": "Strong duplicate canonicalization (see duplicates.py)", "Automatic": "Yes, before matching",
             "Requirement": (
                 "For same-file rows with populated PO and invoice and identical signed cents, "
                 "the earliest source-position row is retained as the canonical transaction and "
                 "only later payload-identical copies are excluded. The candidate business-key group "
                 "and payload-confirmed copy set receive separate stable IDs. Automatic exclusion "
                 "assumes the source report grain does not permit legitimate repeated identical lines; "
                 "that report-grain assertion must be documented by the process owner."
             )},
            {"Priority": 14, "Rule": "Weak duplicate candidates: matched normally, held if unresolved",
             "Automatic": "Conditional",
             "Requirement": (
                 "Rows with identical signed cents but only a PO or only an invoice are never "
                 "auto-excluded up front -- they remain fully active and eligible to match through "
                 "every pass above. A candidate that matches proceeds normally with no exclusion. A "
                 "candidate that is still unresolved once matching and historical clearance complete is "
                 "removed from provisional JE support and placed in Duplicate Review Hold. The run is "
                 "non-postable until a documented human disposition is completed. Rows sharing a reference with "
                 "different amounts are never assumed duplicates and remain eligible for controlled "
                 "aggregate matching (Priorities 4-6)."
             )},
            {"Priority": 15, "Rule": "Historical overlap exclusion", "Automatic": "Yes, before clearance",
             "Requirement": (
                 "Historical files are canonicalized within-file first and only then compared with their "
                 "own primary dataset. Confirmed overlap and every unresolved historical duplicate candidate "
                 "are excluded from clearance while the primary row remains untouched."
             )},
            {"Priority": 16, "Rule": "Posting hard stops", "Automatic": "Yes",
             "Requirement": (
                 "Invalid financial amounts, unresolved primary duplicate candidates, and unresolved "
                 "historical duplicate candidates prevent READY TO POST status. Historical review rows "
                 "cannot participate in a clearance."
             )},
            {"Priority": 17, "Rule": "Reference-matched amount variance review", "Automatic": "Classification only",
             "Requirement": (
                 "After primary and historical matching, mutually unique unresolved rows sharing an exact "
                 "PO and/or invoice but different signed-cent amounts are moved to a separate review hold. "
                 "The full QuickBooks amount, Infinium amount, and potential difference are displayed, but "
                 "the engine never infers which system is correct and never automatically posts either value."
             )},
            {"Priority": 18, "Rule": "Standardized exception-cause classification", "Automatic": "Reporting only",
             "Requirement": (
                 "Every unresolved, duplicate, and amount-variance row receives a controlled cause, confidence, "
                 "financial treatment, and related-row evidence. Reconciled Data retains the actual duplicate PO, "
                 "invoice, signed amount, group ID, copy-set ID, and canonical row ID for audit inspection."
             )},
        ]
    )
    config_records = [
        {"Setting": "Application Version", "Value": APP_VERSION},
        {"Setting": "Matching Rule Version", "Value": MATCHING_RULE_VERSION},
        {"Setting": "Duplicate Rule Version", "Value": DUPLICATE_RULE_VERSION},
        {
            "Setting": "Duplicate source-report grain validated",
            "Value": bool(metadata.get("duplicate_source_grain_validated", False)),
        },
        {"Setting": "Run ID", "Value": metadata["run_id"]},
        {"Setting": "Run Timestamp (Central Time)", "Value": metadata["run_timestamp"]},
        {"Setting": "QuickBooks Filename", "Value": metadata["qb_filename"]},
        {"Setting": "QuickBooks SHA-256", "Value": metadata["qb_sha256"]},
        {"Setting": "Infinium Filename", "Value": metadata["inf_filename"]},
        {"Setting": "Infinium SHA-256", "Value": metadata["inf_sha256"]},
        {"Setting": "Automatic Amount Tolerance", "Value": "$0.00; automatic matches require exact signed cents"},
        {"Setting": "Match Cardinality", "Value": "One-to-one first; then unique one-to-many or many-to-one exact aggregates"},
        {"Setting": "Maximum automatic grouped rows", "Value": MAX_GROUP_SIZE},
        {"Setting": "Maximum rows evaluated per shared reference", "Value": MAX_GROUP_POOL_ROWS},
        {"Setting": "Selected Fiscal Year", "Value": metadata["fiscal_year"]},
        {"Setting": "Selected Fiscal Period", "Value": metadata["fiscal_period"]},
        {"Setting": "QuickBooks Secondary Filename", "Value": metadata.get("qb_secondary_filename") or "Not provided"},
        {"Setting": "QuickBooks Secondary SHA-256", "Value": metadata.get("qb_secondary_sha256") or "Not provided"},
        {"Setting": "Infinium Secondary Filename", "Value": metadata.get("inf_secondary_filename") or "Not provided"},
        {"Setting": "Infinium Secondary SHA-256", "Value": metadata.get("inf_secondary_sha256") or "Not provided"},
        {"Setting": "QuickBooks subtotal rows excluded", "Value": metadata.get("qb_subtotal_rows_excluded", 0)},
        {"Setting": "Source validation warnings", "Value": metadata.get("source_validation_warnings", 0)},
        {
            "Setting": "Fiscal Period Treatment",
            "Value": (
                "Selected period filters the Product Aggregate Summary and classifies exception urgency. "
                "Fiscal period never changes strict transaction matching."
            ),
        },
        {
            "Setting": "Duplicate Treatment",
            "Value": (
                "Strong same-file groups retain one deterministic canonical row per payload-confirmed copy "
                "set and exclude only excess copies. PO-only and invoice-only candidates remain active for "
                "primary matching but block posting if unresolved. Historical files are screened within-file "
                "before cross-scope overlap; every unresolved historical review row is withheld from clearance."
            ),
        },
    ]
    for source_name, audit_key in (
        ("QuickBooks", "qb_import_audit"),
        ("Infinium", "inf_import_audit"),
        ("QuickBooks Secondary", "qb_secondary_import_audit"),
        ("Infinium Secondary", "inf_secondary_import_audit"),
    ):
        audit = metadata.get(audit_key, {}) or {}
        config_records.extend(
            [
                {
                    "Setting": f"{source_name} blank columns removed",
                    "Value": len(audit.get("blank_columns_removed", [])),
                },
                {
                    "Setting": f"{source_name} repeated header rows removed",
                    "Value": int(audit.get("repeated_header_rows_removed", 0) or 0),
                },
                {
                    "Setting": f"{source_name} duplicate headers renamed",
                    "Value": len(audit.get("duplicate_headers_renamed", [])),
                },
            ]
        )
    for source, mapping in (("QuickBooks", qb_mapping), ("Infinium", inf_mapping)):
        for field_name, column_name in mapping.items():
            config_records.append(
                {"Setting": f"{source} mapping - {field_name}", "Value": column_name or "Not mapped"}
            )
    for source, mapping in (
        ("QuickBooks Secondary", qb_secondary_mapping),
        ("Infinium Secondary", inf_secondary_mapping),
    ):
        for field_name, column_name in (mapping or {}).items():
            config_records.append(
                {"Setting": f"{source} mapping - {field_name}", "Value": column_name or "Not mapped"}
            )
    return rules, pd.DataFrame(config_records)


def build_fiscal_exception_summary(result: ReconciliationResult) -> pd.DataFrame:
    """Summarize unresolved primary QuickBooks exceptions by source period."""
    columns = [
        "Fiscal Period", "Period Classification", "Exception Count",
        "Exception Quantity", "Net Exception Amount",
    ]
    if not result.unmatched_qb:
        return pd.DataFrame(columns=columns)

    period_column = result.qb_mapping.get("period")
    quantity_column = result.qb_mapping.get("quantity")
    selected_period = result.metadata.get("fiscal_period")
    default_year = int(result.metadata.get("fiscal_year", result.run_timestamp.year))
    work = result.qb_work.loc[result.unmatched_qb].copy()
    work["__EXCEPTION_QUANTITY"] = (
        pd.to_numeric(work[quantity_column], errors="coerce").fillna(0)
        if quantity_column and quantity_column in work.columns
        else 0.0
    )
    work["__EXCEPTION_AMOUNT"] = work[AMOUNT_CENTS].map(cents_to_float)
    if not period_column or period_column not in work.columns:
        return pd.DataFrame(
            [{
                "Fiscal Period": "Not available",
                "Period Classification": "Fiscal Period Not Available",
                "Exception Count": len(work),
                "Exception Quantity": float(work["__EXCEPTION_QUANTITY"].sum()),
                "Net Exception Amount": float(work["__EXCEPTION_AMOUNT"].sum()),
            }],
            columns=columns,
        )

    work["__PERIOD_NUMBER"] = work[period_column].map(
        lambda value: parse_fiscal_period(value, default_year)[0]
    )
    work["Fiscal Period"] = work["__PERIOD_NUMBER"].map(
        lambda value: f"PD-{int(value):02d}" if pd.notna(value) else "Unspecified"
    )

    def classify(value: Any) -> str:
        if pd.isna(value):
            return "Unspecified Period - Review"
        if selected_period is None:
            return "Reporting Period Not Selected"
        if int(value) == int(selected_period):
            return "Current Reporting Period"
        return "Prior-Period Urgent Exception"

    work["Period Classification"] = work["__PERIOD_NUMBER"].map(classify)
    work["__URGENCY_SORT"] = work["Period Classification"].map(
        {
            "Prior-Period Urgent Exception": 1,
            "Unspecified Period - Review": 2,
            "Reporting Period Not Selected": 3,
            "Current Reporting Period": 4,
        }
    )
    summary = (
        work.groupby(
            ["Fiscal Period", "Period Classification", "__URGENCY_SORT", "__PERIOD_NUMBER"],
            as_index=False,
            dropna=False,
        )
        .agg(**{
            "Exception Count": (QB_ID, "size"),
            "Exception Quantity": ("__EXCEPTION_QUANTITY", "sum"),
            "Net Exception Amount": ("__EXCEPTION_AMOUNT", "sum"),
        })
        .sort_values(["__URGENCY_SORT", "__PERIOD_NUMBER"], na_position="last")
    )
    return summary.reindex(columns=columns).reset_index(drop=True)


def build_controls(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    matches: list[MatchGroup],
    historical_clearances: pd.DataFrame,
    unmatched_qb: list[int],
    unmatched_inf: list[int],
    duplicate_qb_rows: list[int],
    duplicate_inf_rows: list[int],
    duplicate_review_hold_qb_rows: list[int],
    duplicate_review_hold_inf_rows: list[int],
    amount_variance_review_hold_qb_rows: list[int],
    amount_variance_review_hold_inf_rows: list[int],
    fuzzy_match_review_hold_qb_rows: list[int],
    fuzzy_match_review_hold_inf_rows: list[int],
) -> pd.DataFrame:
    matched_q = _matched_row_indexes(matches, "QB")
    matched_i = _matched_row_indexes(matches, "INF")
    historical_q = _historical_row_indexes(
        historical_clearances,
        "QuickBooks Primary",
        "Primary Row Index",
    )
    historical_i = _historical_row_indexes(
        historical_clearances,
        "Infinium Primary",
        "Primary Row Index",
    )
    qb_total = _amount_total(qb, qb.index)
    inf_total = _amount_total(inf, inf.index)
    matched_q_total = _amount_total(qb, matched_q)
    matched_i_total = _amount_total(inf, matched_i)
    historical_q_total = _amount_total(qb, historical_q)
    historical_i_total = _amount_total(inf, historical_i)
    unresolved_q_total = _amount_total(qb, unmatched_qb)
    unresolved_i_total = _amount_total(inf, unmatched_inf)
    duplicate_q_total = _amount_total(qb, duplicate_qb_rows)
    duplicate_i_total = _amount_total(inf, duplicate_inf_rows)
    review_hold_q_total = _amount_total(qb, duplicate_review_hold_qb_rows)
    review_hold_i_total = _amount_total(inf, duplicate_review_hold_inf_rows)
    variance_hold_q_total = _amount_total(qb, amount_variance_review_hold_qb_rows)
    variance_hold_i_total = _amount_total(inf, amount_variance_review_hold_inf_rows)
    fuzzy_hold_q_total = _amount_total(qb, fuzzy_match_review_hold_qb_rows)
    fuzzy_hold_i_total = _amount_total(inf, fuzzy_match_review_hold_inf_rows)
    historical_difference = (
        round(float(historical_clearances["Amount Difference"].sum()), 2)
        if not historical_clearances.empty else 0.0
    )
    records = [
        ("QuickBooks row completeness", len(qb),
         len(matched_q) + len(historical_q) + len(unmatched_qb) + len(duplicate_qb_rows)
         + len(duplicate_review_hold_qb_rows) + len(amount_variance_review_hold_qb_rows)
         + len(fuzzy_match_review_hold_qb_rows)),
        ("Infinium row completeness", len(inf),
         len(matched_i) + len(historical_i) + len(unmatched_inf) + len(duplicate_inf_rows)
         + len(duplicate_review_hold_inf_rows) + len(amount_variance_review_hold_inf_rows)
         + len(fuzzy_match_review_hold_inf_rows)),
        ("QuickBooks amount roll-forward", cents_to_float(qb_total),
         cents_to_float(
             matched_q_total + historical_q_total + unresolved_q_total
             + duplicate_q_total + review_hold_q_total + variance_hold_q_total
             + fuzzy_hold_q_total
         )),
        ("Infinium amount roll-forward", cents_to_float(inf_total),
         cents_to_float(
             matched_i_total + historical_i_total + unresolved_i_total
             + duplicate_i_total + review_hold_i_total + variance_hold_i_total
             + fuzzy_hold_i_total
         )),
        ("Primary-to-primary matched totals", cents_to_float(matched_q_total), cents_to_float(matched_i_total)),
        ("Historical clearance amount difference", 0.0, historical_difference),
        ("Unresolved JE support", cents_to_float(unresolved_q_total),
         cents_to_float(
             qb_total - matched_q_total - historical_q_total
             - duplicate_q_total - review_hold_q_total - variance_hold_q_total
             - fuzzy_hold_q_total
         )),
        ("Excess QuickBooks copies excluded from JE", cents_to_float(duplicate_q_total),
         cents_to_float(duplicate_q_total)),
        ("Duplicate Review Hold QuickBooks items excluded from JE",
         cents_to_float(review_hold_q_total), cents_to_float(review_hold_q_total)),
        ("Reference-matched amount variance QuickBooks items excluded from automatic JE",
         cents_to_float(variance_hold_q_total), cents_to_float(variance_hold_q_total)),
        ("Fuzzy match review hold QuickBooks items excluded from automatic JE",
         cents_to_float(fuzzy_hold_q_total), cents_to_float(fuzzy_hold_q_total)),
    ]
    output = []
    for check, expected, actual in records:
        difference = float(expected) - float(actual)
        output.append(
            {
                "Check": check,
                "Expected": expected,
                "Actual": actual,
                "Difference": difference,
                "Status": "PASS" if difference == 0 else "FAIL",
            }
        )
    return pd.DataFrame(output)


def build_reconciliation(
    qb_raw: pd.DataFrame,
    inf_raw: pd.DataFrame,
    qb_mapping: dict[str, Optional[str]],
    inf_mapping: dict[str, Optional[str]],
    metadata: dict[str, Any],
    fiscal_year: int,
    qb_secondary_raw: Optional[pd.DataFrame] = None,
    inf_secondary_raw: Optional[pd.DataFrame] = None,
    qb_secondary_mapping: Optional[dict[str, Optional[str]]] = None,
    inf_secondary_mapping: Optional[dict[str, Optional[str]]] = None,
) -> ReconciliationResult:
    qb = prepare_working_frame(qb_raw, qb_mapping, "QB", fiscal_year)
    inf = prepare_working_frame(inf_raw, inf_mapping, "INF", fiscal_year)
    qb_secondary = (
        prepare_working_frame(
            qb_secondary_raw, qb_secondary_mapping, "QB", fiscal_year,
            id_prefix="QB-HIST",
        )
        if qb_secondary_raw is not None and qb_secondary_mapping is not None
        else None
    )
    inf_secondary = (
        prepare_working_frame(
            inf_secondary_raw, inf_secondary_mapping, "INF", fiscal_year,
            id_prefix="INF-HIST",
        )
        if inf_secondary_raw is not None and inf_secondary_mapping is not None
        else None
    )

    # Canonicalize strong same-file groups before matching. Only excess copies
    # are removed; the earliest source-position row remains active. Weak groups
    # are reported for review without automatic exclusion. Historical frames
    # are additionally compared with their own primary population so an
    # overlapping historical row cannot clear an opposing primary exception.
    qb_screen = screen_duplicates(
        qb, QB_ID, "QuickBooks", "Primary",
        "Strong groups retain one canonical row; only excess copies are excluded from the JE.",
        auto_exclude_strict=True,
    )
    inf_screen = screen_duplicates(
        inf, INF_ID, "Infinium", "Primary",
        "Strong groups retain one canonical row; only excess copies are excluded from matching.",
        auto_exclude_strict=True,
    )
    qb_secondary_screen = (
        screen_duplicates(
            qb_secondary, QB_ID, "QuickBooks", "Historical (Secondary)",
            "Excess copies and rows overlapping QuickBooks primary are excluded from clearance.",
            auto_exclude_strict=True,
            reference_frame=qb_screen.active_frame,
            reference_id_column=QB_ID,
        )
        if qb_secondary is not None else None
    )
    inf_secondary_screen = (
        screen_duplicates(
            inf_secondary, INF_ID, "Infinium", "Historical (Secondary)",
            "Excess copies and rows overlapping Infinium primary are excluded from clearance.",
            auto_exclude_strict=True,
            reference_frame=inf_screen.active_frame,
            reference_id_column=INF_ID,
        )
        if inf_secondary is not None else None
    )

    qb_active = qb_screen.active_frame
    inf_active = inf_screen.active_frame
    # A historical review candidate is never allowed to clear a primary
    # exception. Unlike primary candidates, historical rows are optional
    # evidence rather than accrual-source rows, so the safe treatment is to
    # withhold every unresolved candidate from clearance until a documented
    # disposition is supplied and the reconciliation is rerun.
    qb_secondary_active = (
        qb_secondary_screen.active_frame.drop(
            index=qb_secondary_screen.suspected_rows, errors="ignore"
        )
        if qb_secondary_screen else qb_secondary
    )
    inf_secondary_active = (
        inf_secondary_screen.active_frame.drop(
            index=inf_secondary_screen.suspected_rows, errors="ignore"
        )
        if inf_secondary_screen else inf_secondary
    )

    matches, initially_unmatched_qb, initially_unmatched_inf, candidates = perform_matching(
        qb_active, inf_active
    )

    # A fuzzy PO/text match is a similarity guess, not a certain relationship,
    # so it is never posted the way an exact match is. Pull every fuzzy group
    # out of the accepted matches and into its own review-hold population --
    # see build_fuzzy_match_review_holds for the reviewer-facing report.
    fuzzy_hold_groups = [group for group in matches if group.confidence == FUZZY_PO_CONFIDENCE]
    matches = [group for group in matches if group.confidence != FUZZY_PO_CONFIDENCE]
    fuzzy_hold_groups.sort(
        key=lambda group: min(qb_active.at[idx, SOURCE_POS] for idx in group.qb_rows)
    )
    for number, group in enumerate(fuzzy_hold_groups, 1):
        group.match_id = f"FUZZY-{number:06d}"
    fuzzy_match_review_hold_qb = sorted({idx for group in fuzzy_hold_groups for idx in group.qb_rows})
    fuzzy_match_review_hold_inf = sorted({idx for group in fuzzy_hold_groups for idx in group.inf_rows})
    fuzzy_match_review_hold_analysis = build_fuzzy_match_review_holds(
        qb_active, inf_active, fuzzy_hold_groups
    )

    historical_clearances, unmatched_qb, unmatched_inf = build_historical_clearances(
        qb,
        inf,
        initially_unmatched_qb,
        initially_unmatched_inf,
        qb_secondary_active,
        inf_secondary_active,
    )

    # Weak-basis (PO-only or invoice-only) duplicate candidates were never
    # excluded from matching -- they stayed active and could resolve
    # normally. Now that matching and historical clearance are both final,
    # split them by outcome: one that matched needs no further action, but
    # one that is still unresolved must never quietly inflate the proposed
    # JE just because nobody happened to open the duplicates worksheet. It
    # is pulled into its own Duplicate Review Hold population instead.
    duplicate_review_hold_qb = sorted(set(unmatched_qb) & set(qb_screen.suspected_rows))
    duplicate_review_hold_inf = sorted(set(unmatched_inf) & set(inf_screen.suspected_rows))
    unmatched_qb = sorted(set(unmatched_qb) - set(duplicate_review_hold_qb))
    unmatched_inf = sorted(set(unmatched_inf) - set(duplicate_review_hold_inf))
    resolved_suspected_qb_ids = {
        qb.at[idx, QB_ID] for idx in qb_screen.suspected_rows if idx not in duplicate_review_hold_qb
    }
    held_suspected_qb_ids = {qb.at[idx, QB_ID] for idx in duplicate_review_hold_qb}
    resolved_suspected_inf_ids = {
        inf.at[idx, INF_ID] for idx in inf_screen.suspected_rows if idx not in duplicate_review_hold_inf
    }
    held_suspected_inf_ids = {inf.at[idx, INF_ID] for idx in duplicate_review_hold_inf}

    # Final unresolved rows are now examined for mutually unique reference
    # relationships with differing amounts. These rows are not ordinary
    # missing-record accruals: both sides are withheld from automatic posting
    # and itemized in their own auditor-facing review population.
    (
        amount_variance_analysis,
        amount_variance_review_hold_qb,
        amount_variance_review_hold_inf,
    ) = build_reference_amount_variances(qb, inf, unmatched_qb, unmatched_inf)
    unmatched_qb = sorted(
        set(unmatched_qb).difference(amount_variance_review_hold_qb)
    )
    unmatched_inf = sorted(
        set(unmatched_inf).difference(amount_variance_review_hold_inf)
    )

    paired_rows = build_paired_rows(
        matches, historical_clearances, unmatched_qb, unmatched_inf, qb, inf, candidates,
        qb_screen.duplicate_rows, inf_screen.duplicate_rows,
        duplicate_review_hold_qb, duplicate_review_hold_inf,
        amount_variance_analysis,
        fuzzy_hold_groups,
        qb_screen.report,
        inf_screen.report,
    )
    normalization = build_normalization_detail(qb, inf, qb_mapping, inf_mapping)
    assessments = build_match_assessments(
        matches, historical_clearances, unmatched_qb, qb, inf, candidates,
        amount_variance_analysis,
    )
    method_summary = build_method_summary(
        matches, historical_clearances, unmatched_qb, unmatched_inf, qb, inf,
        qb_screen.duplicate_rows, inf_screen.duplicate_rows,
        duplicate_review_hold_qb, duplicate_review_hold_inf,
        amount_variance_review_hold_qb, amount_variance_review_hold_inf,
        fuzzy_match_review_hold_qb, fuzzy_match_review_hold_inf,
    )
    exception_analysis = build_exception_analysis(
        qb, inf, unmatched_qb, unmatched_inf, candidates,
        amount_variance_analysis, paired_rows,
        fuzzy_match_review_hold_analysis,
    )
    qb_primary_duplicate_report = finalize_review_dispositions(
        qb_screen.report,
        resolved_ids=resolved_suspected_qb_ids,
        held_ids=held_suspected_qb_ids,
    )
    qb_secondary_duplicate_report = (
        finalize_review_dispositions(
            qb_secondary_screen.report,
            historical_hold_ids={
                qb_secondary.at[idx, QB_ID]
                for idx in qb_secondary_screen.suspected_rows
            },
        )
        if qb_secondary_screen else None
    )
    duplicate_analysis = combine_duplicate_reports(
        qb_primary_duplicate_report,
        qb_secondary_duplicate_report,
    )
    inf_primary_duplicate_report = finalize_review_dispositions(
        inf_screen.report,
        resolved_ids=resolved_suspected_inf_ids,
        held_ids=held_suspected_inf_ids,
    )
    inf_secondary_duplicate_report = (
        finalize_review_dispositions(
            inf_secondary_screen.report,
            historical_hold_ids={
                inf_secondary.at[idx, INF_ID]
                for idx in inf_secondary_screen.suspected_rows
            },
        )
        if inf_secondary_screen else None
    )
    infinium_duplicate_analysis = combine_duplicate_reports(
        inf_primary_duplicate_report,
        inf_secondary_duplicate_report,
    )
    product_summary = build_product_summary(
        qb,
        qb_mapping,
        metadata.get("fiscal_period"),
        fiscal_year,
    )
    controls = build_controls(
        qb, inf, matches, historical_clearances, unmatched_qb, unmatched_inf,
        qb_screen.duplicate_rows, inf_screen.duplicate_rows,
        duplicate_review_hold_qb, duplicate_review_hold_inf,
        amount_variance_review_hold_qb, amount_variance_review_hold_inf,
        fuzzy_match_review_hold_qb, fuzzy_match_review_hold_inf,
    )
    rules, config = build_rules_and_config(
        qb_mapping,
        inf_mapping,
        metadata,
        qb_secondary_mapping,
        inf_secondary_mapping,
    )

    matched_q = _matched_row_indexes(matches, "QB")
    matched_i = _matched_row_indexes(matches, "INF")
    historical_matched_q = _historical_row_indexes(
        historical_clearances,
        "QuickBooks Primary",
        "Primary Row Index",
    )
    historical_matched_i = _historical_row_indexes(
        historical_clearances,
        "Infinium Primary",
        "Primary Row Index",
    )
    qb_secondary_used = _historical_row_indexes(
        historical_clearances,
        "Infinium Primary",
        "Secondary Row Index",
    )
    inf_secondary_used = _historical_row_indexes(
        historical_clearances,
        "QuickBooks Primary",
        "Secondary Row Index",
    )
    qb_total_cents = _amount_total(qb, qb.index)
    inf_total_cents = _amount_total(inf, inf.index)
    unmatched_q_cents = _amount_total(qb, unmatched_qb)
    matched_q_cents = _amount_total(qb, matched_q)
    matched_i_cents = _amount_total(inf, matched_i)
    cleared_q_cents = _amount_total(qb, historical_matched_q)
    cleared_i_cents = _amount_total(inf, historical_matched_i)
    qb_gross = _gross_amount_total(qb, qb.index)
    matched_q_gross = _gross_amount_total(qb, matched_q)
    posting_blockers: list[str] = []
    if duplicate_review_hold_qb:
        posting_blockers.append("unresolved QuickBooks duplicate review holds")
    if duplicate_review_hold_inf:
        posting_blockers.append("unresolved Infinium duplicate review holds")
    if qb_secondary_screen and qb_secondary_screen.suspected_rows:
        posting_blockers.append("unresolved QuickBooks historical duplicate review holds")
    if inf_secondary_screen and inf_secondary_screen.suspected_rows:
        posting_blockers.append("unresolved Infinium historical duplicate review holds")
    if not amount_variance_analysis.empty:
        posting_blockers.append("unresolved reference-matched amount variance review holds")
    if not fuzzy_match_review_hold_analysis.empty:
        posting_blockers.append("unresolved fuzzy PO match review holds")
    if qb[AMOUNT_CENTS].isna().any():
        posting_blockers.append("invalid QuickBooks amounts")
    if inf[AMOUNT_CENTS].isna().any():
        posting_blockers.append("invalid Infinium amounts")
    automatic_duplicate_exclusions = bool(
        qb_screen.duplicate_rows
        or inf_screen.duplicate_rows
        or (qb_secondary_screen and qb_secondary_screen.duplicate_rows)
        or (inf_secondary_screen and inf_secondary_screen.duplicate_rows)
    )
    if automatic_duplicate_exclusions and not bool(
        metadata.get("duplicate_source_grain_validated", False)
    ):
        posting_blockers.append(
            "source-report grain has not been documented as safe for automatic duplicate exclusion"
        )
    metrics = {
        "QuickBooks Rows": len(qb),
        "Infinium Rows": len(inf),
        "QuickBooks Source Total": cents_to_float(qb_total_cents),
        "Infinium Source Total": cents_to_float(inf_total_cents),
        "Source Difference": cents_to_float(qb_total_cents - inf_total_cents),
        "Matched QuickBooks Rows": len(matched_q) + len(historical_matched_q),
        "Matched Infinium Rows": len(matched_i) + len(historical_matched_i),
        "Matched QuickBooks Amount": cents_to_float(matched_q_cents + cleared_q_cents),
        "Matched Infinium Amount": cents_to_float(matched_i_cents + cleared_i_cents),
        "Matched Amount Difference": cents_to_float(matched_q_cents - matched_i_cents),
        "Historical QuickBooks Rows Cleared": len(historical_matched_q),
        "Historical Infinium Rows Cleared": len(historical_matched_i),
        "Historical Clearances": (
            int(historical_clearances["Clearance ID"].nunique())
            if not historical_clearances.empty else 0
        ),
        "QuickBooks Secondary Rows": len(qb_secondary) if qb_secondary is not None else 0,
        "QuickBooks Secondary Rows Used": len(qb_secondary_used),
        "QuickBooks Secondary Rows Ignored": (
            len(qb_secondary) - len(qb_secondary_used) if qb_secondary is not None else 0
        ),
        "Infinium Secondary Rows": len(inf_secondary) if inf_secondary is not None else 0,
        "Infinium Secondary Rows Used": len(inf_secondary_used),
        "Infinium Secondary Rows Ignored": (
            len(inf_secondary) - len(inf_secondary_used) if inf_secondary is not None else 0
        ),
        "Unresolved QuickBooks Rows": len(unmatched_qb),
        "Unresolved QuickBooks Amount": cents_to_float(unmatched_q_cents),
        "Unmatched Infinium Rows": len(unmatched_inf),
        "Unmatched Infinium Amount": cents_to_float(_amount_total(inf, unmatched_inf)),
        "Duplicate QuickBooks Rows": len(qb_screen.duplicate_rows),
        "Duplicate QuickBooks Amount": cents_to_float(_amount_total(qb, qb_screen.duplicate_rows)),
        "Duplicate Infinium Rows": len(inf_screen.duplicate_rows),
        "Duplicate Infinium Amount": cents_to_float(_amount_total(inf, inf_screen.duplicate_rows)),
        "Suspected QuickBooks Duplicate Rows": len(qb_screen.suspected_rows),
        "Suspected Infinium Duplicate Rows": len(inf_screen.suspected_rows),
        "Suspected QuickBooks Secondary Duplicate Rows": (
            len(qb_secondary_screen.suspected_rows) if qb_secondary_screen else 0
        ),
        "Suspected Infinium Secondary Duplicate Rows": (
            len(inf_secondary_screen.suspected_rows) if inf_secondary_screen else 0
        ),
        "Duplicate QuickBooks Secondary Rows Excluded": (
            len(qb_secondary_screen.duplicate_rows) if qb_secondary_screen else 0
        ),
        "Duplicate Infinium Secondary Rows Excluded": (
            len(inf_secondary_screen.duplicate_rows) if inf_secondary_screen else 0
        ),
        "Duplicate Review Hold QuickBooks Rows": len(duplicate_review_hold_qb),
        "Duplicate Review Hold QuickBooks Amount": cents_to_float(
            _amount_total(qb, duplicate_review_hold_qb)
        ),
        "Duplicate Review Hold Infinium Rows": len(duplicate_review_hold_inf),
        "Duplicate Review Hold Infinium Amount": cents_to_float(
            _amount_total(inf, duplicate_review_hold_inf)
        ),
        "Historical Duplicate Review Hold QuickBooks Rows": (
            len(qb_secondary_screen.suspected_rows) if qb_secondary_screen else 0
        ),
        "Historical Duplicate Review Hold Infinium Rows": (
            len(inf_secondary_screen.suspected_rows) if inf_secondary_screen else 0
        ),
        "Reference-Matched Amount Variance Rows": len(amount_variance_analysis),
        "Amount Variance Review Hold QuickBooks Amount": cents_to_float(
            _amount_total(qb, amount_variance_review_hold_qb)
        ),
        "Amount Variance Review Hold Infinium Amount": cents_to_float(
            _amount_total(inf, amount_variance_review_hold_inf)
        ),
        "Amount Variance Potential Net Difference": (
            float(amount_variance_analysis["Potential Difference"].sum())
            if not amount_variance_analysis.empty else 0.0
        ),
        "Amount Variance Potential Gross Difference": (
            float(amount_variance_analysis["Absolute Difference"].sum())
            if not amount_variance_analysis.empty else 0.0
        ),
        "Fuzzy Match Review Hold Rows": len(fuzzy_match_review_hold_analysis),
        "Fuzzy Match Review Hold QuickBooks Amount": cents_to_float(
            _amount_total(qb, fuzzy_match_review_hold_qb)
        ),
        "Fuzzy Match Review Hold Infinium Amount": cents_to_float(
            _amount_total(inf, fuzzy_match_review_hold_inf)
        ),
        "Posting Blockers": "; ".join(posting_blockers) if posting_blockers else "None",
        "Posting Status": "REVIEW REQUIRED" if posting_blockers else "READY TO POST",
        "Posting Authorization": "DO NOT POST" if posting_blockers else "AUTHORIZED BY AUTOMATED CONTROLS",
        "QuickBooks Match Rate by Row": (
            (len(matched_q) + len(historical_matched_q)) / len(qb) if len(qb) else 0
        ),
        "QuickBooks Match Rate by Gross Amount": (
            (matched_q_gross + _gross_amount_total(qb, historical_matched_q))
            / qb_gross if qb_gross else 0
        ),
        "Invalid QuickBooks Amounts": int(qb[AMOUNT_CENTS].isna().sum()),
        "Invalid Infinium Amounts": int(inf[AMOUNT_CENTS].isna().sum()),
        "QuickBooks Subtotal Rows Excluded": int(metadata.get("qb_subtotal_rows_excluded", 0)),
        "Duplicate QuickBooks Item Groups": (
            int(duplicate_analysis["Duplicate Group ID"].nunique()) if not duplicate_analysis.empty else 0
        ),
        "Duplicate QuickBooks Confirmed Copy Sets": (
            int(duplicate_analysis["Confirmed Copy Set ID"].nunique())
            if not duplicate_analysis.empty else 0
        ),
        "Duplicate Infinium Item Groups": (
            int(infinium_duplicate_analysis["Duplicate Group ID"].nunique())
            if not infinium_duplicate_analysis.empty else 0
        ),
        "Duplicate Infinium Confirmed Copy Sets": (
            int(infinium_duplicate_analysis["Confirmed Copy Set ID"].nunique())
            if not infinium_duplicate_analysis.empty else 0
        ),
        "Control Status": "PASS" if controls["Status"].eq("PASS").all() else "FAIL",
    }

    result = ReconciliationResult(
        run_id=metadata["run_id"],
        run_timestamp=metadata["run_timestamp_dt"],
        qb_raw=qb_raw.copy(),
        inf_raw=inf_raw.copy(),
        qb_work=qb,
        inf_work=inf,
        matches=matches,
        paired_rows=paired_rows,
        candidates=candidates,
        normalization=normalization,
        assessments=assessments,
        method_summary=method_summary,
        exception_analysis=exception_analysis,
        amount_variance_analysis=amount_variance_analysis,
        duplicate_analysis=duplicate_analysis,
        infinium_duplicate_analysis=infinium_duplicate_analysis,
        product_summary=product_summary,
        controls=controls,
        metrics=metrics,
        rules=rules,
        config=config,
        qb_mapping=qb_mapping,
        inf_mapping=inf_mapping,
        metadata=metadata,
        unmatched_qb=unmatched_qb,
        unmatched_inf=unmatched_inf,
        historical_clearances=historical_clearances,
        qb_secondary_raw=qb_secondary_raw.copy() if qb_secondary_raw is not None else None,
        inf_secondary_raw=inf_secondary_raw.copy() if inf_secondary_raw is not None else None,
        qb_secondary_work=qb_secondary,
        inf_secondary_work=inf_secondary,
        qb_secondary_mapping=qb_secondary_mapping,
        inf_secondary_mapping=inf_secondary_mapping,
        duplicate_qb_rows=qb_screen.duplicate_rows,
        duplicate_inf_rows=inf_screen.duplicate_rows,
        duplicate_qb_secondary_rows=qb_secondary_screen.duplicate_rows if qb_secondary_screen else [],
        duplicate_inf_secondary_rows=inf_secondary_screen.duplicate_rows if inf_secondary_screen else [],
        suspected_qb_rows=qb_screen.suspected_rows,
        suspected_inf_rows=inf_screen.suspected_rows,
        suspected_qb_secondary_rows=qb_secondary_screen.suspected_rows if qb_secondary_screen else [],
        suspected_inf_secondary_rows=inf_secondary_screen.suspected_rows if inf_secondary_screen else [],
        duplicate_review_hold_qb_rows=duplicate_review_hold_qb,
        duplicate_review_hold_inf_rows=duplicate_review_hold_inf,
        amount_variance_review_hold_qb_rows=amount_variance_review_hold_qb,
        amount_variance_review_hold_inf_rows=amount_variance_review_hold_inf,
        fuzzy_match_review_hold_analysis=fuzzy_match_review_hold_analysis,
        fuzzy_match_review_hold_qb_rows=fuzzy_match_review_hold_qb,
        fuzzy_match_review_hold_inf_rows=fuzzy_match_review_hold_inf,
    )
    validate_reconciliation(result)
    return result


def validate_reconciliation(result: ReconciliationResult) -> None:
    if result.controls["Status"].ne("PASS").any():
        failures = result.controls.loc[result.controls["Status"] != "PASS", "Check"].tolist()
        raise ValueError(f"Reconciliation control failure: {', '.join(failures)}")
    duplicate_qb = set(result.duplicate_qb_rows)
    duplicate_inf = set(result.duplicate_inf_rows)
    review_hold_qb = set(result.duplicate_review_hold_qb_rows)
    review_hold_inf = set(result.duplicate_review_hold_inf_rows)
    variance_hold_qb = set(result.amount_variance_review_hold_qb_rows)
    variance_hold_inf = set(result.amount_variance_review_hold_inf_rows)
    fuzzy_hold_qb = set(result.fuzzy_match_review_hold_qb_rows)
    fuzzy_hold_inf = set(result.fuzzy_match_review_hold_inf_rows)
    if duplicate_qb.intersection(result.suspected_qb_rows):
        raise ValueError(
            "Duplicate disposition control failure: a QuickBooks row was both "
            "excluded and retained for review."
        )
    if duplicate_inf.intersection(result.suspected_inf_rows):
        raise ValueError(
            "Duplicate disposition control failure: an Infinium row was both "
            "excluded and retained for review."
        )
    if not review_hold_qb.issubset(result.suspected_qb_rows):
        raise ValueError(
            "Duplicate Review Hold control failure: a held QuickBooks row was never "
            "a suspected duplicate candidate."
        )
    if not review_hold_inf.issubset(result.suspected_inf_rows):
        raise ValueError(
            "Duplicate Review Hold control failure: a held Infinium row was never "
            "a suspected duplicate candidate."
        )
    if review_hold_qb.intersection(result.unmatched_qb) or review_hold_inf.intersection(result.unmatched_inf):
        raise ValueError(
            "Duplicate Review Hold control failure: a held row also remained in the "
            "unresolved population used for the accrual/journal entry total."
        )
    if duplicate_qb.intersection(review_hold_qb) or duplicate_inf.intersection(review_hold_inf):
        raise ValueError(
            "Duplicate Review Hold control failure: a row was both auto-excluded and "
            "placed in Duplicate Review Hold."
        )
    if (
        variance_hold_qb.intersection(duplicate_qb | review_hold_qb | set(result.unmatched_qb))
        or variance_hold_inf.intersection(duplicate_inf | review_hold_inf | set(result.unmatched_inf))
    ):
        raise ValueError(
            "Amount variance control failure: a variance-held row also appears in another "
            "financial disposition population."
        )
    variance_report = result.amount_variance_analysis
    if len(variance_report) != len(variance_hold_qb) or len(variance_report) != len(variance_hold_inf):
        raise ValueError(
            "Amount variance audit control failure: report and held-row counts do not agree."
        )
    if not variance_report.empty:
        if (
            variance_report["Variance ID"].duplicated().any()
            or variance_report["QuickBooks Row Index"].duplicated().any()
            or variance_report["Infinium Row Index"].duplicated().any()
        ):
            raise ValueError(
                "Amount variance audit control failure: variance relationships are not one-to-one."
            )
        if set(variance_report["QuickBooks Row Index"].astype(int)) != variance_hold_qb:
            raise ValueError("Amount variance audit control failure: QuickBooks row indexes disagree.")
        if set(variance_report["Infinium Row Index"].astype(int)) != variance_hold_inf:
            raise ValueError("Amount variance audit control failure: Infinium row indexes disagree.")
        for variance in variance_report.to_dict("records"):
            qidx = int(variance["QuickBooks Row Index"])
            iidx = int(variance["Infinium Row Index"])
            q_cents = int(result.qb_work.at[qidx, AMOUNT_CENTS])
            i_cents = int(result.inf_work.at[iidx, AMOUNT_CENTS])
            if q_cents == i_cents:
                raise ValueError("Amount variance control failure: held amounts unexpectedly agree.")
            po_agrees = bool(
                result.qb_work.at[qidx, NORM_PO]
                and result.qb_work.at[qidx, NORM_PO] == result.inf_work.at[iidx, NORM_PO]
            )
            invoice_agrees = bool(
                result.qb_work.at[qidx, NORM_INV]
                and result.qb_work.at[qidx, NORM_INV] == result.inf_work.at[iidx, NORM_INV]
            )
            if not (po_agrees or invoice_agrees):
                raise ValueError("Amount variance control failure: no exact reference agrees.")
            reported_difference = Decimal(str(variance["Potential Difference"])) * 100
            if reported_difference != Decimal(q_cents - i_cents):
                raise ValueError("Amount variance control failure: reported difference is incorrect.")
    if (
        fuzzy_hold_qb.intersection(
            duplicate_qb | review_hold_qb | variance_hold_qb | set(result.unmatched_qb)
        )
        or fuzzy_hold_inf.intersection(
            duplicate_inf | review_hold_inf | variance_hold_inf | set(result.unmatched_inf)
        )
    ):
        raise ValueError(
            "Fuzzy match review hold control failure: a fuzzy-held row also appears in "
            "another financial disposition population."
        )
    fuzzy_report = result.fuzzy_match_review_hold_analysis
    if fuzzy_report.empty != (not fuzzy_hold_qb and not fuzzy_hold_inf):
        raise ValueError(
            "Fuzzy match review hold audit control failure: report and held-row "
            "populations do not agree."
        )
    if not fuzzy_report.empty:
        reported_qb_rows: set[int] = set()
        reported_inf_rows: set[int] = set()
        for record in fuzzy_report.to_dict("records"):
            reported_qb_rows.update(
                int(idx) for idx in str(record["QuickBooks Row Indexes"]).split(";") if idx.strip()
            )
            reported_inf_rows.update(
                int(idx) for idx in str(record["Infinium Row Indexes"]).split(";") if idx.strip()
            )
        if reported_qb_rows != fuzzy_hold_qb:
            raise ValueError(
                "Fuzzy match review hold audit control failure: QuickBooks row indexes disagree."
            )
        if reported_inf_rows != fuzzy_hold_inf:
            raise ValueError(
                "Fuzzy match review hold audit control failure: Infinium row indexes disagree."
            )
    for group in result.matches:
        if review_hold_qb.intersection(group.qb_rows) or review_hold_inf.intersection(group.inf_rows):
            raise ValueError(
                "Duplicate Review Hold control failure: a held row was included in an "
                "accepted match."
            )
        if variance_hold_qb.intersection(group.qb_rows) or variance_hold_inf.intersection(group.inf_rows):
            raise ValueError(
                "Amount variance control failure: a variance-held row was included in an "
                "accepted match."
            )
        if fuzzy_hold_qb.intersection(group.qb_rows) or fuzzy_hold_inf.intersection(group.inf_rows):
            raise ValueError(
                "Fuzzy match review hold control failure: a fuzzy-held row was included in "
                "an accepted match."
            )
        if group.confidence == FUZZY_PO_CONFIDENCE:
            raise ValueError(
                "Fuzzy match review hold control failure: a fuzzy match was accepted "
                "instead of held for review."
            )
    for dataset, report, frame, id_column, excluded_rows in (
        ("QuickBooks", result.duplicate_analysis, result.qb_work, QB_ID, duplicate_qb | review_hold_qb),
        ("Infinium", result.infinium_duplicate_analysis, result.inf_work, INF_ID, duplicate_inf | review_hold_inf),
    ):
        if report.empty:
            if excluded_rows:
                raise ValueError(
                    f"Duplicate audit control failure: excluded {dataset} rows have no report."
                )
            continue
        required_columns = {
            "Duplicate Group ID", "Confirmed Copy Set ID", "Screening Stage",
            "Disposition", "Automatically Excluded", "Source Row ID",
            "Source Scope", "Potential Duplicate Reason",
            "Differing Confirmation Fields", "Duplicate Rule Version",
        }
        if not required_columns.issubset(report.columns):
            raise ValueError(f"Duplicate audit control failure: {dataset} report schema is incomplete.")
        if report["Duplicate Group ID"].isna().any() or report["Duplicate Rule Version"].ne(DUPLICATE_RULE_VERSION).any():
            raise ValueError(f"Duplicate audit control failure: {dataset} group IDs or rule versions are invalid.")
        primary_report = report.loc[report["Source Scope"] == "Primary"]
        reported_excluded_ids = set(
            primary_report.loc[
                primary_report["Automatically Excluded"].fillna(False).astype(bool),
                "Source Row ID",
            ]
        )
        expected_excluded_ids = {frame.at[index, id_column] for index in excluded_rows}
        if reported_excluded_ids != expected_excluded_ids:
            raise ValueError(
                f"Duplicate audit control failure: excluded {dataset} rows do not agree with the report."
            )
        confirmed_primary = primary_report.loc[
            primary_report["Payload Confirmed"].fillna(False).astype(bool)
            & primary_report["Duplicate Basis"].ne("Primary/Historical overlap")
        ]
        for _, group in confirmed_primary.groupby("Confirmed Copy Set ID", sort=False):
            if group["Disposition"].eq("Retained canonical row").sum() != 1:
                raise ValueError(
                    f"Duplicate canonicalization control failure: {dataset} copy set does not retain exactly one canonical row."
                )
    for dataset, report, frame, id_column, excluded_rows, suspected_rows in (
        (
            "QuickBooks Historical", result.duplicate_analysis,
            result.qb_secondary_work, QB_ID,
            set(result.duplicate_qb_secondary_rows),
            set(result.suspected_qb_secondary_rows),
        ),
        (
            "Infinium Historical", result.infinium_duplicate_analysis,
            result.inf_secondary_work, INF_ID,
            set(result.duplicate_inf_secondary_rows),
            set(result.suspected_inf_secondary_rows),
        ),
    ):
        if frame is None:
            if excluded_rows or suspected_rows:
                raise ValueError(
                    f"Historical duplicate audit control failure: {dataset} has dispositions "
                    "without a working source frame."
                )
            continue
        historical_report = report.loc[
            report["Source Scope"].eq("Historical (Secondary)")
        ] if not report.empty else report
        reported_excluded_ids = set(
            historical_report.loc[
                historical_report["Automatically Excluded"].fillna(False).astype(bool),
                "Source Row ID",
            ]
        ) if not historical_report.empty else set()
        expected_excluded_ids = {
            frame.at[index, id_column]
            for index in excluded_rows | suspected_rows
        }
        if reported_excluded_ids != expected_excluded_ids:
            raise ValueError(
                f"Historical duplicate audit control failure: withheld {dataset} rows "
                "do not agree with the duplicate report."
            )
    has_posting_blockers = result.metrics.get("Posting Blockers", "None") != "None"
    if has_posting_blockers != (result.metrics.get("Posting Status") == "REVIEW REQUIRED"):
        raise ValueError("Posting control failure: status does not agree with posting blockers.")
    expected_authorization = "DO NOT POST" if has_posting_blockers else "AUTHORIZED BY AUTOMATED CONTROLS"
    if result.metrics.get("Posting Authorization") != expected_authorization:
        raise ValueError("Posting control failure: authorization does not agree with posting blockers.")
    for group in result.matches:
        q_count = len(group.qb_rows)
        i_count = len(group.inf_rows)
        if duplicate_qb.intersection(group.qb_rows) or duplicate_inf.intersection(group.inf_rows):
            raise ValueError(
                "Duplicate exclusion control failure: a duplicate row was included in an "
                "accepted match."
            )
        if group.group_level:
            if not (
                (q_count == 1 and 2 <= i_count <= MAX_GROUP_SIZE)
                or (i_count == 1 and 2 <= q_count <= MAX_GROUP_SIZE)
            ):
                raise ValueError(
                    "Grouped matching control failure: only bounded one-to-many "
                    "or many-to-one relationships are permitted."
                )
        elif q_count != 1 or i_count != 1:
            raise ValueError(
                "One-to-one matching control failure: a non-group match has "
                "unexpected cardinality."
            )
        q_total = _amount_total(result.qb_work, group.qb_rows)
        i_total = _amount_total(result.inf_work, group.inf_rows)
        if q_total != i_total:
            raise ValueError(
                "Matching control failure: accepted relationship totals do not agree."
            )
    if duplicate_qb.intersection(result.unmatched_qb) or duplicate_inf.intersection(result.unmatched_inf):
        raise ValueError(
            "Duplicate exclusion control failure: a duplicate row remained in the "
            "unresolved population used for the accrual/journal entry total."
        )
    if not result.historical_clearances.empty:
        clearance_differences = result.historical_clearances.groupby(
            "Clearance ID", sort=False
        )["Amount Difference"].sum().round(2)
        if clearance_differences.ne(0).any():
            raise ValueError(
                "Historical clearance control failure: aggregate signed amounts "
                "must agree exactly."
            )
        primary_rows = result.historical_clearances.loc[
            result.historical_clearances["Primary Row Index"].notna(),
            ["Primary Dataset", "Primary Row Index"],
        ]
        if primary_rows.duplicated().any():
            raise ValueError("Historical clearance control failure: a primary row was cleared more than once.")
        secondary_rows = result.historical_clearances.loc[
            result.historical_clearances["Secondary Row Index"].notna(),
            ["Secondary Dataset", "Secondary Row Index"],
        ]
        if secondary_rows.duplicated().any():
            raise ValueError("Historical clearance control failure: a secondary row was used more than once.")
        duplicate_qb_secondary = set(result.duplicate_qb_secondary_rows)
        duplicate_inf_secondary = set(result.duplicate_inf_secondary_rows)
        review_qb_secondary = set(result.suspected_qb_secondary_rows)
        review_inf_secondary = set(result.suspected_inf_secondary_rows)
        used_qb_secondary = {
            int(value) for value in result.historical_clearances.loc[
                result.historical_clearances["Secondary Dataset"] == "QuickBooks Secondary (Historical)",
                "Secondary Row Index",
            ].dropna()
        }
        used_inf_secondary = {
            int(value) for value in result.historical_clearances.loc[
                result.historical_clearances["Secondary Dataset"] == "Infinium Secondary (Historical)",
                "Secondary Row Index",
            ].dropna()
        }
        if duplicate_qb_secondary.intersection(used_qb_secondary) or duplicate_inf_secondary.intersection(used_inf_secondary):
            raise ValueError(
                "Duplicate exclusion control failure: a duplicated historical row was used to "
                "clear a primary exception."
            )
        if review_qb_secondary.intersection(used_qb_secondary) or review_inf_secondary.intersection(used_inf_secondary):
            raise ValueError(
                "Historical duplicate review control failure: an unresolved historical "
                "candidate was used to clear a primary exception."
            )
        for _, clearance in result.historical_clearances.groupby(
            "Clearance ID", sort=False
        ):
            primary_count = int(clearance["Primary Row Index"].notna().sum())
            secondary_count = int(clearance["Secondary Row Index"].notna().sum())
            group_level = bool(clearance["Group-Level Match"].iloc[0])
            if group_level and not (
                (primary_count == 1 and 2 <= secondary_count <= MAX_GROUP_SIZE)
                or (secondary_count == 1 and 2 <= primary_count <= MAX_GROUP_SIZE)
            ):
                raise ValueError(
                    "Historical grouped clearance has invalid cardinality."
                )
            if not group_level and (primary_count != 1 or secondary_count != 1):
                raise ValueError(
                    "Historical one-to-one clearance has invalid cardinality."
                )
    exception_sections = {
        "02 Unmatched QuickBooks",
        "03 Unmatched Infinium",
        "04 Duplicate QuickBooks",
        "05 Duplicate Infinium",
        "06 Duplicate Review Hold QuickBooks",
        "07 Duplicate Review Hold Infinium",
        "08 Reference-Matched Amount Variance Review Hold",
        "09 Fuzzy Match Review Hold",
    }
    duplicate_sections = {"04 Duplicate QuickBooks", "05 Duplicate Infinium"}
    for row in result.paired_rows:
        section = row.get("Section")
        if section not in exception_sections:
            continue
        if not row.get("Exception Cause") or not row.get("Financial Treatment"):
            raise ValueError(
                "Exception-cause audit control failure: an exception row lacks its "
                "standardized cause or financial treatment."
            )
        if section in duplicate_sections and not all(
            row.get(field)
            for field in (
                "Duplicate Values", "Duplicate Group ID",
                "Confirmed Copy Set ID", "Canonical Source Row ID",
            )
        ):
            raise ValueError(
                "Duplicate-value audit control failure: a confirmed duplicate row lacks "
                "the values or canonical linkage required in Reconciled Data."
            )
        if section in {
            "06 Duplicate Review Hold QuickBooks",
            "07 Duplicate Review Hold Infinium",
        } and not row.get("Potential Duplicate Reason"):
            raise ValueError(
                "Weak-basis duplicate audit control failure: a review-held row lacks "
                "the specific evidence that caused its classification."
            )
    qb_occurrences = Counter(
        row["QB Index"]
        for row in result.paired_rows
        if row["QB Index"] is not None and row.get("QB Record Scope") == "Primary"
    )
    inf_occurrences = Counter(
        row["Infinium Index"]
        for row in result.paired_rows
        if row["Infinium Index"] is not None
        and row.get("Infinium Record Scope") == "Primary"
    )
    expected_qb = set(result.qb_work.index)
    expected_inf = set(result.inf_work.index)
    if set(qb_occurrences) != expected_qb or any(count != 1 for count in qb_occurrences.values()):
        raise ValueError("QuickBooks rows were dropped or duplicated while building the reconciled output.")
    if set(inf_occurrences) != expected_inf or any(count != 1 for count in inf_occurrences.values()):
        raise ValueError("Infinium rows were dropped or duplicated while building the reconciled output.")
