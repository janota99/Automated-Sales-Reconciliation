"""Tests for matching.py: normalization helpers, the matching engine's core
guarantees, and full end-to-end reconciliation runs.

The full-reconciliation tests double as regression tests for the four
duplicate rules and the two reliability gaps (historical-duplicate
screening, blank-reference duplicates) fixed in duplicates.py -- if a
future change to either module reintroduces one of those issues, one of
these tests should fail.
"""

import pandas as pd
import pytest

from matching import (
    FUZZY_MATCH_CLASSIFICATION_GROUPED,
    FUZZY_MATCH_CLASSIFICATION_SINGLE,
    MatchGroup,
    build_fiscal_exception_summary,
    build_fuzzy_match_review_holds,
    build_reconciliation,
    cents_or_zero,
    cents_to_float,
    clean_alphanumeric,
    clean_po,
    get_fuzzy_lexicon_match,
    parse_amount_cents,
    parse_fiscal_period,
    perform_matching,
    prepare_working_frame,
    valid_cents,
    validate_reconciliation,
)
from vendor_aliases import ALIAS_CONFIDENCE, ALIAS_METHOD, VendorAlias


def _prepare(rows, mapping, source, fiscal_year=2026):
    return prepare_working_frame(pd.DataFrame(rows), mapping, source, fiscal_year)


# ---------------------------------------------------------------------------
# Normalization helpers
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("raw, expected", [
    ("PO# 12345", "12345"),
    ("P.O. 12345", "12345"),
    ("12345.00", "12345"),
    ("PO-12345", "12345"),
    (None, ""),
])
def test_clean_po(raw, expected):
    assert clean_po(raw) == expected


@pytest.mark.parametrize("raw, expected", [
    ("INV-100", "INV100"),
    ("100.00", "100"),
    ("  inv 001  ", "INV001"),
    (None, ""),
])
def test_clean_alphanumeric(raw, expected):
    assert clean_alphanumeric(raw) == expected


@pytest.mark.parametrize("raw, expected_cents", [
    ("$1,234.56", 123456),
    ("(100.00)", -10000),
    ("100", 10000),
    ("abc", None),
    (None, None),
    (True, None),
])
def test_parse_amount_cents(raw, expected_cents):
    assert parse_amount_cents(raw) == expected_cents


def test_cents_helpers():
    assert cents_to_float(10000) == 100.0
    assert cents_to_float(None) == 0.0
    assert cents_or_zero(None) == 0
    assert cents_or_zero(500) == 500
    assert valid_cents(None) is False
    assert valid_cents(0) is True


@pytest.mark.parametrize("raw, expected", [
    ("P01", (1, 2026)),
    ("PD 3-2025", (3, 2025)),
    ("garbage!!", (None, None)),
])
def test_parse_fiscal_period(raw, expected):
    assert parse_fiscal_period(raw, 2026) == expected


def test_fuzzy_lexicon_match_never_affects_matching_only_labels_products():
    assert get_fuzzy_lexicon_match("FC 24") == "Food Club 24 Case"
    assert get_fuzzy_lexicon_match("PPL24") == "Panhandle Pure 24 Case"
    assert get_fuzzy_lexicon_match("totally unknown product") is None


# ---------------------------------------------------------------------------
# perform_matching -- the core engine, in isolation from ingestion/duplicates
# ---------------------------------------------------------------------------

def test_one_to_one_strong_match(qb_mapping, inf_mapping):
    qb = _prepare(
        [{"PO": "PO1", "Invoice": "INV1", "Amount": 100.0, "Qty": 1, "Period": "1"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "PO1", "Invoice": "INV1", "Amount": 100.0, "Period": "1"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert len(matches) == 1
    assert matches[0].method == "PO + Invoice + Amount"
    assert matches[0].qb_rows == [0] and matches[0].inf_rows == [0]
    assert unmatched_qb == [] and unmatched_inf == []


def test_invoice_and_amount_match_is_strong_confidence_even_with_no_po(qb_mapping, inf_mapping):
    """Invoice number is the designated secondary reconciliation check, not
    a lesser one -- a unique match on invoice + amount must be tagged
    'Strong' confidence the same as a PO match, even when the PO itself is
    missing or doesn't agree between the two sides."""
    qb = _prepare(
        [{"PO": "", "Invoice": "INV1", "Amount": 100.0, "Qty": 1, "Period": "1"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "SOMETHING-ELSE", "Invoice": "INV1", "Amount": 100.0, "Period": "1"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert len(matches) == 1
    assert matches[0].method == "Invoice + Amount"
    assert matches[0].confidence == "Strong"
    assert unmatched_qb == [] and unmatched_inf == []


def test_grouped_aggregate_sums_two_differing_amounts_to_one_entry(qb_mapping, inf_mapping):
    """Rule 2 at the engine level: two QuickBooks rows sharing a PO with
    different amounts are matched as a group when they sum exactly to one
    Infinium entry -- never assumed to be duplicates of each other."""
    qb = _prepare(
        [
            {"PO": "PO1", "Invoice": "INVA", "Amount": 30.0, "Qty": 1, "Period": "1"},
            {"PO": "PO1", "Invoice": "INVB", "Amount": 70.0, "Qty": 1, "Period": "1"},
        ],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "PO1", "Invoice": "INVAGG", "Amount": 100.0, "Period": "1"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert len(matches) == 1
    assert matches[0].group_level is True
    assert sorted(matches[0].qb_rows) == [0, 1]
    assert unmatched_qb == [] and unmatched_inf == []


def test_ambiguous_grouped_candidates_are_left_unresolved(qb_mapping, inf_mapping):
    """Two equally valid ways to sum to the same target (10+40 or 20+30 both
    equal 50) must never be guessed -- both stay unresolved for review."""
    qb = _prepare(
        [
            {"PO": "PO1", "Invoice": "", "Amount": 10.0, "Qty": 1, "Period": "1"},
            {"PO": "PO1", "Invoice": "", "Amount": 20.0, "Qty": 1, "Period": "1"},
            {"PO": "PO1", "Invoice": "", "Amount": 30.0, "Qty": 1, "Period": "1"},
            {"PO": "PO1", "Invoice": "", "Amount": 40.0, "Qty": 1, "Period": "1"},
        ],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "PO1", "Invoice": "", "Amount": 50.0, "Period": "1"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert matches == []
    assert sorted(unmatched_qb) == [0, 1, 2, 3]
    assert unmatched_inf == [0]


def test_invalid_amount_rows_are_never_matched(qb_mapping, inf_mapping):
    qb = _prepare(
        [{"PO": "PO1", "Invoice": "INV1", "Amount": "not-a-number", "Qty": 1, "Period": "1"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "PO1", "Invoice": "INV1", "Amount": "not-a-number", "Period": "1"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, candidates = perform_matching(qb, inf)
    assert matches == []
    assert unmatched_qb == [0] and unmatched_inf == [0]
    assert candidates.iloc[0]["Disposition"] == "Invalid or missing QuickBooks amount"


def test_fuzzy_po_pass_resolves_the_hopper_scenario_after_exact_passes(qb_mapping, inf_mapping):
    """QuickBooks 'Hopper' vs Infinium 'DAVID HOPPER 2.2' -- the reported
    real-world case. No exact pass can match these (different normalized
    PO, no invoice overlap), so only the fuzzy PO pass should resolve it."""
    qb = _prepare(
        [{"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 225.00, "Period": "6"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert len(matches) == 1
    assert matches[0].method == "Fuzzy PO + Amount (Token Intersection & Aggregate)"
    assert matches[0].confidence == "Fuzzy"
    assert matches[0].qb_rows == [0] and matches[0].inf_rows == [0]
    assert unmatched_qb == [] and unmatched_inf == []


def test_fuzzy_po_pass_never_overrides_amount_mismatch(qb_mapping, inf_mapping):
    qb = _prepare(
        [{"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 999.00, "Period": "6"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert matches == []
    assert unmatched_qb == [0] and unmatched_inf == [0]


def test_fuzzy_po_pass_never_guesses_among_ambiguous_candidates(qb_mapping, inf_mapping):
    """One QB row fuzzy-matches two Infinium rows at the same amount --
    the engine must leave all three unresolved rather than guess."""
    qb = _prepare(
        [{"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [
            {"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 225.00, "Period": "6"},
            {"PO": "HOPPER LOGISTICS", "Invoice": "88888", "Amount": 225.00, "Period": "6"},
        ],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert matches == []
    assert unmatched_qb == [0]
    assert sorted(unmatched_inf) == [0, 1]


def test_fuzzy_po_pass_ignores_an_unrelated_near_miss_row(qb_mapping, inf_mapping):
    """The reported real-world gap: a clean, self-contained Hopper/David
    Hopper pair was landing in Unmatched/Exceptions because an unrelated
    row elsewhere in the population ('MYSTERY SHOPPER PROGRAM') near-missed
    the 'HOPPER' token by raw string similarity and dragged the true pair
    into an unbounded graph component. Exact-token matches must resolve
    before near-miss ones are considered, so the unrelated row is left
    alone and the true pair still clears."""
    qb = _prepare(
        [{"PO": "hopper", "Invoice": "20044", "Amount": 675.00, "Qty": 1, "Period": "6"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [
            {"PO": "DAVID HOPPER", "Invoice": "99999", "Amount": 675.00, "Period": "6"},
            {"PO": "MYSTERY SHOPPER PROGRAM", "Invoice": "12345", "Amount": 300.00, "Period": "6"},
        ],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert len(matches) == 1
    assert matches[0].confidence == "Fuzzy"
    assert matches[0].qb_rows == [0] and matches[0].inf_rows == [0]
    assert unmatched_qb == []
    assert unmatched_inf == [1]


def test_vendor_alias_resolves_a_pair_no_fuzzy_rule_can_bridge(qb_mapping, inf_mapping):
    """The reported real-world gap, confirmed against the actual raw data:
    QuickBooks PO field literally says 'Hopper' (surname) while Infinium's
    field literally says 'David' (first name) for the same recurring
    dock-sale customer -- zero shared characters, so no fuzzy-text rule
    could ever have matched them. A confirmed vendor alias resolves it and
    is posted like an exact match rather than held for review."""
    qb = _prepare(
        [{"PO": "Hopper", "Invoice": "20044", "Amount": 675.00, "Qty": 1, "Period": "6"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "David", "Invoice": "99999", "Amount": 675.00, "Period": "6"}],
        inf_mapping, "INF",
    )
    alias = VendorAlias(
        "ALIAS-0001", ["HOPPER", "DAVID"], "David Hopper (dock sales)",
        "J. Reviewer", "2026-09-15", "Confirmed via golden-master validation.",
    )
    # Without the alias, this pair is unsolvable by fuzzy text matching.
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf)
    assert matches == []
    assert unmatched_qb == [0] and unmatched_inf == [0]

    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf, vendor_aliases=[alias])
    assert len(matches) == 1
    assert matches[0].method == ALIAS_METHOD
    assert matches[0].confidence == ALIAS_CONFIDENCE
    assert matches[0].qb_rows == [0] and matches[0].inf_rows == [0]
    assert unmatched_qb == [] and unmatched_inf == []


def test_vendor_alias_match_is_posted_not_held_for_review(qb_mapping, inf_mapping, make_metadata):
    """Unlike a fuzzy text guess, a confirmed alias is a decided fact and
    must never be pulled into the fuzzy match review hold."""
    qb_rows = [{"PO": "Hopper", "Invoice": "20044", "Amount": 675.00, "Qty": 1, "Period": "6"}]
    inf_rows = [{"PO": "David", "Invoice": "99999", "Amount": 675.00, "Period": "6"}]
    alias = VendorAlias(
        "ALIAS-0001", ["HOPPER", "DAVID"], "David Hopper (dock sales)",
        "J. Reviewer", "2026-09-15", "Confirmed via golden-master validation.",
    )
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026, vendor_aliases=[alias],
    )
    assert result.metrics["Fuzzy Match Review Hold Rows"] == 0
    assert result.metrics["Unresolved QuickBooks Rows"] == 0
    assert result.metrics["Control Status"] == "PASS"


def test_perform_matching_enable_fuzzy_false_leaves_hopper_row_unmatched(qb_mapping, inf_mapping):
    """Fuzzy matching is disabled for historical-clearance matching -- a
    Hopper-style row must stay unmatched rather than fuzzy-clear against a
    population that is not the accrual source of truth."""
    qb = _prepare(
        [{"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"}],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [{"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 225.00, "Period": "6"}],
        inf_mapping, "INF",
    )
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf, enable_fuzzy=False)
    assert matches == []
    assert unmatched_qb == [0] and unmatched_inf == [0]


def test_build_fuzzy_match_review_holds_classifies_single_and_grouped(qb_mapping, inf_mapping):
    qb = _prepare(
        [
            {"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"},
            {"PO": "Hopper", "Invoice": "20055", "Amount": 100.00, "Qty": 1, "Period": "6"},
            {"PO": "Hopper", "Invoice": "20066", "Amount": 125.00, "Qty": 1, "Period": "6"},
        ],
        qb_mapping, "QB",
    )
    inf = _prepare(
        [
            {"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 225.00, "Period": "6"},
            {"PO": "DAVID HOPPER 3.1", "Invoice": "88888", "Amount": 225.00, "Period": "6"},
        ],
        inf_mapping, "INF",
    )
    single_group = MatchGroup([0], [0], "Fuzzy PO", "Fuzzy", "single-row text variant")
    grouped_group = MatchGroup([1, 2], [1], "Fuzzy PO", "Fuzzy", "multi-row netting", group_level=True)
    analysis = build_fuzzy_match_review_holds(qb, inf, [single_group, grouped_group])
    assert len(analysis) == 2
    single_row = analysis.iloc[0]
    grouped_row = analysis.iloc[1]
    assert single_row["Classification"] == FUZZY_MATCH_CLASSIFICATION_SINGLE
    assert single_row["QuickBooks Row Count"] == 1
    assert single_row["Infinium Row Count"] == 1
    assert grouped_row["Classification"] == FUZZY_MATCH_CLASSIFICATION_GROUPED
    assert grouped_row["QuickBooks Row Count"] == 2
    assert grouped_row["Infinium Row Count"] == 1
    assert grouped_row["QuickBooks Amount"] == pytest.approx(225.00)
    assert grouped_row["Infinium Amount"] == pytest.approx(225.00)
    assert "HOPPER" in grouped_row["Matched Text Evidence"]


def test_full_reconciliation_resolves_hopper_style_po_mismatch(qb_mapping, inf_mapping, make_metadata):
    """The fuzzy pass must also fire through the full build_reconciliation
    pipeline, but a fuzzy match is a text-similarity guess, not a certain
    relationship -- it must never be posted like an exact match. It goes to
    its own review-hold population instead, and blocks posting."""
    qb_rows = [
        {"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"},
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "6"},
    ]
    inf_rows = [
        {"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 225.00, "Period": "6"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.metrics["Control Status"] == "PASS"
    assert result.metrics["Unresolved QuickBooks Rows"] == 1
    assert not any(g.confidence == "Fuzzy" for g in result.matches)
    assert result.metrics["Fuzzy Match Review Hold Rows"] == 1
    assert len(result.fuzzy_match_review_hold_qb_rows) == 1
    assert len(result.fuzzy_match_review_hold_inf_rows) == 1
    assert "Fuzzy Match" in result.fuzzy_match_review_hold_analysis.iloc[0]["Classification"]
    assert result.metrics["Posting Status"] == "REVIEW REQUIRED"


# ---------------------------------------------------------------------------
# Full end-to-end build_reconciliation runs -- the four duplicate rules and
# the two reliability fixes, exercised together the way app.py would.
# ---------------------------------------------------------------------------

def test_duplicate_group_one_of_two_matches_other_becomes_ordinary_exception(
    qb_mapping, inf_mapping, make_metadata,
):
    """The reported example 1: two $500 QuickBooks entries share a PO. One
    has a real Infinium counterpart and matches; the other does not. The
    unmatched one must be included in the accrual as an ordinary exception
    -- not silently collapsed away as if it were a data-entry duplicate --
    because its sibling matching is real evidence the group is genuinely
    two transactions, not one entered twice."""
    qb_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert len(result.matches) == 1
    assert result.metrics["Duplicate QuickBooks Rows"] == 0
    assert result.metrics["Unresolved QuickBooks Rows"] == 1
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(500.00)
    assert result.metrics["Control Status"] == "PASS"


def test_duplicate_group_none_match_only_one_included_in_accrual(
    qb_mapping, inf_mapping, make_metadata,
):
    """The reported example 2: two (or more) $500 QuickBooks entries share
    a PO and NEITHER reconciles to Infinium. With zero corroborating
    evidence more than one real transaction exists, only one instance is
    included in the accrual; the rest are excluded as excess -- the
    original duplicate-canonicalization outcome, just decided after
    matching is attempted instead of before."""
    qb_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "UNRELATED", "Invoice": "UNRELATED", "Amount": 1.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.matches == []
    assert result.metrics["Duplicate QuickBooks Rows"] == 1
    assert result.metrics["Duplicate QuickBooks Amount"] == pytest.approx(500.00)
    assert result.metrics["Unresolved QuickBooks Rows"] == 1
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(500.00)
    assert result.metrics["Control Status"] == "PASS"


def test_duplicate_group_of_three_one_matches_both_others_become_exceptions(
    qb_mapping, inf_mapping, make_metadata,
):
    """Generalizing example 1 to 3+ duplicates: once ANY member of the
    group matches, every other unmatched member becomes its own ordinary
    accrual exception -- not capped at one leftover the way the
    zero-evidence case is."""
    qb_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 500.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert len(result.matches) == 1
    assert result.metrics["Duplicate QuickBooks Rows"] == 0
    assert result.metrics["Unresolved QuickBooks Rows"] == 2
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(1000.00)
    assert result.metrics["Control Status"] == "PASS"


def test_full_reconciliation_applies_all_four_duplicate_rules(qb_mapping, inf_mapping, make_metadata):
    qb_rows = [
        {"PO": "PO100", "Invoice": "INV100", "Amount": 100.00, "Qty": 1, "Period": "1"},  # 1:1 match
        {"PO": "PO200", "Invoice": "INV200", "Amount": 50.00, "Qty": 1, "Period": "1"},    # true duplicate pair
        {"PO": "PO200", "Invoice": "INV200", "Amount": 50.00, "Qty": 1, "Period": "1"},
        {"PO": "PO300", "Invoice": "INV301", "Amount": 30.00, "Qty": 1, "Period": "1"},    # sums to one Infinium row
        {"PO": "PO300", "Invoice": "INV302", "Amount": 70.00, "Qty": 1, "Period": "1"},
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "1"},    # genuine exception
    ]
    inf_rows = [
        {"PO": "PO100", "Invoice": "INV100", "Amount": 100.00, "Period": "1"},
        {"PO": "PO400", "Invoice": "INV400", "Amount": 20.00, "Period": "1"},              # Infinium duplicate pair
        {"PO": "PO400", "Invoice": "INV400", "Amount": 20.00, "Period": "1"},
        {"PO": "PO300", "Invoice": "INV300AGG", "Amount": 100.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )

    # Rule 1: each duplicate pair is a full-payload strict match, so one row
    # is retained as canonical and only the excess copy is excluded from
    # the accrual/JE total. The canonical copy matches nothing else here,
    # so it becomes its own genuine unresolved exception -- it is not lost.
    assert result.metrics["Duplicate QuickBooks Rows"] == 1
    assert result.metrics["Duplicate QuickBooks Amount"] == pytest.approx(50.00)
    assert result.metrics["Unresolved QuickBooks Rows"] == 2  # PO999 + canonical PO200 copy
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(65.00)

    # Rule 2: differing-amount same-PO rows are not duplicates; the grouped
    # aggregate pass matched them to the Infinium entry instead.
    po300_matches = [
        g for g in result.matches
        if set(result.qb_work.loc[g.qb_rows, "PO"]) == {"PO300"}
    ]
    assert len(po300_matches) == 1
    assert po300_matches[0].group_level is True

    # Rule 3: QuickBooks and Infinium duplicates are two separate reports,
    # each itemizing both the canonical row and its excess copy.
    assert len(result.duplicate_analysis) == 2
    assert set(result.duplicate_analysis["Disposition"]) == {
        "Retained canonical row", "Excluded excess copy",
    }
    assert len(result.infinium_duplicate_analysis) == 2

    # Rule 4 / traceability: every dollar and row is still accounted for.
    assert result.metrics["Control Status"] == "PASS"
    assert result.controls["Status"].eq("PASS").all()


def test_full_reconciliation_canonicalizes_duplicated_historical_rows_before_clearance(
    qb_mapping, inf_mapping, make_metadata,
):
    """Two byte-identical historical Infinium rows are the same real
    prior-period transaction uploaded twice, not two independent
    transactions -- one is retained as canonical and may still clear a
    primary exception; only the excess copy is excluded. (Genuine
    cross-scope overlap -- a historical row duplicating something in the
    *primary* file -- is covered at the unit level in test_duplicates.py.)
    """
    qb_rows = [
        {"PO": "POHIST", "Invoice": "INVHIST", "Amount": 555.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "POOTHER", "Invoice": "INVOTHER", "Amount": 1.00, "Period": "1"},
    ]
    inf_secondary_rows = [
        {"PO": "POHIST", "Invoice": "INVHIST", "Amount": 555.00, "Period": "12"},
        {"PO": "POHIST", "Invoice": "INVHIST", "Amount": 555.00, "Period": "12"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
        inf_secondary_raw=pd.DataFrame(inf_secondary_rows),
        inf_secondary_mapping=inf_mapping,
    )
    assert result.metrics["Duplicate Infinium Secondary Rows Excluded"] == 1
    assert not result.historical_clearances.empty
    assert "POHIST" not in set(result.qb_work.loc[result.unmatched_qb, "PO"])
    assert result.metrics["Control Status"] == "PASS"


def test_full_reconciliation_detects_blank_reference_duplicates(qb_mapping, inf_mapping, make_metadata):
    """A duplicate pair missing its PO is a weak-basis (invoice-only)
    candidate: it is never auto-excluded up front, but since neither copy
    matches anything here, there is zero evidence more than one real
    transaction exists -- the earliest is retained as the sole accrual
    exception (canonical) and the other is excluded as excess, unified
    with how a strict PO+invoice+amount duplicate has always been treated
    (see test_full_reconciliation_applies_all_four_duplicate_rules)."""
    qb_rows = [
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "POX", "Invoice": "INVX", "Amount": 1.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.metrics["Duplicate QuickBooks Rows"] == 1
    assert result.metrics["Duplicate QuickBooks Amount"] == pytest.approx(25.00)
    assert result.metrics["Duplicate Review Hold QuickBooks Rows"] == 0
    assert result.metrics["Unresolved QuickBooks Rows"] == 1
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(25.00)
    # An automatic duplicate exclusion always requires the source-report
    # grain to be documented as validated before posting, same as any
    # other confirmed duplicate exclusion -- unrelated to this policy.
    assert result.metrics["Posting Status"] == "REVIEW REQUIRED"
    assert set(result.duplicate_analysis["Disposition"]) == {
        "Retained canonical row", "Excluded excess copy",
    }
    basis_values = set(result.duplicate_analysis["Duplicate Basis"])
    assert "Invoice + Amount (PO blank on both rows)" in basis_values


def test_fiscal_exception_summary_treats_one_and_two_periods_behind_as_routine(
    qb_mapping, inf_mapping, make_metadata,
):
    """A prior-period-close trickle is normal: an unresolved QuickBooks
    exception dated 1 or 2 fiscal periods behind the selected reporting
    period is a routine "Prior Period" exception, not the alarming
    "Urgent Prior Period" label -- that label is reserved for a
    gap wider than PRIOR_PERIOD_URGENT_THRESHOLD periods, which signals a
    genuinely stale item worth investigating now."""
    qb_rows = [
        {"PO": "PO-CUR", "Invoice": "INV-CUR", "Amount": 10.00, "Qty": 1, "Period": "5"},   # current
        {"PO": "PO-P4", "Invoice": "INV-P4", "Amount": 20.00, "Qty": 1, "Period": "4"},     # 1 behind
        {"PO": "PO-P3", "Invoice": "INV-P3", "Amount": 30.00, "Qty": 1, "Period": "3"},     # 2 behind
        {"PO": "PO-P2", "Invoice": "INV-P2", "Amount": 40.00, "Qty": 1, "Period": "2"},     # 3 behind
        {"PO": "PO-P1", "Invoice": "INV-P1", "Amount": 50.00, "Qty": 1, "Period": "1"},     # 4 behind
    ]
    inf_rows = [
        {"PO": "POX", "Invoice": "INVX", "Amount": 1.00, "Period": "5"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(fiscal_period=5), 2026,
    )
    summary = build_fiscal_exception_summary(result)
    classification_by_period = dict(zip(summary["Fiscal Period"], summary["Period Classification"]))
    assert classification_by_period["PD-05"] == "Current Period"
    assert classification_by_period["PD-04"] == "Prior Period"
    assert classification_by_period["PD-03"] == "Prior Period"
    assert classification_by_period["PD-02"] == "Urgent Prior Period"
    assert classification_by_period["PD-01"] == "Urgent Prior Period"


def test_ambiguous_duplicate_candidates_are_withheld_from_accrual(qb_mapping, inf_mapping, make_metadata):
    """A QuickBooks row whose PO/invoice is shared by two or more still-
    unresolved Infinium rows has no single correspondence that can be
    established -- it must not be posted as an ordinary exception (or
    silently guessed at), but withheld and reported as an ambiguous
    duplicate, distinct from both a confirmed same-side duplicate and a
    (single-candidate) reference-matched amount variance."""
    qb_rows = [
        {"PO": "PO-AMBIG", "Invoice": "INV-AMBIG", "Amount": 100.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO-AMBIG", "Invoice": "INV-AMBIG", "Amount": 90.00, "Period": "1"},
        {"PO": "PO-AMBIG", "Invoice": "INV-AMBIG", "Amount": 80.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.unmatched_qb == []
    assert len(result.ambiguous_duplicate_analysis) == 1
    assert result.metrics["Ambiguous Duplicate QuickBooks Rows"] == 1
    assert result.metrics["Ambiguous Duplicate QuickBooks Amount"] == pytest.approx(100.00)
    assert result.metrics["Unresolved QuickBooks Rows"] == 0
    assert result.amount_variance_analysis.empty
    ambiguous = result.ambiguous_duplicate_analysis.iloc[0]
    assert ambiguous["Candidate Count"] == 2
    assert ambiguous["Classification"] == "Ambiguous Duplicate - Multiple Candidates"


def test_po_reuse_error_flags_repeated_po_with_disagreeing_grouped_totals(
    qb_mapping, inf_mapping, make_metadata,
):
    """A PO reused across 2+ still-unresolved QuickBooks rows whose grouped
    total does not tie to the grouped Infinium total for that PO must be
    classified as a PO Re-use Error and reported with PO/QuickBooks
    total/Infinium total/difference/row counts -- but, unlike a duplicate
    or a review-hold row, it stays in the accrual rather than being
    withheld."""
    qb_rows = [
        {"PO": "PO-REUSE1", "Invoice": "INV-A", "Amount": 100.00, "Qty": 1, "Period": "1"},
        {"PO": "PO-REUSE1", "Invoice": "INV-B", "Amount": 50.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO-REUSE1", "Invoice": "", "Amount": 140.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert sorted(result.unmatched_qb) == [0, 1]
    assert result.metrics["Unresolved QuickBooks Rows"] == 2
    assert result.amount_variance_analysis.empty
    assert result.ambiguous_duplicate_analysis.empty
    assert len(result.po_reuse_errors) == 1
    po_reuse = result.po_reuse_errors.iloc[0]
    assert po_reuse["QuickBooks Total"] == pytest.approx(150.00)
    assert po_reuse["Infinium Total"] == pytest.approx(140.00)
    assert po_reuse["Difference"] == pytest.approx(10.00)
    assert po_reuse["QuickBooks Row Count"] == 2
    assert po_reuse["Infinium Row Count"] == 1
    assert result.metrics["PO Re-use Error Groups"] == 1
    assert result.metrics["PO Re-use Error QuickBooks Rows"] == 2
    assert result.metrics["PO Re-use Error Net Difference"] == pytest.approx(10.00)


def test_po_reuse_error_does_not_flag_a_repeated_po_whose_grouped_totals_tie_exactly(
    qb_mapping, inf_mapping, make_metadata,
):
    """When a repeated PO's QuickBooks total ties exactly to the Infinium
    total, the existing grouped-aggregate matching pass already resolves it
    as a real match (Priority 5) -- it must never reach the PO Re-use
    Error classification, per the requirement to preserve the existing
    resolution/review treatment when totals agree."""
    qb_rows = [
        {"PO": "PO-REUSE2", "Invoice": "INV-C", "Amount": 100.00, "Qty": 1, "Period": "1"},
        {"PO": "PO-REUSE2", "Invoice": "INV-D", "Amount": 50.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO-REUSE2", "Invoice": "", "Amount": 150.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.unmatched_qb == []
    assert result.po_reuse_errors.empty
    assert result.metrics["PO Re-use Error Groups"] == 0


def test_po_reuse_error_applies_zero_tolerance_at_one_cent(qb_mapping, inf_mapping, make_metadata):
    """Consistent with this application's established zero-tolerance amount
    policy (no rounding allowance anywhere else in the engine), a
    one-cent grouped-total difference on a reused PO must still be flagged
    -- not silently accepted as immaterial rounding."""
    qb_rows = [
        {"PO": "PO-REUSE3", "Invoice": "INV-E", "Amount": 100.00, "Qty": 1, "Period": "1"},
        {"PO": "PO-REUSE3", "Invoice": "INV-F", "Amount": 50.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO-REUSE3", "Invoice": "", "Amount": 149.99, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert len(result.po_reuse_errors) == 1
    assert result.po_reuse_errors.iloc[0]["Difference"] == pytest.approx(0.01)


def test_po_reuse_error_ignores_blank_po_values(qb_mapping, inf_mapping, make_metadata):
    """Two QuickBooks rows with a blank PO are not "the same PO reused" --
    grouping must key on a populated normalized PO, never on blank."""
    qb_rows = [
        {"PO": "", "Invoice": "INV-G", "Amount": 100.00, "Qty": 1, "Period": "1"},
        {"PO": "", "Invoice": "INV-H", "Amount": 50.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "", "Invoice": "INV-Z", "Amount": 1.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.po_reuse_errors.empty


def test_weak_basis_pair_resolved_via_match_proceeds_normally(qb_mapping, inf_mapping, make_metadata):
    """When a weak-basis (invoice-only) duplicate pair together satisfies a
    legitimate grouped aggregate match, both rows proceed normally: neither
    is excluded from the JE, neither is held for review, and their
    duplicate-report disposition reflects the successful resolution."""
    qb_rows = [
        {"PO": "", "Invoice": "INV-DUP", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "", "Invoice": "INV-DUP", "Amount": 25.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "", "Invoice": "INV-DUP", "Amount": 50.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.metrics["Duplicate Review Hold QuickBooks Rows"] == 0
    assert result.metrics["Unresolved QuickBooks Rows"] == 0
    assert result.metrics["Posting Status"] == "READY TO POST"
    assert len(result.matches) == 1
    assert result.matches[0].group_level is True
    assert sorted(result.matches[0].qb_rows) == [0, 1]
    assert len(result.duplicate_analysis) == 2
    assert set(result.duplicate_analysis["Disposition"]) == {"Resolved via match - no exclusion applied"}


def test_weak_basis_review_hold_does_not_block_control_status(qb_mapping, inf_mapping, make_metadata):
    """A pending Duplicate Review Hold item changes Posting Status but must
    never fail the underlying accounting controls -- every dollar and row
    is still fully accounted for. Infinium carries no accrual effect of its
    own, so its weak-basis duplicates keep the original review-hold policy
    (QuickBooks duplicates no longer use this path -- see
    test_full_reconciliation_detects_blank_reference_duplicates)."""
    qb_rows = [
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Period": "1"},
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.metrics["Control Status"] == "PASS"
    assert result.controls["Status"].eq("PASS").all()
    assert result.metrics["Posting Status"] == "REVIEW REQUIRED"
    assert result.metrics["Duplicate Review Hold Infinium Rows"] == 2
    assert result.metrics["Duplicate Review Hold Infinium Amount"] == pytest.approx(50.00)
    # QuickBooks accounting is entirely unaffected by the Infinium-side hold.
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(15.00)
    assert result.metrics["Duplicate QuickBooks Rows"] == 0


def test_rules_table_documents_duplicate_handling(qb_mapping, inf_mapping, make_metadata):
    result = build_reconciliation(
        pd.DataFrame([{"PO": "PO1", "Invoice": "INV1", "Amount": 1.0, "Qty": 1, "Period": "1"}]),
        pd.DataFrame([{"PO": "PO1", "Invoice": "INV1", "Amount": 1.0, "Period": "1"}]),
        qb_mapping, inf_mapping, make_metadata(), 2026,
    )
    rule_names = set(result.rules["Rule"])
    assert (
        "QuickBooks duplicate-key groups: matched normally, evidence-based accrual (see duplicates.py)"
        in rule_names
    )
    assert "Infinium duplicate-key groups: matched normally, held if unresolved" in rule_names
    assert "Historical overlap exclusion" in rule_names


def test_validate_reconciliation_rejects_duplicate_leaking_into_unresolved(
    qb_mapping, inf_mapping, make_metadata,
):
    """Direct test of the safety-net control added alongside duplicate
    screening: if a duplicate ever ends up back in the unresolved
    (accrual-driving) population, validation must fail loudly."""
    qb_rows = [
        {"PO": "PO1", "Invoice": "INV1", "Amount": 10.00, "Qty": 1, "Period": "1"},
        {"PO": "PO1", "Invoice": "INV1", "Amount": 10.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [{"PO": "POX", "Invoice": "INVX", "Amount": 1.00, "Period": "1"}]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.duplicate_qb_rows, "fixture must actually produce a duplicate to tamper with"
    result.unmatched_qb.append(result.duplicate_qb_rows[0])
    with pytest.raises(ValueError, match="Duplicate exclusion control failure"):
        validate_reconciliation(result)


def test_validate_reconciliation_rejects_duplicate_inside_an_accepted_match(
    qb_mapping, inf_mapping, make_metadata,
):
    qb_rows = [
        {"PO": "PO1", "Invoice": "INV1", "Amount": 10.00, "Qty": 1, "Period": "1"},
        {"PO": "PO1", "Invoice": "INV1", "Amount": 10.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [{"PO": "POX", "Invoice": "INVX", "Amount": 1.00, "Period": "1"}]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    result.matches.append(
        MatchGroup([result.duplicate_qb_rows[0]], [0], "Fabricated", "Strong", "test")
    )
    with pytest.raises(ValueError, match="Duplicate exclusion control failure"):
        validate_reconciliation(result)


def test_validate_reconciliation_rejects_fuzzy_held_row_in_accepted_match(
    qb_mapping, inf_mapping, make_metadata,
):
    """Direct test of the safety net added alongside the fuzzy match review
    hold: if a fuzzy-held row were ever also posted inside an accepted
    match, validation must fail loudly rather than silently double-count it."""
    qb_rows = [
        {"PO": "Hopper", "Invoice": "20044", "Amount": 225.00, "Qty": 3, "Period": "6"},
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "6"},
    ]
    inf_rows = [
        {"PO": "DAVID HOPPER 2.2", "Invoice": "99999", "Amount": 225.00, "Period": "6"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.fuzzy_match_review_hold_qb_rows, "fixture must actually produce a fuzzy hold to tamper with"
    result.matches.append(
        MatchGroup(
            [result.fuzzy_match_review_hold_qb_rows[0]],
            [result.fuzzy_match_review_hold_inf_rows[0]],
            "Fabricated", "Strong", "test",
        )
    )
    with pytest.raises(ValueError, match="Fuzzy match review hold control failure"):
        validate_reconciliation(result)
