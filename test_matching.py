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
    assert matches[0].method == "Fuzzy PO + Amount (Word Match, Unique)"
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


def test_full_reconciliation_resolves_hopper_style_po_mismatch(qb_mapping, inf_mapping, make_metadata):
    """The fuzzy pass must also fire through the full build_reconciliation
    pipeline, and its match must never be mistaken for an exact one."""
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
    fuzzy_matches = [g for g in result.matches if g.confidence == "Fuzzy"]
    assert len(fuzzy_matches) == 1
    assert "Fuzzy PO" in fuzzy_matches[0].method


# ---------------------------------------------------------------------------
# Full end-to-end build_reconciliation runs -- the four duplicate rules and
# the two reliability fixes, exercised together the way app.py would.
# ---------------------------------------------------------------------------

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

    # Rule 1: duplicates excluded from the accrual/JE total, but present.
    assert result.metrics["Duplicate QuickBooks Rows"] == 2
    assert result.metrics["Duplicate QuickBooks Amount"] == pytest.approx(100.00)
    assert result.metrics["Unresolved QuickBooks Rows"] == 1  # only PO999
    assert result.metrics["Unresolved QuickBooks Amount"] == pytest.approx(15.00)

    # Rule 2: differing-amount same-PO rows are not duplicates; the grouped
    # aggregate pass matched them to the Infinium entry instead.
    po300_matches = [
        g for g in result.matches
        if set(result.qb_work.loc[g.qb_rows, "PO"]) == {"PO300"}
    ]
    assert len(po300_matches) == 1
    assert po300_matches[0].group_level is True

    # Rule 3: QuickBooks and Infinium duplicates are two separate reports.
    assert len(result.duplicate_analysis) == 2
    assert set(result.duplicate_analysis["Dataset"]) == {"QuickBooks"}
    assert len(result.infinium_duplicate_analysis) == 2
    assert set(result.infinium_duplicate_analysis["Dataset"]) == {"Infinium"}

    # Rule 4 / traceability: every dollar and row is still accounted for.
    assert result.metrics["Control Status"] == "PASS"
    assert result.controls["Status"].eq("PASS").all()


def test_full_reconciliation_screens_historical_duplicates(qb_mapping, inf_mapping, make_metadata):
    """A duplicated historical Infinium row must never be allowed to clear
    a real primary QuickBooks exception."""
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
    assert result.metrics["Duplicate Infinium Secondary Rows Excluded"] == 2
    assert result.historical_clearances.empty
    assert "POHIST" in set(result.qb_work.loc[result.unmatched_qb, "PO"])
    assert result.metrics["Control Status"] == "PASS"


def test_full_reconciliation_detects_blank_reference_duplicates(qb_mapping, inf_mapping, make_metadata):
    """A duplicate pair missing its PO must still be caught via the
    invoice-only fallback, not silently left in the accrual."""
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
    assert result.metrics["Duplicate QuickBooks Rows"] == 2
    assert result.metrics["Unresolved QuickBooks Rows"] == 0
    basis_values = set(result.duplicate_analysis["Duplicate Basis"])
    assert "Invoice + Amount (PO blank on both rows)" in basis_values


def test_rules_table_documents_duplicate_handling(qb_mapping, inf_mapping, make_metadata):
    result = build_reconciliation(
        pd.DataFrame([{"PO": "PO1", "Invoice": "INV1", "Amount": 1.0, "Qty": 1, "Period": "1"}]),
        pd.DataFrame([{"PO": "PO1", "Invoice": "INV1", "Amount": 1.0, "Period": "1"}]),
        qb_mapping, inf_mapping, make_metadata(), 2026,
    )
    rule_names = set(result.rules["Rule"])
    assert "Exact duplicate exclusion (see duplicates.py)" in rule_names
    assert "Shared PO/invoice with differing amounts is not a duplicate" in rule_names
    assert "Historical/secondary duplicate exclusion" in rule_names


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
    from matching import MatchGroup
    result.matches.append(
        MatchGroup([result.duplicate_qb_rows[0]], [0], "Fabricated", "Strong", "test")
    )
    with pytest.raises(ValueError, match="Duplicate exclusion control failure"):
        validate_reconciliation(result)
