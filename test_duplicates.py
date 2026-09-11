"""Unit tests for duplicates.py -- the duplicate detection/exclusion rules.

These operate directly on the module's working-frame column constants so
they stay isolated from matching.py's ingestion/normalization pipeline
(that pipeline is covered in test_matching.py's full-reconciliation tests).
"""

import pandas as pd
import pytest

from duplicates import (
    AMOUNT_CENTS,
    DUPLICATE_ANALYSIS_COLUMNS,
    DUPLICATE_BASIS_INVOICE_ONLY,
    DUPLICATE_BASIS_PO_ONLY,
    DUPLICATE_BASIS_STRICT,
    NORM_INV,
    NORM_PO,
    SOURCE_POS,
    build_duplicate_item_report,
    combine_duplicate_reports,
    duplicate_row_basis,
    duplicate_row_indexes,
    screen_duplicates,
)

ID_COL = "ID"


def make_frame(*rows):
    """rows: (id, po, invoice, amount_cents_or_None) tuples."""
    records = [
        {ID_COL: row_id, NORM_PO: po, NORM_INV: inv, AMOUNT_CENTS: amount}
        for row_id, po, inv, amount in rows
    ]
    frame = pd.DataFrame(records)
    frame[SOURCE_POS] = range(len(frame))
    return frame


# ---------------------------------------------------------------------------
# duplicate_row_basis / duplicate_row_indexes -- what counts as a duplicate
# ---------------------------------------------------------------------------

def test_strict_duplicate_flagged_when_po_invoice_and_amount_all_match():
    frame = make_frame(
        ("A", "PO1", "INV1", 1000),
        ("B", "PO1", "INV1", 1000),
        ("C", "PO2", "INV2", 2000),
    )
    basis = duplicate_row_basis(frame)
    assert basis == {0: DUPLICATE_BASIS_STRICT, 1: DUPLICATE_BASIS_STRICT}
    assert duplicate_row_indexes(frame) == {0, 1}


def test_shared_po_with_different_amount_is_never_a_duplicate():
    """Rule 2: a shared PO with a different amount is not assumed duplicate."""
    frame = make_frame(
        ("A", "PO1", "INV1", 1000),
        ("B", "PO1", "INV2", 2000),
    )
    assert duplicate_row_indexes(frame) == set()


def test_shared_po_and_amount_with_different_invoice_is_not_flagged():
    """Both rows have a populated, differing invoice -- too weak to assume
    duplication even though PO and amount agree; conservative by design."""
    frame = make_frame(
        ("A", "PO1", "INV1", 1000),
        ("B", "PO1", "INV2", 1000),
    )
    assert duplicate_row_indexes(frame) == set()


def test_po_only_fallback_when_invoice_blank_on_both_rows():
    frame = make_frame(
        ("A", "PO1", "", 1000),
        ("B", "PO1", "", 1000),
        ("C", "PO1", "INV3", 1000),  # populated invoice: must not join the blank-invoice group
    )
    basis = duplicate_row_basis(frame)
    assert basis == {0: DUPLICATE_BASIS_PO_ONLY, 1: DUPLICATE_BASIS_PO_ONLY}
    assert 2 not in basis


def test_invoice_only_fallback_when_po_blank_on_both_rows():
    frame = make_frame(
        ("A", "", "INV1", 1000),
        ("B", "", "INV1", 1000),
        ("C", "PO9", "INV1", 1000),  # populated PO: must not join the blank-PO group
    )
    basis = duplicate_row_basis(frame)
    assert basis == {0: DUPLICATE_BASIS_INVOICE_ONLY, 1: DUPLICATE_BASIS_INVOICE_ONLY}
    assert 2 not in basis


def test_both_references_blank_is_never_flagged():
    """A shared amount alone, with no PO or invoice on either side, is too
    weak a signal to safely exclude from the accrual."""
    frame = make_frame(
        ("A", "", "", 1000),
        ("B", "", "", 1000),
    )
    assert duplicate_row_indexes(frame) == set()


def test_missing_amount_rows_are_never_flagged():
    frame = make_frame(
        ("A", "PO1", "INV1", None),
        ("B", "PO1", "INV1", None),
    )
    assert duplicate_row_indexes(frame) == set()


def test_single_row_group_is_never_flagged():
    frame = make_frame(("A", "PO1", "INV1", 1000))
    assert duplicate_row_indexes(frame) == set()


def test_group_of_three_all_flagged():
    frame = make_frame(
        ("A", "PO1", "INV1", 1000),
        ("B", "PO1", "INV1", 1000),
        ("C", "PO1", "INV1", 1000),
    )
    assert duplicate_row_indexes(frame) == {0, 1, 2}


# ---------------------------------------------------------------------------
# build_duplicate_item_report -- itemized, per-row reporting
# ---------------------------------------------------------------------------

def test_item_report_has_one_row_per_duplicate_and_correct_columns():
    frame = make_frame(
        ("A", "PO1", "INV1", 1000),
        ("B", "PO1", "INV1", 1000),
        ("C", "PO1", "INV1", 1000),
        ("D", "PO2", "INV2", 5000),
    )
    report = build_duplicate_item_report(
        frame, duplicate_row_indexes(frame), ID_COL, "QuickBooks", "Primary",
        "Excluded from the accrual total.",
    )
    assert list(report.columns) == DUPLICATE_ANALYSIS_COLUMNS
    assert len(report) == 3
    assert set(report["Source Row ID"]) == {"A", "B", "C"}
    assert set(report["Duplicate Group Size"]) == {3}
    assert set(report["Dataset"]) == {"QuickBooks"}
    assert set(report["Source Scope"]) == {"Primary"}

    row_a = report.loc[report["Source Row ID"] == "A"].iloc[0]
    assert set(row_a["Other Source Row IDs In Group"].split("; ")) == {"B", "C"}
    assert row_a["Amount"] == 10.0


def test_item_report_is_empty_and_correctly_columned_when_no_duplicates():
    frame = make_frame(("A", "PO1", "INV1", 1000))
    report = build_duplicate_item_report(
        frame, set(), ID_COL, "QuickBooks", "Primary", "n/a",
    )
    assert report.empty
    assert list(report.columns) == DUPLICATE_ANALYSIS_COLUMNS


def test_item_report_distinguishes_basis_and_uses_correct_group_key():
    """A PO-only group and an unrelated invoice-only group must not bleed
    into each other's "Other Source Row IDs In Group" listing."""
    frame = make_frame(
        ("A", "PO1", "", 1000),
        ("B", "PO1", "", 1000),
        ("C", "", "INV9", 2000),
        ("D", "", "INV9", 2000),
    )
    report = build_duplicate_item_report(
        frame, duplicate_row_indexes(frame), ID_COL, "QuickBooks", "Primary", "x",
    )
    assert len(report) == 4
    row_a = report.loc[report["Source Row ID"] == "A"].iloc[0]
    row_c = report.loc[report["Source Row ID"] == "C"].iloc[0]
    assert row_a["Duplicate Basis"] == DUPLICATE_BASIS_PO_ONLY
    assert row_a["Other Source Row IDs In Group"] == "B"
    assert row_c["Duplicate Basis"] == DUPLICATE_BASIS_INVOICE_ONLY
    assert row_c["Other Source Row IDs In Group"] == "D"


# ---------------------------------------------------------------------------
# combine_duplicate_reports
# ---------------------------------------------------------------------------

def test_combine_duplicate_reports_merges_nonempty_reports():
    frame = make_frame(("A", "PO1", "INV1", 1000), ("B", "PO1", "INV1", 1000))
    report = build_duplicate_item_report(
        frame, duplicate_row_indexes(frame), ID_COL, "QuickBooks", "Primary", "x",
    )
    empty = pd.DataFrame(columns=DUPLICATE_ANALYSIS_COLUMNS)
    combined = combine_duplicate_reports(report, empty, None)
    assert len(combined) == 2


def test_combine_duplicate_reports_all_empty_returns_empty_with_columns():
    empty = pd.DataFrame(columns=DUPLICATE_ANALYSIS_COLUMNS)
    combined = combine_duplicate_reports(empty, None)
    assert combined.empty
    assert list(combined.columns) == DUPLICATE_ANALYSIS_COLUMNS


# ---------------------------------------------------------------------------
# screen_duplicates -- the function matching.py actually calls
# ---------------------------------------------------------------------------

def test_screen_duplicates_removes_duplicates_from_active_frame():
    frame = make_frame(
        ("A", "PO1", "INV1", 1000),
        ("B", "PO1", "INV1", 1000),
        ("C", "PO2", "INV2", 2000),
    )
    result = screen_duplicates(frame, ID_COL, "QuickBooks", "Primary", "excluded")
    assert result.duplicate_rows == [0, 1]
    assert list(result.active_frame[ID_COL]) == ["C"]
    assert len(result.report) == 2
    assert set(result.report["Source Row ID"]) == {"A", "B"}


def test_screen_duplicates_with_no_duplicates_keeps_full_frame():
    frame = make_frame(("A", "PO1", "INV1", 1000), ("B", "PO2", "INV2", 2000))
    result = screen_duplicates(frame, ID_COL, "QuickBooks", "Primary", "excluded")
    assert result.duplicate_rows == []
    assert list(result.active_frame[ID_COL]) == ["A", "B"]
    assert result.report.empty
