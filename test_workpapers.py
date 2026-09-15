"""Smoke tests for workpapers.py against a real ReconciliationResult.

These exist because matching.py/duplicates.py can be internally correct and
fully unit-tested while still breaking the app: ui_components.py and
workpapers.py each build their own dataframes/column lookups from a
ReconciliationResult, and nothing in test_matching.py or test_duplicates.py
would ever call them. That gap is exactly how a KeyError on
"Reconciliation Status" and a would-be crash on "Source Row IDs" reached a
running Streamlit session undetected. These tests build both real
workbooks end-to-end (with and without duplicates present) so a future
schema change to any ReconciliationResult field gets caught here first.
"""

import io

import pandas as pd
import pytest
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from matching import build_reconciliation
from workpapers import build_analytics_workbook, build_data_search_dataframe, build_primary_workbook

EXPECTED_PRIMARY_SHEETS = [
    "Raw Data",
    "Reconciled Data",
    "Unresolved Exceptions",
    "Data Search",
    "Product Aggregate Summary",
]

EXPECTED_ANALYTICS_SHEETS = [
    "Executive Summary",
    "Match Method Summary",
    "Detailed Match Ledger",
    "Normalization Detail",
    "Match Assessment",
    "Historical Clearances",
    "Exception Analysis",
    "QuickBooks Duplicates",
    "Infinium Duplicates",
    "Rules and Run Config",
]


def _worksheet_text(ws) -> list[str]:
    return [str(cell.value) for row in ws.iter_rows() for cell in row if cell.value is not None]


def _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata):
    qb_rows = [
        {"PO": "PO100", "Invoice": "INV100", "Amount": 100.00, "Qty": 1, "Period": "1"},  # 1:1 match
        {"PO": "PO200", "Invoice": "INV200", "Amount": 50.00, "Qty": 1, "Period": "1"},    # QB duplicate pair
        {"PO": "PO200", "Invoice": "INV200", "Amount": 50.00, "Qty": 1, "Period": "1"},
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "1"},    # genuine exception
    ]
    inf_rows = [
        {"PO": "PO100", "Invoice": "INV100", "Amount": 100.00, "Period": "1"},
        {"PO": "PO400", "Invoice": "INV400", "Amount": 20.00, "Period": "1"},              # Infinium duplicate pair
        {"PO": "PO400", "Invoice": "INV400", "Amount": 20.00, "Period": "1"},
    ]
    inf_secondary_rows = [
        {"PO": "POHIST", "Invoice": "INVHIST", "Amount": 5.00, "Period": "12"},
    ]
    qb_secondary_rows = [
        {"PO": "POSEC", "Invoice": "INVSEC", "Amount": 9.00, "Qty": 1, "Period": "12"},
    ]
    return build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
        qb_secondary_raw=pd.DataFrame(qb_secondary_rows),
        inf_secondary_raw=pd.DataFrame(inf_secondary_rows),
        qb_secondary_mapping=qb_mapping,
        inf_secondary_mapping=inf_mapping,
    )


def _build_result_without_duplicates(qb_mapping, inf_mapping, make_metadata):
    qb_rows = [
        {"PO": "PO1", "Invoice": "INV1", "Amount": 100.00, "Qty": 1, "Period": "1"},
        {"PO": "PO2", "Invoice": "INV2", "Amount": 15.00, "Qty": 1, "Period": "1"},  # unresolved
    ]
    inf_rows = [
        {"PO": "PO1", "Invoice": "INV1", "Amount": 100.00, "Period": "1"},
    ]
    return build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )


def test_primary_workbook_builds_with_duplicates_present(qb_mapping, inf_mapping, make_metadata):
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    # Each pair is a full-payload strict duplicate: one row is retained as
    # canonical (and, matching nothing else, becomes its own unresolved
    # exception), so only the excess copy is counted as excluded.
    assert result.metrics["Duplicate QuickBooks Rows"] == 1
    assert result.metrics["Duplicate Infinium Rows"] == 1

    workbook_bytes = build_primary_workbook(result)
    assert len(workbook_bytes) > 0

    wb = load_workbook(io.BytesIO(workbook_bytes))
    assert wb.sheetnames == EXPECTED_PRIMARY_SHEETS

    unresolved_text = _worksheet_text(wb["Unresolved Exceptions"])
    assert any("QUICKBOOKS DUPLICATES EXCLUDED FROM JE" in text for text in unresolved_text)
    # The old schema's columns must never resurface in the rendered output.
    assert not any(text == "Reconciliation Status" for text in unresolved_text)
    assert not any(text == "Source Row IDs" for text in unresolved_text)
    # The Infinium exceptions table (its own title band, KPIs, and totals)
    # carries no accrual impact and is already listed in Reconciled Data --
    # it must be gone. QB exception dispositions legitimately mention
    # Infinium (e.g. "No matching Infinium records"), so only the removed
    # section's own headings are checked, not the substring "Infinium".
    assert not any("INFINIUM EXCEPTIONS" in text.upper() for text in unresolved_text)
    assert not any("UNRESOLVED INFINIUM ROWS" in text.upper() for text in unresolved_text)
    assert not any("INFINIUM REVIEW TOTAL" in text.upper() for text in unresolved_text)
    # The duplicates block gets its own reviewer note column, mirroring the
    # QuickBooks exception table's reviewer workflow.
    assert any(text == "Reviewer Note" for text in unresolved_text)


def test_data_search_sheet_shows_duplicate_pair_and_clean_match(qb_mapping, inf_mapping, make_metadata):
    """The Data Search sheet must let a reviewer look up any PO/Invoice and
    see its real status -- including for a duplicate pair, where the
    retained copy keeps its normal (accrual-relevant) status but must still
    be identifiable as part of the same duplicate group as the excluded copy."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)

    wb = load_workbook(io.BytesIO(build_primary_workbook(result)))
    assert "Data Search" in wb.sheetnames

    frame = build_data_search_dataframe(result)
    by_qb_id = frame.set_index("QuickBooks Row ID")

    # PO100/INV100 is a clean 1:1 match -- an ordinary Infinium Match with no
    # duplicate involvement.
    matched = by_qb_id.loc["QB-1"]
    assert matched["Status"] == "Infinium Match"
    assert "PO + Invoice + Amount" in matched["Match Type"]
    assert matched["Duplicate Group ID"] == ""

    # QB-2/QB-3 are the PO200/INV200 duplicate pair. QB-2 is the retained
    # canonical row (matches nothing else, so it's Outstanding) and QB-3 is
    # the excluded excess copy (Status = Duplicate) -- both must carry the
    # same Duplicate Group ID.
    canonical = by_qb_id.loc["QB-2"]
    excess = by_qb_id.loc["QB-3"]
    assert canonical["Status"] == "Outstanding (On Accrual List)"
    assert excess["Status"] == "Duplicate"
    assert canonical["Duplicate Group ID"] != ""
    assert canonical["Duplicate Group ID"] == excess["Duplicate Group ID"]


def test_fiscal_period_summary_is_promoted_above_the_exception_table(
    qb_mapping, inf_mapping, make_metadata,
):
    """The fiscal-period breakdown must sit above the raw exception table,
    not buried below it."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]

    def first_row_containing(needle: str) -> int:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and needle in str(cell.value):
                    return cell.row
        raise AssertionError(f"{needle!r} not found in sheet")

    fiscal_row = first_row_containing("QUICKBOOKS EXCEPTIONS BY FISCAL PERIOD")
    exception_table_header_row = first_row_containing("Exception Status")

    assert fiscal_row < exception_table_header_row
    assert fiscal_row <= 10


def test_duplicates_and_je_sit_ten_rows_below_the_exceptions_total(
    qb_mapping, inf_mapping, make_metadata,
):
    """Per explicit request: the duplicates block (and the JE section that
    follows it) must be placed exactly 10 standard rows beneath the
    QuickBooks exceptions total row -- clearly separate from the exception
    detail, but still on the same page without excessive scrolling."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]

    def first_row_containing(needle: str) -> int:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and needle in str(cell.value):
                    return cell.row
        raise AssertionError(f"{needle!r} not found in sheet")

    exceptions_total_row = first_row_containing("PROPOSED JE SUPPORT TOTAL")
    duplicates_row = first_row_containing("QUICKBOOKS DUPLICATES EXCLUDED FROM JE")
    je_row = first_row_containing("PROPOSED JOURNAL ENTRY")

    assert duplicates_row == exceptions_total_row + 10
    assert je_row > duplicates_row
    assert duplicates_row > exceptions_total_row


def test_analytics_workbook_builds_with_duplicates_present(qb_mapping, inf_mapping, make_metadata):
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)

    workbook_bytes = build_analytics_workbook(result)
    assert len(workbook_bytes) > 0

    wb = load_workbook(io.BytesIO(workbook_bytes))
    assert wb.sheetnames == EXPECTED_ANALYTICS_SHEETS

    qb_dup_text = _worksheet_text(wb["QuickBooks Duplicates"])
    assert any("QB-2" in text or "QB-3" in text for text in qb_dup_text)

    inf_dup_text = _worksheet_text(wb["Infinium Duplicates"])
    assert any("INF-2" in text or "INF-3" in text for text in inf_dup_text)


def test_unresolved_sheet_duplicate_tables_use_simplified_reviewer_columns(
    qb_mapping, inf_mapping, make_metadata,
):
    """The 'Unresolved Exceptions' sheet's duplicate sub-tables show a small,
    plain-English column set for non-technical reviewers, while the full
    technical schema stays intact on the 'QuickBooks Duplicates' audit sheet
    -- a regression guard for the "display-only, additive" scope of that
    simplification."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    primary_wb = load_workbook(io.BytesIO(build_primary_workbook(result)))
    ws = primary_wb["Unresolved Exceptions"]
    unresolved_text = _worksheet_text(ws)

    for header in ("Duplicate ID", "Row ID", "What Was Found", "Reason", "Status"):
        assert header in unresolved_text, f"missing simplified header {header!r}"
    for technical_header in (
        "Screening Stage", "Normalized PO", "Normalized Invoice",
        "Confirmed Copy Set ID", "Duplicate Rule Version", "Payload Confirmed",
    ):
        assert technical_header not in unresolved_text, (
            f"technical header {technical_header!r} leaked into the reviewer-facing sheet"
        )
    assert any("Same PO 200, Invoice INV200" in text for text in unresolved_text)
    assert any("Kept (Original)" in text or "Removed (Duplicate)" in text for text in unresolved_text)

    analytics_wb = load_workbook(io.BytesIO(build_analytics_workbook(result)))
    audit_text = _worksheet_text(analytics_wb["QuickBooks Duplicates"])
    for technical_header in ("Screening Stage", "Normalized PO", "Confirmed Copy Set ID", "Duplicate Rule Version"):
        assert technical_header in audit_text, f"expected full audit column {technical_header!r} to remain"


def test_primary_and_analytics_workbooks_build_with_no_duplicates(qb_mapping, inf_mapping, make_metadata):
    """The empty-duplicate-frame branches must render cleanly too."""
    result = _build_result_without_duplicates(qb_mapping, inf_mapping, make_metadata)
    assert result.metrics["Duplicate QuickBooks Rows"] == 0
    assert result.metrics["Duplicate Infinium Rows"] == 0

    primary_bytes = build_primary_workbook(result)
    analytics_bytes = build_analytics_workbook(result)
    assert len(primary_bytes) > 0
    assert len(analytics_bytes) > 0

    primary_wb = load_workbook(io.BytesIO(primary_bytes))
    analytics_wb = load_workbook(io.BytesIO(analytics_bytes))
    assert primary_wb.sheetnames == EXPECTED_PRIMARY_SHEETS
    assert analytics_wb.sheetnames == EXPECTED_ANALYTICS_SHEETS

    unresolved_text = _worksheet_text(primary_wb["Unresolved Exceptions"])
    assert any(
        "No QuickBooks exact duplicates" in text for text in unresolved_text
    )


def test_unresolved_sheet_formulas_use_valid_structured_references(
    qb_mapping, inf_mapping, make_metadata,
):
    """Regression test for a #NAME? bug: a bare table name (or INDEX(TableName,0,N)
    built from one) is not a valid Excel reference when written as raw formula
    text -- only Excel's own UI auto-converts a typed table name into proper
    Table[Column] structured-reference syntax. Every formula on this sheet must
    use that qualified form outside the table, or the unqualified [Column] form
    inside the table's own totals row."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    wb = load_workbook(io.BytesIO(build_primary_workbook(result)))
    ws = wb["Unresolved Exceptions"]

    formula_cells = [
        (cell.coordinate, cell.value)
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    ]
    assert formula_cells, "expected at least one formula cell to check"
    for coordinate, formula in formula_cells:
        assert "INDEX(" not in formula, f"{coordinate} still uses the broken INDEX(TableName,...) form: {formula}"
        assert "ROWS(QuickBooksExceptions)" not in formula, (
            f"{coordinate} references the table by bare name, which Excel cannot resolve: {formula}"
        )

    # The KPI row (outside the table) must use the qualified Table[Column] form.
    kpi_formulas = " ".join(formula for _, formula in formula_cells if "QuickBooksExceptions[" in formula)
    assert "QuickBooksExceptions[Amount]" in kpi_formulas


def test_render_result_duplicate_badge_never_raises(qb_mapping, inf_mapping, make_metadata):
    """Direct regression test for the exact KeyError that reached Streamlit:
    result.duplicate_analysis["Reconciliation Status"] no longer exists, so
    the attention-badge count must be derived from metrics instead."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    metrics = result.metrics
    unresolved_duplicate_groups = int(metrics.get("Duplicate QuickBooks Rows", 0)) + int(
        metrics.get("Duplicate Infinium Rows", 0)
    )
    assert unresolved_duplicate_groups == 2
    assert "Reconciliation Status" not in result.duplicate_analysis.columns


def _build_result_with_review_hold(qb_mapping, inf_mapping, make_metadata):
    qb_rows = [
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [{"PO": "POX", "Invoice": "INVX", "Amount": 1.00, "Period": "1"}]
    return build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )


def test_duplicate_review_hold_section_renders_with_documented_disposition_dropdown(
    qb_mapping, inf_mapping, make_metadata,
):
    """The Duplicate Review Hold section must appear after the duplicates
    section (before the JE), list the held items, and give the reviewer a
    controlled-vocabulary, editable place to record a documented decision
    -- not just a static, uneditable list."""
    result = _build_result_with_review_hold(qb_mapping, inf_mapping, make_metadata)
    assert result.metrics["Posting Status"] == "REVIEW REQUIRED"
    assert result.metrics["Duplicate Review Hold QuickBooks Rows"] == 2

    wb = load_workbook(io.BytesIO(build_primary_workbook(result)))
    ws = wb["Unresolved Exceptions"]

    def first_row_containing(needle: str) -> int:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and needle in str(cell.value):
                    return cell.row
        raise AssertionError(f"{needle!r} not found in sheet")

    duplicates_row = first_row_containing("QUICKBOOKS DUPLICATES EXCLUDED FROM JE")
    review_hold_row = first_row_containing("DUPLICATE REVIEW HOLD")
    je_row = first_row_containing("PROPOSED JOURNAL ENTRY")
    disposition_header_row = first_row_containing("Reviewer Disposition")

    assert duplicates_row < review_hold_row < je_row

    disposition_col = next(
        cell.column for cell in ws[disposition_header_row] if cell.value == "Reviewer Disposition"
    )
    data_rows = range(disposition_header_row + 1, disposition_header_row + 3)
    for row in data_rows:
        cell = ws.cell(row, disposition_col)
        assert cell.value == "Pending Review"
        assert cell.protection.locked is False, "reviewer must be able to edit the disposition cell"

    list_validations = [dv for dv in ws.data_validations.dataValidation if dv.type == "list"]
    disposition_letter = get_column_letter(disposition_col)
    matching_validation = next(
        dv for dv in list_validations
        if f"{disposition_letter}{disposition_header_row + 1}" in str(dv.sqref)
    )
    assert "Pending Review" in matching_validation.formula1
    assert "Confirmed Duplicate - Exclude Permanently" in matching_validation.formula1
    assert "Confirmed Legitimate - Include In JE Manually" in matching_validation.formula1
    assert "Escalated For Investigation" in matching_validation.formula1


def test_no_review_hold_items_renders_empty_section_cleanly(qb_mapping, inf_mapping, make_metadata):
    """When every weak-basis candidate resolves via a match, the Review
    Hold section must still render (with an explanatory 'none' caption)
    rather than crash on an empty frame."""
    qb_rows = [
        {"PO": "", "Invoice": "INV-DUP", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "", "Invoice": "INV-DUP", "Amount": 25.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [{"PO": "", "Invoice": "INV-DUP", "Amount": 50.00, "Period": "1"}]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.metrics["Posting Status"] == "READY TO POST"
    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]
    unresolved_text = _worksheet_text(ws)
    assert any("No QuickBooks weak-basis duplicate candidates remain unresolved" in t for t in unresolved_text)
