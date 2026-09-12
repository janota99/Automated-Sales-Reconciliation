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

from matching import build_reconciliation
from workpapers import build_analytics_workbook, build_primary_workbook

EXPECTED_PRIMARY_SHEETS = [
    "Raw Data",
    "Reconciled Data",
    "Unresolved Exceptions",
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
    assert result.metrics["Duplicate QuickBooks Rows"] == 2
    assert result.metrics["Duplicate Infinium Rows"] == 2

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


def test_fiscal_period_summary_and_duplicates_are_promoted_above_detail_tables(
    qb_mapping, inf_mapping, make_metadata,
):
    """Both the fiscal-period breakdown and the duplicates block must sit
    above the raw exception/duplicate detail tables, not buried below them."""
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
    duplicates_row = first_row_containing("QUICKBOOKS DUPLICATES EXCLUDED FROM JE")
    exception_table_header_row = first_row_containing("Exception Status")

    assert fiscal_row < exception_table_header_row
    assert duplicates_row < exception_table_header_row
    # Both promoted sections must appear near the very top of the sheet.
    assert fiscal_row <= 10
    assert duplicates_row <= 5


def test_analytics_workbook_builds_with_duplicates_present(qb_mapping, inf_mapping, make_metadata):
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)

    workbook_bytes = build_analytics_workbook(result)
    assert len(workbook_bytes) > 0

    wb = load_workbook(io.BytesIO(workbook_bytes))
    assert wb.sheetnames == EXPECTED_ANALYTICS_SHEETS

    qb_dup_text = _worksheet_text(wb["QuickBooks Duplicates"])
    assert any("QB-000002" in text or "QB-000003" in text for text in qb_dup_text)

    inf_dup_text = _worksheet_text(wb["Infinium Duplicates"])
    assert any("INF-000002" in text or "INF-000003" in text for text in inf_dup_text)


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


def test_render_result_duplicate_badge_never_raises(qb_mapping, inf_mapping, make_metadata):
    """Direct regression test for the exact KeyError that reached Streamlit:
    result.duplicate_analysis["Reconciliation Status"] no longer exists, so
    the attention-badge count must be derived from metrics instead."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    metrics = result.metrics
    unresolved_duplicate_groups = int(metrics.get("Duplicate QuickBooks Rows", 0)) + int(
        metrics.get("Duplicate Infinium Rows", 0)
    )
    assert unresolved_duplicate_groups == 4
    assert "Reconciliation Status" not in result.duplicate_analysis.columns
