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
from typing import Any

import pandas as pd
import pytest
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from matching import QB_ID, build_reconciliation
from workpapers import (
    _legacy_matched_label,
    build_analytics_workbook,
    build_data_search_dataframe,
    build_legacy_workbook,
    build_primary_workbook,
)

EXPECTED_PRIMARY_SHEETS = [
    "Data Search",
    "Data Search QB Source",
    "Data Search INF Source",
    "Raw Data",
    "Reconciled Data",
    "Unresolved Exceptions",
    "Product Aggregate Summary",
]

EXPECTED_LEGACY_SHEETS = [
    "Legacy Reconciliation",
    "Exceptions",
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


def _fill_matches(cell, hex_color: str) -> bool:
    rgb = cell.fill.fgColor.rgb
    return isinstance(rgb, str) and rgb.endswith(hex_color)


_LEGACY_MATCHED_SECTIONS_FOR_TEST = {
    "01 Matched", "01 Matched - Historical Clearance", "09 Fuzzy Match Review Hold",
}


def test_legacy_workbook_builds_with_duplicates_present(qb_mapping, inf_mapping, make_metadata):
    from config import LEGACY_EXCLUDED_FILL, LEGACY_MATCHED_FILL, LEGACY_METHOD_FILL, LEGACY_REVIEW_FILL

    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    # Sanity check the fixture actually has an Infinium-only exception (the
    # PO400/INV400 duplicate pair, which has no QuickBooks counterpart) --
    # otherwise the QB-only filter below would pass vacuously.
    assert result.metrics["Duplicate Infinium Rows"] == 1
    infinium_only_exceptions = [
        r for r in result.paired_rows
        if r.get("Section") not in _LEGACY_MATCHED_SECTIONS_FOR_TEST and r.get("QB Index") is None
    ]
    assert len(infinium_only_exceptions) >= 1

    workbook_bytes = build_legacy_workbook(result)
    assert len(workbook_bytes) > 0

    wb = load_workbook(io.BytesIO(workbook_bytes))
    assert wb.sheetnames == EXPECTED_LEGACY_SHEETS

    recon_ws = wb["Legacy Reconciliation"]
    recon_text = _worksheet_text(recon_ws)
    assert any("Match Method" in text for text in recon_text)
    # The one clean 1:1 match (PO100/INV100/$100) must appear, in a pale
    # green tint, with its method labeled by how it matched -- no
    # confidence tier attached.
    assert any(text == "Unique Match: PO + Invoice + Amount" for text in recon_text)
    # Row 3 is the color legend; row 4 is the header row (unchanged headers).
    header_cells = {cell.value: cell.column for cell in next(recon_ws.iter_rows(min_row=4, max_row=4))}
    method_col = header_cells["Match Method"]
    data_row_cells = list(recon_ws.iter_rows(min_row=5, max_row=5))[0]
    non_method_cells = [c for c in data_row_cells if c.column != method_col]
    assert any(_fill_matches(cell, LEGACY_MATCHED_FILL) for cell in non_method_cells)
    # Body text is regular-weight black -- never bold, never colored.
    for cell in non_method_cells:
        if cell.value is not None:
            assert cell.font.bold is not True
            assert cell.font.color.rgb.endswith("000000")
    # The Match Method data cell is RGB(234,234,234); a clean match isn't bolded.
    method_cell = next(c for c in data_row_cells if c.column == method_col)
    assert _fill_matches(method_cell, LEGACY_METHOD_FILL)
    assert method_cell.font.bold is not True
    # Row 2 (the caption band) stays fixed at height 30; row 1 is pinned to 27.
    assert recon_ws.row_dimensions[2].height == 30
    assert recon_ws.row_dimensions[1].height == 27

    # Every QB row (matched or not) is on Legacy Reconciliation, plus the
    # qualifying Infinium-only rows (any-period duplicate, current-period
    # unmatched) -- not just confirmed matches.
    qb_headers = list(result.qb_raw.columns)
    n_qb = len(qb_headers)
    qualifying_inf_only = [
        r for r in result.paired_rows
        if r.get("QB Index") is None and r.get("Infinium Index") is not None
        and r.get("Section") in {"03 Unmatched Infinium", "05 Duplicate Infinium"}
    ]
    all_qb_rows = [r for r in result.paired_rows if r.get("QB Index") is not None]
    expected_recon_row_count = len(all_qb_rows) + len(qualifying_inf_only)
    total_row_number = next(
        row[0].row for row in recon_ws.iter_rows(min_row=5)
        if any(str(cell.value or "").endswith("TOTAL") for cell in row)
    )
    recon_data_rows = sum(
        1 for row in recon_ws.iter_rows(min_row=5, max_row=total_row_number - 1)
        if any(cell.value is not None for cell in row)
    )
    assert recon_data_rows == expected_recon_row_count

    exceptions_ws = wb["Exceptions"]
    exceptions_text = _worksheet_text(exceptions_ws)
    assert any("Fiscal Period" in text for text in exceptions_text)
    assert any("Exception Type" in text for text in exceptions_text)
    assert "QUICKBOOKS SIDE" in " ".join(exceptions_text).upper()
    fill_colors = {
        cell.fill.fgColor.rgb
        for row in exceptions_ws.iter_rows()
        for cell in row
        if isinstance(cell.fill.fgColor.rgb, str)
    }
    assert any(color.endswith(LEGACY_REVIEW_FILL) for color in fill_colors)
    assert any(color.endswith(LEGACY_EXCLUDED_FILL) for color in fill_colors)
    assert exceptions_ws.row_dimensions[1].height == 27

    # Two independent blocks: general QuickBooks exceptions on the left,
    # excluded QuickBooks duplicate copies on the right, separated by a
    # blank column -- no Infinium columns anywhere on this sheet.
    header_row_cells = next(
        row for row in exceptions_ws.iter_rows()
        if any(cell.value == "Exception Type" for cell in row)
    )
    header_row_number = header_row_cells[0].row
    exception_type_cols = [cell.column for cell in header_row_cells if cell.value == "Exception Type"]
    assert len(exception_type_cols) == 2  # one per block
    left_col, right_col = sorted(exception_type_cols)
    # The right block starts n_qb + 3 (trailer) + 1 (separator) columns
    # after the left block's own "Exception Type" column.
    assert right_col - left_col == n_qb + 3 + 1

    right_qb_start_col = right_col - 1 - n_qb  # "Exception Type" is trailer offset +1
    general_pos = [
        row[0].value for row in exceptions_ws.iter_rows(
            min_row=header_row_number + 1, min_col=1, max_col=1,
        ) if row[0].value is not None
    ]
    duplicate_pos = [
        row[0].value for row in exceptions_ws.iter_rows(
            min_row=header_row_number + 1, min_col=right_qb_start_col, max_col=right_qb_start_col,
        ) if row[0].value is not None
    ]
    expected_general = [
        r for r in result.paired_rows
        if r.get("Section") not in _LEGACY_MATCHED_SECTIONS_FOR_TEST
        and r.get("QB Index") is not None
        and r.get("Section") != "04 Duplicate QuickBooks"
    ]
    expected_duplicate = [
        r for r in result.paired_rows if r.get("Section") == "04 Duplicate QuickBooks"
    ]
    assert len(general_pos) == len(expected_general)
    assert len(duplicate_pos) == len(expected_duplicate)

    assert "Product Aggregate Summary" in wb.sheetnames


@pytest.mark.parametrize("engine_result, expected", [
    ("PO + Invoice + Amount", "Unique Match: PO + Invoice + Amount"),
    ("PO + Amount", "Unique Match: PO + Amount"),
    ("Invoice + Amount", "Unique Match: Invoice + Amount"),
    ("PO + Invoice + Aggregate Amount (Grouped) [group-level; no line allocation]", "Group Match: PO + Invoice + Amount"),
    ("PO + Aggregate Amount (Grouped) [group-level; no line allocation]", "Group Match: PO + Amount"),
    ("Invoice + Aggregate Amount (Grouped) [group-level; no line allocation]", "Group Match: Invoice + Amount"),
    # A confirmed vendor alias is a PO-field + amount match.
    ("Confirmed Vendor Alias + Amount", "Unique Match: PO + Amount"),
    # A prior-period clearance keeps only the underlying rule, not its prefix.
    ("QuickBooks primary ↔ Infinium prior-period match | Invoice + Amount", "Unique Match: Invoice + Amount"),
])
def test_legacy_match_method_labels_state_only_how_the_match_was_made(engine_result, expected):
    label = _legacy_matched_label(engine_result)
    assert label == expected
    for confidence in ("Strong", "Moderate", "Weak", "Review", "Confirmed"):
        assert confidence not in label


def test_legacy_reconciliation_styling_labels_and_legend(qb_mapping, inf_mapping, make_metadata):
    from config import (
        LEGACY_EXCLUDED_FILL, LEGACY_MATCHED_FILL, LEGACY_METHOD_FILL,
        LEGACY_NO_PAIR_FILL, LEGACY_REVIEW_FILL,
    )

    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    ws = load_workbook(io.BytesIO(build_legacy_workbook(result)))["Legacy Reconciliation"]

    # QuickBooks and Infinium share header names in this fixture; setdefault
    # keeps the first (QuickBooks-side) column for each.
    headers: dict = {}
    for cell in ws[4]:
        headers.setdefault(cell.value, cell.column)
    method_col = headers["Match Method"]
    amount_col = headers["Amount"]
    po_col = headers["PO"]
    inf_start = method_col + 1

    rows = list(ws.iter_rows(min_row=5))
    by_label = {}
    for row in rows:
        label = row[method_col - 1].value
        if label and not str(label).endswith("TOTAL"):
            by_label.setdefault(label, []).append(row)

    # Only "how it matched" / why not -- never a confidence tier.
    for label in by_label:
        assert not any(tag in label for tag in ("(Strong)", "(Moderate)", "(Weak)", "(Review)", "(Needs Review)"))

    # Matched: pale green, regular black text, method cell not bold.
    matched = by_label["Unique Match: PO + Invoice + Amount"][0]
    assert _fill_matches(matched[po_col - 1], LEGACY_MATCHED_FILL)
    assert _fill_matches(matched[inf_start - 1], LEGACY_MATCHED_FILL)
    assert _fill_matches(matched[method_col - 1], LEGACY_METHOD_FILL)
    assert matched[method_col - 1].font.bold is not True
    # Descriptive data stays Segoe UI regular; amounts are regular Consolas.
    assert matched[po_col - 1].font.name == "Segoe UI" and matched[po_col - 1].font.bold is not True
    assert matched[amount_col - 1].font.name == "Consolas" and matched[amount_col - 1].font.bold is not True

    # Unmatched QuickBooks row: gold review tint, bold status text, and a
    # gray (not gold) blank Infinium side because nothing is paired there.
    unmatched = by_label["No Matching Infinium records"][0]
    assert _fill_matches(unmatched[po_col - 1], LEGACY_REVIEW_FILL)
    assert unmatched[method_col - 1].font.bold is True
    assert _fill_matches(unmatched[method_col - 1], LEGACY_METHOD_FILL)
    assert _fill_matches(unmatched[inf_start - 1], LEGACY_NO_PAIR_FILL)

    # Excluded duplicate copy: pale red, informational (not bolded).
    duplicate = by_label["Duplicate: Excess Copy Excluded"][0]
    assert _fill_matches(duplicate[po_col - 1], LEGACY_EXCLUDED_FILL)
    assert duplicate[method_col - 1].font.bold is not True

    # Compact legend directly under the introductory note.
    legend = [cell.value for cell in ws[3] if cell.value]
    assert legend == ["Reconciled", "Review", "No paired record"]

    # Infinium-only row (no QuickBooks record): its blank QuickBooks side is
    # the same near-white gray, while the Infinium side keeps the status tint.
    inf_only = by_label["No Matching QuickBooks records"][0]
    assert _fill_matches(inf_only[po_col - 1], LEGACY_NO_PAIR_FILL)
    assert _fill_matches(inf_only[inf_start - 1], LEGACY_REVIEW_FILL) or _fill_matches(
        inf_only[inf_start - 1], LEGACY_EXCLUDED_FILL
    )
    assert ws.print_title_rows == "$1:$4"


def test_legacy_reconciliation_normalizes_infinium_column_names(qb_mapping, make_metadata):
    qb_rows = [{"PO": "PO100", "Invoice": "INV100", "Amount": 100.00, "Qty": 1, "Period": "1"}]
    inf_rows = [{
        "OHAPD": "1", "OHOBDE": "1/09/2026", "OHCO": "WP", "CUNO": 50001, "OHOBNO": "INV100",
        "OHTOTA": 100.00, "OHDESC": "PO 100", "OHPONO": "PO100",
    }]
    inf_mapping = {"po": "OHPONO", "invoice": "OHOBNO", "amount": "OHTOTA", "period": "OHAPD"}
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    ws = load_workbook(io.BytesIO(build_legacy_workbook(result)))["Legacy Reconciliation"]
    header_values = [cell.value for cell in ws[4]]
    for raw in ("OHAPD", "OHOBDE", "OHCO", "CUNO", "OHOBNO", "OHTOTA", "OHDESC", "OHPONO"):
        assert raw not in header_values
    inf_headers = header_values[header_values.index("Match Method") + 1:]
    assert inf_headers[:8] == [
        "Period", "Date", "Type", "Customer No", "Invoice No", "Amount", "Description", "PO No.",
    ]
    # The underlying values still come through under the friendly names.
    data_values = [cell.value for cell in ws[5]]
    assert 50001 in data_values and "INV100" in data_values


def test_legacy_exceptions_bolds_exception_type_only_for_open_items(qb_mapping, inf_mapping, make_metadata):
    from config import LEGACY_EXCLUDED_FILL, LEGACY_REVIEW_FILL

    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    ws = load_workbook(io.BytesIO(build_legacy_workbook(result)))["Exceptions"]
    header_cells = next(row for row in ws.iter_rows() if any(c.value == "Exception Type" for c in row))
    header_row = header_cells[0].row
    left_col, right_col = sorted(c.column for c in header_cells if c.value == "Exception Type")

    open_item = ws.cell(header_row + 1, left_col)
    assert open_item.value is not None
    assert open_item.font.bold is True and _fill_matches(open_item, LEGACY_REVIEW_FILL)
    duplicate_copy = ws.cell(header_row + 1, right_col)
    assert duplicate_copy.value is not None
    assert duplicate_copy.font.bold is not True and _fill_matches(duplicate_copy, LEGACY_EXCLUDED_FILL)
    # Other body cells in an exception row are regular-weight black.
    body_cell = ws.cell(header_row + 1, 1)
    assert body_cell.font.bold is not True and body_cell.font.color.rgb.endswith("000000")


def test_legacy_workbook_builds_with_no_duplicates(qb_mapping, inf_mapping, make_metadata):
    result = _build_result_without_duplicates(qb_mapping, inf_mapping, make_metadata)
    workbook_bytes = build_legacy_workbook(result)
    assert len(workbook_bytes) > 0
    wb = load_workbook(io.BytesIO(workbook_bytes))
    assert wb.sheetnames == EXPECTED_LEGACY_SHEETS


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


def test_data_search_sheet_is_first_and_live_searchable(qb_mapping, inf_mapping, make_metadata):
    """Data Search must be the first sheet a reviewer sees, back its two
    live-search panels with hidden source sheets (not visible clutter), and
    drive its results with a FILTER() formula keyed off the two input cells
    -- not a static table requiring manual filtering."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    wb = load_workbook(io.BytesIO(build_primary_workbook(result)))

    assert wb.sheetnames[0] == "Data Search"
    assert wb["Data Search QB Source"].sheet_state == "hidden"
    assert wb["Data Search INF Source"].sheet_state == "hidden"

    ws = wb["Data Search"]
    search_text = _worksheet_text(ws)
    assert "Enter Invoice #" in search_text
    assert "Enter PO #" in search_text
    assert "QUICKBOOKS ITEMS" in search_text
    assert "INFINIUM ITEMS" in search_text

    qb_formula = ws.cell(9, 1).value
    assert isinstance(qb_formula, str) and qb_formula.startswith("=")
    assert "FILTER(" in qb_formula
    # FILTER is a post-2016 dynamic-array "future function" -- Excel always
    # stores it internally with an _xlfn. prefix, and a file written without
    # that prefix (openpyxl's default) is exactly what triggers Excel's "we
    # found a problem with some content" repair prompt on open.
    assert "_xlfn.FILTER(" in qb_formula
    assert "Data Search QB Source" in qb_formula
    assert "SEARCH(" in qb_formula and "ISNUMBER(" in qb_formula
    # No stray unbalanced parens -- a direct regression guard for the class
    # of hand-built-formula bug this codebase has hit before (#NAME?/#REF!
    # errors from malformed table/range references).
    assert qb_formula.count("(") == qb_formula.count(")")

    inf_col_count = 9  # Row ID, PO, Invoice, Amount, Status, Match Type, Duplicate Group ID, Matched * Row ID, Detail
    inf_formula = ws.cell(9, inf_col_count + 2).value
    assert isinstance(inf_formula, str) and inf_formula.startswith("=")
    assert "Data Search INF Source" in inf_formula
    assert inf_formula.count("(") == inf_formula.count(")")


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


def test_unresolved_sheet_has_a_status_color_legend(qb_mapping, inf_mapping, make_metadata):
    """A reviewer opening this workpaper cold has no other way to learn what
    each status fill means, so a color key must be present, sit between the
    KPI cards and the fiscal-period section, and cover the sheet's actual
    status colors (duplicate red, urgent red, pending amber, current
    green)."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]

    def first_row_containing(needle: str) -> int:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and needle in str(cell.value):
                    return cell.row
        raise AssertionError(f"{needle!r} not found in sheet")

    legend_row = first_row_containing("Confirmed duplicate")
    fiscal_row = first_row_containing("QUICKBOOKS EXCEPTIONS BY FISCAL PERIOD")
    kpi_row = first_row_containing("Proposed JE support total")

    assert kpi_row < legend_row < fiscal_row
    legend_texts = [str(cell.value) for cell in ws[legend_row] if cell.value]
    assert any("duplicate" in text.lower() for text in legend_texts)
    assert any("urgent" in text.lower() for text in legend_texts)


def test_reference_matched_amount_variance_is_excluded_from_accrual_and_shown_separately(
    qb_mapping, inf_mapping, make_metadata,
):
    """A QuickBooks row and an Infinium row that are each other's only
    candidate on a shared PO/invoice, but disagree on amount, must never be
    posted at either amount: not counted as an ordinary unresolved
    QuickBooks exception (and therefore never accrued), but still visible
    to a reviewer in a dedicated "likely data entry error" section rather
    than silently disappearing from the workpaper."""
    qb_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 100.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [
        {"PO": "PO500", "Invoice": "INV500", "Amount": 90.00, "Period": "1"},
    ]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    assert result.unmatched_qb == []
    assert len(result.amount_variance_analysis) == 1
    assert result.metrics["Reference-Matched Amount Variance Rows"] == 1
    assert result.metrics["Unresolved QuickBooks Rows"] == 0

    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]
    sheet_text = _worksheet_text(ws)

    assert any("REFERENCE-MATCHED AMOUNT VARIANCE" in text for text in sheet_text)
    assert any(text == "QB-1" for text in sheet_text)
    assert any(text == "INF-1" for text in sheet_text)
    assert any(isinstance(v, (int, float)) and v == 100.0 for v in [c.value for row in ws.iter_rows() for c in row])
    assert any(isinstance(v, (int, float)) and v == 90.0 for v in [c.value for row in ws.iter_rows() for c in row])
    # It must not also appear as an ordinary exception row/PROPOSED JE total contributor.
    assert not any(text == "Unmatched QuickBooks" for text in sheet_text)


def test_ambiguous_duplicate_is_excluded_from_accrual_and_labeled_separately(
    qb_mapping, inf_mapping, make_metadata,
):
    """A QuickBooks row whose PO/invoice matches more than one still-
    unresolved Infinium row must never be posted as an ordinary exception --
    it is withheld from accrual and must appear, explicitly labeled
    "ambiguous duplicate" (not folded into the confirmed-duplicate or
    amount-variance sections), so a reviewer can research the candidates."""
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
    assert result.amount_variance_analysis.empty

    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]
    sheet_text = _worksheet_text(ws)

    assert any("AMBIGUOUS DUPLICATE" in text for text in sheet_text)
    assert any(text == "QB-1" for text in sheet_text)
    # Both Infinium candidates are listed for the reviewer to research.
    assert any("INF-1" in text and "INF-2" in text for text in sheet_text)
    # It must not also appear as an ordinary exception row, nor be folded
    # into the reference-matched amount variance section (no variance rows
    # were generated at all, per the assertion above).
    assert not any(text == "Unmatched QuickBooks" for text in sheet_text)
    assert not any(str(text).startswith("VAR-") for text in sheet_text)


def test_po_reuse_error_stays_in_accrual_and_shows_grouped_detail(
    qb_mapping, inf_mapping, make_metadata,
):
    """A PO reused across 2+ unresolved QuickBooks rows whose grouped total
    disagrees with Infinium's must remain in the exceptions table and
    accrual total (unlike a duplicate, amount variance, or ambiguous
    duplicate, none of which stay), labeled "PO Re-use Error" at the row
    level, with a supplementary grouped-detail section showing the PO,
    both totals, the difference, and both row counts."""
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
    assert len(result.po_reuse_errors) == 1

    workbook_bytes = build_primary_workbook(result)
    ws = load_workbook(io.BytesIO(workbook_bytes))["Unresolved Exceptions"]
    sheet_text = _worksheet_text(ws)

    # Both rows stay as ordinary exceptions, labeled PO Re-use Error, and
    # are still part of the JE-support total (they are never pulled into a
    # withheld/excluded section like the other review-hold categories).
    assert sheet_text.count("PO Re-use Error") >= 2
    assert any("PO RE-USE ERROR" in text for text in sheet_text)
    assert any("QB-1" in text and "QB-2" in text for text in sheet_text)
    assert any(isinstance(v, (int, float)) and v == 150.0 for v in [c.value for row in ws.iter_rows() for c in row])
    assert any(isinstance(v, (int, float)) and v == 140.0 for v in [c.value for row in ws.iter_rows() for c in row])


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


def test_unresolved_sheet_row_id_links_to_the_correct_reconciled_data_row(
    qb_mapping, inf_mapping, make_metadata,
):
    """Row ID cells in the duplicates and duplicate-review-hold sections
    must link straight to where that same row was originally listed on
    Reconciled Data -- and it must be the CORRECT row, not just any link,
    since a wrong target would be worse than no link at all."""
    result = _build_result_with_duplicates(qb_mapping, inf_mapping, make_metadata)
    wb = load_workbook(io.BytesIO(build_primary_workbook(result)))
    unresolved_ws = wb["Unresolved Exceptions"]
    reconciled_ws = wb["Reconciled Data"]

    def reconciled_po_at(row: int) -> Any:
        return reconciled_ws.cell(row, 1).value  # QB block starts at column A

    qb_po_by_id = result.qb_work.set_index(QB_ID)[qb_mapping["po"]].to_dict()

    def qb_po_for(qb_id: str) -> Any:
        return qb_po_by_id.get(qb_id)

    linked_cells = [
        cell for row in unresolved_ws.iter_rows() for cell in row
        if cell.hyperlink is not None
    ]
    assert linked_cells, "expected at least one Row ID hyperlink on Unresolved Exceptions"

    match_col = next(
        cell.column for cell in reconciled_ws[3] if cell.value == "Match Result"
    )

    checked = 0
    round_trips_checked = 0
    for cell in linked_cells:
        qb_id = str(cell.value)
        target = cell.hyperlink.target
        assert target.startswith("#'Reconciled Data'!A")
        target_row = int(target.rsplit("A", 1)[1])
        assert reconciled_po_at(target_row) == qb_po_for(qb_id), (
            f"Row ID {qb_id} links to Reconciled Data row {target_row}, which doesn't match"
        )
        # Underline signals "clickable" without erasing a meaningful color
        # (e.g. a duplicate-excluded row's Row ID stays red, just underlined),
        # and bold makes every link easy to spot regardless of its color.
        assert cell.font.underline == "single"
        assert cell.font.bold is True
        checked += 1

        # Round trip: Reconciled Data's Match Result cell on that same row
        # must link back to exactly the Unresolved Exceptions row we started
        # from -- bidirectional, not just forward.
        reverse_cell = reconciled_ws.cell(target_row, match_col)
        if reverse_cell.hyperlink is not None:
            reverse_target = reverse_cell.hyperlink.target
            assert reverse_target == f"#'Unresolved Exceptions'!A{cell.row}"
            assert reverse_cell.font.bold is True
            assert reverse_cell.font.underline == "single"
            round_trips_checked += 1
    assert checked >= 2  # both the duplicates and the review-hold section have at least one row here
    assert round_trips_checked >= 2  # every linked row above must round-trip back


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
    """QuickBooks weak-basis duplicates no longer land in Duplicate Review
    Hold under the current policy (see test_full_reconciliation_detects_
    blank_reference_duplicates and test_weak_basis_review_hold_does_not_
    block_control_status in test_matching.py) -- a zero-evidence group is
    now resolved immediately (one canonical exception, the rest excluded)
    rather than deferred to a human. The "Duplicate Review Hold" section on
    Unresolved Exceptions is kept for the structural case of a future
    QuickBooks policy that defers again, so its rendering is tested here
    by patching a real result's duplicate_analysis into that state directly
    rather than by re-deriving it through build_reconciliation, since no
    QuickBooks input can produce it anymore.
    """
    qb_rows = [
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "", "Invoice": "INVBLANK", "Amount": 25.00, "Qty": 1, "Period": "1"},
        {"PO": "PO999", "Invoice": "INV999", "Amount": 15.00, "Qty": 1, "Period": "1"},
    ]
    inf_rows = [{"PO": "POX", "Invoice": "INVX", "Amount": 1.00, "Period": "1"}]
    result = build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026,
    )
    held_qb_ids = set(result.qb_work.loc[[0, 1], QB_ID])
    patched = result.duplicate_analysis.copy()
    held_mask = patched["Source Row ID"].isin(held_qb_ids)
    patched.loc[held_mask, "Disposition"] = "Held for review - excluded from proposed JE pending disposition"
    patched.loc[held_mask, "Confidence"] = "Hold"
    patched.loc[held_mask, "Automatically Excluded"] = True
    patched.loc[held_mask, "Canonical Source Row ID"] = ""
    result.duplicate_analysis = patched
    result.duplicate_review_hold_qb_rows = [0, 1]
    result.duplicate_qb_rows = []
    result.metrics["Duplicate Review Hold QuickBooks Rows"] = 2
    result.metrics["Duplicate Review Hold QuickBooks Amount"] = 50.0
    result.metrics["Duplicate QuickBooks Rows"] = 0
    result.metrics["Duplicate QuickBooks Amount"] = 0.0
    result.metrics["Posting Status"] = "REVIEW REQUIRED"
    return result


def test_duplicate_review_hold_section_renders_with_documented_disposition_dropdown(
    qb_mapping, inf_mapping, make_metadata,
):
    """The Duplicate Review Hold section must appear after the duplicates
    section (before the JE), list the held items, and give the reviewer a
    controlled-vocabulary, editable place to record a documented decision
    -- not just a static, uneditable list. (See _build_result_with_review_
    hold for why this state is patched in rather than produced by a real
    QuickBooks scenario under the current policy.)"""
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
