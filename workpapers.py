"""Excel workpaper construction.

This module is the high-level Excel orchestration layer: it builds the two
downloadable workbooks (the primary accounting workpaper and the optional
analytics evidence package) sheet by sheet. Every function here composes the
styling primitives in ``excel_styles.py`` and pulls its data from a
``ReconciliationResult``. It does not know about Streamlit at all -- the UI
layer (``ui_components.py``) is the only caller that talks to both this
module and the browser.
"""

from __future__ import annotations

import io
import re
import zipfile
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from config import (
    ACCOUNTING_COUNT_FORMAT,
    ACCOUNTING_CURRENCY_FORMAT,
    ACCOUNTING_QUANTITY_FORMAT,
    AMBER,
    CENTRAL_TIMEZONE,
    DATE_NUMBER_FORMAT,
    DUPLICATE_RED_FILL,
    DUPLICATE_RED_TEXT,
    FONT_NAME,
    FONT_NAME_NUMERIC,
    GREEN_LIGHT,
    INFINIUM_FRIENDLY_HEADERS,
    LEGACY_BODY_TEXT,
    LEGACY_EXCLUDED_FILL,
    LEGACY_MATCHED_FILL,
    LEGACY_METHOD_FILL,
    LEGACY_NO_PAIR_FILL,
    LEGACY_REVIEW_FILL,
    METHOD_GREY_DARK,
    METHOD_GREY_FILL,
    NAVY,
    NAVY_LIGHT,
    NEUTRAL_GOLD_FILL,
    NEUTRAL_GOLD_TEXT,
    ORANGE,
    RED_LIGHT,
    SLATE,
    SLATE_LIGHT,
    TEAL,
    TEAL_LIGHT,
    TEXT,
    WHITE,
)
from excel_styles import (
    _apply_default_alignment,
    _apply_duplicate_style,
    _apply_legacy_status_cell,
    _apply_legacy_status_fill,
    _apply_number_formats,
    _autofit_workbook_columns,
    _pin_column_width,
    _format_body_block,
    _format_header,
    _prepare_sheet,
    _set_widths,
    _standardize_column_widths,
    _thin_border,
    _total_border,
    _write_caption_band,
    _write_title_band,
    _write_total_row,
    fix_row_height,
)
from duplicates import (
    DUPLICATE_BASIS_CROSS_SCOPE,
    DUPLICATE_BASIS_INVOICE_ONLY,
    DUPLICATE_BASIS_PO_ONLY,
    DUPLICATE_BASIS_STRICT,
    NORM_INV,
    NORM_PO,
)
from matching import (
    AMOUNT_CENTS,
    INF_ID,
    PRIOR_PERIOD_URGENT_THRESHOLD,
    QB_ID,
    REASON_CODE_GLOSSARY,
    REFERENCE_HOLD_SECTION,
    SHORT_REASON_CODES,
    ReconciliationResult,
    build_fiscal_exception_summary,
    cents_or_zero,
    cents_to_float,
    numeric_quantity_sum,
    numeric_sum,
    describe_match_references,
    parse_fiscal_period,
    po_reuse_error_qb_index_map,
    short_reason_code,
    valid_cents,
    validate_match_references,
)
from utils import excel_safe, format_central_timestamp, format_currency

RECONCILIATION_DETAIL_SHEET = "Reconciliation Detail"
UNRESOLVED_EXCEPTIONS_SHEET = "Unresolved Exceptions"


def _fiscal_period_prefix(result: ReconciliationResult) -> str:
    """"FISCAL PERIOD 05 - 2026": the leading words of every main-sheet title,
    so a reviewer holding two runs of the same workbook open at once knows at
    a glance which period each one belongs to."""
    period = result.metadata.get("fiscal_period")
    year = result.metadata.get("fiscal_year")
    if period is None:
        return f"FISCAL PERIOD NOT SELECTED - {int(year)}" if year else "FISCAL PERIOD NOT SELECTED"
    label = f"FISCAL PERIOD {int(period):02d}"
    return f"{label} - {int(year)}" if year else label


def _write_dataframe_values(ws, frame: pd.DataFrame, start_row: int, start_col: int) -> None:
    for col_offset, header in enumerate(frame.columns):
        ws.cell(start_row, start_col + col_offset, excel_safe(str(header)))
    for row_offset, row in enumerate(frame.values.tolist(), 1):
        for col_offset, value in enumerate(row):
            ws.cell(start_row + row_offset, start_col + col_offset, excel_safe(value))


def _escape_structured_ref_component(text: str) -> str:
    """Escape characters with special meaning inside an Excel structured-table reference."""
    escaped = str(text)
    for char in ("'", "#", "[", "]"):
        escaped = escaped.replace(char, f"'{char}")
    return escaped


def _table_column_reference(table_name: str, column_header: str) -> str:
    """Return a proper qualified structured reference, e.g. ``Table1[Amount]``.

    A bare table name (or ``INDEX(TableName,0,N)`` built from one) is not a
    valid Excel reference when written directly as raw formula text -- only
    Excel's own UI auto-converts a typed table name into this bracketed
    structured-reference form. Writing the bracketed form ourselves is what
    makes formulas outside the table (KPI cards, the proposed JE amount)
    actually resolve instead of showing #NAME?.
    """
    return f"{table_name}[{_escape_structured_ref_component(column_header)}]"


def _table_totals_row_formula(column_header: str) -> str:
    """Return the native Excel table totals-row SUM formula for one column.

    Matches exactly what Excel's own UI writes when a table's Total Row is
    enabled and "Sum" is selected: an *unqualified* single-column reference
    (no table name -- it is implicit from the cell's own position in that
    table's totals row) wrapped in SUBTOTAL so the total also respects any
    filter applied to the table.
    """
    return f"SUBTOTAL(109,[{_escape_structured_ref_component(column_header)}])"


def _review_holds_released_expr() -> str:
    """Review Holds items a reviewer released to the JE (structured
    reference, so it resolves from any sheet, not just Unresolved
    Exceptions -- see the Posting Summary journal-entry bridge)."""
    amount_col = _table_column_reference("ReviewHolds", "Amount")
    disposition_col = _table_column_reference("ReviewHolds", "Reviewer Disposition")
    return f'SUMIFS({amount_col},{disposition_col},"Release to JE")'


def _je_support_manual_exclusions_expr(result: "ReconciliationResult") -> str:
    """True-unmatched JE Support items a reviewer manually excluded, with a
    documented reason -- the engine's own Final Disposition for these rows
    stays TRUE_UNMATCHED; this is a separate, additive manual adjustment."""
    amount_col = _table_column_reference("QuickBooksExceptions", result.qb_mapping["amount"])
    disposition_col = _table_column_reference("QuickBooksExceptions", "Reviewer Disposition")
    return f'SUMIFS({amount_col},{disposition_col},"Exclude")'


def _add_exception_table(
    ws,
    *,
    table_name: str,
    headers: list[str],
    header_row: int,
    total_row: int,
    start_col: int,
    total_label: str,
    summed_headers: set[str],
    style_name: str,
) -> dict[str, int]:
    """Create a filterable exception table with protected, dynamic SUM totals."""
    end_col = start_col + len(headers) - 1
    table = Table(
        displayName=table_name,
        ref=(
            f"{get_column_letter(start_col)}{header_row}:"
            f"{get_column_letter(end_col)}{total_row}"
        ),
        totalsRowCount=1,
        totalsRowShown=True,
    )
    table.tableStyleInfo = TableStyleInfo(
        name=style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    formula_columns: dict[str, int] = {}
    table.tableColumns = []
    for offset, header in enumerate(headers, start=1):
        column = TableColumn(id=offset, name=str(header))
        if offset == 1:
            column.totalsRowLabel = total_label
        if header in summed_headers:
            formula_text = _table_totals_row_formula(header)
            column.totalsRowFunction = "sum"
            formula_columns[header] = start_col + offset - 1
            formula_cell = ws.cell(total_row, start_col + offset - 1)
            formula_cell.value = f"={formula_text}"
            formula_cell.protection = Protection(locked=True)
        table.tableColumns.append(column)

    ws.add_table(table)

    # Users may add, remove, classify, and annotate exception rows. The totals
    # row and every other report formula remain locked by worksheet protection.
    for row in range(header_row + 1, total_row):
        for col in range(start_col, end_col + 1):
            ws.cell(row, col).protection = Protection(locked=False)

    for col in range(start_col, end_col + 1):
        ws.cell(total_row, col).protection = Protection(locked=True)
    return formula_columns


def _source_totals(frame: pd.DataFrame, mapping: dict[str, Optional[str]]) -> dict[str, float]:
    totals: dict[str, float] = {}
    amount_col = mapping.get("amount")
    quantity_col = mapping.get("quantity")
    if amount_col:
        totals[amount_col] = numeric_sum(frame[amount_col])
    if quantity_col:
        totals[quantity_col] = numeric_quantity_sum(frame[quantity_col])
    return totals


def _duplicate_source_indexes(
    result: ReconciliationResult,
    dataset: str,
) -> set[int]:
    """Return duplicate primary-row indexes, including for pre-2.9 session results."""
    attribute = "duplicate_qb_rows" if dataset == "QuickBooks" else "duplicate_inf_rows"
    stored_indexes = getattr(result, attribute, None)
    if stored_indexes is not None:
        return {int(index) for index in stored_indexes}

    analysis = getattr(result, "duplicate_analysis", pd.DataFrame())
    if analysis.empty or not {"Dataset", "Source Row IDs"}.issubset(analysis.columns):
        return set()
    duplicate_ids: set[str] = set()
    source_rows = analysis.loc[analysis["Dataset"].eq(dataset), "Source Row IDs"]
    for value in source_rows.dropna().astype(str):
        duplicate_ids.update(item.strip() for item in value.split(";") if item.strip())
    frame = result.qb_work if dataset == "QuickBooks" else result.inf_work
    id_column = QB_ID if dataset == "QuickBooks" else INF_ID
    return {
        int(index)
        for index in frame.index
        if str(frame.at[index, id_column]) in duplicate_ids
    }


def _resolve_paired_records_bulk(result: ReconciliationResult) -> list[dict[str, Any]]:
    """Bulk upgrade paired-row records using native dicts to prevent O(N) DataFrame lookups."""
    resolved_list = [dict(record) for record in result.paired_rows]

    # Current matching results already contain explicit source scopes. Return
    # them immediately and avoid rebuilding information that is already known.
    if all(
        "QB Record Scope" in record and "Infinium Record Scope" in record
        for record in resolved_list
    ):
        return resolved_list

    # Legacy results may require their scopes and indexes to be reconstructed
    # from historical-clearance evidence. A grouped clearance intentionally
    # repeats its Clearance ID across multiple sequences, so Clearance ID alone
    # is not unique; the sequence is part of the lookup key.
    clearances = getattr(result, "historical_clearances", pd.DataFrame())
    clearance_map: dict[tuple[Any, Any], dict[str, Any]] = {}
    clearance_fallback: dict[Any, dict[str, Any]] = {}
    if not clearances.empty and "Clearance ID" in clearances.columns:
        for clearance in clearances.to_dict("records"):
            clearance_id = clearance.get("Clearance ID")
            sequence = clearance.get("Group Sequence", 1)
            clearance_map[(clearance_id, sequence)] = clearance
            clearance_fallback.setdefault(clearance_id, clearance)

    for resolved in resolved_list:
        if "QB Record Scope" in resolved and "Infinium Record Scope" in resolved:
            continue

        resolved["QB Record Scope"] = "Primary" if resolved.get("QB Index") is not None else None
        resolved["Infinium Record Scope"] = (
            "Primary" if resolved.get("Infinium Index") is not None else None
        )
        if resolved.get("Section") == "01 Matched - Historical Clearance":
            clearance_id = resolved.get("Match ID")
            sequence = resolved.get("Group Sequence", 1)
            clearance = clearance_map.get(
                (clearance_id, sequence), clearance_fallback.get(clearance_id)
            )
            if clearance:
                primary_is_qb = clearance["Primary Dataset"] == "QuickBooks Primary"
                qb_index = (
                    clearance["Primary Row Index"]
                    if primary_is_qb else clearance["Secondary Row Index"]
                )
                inf_index = (
                    clearance["Secondary Row Index"]
                    if primary_is_qb else clearance["Primary Row Index"]
                )
                resolved["QB Index"] = (
                    None if qb_index is None or pd.isna(qb_index) else int(qb_index)
                )
                resolved["Infinium Index"] = (
                    None if inf_index is None or pd.isna(inf_index) else int(inf_index)
                )
                resolved["QB Record Scope"] = (
                    ("Primary" if primary_is_qb else "Historical")
                    if resolved["QB Index"] is not None else None
                )
                resolved["Infinium Record Scope"] = (
                    ("Historical" if primary_is_qb else "Primary")
                    if resolved["Infinium Index"] is not None else None
                )

    return resolved_list


def build_raw_data_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.create_sheet("Raw Data")
    qb_headers = list(result.qb_raw.columns)
    inf_headers = list(result.inf_raw.columns)
    qb_start = 1
    separator_col = len(qb_headers) + 1
    inf_start = separator_col + 1
    header_row = 3
    data_row = 4
    qb_end = len(qb_headers)
    inf_end = inf_start + len(inf_headers) - 1

    _write_title_band(
        ws, 1, qb_start, qb_end, f"{_fiscal_period_prefix(result)} | QUICKBOOKS | RAW TRANSACTION DETAIL", NAVY,
    )
    _write_title_band(ws, 1, inf_start, inf_end, "INFINIUM | RAW UPLOAD", TEAL)
    _write_caption_band(
        ws, 2, qb_start, qb_end,
        f"{len(result.qb_raw):,} rows | Source control total: ${result.metrics['QuickBooks Source Total']:,.2f} | "
        f"{result.metrics['QuickBooks Subtotal Rows Excluded']:,} subtotal row(s) excluded before matching | "
        f"Generated {format_central_timestamp(result.run_timestamp)}",
        NAVY,
    )
    _write_caption_band(
        ws, 2, inf_start, inf_end,
        f"{len(result.inf_raw):,} rows | Source control total: ${result.metrics['Infinium Source Total']:,.2f} | "
        f"Values preserved before matching | Generated {format_central_timestamp(result.run_timestamp)}",
        TEAL,
    )
    _write_dataframe_values(ws, result.qb_raw, header_row, qb_start)
    _write_dataframe_values(ws, result.inf_raw, header_row, inf_start)
    qb_amount_cols = {result.qb_mapping["amount"]}
    qb_quantity_cols = {result.qb_mapping.get("quantity") or ""}
    inf_amount_cols = {result.inf_mapping["amount"]}
    _format_header(ws, header_row, qb_start, qb_end, NAVY, qb_headers, qb_amount_cols, qb_quantity_cols)
    _format_header(ws, header_row, inf_start, inf_end, TEAL, inf_headers, inf_amount_cols)
    _format_body_block(ws, data_row, data_row + len(result.qb_raw) - 1, qb_start, qb_end, NAVY_LIGHT)
    _format_body_block(ws, data_row, data_row + len(result.inf_raw) - 1, inf_start, inf_end, TEAL_LIGHT)
    for source_index in _duplicate_source_indexes(result, "QuickBooks"):
        if 0 <= source_index < len(result.qb_raw):
            _apply_duplicate_style(ws, data_row + source_index, qb_start, qb_end)
    for source_index in _duplicate_source_indexes(result, "Infinium"):
        if 0 <= source_index < len(result.inf_raw):
            _apply_duplicate_style(ws, data_row + source_index, inf_start, inf_end)
    qb_total_row = data_row + len(result.qb_raw)
    inf_total_row = data_row + len(result.inf_raw)
    _write_total_row(ws, qb_total_row, qb_start, qb_end, _source_totals(result.qb_raw, result.qb_mapping), qb_headers, "SOURCE TOTAL")
    _write_total_row(ws, inf_total_row, inf_start, inf_end, _source_totals(result.inf_raw, result.inf_mapping), inf_headers, "SOURCE TOTAL")
    _apply_number_formats(ws, qb_headers, data_row, qb_total_row, qb_start,
                          qb_amount_cols, qb_quantity_cols)
    _apply_number_formats(ws, inf_headers, data_row, inf_total_row, inf_start,
                          inf_amount_cols, set())
    ws.column_dimensions[get_column_letter(separator_col)].width = 3.5
    ws.column_dimensions[get_column_letter(separator_col)].fill = PatternFill("solid", fgColor=WHITE)
    _set_widths(ws, qb_start, qb_end, header_row, qb_total_row)
    _set_widths(ws, inf_start, inf_end, header_row, inf_total_row)
    ws.freeze_panes = f"{get_column_letter(inf_start)}{data_row}"
    ws.print_title_rows = "1:3"
    _prepare_sheet(ws)


ACCEPTED_PRIOR_PERIOD_LABEL = "Accepted Prior-Period Record"


def _paired_display_frames(result: ReconciliationResult) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    def ordered_union(primary_headers: list[str], historical_headers: list[str]) -> list[str]:
        output = list(primary_headers)
        output.extend(header for header in historical_headers if header not in output)
        return output

    def unique_context_header(base: str, headers: list[str]) -> str:
        candidate = base
        suffix = 2
        while candidate in headers:
            candidate = f"{base} {suffix}"
            suffix += 1
        return candidate

    paired_records = _resolve_paired_records_bulk(result)
    qb_historical_used = any(
        record.get("QB Record Scope") == "Historical" for record in paired_records
    )
    inf_historical_used = any(
        record.get("Infinium Record Scope") == "Historical" for record in paired_records
    )
    qb_historical_headers = (
        list(result.qb_secondary_raw.columns)
        if qb_historical_used and result.qb_secondary_raw is not None else []
    )
    inf_historical_headers = (
        list(result.inf_secondary_raw.columns)
        if inf_historical_used and result.inf_secondary_raw is not None else []
    )
    qb_headers = ordered_union(list(result.qb_raw.columns), qb_historical_headers)
    inf_headers = ordered_union(list(result.inf_raw.columns), inf_historical_headers)
    qb_context_header = unique_context_header("QuickBooks Record Context", qb_headers)
    inf_context_header = unique_context_header("Infinium Record Context", inf_headers)

    qb_work_dict = result.qb_work.to_dict("index") if result.qb_work is not None else {}
    inf_work_dict = result.inf_work.to_dict("index") if result.inf_work is not None else {}
    qb_sec_dict = result.qb_secondary_work.to_dict("index") if result.qb_secondary_work is not None else {}
    inf_sec_dict = result.inf_secondary_work.to_dict("index") if result.inf_secondary_work is not None else {}

    def row_values(
        index: Optional[int],
        scope: Optional[str],
        primary_dict: dict,
        historical_dict: dict,
        headers: list[str],
    ) -> list[Any]:
        if index is None:
            return [None] * len(headers)
        source_dict = historical_dict if scope == "Historical" else primary_dict
        row_data = source_dict.get(index, {})
        return [row_data.get(header) for header in headers]

    qb_rows, inf_rows, match_results = [], [], []
    for record in paired_records:
        qidx, iidx = record["QB Index"], record["Infinium Index"]
        qb_scope = record.get("QB Record Scope")
        inf_scope = record.get("Infinium Record Scope")
        qb_values = row_values(
            qidx, qb_scope, qb_work_dict, qb_sec_dict, qb_headers
        )
        inf_values = row_values(
            iidx, inf_scope, inf_work_dict, inf_sec_dict, inf_headers
        )
        # Record Context is only filled for a deviation from the norm -- a
        # record accepted from the prior-period upload -- so an ordinary
        # primary-upload record stays blank and the exceptions stand out.
        qb_values.append(ACCEPTED_PRIOR_PERIOD_LABEL if qb_scope == "Historical" else None)
        inf_values.append(ACCEPTED_PRIOR_PERIOD_LABEL if inf_scope == "Historical" else None)
        qb_rows.append(qb_values)
        inf_rows.append(inf_values)
        match_results.append(record["Match Result"])
    return (
        pd.DataFrame(qb_rows, columns=qb_headers + [qb_context_header]),
        pd.DataFrame(inf_rows, columns=inf_headers + [inf_context_header]),
        match_results,
    )


# Reconciliation Detail writes exactly one row per result.paired_rows entry,
# in that same order, starting at this row (row 3 is the control strip).
# Other sheets (Unresolved Exceptions) rely on this exact constant to compute
# a hyperlink target without re-deriving this sheet's own layout -- keep them
# in sync.
RECONCILIATION_DETAIL_HEADER_ROW = 4
RECONCILIATION_DETAIL_DATA_ROW = 5


def _qb_id_detail_row_map(result: ReconciliationResult) -> dict[str, int]:
    """Map every QuickBooks Row ID to the row it occupies on Reconciliation
    Detail, so another sheet can link straight to where a row was originally
    listed instead of leaving a reader to search for it by hand."""
    mapping: dict[str, int] = {}
    for offset, record in enumerate(result.paired_rows):
        qidx = record.get("QB Index")
        if qidx is None:
            continue
        scope = record.get("QB Record Scope")
        source = (
            result.qb_secondary_work
            if scope == "Historical" and result.qb_secondary_work is not None
            else result.qb_work
        )
        if source is None or qidx not in source.index:
            continue
        qb_id = source.at[qidx, QB_ID]
        mapping.setdefault(str(qb_id), RECONCILIATION_DETAIL_DATA_ROW + offset)
    return mapping


def _apply_row_id_hyperlink(
    ws, row: int, col: int, target_row: Optional[int], target_sheet: str = RECONCILIATION_DETAIL_SHEET,
) -> None:
    """Turn a cell into a link straight to a specific row on another sheet
    -- if a target couldn't be resolved, leave the cell as plain text
    rather than link to nothing.

    Keeps whatever font color is already on the cell rather than
    replacing it outright -- a duplicate-excluded row's Row ID must stay
    visibly red, not turn hyperlink-blue and lose that signal -- but
    always forces bold plus an underline, so a link is easy to spot at a
    glance regardless of which color/status it's sitting on top of.
    """
    if target_row is None:
        return
    cell = ws.cell(row, col)
    cell.hyperlink = f"#'{target_sheet}'!A{target_row}"
    current = cell.font
    cell.font = Font(
        name=current.name or FONT_NAME,
        size=current.size or 10,
        bold=True,
        color=current.color or NAVY,
        underline="single",
    )


def _find_rows_by_cell_value(ws, target_values: set[str]) -> dict[str, int]:
    """Scan an already-built worksheet for the row each of these exact
    cell values landed on -- used to link back to a sheet whose row
    layout (fiscal-period summary length, KPI rows, stacked sections)
    isn't a simple, safely-reusable formula the way Reconciliation Detail's is.
    """
    found: dict[str, int] = {}
    if not target_values:
        return found
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if value is None:
                continue
            text = str(value)
            if text in target_values and text not in found:
                found[text] = cell.row
    return found


_UNRESOLVED_SHEET_LINKABLE_DISPOSITIONS = frozenset({
    "Retained canonical row",
    "Excluded excess copy",
    "Held for review - excluded from proposed JE pending disposition",
})


def _detail_match_ref_letter(wb: Workbook) -> Optional[str]:
    """Column letter of Match Ref. on the already-built Reconciliation Detail
    sheet, or None if that sheet is absent (links then stay plain text)."""
    if RECONCILIATION_DETAIL_SHEET not in wb.sheetnames:
        return None
    for cell in wb[RECONCILIATION_DETAIL_SHEET][RECONCILIATION_DETAIL_HEADER_ROW]:
        if cell.value == "Match Ref.":
            return get_column_letter(cell.column)
    return None


def _link_reconciled_data_to_unresolved_exceptions(wb: Workbook, result: ReconciliationResult) -> None:
    """The reverse direction of the Row ID links on Unresolved Exceptions:
    for every QuickBooks row shown there as a duplicate or a duplicate
    review-hold item, add a matching link on its Reconciliation Detail row back
    to where it was originally listed as an exception -- so a reviewer
    working from either sheet can always jump to the other.
    """
    if result.duplicate_analysis.empty:
        return
    linkable = result.duplicate_analysis.loc[
        result.duplicate_analysis["Disposition"].isin(_UNRESOLVED_SHEET_LINKABLE_DISPOSITIONS)
    ]
    target_qb_ids = set(linkable["Source Row ID"].astype(str))
    if not target_qb_ids:
        return

    unresolved_ws = wb[UNRESOLVED_EXCEPTIONS_SHEET]
    reconciled_ws = wb[RECONCILIATION_DETAIL_SHEET]
    unresolved_row_by_id = _find_rows_by_cell_value(unresolved_ws, target_qb_ids)
    if not unresolved_row_by_id:
        return

    match_col = next(
        (
            cell.column for cell in reconciled_ws[RECONCILIATION_DETAIL_HEADER_ROW]
            if cell.value == "Match Result"
        ),
        None,
    )
    if match_col is None:
        return
    reconciled_row_by_id = _qb_id_detail_row_map(result)
    for qb_id, unresolved_row in unresolved_row_by_id.items():
        reconciled_row = reconciled_row_by_id.get(qb_id)
        if reconciled_row is None:
            continue
        _apply_row_id_hyperlink(
            reconciled_ws, reconciled_row, match_col, unresolved_row,
            target_sheet=UNRESOLVED_EXCEPTIONS_SHEET,
        )


_MATCHED_SECTIONS = frozenset({"01 Matched", "01 Matched - Historical Clearance"})
_QB_DUPLICATE_EXCLUDED_SECTION = "04 Duplicate QuickBooks"


_QB_REVIEW_HOLD_SECTIONS = frozenset({
    "06 Duplicate Review Hold QuickBooks",
    "08 Reference-Matched Amount Variance Review Hold",
    "09 Fuzzy Match Review Hold",
    "10 Ambiguous Duplicate QuickBooks",
    REFERENCE_HOLD_SECTION,
})


def qb_record_outcomes(result: ReconciliationResult) -> dict[str, int]:
    """Every primary QuickBooks record falls into exactly one outcome -- the
    same four final dispositions the disposition ledger records: reconciled
    (an accepted match, including a prior-period clearance), an excluded
    exact duplicate, a review hold (withheld from the journal entry pending a
    decision), or a true unmatched transaction (the only kind that feeds the
    JE). Counted from the same paired rows the sheet lists, so the control
    strip can never disagree with the rows beneath it."""
    records = reconciled = excluded = review_hold = 0
    for row in result.paired_rows:
        if row.get("QB Index") is None or row.get("QB Record Scope") != "Primary":
            continue
        records += 1
        section = row.get("Section")
        if section in _MATCHED_SECTIONS:
            reconciled += 1
        elif section == _QB_DUPLICATE_EXCLUDED_SECTION:
            excluded += 1
        elif section in _QB_REVIEW_HOLD_SECTIONS:
            review_hold += 1
    return {
        "records": records,
        "reconciled": reconciled,
        "excluded": excluded,
        "review_hold": review_hold,
        "unmatched": records - reconciled - excluded - review_hold,
    }


def _control_strip_text(result: ReconciliationResult) -> str:
    """"661 QuickBooks records | 449 reconciled | 67.9% | 30 review hold |
    178 unmatched | 4 duplicates excluded | JE support: $747,822.02" -- the
    run's outcome in one line, above the detail it summarizes."""
    outcomes = qb_record_outcomes(result)
    rate = outcomes["reconciled"] / outcomes["records"] * 100 if outcomes["records"] else 0.0
    noun = "record" if outcomes["records"] == 1 else "records"
    duplicates = "duplicate" if outcomes["excluded"] == 1 else "duplicates"
    return "   |   ".join((
        f"{outcomes['records']:,} QuickBooks {noun}",
        f"{outcomes['reconciled']:,} reconciled",
        f"{rate:.1f}%",
        f"{outcomes['review_hold']:,} review hold",
        f"{outcomes['unmatched']:,} unmatched",
        f"{outcomes['excluded']:,} {duplicates} excluded",
        f"Engine JE support: {format_currency(result.metrics['Unresolved QuickBooks Amount'])}",
    ))


def _match_ref_link_formula(
    text: str, target_ref_column: str, target_sheet: str = RECONCILIATION_DETAIL_SHEET,
) -> str:
    """A live link from a Referenced Match Ref. cell to the accepted match it
    names on Reconciliation Detail. The row is looked up by reference at
    open time (MATCH over the Match Ref. column), not written as a fixed
    address, so the link still lands on the right relationship after
    Reconciliation Detail is sorted or filtered. A cell naming several
    matches shows all of them and links to the first. Falls back to the
    plain text if the reference cannot be found."""
    shown = str(text).replace('"', '""')
    first = str(text).split(";")[0].strip().replace('"', '""')
    column = target_ref_column
    return (
        f'=IFERROR(HYPERLINK("#\'{target_sheet}\'!{column}"&MATCH("{first}",'
        f"'{target_sheet}'!${column}:${column},0),"
        f'"{shown}"),"{shown}")'
    )


def _style_match_ref_link(cell, text: str, target_ref_column: str) -> None:
    cell.value = _match_ref_link_formula(text, target_ref_column)
    cell.font = Font(name=FONT_NAME, size=10, bold=True, underline="single", color=NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _friendly_infinium_headers(ws, header_row: int, start_col: int, raw_headers: list[str]) -> None:
    """Show Infinium's plain field names (Period, Customer No., Amount ...) in
    place of its system codes. The original code stays one hover away in a
    cell comment; the Raw Data sheet keeps every code exactly as uploaded."""
    for offset, raw_header in enumerate(raw_headers):
        friendly = INFINIUM_FRIENDLY_HEADERS.get(str(raw_header).strip().upper())
        if not friendly:
            continue
        cell = ws.cell(header_row, start_col + offset)
        cell.value = friendly
        comment = Comment(f"Infinium field code: {raw_header}", "Sales Reconciliation")
        comment.width, comment.height = 190, 44
        cell.comment = comment


def build_reconciliation_detail_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.create_sheet(RECONCILIATION_DETAIL_SHEET)
    qb_display, inf_display, match_results = _paired_display_frames(result)
    qb_headers = list(qb_display.columns)
    inf_headers = list(inf_display.columns)
    qb_start = 1
    # Match panel: Match Ref. sits immediately before Match Result; the
    # Referenced Match Ref. pointer (an exception naming a match that already
    # consumed a record) follows it.
    ref_col = len(qb_headers) + 1
    match_col = ref_col + 1
    referenced_col = match_col + 1
    inf_start = referenced_col + 1
    qb_end = len(qb_headers)
    inf_end = inf_start + len(inf_headers) - 1
    strip_row = RECONCILIATION_DETAIL_HEADER_ROW - 1
    header_row, data_row = RECONCILIATION_DETAIL_HEADER_ROW, RECONCILIATION_DETAIL_DATA_ROW
    final_data_row = data_row + len(match_results) - 1
    ref_letter = get_column_letter(ref_col)

    _write_title_band(
        ws, 1, qb_start, qb_end, f"{_fiscal_period_prefix(result)} | QUICKBOOKS | RECONCILIATION DETAIL", NAVY,
    )
    _write_title_band(ws, 1, ref_col, referenced_col, "MATCH RESULT", METHOD_GREY_DARK)
    _write_title_band(ws, 1, inf_start, inf_end, "INFINIUM | RECONCILIATION DETAIL", TEAL)
    _write_caption_band(
        ws, 2, qb_start, qb_end,
        f"Every primary QuickBooks record appears once, reconciled or not. Any accepted QuickBooks prior-period match is displayed on this side and labeled in Record Context. Generated {format_central_timestamp(result.run_timestamp)}.",
        NAVY,
    )
    _write_caption_band(ws, 2, ref_col, referenced_col, "Matching Methodology", METHOD_GREY_DARK)
    _write_caption_band(
        ws, 2, inf_start, inf_end,
        "Every primary Infinium record appears once. Accepted prior-period matches are displayed; unused historical rows are excluded.",
        TEAL,
    )

    # Control strip: the run's outcome in one line, plus the control result.
    strip_cell = ws.cell(strip_row, qb_start, _control_strip_text(result))
    ws.merge_cells(start_row=strip_row, start_column=qb_start, end_row=strip_row, end_column=match_col)
    for col in range(qb_start, match_col + 1):
        ws.cell(strip_row, col).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
        ws.cell(strip_row, col).border = _thin_border()
    strip_cell.font = Font(name=FONT_NAME, size=10, bold=True, color=NAVY)
    strip_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1, shrink_to_fit=True)
    control_passed = result.metrics["Control Status"] == "PASS"
    control_cell = ws.cell(strip_row, referenced_col, f"Control: {result.metrics['Control Status']}")
    control_cell.fill = PatternFill("solid", fgColor=GREEN_LIGHT if control_passed else RED_LIGHT)
    control_cell.font = Font(name=FONT_NAME, size=10, bold=True, color=TEXT)
    control_cell.border = _thin_border()
    control_cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    fix_row_height(ws, strip_row, 22)

    _write_dataframe_values(ws, qb_display, header_row, qb_start)
    ws.cell(header_row, ref_col, "Match Ref.")
    ws.cell(header_row, match_col, "Match Result")
    ws.cell(header_row, referenced_col, "Referenced Match Ref.")
    resolved_records = _resolve_paired_records_bulk(result)
    for offset, (value, paired) in enumerate(zip(match_results, resolved_records), 1):
        ws.cell(header_row + offset, ref_col, paired.get("Match Ref.") or None)
        ws.cell(header_row + offset, match_col, value)
        ws.cell(header_row + offset, referenced_col, paired.get("Referenced Match Ref.") or None)
    _write_dataframe_values(ws, inf_display, header_row, inf_start)
    qb_amount_cols = {result.qb_mapping["amount"]}
    qb_quantity_cols = {result.qb_mapping.get("quantity") or ""}
    inf_amount_cols = {result.inf_mapping["amount"]}
    # Alignment and number formats key off the original headers (the amount
    # column is literally "OHTOTA"); the friendly names are painted on last.
    _format_header(ws, header_row, qb_start, qb_end, NAVY, qb_headers, qb_amount_cols, qb_quantity_cols)
    _format_header(ws, header_row, ref_col, referenced_col, METHOD_GREY_DARK)
    _format_header(ws, header_row, inf_start, inf_end, TEAL, inf_headers, inf_amount_cols)
    for code_col in (ref_col, referenced_col):
        ws.cell(header_row, code_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _format_body_block(ws, data_row, final_data_row, qb_start, qb_end, NAVY_LIGHT)
    _format_body_block(ws, data_row, final_data_row, ref_col, referenced_col, METHOD_GREY_FILL)
    _format_body_block(ws, data_row, final_data_row, inf_start, inf_end, TEAL_LIGHT)

    duplicate_qb_rows = _duplicate_source_indexes(result, "QuickBooks")
    duplicate_inf_rows = _duplicate_source_indexes(result, "Infinium")
    for offset, record in enumerate(resolved_records):
        row = data_row + offset
        status = str(record["Section"])
        if status == "02 Unmatched QuickBooks":
            for col in range(qb_start, referenced_col + 1):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=AMBER)
        elif status == "03 Unmatched Infinium":
            for col in range(ref_col, inf_end + 1):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=ORANGE)
        qidx = record["QB Index"]
        iidx = record["Infinium Index"]
        if (
            record.get("QB Record Scope") == "Primary"
            and qidx is not None
            and int(qidx) in duplicate_qb_rows
        ):
            _apply_duplicate_style(ws, row, qb_start, qb_end)
        if (
            record.get("Infinium Record Scope") == "Primary"
            and iidx is not None
            and int(iidx) in duplicate_inf_rows
        ):
            _apply_duplicate_style(ws, row, inf_start, inf_end)
        if record.get("QB Record Scope") == "Historical":
            ws.cell(row, qb_end).font = Font(
                name=FONT_NAME, size=10, bold=True, color=NAVY
            )
        if record.get("Infinium Record Scope") == "Historical":
            ws.cell(row, inf_end).font = Font(
                name=FONT_NAME, size=10, bold=True, color=TEAL
            )
        ws.cell(row, match_col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.cell(row, ref_col).alignment = Alignment(horizontal="center", vertical="center")
        pointer = record.get("Referenced Match Ref.")
        if pointer:
            # An exception naming an accepted match links straight to it.
            _style_match_ref_link(ws.cell(row, referenced_col), pointer, ref_letter)
        else:
            ws.cell(row, referenced_col).alignment = Alignment(horizontal="center", vertical="center")

    total_row = final_data_row + 1
    _write_total_row(ws, total_row, qb_start, qb_end,
                     _source_totals(result.qb_raw, result.qb_mapping), qb_headers, "RECONCILIATION TOTAL")
    _write_total_row(ws, total_row, inf_start, inf_end,
                     _source_totals(result.inf_raw, result.inf_mapping), inf_headers, "RECONCILIATION TOTAL")
    ws.cell(total_row, match_col, f"Control: {result.metrics['Control Status']}")
    ws.cell(total_row, match_col).fill = PatternFill("solid", fgColor=GREEN_LIGHT if control_passed else RED_LIGHT)
    ws.cell(total_row, match_col).font = Font(name=FONT_NAME, size=10, bold=True, color=TEXT)
    ws.cell(total_row, match_col).border = _total_border()
    ws.cell(total_row, match_col).alignment = Alignment(horizontal="center", vertical="center")
    _apply_number_formats(ws, qb_headers, data_row, total_row, qb_start,
                          qb_amount_cols, qb_quantity_cols)
    _apply_number_formats(ws, inf_headers, data_row, total_row, inf_start,
                          inf_amount_cols, set())
    _friendly_infinium_headers(ws, header_row, inf_start, inf_headers)
    _set_widths(ws, qb_start, qb_end, header_row, total_row)
    _set_widths(ws, inf_start, inf_end, header_row, total_row)
    ws.column_dimensions[get_column_letter(match_col)].width = 43
    # Wide enough for the full heading and a reference: the workbook-wide
    # autofit must not stretch (or, for the link formulas, misread) them.
    _pin_column_width(ws, ref_letter, 12)
    _pin_column_width(ws, get_column_letter(referenced_col), 22)
    ws.freeze_panes = f"{get_column_letter(inf_start)}{data_row}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(inf_end)}{final_data_row}"
    ws.print_title_rows = f"1:{header_row}"
    _prepare_sheet(ws)


_DUPLICATE_STATUS_LABELS = {
    "Retained canonical row": "Kept (Original)",
    "Excluded excess copy": "Removed (Duplicate)",
    "Held for review - excluded from proposed JE pending disposition": "Pending Review",
}

_DUPLICATE_BASIS_PHRASES = {
    DUPLICATE_BASIS_STRICT: "PO, Invoice, and Amount",
    DUPLICATE_BASIS_PO_ONLY: "PO and Amount",
    DUPLICATE_BASIS_INVOICE_ONLY: "Invoice and Amount",
    DUPLICATE_BASIS_CROSS_SCOPE: "PO, Invoice, and Amount across periods",
}


def _what_was_found(row: dict) -> str:
    basis = row.get("Duplicate Basis", "")
    po = row.get("Normalized PO") or ""
    invoice = row.get("Normalized Invoice") or ""
    amount_str = format_currency(row.get("Amount"))
    if basis == DUPLICATE_BASIS_PO_ONLY:
        shared = f"Same PO {po} and Amount {amount_str} (Invoice blank on both rows)"
    elif basis == DUPLICATE_BASIS_INVOICE_ONLY:
        shared = f"Same Invoice {invoice} and Amount {amount_str} (PO blank on both rows)"
    elif basis == DUPLICATE_BASIS_CROSS_SCOPE:
        shared = f"Same PO {po}, Invoice {invoice}, and Amount {amount_str} as a row in the other period's data"
    else:
        shared = f"Same PO {po}, Invoice {invoice}, and Amount {amount_str}"
    other_ids = [
        piece.strip() for piece in str(row.get("Other Source Row IDs In Group", "")).split(";")
        if piece.strip()
    ]
    if not other_ids:
        as_clause = ""
    elif len(other_ids) == 1:
        as_clause = f" as row {other_ids[0]}"
    elif len(other_ids) <= 3:
        as_clause = f" as rows {', '.join(other_ids)}"
    else:
        as_clause = f" as rows {', '.join(other_ids[:3])}, and {len(other_ids) - 3} more"
    return f"{shared}{as_clause}."


def _duplicate_reason(row: dict) -> str:
    basis_phrase = _DUPLICATE_BASIS_PHRASES.get(row.get("Duplicate Basis", ""), "PO, Invoice, and Amount")
    if row.get("Payload Confirmed"):
        return "These rows appear identical in every field compared."
    differing = str(row.get("Differing Confirmation Fields", "") or "").strip()
    if differing:
        return f"These rows match on {basis_phrase}, but differ in: {differing.replace('; ', ', ')}."
    return f"These rows share the same {basis_phrase}; no other confirmation fields were available to compare."


def _simplify_duplicate_display(frame: pd.DataFrame) -> pd.DataFrame:
    """Reduce a duplicate_analysis-shaped frame to a small, plain-English view
    for non-technical reviewers -- unique ID, what was found, and why.

    This is purely a display transform for the "Unresolved Exceptions" sheet.
    The full technical schema (Screening Stage, Duplicate Basis, Normalized
    PO/Invoice, Confirmed Copy Set ID, etc.) stays intact everywhere else --
    the "QuickBooks/Infinium Duplicates" audit sheets, `finalize_review_
    dispositions`, and `validate_reconciliation` all keep reading the
    original `duplicate_analysis` frame untouched.
    """
    optional_columns = [col for col in ("Reviewer Note", "Reviewer Disposition") if col in frame.columns]
    if frame.empty:
        return pd.DataFrame(columns=["Duplicate ID", "Row ID", "What Was Found", "Reason", "Amount", "Status"] + optional_columns)
    records = frame.to_dict("records")
    simplified = pd.DataFrame({
        "Duplicate ID": frame["Duplicate Group ID"].values,
        "Row ID": frame["Source Row ID"].values,
        "What Was Found": [_what_was_found(row) for row in records],
        "Reason": [_duplicate_reason(row) for row in records],
        "Amount": frame["Amount"].values,
        "Status": frame["Disposition"].map(_DUPLICATE_STATUS_LABELS).fillna(frame["Disposition"]).values,
    })
    for column in optional_columns:
        simplified[column] = frame[column].values
    return simplified


_AMOUNT_VARIANCE_WHAT_MATCHED = {
    "High-likelihood amount variance": "Same PO and Invoice",
    "Critical possible sign reversal": "Same PO and Invoice",
    "Strong invoice-linked amount variance": "Same Invoice only",
    "PO-linked amount variance": "Same PO only",
}


def _amount_variance_reason(row: dict) -> str:
    if row.get("Possible Sign Reversal"):
        return (
            "Same magnitude, opposite sign -- check whether one system recorded this "
            "as a credit and the other as a debit before treating it as a plain typo."
        )
    return (
        "References agree but the dollar amount does not -- most likely a data-entry "
        "error on one side. Verify against source documents; do not accrue either "
        "amount until it's resolved."
    )


def _simplify_amount_variance_display(frame: pd.DataFrame) -> pd.DataFrame:
    """Reduce an amount_variance_analysis-shaped frame (see
    build_reference_amount_variances in matching.py) to a small, plain-
    English view for the "Unresolved Exceptions" sheet -- the full
    technical schema (Reference Evidence, Mutually Unique Reference,
    Posting Disposition, etc.) stays intact on result.amount_variance_analysis
    for the Analytics workbook and validate_reconciliation."""
    columns = [
        "Variance ID", "QuickBooks Row ID", "Infinium Row ID", "What Matched",
        "QuickBooks Amount", "Infinium Amount", "Difference", "Likely Cause", "Reviewer Note",
    ]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    records = frame.to_dict("records")
    simplified = pd.DataFrame({
        "Variance ID": frame["Variance ID"].values,
        "QuickBooks Row ID": frame["QuickBooks Row ID"].values,
        "Infinium Row ID": frame["Infinium Row ID"].values,
        "What Matched": [
            _AMOUNT_VARIANCE_WHAT_MATCHED.get(row.get("Classification"), "Reference agreement")
            for row in records
        ],
        "QuickBooks Amount": frame["QuickBooks Amount"].values,
        "Infinium Amount": frame["Infinium Amount"].values,
        "Difference": frame["Potential Difference"].values,
        "Likely Cause": [_amount_variance_reason(row) for row in records],
    })
    simplified["Reviewer Note"] = ""
    return simplified


def _simplify_ambiguous_duplicate_display(frame: pd.DataFrame) -> pd.DataFrame:
    """Reduce an ambiguous_duplicate_analysis-shaped frame (see
    build_ambiguous_duplicate_candidates in matching.py) to a small, plain-
    English view for the "Unresolved Exceptions" sheet -- the full
    technical schema (Posting Disposition, Manual Decision, etc.) stays
    intact on result.ambiguous_duplicate_analysis for the Analytics
    workbook and validate_reconciliation."""
    columns = [
        "Ambiguous ID", "QuickBooks Row ID", "Candidate Count", "Candidate Infinium Row IDs",
        "Candidate Infinium Amounts", "QuickBooks Amount", "Likely Cause", "Reviewer Note",
    ]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    simplified = pd.DataFrame({
        "Ambiguous ID": frame["Ambiguous ID"].values,
        "QuickBooks Row ID": frame["QuickBooks Row ID"].values,
        "Candidate Count": frame["Candidate Count"].values,
        "Candidate Infinium Row IDs": frame["Candidate Infinium Row IDs"].values,
        "Candidate Infinium Amounts": frame["Candidate Infinium Amounts"].values,
        "QuickBooks Amount": frame["QuickBooks Amount"].values,
        "Likely Cause": frame["Explanation"].values,
    })
    simplified["Reviewer Note"] = ""
    return simplified


def _simplify_reference_hold_display(frame: pd.DataFrame) -> pd.DataFrame:
    """Reduce a reference_hold_analysis-shaped frame (see
    build_reference_evidence_review_holds in matching.py) to the reviewer-
    facing columns: what was held and why, the match it points at, and the
    amounts. The full record stays on result.reference_hold_analysis."""
    columns = [
        "Hold ID", "QuickBooks Row ID", "Why Held", "PO / Invoice", "Referenced Match Ref.",
        "Related Infinium Row IDs", "QuickBooks Amount", "Infinium Amount", "Difference",
        "Reviewer Note",
    ]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    simplified = pd.DataFrame({
        "Hold ID": frame["Hold ID"].values,
        "QuickBooks Row ID": frame["QuickBooks Row ID"].values,
        "Why Held": frame["Classification"].values,
        "PO / Invoice": [
            f"{po or '(blank)'} / {invoice or '(blank)'}"
            for po, invoice in zip(frame["Normalized PO"], frame["Normalized Invoice"])
        ],
        "Referenced Match Ref.": frame["Related Match Ref."].values,
        "Related Infinium Row IDs": frame["Related Infinium Row IDs"].values,
        "QuickBooks Amount": frame["QuickBooks Amount"].values,
        "Infinium Amount": frame["Infinium Amount"].values,
        "Difference": frame["Amount Difference"].values,
    })
    simplified["Reviewer Note"] = ""
    return simplified


def _simplify_po_reuse_display(frame: pd.DataFrame) -> pd.DataFrame:
    """Reduce a po_reuse_errors-shaped frame (see build_po_reuse_errors in
    matching.py) to the reviewer-facing grouped detail columns -- plain-
    English headers, and drops the internal row-index columns used only
    for lookups elsewhere. The full technical schema (Normalized PO,
    QuickBooks/Infinium Row Indexes, etc.) stays intact on
    result.po_reuse_errors for validate_reconciliation and any future
    audit-sheet use."""
    columns = [
        "PO Reuse ID", "PO", "QuickBooks Row IDs", "QuickBooks Row Count",
        "QuickBooks Total", "Infinium Row IDs", "Infinium Row Count", "Infinium Total",
        "Difference", "Explanation",
    ]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    simplified = frame.rename(columns={"Normalized PO": "PO"})
    return simplified.reindex(columns=columns)


def _write_kpi_band(
    ws, label_row: int, value_row: int, kpis: list[tuple], end_col: int, start_col: int = 1,
) -> None:
    """Render a row of KPI cards packed tightly at the left: each card
    merges two columns so its label and value keep a fixed, compact width
    regardless of how wide the data columns beneath happen to be, and
    cards sit directly against each other with no empty gap column between
    them -- unlike spacing cards across every-other column of a wide data
    table, which stretches the band across the full sheet width."""
    col = start_col
    for label, value, number_format in kpis:
        if col + 1 > end_col:
            break
        ws.merge_cells(start_row=label_row, start_column=col, end_row=label_row, end_column=col + 1)
        ws.merge_cells(start_row=value_row, start_column=col, end_row=value_row, end_column=col + 1)
        ws.cell(label_row, col, label)
        ws.cell(value_row, col, value)
        ws.cell(label_row, col).font = Font(name=FONT_NAME, size=9, bold=True, color=SLATE)
        ws.cell(value_row, col).font = Font(name=FONT_NAME, size=12, bold=True, color=NAVY)
        # A card is a fixed, compact block: a long label or value shrinks to
        # fit its card instead of wrapping and stretching the row.
        ws.cell(label_row, col).alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
        ws.cell(value_row, col).alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
        ws.cell(value_row, col).number_format = number_format
        ws.cell(value_row, col).protection = Protection(locked=True)
        for row in (label_row, value_row):
            for c in (col, col + 1):
                ws.cell(row, c).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
                ws.cell(row, c).border = _thin_border()
        col += 2
    fix_row_height(ws, label_row, 15)
    fix_row_height(ws, value_row, 21)


# Status-fill meanings used across this sheet -- a swatch and a short label
# per entry, so a reviewer opening the file cold doesn't need tribal
# knowledge of what each color means. The first three name the period
# classes exactly as the fiscal-period summary does. (fill, text_color_or_None, label)
_STATUS_COLOR_LEGEND: list[tuple] = [
    (GREEN_LIGHT, None, "Current Period"),
    (AMBER, None, "Prior Period / pending review"),
    (RED_LIGHT, None, "Urgent Prior Period"),
    (DUPLICATE_RED_FILL, DUPLICATE_RED_TEXT, "Confirmed duplicate - excluded from JE"),
    (ORANGE, None, "Reference matches, amount differs - likely data entry error"),
    (NEUTRAL_GOLD_FILL, NEUTRAL_GOLD_TEXT, "Ambiguous duplicate - multiple candidates, not accrued"),
]


def _write_color_legend(ws, first_row: int, start_col: int, end_col: int) -> int:
    """A compact color key: a small filled swatch immediately followed by
    its label, packed left to right with no gap between entries -- same
    tight-packing idea as _write_kpi_band, applied to a legend instead of a
    KPI card. Each entry takes four columns (swatch plus a three-column
    label), as many per row as the sheet is wide, wrapping onto a second
    row rather than dropping entries. Returns the last row used."""
    per_row = max(1, (end_col - start_col + 1) // 4)
    last_row = first_row
    for index, (fill_color, text_color, label) in enumerate(_STATUS_COLOR_LEGEND):
        row = first_row + index // per_row
        col = start_col + (index % per_row) * 4
        last_row = row
        swatch = ws.cell(row, col)
        swatch.fill = PatternFill("solid", fgColor=fill_color)
        swatch.border = _thin_border()
        label_end = min(col + 3, end_col)
        if label_end > col + 1:
            ws.merge_cells(start_row=row, start_column=col + 1, end_row=row, end_column=label_end)
        label_cell = ws.cell(row, col + 1, label)
        label_cell.font = Font(name=FONT_NAME, size=8, italic=True, color=text_color or TEXT)
        label_cell.alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
    for row in range(first_row, last_row + 1):
        fix_row_height(ws, row, 15)
    return last_row


def build_unresolved_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.create_sheet(UNRESOLVED_EXCEPTIONS_SHEET)

    # QuickBooks is the sole accrual and journal-entry basis, so this sheet
    # keeps everything with accrual relevance -- the raw QuickBooks exception
    # population on the left, and the QuickBooks items excluded from that
    # same population as duplicates on the right. Infinium exceptions carry
    # no accrual impact and are already listed in Reconciliation Detail, so they
    # are intentionally not repeated here.
    source_headers = list(result.qb_raw.columns)
    headers = source_headers + [
        "Referenced Match Ref.", "Exception Status", "Reference Amount Difference",
        "Reviewer Disposition", "Reviewer", "Review Date", "Comment",
    ]
    referenced_offset = len(source_headers) + 1
    status_offset = referenced_offset + 1
    difference_offset = status_offset + 1
    # This row's Final Disposition is always TRUE_UNMATCHED (the engine's own,
    # immutable conclusion -- see the QB Disposition Ledger). Reviewer
    # Disposition is a SEPARATE manual layer: normally the row simply stays in
    # the automated JE, and only "Exclude" (with a documented reason) pulls
    # it out -- see the Posting Summary bridge (Manual JE Exclusions).
    disposition_offset = difference_offset + 1
    reviewer_offset = disposition_offset + 1
    review_date_offset = reviewer_offset + 1
    note_offset = review_date_offset + 1
    # An exception can point at an accepted match two ways -- a consumed
    # PO/invoice candidate, or a duplicate group with a matched member --
    # both already resolved onto its paired row.
    referenced_by_qb_index = {
        int(row["QB Index"]): row.get("Referenced Match Ref.", "")
        for row in result.paired_rows
        if row.get("Section") == "02 Unmatched QuickBooks" and row.get("QB Index") is not None
    }
    candidate_map = result.candidates.set_index("QuickBooks Row ID").to_dict("index") if not result.candidates.empty else {}
    qb_id_row_map = _qb_id_detail_row_map(result)
    # PO Re-use Error rows are never withheld from the accrual (unlike
    # every other review-hold classification on this sheet), so they still
    # come through result.unmatched_qb in the loop below -- this lookup
    # just overrides their Exception Status label to make the reused-PO
    # grouping traceable instead of showing a generic "no match" reason.
    po_reuse_qb_map = po_reuse_error_qb_index_map(result.po_reuse_errors)
    po_reuse_detail = (
        result.po_reuse_errors.set_index("PO Reuse ID").to_dict("index")
        if not result.po_reuse_errors.empty else {}
    )

    qb_subset_dict = result.qb_work.loc[result.unmatched_qb].to_dict("index")
    records = []
    for qidx in result.unmatched_qb:
        row_data = qb_subset_dict[qidx]
        candidate = candidate_map.get(row_data[QB_ID], {})
        po_reuse_id = po_reuse_qb_map.get(qidx)
        if po_reuse_id:
            exception_status = "PO Re-use Error"
            reference_amount_difference = po_reuse_detail.get(po_reuse_id, {}).get("Difference")
        else:
            exception_status = candidate.get("Disposition", "Unmatched QuickBooks")
            reference_amount_difference = candidate.get("Minimum Amount Difference")
        records.append(
            [row_data.get(col) for col in source_headers]
            + [
                referenced_by_qb_index.get(int(qidx)) or None,
                exception_status,
                reference_amount_difference,
                "Pending Review", "", "", "",
            ]
        )
    frame = pd.DataFrame(records, columns=headers)
    end_col = len(headers)
    amount_col_position = source_headers.index(result.qb_mapping["amount"]) + 1

    # duplicate_analysis carries every disposition tier (canonical, excess,
    # resolved-via-match, and held-for-review) in one report for full audit
    # traceability. Each tier means something different for the JE, so the
    # Excel output splits them: canonical/excess get their own section
    # below, review-hold gets a distinct section further down, and
    # resolved-via-match candidates need no special display at all since
    # they proceeded normally with no exclusion.
    duplicate_frame = result.duplicate_analysis.loc[
        result.duplicate_analysis["Disposition"].isin(
            ["Retained canonical row", "Excluded excess copy"]
        )
    ].copy()
    duplicate_frame["Reviewer Note"] = ""
    # Capture styling inputs from the full technical frame before reducing it
    # to the reviewer-facing view below -- only the excess copy is actually
    # excluded from the JE, and that flag isn't part of the simplified columns.
    duplicate_excluded_flags = duplicate_frame["Automatically Excluded"].fillna(False).astype(bool).tolist()
    duplicate_frame = _simplify_duplicate_display(duplicate_frame)
    duplicate_headers = list(duplicate_frame.columns)
    dup_end_col = len(duplicate_headers)
    duplicate_excluded_count = result.metrics["Duplicate QuickBooks Rows"]
    duplicate_amount_total = result.metrics["Duplicate QuickBooks Amount"]

    # Every QuickBooks row on REVIEW_HOLD -- the same population the
    # disposition ledger, Posting Summary, and Legacy sheet all show -- in ONE
    # table, replacing what used to be four separately-labeled sections
    # (Duplicate Review Hold, Amount Variance, Ambiguous, Reference Evidence).
    # A short Reason Code plus a one-line Reason keep the sheet scannable; the
    # full technical explanation for each code lives on Reason Code Glossary.
    review_ledger = result.qb_dispositions.loc[
        result.qb_dispositions["Final Disposition"] == "REVIEW_HOLD"
    ].copy()
    qb_by_id = result.qb_work.set_index(QB_ID)
    po_field, invoice_field = result.qb_mapping["po"], result.qb_mapping["invoice"]

    def _field(row_id: str, column: Optional[str]) -> Any:
        if not column or row_id not in qb_by_id.index:
            return None
        return qb_by_id.at[row_id, column]

    # Reference-evidence holds and reference-matched amount variances each
    # carry the specific Infinium amount and difference; folded in here so a
    # reviewer sees both sides of an amount conflict without leaving this table.
    reference_amount_frames = [
        result.reference_hold_analysis[["QuickBooks Row ID", "Infinium Amount", "Amount Difference"]]
        if not result.reference_hold_analysis.empty else None,
        result.amount_variance_analysis.rename(
            columns={"Potential Difference": "Amount Difference"},
        )[["QuickBooks Row ID", "Infinium Amount", "Amount Difference"]]
        if not result.amount_variance_analysis.empty else None,
    ]
    reference_amount_frames = [frame for frame in reference_amount_frames if frame is not None]
    reference_amounts = (
        pd.concat(reference_amount_frames, ignore_index=True).set_index("QuickBooks Row ID")
        if reference_amount_frames else pd.DataFrame(columns=["Infinium Amount", "Amount Difference"])
    )
    review_frame = pd.DataFrame({
        "Review ID": review_ledger["Review ID"].values,
        "Row ID": review_ledger["QBO Row ID"].values,
        "PO": [_field(row_id, po_field) for row_id in review_ledger["QBO Row ID"]],
        "Invoice": [_field(row_id, invoice_field) for row_id in review_ledger["QBO Row ID"]],
        "Amount": review_ledger["Amount"].values,
        "Infinium Amount": [
            reference_amounts.at[row_id, "Infinium Amount"] if row_id in reference_amounts.index else None
            for row_id in review_ledger["QBO Row ID"]
        ],
        "Difference": [
            reference_amounts.at[row_id, "Amount Difference"] if row_id in reference_amounts.index else None
            for row_id in review_ledger["QBO Row ID"]
        ],
        "Reason Code": [short_reason_code(code) for code in review_ledger["Reason Code"]],
        "Reason": review_ledger["Final Reason"].values,
        "Referenced Match Ref.": review_ledger["Related Match Ref."].values,
        "Related Infinium Row IDs": review_ledger["Related Infinium Row IDs"].values,
    })
    review_frame["Reviewer Disposition"] = "Pending Review"
    review_frame["Reviewer"] = ""
    review_frame["Review Date"] = ""
    review_frame["Comment"] = ""
    review_headers = list(review_frame.columns)
    review_end_col = len(review_headers)
    review_hold_count = result.metrics["Final Disposition - Review Hold Rows"]
    review_hold_amount = result.metrics["Final Disposition - Review Hold Amount"]

    # PO Re-use Error (see build_po_reuse_errors in matching.py): a
    # normalized PO reused across 2+ still-unresolved QuickBooks rows whose
    # grouped total does not tie exactly to the grouped Infinium total for
    # that PO. Unlike every section above, these rows are NOT withheld --
    # they already appear in the exceptions table above, counted in the
    # accrual total. This is purely the grouped detail view.
    po_reuse_frame = _simplify_po_reuse_display(result.po_reuse_errors)
    po_reuse_headers = list(po_reuse_frame.columns)
    po_reuse_end_col = len(po_reuse_headers)
    po_reuse_group_count = result.metrics["PO Re-use Error Groups"]
    po_reuse_qb_row_count = result.metrics["PO Re-use Error QuickBooks Rows"]
    po_reuse_net_difference = result.metrics["PO Re-use Error Net Difference"]

    unmatched_qb_amounts = result.qb_work.loc[result.unmatched_qb, AMOUNT_CENTS].tolist()
    amounts = [cents_or_zero(val) for val in unmatched_qb_amounts]
    net = sum(amounts)

    qb_table_name = "QuickBooksExceptions"
    qb_amount_column_expr = _table_column_reference(qb_table_name, result.qb_mapping["amount"])
    qb_amount_sum_expr = f"SUM({qb_amount_column_expr})"

    _write_title_band(
        ws, 1, 1, end_col,
        f"{_fiscal_period_prefix(result)} | QUICKBOOKS EXCEPTIONS | JOURNAL ENTRY SUPPORT", NAVY,
    )
    _write_caption_band(
        ws, 2, 1, end_col,
        f"Only TRUE UNMATCHED QuickBooks transactions -- no Infinium support after every matching pass -- feed the "
        f"proposed journal entry. Duplicates and review holds are excluded and itemized below. "
        f"Review every exception before posting. Generated {format_central_timestamp(result.run_timestamp)}.",
        NAVY,
    )
    duplicate_caption = (
        f"{len(duplicate_frame):,} QuickBooks item(s) belong to a strong duplicate group (identical "
        "normalized PO, invoice, and signed amount). One canonical row per group is retained and remains "
        f"active; {duplicate_excluded_count:,} excess "
        f"{'copy is' if duplicate_excluded_count == 1 else 'copies are'} excluded from the accrual/JE "
        "support total in the exceptions table above. If a pair turns out to be a legitimate repeated "
        "transaction rather than a duplicate entry, the excess copy must be added to the JE support "
        "manually."
        if len(duplicate_frame)
        else "No QuickBooks exact duplicates (matching PO, invoice, and amount) were identified."
    )

    kpis = [
        ("Unresolved Rows", f"=IFERROR(ROWS({qb_amount_column_expr}),0)", ACCOUNTING_COUNT_FORMAT),
        (
            "Proposed Debit Support",
            f'=SUMIF({qb_amount_column_expr},">0",{qb_amount_column_expr})',
            ACCOUNTING_CURRENCY_FORMAT,
        ),
        (
            "Proposed Credit Support",
            f'=ABS(SUMIF({qb_amount_column_expr},"<0",{qb_amount_column_expr}))',
            ACCOUNTING_CURRENCY_FORMAT,
        ),
        (
            "Engine Proposed JE Support",
            f"={qb_amount_sum_expr}",
            ACCOUNTING_CURRENCY_FORMAT,
        ),
    ]
    _write_kpi_band(ws, 3, 4, kpis, end_col)
    legend_last_row = _write_color_legend(ws, 5, 1, end_col)
    # One slim spacer row, then the fiscal-period summary.
    fix_row_height(ws, legend_last_row + 1, 8)

    # QuickBooks exceptions by fiscal period -- promoted above the detail
    # tables so period-level review (count and net amount per period) never
    # requires scrolling past the full exception and duplicate lists.
    fiscal_summary = build_fiscal_exception_summary(result)
    fiscal_headers = list(fiscal_summary.columns)
    fiscal_title_row = legend_last_row + 2
    fiscal_caption_row = fiscal_title_row + 1
    fiscal_header_row = fiscal_title_row + 2
    fiscal_data_row = fiscal_header_row + 1
    fiscal_end_col = len(fiscal_headers)
    fiscal_section_end_col = end_col
    selected_period = result.metadata.get("fiscal_period")
    has_fiscal_period = bool(result.qb_mapping.get("period"))
    _write_title_band(
        ws, fiscal_title_row, 1, fiscal_section_end_col,
        (
            "QUICKBOOKS EXCEPTIONS BY FISCAL PERIOD | CURRENT VS PRIOR PERIODS"
            if has_fiscal_period
            else "QUICKBOOKS EXCEPTIONS BY FISCAL PERIOD | FISCAL PERIOD NOT AVAILABLE"
        ),
        NAVY,
    )
    quantity_note = (
        "Exception quantity is sourced from the mapped QuickBooks quantity column."
        if result.qb_mapping.get("quantity")
        else "No QuickBooks quantity column was mapped; exception quantities are shown as zero."
    )
    _write_caption_band(
        ws, fiscal_caption_row, 1, fiscal_section_end_col,
        (
            f"Selected current reporting period: PD-{int(selected_period):02d}. A QuickBooks fiscal period up to "
            f"{PRIOR_PERIOD_URGENT_THRESHOLD} period(s) behind is a Prior Period exception; anything "
            f"older is an Urgent Prior Period exception. {quantity_note}"
            if has_fiscal_period and selected_period is not None
            else f"No current reporting period was selected. Exceptions are summarized by source period without current/prior classification. {quantity_note}"
            if has_fiscal_period
            else "No credible QuickBooks fiscal-period identifier was found or mapped. Period-based "
            f"classification is disabled and all exceptions are summarized together. {quantity_note}"
        ),
        NAVY,
    )
    _write_dataframe_values(ws, fiscal_summary, fiscal_header_row, 1)
    _format_header(
        ws, fiscal_header_row, 1, fiscal_end_col, NAVY,
        headers=fiscal_headers, amount_columns={"Net Exception Amount"},
        quantity_columns={"Exception Count", "Exception Quantity"},
    )
    if len(fiscal_summary):
        fiscal_last_row = fiscal_data_row + len(fiscal_summary) - 1
        _format_body_block(ws, fiscal_data_row, fiscal_last_row, 1, fiscal_end_col, NAVY_LIGHT)
        # Only the classification cell carries the color: tinting the whole
        # row makes the summary compete with the exception table below it.
        classification_col = fiscal_headers.index("Period Classification") + 1
        for offset, classification in enumerate(fiscal_summary["Period Classification"], start=fiscal_data_row):
            fill = (
                RED_LIGHT
                if classification == "Urgent Prior Period"
                else GREEN_LIGHT
                if classification == "Current Period"
                else AMBER
            )
            ws.cell(offset, classification_col).fill = PatternFill("solid", fgColor=fill)
        _apply_number_formats(
            ws, fiscal_headers, fiscal_data_row, fiscal_last_row, 1,
            {"Net Exception Amount"}, {"Exception Count", "Exception Quantity"},
        )
    else:
        fiscal_last_row = fiscal_header_row
    fiscal_total_row = fiscal_last_row + 1
    _write_total_row(
        ws, fiscal_total_row, 1, fiscal_end_col,
        {
            "Exception Count": float(fiscal_summary["Exception Count"].sum()) if len(fiscal_summary) else 0,
            "Exception Quantity": float(fiscal_summary["Exception Quantity"].sum()) if len(fiscal_summary) else 0,
            "Net Exception Amount": float(fiscal_summary["Net Exception Amount"].sum()) if len(fiscal_summary) else 0,
        },
        fiscal_headers,
        "TOTAL EXCEPTIONS",
    )
    _set_widths(ws, 1, fiscal_end_col, fiscal_header_row, fiscal_total_row)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 34)

    header_row = fiscal_total_row + 3
    data_row = header_row + 1
    qb_quantity_header = result.qb_mapping.get("quantity")
    _write_dataframe_values(ws, frame, header_row, 1)
    _format_header(
        ws, header_row, 1, end_col, NAVY,
        headers=headers, amount_columns={result.qb_mapping["amount"]},
        quantity_columns={qb_quantity_header or ""},
    )
    # Centered, like the references beneath it, so it never reads as one run
    # of text with the right-aligned amount heading beside it.
    ws.cell(header_row, referenced_offset).alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True,
    )
    detail_ref_letter = _detail_match_ref_letter(wb)
    if len(frame):
        _format_body_block(ws, data_row, data_row + len(frame) - 1, 1, end_col, NAVY_LIGHT)
        duplicate_qb_rows = _duplicate_source_indexes(result, "QuickBooks")
        for offset, qidx in enumerate(result.unmatched_qb):
            row = data_row + offset
            ws.cell(row, source_headers.index(result.qb_mapping["amount"]) + 1).number_format = ACCOUNTING_CURRENCY_FORMAT
            ws.cell(row, status_offset).fill = PatternFill("solid", fgColor=AMBER)
            ws.cell(row, status_offset).alignment = Alignment(wrap_text=True, vertical="center")
            ws.cell(row, referenced_offset).alignment = Alignment(horizontal="center", vertical="center")
            pointer = referenced_by_qb_index.get(int(qidx))
            if pointer and detail_ref_letter:
                # Jump straight to the accepted match this exception points at.
                _style_match_ref_link(ws.cell(row, referenced_offset), pointer, detail_ref_letter)
            ws.cell(row, difference_offset).number_format = ACCOUNTING_CURRENCY_FORMAT
            if int(qidx) in duplicate_qb_rows:
                _apply_duplicate_style(ws, row, 1, end_col)
    total_row = data_row + len(frame)
    _write_total_row(
        ws, total_row, 1, end_col,
        {result.qb_mapping["amount"]: cents_to_float(net)}, headers,
        "PROPOSED JE SUPPORT TOTAL",
    )
    ws.cell(total_row, amount_col_position).number_format = ACCOUNTING_CURRENCY_FORMAT

    qb_summed_headers = {result.qb_mapping["amount"]}
    if qb_quantity_header and qb_quantity_header in headers:
        qb_summed_headers.add(qb_quantity_header)
    _add_exception_table(
        ws,
        table_name=qb_table_name,
        headers=headers,
        header_row=header_row,
        total_row=total_row,
        start_col=1,
        total_label="PROPOSED JE SUPPORT TOTAL",
        summed_headers=qb_summed_headers,
        style_name="TableStyleMedium2",
    )

    _apply_number_formats(
        ws, headers, data_row, total_row, 1,
        {result.qb_mapping["amount"]}, {qb_quantity_header or ""},
    )
    ws.column_dimensions[get_column_letter(status_offset)].width = 48
    ws.column_dimensions[get_column_letter(difference_offset)].width = 24
    ws.column_dimensions[get_column_letter(disposition_offset)].width = 22
    ws.column_dimensions[get_column_letter(reviewer_offset)].width = 18
    ws.column_dimensions[get_column_letter(review_date_offset)].width = 14
    ws.column_dimensions[get_column_letter(note_offset)].width = 36
    _set_widths(ws, 1, len(source_headers), header_row, total_row)
    ws.freeze_panes = f"A{data_row}"
    if len(frame):
        note_col = get_column_letter(note_offset)
        validation = DataValidation(
            type="textLength", operator="lessThanOrEqual", formula1="1000", allow_blank=True
        )
        validation.error = "Comments are limited to 1,000 characters."
        validation.errorTitle = "Comment too long"
        ws.add_data_validation(validation)
        validation.add(f"{note_col}{data_row}:{note_col}{header_row + len(frame)}")

        je_disposition_validation = DataValidation(
            type="list", formula1='"Pending Review,Exclude"', allow_blank=False,
        )
        je_disposition_validation.error = (
            "Select Pending Review (the automated conclusion stands) or Exclude (documented reason "
            "required) -- the engine's own Final Disposition is never changed by this selection."
        )
        je_disposition_validation.errorTitle = "Reviewer disposition required"
        ws.add_data_validation(je_disposition_validation)
        disposition_col_letter = get_column_letter(disposition_offset)
        je_disposition_validation.add(
            f"{disposition_col_letter}{data_row}:{disposition_col_letter}{header_row + len(frame)}"
        )
        # Excluding a row from the automatic JE is itself an accounting decision,
        # so it must carry a documented reason -- flagged, not silently accepted,
        # if the reviewer picks Exclude and leaves the comment blank.
        je_exclude_rule = FormulaRule(
            formula=[f'AND(${disposition_col_letter}{data_row}="Exclude",${note_col}{data_row}="")'],
            fill=PatternFill("solid", fgColor=RED_LIGHT),
        )
        ws.conditional_formatting.add(f"{note_col}{data_row}:{note_col}{header_row + len(frame)}", je_exclude_rule)

    # QuickBooks duplicates excluded from the JE above, and the proposed JE
    # itself, are placed a fixed 10 rows below the exceptions total row --
    # far enough to read as clearly separate from the exception detail,
    # close enough to stay on the same review pass.
    duplicate_kpis = [
        ("Duplicate QuickBooks items excluded", duplicate_excluded_count, ACCOUNTING_COUNT_FORMAT),
        ("Amount excluded from JE", duplicate_amount_total, ACCOUNTING_CURRENCY_FORMAT),
        ("JE inclusion", "Excluded", 'General'),
    ]
    dup_title_row = total_row + 10
    dup_caption_row = dup_title_row + 1
    dup_kpi_label_row = dup_title_row + 2
    dup_kpi_value_row = dup_title_row + 3
    dup_header_row = dup_title_row + 5
    dup_data_row = dup_header_row + 1
    section_end_col = max(end_col, dup_end_col, review_end_col, po_reuse_end_col)

    _write_title_band(
        ws, dup_title_row, 1, section_end_col,
        "DUPLICATE EXCLUDED | CONFIRMED COPY - EXCLUDED FROM PROPOSED JE", NAVY,
    )
    _write_caption_band(ws, dup_caption_row, 1, section_end_col, duplicate_caption, NAVY)
    _write_kpi_band(ws, dup_kpi_label_row, dup_kpi_value_row, duplicate_kpis, dup_end_col)

    _write_dataframe_values(ws, duplicate_frame, dup_header_row, 1)
    _format_header(
        ws, dup_header_row, 1, dup_end_col, NAVY,
        headers=duplicate_headers, amount_columns={"Amount"},
    )
    if len(duplicate_frame):
        dup_last_row = dup_data_row + len(duplicate_frame) - 1
        _format_body_block(ws, dup_data_row, dup_last_row, 1, dup_end_col, NAVY_LIGHT)
        _apply_number_formats(
            ws, duplicate_headers, dup_data_row, dup_last_row, 1,
            {"Amount"}, set(),
        )
        # Only the excess copy is actually excluded from the JE -- the
        # retained canonical row is shown for audit context but must not be
        # styled as if it, too, had been dropped from the accrual.
        for offset, is_excluded in enumerate(duplicate_excluded_flags):
            if is_excluded:
                _apply_duplicate_style(ws, dup_data_row + offset, 1, dup_end_col)
        # Row ID links straight to where this row was originally listed on
        # Reconciliation Detail, so a reviewer never has to search for it by hand.
        row_id_col = duplicate_headers.index("Row ID") + 1
        for offset, qb_id in enumerate(duplicate_frame["Row ID"]):
            _apply_row_id_hyperlink(ws, dup_data_row + offset, row_id_col, qb_id_row_map.get(str(qb_id)))
        if "Reviewer Note" in duplicate_headers:
            note_col = duplicate_headers.index("Reviewer Note") + 1
            for row in range(dup_data_row, dup_last_row + 1):
                ws.cell(row, note_col).protection = Protection(locked=False)
    else:
        dup_last_row = dup_header_row

    _set_widths(ws, 1, dup_end_col, dup_header_row, dup_last_row)
    if "What Was Found" in duplicate_headers:
        ws.column_dimensions[
            get_column_letter(duplicate_headers.index("What Was Found") + 1)
        ].width = 52
    if "Reason" in duplicate_headers:
        ws.column_dimensions[
            get_column_letter(duplicate_headers.index("Reason") + 1)
        ].width = 46
    if "Reviewer Note" in duplicate_headers:
        ws.column_dimensions[
            get_column_letter(duplicate_headers.index("Reviewer Note") + 1)
        ].width = 36
    if len(duplicate_frame):
        dup_note_col = get_column_letter(dup_end_col)
        dup_validation = DataValidation(
            type="textLength", operator="lessThanOrEqual", formula1="1000", allow_blank=True
        )
        dup_validation.error = "Reviewer notes are limited to 1,000 characters."
        dup_validation.errorTitle = "Note too long"
        ws.add_data_validation(dup_validation)
        dup_validation.add(f"{dup_note_col}{dup_data_row}:{dup_note_col}{dup_last_row}")

    # Review Holds: every QuickBooks row the engine could not safely match
    # but that has SOME evidence of a possible Infinium counterpart -- a
    # potential duplicate, a PO already represented, an amount variance, a
    # non-unique or ambiguous candidate, controlled-typo candidates, or a
    # historical clearance blocked by an unresolved historical duplicate.
    # "Cannot safely match" is never treated as "does not exist in
    # Infinium": every row here is excluded from the proposed JE and stays
    # here, visible, until a reviewer records a disposition.
    review_title_row = dup_last_row + 3
    review_caption_row = review_title_row + 1
    review_kpi_label_row = review_title_row + 2
    review_kpi_value_row = review_title_row + 3
    review_header_row = review_title_row + 5
    review_data_row = review_header_row + 1

    review_caption = (
        f"{review_hold_count:,} row(s) are held for review, out of the proposed JE, pending a "
        "documented decision -- see Reason Code for why each was held and Reason Code Glossary for "
        "the full explanation of every code. Suggested dispositions: Release to JE (a confirmed "
        "genuine transaction), Exclude (a confirmed duplicate or already-represented transaction), "
        "Confirm Match (accept a candidate manually), or Carry Forward (needs more investigation)."
        if review_hold_count
        else "No QuickBooks rows are held for review."
    )
    _write_title_band(
        ws, review_title_row, 1, section_end_col,
        "REVIEW HOLDS | EXCLUDED FROM PROPOSED JE - REQUIRES DOCUMENTED DISPOSITION", SLATE,
    )
    _write_caption_band(ws, review_caption_row, 1, section_end_col, review_caption, SLATE)

    review_kpis = [
        ("Items held for review", review_hold_count, ACCOUNTING_COUNT_FORMAT),
        ("Amount excluded from JE", review_hold_amount, ACCOUNTING_CURRENCY_FORMAT),
        ("JE inclusion", "Excluded pending disposition", 'General'),
    ]
    _write_kpi_band(ws, review_kpi_label_row, review_kpi_value_row, review_kpis, review_end_col)

    _write_dataframe_values(ws, review_frame, review_header_row, 1)
    _format_header(
        ws, review_header_row, 1, review_end_col, SLATE,
        headers=review_headers, amount_columns={"Amount", "Infinium Amount", "Difference"},
    )
    review_ref_col = review_headers.index("Referenced Match Ref.") + 1
    review_ref_col_letter = get_column_letter(review_ref_col)
    ws.cell(review_header_row, review_ref_col).alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True,
    )
    if len(review_frame):
        review_last_row = review_data_row + len(review_frame) - 1
        _format_body_block(ws, review_data_row, review_last_row, 1, review_end_col, SLATE_LIGHT)
        _apply_number_formats(
            ws, review_headers, review_data_row, review_last_row, 1,
            {"Amount", "Infinium Amount", "Difference"}, set(),
        )
        # One consistent amber tint on the Reason cell -- a single disposition
        # (REVIEW HOLD) no longer needs a different color per sub-category.
        reason_col = review_headers.index("Reason") + 1
        row_id_col = review_headers.index("Row ID") + 1
        disposition_col = review_headers.index("Reviewer Disposition") + 1
        reviewer_col = review_headers.index("Reviewer") + 1
        review_date_col = review_headers.index("Review Date") + 1
        comment_col = review_headers.index("Comment") + 1
        for offset, record in enumerate(review_frame.to_dict("records")):
            row = review_data_row + offset
            ws.cell(row, reason_col).fill = PatternFill("solid", fgColor=AMBER)
            ws.cell(row, reason_col).alignment = Alignment(wrap_text=True, vertical="center")
            _apply_row_id_hyperlink(ws, row, row_id_col, qb_id_row_map.get(str(record["Row ID"])))
            pointer = record["Referenced Match Ref."]
            if pointer and detail_ref_letter:
                _style_match_ref_link(ws.cell(row, review_ref_col), pointer, detail_ref_letter)
            else:
                ws.cell(row, review_ref_col).alignment = Alignment(horizontal="center", vertical="center")
            for col in (reviewer_col, review_date_col, comment_col):
                ws.cell(row, col).protection = Protection(locked=False)
        disposition_validation = DataValidation(
            type="list",
            formula1='"Pending Review,Release to JE,Exclude,Confirm Match,Carry Forward"',
            allow_blank=False,
        )
        disposition_validation.error = "Select a disposition from the list before posting."
        disposition_validation.errorTitle = "Disposition required"
        ws.add_data_validation(disposition_validation)
        disposition_letter = get_column_letter(disposition_col)
        disposition_validation.add(f"{disposition_letter}{review_data_row}:{disposition_letter}{review_last_row}")
        for row in range(review_data_row, review_last_row + 1):
            ws.cell(row, disposition_col).protection = Protection(locked=False)
        comment_validation = DataValidation(
            type="textLength", operator="lessThanOrEqual", formula1="1000", allow_blank=True
        )
        comment_validation.error = "Comments are limited to 1,000 characters."
        comment_validation.errorTitle = "Comment too long"
        ws.add_data_validation(comment_validation)
        comment_letter = get_column_letter(comment_col)
        comment_validation.add(f"{comment_letter}{review_data_row}:{comment_letter}{review_last_row}")
        # Confirm Match asserts this row IS a match -- the reviewer must
        # identify what it matched to (the engine's own related reference, if
        # any, or a documented comment), so an unresolved row can never be
        # marked matched without saying against what.
        related_inf_col = review_headers.index("Related Infinium Row IDs") + 1
        related_inf_letter = get_column_letter(related_inf_col)
        confirm_match_rule = FormulaRule(
            formula=[
                f'AND(${disposition_letter}{review_data_row}="Confirm Match",'
                f'${review_ref_col_letter}{review_data_row}="",'
                f'${related_inf_letter}{review_data_row}="",'
                f'${comment_letter}{review_data_row}="")'
            ],
            fill=PatternFill("solid", fgColor=RED_LIGHT),
        )
        # openpyxl/Excel sqref multi-area syntax is space-separated, not comma-separated.
        confirm_match_range = (
            f"{review_ref_col_letter}{review_data_row}:{review_ref_col_letter}{review_last_row} "
            f"{related_inf_letter}{review_data_row}:{related_inf_letter}{review_last_row}"
        )
        ws.conditional_formatting.add(confirm_match_range, confirm_match_rule)
        # Exclude removes an evidence-backed hold from consideration entirely
        # -- also requires a documented reason.
        review_exclude_rule = FormulaRule(
            formula=[f'AND(${disposition_letter}{review_data_row}="Exclude",${comment_letter}{review_data_row}="")'],
            fill=PatternFill("solid", fgColor=RED_LIGHT),
        )
        ws.conditional_formatting.add(
            f"{comment_letter}{review_data_row}:{comment_letter}{review_last_row}", review_exclude_rule,
        )
    else:
        review_last_row = review_header_row

    review_total_row = review_last_row + 1
    _write_total_row(
        ws, review_total_row, 1, review_end_col,
        {"Amount": float(review_frame["Amount"].sum()) if len(review_frame) else 0.0},
        review_headers, "REVIEW HOLDS TOTAL",
    )
    _add_exception_table(
        ws,
        table_name="ReviewHolds",
        headers=review_headers,
        header_row=review_header_row,
        total_row=review_total_row,
        start_col=1,
        total_label="REVIEW HOLDS TOTAL",
        summed_headers={"Amount"},
        style_name="TableStyleMedium2",
    )

    _set_widths(ws, 1, review_end_col, review_header_row, review_total_row)
    ws.column_dimensions[get_column_letter(review_headers.index("Reason Code") + 1)].width = 26
    ws.column_dimensions[get_column_letter(review_headers.index("Reason") + 1)].width = 52
    ws.column_dimensions[get_column_letter(review_headers.index("Reviewer Disposition") + 1)].width = 22
    ws.column_dimensions[get_column_letter(review_headers.index("Reviewer") + 1)].width = 18
    ws.column_dimensions[get_column_letter(review_headers.index("Review Date") + 1)].width = 14
    ws.column_dimensions[get_column_letter(review_headers.index("Comment") + 1)].width = 36

    # PO Re-use Error: unlike every section above, these rows are NOT
    # withheld from the accrual -- they already appear, and are already
    # counted, in the QuickBooks exceptions table at the top of this
    # sheet. This section is purely supplementary grouped detail (PO,
    # QuickBooks total, Infinium total, difference, row counts) so a
    # reused PO doesn't read as several unrelated individual exceptions.
    po_reuse_title_row = review_total_row + 3
    po_reuse_caption_row = po_reuse_title_row + 1
    po_reuse_kpi_label_row = po_reuse_title_row + 2
    po_reuse_kpi_value_row = po_reuse_title_row + 3
    po_reuse_header_row = po_reuse_title_row + 5
    po_reuse_data_row = po_reuse_header_row + 1
    po_reuse_section_end_col = section_end_col

    po_reuse_caption = (
        f"{po_reuse_group_count:,} PO(s) appear more than once in the unresolved QuickBooks pool "
        "with a grouped total that does not tie exactly to the grouped Infinium total for the same "
        "PO. A row with Infinium evidence for its PO is held (see the Review Hold section) and is not "
        "accrued; a row with none remains in the exceptions table. This section shows the grouped PO "
        "detail so the pattern is traceable instead of reading as several unrelated exceptions."
        if po_reuse_group_count
        else "No PO Re-use Errors were identified (every PO repeated in the unresolved QuickBooks "
        "pool either ties exactly to Infinium -- and was already matched -- or appears only once)."
    )
    _write_title_band(
        ws, po_reuse_title_row, 1, po_reuse_section_end_col,
        "PO RE-USE ERROR | GROUPED DETAIL - HELD WHEN INFINIUM HAS EVIDENCE", SLATE,
    )
    _write_caption_band(ws, po_reuse_caption_row, 1, po_reuse_section_end_col, po_reuse_caption, SLATE)

    po_reuse_kpis = [
        ("PO groups flagged", po_reuse_group_count, ACCOUNTING_COUNT_FORMAT),
        ("QuickBooks rows involved", po_reuse_qb_row_count, ACCOUNTING_COUNT_FORMAT),
        ("Net difference", po_reuse_net_difference, ACCOUNTING_CURRENCY_FORMAT),
    ]
    _write_kpi_band(ws, po_reuse_kpi_label_row, po_reuse_kpi_value_row, po_reuse_kpis, po_reuse_end_col)

    _write_dataframe_values(ws, po_reuse_frame, po_reuse_header_row, 1)
    _format_header(
        ws, po_reuse_header_row, 1, po_reuse_end_col, SLATE,
        headers=po_reuse_headers,
        amount_columns={"QuickBooks Total", "Infinium Total", "Difference"},
        quantity_columns={"QuickBooks Row Count", "Infinium Row Count"},
    )
    if len(po_reuse_frame):
        po_reuse_last_row = po_reuse_data_row + len(po_reuse_frame) - 1
        _format_body_block(ws, po_reuse_data_row, po_reuse_last_row, 1, po_reuse_end_col, SLATE_LIGHT)
        _apply_number_formats(
            ws, po_reuse_headers, po_reuse_data_row, po_reuse_last_row, 1,
            {"QuickBooks Total", "Infinium Total", "Difference"},
            {"QuickBooks Row Count", "Infinium Row Count"},
        )
    else:
        po_reuse_last_row = po_reuse_header_row

    _set_widths(ws, 1, po_reuse_end_col, po_reuse_header_row, po_reuse_last_row)
    if "Explanation" in po_reuse_headers:
        ws.column_dimensions[
            get_column_letter(po_reuse_headers.index("Explanation") + 1)
        ].width = 52
    if "QuickBooks Row IDs" in po_reuse_headers:
        ws.column_dimensions[
            get_column_letter(po_reuse_headers.index("QuickBooks Row IDs") + 1)
        ].width = 30
    if "Infinium Row IDs" in po_reuse_headers:
        ws.column_dimensions[
            get_column_letter(po_reuse_headers.index("Infinium Row IDs") + 1)
        ].width = 30

    je_title_row = po_reuse_last_row + 3
    je_caption_row = je_title_row + 1
    je_header_row = je_title_row + 2
    je_data_row = je_header_row + 1
    je_headers = [
        "Entry Name", "GL Account", "Account Name", "Debit", "Credit", "Entry Basis",
    ]
    je_frame = pd.DataFrame(
        [
            [
                "AC001 Sales Accrual",
                "017-00000-110160.0",
                "Accrued Income",
                0.0,
                0.0,
                "Unresolved QuickBooks net exception support",
            ],
            [
                "AC001 Sales Accrual",
                "017-91000-400000-0",
                "Income-Manufacturing",
                0.0,
                0.0,
                "Balanced offset",
            ],
        ],
        columns=je_headers,
    )
    _write_title_band(
        ws, je_title_row, 1, section_end_col,
        "PROPOSED JOURNAL ENTRY | AC001 SALES ACCRUAL",
        SLATE,
    )
    _write_caption_band(
        ws, je_caption_row, 1, section_end_col,
        "Post only after review and approval. Debit 017-00000-110160.0 Accrued Income and credit "
        "017-91000-400000-0 Income-Manufacturing for the Final Approved JE (the engine's TRUE_UNMATCHED "
        "total, plus Review Holds released to JE, less any documented manual exclusions -- see the "
        "Posting Summary bridge); evaluate reversals and negative source values before posting.",
        SLATE,
    )
    _write_dataframe_values(ws, je_frame, je_header_row, 1)
    je_amount_formula = (
        f"=ABS({qb_amount_sum_expr}+{_review_holds_released_expr()}"
        f"-{_je_support_manual_exclusions_expr(result)})"
    )
    ws.cell(je_data_row, 4, je_amount_formula)
    ws.cell(je_data_row, 5, 0.0)
    ws.cell(je_data_row + 1, 4, 0.0)
    ws.cell(je_data_row + 1, 5, je_amount_formula)
    _format_header(
        ws, je_header_row, 1, len(je_headers), SLATE,
        headers=je_headers, amount_columns={"Debit", "Credit"},
    )
    _format_body_block(ws, je_data_row, je_data_row + len(je_frame) - 1, 1, len(je_headers), SLATE_LIGHT)
    for row in range(je_data_row, je_data_row + len(je_frame)):
        ws.cell(row, 4).number_format = ACCOUNTING_CURRENCY_FORMAT
        ws.cell(row, 5).number_format = ACCOUNTING_CURRENCY_FORMAT
    je_total_row = je_data_row + len(je_frame)
    _write_total_row(
        ws, je_total_row, 1, len(je_headers),
        {"Debit": 0.0, "Credit": 0.0}, je_headers, "BALANCED TOTAL",
    )
    ws.cell(je_total_row, 4, f"=SUM(D{je_data_row}:D{je_data_row + len(je_frame) - 1})")
    ws.cell(je_total_row, 5, f"=SUM(E{je_data_row}:E{je_data_row + len(je_frame) - 1})")
    for row in range(je_data_row, je_total_row + 1):
        for col in (4, 5):
            ws.cell(row, col).protection = Protection(locked=True)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 24)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 23)
    ws.column_dimensions["C"].width = max(ws.column_dimensions["C"].width or 0, 28)
    ws.column_dimensions["F"].width = max(ws.column_dimensions["F"].width or 0, 52)

    ws.print_title_rows = "1:4"
    _prepare_sheet(ws)


_LEGACY_MATCHED_SECTIONS = {
    "01 Matched",
    "01 Matched - Historical Clearance",
}
_LEGACY_QB_DUPLICATE_SECTIONS = {"04 Duplicate QuickBooks"}
_LEGACY_INF_DUPLICATE_SECTIONS = {"05 Duplicate Infinium"}
_LEGACY_DUPLICATE_SECTIONS = _LEGACY_QB_DUPLICATE_SECTIONS | _LEGACY_INF_DUPLICATE_SECTIONS
_LEGACY_INF_UNMATCHED_SECTION = "03 Unmatched Infinium"


def _legacy_section_label(section: str) -> str:
    """Strip the sort-order prefix (e.g. "04 ") from a Section value for a
    plain, management-facing exception type label."""
    prefix, _, remainder = str(section).partition(" ")
    return remainder if prefix.isdigit() and remainder else str(section)


def _legacy_row_values(
    index: Optional[int],
    scope: Optional[str],
    primary_frame: Optional[pd.DataFrame],
    historical_frame: Optional[pd.DataFrame],
    headers: list[str],
) -> list[Any]:
    if index is None:
        return [None] * len(headers)
    source = historical_frame if scope == "Historical" and historical_frame is not None else primary_frame
    if source is None or index not in source.index:
        return [None] * len(headers)
    row = source.loc[index]
    return [row.get(header) for header in headers]


def _legacy_norm_po_sort_key(index: Optional[int], frame: pd.DataFrame) -> tuple[bool, str]:
    """Sort key that puts a blank/unavailable normalized PO last."""
    value = ""
    if index is not None and NORM_PO in frame.columns and index in frame.index:
        value = str(frame.at[index, NORM_PO] or "")
    return (value == "", value)


# Fixed, type-appropriate widths for a raw source frame's mapped fields --
# same idea as _apply_number_formats' amount/quantity/date treatment, but
# for column width. Keyed by role rather than header text so it works
# whatever the export happens to call these columns (e.g. Infinium's
# amount field is literally "OHTOTA", which doesn't textually resemble
# "amount" at all).
_LEGACY_ROLE_WIDTHS = {"period": 10, "invoice": 16, "po": 18, "amount": 15, "quantity": 12}


def _legacy_fixed_widths(mapping: dict[str, Optional[str]]) -> dict[str, float]:
    widths: dict[str, float] = {}
    for role, width in _LEGACY_ROLE_WIDTHS.items():
        column = mapping.get(role)
        if column:
            widths[column] = width
    return widths


def _standardize_legacy_widths(ws, headers: list[str], start_col: int, mapping: dict[str, Optional[str]]) -> None:
    """Apply role-based fixed widths for mapped fields, plus a generic
    date-column width for any remaining header that looks like a date --
    the one column type with no dedicated mapping key of its own."""
    _standardize_column_widths(ws, headers, start_col, _legacy_fixed_widths(mapping))
    for offset, header in enumerate(headers):
        header_upper = str(header).upper()
        if "DATE" in header_upper or "TIMESTAMP" in header_upper:
            ws.column_dimensions[get_column_letter(start_col + offset)].width = 13


def _legacy_infinium_display_headers(headers: list[str]) -> list[str]:
    return [INFINIUM_FRIENDLY_HEADERS.get(str(header).strip().upper(), header) for header in headers]


def _legacy_matched_label(match_result: str) -> str:
    """"Unique Match: PO + Amount", "Group Match: Invoice + Amount", etc. --
    only how the match was made, never the confidence tier. A vendor-alias
    match is a PO-field + amount match, and a historical (prior-period)
    clearance uses the same underlying rule after its "... | " prefix."""
    text = str(match_result).split("|")[-1].strip()
    is_group = "Grouped" in text or "group-level" in str(match_result)
    if "PO + Invoice" in text:
        keys = "PO + Invoice + Amount"
    elif text.startswith("Invoice +"):
        keys = "Invoice + Amount"
    else:
        keys = "PO + Amount"
    return f"{'Group' if is_group else 'Unique'} Match: {keys}"


def _legacy_reference_label(record: dict) -> Optional[str]:
    """Legacy wording for an exception that points at an accepted match --
    built from the canonical reference stored on the paired row, never
    re-derived here, so it always names a match that exists."""
    references = _split_cell_references(record.get("Referenced Match Ref."))
    if not references:
        return None
    basis = record.get("Reference Basis")
    if basis == "Group":
        return f"Review: Candidate Belongs to {describe_match_references(references, group=True)}"
    if basis in ("PO", "Invoice"):
        return f"Review: {basis} Already Used by {describe_match_references(references)}"
    if basis == "Record":
        return f"Infinium Record Already Assigned to {describe_match_references(references)}"
    if record.get("Section") in _LEGACY_DUPLICATE_SECTIONS:
        return f"Exact Duplicate of {describe_match_references(references)} - Excluded"
    return f"Potential Duplicate of {describe_match_references(references)}"


def _split_cell_references(text: Any) -> list[str]:
    return [piece.strip() for piece in str(text or "").split(";") if piece.strip()]


def _legacy_final_disposition(record: dict) -> str:
    """The same four top-level dispositions shown everywhere else in the
    workbook (MATCHED / TRUE UNMATCHED / REVIEW HOLD / DUPLICATE EXCLUDED),
    derived straight from the paired row's Section -- so this ledger and the
    Unresolved Exceptions / Reconciliation Detail sheets can never disagree.
    Blank for an Infinium-only row, which carries no QuickBooks disposition."""
    section = str(record.get("Section", ""))
    if section in _LEGACY_MATCHED_SECTIONS:
        return "MATCHED"
    if record.get("QB Index") is None:
        return ""
    if section in _LEGACY_QB_DUPLICATE_SECTIONS:
        return "DUPLICATE EXCLUDED"
    if section == "02 Unmatched QuickBooks":
        return "TRUE UNMATCHED"
    return "REVIEW HOLD"


_LEGACY_REFERENCE_HOLD_LABELS = {
    "REVIEW_HOLD_AMOUNT_VARIANCE": "Amount Differs: Same PO/Invoice",
    "REVIEW_HOLD_EXACT_CANDIDATE_NOT_UNIQUE": "Review: Exact Candidate Not Unique",
    "REVIEW_HOLD_MULTIPLE_CANDIDATES": "Potential Duplicate: Multiple Infinium Candidates",
    "REVIEW_HOLD_INVALID_AMOUNT": "Review: Invalid Amount",
    "REVIEW_HOLD_CANDIDATE_INVALID_AMOUNT": "Review: Candidate Amount Invalid",
    "REVIEW_HOLD_PO_ALREADY_REPRESENTED": "Review: PO Already Represented",
}


def _legacy_match_method_label(record: dict) -> str:
    """The Legacy Reconciliation "Match Result" text: just how the match
    was made (or why there isn't one). Confidence tiers stay on the primary
    workpaper and analytics package -- the accountant's legacy view is a
    read-on-sight summary, and the row's fill already carries the status."""
    section = str(record.get("Section", ""))
    match_result = str(record.get("Match Result", ""))
    if section in _LEGACY_MATCHED_SECTIONS:
        return _legacy_matched_label(match_result)
    if section == "09 Fuzzy Match Review Hold":
        # A fuzzy match is always held for review -- never counted as
        # reconciled -- so its label lives on the review path, not the
        # matched one, even though the wording is unchanged.
        return "Possible Match: Similar PO + Amount"
    pointer = _legacy_reference_label(record)
    if pointer:
        return pointer
    if section == "02 Unmatched QuickBooks":
        return "No Matching Infinium Records"
    if section == "03 Unmatched Infinium":
        return "No Matching QuickBooks Records"
    if section in _LEGACY_DUPLICATE_SECTIONS:
        return "Duplicate: Excess Copy Excluded"
    if section in {"06 Duplicate Review Hold QuickBooks", "07 Duplicate Review Hold Infinium"}:
        basis = record.get("Duplicate Basis")
        if basis == DUPLICATE_BASIS_PO_ONLY:
            return "Potential Duplicate: PO + Amount"
        if basis == DUPLICATE_BASIS_INVOICE_ONLY:
            return "Potential Duplicate: Invoice + Amount"
        return "Potential Duplicate: PO + Invoice + Amount"
    if section == "08 Reference-Matched Amount Variance Review Hold":
        return "Amount Differs: Same PO/Invoice"
    if section == "10 Ambiguous Duplicate QuickBooks":
        return "Potential Duplicate: Multiple Infinium Candidates"
    if section == REFERENCE_HOLD_SECTION:
        return _LEGACY_REFERENCE_HOLD_LABELS.get(str(record.get("Reason Code")), match_result)
    return match_result


def _legacy_row_needs_attention(section: str) -> bool:
    """A review row (gold) or a fuzzy possible match is the only kind of row
    whose status text is bolded; a clean match and an already-decided
    excluded duplicate are informational."""
    return (
        section not in _LEGACY_MATCHED_SECTIONS and section not in _LEGACY_DUPLICATE_SECTIONS
    ) or section == "09 Fuzzy Match Review Hold"


def _write_legacy_legend(ws, row: int, start_col: int) -> None:
    """A compact color key just under the introductory note: three adjacent
    chips, each filled with the exact row tint it explains and carrying its
    own label, so the key and the rows can never disagree. (A colored "■"
    glyph would be invisible for the paler tints.) Text shrinks to fit
    whatever width its column happens to have."""
    chip_border_side = Side(style="thin", color="BFBFBF")
    chips = [
        (LEGACY_MATCHED_FILL, "Reconciled"),
        (LEGACY_REVIEW_FILL, "Review required"),
        (LEGACY_EXCLUDED_FILL, "Excluded duplicate"),
        (LEGACY_NO_PAIR_FILL, "No paired record"),
    ]
    for offset, (fill_color, label) in enumerate(chips):
        cell = ws.cell(row, start_col + offset, label)
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.font = Font(name=FONT_NAME, size=9, color=LEGACY_BODY_TEXT)
        cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        cell.border = Border(
            left=chip_border_side, right=chip_border_side, top=chip_border_side, bottom=chip_border_side,
        )
    ws.row_dimensions[row].height = 18


def _legacy_generated_stamp(run_timestamp: datetime) -> str:
    """"09/19/2026 3:21 PM CDT" -- MM/DD/YYYY and a 12-hour clock, in
    Central time, matching how every date on the legacy sheets displays."""
    local = run_timestamp.astimezone(CENTRAL_TIMEZONE)
    hour = local.strftime("%I").lstrip("0") or "12"
    return f"{local:%m/%d/%Y} {hour}:{local:%M} {local:%p} {local:%Z}"


_LEGACY_DATE_INPUT_FORMATS = ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%Y-%m-%d %H:%M:%S", "%m-%d-%Y")


def _legacy_parse_date(value: Any) -> Optional[datetime]:
    """A source date as a real date: an existing datetime/date passes
    through, and text like "1/09/2026" or "2026-03-03" is parsed month-first
    (the export's own convention). Anything else is not treated as a date."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if not isinstance(value, str) or not value.strip():
        return None
    text = value.strip()
    for pattern in _LEGACY_DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            continue
    return None


def _legacy_ensure_headers_fit(ws, headers: list[str], start_col: int) -> None:
    """Widen any column too narrow for its own bold heading plus the
    autofilter button Excel draws inside it -- otherwise a short-valued
    column such as an Infinium "Customer No" (five-digit values) is sized
    to its data and its heading is clipped to "Customer N"."""
    for offset, header in enumerate(headers):
        letter = get_column_letter(start_col + offset)
        needed = len(str(header)) + 5
        current = ws.column_dimensions[letter].width or 0
        if current < needed:
            ws.column_dimensions[letter].width = needed


def _legacy_standardize_dates(
    ws, headers: list[str], start_col: int, first_row: int, last_row: int,
) -> None:
    """Every date-column cell on a legacy sheet becomes a real date shown as
    MM/DD/YYYY. QuickBooks delivers some dates as real dates and others as
    text, which otherwise render as "2026-03-03" beside "08/17/2026"; a
    text date is converted, and all date cells are centered alike."""
    for offset, header in enumerate(headers):
        if "DATE" not in str(header).upper():
            continue
        for row in range(first_row, last_row + 1):
            cell = ws.cell(row, start_col + offset)
            parsed = _legacy_parse_date(cell.value)
            if parsed is None:
                continue
            cell.value = parsed
            cell.number_format = DATE_NUMBER_FORMAT
            cell.alignment = Alignment(horizontal="center", vertical="center")


def build_legacy_reconciliation_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    """A simplified, side-by-side QuickBooks/Infinium sheet styled after the
    accountant's original hand-built workbook: every QuickBooks row (sorted
    by normalized PO), colored by outcome, with its matched Infinium row
    riding along on the same line when one exists and blank when it
    doesn't. QuickBooks exceptions repeat on their own Exceptions sheet;
    Infinium is only shown here (never on Exceptions) and only for rows
    that actually matter on this sheet -- a confirmed match, an Infinium
    duplicate (any period), or an unmatched Infinium row from the current
    selected period. Older-period Infinium noise with no QuickBooks tie
    is intentionally left out.
    """
    ws = wb.active
    ws.title = "Legacy Reconciliation"
    qb_headers = list(result.qb_raw.columns)
    inf_headers = list(result.inf_raw.columns)
    qb_start = 1
    # Final Disposition (the same four top-level outcomes shown everywhere
    # else) leads the panel, followed by Match Ref., Match Result, and
    # Referenced Match Ref. -- prioritized in that order so the outcome a
    # reviewer needs is the first thing to the right of the QuickBooks block,
    # ahead of the how-it-resolved and why-it-points-elsewhere detail.
    disposition_col = len(qb_headers) + 1
    ref_col = disposition_col + 1
    method_col = ref_col + 1
    referenced_col = method_col + 1
    inf_start = referenced_col + 1
    qb_end = len(qb_headers)
    inf_end = inf_start + len(inf_headers) - 1
    # Row 3 holds the color legend, directly under the introductory note.
    legend_row, header_row, data_row = 3, 4, 5

    default_year = int(result.metadata.get("fiscal_year") or result.run_timestamp.year)
    selected_period = result.metadata.get("fiscal_period")
    inf_period_col = result.inf_mapping.get("period")

    def inf_row_period(record: dict) -> Any:
        iidx = record.get("Infinium Index")
        if iidx is not None and inf_period_col and iidx in result.inf_work.index:
            period, _ = parse_fiscal_period(result.inf_work.at[iidx, inf_period_col], default_year)
            return period
        return None

    qb_rows = [record for record in result.paired_rows if record.get("QB Index") is not None]
    inf_only_rows = [
        record for record in result.paired_rows
        if record.get("QB Index") is None
        and record.get("Infinium Index") is not None
        and (
            record.get("Section") in _LEGACY_INF_DUPLICATE_SECTIONS
            or (
                record.get("Section") == _LEGACY_INF_UNMATCHED_SECTION
                and selected_period is not None
                and inf_row_period(record) == int(selected_period)
            )
        )
    ]
    qb_rows.sort(key=lambda record: _legacy_norm_po_sort_key(record.get("QB Index"), result.qb_work))
    inf_only_rows.sort(key=lambda record: _legacy_norm_po_sort_key(record.get("Infinium Index"), result.inf_work))
    all_rows = qb_rows + inf_only_rows
    final_data_row = data_row + max(len(all_rows), 1) - 1
    matched_count = sum(1 for record in all_rows if record.get("Section") in _LEGACY_MATCHED_SECTIONS)
    # The red-row sentence in the note is only meaningful if a red row exists.
    has_excluded_duplicates = any(record.get("Section") in _LEGACY_DUPLICATE_SECTIONS for record in all_rows)

    _write_title_band(ws, 1, qb_start, qb_end, "QUICKBOOKS | SORTED BY PO", NAVY)
    _write_title_band(ws, 1, disposition_col, referenced_col, "MATCH RESULT", SLATE)
    _write_title_band(ws, 1, inf_start, inf_end, "INFINIUM", TEAL)
    _write_caption_band(
        ws, 2, qb_start, qb_end,
        f"All {len(qb_rows):,} QuickBooks records accounted for -- {matched_count:,} matched "
        f"({(matched_count / len(qb_rows) * 100) if qb_rows else 0:.1f}%). "
        f"{'Gold rows require review; red rows are excluded duplicates.' if has_excluded_duplicates else 'Gold rows require review.'} "
        f"See Exceptions for details. Generated {_legacy_generated_stamp(result.run_timestamp)}.",
        NAVY,
    )
    _write_caption_band(
        ws, 2, disposition_col, referenced_col,
        "Final Disposition (MATCHED / TRUE UNMATCHED / REVIEW HOLD / DUPLICATE EXCLUDED), then how each "
        "row resolved, or why it did not.",
        SLATE,
    )
    _write_caption_band(
        ws, 2, inf_start, inf_end,
        "Blank unless matched. Unmatched Infinium rows are shown only for the currently "
        "selected fiscal period; an Infinium duplicate is shown for any period.",
        TEAL,
    )

    for offset, record in enumerate(all_rows):
        row = data_row + offset
        qb_values = _legacy_row_values(
            record.get("QB Index"), record.get("QB Record Scope"),
            result.qb_work, result.qb_secondary_work, qb_headers,
        )
        inf_values = _legacy_row_values(
            record.get("Infinium Index"), record.get("Infinium Record Scope"),
            result.inf_work, result.inf_secondary_work, inf_headers,
        )
        for col_offset, value in enumerate(qb_values):
            ws.cell(row, qb_start + col_offset, excel_safe(value))
        for col_offset, value in enumerate(inf_values):
            ws.cell(row, inf_start + col_offset, excel_safe(value))
        ws.cell(row, disposition_col, _legacy_final_disposition(record) or None)
        ws.cell(row, ref_col, record.get("Match Ref.") or None)
        ws.cell(row, method_col, _legacy_match_method_label(record))
        ws.cell(row, referenced_col, record.get("Referenced Match Ref.") or None)

    ws.cell(header_row, disposition_col, "Final Disposition")
    ws.cell(header_row, ref_col, "Match Ref.")
    ws.cell(header_row, method_col, "Match Result")
    ws.cell(header_row, referenced_col, "Referenced Match Ref.")
    _write_legacy_legend(ws, legend_row, qb_start)
    _write_dataframe_values(ws, pd.DataFrame(columns=qb_headers), header_row, qb_start)
    _write_dataframe_values(
        ws, pd.DataFrame(columns=_legacy_infinium_display_headers(inf_headers)), header_row, inf_start,
    )
    _format_header(ws, header_row, qb_start, qb_end, NAVY)
    _format_header(ws, header_row, disposition_col, referenced_col, SLATE)
    _format_header(ws, header_row, inf_start, inf_end, TEAL)

    blank_side_panels: list[tuple[int, int, int]] = []
    for offset, record in enumerate(all_rows):
        row = data_row + offset
        section = record.get("Section", "")
        # Alignment/border first, uniform across the whole row regardless
        # of outcome, so every cell has a consistent look; the color style
        # applied next only ever touches fill/font, never alignment.
        _apply_default_alignment(ws, row, qb_start, inf_end)
        if section in _LEGACY_MATCHED_SECTIONS:
            status_fill = LEGACY_MATCHED_FILL
        elif section in _LEGACY_DUPLICATE_SECTIONS:
            status_fill = LEGACY_EXCLUDED_FILL
        else:
            status_fill = LEGACY_REVIEW_FILL
        # A side with no record in its own dataset (no paired Infinium row,
        # or an Infinium-only row with no QuickBooks row) stays blank in a
        # near-white gray, so it reads as "nothing here" rather than as a
        # second, empty status band.
        has_qb = record.get("QB Index") is not None
        has_inf = record.get("Infinium Index") is not None
        _apply_legacy_status_fill(ws, row, qb_start, qb_end, status_fill if has_qb else LEGACY_NO_PAIR_FILL)
        _apply_legacy_status_fill(ws, row, inf_start, inf_end, status_fill if has_inf else LEGACY_NO_PAIR_FILL)
        if not has_qb:
            blank_side_panels.append((row, qb_start, qb_end))
        if not has_inf:
            blank_side_panels.append((row, inf_start, inf_end))
        needs_attention = _legacy_row_needs_attention(section)
        _apply_legacy_status_cell(ws, row, disposition_col, LEGACY_METHOD_FILL, needs_attention)
        _apply_legacy_status_cell(ws, row, ref_col, LEGACY_METHOD_FILL, False)
        _apply_legacy_status_cell(ws, row, method_col, LEGACY_METHOD_FILL, needs_attention)
        _apply_legacy_status_cell(ws, row, referenced_col, LEGACY_METHOD_FILL, False)
        ws.cell(row, method_col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for code_col in (disposition_col, ref_col, referenced_col):
            ws.cell(row, code_col).alignment = Alignment(horizontal="center", vertical="center")
        for panel_col in (disposition_col, ref_col, method_col, referenced_col):
            ws.cell(row, panel_col).border = _thin_border()
    if not all_rows:
        _apply_default_alignment(ws, data_row, qb_start, inf_end)
        for panel_col in (disposition_col, ref_col, method_col, referenced_col):
            _apply_legacy_status_cell(ws, data_row, panel_col, LEGACY_METHOD_FILL, False)

    total_row = final_data_row + 1
    qb_display = pd.DataFrame(
        [_legacy_row_values(r.get("QB Index"), r.get("QB Record Scope"), result.qb_work, result.qb_secondary_work, qb_headers) for r in qb_rows],
        columns=qb_headers,
    )
    inf_display = pd.DataFrame(
        [_legacy_row_values(r.get("Infinium Index"), r.get("Infinium Record Scope"), result.inf_work, result.inf_secondary_work, inf_headers) for r in all_rows],
        columns=inf_headers,
    )
    _write_total_row(ws, total_row, qb_start, qb_end,
                     _source_totals(qb_display, result.qb_mapping), qb_headers, "QUICKBOOKS TOTAL")
    _write_total_row(ws, total_row, inf_start, inf_end,
                     _source_totals(inf_display, result.inf_mapping), inf_headers, "INFINIUM TOTAL (SHOWN)")
    for panel_col in (disposition_col, ref_col, method_col, referenced_col):
        ws.cell(total_row, panel_col).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
        ws.cell(total_row, panel_col).border = _total_border()
    _apply_number_formats(ws, qb_headers, data_row, total_row, qb_start,
                          {result.qb_mapping["amount"]}, {result.qb_mapping.get("quantity") or ""})
    _apply_number_formats(ws, inf_headers, data_row, total_row, inf_start,
                          {result.inf_mapping["amount"]}, set())
    _set_widths(ws, qb_start, qb_end, header_row, total_row, maximum=40)
    _set_widths(ws, inf_start, inf_end, header_row, total_row, maximum=40)
    _standardize_legacy_widths(ws, qb_headers, qb_start, result.qb_mapping)
    _standardize_legacy_widths(ws, inf_headers, inf_start, result.inf_mapping)
    _legacy_ensure_headers_fit(ws, qb_headers, qb_start)
    _legacy_ensure_headers_fit(ws, _legacy_infinium_display_headers(inf_headers), inf_start)
    ws.column_dimensions[get_column_letter(method_col)].width = 46
    # Narrow, but wide enough for a reference and the full (wrapping) heading.
    ws.column_dimensions[get_column_letter(disposition_col)].width = 18
    ws.column_dimensions[get_column_letter(ref_col)].width = 12
    ws.column_dimensions[get_column_letter(referenced_col)].width = 16
    _legacy_standardize_dates(ws, qb_headers, qb_start, data_row, final_data_row)
    _legacy_standardize_dates(ws, _legacy_infinium_display_headers(inf_headers), inf_start, data_row, final_data_row)
    # Each blank side reads as one quiet panel instead of a row of empty
    # gridlined cells -- by dropping the borders BETWEEN its cells and keeping
    # only the block's outline, not by merging them: Excel refuses to sort a
    # range containing merged cells of different sizes ("all the merged cells
    # need to be the same size"), and this sheet has to stay sortable and
    # filterable. Formats travel with their rows through a sort.
    edge = _thin_border().top
    for panel_row, panel_start, panel_end in blank_side_panels:
        for panel_col in range(panel_start, panel_end + 1):
            ws.cell(panel_row, panel_col).border = Border(
                top=edge,
                bottom=edge,
                left=edge if panel_col == panel_start else Side(),
                right=edge if panel_col == panel_end else Side(),
            )
    ws.freeze_panes = f"{get_column_letter(inf_start)}{data_row}"
    if all_rows:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(inf_end)}{final_data_row}"
    ws.print_title_rows = f"1:{header_row}"
    _prepare_sheet(ws)


def build_legacy_exceptions_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    """A plain, single listing of QuickBooks-side exceptions only --
    unmatched QuickBooks rows, excluded duplicate copies, and every
    QuickBooks review-hold item -- grouped by fiscal period, styled after
    the accountant's original exceptions tab. Infinium-only exceptions
    (an unmatched or duplicate Infinium row with no QuickBooks
    counterpart) carry no accrual impact and are intentionally not
    repeated here. Genuine unresolved items are Excel's standard
    "Neutral" gold; excluded duplicate copies are "Bad" red.
    """
    ws = wb.create_sheet("Exceptions")
    qb_headers = list(result.qb_raw.columns)
    trailer_headers = ["Fiscal Period", "Exception Type", "Referenced Match Ref.", "Explanation"]
    n_qb = len(qb_headers)
    n_trailer = len(trailer_headers)

    # Two independent blocks on one sheet: general QuickBooks exceptions on
    # the left, excluded QuickBooks duplicate copies on the right -- a
    # duplicate is a definite, already-decided exclusion, not an open
    # question like the rest, so it gets its own space rather than being
    # mixed into the same list.
    qb_start = 1
    qb_end = n_qb
    left_trailer_start = qb_end + 1
    left_trailer_end = left_trailer_start + n_trailer - 1
    separator_col = left_trailer_end + 1
    dup_qb_start = separator_col + 1
    dup_qb_end = dup_qb_start + n_qb - 1
    dup_trailer_start = dup_qb_end + 1
    dup_trailer_end = dup_trailer_start + n_trailer - 1

    default_year = int(result.metadata.get("fiscal_year") or result.run_timestamp.year)
    qb_period_col = result.qb_mapping.get("period")

    def row_period(record: dict) -> Any:
        qidx = record.get("QB Index")
        if qidx is not None and qb_period_col and qidx in result.qb_work.index:
            period, _ = parse_fiscal_period(result.qb_work.at[qidx, qb_period_col], default_year)
            if period is not None:
                return period
        return None

    def sort_by_period(records: list[dict]) -> list[dict]:
        return sorted(records, key=lambda record: (row_period(record) is None, row_period(record) or 0))

    qb_side_rows = [
        record for record in result.paired_rows
        if record.get("Section") not in _LEGACY_MATCHED_SECTIONS
        and record.get("QB Index") is not None
    ]
    general_rows = sort_by_period(
        [r for r in qb_side_rows if r.get("Section") not in _LEGACY_QB_DUPLICATE_SECTIONS]
    )
    duplicate_rows = sort_by_period(
        [r for r in qb_side_rows if r.get("Section") in _LEGACY_QB_DUPLICATE_SECTIONS]
    )

    fiscal_summary = build_fiscal_exception_summary(result)
    fiscal_headers = list(fiscal_summary.columns)
    fiscal_end_col = max(len(fiscal_headers), 1)
    section_end_col = dup_trailer_end

    _write_title_band(ws, 1, qb_start, section_end_col, "EXCEPTIONS | QUICKBOOKS SIDE | BY FISCAL PERIOD", NAVY)
    _write_caption_band(
        ws, 2, qb_start, section_end_col,
        f"{len(general_rows):,} QuickBooks exception(s) at left (unmatched and review-hold items, shaded "
        f"gold) and {len(duplicate_rows):,} excluded QuickBooks duplicate copy(ies) at right (shaded red). "
        f"Infinium-only exceptions carry no accrual impact and are not repeated here. Generated "
        f"{_legacy_generated_stamp(result.run_timestamp)}.",
        NAVY,
    )

    summary_header_row = 4
    summary_data_row = summary_header_row + 1
    _write_dataframe_values(ws, fiscal_summary, summary_header_row, 1)
    _format_header(ws, summary_header_row, 1, fiscal_end_col, NAVY)
    if len(fiscal_summary):
        summary_last_row = summary_data_row + len(fiscal_summary) - 1
        _format_body_block(ws, summary_data_row, summary_last_row, 1, fiscal_end_col, NAVY_LIGHT)
        _apply_number_formats(
            ws, fiscal_headers, summary_data_row, summary_last_row, 1,
            {"Net Exception Amount"}, {"Exception Count", "Exception Quantity"},
        )
    else:
        summary_last_row = summary_header_row
    summary_total_row = summary_last_row + 1
    _write_total_row(
        ws, summary_total_row, 1, fiscal_end_col,
        {
            "Exception Count": float(fiscal_summary["Exception Count"].sum()) if len(fiscal_summary) else 0,
            "Exception Quantity": float(fiscal_summary["Exception Quantity"].sum()) if len(fiscal_summary) else 0,
            "Net Exception Amount": float(fiscal_summary["Net Exception Amount"].sum()) if len(fiscal_summary) else 0,
        },
        fiscal_headers, "TOTAL EXCEPTIONS",
    )
    _set_widths(ws, 1, fiscal_end_col, summary_header_row, summary_total_row)

    header_row = summary_total_row + 3
    data_row = header_row + 1

    def write_block(
        records: list[dict], block_qb_start: int, trailer_start: int, trailer_end: int,
        status_fill: str, bold_exception_type: bool,
    ) -> int:
        block_headers = qb_headers + trailer_headers
        block = pd.DataFrame(
            [
                _legacy_row_values(record.get("QB Index"), record.get("QB Record Scope"), result.qb_work, None, qb_headers)
                + [
                    row_period(record),
                    _legacy_section_label(str(record.get("Section", ""))),
                    record.get("Referenced Match Ref.") or None,
                    f"{record.get('Match Result', '')} -- {record.get('Explanation', '')}",
                ]
                for record in records
            ],
            columns=block_headers,
        )
        _write_dataframe_values(ws, block, header_row, block_qb_start)
        _format_header(ws, header_row, block_qb_start, block_qb_start + n_qb - 1, NAVY)
        _format_header(ws, header_row, trailer_start, trailer_end, SLATE)
        block_final_row = data_row + max(len(records), 1) - 1
        for offset in range(len(records)):
            row = data_row + offset
            _apply_default_alignment(ws, row, block_qb_start, trailer_end)
            _apply_legacy_status_fill(ws, row, block_qb_start, trailer_end, status_fill)
            # The Exception Type cell (trailer's second column) is the one
            # status cell worth bolding, and only for a genuine open item.
            _apply_legacy_status_cell(ws, row, trailer_start + 1, status_fill, bold_exception_type)
            ws.cell(row, trailer_end).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(row, trailer_start + 2).alignment = Alignment(horizontal="center", vertical="center")
        _apply_number_formats(ws, qb_headers, data_row, block_final_row, block_qb_start,
                              {result.qb_mapping["amount"]}, {result.qb_mapping.get("quantity") or ""})
        _legacy_standardize_dates(ws, qb_headers, block_qb_start, data_row, block_final_row)
        _set_widths(ws, block_qb_start, block_qb_start + n_qb - 1, header_row, block_final_row, maximum=40)
        _standardize_legacy_widths(ws, qb_headers, block_qb_start, result.qb_mapping)
        _legacy_ensure_headers_fit(ws, qb_headers, block_qb_start)
        ws.column_dimensions[get_column_letter(trailer_start)].width = 14
        ws.column_dimensions[get_column_letter(trailer_start + 1)].width = 34
        ws.column_dimensions[get_column_letter(trailer_start + 2)].width = 16
        ws.column_dimensions[get_column_letter(trailer_end)].width = 60
        return block_final_row

    general_final_row = write_block(
        general_rows, qb_start, left_trailer_start, left_trailer_end, LEGACY_REVIEW_FILL, True,
    )
    duplicate_final_row = write_block(
        duplicate_rows, dup_qb_start, dup_trailer_start, dup_trailer_end, LEGACY_EXCLUDED_FILL, False,
    )
    final_data_row = max(general_final_row, duplicate_final_row)

    ws.column_dimensions[get_column_letter(separator_col)].width = 3.5
    ws.column_dimensions[get_column_letter(separator_col)].fill = PatternFill("solid", fgColor=WHITE)
    ws.freeze_panes = f"{get_column_letter(qb_start)}{data_row}"
    if general_rows:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(left_trailer_end)}{general_final_row}"
    ws.print_title_rows = "1:2"
    _prepare_sheet(ws)


def build_legacy_workbook(result: ReconciliationResult) -> bytes:
    """The 'Accountant's Legacy Download' -- a simplified export mirroring
    the reviewer's original hand-built workbook (QuickBooks left, Infinium
    right, exceptions on their own tab by fiscal period) with Excel's
    standard Good/Neutral/Bad coloring, meant to be read on sight without
    the audit-trail depth of the primary workpaper. The primary workpaper
    and analytics package are unaffected by this export.
    """
    validate_match_references(result)
    wb = Workbook()
    wb.properties.creator = "Sales Reconciliation Application"
    wb.properties.title = f"Sales Reconciliation (Legacy Format) {result.run_id}"
    wb.properties.subject = "QuickBooks to Infinium reconciliation, accountant's legacy layout"
    wb.properties.description = (
        "Simplified accountant's legacy-format export generated from one controlled reconciliation run."
    )
    build_legacy_reconciliation_sheet(wb, result)
    build_legacy_exceptions_sheet(wb, result)
    build_product_sheet(wb, result)
    _apply_workbook_run_metadata(wb, result)
    # Legacy Reconciliation and Exceptions set their own deliberate,
    # type-appropriate column widths (see _standardize_legacy_widths) --
    # skip the workbook-wide content-driven autofit pass for them so those
    # widths actually stick instead of being overwritten by it.
    return _save_workbook_bytes(
        wb, apply_accountant_row_heights=True,
        skip_autofit_titles=frozenset({"Legacy Reconciliation", "Exceptions"}),
        suppress_text_number_warnings=True,
    )


def build_product_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    """Decision-oriented, not just a quantity/value total that "does not
    affect matching" and stops there: JE Support, Review Hold, and Matched %
    by product, so a reviewer can see where a product's dollars actually
    sit. Falls back to the plain quantity/value view only when no product
    classification is possible at all (see build_product_disposition_summary
    in matching.py)."""
    ws = wb.create_sheet("Product Aggregate Summary")
    frame = result.product_disposition_summary
    selected_period = result.metadata.get("fiscal_period")
    period_scope = (
        f"Only primary QuickBooks rows from Period {int(selected_period):02d} are included."
        if selected_period is not None
        else "All primary QuickBooks fiscal periods are included."
    )
    if frame.empty and result.product_summary.empty:
        headers = ["Product Name", "Product Quantity", "Product Value"]
        end_col = len(headers)
        _write_title_band(ws, 1, 1, end_col, "PRODUCT AGGREGATE SUMMARY", NAVY)
        _write_caption_band(
            ws, 2, 1, end_col,
            "A QuickBooks quantity column, an amount column, and a recognizable product description "
            f"are not all mapped, so no product breakdown is available. Generated "
            f"{format_central_timestamp(result.run_timestamp)}.",
            NAVY,
        )
        _write_dataframe_values(ws, pd.DataFrame(columns=headers), 3, 1)
        _format_header(ws, 3, 1, end_col, NAVY, headers=headers, amount_columns={"Product Value"})
        _write_total_row(ws, 4, 1, end_col, {}, headers)
        for col, width in zip(range(1, end_col + 1), (34, 20, 20)):
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes = "A4"
        _prepare_sheet(ws, landscape=False)
        return

    headers = list(frame.columns)
    end_col = len(headers)
    _write_title_band(ws, 1, 1, end_col, "PRODUCT AGGREGATE SUMMARY | JE SUPPORT AND REVIEW HOLD BY PRODUCT", NAVY)
    _write_caption_band(
        ws, 2, 1, end_col,
        "Products are classified from the QuickBooks description column; this classification does not "
        f"influence data matching. {period_scope} Matched % = matched rows / QuickBooks rows for that "
        f"product. Generated {format_central_timestamp(result.run_timestamp)}.",
        NAVY,
    )
    _write_dataframe_values(ws, frame, 3, 1)
    _format_header(
        ws, 3, 1, end_col, NAVY, headers=headers,
        amount_columns={"JE Support Amount", "Review Hold Amount"},
        quantity_columns={"QuickBooks Rows", "Matched Rows", "JE Support Rows", "Review Hold Rows", "Duplicate Excluded Rows"},
    )
    last_row = 3 + len(frame)
    if len(frame):
        _format_body_block(ws, 4, last_row, 1, end_col, NAVY_LIGHT)
        _apply_number_formats(
            ws, headers, 4, last_row, 1,
            {"JE Support Amount", "Review Hold Amount"},
            {"QuickBooks Rows", "Matched Rows", "JE Support Rows", "Review Hold Rows", "Duplicate Excluded Rows"},
        )
    total_row = last_row + 1
    totals = {
        column: float(frame[column].sum())
        for column in (
            "QuickBooks Rows", "Matched Rows", "JE Support Rows", "JE Support Amount",
            "Review Hold Rows", "Review Hold Amount", "Duplicate Excluded Rows",
        )
    } if len(frame) else {}
    _write_total_row(ws, total_row, 1, end_col, totals, headers)
    matched_col = headers.index("Matched %") + 1
    matched_rows_letter = get_column_letter(headers.index("Matched Rows") + 1)
    qb_rows_letter = get_column_letter(headers.index("QuickBooks Rows") + 1)
    ws.cell(
        total_row, matched_col,
        f"=IFERROR(SUM({matched_rows_letter}4:{matched_rows_letter}{last_row})/"
        f"SUM({qb_rows_letter}4:{qb_rows_letter}{last_row}),0)" if len(frame) else 0,
    )
    for row in range(4, total_row + 1):
        ws.cell(row, matched_col).number_format = "0.0%"
    _set_widths(ws, 1, end_col, 3, total_row)
    ws.column_dimensions["A"].width = 34
    ws.freeze_panes = "A4"
    if len(frame):
        ws.auto_filter.ref = f"A3:{get_column_letter(end_col)}{last_row}"
    _prepare_sheet(ws, landscape=False)


def _autofit_workbook_rows(wb: Workbook) -> None:
    """
    Dynamically calculates and explicitly sets row heights based on text wrapping 
    and column widths. This forces Excel to auto-expand rows even when the 
    workbook opens in Protected View.
    """
    for ws in wb.worksheets:
        # Capture custom column widths to estimate text wrapping constraints
        col_widths = {}
        for col_letter, dim in ws.column_dimensions.items():
            col_widths[col_letter] = dim.width or 15

        # A merged cell wraps across the whole width of its merge, not the
        # width of its first column -- measured against that one column, a
        # long caption or title reads as a dozen wrapped lines and the row
        # balloons to several times its real height.
        merged_widths: dict[tuple[int, int], float] = {}
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row == merged_range.max_row:
                merged_widths[(merged_range.min_row, merged_range.min_col)] = sum(
                    col_widths.get(get_column_letter(col), 15)
                    for col in range(merged_range.min_col, merged_range.max_col + 1)
                )
        designed_heights = {**_FIXED_ROW_HEIGHTS.get(ws.title, {}), **getattr(ws, "_fixed_row_heights", {})}

        for row in ws.iter_rows():
            max_lines = 1
            row_idx = row[0].row

            # Row 2 is a controlled report caption band. Do not let its long
            # merged-cell text enter the generic wrapping calculation, which
            # can otherwise expand it to several times the requested height.
            if row_idx == 2:
                ws.row_dimensions[row_idx].height = _controlled_row_two_height(ws.title)
                continue

            fixed_height = designed_heights.get(row_idx)
            if fixed_height is not None:
                ws.row_dimensions[row_idx].height = fixed_height
                continue

            for cell in row:
                # A date renders as MM/DD/YYYY (10 characters), not as the
                # 19-character "YYYY-MM-DD HH:MM:SS" str() of the datetime --
                # measuring the latter wraps every date cell and inflates
                # the whole row.
                if isinstance(cell.value, (datetime, date)):
                    text = "00/00/0000"
                else:
                    text = str(cell.value) if cell.value is not None else ""
                # A formula's text is not what the cell displays, and its result
                # is unknown until Excel calculates -- measuring the formula
                # would wrap a one-line figure or link into a tall row.
                if not text or text.startswith("="):
                    continue

                # Estimate character capacity based on column width
                col_width = merged_widths.get((row_idx, cell.column)) or col_widths.get(cell.column_letter, 15)
                # Approximation: ~1.1 to 1.2 chars fit per Excel width unit (10pt font)
                chars_per_line = max(int(col_width * 1.1), 10)

                cell_lines = 0
                for line in text.split("\n"):
                    # Calculate how many times this line will wrap
                    cell_lines += max(1, (len(line) // chars_per_line) + 1)
                
                if cell_lines > max_lines:
                    max_lines = cell_lines

                # Enable wrap_text for cells taking up multiple lines
                if cell_lines > 1:
                    curr_align = cell.alignment
                    if not curr_align or not curr_align.wrap_text:
                        cell.alignment = Alignment(
                            horizontal=curr_align.horizontal if curr_align else "left",
                            vertical=curr_align.vertical if curr_align else "center",
                            wrap_text=True,
                            shrink_to_fit=curr_align.shrink_to_fit if curr_align else False,
                            indent=curr_align.indent if curr_align else 0
                        )

            # Assign row height. (15 points per line is standard padding)
            if max_lines > 1:
                ws.row_dimensions[row_idx].height = max_lines * 15
            elif row_idx in ws.row_dimensions:
                # Let Excel manage the single-line rows natively
                ws.row_dimensions[row_idx].height = None
                

# Rows whose height is a deliberate design value rather than something to be
# measured from their text. Row 1 is a single-line title band; measured
# against only its first (narrow) column it reads as 3-5 wrapped lines and
# balloons to 45-75pt. Legacy Reconciliation row 3 is the one-line legend.
# Builders mark their own designed rows (title bands, KPI cards, legends, the
# control strip) with fix_row_height; this table covers only what they do not.
_FIXED_ROW_HEIGHTS = {
    "Legacy Reconciliation": {1: 27, 3: 18},
    "Exceptions": {1: 27},
}


def _controlled_row_two_height(sheet_title: str) -> int:
    """Fixed row-2 caption-band height per sheet, overriding autofit."""
    if sheet_title == "Product Aggregate Summary":
        return 60
    if sheet_title == "Legacy Reconciliation":
        return 30
    return 15


def _apply_accountant_output_row_heights(wb: Workbook) -> None:
    """Apply the controlled row-two presentation required by the workpaper."""
    for ws in wb.worksheets:
        ws.row_dimensions[2].height = _controlled_row_two_height(ws.title)


def _save_workbook_bytes(
    wb: Workbook,
    *,
    apply_accountant_row_heights: bool = False,
    skip_autofit_titles: frozenset = frozenset(),
    suppress_text_number_warnings: bool = False,
) -> bytes:
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    _autofit_workbook_columns(wb, skip_titles=skip_autofit_titles)
    _autofit_workbook_rows(wb)  # Dynamically auto-expand the row heights
    # Apply fixed reporting requirements after autofit so they cannot be
    # overwritten by content-dependent height calculations.
    if apply_accountant_row_heights:
        _apply_accountant_output_row_heights(wb)
    buffer = io.BytesIO()
    wb.save(buffer)
    saved = buffer.getvalue()
    return _add_ignored_errors(saved) if suppress_text_number_warnings else saved


# Worksheet-XML children that must come after <ignoredErrors> in the schema.
_AFTER_IGNORED_ERRORS = (
    "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
    "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
)
_IGNORED_ERRORS_XML = (
    '<ignoredErrors><ignoredError sqref="A1:XFD1048576" numberStoredAsText="1"/></ignoredErrors>'
)


def _add_ignored_errors(xlsx_bytes: bytes) -> bytes:
    """Suppress Excel's green "number stored as text" triangles on every
    worksheet. Invoice, PO, and customer numbers are identifiers, so storing
    them as text is correct -- but openpyxl has no API for the worksheet's
    <ignoredErrors> element, so it is inserted into each sheet's XML at its
    schema-mandated position after the workbook is saved."""
    source = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    output = io.BytesIO()
    boundary = re.compile(r"<(?:%s)[\s>/]" % "|".join(_AFTER_IGNORED_ERRORS))
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", item.filename):
                xml = data.decode("utf-8")
                match = boundary.search(xml, xml.index("</sheetData>"))
                position = match.start() if match else xml.rindex("</worksheet>")
                data = (xml[:position] + _IGNORED_ERRORS_XML + xml[position:]).encode("utf-8")
            target.writestr(item, data)
    return output.getvalue()


def _apply_workbook_run_metadata(wb: Workbook, result: ReconciliationResult) -> None:
    central_timestamp = result.run_timestamp.astimezone(CENTRAL_TIMEZONE)
    excel_timestamp = central_timestamp.replace(tzinfo=None)
    display_timestamp = format_central_timestamp(central_timestamp)
    # Visible report captions carry the controlled Central timestamp. Core
    # properties use the same run time, while page headers/footers are blank so
    # Excel cannot surface a stale or locale-generated tag on printed sheets.
    wb.properties.created = excel_timestamp
    wb.properties.modified = excel_timestamp
    existing_description = wb.properties.description or ""
    wb.properties.description = (
        f"{existing_description} Generated {display_timestamp}. Run ID: {result.run_id}."
    ).strip()
    for ws in wb.worksheets:
        for section in (
            ws.oddHeader, ws.evenHeader, ws.firstHeader,
            ws.oddFooter, ws.evenFooter, ws.firstFooter,
        ):
            section.left.text = None
            section.center.text = None
            section.right.text = None


_POSTING_SUMMARY_END_COL = 8

# The four top-level dispositions, in this fixed order, everywhere the
# workbook shows them as a set -- the Posting Summary cards, the equation, and
# the color used for each on both.
_DISPOSITION_CARD_ORDER = (
    ("Matched", GREEN_LIGHT),
    ("True Unmatched", NAVY_LIGHT),
    ("Review Hold", AMBER),
    ("Duplicate Excluded", RED_LIGHT),
)


def _fiscal_period_consistency_warning(result: ReconciliationResult) -> Optional[str]:
    """None if the selected fiscal period looks consistent with the primary
    QuickBooks data, otherwise a one-sentence warning naming the mismatch --
    a large, unmissable banner is more useful before export than a silent
    misclassification (see the fiscal-period rule in matching.py)."""
    selected_period = result.metadata.get("fiscal_period")
    period_col = result.qb_mapping.get("period")
    if selected_period is None or not period_col or not len(result.qb_work):
        return None
    default_year = int(result.metadata.get("fiscal_year") or result.run_timestamp.year)
    periods = [
        parse_fiscal_period(value, default_year)[0]
        for value in result.qb_work[period_col]
    ]
    read = [period for period in periods if period is not None]
    if not read:
        return None
    share_selected = sum(1 for period in read if period == int(selected_period)) / len(read)
    if share_selected >= 0.5:
        return None
    from collections import Counter

    most_common_period, most_common_count = Counter(read).most_common(1)[0]
    return (
        f"Only {share_selected:.0%} of QuickBooks rows with a readable fiscal period belong to the "
        f"selected period PD-{int(selected_period):02d} -- most ({most_common_count:,} of {len(read):,}) are "
        f"PD-{most_common_period:02d}. Verify PD-{int(selected_period):02d} is the intended reporting period "
        "before relying on this run's period classifications."
    )


def build_posting_summary_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    """The landing page: everything a reviewer needs before opening any other
    sheet -- the selected fiscal period (impossible to miss), a single
    accounted-for-rows control, the posting equation, and one card per final
    disposition. Every number here is read straight from result.metrics /
    result.qb_dispositions, the same source every other sheet uses, so this
    page can never disagree with the detail behind it."""
    # Appended like every other sheet, then moved to the front by the caller
    # (build_primary_workbook) once every sheet exists -- wb.active is still
    # claimed by build_data_search_sheet's own "grab the default sheet" step,
    # so this cannot itself occupy index 0 without colliding with that.
    ws = wb.create_sheet("Posting Summary")
    end_col = _POSTING_SUMMARY_END_COL
    metrics = result.metrics
    control_ok = metrics["Control Status"] == "PASS"

    _write_title_band(ws, 1, 1, end_col, "QBO RECONCILIATION POSTING SUMMARY", NAVY)
    _write_caption_band(
        ws, 2, 1, end_col,
        f"Run ID: {result.run_id} | Generated {format_central_timestamp(result.run_timestamp)}. "
        "Every figure on this page is read from the same controls shown in the Analytics workbook's "
        "Executive Summary and this workbook's Unresolved Exceptions and Reconciliation Detail sheets.",
        NAVY,
    )

    # The selected fiscal period, impossible to miss.
    period_row = 4
    selected_period = result.metadata.get("fiscal_period")
    fiscal_year = result.metadata.get("fiscal_year")
    period_text = (
        f"CURRENT RECONCILIATION PERIOD: PD-{int(selected_period):02d}"
        + (f" - {int(fiscal_year)}" if fiscal_year else "")
        if selected_period is not None
        else "CURRENT RECONCILIATION PERIOD: NOT SELECTED"
    )
    ws.merge_cells(start_row=period_row, start_column=1, end_row=period_row, end_column=end_col)
    period_cell = ws.cell(period_row, 1, period_text)
    period_cell.font = Font(name=FONT_NAME, size=16, bold=True, color=WHITE)
    period_cell.fill = PatternFill("solid", fgColor=SLATE)
    period_cell.alignment = Alignment(horizontal="center", vertical="center")
    fix_row_height(ws, period_row, 30)
    next_row = period_row + 1

    warning = _fiscal_period_consistency_warning(result)
    if warning:
        ws.merge_cells(start_row=next_row, start_column=1, end_row=next_row, end_column=end_col)
        warning_cell = ws.cell(next_row, 1, f"⚠ {warning}")
        warning_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="9C6500")
        warning_cell.fill = PatternFill("solid", fgColor=NEUTRAL_GOLD_FILL)
        warning_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        fix_row_height(ws, next_row, 30)
        next_row += 1
    next_row += 1  # spacer

    # The 100%-disposition control headline -- every row is accounted for by
    # construction (see build_qb_dispositions); the control fails only if the
    # disposition ledger and the source population disagree.
    headline_row = next_row
    total_rows = int(metrics["QuickBooks Rows"])
    headline_text = f"{total_rows:,} of {total_rows:,} QBO rows accounted for — CONTROL: {metrics['Control Status']}"
    ws.merge_cells(start_row=headline_row, start_column=1, end_row=headline_row, end_column=end_col)
    headline_cell = ws.cell(headline_row, 1, headline_text)
    headline_cell.font = Font(name=FONT_NAME, size=15, bold=True, color=TEXT if control_ok else "9C0006")
    headline_cell.fill = PatternFill("solid", fgColor=GREEN_LIGHT if control_ok else RED_LIGHT)
    headline_cell.alignment = Alignment(horizontal="center", vertical="center")
    fix_row_height(ws, headline_row, 26)
    next_row = headline_row + 1

    # Secondary: the match rate, explicitly not the headline any more.
    rate_row = next_row
    rate_cell = ws.cell(
        rate_row, 1,
        f"Match rate (secondary measure): {metrics['QuickBooks Match Rate by Row']:.1%}",
    )
    ws.merge_cells(start_row=rate_row, start_column=1, end_row=rate_row, end_column=end_col)
    rate_cell.font = Font(name=FONT_NAME, size=9, italic=True, color=SLATE)
    rate_cell.alignment = Alignment(horizontal="center", vertical="center")
    fix_row_height(ws, rate_row, 15)
    next_row = rate_row + 2

    # The posting equation, front and center: source rows/dollars equal the
    # sum of the four final dispositions.
    def card_value(label: str, kind: str) -> float:
        return metrics[f"Final Disposition - {label} {kind}"]

    equation_row = next_row
    equation_text = (
        f"{total_rows:,} QBO Rows ({format_currency(metrics['QuickBooks Source Total'])})  =  "
        + "  +  ".join(
            f"{int(card_value(label, 'Rows')):,} {label} ({format_currency(card_value(label, 'Amount'))})"
            for label, _ in _DISPOSITION_CARD_ORDER
        )
    )
    ws.merge_cells(start_row=equation_row, start_column=1, end_row=equation_row, end_column=end_col)
    equation_cell = ws.cell(equation_row, 1, equation_text)
    equation_cell.font = Font(name=FONT_NAME_NUMERIC, size=10, bold=True, color=TEXT)
    equation_cell.fill = PatternFill("solid", fgColor=SLATE_LIGHT)
    equation_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, shrink_to_fit=True)
    equation_cell.border = _thin_border()
    fix_row_height(ws, equation_row, 24)
    next_row = equation_row + 2

    # One card per final disposition -- the same four categories, in the same
    # order and color, as the equation above.
    label_row, value_row = next_row, next_row + 1
    col = 1
    for label, fill_color in _DISPOSITION_CARD_ORDER:
        ws.merge_cells(start_row=label_row, start_column=col, end_row=label_row, end_column=col + 1)
        ws.merge_cells(start_row=value_row, start_column=col, end_row=value_row, end_column=col + 1)
        display_label = "Engine Proposed JE (True Unmatched)" if label == "True Unmatched" else label.upper()
        ws.cell(label_row, col, display_label)
        ws.cell(
            value_row, col,
            f"{int(card_value(label, 'Rows')):,} rows | {format_currency(card_value(label, 'Amount'))}",
        )
        for row in (label_row, value_row):
            for c in (col, col + 1):
                ws.cell(row, c).fill = PatternFill("solid", fgColor=fill_color)
                ws.cell(row, c).border = _thin_border()
        ws.cell(label_row, col).font = Font(name=FONT_NAME, size=9, bold=True, color=TEXT)
        ws.cell(value_row, col).font = Font(name=FONT_NAME, size=12, bold=True, color=TEXT)
        for row in (label_row, value_row):
            ws.cell(row, col).alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
        col += 2
    fix_row_height(ws, label_row, 15)
    fix_row_height(ws, value_row, 24)
    next_row = value_row + 2

    # Journal Entry Bridge: the engine's automated total is immutable and
    # never overwritten by a reviewer selection (see the Reviewer Disposition
    # columns on Unresolved Exceptions) -- it is bridged to what actually
    # posts through two purely additive manual adjustments. With no reviewer
    # overrides, both adjustments are zero and Final Approved JE = Engine
    # Proposed JE exactly.
    bridge_title_row = next_row
    ws.merge_cells(start_row=bridge_title_row, start_column=1, end_row=bridge_title_row, end_column=end_col)
    bridge_title_cell = ws.cell(bridge_title_row, 1, "JOURNAL ENTRY BRIDGE | ENGINE TOTAL TO FINAL APPROVED JE")
    bridge_title_cell.font = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
    bridge_title_cell.fill = PatternFill("solid", fgColor=SLATE)
    bridge_title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    fix_row_height(ws, bridge_title_row, 20)

    released_expr = _review_holds_released_expr()
    excluded_expr = _je_support_manual_exclusions_expr(result)
    bridge_rows = [
        ("Engine Proposed JE (automated TRUE_UNMATCHED total -- immutable)", metrics["Proposed JE Amount"]),
        ("+  Review Holds Released to JE (reviewer)", f"={released_expr}"),
        ("-  Manual JE Exclusions (reviewer, documented reason required)", f"=-({excluded_expr})"),
    ]
    bridge_value_cells: list[str] = []
    row = bridge_title_row + 1
    for label, value in bridge_rows:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col - 3)
        ws.merge_cells(start_row=row, start_column=end_col - 2, end_row=row, end_column=end_col)
        label_cell = ws.cell(row, 1, label)
        value_cell = ws.cell(row, end_col - 2, value)
        label_cell.font = Font(name=FONT_NAME, size=10, color=TEXT)
        value_cell.font = Font(name=FONT_NAME_NUMERIC, size=10, color=TEXT)
        value_cell.number_format = ACCOUNTING_CURRENCY_FORMAT
        for col in (1, end_col - 2):
            ws.cell(row, col).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
            ws.cell(row, col).border = _thin_border()
        label_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        value_cell.alignment = Alignment(horizontal="right", vertical="center")
        bridge_value_cells.append(f"{get_column_letter(end_col - 2)}{row}")
        fix_row_height(ws, row, 18)
        row += 1

    final_row = row
    ws.merge_cells(start_row=final_row, start_column=1, end_row=final_row, end_column=end_col - 3)
    ws.merge_cells(start_row=final_row, start_column=end_col - 2, end_row=final_row, end_column=end_col)
    final_label_cell = ws.cell(final_row, 1, "=  FINAL APPROVED JE")
    final_value_cell = ws.cell(final_row, end_col - 2, "=" + "+".join(bridge_value_cells))
    final_label_cell.font = Font(name=FONT_NAME, size=11, bold=True, color=TEXT)
    final_value_cell.font = Font(name=FONT_NAME_NUMERIC, size=12, bold=True, color=TEXT)
    final_value_cell.number_format = ACCOUNTING_CURRENCY_FORMAT
    for col in (1, end_col - 2):
        ws.cell(final_row, col).fill = PatternFill("solid", fgColor=GREEN_LIGHT)
        ws.cell(final_row, col).border = _total_border()
    final_label_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    final_value_cell.alignment = Alignment(horizontal="right", vertical="center")
    fix_row_height(ws, final_row, 22)
    next_row = final_row + 2

    nav_row = next_row
    ws.merge_cells(start_row=nav_row, start_column=1, end_row=nav_row, end_column=end_col)
    nav_cell = ws.cell(
        nav_row, 1,
        "Review Hold detail and reviewer actions: Unresolved Exceptions. Every source row: "
        "Reconciliation Detail. Full explanation of every reason code above: Reason Code Glossary.",
    )
    nav_cell.font = Font(name="Segoe UI", size=9, italic=True, color=SLATE)
    nav_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    fix_row_height(ws, nav_row, 26)

    for col in range(1, end_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build_reason_code_glossary_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    """Every short reason code shown anywhere in the workbook, its full
    audit-grade explanation, and whether the disposition it belongs to feeds
    the proposed JE -- the one place the long technical text lives now that
    the Unresolved Exceptions and Legacy sheets show only the short code plus
    a concise reason (see SHORT_REASON_CODES / REASON_CODE_GLOSSARY in
    matching.py)."""
    ws = wb.create_sheet("Reason Code Glossary")
    headers = ["Reason Code", "Disposition", "Feeds Proposed JE", "Full Explanation"]
    end_col = len(headers)
    _write_title_band(ws, 1, 1, end_col, "REASON CODE GLOSSARY", SLATE)
    _write_caption_band(
        ws, 2, 1, end_col,
        "Every reason code shown on Unresolved Exceptions, Reconciliation Detail, and Legacy "
        "Reconciliation, in full. Only TRUE_UNMATCHED rows feed the proposed journal entry.",
        SLATE,
    )
    disposition_by_code = {
        "NO_INFINIUM_CANDIDATE": "TRUE UNMATCHED", "PO_REUSE_UNSUPPORTED": "TRUE UNMATCHED",
        "DUPLICATE_EXCLUDED": "DUPLICATE EXCLUDED",
    }
    header_row, data_row = 3, 4
    records = []
    for short_code, description in sorted(REASON_CODE_GLOSSARY.items()):
        disposition = disposition_by_code.get(short_code, "REVIEW HOLD")
        records.append({
            "Reason Code": short_code,
            "Disposition": disposition,
            "Feeds Proposed JE": "Yes" if disposition == "TRUE UNMATCHED" else "No",
            "Full Explanation": description,
        })
    frame = pd.DataFrame(records, columns=headers)
    _write_dataframe_values(ws, frame, header_row, 1)
    _format_header(ws, header_row, 1, end_col, SLATE, headers=headers)
    last_row = data_row + len(frame) - 1
    _format_body_block(ws, data_row, last_row, 1, end_col, SLATE_LIGHT)
    for row in range(data_row, last_row + 1):
        ws.cell(row, 4).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _set_widths(ws, 1, end_col, header_row, last_row, maximum=30)
    ws.column_dimensions[get_column_letter(4)].width = 90
    ws.freeze_panes = f"A{data_row}"
    ws.print_title_rows = f"1:{header_row}"
    _prepare_sheet(ws, landscape=False)


def build_primary_workbook(result: ReconciliationResult) -> bytes:
    validate_match_references(result)
    wb = Workbook()
    wb.properties.creator = "Sales Reconciliation Application"
    wb.properties.title = f"Sales Reconciliation {result.run_id}"
    wb.properties.subject = "QuickBooks to Infinium reconciliation and journal-entry support"
    wb.properties.description = "Accounting workpaper generated from one controlled reconciliation run."
    build_data_search_sheet(wb, result)
    build_raw_data_sheet(wb, result)
    build_reconciliation_detail_sheet(wb, result)
    build_unresolved_sheet(wb, result)
    _link_reconciled_data_to_unresolved_exceptions(wb, result)
    build_product_sheet(wb, result)
    build_reason_code_glossary_sheet(wb, result)
    build_posting_summary_sheet(wb, result)
    # The landing page: moved to the very front now that every other sheet
    # (and the default sheet build_data_search_sheet claimed) exists.
    wb.move_sheet("Posting Summary", offset=-wb.sheetnames.index("Posting Summary"))
    wb.active = 0
    _apply_workbook_run_metadata(wb, result)
    return _save_workbook_bytes(wb, apply_accountant_row_heights=True, suppress_text_number_warnings=True)


def detailed_ledger_dataframe(result: ReconciliationResult) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    qb_display, inf_display, _ = _paired_display_frames(result)
    resolved_records = _resolve_paired_records_bulk(result)
    
    qb_disp_dicts = qb_display.to_dict("records")
    inf_disp_dicts = inf_display.to_dict("records")
    
    qb_id_map = result.qb_work[QB_ID].to_dict() if result.qb_work is not None else {}
    inf_id_map = result.inf_work[INF_ID].to_dict() if result.inf_work is not None else {}
    qb_sec_id_map = result.qb_secondary_work[QB_ID].to_dict() if result.qb_secondary_work is not None else {}
    inf_sec_id_map = result.inf_secondary_work[INF_ID].to_dict() if result.inf_secondary_work is not None else {}

    for position, record in enumerate(resolved_records):
        output = {
            "Run ID": result.run_id,
            "Section": record["Section"],
            "Match ID": record["Match ID"],
            "Match Ref.": record.get("Match Ref.", ""),
            "Match Result": record["Match Result"],
            "Referenced Match Ref.": record.get("Referenced Match Ref.", ""),
            "Confidence": record["Confidence"],
            "Group Sequence": record["Group Sequence"],
            "Assessment Explanation": record["Explanation"],
        }
        qidx, iidx = record["QB Index"], record["Infinium Index"]
        qb_scope = record.get("QB Record Scope")
        inf_scope = record.get("Infinium Record Scope")
        
        active_q_map = qb_sec_id_map if qb_scope == "Historical" else qb_id_map
        output["QB | Source Row ID"] = active_q_map.get(qidx) if qidx is not None else None
        
        for header, value in qb_disp_dicts[position].items():
            output[f"QB | {header}"] = value
            
        active_i_map = inf_sec_id_map if inf_scope == "Historical" else inf_id_map
        output["INF | Source Row ID"] = active_i_map.get(iidx) if iidx is not None else None
        
        for header, value in inf_disp_dicts[position].items():
            output[f"INF | {header}"] = value
            
        records.append(output)
    return pd.DataFrame(records)


DATA_SEARCH_COLUMNS = [
    "Status", "Match Ref.", "Match Type", "Referenced Match Ref.", "Duplicate Group ID",
    "QuickBooks Row ID", "QuickBooks PO", "QuickBooks Invoice", "QuickBooks Amount",
    "Infinium Row ID", "Infinium PO", "Infinium Invoice", "Infinium Amount",
    "Detail",
]

_DATA_SEARCH_STATUS_BY_SECTION = {
    "01 Matched": "Infinium Match",
    "01 Matched - Historical Clearance": "Infinium Match",
    "02 Unmatched QuickBooks": "Outstanding (On Accrual List)",
    "03 Unmatched Infinium": "Error",
    "04 Duplicate QuickBooks": "Duplicate",
    "05 Duplicate Infinium": "Duplicate",
    "06 Duplicate Review Hold QuickBooks": "Duplicate",
    "07 Duplicate Review Hold Infinium": "Duplicate",
    "08 Reference-Matched Amount Variance Review Hold": "Error",
    "09 Fuzzy Match Review Hold": "Fuzzy Match",
    REFERENCE_HOLD_SECTION: "Review Hold",
}

_DATA_SEARCH_MATCH_TYPE_SECTIONS = {
    "01 Matched", "01 Matched - Historical Clearance", "09 Fuzzy Match Review Hold",
}


def _data_search_status(record: dict[str, Any]) -> str:
    section = record.get("Section", "")
    if section == "02 Unmatched QuickBooks" and "Invalid" in str(record.get("Exception Cause") or ""):
        return "Error"
    return _DATA_SEARCH_STATUS_BY_SECTION.get(section, "Error")


def build_data_search_dataframe(result: ReconciliationResult) -> pd.DataFrame:
    """One row per QuickBooks/Infinium line across every reconciliation
    population, for a plain PO/Invoice lookup: what is this item's status,
    and if it's a match, what kind.

    Built from the same paired-row resolution used for the Detailed Match
    Ledger, so every row appearing anywhere in the reconciliation appears
    here exactly once (validate_reconciliation guarantees this). A row can
    also carry a Duplicate Group ID even when its own Status is something
    else -- the retained copy of a confirmed duplicate keeps its real
    Outstanding/Infinium Match status (it still feeds the accrual) but is
    still identifiable as part of that duplicate pair.
    """
    resolved_records = _resolve_paired_records_bulk(result)
    qb_display, inf_display, _ = _paired_display_frames(result)
    qb_disp_dicts = qb_display.to_dict("records")
    inf_disp_dicts = inf_display.to_dict("records")

    qb_id_map = result.qb_work[QB_ID].to_dict() if result.qb_work is not None else {}
    inf_id_map = result.inf_work[INF_ID].to_dict() if result.inf_work is not None else {}
    qb_sec_id_map = result.qb_secondary_work[QB_ID].to_dict() if result.qb_secondary_work is not None else {}
    inf_sec_id_map = result.inf_secondary_work[INF_ID].to_dict() if result.inf_secondary_work is not None else {}

    qb_po_header = result.qb_mapping.get("po")
    qb_invoice_header = result.qb_mapping.get("invoice")
    qb_amount_header = result.qb_mapping.get("amount")
    inf_po_header = result.inf_mapping.get("po")
    inf_invoice_header = result.inf_mapping.get("invoice")
    inf_amount_header = result.inf_mapping.get("amount")

    def duplicate_group_map(report: pd.DataFrame) -> dict[str, str]:
        if report is None or report.empty:
            return {}
        return dict(zip(report["Source Row ID"].astype(str), report["Duplicate Group ID"]))

    qb_dup_groups = duplicate_group_map(result.duplicate_analysis)
    inf_dup_groups = duplicate_group_map(result.infinium_duplicate_analysis)

    records: list[dict[str, Any]] = []
    for position, record in enumerate(resolved_records):
        qidx, iidx = record["QB Index"], record["Infinium Index"]
        qb_scope = record.get("QB Record Scope")
        inf_scope = record.get("Infinium Record Scope")
        active_q_map = qb_sec_id_map if qb_scope == "Historical" else qb_id_map
        active_i_map = inf_sec_id_map if inf_scope == "Historical" else inf_id_map
        qb_row_id = active_q_map.get(qidx) if qidx is not None else None
        inf_row_id = active_i_map.get(iidx) if iidx is not None else None
        qb_values = qb_disp_dicts[position]
        inf_values = inf_disp_dicts[position]

        section = record.get("Section", "")
        match_type = record.get("Match Result", "") if section in _DATA_SEARCH_MATCH_TYPE_SECTIONS else ""
        duplicate_group_id = (
            (qb_dup_groups.get(str(qb_row_id)) if qb_row_id is not None else None)
            or (inf_dup_groups.get(str(inf_row_id)) if inf_row_id is not None else None)
            or ""
        )

        records.append({
            "Status": _data_search_status(record),
            "Match Ref.": record.get("Match Ref.", ""),
            "Match Type": match_type,
            "Referenced Match Ref.": record.get("Referenced Match Ref.", ""),
            "Duplicate Group ID": duplicate_group_id,
            "QuickBooks Row ID": qb_row_id,
            "QuickBooks PO": qb_values.get(qb_po_header) if qb_po_header else None,
            "QuickBooks Invoice": qb_values.get(qb_invoice_header) if qb_invoice_header else None,
            "QuickBooks Amount": qb_values.get(qb_amount_header) if qb_amount_header else None,
            "Infinium Row ID": inf_row_id,
            "Infinium PO": inf_values.get(inf_po_header) if inf_po_header else None,
            "Infinium Invoice": inf_values.get(inf_invoice_header) if inf_invoice_header else None,
            "Infinium Amount": inf_values.get(inf_amount_header) if inf_amount_header else None,
            "Detail": record.get("Exception Cause") or record.get("Explanation") or "",
        })
    return pd.DataFrame(records, columns=DATA_SEARCH_COLUMNS)


def build_data_search_indexes(result: ReconciliationResult) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Split the joined Data Search table into one row-per-QuickBooks-row
    table and one row-per-Infinium-row table, for the two live search panels.

    Every QuickBooks row appears in exactly one row of the QB table, and
    every Infinium row appears in exactly one row of the Infinium table --
    same guarantee as the joined table, since it's just a filtered view of it.
    """
    joined = build_data_search_dataframe(result)
    qb_index = joined.loc[joined["QuickBooks Row ID"].notna(), [
        "QuickBooks Row ID", "QuickBooks PO", "QuickBooks Invoice", "QuickBooks Amount",
        "Status", "Match Ref.", "Match Type", "Referenced Match Ref.", "Duplicate Group ID",
        "Infinium Row ID", "Detail",
    ]].rename(columns={
        "QuickBooks Row ID": "Row ID", "QuickBooks PO": "PO", "QuickBooks Invoice": "Invoice",
        "QuickBooks Amount": "Amount", "Infinium Row ID": "Matched Infinium Row ID",
    }).reset_index(drop=True)
    inf_index = joined.loc[joined["Infinium Row ID"].notna(), [
        "Infinium Row ID", "Infinium PO", "Infinium Invoice", "Infinium Amount",
        "Status", "Match Ref.", "Match Type", "Referenced Match Ref.", "Duplicate Group ID",
        "QuickBooks Row ID", "Detail",
    ]].rename(columns={
        "Infinium Row ID": "Row ID", "Infinium PO": "PO", "Infinium Invoice": "Invoice",
        "Infinium Amount": "Amount", "QuickBooks Row ID": "Matched QuickBooks Row ID",
    }).reset_index(drop=True)
    return qb_index, inf_index


def _search_criteria_formula(search_cell_ref: str, column_range: str) -> str:
    """Boolean array: TRUE for every row when the search cell is blank
    (no filter applied), else TRUE only where that row contains the typed
    text (case-insensitive substring match)."""
    return f'(({search_cell_ref}="")+(({search_cell_ref}<>"")*ISNUMBER(SEARCH({search_cell_ref},{column_range}))))'


def _write_search_input(ws, label: str, label_row: int, input_row: int, start_col: int) -> str:
    ws.cell(label_row, start_col, label)
    ws.cell(label_row, start_col).font = Font(name=FONT_NAME, size=11, bold=True, color=TEXT)
    ws.merge_cells(start_row=label_row, start_column=start_col, end_row=label_row, end_column=start_col + 2)
    input_col = start_col + 3
    input_cell = ws.cell(input_row, input_col, "")
    ws.merge_cells(start_row=input_row, start_column=input_col, end_row=input_row, end_column=input_col + 2)
    for col in range(input_col, input_col + 3):
        cell = ws.cell(input_row, col)
        cell.fill = PatternFill("solid", fgColor=WHITE)
        cell.border = _thin_border()
        cell.font = Font(name=FONT_NAME, size=11, color=TEXT)
    ws.row_dimensions[input_row].height = 20
    return f"${get_column_letter(input_col)}${input_row}"


def _write_search_panel(
    ws, source_sheet_name: str, index: pd.DataFrame,
    start_col: int, header_color: str, panel_title: str,
    panel_header_row: int, column_header_row: int, data_row: int,
    po_cell_ref: str, invoice_cell_ref: str,
) -> None:
    end_col = start_col + len(index.columns) - 1
    _write_title_band(ws, panel_header_row, start_col, end_col, panel_title, header_color)
    for offset, header in enumerate(index.columns):
        ws.cell(column_header_row, start_col + offset, header)
    _format_header(
        ws, column_header_row, start_col, end_col, header_color,
        headers=list(index.columns), amount_columns={"Amount"},
    )

    if index.empty:
        ws.cell(data_row, start_col, "No rows to search.")
        return

    last_source_row = len(index) + 1  # row 1 on the source sheet is its header
    columns = list(index.columns)
    po_col_letter = get_column_letter(columns.index("PO") + 1)
    invoice_col_letter = get_column_letter(columns.index("Invoice") + 1)
    data_range = (
        f"'{source_sheet_name}'!A2:{get_column_letter(len(columns))}{last_source_row}"
    )
    po_range = f"'{source_sheet_name}'!{po_col_letter}2:{po_col_letter}{last_source_row}"
    invoice_range = f"'{source_sheet_name}'!{invoice_col_letter}2:{invoice_col_letter}{last_source_row}"
    criteria = (
        f"{_search_criteria_formula(po_cell_ref, po_range)}*"
        f"{_search_criteria_formula(invoice_cell_ref, invoice_range)}"
    )
    formula = (
        f'=IF(AND({po_cell_ref}="",{invoice_cell_ref}=""),'
        f'"Type a PO or Invoice # above to search",'
        # FILTER is a post-2016 "future function" -- Excel itself always
        # stores these internally with an _xlfn. prefix, and a file written
        # without it (as openpyxl does by default) reads as a malformed
        # formula, which is exactly what triggers Excel's "we found a
        # problem with some content" repair prompt on open.
        f'_xlfn.FILTER({data_range},{criteria},"No matching items found"))'
    )
    ws.cell(data_row, start_col, formula)
    _set_widths(ws, start_col, end_col, panel_header_row, data_row)
    for offset, header in enumerate(columns):
        if header in ("Detail",):
            ws.column_dimensions[get_column_letter(start_col + offset)].width = 46
        elif header in ("PO", "Invoice"):
            ws.column_dimensions[get_column_letter(start_col + offset)].width = 22


def build_data_search_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    """A live-search sheet: type a PO or Invoice # once, and matching
    QuickBooks and Infinium items spill in automatically via Excel's FILTER()
    dynamic array function -- no manual filtering. Requires Excel 365 or
    Excel 2021+ (or the free Excel for the web), since FILTER() is a dynamic-
    array function not available in older desktop Excel.
    """
    qb_index, inf_index = build_data_search_indexes(result)

    qb_source_ws = wb.create_sheet("Data Search QB Source")
    _write_dataframe_values(qb_source_ws, qb_index, 1, 1)
    qb_source_ws.sheet_state = "hidden"

    inf_source_ws = wb.create_sheet("Data Search INF Source")
    _write_dataframe_values(inf_source_ws, inf_index, 1, 1)
    inf_source_ws.sheet_state = "hidden"

    ws = wb.active
    ws.title = "Data Search"

    qb_cols, inf_cols = len(qb_index.columns), len(inf_index.columns)
    qb_start, qb_end = 1, qb_cols
    inf_start = qb_end + 2
    inf_end = inf_start + inf_cols - 1
    end_col = inf_end

    _write_title_band(ws, 1, 1, end_col, "DATA SEARCH | TYPE A PO OR INVOICE NUMBER", NAVY)
    _write_caption_band(
        ws, 2, 1, end_col,
        "Type an Invoice # or PO # below (either one, or both to narrow the results) and "
        "matching QuickBooks and Infinium items appear automatically. Partial text matches "
        "count, so a few digits are enough. Status shows whether an item is outstanding on the "
        "accrual list, an error, a duplicate, a fuzzy match, or matched to Infinium; Match Type "
        "shows how a match was made. Duplicate Group ID is populated even on a duplicate pair's "
        "retained (accrual) copy, so both members of a pair stay traceable together. Requires "
        "Excel 365 / Excel 2021+ for the live results (FILTER is a dynamic-array function).",
        NAVY,
    )

    invoice_label_row = invoice_input_row = 4
    po_label_row = po_input_row = 5
    invoice_cell_ref = _write_search_input(ws, "Enter Invoice #", invoice_label_row, invoice_input_row, 1)
    po_cell_ref = _write_search_input(ws, "Enter PO #", po_label_row, po_input_row, 1)

    panel_header_row = 7
    column_header_row = 8
    data_row = 9

    _write_search_panel(
        ws, "Data Search QB Source", qb_index, qb_start, NAVY, "QUICKBOOKS ITEMS",
        panel_header_row, column_header_row, data_row, po_cell_ref, invoice_cell_ref,
    )
    _write_search_panel(
        ws, "Data Search INF Source", inf_index, inf_start, TEAL, "INFINIUM ITEMS",
        panel_header_row, column_header_row, data_row, po_cell_ref, invoice_cell_ref,
    )
    ws.column_dimensions[get_column_letter(qb_end + 1)].width = 3.5
    ws.freeze_panes = f"A{data_row}"
    _prepare_sheet(ws)


def _add_standard_data_sheet(
    wb: Workbook,
    name: str,
    title: str,
    caption: str,
    frame: pd.DataFrame,
    header_color: str = NAVY,
    chart_column: Optional[str] = None,
) -> Any:
    ws = wb.create_sheet(name)
    end_col = max(len(frame.columns), 1)
    _write_title_band(ws, 1, 1, end_col, title, header_color)
    _write_caption_band(ws, 2, 1, end_col, caption, header_color)
    if frame.empty and len(frame.columns) == 0:
        frame = pd.DataFrame({"Result": ["No records"]})
        end_col = 1
    _write_dataframe_values(ws, frame, 3, 1)
    _format_header(ws, 3, 1, len(frame.columns), header_color)
    if len(frame):
        _format_body_block(ws, 4, 3 + len(frame), 1, len(frame.columns), NAVY_LIGHT)
        _apply_number_formats(ws, list(frame.columns), 4, 3 + len(frame), 1, set(), set())
    _set_widths(ws, 1, len(frame.columns), 3, max(3 + len(frame), 3), maximum=48)
    wrap_terms = ("EXPLANATION", "REQUIREMENT", "DISPOSITION", "MATCH RESULT", "CANDIDATE IDS")
    for col, header in enumerate(frame.columns, 1):
        if any(term in str(header).upper() for term in wrap_terms):
            ws.column_dimensions[get_column_letter(col)].width = 42
            for row in range(4, 4 + len(frame)):
                ws.cell(row, col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.freeze_panes = "A4"
    if len(frame):
        ws.auto_filter.ref = f"A3:{get_column_letter(len(frame.columns))}{3 + len(frame)}"
    if chart_column and chart_column in frame.columns and len(frame):
        data_col = list(frame.columns).index(chart_column) + 1
        category_col = 1
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = f"{chart_column} by Match Method"
        chart.y_axis.title = "Match Method"
        chart.x_axis.title = chart_column
        chart.height = 7.5
        chart.width = 15
        data = Reference(ws, min_col=data_col, min_row=3, max_row=3 + len(frame))
        categories = Reference(ws, min_col=category_col, min_row=4, max_row=3 + len(frame))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.legend = None
        ws.add_chart(chart, f"{get_column_letter(len(frame.columns) + 2)}3")
    _prepare_sheet(ws)
    return ws


def build_analytics_summary_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.active
    ws.title = "Executive Summary"
    _write_title_band(ws, 1, 1, 8, "RECONCILIATION ANALYTICS | EXECUTIVE SUMMARY", NAVY)
    _write_caption_band(
        ws, 2, 1, 8,
        f"Run ID: {result.run_id} | Generated {format_central_timestamp(result.run_timestamp)} | Control status: {result.metrics['Control Status']}",
        NAVY,
    )
    cards = [
        ("QuickBooks rows", result.metrics["QuickBooks Rows"], '#,##0'),
        ("Infinium rows", result.metrics["Infinium Rows"], '#,##0'),
        ("QB match rate", result.metrics["QuickBooks Match Rate by Row"], '0.0%'),
        ("Proposed JE (true unmatched)", result.metrics["Proposed JE Amount"], '$#,##0.00;[Red]($#,##0.00);-'),
        ("QB source total", result.metrics["QuickBooks Source Total"], '$#,##0.00;[Red]($#,##0.00);-'),
        ("Infinium source total", result.metrics["Infinium Source Total"], '$#,##0.00;[Red]($#,##0.00);-'),
        ("Matched amount difference", result.metrics["Matched Amount Difference"], '$#,##0.00;[Red]($#,##0.00);-'),
        ("Control status", result.metrics["Control Status"], '@'),
    ]
    for idx, (label, value, number_format) in enumerate(cards):
        row = 4 if idx < 4 else 7
        col = 1 + (idx % 4) * 2
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
        ws.cell(row, col, label)
        ws.cell(row + 1, col, excel_safe(value))
        for r in (row, row + 1):
            for c in (col, col + 1):
                ws.cell(r, c).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
                ws.cell(r, c).border = _thin_border()
        ws.cell(row, col).font = Font(name=FONT_NAME, size=9, bold=True, color=SLATE)
        ws.cell(row + 1, col).font = Font(
            name=FONT_NAME, size=13, bold=True,
            color=NAVY if value != "FAIL" else "9C0006",
        )
        ws.cell(row + 1, col).number_format = number_format

    start = 11
    posting_status = result.metrics.get("Posting Status", "READY TO POST")
    ws.cell(start, 1, "MODEL STATUS")
    ws.cell(start, 2, result.metrics["Control Status"])
    ws.cell(start, 1).font = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
    ws.cell(start, 1).fill = PatternFill("solid", fgColor=SLATE)
    ws.cell(start, 2).font = Font(name=FONT_NAME, size=11, bold=True, color=TEXT)
    ws.cell(start, 2).fill = PatternFill("solid", fgColor=GREEN_LIGHT if result.metrics["Control Status"] == "PASS" else RED_LIGHT)
    ws.cell(start, 3, "POSTING STATUS")
    ws.cell(start, 4, posting_status)
    ws.cell(start, 3).font = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
    ws.cell(start, 3).fill = PatternFill("solid", fgColor=SLATE)
    ws.cell(start, 4).font = Font(name=FONT_NAME, size=11, bold=True, color=TEXT)
    ws.cell(start, 4).fill = PatternFill(
        "solid", fgColor=GREEN_LIGHT if posting_status == "READY TO POST" else AMBER
    )
    _write_dataframe_values(ws, result.controls, start + 2, 1)
    _format_header(ws, start + 2, 1, len(result.controls.columns), SLATE)
    _format_body_block(ws, start + 3, start + 2 + len(result.controls), 1, len(result.controls.columns), SLATE_LIGHT)
    _apply_number_formats(ws, list(result.controls.columns), start + 3, start + 2 + len(result.controls), 1, set(), set())
    status_col = list(result.controls.columns).index("Status") + 1
    if len(result.controls):
        green = PatternFill("solid", fgColor=GREEN_LIGHT)
        red = PatternFill("solid", fgColor=RED_LIGHT)
        ws.conditional_formatting.add(
            f"{get_column_letter(status_col)}{start + 3}:{get_column_letter(status_col)}{start + 2 + len(result.controls)}",
            FormulaRule(formula=[f'{get_column_letter(status_col)}{start + 3}="PASS"'], fill=green),
        )
        ws.conditional_formatting.add(
            f"{get_column_letter(status_col)}{start + 3}:{get_column_letter(status_col)}{start + 2 + len(result.controls)}",
            FormulaRule(formula=[f'{get_column_letter(status_col)}{start + 3}="FAIL"'], fill=red),
        )
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.freeze_panes = "A3"
    _prepare_sheet(ws)


def build_analytics_workbook(result: ReconciliationResult) -> bytes:
    validate_match_references(result)
    wb = Workbook()
    wb.properties.creator = "Sales Reconciliation Application"
    wb.properties.title = f"Sales Reconciliation Analytics {result.run_id}"
    wb.properties.subject = "Detailed matching evidence and reconciliation controls"
    build_analytics_summary_sheet(wb, result)
    _add_standard_data_sheet(
        wb, "Match Method Summary", "MATCH METHOD SUMMARY",
        "Distribution of matched and unresolved rows. Percentages use QuickBooks row count as the denominator.",
        result.method_summary, NAVY, chart_column="QuickBooks Rows",
    )
    _add_standard_data_sheet(
        wb, "Match Register", "MATCH REGISTER",
        "One row per accepted match relationship -- M-### for a one-to-one match, G-### for a grouped "
        "match accepted as a whole -- with the exact QuickBooks and Infinium records it contains and its "
        "amounts. Every Match Ref. and Referenced Match Ref. shown anywhere in the generated workbooks "
        "is listed here.",
        result.match_register, SLATE,
    )
    _add_standard_data_sheet(
        wb, "QB Disposition Ledger", "QUICKBOOKS DISPOSITION LEDGER",
        "Every QuickBooks source row appears exactly once with one final disposition -- MATCHED, "
        "EXACT_QBO_DUPLICATE_EXCLUDED, REVIEW_HOLD, or TRUE_UNMATCHED -- and the precise reason. Only "
        "TRUE_UNMATCHED rows feed the proposed journal entry; the row and dollar totals of these "
        "dispositions are proven equal to the source in the Executive Summary controls.",
        result.qb_dispositions, SLATE,
    )
    _add_standard_data_sheet(
        wb, "Detailed Match Ledger", "DETAILED MATCH Ledger",
        "One record per displayed reconciliation line. Historical clearances show the specific accepted prior-period row on the opposing side; unused historical rows are excluded.",
        detailed_ledger_dataframe(result), SLATE,
    )
    _add_standard_data_sheet(
        wb, "Normalization Detail", "NORMALIZATION DETAIL",
        "Original source values and the exact normalized references and signed amounts considered by the matching engine.",
        result.normalization, TEAL,
    )
    _add_standard_data_sheet(
        wb, "Match Assessment", "MATCH ASSESSMENT",
        "One row per accepted match group or unresolved QuickBooks decision, including criteria and evidence.",
        result.assessments, NAVY,
    )
    _add_standard_data_sheet(
        wb, "Historical Clearances", "SECONDARY HISTORICAL CLEARANCES",
        "Accepted strict matches between an unresolved primary row and the opposing historical source. Unused secondary rows are intentionally excluded.",
        result.historical_clearances, TEAL,
    )
    _add_standard_data_sheet(
        wb, "Exception Analysis", "EXCEPTION ANALYSIS",
        "Unresolved population summarized by source period, reason, and source ledger. Period is reporting metadata only.",
        result.exception_analysis, NAVY,
    )
    _add_standard_data_sheet(
        wb, "QuickBooks Duplicates", "QUICKBOOKS DUPLICATES",
        "Every QuickBooks row excluded from matching and from the accrual/journal-entry total because it "
        "shares an identical normalized PO, invoice, and signed amount with another QuickBooks row.",
        result.duplicate_analysis, NAVY,
    )
    _add_standard_data_sheet(
        wb, "Infinium Duplicates", "INFINIUM DUPLICATES",
        "Every Infinium row excluded from matching because it shares an identical normalized PO, invoice, "
        "and signed amount with another Infinium row. Kept on a separate worksheet from QuickBooks "
        "duplicates because the treatment is entirely different: Infinium duplicates are reviewed "
        "independently and never feed the QuickBooks accrual or journal entry.",
        result.infinium_duplicate_analysis, TEAL,
    )
    _add_standard_data_sheet(
        wb, "Rules and Run Config", "RULES AND RUN CONFIGURATION",
        "Exact rules, file fingerprints, mappings, versions, and parameters used for this run.",
        pd.concat(
            [
                result.config,
                pd.DataFrame([{"Setting": "", "Value": ""}, {"Setting": "MATCHING RULES", "Value": ""}]),
                result.rules.rename(columns={
                    "Priority": "Setting",
                    "Rule": "Value",
                    "Automatic": "Automatic",
                    "Requirement": "Requirement",
                }),
            ],
            ignore_index=True,
        ),
        SLATE,
    )
    _apply_workbook_run_metadata(wb, result)
    return _save_workbook_bytes(wb)


# Public alias: ui_components.py needs the paired display frames to build the
# in-app "Reconciled View" preview without duplicating this logic.
paired_display_frames = _paired_display_frames