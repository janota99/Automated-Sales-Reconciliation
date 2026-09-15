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
from datetime import datetime
from decimal import Decimal
from typing import Any, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from config import (
    AMBER,
    CENTRAL_TIMEZONE,
    GREEN_LIGHT,
    NAVY,
    NAVY_LIGHT,
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
    _apply_good_style,
    _apply_method_style,
    _apply_neutral_style,
    _apply_number_formats,
    _autofit_workbook_columns,
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
)
from duplicates import (
    DUPLICATE_BASIS_CROSS_SCOPE,
    DUPLICATE_BASIS_INVOICE_ONLY,
    DUPLICATE_BASIS_PO_ONLY,
    DUPLICATE_BASIS_STRICT,
    NORM_PO,
)
from matching import (
    AMOUNT_CENTS,
    INF_ID,
    QB_ID,
    ReconciliationResult,
    build_fiscal_exception_summary,
    cents_or_zero,
    cents_to_float,
    numeric_quantity_sum,
    numeric_sum,
    parse_fiscal_period,
    valid_cents,
)
from utils import excel_safe, format_central_timestamp, format_currency


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

    _write_title_band(ws, 1, qb_start, qb_end, "QUICKBOOKS | RAW TRANSACTION DETAIL", NAVY)
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
    _format_header(ws, header_row, qb_start, qb_end, NAVY)
    _format_header(ws, header_row, inf_start, inf_end, TEAL)
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
                          {result.qb_mapping["amount"]}, {result.qb_mapping.get("quantity") or ""})
    _apply_number_formats(ws, inf_headers, data_row, inf_total_row, inf_start,
                          {result.inf_mapping["amount"]}, set())
    ws.column_dimensions[get_column_letter(separator_col)].width = 3.5
    ws.column_dimensions[get_column_letter(separator_col)].fill = PatternFill("solid", fgColor=WHITE)
    _set_widths(ws, qb_start, qb_end, header_row, qb_total_row)
    _set_widths(ws, inf_start, inf_end, header_row, inf_total_row)
    ws.freeze_panes = f"{get_column_letter(inf_start)}{data_row}"
    ws.print_title_rows = "1:3"
    _prepare_sheet(ws)


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
        qb_values.append(
            "QuickBooks Prior Period Match"
            if qb_scope == "Historical" else "Primary QuickBooks Upload" if qidx is not None else None
        )
        inf_values.append(
            "Infinium Prior Period Match"
            if inf_scope == "Historical" else "Primary Infinium Upload" if iidx is not None else None
        )
        qb_rows.append(qb_values)
        inf_rows.append(inf_values)
        match_results.append(record["Match Result"])
    return (
        pd.DataFrame(qb_rows, columns=qb_headers + [qb_context_header]),
        pd.DataFrame(inf_rows, columns=inf_headers + [inf_context_header]),
        match_results,
    )


# Reconciled Data writes exactly one row per result.paired_rows entry, in
# that same order, starting at this row. Other sheets (Unresolved
# Exceptions) rely on this exact constant to compute a hyperlink target
# without re-deriving Reconciled Data's own layout -- keep them in sync.
RECONCILED_DATA_HEADER_ROW = 3
RECONCILED_DATA_DATA_ROW = 4


def _qb_id_reconciled_data_row_map(result: ReconciliationResult) -> dict[str, int]:
    """Map every QuickBooks Row ID to the row it occupies on Reconciled
    Data, so another sheet can link straight to where a row was originally
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
        mapping.setdefault(str(qb_id), RECONCILED_DATA_DATA_ROW + offset)
    return mapping


def _apply_row_id_hyperlink(ws, row: int, col: int, target_row: Optional[int]) -> None:
    """Turn a Row ID cell into a link straight to that row on Reconciled
    Data -- if a target couldn't be resolved (e.g. an Infinium-only row,
    which Reconciled Data still lists but this map doesn't cover), leave
    the cell as plain text rather than link to nothing.

    Only adds underline to whatever font is already on the cell rather
    than replacing it outright -- a duplicate-excluded row's Row ID must
    stay visibly red, not turn hyperlink-blue and lose that signal, while
    still being clickable.
    """
    if target_row is None:
        return
    cell = ws.cell(row, col)
    cell.hyperlink = f"#'Reconciled Data'!A{target_row}"
    current = cell.font
    cell.font = Font(
        name=current.name or "Segoe UI",
        size=current.size or 10,
        bold=current.bold,
        color=current.color or NAVY,
        underline="single",
    )


def build_reconciled_data_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.create_sheet("Reconciled Data")
    qb_display, inf_display, match_results = _paired_display_frames(result)
    qb_headers = list(qb_display.columns)
    inf_headers = list(inf_display.columns)
    qb_start = 1
    match_col = len(qb_headers) + 1
    inf_start = match_col + 1
    qb_end = len(qb_headers)
    inf_end = inf_start + len(inf_headers) - 1
    header_row, data_row = RECONCILED_DATA_HEADER_ROW, RECONCILED_DATA_DATA_ROW
    final_data_row = data_row + len(match_results) - 1

    _write_title_band(ws, 1, qb_start, qb_end, "QUICKBOOKS | RECONCILED", NAVY)
    _write_title_band(ws, 1, match_col, match_col, "MATCH RESULT", SLATE)
    _write_title_band(ws, 1, inf_start, inf_end, "INFINIUM | RECONCILED", TEAL)
    _write_caption_band(
        ws, 2, qb_start, qb_end,
        f"Every primary QuickBooks record appears once. Any accepted QuickBooks prior-period match is displayed on this side and labeled in Record Context. Generated {format_central_timestamp(result.run_timestamp)}.",
        NAVY,
    )
    _write_caption_band(ws, 2, match_col, match_col, "Matching Methodology", SLATE)
    _write_caption_band(
        ws, 2, inf_start, inf_end,
        "Every primary Infinium record appears once. Accepted prior-period matches are displayed; unused historical rows are excluded.",
        TEAL,
    )
    _write_dataframe_values(ws, qb_display, header_row, qb_start)
    ws.cell(header_row, match_col, "Match Result")
    for offset, value in enumerate(match_results, 1):
        ws.cell(header_row + offset, match_col, value)
    _write_dataframe_values(ws, inf_display, header_row, inf_start)
    _format_header(ws, header_row, qb_start, qb_end, NAVY)
    _format_header(ws, header_row, match_col, match_col, SLATE)
    _format_header(ws, header_row, inf_start, inf_end, TEAL)
    _format_body_block(ws, data_row, final_data_row, qb_start, qb_end, NAVY_LIGHT)
    _format_body_block(ws, data_row, final_data_row, match_col, match_col, SLATE_LIGHT)
    _format_body_block(ws, data_row, final_data_row, inf_start, inf_end, TEAL_LIGHT)

    duplicate_qb_rows = _duplicate_source_indexes(result, "QuickBooks")
    duplicate_inf_rows = _duplicate_source_indexes(result, "Infinium")
    resolved_records = _resolve_paired_records_bulk(result)
    for offset, record in enumerate(resolved_records):
        row = data_row + offset
        status = str(record["Section"])
        if status == "02 Unmatched QuickBooks":
            for col in range(qb_start, match_col + 1):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=AMBER)
        elif status == "03 Unmatched Infinium":
            for col in range(match_col, inf_end + 1):
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
                name="Segoe UI", size=10, bold=True, color=NAVY
            )
        if record.get("Infinium Record Scope") == "Historical":
            ws.cell(row, inf_end).font = Font(
                name="Segoe UI", size=10, bold=True, color=TEAL
            )
        ws.cell(row, match_col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    total_row = final_data_row + 1
    _write_total_row(ws, total_row, qb_start, qb_end,
                     _source_totals(result.qb_raw, result.qb_mapping), qb_headers, "RECONCILED TOTAL")
    _write_total_row(ws, total_row, inf_start, inf_end,
                     _source_totals(result.inf_raw, result.inf_mapping), inf_headers, "RECONCILED TOTAL")
    ws.cell(total_row, match_col, f"Control: {result.metrics['Control Status']}")
    ws.cell(total_row, match_col).fill = PatternFill("solid", fgColor=GREEN_LIGHT if result.metrics["Control Status"] == "PASS" else RED_LIGHT)
    ws.cell(total_row, match_col).font = Font(name="Segoe UI", size=10, bold=True, color=TEXT)
    ws.cell(total_row, match_col).border = _total_border()
    ws.cell(total_row, match_col).alignment = Alignment(horizontal="center", vertical="center")
    _apply_number_formats(ws, qb_headers, data_row, total_row, qb_start,
                          {result.qb_mapping["amount"]}, {result.qb_mapping.get("quantity") or ""})
    _apply_number_formats(ws, inf_headers, data_row, total_row, inf_start,
                          {result.inf_mapping["amount"]}, set())
    _set_widths(ws, qb_start, qb_end, header_row, total_row)
    _set_widths(ws, inf_start, inf_end, header_row, total_row)
    ws.column_dimensions[get_column_letter(match_col)].width = 43
    ws.freeze_panes = f"{get_column_letter(inf_start)}{data_row}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(inf_end)}{final_data_row}"
    ws.print_title_rows = "1:3"
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


def build_unresolved_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.create_sheet("Unresolved Exceptions")

    # QuickBooks is the sole accrual and journal-entry basis, so this sheet
    # keeps everything with accrual relevance -- the raw QuickBooks exception
    # population on the left, and the QuickBooks items excluded from that
    # same population as duplicates on the right. Infinium exceptions carry
    # no accrual impact and are already listed in Reconciled Data, so they
    # are intentionally not repeated here.
    source_headers = list(result.qb_raw.columns)
    headers = source_headers + ["Exception Status", "Reference Amount Difference", "Reviewer Note"]
    candidate_map = result.candidates.set_index("QuickBooks Row ID").to_dict("index") if not result.candidates.empty else {}
    qb_id_row_map = _qb_id_reconciled_data_row_map(result)

    qb_subset_dict = result.qb_work.loc[result.unmatched_qb].to_dict("index")
    records = []
    for qidx in result.unmatched_qb:
        row_data = qb_subset_dict[qidx]
        candidate = candidate_map.get(row_data[QB_ID], {})
        records.append(
            [row_data.get(col) for col in source_headers]
            + [
                candidate.get("Disposition", "Unmatched QuickBooks"),
                candidate.get("Minimum Amount Difference"),
                "",
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

    review_hold_frame = result.duplicate_analysis.loc[
        result.duplicate_analysis["Disposition"] == "Held for review - excluded from proposed JE pending disposition"
    ].copy()
    review_hold_frame = _simplify_duplicate_display(review_hold_frame)
    review_hold_frame["Reviewer Disposition"] = "Pending Review"
    review_headers = list(review_hold_frame.columns)
    review_end_col = len(review_headers)
    review_hold_count = result.metrics["Duplicate Review Hold QuickBooks Rows"]
    review_hold_amount = result.metrics["Duplicate Review Hold QuickBooks Amount"]

    unmatched_qb_amounts = result.qb_work.loc[result.unmatched_qb, AMOUNT_CENTS].tolist()
    amounts = [cents_or_zero(val) for val in unmatched_qb_amounts]
    net = sum(amounts)

    qb_table_name = "QuickBooksExceptions"
    qb_amount_column_expr = _table_column_reference(qb_table_name, result.qb_mapping["amount"])
    qb_amount_sum_expr = f"SUM({qb_amount_column_expr})"

    _write_title_band(ws, 1, 1, end_col, "QUICKBOOKS EXCEPTIONS | JOURNAL ENTRY SUPPORT", NAVY)
    invalid_unresolved = sum(
        1 for idx in result.unmatched_qb if not valid_cents(result.qb_work.at[idx, AMOUNT_CENTS])
    )
    _write_caption_band(
        ws, 2, 1, end_col,
        f"QuickBooks is the sole accrual and proposed-journal-entry basis. Net signed support excludes "
        f"{invalid_unresolved} row(s) with invalid or missing amounts. Review every exception before posting. "
        f"Generated {format_central_timestamp(result.run_timestamp)}.",
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
        ("Unresolved rows", f"=IFERROR(ROWS({qb_amount_column_expr}),0)", '#,##0'),
        (
            "Gross debits",
            f'=SUMIF({qb_amount_column_expr},">0",{qb_amount_column_expr})',
            '$#,##0.00;[Red]($#,##0.00);-',
        ),
        (
            "Credits",
            f'=ABS(SUMIF({qb_amount_column_expr},"<0",{qb_amount_column_expr}))',
            '$#,##0.00;[Red]($#,##0.00);-',
        ),
        (
            "Proposed JE support total",
            f"={qb_amount_sum_expr}",
            '$#,##0.00;[Red]($#,##0.00);-',
        ),
    ]
    for idx, (label, value, number_format) in enumerate(kpis):
        start = 1 + idx * 2
        if start > end_col:
            break
        ws.cell(3, start, label)
        ws.cell(4, start, value)
        ws.cell(3, start).font = Font(name="Segoe UI", size=9, bold=True, color=SLATE)
        ws.cell(4, start).font = Font(name="Segoe UI", size=12, bold=True, color=NAVY)
        ws.cell(4, start).number_format = number_format
        ws.cell(4, start).protection = Protection(locked=True)
        for row in (3, 4):
            ws.cell(row, start).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
            ws.cell(row, start).border = _thin_border()

    # QuickBooks exceptions by fiscal period -- promoted above the detail
    # tables so period-level review (count and net amount per period) never
    # requires scrolling past the full exception and duplicate lists.
    fiscal_summary = build_fiscal_exception_summary(result)
    fiscal_headers = list(fiscal_summary.columns)
    fiscal_title_row = 7
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
            f"Selected current reporting period: PD-{int(selected_period):02d}. Every other valid QuickBooks fiscal "
            f"period is classified as a prior-period urgent exception. {quantity_note}"
            if has_fiscal_period and selected_period is not None
            else f"No current reporting period was selected. Exceptions are summarized by source period without current/prior classification. {quantity_note}"
            if has_fiscal_period
            else "No credible QuickBooks fiscal-period identifier was found or mapped. Period-based "
            f"classification is disabled and all exceptions are summarized together. {quantity_note}"
        ),
        NAVY,
    )
    _write_dataframe_values(ws, fiscal_summary, fiscal_header_row, 1)
    _format_header(ws, fiscal_header_row, 1, fiscal_end_col, NAVY)
    if len(fiscal_summary):
        fiscal_last_row = fiscal_data_row + len(fiscal_summary) - 1
        _format_body_block(ws, fiscal_data_row, fiscal_last_row, 1, fiscal_end_col, NAVY_LIGHT)
        for offset, classification in enumerate(fiscal_summary["Period Classification"], start=fiscal_data_row):
            fill = (
                RED_LIGHT
                if classification == "Prior-Period Urgent Exception"
                else GREEN_LIGHT
                if classification == "Current Reporting Period"
                else AMBER
            )
            for col in range(1, fiscal_end_col + 1):
                ws.cell(offset, col).fill = PatternFill("solid", fgColor=fill)
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
    _write_dataframe_values(ws, frame, header_row, 1)
    _format_header(ws, header_row, 1, end_col, NAVY)
    if len(frame):
        _format_body_block(ws, data_row, data_row + len(frame) - 1, 1, end_col, NAVY_LIGHT)
        duplicate_qb_rows = _duplicate_source_indexes(result, "QuickBooks")
        for offset, qidx in enumerate(result.unmatched_qb):
            row = data_row + offset
            ws.cell(row, source_headers.index(result.qb_mapping["amount"]) + 1).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            ws.cell(row, len(source_headers) + 1).fill = PatternFill("solid", fgColor=AMBER)
            ws.cell(row, len(source_headers) + 1).alignment = Alignment(wrap_text=True, vertical="center")
            ws.cell(row, len(source_headers) + 2).number_format = '$#,##0.00;[Red]($#,##0.00);-'
            if int(qidx) in duplicate_qb_rows:
                _apply_duplicate_style(ws, row, 1, end_col)
    total_row = data_row + len(frame)
    _write_total_row(
        ws, total_row, 1, end_col,
        {result.qb_mapping["amount"]: cents_to_float(net)}, headers,
        "PROPOSED JE SUPPORT TOTAL",
    )
    ws.cell(total_row, amount_col_position).number_format = '$#,##0.00;[Red]($#,##0.00);-'

    qb_summed_headers = {result.qb_mapping["amount"]}
    qb_quantity_header = result.qb_mapping.get("quantity")
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
    ws.column_dimensions[get_column_letter(len(source_headers) + 1)].width = 48
    ws.column_dimensions[get_column_letter(len(source_headers) + 2)].width = 24
    ws.column_dimensions[get_column_letter(len(source_headers) + 3)].width = 36
    _set_widths(ws, 1, len(source_headers), header_row, total_row)
    ws.freeze_panes = f"A{data_row}"
    if len(frame):
        note_col = get_column_letter(len(source_headers) + 3)
        validation = DataValidation(
            type="textLength", operator="lessThanOrEqual", formula1="1000", allow_blank=True
        )
        validation.error = "Reviewer notes are limited to 1,000 characters."
        validation.errorTitle = "Note too long"
        ws.add_data_validation(validation)
        validation.add(f"{note_col}{data_row}:{note_col}{header_row + len(frame)}")

    # QuickBooks duplicates excluded from the JE above, and the proposed JE
    # itself, are placed a fixed 10 rows below the exceptions total row --
    # far enough to read as clearly separate from the exception detail,
    # close enough to stay on the same review pass.
    duplicate_kpis = [
        ("Duplicate QuickBooks items excluded", duplicate_excluded_count, '#,##0'),
        ("Amount excluded from JE", duplicate_amount_total, '$#,##0.00;[Red]($#,##0.00);-'),
        ("JE inclusion", "Excluded", 'General'),
    ]
    dup_title_row = total_row + 10
    dup_caption_row = dup_title_row + 1
    dup_kpi_label_row = dup_title_row + 2
    dup_kpi_value_row = dup_title_row + 3
    dup_header_row = dup_title_row + 5
    dup_data_row = dup_header_row + 1
    section_end_col = max(end_col, dup_end_col, review_end_col)

    _write_title_band(
        ws, dup_title_row, 1, section_end_col,
        "QUICKBOOKS DUPLICATES EXCLUDED FROM JE | REVIEW", NAVY,
    )
    _write_caption_band(ws, dup_caption_row, 1, section_end_col, duplicate_caption, NAVY)
    for idx, (label, value, number_format) in enumerate(duplicate_kpis):
        start = 1 + idx * 2
        if start > dup_end_col:
            break
        ws.cell(dup_kpi_label_row, start, label)
        ws.cell(dup_kpi_value_row, start, value)
        ws.cell(dup_kpi_label_row, start).font = Font(name="Segoe UI", size=9, bold=True, color=SLATE)
        ws.cell(dup_kpi_value_row, start).font = Font(name="Segoe UI", size=12, bold=True, color=NAVY)
        ws.cell(dup_kpi_value_row, start).number_format = number_format
        ws.cell(dup_kpi_value_row, start).protection = Protection(locked=True)
        for row in (dup_kpi_label_row, dup_kpi_value_row):
            ws.cell(row, start).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
            ws.cell(row, start).border = _thin_border()

    _write_dataframe_values(ws, duplicate_frame, dup_header_row, 1)
    _format_header(ws, dup_header_row, 1, dup_end_col, NAVY)
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
        # Reconciled Data, so a reviewer never has to search for it by hand.
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

    # Duplicate Review Hold: weak-basis (PO-only or invoice-only) candidates
    # that stayed active for matching but never resolved. They are already
    # excluded from the JE support total above; this section is where a
    # human records what actually happens to them next.
    review_title_row = dup_last_row + 3
    review_caption_row = review_title_row + 1
    review_kpi_label_row = review_title_row + 2
    review_kpi_value_row = review_title_row + 3
    review_header_row = review_title_row + 5
    review_data_row = review_header_row + 1
    review_section_end_col = section_end_col

    review_caption = (
        f"{review_hold_count:,} weak-basis duplicate candidate(s) (sharing only a PO or only an "
        "invoice with another QuickBooks row at the same signed amount) remained unresolved after "
        "every matching pass. Rather than silently inflating the accrual, each is excluded from the "
        "JE support total above and held here pending a documented human decision -- confirm as a "
        "genuine duplicate, confirm as legitimate and add to the JE manually, or escalate for "
        "investigation -- before this journal entry is posted."
        if review_hold_count
        else "No QuickBooks weak-basis duplicate candidates remain unresolved."
    )
    _write_title_band(
        ws, review_title_row, 1, review_section_end_col,
        "DUPLICATE REVIEW HOLD | REQUIRES DOCUMENTED DISPOSITION BEFORE POSTING", SLATE,
    )
    _write_caption_band(ws, review_caption_row, 1, review_section_end_col, review_caption, SLATE)

    review_kpis = [
        ("Items held for review", review_hold_count, '#,##0'),
        ("Amount excluded from JE", review_hold_amount, '$#,##0.00;[Red]($#,##0.00);-'),
        ("JE inclusion", "Excluded pending disposition", 'General'),
    ]
    for idx, (label, value, number_format) in enumerate(review_kpis):
        start = 1 + idx * 2
        if start > review_end_col:
            break
        ws.cell(review_kpi_label_row, start, label)
        ws.cell(review_kpi_value_row, start, value)
        ws.cell(review_kpi_label_row, start).font = Font(name="Segoe UI", size=9, bold=True, color=SLATE)
        ws.cell(review_kpi_value_row, start).font = Font(name="Segoe UI", size=12, bold=True, color=NAVY)
        ws.cell(review_kpi_value_row, start).number_format = number_format
        ws.cell(review_kpi_value_row, start).protection = Protection(locked=True)
        for row in (review_kpi_label_row, review_kpi_value_row):
            ws.cell(row, start).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
            ws.cell(row, start).border = _thin_border()

    _write_dataframe_values(ws, review_hold_frame, review_header_row, 1)
    _format_header(ws, review_header_row, 1, review_end_col, SLATE)
    if len(review_hold_frame):
        review_last_row = review_data_row + len(review_hold_frame) - 1
        _format_body_block(ws, review_data_row, review_last_row, 1, review_end_col, SLATE_LIGHT)
        _apply_number_formats(
            ws, review_headers, review_data_row, review_last_row, 1,
            {"Amount"}, set(),
        )
        # Amber, not the red duplicate style: this needs attention, but it
        # is not yet a confirmed duplicate the way an excess copy is.
        for row in range(review_data_row, review_last_row + 1):
            for col in range(1, review_end_col + 1):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=AMBER)
        # Row ID links straight to where this row was originally listed on
        # Reconciled Data, same as the duplicates section above.
        review_row_id_col = review_headers.index("Row ID") + 1
        for offset, qb_id in enumerate(review_hold_frame["Row ID"]):
            _apply_row_id_hyperlink(
                ws, review_data_row + offset, review_row_id_col, qb_id_row_map.get(str(qb_id)),
            )
    else:
        review_last_row = review_header_row

    _set_widths(ws, 1, review_end_col, review_header_row, review_last_row)
    if "What Was Found" in review_headers:
        ws.column_dimensions[
            get_column_letter(review_headers.index("What Was Found") + 1)
        ].width = 52
    if "Reason" in review_headers:
        ws.column_dimensions[
            get_column_letter(review_headers.index("Reason") + 1)
        ].width = 46
    if "Reviewer Disposition" in review_headers:
        disposition_col = review_headers.index("Reviewer Disposition") + 1
        ws.column_dimensions[get_column_letter(disposition_col)].width = 40
        if len(review_hold_frame):
            disposition_validation = DataValidation(
                type="list",
                formula1='"Pending Review,Confirmed Duplicate - Exclude Permanently,'
                         'Confirmed Legitimate - Include In JE Manually,Escalated For Investigation"',
                allow_blank=False,
            )
            disposition_validation.error = "Select a disposition from the list before posting."
            disposition_validation.errorTitle = "Disposition required"
            ws.add_data_validation(disposition_validation)
            disposition_letter = get_column_letter(disposition_col)
            disposition_validation.add(
                f"{disposition_letter}{review_data_row}:{disposition_letter}{review_last_row}"
            )
            for row in range(review_data_row, review_last_row + 1):
                ws.cell(row, disposition_col).protection = Protection(locked=False)

    je_title_row = review_last_row + 3
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
        "017-91000-400000-0 Income-Manufacturing for the absolute unresolved net amount; evaluate "
        "reversals and negative source values before posting.",
        SLATE,
    )
    _write_dataframe_values(ws, je_frame, je_header_row, 1)
    je_amount_formula = f"=ABS({qb_amount_sum_expr})"
    ws.cell(je_data_row, 4, je_amount_formula)
    ws.cell(je_data_row, 5, 0.0)
    ws.cell(je_data_row + 1, 4, 0.0)
    ws.cell(je_data_row + 1, 5, je_amount_formula)
    _format_header(ws, je_header_row, 1, len(je_headers), SLATE)
    _format_body_block(ws, je_data_row, je_data_row + len(je_frame) - 1, 1, len(je_headers), SLATE_LIGHT)
    for row in range(je_data_row, je_data_row + len(je_frame)):
        ws.cell(row, 4).number_format = '$#,##0.00;[Red]($#,##0.00);-'
        ws.cell(row, 5).number_format = '$#,##0.00;[Red]($#,##0.00);-'
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
    "09 Fuzzy Match Review Hold",
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
    method_col = len(qb_headers) + 1
    inf_start = method_col + 1
    qb_end = len(qb_headers)
    inf_end = inf_start + len(inf_headers) - 1
    header_row, data_row = 3, 4

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

    _write_title_band(ws, 1, qb_start, qb_end, "QUICKBOOKS | SORTED BY PO", NAVY)
    _write_title_band(ws, 1, method_col, method_col, "MATCH METHOD", SLATE)
    _write_title_band(ws, 1, inf_start, inf_end, "INFINIUM", TEAL)
    _write_caption_band(
        ws, 2, qb_start, qb_end,
        f"Every QuickBooks row, sorted by normalized PO. {matched_count:,} of {len(qb_rows):,} matched -- see "
        f"the Exceptions sheet for QuickBooks items shown gold/red here. Generated "
        f"{format_central_timestamp(result.run_timestamp)}.",
        NAVY,
    )
    _write_caption_band(
        ws, 2, method_col, method_col,
        "States the exact rule that resolved the match, including whether it was a fuzzy "
        "text-similarity match rather than an exact one.",
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
        ws.cell(row, method_col, f"{record.get('Match Result', '')} ({record.get('Confidence', '')})")

    ws.cell(header_row, method_col, "Match Method")
    _write_dataframe_values(ws, pd.DataFrame(columns=qb_headers), header_row, qb_start)
    _write_dataframe_values(ws, pd.DataFrame(columns=inf_headers), header_row, inf_start)
    _format_header(ws, header_row, qb_start, qb_end, NAVY)
    _format_header(ws, header_row, method_col, method_col, SLATE)
    _format_header(ws, header_row, inf_start, inf_end, TEAL)

    for offset, record in enumerate(all_rows):
        row = data_row + offset
        section = record.get("Section", "")
        # Alignment/border first, uniform across the whole row regardless
        # of outcome, so every cell has a consistent look; the color style
        # applied next only ever touches fill/font, never alignment.
        _apply_default_alignment(ws, row, qb_start, inf_end)
        if section in _LEGACY_MATCHED_SECTIONS:
            _apply_good_style(ws, row, qb_start, qb_end)
            _apply_good_style(ws, row, inf_start, inf_end)
        elif section in _LEGACY_DUPLICATE_SECTIONS:
            _apply_duplicate_style(ws, row, qb_start, qb_end)
            _apply_duplicate_style(ws, row, inf_start, inf_end)
        else:
            _apply_neutral_style(ws, row, qb_start, qb_end)
            _apply_neutral_style(ws, row, inf_start, inf_end)
        _apply_method_style(ws, row, method_col, method_col)
        ws.cell(row, method_col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.cell(row, method_col).border = _thin_border()
    if not all_rows:
        _apply_default_alignment(ws, data_row, qb_start, inf_end)
        _apply_method_style(ws, data_row, method_col, method_col)

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
    ws.cell(total_row, method_col).fill = PatternFill("solid", fgColor=SLATE_LIGHT)
    ws.cell(total_row, method_col).border = _total_border()
    _apply_number_formats(ws, qb_headers, data_row, total_row, qb_start,
                          {result.qb_mapping["amount"]}, {result.qb_mapping.get("quantity") or ""})
    _apply_number_formats(ws, inf_headers, data_row, total_row, inf_start,
                          {result.inf_mapping["amount"]}, set())
    _set_widths(ws, qb_start, qb_end, header_row, total_row, maximum=40)
    _set_widths(ws, inf_start, inf_end, header_row, total_row, maximum=40)
    _standardize_legacy_widths(ws, qb_headers, qb_start, result.qb_mapping)
    _standardize_legacy_widths(ws, inf_headers, inf_start, result.inf_mapping)
    ws.column_dimensions[get_column_letter(method_col)].width = 46
    ws.freeze_panes = f"{get_column_letter(inf_start)}{data_row}"
    if all_rows:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(inf_end)}{final_data_row}"
    ws.print_title_rows = "1:3"
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
    trailer_headers = ["Fiscal Period", "Exception Type", "Explanation"]
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
        f"{format_central_timestamp(result.run_timestamp)}.",
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
        records: list[dict], block_qb_start: int, trailer_start: int, trailer_end: int, style_fn,
    ) -> int:
        block_headers = qb_headers + trailer_headers
        block = pd.DataFrame(
            [
                _legacy_row_values(record.get("QB Index"), record.get("QB Record Scope"), result.qb_work, None, qb_headers)
                + [
                    row_period(record),
                    _legacy_section_label(str(record.get("Section", ""))),
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
            style_fn(ws, row, block_qb_start, trailer_end)
            ws.cell(row, trailer_end).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        _apply_number_formats(ws, qb_headers, data_row, block_final_row, block_qb_start,
                              {result.qb_mapping["amount"]}, {result.qb_mapping.get("quantity") or ""})
        _set_widths(ws, block_qb_start, block_qb_start + n_qb - 1, header_row, block_final_row, maximum=40)
        _standardize_legacy_widths(ws, qb_headers, block_qb_start, result.qb_mapping)
        ws.column_dimensions[get_column_letter(trailer_start)].width = 14
        ws.column_dimensions[get_column_letter(trailer_start + 1)].width = 34
        ws.column_dimensions[get_column_letter(trailer_end)].width = 60
        return block_final_row

    general_final_row = write_block(
        general_rows, qb_start, left_trailer_start, left_trailer_end, _apply_neutral_style,
    )
    duplicate_final_row = write_block(
        duplicate_rows, dup_qb_start, dup_trailer_start, dup_trailer_end, _apply_duplicate_style,
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
    )


def build_product_sheet(wb: Workbook, result: ReconciliationResult) -> None:
    ws = wb.create_sheet("Product Aggregate Summary")
    frame = result.product_summary
    headers = ["Product Name", "Product Quantity", "Product Value"]
    _write_title_band(ws, 1, 1, 3, "PRODUCT AGGREGATE SUMMARY", NAVY)
    selected_period = result.metadata.get("fiscal_period")
    period_scope = (
        f"Only primary QuickBooks rows from Period {int(selected_period):02d} are included."
        if selected_period is not None
        else "All primary QuickBooks fiscal periods are included."
    )
    _write_caption_band(
        ws, 2, 1, 3,
        "Products are classified from the QuickBooks description column. "
        f"{period_scope} This classification does not influence data matching. "
        f"Generated {format_central_timestamp(result.run_timestamp)}.",
        NAVY,
    )
    _write_dataframe_values(ws, frame.reindex(columns=headers), 3, 1)
    _format_header(ws, 3, 1, 3, NAVY)
    if len(frame):
        _format_body_block(ws, 4, 3 + len(frame), 1, 3, NAVY_LIGHT)
    total_row = 4 + len(frame)
    totals = {
        "Product Quantity": float(frame["Product Quantity"].sum()) if not frame.empty else 0,
        "Product Value": float(frame["Product Value"].sum()) if not frame.empty else 0,
    }
    _write_total_row(ws, total_row, 1, 3, totals, headers)
    _apply_number_formats(ws, headers, 4, total_row, 1, {"Product Value"}, {"Product Quantity"})
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.freeze_panes = "A4"
    if len(frame):
        ws.auto_filter.ref = f"A3:C{3 + len(frame)}"
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

        for row in ws.iter_rows():
            max_lines = 1
            row_idx = row[0].row

            # Row 2 is a controlled report caption band. Do not let its long
            # merged-cell text enter the generic wrapping calculation, which
            # can otherwise expand it to several times the requested height.
            if row_idx == 2:
                ws.row_dimensions[row_idx].height = _controlled_row_two_height(ws.title)
                continue

            for cell in row:
                text = str(cell.value) if cell.value is not None else ""
                if not text:
                    continue

                # Estimate character capacity based on column width
                col_width = col_widths.get(cell.column_letter, 15)
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
    return buffer.getvalue()


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


def build_primary_workbook(result: ReconciliationResult) -> bytes:
    wb = Workbook()
    wb.properties.creator = "Sales Reconciliation Application"
    wb.properties.title = f"Sales Reconciliation {result.run_id}"
    wb.properties.subject = "QuickBooks to Infinium reconciliation and journal-entry support"
    wb.properties.description = "Accounting workpaper generated from one controlled reconciliation run."
    build_data_search_sheet(wb, result)
    build_raw_data_sheet(wb, result)
    build_reconciled_data_sheet(wb, result)
    build_unresolved_sheet(wb, result)
    build_product_sheet(wb, result)
    _apply_workbook_run_metadata(wb, result)
    return _save_workbook_bytes(wb, apply_accountant_row_heights=True)


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
            "Match Result": record["Match Result"],
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
    "Status", "Match Type", "Duplicate Group ID",
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
            "Match Type": match_type,
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
        "Status", "Match Type", "Duplicate Group ID", "Infinium Row ID", "Detail",
    ]].rename(columns={
        "QuickBooks Row ID": "Row ID", "QuickBooks PO": "PO", "QuickBooks Invoice": "Invoice",
        "QuickBooks Amount": "Amount", "Infinium Row ID": "Matched Infinium Row ID",
    }).reset_index(drop=True)
    inf_index = joined.loc[joined["Infinium Row ID"].notna(), [
        "Infinium Row ID", "Infinium PO", "Infinium Invoice", "Infinium Amount",
        "Status", "Match Type", "Duplicate Group ID", "QuickBooks Row ID", "Detail",
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
    ws.cell(label_row, start_col).font = Font(name="Segoe UI", size=11, bold=True, color=TEXT)
    ws.merge_cells(start_row=label_row, start_column=start_col, end_row=label_row, end_column=start_col + 2)
    input_col = start_col + 3
    input_cell = ws.cell(input_row, input_col, "")
    ws.merge_cells(start_row=input_row, start_column=input_col, end_row=input_row, end_column=input_col + 2)
    for col in range(input_col, input_col + 3):
        cell = ws.cell(input_row, col)
        cell.fill = PatternFill("solid", fgColor=WHITE)
        cell.border = _thin_border()
        cell.font = Font(name="Segoe UI", size=11, color=TEXT)
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
    _format_header(ws, column_header_row, start_col, end_col, header_color)

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
        f'FILTER({data_range},{criteria},"No matching items found"))'
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
        ("Unresolved QB amount", result.metrics["Unresolved QuickBooks Amount"], '$#,##0.00;[Red]($#,##0.00);-'),
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
        ws.cell(row, col).font = Font(name="Segoe UI", size=9, bold=True, color=SLATE)
        ws.cell(row + 1, col).font = Font(
            name="Segoe UI", size=13, bold=True,
            color=NAVY if value != "FAIL" else "9C0006",
        )
        ws.cell(row + 1, col).number_format = number_format

    start = 11
    posting_status = result.metrics.get("Posting Status", "READY TO POST")
    ws.cell(start, 1, "MODEL STATUS")
    ws.cell(start, 2, result.metrics["Control Status"])
    ws.cell(start, 1).font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    ws.cell(start, 1).fill = PatternFill("solid", fgColor=SLATE)
    ws.cell(start, 2).font = Font(name="Segoe UI", size=11, bold=True, color=TEXT)
    ws.cell(start, 2).fill = PatternFill("solid", fgColor=GREEN_LIGHT if result.metrics["Control Status"] == "PASS" else RED_LIGHT)
    ws.cell(start, 3, "POSTING STATUS")
    ws.cell(start, 4, posting_status)
    ws.cell(start, 3).font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    ws.cell(start, 3).fill = PatternFill("solid", fgColor=SLATE)
    ws.cell(start, 4).font = Font(name="Segoe UI", size=11, bold=True, color=TEXT)
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