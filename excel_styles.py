"""OpenPyXL formatting engine for the reconciliation workbooks.

Every visual primitive used by the sheet builders in ``workpapers.py`` lives
here: thin/total borders, header and body block styling, duplicate-row
highlighting, number formats, column-width sizing (including the bounded
workbook-wide autofit pass), title/caption bands, and total rows. Keeping
these out of workpapers.py keeps that module focused on "what goes on each
sheet" rather than "how a cell block is painted."
"""

from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import (
    ACCOUNTING_COUNT_FORMAT,
    ACCOUNTING_CURRENCY_FORMAT,
    ACCOUNTING_QUANTITY_FORMAT,
    BORDER,
    DATE_NUMBER_FORMAT,
    DUPLICATE_RED_FILL,
    DUPLICATE_RED_TEXT,
    FONT_NAME,
    FONT_NAME_NUMERIC,
    LEGACY_BODY_TEXT,
    SLATE,
    SLATE_LIGHT,
    TEXT,
    TOTAL_FILL,
    WHITE,
)

# ---------------------------------------------------------------------------
# Global Style Singletons (Prevents massive instantiation overhead in loops)
# ---------------------------------------------------------------------------

ALIGN_LEFT_CENTER = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT_CENTER = Alignment(horizontal="right", vertical="center")
ALIGN_CENTER_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_WRAP_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

FONT_BODY = Font(name=FONT_NAME, size=10, color=TEXT)
FONT_DUPLICATE = Font(name=FONT_NAME, size=10, bold=True, color=DUPLICATE_RED_TEXT)
FONT_LEGACY_BODY = Font(name=FONT_NAME, size=10, color=LEGACY_BODY_TEXT)
FONT_LEGACY_BODY_BOLD = Font(name=FONT_NAME, size=10, bold=True, color=LEGACY_BODY_TEXT)
FONT_TOTAL = Font(name=FONT_NAME, size=10, bold=True, color=TEXT)
FONT_HEADER = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
FONT_TITLE = Font(name=FONT_NAME, size=12, bold=True, color=WHITE)

FILL_DUPLICATE = PatternFill("solid", fgColor=DUPLICATE_RED_FILL)
FILL_TOTAL = PatternFill("solid", fgColor=TOTAL_FILL)
FILL_NONE = PatternFill(fill_type=None)
FILL_CAPTION_BAND = PatternFill("solid", fgColor=SLATE_LIGHT)

_SIDE_THIN_BORDER = Side(style="thin", color=BORDER)
BORDER_THIN = Border(
    left=_SIDE_THIN_BORDER, right=_SIDE_THIN_BORDER, 
    top=_SIDE_THIN_BORDER, bottom=_SIDE_THIN_BORDER
)

BORDER_TOTAL = Border(
    top=Side(style="thin", color=SLATE),
    bottom=Side(style="double", color=SLATE),
)

# ---------------------------------------------------------------------------
# Formatting Engine
# ---------------------------------------------------------------------------

def fix_row_height(ws, row: int, height: float) -> None:
    """Give a row a deliberate height and mark it, so the workbook-wide row
    autofit -- which estimates wrapping from text length -- leaves it alone.
    Title bands, KPI cards, legends, and control strips are single designed
    lines; measured as text they would balloon."""
    ws.row_dimensions[row].height = height
    fixed = getattr(ws, "_fixed_row_heights", {})
    fixed[row] = height
    ws._fixed_row_heights = fixed


def _thin_border() -> Border:
    return BORDER_THIN


def _total_border() -> Border:
    return BORDER_TOTAL


ALIGN_WRAP_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
ALIGN_WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _infer_column_style(
    header: str,
    amount_columns: frozenset = frozenset(),
    quantity_columns: frozenset = frozenset(),
) -> tuple:
    """Classify a column from its header text -- or explicit membership in
    amount_columns/quantity_columns, which name the real mapped field so
    this still works when a header doesn't textually resemble its meaning
    (e.g. Infinium's amount field is literally "OHTOTA"). Returns
    (alignment, number_format, is_numeric_dense); is_numeric_dense marks a
    measured, calculated value (amount, quantity, count, rate) for the
    monospaced/tabular font, so a column of figures lines up for scanning.
    An identifier -- PO, invoice, period, customer number, or any other
    reference number that is never summed or compared numerically -- is
    deliberately excluded even when it reads as "all digits": the
    distinction that matters is measurement vs. identifier, not text vs.
    number. Shared by _apply_number_formats (the data cells) and
    _format_header (so the header lines up with its data).
    """
    header_upper = str(header).upper()
    if any(term in header_upper for term in ("VALID", "GROUP-LEVEL", "AUTOMATIC")):
        return ALIGN_CENTER_CENTER, None, False
    if any(term in header_upper for term in ("PERCENT", "RATE", "SHARE")):
        return ALIGN_RIGHT_CENTER, "0.0%", True
    if any(term in header_upper for term in ("COUNT", "ROWS", "CANDIDATE COUNT")):
        return ALIGN_RIGHT_CENTER, ACCOUNTING_COUNT_FORMAT, True
    if header in amount_columns or any(
        term in header_upper
        for term in ("AMOUNT", "VALUE", "VARIANCE", "DIFFERENCE", "BALANCE", "EXPOSURE", "TOLERANCE")
    ):
        return ALIGN_RIGHT_CENTER, ACCOUNTING_CURRENCY_FORMAT, True
    if header in quantity_columns or any(term in header_upper for term in ("QUANTITY", "QTY")):
        return ALIGN_RIGHT_CENTER, ACCOUNTING_QUANTITY_FORMAT, True
    if "DATE" in header_upper or "TIMESTAMP" in header_upper:
        return ALIGN_CENTER_CENTER, DATE_NUMBER_FORMAT, False
    return None, None, False


def _numeric_font(cell) -> Font:
    """The cell's current font, swapped to the monospaced numeric face --
    every digit the same width, so a column of measured figures lines up
    for scanning -- while keeping its bold/italic/color/underline. Sized
    one point smaller than the surrounding Segoe UI: Consolas's heavier,
    more uniform stroke weight reads noticeably darker and larger than
    Segoe UI at the same point size, so dropping a point keeps the
    numbers visually level with the descriptive text around them."""
    current = cell.font
    current_size = current.size or 10
    return Font(
        name=FONT_NAME_NUMERIC,
        size=max(current_size - 1, 1),
        bold=current.bold,
        italic=current.italic,
        color=current.color,
        underline=current.underline,
    )


def _format_header(
    ws,
    row: int,
    start_col: int,
    end_col: int,
    color: str,
    headers: list = None,
    amount_columns: frozenset = frozenset(),
    quantity_columns: frozenset = frozenset(),
) -> None:
    """Paint a header band. Pass `headers` (left-to-right, matching the
    column range) to align each header with its own column's data --
    right over a right-aligned numeric column, centered over a centered
    one, left (wrapped) otherwise -- rather than every header defaulting
    to left, which reads oddly sitting over right-aligned figures. Without
    `headers`, every column in the range is left-aligned and wrapped as
    before.
    """
    header_fill = PatternFill("solid", fgColor=color)
    header_border = Border(bottom=Side(style="medium", color=color))

    for offset, col in enumerate(range(start_col, end_col + 1)):
        cell = ws.cell(row, col)
        cell.fill = header_fill
        cell.font = FONT_HEADER
        cell.border = header_border
        alignment = ALIGN_WRAP_LEFT
        if headers is not None and offset < len(headers):
            inferred_align, _, _ = _infer_column_style(headers[offset], amount_columns, quantity_columns)
            if inferred_align is ALIGN_RIGHT_CENTER:
                alignment = ALIGN_WRAP_RIGHT
            elif inferred_align is ALIGN_CENTER_CENTER:
                alignment = ALIGN_WRAP_CENTER
        cell.alignment = alignment
    ws.row_dimensions[row].height = 34


def _format_body_block(
    ws,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    light_fill: str,
) -> None:
    fill_even = PatternFill("solid", fgColor=light_fill)
    
    for row in range(start_row, end_row + 1):
        row_fill = fill_even if row % 2 == 0 else FILL_NONE
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row, col)
            cell.font = FONT_BODY
            cell.border = BORDER_THIN
            cell.fill = row_fill
            cell.alignment = ALIGN_LEFT_CENTER
        ws.row_dimensions[row].height = 20


def _apply_duplicate_style(ws, row: int, start_col: int, end_col: int) -> None:
    """Apply Excel's traditional red bad-value style to a duplicate source row."""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = FILL_DUPLICATE
        cell.font = FONT_DUPLICATE


def _apply_default_alignment(ws, row: int, start_col: int, end_col: int) -> None:
    """Left-align and border a row the same way _format_body_block's plain
    cells look, without touching fill or font -- call this alongside a
    color-status style (_apply_legacy_status_fill, etc.) so every cell in a row
    gets a consistent alignment/border regardless of which style painted
    its color, then let _apply_number_formats override specific columns
    (amounts, dates, quantities) to right/center afterward."""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.alignment = ALIGN_LEFT_CENTER
        cell.border = BORDER_THIN


def _apply_legacy_status_fill(ws, row: int, start_col: int, end_col: int, fill_color: str) -> None:
    """Accountant's Legacy Format row styling: a pale full-row tint with
    plain regular-weight black text. Color alone carries the status --
    bold is reserved for headers, totals, and the one status cell that
    needs attention (see _apply_legacy_status_cell)."""
    fill = PatternFill("solid", fgColor=fill_color)
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.font = FONT_LEGACY_BODY


def _apply_legacy_status_cell(ws, row: int, col: int, fill_color: str, bold: bool) -> None:
    """A single status cell (Match Method / Exception Type) -- bold only
    when the row genuinely needs attention."""
    cell = ws.cell(row, col)
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.font = FONT_LEGACY_BODY_BOLD if bold else FONT_LEGACY_BODY


def _apply_number_formats(
    ws,
    headers: list[str],
    start_row: int,
    end_row: int,
    start_col: int,
    amount_columns: set[str],
    quantity_columns: set[str],
) -> None:
    for offset, header in enumerate(headers):
        col = start_col + offset
        target_align, target_format, is_numeric_dense = _infer_column_style(
            header, amount_columns, quantity_columns,
        )
        if not (target_align or target_format or is_numeric_dense):
            continue
        for row in range(start_row, end_row + 1):
            cell = ws.cell(row, col)
            if target_align:
                cell.alignment = target_align
            if target_format:
                cell.number_format = target_format
            if is_numeric_dense:
                cell.font = _numeric_font(cell)


def _set_widths(
    ws,
    start_col: int,
    end_col: int,
    start_row: int,
    end_row: int,
    minimum: float = 11,
    maximum: float = 34,
) -> None:
    for col in range(start_col, end_col + 1):
        lengths = [len(str(ws.cell(row, col).value or "")) for row in range(start_row, end_row + 1)]
        ws.column_dimensions[get_column_letter(col)].width = min(max(max(lengths, default=0) + 2, minimum), maximum)


def _standardize_column_widths(
    ws, headers: list, start_col: int, fixed_widths: dict,
) -> None:
    """After _set_widths has sized columns to their content, override
    specific columns (matched by exact header name) to a fixed, appropriate
    width -- so a short code column and a currency column each get a
    consistent width regardless of what their own content happened to look
    like, instead of every column being sized independently by chance."""
    for offset, header in enumerate(headers):
        if header in fixed_widths:
            col = start_col + offset
            ws.column_dimensions[get_column_letter(col)].width = fixed_widths[header]


def _pin_column_width(ws, letter: str, width: float) -> None:
    """Set a column's width and mark it deliberate, so the workbook-wide
    content-driven autofit leaves it alone (a narrow reference column whose
    long heading is meant to wrap, for instance, must not be stretched to
    fit that heading on one line)."""
    ws.column_dimensions[letter].width = width
    pinned = getattr(ws, "_pinned_column_letters", set())
    pinned.add(letter)
    ws._pinned_column_letters = pinned


# A Referenced Match Ref. cell is a live-link formula whose length has nothing
# to do with the short reference it displays.
_MATCH_LINK_FORMULA = re.compile(r'^=IFERROR\(HYPERLINK\(.*,"([^"]*)"\),"[^"]*"\)$')


def _autofit_display_text(value) -> str:
    """The text a cell shows, for sizing its column: a match-reference link
    formula counts as its short displayed reference, not its formula text."""
    text = str(value)
    link = _MATCH_LINK_FORMULA.match(text)
    return link.group(1) if link else text


def _autofit_workbook_columns(
    wb: Workbook,
    minimum: float = 10,
    maximum: float = 38,
    sample_rows: int = 750,
    skip_titles: frozenset = frozenset(),
) -> None:
    """Auto-size columns from a bounded sample instead of rescanning every cell.

    A sheet in `skip_titles` keeps whatever widths its own builder function
    already set instead of having them overwritten here -- for a sheet with
    deliberately standardized, type-appropriate widths (see
    _standardize_column_widths), this content-driven pass would otherwise
    always win since it runs after every sheet is built.
    """
    for ws in wb.worksheets:
        if ws.title in skip_titles:
            continue
        merged_coordinates: set[str] = set()
        for merged_range in ws.merged_cells.ranges:
            for row in ws.iter_rows(
                min_row=merged_range.min_row,
                max_row=merged_range.max_row,
                min_col=merged_range.min_col,
                max_col=merged_range.max_col,
            ):
                merged_coordinates.update(cell.coordinate for cell in row)

        sampled_row_numbers = list(range(1, min(ws.max_row, sample_rows) + 1))
        if ws.max_row > sample_rows:
            sampled_row_numbers.append(ws.max_row)

        pinned_letters = getattr(ws, "_pinned_column_letters", set())
        for column_index in range(1, ws.max_column + 1):
            letter = get_column_letter(column_index)
            if letter in pinned_letters:
                continue
            maximum_length = 0
            has_unmerged_value = False
            for row_index in sampled_row_numbers:
                cell = ws.cell(row_index, column_index)
                if cell.coordinate in merged_coordinates or cell.value is None:
                    continue
                has_unmerged_value = True
                value = cell.value
                if isinstance(value, datetime):
                    display = value.strftime("%Y-%m-%d %I:%M:%S %p")
                elif isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
                    if "$" in str(cell.number_format):
                        display = f"${float(value):,.2f}"
                    elif any(token in str(cell.number_format) for token in ("0.00", "#,##0")):
                        display = f"{float(value):,.2f}"
                    else:
                        display = str(value)
                else:
                    display = _autofit_display_text(value)
                maximum_length = max(
                    maximum_length,
                    max((len(line) for line in display.splitlines()), default=0),
                )

            if has_unmerged_value:
                ws.column_dimensions[letter].width = min(max(maximum_length + 2, minimum), maximum)
            else:
                existing = ws.column_dimensions[letter].width or minimum
                ws.column_dimensions[letter].width = min(max(existing, 3.5), maximum)


def _write_title_band(ws, row: int, start_col: int, end_col: int, title: str, color: str) -> None:
    if start_col < end_col:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    
    title_fill = PatternFill("solid", fgColor=color)
    cell = ws.cell(row, start_col, title)
    cell.fill = title_fill
    cell.font = FONT_TITLE
    # A title longer than its band shrinks to fit instead of being clipped.
    cell.alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)

    for col in range(start_col + 1, end_col + 1):
        ws.cell(row, col).fill = title_fill
    fix_row_height(ws, row, 27)


def _write_caption_band(ws, row: int, start_col: int, end_col: int, caption: str, color: str) -> None:
    if start_col < end_col:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        
    caption_font = Font(name="Segoe UI", size=9, italic=True, color=color)
    cell = ws.cell(row, start_col, caption)
    cell.fill = FILL_CAPTION_BAND
    cell.font = caption_font
    cell.alignment = ALIGN_WRAP_LEFT
    
    for col in range(start_col + 1, end_col + 1):
        ws.cell(row, col).fill = FILL_CAPTION_BAND
    ws.row_dimensions[row].height = 30


def _write_total_row(
    ws,
    row: int,
    start_col: int,
    end_col: int,
    totals: dict[str, float],
    headers: list[str],
    label: str = "TOTAL",
) -> None:
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = FILL_TOTAL
        cell.font = FONT_TOTAL
        cell.border = BORDER_TOTAL
        
    ws.cell(row, start_col, label)
    for offset, header in enumerate(headers):
        if header in totals:
            cell = ws.cell(row, start_col + offset, totals[header])
            cell.number_format = (
                ACCOUNTING_CURRENCY_FORMAT
                if "amount" in header.lower() or "value" in header.lower()
                else ACCOUNTING_QUANTITY_FORMAT
            )
            cell.alignment = ALIGN_RIGHT_CENTER
            cell.font = _numeric_font(cell)
    ws.row_dimensions[row].height = 23


def _prepare_sheet(ws, landscape: bool = True) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if landscape else ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.autoPageBreaks = False