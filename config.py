"""Global configuration constants for the sales reconciliation application.

Centralizing the color palette and shared settings here means every other
module (excel_styles, workpapers, ui_components, utils) can import them
without creating circular dependencies back on app.py.
"""

from __future__ import annotations

from zoneinfo import ZoneInfo

CENTRAL_TIMEZONE = ZoneInfo("America/Chicago")

NAVY = "1B365D"
NAVY_LIGHT = "E8EEF5"
TEAL = "27666B"
TEAL_LIGHT = "E7F1F1"
SLATE = "475569"
SLATE_LIGHT = "F2F5F8"
AMBER = "FFF3CD"
ORANGE = "FCE8D5"
# Softened from the original bright Excel "Bad" red (FFC7CE/9C0006) -- eye
# strain when a reviewer is scanning hundreds of exception rows was the
# reported problem. A muted pastel background with a dark, legible maroon
# text keeps the "this is a problem" signal without the harshness.
DUPLICATE_RED_FILL = "F8D7DA"
DUPLICATE_RED_TEXT = "72232B"
GOOD_GREEN_FILL = "C6EFCE"
GOOD_GREEN_TEXT = "006100"
NEUTRAL_GOLD_FILL = "FFEB9C"
NEUTRAL_GOLD_TEXT = "9C6500"
METHOD_GREY_FILL = "D9D9D9"
METHOD_GREY_TEXT = "000000"
# A true neutral grey (no blue or teal undertone) for a column that exists
# purely to divide two colored sides -- e.g. Reconciliation Detail's "Match
# Result" column, sitting between the QuickBooks (navy) and Infinium
# (teal) blocks. SLATE reads as a third shade of blue next to those two,
# not as a clear divider; this is deliberately desaturated so it never
# competes with either side's identity color.
METHOD_GREY_DARK = "595959"

# Infinium's raw extract headers are cryptic system codes; every reconciliation
# sheet that shows an Infinium record uses these plain field names instead.
# Display-only -- the values, widths, and number formats still key off the
# original header, and the raw upload keeps the codes untouched.
INFINIUM_FRIENDLY_HEADERS = {
    "OHAPD": "Period",
    "OHOBDE": "Date",
    "OHCO": "Type",
    "CUNO": "Customer No.",
    "OHOBNO": "Invoice No.",
    "OHTOTA": "Amount",
    "OHDESC": "Description",
    "OHPONO": "PO No.",
}

# Accountant's Legacy Format workbook only. A full-row fill repeated across
# hundreds of rows has to stay quiet, so these are much lighter tints than
# the primary workpaper's status colors, paired with plain black regular
# text instead of a colored bold font.
LEGACY_MATCHED_FILL = "E7F3E8"
LEGACY_REVIEW_FILL = "FFF2CC"
LEGACY_EXCLUDED_FILL = "FCE4E4"
# The empty side of a row with no paired record in that dataset -- the blank
# Infinium block of an unmatched QuickBooks row, or the blank QuickBooks block
# of an Infinium-only row. Almost white, so "nothing here" recedes.
LEGACY_NO_PAIR_FILL = "F8FAF9"
# The Match Method data column: RGB(234, 234, 234).
LEGACY_METHOD_FILL = "EAEAEA"
LEGACY_BODY_TEXT = "000000"

# Every real date cell in every workbook is displayed this way.
DATE_NUMBER_FORMAT = "mm/dd/yyyy"
RED_LIGHT = "FDECEC"
GREEN_LIGHT = "E8F3EC"
WHITE = "FFFFFF"
TEXT = "172B4D"
BORDER = "D5DDE4"
TOTAL_FILL = "E2E8F0"

# Every workbook cell is explicitly set to this UI font (see
# _apply_workbook_default_font in excel_styles.py, plus each Font(...) call
# throughout this module) rather than leaving anything on Excel's Calibri
# theme default.
FONT_NAME = "Segoe UI"
# Applied specifically to amount/quantity/invoice-and-PO-like columns: a
# true monospaced font so every digit occupies the same width, making a
# column of numbers easy to scan straight down rather than a UI font's
# proportional (if fairly tabular) figures.
FONT_NAME_NUMERIC = "Consolas"

# True Excel "Accounting" format: the currency symbol is locked to the far
# left of the cell, thousands separators and a fixed two-decimal tail keep
# every value's decimal point vertically aligned with its neighbors, and a
# lone "-" (not "$0.00") marks a zero. Negatives still render in red,
# consistent with the rest of the workbook's number formats.
ACCOUNTING_CURRENCY_FORMAT = '_($* #,##0.00_);[Red]_($* (#,##0.00);_($* "-"??_);_(@_)'
# Same alignment behavior, without the currency symbol, for quantity
# columns -- whole numbers only (a quantity of units sold has no
# fractional part), still comma-separated and accounting-aligned.
ACCOUNTING_QUANTITY_FORMAT = '_(* #,##0_);[Red]_(* (#,##0);_(* "-"??_);_(@_)'
ACCOUNTING_COUNT_FORMAT = '_(* #,##0_);[Red]_(* (#,##0);_(* "-"??_);_(@_)'