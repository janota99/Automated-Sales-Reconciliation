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
DUPLICATE_RED_FILL = "FFC7CE"
DUPLICATE_RED_TEXT = "9C0006"
RED_LIGHT = "FDECEC"
GREEN_LIGHT = "E8F3EC"
WHITE = "FFFFFF"
TEXT = "172B4D"
BORDER = "D5DDE4"
TOTAL_FILL = "E2E8F0"