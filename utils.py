"""Generic helper functions used across the application.

This module holds standalone data sanitization, timestamp formatting, and
display formatting helpers, plus the small run-identity / session-signature
functions used by app.py to decide when a cached reconciliation result must
be discarded. None of these functions build UI or Excel output themselves,
which is what keeps them reusable from every other module.
"""

from __future__ import annotations

import hashlib
from datetime import datetime, timezone
from decimal import Decimal
from typing import Any

import pandas as pd
import streamlit as st

from config import CENTRAL_TIMEZONE


def format_central_timestamp(value: datetime) -> str:
    """Format a timezone-aware report timestamp in U.S. Central Time."""
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    central_value = value.astimezone(CENTRAL_TIMEZONE)
    return central_value.strftime("%Y-%m-%d %I:%M:%S %p %Z")


def excel_safe(value: Any) -> Any:
    """Convert values safely for Excel and neutralize formula-like source text."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime) and value.tzinfo is not None:
        value = value.replace(tzinfo=None)
    if isinstance(value, Decimal):
        return float(value)
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass
    if isinstance(value, (list, tuple, set)):
        value = "; ".join(str(item) for item in value)
    if isinstance(value, str):
        # Excel treats strings beginning with these characters as formulas or
        # external commands in several import/opening contexts. Prefixing an
        # apostrophe forces a literal text cell without changing its display.
        if value.lstrip().startswith(("=", "+", "-", "@")):
            return "'" + value
    return value


def format_currency(value: Any) -> str:
    numeric = float(value or 0)
    return f"${numeric:,.2f}" if numeric >= 0 else f"(${abs(numeric):,.2f})"


def run_id_for_inputs(*source_hashes: str) -> tuple[str, datetime]:
    now_utc = datetime.now(timezone.utc)
    now_central = now_utc.astimezone(CENTRAL_TIMEZONE)
    digest_source = "|".join([*source_hashes, now_utc.isoformat()])
    digest = hashlib.sha256(digest_source.encode()).hexdigest()[:8].upper()
    return f"REC-{now_central.strftime('%Y%m%d-%H%M%S')}-{digest}", now_central


def clear_results_if_signature_changed(signature: tuple[Any, ...]) -> None:
    if st.session_state.get("input_signature") != signature:
        st.session_state.input_signature = signature
        st.session_state.pop("reconciliation_result", None)
        st.session_state.pop("primary_workbook", None)
        st.session_state.pop("analytics_workbook", None)