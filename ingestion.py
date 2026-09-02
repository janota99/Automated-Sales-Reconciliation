"""Source ingestion: file parsing, header detection, cleanup, and column mapping.

This module owns everything about getting a raw QuickBooks or Infinium
export into a clean, well-mapped DataFrame before it is handed to
``matching.py``:

* reading uploaded CSV/XLSX bytes and listing worksheets
* detecting the header row and removing blank columns, repeated header
  rows, and standalone legacy report-timestamp rows
* excluding QuickBooks subtotal / non-detail rows
* inferring and letting the user confirm standard column mappings
* validating a mapped source before it is allowed into reconciliation

It depends only on ``matching.py`` for the normalization primitives needed
by the validation report (clean_po, clean_alphanumeric, parse_amount_cents).
"""

from __future__ import annotations

import hashlib
import io
import re
from collections import defaultdict
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable, Optional

import pandas as pd
import streamlit as st

from matching import clean_alphanumeric, clean_po, parse_amount_cents

__all__ = [
    "INF_COLUMN_PATTERNS",
    "NONE_OPTION",
    "QB_COLUMN_PATTERNS",
    "REPORT_TIMESTAMP_PATTERN",
    "SELECT_OPTION",
    "build_source_validation_report",
    "detect_header_row",
    "file_sha256",
    "filter_qb_subtotal_rows",
    "fiscal_period_column_profile",
    "infer_column",
    "infer_qb_period_column",
    "list_source_sheets",
    "mapping_panel",
    "read_raw_source",
    "read_source_file",
    "select_column",
    "secondary_mapping_panel",
]


REPORT_TIMESTAMP_PATTERN = re.compile(
    r"^(?:MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)\s*,\s*"
    r"(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+"
    r"\d{1,2}\s*,\s*\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s+(?:AM|PM)\s+"
    r"(?:GMT|UTC)(?:[+-]\d{2}:?\d{2})?$",
    re.IGNORECASE,
)

_RE_NON_ALNUM = re.compile(r"[^A-Z0-9]")

NONE_OPTION = "(Not mapped)"
SELECT_OPTION = "(Select a column)"

# Standard source schemas are attempted first on every upload. Users only need
# the override controls when a report's headers depart from these conventions.
QB_COLUMN_PATTERNS = {
    "po": ["P.O. NUMBER", "PO NUMBER", "PO#", "PO"],
    "invoice": ["NUM", "INVOICE", "INVOICE NUMBER", "INVOICE #"],
    "amount": ["AMOUNT", "TOTAL AMOUNT", "TOTAL"],
    "quantity": ["QTY", "QUANTITY"],
    "product": ["MEMO/DESCRIPTION", "MEMO", "DESCRIPTION", "PRODUCT"],
    "period": ["FISCAL PERIOD", "SOURCE PERIOD", "PERIOD", "PD"],
}
INF_COLUMN_PATTERNS = {
    "po": ["OHDESC", "PO", "PO NUMBER", "PO#"],
    "invoice": ["OHOBNO", "INVOICE", "INVOICE NUMBER", "INVOICE #"],
    "amount": ["OHTOTA", "AMOUNT", "TOTAL AMOUNT", "TOTAL"],
    "period": ["FISCAL PERIOD", "SOURCE PERIOD", "PERIOD", "PD"],
}

def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _header_score(values: Iterable[Any], source: str) -> int:
    tokens = {str(value).strip().upper() for value in values if value is not None and not pd.isna(value)}
    joined = " | ".join(tokens)
    score = len(tokens)
    if source == "QB":
        for term, weight in (("AMOUNT", 10), ("P.O", 7), ("PO", 5), ("NUM", 5), ("QTY", 4), ("MEMO", 4)):
            if term in joined:
                score += weight
    else:
        for term, weight in (("OHTOTA", 12), ("OHOBNO", 8), ("OHDESC", 8), ("AMOUNT", 8), ("INVOICE", 5)):
            if term in joined:
                score += weight
    return score


@st.cache_data(show_spinner=False, max_entries=8)
def list_source_sheets(file_bytes: bytes, filename: str) -> tuple[str, ...]:
    """Return Excel worksheet names; CSV sources have no worksheet selector."""
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        return ()
    if suffix != ".xlsx":
        raise ValueError("Only CSV and XLSX source files are supported.")
    with pd.ExcelFile(io.BytesIO(file_bytes)) as workbook:
        sheets = tuple(str(name) for name in workbook.sheet_names)
    if not sheets:
        raise ValueError("The Excel workbook does not contain any worksheets.")
    return sheets


@st.cache_data(show_spinner=False, max_entries=16)
def read_raw_source(
    file_bytes: bytes,
    filename: str,
    sheet_name: Optional[str] = None,
) -> pd.DataFrame:
    """Parse an uploaded source once and reuse it for detection and cleanup."""
    suffix = Path(filename).suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise ValueError("Only CSV and XLSX source files are supported.")
    stream = io.BytesIO(file_bytes)
    if suffix == ".csv":
        return pd.read_csv(stream, header=None, dtype=object)
    return pd.read_excel(
        stream,
        sheet_name=sheet_name if sheet_name is not None else 0,
        header=None,
        dtype=object,
    )


@st.cache_data(show_spinner=False, max_entries=16)
def detect_header_row(
    file_bytes: bytes,
    filename: str,
    source: str,
    sheet_name: Optional[str] = None,
) -> int:
    probe = read_raw_source(file_bytes, filename, sheet_name).head(25)
    if probe.empty:
        return 0
    scores = [_header_score(probe.iloc[row].tolist(), source) for row in range(len(probe))]
    return int(max(range(len(scores)), key=lambda idx: scores[idx]))


@lru_cache(maxsize=4096)
def _cached_header_token(text: str) -> str:
    return _RE_NON_ALNUM.sub("", text)


def _header_token(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return _cached_header_token(str(value).strip().upper())


def _stripped_text_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Return one reusable, NA-safe string view of a source frame."""
    if frame.empty:
        return pd.DataFrame("", index=frame.index, columns=frame.columns, dtype="string")
    return frame.astype("string").fillna("").apply(lambda column: column.str.strip())


def _substantive_cells(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(False, index=frame.index, columns=frame.columns)
    return _stripped_text_frame(frame).ne("").astype(bool)


def _standalone_report_timestamp_mask(
    frame: pd.DataFrame,
    *,
    stripped: Optional[pd.DataFrame] = None,
    populated: Optional[pd.DataFrame] = None,
) -> pd.Series:
    """Identify nontransaction report metadata rows such as stale export tags."""
    if frame.empty:
        return pd.Series(False, index=frame.index, dtype=bool)

    stripped = stripped if stripped is not None else _stripped_text_frame(frame)
    populated = populated if populated is not None else stripped.ne("")
    single_value = populated.sum(axis=1).eq(1)
    result = pd.Series(False, index=frame.index, dtype=bool)
    if not single_value.any():
        return result

    # Only timestamp candidates have one populated cell, so stack just that
    # small subset instead of joining every cell in every transaction row.
    candidates = stripped.loc[single_value].where(populated.loc[single_value])
    candidate_text = candidates.bfill(axis=1).iloc[:, 0]
    result.loc[candidate_text.index] = candidate_text.str.fullmatch(
        REPORT_TIMESTAMP_PATTERN,
        na=False,
    ).to_numpy(dtype=bool)
    return result


def _header_token_frame(stripped: pd.DataFrame) -> pd.DataFrame:
    """Vectorized equivalent of applying ``_header_token`` cell by cell."""
    if stripped.empty:
        return stripped.copy()
    return stripped.apply(
        lambda column: column.str.upper().str.replace(_RE_NON_ALNUM, "", regex=True)
    )


def _unique_source_headers(values: list[Any]) -> tuple[list[str], list[str], list[str]]:
    headers: list[str] = []
    original_headers: list[str] = []
    duplicate_renames: list[str] = []
    occurrences: dict[str, int] = defaultdict(int)
    for position, value in enumerate(values, 1):
        try:
            missing = value is None or bool(pd.isna(value))
        except (TypeError, ValueError):
            missing = value is None
        original = "" if missing else str(value).strip()
        base = original or f"Unnamed Column {position}"
        key = _header_token(base) or f"UNNAMEDCOLUMN{position}"
        occurrences[key] += 1
        header = base if occurrences[key] == 1 else f"{base} [{occurrences[key]}]"
        if occurrences[key] > 1:
            duplicate_renames.append(f"{base} renamed to {header}")
        headers.append(header)
        original_headers.append(original)
    return headers, original_headers, duplicate_renames


@st.cache_data(show_spinner=False, max_entries=16)
def read_source_file(
    file_bytes: bytes,
    filename: str,
    header_row: int,
    sheet_name: Optional[str] = None,
) -> pd.DataFrame:
    raw = read_raw_source(file_bytes, filename, sheet_name)
    if raw.empty or header_row < 0 or header_row >= len(raw):
        raise ValueError("The selected header row is outside the populated source data.")

    headers, original_headers, duplicate_renames = _unique_source_headers(raw.iloc[header_row].tolist())
    frame = raw.iloc[header_row + 1:].copy()
    frame.columns = headers

    # Build the cleaned string representation once and keep its masks aligned
    # with each filtering pass. This avoids rescanning the full frame three
    # times during a normal import.
    stripped = _stripped_text_frame(frame)
    populated = stripped.ne("")
    retained_rows = populated.any(axis=1)
    blank_rows_removed = int((~retained_rows).sum())
    frame = frame.loc[retained_rows].copy()
    stripped = stripped.loc[retained_rows]
    populated = populated.loc[retained_rows]

    standalone_timestamp_mask = _standalone_report_timestamp_mask(
        frame,
        stripped=stripped,
        populated=populated,
    )
    standalone_timestamp_rows_removed = int(standalone_timestamp_mask.sum())
    retained_rows = ~standalone_timestamp_mask
    frame = frame.loc[retained_rows].copy()
    stripped = stripped.loc[retained_rows]
    populated = populated.loc[retained_rows]

    expected = pd.Series([_header_token(value) for value in original_headers], index=headers)
    repeated_header_mask = pd.Series(False, index=frame.index, dtype=bool)
    header_candidates = populated.sum(axis=1).ge(2)
    if header_candidates.any():
        tokenized = _header_token_frame(stripped.loc[header_candidates])
        header_matches = tokenized.eq(expected, axis="columns") & expected.ne("")
        nonblank_tokens = tokenized.ne("").sum(axis=1)
        candidate_matches = header_matches.sum(axis=1)
        repeated_header_mask.loc[header_candidates] = (
            candidate_matches.ge(2)
            & candidate_matches.div(nonblank_tokens).ge(0.60)
        ).to_numpy(dtype=bool)
    repeated_header_rows_removed = int(repeated_header_mask.sum())
    retained_rows = ~repeated_header_mask
    frame = frame.loc[retained_rows].copy()
    populated = populated.loc[retained_rows]

    blank_columns = [column for column in frame.columns if not populated[column].any()]
    frame = frame.drop(columns=blank_columns).reset_index(drop=True)
    frame.attrs["ingestion_audit"] = {
        "original_column_count": len(headers),
        "retained_column_count": len(frame.columns),
        "blank_columns_removed": blank_columns,
        "duplicate_headers_renamed": duplicate_renames,
        "repeated_header_rows_removed": repeated_header_rows_removed,
        "standalone_timestamp_rows_removed": standalone_timestamp_rows_removed,
        "blank_rows_removed": blank_rows_removed,
        "selected_worksheet": sheet_name or "CSV source",
    }
    return frame


@st.cache_data(show_spinner=False, max_entries=16)
def filter_qb_subtotal_rows(
    frame: pd.DataFrame,
    mapping: dict[str, Optional[str]],
) -> tuple[pd.DataFrame, int]:
    """Retain only complete QuickBooks detail rows for reconciliation.

    Quantity and amount may be blank and are validated separately. Every other
    QuickBooks field must contain a substantive value. A row with any blank
    non-quantity/non-amount field is classified as a subtotal or product-summary
    row and excluded from all reconciliation populations and outputs.
    """
    if frame.empty:
        return frame.copy(), 0

    stripped = _stripped_text_frame(frame)
    populated = stripped.ne("")
    first_column = frame.columns[0]
    first_values = stripped[first_column]
    normalized_first = (
        first_values.str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip(" .:-")
    )
    labeled_subtotal = (
        normalized_first.str.contains(r"\bTOTAL\s+FOR\b", regex=True, na=False)
        | normalized_first.str.fullmatch(r"SERVICES?", na=False)
        | normalized_first.str.contains(r"\bCASE\s*$", regex=True, na=False)
    )

    numeric_columns = {
        column for column in (mapping.get("quantity"), mapping.get("amount"))
        if column in frame.columns
    }
    other_columns = [column for column in frame.columns if column not in numeric_columns]
    incomplete_non_numeric = (
        ~populated[other_columns].all(axis=1)
        if other_columns
        else pd.Series(False, index=frame.index)
    )

    mask = (labeled_subtotal | incomplete_non_numeric).astype(bool)
    filtered = frame.loc[~mask].reset_index(drop=True)
    filtered.attrs.update(frame.attrs)
    filtered.attrs["qb_subtotal_filter_audit"] = {
        "labeled_subtotal_rows_removed": int(labeled_subtotal.sum()),
        "incomplete_non_quantity_amount_rows_removed": int(
            (incomplete_non_numeric & ~labeled_subtotal).sum()
        ),
        "total_rows_removed": int(mask.sum()),
    }
    return filtered, int(mask.sum())


def fiscal_period_column_profile(series: pd.Series) -> tuple[int, int, bool]:
    """Return populated count, valid 1-13 count, and strict period validity."""
    text = series.astype("string").str.strip()
    populated = series.notna() & text.ne("").fillna(False)
    populated_count = int(populated.sum())
    if populated_count == 0:
        return 0, 0, False
    numeric = pd.to_numeric(series.where(populated), errors="coerce")
    valid = numeric.notna() & numeric.between(1, 13) & numeric.mod(1).eq(0)
    valid_count = int((valid & populated).sum())
    strictly_valid = populated_count == len(series) and valid_count == populated_count
    return populated_count, valid_count, strictly_valid


def infer_qb_period_column(
    frame: pd.DataFrame,
    excluded_columns: Optional[Iterable[Optional[str]]] = None,
) -> Optional[str]:
    """Return the first eligible column containing only numeric periods 1-13."""
    columns = list(frame.columns)
    if frame.empty or not columns:
        return None

    excluded = {column for column in (excluded_columns or []) if column}
    named_column = infer_column(columns, QB_COLUMN_PATTERNS["period"], optional=True)
    ordered_candidates = (
        ([named_column] if named_column and named_column not in excluded else [])
        + [column for column in columns if column != named_column and column not in excluded]
    )
    for column in ordered_candidates:
        _, _, strictly_valid = fiscal_period_column_profile(frame[column])
        if strictly_valid:
            return column
    return None


def _infer_from_normalized_headers(
    normalized: dict[str, str],
    patterns: Iterable[str],
) -> Optional[str]:
    for pattern in patterns:
        target = _cached_header_token(str(pattern).strip().upper())
        for column, cleaned in normalized.items():
            if cleaned == target:
                return column
    for pattern in patterns:
        target = _cached_header_token(str(pattern).strip().upper())
        for column, cleaned in normalized.items():
            # Short tokens such as NUM, PO, and PD are exact-only because
            # substring matching could map NUM to P.O. NUMBER or PO to PRODUCT.
            if len(target) >= 4 and target in cleaned:
                return column
    # Never silently map a required accounting field to an unrelated first
    # column. A missing standard field must be selected deliberately.
    return None


def _infer_pattern_map(
    columns: list[str],
    pattern_map: dict[str, list[str]],
) -> dict[str, Optional[str]]:
    """Infer several fields while normalizing the source headers only once."""
    normalized = {col: _cached_header_token(str(col).strip().upper()) for col in columns}
    return {
        field: _infer_from_normalized_headers(normalized, patterns)
        for field, patterns in pattern_map.items()
    }


def infer_column(columns: list[str], patterns: list[str], optional: bool = False) -> Optional[str]:
    # ``optional`` remains in the public signature for backward compatibility.
    normalized = {col: _cached_header_token(str(col).strip().upper()) for col in columns}
    return _infer_from_normalized_headers(normalized, patterns)


def select_column(label: str, columns: list[str], patterns: list[str], key: str,
                  optional: bool = False, default_column: Optional[str] = None) -> Optional[str]:
    inferred = default_column if default_column in columns else infer_column(columns, patterns, optional=optional)
    prefix = [NONE_OPTION] if optional else ([SELECT_OPTION] if inferred is None else [])
    options = prefix + columns
    default_value = inferred if inferred in options else (NONE_OPTION if optional else SELECT_OPTION)
    selected = st.selectbox(label, options, index=options.index(default_value), key=key)
    return None if selected in {NONE_OPTION, SELECT_OPTION} else selected


def build_source_validation_report(
    frame: pd.DataFrame,
    mapping: dict[str, Optional[str]],
    source: str,
    audit: Optional[dict[str, Any]] = None,
    *,
    require_period: bool = True,
    dataset_label: Optional[str] = None,
) -> pd.DataFrame:
    """Return formal pre-match validation checks for one source dataset."""
    dataset = dataset_label or ("QuickBooks" if source == "QB" else "Infinium")
    records: list[dict[str, Any]] = []

    def add(check: str, status: str, details: str) -> None:
        records.append({"Dataset": dataset, "Check": check, "Status": status, "Details": details})

    audit = audit or {}
    if frame.empty:
        add("Usable data rows", "FAIL", "No transaction rows remain after import cleanup.")
        return pd.DataFrame(records)
    add("Usable data rows", "PASS", f"{len(frame):,} transaction row(s) available for validation.")

    required_fields = ["po", "invoice", "amount"]
    if source == "INF" and require_period:
        required_fields.append("period")
    missing_mappings = [field for field in required_fields if not mapping.get(field)]
    if missing_mappings:
        add(
            "Required mappings", "FAIL",
            "Missing mapping(s): " + ", ".join(missing_mappings) + ".",
        )
        return pd.DataFrame(records)

    mapped_columns = [str(mapping[field]) for field in required_fields]
    absent_columns = [column for column in mapped_columns if column not in frame.columns]
    add(
        "Mapped columns exist",
        "FAIL" if absent_columns else "PASS",
        "Missing mapped column(s): " + ", ".join(absent_columns)
        if absent_columns else "All required mapped columns are present.",
    )
    if absent_columns:
        return pd.DataFrame(records)

    duplicate_mappings = sorted({column for column in mapped_columns if mapped_columns.count(column) > 1})
    add(
        "Distinct mapping fields",
        "FAIL" if duplicate_mappings else "PASS",
        "The same source column cannot serve multiple required fields: " + ", ".join(duplicate_mappings)
        if duplicate_mappings else "All required fields use distinct source columns.",
    )

    amount_column = str(mapping["amount"])
    parsed_amounts = frame[amount_column].map(parse_amount_cents)
    valid_amounts = parsed_amounts.notna()
    valid_amount_count = int(valid_amounts.sum())
    valid_amount_rate = valid_amount_count / len(frame)
    if valid_amount_count == 0 or valid_amount_rate < 0.50:
        amount_status = "FAIL"
    elif valid_amount_rate < 0.90:
        amount_status = "WARNING"
    else:
        amount_status = "PASS"
    add(
        "Amount field parseability",
        amount_status,
        f"{valid_amount_count:,} of {len(frame):,} row(s) ({valid_amount_rate:.1%}) contain valid signed-cent amounts. "
        "At least 50% is required to proceed.",
    )

    po_values = frame[str(mapping["po"])].map(clean_po)
    invoice_values = frame[str(mapping["invoice"])].map(clean_alphanumeric)
    usable_references = po_values.ne("") | invoice_values.ne("")
    usable_reference_count = int(usable_references.sum())
    reference_rate = usable_reference_count / len(frame)
    reference_status = "FAIL" if usable_reference_count == 0 else ("WARNING" if reference_rate < 0.50 else "PASS")
    add(
        "Reference availability",
        reference_status,
        f"{usable_reference_count:,} of {len(frame):,} row(s) ({reference_rate:.1%}) contain a usable PO or invoice reference.",
    )

    period_column = mapping.get("period")
    if period_column and period_column in frame.columns:
        period_populated, period_valid, period_is_strict = fiscal_period_column_profile(
            frame[str(period_column)]
        )
        add(
            "Fiscal period values",
            "PASS" if period_is_strict else "FAIL",
            f"{period_valid:,} of {period_populated:,} populated value(s) are whole-number periods from 1 to 13. "
            "Every retained row must contain a valid period when a period column is mapped.",
        )
    elif source == "QB":
        add(
            "Fiscal period values",
            "PASS",
            "No QuickBooks fiscal-period column was detected or mapped; period-based exception classification is disabled.",
        )

    blank_columns = list(audit.get("blank_columns_removed", []))
    add(
        "Blank-column cleanup", "PASS",
        f"Ignored {len(blank_columns):,} entirely blank column(s)."
        if blank_columns else "No entirely blank columns required removal.",
    )
    repeated_headers = int(audit.get("repeated_header_rows_removed", 0) or 0)
    add(
        "Repeated header-row cleanup", "PASS",
        f"Ignored {repeated_headers:,} repeated header row(s) found inside the data."
        if repeated_headers else "No repeated header rows were found inside the data.",
    )
    timestamp_rows = int(audit.get("standalone_timestamp_rows_removed", 0) or 0)
    add(
        "Standalone report timestamp cleanup", "PASS",
        f"Ignored {timestamp_rows:,} standalone legacy report timestamp row(s)."
        if timestamp_rows else "No standalone legacy report timestamp rows were found.",
    )
    duplicate_headers = list(audit.get("duplicate_headers_renamed", []))
    add(
        "Duplicate column headers",
        "WARNING" if duplicate_headers else "PASS",
        "Preserved duplicate header(s) with unique display names: " + "; ".join(duplicate_headers)
        if duplicate_headers else "Column headers are unique.",
    )
    return pd.DataFrame(records)


def mapping_panel(
    qb_raw: pd.DataFrame,
    inf_raw: pd.DataFrame,
    key_suffix: str,
) -> tuple[dict[str, Optional[str]], dict[str, Optional[str]]]:
    """Render the sidebar column-mapping controls and return both mappings."""
    qb_columns = list(qb_raw.columns)
    inf_columns = list(inf_raw.columns)
    qb_defaults = _infer_pattern_map(qb_columns, QB_COLUMN_PATTERNS)
    inf_defaults = _infer_pattern_map(inf_columns, INF_COLUMN_PATTERNS)
    inf_defaults["period"] = inf_columns[0] if inf_columns else None
    standard_required_ready = all(
        qb_defaults.get(field) for field in ("po", "invoice", "amount")
    ) and all(inf_defaults.get(field) for field in ("po", "invoice", "amount"))

    st.sidebar.markdown("### Field mapping")
    if standard_required_ready:
        st.sidebar.markdown(
            '<div class="source-status complete">Standard column mappings populated automatically</div>',
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.markdown(
            '<div class="source-status active">Action required: select any missing required columns</div>',
            unsafe_allow_html=True,
        )

    with st.sidebar.expander("Review or change column mappings", expanded=not standard_required_ready):
        st.markdown("**QuickBooks**")
        qb_mapping = {
            "po": select_column("PO", qb_columns, QB_COLUMN_PATTERNS["po"], f"qb_map_po_{key_suffix}", default_column=qb_defaults["po"]),
            "invoice": select_column("Invoice", qb_columns, QB_COLUMN_PATTERNS["invoice"], f"qb_map_inv_{key_suffix}", default_column=qb_defaults["invoice"]),
            "amount": select_column("Amount", qb_columns, QB_COLUMN_PATTERNS["amount"], f"qb_map_amount_{key_suffix}", default_column=qb_defaults["amount"]),
            "quantity": select_column("Quantity", qb_columns, QB_COLUMN_PATTERNS["quantity"], f"qb_map_qty_{key_suffix}", optional=True, default_column=qb_defaults["quantity"]),
            "product": select_column(
                "Product description", qb_columns, QB_COLUMN_PATTERNS["product"],
                f"qb_map_product_{key_suffix}", optional=True, default_column=qb_defaults["product"],
            ),
        }
        qb_period_detail, _ = filter_qb_subtotal_rows(qb_raw, qb_mapping)
        detected_qb_period = infer_qb_period_column(
            qb_period_detail,
            excluded_columns=qb_mapping.values(),
        )
        qb_mapping["period"] = select_column(
            "Source fiscal period", qb_columns, QB_COLUMN_PATTERNS["period"],
            f"qb_map_period_{key_suffix}", optional=True, default_column=detected_qb_period,
        )
        st.markdown("**Infinium**")
        inf_mapping = {
            "po": select_column("PO", inf_columns, INF_COLUMN_PATTERNS["po"], f"inf_map_po_{key_suffix}", default_column=inf_defaults["po"]),
            "invoice": select_column("Invoice", inf_columns, INF_COLUMN_PATTERNS["invoice"], f"inf_map_inv_{key_suffix}", default_column=inf_defaults["invoice"]),
            "amount": select_column("Amount", inf_columns, INF_COLUMN_PATTERNS["amount"], f"inf_map_amount_{key_suffix}", default_column=inf_defaults["amount"]),
            "period": select_column(
                "Source fiscal period", inf_columns, INF_COLUMN_PATTERNS["period"],
                f"inf_map_period_{key_suffix}", default_column=inf_defaults["period"],
            ),
        }

    return qb_mapping, inf_mapping


def secondary_mapping_panel(
    qb_secondary_raw: Optional[pd.DataFrame],
    inf_secondary_raw: Optional[pd.DataFrame],
    key_suffix: str,
) -> tuple[Optional[dict[str, Optional[str]]], Optional[dict[str, Optional[str]]]]:
    """Map the three fields used by optional historical matching sources.

    Secondary files do not require fiscal-period, product, or quantity fields:
    they are background evidence used only to clear an opposing primary
    exception through the same strict reference-and-amount rules.
    """

    def render_mapping(
        frame: Optional[pd.DataFrame],
        source: str,
    ) -> Optional[dict[str, Optional[str]]]:
        if frame is None:
            return None
        columns = list(frame.columns)
        patterns = QB_COLUMN_PATTERNS if source == "QB" else INF_COLUMN_PATTERNS
        defaults = _infer_pattern_map(columns, patterns)
        prefix = "qb_secondary" if source == "QB" else "inf_secondary"
        mapping: dict[str, Optional[str]] = {
            "po": select_column(
                "PO", columns, patterns["po"],
                f"{prefix}_map_po_{key_suffix}",
                default_column=defaults["po"],
            ),
            "invoice": select_column(
                "Invoice", columns, patterns["invoice"],
                f"{prefix}_map_inv_{key_suffix}",
                default_column=defaults["invoice"],
            ),
            "amount": select_column(
                "Amount", columns, patterns["amount"],
                f"{prefix}_map_amount_{key_suffix}",
                default_column=defaults["amount"],
            ),
        }
        if source == "QB":
            mapping["quantity"] = select_column(
                "Quantity", columns, patterns["quantity"],
                f"{prefix}_map_quantity_{key_suffix}", optional=True,
                default_column=defaults["quantity"],
            )
        return mapping

    qb_mapping: Optional[dict[str, Optional[str]]] = None
    inf_mapping: Optional[dict[str, Optional[str]]] = None
    if qb_secondary_raw is not None or inf_secondary_raw is not None:
        with st.sidebar.expander("Secondary historical mappings", expanded=False):
            st.caption(
                "Secondary rows are historical background data. Only PO, invoice, "
                "and exact signed amount are considered."
            )
            if qb_secondary_raw is not None:
                st.markdown("**QuickBooks secondary**")
                qb_mapping = render_mapping(qb_secondary_raw, "QB")
            if inf_secondary_raw is not None:
                st.markdown("**Infinium secondary**")
                inf_mapping = render_mapping(inf_secondary_raw, "INF")
    return qb_mapping, inf_mapping
