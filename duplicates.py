"""Production duplicate controls for the QB-to-Infinium reconciliation.

Detection, classification, and disposition are deliberately separate. A shared
PO + invoice + signed-cents key only IDENTIFIES a duplicate relationship; an
excess row is excluded automatically only when it is CONFIRMED to be a copy of
the same underlying LINE, by this hierarchy:

  1. a LINE-LEVEL source ID (one that identifies a single source row) shared by
     the rows -- but only when the data does not contradict that it is line-
     level. A mapped ID that is shared by rows that are not otherwise the same
     line (different key or fingerprint) is invoice/transaction-level, and is
     demoted to supporting evidence;
  2. otherwise an identical fingerprint over an EXPLICIT set of stable, line-
     level transaction attributes (see STABLE_FINGERPRINT_ROLES) in addition to
     the normalized PO + invoice + signed amount -- never "whatever other
     columns the report happens to contain" -- backed by enough identity-grade
     evidence (customer, date, item, rate: see MIN_IDENTITY_FIELDS), and, when a
     transaction-level ID is mapped, that ID agreeing too. Quantity alone, or any
     reporting-context field such as the fiscal period, is never enough;
  3. if neither can be established, the rows stay potential duplicates.

Rows that share the key but are not confirmed copies stay active and are
reported for review (a potential duplicate is never discarded on the key alone).
Weaker PO-only or invoice-only groups likewise remain active and are reported
for review.
Historical screening is explicitly two-stage: same-file copies are handled
first, then the surviving rows are compared with their primary dataset so
upload overlap cannot improperly clear an exception.

This module intentionally makes no cross-system amount-variance decisions.
``matching.py`` consumes these duplicate dispositions first and only then
classifies mutually unique PO/invoice relationships whose amounts differ,
ensuring a duplicate candidate can never be repurposed as an amount error.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from hashlib import sha256
from types import MappingProxyType
from typing import Any, Hashable, Iterable, Mapping, Optional

import pandas as pd

__all__ = [
    "AMOUNT_CENTS", "NORM_INV", "NORM_PO", "SOURCE_POS", "STABLE_FINGERPRINT_ROLES", "IDENTITY_GRADE_ROLES", "MIN_IDENTITY_FIELDS", "LINE_ID", "TXN_ID",
    "DUPLICATE_ANALYSIS_COLUMNS", "DUPLICATE_BASIS_CROSS_SCOPE",
    "DUPLICATE_BASIS_INVOICE_ONLY", "DUPLICATE_BASIS_PO_ONLY",
    "DUPLICATE_BASIS_STRICT", "DUPLICATE_RULE_VERSION",
    "DISPOSITION_HISTORICAL_REVIEW_HOLD", "DISPOSITION_REVIEW",
    "DISPOSITION_REVIEW_HOLD", "DISPOSITION_REVIEW_RESOLVED",
    "DuplicateScreeningError", "DuplicateScreeningResult",
    "build_duplicate_item_report", "combine_duplicate_reports",
    "duplicate_row_basis", "duplicate_row_indexes", "finalize_review_dispositions",
    "screen_duplicates", "validate_duplicate_input",
]

# The mapping roles whose source columns form the duplicate fingerprint. The
# list is deliberate and fixed: only line-level attributes of the transaction
# itself (who, when, which item, how many, at what rate) -- not report or export
# metadata, and not the fiscal period, which is reporting CONTEXT: it can change
# with carryforward or reclassification, or with which file holds the row.
STABLE_FINGERPRINT_ROLES = ("customer", "date", "product", "quantity", "rate")
# The subset that actually distinguishes one line from another. Quantity alone
# is weak (many lines share it), so an automatic exclusion needs at least
# MIN_IDENTITY_FIELDS of these present and identical.
IDENTITY_GRADE_ROLES = ("customer", "date", "product", "rate")
MIN_IDENTITY_FIELDS = 2

SOURCE_POS = "__REC_SOURCE_POS"
LINE_ID = "__REC_LINE_ID"
TXN_ID = "__REC_TXN_ID"
NORM_PO = "__REC_NORM_PO"
NORM_INV = "__REC_NORM_INV"
AMOUNT_CENTS = "__REC_AMOUNT_CENTS"

DUPLICATE_RULE_VERSION = "2026.09-STREAMLINED-DUPLICATE-OUTPUT"
DUPLICATE_BASIS_STRICT = "PO + Invoice + Amount"
DUPLICATE_BASIS_PO_ONLY = "PO + Amount (invoice blank on both rows)"
DUPLICATE_BASIS_INVOICE_ONLY = "Invoice + Amount (PO blank on both rows)"
DUPLICATE_BASIS_CROSS_SCOPE = "Primary/Historical overlap"

DISPOSITION_CANONICAL = "Retained canonical row"
DISPOSITION_EXCLUDED_EXCESS = "Excluded excess copy"
DISPOSITION_REVIEW = "Review required - retained for matching"
DISPOSITION_EXCLUDED_OVERLAP = "Excluded historical overlap"
DISPOSITION_REVIEW_RESOLVED = "Resolved via match - no exclusion applied"
DISPOSITION_REVIEW_HOLD = "Held for review - excluded from proposed JE pending disposition"
DISPOSITION_HISTORICAL_REVIEW_HOLD = (
    "Held for review - excluded from historical clearance pending disposition"
)
CONFIDENCE_CONFIRMED = "High"
CONFIDENCE_SUSPECTED = "Review"
CONFIDENCE_HELD = "Hold"
CONFIDENCE_RESOLVED = "Resolved"

# This is a user-facing review report. Dataset is intentionally omitted because
# QuickBooks and Infinium reports are delivered separately. Disposition is the
# single authoritative status field; the former Treatment duplicate was removed.
# The full confirmation-field inventory is also internal-only; the report shows
# only fields that actually differ, which is the evidence a reviewer needs.
DUPLICATE_ANALYSIS_COLUMNS = [
    "Source Scope", "Screening Stage", "Source Row ID", "Duplicate Basis",
    "Potential Duplicate Reason", "Differing Confirmation Fields",
    "Normalized PO", "Normalized Invoice", "Amount", "Duplicate Values",
    "Duplicate Group Size",
    "Confirmed Copy Set Size", "Other Source Row IDs In Group",
    "Duplicate Group ID", "Confirmed Copy Set ID",
    "Confidence", "Disposition", "Canonical Source Row ID",
    "Reference Source Row IDs", "Automatically Excluded", "Excluded Amount",
    "Payload Confirmed", "Confirmation Basis", "Policy Note",
    "Manual Decision", "Reviewed By", "Review Timestamp", "Review Rationale",
    "Duplicate Rule Version",
]


class DuplicateScreeningError(ValueError):
    """Controlled input-contract failure raised before screening."""


RowLabel = Hashable
GroupKey = tuple[str, str, str, int]


@dataclass(frozen=True, slots=True)
class _Decision:
    row_index: RowLabel
    basis: str
    key: GroupKey
    group_id: str
    copy_set_id: str
    group_rows: tuple[RowLabel, ...]
    copy_set_rows: tuple[RowLabel, ...]
    reference_rows: tuple[RowLabel, ...]
    disposition: str
    confidence: str
    canonical_index: Optional[RowLabel]
    payload_confirmed: bool
    screening_stage: str
    confirmation_basis: str = ""
    insufficient_evidence: bool = False


@dataclass(frozen=True, slots=True)
class DuplicateScreeningResult:
    """Duplicate outcome with compatibility fields for current callers.

    ``duplicate_rows`` now means only automatically excluded rows.
    ``active_frame`` contains canonical and review-required rows.
    """

    duplicate_rows: list[RowLabel]
    active_frame: pd.DataFrame
    report: pd.DataFrame
    canonical_rows: list[RowLabel]
    suspected_rows: list[RowLabel]
    excluded_frame: pd.DataFrame
    suspected_frame: pd.DataFrame

    @property
    def excluded_rows(self) -> list[RowLabel]:
        return list(self.duplicate_rows)

    @property
    def retained_frame(self) -> pd.DataFrame:
        return self.active_frame


def _is_missing(value: Any) -> bool:
    if value is None or value is pd.NA:
        return True
    result = pd.isna(value)
    try:
        return bool(result)
    except (TypeError, ValueError):
        return False


def _amount_as_int(value: Any, *, field_name: str = AMOUNT_CENTS) -> int:
    if isinstance(value, bool):
        raise DuplicateScreeningError(f"{field_name} cannot contain Boolean values.")
    try:
        number = Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise DuplicateScreeningError(
            f"{field_name} contains a nonnumeric value: {value!r}."
        ) from exc
    if not number.is_finite() or number != number.to_integral_value():
        raise DuplicateScreeningError(
            f"{field_name} must contain whole signed cents; received {value!r}."
        )
    return int(number)


def cents_to_float(value: Any) -> Optional[float]:
    """Convert validated integer cents for presentation without masking nulls."""
    if _is_missing(value):
        return None
    return float(Decimal(_amount_as_int(value)) / Decimal(100))


def validate_duplicate_input(
    frame: pd.DataFrame,
    id_column: str,
    *,
    frame_label: str = "dataset",
    allow_missing_amounts: bool = True,
) -> None:
    """Validate required schema and identity invariants."""
    if not isinstance(frame, pd.DataFrame):
        raise DuplicateScreeningError(f"{frame_label} must be a pandas DataFrame.")
    required = {SOURCE_POS, NORM_PO, NORM_INV, AMOUNT_CENTS, id_column}
    missing = sorted(required.difference(frame.columns))
    if missing:
        raise DuplicateScreeningError(
            f"{frame_label} is missing duplicate-screening columns: {missing}."
        )
    if not frame.index.is_unique:
        raise DuplicateScreeningError(
            f"{frame_label} has repeated DataFrame index labels. Reset or uniquely "
            "key the index before duplicate screening."
        )
    ids = frame[id_column].astype("string")
    if frame[id_column].isna().any() or ids.str.strip().eq("").any():
        raise DuplicateScreeningError(f"{frame_label} contains blank source row IDs.")
    if frame[id_column].duplicated().any():
        raise DuplicateScreeningError(f"{frame_label} contains repeated source row IDs.")

    positions = pd.to_numeric(frame[SOURCE_POS], errors="coerce")
    if positions.isna().any() or positions.mod(1).ne(0).any():
        raise DuplicateScreeningError(
            f"{frame_label} source positions must be populated whole numbers."
        )
    if positions.duplicated().any():
        raise DuplicateScreeningError(f"{frame_label} contains repeated source positions.")

    for column in (NORM_PO, NORM_INV):
        for value in frame[column]:
            if _is_missing(value):
                continue
            if not isinstance(value, str):
                raise DuplicateScreeningError(
                    f"{frame_label} column {column} must contain normalized strings or nulls."
                )
            if value != value.strip():
                raise DuplicateScreeningError(
                    f"{frame_label} column {column} contains untrimmed values."
                )

    for value in frame[AMOUNT_CENTS]:
        if _is_missing(value):
            if allow_missing_amounts:
                continue
            raise DuplicateScreeningError(f"{frame_label} contains missing amount cents.")
        _amount_as_int(value)


def _basis_for_values(po: Any, invoice: Any, amount: Any) -> Optional[str]:
    if _is_missing(amount):
        return None
    has_po = not _is_missing(po) and po != ""
    has_invoice = not _is_missing(invoice) and invoice != ""
    if has_po and has_invoice:
        return DUPLICATE_BASIS_STRICT
    if has_po:
        return DUPLICATE_BASIS_PO_ONLY
    if has_invoice:
        return DUPLICATE_BASIS_INVOICE_ONLY
    return None


def _key_for_row(frame: pd.DataFrame, index: RowLabel) -> Optional[GroupKey]:
    po = frame.at[index, NORM_PO]
    invoice = frame.at[index, NORM_INV]
    amount = frame.at[index, AMOUNT_CENTS]
    basis = _basis_for_values(po, invoice, amount)
    if basis is None:
        return None
    return (
        basis,
        str(po) if basis != DUPLICATE_BASIS_INVOICE_ONLY else "",
        str(invoice) if basis != DUPLICATE_BASIS_PO_ONLY else "",
        _amount_as_int(amount),
    )


def _ordered_indexes(frame: pd.DataFrame, indexes: Iterable[RowLabel]) -> list[RowLabel]:
    return sorted(
        indexes,
        key=lambda idx: (int(Decimal(str(frame.at[idx, SOURCE_POS]))), repr(idx)),
    )


def _group_id(key: GroupKey) -> str:
    payload = "\x1f".join(map(str, key))
    return f"DUP-{sha256(payload.encode('utf-8')).hexdigest()[:12].upper()}"


def _copy_set_id(key: GroupKey, signature: tuple[str, ...]) -> str:
    """Identify one payload-specific copy set within a broader business key."""
    payload = "\x1f".join([*(map(str, key)), "\x1e", *signature])
    return f"COPY-{sha256(payload.encode('utf-8')).hexdigest()[:16].upper()}"


def _confirmation_columns(
    frame: pd.DataFrame, id_column: str, fingerprint_columns: Iterable[str] = (),
) -> tuple[str, ...]:
    """The explicit stable-field columns that make up the fingerprint, in the
    order given, limited to those actually present. The frame's other columns
    are never consulted -- a report-layout change or an added informational
    column cannot alter duplicate behavior. With none, only a native
    transaction ID can confirm a copy."""
    internal = {SOURCE_POS, NORM_PO, NORM_INV, AMOUNT_CENTS, id_column}
    return tuple(
        column for column in dict.fromkeys(fingerprint_columns)
        if column and column in frame.columns and column not in internal
        and not str(column).startswith("__REC_")
    )


_RE_DATE_TEXT = re.compile(
    r"^\s*(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})/(\d{1,2})/(\d{2}|\d{4}))"
    r"(?:[ T]\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)?\s*$", re.IGNORECASE,
)


def _payload_value(value: Any) -> str:
    """A stable, formatting-insensitive form of one fingerprint value, so that
    "Acme  Corp" vs "ACME CORP", 1 vs 1.0, or 1/5/2026 vs 2026-01-05 (formatting
    only) never keep two copies of one record from being confirmed."""
    if _is_missing(value):
        return "<NULL>"
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return f"date:{pd.Timestamp(value).date().isoformat()}"
    if isinstance(value, bool):
        return f"bool:{value}"
    if isinstance(value, (int, float, Decimal)):
        return f"num:{Decimal(str(value)).normalize()}"
    text = " ".join(str(value).split())
    if text == "":
        return "<NULL>"
    if _RE_DATE_TEXT.match(text):
        parsed = pd.to_datetime(text, errors="coerce")
        if not pd.isna(parsed):
            return f"date:{parsed.date().isoformat()}"
    try:
        return f"num:{Decimal(text.replace(',', '')).normalize()}"
    except InvalidOperation:
        return f"text:{text.upper()}"


def _id_values(
    frame: pd.DataFrame, rows: Iterable[RowLabel], column: str,
) -> Optional[dict[RowLabel, str]]:
    """The value of an ID column on every row of a duplicate-key group, or None
    unless the column is mapped AND populated on every row -- a group with even
    one blank ID never treats blank as a shared identity."""
    if column not in frame.columns:
        return None
    found: dict[RowLabel, str] = {}
    for index in rows:
        value = frame.at[index, column]
        if _is_missing(value) or str(value).strip() == "":
            return None
        found[index] = str(value).strip()
    return found


def _line_id_is_trusted(frame: pd.DataFrame, confirmation_columns: tuple[str, ...]) -> bool:
    """Whether a mapped line-level ID may independently confirm a copy.

    A genuine line-level ID is shared only by copies of the same line, so every
    set of rows carrying one ID must share the duplicate key and the stable
    fingerprint. If any shared ID spans rows that differ -- two different sales
    lines on one invoice, say -- the column is invoice/transaction-level and
    cannot be trusted to identify a line; it is demoted to supporting evidence."""
    if LINE_ID not in frame.columns:
        return False
    by_id: dict[str, list[RowLabel]] = defaultdict(list)
    for index in frame.index:
        value = frame.at[index, LINE_ID]
        if not _is_missing(value) and str(value).strip() != "":
            by_id[str(value).strip()].append(index)
    for rows in by_id.values():
        if len(rows) < 2:
            continue
        keys = {_key_for_row(frame, index) for index in rows}
        prints = {_payload_signature(frame, index, confirmation_columns) for index in rows}
        if len(keys) != 1 or None in keys or len(prints) != 1:
            return False
    return True


def _payload_signature(
    frame: pd.DataFrame,
    index: RowLabel,
    columns: tuple[str, ...],
) -> tuple[str, ...]:
    return tuple(_payload_value(frame.at[index, column]) for column in columns)


def _groups(frame: pd.DataFrame) -> dict[GroupKey, list[RowLabel]]:
    grouped: dict[GroupKey, list[RowLabel]] = defaultdict(list)
    for index in frame.index:
        key = _key_for_row(frame, index)
        if key is not None:
            grouped[key].append(index)
    return {key: _ordered_indexes(frame, rows) for key, rows in grouped.items()}


def duplicate_row_basis(frame: pd.DataFrame) -> dict[RowLabel, str]:
    """Return same-frame candidates and evidence basis; no exclusion implied."""
    required = {SOURCE_POS, NORM_PO, NORM_INV, AMOUNT_CENTS}
    missing = sorted(required.difference(frame.columns))
    if missing:
        raise DuplicateScreeningError(f"Dataset is missing columns: {missing}.")
    if not frame.index.is_unique:
        raise DuplicateScreeningError("Dataset has repeated DataFrame index labels.")
    result: dict[RowLabel, str] = {}
    for key, rows in _groups(frame).items():
        if len(rows) > 1:
            for index in rows:
                result[index] = key[0]
    return result


def duplicate_row_indexes(frame: pd.DataFrame) -> set[RowLabel]:
    """Return detected candidates, without implying automatic exclusion."""
    return set(duplicate_row_basis(frame))


def _decisions(
    frame: pd.DataFrame,
    *,
    auto_exclude_strict: bool,
    confirmation_columns: tuple[str, ...],
    identity_columns: tuple[str, ...] = (),
    min_identity_fields: int = 0,
) -> list[_Decision]:
    """Classify duplicates within one file only.

    Cross-scope overlap is intentionally handled by
    ``_cross_scope_decisions`` after same-file excess copies are removed.
    Keeping the stages separate prevents the existence of a primary reference
    from bypassing historical same-file canonicalization.
    """
    current_groups = _groups(frame)
    line_id_trusted = _line_id_is_trusted(frame, confirmation_columns)
    decisions: list[_Decision] = []
    for key, rows in current_groups.items():
        if len(rows) < 2:
            continue
        group_id = _group_id(key)
        # The key only identifies a duplicate RELATIONSHIP. Rows are confirmed
        # copies of one underlying line only through (1) a trusted line-level ID
        # or (2) a sufficient stable-field fingerprint (with any mapped
        # transaction-level or untrusted ID agreeing). Rows that share the key
        # but not that identity are distinct copy sets -- potential duplicates --
        # and are never excluded on the key alone.
        line_ids = _id_values(frame, rows, LINE_ID) if line_id_trusted else None
        supporting: list[dict[RowLabel, str]] = []
        if line_ids is None:
            for column in (TXN_ID, LINE_ID):
                values = _id_values(frame, rows, column)
                if values is not None:
                    supporting.append(values)
        identity_present = sum(
            1 for column in identity_columns
            if column in confirmation_columns
            and all(_payload_value(frame.at[index, column]) != "<NULL>" for index in rows)
        )
        # Callers pass MIN_IDENTITY_FIELDS for every dataset: an exclusion needs a
        # trusted line-level ID or enough identity-grade evidence. (0 disables the
        # requirement and exists only for direct unit tests of the mechanism.)
        sufficient = line_ids is not None or identity_present >= min_identity_fields
        strong = key[0] == DUPLICATE_BASIS_STRICT and auto_exclude_strict and sufficient
        if line_ids is not None:
            confirmation_basis = "Line-level source ID"
        else:
            confirmation_basis = f"Stable-field fingerprint ({', '.join(confirmation_columns)})"
            if supporting:
                confirmation_basis += " + transaction ID"

        def signature_for(index: RowLabel) -> tuple[str, ...]:
            if line_ids is not None:
                return (line_ids[index],)
            return (
                *(values[index] for values in supporting),
                *_payload_signature(frame, index, confirmation_columns),
            )

        payload_groups: dict[tuple[str, ...], list[RowLabel]] = defaultdict(list)
        for index in rows:
            payload_groups[signature_for(index)].append(index)
        for payload_rows in payload_groups.values():
            signature = signature_for(payload_rows[0])
            copy_set_id = _copy_set_id(key, signature)
            payload_confirmed = strong and len(payload_rows) > 1
            canonical = payload_rows[0] if payload_confirmed else None
            for index in payload_rows:
                if payload_confirmed and index == canonical:
                    disposition, confidence = DISPOSITION_CANONICAL, CONFIDENCE_CONFIRMED
                elif payload_confirmed:
                    disposition, confidence = DISPOSITION_EXCLUDED_EXCESS, CONFIDENCE_CONFIRMED
                else:
                    disposition, confidence = DISPOSITION_REVIEW, CONFIDENCE_SUSPECTED
                decisions.append(_Decision(
                    index, key[0], key, group_id, copy_set_id, tuple(rows),
                    tuple(payload_rows), (),
                    disposition, confidence, canonical, payload_confirmed, "Same-file",
                    confirmation_basis if payload_confirmed else "",
                    key[0] == DUPLICATE_BASIS_STRICT and auto_exclude_strict and not sufficient,
                ))
    return decisions


def _cross_scope_decisions(
    frame: pd.DataFrame,
    reference_frame: pd.DataFrame,
    *,
    confirmation_columns: tuple[str, ...],
) -> list[_Decision]:
    """Compare already canonicalized historical rows with primary rows."""
    current_groups = _groups(frame)
    reference_groups = _groups(reference_frame)
    decisions: list[_Decision] = []
    for key, rows in current_groups.items():
        reference_rows = tuple(reference_groups.get(key, ()))
        if not reference_rows:
            continue
        group_id = _group_id(key)
        reference_by_payload: dict[tuple[str, ...], list[RowLabel]] = defaultdict(list)
        current_by_payload: dict[tuple[str, ...], list[RowLabel]] = defaultdict(list)
        for current_index in rows:
            current_by_payload[
                _payload_signature(frame, current_index, confirmation_columns)
            ].append(current_index)
        for reference_index in reference_rows:
            signature = _payload_signature(reference_frame, reference_index, confirmation_columns)
            reference_by_payload[signature].append(reference_index)
        for index in rows:
            signature = _payload_signature(frame, index, confirmation_columns)
            confirmed_reference_rows = tuple(reference_by_payload.get(signature, ()))
            payload_confirmed = bool(confirmation_columns and confirmed_reference_rows)
            if payload_confirmed:
                disposition = DISPOSITION_EXCLUDED_OVERLAP
                confidence = CONFIDENCE_CONFIRMED
                canonical = confirmed_reference_rows[0]
            else:
                disposition = DISPOSITION_REVIEW
                confidence = CONFIDENCE_SUSPECTED
                canonical = None
            decisions.append(_Decision(
                index, DUPLICATE_BASIS_CROSS_SCOPE, key, group_id,
                _copy_set_id(key, signature), tuple(rows),
                tuple(current_by_payload[signature]), reference_rows,
                disposition, confidence, canonical, payload_confirmed, "Cross-scope",
            ))
    return decisions


def _report_from_decisions(
    frame: pd.DataFrame,
    decisions: Iterable[_Decision],
    *,
    id_column: str,
    dataset_label: str,
    source_scope: str,
    policy_note: str,
    reference_frame: Optional[pd.DataFrame],
    reference_id_column: Optional[str],
    confirmation_columns: tuple[str, ...],
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    ordered = sorted(
        decisions,
        key=lambda item: (item.group_id, int(frame.at[item.row_index, SOURCE_POS])),
    )
    for decision in ordered:
        index = decision.row_index
        other_ids = [str(frame.at[row, id_column]) for row in decision.group_rows if row != index]
        reference_ids: list[str] = []
        canonical_id: Any = None
        if decision.reference_rows and reference_frame is not None and reference_id_column:
            reference_ids = [str(reference_frame.at[row, reference_id_column]) for row in decision.reference_rows]
            if decision.canonical_index is not None:
                canonical_id = reference_frame.at[decision.canonical_index, reference_id_column]
        elif decision.canonical_index is not None:
            canonical_id = frame.at[decision.canonical_index, id_column]
        excluded = decision.disposition in {DISPOSITION_EXCLUDED_EXCESS, DISPOSITION_EXCLUDED_OVERLAP}
        amount = cents_to_float(frame.at[index, AMOUNT_CENTS])
        if decision.basis == DUPLICATE_BASIS_PO_ONLY:
            duplicate_reason = (
                "Same normalized PO and signed amount; invoice is blank on all "
                "candidate rows."
            )
        elif decision.basis == DUPLICATE_BASIS_INVOICE_ONLY:
            duplicate_reason = (
                "Same normalized invoice and signed amount; PO is blank on all "
                "candidate rows."
            )
        elif decision.basis == DUPLICATE_BASIS_CROSS_SCOPE:
            duplicate_reason = (
                "Historical and primary rows share the same duplicate business key, "
                + (
                    "and their confirmation payloads agree."
                    if decision.payload_confirmed
                    else "but their confirmation payloads do not fully agree."
                )
            )
        elif decision.payload_confirmed:
            duplicate_reason = (
                "Same normalized PO, invoice, signed amount, and line-level source ID."
                if decision.confirmation_basis == "Line-level source ID"
                else "Same normalized PO, invoice, signed amount, and stable-field fingerprint."
            )
        elif decision.insufficient_evidence:
            duplicate_reason = (
                "Same normalized PO, invoice, and signed amount, but the available fields are not "
                "sufficient to establish that these rows are copies of one underlying line."
            )
        else:
            duplicate_reason = (
                "Same normalized PO, invoice, and signed amount, but one or more "
                "confirmation fields differ."
            )

        differing_fields: list[str] = []
        for column in confirmation_columns:
            current_value = _payload_value(frame.at[index, column])
            if decision.reference_rows and reference_frame is not None:
                comparison_values = {
                    _payload_value(reference_frame.at[row, column])
                    for row in decision.reference_rows
                }
            else:
                # A confirmed copy set is compared against itself only, so an
                # unrelated payload subgroup sharing the same duplicate key
                # can't be mistaken for a differing field on a confirmed row.
                # An unconfirmed (Review) row is compared against the full
                # duplicate-key group, to surface which field(s) disagree.
                comparison_rows = (
                    decision.copy_set_rows if decision.payload_confirmed else decision.group_rows
                )
                comparison_values = {
                    _payload_value(frame.at[row, column])
                    for row in comparison_rows
                }
            if current_value not in comparison_values or len(comparison_values) > 1:
                differing_fields.append(str(column))
        records.append({
            "Source Scope": source_scope,
            "Screening Stage": decision.screening_stage,
            "Source Row ID": frame.at[index, id_column],
            "Duplicate Basis": decision.basis,
            "Potential Duplicate Reason": duplicate_reason,
            "Differing Confirmation Fields": "; ".join(differing_fields),
            "Normalized PO": frame.at[index, NORM_PO],
            "Normalized Invoice": frame.at[index, NORM_INV],
            "Amount": amount,
            "Duplicate Values": (
                f"PO={frame.at[index, NORM_PO] or '<BLANK>'} | "
                f"Invoice={frame.at[index, NORM_INV] or '<BLANK>'} | "
                f"Signed Amount={amount if amount is not None else '<INVALID>'}"
            ),
            "Duplicate Group Size": len(decision.group_rows) + len(decision.reference_rows),
            "Confirmed Copy Set Size": len(decision.copy_set_rows),
            "Other Source Row IDs In Group": "; ".join(other_ids),
            "Duplicate Group ID": decision.group_id,
            "Confirmed Copy Set ID": decision.copy_set_id,
            "Confidence": decision.confidence,
            "Disposition": decision.disposition,
            "Canonical Source Row ID": canonical_id,
            "Reference Source Row IDs": "; ".join(reference_ids),
            "Automatically Excluded": excluded,
            "Excluded Amount": amount if excluded else None,
            "Payload Confirmed": decision.payload_confirmed,
            "Confirmation Basis": decision.confirmation_basis,
            "Policy Note": policy_note,
            "Manual Decision": None,
            "Reviewed By": None,
            "Review Timestamp": None,
            "Review Rationale": None,
            "Duplicate Rule Version": DUPLICATE_RULE_VERSION,
        })
    return pd.DataFrame(records, columns=DUPLICATE_ANALYSIS_COLUMNS)


def build_duplicate_item_report(
    frame: pd.DataFrame,
    duplicate_rows: set[RowLabel],
    id_column: str,
    dataset_label: str,
    source_scope: str,
    treatment: str,
) -> pd.DataFrame:
    """Compatibility builder for already-detected same-frame candidates."""
    validate_duplicate_input(frame, id_column, frame_label=f"{dataset_label} {source_scope}")
    confirmation_columns = _confirmation_columns(frame, id_column)
    decisions = _decisions(
        frame, auto_exclude_strict=True, confirmation_columns=confirmation_columns,
    )
    selected = [decision for decision in decisions if decision.row_index in duplicate_rows]
    return _report_from_decisions(
        frame, selected, id_column=id_column, dataset_label=dataset_label,
        source_scope=source_scope, policy_note=treatment, reference_frame=None,
        reference_id_column=None, confirmation_columns=confirmation_columns,
    )


def combine_duplicate_reports(*reports: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Combine reports with stable columns and deterministic ordering."""
    populated = [report for report in reports if report is not None and not report.empty]
    if not populated:
        return pd.DataFrame(columns=DUPLICATE_ANALYSIS_COLUMNS)
    combined = pd.concat(populated, ignore_index=True).reindex(columns=DUPLICATE_ANALYSIS_COLUMNS)
    # IDs are deliberately short (QB-1, QB-2, ...), so sort by their numeric
    # suffix to keep QB-10 after QB-9 instead of after QB-1.
    combined["__SOURCE_ROW_SORT"] = pd.to_numeric(
        combined["Source Row ID"].astype("string").str.extract(r"(\d+)$", expand=False),
        errors="coerce",
    )
    return combined.sort_values(
        ["Source Scope", "Duplicate Group ID", "__SOURCE_ROW_SORT", "Source Row ID"],
        kind="stable",
    ).drop(columns="__SOURCE_ROW_SORT").reset_index(drop=True)


def screen_duplicates(
    frame: pd.DataFrame,
    id_column: str,
    dataset_label: str,
    source_scope: str,
    treatment: str = "",
    *,
    auto_exclude_strict: bool = True,
    reference_frame: Optional[pd.DataFrame] = None,
    reference_id_column: Optional[str] = None,
    allow_missing_amounts: bool = True,
    fingerprint_columns: Iterable[str] = (),
    identity_columns: Iterable[str] = (),
    min_identity_fields: int = 0,
) -> DuplicateScreeningResult:
    """Detect candidates, apply policy, and return active/excluded populations."""
    frame_label = f"{dataset_label} {source_scope}".strip()
    validate_duplicate_input(
        frame, id_column, frame_label=frame_label,
        allow_missing_amounts=allow_missing_amounts,
    )
    if reference_frame is not None:
        if not reference_id_column:
            raise DuplicateScreeningError(
                "reference_id_column is required when reference_frame is supplied."
            )
        validate_duplicate_input(
            reference_frame, reference_id_column,
            frame_label=f"{dataset_label} Primary Reference",
            allow_missing_amounts=allow_missing_amounts,
        )
    fingerprint_columns = tuple(column for column in fingerprint_columns if column)
    candidate_confirmation_columns = _confirmation_columns(frame, id_column, fingerprint_columns)
    if reference_frame is None:
        confirmation_columns = candidate_confirmation_columns
    else:
        reference_candidates = set(
            _confirmation_columns(reference_frame, reference_id_column, fingerprint_columns)
        )
        confirmation_columns = tuple(
            column for column in candidate_confirmation_columns
            if column in reference_candidates
        )
    same_file_decisions = _decisions(
        frame, auto_exclude_strict=auto_exclude_strict,
        confirmation_columns=candidate_confirmation_columns,
        identity_columns=tuple(column for column in identity_columns if column),
        min_identity_fields=min_identity_fields,
    )
    same_file_excluded = {
        decision.row_index for decision in same_file_decisions
        if decision.disposition == DISPOSITION_EXCLUDED_EXCESS
    }
    overlap_decisions: list[_Decision] = []
    if reference_frame is not None:
        overlap_candidates = frame.drop(index=list(same_file_excluded)).copy()
        overlap_decisions = _cross_scope_decisions(
            overlap_candidates,
            reference_frame,
            confirmation_columns=confirmation_columns,
        )
    decisions = [*same_file_decisions, *overlap_decisions]
    excluded = _ordered_indexes(frame, (
        decision.row_index for decision in decisions
        if decision.disposition in {DISPOSITION_EXCLUDED_EXCESS, DISPOSITION_EXCLUDED_OVERLAP}
    ))
    excluded_set = set(excluded)
    canonical = _ordered_indexes(frame, (
        decision.row_index for decision in decisions
        if decision.disposition == DISPOSITION_CANONICAL
        and decision.row_index not in excluded_set
    ))
    suspected = _ordered_indexes(frame, set(
        decision.row_index for decision in decisions
        if decision.disposition == DISPOSITION_REVIEW
        and decision.row_index not in excluded_set
    ))
    # Same-file decisions were classified against the candidate frame's own
    # (wider) confirmation columns, so they must be reported against that same
    # set -- reporting them against the reference-narrowed set could hide the
    # very field that drove the classification. Cross-scope decisions were
    # classified against the narrower set and are reported the same way.
    report = pd.concat(
        [
            _report_from_decisions(
                frame, same_file_decisions, id_column=id_column, dataset_label=dataset_label,
                source_scope=source_scope, policy_note=treatment,
                reference_frame=None, reference_id_column=None,
                confirmation_columns=candidate_confirmation_columns,
            ),
            _report_from_decisions(
                frame, overlap_decisions, id_column=id_column, dataset_label=dataset_label,
                source_scope=source_scope, policy_note=treatment,
                reference_frame=reference_frame, reference_id_column=reference_id_column,
                confirmation_columns=confirmation_columns,
            ),
        ],
        ignore_index=True,
    )
    return DuplicateScreeningResult(
        duplicate_rows=excluded,
        active_frame=frame.drop(index=excluded).copy(),
        report=report,
        canonical_rows=canonical,
        suspected_rows=suspected,
        excluded_frame=frame.loc[excluded].copy(),
        suspected_frame=frame.loc[suspected].copy(),
    )


def finalize_review_dispositions(
    report: pd.DataFrame,
    *,
    resolved_ids: Iterable[Any] = (),
    held_ids: Iterable[Any] = (),
    historical_hold_ids: Iterable[Any] = (),
    excess_ids: Mapping[Any, Any] = MappingProxyType({}),
    canonical_survivor_ids: Iterable[Any] = (),
    held_canonical_ids: Mapping[Any, Any] = MappingProxyType({}),
) -> pd.DataFrame:
    """Patch review-tier rows with their final match outcome.

    Screening runs before matching, so a weak-basis candidate's initial
    ``DISPOSITION_REVIEW`` cannot yet know whether it will end up matched or
    left unresolved. Call this once matching (and historical clearance)
    completes:

      * A row in ``resolved_ids`` matched normally -- no exclusion was ever
        needed, so its disposition becomes ``DISPOSITION_REVIEW_RESOLVED``.
        Only rows that themselves matched belong here; an unmatched sibling
        of a matched member is passed in ``held_ids`` instead -- it is
        already-represented evidence, not a fresh transaction.
      * A row in ``held_ids`` stayed unresolved -- it is pulled from the
        accrual and marked ``DISPOSITION_REVIEW_HOLD``, requiring a
        documented human decision before the proposed JE is posted.
      * A row keyed in ``excess_ids`` (mapping excess row ID -> the
        canonical row ID retained in its place) belongs to a group where NO
        member matched at all -- with zero corroborating evidence the group
        is more than one real transaction, only the earliest member is kept
        and every other member is excluded here, exactly as a confirmed
        duplicate always has been (``DISPOSITION_EXCLUDED_EXCESS``, no
        review needed, with ``Canonical Source Row ID`` set to the survivor).
      * A row in ``canonical_survivor_ids`` is that earliest-listed
        survivor -- nothing in its group matched, so calling it "resolved
        via match" would be inaccurate; it becomes ``DISPOSITION_CANONICAL``
        instead, same label a pre-matching canonical row has always had.

    Only rows currently at ``DISPOSITION_REVIEW`` are touched; canonical,
    excess, and cross-scope-overlap rows are left exactly as screening
    decided them.
    """
    if report.empty:
        return report
    resolved = {str(value) for value in resolved_ids}
    held = {str(value) for value in held_ids}
    historical_held = {str(value) for value in historical_hold_ids}
    excess_canonical_by_id = {str(key): str(value) for key, value in excess_ids.items()}
    excess = set(excess_canonical_by_id)
    canonical_survivors = {str(value) for value in canonical_survivor_ids}
    all_sets = (resolved, held, historical_held, excess, canonical_survivors)
    overlaps: set[str] = set()
    for i, left in enumerate(all_sets):
        for right in all_sets[i + 1:]:
            overlaps |= left & right
    if overlaps:
        raise DuplicateScreeningError(
            "Duplicate disposition IDs must be mutually exclusive; overlapping IDs: "
            f"{sorted(overlaps)}."
        )
    updated = report.copy()
    is_review = updated["Disposition"] == DISPOSITION_REVIEW
    row_ids = updated["Source Row ID"].astype(str)
    resolved_mask = is_review & row_ids.isin(resolved)
    held_mask = is_review & row_ids.isin(held)
    historical_held_mask = is_review & row_ids.isin(historical_held)
    excess_mask = is_review & row_ids.isin(excess)
    canonical_mask = is_review & row_ids.isin(canonical_survivors)

    updated.loc[resolved_mask, "Disposition"] = DISPOSITION_REVIEW_RESOLVED
    updated.loc[resolved_mask, "Confidence"] = CONFIDENCE_RESOLVED
    updated.loc[resolved_mask, "Policy Note"] = (
        "This candidate matched normally during reconciliation, or a group sibling did, "
        "which is real evidence the pattern recurs; no duplicate exclusion was applied."
    )

    updated.loc[held_mask, "Disposition"] = DISPOSITION_REVIEW_HOLD
    updated.loc[held_mask, "Confidence"] = CONFIDENCE_HELD
    updated.loc[held_mask, "Automatically Excluded"] = True
    updated.loc[held_mask, "Excluded Amount"] = updated.loc[held_mask, "Amount"]
    updated.loc[held_mask, "Policy Note"] = (
        "This row shares its duplicate key with another row but is not confirmed to be a copy of "
        "the same underlying transaction. It is not discarded; it is excluded from the accrual/"
        "proposed journal entry and held in Duplicate Review Hold pending a documented human "
        "disposition."
    )
    held_canonical = {str(key): str(value) for key, value in held_canonical_ids.items()}
    if held_canonical:
        held_ids_present = row_ids.loc[held_mask]
        updated.loc[held_mask, "Canonical Source Row ID"] = held_ids_present.map(held_canonical).fillna("")

    updated.loc[historical_held_mask, "Disposition"] = (
        DISPOSITION_HISTORICAL_REVIEW_HOLD
    )
    updated.loc[historical_held_mask, "Confidence"] = CONFIDENCE_HELD
    updated.loc[historical_held_mask, "Automatically Excluded"] = True
    updated.loc[historical_held_mask, "Excluded Amount"] = updated.loc[
        historical_held_mask, "Amount"
    ]
    updated.loc[historical_held_mask, "Policy Note"] = (
        "This historical candidate requires review and was not permitted to clear a "
        "primary exception. A documented disposition is required before relying on "
        "the historical-clearance result."
    )

    updated.loc[excess_mask, "Disposition"] = DISPOSITION_EXCLUDED_EXCESS
    updated.loc[excess_mask, "Confidence"] = CONFIDENCE_CONFIRMED
    updated.loc[excess_mask, "Automatically Excluded"] = True
    updated.loc[excess_mask, "Excluded Amount"] = updated.loc[excess_mask, "Amount"]
    updated.loc[excess_mask, "Canonical Source Row ID"] = row_ids.loc[excess_mask].map(excess_canonical_by_id)
    updated.loc[excess_mask, "Policy Note"] = (
        "No member of this duplicate-key group matched during reconciliation, so there is no "
        "evidence more than one real transaction exists. The earliest-listed member is retained "
        "in the accrual; this excess copy is excluded, matching how a confirmed duplicate has "
        "always been treated."
    )

    updated.loc[canonical_mask, "Disposition"] = DISPOSITION_CANONICAL
    updated.loc[canonical_mask, "Confidence"] = CONFIDENCE_CONFIRMED
    updated.loc[canonical_mask, "Policy Note"] = (
        "No member of this duplicate-key group matched during reconciliation. This earliest-"
        "listed member is retained as the group's representative; every other member is a "
        "potential duplicate held for review."
    )
    return updated
