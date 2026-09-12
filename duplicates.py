"""Production duplicate controls for the QB-to-Infinium reconciliation.

Detection, classification, and disposition are deliberately separate. Strong
same-file groups (PO + invoice + signed cents) retain one deterministic
canonical row and exclude only excess copies. Weaker PO-only or invoice-only
groups remain active and are reported for review. Historical rows may also be
screened against their primary dataset so upload overlap cannot improperly
clear an exception.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from hashlib import sha256
from typing import Any, Hashable, Iterable, Optional

import pandas as pd

__all__ = [
    "AMOUNT_CENTS", "NORM_INV", "NORM_PO", "SOURCE_POS",
    "DUPLICATE_ANALYSIS_COLUMNS", "DUPLICATE_BASIS_CROSS_SCOPE",
    "DUPLICATE_BASIS_INVOICE_ONLY", "DUPLICATE_BASIS_PO_ONLY",
    "DUPLICATE_BASIS_STRICT", "DUPLICATE_RULE_VERSION",
    "DISPOSITION_REVIEW", "DISPOSITION_REVIEW_HOLD", "DISPOSITION_REVIEW_RESOLVED",
    "DuplicateScreeningError", "DuplicateScreeningResult",
    "build_duplicate_item_report", "combine_duplicate_reports",
    "duplicate_row_basis", "duplicate_row_indexes", "finalize_review_dispositions",
    "screen_duplicates", "validate_duplicate_input",
]

SOURCE_POS = "__REC_SOURCE_POS"
NORM_PO = "__REC_NORM_PO"
NORM_INV = "__REC_NORM_INV"
AMOUNT_CENTS = "__REC_AMOUNT_CENTS"

DUPLICATE_RULE_VERSION = "2026.09-CANONICAL-EXCESS-WEAK-REVIEW-CROSS-SCOPE"
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
CONFIDENCE_CONFIRMED = "High"
CONFIDENCE_SUSPECTED = "Review"
CONFIDENCE_HELD = "Hold"
CONFIDENCE_RESOLVED = "Resolved"

# Original columns stay first for compatibility with existing exporters.
DUPLICATE_ANALYSIS_COLUMNS = [
    "Dataset", "Source Scope", "Source Row ID", "Duplicate Basis",
    "Normalized PO", "Normalized Invoice", "Amount", "Duplicate Group Size",
    "Other Source Row IDs In Group", "Treatment", "Duplicate Group ID",
    "Confidence", "Disposition", "Canonical Source Row ID",
    "Reference Source Row IDs", "Automatically Excluded", "Excluded Amount",
    "Payload Confirmed", "Confirmation Fields", "Policy Note",
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
    group_rows: tuple[RowLabel, ...]
    reference_rows: tuple[RowLabel, ...]
    disposition: str
    confidence: str
    canonical_index: Optional[RowLabel]
    payload_confirmed: bool


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


def _confirmation_columns(frame: pd.DataFrame, id_column: str) -> tuple[str, ...]:
    """Use original source columns as corroborating duplicate evidence."""
    internal = {SOURCE_POS, NORM_PO, NORM_INV, AMOUNT_CENTS, id_column}
    return tuple(
        column for column in frame.columns
        if column not in internal and not str(column).startswith("__REC_")
    )


def _payload_value(value: Any) -> str:
    if _is_missing(value):
        return "<NULL>"
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    return f"{type(value).__name__}:{value!s}"


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
    reference_frame: Optional[pd.DataFrame],
    confirmation_columns: tuple[str, ...],
    reference_confirmation_columns: tuple[str, ...],
) -> list[_Decision]:
    current_groups = _groups(frame)
    reference_groups = _groups(reference_frame) if reference_frame is not None else {}
    decisions: list[_Decision] = []
    for key, rows in current_groups.items():
        reference_rows = tuple(reference_groups.get(key, ()))
        if len(rows) < 2 and not reference_rows:
            continue
        group_id = _group_id(key)
        if reference_rows:
            reference_by_payload: dict[tuple[str, ...], list[RowLabel]] = defaultdict(list)
            for reference_index in reference_rows:
                signature = _payload_signature(
                    reference_frame, reference_index, reference_confirmation_columns
                )
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
                    index, DUPLICATE_BASIS_CROSS_SCOPE, key, group_id, tuple(rows),
                    reference_rows, disposition, confidence, canonical,
                    payload_confirmed,
                ))
            continue
        strong = (
            key[0] == DUPLICATE_BASIS_STRICT
            and auto_exclude_strict
            and bool(confirmation_columns)
        )
        payload_groups: dict[tuple[str, ...], list[RowLabel]] = defaultdict(list)
        for index in rows:
            payload_groups[_payload_signature(frame, index, confirmation_columns)].append(index)
        for payload_rows in payload_groups.values():
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
                    index, key[0], key, group_id, tuple(rows), (), disposition,
                    confidence, canonical, payload_confirmed,
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
        records.append({
            "Dataset": dataset_label,
            "Source Scope": source_scope,
            "Source Row ID": frame.at[index, id_column],
            "Duplicate Basis": decision.basis,
            "Normalized PO": frame.at[index, NORM_PO],
            "Normalized Invoice": frame.at[index, NORM_INV],
            "Amount": amount,
            "Duplicate Group Size": len(decision.group_rows) + len(decision.reference_rows),
            "Other Source Row IDs In Group": "; ".join(other_ids),
            "Treatment": decision.disposition,
            "Duplicate Group ID": decision.group_id,
            "Confidence": decision.confidence,
            "Disposition": decision.disposition,
            "Canonical Source Row ID": canonical_id,
            "Reference Source Row IDs": "; ".join(reference_ids),
            "Automatically Excluded": excluded,
            "Excluded Amount": amount if excluded else None,
            "Payload Confirmed": decision.payload_confirmed,
            "Confirmation Fields": "; ".join(map(str, confirmation_columns)),
            "Policy Note": policy_note,
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
        frame, auto_exclude_strict=True, reference_frame=None,
        confirmation_columns=confirmation_columns,
        reference_confirmation_columns=(),
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
    return combined.sort_values(
        ["Dataset", "Source Scope", "Duplicate Group ID", "Source Row ID"],
        kind="stable",
        key=lambda values: values.astype("string"),
    ).reset_index(drop=True)


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
) -> DuplicateScreeningResult:
    """Detect candidates, apply policy, and return active/excluded populations."""
    frame_label = f"{dataset_label} {source_scope}".strip()
    validate_duplicate_input(frame, id_column, frame_label=frame_label)
    if reference_frame is not None:
        if not reference_id_column:
            raise DuplicateScreeningError(
                "reference_id_column is required when reference_frame is supplied."
            )
        validate_duplicate_input(
            reference_frame, reference_id_column,
            frame_label=f"{dataset_label} Primary Reference",
        )
    candidate_confirmation_columns = _confirmation_columns(frame, id_column)
    if reference_frame is None:
        confirmation_columns = candidate_confirmation_columns
    else:
        reference_candidates = set(_confirmation_columns(reference_frame, reference_id_column))
        confirmation_columns = tuple(
            column for column in candidate_confirmation_columns
            if column in reference_candidates
        )
    decisions = _decisions(
        frame, auto_exclude_strict=auto_exclude_strict,
        reference_frame=reference_frame,
        confirmation_columns=confirmation_columns,
        reference_confirmation_columns=confirmation_columns,
    )
    excluded = _ordered_indexes(frame, (
        decision.row_index for decision in decisions
        if decision.disposition in {DISPOSITION_EXCLUDED_EXCESS, DISPOSITION_EXCLUDED_OVERLAP}
    ))
    canonical = _ordered_indexes(frame, (
        decision.row_index for decision in decisions
        if decision.disposition == DISPOSITION_CANONICAL
    ))
    suspected = _ordered_indexes(frame, (
        decision.row_index for decision in decisions
        if decision.disposition == DISPOSITION_REVIEW
    ))
    report = _report_from_decisions(
        frame, decisions, id_column=id_column, dataset_label=dataset_label,
        source_scope=source_scope, policy_note=treatment,
        reference_frame=reference_frame, reference_id_column=reference_id_column,
        confirmation_columns=confirmation_columns,
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
) -> pd.DataFrame:
    """Patch review-tier rows with their final match outcome.

    Screening runs before matching, so a weak-basis candidate's initial
    ``DISPOSITION_REVIEW`` cannot yet know whether it will end up matched or
    left unresolved. Call this once matching (and historical clearance)
    completes:

      * A row in ``resolved_ids`` matched normally -- no exclusion was ever
        needed, so its disposition becomes ``DISPOSITION_REVIEW_RESOLVED``.
      * A row in ``held_ids`` stayed unresolved -- it is pulled from the
        accrual and marked ``DISPOSITION_REVIEW_HOLD``, requiring a
        documented human decision before the proposed JE is posted.

    Only rows currently at ``DISPOSITION_REVIEW`` are touched; canonical,
    excess, and cross-scope-overlap rows are left exactly as screening
    decided them.
    """
    if report.empty:
        return report
    resolved = {str(value) for value in resolved_ids}
    held = {str(value) for value in held_ids}
    updated = report.copy()
    is_review = updated["Disposition"] == DISPOSITION_REVIEW
    row_ids = updated["Source Row ID"].astype(str)
    resolved_mask = is_review & row_ids.isin(resolved)
    held_mask = is_review & row_ids.isin(held)

    updated.loc[resolved_mask, ["Disposition", "Treatment"]] = DISPOSITION_REVIEW_RESOLVED
    updated.loc[resolved_mask, "Confidence"] = CONFIDENCE_RESOLVED
    updated.loc[resolved_mask, "Policy Note"] = (
        "This candidate matched normally during reconciliation; no duplicate "
        "exclusion was ever applied."
    )

    updated.loc[held_mask, ["Disposition", "Treatment"]] = DISPOSITION_REVIEW_HOLD
    updated.loc[held_mask, "Confidence"] = CONFIDENCE_HELD
    updated.loc[held_mask, "Automatically Excluded"] = True
    updated.loc[held_mask, "Excluded Amount"] = updated.loc[held_mask, "Amount"]
    updated.loc[held_mask, "Policy Note"] = (
        "This candidate remained unresolved after matching. It is excluded from the "
        "accrual/proposed journal entry and held in Duplicate Review Hold pending a "
        "documented human disposition."
    )
    return updated
