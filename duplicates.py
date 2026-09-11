"""Duplicate detection, exclusion, and reporting for the QB-to-Infinium engine.

This module owns every rule about what counts as a duplicate and how a
duplicate is treated once found. ``matching.py`` depends on it (for the
working-frame column-name constants as well as the screening functions) but
this module never depends on ``matching.py``, so there is no import cycle.

Rules implemented here:

1. A duplicate is excluded from the accrual total and the proposed journal
   entry amount, but is itemized and listed as its own item rather than
   silently dropped.
2. Two rows that merely share a PO (or an invoice) with a *different*
   amount are never assumed to be duplicates -- they are left for
   ``matching.py``'s grouped aggregate passes, which test whether several
   such rows sum exactly to one matching entry on the opposing side.
3. QuickBooks duplicates and Infinium duplicates are reported separately
   (two distinct, independently filterable reports) because their
   downstream treatment is entirely different.
4. Duplicate screening is applied uniformly to every dataset that enters
   the reconciliation -- the two primary files and any optional historical
   (secondary) files -- because a duplicated historical row is just as
   capable of improperly clearing a real primary exception as a duplicated
   primary row is of misstating the accrual.
5. A row only qualifies as a duplicate when its amount matches another row
   exactly AND at least one populated reference field (PO or invoice) also
   matches exactly. A row missing one reference field can still be caught
   against other rows that are consistently missing that same field; a row
   missing *both* reference fields is never flagged, since a shared amount
   alone is too weak a signal to safely exclude from the accrual.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from typing import Any

import pandas as pd

__all__ = [
    "AMOUNT_CENTS",
    "NORM_INV",
    "NORM_PO",
    "SOURCE_POS",
    "DUPLICATE_ANALYSIS_COLUMNS",
    "DUPLICATE_BASIS_INVOICE_ONLY",
    "DUPLICATE_BASIS_PO_ONLY",
    "DUPLICATE_BASIS_STRICT",
    "DuplicateScreeningResult",
    "build_duplicate_item_report",
    "combine_duplicate_reports",
    "duplicate_row_basis",
    "duplicate_row_indexes",
    "screen_duplicates",
]

# Working-frame column names. These are the single source of truth for the
# reconciliation engine -- matching.py imports them from here rather than
# defining its own copies, so the two modules can never drift apart.
SOURCE_POS = "__REC_SOURCE_POS"
NORM_PO = "__REC_NORM_PO"
NORM_INV = "__REC_NORM_INV"
AMOUNT_CENTS = "__REC_AMOUNT_CENTS"

DUPLICATE_BASIS_STRICT = "PO + Invoice + Amount"
DUPLICATE_BASIS_PO_ONLY = "PO + Amount (invoice blank on both rows)"
DUPLICATE_BASIS_INVOICE_ONLY = "Invoice + Amount (PO blank on both rows)"

DUPLICATE_ANALYSIS_COLUMNS = [
    "Dataset",
    "Source Scope",
    "Source Row ID",
    "Duplicate Basis",
    "Normalized PO",
    "Normalized Invoice",
    "Amount",
    "Duplicate Group Size",
    "Other Source Row IDs In Group",
    "Treatment",
]


def cents_to_float(value: Any) -> float:
    if value is None or pd.isna(value):
        return 0.0
    return float(Decimal(int(value)) / Decimal(100))


def duplicate_row_basis(frame: pd.DataFrame) -> dict[int, str]:
    """Map each duplicate row's index to the reference basis that flagged it.

    A row counts as an exact duplicate only when its signed-cent amount
    matches another row's amount AND at least one populated reference
    field also matches exactly:

      * Both PO and invoice populated -> require all three (PO, invoice,
        amount) to match another row. The strictest and most common case.
      * Invoice blank, PO populated -> fall back to PO + amount, but only
        among rows that are *consistently* missing an invoice, so a row
        that does carry an invoice never collides with one that doesn't.
      * PO blank, invoice populated -> the mirror image: invoice + amount,
        restricted to rows consistently missing a PO.
      * Both PO and invoice blank -> never flagged. A shared amount alone
        is too weak a signal to safely treat as a duplicate.

    Two rows sharing only a PO (or only an invoice) with a *different*
    amount are never flagged under any basis -- see matching.py's grouped
    aggregate passes, which handle that scenario deliberately instead of
    assuming duplication.
    """
    basis: dict[int, str] = {}
    has_po = frame[NORM_PO].notna() & frame[NORM_PO].ne("")
    has_inv = frame[NORM_INV].notna() & frame[NORM_INV].ne("")
    has_amount = frame[AMOUNT_CENTS].notna()

    strict_subset = frame.loc[has_po & has_inv & has_amount, [NORM_PO, NORM_INV, AMOUNT_CENTS]]
    strict_dupe = strict_subset.duplicated(subset=[NORM_PO, NORM_INV, AMOUNT_CENTS], keep=False)
    for idx in strict_dupe.index[strict_dupe]:
        basis[int(idx)] = DUPLICATE_BASIS_STRICT

    po_only_subset = frame.loc[has_po & ~has_inv & has_amount, [NORM_PO, AMOUNT_CENTS]]
    po_only_dupe = po_only_subset.duplicated(subset=[NORM_PO, AMOUNT_CENTS], keep=False)
    for idx in po_only_dupe.index[po_only_dupe]:
        basis[int(idx)] = DUPLICATE_BASIS_PO_ONLY

    inv_only_subset = frame.loc[~has_po & has_inv & has_amount, [NORM_INV, AMOUNT_CENTS]]
    inv_only_dupe = inv_only_subset.duplicated(subset=[NORM_INV, AMOUNT_CENTS], keep=False)
    for idx in inv_only_dupe.index[inv_only_dupe]:
        basis[int(idx)] = DUPLICATE_BASIS_INVOICE_ONLY

    return basis


def duplicate_row_indexes(frame: pd.DataFrame) -> set[int]:
    """Return source indexes that are exact duplicate entries.

    See ``duplicate_row_basis`` for exactly what qualifies.
    """
    return set(duplicate_row_basis(frame))


def _group_key(frame: pd.DataFrame, idx: int, basis: str) -> tuple[str, ...]:
    amount_key = str(int(frame.at[idx, AMOUNT_CENTS]))
    if basis == DUPLICATE_BASIS_PO_ONLY:
        return (basis, frame.at[idx, NORM_PO], amount_key)
    if basis == DUPLICATE_BASIS_INVOICE_ONLY:
        return (basis, frame.at[idx, NORM_INV], amount_key)
    return (basis, frame.at[idx, NORM_PO], frame.at[idx, NORM_INV], amount_key)


def build_duplicate_item_report(
    frame: pd.DataFrame,
    duplicate_rows: set[int],
    id_column: str,
    dataset_label: str,
    source_scope: str,
    treatment: str,
) -> pd.DataFrame:
    """Itemize duplicate rows for one dataset/scope's own worksheet.

    Every duplicate item is listed individually rather than folded into a
    single summary line, so each excluded item stays independently
    traceable back to its source row. ``dataset_label`` keeps QuickBooks and
    Infinium reports independent of one another; ``source_scope`` (e.g.
    "Primary" vs. "Historical (Secondary)") keeps a duplicated historical
    row visibly distinct from a duplicated current-period row.
    """
    if not duplicate_rows:
        return pd.DataFrame(columns=DUPLICATE_ANALYSIS_COLUMNS)
    basis_by_row = duplicate_row_basis(frame)
    ordered = sorted(duplicate_rows, key=lambda idx: frame.at[idx, SOURCE_POS])

    groups: dict[tuple[str, ...], list[int]] = defaultdict(list)
    for idx in ordered:
        groups[_group_key(frame, idx, basis_by_row[idx])].append(idx)

    records: list[dict[str, Any]] = []
    for idx in ordered:
        basis = basis_by_row[idx]
        group = groups[_group_key(frame, idx, basis)]
        others = [str(frame.at[gidx, id_column]) for gidx in group if gidx != idx]
        records.append(
            {
                "Dataset": dataset_label,
                "Source Scope": source_scope,
                "Source Row ID": frame.at[idx, id_column],
                "Duplicate Basis": basis,
                "Normalized PO": frame.at[idx, NORM_PO],
                "Normalized Invoice": frame.at[idx, NORM_INV],
                "Amount": cents_to_float(frame.at[idx, AMOUNT_CENTS]),
                "Duplicate Group Size": len(group),
                "Other Source Row IDs In Group": "; ".join(others),
                "Treatment": treatment,
            }
        )
    return pd.DataFrame(records, columns=DUPLICATE_ANALYSIS_COLUMNS)


def combine_duplicate_reports(*reports: pd.DataFrame) -> pd.DataFrame:
    """Concatenate per-scope duplicate reports (e.g. primary + historical).

    Used so one dataset's worksheet (QuickBooks, or Infinium) can show its
    primary-file and historical-file duplicates together, distinguished by
    the "Source Scope" column, while the two datasets stay on separate
    reports per Rule 3.
    """
    populated = [report for report in reports if report is not None and not report.empty]
    if not populated:
        return pd.DataFrame(columns=DUPLICATE_ANALYSIS_COLUMNS)
    return pd.concat(populated, ignore_index=True)


@dataclass
class DuplicateScreeningResult:
    """One dataset's duplicate screening outcome.

    ``active_frame`` is the input frame with every duplicate row dropped.
    Callers must use this (never the original frame) for any matching or
    historical-clearance pass, so a duplicate -- current-period or
    historical -- can never be matched, never clear an exception, and never
    be folded into an accrual or journal-entry total.
    """

    duplicate_rows: list[int]
    active_frame: pd.DataFrame
    report: pd.DataFrame


def screen_duplicates(
    frame: pd.DataFrame,
    id_column: str,
    dataset_label: str,
    source_scope: str,
    treatment: str,
) -> DuplicateScreeningResult:
    """Identify duplicates, remove them from the working population, and itemize them.

    Apply this to every dataset entering the reconciliation -- primary
    QuickBooks and Infinium files as well as any optional historical
    (secondary) files -- before it is used for matching or clearance.
    """
    duplicate_rows = sorted(duplicate_row_indexes(frame))
    active_frame = frame.drop(index=duplicate_rows)
    report = build_duplicate_item_report(
        frame, set(duplicate_rows), id_column, dataset_label, source_scope, treatment
    )
    return DuplicateScreeningResult(
        duplicate_rows=duplicate_rows, active_frame=active_frame, report=report
    )
