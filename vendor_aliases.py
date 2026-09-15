"""Persistent, human-confirmed vendor/customer identity equivalences.

Fuzzy text matching -- exact-token or typo-level near-miss -- can never
bridge two PO references that share no characters at all, e.g. QuickBooks
"Hopper" vs. Infinium "David": the same dock-sale customer, referenced by
surname on one system and first name on the other. No string-similarity
rule can respond to that; it requires someone who knows David Hopper is
the same person on both sides.

When a reviewer confirms a pairing like that, it belongs here rather than
in the fuzzy-match review-hold report -- a hold is for a guess that still
needs a human to look at it every time it recurs, but this is a decided
fact. Once recorded, ``find_alias_po_matches`` treats the aliased terms as
interchangeable and resolves the pair like any other clean, bounded,
amount-tying match. Every alias entry keeps a durable audit trail --
who confirmed it, when, and why -- so the reconciliation never has to
re-litigate the same identity question period after period.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from fuzzy_po_matching import PO_TOKENS, _build_candidate_edges, _resolve_components
from duplicates import SOURCE_POS

import pandas as pd

__all__ = [
    "ALIAS_CONFIDENCE",
    "ALIAS_METHOD",
    "DEFAULT_ALIAS_STORE_PATH",
    "VendorAlias",
    "add_vendor_alias",
    "build_alias_token_map",
    "find_alias_po_matches",
    "load_vendor_aliases",
    "save_vendor_aliases",
]

DEFAULT_ALIAS_STORE_PATH = Path(__file__).resolve().parent / "vendor_aliases.json"

ALIAS_METHOD = "Confirmed Vendor Alias + Amount"
ALIAS_CONFIDENCE = "Confirmed"


@dataclass
class VendorAlias:
    """One confirmed identity equivalence, e.g. {"HOPPER", "DAVID"} both
    referring to the same real-world dock-sale customer."""

    id: str
    terms: list[str]
    label: str
    confirmed_by: str
    confirmed_date: str
    rationale: str
    active: bool = True

    def explanation(self) -> str:
        term_list = ", ".join(sorted(self.terms))
        return (
            f'Confirmed vendor alias "{self.label}" ({self.id}): terms '
            f"[{term_list}] treated as the same party. Confirmed by "
            f"{self.confirmed_by} on {self.confirmed_date} -- {self.rationale}"
        )


def load_vendor_aliases(path: Path | str = DEFAULT_ALIAS_STORE_PATH) -> list[VendorAlias]:
    """Load the alias store. Missing file means no confirmed aliases yet --
    that's a normal starting state, not an error."""
    path = Path(path)
    if not path.exists():
        return []
    raw = json.loads(path.read_text(encoding="utf-8"))
    return [VendorAlias(**entry) for entry in raw.get("aliases", [])]


def save_vendor_aliases(aliases: list[VendorAlias], path: Path | str = DEFAULT_ALIAS_STORE_PATH) -> None:
    path = Path(path)
    payload = {"aliases": [asdict(a) for a in aliases]}
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def add_vendor_alias(
    terms: list[str],
    label: str,
    confirmed_by: str,
    rationale: str,
    *,
    path: Path | str = DEFAULT_ALIAS_STORE_PATH,
    confirmed_date: Optional[str] = None,
) -> VendorAlias:
    """Confirm a new identity equivalence and persist it permanently.

    ``terms`` must include at least two distinct significant tokens (see
    ``significant_po_tokens`` in fuzzy_po_matching.py) -- one from each
    side is the normal case, e.g. terms=["HOPPER", "DAVID"].
    """
    normalized_terms = sorted({t.strip().upper() for t in terms if t.strip()})
    if len(normalized_terms) < 2:
        raise ValueError("A vendor alias needs at least two distinct terms to link.")

    aliases = load_vendor_aliases(path)
    existing_ids = {a.id for a in aliases}
    next_number = len(aliases) + 1
    next_id = f"ALIAS-{next_number:04d}"
    while next_id in existing_ids:
        next_number += 1
        next_id = f"ALIAS-{next_number:04d}"

    entry = VendorAlias(
        id=next_id,
        terms=normalized_terms,
        label=label,
        confirmed_by=confirmed_by,
        confirmed_date=confirmed_date or datetime.now(timezone.utc).date().isoformat(),
        rationale=rationale,
    )
    aliases.append(entry)
    save_vendor_aliases(aliases, path)
    return entry


def build_alias_token_map(aliases: list[VendorAlias]) -> dict[str, str]:
    """Map every confirmed term (upper-cased) to its alias group id, active
    aliases only. A row's PO tokens run through this map to find whether
    they belong to a confirmed identity group."""
    token_map: dict[str, str] = {}
    for alias in aliases:
        if not alias.active:
            continue
        for term in alias.terms:
            token_map[term.upper()] = alias.id
    return token_map


def _alias_groups(tokens: frozenset[str], alias_token_map: dict[str, str]) -> frozenset[str]:
    return frozenset(alias_token_map[t] for t in tokens if t in alias_token_map)


def _alias_match(
    tokens_a: frozenset[str], tokens_b: frozenset[str], alias_token_map: dict[str, str]
) -> bool:
    """True if the two token sets touch the same confirmed alias group.
    Deliberately simpler than is_fuzzy_po_match's ratio/length rules --
    a confirmed alias is a decided fact, not a text-similarity guess, so
    it doesn't need those false-positive guards."""
    if not alias_token_map:
        return False
    return bool(_alias_groups(tokens_a, alias_token_map) & _alias_groups(tokens_b, alias_token_map))


def find_alias_po_matches(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    remaining_q: set[int],
    remaining_i: set[int],
    alias_token_map: dict[str, str],
    qb_date_col: Optional[str] = None,
    inf_date_col: Optional[str] = None,
    max_days_variance: int = 30,
) -> list[tuple[tuple[int, ...], tuple[int, ...]]]:
    """Return grouped pairs where the two sides share a confirmed vendor
    alias, using the same bipartite-graph resolution (and the same
    aggregate-amount and bounded-shape safeguards) as the fuzzy PO pass --
    see fuzzy_po_matching._resolve_components. An alias match still has to
    tie out on amount; only the identity-equivalence check is replaced by
    the confirmed alias lookup instead of text similarity.
    """
    if not alias_token_map:
        return []

    q_rows = sorted(remaining_q, key=lambda idx: qb.at[idx, SOURCE_POS])
    i_rows = sorted(remaining_i, key=lambda idx: inf.at[idx, SOURCE_POS])

    edges = _build_candidate_edges(
        qb, inf, q_rows, i_rows, qb_date_col, inf_date_col, max_days_variance,
        match_fn=lambda a, b: _alias_match(a, b, alias_token_map),
    )
    return _resolve_components(qb, inf, edges)


def alias_explanation(
    qb: pd.DataFrame,
    inf: pd.DataFrame,
    q_rows: tuple[int, ...],
    i_rows: tuple[int, ...],
    aliases: list[VendorAlias],
    alias_token_map: dict[str, str],
) -> str:
    """Build a human-readable explanation naming which confirmed alias
    record justified this match, for the audit trail."""
    alias_by_id = {a.id: a for a in aliases}
    matched_ids: set[str] = set()
    for qidx in q_rows:
        matched_ids |= _alias_groups(qb.at[qidx, PO_TOKENS], alias_token_map)
    for iidx in i_rows:
        matched_ids |= _alias_groups(inf.at[iidx, PO_TOKENS], alias_token_map)
    explanations = [alias_by_id[aid].explanation() for aid in sorted(matched_ids) if aid in alias_by_id]
    return " ".join(explanations) if explanations else "Confirmed vendor alias match."
