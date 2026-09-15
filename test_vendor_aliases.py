"""Unit tests for vendor_aliases.py -- confirmed, persistent identity
equivalences that bridge PO references sharing no text similarity at all
(e.g. QuickBooks "Hopper" vs. Infinium "David")."""

import json

import pandas as pd
import pytest

from duplicates import AMOUNT_CENTS, SOURCE_POS
from fuzzy_po_matching import PO_TOKENS, significant_po_tokens
from vendor_aliases import (
    VendorAlias,
    add_vendor_alias,
    alias_explanation,
    build_alias_token_map,
    find_alias_po_matches,
    load_vendor_aliases,
    save_vendor_aliases,
)

ID_COL = "ID"


def make_frame(*rows):
    """rows: (id, raw_po, amount_cents_or_None) tuples."""
    records = [
        {ID_COL: row_id, PO_TOKENS: significant_po_tokens(po), AMOUNT_CENTS: amount}
        for row_id, po, amount in rows
    ]
    frame = pd.DataFrame(records)
    frame[SOURCE_POS] = range(len(frame))
    return frame


# ---------------------------------------------------------------------------
# Store persistence
# ---------------------------------------------------------------------------

def test_load_vendor_aliases_missing_file_returns_empty_list(tmp_path):
    assert load_vendor_aliases(tmp_path / "does_not_exist.json") == []


def test_add_vendor_alias_persists_and_round_trips(tmp_path):
    path = tmp_path / "aliases.json"
    entry = add_vendor_alias(
        ["Hopper", "David"], "David Hopper (dock sales)", "J. Reviewer",
        "Confirmed via golden-master validation.", path=path,
        confirmed_date="2026-09-15",
    )
    assert entry.id == "ALIAS-0001"
    assert entry.terms == ["DAVID", "HOPPER"]

    reloaded = load_vendor_aliases(path)
    assert len(reloaded) == 1
    assert reloaded[0] == entry


def test_add_vendor_alias_assigns_sequential_ids(tmp_path):
    path = tmp_path / "aliases.json"
    add_vendor_alias(["Hopper", "David"], "A", "Reviewer", "r", path=path)
    second = add_vendor_alias(["Eliot", "Electric"], "B", "Reviewer", "r", path=path)
    assert second.id == "ALIAS-0002"
    assert len(load_vendor_aliases(path)) == 2


def test_add_vendor_alias_requires_at_least_two_distinct_terms(tmp_path):
    path = tmp_path / "aliases.json"
    with pytest.raises(ValueError):
        add_vendor_alias(["Hopper"], "A", "Reviewer", "r", path=path)
    with pytest.raises(ValueError):
        add_vendor_alias(["Hopper", "hopper"], "A", "Reviewer", "r", path=path)


def test_save_and_load_preserve_inactive_flag(tmp_path):
    path = tmp_path / "aliases.json"
    alias = VendorAlias(
        id="ALIAS-0001", terms=["A", "B"], label="x", confirmed_by="y",
        confirmed_date="2026-01-01", rationale="z", active=False,
    )
    save_vendor_aliases([alias], path)
    reloaded = load_vendor_aliases(path)
    assert reloaded[0].active is False


# ---------------------------------------------------------------------------
# Token map / matching
# ---------------------------------------------------------------------------

def test_build_alias_token_map_ignores_inactive_aliases():
    active = VendorAlias("ALIAS-0001", ["HOPPER", "DAVID"], "x", "y", "2026-01-01", "z")
    inactive = VendorAlias("ALIAS-0002", ["ELIOT", "ELECTRIC"], "x", "y", "2026-01-01", "z", active=False)
    token_map = build_alias_token_map([active, inactive])
    assert token_map == {"HOPPER": "ALIAS-0001", "DAVID": "ALIAS-0001"}


def test_find_alias_po_matches_resolves_the_hopper_david_case():
    """The reported real-world gap: QuickBooks 'Hopper' and Infinium
    'David' name the same dock-sale customer by surname vs. first name --
    zero shared characters, unsolvable by any text-similarity rule."""
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(("INF1", "David", 67500))
    alias = VendorAlias("ALIAS-0001", ["HOPPER", "DAVID"], "David Hopper", "Reviewer", "2026-01-01", "confirmed")
    token_map = build_alias_token_map([alias])

    groups = find_alias_po_matches(qb, inf, {0}, {0}, token_map)
    assert groups == [((0,), (0,))]


def test_find_alias_po_matches_without_alias_finds_nothing():
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(("INF1", "David", 67500))
    groups = find_alias_po_matches(qb, inf, {0}, {0}, {})
    assert groups == []


def test_find_alias_po_matches_still_requires_amount_tie_out():
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(("INF1", "David", 30000))
    alias = VendorAlias("ALIAS-0001", ["HOPPER", "DAVID"], "David Hopper", "Reviewer", "2026-01-01", "confirmed")
    groups = find_alias_po_matches(qb, inf, {0}, {0}, build_alias_token_map([alias]))
    assert groups == []


def test_find_alias_po_matches_ignores_unrelated_rows():
    """A row with no alias-covered token must never be pulled in just
    because it shares a graph with an aliased pair."""
    qb = make_frame(("QB1", "Hopper", 67500), ("QB2", "Acme Corp", 5000))
    inf = make_frame(("INF1", "David", 67500), ("INF2", "Unrelated Vendor", 5000))
    alias = VendorAlias("ALIAS-0001", ["HOPPER", "DAVID"], "David Hopper", "Reviewer", "2026-01-01", "confirmed")
    groups = find_alias_po_matches(qb, inf, {0, 1}, {0, 1}, build_alias_token_map([alias]))
    assert groups == [((0,), (0,))]


def test_alias_explanation_names_the_confirming_record():
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(("INF1", "David", 67500))
    alias = VendorAlias(
        "ALIAS-0001", ["HOPPER", "DAVID"], "David Hopper (dock sales)",
        "J. Reviewer", "2026-09-15", "Confirmed via golden-master validation.",
    )
    token_map = build_alias_token_map([alias])
    explanation = alias_explanation(qb, inf, (0,), (0,), [alias], token_map)
    assert "David Hopper (dock sales)" in explanation
    assert "J. Reviewer" in explanation
    assert "2026-09-15" in explanation
