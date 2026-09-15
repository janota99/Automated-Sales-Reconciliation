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


def test_alias_term_scoped_to_the_specific_vendor_not_a_shared_first_name():
    """The reported real-world lesson: 'JERRY' alone was nearly included in
    the Jerry Wood alias, but the same real dataset has two other,
    unrelated dock-sale customers who also happen to be named Jerry
    ('Jerry Potter', 'Jerry - Gibbs Electric') at the same $225. A first
    name that generic must never be a standalone alias term -- only the
    distinctive compound/abbreviated forms ('WOOD', 'JWOOD', 'JERRYWOOD')
    are linked, so an unrelated same-first-name customer is never pulled
    into this alias group."""
    jerry_wood_alias = VendorAlias(
        "ALIAS-0003", ["WOOD", "JWOOD", "JERRYWOOD"], "Jerry Wood (dock sales)",
        "Reviewer", "2026-09-15", "confirmed",
    )
    token_map = build_alias_token_map([jerry_wood_alias])

    qb = make_frame(
        ("QB-POTTER", "Jerry Potter", 22500),
        ("QB-WOOD", "Jerry Wood 6.4.26", 22500),
    )
    inf = make_frame(("INF-JWOOD", "JWOOD7060", 22500))

    groups = find_alias_po_matches(qb, inf, {0, 1}, {0}, token_map)
    # Only the Wood row may match; Jerry Potter must never be pulled in
    # just because both QB rows happen to share the first name "Jerry".
    assert groups == [((1,), (0,))]


def test_incomplete_alias_terms_must_not_falsely_resolve_a_real_ambiguity():
    """The reported real-world lesson: Infinium had TWO $225 candidates for
    one QuickBooks 'Jerry Wood' row that period -- 'JWOOD7060' and
    'POJERRYWOOD' (a form the letter/digit tokenizer can't split, since it
    has no digit boundary). When the alias only covered 'JWOOD' and not
    the literal 'POJERRYWOOD' blob, it silently resolved to JWOOD7060 with
    'Confirmed' certainty -- a false confidence, since either candidate
    could plausibly be the real one and there was no way to tell which.
    Once both literal forms are in the alias, the pair correctly reverts
    to a genuine, unresolved ambiguity instead of an accidental pick."""
    qb = make_frame(("QB-WOOD", "Jerry Wood 6.4.26", 22500))
    inf = make_frame(
        ("INF-A", "JWOOD7060", 22500),
        ("INF-B", "POJERRYWOOD", 22500),
    )

    incomplete_alias = VendorAlias(
        "ALIAS-0003", ["WOOD", "JWOOD", "JERRYWOOD"], "Jerry Wood", "Reviewer", "2026-01-01", "confirmed",
    )
    groups = find_alias_po_matches(qb, inf, {0}, {0, 1}, build_alias_token_map([incomplete_alias]))
    assert groups == [((0,), (0,))]  # the bug: falsely "confirms" JWOOD7060 alone

    complete_alias = VendorAlias(
        "ALIAS-0003", ["WOOD", "JWOOD", "JERRYWOOD", "POJERRYWOOD"], "Jerry Wood",
        "Reviewer", "2026-01-01", "confirmed",
    )
    groups = find_alias_po_matches(qb, inf, {0}, {0, 1}, build_alias_token_map([complete_alias]))
    assert groups == []  # the fix: both candidates tie, correctly left unresolved


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
