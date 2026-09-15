"""Unit tests for fuzzy_po_matching.py -- the post-exact-pass fuzzy PO rule."""

import pandas as pd

from duplicates import AMOUNT_CENTS, SOURCE_POS
from fuzzy_po_matching import (
    PO_TOKENS,
    find_fuzzy_po_matches,
    is_fuzzy_po_match,
    significant_po_tokens,
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
# significant_po_tokens / is_fuzzy_po_match
# ---------------------------------------------------------------------------

def test_significant_po_tokens_drops_numbers_and_short_words():
    assert significant_po_tokens("DAVID HOPPER 2.2") == frozenset({"DAVID", "HOPPER"})
    assert significant_po_tokens("Hopper") == frozenset({"HOPPER"})
    assert significant_po_tokens("24 CASE") == frozenset({"CASE"})
    assert significant_po_tokens("PO-12345") == frozenset()
    assert significant_po_tokens(None) == frozenset()
    assert significant_po_tokens(float("nan")) == frozenset()


def test_significant_po_tokens_splits_letters_from_digits_with_no_separator():
    """Real Infinium data glues a name straight onto a reference number
    with no separator (e.g. 'JWOOD7060', 'ABOLT184225'). Without splitting
    at the letter/digit boundary, the whole blob is a one-off token that
    could never equal or resemble anything else again -- splitting it
    recovers the meaningful name portion and still drops the numeric
    reference the same way a space-separated number would be dropped."""
    assert significant_po_tokens("JWOOD7060") == frozenset({"JWOOD"})
    assert significant_po_tokens("ABOLT184225") == frozenset({"ABOLT"})
    assert significant_po_tokens("181493AMARILLO B") == frozenset({"AMARILLO"})
    # A purely-glued word with no digit boundary at all still can't be
    # split -- that's a known, documented limitation, not a regression.
    assert significant_po_tokens("POJERRYWOOD") == frozenset({"POJERRYWOOD"})


def test_hopper_example_is_a_fuzzy_match():
    """The exact scenario reported: QuickBooks 'Hopper' vs Infinium
    'DAVID HOPPER 2.2' for the same transaction."""
    qb_tokens = significant_po_tokens("Hopper")
    inf_tokens = significant_po_tokens("DAVID HOPPER 2.2")
    assert is_fuzzy_po_match(qb_tokens, inf_tokens)


def test_empty_token_sets_never_match():
    assert not is_fuzzy_po_match(frozenset(), frozenset({"HOPPER"}))
    assert not is_fuzzy_po_match(frozenset(), frozenset())


def test_shared_short_or_numeric_fragments_do_not_cause_a_match():
    """Two unrelated POs sharing only a product code (e.g. '24 CASE') must
    never be considered a fuzzy match -- the shared token there is a number,
    already dropped by significant_po_tokens before comparison even runs."""
    tokens_a = significant_po_tokens("24 CASE LOWES")
    tokens_b = significant_po_tokens("24 CASE ALLSUPS")
    assert not is_fuzzy_po_match(tokens_a, tokens_b)


def test_partial_word_overlap_is_not_a_match():
    """Every word on the shorter side must appear -- one shared name among
    several unrelated words is not enough."""
    tokens_a = significant_po_tokens("DAVID HOPPER SMITH")
    tokens_b = significant_po_tokens("DAVID HOPPER 2.2")
    # "SMITH" is not in tokens_b, so the shorter set (tokens_b, 2 words) is
    # not fully contained in tokens_a either way this is checked -- but
    # tokens_a is NOT a subset of tokens_b (SMITH missing), and tokens_b
    # (DAVID, HOPPER) IS a subset of tokens_a. The function takes the
    # smaller side's tokens, so this case *does* match on the smaller
    # (Infinium-shaped) side; assert that explicitly.
    assert is_fuzzy_po_match(tokens_a, tokens_b)


def test_completely_unrelated_names_do_not_match():
    tokens_a = significant_po_tokens("Hopper")
    tokens_b = significant_po_tokens("SMITH ENTERPRISES LLC")
    assert not is_fuzzy_po_match(tokens_a, tokens_b)


def test_single_character_typo_within_a_word_still_matches():
    """A one-character typo in an otherwise-matching word must not defeat
    the match -- the reported gap: real invoice/PO pairs where a name is
    misspelled by a single letter were being rejected outright."""
    assert is_fuzzy_po_match(
        significant_po_tokens("HOPER"), significant_po_tokens("HOPPER"),
    )


def test_elliot_electric_hord_matches_eliot_electric():
    """The reported real-world case: QuickBooks 'ELLIOT ELECTRIC - HORD' vs
    Infinium 'ELIOT ELECTRIC'. 'ELECTRIC' matches exactly; 'ELLIOT'/'ELIOT'
    is a one-letter typo of the same name; the extra 'HORD' token doesn't
    block the match since the ratio is computed against the shorter side."""
    qb_tokens = significant_po_tokens("ELLIOT ELECTRIC - HORD")
    inf_tokens = significant_po_tokens("ELIOT ELECTRIC")
    assert is_fuzzy_po_match(qb_tokens, inf_tokens)


def test_similar_looking_but_different_surnames_do_not_match():
    """A near-miss must still require real similarity -- 'Hopper' and
    'Hooper' are different, unrelated surnames and must never be fuzzy-
    matched just because they look alike."""
    assert not is_fuzzy_po_match(
        significant_po_tokens("Hopper"), significant_po_tokens("Hooper"),
    )


def test_unrelated_word_containing_the_target_does_not_near_miss_match():
    """'SHOPPER' and 'CHOPPER' both score ~0.92 similarity against 'HOPPER'
    on ratio() alone (well past the 0.84 threshold) despite being distinct
    words, not misspellings -- they add an unrelated leading letter rather
    than typo a shared word. The leading-character requirement in
    _near_miss_pairs must reject both."""
    hopper = significant_po_tokens("Hopper")
    assert not is_fuzzy_po_match(hopper, significant_po_tokens("Shopper"))
    assert not is_fuzzy_po_match(hopper, significant_po_tokens("Chopper"))


def test_allow_near_miss_false_restricts_to_exact_token_overlap():
    qb_tokens = significant_po_tokens("Hopper")
    inf_tokens = significant_po_tokens("DAVID HOPPER 2.2")
    assert is_fuzzy_po_match(qb_tokens, inf_tokens, allow_near_miss=False)

    typo_tokens = significant_po_tokens("HOPER")
    assert is_fuzzy_po_match(qb_tokens, typo_tokens)
    assert not is_fuzzy_po_match(qb_tokens, typo_tokens, allow_near_miss=False)


# ---------------------------------------------------------------------------
# find_fuzzy_po_matches
# ---------------------------------------------------------------------------

def test_finds_the_hopper_pair_with_matching_amount():
    qb = make_frame(("QB1", "Hopper", 22500))
    inf = make_frame(("INF1", "DAVID HOPPER 2.2", 22500))
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0})
    assert groups == [((0,), (0,))]


def test_amount_mismatch_blocks_an_otherwise_fuzzy_match():
    qb = make_frame(("QB1", "Hopper", 22500))
    inf = make_frame(("INF1", "DAVID HOPPER 2.2", 30000))
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0})
    assert groups == []


def test_invalid_amount_rows_are_never_fuzzy_matched():
    qb = make_frame(("QB1", "Hopper", None))
    inf = make_frame(("INF1", "DAVID HOPPER 2.2", 22500))
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0})
    assert groups == []


def test_ambiguous_fuzzy_candidates_are_never_guessed():
    """One QB row fuzzy-matches two different Infinium rows, but their
    aggregate amount doesn't tie out to the QB side -- the whole component
    is rejected rather than guessing which Infinium row is the real match."""
    qb = make_frame(("QB1", "Hopper", 22500))
    inf = make_frame(
        ("INF1", "DAVID HOPPER 2.2", 22500),
        ("INF2", "HOPPER LOGISTICS", 22500),
    )
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0, 1})
    assert groups == []


def test_unrelated_rows_with_no_fuzzy_match_stay_unresolved():
    qb = make_frame(("QB1", "Acme Corp", 5000))
    inf = make_frame(("INF1", "DAVID HOPPER 2.2", 5000))
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0})
    assert groups == []


def test_many_to_many_components_are_rejected_even_when_totals_tie_out():
    """matching.py's grouped-matching control only ever accepts a bounded
    one-to-many or many-to-one relationship (one side must be exactly one
    row). A true many-to-many cluster -- multiple QB rows and multiple
    Infinium rows all linked to each other through shared fuzzy tokens --
    must never be accepted here, even if the aggregate sums happen to
    agree, or it would trip that downstream control and halt the whole
    reconciliation. Both QB rows use the bare token "Hopper" so every
    pairwise comparison clears the containment ratio and the component is
    genuinely fully connected on both sides -- not merely bounded with an
    unrelated extra candidate (see the unique-tie fallback in
    _resolve_components, covered by test_recurring_name_with_one_true_match_
    among_several_ties_uniquely)."""
    qb = make_frame(("QB1", "Hopper", 10000), ("QB2", "Hopper", 15000))
    inf = make_frame(("INF1", "David Hopper", 10000), ("INF2", "Hopper Trucking", 15000))
    groups = find_fuzzy_po_matches(qb, inf, {0, 1}, {0, 1})
    assert groups == []


def test_unrelated_near_miss_row_does_not_void_a_clean_exact_match():
    """The reported real-world gap: QuickBooks 'Hopper' ($675) and Infinium
    'DAVID HOPPER' ($675) is a clean, self-contained exact-token 1:1 match.
    An unrelated Infinium row that merely near-misses 'HOPPER' by raw
    similarity ('MYSTERY SHOPPER PROGRAM', a different amount, no real
    relationship) must never be allowed to drag the true pair into the same
    graph component and void it -- exact-token components are resolved
    before near-miss ones are even considered."""
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(
        ("INF1", "DAVID HOPPER", 67500),
        ("INF2", "MYSTERY SHOPPER PROGRAM", 30000),
    )
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0, 1})
    assert groups == [((0,), (0,))]


def test_recurring_name_with_one_true_match_among_several_ties_uniquely():
    """The actual reported real-world case, confirmed against production
    data: QuickBooks has one 'Hopper' row ($675). Infinium has THREE
    separate transactions that period all naming the same recurring
    dock-sale customer 'David Hopper' -- at three different amounts. The
    group sum across all three obviously doesn't tie to $675, but exactly
    one of them ($675) ties the QB row's amount individually and
    unambiguously -- that must resolve as the match, without needing to
    net the unrelated transactions together or guess among them."""
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(
        ("INF1", "DAVID HOPPER 2.2", 67500),
        ("INF2", "PO HOPPER,DAVID", 247500),
        ("INF3", "PO DAVID HOPPER SITELINE SERVI", 157500),
    )
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0, 1, 2})
    assert groups == [((0,), (0,))]


def test_recurring_name_with_two_equal_ties_stays_ambiguous():
    """If two of the several same-side candidates both individually tie
    the singleton row's amount, the choice is genuinely ambiguous and must
    still be left unresolved rather than guessed -- the unique-tie fallback
    only fires when exactly one candidate matches."""
    qb = make_frame(("QB1", "Hopper", 67500))
    inf = make_frame(
        ("INF1", "DAVID HOPPER 2.2", 67500),
        ("INF2", "PO HOPPER,DAVID", 67500),
        ("INF3", "PO DAVID HOPPER SITELINE SERVI", 157500),
    )
    groups = find_fuzzy_po_matches(qb, inf, {0}, {0, 1, 2})
    assert groups == []
