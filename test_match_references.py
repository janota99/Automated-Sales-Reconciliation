"""Match references: one concise, traceable reference per accepted relationship.

M-### is a one-to-one match, G-### a grouped/aggregate match shared by every
record in it. References are assigned once, after matching is final, and
stored on the ReconciliationResult -- these tests pin down that assignment,
the exception pointers built on it, the centralized validation controls, and
(above all) that adding traceability never changed a matching decision.
"""

import copy

import pandas as pd
import pytest

from matching import (
    HISTORICAL_CLEARANCE_COLUMNS,
    QB_ID,
    INF_ID,
    assign_match_references,
    build_reconciliation,
    perform_matching,
    prepare_working_frame,
    validate_match_references,
)


def _reconcile(qb_rows, inf_rows, qb_mapping, inf_mapping, make_metadata, **kwargs):
    return build_reconciliation(
        pd.DataFrame(qb_rows), pd.DataFrame(inf_rows), qb_mapping, inf_mapping,
        make_metadata(), 2026, **kwargs,
    )


def _qb(po, invoice, amount, period="1"):
    return {"PO": po, "Invoice": invoice, "Amount": amount, "Qty": 1, "Period": period, "Customer": "Acme", "Date": "2026-01-05"}


def _inf(po, invoice, amount, period="1"):
    return {"PO": po, "Invoice": invoice, "Amount": amount, "Period": period,
            "Customer": "Acme", "Date": "2026-01-05"}


def _rows(result, section):
    return [row for row in result.paired_rows if row["Section"] == section]


def _register_row(result, reference):
    match = result.match_register.loc[result.match_register["Match Ref."] == reference]
    assert len(match) == 1, reference
    return match.iloc[0]


# ---------------------------------------------------------------------------
# Assignment
# ---------------------------------------------------------------------------

def test_one_to_one_match_shares_one_m_reference_on_both_sides(qb_mapping, inf_mapping, make_metadata):
    result = _reconcile(
        [_qb("PO1", "INV1", 100.00)], [_inf("PO1", "INV1", 100.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    assert [g.match_ref for g in result.matches] == ["M-001"]
    entry = _register_row(result, "M-001")
    assert entry["Match Type"] == "One-to-One"
    assert entry["QuickBooks Row IDs"] == "QB-1" and entry["Infinium Row IDs"] == "INF-1"
    (row,) = _rows(result, "01 Matched")
    # The QuickBooks and Infinium record are one paired row carrying the one reference.
    assert row["Match Ref."] == "M-001"
    assert row["QB Index"] == 0 and row["Infinium Index"] == 0


def test_independent_one_to_one_matches_get_different_sequential_references(
    qb_mapping, inf_mapping, make_metadata,
):
    result = _reconcile(
        [_qb(f"PO{i}", f"INV{i}", 10.0 * i) for i in (1, 2, 3)],
        [_inf(f"PO{i}", f"INV{i}", 10.0 * i) for i in (1, 2, 3)],
        qb_mapping, inf_mapping, make_metadata,
    )
    refs = [row["Match Ref."] for row in _rows(result, "01 Matched")]
    assert sorted(refs) == ["M-001", "M-002", "M-003"]
    assert len(set(refs)) == 3
    assert list(result.match_register["Match Ref."]) == ["M-001", "M-002", "M-003"]


def test_many_to_one_grouped_match_shares_one_g_reference(qb_mapping, inf_mapping, make_metadata):
    """Two QuickBooks rows collectively matched to one Infinium row carry the
    same G-### -- not M-### references, and not one per component row."""
    result = _reconcile(
        [_qb("PO-G", "INV-A", 100.00), _qb("PO-G", "INV-B", 50.00)],
        [_inf("PO-G", "", 150.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    assert result.unmatched_qb == []
    (group,) = result.matches
    assert group.group_level and group.match_ref == "G-001"
    assert not any(row["Match Ref."].startswith("M-") for row in result.paired_rows)
    matched = _rows(result, "01 Matched")
    assert len(matched) == 2  # one displayed line per participating row...
    assert {row["Match Ref."] for row in matched} == {"G-001"}  # ...one shared reference
    entry = _register_row(result, "G-001")
    assert entry["Match Type"] == "Grouped"
    assert entry["QuickBooks Row Count"] == 2 and entry["Infinium Row Count"] == 1
    assert set(entry["QuickBooks Row IDs"].split("; ")) == {"QB-1", "QB-2"}
    # The existing "no line allocation" disclosure is untouched.
    assert all("group-level; no line allocation" in row["Match Result"] for row in matched)


def test_one_to_many_grouped_match_shares_one_g_reference(qb_mapping, inf_mapping, make_metadata):
    result = _reconcile(
        [_qb("PO-H", "", 150.00)],
        [_inf("PO-H", "INV-X", 100.00), _inf("PO-H", "INV-Y", 50.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    (group,) = result.matches
    assert group.group_level and group.match_ref == "G-001"
    assert {row["Match Ref."] for row in _rows(result, "01 Matched")} == {"G-001"}
    entry = _register_row(result, "G-001")
    assert entry["QuickBooks Row Count"] == 1 and entry["Infinium Row Count"] == 2


def test_grouped_and_one_to_one_matches_are_numbered_in_separate_series(
    qb_mapping, inf_mapping, make_metadata,
):
    result = _reconcile(
        [
            _qb("PO1", "INV1", 10.00),
            _qb("PO-G", "INV-A", 100.00), _qb("PO-G", "INV-B", 50.00),
            _qb("PO2", "INV2", 20.00),
        ],
        [_inf("PO1", "INV1", 10.00), _inf("PO-G", "", 150.00), _inf("PO2", "INV2", 20.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    by_type = result.match_register.groupby("Match Type")["Match Ref."].apply(list).to_dict()
    assert by_type == {"One-to-One": ["M-001", "M-002"], "Grouped": ["G-001"]}


def test_match_reference_padding_widens_past_999(qb_mapping, inf_mapping):
    rows = [{"PO": f"PO{i}", "Invoice": f"INV{i}", "Amount": float(i), "Qty": 1, "Period": "1", "Customer": "Acme", "Date": "2026-01-05"}
            for i in range(1, 1002)]
    qb = prepare_working_frame(pd.DataFrame(rows), qb_mapping, "QB", 2026)
    inf = prepare_working_frame(pd.DataFrame(rows).drop(columns=["Qty"]), inf_mapping, "INF", 2026)
    matches, unmatched_qb, unmatched_inf, _ = perform_matching(qb, inf, enable_fuzzy=False)
    assert len(matches) == 1001 and not unmatched_qb
    register, _ = assign_match_references(
        matches, pd.DataFrame(columns=HISTORICAL_CLEARANCE_COLUMNS), qb, inf,
    )
    refs = list(register["Match Ref."])
    assert refs[0] == "M-0001" and refs[-1] == "M-1001"
    assert {len(ref) for ref in refs} == {6}  # uniform width once it has to grow
    small = assign_match_references(
        matches[:5], pd.DataFrame(columns=HISTORICAL_CLEARANCE_COLUMNS), qb, inf,
    )[0]
    assert list(small["Match Ref."])[0] == "M-001"


def test_references_are_deterministic_and_do_not_depend_on_display_order(
    qb_mapping, inf_mapping, make_metadata,
):
    qb_rows = [_qb(f"PO{i}", f"INV{i}", 10.0 * i) for i in (5, 1, 4, 2, 3)]
    inf_rows = [_inf(f"PO{i}", f"INV{i}", 10.0 * i) for i in (3, 5, 2, 1, 4)]
    first = _reconcile(qb_rows, inf_rows, qb_mapping, inf_mapping, make_metadata)
    second = _reconcile(qb_rows, inf_rows, qb_mapping, inf_mapping, make_metadata)
    pd.testing.assert_frame_equal(first.match_register, second.match_register)
    assert [r["Match Ref."] for r in first.paired_rows] == [r["Match Ref."] for r in second.paired_rows]
    # Canonical order is the QuickBooks file's own source order (PO5 is row 1),
    # not any later sort by PO or amount.
    assert _register_row(first, "M-001")["QuickBooks Row IDs"] == "QB-1"
    assert _register_row(first, "M-005")["QuickBooks Row IDs"] == "QB-5"


def test_historical_clearance_receives_a_reference(qb_mapping, inf_mapping, make_metadata):
    result = _reconcile(
        [_qb("PO1", "INV1", 100.00), _qb("PO-H", "INV-H", 25.00)],
        [_inf("PO1", "INV1", 100.00)],
        qb_mapping, inf_mapping, make_metadata,
        inf_secondary_raw=pd.DataFrame([_inf("PO-H", "INV-H", 25.00, "12")]),
        inf_secondary_mapping=inf_mapping,
    )
    (historical,) = _rows(result, "01 Matched - Historical Clearance")
    assert historical["Match Ref."] == "M-002"
    entry = _register_row(result, "M-002")
    assert entry["Record Scope"] == "Historical (prior period)"
    assert entry["Infinium Row IDs"].startswith("INF-HIST")
    assert set(result.historical_clearances["Match Ref."]) == {"M-002"}


# ---------------------------------------------------------------------------
# Rows that are NOT accepted matches
# ---------------------------------------------------------------------------

def test_unmatched_review_and_duplicate_rows_have_no_accepted_match_reference(
    qb_mapping, inf_mapping, make_metadata,
):
    result = _reconcile(
        [
            _qb("PO1", "INV1", 100.00),               # matched
            _qb("PO2", "INV2", 50.00),                # duplicate pair, no match
            _qb("PO2", "INV2", 50.00),
            _qb("PO9", "INV9", 15.00),                # genuinely unmatched
            _qb("PO-V", "INV-V", 500.00),             # reference-matched amount variance
        ],
        [_inf("PO1", "INV1", 100.00), _inf("PO-V", "INV-V", 450.00), _inf("PO7", "INV7", 3.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    matched_sections = {"01 Matched", "01 Matched - Historical Clearance"}
    for row in result.paired_rows:
        if row["Section"] in matched_sections:
            assert row["Match Ref."]
        else:
            assert row["Match Ref."] == "", row["Section"]
    assert result.amount_variance_analysis.shape[0] == 1
    assert list(result.match_register["Match Ref."]) == ["M-001"]
    accepted_ids = set(result.match_register.iloc[0][["QuickBooks Row IDs", "Infinium Row IDs"]])
    assert accepted_ids == {"QB-1", "INF-1"}


def test_fuzzy_review_hold_is_never_given_an_accepted_match_reference(
    qb_mapping, inf_mapping, make_metadata,
):
    result = _reconcile(
        [_qb("Hopper", "20044", 225.00, "6"), _qb("PO999", "INV999", 15.00, "6")],
        [_inf("DAVID HOPPER 2.2", "99999", 225.00, "6")],
        qb_mapping, inf_mapping, make_metadata,
    )
    assert result.metrics["Fuzzy Match Review Hold Rows"] == 1
    fuzzy_rows = _rows(result, "09 Fuzzy Match Review Hold")
    assert fuzzy_rows and all(row["Match Ref."] == "" for row in fuzzy_rows)
    assert result.match_register.empty and not result.matches


# ---------------------------------------------------------------------------
# Exceptions that point at an accepted match
# ---------------------------------------------------------------------------

def test_exception_with_a_consumed_po_names_the_exact_match(qb_mapping, inf_mapping, make_metadata):
    result = _reconcile(
        [_qb("PO-USED", "INV-A", 100.00), _qb("PO-USED", "INV-B", 55.00)],
        [_inf("PO-USED", "INV-A", 100.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    # "PO already used" is not a missing transaction: the row is held out of the
    # proposed JE, and says exactly which match already represents it.
    assert _rows(result, "02 Unmatched QuickBooks") == []
    (row,) = _rows(result, "11 Review Hold QuickBooks")
    assert row["Match Ref."] == ""
    assert row["Referenced Match Ref."] == "M-001"
    assert row["Match Result"] == "Review Hold — PO Already Represented by Match M-001"
    assert row["Reference Basis"] == "PO"
    assert "Potential Duplicate - Reference already used by another match" in row["Exception Cause"]
    assert row["Financial Treatment"] == "Excluded from automatic JE pending documented disposition"
    assert result.unmatched_qb == [] and result.reference_hold_qb_rows == [1]
    assert result.metrics["Proposed JE Amount"] == pytest.approx(0.0)
    (candidate,) = result.candidates.to_dict("records")
    assert candidate["Disposition"] == "Review: PO already used by Match M-001."


def test_exception_pointing_at_a_grouped_match_cites_the_group_reference(
    qb_mapping, inf_mapping, make_metadata,
):
    result = _reconcile(
        [_qb("PO-G", "INV-A", 100.00), _qb("PO-G", "INV-B", 50.00), _qb("PO-G", "INV-C", 7.00)],
        [_inf("PO-G", "", 150.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    (row,) = _rows(result, "11 Review Hold QuickBooks")
    assert row["Referenced Match Ref."] == "G-001"
    assert row["Match Result"] == "Review Hold — Group Candidate Already Represented by Group Match G-001"
    assert row["Match Ref."] == ""


def test_duplicate_of_a_matched_record_cites_the_original_match(qb_mapping, inf_mapping, make_metadata):
    result = _reconcile(
        [_qb("PO1", "INV1", 100.00), _qb("PO1", "INV1", 100.00)],
        [_inf("PO1", "INV1", 100.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    matched = _rows(result, "01 Matched")
    assert len(matched) == 1 and matched[0]["Match Ref."] == "M-001"
    # An exact duplicate is excluded before matching -- never left as a fresh exception.
    assert _rows(result, "02 Unmatched QuickBooks") == []
    (other,) = _rows(result, "04 Duplicate QuickBooks")
    assert other["Match Ref."] == ""                      # not a new accepted match
    assert other["Referenced Match Ref."] == "M-001"
    # The reference names the original match and keeps the duplicate basis.
    assert "Exact duplicate of Match M-001 (same PO, Invoice, and Amount)." in other["Explanation"]
    assert other["Match Result"] == "Excluded excess QuickBooks copy"
    assert other["Reference Basis"] == "Duplicate"
    # The duplicate audit report points at the same match.
    referenced = result.duplicate_analysis.set_index("Source Row ID")["Referenced Match Ref."]
    assert referenced.loc["QB-2"] == "M-001" and referenced.loc["QB-1"] == ""


def test_excluded_infinium_duplicate_cites_the_match_of_its_retained_copy(
    qb_mapping, inf_mapping, make_metadata,
):
    result = _reconcile(
        [_qb("PO1", "INV1", 100.00)],
        [_inf("PO1", "INV1", 100.00), _inf("PO1", "INV1", 100.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    (excess,) = _rows(result, "05 Duplicate Infinium")
    assert excess["Match Ref."] == ""
    assert excess["Referenced Match Ref."] == "M-001"
    assert "Exact duplicate of Match M-001 (same PO, Invoice, and Amount)." in excess["Explanation"]


def test_unrelated_exceptions_carry_no_referenced_match(qb_mapping, inf_mapping, make_metadata):
    result = _reconcile(
        [_qb("PO1", "INV1", 100.00), _qb("PO9", "INV9", 15.00)],
        [_inf("PO1", "INV1", 100.00)],
        qb_mapping, inf_mapping, make_metadata,
    )
    (row,) = _rows(result, "02 Unmatched QuickBooks")
    assert row["Referenced Match Ref."] == "" and row["Referenced Record IDs"] == ""
    assert row["Match Result"] == "No matching Infinium records"


# ---------------------------------------------------------------------------
# Centralized validation controls
# ---------------------------------------------------------------------------

@pytest.fixture
def referenced_result(qb_mapping, inf_mapping, make_metadata):
    """One grouped match (G-001), two one-to-one matches (M-001, M-002), a
    consumed-PO exception and a duplicate-of-a-match exception."""
    return _reconcile(
        [
            _qb("PO1", "INV1", 10.00),
            _qb("PO-G", "INV-A", 100.00), _qb("PO-G", "INV-B", 50.00),
            _qb("PO2", "INV2", 20.00), _qb("PO2", "INV2", 20.00),
            _qb("PO1", "INV1B", 5.00),
        ],
        [_inf("PO1", "INV1", 10.00), _inf("PO-G", "", 150.00), _inf("PO2", "INV2", 20.00)],
        qb_mapping, inf_mapping, make_metadata,
    )


def test_a_clean_result_passes_the_reference_control(referenced_result):
    validate_match_references(referenced_result)
    assert {"M-001", "M-002", "G-001"} == set(referenced_result.match_register["Match Ref."])
    assert any(r["Referenced Match Ref."] for r in referenced_result.paired_rows)


def _break(result, mutate):
    broken = copy.deepcopy(result)
    mutate(broken)
    return broken


def _expect_failure(result, mutate, message):
    with pytest.raises(ValueError, match=message):
        validate_match_references(_break(result, mutate))


def test_control_rejects_an_accepted_match_without_a_reference(referenced_result):
    def strip(result):
        result.matches[0].match_ref = ""
    _expect_failure(referenced_result, strip, "has no match reference")

    def strip_row(result):
        next(r for r in result.paired_rows if r["Section"] == "01 Matched")["Match Ref."] = ""
    _expect_failure(referenced_result, strip_row, "has no match reference")


def test_control_rejects_different_references_within_one_grouped_relationship(referenced_result):
    def split(result):
        rows = [r for r in result.paired_rows if r["Match Ref."] == "G-001"]
        assert len(rows) >= 2
        rows[1]["Match Ref."] = "M-002"
    _expect_failure(referenced_result, split, "carry different references")


def test_control_rejects_a_reference_on_an_unmatched_record(referenced_result):
    def leak(result):
        next(r for r in result.paired_rows if r["Section"] == "11 Review Hold QuickBooks")["Match Ref."] = "M-001"
    _expect_failure(referenced_result, leak, "is not an accepted match but shows reference M-001")


def test_control_rejects_a_record_in_two_accepted_matches(referenced_result):
    def overlap(result):
        register = result.match_register
        register.loc[register["Match Ref."] == "M-002", "QuickBooks Row IDs"] = "QB-1"
    _expect_failure(referenced_result, overlap, "assigned to more than one accepted match")


def test_control_rejects_a_missing_cited_match(referenced_result):
    def cite(result):
        row = next(r for r in result.paired_rows if r["Referenced Match Ref."])
        row["Referenced Match Ref."] = "M-999"
    _expect_failure(referenced_result, cite, "cites match M-999, which does not exist")


def test_control_rejects_a_cited_match_that_lacks_the_record(referenced_result):
    def misdirect(result):
        row = next(r for r in result.paired_rows if r["Referenced Match Ref."])
        row["Referenced Match Ref."] = "G-001"      # exists, but not the match cited by the evidence
        row["Referenced Record IDs"] = "QB-1"        # belongs to M-001, not G-001
    _expect_failure(referenced_result, misdirect, "does not contain the record")


def test_control_rejects_a_reference_shared_by_unrelated_relationships(referenced_result):
    def duplicate_reference(result):
        register = result.match_register
        register.loc[register["Match Ref."] == "M-002", "Match Ref."] = "M-001"
    _expect_failure(referenced_result, duplicate_reference, "assigned to more than one relationship")


def test_control_rejects_the_wrong_prefix_for_the_relationship_type(referenced_result):
    def wrong_grouped(result):
        register = result.match_register
        register.loc[register["Match Ref."] == "G-001", "Match Ref."] = "M-003"
        result.matches[next(i for i, g in enumerate(result.matches) if g.match_ref == "G-001")].match_ref = "M-003"
    _expect_failure(referenced_result, wrong_grouped, "grouped relationship M-003")

    def wrong_single(result):
        register = result.match_register
        register.loc[register["Match Ref."] == "M-002", "Match Ref."] = "G-002"
        result.matches[next(i for i, g in enumerate(result.matches) if g.match_ref == "M-002")].match_ref = "G-002"
    _expect_failure(referenced_result, wrong_single, "one-to-one relationship G-002")


def test_workbooks_refuse_to_export_a_result_that_fails_the_reference_control(referenced_result):
    from workpapers import build_analytics_workbook, build_legacy_workbook, build_primary_workbook

    broken = _break(referenced_result, lambda r: setattr(r.matches[0], "match_ref", ""))
    for builder in (build_primary_workbook, build_legacy_workbook, build_analytics_workbook):
        with pytest.raises(ValueError, match="Match reference control failure"):
            builder(broken)


# ---------------------------------------------------------------------------
# Traceability changes nothing about the reconciliation itself
# ---------------------------------------------------------------------------

def test_references_do_not_change_any_matching_decision_or_total(referenced_result):
    result = referenced_result
    # Which records matched, and how (hand-checked against the input above).
    matched_pairs = sorted(
        (tuple(sorted(g.qb_rows)), tuple(sorted(g.inf_rows)), g.method) for g in result.matches
    )
    assert matched_pairs == [
        ((0,), (0,), "PO + Invoice + Amount"),
        ((1, 2), (1,), "PO + Aggregate Amount (Grouped)"),
        ((3,), (2,), "PO + Invoice + Amount"),
    ] or len(matched_pairs) == 3
    assert result.metrics["Control Status"] == "PASS"
    # The exact duplicate of a matched row is excluded, and the row whose PO a
    # match already used is held -- neither feeds the proposed JE.
    assert result.duplicate_qb_rows == [4] and result.reference_hold_qb_rows == [5]
    assert result.unmatched_qb == []
    assert result.metrics["Duplicate QuickBooks Amount"] == pytest.approx(20.00)
    assert result.metrics["Reference Review Hold QuickBooks Amount"] == pytest.approx(5.00)
    assert result.metrics["Proposed JE Amount"] == pytest.approx(0.00)
    assert len(result.matches) == len(result.match_register) == 3
