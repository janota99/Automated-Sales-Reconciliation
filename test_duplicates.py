"""Regression tests for duplicates.py.

Run from the application directory with:
    python -m unittest -v test_duplicates.py
"""

from __future__ import annotations

import unittest

import pandas as pd

from duplicates import (
    AMOUNT_CENTS,
    NORM_INV,
    NORM_PO,
    LINE_ID,
    SOURCE_POS,
    TXN_ID,
    DISPOSITION_REVIEW,
    DISPOSITION_REVIEW_HOLD,
    DISPOSITION_REVIEW_RESOLVED,
    DuplicateScreeningError,
    cents_to_float,
    finalize_review_dispositions,
)
from duplicates import screen_duplicates as _screen_duplicates


def screen_duplicates(*args, **kwargs):
    """These fixtures fingerprint copies on "Raw Note" -- the one stable field
    they carry. Screening never looks at any other column."""
    kwargs.setdefault("fingerprint_columns", ("Raw Note",))
    return _screen_duplicates(*args, **kwargs)


def make_frame(rows, *, index=None, prefix="ROW", notes=None):
    return pd.DataFrame(
        {
            SOURCE_POS: range(len(rows)),
            NORM_PO: [row[0] for row in rows],
            NORM_INV: [row[1] for row in rows],
            AMOUNT_CENTS: [row[2] for row in rows],
            "Source Row ID": [f"{prefix}-{number + 1}" for number in range(len(rows))],
            "Raw Note": notes or ["same"] * len(rows),
        },
        index=index,
    )


class DuplicateScreeningTests(unittest.TestCase):
    def screen(self, frame, **kwargs):
        return screen_duplicates(
            frame, "Source Row ID", "QuickBooks", "Primary", **kwargs
        )

    def test_strong_group_retains_one_and_excludes_only_excess(self):
        result = self.screen(make_frame([("100", "A", 10000)] * 3))
        self.assertEqual(result.canonical_rows, [0])
        self.assertEqual(result.duplicate_rows, [1, 2])
        self.assertEqual(list(result.active_frame.index), [0])
        self.assertEqual(int(result.report["Automatically Excluded"].sum()), 2)

    def test_same_business_key_with_different_payload_is_review_only(self):
        frame = make_frame(
            [("100", "A", 10000), ("100", "A", 10000)],
            notes=["product one", "product two"],
        )
        result = self.screen(frame)
        self.assertEqual(result.duplicate_rows, [])
        self.assertEqual(result.suspected_rows, [0, 1])
        self.assertEqual(len(result.active_frame), 2)

    def test_po_only_and_invoice_only_groups_remain_active(self):
        for rows in (
            [("100", "", 10000), ("100", "", 10000)],
            [("", "A", 10000), ("", "A", 10000)],
        ):
            with self.subTest(rows=rows):
                result = self.screen(make_frame(rows))
                self.assertEqual(result.duplicate_rows, [])
                self.assertEqual(result.suspected_rows, [0, 1])

    def test_amount_only_is_not_a_duplicate(self):
        result = self.screen(make_frame([("", "", 10000)] * 2))
        self.assertTrue(result.report.empty)
        self.assertEqual(len(result.active_frame), 2)

    def test_positive_and_negative_amounts_do_not_collide(self):
        result = self.screen(
            make_frame([("100", "A", 10000), ("100", "A", -10000)])
        )
        self.assertTrue(result.report.empty)

    def test_arbitrary_unique_indexes_are_supported(self):
        frame = make_frame(
            [("100", "A", 10000)] * 2,
            index=["row-a", "row-b"],
        )
        result = self.screen(frame)
        self.assertEqual(result.canonical_rows, ["row-a"])
        self.assertEqual(result.duplicate_rows, ["row-b"])

    def test_repeated_indexes_raise_controlled_error(self):
        frame = make_frame(
            [("100", "A", 10000)] * 2,
            index=[0, 0],
        )
        with self.assertRaisesRegex(DuplicateScreeningError, "repeated DataFrame index"):
            self.screen(frame)

    def test_fractional_cents_raise_controlled_error(self):
        with self.assertRaisesRegex(DuplicateScreeningError, "whole signed cents"):
            self.screen(make_frame([("100", "A", 10000.5)]))

    def test_missing_amount_is_not_zero(self):
        result = self.screen(
            make_frame([("100", "A", None), ("100", "A", 0)])
        )
        self.assertTrue(result.report.empty)
        self.assertIsNone(cents_to_float(None))
        self.assertEqual(cents_to_float(0), 0.0)

    def test_primary_historical_exact_payload_overlap_is_excluded(self):
        primary = make_frame([("100", "A", 10000)], prefix="PRIMARY")
        historical = make_frame([("100", "A", 10000)], prefix="HIST")
        result = screen_duplicates(
            historical,
            "Source Row ID",
            "QuickBooks",
            "Historical (Secondary)",
            reference_frame=primary,
            reference_id_column="Source Row ID",
        )
        self.assertEqual(result.duplicate_rows, [0])
        self.assertEqual(result.report.iloc[0]["Reference Source Row IDs"], "PRIMARY-1")

    def test_cross_scope_key_collision_with_different_payload_is_review_only(self):
        primary = make_frame([("100", "A", 10000)], prefix="PRIMARY")
        historical = make_frame(
            [("100", "A", 10000)], prefix="HIST", notes=["changed"]
        )
        result = screen_duplicates(
            historical,
            "Source Row ID",
            "QuickBooks",
            "Historical (Secondary)",
            reference_frame=primary,
            reference_id_column="Source Row ID",
        )
        self.assertEqual(result.duplicate_rows, [])
        self.assertEqual(result.suspected_rows, [0])

    def test_only_the_explicit_stable_fields_form_the_fingerprint(self):
        """Report/export metadata -- row numbers, sequences, timestamps, generated
        IDs, or any column added later -- neither blocks nor creates a confirmation."""
        frame = make_frame([("100", "A", 10000)] * 2)
        frame["Export Row #"] = [7, 8]
        frame["Import Timestamp"] = ["2026-09-01 10:00", "2026-09-02 11:30"]
        frame["Generated ID"] = ["g-1", "g-2"]
        result = self.screen(frame)
        self.assertEqual(result.duplicate_rows, [1])
        self.assertEqual(result.report.iloc[0]["Confirmation Basis"], "Stable-field fingerprint (Raw Note)")

    def test_a_real_difference_in_a_stable_field_prevents_confirmation(self):
        result = self.screen(make_frame([("100", "A", 10000)] * 2, notes=["Acme", "Beta"]))
        self.assertEqual(result.duplicate_rows, [])
        self.assertEqual(result.suspected_rows, [0, 1])

    def test_formatting_only_differences_do_not_prevent_confirmation(self):
        result = self.screen(make_frame([("100", "A", 10000)] * 2, notes=["Acme  Corp", "ACME CORP"]))
        self.assertEqual(result.duplicate_rows, [1])
        dates = self.screen(make_frame([("100", "A", 10000)] * 2, notes=["1/5/2026", "2026-01-05"]))
        self.assertEqual(dates.duplicate_rows, [1])
        numbers = self.screen(make_frame([("100", "A", 10000)] * 2, notes=[1, 1.0]))
        self.assertEqual(numbers.duplicate_rows, [1])

    def test_without_enough_identity_evidence_nothing_is_confirmed(self):
        # Quantity-like context alone is not identity: with a requirement of two
        # identity-grade fields and none supplied, the rows stay potential duplicates.
        frame = make_frame([("100", "A", 10000)] * 2)
        result = self.screen(frame, fingerprint_columns=("Raw Note",), min_identity_fields=2)
        self.assertEqual(result.duplicate_rows, [])
        self.assertEqual(result.suspected_rows, [0, 1])
        self.assertIn("not sufficient", result.report.iloc[0]["Potential Duplicate Reason"])

    def test_two_identity_grade_fields_are_enough_and_weak_fields_are_not(self):
        frame = make_frame([("100", "A", 10000)] * 2)
        frame["Customer"] = ["Acme", "Acme"]
        frame["Date"] = ["2026-01-05", "2026-01-05"]
        frame["Qty"] = [1, 1]
        strong = self.screen(
            frame, fingerprint_columns=("Customer", "Date", "Qty"),
            identity_columns=("Customer", "Date"), min_identity_fields=2,
        )
        self.assertEqual(strong.duplicate_rows, [1])
        weak = self.screen(
            frame, fingerprint_columns=("Qty",), identity_columns=(), min_identity_fields=2,
        )
        self.assertEqual(weak.duplicate_rows, [])
        one_field = self.screen(
            frame, fingerprint_columns=("Customer", "Qty"),
            identity_columns=("Customer",), min_identity_fields=2,
        )
        self.assertEqual(one_field.duplicate_rows, [])

    def test_a_line_level_id_confirms_a_copy_on_its_own_when_nothing_contradicts_it(self):
        frame = make_frame([("100", "A", 10000)] * 2)
        frame[LINE_ID] = ["L-1", "L-1"]
        result = self.screen(frame, fingerprint_columns=(), min_identity_fields=2)
        self.assertEqual(result.duplicate_rows, [1])
        self.assertEqual(result.report.iloc[0]["Confirmation Basis"], "Line-level source ID")

    def test_an_id_shared_by_rows_that_differ_is_not_line_level_and_never_confirms(self):
        # Two different sales lines (different item) share one invoice-level ID.
        frame = make_frame([("100", "A", 10000)] * 2, notes=["Item X", "Item Y"])
        frame[LINE_ID] = ["INV-1", "INV-1"]
        result = self.screen(frame, fingerprint_columns=("Raw Note",))
        self.assertEqual(result.duplicate_rows, [])
        self.assertEqual(result.suspected_rows, [0, 1])

    def test_one_contradiction_demotes_the_whole_column(self):
        frame = make_frame(
            [("100", "A", 10000), ("100", "A", 10000), ("200", "B", 500), ("200", "B", 500)],
            notes=["same", "same", "Item X", "Item Y"],
        )
        frame[LINE_ID] = ["L-1", "L-1", "INV-9", "INV-9"]
        result = self.screen(frame, fingerprint_columns=("Raw Note",), min_identity_fields=2)
        # The second pair shares an ID across different lines, so the column is not
        # line-level; the first pair's ID therefore cannot confirm it either.
        self.assertEqual(result.duplicate_rows, [])

    def test_a_transaction_level_id_only_supports_a_fingerprint_match(self):
        base = make_frame([("100", "A", 10000)] * 2)
        base["Customer"] = ["Acme", "Acme"]
        base["Date"] = ["2026-01-05", "2026-01-05"]
        kwargs = dict(
            fingerprint_columns=("Customer", "Date"), identity_columns=("Customer", "Date"),
            min_identity_fields=2,
        )
        # Same transaction ID + identical line fingerprint => confirmed.
        same = base.copy()
        same[TXN_ID] = ["T-1", "T-1"]
        result = self.screen(same, **kwargs)
        self.assertEqual(result.duplicate_rows, [1])
        self.assertTrue(result.report.iloc[0]["Confirmation Basis"].endswith("+ transaction ID"))
        # Different transaction IDs => different transactions, however alike.
        different = base.copy()
        different[TXN_ID] = ["T-1", "T-2"]
        self.assertEqual(self.screen(different, **kwargs).duplicate_rows, [])
        # The same transaction ID with NO line-level evidence is not enough.
        alone = make_frame([("100", "A", 10000)] * 2)
        alone[TXN_ID] = ["T-1", "T-1"]
        self.assertEqual(
            self.screen(alone, fingerprint_columns=(), min_identity_fields=2).duplicate_rows, [],
        )
        # ...and a shared transaction ID never overrides a line-level difference.
        lines = base.copy()
        lines[TXN_ID] = ["T-1", "T-1"]
        lines["Customer"] = ["Acme", "Beta"]
        self.assertEqual(self.screen(lines, **kwargs).duplicate_rows, [])

    def test_screening_does_not_mutate_input(self):
        frame = make_frame([("100", "A", 10000)] * 2)
        before = frame.copy(deep=True)
        self.screen(frame)
        pd.testing.assert_frame_equal(frame, before)

    def test_group_id_is_deterministic(self):
        frame = make_frame([("100", "A", 10000)] * 2)
        first = self.screen(frame).report["Duplicate Group ID"].tolist()
        second = self.screen(frame).report["Duplicate Group ID"].tolist()
        self.assertEqual(first, second)


class FinalizeReviewDispositionTests(unittest.TestCase):
    """Weak-basis (PO-only/invoice-only) candidates are screened before
    matching runs, so screening alone cannot know whether a candidate will
    end up matched or left unresolved. finalize_review_dispositions patches
    the report with that outcome once matching completes.
    """

    def review_report(self):
        # PO-only basis: both rows have blank invoice, populated PO -- never
        # auto-excluded by screening alone, both start at DISPOSITION_REVIEW.
        frame = make_frame([("100", "", 10000), ("100", "", 10000)])
        result = screen_duplicates(frame, "Source Row ID", "QuickBooks", "Primary")
        self.assertEqual(result.duplicate_rows, [])
        self.assertEqual(result.suspected_rows, [0, 1])
        self.assertTrue((result.report["Disposition"] == DISPOSITION_REVIEW).all())
        return result.report

    def test_matched_candidate_is_marked_resolved_and_not_excluded(self):
        report = self.review_report()
        updated = finalize_review_dispositions(report, resolved_ids=["ROW-1"])
        row = updated.loc[updated["Source Row ID"] == "ROW-1"].iloc[0]
        self.assertEqual(row["Disposition"], DISPOSITION_REVIEW_RESOLVED)
        self.assertFalse(bool(row["Automatically Excluded"]))
        self.assertIsNone(row["Excluded Amount"])

    def test_unresolved_candidate_is_held_and_excluded(self):
        report = self.review_report()
        updated = finalize_review_dispositions(report, held_ids=["ROW-2"])
        row = updated.loc[updated["Source Row ID"] == "ROW-2"].iloc[0]
        self.assertEqual(row["Disposition"], DISPOSITION_REVIEW_HOLD)
        self.assertTrue(bool(row["Automatically Excluded"]))
        self.assertEqual(row["Excluded Amount"], row["Amount"])

    def test_untouched_rows_keep_their_original_review_disposition(self):
        report = self.review_report()
        updated = finalize_review_dispositions(report, resolved_ids=["ROW-1"])
        untouched = updated.loc[updated["Source Row ID"] == "ROW-2"].iloc[0]
        self.assertEqual(untouched["Disposition"], DISPOSITION_REVIEW)

    def test_canonical_and_excess_rows_are_never_touched(self):
        # A strong, full-payload duplicate pair -- canonical/excess, not
        # review-tier -- must be immune to finalize_review_dispositions
        # even if its Source Row ID happens to be passed in by mistake.
        frame = make_frame([("100", "A", 10000)] * 2)
        result = screen_duplicates(frame, "Source Row ID", "QuickBooks", "Primary")
        updated = finalize_review_dispositions(
            result.report, resolved_ids=["ROW-1"], held_ids=["ROW-2"],
        )
        pd.testing.assert_frame_equal(updated, result.report)

    def test_empty_report_is_returned_unchanged(self):
        empty = make_frame([("", "", None)]).iloc[0:0]
        result = screen_duplicates(empty, "Source Row ID", "QuickBooks", "Primary")
        self.assertTrue(result.report.empty)
        updated = finalize_review_dispositions(result.report, resolved_ids=["X"])
        self.assertTrue(updated.empty)


if __name__ == "__main__":
    unittest.main()
