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
    SOURCE_POS,
    DISPOSITION_REVIEW,
    DISPOSITION_REVIEW_HOLD,
    DISPOSITION_REVIEW_RESOLVED,
    DuplicateScreeningError,
    cents_to_float,
    finalize_review_dispositions,
    screen_duplicates,
)


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
        self.assertEqual(row["Treatment"], DISPOSITION_REVIEW_RESOLVED)
        self.assertFalse(bool(row["Automatically Excluded"]))
        self.assertIsNone(row["Excluded Amount"])

    def test_unresolved_candidate_is_held_and_excluded(self):
        report = self.review_report()
        updated = finalize_review_dispositions(report, held_ids=["ROW-2"])
        row = updated.loc[updated["Source Row ID"] == "ROW-2"].iloc[0]
        self.assertEqual(row["Disposition"], DISPOSITION_REVIEW_HOLD)
        self.assertEqual(row["Treatment"], DISPOSITION_REVIEW_HOLD)
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
