from datetime import datetime, timezone

import pytest


@pytest.fixture
def qb_mapping():
    return {
        "po": "PO",
        "invoice": "Invoice",
        "amount": "Amount",
        "quantity": "Qty",
        "period": "Period",
        "product": None,
        # Line-level identity fields for the duplicate fingerprint.
        "customer": "Customer",
        "date": "Date",
    }


@pytest.fixture
def inf_mapping():
    return {
        "po": "PO", "invoice": "Invoice", "amount": "Amount", "period": "Period",
        "customer": "Customer", "date": "Date",
    }


@pytest.fixture
def make_metadata():
    """Factory fixture: make_metadata() -> a valid metadata dict for build_reconciliation."""

    def _make(run_id="TEST-RUN", fiscal_year=2026, fiscal_period=1, **overrides):
        metadata = {
            "run_id": run_id,
            "run_timestamp_dt": datetime.now(timezone.utc),
            "run_timestamp": "test",
            "qb_filename": "qb.xlsx",
            "qb_sha256": "x",
            "inf_filename": "inf.xlsx",
            "inf_sha256": "y",
            "fiscal_year": fiscal_year,
            "fiscal_period": fiscal_period,
            "qb_subtotal_rows_excluded": 0,
        }
        metadata.update(overrides)
        return metadata

    return _make
