"""Reusable Streamlit interface blocks.

Every function that renders something onto the page lives here: CSS
loading, KPI cards, dataframe previews, the ingestion-flow stepper, upload
status badges, and the full tabbed results view. Workbook bytes themselves
are still produced by ``workpapers.py``; this module only calls those
builders from the Downloads tab and displays the result.
"""

from __future__ import annotations

import html
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st

from matching import QB_ID, ReconciliationResult, numeric_sum
from utils import format_currency
from workpapers import build_analytics_workbook, build_primary_workbook, paired_display_frames


def load_app_css() -> None:
    """Load the stylesheet located beside the application module."""
    css_path = Path(__file__).resolve().with_name("style.css")
    try:
        css = css_path.read_text(encoding="utf-8")
        css = (
            css.replace("\u00a0", " ")
            .replace("\u2007", " ")
            .replace("\u202f", " ")
        )
    except OSError as exc:
        st.warning(f"The application stylesheet could not be loaded: {exc}")
        return

    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def render_kpi(label: str, value: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="rec-card">
            <div class="rec-kpi-label">{label}</div>
            <div class="rec-kpi-value">{value}</div>
            <div class="rec-kpi-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def display_primary_preview(result: ReconciliationResult) -> pd.DataFrame:
    qb_display, inf_display, methods = paired_display_frames(result)
    qb_display = qb_display.add_prefix("QB | ")
    inf_display = inf_display.add_prefix("INF | ")
    return pd.concat(
        [qb_display, pd.DataFrame({"Match Result": methods}), inf_display], axis=1
    )


def render_limited_dataframe(
    frame: pd.DataFrame,
    *,
    height: Optional[int] = None,
    row_limit: int = 2000,
) -> None:
    """Keep browser previews responsive while preserving full Excel outputs."""
    displayed = frame.head(row_limit)
    kwargs: dict[str, Any] = {
        "use_container_width": True,
        "hide_index": True,
    }
    if height is not None:
        kwargs["height"] = height
    st.dataframe(displayed, **kwargs)
    if len(frame) > row_limit:
        st.caption(
            f"Showing the first {row_limit:,} of {len(frame):,} rows. "
            "The complete population is retained in the downloadable workbook."
        )


def show_toast_once(state_key: str, message: str, icon: str = "✅") -> None:
    """Show a native toast once without repeating it on every app rerun."""
    if not st.session_state.get(state_key, False):
        st.toast(message, icon=icon)
        st.session_state[state_key] = True


def attention_tab_label(label: str, count: int) -> str:
    """Add a compatible text badge only when a tab has review items."""
    return f"{label} ({count:,})" if count > 0 else label


def render_source_status(filename: Optional[str], next_step: str) -> None:
    """Renders the file ingestion status directly beneath the file uploader in the main interface."""
    if filename:
        st.markdown(
            f'<div class="source-status complete">Loaded: {html.escape(filename)}</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div class="source-status active">Current step: {next_step}</div>',
            unsafe_allow_html=True,
        )


def render_upload_source_heading(
    step: int,
    label: str,
    logo_uri: str,
    logo_class: str,
) -> None:
    """Renders the file upload headings in the main interface columns."""
    st.markdown(
        f'<div class="upload-source-heading">'
        f'<div class="upload-source-copy"><span class="upload-step-number">{step}</span>'
        f'<span class="upload-source-title">{html.escape(label)}</span></div>'
        f'<img class="upload-source-logo {html.escape(logo_class)}" '
        f'src="{logo_uri}" alt="{html.escape(label)} logo">'
        f'</div>',
        unsafe_allow_html=True,
    )


def render_ingestion_flow(
    qb_loaded: bool,
    inf_loaded: bool,
    mappings_ready: bool,
    reconciliation_complete: bool,
    source_validation_failed: bool = False,
) -> None:
    step1_class = "complete" if qb_loaded else "active"
    step1_status = "Upload confirmed" if qb_loaded else "Upload QuickBooks first"
    if inf_loaded:
        step2_class, step2_status = "complete", "Upload confirmed"
    elif qb_loaded:
        step2_class, step2_status = "active", "Upload Infinium next"
    else:
        step2_class, step2_status = "upcoming", "Available after QuickBooks"
    if reconciliation_complete:
        step3_class, step3_status = "complete", "Controls validated"
    elif qb_loaded and inf_loaded and source_validation_failed:
        step3_class, step3_status = "active", "Resolve source validation"
    elif qb_loaded and inf_loaded and mappings_ready:
        step3_class, step3_status = "active", "Ready to run"
    elif qb_loaded and inf_loaded:
        step3_class, step3_status = "active", "Confirm required mappings"
    else:
        step3_class, step3_status = "upcoming", "Waiting for both files"
    if reconciliation_complete:
        step4_class, step4_status = "active", "Review and download"
    else:
        step4_class, step4_status = "upcoming", "Available after reconciliation"

    steps = [
        (1, "Ingest QuickBooks", step1_status, step1_class),
        (2, "Ingest Infinium", step2_status, step2_class),
        (3, "Run Reconciliation", step3_status, step3_class),
        (4, "Review Results", step4_status, step4_class),
    ]
    # Keep this markup on one physical line. Markdown treats indented HTML
    # after a blank line as a code block, which previously rendered card 1 but
    # displayed cards 2-4 as literal <div> text.
    cards = "".join(
        f'<div class="ingestion-step {css_class}">'
        f'<span class="ingestion-number">{number}</span>'
        f'<span class="ingestion-title">{title}</span>'
        f'<div class="ingestion-status">{status}</div>'
        f'</div>'
        for number, title, status, css_class in steps
    )
    flow_html = f'<div class="ingestion-flow">{cards}</div>'
    st.markdown(flow_html, unsafe_allow_html=True)


def render_pre_execution_controls(
    qb_raw: pd.DataFrame,
    inf_raw: pd.DataFrame,
    qb_mapping: dict[str, Optional[str]],
    inf_mapping: dict[str, Optional[str]],
) -> None:
    qb_total = numeric_sum(qb_raw[qb_mapping["amount"]])
    inf_total = numeric_sum(inf_raw[inf_mapping["amount"]])
    st.subheader("Source verification")
    cols = st.columns(4)
    with cols[0]:
        render_kpi("QuickBooks rows", f"{len(qb_raw):,}", format_currency(qb_total))
    with cols[1]:
        render_kpi("Infinium rows", f"{len(inf_raw):,}", format_currency(inf_total))
    with cols[2]:
        render_kpi("Row difference", f"{len(qb_raw) - len(inf_raw):,}", "QB minus Infinium")
    with cols[3]:
        render_kpi("Source value difference", format_currency(qb_total - inf_total), "QB minus Infinium")
    st.caption(
        "Pre-reconciliation control totals include only QuickBooks rows populated in every field other than Quantity and Amount. "
        "Retained rows with invalid amounts are flagged and remain unreconciled."
    )


def render_result(result: ReconciliationResult) -> None:
    metrics = result.metrics
    control_status = metrics["Control Status"]
    if control_status == "PASS":
        show_toast_once(
            f"reconciliation_complete_{result.run_id}",
            f"Reconciliation complete. All controls passed. Run ID: {result.run_id}",
        )
    else:
        st.error("Reconciliation completed with failed controls. Downloads are withheld until controls pass.")

    failed_controls = int(result.controls["Status"].ne("PASS").sum())
    invalid_amount_rows = int(
        metrics["Invalid QuickBooks Amounts"] + metrics["Invalid Infinium Amounts"]
    )
    unmatched_qb = int(metrics["Unresolved QuickBooks Rows"])
    unmatched_inf = int(metrics["Unmatched Infinium Rows"])
    unresolved_duplicate_groups = 0
    if not result.duplicate_analysis.empty:
        unresolved_duplicate_groups = int(
            result.duplicate_analysis["Reconciliation Status"]
            .ne("All rows matched through a more specific unique key")
            .sum()
        )

    tab_labels = [
        attention_tab_label("Overview", failed_controls + invalid_amount_rows),
        attention_tab_label("Reconciled View", unmatched_qb + unmatched_inf),
        attention_tab_label("Exception Review", unmatched_qb),
        attention_tab_label("Analytics", unresolved_duplicate_groups),
        "Downloads",
    ]
    overview_tab, reconciled_tab, exceptions_tab, analytics_tab, downloads_tab = st.tabs(
        tab_labels
    )
    with overview_tab:
        cols = st.columns(4)
        with cols[0]:
            render_kpi(
                "QB match rate",
                f"{metrics['QuickBooks Match Rate by Row']:.1%}",
                f"{metrics['Matched QuickBooks Rows']:,} of {metrics['QuickBooks Rows']:,} rows",
            )
        with cols[1]:
            render_kpi(
                "Unresolved QB",
                format_currency(metrics["Unresolved QuickBooks Amount"]),
                f"{metrics['Unresolved QuickBooks Rows']:,} transactions",
            )
        with cols[2]:
            render_kpi(
                "Matched control difference",
                format_currency(metrics["Matched Amount Difference"]),
                "QuickBooks minus Infinium",
            )
        with cols[3]:
            render_kpi(
                "Control status",
                control_status,
                f"{len(result.controls)} required controls",
            )
        st.markdown("#### Required controls")
        st.dataframe(result.controls, use_container_width=True, hide_index=True)
        if metrics.get("Historical Clearances", 0):
            st.info(
                f"Historical secondary data cleared {metrics['Historical Clearances']:,} opposing-primary "
                "exception item(s). "
                f"{metrics['QuickBooks Secondary Rows Ignored'] + metrics['Infinium Secondary Rows Ignored']:,} "
                "unused secondary row(s) were excluded from exception reporting."
            )
        if metrics["Invalid QuickBooks Amounts"] or metrics["Invalid Infinium Amounts"]:
            st.markdown(
                '<div class="data-quality-banner">Data quality review: '
                f"{metrics['Invalid QuickBooks Amounts']} QuickBooks and "
                f"{metrics['Invalid Infinium Amounts']} Infinium rows have invalid or missing amounts. "
                "Those rows were not automatically matched.</div>",
                unsafe_allow_html=True,
            )

    with reconciled_tab:
        st.caption(
            "Matched rows appear first. Unmatched QuickBooks and unmatched Infinium rows are kept in separate "
            "sections so unrelated exceptions are never displayed as a pair. Historical clearances display the "
            "specific accepted prior-period row on the opposing side and label it in Record Context. Unused "
            "historical rows are never added to this ledger."
        )
        render_limited_dataframe(display_primary_preview(result), height=520)

    with exceptions_tab:
        unresolved = result.qb_work.loc[result.unmatched_qb, list(result.qb_raw.columns)].copy()
        if not result.candidates.empty:
            candidate_map = result.candidates.set_index("QuickBooks Row ID").to_dict("index")
            unresolved["Exception Status"] = [
                candidate_map.get(result.qb_work.at[idx, QB_ID], {}).get(
                    "Disposition", "Unmatched QuickBooks"
                )
                for idx in result.unmatched_qb
            ]
            unresolved["Reference Amount Difference"] = [
                candidate_map.get(result.qb_work.at[idx, QB_ID], {}).get(
                    "Minimum Amount Difference"
                )
                for idx in result.unmatched_qb
            ]
        if unresolved.empty:
            st.caption("No unresolved QuickBooks exceptions remain.")
        else:
            render_limited_dataframe(unresolved, height=440)
        st.metric("Proposed journal-entry support total", format_currency(metrics["Unresolved QuickBooks Amount"]))
        st.caption("This is the net signed amount of unresolved QuickBooks transactions. Review credits and reversals before posting.")

    with analytics_tab:
        left, right = st.columns(2)
        with left:
            st.markdown("#### Match-method distribution")
            st.dataframe(result.method_summary, use_container_width=True, hide_index=True)
        with right:
            st.markdown("#### Exception analysis")
            st.dataframe(result.exception_analysis, use_container_width=True, hide_index=True)
        with st.expander("Normalization and assessment detail", expanded=False):
            st.dataframe(result.assessments, use_container_width=True, hide_index=True, height=320)
    with downloads_tab:
        st.markdown("#### Accounting workpaper")
        st.caption("Four sheets: Raw Data, Reconciled Data, Unresolved Exceptions, and Product Aggregate Summary.")
        if "primary_workbook" not in st.session_state:
            if st.button(
                "Prepare Sales Reconciliation",
                type="primary",
                use_container_width=True,
                key=f"prepare_primary_{result.run_id}",
            ):
                try:
                    with st.spinner("Preparing the accounting workpaper..."):
                        st.session_state.primary_workbook = build_primary_workbook(result)
                    st.toast("Accounting workpaper prepared.", icon="✅")
                except Exception as exc:
                    st.error(f"The accounting workpaper could not be prepared: {exc}")
        if "primary_workbook" in st.session_state:
            st.download_button(
                "Download Sales Reconciliation",
                data=st.session_state.primary_workbook,
                file_name=f"Sales_Reconciliation_{result.run_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        st.markdown("#### Detailed analytics")
        st.caption("Optional evidence package with normalization, assessments, controls, and exact run configuration.")
        if "analytics_workbook" not in st.session_state:
            if st.button(
                "Prepare Reconciliation Analytics",
                use_container_width=True,
                key=f"prepare_analytics_{result.run_id}",
            ):
                try:
                    with st.spinner("Preparing the detailed analytics workbook..."):
                        st.session_state.analytics_workbook = build_analytics_workbook(result)
                    st.toast("Analytics workbook prepared.", icon="✅")
                except Exception as exc:
                    st.error(f"The analytics workbook could not be prepared: {exc}")
        if "analytics_workbook" in st.session_state:
            st.download_button(
                "Download Reconciliation Analytics",
                data=st.session_state.analytics_workbook,
                file_name=f"Sales_Reconciliation_Analytics_{result.run_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )