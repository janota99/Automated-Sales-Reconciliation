import streamlit as st
import pandas as pd
import io
import time
from datetime import datetime

# Openpyxl Core Styling & Calculation Engines
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# MOCK ENGINES (Replace these with your actual local imports if available)
try:
    from engine.matching import execute_sales_reconciliation, get_fuzzy_lexicon_match
except ImportError:
    def execute_sales_reconciliation(df_qb, df_inf, qb_cols, inf_cols):
        df_qb = df_qb.copy()
        df_qb['IS_MATCHED'] = False  # Default simulation fallback layer
        return df_qb, df_inf
    def get_fuzzy_lexicon_match(memo):
        return "Standard Product Line" if pd.notna(memo) else "Unknown"

st.set_page_config(page_title="FinTech Rec Engine", layout="wide")

# -----------------------------------------------------------------
# 🏛️ INSTITUTIONAL AS400 EXTERNAL STYLESHEET INGESTION ENGINE
# -----------------------------------------------------------------
def load_external_css(css_file_path):
    try:
        with open(css_file_path, "r") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error(f"⚠️ Core Asset Missing: Could not find workspace sheet layout rules at '{css_file_path}'")

# Call layout template array injection
load_external_css("style.css")

# Cache & state configurations initialization loop
for key, default in {
    "processing_complete": False,
    "execution_time": 0.0,
    "last_run_time": "N/A",
    "selected_workbench_idx": None,
    "workbench_comments": {},
    "workbench_resolutions": {},
    "active_panel": "Dashboard Workspace",
    "net_variance": 0.0,
    "parsed_period": "05",
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

if "system_event_log" not in st.session_state:
    st.session_state.system_event_log = [
        f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SYSTEM: Core AS400 Platform Initialized. Awaiting Data Ingestion Streams."
    ]

# Backing data persistence slots
if "exceptions_df" not in st.session_state:
    st.session_state.exceptions_df = pd.DataFrame()
if "matched_df" not in st.session_state:
    st.session_state.matched_df = pd.DataFrame()
if "reconciled_qb_df" not in st.session_state:
    st.session_state.reconciled_qb_df = pd.DataFrame()  
if "agg_df" not in st.session_state:
    st.session_state.agg_df = pd.DataFrame()
if "mapped_columns" not in st.session_state:
    st.session_state.mapped_columns = {}
if "metrics_cache" not in st.session_state:
    st.session_state.metrics_cache = {
        "count": 0, "val": 0.0, "pct": 0.0,
        "qb_rows": 0, "qb_vol": 0.0,
        "inf_rows": 0, "inf_vol": 0.0,
        "matched_rows": 0, "matched_vol": 0.0
    }

def append_system_event(event_message):
    st.session_state.system_event_log.append(
        f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {event_message}"
    )

PRIOR_PERIOD_EXCEPTION_COUNT = 71
PRIOR_PERIOD_EXCEPTION_VALUE = 243200.00

# Sidebar Layout Paint Elements
st.sidebar.markdown("<h3 class='sidebar-header'>Workspace Routing</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<p class='sidebar-subheader'>OPERATIONAL MODULES</p>", unsafe_allow_html=True)

current_panel = st.sidebar.radio(
    "Operational Workspace Modules",
    ["Dashboard Workspace", "Exception Workbench [Review]"],
    label_visibility="collapsed",
    index=["Dashboard Workspace", "Exception Workbench [Review]"].index(st.session_state.active_panel) if st.session_state.active_panel in ["Dashboard Workspace", "Exception Workbench [Review]"] else 0
)
st.session_state.active_panel = current_panel

st.sidebar.markdown("<br><p class='sidebar-subheader'>ADMINISTRATION & CONTROLS</p>", unsafe_allow_html=True)

admin_panel = st.sidebar.radio(
    "Administrative Workspace Modules",
    ["System Audit Log", "Mapping Configuration"],
    label_visibility="collapsed"
)

if st.sidebar.button("🔓 Route to Selected Admin Panel", use_container_width=True):
    st.session_state.active_panel = admin_panel
    st.rerun()

current_panel = st.session_state.active_panel

with st.sidebar.expander("⚙️ Schema Field Mapping", expanded=False):
    qb_po = st.text_input("QB PO Column", value="P.O. NUMBER")
    qb_inv = st.text_input("QB Invoice Column", value="NUM")
    qb_amt = st.text_input("QB Amount Column", value="AMOUNT")
    qb_qty = st.text_input("QB QTY Column", value="QTY")
    st.markdown("<hr style='border-top:1px solid #334155; margin:10px 0;'>", unsafe_allow_html=True)
    inf_po = st.text_input("Infinium PO Column", value="OHDESC")
    inf_inv = st.text_input("Infinium Invoice Column", value="OHOBNO")
    inf_amt = st.text_input("Infinium Amount Column", value="OHTOTA")

st.sidebar.markdown("<hr style='border-top:1px solid #1E293B; margin:15px 0;'>", unsafe_allow_html=True)
st.sidebar.markdown("<h4 class='sidebar-section-title'>📥 Data Source Feeds</h4>", unsafe_allow_html=True)

file_qb = st.sidebar.file_uploader("QuickBooks Export (.xlsx)", type=["xlsx"])
file_inf = st.sidebar.file_uploader("Infinium Data Pull (.xlsx)", type=["xlsx"])

qb_rows, qb_vol, inf_rows, inf_vol = 0, 0.0, 0, 0.0
raw_qb = None

# Pure Data Extraction (Runs only when files are actively selected inside uploader buffers)
if file_qb:
    raw_qb = pd.read_excel(file_qb, header=None)
    h_idx = 0

    for idx, row in raw_qb.head(15).iterrows():
        row_str_list = [str(val).strip().upper() for val in row.values if pd.notna(val)]
        if "AMOUNT" in row_str_list or "P.O. NUMBER" in row_str_list:
            h_idx = idx
            break

    df_qb_raw = pd.read_excel(file_qb, header=h_idx)
    df_qb_raw.columns = df_qb_raw.columns.astype(str).str.strip().str.upper()

    memo_candidates = [c for c in df_qb_raw.columns if "MEMO" in c]
    qb_memo_col = "MEMO/DESCRIPTION" if "MEMO/DESCRIPTION" in df_qb_raw.columns else (memo_candidates[0] if memo_candidates else df_qb_raw.columns[0])

    df_qb_clean_rows = df_qb_raw[
        df_qb_raw[qb_inv.upper()].notna() | df_qb_raw[qb_memo_col].notna()
    ].dropna(how="all")

    qb_rows = len(df_qb_clean_rows)
    qb_vol = pd.to_numeric(df_qb_clean_rows[qb_amt.upper()], errors="coerce").sum()

if file_inf:
    df_inf_raw = pd.read_excel(file_inf)
    df_inf_raw.columns = df_inf_raw.columns.astype(str).str.strip().str.upper()
    df_inf_clean_rows = df_inf_raw.dropna(how="all")
    inf_rows = len(df_inf_clean_rows)
    inf_vol = pd.to_numeric(df_inf_clean_rows[inf_amt.upper()], errors="coerce").sum()

if file_qb and file_inf:
    st.session_state.net_variance = qb_vol - inf_vol
    if raw_qb is not None and len(raw_qb) > 1 and pd.notna(raw_qb.iloc[1, 2]):
        st.session_state.parsed_period = str(raw_qb.iloc[1, 2]).strip()

    QB_COLUMNS = {
        "po": qb_po.upper(),
        "invoice": qb_inv.upper(),
        "amount": qb_amt.upper(),
        "qty": qb_qty.upper(),
    }
    AS400_COLUMNS = {
        "po": inf_po.upper(),
        "invoice": inf_inv.upper(),
        "amount": inf_amt.upper(),
    }
    st.session_state.mapped_columns = {"qb": QB_COLUMNS, "inf": AS400_COLUMNS, "qb_memo": qb_memo_col}

# PANEL DISPATCH ROUTER
if current_panel == "Dashboard Workspace":
    st.markdown(f"""
        <div class='command-center-header'>
            <h1 class='command-center-title'>Sales Reconciliation Command Center</h1>
            <div class='command-center-subtitle'>
                Period {st.session_state.parsed_period} Close &nbsp; | &nbsp; QuickBooks Online ↔ Infinium AS400
            </div>
        </div>
    """, unsafe_allow_html=True)

    style_step1 = "background-color:#DBEAFE; border-color:#3B82F6; color:#1E40AF;" if file_qb else "background-color:#0F172A; border-color:#334155; color:#94A3B8;"
    style_step2 = "background-color:#DBEAFE; border-color:#3B82F6; color:#1E40AF;" if file_inf else "background-color:#0F172A; border-color:#334155; color:#94A3B8;"
    style_step3 = "background-color:#D1FAE5; border-color:#10B981; color:#065F46;" if st.session_state.processing_complete else "background-color:#0F172A; border-color:#334155; color:#94A3B8;"
    style_step4 = "background-color:#F1F5F9; border-color:#475569; color:#0F172A;" if st.session_state.processing_complete else "background-color:#0F172A; border-color:#334155; color:#94A3B8;"

    st.markdown(f"""
        <div class='process-flow-container'>
            <div class='flow-step' style='{style_step1}'>
                <b>1. Upload QuickBooks</b><br><span>{"✓ Ingested" if file_qb else "Pending Feed"}</span>
            </div>
            <div class='flow-arrow'>➔</div>
            <div class='flow-step' style='{style_step2}'>
                <b>2. Upload Infinium</b><br><span>{"✓ Ingested" if file_inf else "Pending Feed"}</span>
            </div>
            <div class='flow-arrow'>➔</div>
            <div class='flow-step' style='{style_step3}'>
                <b>3. Run Reconciliation</b><br><span>{st.session_state.last_run_time if st.session_state.processing_complete else "Awaiting Click"}</span>
            </div>
            <div class='flow-arrow'>➔</div>
            <div class='flow-step' style='{style_step4}'>
                <b>4. Review Exceptions</b><br><span>{f"{st.session_state.metrics_cache['count']} Items Open" if st.session_state.processing_complete else "Awaiting Close"}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)

    if not file_qb or not file_inf:
        st.markdown(
            "<div class='info-panel-box'><b>Data feeds open.</b> Populate source files in sidebar to initialize verification matrix columns.</div>",
            unsafe_allow_html=True
        )
    else:
        btn_col1, btn_col2 = st.columns([1.6, 5])

        with btn_col1:
            if st.button("Run Reconciliation", type="primary", use_container_width=True):
                start_timer = time.time()
                
                QB_COLUMNS = st.session_state.mapped_columns['qb']
                AS400_COLUMNS = st.session_state.mapped_columns['inf']
                qb_memo_col = st.session_state.mapped_columns['qb_memo']
                
                df_qb_pass = df_qb_raw[df_qb_raw[QB_COLUMNS['invoice']].notna() | df_qb_raw[qb_memo_col].notna()].copy()
                df_qb_pass = df_qb_pass.dropna(how="all").reset_index(drop=True)
                df_inf_pass = df_inf_clean_rows.reset_index(drop=True)

                reconciled_qb, reconciled_inf = execute_sales_reconciliation(
                    df_qb_pass, df_inf_pass, QB_COLUMNS, AS400_COLUMNS
                )

                st.session_state.exceptions_df = reconciled_qb[reconciled_qb["IS_MATCHED"] == False].copy()
                st.session_state.matched_df = reconciled_qb[reconciled_qb["IS_MATCHED"] == True].copy()
                st.session_state.reconciled_qb_df = reconciled_qb.copy()

                df_qb_agg_source = df_qb_pass.copy()
                df_qb_agg_source[QB_COLUMNS["qty"]] = pd.to_numeric(df_qb_agg_source[QB_COLUMNS["qty"]], errors="coerce").fillna(0)
                df_qb_agg_source[QB_COLUMNS["amount"]] = pd.to_numeric(df_qb_agg_source[QB_COLUMNS["amount"]], errors="coerce").fillna(0)
                df_qb_agg_source["Product Name"] = df_qb_agg_source[qb_memo_col].apply(get_fuzzy_lexicon_match)
                df_qb_agg_filtered = df_qb_agg_source[df_qb_agg_source["Product Name"].notna()].copy()

                agg_df = df_qb_agg_filtered.groupby("Product Name").agg(
                    Product_Quantity=(QB_COLUMNS["qty"], "sum"),
                    Product_Value=(QB_COLUMNS["amount"], "sum"),
                ).reset_index()

                st.session_state.agg_df = agg_df.sort_values(by="Product Name", ascending=True).rename(
                    columns={"Product_Quantity": "Product Quantity", "Product_Value": "Product Value"}
                )

                total_scanned_qb = len(reconciled_qb)
                cleared_count_qb = len(st.session_state.matched_df)
                exception_count_qb = len(st.session_state.exceptions_df)
                pct_matched_qb = (cleared_count_qb / total_scanned_qb) * 100 if total_scanned_qb else 0.0
                exception_value_qb = pd.to_numeric(st.session_state.exceptions_df[QB_COLUMNS["amount"]], errors="coerce").sum()
                matched_value_qb = pd.to_numeric(st.session_state.matched_df[QB_COLUMNS["amount"]], errors="coerce").sum()

                st.session_state.metrics_cache = {
                    "count": exception_count_qb,
                    "val": exception_value_qb,
                    "pct": pct_matched_qb,
                    "qb_rows": qb_rows,
                    "qb_vol": qb_vol,
                    "inf_rows": inf_rows,
                    "inf_vol": inf_vol,
                    "matched_rows": cleared_count_qb,
                    "matched_vol": matched_value_qb
                }

                st.session_state.processing_complete = True
                st.session_state.execution_time = max(round(time.time() - start_timer, 2), 0.02)
                st.session_state.last_run_time = datetime.now().strftime("%I:%M %p")
                append_system_event(f"VALIDATION: Extracted {qb_rows} records from QBO workbook.")
                append_system_event(f"EXECUTION: 1-to-1 match loops compiled in {st.session_state.execution_time}s.")
                st.rerun()

        with btn_col2:
            if st.button("Clear Workspace", key="clear_btn"):
                st.session_state.processing_complete = False
                st.session_state.execution_time = 0.0
                st.session_state.last_run_time = "N/A"
                st.session_state.selected_workbench_idx = None
                st.session_state.net_variance = 0.0
                st.session_state.exceptions_df = pd.DataFrame()
                st.session_state.matched_df = pd.DataFrame()
                st.session_state.reconciled_qb_df = pd.DataFrame()
                st.session_state.agg_df = pd.DataFrame()
                st.session_state.metrics_cache = {
                    "count": 0, "val": 0.0, "pct": 0.0,
                    "qb_rows": 0, "qb_vol": 0.0,
                    "inf_rows": 0, "inf_vol": 0.0,
                    "matched_rows": 0, "matched_vol": 0.0
                }
                st.rerun()

        if not st.session_state.processing_complete:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"""
                <div class='table-card-panel'>
                    <h4 class='panel-header-text'>🗂️ Pre-Execution Ledger Footprint Verification</h4>
                    <p class='panel-subheader-text'>Confirm ledger control totals before calling matching arrays:</p>
                    <table class='control-data-table'>
                        <thead>
                            <tr class='control-table-header'>
                                <th style='padding:8px;'>Source Ledger Stream</th>
                                <th style='padding:8px; text-align:right;'>Transaction Volume</th>
                                <th style='padding:8px; text-align:right;'>Control Gross Value</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr class='control-table-row'>
                                <td style='padding:10px 8px;'>QuickBooks Online Subledger</td>
                                <td style='padding:10px 8px; text-align:right;'>{qb_rows:,} rows</td>
                                <td style='padding:10px 8px; text-align:right;'><b>\${qb_vol:,.2f}</b></td>
                            </tr>
                            <tr class='control-table-row'>
                                <td style='padding:10px 8px;'>Infinium AS400 Base Ledger</td>
                                <td style='padding:10px 8px; text-align:right;'>{inf_rows:,} rows</td>
                                <td style='padding:10px 8px; text-align:right;'><b>\${inf_vol:,.2f}</b></td>
                            </tr>
                            <tr class='control-table-row-alert'>
                                <td class='control-table-cell-alert'>Net Reconciliation Variance</td>
                                <td class='control-table-cell-alert-right'>{abs(qb_rows - inf_rows)} rows net difference</td>
                                <td class='control-table-cell-alert-right'><b>\${st.session_state.net_variance:,.2f}</b></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            """, unsafe_allow_html=True)

        if st.session_state.processing_complete:
            st.markdown("<br>", unsafe_allow_html=True)

            variance_color = "background-color:#065F46; border:1px solid #047857;" if round(st.session_state.net_variance, 2) == 0 else "background-color:#8B1E1E; border:1px solid #B91C1C;"
            variance_label = "SYSTEM IN BALANCE" if round(st.session_state.net_variance, 2) == 0 else "OUT OF BALANCE"

            c = st.session_state.metrics_cache

            # Core Variance Banner
            st.markdown(f"""
                <div class='variance-banner' style='{variance_color}'>
                    <div class='variance-banner-label'>Subledger Reconciliation Status</div>
                    <div class='variance-banner-amount'>\${st.session_state.net_variance:,.2f}</div>
                    <div class='variance-banner-status'>{variance_label}</div>
                </div>
            """, unsafe_allow_html=True)

            # Balanced native Streamlit row layout blocks container
            sm1, sm2, sm3 = st.columns(3)
            with sm1:
                st.metric("Open Exceptions Count", f"{c['count']} Rows")
            with sm2:
                st.markdown(f"""
                    <div class='nested-metric-card'>
                        <span class='nested-metric-label'>Open Exception Value</span>
                        <h2 class='nested-metric-value'>\${c['val']:,.2f}</h2>
                    </div>
                """, unsafe_allow_html=True)
            with sm3:
                st.metric("Auto-Reconciled Percentage", f"{c['pct']:.1f}%")

            driver_agg = st.session_state.agg_df.sort_values(by="Product Value", ascending=False).head(3)

            st.markdown("<br>", unsafe_allow_html=True)
            grid_col1, grid_col2 = st.columns(2)

            with grid_col1:
                exposure_change = (
                    ((PRIOR_PERIOD_EXCEPTION_VALUE - c['val']) / PRIOR_PERIOD_EXCEPTION_VALUE) * 100
                    if PRIOR_PERIOD_EXCEPTION_VALUE
                    else 0
                )

                st.markdown(f"""
                    <div class='table-card-panel-small'>
                        <h5 class='panel-header-text' style='font-size:13px;'>Closing Cycle Variance Trend</h5>
                        <table class='control-data-table' style='font-size:13px; margin-top:18px;'>
                            <thead>
                                <tr class='control-table-header' style='font-size:11px;'>
                                    <th style='padding:6px;'>Metric</th>
                                    <th style='padding:6px; text-align:right;'>Prior P04</th>
                                    <th style='padding:6px; text-align:right;'>Current P05</th>
                                    <th style='padding:6px; text-align:right;'>Change</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr class='control-table-row'>
                                    <td style='padding:10px 6px;'>Exceptions Count</td>
                                    <td style='padding:10px 6px; text-align:right;'>{PRIOR_PERIOD_EXCEPTION_COUNT:,}</td>
                                    <td style='padding:10px 6px; text-align:right;'>{c['count']:,}</td>
                                    <td style='padding:10px 6px; text-align:right; color:#10B981; font-weight:700;'>
                                        {c['count'] - PRIOR_PERIOD_EXCEPTION_COUNT:+,}
                                    </td>
                                </tr>
                                <tr>
                                    <td style='padding:10px 6px;'>Exposure Value</td>
<tr class='control-table-row'>
    <td style='padding:10px 6px;'>Exposure Value</td>
    <td style='padding:10px 6px; text-align:right;'>&#36;{PRIOR_PERIOD_EXCEPTION_VALUE:,.2f}</td>
    <td style='padding:10px 6px; text-align:right;'>&#36;{c['val']:,.2f}</td>
    <td style='padding:10px 6px; text-align:right; color:#10B981; font-weight:700;'>
        -{exposure_change:.1f}%
    </td>
</tr>
                                    <td style='padding:10px 6px; text-align:right; color:#10B981; font-weight:700;'>
                                        -{exposure_change:.1f}%
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                """, unsafe_allow_html=True)

            with grid_col2:
                st.markdown("""
                    <div class='table-card-panel-small'>
                        <h5 class='panel-header-text' style='font-size:13px;'>Top Exception Drivers</h5>
                """, unsafe_allow_html=True)

                if driver_agg.empty:
                    st.markdown(
                        "<p class='panel-subheader-text' style='padding-top:20px;'>No product aggregate drivers identified.</p>",
                        unsafe_allow_html=True,
                    )
                else:
                    max_driver_value = driver_agg["Product Value"].max()

                    for _, d_row in driver_agg.iterrows():
                        bar_width = 0 if max_driver_value == 0 else (d_row["Product Value"] / max_driver_value) * 100
                        st.markdown(f"""
                            <div class='driver-row-wrapper'>
                                <div class='driver-row-data'>
                                    <span>{d_row["Product Name"]}</span>
                                    <b>\${d_row["Product Value"]:,.2f}</b>
                                </div>
                                <div class='driver-bar-track'>
                                    <div class='driver-bar-fill' style='width:{bar_width:.1f}%;'></div>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)

                st.markdown("</div>", unsafe_allow_html=True)

            # -----------------------------------------------------------------
            # 📊 UNIFIED INSTITUTIONAL EXCEL GENERATION ENGINE
            # -----------------------------------------------------------------
            pkg_buffer = io.BytesIO()
            with pd.ExcelWriter(pkg_buffer, engine="openpyxl") as writer:
                st.session_state.exceptions_df.to_excel(writer, index=False, sheet_name="1. Unresolved Exceptions")
                st.session_state.matched_df.to_excel(writer, index=False, sheet_name="2. Matched Transactions")
                st.session_state.agg_df.to_excel(writer, index=False, sheet_name="3. Product Aggregates Summary")
                st.session_state.reconciled_qb_df.to_excel(writer, index=False, sheet_name="Raw QuickBooks Sales")

                workbook = writer.book
                workbook.calculation.calcMode = "auto"

                font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                font_body = Font(name="Segoe UI", size=11)
                font_total = Font(name="Segoe UI", size=11, bold=True)
                fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
                fill_zebra = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
                fill_total = PatternFill(start_color="EAECEF", end_color="EAECEF", fill_type="solid")
                
                fill_matched_grey = PatternFill(start_color="EBEBEB", end_color="EBEBEB", fill_type="solid")
                fill_exception_amber = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                
                border_thin = Side(border_style="thin", color="D3D3D3")
                border_top_thick = Side(border_style="thin", color="000000")
                border_bottom_double = Side(border_style="double", color="000000")
                cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
                total_border = Border(top=border_top_thick, bottom=border_bottom_double)
                align_center = Alignment(horizontal="center", vertical="center")
                align_right = Alignment(horizontal="right", vertical="center")
                align_left = Alignment(horizontal="left", vertical="center")

                for sheet_name in workbook.sheetnames:
                    ws = workbook[sheet_name]
                    ws.views.sheetView[0].showGridLines = True
                    ws.row_dimensions[1].height = 26

                    for cell in ws[1]:
                        cell.font = font_header
                        cell.fill = fill_header
                        cell.alignment = align_center

                    max_row = ws.max_row
                    max_col = ws.max_column

                    is_raw_sheet = (sheet_name == "Raw QuickBooks Sales")
                    match_flag_col_idx = None
                    
                    if is_raw_sheet:
                        headers = [str(ws.cell(row=1, column=c).value).upper() for c in range(1, max_col + 1)]
                        if "IS_MATCHED" in headers:
                            match_flag_col_idx = headers.index("IS_MATCHED") + 1

                    for row_idx in range(2, max_row + 1):
                        ws.row_dimensions[row_idx].height = 20
                        
                        if is_raw_sheet and match_flag_col_idx:
                            is_matched_val = ws.cell(row=row_idx, column=match_flag_col_idx).value
                            if str(is_matched_val).strip().upper() in ["TRUE", "1", "MATCHED"]:
                                current_row_fill = fill_matched_grey
                            else:
                                current_row_fill = fill_exception_amber
                        else:
                            current_row_fill = fill_zebra if row_idx % 2 == 0 else None

                        for col_idx in range(1, max_col + 1):
                            cell = ws.cell(row=row_idx, column=col_idx)
                            cell.font = font_body
                            cell.border = cell_border
                            
                            if current_row_fill:
                                cell.fill = current_row_fill

                            header_text = str(ws.cell(row=1, column=col_idx).value).upper()

                            if "AMOUNT" in header_text or "VALUE" in header_text or "OHTOTA" in header_text:
                                cell.number_format = "$#,##0.00"
                                cell.alignment = align_right
                            elif "QTY" in header_text or "QUANTITY" in header_text:
                                cell.number_format = "#,##0"
                                cell.alignment = align_right
                            elif "DATE" in header_text or "IS_MATCHED" in header_text:
                                cell.alignment = align_center
                            else:
                                cell.alignment = align_left

                    if max_row > 1 and not is_raw_sheet:
                        tot_row_idx = max_row + 1
                        ws.row_dimensions[tot_row_idx].height = 22
                        ws.cell(row=tot_row_idx, column=1, value="Total").font = font_total

                        for col_idx in range(1, max_col + 1):
                            t_cell = ws.cell(row=tot_row_idx, column=col_idx)
                            t_cell.fill = fill_total
                            t_cell.border = total_border
                            t_cell.font = font_total

                            h_text = str(ws.cell(row=1, column=col_idx).value).upper()
                            col_letter = get_column_letter(col_idx)

                            if col_idx > 1:
                                if (
                                    "AMOUNT" in h_text
                                    or "VALUE" in h_text
                                    or col_letter in ["G", "H"]
                                ):
                                    t_cell.value = f"=SUM({col_letter}2:{col_letter}{max_row})"
                                    t_cell.number_format = "$#,##0.00"
                                    t_cell.alignment = align_right

                                if "QTY" in h_text or "QUANTITY" in h_text:
                                    t_cell.value = f"=SUM({col_letter}2:{col_letter}{max_row})"
                                    t_cell.number_format = "#,##0"
                                    t_cell.alignment = align_right

                    for col in ws.columns:
                        max_len = 0
                        col_letter = get_column_letter(col[0].column)

                        for cell in col:
                            if cell.value is not None:
                                val_str = str(cell.value)
                                if cell.number_format and "$" in cell.number_format:
                                    val_str += "   "
                                max_len = max(max_len, len(val_str))

                        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            st.markdown("<br><h5>System Package Actions</h5>", unsafe_allow_html=True)

            if st.download_button(
                label="📥 Generate Institutional Audit Package (.XLSX)",
                data=pkg_buffer.getvalue(),
                file_name="Sales_Reconciliation_Master_Package.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            ):
                append_system_event("EXPORTS: Unified corporate excel package compiled cleanly.")

# PANEL 2: AUDIT EXCEPTION WORKSHOP TERMINAL VIEW
elif current_panel == "Exception Workbench [Review]":
    if not st.session_state.processing_complete or st.session_state.exceptions_df.empty:
        st.warning("⚠️ Access Denied: Execute close processes on the dashboard panel to load open subledger items.")
    else:
        st.markdown("### 🔍 Exception Review Workbench")

        list_col, work_col = st.columns([2, 2])
        
        QB_COLUMNS = st.session_state.mapped_columns['qb']
        exceptions_df_raw = st.session_state.exceptions_df

        with list_col:
            st.markdown("#### Open Exceptions Vector")

            selection_options = [
                f"Row {idx} | PO: {row[QB_COLUMNS['po']]} | Amt: \${row[QB_COLUMNS['amount']]:,.2f}"
                for idx, row in exceptions_df_raw.iterrows()
            ]

            if not selection_options:
                st.success("🎉 Zero exceptions detected. All ledger lines cleared autonomously.")
            else:
                default_sel_idx = 0
                if st.session_state.selected_workbench_idx in exceptions_df_raw.index:
                    current_idx_str = f"Row {st.session_state.selected_workbench_idx} "
                    matching_opt = [o for o in selection_options if o.startswith(current_idx_str)]
                    if matching_opt:
                        default_sel_idx = selection_options.index(matching_opt[0])

                selected_item_str = st.selectbox(
                    "Select Variance Row Target",
                    selection_options,
                    label_visibility="collapsed",
                    index=default_sel_idx
                )

                selected_idx = int(selected_item_str.split(" | ")[0].replace("Row ", ""))
                st.session_state.selected_workbench_idx = selected_idx

                st.markdown("<div class='quick-entry-note-title'><b>📝 Quick Entry Item Note:</b></div>", unsafe_allow_html=True)
                
                existing_quick_note = st.session_state.workbench_comments.get(selected_idx, "")
                updated_note = st.text_input(
                    "Add / Edit clearance details for selected row",
                    value=existing_quick_note,
                    key=f"quick_note_{selected_idx}",
                    label_visibility="collapsed",
                    placeholder="e.g., AP is researching, received via wire, transit variance..."
                )
                
                if updated_note != existing_quick_note:
                    st.session_state.workbench_comments[selected_idx] = updated_note
                    append_system_event(f"NOTES: Captured audit text statement on index row #{selected_idx}.")

                st.dataframe(exceptions_df_raw, use_container_width=True)

        with work_col:
            if (
                st.session_state.selected_workbench_idx is not None
                and st.session_state.selected_workbench_idx in exceptions_df_raw.index
            ):
                target_row = exceptions_df_raw.loc[st.session_state.selected_workbench_idx]

                ranked_exceptions = exceptions_df_raw.copy()
                ranked_exceptions["_ABS_AMOUNT"] = pd.to_numeric(
                    ranked_exceptions[QB_COLUMNS["amount"]],
                    errors="coerce",
                ).abs()

                ranked_exceptions["_RANK_BY_VALUE"] = ranked_exceptions["_ABS_AMOUNT"].rank(
                    method="dense",
                    ascending=False,
                ).fillna(1).astype(int)

                target_abs_amount = abs(pd.to_numeric(target_row[QB_COLUMNS["amount"]], errors="coerce"))
                target_rank = int(
                    ranked_exceptions.loc[
                        st.session_state.selected_workbench_idx,
                        "_RANK_BY_VALUE",
                    ]
                )
                
                exception_value_qb = st.session_state.metrics_cache['val']
                target_exposure_pct = (
                    target_abs_amount / abs(exception_value_qb) * 100
                    if exception_value_qb
                    else 0
                )

                # -------------------------------------------------------------
                # 🛠️ FIXED: CLEAN NATIVE MITIGATION VISUALIZATION
                # -------------------------------------------------------------
                st.markdown(f"### Mitigate Row Entry Reference #{st.session_state.selected_workbench_idx}")
                
                # Render key risk analytics metrics horizontally via native components
                m_col1, m_col2, m_col3 = st.columns(3)
                m_col1.metric("Total Open Exceptions", f"{st.session_state.metrics_cache['count']}")
                m_col2.metric("Exposure Value Rank", f"#{target_rank}")
                m_col3.metric("Total Exposure Contribution", f"{target_exposure_pct:.2f}%")
                
                st.markdown("---")
                
                # Fetch data directly from mapped dictionary targets safely
                po_num_val = target_row.get(QB_COLUMNS['po'], "N/A")
                inv_num_val = target_row.get(QB_COLUMNS['invoice'], "N/A")
                raw_amt_val = target_row.get(QB_COLUMNS['amount'], 0.0)
                
                # Display clean, scannable field references
                st.markdown(f"**Purchase Order Number:** `{po_num_val}`")
                st.markdown(f"**Invoice Number:** `{inv_num_val}`")
                st.markdown(f"**QuickBooks Base Amount:** `\${pd.to_numeric(raw_amt_val, errors='coerce'):,.2f}`")
                st.markdown("**Exception Classification:** :red[Missing AS400 AS400 Transaction]")
                st.markdown(f"**Granular Item Variance:** :red[\${pd.to_numeric(raw_amt_val, errors='coerce'):,.2f}]")
                
                st.markdown("---")
                st.markdown("##### Resolution Actions Management")

                current_saved_res = st.session_state.workbench_resolutions.get(
                    st.session_state.selected_workbench_idx,
                    "Unassigned Open Item",
                )
                current_saved_comment = st.session_state.workbench_comments.get(
                    st.session_state.selected_workbench_idx,
                    "",
                )

                resolution_options = [
                    "Unassigned Open Item",
                    "Timing Difference - Transit Accrual Required",
                    "Freight Classification Variance Adjustment",
                    "Approved AS400 Variance Write-off Entry",
                ]

                resolution_type = st.selectbox(
                    "Assign Action Status",
                    resolution_options,
                    index=resolution_options.index(current_saved_res),
                )

                audit_comment = st.text_area(
                    "Audit Log Clearance Comments (Mirrored)",
                    value=current_saved_comment,
                    key=f"detailed_note_{st.session_state.selected_workbench_idx}"
                )
                
                if audit_comment != current_saved_comment:
                    st.session_state.workbench_comments[st.session_state.selected_workbench_idx] = audit_comment

                w_btn1, w_btn2 = st.columns(2)

                with w_btn1:
                    if st.button("Commit Resolution State", type="primary", use_container_width=True):
                        st.session_state.workbench_resolutions[st.session_state.selected_workbench_idx] = resolution_type
                        append_system_event(
                            f"MITIGATION: Override status locked on index row #{st.session_state.selected_workbench_idx} [Status: {resolution_type}]."
                        )
                        st.success(f"✓ Resolution assigned to row #{st.session_state.selected_workbench_idx}.")

                with w_btn2:
                    if st.button("Clear Selection", use_container_width=True):
                        st.session_state.selected_workbench_idx = None
                        st.rerun()

                if st.session_state.selected_workbench_idx in st.session_state.workbench_resolutions:
                    st.markdown(f"""
                        <div class='resolution-profile-lockbox'>
                            <b>🔒 Documented Resolution Profile:</b><br>
                            • Status Type: {st.session_state.workbench_resolutions[st.session_state.selected_workbench_idx]}<br>
                            • Comment: "{st.session_state.workbench_comments[st.session_state.selected_workbench_idx]}"
                        </div>
                    """, unsafe_allow_html=True)

# SYSTEM AUDIT LOG PANEL
elif current_panel == "System Audit Log":
    st.markdown("### 📋 AS400 System Event Log & Security Audit Trail")
    st.markdown(
        "<p style='color:#94A3B8; font-size:14px; margin-bottom:20px;'>SOC 1/2 Compliance Stream: Verifiable platform orchestration and system matching execution logs:</p>",
        unsafe_allow_html=True,
    )
    st.markdown("<div class='system-log-container'>", unsafe_allow_html=True)

    for log_record in st.session_state.system_event_log:
        st.markdown(f"<div class='system-log-line'>{log_record}</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# CONFIGURATION PANEL
elif current_panel == "Mapping Configuration":
    st.markdown("### ⚙️ Mapping Configuration Settings")
    st.markdown("""
        <div class='table-card-panel' style='padding:20px;'>
            <h5 style='color:#3B82F6; margin-top:0;'>🔒 Environment Execution Constants</h5>
            <table class='control-data-table' style='font-size:13px;'>
                <tr class='control-table-row'><td style='padding:8px 0; font-weight:600;'>System Hosting Stack:</td><td>Local Port Server Endpoint</td></tr>
                <tr class='control-table-row'><td style='padding:8px 0; font-weight:600;'>Array Engine Driver:</td><td>Multi-Pass Dataframe Matching Engine</td></tr>
                <tr class='control-table-row'><td style='padding:8px 0; font-weight:600;'>Reconciliation Security State:</td><td>Stable Session Container Confirmed</td></tr>
            </table>
        </div>
    """, unsafe_allow_html=True)