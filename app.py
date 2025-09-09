# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from datetime import datetime

# ---------------- Page configuration ----------------
st.set_page_config(
    page_title="LG Branch Summary Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- Custom CSS ----------------
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1rem;
        background: linear-gradient(90deg, #f0f8ff, #e6f3ff);
        border-radius: 10px;
        border-left: 5px solid #1f4e79;
    }
    .metric-container {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #1f4e79;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 0.5rem 0;
    }
    .section-header {
        color: #1f4e79;
        font-size: 1.5rem;
        font-weight: bold;
        margin: 1.5rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #e6f3ff;
    }
    .upload-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px dashed #1f4e79;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ---------------- Helpers ----------------
def _std(s: str) -> str:
    """Normalize header to UPPER_SNAKE."""
    return str(s).strip().upper().replace(" ", "_").replace("-", "_")

def _clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Drop empty rows/cols; if headers look 'UNNAMED', promote the best row as header."""
    d = df.copy()
    d = d.dropna(how='all')
    d = d.dropna(axis=1, how='all')

    if any(str(c).upper().startswith("UNNAMED") for c in d.columns):
        hdr_idx = d.notna().sum(axis=1).idxmax()
        new_cols = [str(x).strip() if pd.notna(x) else "" for x in d.loc[hdr_idx].tolist()]
        d = d.iloc[hdr_idx + 1:].reset_index(drop=True)
        if len(new_cols) == d.shape[1]:
            d.columns = new_cols

    d = d.loc[:, ~d.columns.duplicated()]
    d = d[[c for c in d.columns if str(c).strip() != "" and not str(c).upper().startswith("UNNAMED")]]
    return d

def _format_dates_for_display(df: pd.DataFrame, cols=('ISSUE_DATE', 'EXPIRY_DATE')) -> pd.DataFrame:
    d = df.copy()
    for c in cols:
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')
    return d

def _ensure_rates_only(df_sum: pd.DataFrame) -> pd.DataFrame:
    """Only compute percentage rates if source columns exist. Never fabricate totals."""
    d = df_sum.copy()
    if {"AMOUNT_UTILIZED", "TOTAL_FACILITIES"}.issubset(d.columns):
        with np.errstate(divide='ignore', invalid='ignore'):
            d["UTILIZATION_RATE"] = (d["AMOUNT_UTILIZED"] / d["TOTAL_FACILITIES"] * 100).round(2)
    if {"OUTSTANDING", "TOTAL_FACILITIES"}.issubset(d.columns):
        with np.errstate(divide='ignore', invalid='ignore'):
            d["OUTSTANDING_RATE"] = (d["OUTSTANDING"] / d["TOTAL_FACILITIES"] * 100).round(2)
    return d

# ---------------- Data loaders ----------------
@st.cache_data
def load_summary_data_from_file(uploaded_file):
    """
    Robustly extract the BANKS summary block from the 'Summary' sheet:
    headers anywhere on the sheet like [BANKS | TOTAL FACILITIES | AMOUNT UTILIZED | OUTSTANDING].
    Never fabricate numbers.
    """
    wanted = ["BANKS", "TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"]

    def _find_summary_block(df_raw: pd.DataFrame):
        # return (start_row_index, col_indices_dict) or (None, None)
        for r in range(len(df_raw)):
            row_vals = [str(v).strip() if pd.notna(v) else "" for v in df_raw.iloc[r].tolist()]
            cols_map = {}
            for ci, val in enumerate(row_vals):
                v = _std(val)
                if not v:
                    continue
                if v.startswith("BANK"):
                    cols_map["BANKS"] = ci
                if "TOTAL" in v and ("FACILITY" in v or "FACILITIES" in v):
                    cols_map["TOTAL FACILITIES"] = ci
                if ("AMOUNT" in v and ("UTILIZED" in v or "UTILISED" in v or "USED" in v)):
                    cols_map["AMOUNT UTILIZED"] = ci
                if v.startswith("OUTSTANDING") or v == "AVAILABLE" or v == "BALANCE":
                    cols_map["OUTSTANDING"] = ci
            if len(set(cols_map.keys()) & set(wanted)) >= 2:
                return r, cols_map
        return None, None

    # 1) Read summary sheet raw (no header)
    try:
        raw = pd.read_excel(uploaded_file, sheet_name="Summary", header=None)
    except Exception as e:
        st.error(f"Failed to open 'Summary' sheet: {e}")
        return None

    start_row, cmap = _find_summary_block(raw)
    if start_row is not None:
        data = raw.iloc[start_row + 1:].copy()
        ordered = [c for c in ["BANKS", "TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"] if c in cmap]
        use_cols = [cmap[c] for c in ordered]
        d = data[use_cols].copy()
        d.columns = ordered

        # drop rows where BANKS is blank or totals
        d["BANKS"] = d["BANKS"].astype(str).str.strip()
        d = d[~d["BANKS"].str.upper().isin(["", "TOTAL", "TOTALS", "SUM"])]

        # numeric coercions
        for c in ["TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"]:
            if c in d.columns:
                d[c] = pd.to_numeric(d[c], errors="coerce")

        # standardize names
        d = d.rename(columns={
            "TOTAL FACILITIES": "TOTAL_FACILITIES",
            "AMOUNT UTILIZED": "AMOUNT_UTILIZED",
            "OUTSTANDING": "OUTSTANDING",
        })

        d = _ensure_rates_only(d)
        if "BANKS" in d.columns and "AMOUNT_UTILIZED" in d.columns:
            return d

    # 2) Fallback: build minimal summary from detailed sheet
    try:
        df_det = pd.read_excel(uploaded_file, sheet_name="LG BRANCH SUMMARY_2025", usecols="A:K")
        df_det = _clean_frame(df_det)
        df_det.columns = [_std(c) for c in df_det.columns]
        if "BANK" not in df_det.columns or "AMOUNT" not in df_det.columns:
            raise KeyError("Could not find BANK/AMOUNT in detailed sheet to build summary.")
        tmp = (
            df_det.groupby("BANK", dropna=True)["AMOUNT"]
            .sum()
            .reset_index()
            .rename(columns={"BANK": "BANKS", "AMOUNT": "AMOUNT_UTILIZED"})
        )
        st.info("Built minimal Summary from detailed sheet (BANKS + AMOUNT_UTILIZED).")
        return tmp
    except Exception as e_fallback:
        st.error(f"Error loading summary data (Summary + fallback failed): {e_fallback}")
        return None

@st.cache_data
def load_detailed_data_from_file(uploaded_file):
    """Load detailed data (LG BRANCH SUMMARY_2025, A:K), robust to messy sheets."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name='LG BRANCH SUMMARY_2025', usecols="A:K")
        df = _clean_frame(df)

        # Map by position to stable names
        expected_columns = {
            0: 'BANK',            # A
            1: 'LG_REF',          # B
            2: 'CUSTOMER_NAME',   # C
            3: 'GUARANTEE_TYPE',  # D
            4: 'ISSUE_DATE',      # E
            5: 'EXPIRY_DATE',     # F
            6: 'AMOUNT',          # G
            7: 'CURRENCY',        # H
            8: 'BRANCH',          # I
            9: 'BANK_2',          # J (duplicate)
            10: 'DAYS_TO_MATURE'  # K
        }
        df.columns = [expected_columns.get(i, f'Column_{i}') for i, _ in enumerate(df.columns)]
        df = df.drop('BANK_2', axis=1, errors='ignore')

        # Remove empty/header rows
        df = df[df['BANK'].notna()]
        df = df[df['BANK'] != 'BANK']

        # Clean types
        if 'GUARANTEE_TYPE' in df.columns:
            df['GUARANTEE_TYPE'] = df['GUARANTEE_TYPE'].astype(str).str.strip()
        if 'AMOUNT' in df.columns:
            df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce').fillna(0)
        if 'DAYS_TO_MATURE' in df.columns:
            df['DAYS_TO_MATURE'] = pd.to_numeric(df['DAYS_TO_MATURE'], errors='coerce').fillna(30)
        for date_col in ['ISSUE_DATE', 'EXPIRY_DATE']:
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        # Optional fills
        if 'LG_REF' in df.columns:
            df['LG_REF'] = df['LG_REF'].fillna('N/A')
        if 'CUSTOMER_NAME' in df.columns:
            df['CUSTOMER_NAME'] = df['CUSTOMER_NAME'].fillna('Unknown Customer')
        if 'BRANCH' in df.columns:
            df['BRANCH'] = df['BRANCH'].fillna('Main Branch')
        if 'CURRENCY' in df.columns:
            df['CURRENCY'] = df['CURRENCY'].fillna('SAR')

        # Keep essential rows only
        df = df.dropna(subset=['BANK', 'GUARANTEE_TYPE'])
        df = df[df['BANK'].astype(str).str.strip() != '']
        df = df[df['GUARANTEE_TYPE'].astype(str).str.strip() != '']
        return df

    except Exception as e:
        st.error(f"Error loading detailed data: {str(e)}")
        try:
            df_fallback = pd.read_excel(uploaded_file, sheet_name='LG BRANCH SUMMARY_2025')
            df_fallback = _clean_frame(df_fallback)
            # best-effort rename by name patterns
            column_mapping = {}
            for col in df_fallback.columns:
                col_str = _std(col)
                if col_str == 'BANK':
                    column_mapping[col] = 'BANK'
                elif 'LG_REF' in col_str or col_str == 'LG_REF':
                    column_mapping[col] = 'LG_REF'
                elif 'CUSTOMER' in col_str:
                    column_mapping[col] = 'CUSTOMER_NAME'
                elif 'GUARR' in col_str or 'TYPE' in col_str:
                    column_mapping[col] = 'GUARANTEE_TYPE'
                elif col_str == 'AMOUNT':
                    column_mapping[col] = 'AMOUNT'
                elif 'BRANCH' in col_str:
                    column_mapping[col] = 'BRANCH'
                elif 'CURRENCY' in col_str:
                    column_mapping[col] = 'CURRENCY'
                elif 'DAYS' in col_str:
                    column_mapping[col] = 'DAYS_TO_MATURE'
                elif 'ISSUE' in col_str:
                    column_mapping[col] = 'ISSUE_DATE'
                elif 'EXPIRY' in col_str:
                    column_mapping[col] = 'EXPIRY_DATE'
            df_fallback = df_fallback.rename(columns=column_mapping)
            for c in ['ISSUE_DATE', 'EXPIRY_DATE']:
                if c in df_fallback.columns:
                    df_fallback[c] = pd.to_datetime(df_fallback[c], errors='coerce')
            return df_fallback
        except Exception as e2:
            st.error(f"Fallback also failed: {str(e2)}")
            return None

# ---------------- UI/Render helpers ----------------
def create_file_upload_section():
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown("### 📁 Upload Your Excel File")
    st.markdown("Upload **OUTSTANDING LGS_AS OF 2025.xlsx** (sheet: **LG BRANCH SUMMARY_2025**).")

    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx', 'xls'],
        help="Upload Excel file containing LG data"
    )

    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        file_details = {"Filename": uploaded_file.name,
                        "File size": f"{uploaded_file.size} bytes",
                        "File type": uploaded_file.type}
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Filename", file_details["Filename"])
        with c2: st.metric("Size", file_details["File size"])
        with c3: st.metric("Type", file_details["File type"])

    st.markdown('</div>', unsafe_allow_html=True)
    return uploaded_file

def create_summary_metrics(df):
    st.markdown('<div class="section-header">📊 Key Performance Indicators</div>', unsafe_allow_html=True)

    total_fac = df['TOTAL_FACILITIES'].sum() if 'TOTAL_FACILITIES' in df.columns else None
    total_used = df['AMOUNT_UTILIZED'].sum()  if 'AMOUNT_UTILIZED'  in df.columns else None
    total_outs = df['OUTSTANDING'].sum()      if 'OUTSTANDING'      in df.columns else None

    util_pct  = (total_used/total_fac*100) if (total_fac and total_used is not None and total_fac != 0) else None
    avail_pct = (total_outs/total_fac*100) if (total_fac and total_outs is not None and total_fac != 0) else None

    if 'UTILIZATION_RATE' in df.columns and df['UTILIZATION_RATE'].notna().any():
        avg_util = df['UTILIZATION_RATE'].mean()
    elif util_pct is not None:
        avg_util = util_pct
    else:
        avg_util = None

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Total Facilities", f"SAR {total_fac:,.0f}" if total_fac is not None else "—")
    with c2:
        st.metric("Amount Utilized",
                  f"SAR {total_used:,.0f}" if total_used is not None else "—",
                  delta=(f"{util_pct:.1f}% of total" if util_pct is not None else None))
    with c3:
        st.metric("Outstanding Amount",
                  f"SAR {total_outs:,.0f}" if total_outs is not None else "—",
                  delta=(f"{avail_pct:.1f}% available" if avail_pct is not None else None))
    with c4:
        st.metric("Avg Utilization Rate", f"{avg_util:.1f}%" if avg_util is not None else "—")

def charts_for_subset(df_subset, title_prefix):
    """Two charts: Amount by Bank (bar) and Count by Bank (pie)."""
    col1, col2 = st.columns(2)
    with col1:
        if 'BANK' in df_subset.columns and 'AMOUNT' in df_subset.columns:
            bank_amounts = df_subset.groupby('BANK')['AMOUNT'].sum().sort_values(ascending=False)
            fig_bank = px.bar(
                x=bank_amounts.index,
                y=bank_amounts.values,
                title=f'{title_prefix} - Amount by Bank',
                color=bank_amounts.values,
                color_continuous_scale='Blues'
            )
            fig_bank.update_layout(height=400, title_x=0.5, xaxis_title="Bank", yaxis_title="Total Amount (SAR)", showlegend=False)
            st.plotly_chart(fig_bank, use_container_width=True)
        else:
            st.info("Need BANK and AMOUNT columns for bar chart.")
    with col2:
        if 'BANK' in df_subset.columns:
            bank_counts = df_subset.groupby('BANK').size().sort_values(ascending=False)
            fig_count = px.pie(values=bank_counts.values, names=bank_counts.index, title=f'{title_prefix} - Count by Bank')
            fig_count.update_layout(height=400, title_x=0.5)
            st.plotly_chart(fig_count, use_container_width=True)
        else:
            st.info("Need BANK column for pie chart.")

def render_summary_and_detailed_tables(df_subset, summary_by='BANK', key_prefix='main'):
    """Two inner tabs: Summary (grouped stats) and Detailed (rows)."""
    tab_s, tab_d = st.tabs(["📈 Summary", "📋 Detailed"])

    with tab_s:
        if df_subset.empty:
            st.warning("No data.")
        else:
            group_cols = []
            if summary_by in df_subset.columns:
                group_cols = [summary_by]
            elif 'BANK' in df_subset.columns:
                group_cols = ['BANK']

            if group_cols:
                agg_dict = {}
                if 'AMOUNT' in df_subset.columns:
                    agg_dict['AMOUNT'] = ['count', 'sum', 'mean']
                if 'DAYS_TO_MATURE' in df_subset.columns:
                    agg_dict['DAYS_TO_MATURE'] = 'mean'

                if agg_dict:
                    summary_stats = df_subset.groupby(group_cols).agg(agg_dict).round(2)
                    # flatten columns
                    summary_stats.columns = [' '.join([c for c in col if c]).strip().title().replace('_', ' ')
                                             if isinstance(col, tuple) else str(col)
                                             for col in summary_stats.columns]
                    summary_stats = summary_stats.reset_index()
                    summary_stats = summary_stats.rename(columns={
                        'Amount Count': 'Count',
                        'Amount Sum': 'Total Amount',
                        'Amount Mean': 'Avg Amount',
                        'Days To Mature Mean': 'Avg Days to Mature'
                    })
                    st.dataframe(
                        summary_stats.style.format({
                            'Total Amount': '{:,.0f}',
                            'Avg Amount': '{:,.0f}',
                            'Avg Days to Mature': '{:.0f}'
                        }),
                        use_container_width=True
                    )
                else:
                    st.info("No numeric columns available to summarize.")
            else:
                st.info("No valid grouping column to summarize.")

    with tab_d:
        if df_subset.empty:
            st.warning("No rows to display.")
        else:
            df_display = _format_dates_for_display(df_subset, cols=('ISSUE_DATE', 'EXPIRY_DATE'))
            st.dataframe(df_display, use_container_width=True)

            # dd-mm-yyyy export
            export_df = df_subset.copy()
            for c in ('ISSUE_DATE', 'EXPIRY_DATE'):
                if c in export_df.columns:
                    export_df[c] = pd.to_datetime(cast := export_df[c], errors='coerce').dt.strftime('%d-%m-%Y')

            st.download_button(
                label="📥 Download (current tab) CSV",
                data=export_df.to_csv(index=False),
                file_name=f"lg_{summary_by.lower()}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                key=f"dl_{key_prefix}"
            )

def render_type_tab(df_filtered, gtype):
    st.subheader(f"📋 {gtype}")
    if 'GUARANTEE_TYPE' not in df_filtered.columns:
        st.warning("GUARANTEE_TYPE column missing.")
        return
    df_type = df_filtered[df_filtered['GUARANTEE_TYPE'] == gtype]
    if df_type.empty:
        st.warning(f"No data for **{gtype}**")
        return
    charts_for_subset(df_type, gtype)                                   # Charts on top
    render_summary_and_detailed_tables(df_type, 'BANK', f"type_{_std(gtype)}")  # Tables below

def create_maturity_analysis(df_detailed):
    st.markdown('<div class="section-header">⏰ Maturity Analysis</div>', unsafe_allow_html=True)

    def categorize_maturity(days):
        if days <= 30: return '≤ 30 days'
        elif days <= 90: return '31-90 days'
        elif days <= 180: return '91-180 days'
        else: return '> 180 days'

    if 'DAYS_TO_MATURE' not in df_detailed.columns:
        st.info("No DAYS_TO_MATURE column for maturity analysis.")
        return

    d = df_detailed.copy()
    d['MATURITY_CATEGORY'] = d['DAYS_TO_MATURE'].apply(categorize_maturity)

    col1, col2 = st.columns(2)
    with col1:
        maturity_dist = d['MATURITY_CATEGORY'].value_counts()
        fig_maturity = px.pie(values=maturity_dist.values, names=maturity_dist.index, title='LG Distribution by Time to Maturity')
        fig_maturity.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_maturity, use_container_width=True)
    with col2:
        if 'BANK' in d.columns:
            maturity_bank = d.groupby(['BANK', 'MATURITY_CATEGORY']).size().reset_index(name='count')
            fig_bank_maturity = px.bar(
                maturity_bank, x='BANK', y='count', color='MATURITY_CATEGORY',
                title='Maturity Distribution by Bank', barmode='stack'
            )
            fig_bank_maturity.update_layout(height=400, title_x=0.5)
            st.plotly_chart(fig_bank_maturity, use_container_width=True)
        else:
            st.info("Need BANK column to break down maturity by bank.")

# ---------------- Main ----------------
def main():
    st.markdown('<div class="main-header">🏦 LG Branch Summary Dashboard 2025</div>', unsafe_allow_html=True)

    # Reload button to clear caches and refresh
    st.sidebar.button("🔄 Reload data (clear cache)", on_click=st.cache_data.clear)

    uploaded_file = create_file_upload_section()
    if uploaded_file is None:
        st.warning("Please upload your Excel file to proceed. No fixed/sample data will be used.")
        st.stop()

    # Load data
    df_summary = load_summary_data_from_file(uploaded_file)
    df_detailed = load_detailed_data_from_file(uploaded_file)

    if df_summary is None or df_detailed is None or df_detailed.empty:
        st.error("No usable data found. Check sheet names and structure, then re-upload.")
        st.stop()

    # Sidebar filters
    st.sidebar.title("🔧 Dashboard Filters")
    st.sidebar.markdown("---")

    available_banks = df_detailed['BANK'].unique().tolist() if 'BANK' in df_detailed.columns else []
    selected_banks = st.sidebar.multiselect(
        "🏦 Select Banks", options=available_banks, default=available_banks, help="Filter banks to display"
    )

    st.sidebar.markdown("### 🏢 Branch Filter")
    available_branches = ['All']
    if 'BRANCH' in df_detailed.columns:
        available_branches += sorted([b for b in df_detailed['BRANCH'].dropna().unique().tolist()])
    selected_branch = st.sidebar.radio("Select Branch", options=available_branches, help="Filter by branch")

    st.sidebar.markdown("### 🏦 Bank Filter (Secondary)")
    available_banks_radio = ['All'] + available_banks
    selected_bank_radio = st.sidebar.radio("Select Specific Bank", options=available_banks_radio, help="Focus bank")

    # Apply filters to detailed
    df_detailed_filtered = df_detailed.copy()
    try:
        if selected_banks and 'BANK' in df_detailed_filtered.columns:
            df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BANK'].isin(selected_banks)]
        if selected_branch != 'All' and 'BRANCH' in df_detailed_filtered.columns:
            df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BRANCH'] == selected_branch]
        if selected_bank_radio != 'All' and 'BANK' in df_detailed_filtered.columns:
            df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BANK'] == selected_bank_radio]
    except Exception as e:
        st.warning(f"Error applying filters: {str(e)}. Using unfiltered data.")
        df_detailed_filtered = df_detailed.copy()

    # Filter summary to selected banks (handles BANKS or BANK)
    try:
        df_summary_cols = {c.upper(): c for c in df_summary.columns}
        bank_hdr = df_summary_cols.get("BANKS") or df_summary_cols.get("BANK")
        if selected_banks and bank_hdr:
            df_summary_filtered = df_summary[df_summary[bank_hdr].isin(selected_banks)]
        else:
            df_summary_filtered = df_summary.copy()
    except Exception:
        df_summary_filtered = df_summary.copy()

    # Sidebar info
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Current Filters")
    st.sidebar.write(f"**Banks:** {', '.join(selected_banks) if selected_banks else 'None'}")
    st.sidebar.write(f"**Branch:** {selected_branch}")
    st.sidebar.write(f"**Focus Bank:** {selected_bank_radio}")
    st.sidebar.write(f"**Records:** {len(df_detailed_filtered)}")
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**Last Updated:** {datetime.now().strftime('%d-%m-%Y')}")
    st.sidebar.markdown(f"**Data Source:** {uploaded_file.name}")

    # KPIs
    if not df_summary_filtered.empty:
        df_summary_filtered = _ensure_rates_only(df_summary_filtered)
        create_summary_metrics(df_summary_filtered)

    # Tabs by Guarantee Type (EVERY TAB: charts on top, summary & detailed below)
    if 'GUARANTEE_TYPE' in df_detailed_filtered.columns:
        guarantee_types = df_detailed_filtered['GUARANTEE_TYPE'].dropna().unique().tolist()
        if guarantee_types:
            st.markdown('<div class="section-header">📋 Analysis by Guarantee Type</div>', unsafe_allow_html=True)
            tab_names = guarantee_types + ["🔄 All Types", "📊 Summary Tables"]
            tabs = st.tabs(tab_names)

            # Each guarantee type tab
            for i, gtype in enumerate(guarantee_types):
                with tabs[i]:
                    render_type_tab(df_detailed_filtered, gtype)

            # All Types tab — charts + inner Summary/Detailed tables on the whole filtered set
            with tabs[len(guarantee_types)]:
                st.subheader("🔄 All Guarantee Types (Combined)")
                if not df_detailed_filtered.empty:
                    charts_for_subset(df_detailed_filtered, "All Types")
                    render_summary_and_detailed_tables(
                        df_detailed_filtered, summary_by='GUARANTEE_TYPE', key_prefix="all_types"
                    )
                    st.markdown(" ")
                    create_maturity_analysis(df_detailed_filtered)
                else:
                    st.warning("No data after filters.")

            # Summary Tables tab
            with tabs[len(guarantee_types) + 1]:
                st.subheader("📊 Data Tables")
                tab1, tab2 = st.tabs(["📈 Summary Data", "📋 Detailed Transactions"])
                with tab1:
                    st.subheader("Bank Summary")
                    if not df_summary_filtered.empty:
                        st.dataframe(df_summary_filtered, use_container_width=True)
                with tab2:
                    st.subheader("Transaction Details")
                    if not df_detailed_filtered.empty:
                        df_display = _format_dates_for_display(df_detailed_filtered, cols=('ISSUE_DATE', 'EXPIRY_DATE'))
                        st.dataframe(df_display, use_container_width=True)
                        export_df = df_detailed_filtered.copy()
                        for c in ('ISSUE_DATE', 'EXPIRY_DATE'):
                            if c in export_df.columns:
                                export_df[c] = pd.to_datetime(export_df[c], errors='coerce').dt.strftime('%d-%m-%Y')
                        st.download_button(
                            label="📥 Download Filtered Data as CSV",
                            data=export_df.to_csv(index=False),
                            file_name=f"lg_data_filtered_{datetime.now().strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                            use_container_width=True,
                            key="dl_filtered_main"
                        )
        else:
            st.warning("⚠️ No guarantee types found in the data.")
    else:
        st.warning("⚠️ GUARANTEE_TYPE column not found in the data. Please check your Excel file structure.")
        st.write("Available columns:", df_detailed_filtered.columns.tolist())

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "LG Branch Summary Dashboard • Built with Streamlit • Data as of "
        f"{datetime.now().strftime('%d-%m-%Y')}"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
