# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from datetime import datetime

# ========= Source URL (Google Sheets -> Published as XLSX) =========
SOURCE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS0uwaNWlivxOLwohf6kCSAkkfGTUpw5fnzwhGpoXIbymZaC8_QaHa-3ZaYz-gYEw/pub?output=xlsx"

# ========= Page configuration =========
st.set_page_config(
    page_title="LG Branch Summary Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ========= Minimal, clean CSS =========
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1f4e79;
        text-align: center;
        margin: 0.5rem 0 1rem 0;
        padding: .8rem 1rem;
        background: linear-gradient(90deg, #f0f8ff, #e6f3ff);
        border-radius: 12px;
        border-left: 6px solid #1f4e79;
    }
    .section-header {
        color: #1f4e79;
        font-size: 1.25rem;
        font-weight: 700;
        margin: 1.2rem 0 .6rem 0;
        padding-bottom: .4rem;
        border-bottom: 2px solid #e6f3ff;
    }
</style>
""", unsafe_allow_html=True)

# ========= Helpers =========
def _std(s: str) -> str:
    """Normalize header to UPPER_SNAKE."""
    return str(s).strip().upper().replace(" ", "_").replace("-", "_")

def _clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Drop empty rows/cols; if headers look 'UNNAMED', promote the best row as header."""
    d = df.copy()
    d = d.dropna(how='all').dropna(axis=1, how='all')
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

def _style_table(df: pd.DataFrame):
    """
    Format:
      - Money-ish columns -> commas, no decimals
      - DAYS / COUNT columns -> integers, no decimals
    """
    if df is None or df.empty:
        return df
    fmt = {}
    for c in df.columns:
        cu = str(c).upper()
        if any(tok in cu for tok in ["AMOUNT", "TOTAL", "OUTSTANDING", "FACILIT"]):
            if pd.api.types.is_numeric_dtype(df[c]):
                fmt[c] = "{:,.0f}"
        if "DAYS" in cu or cu == "DAYS_TO_MATURE" or cu == "COUNT":
            if pd.api.types.is_numeric_dtype(df[c]):
                fmt[c] = "{:,.0f}"
    return df.style.format(fmt)

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

# ========= Data loaders (from SOURCE_URL) =========
@st.cache_data
def load_summary_data_from_source(source_url: str):
    """
    Robustly extract the BANKS summary block from the 'Summary' sheet:
    headers anywhere like [BANKS | TOTAL FACILITIES | AMOUNT UTILIZED | OUTSTANDING].
    """
    wanted = ["BANKS", "TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"]

    def _find_summary_block(df_raw: pd.DataFrame):
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

    # Read raw (no header) and detect the block
    raw = pd.read_excel(source_url, sheet_name="Summary", header=None)
    start_row, cmap = _find_summary_block(raw)
    if start_row is not None:
        data = raw.iloc[start_row + 1:].copy()
        ordered = [c for c in ["BANKS", "TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"] if c in cmap]
        use_cols = [cmap[c] for c in ordered]
        d = data[use_cols].copy()
        d.columns = ordered

        # drop totals & blanks (fixes double counting)
        d = d[d["BANKS"].notna()]
        d["BANKS"] = d["BANKS"].astype(str).str.strip()
        d = d[~d["BANKS"].str.upper().isin(["", "TOTAL", "TOTALS", "SUM", "GRAND TOTAL"])]

        # numeric coercions
        for c in ["TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"]:
            if c in d.columns:
                d[c] = pd.to_numeric(d[c], errors="coerce")

        d = d.rename(columns={
            "TOTAL FACILITIES": "TOTAL_FACILITIES",
            "AMOUNT UTILIZED": "AMOUNT_UTILIZED",
            "OUTSTANDING": "OUTSTANDING",
        })
        d = _ensure_rates_only(d)
        if "BANKS" in d.columns and "AMOUNT_UTILIZED" in d.columns:
            return d

    # Fallback: build minimal summary from detailed
    df_det = pd.read_excel(source_url, sheet_name="LG BRANCH SUMMARY_2025", usecols="A:K")
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

@st.cache_data
def load_detailed_data_from_source(source_url: str):
    """Load detailed data (LG BRANCH SUMMARY_2025, A:K), robust to messy sheets."""
    try:
        df = pd.read_excel(source_url, sheet_name='LG BRANCH SUMMARY_2025', usecols="A:K")
        df = _clean_frame(df)

        expected_columns = {
            0: 'BANK', 1: 'LG_REF', 2: 'CUSTOMER_NAME', 3: 'GUARANTEE_TYPE',
            4: 'ISSUE_DATE', 5: 'EXPIRY_DATE', 6: 'AMOUNT', 7: 'CURRENCY',
            8: 'BRANCH', 9: 'BANK_2', 10: 'DAYS_TO_MATURE'
        }
        df.columns = [expected_columns.get(i, f'Column_{i}') for i, _ in enumerate(df.columns)]
        df = df.drop('BANK_2', axis=1, errors='ignore')

        df = df[df['BANK'].notna()]
        df = df[df['BANK'] != 'BANK']

        if 'GUARANTEE_TYPE' in df.columns:
            df['GUARANTEE_TYPE'] = df['GUARANTEE_TYPE'].astype(str).str.strip()
        if 'AMOUNT' in df.columns:
            df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce').fillna(0)
        if 'DAYS_TO_MATURE' in df.columns:
            df['DAYS_TO_MATURE'] = pd.to_numeric(df['DAYS_TO_MATURE'], errors='coerce').fillna(0)
        for date_col in ['ISSUE_DATE', 'EXPIRY_DATE']:
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        if 'LG_REF' in df.columns: df['LG_REF'] = df['LG_REF'].fillna('N/A')
        if 'CUSTOMER_NAME' in df.columns: df['CUSTOMER_NAME'] = df['CUSTOMER_NAME'].fillna('Unknown Customer')
        if 'BRANCH' in df.columns: df['BRANCH'] = df['BRANCH'].fillna('Main Branch')
        if 'CURRENCY' in df.columns: df['CURRENCY'] = df['CURRENCY'].fillna('SAR')

        # keep essential rows only
        df = df.dropna(subset=['BANK', 'GUARANTEE_TYPE'])
        df = df[df['BANK'].astype(str).str.strip() != '']
        df = df[df['GUARANTEE_TYPE'].astype(str).str.strip() != '']
        return df

    except Exception as e:
        st.error(f"Error loading detailed data: {str(e)}")
        try:
            df_fallback = pd.read_excel(source_url, sheet_name='LG BRANCH SUMMARY_2025')
            df_fallback = _clean_frame(df_fallback)
            column_mapping = {}
            for col in df_fallback.columns:
                col_str = _std(col)
                if col_str == 'BANK': column_mapping[col] = 'BANK'
                elif 'LG_REF' in col_str or col_str == 'LG_REF': column_mapping[col] = 'LG_REF'
                elif 'CUSTOMER' in col_str: column_mapping[col] = 'CUSTOMER_NAME'
                elif 'GUARR' in col_str or 'TYPE' in col_str: column_mapping[col] = 'GUARANTEE_TYPE'
                elif col_str == 'AMOUNT': column_mapping[col] = 'AMOUNT'
                elif 'BRANCH' in col_str: column_mapping[col] = 'BRANCH'
                elif 'CURRENCY' in col_str: column_mapping[col] = 'CURRENCY'
                elif 'DAYS' in col_str: column_mapping[col] = 'DAYS_TO_MATURE'
                elif 'ISSUE' in col_str: column_mapping[col] = 'ISSUE_DATE'
                elif 'EXPIRY' in col_str: column_mapping[col] = 'EXPIRY_DATE'
            df_fallback = df_fallback.rename(columns=column_mapping)
            for c in ['ISSUE_DATE', 'EXPIRY_DATE']:
                if c in df_fallback.columns:
                    df_fallback[c] = pd.to_datetime(df_fallback[c], errors='coerce')
            return df_fallback
        except Exception as e2:
            st.error(f"Fallback also failed: {str(e2)}")
            return None

# ========= UI helpers =========
def create_source_section():
    """Top-right refresh button (no source text shown)."""
    cols = st.columns([8, 1])
    with cols[1]:
        if st.button("🔄 Refresh data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

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
            fig_bank.update_yaxes(tickformat=",.0f")
            fig_bank.update_traces(hovertemplate="<b>%{x}</b><br>Total Amount: SAR %{y:,.0f}<extra></extra>")
            st.plotly_chart(fig_bank, use_container_width=True)
        else:
            st.info("Need BANK and AMOUNT columns for bar chart.")
    with col2:
        if 'BANK' in df_subset.columns:
            bank_counts = df_subset.groupby('BANK').size().sort_values(ascending=False)
            fig_count = px.pie(values=bank_counts.values, names=bank_counts.index, title=f'{title_prefix} - Count by Bank')
            fig_count.update_layout(height=400, title_x=0.5)
            fig_count.update_traces(hovertemplate="<b>%{label}</b><br>Count: %{value:,.0f}<extra></extra>")
            st.plotly_chart(fig_count, use_container_width=True)
        else:
            st.info("Need BANK column for pie chart.")

def render_summary_and_detailed_tables(df_subset, summary_by='BANK', key_prefix='main'):
    tab_s, tab_d = st.tabs(["📈 Summary", "📋 Detailed"])

    with tab_s:
        if df_subset.empty:
            st.warning("No data.")
        else:
            group_cols = [summary_by] if summary_by in df_subset.columns else (['BANK'] if 'BANK' in df_subset.columns else [])
            if group_cols:
                agg_dict = {}
                if 'AMOUNT' in df_subset.columns:
                    agg_dict['AMOUNT'] = ['count', 'sum']  # no averages
                if agg_dict:
                    summary_stats = df_subset.groupby(group_cols).agg(agg_dict).round(2)
                    summary_stats.columns = [
                        ' '.join([c for c in col if c]).strip().title().replace('_', ' ')
                        if isinstance(col, tuple) else str(col)
                        for col in summary_stats.columns
                    ]
                    summary_stats = summary_stats.reset_index().rename(columns={
                        'Amount Count': 'Count',
                        'Amount Sum': 'Total Amount',
                    })
                    st.dataframe(_style_table(summary_stats), use_container_width=True)
                else:
                    st.info("No numeric columns available to summarize.")
            else:
                st.info("No valid grouping column to summarize.")

    with tab_d:
        if df_subset.empty:
            st.warning("No rows to display.")
        else:
            df_display = _format_dates_for_display(df_subset, cols=('ISSUE_DATE', 'EXPIRY_DATE'))
            st.dataframe(_style_table(df_display), use_container_width=True)

            export_df = df_subset.copy()
            for c in ('ISSUE_DATE', 'EXPIRY_DATE'):
                if c in export_df.columns:
                    export_df[c] = pd.to_datetime(export_df[c], errors='coerce').dt.strftime('%d-%m-%Y')

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
    # 👉 TABLES FIRST, then charts
    render_summary_and_detailed_tables(df_type, 'BANK', f"type_{_std(gtype)}")
    charts_for_subset(df_type, gtype)

def create_maturity_analysis(df_detailed):
    st.markdown('<div class="section-header">⏰ Maturity Analysis</div>', unsafe_allow_html=True)
    def categorize_maturity(days):
        if days <= 30: return '≤ 30 days'
        elif days <= 90: return '31-90 days'
        elif days <= 180: return '91-180 days'
        else: return '> 180 days'
    if 'DAYS_TO_MATURE' not in df_detailed.columns:
        st.info("No DAYS_TO_MATURE column for maturity analysis."); return
    d = df_detailed.copy()
    d['MATURITY_CATEGORY'] = d['DAYS_TO_MATURE'].apply(categorize_maturity)
    col1, col2 = st.columns(2)
    with col1:
        maturity_dist = d['MATURITY_CATEGORY'].value_counts()
        fig_maturity = px.pie(values=maturity_dist.values, names=maturity_dist.index, title='LG Distribution by Time to Maturity')
        fig_maturity.update_layout(height=400, title_x=0.5)
        fig_maturity.update_traces(hovertemplate="<b>%{label}</b><br>Count: %{value:,.0f}<extra></extra>")
        st.plotly_chart(fig_maturity, use_container_width=True)
    with col2:
        if 'BANK' in d.columns:
            maturity_bank = d.groupby(['BANK', 'MATURITY_CATEGORY']).size().reset_index(name='count')
            fig_bank_maturity = px.bar(
                maturity_bank, x='BANK', y='count', color='MATURITY_CATEGORY',
                title='Maturity Distribution by Bank', barmode='stack'
            )
            fig_bank_maturity.update_layout(height=400, title_x=0.5)
            fig_bank_maturity.update_yaxes(tickformat=",.0f")
            fig_bank_maturity.update_traces(hovertemplate="<b>%{x}</b><br>Count: %{y:,.0f}<extra></extra>")
            st.plotly_chart(fig_bank_maturity, use_container_width=True)

def render_current_month_maturity(df_all):
    """Tab: current month maturities (tables first, then charts)."""
    st.subheader("📅 Current Month Maturity")
    if 'EXPIRY_DATE' not in df_all.columns:
        st.info("No EXPIRY_DATE column found.")
        return
    today = pd.Timestamp.today().normalize()
    start = today.replace(day=1)
    end = (start + pd.offsets.MonthEnd(1))  # last day this month
    d = df_all.copy()
    d = d[(pd.to_datetime(d['EXPIRY_DATE'], errors='coerce') >= start) &
          (pd.to_datetime(d['EXPIRY_DATE'], errors='coerce') <= end)]
    if d.empty:
        st.success("✅ No LGs are maturing this month.")
        return

    # Tables first
    render_summary_and_detailed_tables(d, summary_by='BANK', key_prefix="curr_month")
    # Then charts
    charts_for_subset(d, f"Maturity in {start.strftime('%b %Y')}")

# ========= Main =========
def main():
    st.markdown('<div class="main-header">🏦 LG Branch Summary Dashboard 2025</div>', unsafe_allow_html=True)

    # --- Top-right refresh button (no source text) ---
    create_source_section()

    # Load data from SOURCE_URL
    df_summary = load_summary_data_from_source(SOURCE_URL)
    df_detailed = load_detailed_data_from_source(SOURCE_URL)
    if df_summary is None or df_detailed is None or df_detailed.empty:
        st.error("No usable data found. Check the source sheet names/structure.")
        st.stop()

    # --- Simple filters (RADIO): Branch + Bank ---
    st.markdown('<div class="section-header">🔧 Filters</div>', unsafe_allow_html=True)
    cols = st.columns(2)
    with cols[0]:
        branches = ['All']
        if 'BRANCH' in df_detailed.columns:
            branches += sorted([b for b in df_detailed['BRANCH'].dropna().unique().tolist()])
        selected_branch = st.radio("Branch", options=branches, horizontal=True)
    with cols[1]:
        banks_radio = ['All']
        if 'BANK' in df_detailed.columns:
            banks_radio += sorted(df_detailed['BANK'].dropna().unique().tolist())
        selected_bank_radio = st.radio("Bank", options=banks_radio, horizontal=True)

    # Apply filters (detailed)
    df_detailed_filtered = df_detailed.copy()
    if selected_branch != 'All' and 'BRANCH' in df_detailed_filtered.columns:
        df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BRANCH'] == selected_branch]
    if selected_bank_radio != 'All' and 'BANK' in df_detailed_filtered.columns:
        df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BANK'] == selected_bank_radio]

    # Filter summary (if a bank is chosen)
    try:
        df_summary_cols = {c.upper(): c for c in df_summary.columns}
        bank_hdr = df_summary_cols.get("BANKS") or df_summary_cols.get("BANK")
        if selected_bank_radio != 'All' and bank_hdr:
            df_summary_filtered = df_summary[df_summary[bank_hdr] == selected_bank_radio]
        else:
            df_summary_filtered = df_summary.copy()
    except Exception:
        df_summary_filtered = df_summary.copy()

    # KPIs block
    if not df_summary_filtered.empty:
        df_summary_filtered = _ensure_rates_only(df_summary_filtered)
        create_summary_metrics(df_summary_filtered)

    # Tabs by Guarantee Type (+ Current Month Maturity)
    if 'GUARANTEE_TYPE' in df_detailed_filtered.columns:
        guarantee_types = df_detailed_filtered['GUARANTEE_TYPE'].dropna().unique().tolist()
        if guarantee_types:
            st.markdown('<div class="section-header">📋 Analysis by Guarantee Type</div>', unsafe_allow_html=True)
            tabs = st.tabs(["📅 Current Month Maturity"] + guarantee_types + ["🔄 All Types", "📊 Summary Tables"])

            # Current Month tab
            with tabs[0]:
                render_current_month_maturity(df_detailed_filtered)

            # Each guarantee type tab (TABLES FIRST, then CHARTS)
            for idx, gtype in enumerate(guarantee_types, start=1):
                with tabs[idx]:
                    render_type_tab(df_detailed_filtered, gtype)

            # All Types tab — TABLES first, then charts + maturity analysis
            with tabs[len(guarantee_types) + 1]:
                st.subheader("🔄 All Guarantee Types (Combined)")
                render_summary_and_detailed_tables(
                    df_detailed_filtered, summary_by='GUARANTEE_TYPE', key_prefix="all_types"
                )
                charts_for_subset(df_detailed_filtered, "All Types")
                create_maturity_analysis(df_detailed_filtered)

            # Summary Tables tab
            with tabs[len(guarantee_types) + 2]:
                st.subheader("📊 Data Tables")
                tab1, tab2 = st.tabs(["📈 Summary Data", "📋 Detailed Transactions"])
                with tab1:
                    st.subheader("Bank Summary")
                    st.dataframe(_style_table(df_summary_filtered), use_container_width=True)
                with tab2:
                    st.subheader("Transaction Details")
                    df_display = _format_dates_for_display(df_detailed_filtered, cols=('ISSUE_DATE', 'EXPIRY_DATE'))
                    st.dataframe(_style_table(df_display), use_container_width=True)
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

    # Footer (discreet)
    st.markdown("---")
    st.markdown(
        f"<div style='text-align:center;color:#666;'>LG Branch Summary Dashboard ,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()

