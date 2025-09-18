# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from datetime import datetime
from difflib import get_close_matches
import re

# ========= Source URL (Google Sheets -> Published as XLSX) =========
SOURCE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS0uwaNWlivxOLwohf6kCSAkkfGTUpw5fnzwhGpoXIbymZaC8_QaHa-3ZaYz-gYEw/pub?output=xlsx"

# ========= Page configuration =========
st.set_page_config(
    page_title="LETTER OF GUARANTEE",
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
    .pill {
        display:inline-block; padding:4px 10px; border-radius:999px; font-size:.8rem;
        background:#eef5ff; color:#164e8a; border:1px solid #d8e7ff; margin-left:8px;
    }
</style>
""", unsafe_allow_html=True)

# ========= Helpers =========
def _std(s: str) -> str:
    return str(s).strip().upper().replace(" ", "_").replace("-", "_")

def _clean_frame(df: pd.DataFrame) -> pd.DataFrame:
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
    d = df_sum.copy()
    if {"AMOUNT_UTILIZED", "TOTAL_FACILITIES"}.issubset(d.columns):
        with np.errstate(divide='ignore', invalid='ignore'):
            d["UTILIZATION_RATE"] = (d["AMOUNT_UTILIZED"] / d["TOTAL_FACILITIES"] * 100).round(2)
    if {"OUTSTANDING", "TOTAL_FACILITIES"}.issubset(d.columns):
        with np.errstate(divide='ignore', invalid='ignore'):
            d["OUTSTANDING_RATE"] = (d["OUTSTANDING"] / d["TOTAL_FACILITIES"] * 100).round(2)
    return d

# ========= Data loaders from SOURCE_URL =========
@st.cache_data(ttl=180)
def load_summary_data_from_source(source_url: str):
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

    raw = pd.read_excel(source_url, sheet_name="Summary", header=None)
    start_row, cmap = _find_summary_block(raw)
    if start_row is not None:
        data = raw.iloc[start_row + 1:].copy()
        ordered = [c for c in ["BANKS", "TOTAL FACILITIES", "AMOUNT UTILIZED", "OUTSTANDING"] if c in cmap]
        use_cols = [cmap[c] for c in ordered]
        d = data[use_cols].copy()
        d.columns = ordered

        d = d[d["BANKS"].notna()]
        d["BANKS"] = d["BANKS"].astype(str).str.strip()
        d = d[~d["BANKS"].str.upper().isin(["", "TOTAL", "TOTALS", "SUM", "GRAND TOTAL"])]

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

@st.cache_data(ttl=180)
def load_detailed_data_from_source(source_url: str):
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

        # Clean text columns
        for c in ['BANK', 'CUSTOMER_NAME', 'GUARANTEE_TYPE', 'CURRENCY', 'BRANCH', 'LG_REF']:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip()

        if 'AMOUNT' in df.columns:
            df['AMOUNT'] = (
                df['AMOUNT'].astype(str)
                .str.replace(",", "", regex=False)
                .str.replace(" ", "", regex=False)
            )
            df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce').fillna(0)

        if 'DAYS_TO_MATURE' in df.columns:
            df['DAYS_TO_MATURE'] = pd.to_numeric(df['DAYS_TO_MATURE'], errors='coerce').fillna(0)

        for date_col in ['ISSUE_DATE', 'EXPIRY_DATE']:
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        # Fill defaults
        df['LG_REF'] = df.get('LG_REF', pd.Series(dtype=str)).fillna('N/A')
        df['CUSTOMER_NAME'] = df.get('CUSTOMER_NAME', pd.Series(dtype=str)).fillna('Unknown Customer')
        df['BRANCH'] = df.get('BRANCH', pd.Series(dtype=str)).fillna('Main Branch')
        df['CURRENCY'] = df.get('CURRENCY', pd.Series(dtype=str)).fillna('SAR')

        # Drop empties
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

@st.cache_data(ttl=180)
def load_branch_utilized_from_summary(source_url: str) -> pd.DataFrame:
    """
    Read the Summary sheet's two-column table [BRANCH | AMOUNT] and
    return columns: BRANCH, AMOUNT_UTILIZED.
    """
    raw = pd.read_excel(source_url, sheet_name="Summary", header=None)
    found = None
    for r in range(len(raw)):
        row = raw.iloc[r].tolist()
        norm = { _std(v): i for i, v in enumerate(row) if pd.notna(v) and str(v).strip() != "" }
        amt_key = next((k for k in norm.keys() if k.startswith("AMOUNT")), None)
        if "BRANCH" in norm and amt_key:
            found = (r, norm["BRANCH"], norm[amt_key]); break
    if not found:
        return pd.DataFrame(columns=["BRANCH", "AMOUNT_UTILIZED"])
    r, c_br, c_amt = found
    t = raw.iloc[r+1:, [c_br, c_amt]].copy()
    t.columns = ["BRANCH", "AMOUNT"]
    t["BRANCH"] = t["BRANCH"].astype(str).str.strip()
    t = t[t["BRANCH"].notna()]
    t = t[t["BRANCH"] != ""]
    t = t[~t["BRANCH"].str.upper().isin(["TOTAL", "TOTALS", "GRAND TOTAL", "SUM"])]
    t["AMOUNT"] = t["AMOUNT"].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
    t["AMOUNT"] = t["AMOUNT"].replace({"-": 0, "": 0, "nan": 0, "None": 0})
    t["AMOUNT"] = pd.to_numeric(t["AMOUNT"], errors="coerce").fillna(0)
    t = t.rename(columns={"AMOUNT": "AMOUNT_UTILIZED"}).reset_index(drop=True)
    return t

# ========= UI helpers =========
def create_source_section():
    cols = st.columns([7, 2, 1, 1])
    with cols[1]:
        if st.button("🔄 Refresh data", use_container_width=True):
            st.cache_data.clear()
            st.session_state["_force_refresh"] = True
            st.rerun()
    with cols[2]:
        if st.button("🧹 Clear chat", use_container_width=True):
            st.session_state.pop("chat", None)
            st.session_state.pop("_last_filters", None)
            st.rerun()
    with cols[3]:
        st.write("")

def create_summary_metrics(df):
    st.markdown('<div class="section-header">📊 Key Performance Indicators</div>', unsafe_allow_html=True)
    total_fac = df['TOTAL_FACILITIES'].sum() if 'TOTAL_FACILITIES' in df.columns else None
    total_used = df['AMOUNT_UTILIZED'].sum()  if 'AMOUNT_UTILIZED'  in df.columns else None
    total_outs = df['OUTSTANDING'].sum()      if 'OUTSTANDING'      in df.columns else None
    util_pct  = (total_used/total_fac*100) if (total_fac is not None and total_fac != 0 and total_used is not None) else None
    avail_pct = (total_outs/total_fac*100) if (total_fac is not None and total_fac != 0 and total_outs is not None) else None
    if 'UTILIZATION_RATE' in df.columns and df['UTILIZATION_RATE'].notna().any():
        avg_util = df['UTILIZATION_RATE'].mean()
    elif util_pct is not None:
        avg_util = util_pct
    else:
        avg_util = None
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Total Facilities", f"SAR {total_fac:,.0f}" if total_fac is not None else "—")
    with c2: st.metric("Amount Utilized", f"SAR {total_used:,.0f}" if total_used is not None else "—",
                       delta=(f"{util_pct:.1f}% of total" if util_pct is not None else None))
    with c3: st.metric("Outstanding Amount", f"SAR {total_outs:,.0f}" if total_outs is not None else "—",
                       delta=(f"{avail_pct:.1f}% available" if avail_pct is not None else None))
    with c4: st.metric("Avg Utilization Rate", f"{avg_util:.1f}%" if avg_util is not None else "—")

def charts_for_subset(df_subset, title_prefix):
    col1, col2 = st.columns(2)
    with col1:
        if 'BANK' in df_subset.columns and 'AMOUNT' in df_subset.columns:
            bank_amounts = df_subset.groupby('BANK', dropna=True)['AMOUNT'].sum().sort_values(ascending=False)
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
            bank_counts = df_subset.groupby('BANK', dropna=True).size().sort_values(ascending=False)
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
                    agg_dict['AMOUNT'] = ['count', 'sum']
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
    st.subheader("📅 Current Month Maturity")
    if 'EXPIRY_DATE' not in df_all.columns:
        st.info("No EXPIRY_DATE column found.")
        return
    today = pd.Timestamp.today().normalize()
    start = today.replace(day=1)
    end = (start + pd.offsets.MonthEnd(1))
    d = df_all.copy()
    d = d[(pd.to_datetime(d['EXPIRY_DATE'], errors='coerce') >= start) &
          (pd.to_datetime(d['EXPIRY_DATE'], errors='coerce') <= end)]
    if d.empty:
        st.success("✅ No LGs are maturing this month.")
        return
    render_summary_and_detailed_tables(d, summary_by='BANK', key_prefix="curr_month")
    charts_for_subset(d, f"Maturity in {start.strftime('%b %Y')}")

# ========= Chatbot helpers (rule-based, on-data) =========
def _norm_text(s): return str(s).strip().upper()

def _fuzzy_find(name, choices, cutoff=0.7):
    if not name or not choices:
        return None
    # exact (case-insensitive)
    for c in choices:
        if _norm_text(c) == _norm_text(name):
            return c
    m = get_close_matches(name, choices, n=1, cutoff=cutoff)
    return m[0] if m else None

def _parse_days_window(q):
    ql = q.lower()
    m = re.search(r'(?:next|coming|within|in)\s+(\d{1,3})\s*days', ql)
    return int(m.group(1)) if m else None

def _this_month_mask(series_dt):
    today = pd.Timestamp.today().normalize()
    start = today.replace(day=1)
    end = (start + pd.offsets.MonthEnd(1))
    s = pd.to_datetime(series_dt, errors='coerce')
    return (s >= start) & (s <= end)

def _next_n_days_mask(series_dt, n):
    today = pd.Timestamp.today().normalize()
    end = today + pd.Timedelta(days=n)
    s = pd.to_datetime(series_dt, errors='coerce')
    return (s >= today) & (s <= end)

def _apply_scope(df_summary, df_detailed, df_branch_util, scope, selected_branch, selected_bank):
    """Return context dataframes according to chat scope."""
    if scope == "Current filters":
        d = df_detailed.copy()
        if selected_branch and selected_branch != "All" and 'BRANCH' in d.columns:
            d = d[d['BRANCH'] == selected_branch]
        if selected_bank and selected_bank != "All" and 'BANK' in d.columns:
            d = d[d['BANK'] == selected_bank]
        # Filter summary by bank if selected
        s = df_summary.copy() if df_summary is not None else None
        if s is not None:
            cols_map = {c.upper(): c for c in s.columns}
            bank_hdr = cols_map.get("BANKS") or cols_map.get("BANK")
            if bank_hdr and selected_bank and selected_bank != "All":
                s = s[s[bank_hdr] == selected_bank]
        return s, d, df_branch_util
    else:
        return df_summary, df_detailed, df_branch_util

def answer_report(question, df_summary, df_detailed, df_branch_util):
    """
    Returns (answer_text, optional_dataframe or None)
    """
    q = question.strip()
    ql = q.lower()
    if df_detailed is None or df_detailed.empty:
        return ("I couldn't find detailed data loaded.", None)

    banks = sorted(df_detailed['BANK'].dropna().astype(str).unique().tolist()) if 'BANK' in df_detailed.columns else []
    branches = sorted(df_detailed['BRANCH'].dropna().astype(str).unique().tolist()) if 'BRANCH' in df_detailed.columns else []
    types = sorted(df_detailed['GUARANTEE_TYPE'].dropna().astype(str).unique().tolist()) if 'GUARANTEE_TYPE' in df_detailed.columns else []

    # 0) Direct LG_REF lookup
    m_ref = re.search(r'(?:LG|LG_REF|REF)\s*[:#]?\s*([A-Za-z0-9\-_/]+)', q, re.IGNORECASE)
    if m_ref and 'LG_REF' in df_detailed.columns:
        ref = m_ref.group(1)
        rows = df_detailed[df_detailed['LG_REF'].astype(str).str.contains(re.escape(ref), case=False, na=False)]
        if rows.empty:
            return (f"No records found for reference like '{ref}'.", None)
        total = rows['AMOUNT'].sum() if 'AMOUNT' in rows.columns else None
        ans = f"Found {len(rows)} record(s) for LG ref like **{ref}**."
        if total is not None:
            ans += f" Total Amount: **SAR {total:,.0f}**."
        return (ans, _format_dates_for_display(rows, cols=('ISSUE_DATE','EXPIRY_DATE')))

    # 1) Totals (Summary)
    if any(k in ql for k in ["total facilities", "amount utilized", "outstanding", "utilization rate"]):
        if df_summary is None or df_summary.empty:
            return ("Summary sheet not available to compute totals.", None)
        tot_fac = df_summary['TOTAL_FACILITIES'].sum() if 'TOTAL_FACILITIES' in df_summary.columns else None
        tot_used = df_summary['AMOUNT_UTILIZED'].sum() if 'AMOUNT_UTILIZED' in df_summary.columns else None
        tot_out = df_summary['OUTSTANDING'].sum() if 'OUTSTANDING' in df_summary.columns else None
        parts = []
        if tot_fac is not None: parts.append(f"**Total Facilities**: SAR {tot_fac:,.0f}")
        if tot_used is not None: parts.append(f"**Amount Utilized**: SAR {tot_used:,.0f}")
        if tot_out is not None: parts.append(f"**Outstanding**: SAR {tot_out:,.0f}")
        if (tot_fac is not None and tot_fac != 0 and tot_used is not None):
            parts.append(f"**Utilization**: {tot_used/tot_fac*100:,.1f}%")
        if not parts:
            return ("I couldn't compute top-level totals from the Summary sheet.", None)
        return ("; ".join(parts), None)

    # 2) By bank
    m_bank = None
    for tok in ["at ", "of ", "for ", "by "]:
        m = re.search(rf'(?:bank\s*{tok}|{tok})([A-Za-z0-9 &._-]+)', ql)
        if m:
            m_bank = m.group(1).strip()
            break
    bank_name = _fuzzy_find(m_bank, banks) if m_bank else None
    if ("bank" in ql or bank_name) and any(k in ql for k in ["amount", "utilized", "total", "count", "how many", "sum"]):
        d = df_detailed.copy()
        if bank_name:
            d = d[d['BANK'].astype(str) == bank_name]
            if d.empty:
                return (f"No data for bank **{bank_name}**.", None)
            total = d['AMOUNT'].sum() if 'AMOUNT' in d.columns else 0
            return (f"**{bank_name}** — Count: {len(d):,.0f}, Total Amount: **SAR {total:,.0f}**.", _format_dates_for_display(d))
        if 'AMOUNT' in df_detailed.columns and 'BANK' in df_detailed.columns:
            g = df_detailed.groupby('BANK', dropna=True)['AMOUNT'].agg(['count','sum']).reset_index().sort_values('sum', ascending=False)
            g = g.rename(columns={'count':'Count', 'sum':'Total Amount'})
            return ("Bank-wise totals (sorted by amount):", g)

    # 3) By branch
    m_branch = None
    for tok in ["branch ", "for branch ", "at branch "]:
        m = re.search(rf'{tok}([A-Za-z0-9 /._-]+)', ql)
        if m:
            m_branch = m.group(1).strip()
            break
    branch_name = _fuzzy_find(m_branch, branches) if m_branch else None
    if ("branch" in ql or branch_name):
        if "utilized" in ql or "amount" in ql or "total" in ql:
            # Try Summary sheet two-column Branch table first
            if df_branch_util is not None and not df_branch_util.empty:
                if branch_name:
                    row = df_branch_util[df_branch_util['BRANCH'].astype(str).str.upper() == branch_name.upper()]
                    if row.empty:
                        row = df_branch_util[df_branch_util['BRANCH'].astype(str).str.contains(branch_name, case=False, na=False)]
                    if not row.empty:
                        val = float(row['AMOUNT_UTILIZED'].iloc[0])
                        return (f"**{branch_name}** — Amount Utilized: **SAR {val:,.0f}** (Summary sheet).", None)
            # Fallback: compute from detailed
            d = df_detailed.copy()
            if branch_name:
                d = d[d['BRANCH'].astype(str).str.upper() == branch_name.upper()]
                if d.empty:
                    d = df_detailed[df_detailed['BRANCH'].astype(str).str.contains(branch_name, case=False, na=False)]
            if d.empty:
                return (f"No rows found for branch like **{m_branch or 'N/A'}**.", None)
            total = d['AMOUNT'].sum() if 'AMOUNT' in d.columns else 0
            return (f"**{branch_name or 'Selected'}** — Rows: {len(d):,.0f}, Total Amount: **SAR {total:,.0f}**.", _format_dates_for_display(d))

    # 4) By guarantee type
    m = re.search(r'(?:type|guarantee type)\s*([A-Za-z0-9 /._-]+)', ql)
    type_name = _fuzzy_find(m.group(1).strip(), types) if m else None
    if ("by type" in ql or "type " in ql or type_name) and any(k in ql for k in ["sum","amount","total","count"]):
        d = df_detailed.copy()
        if type_name:
            d = d[d['GUARANTEE_TYPE'].astype(str) == type_name]
            if d.empty:
                return (f"No data for guarantee type **{type_name}**.", None
