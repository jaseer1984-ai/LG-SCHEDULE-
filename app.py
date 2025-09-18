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
@st.cache_data
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

@st.cache_data
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
        # allow AMOUNT / AMOUNT_... headers
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
def _norm_text(s):
    return str(s).strip().upper()

def _fuzzy_find(name, choices, cutoff=0.65):
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

def answer_report(question, df_summary, df_detailed, df_branch_util):
    """
    Returns (answer_text, optional_dataframe or None)
    """
    q = question.strip()
    ql = q.lower()
    if df_detailed is None or df_detailed.empty:
        return ("I couldn't find detailed data loaded.", None)

    # vocab
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
        tot_fac = df_summary['TOTAL_FACILITIES'].sum() if (df_summary is not None and 'TOTAL_FACILITIES' in df_summary.columns) else None
        tot_used = df_summary['AMOUNT_UTILIZED'].sum() if (df_summary is not None and 'AMOUNT_UTILIZED' in df_summary.columns) else None
        tot_out = df_summary['OUTSTANDING'].sum() if (df_summary is not None and 'OUTSTANDING' in df_summary.columns) else None
        parts = []
        if tot_fac is not None: parts.append(f"**Total Facilities**: SAR {tot_fac:,.0f}")
        if tot_used is not None: parts.append(f"**Amount Utilized**: SAR {tot_used:,.0f}")
        if tot_out is not None: parts.append(f"**Outstanding**: SAR {tot_out:,.0f}")
        if tot_fac and tot_used is not None:
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
        # Summary across all banks
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
    m_type = None
    m = re.search(r'(?:type|guarantee type)\s*([A-Za-z0-9 /._-]+)', ql)
    if m:
        m_type = m.group(1).strip()
    type_name = _fuzzy_find(m_type, types) if m_type else None
    if ("by type" in ql or "type " in ql or type_name) and any(k in ql for k in ["sum","amount","total","count"]):
        d = df_detailed.copy()
        if type_name:
            d = d[d['GUARANTEE_TYPE'].astype(str) == type_name]
            if d.empty:
                return (f"No data for guarantee type **{type_name}**.", None)
            total = d['AMOUNT'].sum() if 'AMOUNT' in d.columns else 0
            return (f"**{type_name}** — Count: {len(d):,.0f}, Total Amount: **SAR {total:,.0f}**.", _format_dates_for_display(d))
        g = df_detailed.groupby('GUARANTEE_TYPE')['AMOUNT'].agg(['count','sum']).reset_index().sort_values('sum', ascending=False)
        g = g.rename(columns={'count':'Count','sum':'Total Amount'})
        return ("Totals by Guarantee Type:", g)

    # 5) Maturities
    if 'EXPIRY_DATE' in df_detailed.columns and any(k in ql for k in ["mature", "maturity", "expire", "expiry", "due"]):
        n = _parse_days_window(ql)
        d = df_detailed.copy()
        if "this month" in ql or "current month" in ql:
            mask = _this_month_mask(d['EXPIRY_DATE'])
            scope = "this month"
        elif isinstance(n, int):
            mask = _next_n_days_mask(d['EXPIRY_DATE'], n)
            scope = f"next {n} days"
        else:
            mask = _next_n_days_mask(d['EXPIRY_DATE'], 30)
            scope = "next 30 days"
        d = d[mask]
        if d.empty:
            return (f"✅ No LGs expiring in the {scope}.", None)
        total = d['AMOUNT'].sum() if 'AMOUNT' in d.columns else 0
        return (f"LGs expiring in the **{scope}** — Count: {len(d):,.0f}, Total Amount: **SAR {total:,.0f}**.", _format_dates_for_display(d))

    # 6) Top banks
    if "top" in ql and "bank" in ql:
        if 'AMOUNT' in df_detailed.columns and 'BANK' in df_detailed.columns:
            g = df_detailed.groupby('BANK')['AMOUNT'].sum().sort_values(ascending=False).head(5)
            lines = [f"{i+1}. {b} — SAR {amt:,.0f}" for i, (b, amt) in enumerate(g.items())]
            return ("Top banks by total amount:\n" + "\n".join(lines), None)

    # 7) Show/list filters
    if any(k in ql for k in ["show", "list", "display"]) and ('BANK' in df_detailed.columns):
        d = df_detailed.copy()
        if bank_name:
            d = d[d['BANK'].astype(str) == bank_name]
        if branch_name:
            d = d[d['BRANCH'].astype(str).str.upper() == branch_name.upper()]
        if type_name:
            d = d[d['GUARANTEE_TYPE'].astype(str) == type_name]
        if d.empty:
            return ("No matching rows for that filter.", None)
        return (f"Showing {len(d)} row(s).", _format_dates_for_display(d, cols=('ISSUE_DATE','EXPIRY_DATE')))

    # Fallback
    if 'AMOUNT' in df_detailed.columns and 'BANK' in df_detailed.columns:
        g = df_detailed.groupby('BANK')['AMOUNT'].agg(['count','sum']).reset_index().rename(columns={'count':'Count','sum':'Total Amount'})
        return ("I wasn't sure of the exact intent. Here is a bank-wise overview:", g)

    return ("I couldn't interpret the question. Try:\n"
            "• total facilities / amount utilized\n"
            "• amount by bank/branch/type\n"
            "• maturities this month / next 30 days\n"
            "• LG_REF: <ref>", None)

# ========= Main =========
def main():
    st.markdown('<div class="main-header">🏦 LETTER OF GUARANTEE</div>', unsafe_allow_html=True)

    create_source_section()

    df_summary = load_summary_data_from_source(SOURCE_URL)
    df_detailed = load_detailed_data_from_source(SOURCE_URL)
    df_branch_util = load_branch_utilized_from_summary(SOURCE_URL)

    if df_summary is None or df_detailed is None or df_detailed.empty:
        st.error("No usable data found. Check the source sheet names/structure.")
        st.stop()

    # --- 🤖 Ask the Report (Chat) ---
    st.markdown('<div class="section-header">🤖 Ask the Report</div>', unsafe_allow_html=True)
    if "chat" not in st.session_state:
        st.session_state.chat = []

    # Render history
    for role, msg in st.session_state.chat:
        with st.chat_message(role):
            st.markdown(msg)

    user_q = st.chat_input("Ask: 'Total facilities', 'Amount for bank SNB', 'Branch Riyadh utilized', 'Maturities this month', 'LG_REF: 1234'…")
    if user_q:
        st.session_state.chat.append(("user", user_q))
        with st.chat_message("user"):
            st.markdown(user_q)

        ans_text, ans_df = answer_report(user_q, df_summary, df_detailed, df_branch_util)

        with st.chat_message("assistant"):
            st.markdown(ans_text)
            if isinstance(ans_df, pd.DataFrame) and not ans_df.empty:
                st.dataframe(_style_table(ans_df), use_container_width=True)

    # --- Filters: Branch + Bank (radio) ---
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

    # Apply filters to detailed
    df_detailed_filtered = df_detailed.copy()
    if selected_branch != 'All' and 'BRANCH' in df_detailed_filtered.columns:
        df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BRANCH'] == selected_branch]
    if selected_bank_radio != 'All' and 'BANK' in df_detailed_filtered.columns:
        df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BANK'] == selected_bank_radio]

    # Filter summary by bank if chosen
    try:
        df_summary_cols = {c.upper(): c for c in df_summary.columns}
        bank_hdr = df_summary_cols.get("BANKS") or df_summary_cols.get("BANK")
        if selected_bank_radio != 'All' and bank_hdr:
            df_summary_filtered = df_summary[df_summary[bank_hdr] == selected_bank_radio]
        else:
            df_summary_filtered = df_summary.copy()
    except Exception:
        df_summary_filtered = df_summary.copy()

    # KPIs (overall)
    if not df_summary_filtered.empty:
        df_summary_filtered = _ensure_rates_only(df_summary_filtered)
        create_summary_metrics(df_summary_filtered)

    # ---- Branch metric (Summary sheet two-column) ----
    if selected_branch != 'All' and df_branch_util is not None and not df_branch_util.empty:
        match = df_branch_util[df_branch_util['BRANCH'].str.upper() == selected_branch.upper()]
        if match.empty:
            match = df_branch_util[df_branch_util['BRANCH'].str.upper().str.contains(selected_branch.upper(), na=False)]
        branch_amt = float(match['AMOUNT_UTILIZED'].iloc[0]) if not match.empty else None
        st.markdown('<div class="section-header">🏢 Selected Branch Summary</div>', unsafe_allow_html=True)
        st.metric("Amount Utilized (Branch)", f"SAR {branch_amt:,.0f}" if branch_amt is not None else "—")

    # Tabs by Guarantee Type (All Types FIRST now)
    if 'GUARANTEE_TYPE' in df_detailed_filtered.columns:
        guarantee_types = df_detailed_filtered['GUARANTEE_TYPE'].dropna().unique().tolist()
        if guarantee_types:
            st.markdown('<div class="section-header">📋 Analysis by Guarantee Type</div>', unsafe_allow_html=True)

            tab_titles = ["🔄 All Types", "📅 Current Month Maturity"] + guarantee_types + ["📊 Summary Tables"]
            tabs = st.tabs(tab_titles)

            # 0) All Types
            with tabs[0]:
                st.subheader("🔄 All Guarantee Types (Combined)")
                render_summary_and_detailed_tables(
                    df_detailed_filtered, summary_by='GUARANTEE_TYPE', key_prefix="all_types"
                )
                charts_for_subset(df_detailed_filtered, "All Types")
                create_maturity_analysis(df_detailed_filtered)

            # 1) Current Month Maturity
            with tabs[1]:
                render_current_month_maturity(df_detailed_filtered)

            # 2..N) Individual guarantee type tabs
            for idx, gtype in enumerate(guarantee_types, start=2):
                with tabs[idx]:
                    render_type_tab(df_detailed_filtered, gtype)

            # Last) Summary tables
            with tabs[len(tab_titles) - 1]:
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

    # Footer removed (hidden)

if __name__ == "__main__":
    main()
