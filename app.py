# -*- coding: utf-8 -*-
# Tawsif Travel & Tourism — BI Dashboard
# - Hides file-uploader UI & messages after a successful upload
# - Dates shown as dd-mm-yyyy (no time) in tables, footer, and chart axes
# - Safe date picker (single-day vs range), branch filter, KPIs, charts, tables

import io
from datetime import datetime, date, timedelta
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ---------------- Page config ----------------
st.set_page_config(
    page_title="Tawsif Travel & Tourism Dashboard",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------- Styles ----------------
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 30px;
        color: white;
        text-align: center;
    }
    .metric-card {background: white;padding: 20px;border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);border-left: 4px solid #2a5298;
        margin-bottom: 20px;}
    .sidebar .sidebar-content {background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);}
    .stSelectbox > div > div {background-color: #f8f9fa;}
    .success-message {background: #d4edda;border: 1px solid #c3e6cb;border-radius: 5px;
        padding: 10px;margin: 10px 0;color: #155724;}
    .error-message {background: #f8d7da;border: 1px solid #f5c6cb;border-radius: 5px;
        padding: 10px;margin: 10px 0;color: #721c24;}
</style>
""", unsafe_allow_html=True)

# ---------------- Header ----------------
st.markdown("""
<div class="main-header">
    <h1>✈️ Tawsif Travel & Tourism Company</h1>
    <h3>Business Intelligence Dashboard</h3>
</div>
""", unsafe_allow_html=True)

# ---------------- Session init ----------------
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'data_dict' not in st.session_state:
    st.session_state.data_dict = {}
if 'show_uploader' not in st.session_state:
    st.session_state.show_uploader = True  # control showing/hiding the uploader block

# ---------------- Helpers ----------------
SHEET_MAPPING = {
    'Daily_Summary': ['Daily_Summary', 'Daily Summary', 'daily_summary', 'Summary'],
    'Tickets_By_Airline': ['Tickets_By_Airline', 'Tickets By Airline', 'tickets_by_airline', 'Tickets'],
    'Airline_Sales': ['Airline_Sales', 'Airline Sales', 'airline_sales', 'Sales'],
    'Staff_Sales': ['Staff_Sales', 'Staff Sales', 'staff_sales', 'Staff'],
    'Bank_Balances': ['Bank_Balances', 'Bank Balances', 'bank_balances', 'Banks'],
}

NORMALIZE_MAP = {
    'Daily_Summary': {
        'Daily_Sales': 'Daily Sales', 'daily_sales': 'Daily Sales', 'Sales': 'Daily Sales', 'Total Sales': 'Daily Sales',
        'Cash_Balance': 'Cash Balance', 'cash_balance': 'Cash Balance', 'Cash': 'Cash Balance',
        'Bank_Balance': 'Bank Balance', 'bank_balance': 'Bank Balance', 'Bank': 'Bank Balance',
    },
    'Tickets_By_Airline': {'Tickets_Issued': 'Tickets Issued', 'tickets_issued': 'Tickets Issued', 'Tickets': 'Tickets Issued'},
    'Staff_Sales': {'Tickets_Issued': 'Tickets Issued', 'tickets_issued': 'Tickets Issued', 'Tickets': 'Tickets Issued'},
    'Bank_Balances': {'balance': 'Balance', 'Amount': 'Balance', 'amount': 'Balance'},
}

@st.cache_data
def load_excel_data(uploaded_file):
    """
    Read Excel sheets; map sheet names; ensure Date is python date (no time).
    Returns a dict of expected sheets if they exist.
    """
    excel_data = pd.read_excel(uploaded_file, sheet_name=None)
    available = list(excel_data.keys())
    mapped = {}

    for expected, candidates in SHEET_MAPPING.items():
        found = None
        for c in candidates:
            if c in available:
                found = c
                break
        if found:
            df = excel_data[found].copy()
            df.columns = df.columns.str.strip()
            if 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date  # drop time
            mapped[expected] = df

    # Normalize columns
    out = {}
    for sheet_name, df in mapped.items():
        if sheet_name in NORMALIZE_MAP:
            df = df.rename(columns=NORMALIZE_MAP[sheet_name])
        out[sheet_name] = df
    return out

def fmt_dates_for_table(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with Date formatted as dd-mm-yyyy string for display."""
    if 'Date' in df.columns:
        d = df.copy()
        d['Date'] = pd.to_datetime(d['Date'], errors='coerce').dt.strftime('%d-%m-%Y')
        return d
    return df

def chart_xaxis_ddmmyyyy(fig):
    fig.update_xaxes(tickformat="%d-%m-%Y")
    return fig

# ---------------- Sidebar: Re-upload control ----------------
st.sidebar.header("⚙️ Controls")
if st.sidebar.button("🔄 Re-upload Excel"):
    st.session_state.data_loaded = False
    st.session_state.data_dict = {}
    st.session_state.show_uploader = True
    st.rerun()

# ---------------- Upload block (hidden after success) ----------------
if st.session_state.show_uploader:
    uploader_ph = st.empty()
    with uploader_ph.container():
        # Minimal uploader (no label, no extra info)
        uploaded_file = st.file_uploader(
            label="",
            type=['xlsx', 'xls'],
            label_visibility="collapsed",
            help=None
        )
        # (Optional) You can put muted guidance here if you want
        # st.caption("Upload your Excel file to begin.")

        if uploaded_file is not None:
            with st.spinner("📊 Loading data..."):
                try:
                    data_dict = load_excel_data(uploaded_file)
                    if not data_dict:
                        st.error("The workbook is missing expected sheets. Please check your file.")
                    else:
                        st.session_state.data_dict = data_dict
                        st.session_state.data_loaded = True
                        st.session_state.show_uploader = False  # Hide uploader from now on
                        uploader_ph.empty()  # << Hides the entire upload UI & messages
                        st.rerun()           # ensure clean state without the uploader block
                except Exception as e:
                    st.error(f"Failed to read the Excel file: {e}")

# ---------------- Dashboard ----------------
if st.session_state.data_loaded and st.session_state.data_dict:
    data_dict = st.session_state.data_dict

    st.sidebar.header("📊 Dashboard Filters")

    # Date picker (safe: single vs range)
    all_dates = []
    for df in data_dict.values():
        if 'Date' in df.columns:
            all_dates.extend([d for d in df['Date'] if pd.notna(d)])

    if all_dates:
        min_date = min(all_dates)
        max_date = max(all_dates)
        default_start = max(min_date, max_date - timedelta(days=30))

        if min_date == max_date:
            sel = st.sidebar.date_input(
                "Select Date", value=max_date, min_value=min_date, max_value=max_date, key="single_date"
            )
        else:
            sel = st.sidebar.date_input(
                "Select Date Range", value=(default_start, max_date),
                min_value=min_date, max_value=max_date, key="range_date"
            )

        if isinstance(sel, tuple) and len(sel) == 2:
            start_date, end_date = sel
        else:
            start_date = end_date = sel
    else:
        end_date = date.today()
        start_date = end_date - timedelta(days=30)

    # Branch filter
    branches = ['All']
    if 'Tickets_By_Airline' in data_dict and 'Branch' in data_dict['Tickets_By_Airline'].columns:
        branches.extend(list(pd.Series(data_dict['Tickets_By_Airline']['Branch']).dropna().unique()))
    selected_branch = st.sidebar.selectbox("Select Branch", branches)

    # Filter data
    filtered_data = {}
    for name, df in data_dict.items():
        temp = df.copy()
        if 'Date' in temp.columns:
            mask = (temp['Date'] >= start_date) & (temp['Date'] <= end_date)
            temp = temp.loc[mask]
        if selected_branch != 'All' and 'Branch' in temp.columns:
            temp = temp[temp['Branch'] == selected_branch]
        filtered_data[name] = temp

    # ---------------- KPIs ----------------
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if 'Daily_Summary' in filtered_data and not filtered_data['Daily_Summary'].empty:
            total_sales = float(filtered_data['Daily_Summary']['Daily Sales'].sum())
            days_count = max((end_date - start_date).days + 1, 1)
            avg_daily = total_sales / days_count
            st.metric("💰 Total Sales", f"SAR {total_sales:,.0f}", f"{avg_daily:,.0f} avg/day")
        else:
            st.metric("💰 Total Sales", "—")

    with c2:
        if 'Tickets_By_Airline' in filtered_data and not filtered_data['Tickets_By_Airline'].empty:
            total_tickets = int(filtered_data['Tickets_By_Airline']['Tickets Issued'].sum())
            days_count = max((end_date - start_date).days + 1, 1)
            avg_daily_tk = total_tickets / days_count
            st.metric("🎫 Total Tickets", f"{total_tickets:,}", f"{avg_daily_tk:,.0f} avg/day")
        else:
            st.metric("🎫 Total Tickets", "—")

    with c3:
        if 'Daily_Summary' in filtered_data and not filtered_data['Daily_Summary'].empty:
            avg_cash = float(filtered_data['Daily_Summary']['Cash Balance'].mean())
            st.metric("💵 Avg Cash Balance", f"SAR {avg_cash:,.0f}", "Daily Average")
        else:
            st.metric("💵 Avg Cash Balance", "—")

    with c4:
        if 'Bank_Balances' in filtered_data and not filtered_data['Bank_Balances'].empty:
            total_bank = float(filtered_data['Bank_Balances']['Balance'].sum())
            st.metric("🏦 Total Bank Balance", f"SAR {total_bank:,.0f}", "All Banks")
        else:
            st.metric("🏦 Total Bank Balance", "—")

    st.markdown("---")

    # ---------------- Charts ----------------
    colA, colB = st.columns(2)

    with colA:
        st.subheader("📈 Daily Sales Trend")
        if 'Daily_Summary' in filtered_data and not filtered_data['Daily_Summary'].empty:
            df_sales = filtered_data['Daily_Summary'].copy()
            df_sales['Date_dt'] = pd.to_datetime(df_sales['Date'])  # for plotly
            fig_sales = px.line(
                df_sales, x='Date_dt', y='Daily Sales',
                labels={'Daily Sales': 'Sales (SAR)', 'Date_dt': 'Date'},
                title="Daily Sales Over Time"
            )
            fig_sales.update_traces(line_color='#2a5298', line_width=3, showlegend=False)
            chart_xaxis_ddmmyyyy(fig_sales)
            st.plotly_chart(fig_sales, use_container_width=True)
        else:
            st.info("No daily sales data available")

    with colB:
        st.subheader("✈️ Airline Performance")
        if 'Airline_Sales' in filtered_data and not filtered_data['Airline_Sales'].empty:
            airline_summary = (
                filtered_data['Airline_Sales']
                .groupby('Airline', dropna=True)['Sales']
                .sum().reset_index().sort_values('Sales', ascending=False)
            )
            fig_airline = px.bar(
                airline_summary, x='Airline', y='Sales',
                labels={'Sales': 'Total Sales (SAR)', 'Airline': 'Airline'},
                title="Sales by Airline",
                color='Sales', color_continuous_scale='Blues'
            )
            fig_airline.update_layout(showlegend=False, xaxis_tickangle=-45)
            st.plotly_chart(fig_airline, use_container_width=True)
        else:
            st.info("No airline sales data available")

    colC, colD = st.columns(2)

    with colC:
        st.subheader("👥 Top Performing Staff")
        if 'Staff_Sales' in filtered_data and not filtered_data['Staff_Sales'].empty:
            staff_summary = (
                filtered_data['Staff_Sales']
                .groupby('Staff', dropna=True)
                .agg({'Sales': 'sum', 'Tickets Issued': 'sum'})
                .reset_index().sort_values('Sales', ascending=False).head(8)
            )
            fig_staff = px.bar(
                staff_summary, x='Staff', y='Sales',
                labels={'Sales': 'Total Sales (SAR)', 'Staff': 'Staff Member'},
                title="Sales by Staff Member",
                color='Sales', color_continuous_scale='Greens'
            )
            fig_staff.update_layout(showlegend=False, xaxis_tickangle=-45)
            st.plotly_chart(fig_staff, use_container_width=True)
        else:
            st.info("No staff sales data available")

    with colD:
        st.subheader("🎯 Ticket Distribution by Airline")
        if 'Tickets_By_Airline' in filtered_data and not filtered_data['Tickets_By_Airline'].empty:
            ticket_summary = (
                filtered_data['Tickets_By_Airline']
                .groupby('Airline', dropna=True)['Tickets Issued']
                .sum().reset_index()
            )
            fig_tickets = px.pie(
                ticket_summary, values='Tickets Issued', names='Airline',
                title="Ticket Distribution", color_discrete_sequence=px.colors.qualitative.Set3
            )
            st.plotly_chart(fig_tickets, use_container_width=True)
        else:
            st.info("No ticket data available")

    # Financial Overview
    st.subheader("💳 Financial Overview")
    colE, colF = st.columns(2)

    with colE:
        st.write("**Cash vs Bank Balance Trend**")
        if 'Daily_Summary' in filtered_data and not filtered_data['Daily_Summary'].empty:
            df_bal = filtered_data['Daily_Summary'].copy()
            df_bal['Date_dt'] = pd.to_datetime(df_bal['Date'])
            fig_balance = make_subplots(specs=[[{"secondary_y": True}]])
            fig_balance.add_trace(
                go.Scatter(x=df_bal['Date_dt'], y=df_bal['Cash Balance'],
                           name="Cash Balance", line=dict(color='green', width=2)),
                secondary_y=False
            )
            fig_balance.add_trace(
                go.Scatter(x=df_bal['Date_dt'], y=df_bal['Bank Balance'],
                           name="Bank Balance", line=dict(color='blue', width=2)),
                secondary_y=True
            )
            fig_balance.update_xaxes(title_text="Date", tickformat="%d-%m-%Y")
            fig_balance.update_yaxes(title_text="Cash Balance (SAR)", secondary_y=False)
            fig_balance.update_yaxes(title_text="Bank Balance (SAR)", secondary_y=True)
            fig_balance.update_layout(height=400, showlegend=True)
            st.plotly_chart(fig_balance, use_container_width=True)
        else:
            st.info("No financial balance data available")

    with colF:
        st.write("**Bank Balances Distribution**")
        if 'Bank_Balances' in filtered_data and not filtered_data['Bank_Balances'].empty:
            bank_summary = (
                filtered_data['Bank_Balances']
                .groupby('Bank', dropna=True)['Balance']
                .sum().reset_index().sort_values('Balance', ascending=False)
            )
            fig_banks = px.bar(
                bank_summary, x='Bank', y='Balance',
                labels={'Balance': 'Total Balance (SAR)', 'Bank': 'Bank'},
                title="Balance by Bank",
                color='Balance', color_continuous_scale='Oranges'
            )
            fig_banks.update_layout(showlegend=False)
            st.plotly_chart(fig_banks, use_container_width=True)
        else:
            st.info("No bank balance data available")

    # ---------------- Tables ----------------
    st.markdown("---")
    st.subheader("📋 Detailed Data Tables")

    table_order = ['Daily_Summary', 'Airline_Sales', 'Staff_Sales', 'Bank_Balances']
    available_tables = [t for t in table_order if t in filtered_data and not filtered_data[t].empty]

    if available_tables:
        tabs = st.tabs([t.replace('_', ' ') for t in available_tables])
        for i, tname in enumerate(available_tables):
            with tabs[i]:
                df_show = fmt_dates_for_table(filtered_data[tname])

                # Numeric formatting
                numeric_cols = df_show.select_dtypes(include=[np.number]).columns
                fmt = {}
                for col in numeric_cols:
                    if 'Sales' in col or 'Balance' in col:
                        fmt[col] = 'SAR {:,.0f}'
                    elif 'Tickets' in col:
                        fmt[col] = '{:,}'

                if fmt:
                    st.dataframe(df_show.style.format(fmt), use_container_width=True)
                else:
                    st.dataframe(df_show, use_container_width=True)
    else:
        st.info("No data tables available to display")

    # ---------------- Footer ----------------
    st.markdown("---")
    st.markdown(f"""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <p>© 2025 Tawsif Travel & Tourism Company - Business Intelligence Dashboard</p>
        <p>Last Updated: {datetime.now().strftime("%d-%m-%Y %H:%M:%S")} |
           Data Range: {start_date.strftime("%d-%m-%Y")} to {end_date.strftime("%d-%m-%Y")}</p>
    </div>
    """, unsafe_allow_html=True)

else:
    # If no data loaded and uploader is hidden (shouldn't happen), show a gentle nudge
    if not st.session_state.show_uploader:
        st.info("Click **🔄 Re-upload Excel** in the sidebar to upload a file.")
