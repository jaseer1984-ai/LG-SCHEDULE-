import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import openpyxl

# Page configuration
st.set_page_config(
    page_title="LG Branch Summary Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
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
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_summary_data():
    """Load the summary data from the screenshot"""
    summary_data = {
        'BANKS': ['ANB', 'SAB', 'SNB', 'RB', 'NBK', 'INMA'],
        'TOTAL_FACILITIES': [5000000, 5000000, 30911990, 10000000, 10000000, 124367929],
        'AMOUNT_UTILIZED': [3503195, 3877442, 7561901, 8370935, 3774396, 124367929],
        'OUTSTANDING': [1496805, 1122558, 23350090, 1629065, 6225604, 0]
    }
    
    df = pd.DataFrame(summary_data)
    
    # Calculate additional metrics
    df['UTILIZATION_RATE'] = (df['AMOUNT_UTILIZED'] / df['TOTAL_FACILITIES'] * 100).round(2)
    df['OUTSTANDING_RATE'] = (df['OUTSTANDING'] / df['TOTAL_FACILITIES'] * 100).round(2)
    
    return df

@st.cache_data
def load_detailed_data():
    """Load detailed transaction data (sample)"""
    # Fixed number of records
    num_records = 80
    
    # Create data with consistent lengths
    banks = ['ANB', 'SAB', 'SNB', 'RB', 'NBK', 'INMA']
    customers = [
        'AL FOZAN TRADING', 'EL SEIF ENGINEERING', 'BETA CONSTRUCTION',
        'ALPHA CONTRACTING', 'GAMMA TRADING', 'DELTA ENGINEERING'
    ]
    guarantee_types = ['ADVANCE PAYMENT', 'PERFORMANCE BOND', 'BID BOND']
    
    # Ensure all arrays have exactly num_records length
    detailed_data = {
        'BANK': [banks[i % len(banks)] for i in range(num_records)],
        'LG_REF': [f'LG{str(i+1).zfill(6)}' for i in range(num_records)],
        'CUSTOMER_NAME': [customers[i % len(customers)] for i in range(num_records)],
        'GUARANTEE_TYPE': [guarantee_types[i % len(guarantee_types)] for i in range(num_records)],
        'AMOUNT': np.random.uniform(10000, 500000, num_records),
        'CURRENCY': ['SAR'] * num_records,
        'BRANCH': ['BETA RIYADH'] * num_records,
        'DAYS_TO_MATURE': np.random.randint(1, 365, num_records)
    }
    
    df_detailed = pd.DataFrame(detailed_data)
    
    # Add dates
    df_detailed['ISSUE_DATE'] = pd.date_range(start='2020-01-01', periods=num_records, freq='30D')
    df_detailed['EXPIRY_DATE'] = df_detailed['ISSUE_DATE'] + pd.to_timedelta(df_detailed['DAYS_TO_MATURE'], unit='D')
    
    return df_detailed

def create_summary_metrics(df):
    """Create summary metrics section"""
    st.markdown('<div class="section-header">📊 Key Performance Indicators</div>', unsafe_allow_html=True)
    
    # Calculate totals
    total_facilities = df['TOTAL_FACILITIES'].sum()
    total_utilized = df['AMOUNT_UTILIZED'].sum()
    total_outstanding = df['OUTSTANDING'].sum()
    avg_utilization = df['UTILIZATION_RATE'].mean()
    
    # Create columns for metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Total Facilities",
            value=f"SAR {total_facilities:,.0f}",
            help="Total credit facilities across all banks"
        )
    
    with col2:
        st.metric(
            label="Amount Utilized",
            value=f"SAR {total_utilized:,.0f}",
            delta=f"{(total_utilized/total_facilities)*100:.1f}% of total",
            help="Total amount currently utilized"
        )
    
    with col3:
        st.metric(
            label="Outstanding Amount",
            value=f"SAR {total_outstanding:,.0f}",
            delta=f"{(total_outstanding/total_facilities)*100:.1f}% available",
            help="Remaining available credit"
        )
    
    with col4:
        st.metric(
            label="Avg Utilization Rate",
            value=f"{avg_utilization:.1f}%",
            help="Average utilization across all banks"
        )

def create_bank_analysis_charts(df):
    """Create bank analysis charts"""
    st.markdown('<div class="section-header">🏦 Bank-wise Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Utilization Rate by Bank
        fig_util = px.bar(
            df, 
            x='BANKS', 
            y='UTILIZATION_RATE',
            title='Utilization Rate by Bank (%)',
            color='UTILIZATION_RATE',
            color_continuous_scale='RdYlBu_r',
            text='UTILIZATION_RATE'
        )
        fig_util.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_util.update_layout(
            showlegend=False,
            height=400,
            title_x=0.5,
            xaxis_title="Banks",
            yaxis_title="Utilization Rate (%)"
        )
        st.plotly_chart(fig_util, use_container_width=True)
    
    with col2:
        # Outstanding vs Utilized
        fig_comparison = go.Figure()
        
        fig_comparison.add_trace(go.Bar(
            name='Amount Utilized',
            x=df['BANKS'],
            y=df['AMOUNT_UTILIZED'],
            marker_color='#1f77b4'
        ))
        
        fig_comparison.add_trace(go.Bar(
            name='Outstanding',
            x=df['BANKS'],
            y=df['OUTSTANDING'],
            marker_color='#ff7f0e'
        ))
        
        fig_comparison.update_layout(
            title='Utilized vs Outstanding Amounts by Bank',
            title_x=0.5,
            xaxis_title="Banks",
            yaxis_title="Amount (SAR)",
            barmode='stack',
            height=400,
            legend=dict(x=0.7, y=1)
        )
        
        st.plotly_chart(fig_comparison, use_container_width=True)

def create_pie_charts(df):
    """Create pie charts for distribution analysis"""
    col1, col2 = st.columns(2)
    
    with col1:
        # Total Facilities Distribution
        fig_pie1 = px.pie(
            df, 
            values='TOTAL_FACILITIES', 
            names='BANKS',
            title='Total Facilities Distribution by Bank'
        )
        fig_pie1.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie1.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_pie1, use_container_width=True)
    
    with col2:
        # Outstanding Amount Distribution
        fig_pie2 = px.pie(
            df[df['OUTSTANDING'] > 0],  # Only show banks with outstanding amounts
            values='OUTSTANDING', 
            names='BANKS',
            title='Outstanding Amounts Distribution'
        )
        fig_pie2.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie2.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_pie2, use_container_width=True)

def create_detailed_analysis(df_detailed):
    """Create detailed transaction analysis"""
    st.markdown('<div class="section-header">📋 Detailed Transaction Analysis</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Guarantee Type Distribution
        guarantee_counts = df_detailed['GUARANTEE_TYPE'].value_counts()
        fig_guarantee = px.bar(
            x=guarantee_counts.index,
            y=guarantee_counts.values,
            title='Distribution by Guarantee Type',
            color=guarantee_counts.values,
            color_continuous_scale='viridis'
        )
        fig_guarantee.update_layout(
            height=400,
            title_x=0.5,
            xaxis_title="Guarantee Type",
            yaxis_title="Count",
            showlegend=False
        )
        st.plotly_chart(fig_guarantee, use_container_width=True)
    
    with col2:
        # Average Amount by Bank
        avg_amounts = df_detailed.groupby('BANK')['AMOUNT'].mean().sort_values(ascending=False)
        fig_avg = px.bar(
            x=avg_amounts.index,
            y=avg_amounts.values,
            title='Average Transaction Amount by Bank',
            color=avg_amounts.values,
            color_continuous_scale='Blues'
        )
        fig_avg.update_layout(
            height=400,
            title_x=0.5,
            xaxis_title="Bank",
            yaxis_title="Average Amount (SAR)",
            showlegend=False
        )
        st.plotly_chart(fig_avg, use_container_width=True)

def create_maturity_analysis(df_detailed):
    """Create maturity analysis"""
    st.markdown('<div class="section-header">⏰ Maturity Analysis</div>', unsafe_allow_html=True)
    
    # Categorize by days to mature
    def categorize_maturity(days):
        if days <= 30:
            return '≤ 30 days'
        elif days <= 90:
            return '31-90 days'
        elif days <= 180:
            return '91-180 days'
        else:
            return '> 180 days'
    
    df_detailed['MATURITY_CATEGORY'] = df_detailed['DAYS_TO_MATURE'].apply(categorize_maturity)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Maturity distribution
        maturity_dist = df_detailed['MATURITY_CATEGORY'].value_counts()
        fig_maturity = px.pie(
            values=maturity_dist.values,
            names=maturity_dist.index,
            title='LG Distribution by Time to Maturity'
        )
        fig_maturity.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_maturity, use_container_width=True)
    
    with col2:
        # Bank-wise maturity
        maturity_bank = df_detailed.groupby(['BANK', 'MATURITY_CATEGORY']).size().reset_index(name='count')
        fig_bank_maturity = px.bar(
            maturity_bank,
            x='BANK',
            y='count',
            color='MATURITY_CATEGORY',
            title='Maturity Distribution by Bank',
            barmode='stack'
        )
        fig_bank_maturity.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_bank_maturity, use_container_width=True)

def main():
    """Main dashboard function"""
    # Header
    st.markdown('<div class="main-header">🏦 LG Branch Summary Dashboard 2025</div>', unsafe_allow_html=True)
    
    # Load data
    df_summary = load_summary_data()
    df_detailed = load_detailed_data()
    
    # Sidebar
    st.sidebar.title("🔧 Dashboard Controls")
    st.sidebar.markdown("---")
    
    # Bank filter
    selected_banks = st.sidebar.multiselect(
        "Select Banks",
        options=df_summary['BANKS'].tolist(),
        default=df_summary['BANKS'].tolist(),
        help="Filter banks to display in charts"
    )
    
    # Filter data
    df_filtered = df_summary[df_summary['BANKS'].isin(selected_banks)]
    df_detailed_filtered = df_detailed[df_detailed['BANK'].isin(selected_banks)]
    
    # Display date and time
    st.sidebar.markdown(f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.sidebar.markdown("**Data Source:** OUTSTANDING LGS_AS OF 2025.xlsx")
    
    # Summary metrics
    create_summary_metrics(df_filtered)
    
    # Bank analysis charts
    create_bank_analysis_charts(df_filtered)
    
    # Pie charts
    create_pie_charts(df_filtered)
    
    # Detailed analysis
    create_detailed_analysis(df_detailed_filtered)
    
    # Maturity analysis
    create_maturity_analysis(df_detailed_filtered)
    
    # Data tables
    st.markdown('<div class="section-header">📊 Data Tables</div>', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📈 Summary Data", "📋 Detailed Transactions"])
    
    with tab1:
        st.subheader("Bank Summary")
        st.dataframe(
            df_filtered.style.format({
                'TOTAL_FACILITIES': '{:,.0f}',
                'AMOUNT_UTILIZED': '{:,.0f}',
                'OUTSTANDING': '{:,.0f}',
                'UTILIZATION_RATE': '{:.1f}%',
                'OUTSTANDING_RATE': '{:.1f}%'
            }),
            use_container_width=True
        )
    
    with tab2:
        st.subheader("Transaction Details")
        st.dataframe(
            df_detailed_filtered.style.format({
                'AMOUNT': '{:,.2f}',
                'DAYS_TO_MATURE': '{:.0f}'
            }),
            use_container_width=True
        )
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666;'>"
        "LG Branch Summary Dashboard • Built with Streamlit • Data as of September 2025"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
