import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import openpyxl
import io

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
    
    .upload-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px dashed #1f4e79;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_summary_data_from_file(uploaded_file):
    """Load summary data from uploaded Excel file"""
    try:
        # Read the Summary sheet or create from LG BRANCH SUMMARY_2025
        try:
            df = pd.read_excel(uploaded_file, sheet_name='Summary')
        except:
            # If Summary sheet doesn't exist, create it from detailed data
            df_detailed = pd.read_excel(uploaded_file, sheet_name='LG BRANCH SUMMARY_2025')
            
            # Group by bank to create summary
            summary_data = df_detailed.groupby('BANK').agg({
                'AMOUNT': 'sum'
            }).reset_index()
            
            # Create mock summary data structure
            summary_data.rename(columns={'AMOUNT': 'AMOUNT_UTILIZED'}, inplace=True)
            summary_data['TOTAL_FACILITIES'] = summary_data['AMOUNT_UTILIZED'] * 1.5  # Mock total
            summary_data['OUTSTANDING'] = summary_data['TOTAL_FACILITIES'] - summary_data['AMOUNT_UTILIZED']
            summary_data.rename(columns={'BANK': 'BANKS'}, inplace=True)
            df = summary_data
            
        # Calculate additional metrics
        df['UTILIZATION_RATE'] = (df['AMOUNT_UTILIZED'] / df['TOTAL_FACILITIES'] * 100).round(2)
        df['OUTSTANDING_RATE'] = (df['OUTSTANDING'] / df['TOTAL_FACILITIES'] * 100).round(2)
        
        return df
    except Exception as e:
        st.error(f"Error loading summary data: {str(e)}")
        return None

@st.cache_data
def load_detailed_data_from_file(uploaded_file):
    """Load detailed data from uploaded Excel file"""
    try:
        # Read the main transaction data (columns A:K based on your images)
        df = pd.read_excel(uploaded_file, sheet_name='LG BRANCH SUMMARY_2025', usecols="A:K")
        
        st.write("Original columns:", df.columns.tolist())
        st.write("Data shape:", df.shape)
        st.write("First few rows:", df.head())
        
        # Based on your images, the exact column mapping should be:
        expected_columns = {
            0: 'BANK',           # Column A
            1: 'LG_REF',         # Column B - LG REF
            2: 'CUSTOMER_NAME',  # Column C - CUSTOMER NAME
            3: 'GUARANTEE_TYPE', # Column D - GUARRENTY TYPE
            4: 'ISSUE_DATE',     # Column E - ISSUE DATE
            5: 'EXPIRY_DATE',    # Column F - EXPIRY DATE
            6: 'AMOUNT',         # Column G - AMOUNT
            7: 'CURRENCY',       # Column H - CURRENCY
            8: 'BRANCH',         # Column I - BRANCH
            9: 'BANK_2',         # Column J - BANK (duplicate)
            10: 'DAYS_TO_MATURE' # Column K - DAYS TO MATURE
        }
        
        # Create new column names list
        new_columns = []
        for i, col in enumerate(df.columns):
            if i in expected_columns:
                new_columns.append(expected_columns[i])
            else:
                new_columns.append(f'Column_{i}')
        
        # Apply new column names
        df.columns = new_columns
        
        # Use the first BANK column and drop the duplicate
        df = df.drop('BANK_2', axis=1, errors='ignore')
        
        st.write("Mapped columns:", df.columns.tolist())
        
        # Clean and validate the data
        # Remove header rows that might be mixed in the data
        df = df[df['BANK'].notna()]  # Remove rows where BANK is null
        df = df[df['BANK'] != 'BANK']  # Remove header rows
        
        # Clean the GUARANTEE_TYPE column (fix the typo in original)
        if 'GUARANTEE_TYPE' in df.columns:
            df['GUARANTEE_TYPE'] = df['GUARANTEE_TYPE'].astype(str).str.strip()
        
        # Convert numeric columns
        if 'AMOUNT' in df.columns:
            df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce').fillna(0)
        
        if 'DAYS_TO_MATURE' in df.columns:
            df['DAYS_TO_MATURE'] = pd.to_numeric(df['DAYS_TO_MATURE'], errors='coerce').fillna(30)
        
        # Convert date columns
        for date_col in ['ISSUE_DATE', 'EXPIRY_DATE']:
            if date_col in df.columns:
                try:
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                except:
                    st.warning(f"Could not convert {date_col} to datetime")
        
        # Remove rows with missing essential data
        df_cleaned = df.dropna(subset=['BANK', 'GUARANTEE_TYPE'])
        df_cleaned = df_cleaned[df_cleaned['BANK'].str.strip() != '']
        df_cleaned = df_cleaned[df_cleaned['GUARANTEE_TYPE'].str.strip() != '']
        
        # Fill missing values for optional columns
        if 'LG_REF' in df_cleaned.columns:
            df_cleaned['LG_REF'] = df_cleaned['LG_REF'].fillna('N/A')
        
        if 'CUSTOMER_NAME' in df_cleaned.columns:
            df_cleaned['CUSTOMER_NAME'] = df_cleaned['CUSTOMER_NAME'].fillna('Unknown Customer')
        
        if 'BRANCH' in df_cleaned.columns:
            df_cleaned['BRANCH'] = df_cleaned['BRANCH'].fillna('Main Branch')
        
        if 'CURRENCY' in df_cleaned.columns:
            df_cleaned['CURRENCY'] = df_cleaned['CURRENCY'].fillna('SAR')
        
        st.success(f"Successfully loaded {len(df_cleaned)} records")
        st.write("Final columns:", df_cleaned.columns.tolist())
        st.write("Sample data:", df_cleaned.head(3))
        
        return df_cleaned
        
    except Exception as e:
        st.error(f"Error loading detailed data: {str(e)}")
        st.write("Attempting to load with different approach...")
        
        # Fallback: try to read without column restrictions
        try:
            df_fallback = pd.read_excel(uploaded_file, sheet_name='LG BRANCH SUMMARY_2025')
            st.write("Fallback - All columns:", df_fallback.columns.tolist())
            
            # Try to find the right columns by name
            column_mapping = {}
            for col in df_fallback.columns:
                col_str = str(col).upper().strip()
                if col_str == 'BANK':
                    column_mapping[col] = 'BANK'
                elif 'LG REF' in col_str or 'LG_REF' in col_str:
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
            
            df_fallback = df_fallback.rename(columns=column_mapping)
            return df_fallback
            
        except Exception as e2:
            st.error(f"Fallback also failed: {str(e2)}")
            return None

@st.cache_data
def load_default_summary_data():
    """Load default summary data when no file is uploaded"""
    summary_data = {
        'BANKS': ['ANB', 'SAB', 'SNB', 'RB', 'NBK', 'INMA'],
        'TOTAL_FACILITIES': [5000000, 5000000, 30911990, 10000000, 10000000, 124367929],
        'AMOUNT_UTILIZED': [3503195, 3877442, 7561901, 8370935, 3774396, 124367929],
        'OUTSTANDING': [1496805, 1122558, 23350090, 1629065, 6225604, 0]
    }
    
    df = pd.DataFrame(summary_data)
    df['UTILIZATION_RATE'] = (df['AMOUNT_UTILIZED'] / df['TOTAL_FACILITIES'] * 100).round(2)
    df['OUTSTANDING_RATE'] = (df['OUTSTANDING'] / df['TOTAL_FACILITIES'] * 100).round(2)
    
    return df

@st.cache_data
def load_default_detailed_data():
    """Load default detailed data when no file is uploaded"""
    num_records = 80
    banks = ['ANB', 'SAB', 'SNB', 'RB', 'NBK', 'INMA']
    customers = [
        'AL FOZAN TRADING', 'EL SEIF ENGINEERING', 'BETA CONSTRUCTION',
        'ALPHA CONTRACTING', 'GAMMA TRADING', 'DELTA ENGINEERING'
    ]
    guarantee_types = ['ADVANCE PAYMENT', 'PERFORMANCE BOND', 'BID BOND']
    branches = ['BETA RIYADH', 'ALPHA JEDDAH', 'GAMMA DAMMAM']
    
    detailed_data = {
        'BANK': [banks[i % len(banks)] for i in range(num_records)],
        'LG_REF': [f'LG{str(i+1).zfill(6)}' for i in range(num_records)],
        'CUSTOMER_NAME': [customers[i % len(customers)] for i in range(num_records)],
        'GUARANTEE_TYPE': [guarantee_types[i % len(guarantee_types)] for i in range(num_records)],
        'AMOUNT': np.random.uniform(10000, 500000, num_records),
        'CURRENCY': ['SAR'] * num_records,
        'BRANCH': [branches[i % len(branches)] for i in range(num_records)],
        'DAYS_TO_MATURE': np.random.randint(1, 365, num_records)
    }
    
    df_detailed = pd.DataFrame(detailed_data)
    df_detailed['ISSUE_DATE'] = pd.date_range(start='2020-01-01', periods=num_records, freq='30D')
    df_detailed['EXPIRY_DATE'] = df_detailed['ISSUE_DATE'] + pd.to_timedelta(df_detailed['DAYS_TO_MATURE'], unit='D')
    
    return df_detailed

def create_file_upload_section():
    """Create file upload section"""
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown("### 📁 Upload Your Excel File")
    st.markdown("Upload your **OUTSTANDING LGS_AS OF 2025.xlsx** file or any Excel file with **LG BRANCH SUMMARY_2025** sheet")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=['xlsx', 'xls'],
        help="Upload Excel file containing LG data"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded successfully: {uploaded_file.name}")
        
        # Show file info
        file_details = {
            "Filename": uploaded_file.name,
            "File size": f"{uploaded_file.size} bytes",
            "File type": uploaded_file.type
        }
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Filename", file_details["Filename"])
        with col2:
            st.metric("Size", file_details["File size"])
        with col3:
            st.metric("Type", file_details["File type"])
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    return uploaded_file

def create_summary_metrics(df):
    """Create summary metrics section"""
    st.markdown('<div class="section-header">📊 Key Performance Indicators</div>', unsafe_allow_html=True)
    
    total_facilities = df['TOTAL_FACILITIES'].sum()
    total_utilized = df['AMOUNT_UTILIZED'].sum()
    total_outstanding = df['OUTSTANDING'].sum()
    avg_utilization = df['UTILIZATION_RATE'].mean()
    
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

def create_charts_for_guarantee_type(df_filtered, guarantee_type):
    """Create charts for specific guarantee type"""
    
    # Filter data for this guarantee type
    df_type = df_filtered[df_filtered['GUARANTEE_TYPE'] == guarantee_type]
    
    if df_type.empty:
        st.warning(f"No data available for {guarantee_type}")
        return
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Amount distribution by bank
        bank_amounts = df_type.groupby('BANK')['AMOUNT'].sum().sort_values(ascending=False)
        
        fig_bank = px.bar(
            x=bank_amounts.index,
            y=bank_amounts.values,
            title=f'{guarantee_type} - Amount by Bank',
            color=bank_amounts.values,
            color_continuous_scale='Blues'
        )
        fig_bank.update_layout(
            height=400,
            title_x=0.5,
            xaxis_title="Bank",
            yaxis_title="Total Amount (SAR)",
            showlegend=False
        )
        st.plotly_chart(fig_bank, use_container_width=True)
    
    with col2:
        # Count of LGs by bank
        bank_counts = df_type.groupby('BANK').size().sort_values(ascending=False)
        
        fig_count = px.pie(
            values=bank_counts.values,
            names=bank_counts.index,
            title=f'{guarantee_type} - Count by Bank'
        )
        fig_count.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_count, use_container_width=True)
    
    # Summary table for this guarantee type
    st.subheader(f"📊 {guarantee_type} Summary")
    
    summary_stats = df_type.groupby('BANK').agg({
        'AMOUNT': ['count', 'sum', 'mean'],
        'DAYS_TO_MATURE': 'mean'
    }).round(2)
    
    summary_stats.columns = ['Count', 'Total Amount', 'Avg Amount', 'Avg Days to Mature']
    summary_stats = summary_stats.reset_index()
    
    st.dataframe(
        summary_stats.style.format({
            'Total Amount': '{:,.0f}',
            'Avg Amount': '{:,.0f}',
            'Avg Days to Mature': '{:.0f}'
        }),
        use_container_width=True
    )

def create_maturity_analysis(df_detailed):
    """Create maturity analysis"""
    st.markdown('<div class="section-header">⏰ Maturity Analysis</div>', unsafe_allow_html=True)
    
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
        maturity_dist = df_detailed['MATURITY_CATEGORY'].value_counts()
        fig_maturity = px.pie(
            values=maturity_dist.values,
            names=maturity_dist.index,
            title='LG Distribution by Time to Maturity'
        )
        fig_maturity.update_layout(height=400, title_x=0.5)
        st.plotly_chart(fig_maturity, use_container_width=True)
    
    with col2:
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
    try:
        # Header
        st.markdown('<div class="main-header">🏦 LG Branch Summary Dashboard 2025</div>', unsafe_allow_html=True)
        
        # File upload section
        uploaded_file = create_file_upload_section()
        
        # Load data based on file upload
        df_summary = None
        df_detailed = None
        
        if uploaded_file is not None:
            try:
                df_summary = load_summary_data_from_file(uploaded_file)
                df_detailed = load_detailed_data_from_file(uploaded_file)
                
                if df_summary is None or df_detailed is None:
                    st.error("Failed to load data from uploaded file. Using default data.")
                    raise Exception("Failed to load uploaded file")
                    
            except Exception as e:
                st.error(f"Error processing uploaded file: {str(e)}")
                df_summary = None
                df_detailed = None
        
        # Use default data if file loading failed or no file uploaded
        if df_summary is None or df_detailed is None:
            if uploaded_file is None:
                st.info("📝 No file uploaded. Using sample data for demonstration.")
            try:
                df_summary = load_default_summary_data()
                df_detailed = load_default_detailed_data()
            except Exception as e:
                st.error(f"Critical error loading default data: {str(e)}")
                st.stop()
        
        # Verify data is loaded
        if df_summary is None or df_detailed is None or df_detailed.empty:
            st.error("No data could be loaded. Please refresh the page and try again.")
            st.stop()
        
        # Sidebar filters
        st.sidebar.title("🔧 Dashboard Filters")
        st.sidebar.markdown("---")
        
        # Bank filter (multiselect)
        available_banks = df_detailed['BANK'].unique().tolist() if 'BANK' in df_detailed.columns else []
        selected_banks = st.sidebar.multiselect(
            "🏦 Select Banks",
            options=available_banks,
            default=available_banks,
            help="Filter banks to display in charts"
        )
        
        # Branch filter (radio buttons)
        st.sidebar.markdown("### 🏢 Branch Filter")
        available_branches = ['All']
        if 'BRANCH' in df_detailed.columns:
            branch_values = df_detailed['BRANCH'].dropna().unique().tolist()
            available_branches.extend(branch_values)
        
        selected_branch = st.sidebar.radio(
            "Select Branch",
            options=available_branches,
            help="Filter by branch location"
        )
        
        # Bank filter (radio buttons) - secondary filter
        st.sidebar.markdown("### 🏦 Bank Filter (Secondary)")
        available_banks_radio = ['All']
        if 'BANK' in df_detailed.columns:
            bank_values = df_detailed['BANK'].dropna().unique().tolist()
            available_banks_radio.extend(bank_values)
        
        selected_bank_radio = st.sidebar.radio(
            "Select Specific Bank",
            options=available_banks_radio,
            help="Focus on specific bank"
        )
        
        # Apply filters safely
        df_detailed_filtered = df_detailed.copy()
        
        try:
            # Apply multiselect bank filter
            if selected_banks and 'BANK' in df_detailed_filtered.columns:
                df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BANK'].isin(selected_banks)]
            
            # Apply branch filter
            if selected_branch != 'All' and 'BRANCH' in df_detailed_filtered.columns:
                df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BRANCH'] == selected_branch]
            
            # Apply radio bank filter
            if selected_bank_radio != 'All' and 'BANK' in df_detailed_filtered.columns:
                df_detailed_filtered = df_detailed_filtered[df_detailed_filtered['BANK'] == selected_bank_radio]
        
        except Exception as e:
            st.warning(f"Error applying filters: {str(e)}. Using unfiltered data.")
            df_detailed_filtered = df_detailed.copy()
        
        # Filter summary data based on selected banks
        try:
            if selected_banks and 'BANKS' in df_summary.columns:
                df_summary_filtered = df_summary[df_summary['BANKS'].isin(selected_banks)]
            else:
                df_summary_filtered = df_summary.copy()
        except:
            df_summary_filtered = df_summary.copy()
        
        # Display current filters
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📊 Current Filters")
        st.sidebar.write(f"**Banks:** {', '.join(selected_banks) if selected_banks else 'None'}")
        st.sidebar.write(f"**Branch:** {selected_branch}")
        st.sidebar.write(f"**Focus Bank:** {selected_bank_radio}")
        st.sidebar.write(f"**Records:** {len(df_detailed_filtered)}")
        
        # Display date and time
        st.sidebar.markdown("---")
        st.sidebar.markdown(f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if uploaded_file:
            st.sidebar.markdown(f"**Data Source:** {uploaded_file.name}")
        else:
            st.sidebar.markdown("**Data Source:** Sample Data")
        
        # Summary metrics
        if not df_summary_filtered.empty:
            try:
                create_summary_metrics(df_summary_filtered)
            except Exception as e:
                st.warning(f"Error creating summary metrics: {str(e)}")
        
        # Create tabs based on Guarantee Type (Column D)
        if 'GUARANTEE_TYPE' in df_detailed_filtered.columns:
            guarantee_types = df_detailed_filtered['GUARANTEE_TYPE'].dropna().unique().tolist()
            
            if guarantee_types:
                st.markdown('<div class="section-header">📋 Analysis by Guarantee Type</div>', unsafe_allow_html=True)
                
                # Create tabs for each guarantee type
                tab_names = guarantee_types + ["🔄 All Types", "📊 Summary Tables"]
                tabs = st.tabs(tab_names)
                
                # Individual guarantee type tabs
                for i, guarantee_type in enumerate(guarantee_types):
                    with tabs[i]:
                        try:
                            st.subheader(f"📋 {guarantee_type} Analysis")
                            create_charts_for_guarantee_type(df_detailed_filtered, guarantee_type)
                        except Exception as e:
                            st.error(f"Error creating charts for {guarantee_type}: {str(e)}")
                
                # All types combined tab
                with tabs[len(guarantee_types)]:
                    try:
                        st.subheader("🔄 All Guarantee Types Combined")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            # Overall distribution by guarantee type
                            type_counts = df_detailed_filtered['GUARANTEE_TYPE'].value_counts()
                            fig_types = px.pie(
                                values=type_counts.values,
                                names=type_counts.index,
                                title='Distribution by Guarantee Type'
                            )
                            fig_types.update_layout(height=400, title_x=0.5)
                            st.plotly_chart(fig_types, use_container_width=True)
                        
                        with col2:
                            # Amount by guarantee type
                            if 'AMOUNT' in df_detailed_filtered.columns:
                                type_amounts = df_detailed_filtered.groupby('GUARANTEE_TYPE')['AMOUNT'].sum().sort_values(ascending=False)
                                fig_amounts = px.bar(
                                    x=type_amounts.index,
                                    y=type_amounts.values,
                                    title='Total Amount by Guarantee Type',
                                    color=type_amounts.values,
                                    color_continuous_scale='Viridis'
                                )
                                fig_amounts.update_layout(height=400, title_x=0.5, showlegend=False)
                                st.plotly_chart(fig_amounts, use_container_width=True)
                        
                        # Maturity analysis
                        if 'DAYS_TO_MATURE' in df_detailed_filtered.columns:
                            create_maturity_analysis(df_detailed_filtered)
                            
                    except Exception as e:
                        st.error(f"Error creating combined analysis: {str(e)}")
                
                # Summary tables tab
                with tabs[len(guarantee_types) + 1]:
                    try:
                        st.subheader("📊 Data Tables")
                        
                        tab1, tab2 = st.tabs(["📈 Summary Data", "📋 Detailed Transactions"])
                        
                        with tab1:
                            st.subheader("Bank Summary")
                            if not df_summary_filtered.empty:
                                st.dataframe(df_summary_filtered, use_container_width=True)
                        
                        with tab2:
                            st.subheader("Transaction Details")
                            if not df_detailed_filtered.empty:
                                st.dataframe(df_detailed_filtered, use_container_width=True)
                                
                                # Export option
                                if st.button("📥 Download Filtered Data as CSV"):
                                    csv = df_detailed_filtered.to_csv(index=False)
                                    st.download_button(
                                        label="Download CSV",
                                        data=csv,
                                        file_name=f"lg_data_filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                                        mime="text/csv"
                                    )
                    except Exception as e:
                        st.error(f"Error creating data tables: {str(e)}")
            
            else:
                st.warning("⚠️ No guarantee types found in the data.")
        
        else:
            st.warning("⚠️ GUARANTEE_TYPE column not found in the data. Please check your Excel file structure.")
            st.write("Available columns:", df_detailed_filtered.columns.tolist())
        
        # Footer
        st.markdown("---")
        st.markdown(
            "<div style='text-align: center; color: #666;'>"
            "LG Branch Summary Dashboard • Built with Streamlit • Data as of September 2025"
            "</div>",
            unsafe_allow_html=True
        )
        
    except Exception as e:
        st.error(f"Critical application error: {str(e)}")
        st.error("Please refresh the page and try again. If the problem persists, check your data file format.")
        st.stop()

if __name__ == "__main__":
    main()
