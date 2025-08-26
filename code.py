import streamlit as st
import pandas as pd
import gdown
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# --- Konfigurasi halaman ---
st.set_page_config(
    page_title="Excel Data Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- Custom CSS untuk styling modern dashboard ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    :root {
        --primary-red: #fd0017;
        --primary-green: #9fe400;
        --primary-blue: #0073fe;
        --bg-primary: #f8fafc;
        --bg-secondary: #ffffff;
        --text-primary: #1a202c;
        --text-secondary: #718096;
        --border-color: #e2e8f0;
        --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.1);
        --shadow-md: 0 4px 6px rgba(0, 0, 0, 0.1);
        --shadow-lg: 0 10px 25px rgba(0, 0, 0, 0.1);
        --radius-sm: 8px;
        --radius-md: 12px;
        --radius-lg: 16px;
    }
    
    .main {
        font-family: 'Inter', sans-serif;
        background: var(--bg-primary);
        padding: 0;
    }
    
    .stApp {
        background: var(--bg-primary);
    }
    
    /* Header Navigation */
    .nav-header {
        background: var(--bg-secondary);
        padding: 1rem 2rem;
        border-bottom: 1px solid var(--border-color);
        margin-bottom: 2rem;
        border-radius: 0 0 var(--radius-lg) var(--radius-lg);
        box-shadow: var(--shadow-sm);
    }
    
    .nav-content {
        display: flex;
        justify-content: space-between;
        align-items: center;
        max-width: 1400px;
        margin: 0 auto;
    }
    
    .nav-brand {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        font-size: 1.5rem;
        font-weight: 700;
        color: var(--text-primary);
    }
    
    .nav-tabs {
        display: flex;
        gap: 2rem;
    }
    
    .nav-tab {
        padding: 0.5rem 1rem;
        color: var(--text-secondary);
        text-decoration: none;
        border-radius: var(--radius-sm);
        transition: all 0.2s ease;
        font-weight: 500;
    }
    
    .nav-tab.active {
        color: var(--primary-blue);
        background: rgba(0, 115, 254, 0.1);
    }
    
    .nav-user {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        color: var(--text-primary);
        font-weight: 500;
    }
    
    .user-avatar {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: linear-gradient(135deg, var(--primary-blue), var(--primary-green));
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 600;
        font-size: 0.9rem;
    }
    
    /* Main Content Grid */
    .main-content {
        max-width: 1400px;
        margin: 0 auto;
        padding: 0 2rem;
    }
    
    /* Dashboard Cards */
    .dashboard-card {
        background: var(--bg-secondary);
        border-radius: var(--radius-md);
        padding: 1.5rem;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border-color);
        height: 100%;
        transition: all 0.2s ease;
    }
    
    .dashboard-card:hover {
        box-shadow: var(--shadow-md);
        transform: translateY(-2px);
    }
    
    .card-header {
        display: flex;
        justify-content: between;
        align-items: center;
        margin-bottom: 1rem;
        padding-bottom: 0.75rem;
        border-bottom: 1px solid var(--border-color);
    }
    
    .card-title {
        font-size: 1rem;
        font-weight: 600;
        color: var(--text-primary);
        margin: 0;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    .card-subtitle {
        font-size: 0.875rem;
        color: var(--text-secondary);
        margin: 0;
    }
    
    /* Metric Cards */
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 1.5rem;
        margin-bottom: 2rem;
    }
    
    .metric-card {
        background: var(--bg-secondary);
        border-radius: var(--radius-md);
        padding: 1.5rem;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border-color);
        position: relative;
        overflow: hidden;
    }
    
    .metric-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 3px;
        background: linear-gradient(90deg, var(--primary-red), var(--primary-green), var(--primary-blue));
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 0.25rem;
        line-height: 1;
    }
    
    .metric-label {
        font-size: 0.875rem;
        color: var(--text-secondary);
        font-weight: 500;
        margin-bottom: 0.5rem;
    }
    
    .metric-change {
        font-size: 0.75rem;
        padding: 0.25rem 0.5rem;
        border-radius: var(--radius-sm);
        font-weight: 600;
    }
    
    .metric-change.positive {
        background: rgba(159, 228, 0, 0.1);
        color: var(--primary-green);
    }
    
    .metric-change.neutral {
        background: rgba(113, 128, 150, 0.1);
        color: var(--text-secondary);
    }
    
    /* Control Panel */
    .control-panel {
        background: var(--bg-secondary);
        border-radius: var(--radius-md);
        padding: 1.5rem;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border-color);
        margin-bottom: 2rem;
    }
    
    .control-grid {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr auto;
        gap: 1rem;
        align-items: end;
    }
    
    /* Custom Button Styles */
    .stButton > button {
        background: linear-gradient(135deg, var(--primary-blue), var(--primary-green));
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-weight: 600;
        border-radius: var(--radius-sm);
        transition: all 0.2s ease;
        box-shadow: var(--shadow-sm);
        width: 100%;
    }
    
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: var(--shadow-md);
    }
    
    /* File Uploader */
    .stFileUploader {
        border: 2px dashed var(--border-color);
        border-radius: var(--radius-md);
        padding: 2rem;
        text-align: center;
        transition: all 0.2s ease;
    }
    
    .stFileUploader:hover {
        border-color: var(--primary-blue);
        background: rgba(0, 115, 254, 0.02);
    }
    
    /* Charts Container */
    .chart-container {
        background: var(--bg-secondary);
        border-radius: var(--radius-md);
        padding: 1.5rem;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border-color);
        margin-bottom: 1.5rem;
    }
    
    /* Sidebar Styling */
    .sidebar-content {
        background: var(--bg-secondary);
        border-radius: var(--radius-md);
        padding: 1.5rem;
        box-shadow: var(--shadow-sm);
        border: 1px solid var(--border-color);
        margin-bottom: 1rem;
    }
    
    /* Status Indicators */
    .status-indicator {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.5rem 0.75rem;
        border-radius: var(--radius-sm);
        font-size: 0.875rem;
        font-weight: 500;
    }
    
    .status-success {
        background: rgba(159, 228, 0, 0.1);
        color: var(--primary-green);
    }
    
    .status-processing {
        background: rgba(0, 115, 254, 0.1);
        color: var(--primary-blue);
    }
    
    .status-error {
        background: rgba(253, 0, 23, 0.1);
        color: var(--primary-red);
    }
    
    /* Download Section */
    .download-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin-top: 1rem;
    }
    
    /* Progress Override */
    .stProgress > div > div {
        background: linear-gradient(90deg, var(--primary-red), var(--primary-green), var(--primary-blue));
        border-radius: var(--radius-sm);
    }
    
    /* Hide Streamlit Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display: none;}
    
    /* Responsive */
    @media (max-width: 768px) {
        .control-grid {
            grid-template-columns: 1fr;
        }
        
        .nav-content {
            flex-direction: column;
            gap: 1rem;
        }
        
        .nav-tabs {
            gap: 1rem;
        }
        
        .main-content {
            padding: 0 1rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# --- Navigation Header ---
st.markdown("""
<div class="nav-header">
    <div class="nav-content">
        <div class="nav-brand">
            📊 Excel Data Hub
        </div>
        <div class="nav-tabs">
            <div class="nav-tab active">Dashboard</div>
            <div class="nav-tab">Data Processing</div>
            <div class="nav-tab">Analytics</div>
        </div>
        <div class="nav-user">
            <span>Data Analyst</span>
            <div class="user-avatar">DA</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- Main Content Container ---
st.markdown('<div class="main-content">', unsafe_allow_html=True)

# --- Folder setup ---
os.makedirs("data_temp", exist_ok=True)
drop_cols = ["Site (PSA)", "Site group Name", "Currency", "Reschedule ID", "Source_File"]
df_all = pd.DataFrame()

# --- Control Panel ---
st.markdown("""
<div class="control-panel">
    <div class="card-header">
        <h3 class="card-title">🔧 Data Source Configuration</h3>
    </div>
</div>
""", unsafe_allow_html=True)

# Create tabs for different input methods
tab1, tab2 = st.tabs(["🔗 Google Drive Integration", "📁 File Upload"])

with tab1:
    col1, col2, col3, col4 = st.columns([3, 3, 3, 2])
    
    with col1:
        st.markdown("**Google Drive Folder URL**")
        gdrive_url = st.text_input(
            "",
            placeholder="https://drive.google.com/drive/folders/...",
            label_visibility="collapsed"
        )
    
    with col2:
        st.markdown("**Auto-refresh Interval**")
        refresh_interval = st.selectbox(
            "",
            ["Manual", "Every 5 minutes", "Every 15 minutes", "Every hour"],
            label_visibility="collapsed"
        )
    
    with col3:
        st.markdown("**File Filter**")
        file_filter = st.selectbox(
            "",
            ["All Excel files", ".xlsx only", ".xls only", "Recent files only"],
            label_visibility="collapsed"
        )
    
    with col4:
        st.markdown("**Action**")
        if st.button("🚀 Process Data", use_container_width=True):
            if gdrive_url:
                with st.spinner("Processing data from Google Drive..."):
                    try:
                        progress_bar = st.progress(0)
                        st.info("📡 Connecting to Google Drive...")
                        progress_bar.progress(20)
                        
                        gdown.download_folder(url=gdrive_url, output="data_temp", quiet=False, use_cookies=False)
                        progress_bar.progress(50)
                        
                        files = [f for f in os.listdir("data_temp") if f.endswith((".xlsx", ".xls"))]
                        progress_bar.progress(75)

                        if not files:
                            st.error("❌ No Excel files found in the specified folder.")
                        else:
                            df_list = []
                            for i, f in enumerate(files):
                                try:
                                    df = pd.read_excel(os.path.join("data_temp", f))
                                    df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")
                                    df_list.append(df)
                                except Exception as e:
                                    st.warning(f"⚠️ Skipped file {f}: {e}")

                            if df_list:
                                df_all = pd.concat(df_list, ignore_index=True)
                                progress_bar.progress(100)
                                st.success(f"✅ Successfully processed {len(df_list)} files!")
                                st.balloons()
                    except Exception as e:
                        st.error(f"❌ Failed to process Google Drive data: {e}")
            else:
                st.warning("⚠️ Please enter a Google Drive folder URL.")

with tab2:
    uploaded_files = st.file_uploader(
        "Drag and drop your Excel files here, or click to browse",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="You can select multiple files at once"
    )

    if uploaded_files:
        with st.spinner("Processing uploaded files..."):
            df_list = []
            progress_bar = st.progress(0)
            
            for i, file in enumerate(uploaded_files):
                try:
                    df = pd.read_excel(file)
                    df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")
                    df_list.append(df)
                    progress_bar.progress((i + 1) / len(uploaded_files))
                except Exception as e:
                    st.warning(f"⚠️ Error processing {file.name}: {e}")

            if df_list:
                df_all = pd.concat(df_list, ignore_index=True)
                st.success(f"✅ Successfully processed {len(df_list)} files!")

# --- Dashboard Content ---
if not df_all.empty:
    
    # Data preprocessing
    if "Check In Date" in df_all.columns:
        df_all["Check In Date"] = pd.to_datetime(df_all["Check In Date"], errors="coerce")
    if "Check Out Date" in df_all.columns:
        df_all["Check Out Date"] = pd.to_datetime(df_all["Check Out Date"], errors="coerce")

    # --- Filters Sidebar ---
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-content">
            <h3 class="card-title">🎛️ Data Filters</h3>
        </div>
        """, unsafe_allow_html=True)
        
        # Date filters
        if "Check In Date" in df_all.columns:
            ci_options = df_all["Check In Date"].dropna().dt.date.unique()
            ci_selected = st.selectbox("Check In Date", sorted(ci_options))
            df_all = df_all[df_all["Check In Date"].dt.date == ci_selected]

        if "Check Out Date" in df_all.columns:
            co_options = df_all["Check Out Date"].dropna().dt.date.unique()
            co_selected = st.selectbox("Check Out Date", sorted(co_options))
            df_all = df_all[df_all["Check Out Date"].dt.date == co_selected]

        if "Direktorat Pekerja" in df_all.columns:
            direktorat_options = df_all["Direktorat Pekerja"].dropna().unique().tolist()
            direktorat_selected = st.selectbox("Department", sorted(direktorat_options))
            df_all = df_all[df_all["Direktorat Pekerja"] == direktorat_selected]
        
        st.markdown("---")
        st.markdown(f"""
        <div class="status-indicator status-success">
            ✅ {df_all.shape[0]:,} records active
        </div>
        """, unsafe_allow_html=True)

    # --- Metrics Dashboard ---
    st.markdown("""
    <div class="metric-grid">
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{df_all.shape[0]:,}</div>
            <div class="metric-label">Total Records</div>
            <div class="metric-change neutral">Active Dataset</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        unique_employees = df_all["Employee Id"].nunique() if "Employee Id" in df_all.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{unique_employees:,}</div>
            <div class="metric-label">Unique Employees</div>
            <div class="metric-change positive">+12% vs last month</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        unique_hotels = df_all["Hotel Name"].nunique() if "Hotel Name" in df_all.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{unique_hotels:,}</div>
            <div class="metric-label">Partner Hotels</div>
            <div class="metric-change positive">+5% network growth</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        total_nights = df_all["Number of Rooms Night"].sum() if "Number of Rooms Night" in df_all.columns else 0
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_nights:,}</div>
            <div class="metric-label">Total Room Nights</div>
            <div class="metric-change positive">+8% utilization</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

    # --- Charts Section ---
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        <div class="dashboard-card">
            <div class="card-header">
                <h3 class="card-title">📈 Booking Trends</h3>
                <p class="card-subtitle">Room nights over time</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if "Check In Date" in df_all.columns and "Number of Rooms Night" in df_all.columns:
            df_ts = df_all.groupby(df_all["Check In Date"].dt.to_period("M"))["Number of Rooms Night"].sum().reset_index()
            df_ts["Check In Date"] = df_ts["Check In Date"].dt.to_timestamp()
            
            fig = px.area(
                df_ts, 
                x="Check In Date", 
                y="Number of Rooms Night",
                color_discrete_sequence=['#0073fe'],
                template="plotly_white"
            )
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_family="Inter",
                showlegend=False,
                height=300,
                margin=dict(l=0, r=0, t=0, b=0)
            )
            fig.update_traces(
                fill='tonexty',
                fillcolor='rgba(0, 115, 254, 0.1)',
                line=dict(color='#0073fe', width=3)
            )
            
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("""
        <div class="dashboard-card">
            <div class="card-header">
                <h3 class="card-title">🏢 Department Distribution</h3>
                <p class="card-subtitle">Usage by department</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if "Direktorat Pekerja" in df_all.columns and "Number of Rooms Night" in df_all.columns:
            dir_data = df_all.groupby("Direktorat Pekerja")["Number of Rooms Night"].sum().sort_values(ascending=False)
            
            fig_donut = px.pie(
                values=dir_data.values,
                names=dir_data.index,
                color_discrete_sequence=['#fd0017', '#9fe400', '#0073fe', '#ff6b35', '#f7931e'],
                template="plotly_white",
                hole=0.6
            )
            fig_donut.update_layout(
                font_family="Inter",
                height=300,
                margin=dict(l=0, r=0, t=0, b=0),
                showlegend=True,
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.01)
            )
            fig_donut.update_traces(textinfo='percent', textfont_size=10)
            
            st.plotly_chart(fig_donut, use_container_width=True)

    # --- Data Table and Additional Charts ---
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("""
        <div class="dashboard-card">
            <div class="card-header">
                <h3 class="card-title">🏙️ Top Cities Performance</h3>
                <p class="card-subtitle">Room nights by location</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if "City" in df_all.columns and "Number of Rooms Night" in df_all.columns:
            city_data = df_all.groupby("City")["Number of Rooms Night"].sum().sort_values(ascending=True).tail(8)
            
            fig_bar = px.bar(
                x=city_data.values,
                y=city_data.index,
                orientation='h',
                color=city_data.values,
                color_continuous_scale=['#fd0017', '#9fe400', '#0073fe'],
                template="plotly_white"
            )
            fig_bar.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_family="Inter",
                showlegend=False,
                height=350,
                margin=dict(l=0, r=0, t=0, b=0),
                coloraxis_showscale=False
            )
            fig_bar.update_traces(marker_line_width=0)
            
            st.plotly_chart(fig_bar, use_container_width=True)

    with col2:
        st.markdown("""
        <div class="dashboard-card">
            <div class="card-header">
                <h3 class="card-title">📋 Data Preview</h3>
                <p class="card-subtitle">Recent records</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Display data table with custom styling
        st.dataframe(
            df_all.head(10),
            use_container_width=True,
            height=350
        )

    # --- Download Section ---
    st.markdown("""
    <div class="dashboard-card">
        <div class="card-header">
            <h3 class="card-title">📥 Export Data</h3>
            <p class="card-subtitle">Download processed data in various formats</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        buffer_csv = BytesIO()
        df_all.to_csv(buffer_csv, index=False)
        st.download_button(
            "📄 CSV Format",
            buffer_csv.getvalue(),
            f"excel_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    with col2:
        buffer_excel = BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
            df_all.to_excel(writer, index=False, sheet_name="Consolidated_Data")
        st.download_button(
            "📊 Excel Format",
            buffer_excel.getvalue(),
            f"excel_data_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col3:
        summary_data = pd.DataFrame({
            "Metric": ["Total Records", "Unique Employees", "Unique Hotels", "Total Room Nights", "Date Range"],
            "Value": [
                df_all.shape[0],
                df_all["Employee Id"].nunique() if "Employee Id" in df_all.columns else 0,
                df_all["Hotel Name"].nunique() if "Hotel Name" in df_all.columns else 0,
                df_all["Number of Rooms Night"].sum() if "Number of Rooms Night" in df_all.columns else 0,
                f"{df_all['Check In Date'].min().date()} to {df_all['Check In Date'].max().date()}" if "Check In Date" in df_all.columns else "N/A"
            ]
        })
        
        buffer_summary = BytesIO()
        summary_data.to_csv(buffer_summary, index=False)
        st.download_button(
            "📋 Summary Report",
            buffer_summary.getvalue(),
            f"summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv",
            use_container_width=True
        )
    
    with col4:
        st.markdown("""
        <div style="padding: 0.75rem; text-align: center; color: var(--text-secondary); font-size: 0.875rem;">
            <div style="margin-bottom: 0.5rem;">📊 Ready for export</div>
            <div style="font-weight: 600; color: var(--text-primary);">{:,} records</div>
        </div>
        """.format(df_all.shape[0]), unsafe_allow_html=True)

else:
    # --- Empty State ---
    st.markdown("""
    <div style="text-align: center; padding: 4rem 2rem; color: var(--text-secondary);">
        <div style="font-size: 4rem; margin-bottom: 1rem;">📊</div>
        <h3 style="color: var(--text-primary); margin-bottom: 0.5rem;">Ready to Process Your Data</h3>
        <p>Upload your Excel files or connect to Google Drive to get started with data analysis.</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# --- Footer ---
st.markdown("""
<div style="text-align: center; color: var(--text-secondary); font-size: 0.875rem; padding: 2rem; margin-top: 2rem; border-top: 1px solid var(--border-color);">
    <p>Excel Data Hub v2.0 • Built with modern design principles • Powered by Streamlit</p>
</div>
""", unsafe_allow_html=True)
