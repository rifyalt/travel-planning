import streamlit as st
import pandas as pd
import gdown
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# --- Konfigurasi halaman ---
st.set_page_config(
    page_title="Join Data Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS untuk styling ---
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Main color palette */
    :root {
        --primary-red: #fd0017;
        --primary-green: #9fe400;
        --primary-blue: #0073fe;
        --light-gray: #f8f9fa;
        --dark-gray: #343a40;
        --border-color: #dee2e6;
    }
    
    /* Global styling */
    .main {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, rgba(253, 0, 23, 0.02) 0%, rgba(159, 228, 0, 0.02) 50%, rgba(0, 115, 254, 0.02) 100%);
    }
    
    /* Header styling */
    .main-header {
        background: linear-gradient(135deg, var(--primary-red) 0%, var(--primary-blue) 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
    }
    
    .main-header h1 {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.2);
    }
    
    .main-header p {
        font-size: 1.1rem;
        opacity: 0.9;
        margin: 0;
    }
    
    /* Card styling */
    .card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        margin-bottom: 1.5rem;
        border-left: 4px solid var(--primary-green);
    }
    
    .card-header {
        font-size: 1.3rem;
        font-weight: 600;
        color: var(--dark-gray);
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, var(--primary-green) 0%, var(--primary-blue) 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-weight: 600;
        border-radius: 8px;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(159, 228, 0, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(159, 228, 0, 0.4);
    }
    
    /* Download buttons */
    .download-section {
        background: linear-gradient(135deg, rgba(0, 115, 254, 0.1) 0%, rgba(159, 228, 0, 0.1) 100%);
        padding: 1.5rem;
        border-radius: 12px;
        margin-top: 2rem;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, var(--light-gray) 0%, white 100%);
    }
    
    /* Success/Error messages */
    .stSuccess {
        background: rgba(159, 228, 0, 0.1);
        border: 1px solid var(--primary-green);
        border-radius: 8px;
    }
    
    .stError {
        background: rgba(253, 0, 23, 0.1);
        border: 1px solid var(--primary-red);
        border-radius: 8px;
    }
    
    /* Metrics styling */
    .metric-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
        gap: 1rem;
        margin: 1rem 0;
    }
    
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border-left: 3px solid var(--primary-blue);
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }
    
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: var(--primary-blue);
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: var(--dark-gray);
        margin-top: 0.25rem;
    }
    
    /* File uploader styling */
    .stFileUploader {
        border: 2px dashed var(--primary-green);
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, var(--primary-red) 0%, var(--primary-green) 50%, var(--primary-blue) 100%);
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: var(--light-gray);
        border-radius: 8px 8px 0 0;
        padding: 1rem 1.5rem;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background: var(--primary-green);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# --- Header dengan styling ---
st.markdown("""
<div class="main-header">
    <h1>📊 Aplikasi Join Data Excel</h1>
    <p>Platform modern untuk menggabungkan dan menganalisis data Excel dengan mudah</p>
</div>
""", unsafe_allow_html=True)

# --- Folder untuk menyimpan file sementara ---
os.makedirs("data_temp", exist_ok=True)

# --- Kolom yang tidak dipakai ---
drop_cols = ["Site (PSA)", "Site group Name", "Currency", "Reschedule ID", "Source_File"]

df_all = pd.DataFrame()

# --- Layout dengan columns ---
col1, col2 = st.columns([2, 1])

with col1:
    # --- Tabs untuk mode input ---
    tab1, tab2 = st.tabs(["🔗 Google Drive", "📁 Upload Manual"])
    
    with tab1:
        st.markdown("""
        <div class="card">
            <div class="card-header">
                🔗 Import dari Google Drive
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        gdrive_url = st.text_input(
            "Masukkan link Google Drive Folder (pastikan akses publik):",
            placeholder="https://drive.google.com/drive/folders/..."
        )
        
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Download & Gabungkan Data", use_container_width=True):
                if gdrive_url:
                    with st.spinner("Sedang mengunduh data dari Google Drive..."):
                        try:
                            progress_bar = st.progress(0)
                            progress_bar.progress(25)
                            
                            gdown.download_folder(url=gdrive_url, output="data_temp", quiet=False, use_cookies=False)
                            progress_bar.progress(50)
                            
                            files = [f for f in os.listdir("data_temp") if f.endswith((".xlsx", ".xls"))]
                            progress_bar.progress(75)

                            if not files:
                                st.error("❌ Tidak ada file Excel di folder Google Drive.")
                            else:
                                df_list = []
                                for i, f in enumerate(files):
                                    try:
                                        df = pd.read_excel(os.path.join("data_temp", f))
                                        df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")
                                        df_list.append(df)
                                    except Exception as e:
                                        st.warning(f"⚠️ Gagal membaca file {f}: {e}")

                                if df_list:
                                    df_all = pd.concat(df_list, ignore_index=True)
                                    progress_bar.progress(100)
                                    st.success(f"✅ Berhasil menggabungkan {len(df_list)} file Excel!")
                                    st.balloons()
                        except Exception as e:
                            st.error(f"❌ Gagal mengunduh dari Google Drive: {e}")
                else:
                    st.warning("⚠️ Silakan masukkan link Google Drive terlebih dahulu.")

    with tab2:
        st.markdown("""
        <div class="card">
            <div class="card-header">
                📁 Upload File Manual
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_files = st.file_uploader(
            "Pilih file Excel (mendukung multiple files)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Anda dapat memilih beberapa file sekaligus dengan Ctrl+Click"
        )

        if uploaded_files:
            with st.spinner("Memproses file yang diupload..."):
                df_list = []
                progress_bar = st.progress(0)
                
                for i, file in enumerate(uploaded_files):
                    try:
                        df = pd.read_excel(file)
                        df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")
                        df_list.append(df)
                        progress_bar.progress((i + 1) / len(uploaded_files))
                    except Exception as e:
                        st.warning(f"⚠️ Gagal membaca file {file.name}: {e}")

                if df_list:
                    df_all = pd.concat(df_list, ignore_index=True)
                    st.success(f"✅ Berhasil menggabungkan {len(df_list)} file!")
                    st.balloons()

with col2:
    # --- Info Panel ---
    st.markdown("""
    <div class="card">
        <div class="card-header">
            ℹ️ Panduan Penggunaan
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("""
    **Langkah-langkah:**
    1. Pilih sumber data (Google Drive atau Upload)
    2. Masukkan link/upload file Excel
    3. Klik tombol untuk memproses
    4. Gunakan filter untuk analisis
    5. Download hasil yang sudah digabungkan
    """)
    
    st.markdown("---")
    
    st.markdown("""
    <div class="card">
        <div class="card-header">
            🔧 Format File
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("""
    **Kolom yang akan dihapus otomatis:**
    - Site (PSA)
    - Site group Name
    - Currency
    - Reschedule ID
    - Source_File
    """)

# --- Preview & Analisis Data ---
if not df_all.empty:
    st.markdown("---")
    
    # --- Sidebar Filters ---
    with st.sidebar:
        st.markdown("### 🔍 Filter & Analisis Data")
        
        # Pastikan kolom tanggal dalam datetime
        if "Check In Date" in df_all.columns:
            df_all["Check In Date"] = pd.to_datetime(df_all["Check In Date"], errors="coerce")
        if "Check Out Date" in df_all.columns:
            df_all["Check Out Date"] = pd.to_datetime(df_all["Check Out Date"], errors="coerce")

        # Filter Check In Date
        if "Check In Date" in df_all.columns:
            ci_options = df_all["Check In Date"].dropna().dt.date.unique()
            ci_selected = st.selectbox("📅 Check In Date", sorted(ci_options))
            df_all = df_all[df_all["Check In Date"].dt.date == ci_selected]

        # Filter Check Out Date
        if "Check Out Date" in df_all.columns:
            co_options = df_all["Check Out Date"].dropna().dt.date.unique()
            co_selected = st.selectbox("📅 Check Out Date", sorted(co_options))
            df_all = df_all[df_all["Check Out Date"].dt.date == co_selected]

        # Filter Direktorat
        if "Direktorat Pekerja" in df_all.columns:
            direktorat_options = df_all["Direktorat Pekerja"].dropna().unique().tolist()
            direktorat_selected = st.selectbox("🏢 Direktorat Pekerja", sorted(direktorat_options))
            df_all = df_all[df_all["Direktorat Pekerja"] == direktorat_selected]

    # --- Preview Data ---
    st.markdown("""
    <div class="card">
        <div class="card-header">
            👀 Preview Data Gabungan
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    with st.expander("📋 Lihat Data (50 baris pertama)", expanded=True):
        st.dataframe(
            df_all.head(50),
            use_container_width=True,
            height=400
        )

    # --- Metrics Dashboard ---
    st.markdown("""
    <div class="card">
        <div class="card-header">
            📊 Dashboard Metrics
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "📄 Total Rows", 
            f"{df_all.shape[0]:,}",
            delta=None,
            help="Jumlah total baris data"
        )
    
    with col2:
        if "Employee Id" in df_all.columns:
            st.metric(
                "👥 Unique Employees", 
                f"{df_all['Employee Id'].nunique():,}",
                delta=None,
                help="Jumlah karyawan unik"
            )
    
    with col3:
        if "Hotel Name" in df_all.columns:
            st.metric(
                "🏨 Unique Hotels", 
                f"{df_all['Hotel Name'].nunique():,}",
                delta=None,
                help="Jumlah hotel unik"
            )
    
    with col4:
        if "Number of Rooms Night" in df_all.columns:
            st.metric(
                "🛏️ Total Room Nights", 
                f"{df_all['Number of Rooms Night'].sum():,}",
                delta=None,
                help="Total malam kamar"
            )

    # --- Visualizations ---
    if "Check In Date" in df_all.columns and "Number of Rooms Night" in df_all.columns:
        st.markdown("""
        <div class="card">
            <div class="card-header">
                📈 Analisis Time Series
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Group by month
        df_ts = df_all.groupby(df_all["Check In Date"].dt.to_period("M"))["Number of Rooms Night"].sum().reset_index()
        df_ts["Check In Date"] = df_ts["Check In Date"].dt.to_timestamp()
        
        # Create plotly chart
        fig = px.line(
            df_ts, 
            x="Check In Date", 
            y="Number of Rooms Night",
            title="Tren Room Nights per Bulan",
            color_discrete_sequence=['#0073fe']
        )
        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_family="Inter",
            title_font_size=16,
            title_font_color='#343a40'
        )
        fig.update_traces(line=dict(width=3))
        
        st.plotly_chart(fig, use_container_width=True)

    # --- Additional Analysis ---
    col1, col2 = st.columns(2)
    
    with col1:
        if "City" in df_all.columns and "Number of Rooms Night" in df_all.columns:
            st.markdown("""
            <div class="card">
                <div class="card-header">
                    🏙️ Top 10 Cities
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            city_data = df_all.groupby("City")["Number of Rooms Night"].sum().sort_values(ascending=False).head(10)
            
            fig_bar = px.bar(
                x=city_data.values,
                y=city_data.index,
                orientation='h',
                title="Room Nights by City",
                color=city_data.values,
                color_continuous_scale=['#fd0017', '#9fe400', '#0073fe']
            )
            fig_bar.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_family="Inter",
                showlegend=False,
                height=400
            )
            
            st.plotly_chart(fig_bar, use_container_width=True)
    
    with col2:
        if "Direktorat Pekerja" in df_all.columns and "Number of Rooms Night" in df_all.columns:
            st.markdown("""
            <div class="card">
                <div class="card-header">
                    🏢 Distribusi Direktorat
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            dir_data = df_all.groupby("Direktorat Pekerja")["Number of Rooms Night"].sum()
            
            fig_pie = px.pie(
                values=dir_data.values,
                names=dir_data.index,
                title="Room Nights by Direktorat",
                color_discrete_sequence=['#fd0017', '#9fe400', '#0073fe', '#ff6b35', '#f7931e']
            )
            fig_pie.update_layout(
                font_family="Inter",
                height=400
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)

    # --- Download Section ---
    st.markdown("""
    <div class="download-section">
        <div class="card-header">
            ⬇️ Download Hasil Gabungan
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    
    with col1:
        # CSV Download
        buffer_csv = BytesIO()
        df_all.to_csv(buffer_csv, index=False)
        st.download_button(
            "📄 Download CSV",
            buffer_csv.getvalue(),
            f"data_gabungan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv",
            use_container_width=True,
            help="Download dalam format CSV"
        )
    
    with col2:
        # Excel Download
        buffer_excel = BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
            df_all.to_excel(writer, index=False, sheet_name="Data_Gabungan")
        st.download_button(
            "📊 Download Excel",
            buffer_excel.getvalue(),
            f"data_gabungan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Download dalam format Excel"
        )
    
    with col3:
        # Summary Report
        summary_data = {
            "Metric": [
                "Total Rows", "Total Columns", "Unique Employees", 
                "Unique Hotels", "Unique Cities", "Total Room Nights"
            ],
            "Value": [
                df_all.shape[0], df_all.shape[1],
                df_all["Employee Id"].nunique() if "Employee Id" in df_all.columns else 0,
                df_all["Hotel Name"].nunique() if "Hotel Name" in df_all.columns else 0,
                df_all["City"].nunique() if "City" in df_all.columns else 0,
                df_all["Number of Rooms Night"].sum() if "Number of Rooms Night" in df_all.columns else 0
            ]
        }
        
        buffer_summary = BytesIO()
        pd.DataFrame(summary_data).to_csv(buffer_summary, index=False)
        st.download_button(
            "📋 Download Summary",
            buffer_summary.getvalue(),
            f"summary_report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.csv",
            "text/csv",
            use_container_width=True,
            help="Download ringkasan data"
        )

# --- Footer ---
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #6c757d; font-size: 0.9rem; padding: 1rem;">
    <p>🚀 Aplikasi Join Data Excel v2.0 | Dibuat dengan ❤️ menggunakan Streamlit</p>
</div>
""", unsafe_allow_html=True)
