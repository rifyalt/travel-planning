import streamlit as st
import pandas as pd
import gdown
import os
from io import BytesIO
import plotly.express as px

# --- Konfigurasi halaman ---
st.set_page_config(page_title="Join Data Excel", page_icon="📊", layout="wide")

st.title("📂 Aplikasi Join Data Excel")
st.write("Pilih salah satu metode input data:")

# --- Sidebar opsi ---
mode = st.sidebar.radio("Pilih sumber data:", ["Google Drive Folder", "Upload Manual"])

# --- Folder untuk menyimpan file sementara ---
os.makedirs("data_temp", exist_ok=True)

# --- Kolom yang tidak dipakai ---
drop_cols = ["No Trip SAP","Cost Center Pekerja","Site (PSA)", "Site group Name", "Currency", "Reschedule ID", "Source_File"]

df_all = pd.DataFrame()

# --- Mode Google Drive ---
if mode == "Google Drive Folder":
    gdrive_url = st.text_input("Masukkan link Google Drive Folder (publik):")

    if st.button("Download & Gabungkan Data"):
        if gdrive_url:
            try:
                gdown.download_folder(url=gdrive_url, output="data_temp", quiet=False, use_cookies=False)
                files = [f for f in os.listdir("data_temp") if f.endswith((".xlsx", ".xls"))]

                if not files:
                    st.error("Tidak ada file Excel di folder Google Drive.")
                else:
                    df_list = []
                    for f in files:
                        try:
                            df = pd.read_excel(os.path.join("data_temp", f))
                            df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")
                            df_list.append(df)
                        except Exception as e:
                            st.warning(f"Gagal baca file {f}: {e}")

                    if df_list:
                        df_all = pd.concat(df_list, ignore_index=True)
                        st.success(f"Berhasil gabungkan {len(df_list)} file.")
            except Exception as e:
                st.error(f"Gagal download dari Google Drive: {e}")

# --- Mode Upload Manual ---
elif mode == "Upload Manual":
    uploaded_files = st.file_uploader("Upload file Excel (bisa multi upload)", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        df_list = []
        for file in uploaded_files:
            try:
                df = pd.read_excel(file)
                df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")
                df_list.append(df)
            except Exception as e:
                st.warning(f"Gagal baca file {file.name}: {e}")

        if df_list:
            df_all = pd.concat(df_list, ignore_index=True)
            st.success(f"Berhasil gabungkan {len(df_list)} file.")

# --- Preview & Analisis Data ---
if not df_all.empty:
    st.subheader("👀 Preview Data Gabungan")
    st.dataframe(df_all, use_container_width=True)

    # --- Konversi tanggal dengan dayfirst ---
    date_columns = ["Check In Date", "Check Out Date"]
    for col in date_columns:
        if col in df_all.columns:
            df_all[col] = pd.to_datetime(df_all[col], errors="coerce", dayfirst=True)

    # --- Sidebar Filters ---
    st.sidebar.header("🔍 Filter Data")

    # Check In Date filter
    if "Check In Date" in df_all.columns and df_all["Check In Date"].notna().any():
        min_ci, max_ci = df_all["Check In Date"].min(), df_all["Check In Date"].max()
        ci_selected = st.sidebar.date_input("Pilih Rentang Check In Date", [min_ci, max_ci])
        if isinstance(ci_selected, list) and len(ci_selected) == 2:
            df_all = df_all[(df_all["Check In Date"] >= pd.to_datetime(ci_selected[0])) &
                            (df_all["Check In Date"] <= pd.to_datetime(ci_selected[1]))]

    # Check Out Date filter
    if "Check Out Date" in df_all.columns and df_all["Check Out Date"].notna().any():
        min_co, max_co = df_all["Check Out Date"].min(), df_all["Check Out Date"].max()
        co_selected = st.sidebar.date_input("Pilih Rentang Check Out Date", [min_co, max_co])
        if isinstance(co_selected, list) and len(co_selected) == 2:
            df_all = df_all[(df_all["Check Out Date"] >= pd.to_datetime(co_selected[0])) &
                            (df_all["Check Out Date"] <= pd.to_datetime(co_selected[1]))]

    # Direktorat filter
    if "Direktorat Pekerja" in df_all.columns:
        direktorat_options = df_all["Direktorat Pekerja"].dropna().unique().tolist()
        direktorat_selected = st.sidebar.selectbox("Pilih Direktorat Pekerja", ["All"] + sorted(direktorat_options))
        if direktorat_selected != "All":
            df_all = df_all[df_all["Direktorat Pekerja"] == direktorat_selected]

    # --- Data Summary ---
    
    st.subheader("📊 Data Summary")

    summary_list = [
        {"Metric": "Ukuran Data", "Value": f"{df_all.shape[0]} rows, {df_all.shape[1]} cols"},
        {"Metric": "Employee Id (unik)", "Value": df_all["Employee Id"].nunique() if "Employee Id" in df_all.columns else 0},
        {"Metric": "Direktorat Pekerja (unik)", "Value": df_all["Direktorat Pekerja"].nunique() if "Direktorat Pekerja" in df_all.columns else 0},
        {"Metric": "Nama Fungsi (unik)", "Value": df_all["Nama Fungsi"].nunique() if "Nama Fungsi" in df_all.columns else 0},
        {"Metric": "Hotel Name (unik)", "Value": df_all["Hotel Name"].nunique() if "Hotel Name" in df_all.columns else 0},
        {"Metric": "City (unik)", "Value": df_all["City"].nunique() if "City" in df_all.columns else 0},
        {"Metric": "Country (unik)", "Value": df_all["Country"].nunique() if "Country" in df_all.columns else 0},
        {"Metric": "Total Rooms Night", "Value": int(df_all["Number of Rooms Night"].sum()) if "Number of Rooms Night" in df_all.columns else 0},
    ]

    # Grid kolom untuk score cards
    cols = st.columns(4)  # tampil 4 per baris
    for i, item in enumerate(summary_list):
        with cols[i % 4]:
            st.markdown(
                f"""
                <div style="
                    background: linear-gradient(135deg, #4F46E5, #3B82F6);
                    padding: 15px;
                    border-radius: 15px;
                    text-align: center;
                    color: white;
                    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                    margin-bottom: 15px;">
                    <h4 style="margin: 0; font-size: 16px;">{item['Metric']}</h4>
                    <p style="margin: 5px 0 0; font-size: 22px; font-weight: bold;">{item['Value']}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

    # --- Analisa Tambahan (Time Series) ---
    if "Check In Date" in df_all.columns and "Number of Rooms Night" in df_all.columns:
        st.subheader("📈 Analisa Time Series - Rooms Night per Bulan")
        df_ts = df_all.groupby(df_all["Check In Date"].dt.to_period("M"))["Number of Rooms Night"].sum().reset_index()
        df_ts["Check In Date"] = df_ts["Check In Date"].dt.to_timestamp()

        fig = px.line(
            df_ts,
            x="Check In Date",
            y="Number of Rooms Night",
            title="Total Rooms Night per Bulan",
            markers=True
        )
        fig.update_layout(xaxis_title="Bulan", yaxis_title="Rooms Night")
        st.plotly_chart(fig, use_container_width=True)

    # --- Download hasil ---
    st.subheader("⬇️ Download Hasil Gabungan")
    buffer_csv = BytesIO()
    df_all.to_csv(buffer_csv, index=False)
    st.download_button("Download CSV", buffer_csv.getvalue(), "gabungan.csv", "text/csv")

    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
        df_all.to_excel(writer, index=False, sheet_name="Gabungan")
    st.download_button("Download Excel", buffer_excel.getvalue(), "gabungan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



