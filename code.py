import streamlit as st
import pandas as pd
import gdown
import os
from io import BytesIO
import plotly.express as px

# --- Konfigurasi halaman ---
st.set_page_config(page_title="Join & Analisis Data Excel", page_icon="📊", layout="wide")

st.title("📂 Aplikasi Gabung & Analisis Data Excel")
st.write("Pilih salah satu metode input data:")

# --- Sidebar opsi ---
mode = st.sidebar.radio("Pilih sumber data:", ["Google Drive Folder", "Upload Manual"])

# --- Folder untuk menyimpan file sementara ---
os.makedirs("data_temp", exist_ok=True)

# --- Kolom yang tidak dipakai ---
drop_cols = ["Site (PSA)", "Site group Name", "Currency", "Reschedule ID", "Source_File"]

df_all = pd.DataFrame()

# --- Mode Google Drive ---
if mode == "Google Drive Folder":
    gdrive_url = st.text_input("Masukkan link Google Drive Folder (publik):")

    if st.button("Download & Gabungkan Data"):
        if gdrive_url:
            try:
                # Download semua file dari folder
                gdown.download_folder(url=gdrive_url, output="data_temp", quiet=False, use_cookies=False)

                files = [f for f in os.listdir("data_temp") if f.endswith((".xlsx", ".xls"))]

                if not files:
                    st.error("Tidak ada file Excel di folder Google Drive.")
                else:
                    df_list = []
                    for f in files:
                        try:
                            df = pd.read_excel(os.path.join("data_temp", f))
                            # Drop kolom tidak dipakai jika ada
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

# --- Filter dan Analisis Data ---
if not df_all.empty:
    st.sidebar.subheader("🔎 Filter Data")

    # Konversi kolom tanggal jika ada
    try:
        if 'Check in Date' in df_all.columns:
            df_all['Check in Date'] = pd.to_datetime(df_all['Check in Date'])
        if 'Check Out Date' in df_all.columns:
            df_all['Check Out Date'] = pd.to_datetime(df_all['Check Out Date'])
    except Exception as e:
        st.warning(f"Gagal mengonversi kolom tanggal: {e}")

    # --- Filter Check in Date ---
    if 'Check in Date' in df_all.columns:
        min_date = df_all['Check in Date'].min().date()
        max_date = df_all['Check in Date'].max().date()
        
        start_date, end_date = st.sidebar.date_input(
            "Filter Check in Date:",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        df_filtered = df_all[(df_all['Check in Date'].dt.date >= start_date) & (df_all['Check in Date'].dt.date <= end_date)]
    else:
        df_filtered = df_all.copy()

    # --- Filter Direktorat Pekerja ---
    if 'Direktorat Pekerja' in df_all.columns:
        direktorat_options = df_all['Direktorat Pekerja'].unique().tolist()
        selected_direktorat = st.sidebar.multiselect(
            "Filter Direktorat Pekerja:",
            options=direktorat_options,
            default=direktorat_options
        )
        df_filtered = df_filtered[df_filtered['Direktorat Pekerja'].isin(selected_direktorat)]

    st.subheader("👀 Preview Data Gabungan")
    st.dataframe(df_filtered, use_container_width=True)

    # --- Ringkasan Data ---
    st.subheader("📊 Ringkasan Data")
    
    col1, col2, col3, col4 = st.columns(4)
    
    # Ukuran Data
    with col1:
        st.metric("Ukuran Data (Rows)", df_filtered.shape[0])
    
    # Employee ID Unik
    if 'Employee ID' in df_filtered.columns:
        with col2:
            st.metric("Employee ID Unik", df_filtered['Employee ID'].nunique())
    
    # Direktorat Pekerja Unik
    if 'Direktorat Pekerja' in df_filtered.columns:
        with col3:
            st.metric("Direktorat Pekerja Unik", df_filtered['Direktorat Pekerja'].nunique())
    
    # Nama Fungsi Unik
    if 'Nama Fungsi' in df_filtered.columns:
        with col4:
            st.metric("Nama Fungsi Unik", df_filtered['Nama Fungsi'].nunique())
    
    col5, col6, col7, col8 = st.columns(4)

    # Hotel Name Unik
    if 'Hotel Name' in df_filtered.columns:
        with col5:
            st.metric("Hotel Name Unik", df_filtered['Hotel Name'].nunique())

    # City Unik
    if 'City' in df_filtered.columns:
        with col6:
            st.metric("City Unik", df_filtered['City'].nunique())
    
    # Country Unik
    if 'Country' in df_filtered.columns:
        with col7:
            st.metric("Country Unik", df_filtered['Country'].nunique())
    
    # Total Rooms Night
    if 'Rooms Night' in df_filtered.columns:
        with col8:
            st.metric("Total Rooms Night", int(df_filtered['Rooms Night'].sum()))
    
    # --- Analisis Time Series ---
    st.subheader("📈 Analisis Time Series: Total Rooms Night per Bulan")
    if 'Check in Date' in df_filtered.columns and 'Rooms Night' in df_filtered.columns:
        df_ts = df_filtered.copy()
        df_ts['Tahun-Bulan'] = df_ts['Check in Date'].dt.to_period('M').astype(str)
        df_monthly = df_ts.groupby('Tahun-Bulan')['Rooms Night'].sum().reset_index()
        
        fig = px.line(
            df_monthly, 
            x='Tahun-Bulan', 
            y='Rooms Night', 
            title='Jumlah Rooms Night dari Waktu ke Waktu',
            markers=True
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Kolom 'Check in Date' dan 'Rooms Night' tidak ditemukan untuk analisis time series.")

    # --- Download hasil ---
    st.subheader("⬇️ Download Hasil Gabungan")
    
    # Download CSV
    buffer_csv = BytesIO()
    df_filtered.to_csv(buffer_csv, index=False)
    st.download_button("Download CSV", buffer_csv.getvalue(), "gabungan.csv", "text/csv")

    # Download Excel
    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
        df_filtered.to_excel(writer, index=False, sheet_name="Gabungan")
    st.download_button("Download Excel", buffer_excel.getvalue(), "gabungan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")