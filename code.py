import streamlit as st
import pandas as pd
import gdown
import os
from io import BytesIO
import matplotlib.pyplot as plt

# --- Konfigurasi halaman ---
st.set_page_config(page_title="Join Data Excel", page_icon="📊", layout="wide")

st.title("📂 Aplikasi Join Data Excel")
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
    st.dataframe(df_all.head(50), use_container_width=True)

    # Pastikan kolom tanggal dalam datetime
    if "Check In Date" in df_all.columns:
        df_all["Check In Date"] = pd.to_datetime(df_all["Check In Date"], errors="coerce")
    if "Check Out Date" in df_all.columns:
        df_all["Check Out Date"] = pd.to_datetime(df_all["Check Out Date"], errors="coerce")

    # --- Sidebar Filters ---
    st.sidebar.header("🔍 Filter Data")

    if "Check In Date" in df_all.columns:
        min_ci, max_ci = df_all["Check In Date"].min(), df_all["Check In Date"].max()
        checkin_range = st.sidebar.date_input("Check In Date Range", [min_ci, max_ci])
        if len(checkin_range) == 2:
            df_all = df_all[(df_all["Check In Date"] >= pd.to_datetime(checkin_range[0])) &
                            (df_all["Check In Date"] <= pd.to_datetime(checkin_range[1]))]

    if "Check Out Date" in df_all.columns:
        min_co, max_co = df_all["Check Out Date"].min(), df_all["Check Out Date"].max()
        checkout_range = st.sidebar.date_input("Check Out Date Range", [min_co, max_co])
        if len(checkout_range) == 2:
            df_all = df_all[(df_all["Check Out Date"] >= pd.to_datetime(checkout_range[0])) &
                            (df_all["Check Out Date"] <= pd.to_datetime(checkout_range[1]))]

    if "Direktorat Pekerja" in df_all.columns:
        direktorat_list = df_all["Direktorat Pekerja"].dropna().unique().tolist()
        selected_direktorat = st.sidebar.multiselect("Direktorat Pekerja", direktorat_list, default=direktorat_list)
        df_all = df_all[df_all["Direktorat Pekerja"].isin(selected_direktorat)]

    # --- Data Summary ---
    st.subheader("📊 Data Summary")
    summary = {
        "Ukuran Data (rows, cols)": df_all.shape,
        "Employee Id (unik)": df_all["Employee Id"].nunique() if "Employee Id" in df_all.columns else None,
        "Direktorat Pekerja (unik)": df_all["Direktorat Pekerja"].nunique() if "Direktorat Pekerja" in df_all.columns else None,
        "Nama Fungsi (unik)": df_all["Nama Fungsi"].nunique() if "Nama Fungsi" in df_all.columns else None,
        "Hotel Name (unik)": df_all["Hotel Name"].nunique() if "Hotel Name" in df_all.columns else None,
        "City (unik)": df_all["City"].nunique() if "City" in df_all.columns else None,
        "Country (unik)": df_all["Country"].nunique() if "Country" in df_all.columns else None,
        "Total Number of Rooms Night": df_all["Number of Rooms Night"].sum() if "Number of Rooms Night" in df_all.columns else None,
    }
    st.write(pd.DataFrame(summary, index=["Value"]).T)

    # --- Analisa Tambahan (Time Series) ---
    if "Check In Date" in df_all.columns and "Number of Rooms Night" in df_all.columns:
        st.subheader("📈 Analisa Time Series - Rooms Night per Bulan")
        df_ts = df_all.groupby(df_all["Check In Date"].dt.to_period("M"))["Number of Rooms Night"].sum().reset_index()
        df_ts["Check In Date"] = df_ts["Check In Date"].dt.to_timestamp()

        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(df_ts["Check In Date"], df_ts["Number of Rooms Night"], marker="o")
        ax.set_title("Total Rooms Night per Bulan")
        ax.set_xlabel("Bulan")
        ax.set_ylabel("Rooms Night")
        plt.xticks(rotation=45)
        st.pyplot(fig)

    # --- Download hasil ---
    st.subheader("⬇️ Download Hasil Gabungan")
    buffer_csv = BytesIO()
    df_all.to_csv(buffer_csv, index=False)
    st.download_button("Download CSV", buffer_csv.getvalue(), "gabungan.csv", "text/csv")

    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
        df_all.to_excel(writer, index=False, sheet_name="Gabungan")
    st.download_button("Download Excel", buffer_excel.getvalue(), "gabungan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
