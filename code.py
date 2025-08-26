import streamlit as st
import pandas as pd
import gdown
import os
from io import BytesIO

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

# --- Preview Data Gabungan ---
if not df_all.empty:
    st.subheader("👀 Preview Data Gabungan")
    st.dataframe(df_all, use_container_width=True)

    # --- Download hasil ---
    st.subheader("⬇️ Download Hasil Gabungan")
    buffer_csv = BytesIO()
    df_all.to_csv(buffer_csv, index=False)
    st.download_button("Download CSV", buffer_csv.getvalue(), "gabungan.csv", "text/csv")

    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
        df_all.to_excel(writer, index=False, sheet_name="Gabungan")
    st.download_button("Download Excel", buffer_excel.getvalue(), "gabungan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
