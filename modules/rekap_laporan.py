import streamlit as st
import pandas as pd
from modules import kategori, petugas, wali_hakim, wna

# --- KONFIGURASI NAMA BULAN ---
MONTH_MAP = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

st.set_page_config(page_title="App Laporan NR", layout="wide")
st.title("🏛️ App Laporan NR - KUA Tangerang")

uploaded_file = st.file_uploader("Upload Data Master (.xlsx)", type=["xlsx"])

if uploaded_file:
    # --- LOAD DATA & CLEANING ---
    df_raw = pd.read_excel(uploaded_file, dtype=str)
    df_raw.columns = df_raw.columns.str.strip()
    df = df_raw.apply(lambda x: x.str.strip().str.upper() if x.dtype == "object" else x).fillna('')
    
    try:
        df_temp = df.copy()
        df_temp['dt_temp'] = pd.to_datetime(df_temp['Tanggal Akad'], dayfirst=True, errors='coerce')
        valid_dates = df_temp['dt_temp'].dropna()
        nama_bulan = MONTH_MAP.get(valid_dates.dt.month.mode()[0], "JANUARI")
        year_val = valid_dates.dt.year.mode()[0]
    except:
        nama_bulan = "JANUARI"; year_val = "2025"

    st.success(f"✅ Data Terdeteksi: {nama_bulan} {year_val}")

    # --- MENU RUANG TERPISAH (SIDEBAR) ---
    menu = st.sidebar.selectbox("Pilih Ruang Laporan:", [
        "1. Laporan Kategori (Luar/Kantor/Isbat)",
        "2. Laporan Per Petugas",
        "3. Laporan Wali Hakim",
        "4. Laporan WNA"
    ])

    # Oper data ke modul masing-masing
    if "1. Laporan Kategori" in menu:
        kategori.render(df, nama_bulan, year_val)
    elif "2. Laporan Per Petugas" in menu:
        petugas.render(df, nama_bulan, year_val)
    elif "3. Laporan Wali Hakim" in menu:
        wali_hakim.render(df, nama_bulan, year_val)
    elif "4. Laporan WNA" in menu:
        wna.render(df, nama_bulan, year_val)