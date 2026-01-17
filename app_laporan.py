import streamlit as st
import pandas as pd
from io import BytesIO
from modules import kategori, petugas, wali_hakim, wna, pnbp

# 1. KONFIGURASI DASAR
MONTH_MAP = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
    7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

st.set_page_config(page_title="App Laporan NR", layout="wide")

# --- HTML & CSS CUSTOM (MACOS GLASSMORPHISM) ---
st.markdown("""
    <style>
    /* Background Utama */
    .stApp {
        background: linear-gradient(135deg, #fce4ec 0%, #f3e5f5 50%, #e1bee7 100%);
    }

    /* Container Kartu */
    .macos-card-container {
        display: flex; gap: 20px; margin-bottom: 25px;
    }

    .macos-card {
        background: rgba(255, 255, 255, 0.85);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 20px; flex: 1;
        box-shadow: 0 10px 25px rgba(156, 39, 176, 0.15);
        transition: all 0.3s ease;
    }

    .macos-card h4 { margin: 0; color: #7b1fa2; font-size: 0.85rem; font-weight: 700; text-transform: uppercase; }
    .macos-card h2 { 
        margin: 10px 0 0 0; 
        background: -webkit-linear-gradient(#e91e63, #9c27b0);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.3rem; font-weight: 800;
    }

    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] { background: rgba(255, 255, 255, 0.6); padding: 8px; border-radius: 18px; }
    .stTabs [data-baseweb="tab"] { color: #9c27b0 !important; font-weight: 700; }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #e91e63, #9c27b0) !important;
        color: white !important; border-radius: 12px;
    }

    /* --- FIX HEADER TABEL UNGU TEKS PUTIH --- */
    /* Selector untuk Header Dataframe */
    [data-testid="stTableThead"] {
        background-color: #9c27b0 !important;
    }
    
    /* Memaksa warna header tabel */
    .stDataFrame div[data-testid="stTable"] thead tr th {
        background-color: #9c27b0 !important;
        color: white !important;
    }

    /* Styling Header Kolom (Streamlit New Version) */
    div[data-testid="stDataFrameColHeader"] {
        background-color: #9c27b0 !important;
        color: white !important;
    }

    /* Memastikan teks header berwarna putih */
    div[data-testid="stHeader"] {
        background-color: #9c27b0 !important;
        color: white !important;
    }

    /* Tombol Cetak */
    div.stButton > button:first-child {
        background: linear-gradient(90deg, #f06292, #ba68c8);
        color: white; border: none; border-radius: 12px;
        font-weight: bold; padding: 0.6rem 1.5rem;
    }
    </style>
    """, unsafe_allow_html=True)

def cari_kolom(df, keywords):
    for k in keywords:
        for c in df.columns:
            if k.upper() == str(c).upper().strip(): return c
    for k in keywords:
        for c in df.columns:
            if k.upper() in str(c).upper(): return c
    return None

def show_rekap_total(df, bulan, tahun):
    # --- RENDER KARTU VIA HTML ---
    col_lok_asli = cari_kolom(df, ["NIKAH DI", "LOKASI"])
    total_val = len(df)
    k_val = len(df[(df[col_lok_asli].str.contains("KANTOR|KUA", na=False)) & (~df[col_lok_asli].str.contains("LUAR", na=False))]) if col_lok_asli else 0
    lk_val = len(df[df[col_lok_asli].str.contains("LUAR", na=False)]) if col_lok_asli else 0
    isbat_val = len(df[df[col_lok_asli].str.contains("ISBAT", na=False)]) if col_lok_asli else 0

    st.markdown(f"""
        <div class="macos-card-container">
            <div class="macos-card">
                <h4>Total Peristiwa</h4>
                <h2>{total_val}</h2>
            </div>
            <div class="macos-card">
                <h4>Nikah di Kantor</h4>
                <h2>{k_val}</h2>
            </div>
            <div class="macos-card">
                <h4>Nikah di Luar</h4>
                <h2>{lk_val}</h2>
            </div>
            <div class="macos-card">
                <h4>Isbat Nikah</h4>
                <h2>{isbat_val}</h2>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # --- TABEL KERJA (MAPPING ASLI LO) ---
    df_fix = df.copy()
    df_fix['No_Urut'] = range(1, len(df_fix) + 1)
    c_seri = cari_kolom(df, ["NO SERI HURUF", "SERI HURUF"])
    c_perfo = cari_kolom(df, ["NO PERFORASI", "PERFORASI"])
    df_fix['P_Perforasi'] = (df_fix[c_seri].astype(str) + ". " + df_fix[c_perfo].astype(str)) if (c_seri and c_perfo) else ""
    df_fix['Blank'] = ""

    mapping_all = {
        'No': 'No_Urut', 'No Perforasi': 'P_Perforasi', ' ': 'Blank',
        'No Pemeriksaan': ["NO PEMERIKSAAN", "PEMERIKSAAN"],
        'No Aktanikah': ["NO AKTANIKAH", "AKTA NIKAH"],
        'Nama Suami': ["NAMA SUAMI"], 'Nama Istri': ["NAMA ISTRI"],
        'Tanggal Akad': ["TANGGAL AKAD"], 'Nikah Di': ["NIKAH DI", "LOKASI"]
    }

    rekap_df = pd.DataFrame()
    for label, keys in mapping_all.items():
        if isinstance(keys, list):
            col_found = cari_kolom(df, keys)
            rekap_df[label] = df_fix[col_found] if col_found else ""
        else:
            rekap_df[label] = df_fix[keys]

    st.markdown("<div style='background: white; padding: 20px; border-radius: 15px;'>", unsafe_allow_html=True)
    st.dataframe(rekap_df, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🚀 Cetak Excel Rekap Semua"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            rekap_df.to_excel(writer, index=False, sheet_name='Semua_Data', startrow=4)
            from modules.petugas import format_excel
            format_excel(writer, rekap_df, "REKAP SELURUH DATA", bulan, tahun)
        st.download_button(label="📥 Ambil File Excel", data=output.getvalue(), file_name=f"REKAP_{bulan}.xlsx")

# --- MAIN UI ---
st.markdown("<h2 style='text-align: center; color: #1d1d1f; font-weight: 700;'>🏛️ LAPORAN BULANAN & PNBP KUA TANGERANG</h2>", unsafe_allow_html=True)

# Upload Area
uploaded_file = st.file_uploader("", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, dtype=str)
    df_raw.columns = df_raw.columns.str.strip()
    df = df_raw.apply(lambda x: x.str.strip().str.upper() if x.dtype == "object" else x).fillna('')

    col_tgl = cari_kolom(df, ["TANGGAL AKAD", "TGL AKAD"])
    try:
        df_temp = df.copy()
        df_temp['dt_temp'] = pd.to_datetime(df_temp[col_tgl], dayfirst=True, errors='coerce')
        valid_dates = df_temp['dt_temp'].dropna()
        nama_bulan = MONTH_MAP.get(valid_dates.dt.month.mode()[0], "JANUARI")
        year_val = int(valid_dates.dt.year.mode()[0])
    except:
        nama_bulan = "JANUARI"; year_val = "2026"

    st.markdown(f"<p style='text-align: center; color: #007aff;'><b>Periode: {nama_bulan} {year_val}</b></p>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📊 Overview", "📁 Kategori", "👨‍💼 Petugas"])

    with tab1:
        show_rekap_total(df, nama_bulan, year_val)

    with tab2:
        sub = st.radio("Pilih Laporan:", ["UTAMA", "WALI HAKIM", "WNA", "PNBP"], horizontal=True)
        if sub == "UTAMA": kategori.render(df, nama_bulan, year_val)
        elif sub == "WALI HAKIM": wali_hakim.render(df, nama_bulan, year_val)
        elif sub == "WNA": wna.render(df, nama_bulan, year_val)
        elif sub == "PNBP": pnbp.render(df, nama_bulan, year_val)

    with tab3:
        petugas.render(df, nama_bulan, year_val)
else:
    st.info("👋 Unggah file Excel untuk memproses data.")