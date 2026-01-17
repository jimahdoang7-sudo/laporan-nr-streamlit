import streamlit as st
import pandas as pd
from io import BytesIO

# Map Bulan Indonesia
MONTH_MAP = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
}

st.set_page_config(page_title="Mesin Rekap KUA Pro", layout="wide")
st.title("🏛️ Mesin Rekap Laporan KUA Tangerang")

uploaded_file = st.file_uploader("Upload Data Master (.xlsx)", type=["xlsx"])

if uploaded_file:
    # 1. LOAD DATA (TETAP PAKAI PROTEKSI STRING & FILLNA)
    df = pd.read_excel(uploaded_file, dtype=str).fillna('')
    
    # --- DETEKSI BULAN (KODE ASLI) ---
    try:
        df_temp = df.copy()
        df_temp['dt'] = pd.to_datetime(df_temp['Tanggal Akad'], dayfirst=True, errors='coerce')
        valid_dates = df_temp['dt'].dropna()
        if not valid_dates.empty:
            month_idx = valid_dates.dt.month.mode()[0]
            year_val = valid_dates.dt.year.mode()[0]
            nama_bulan = MONTH_MAP.get(month_idx, "JANUARI")
        else:
            nama_bulan = "JANUARI"; year_val = "2025"
    except:
        nama_bulan = "JANUARI"; year_val = "2025"

    st.success(f"✅ Data Terdeteksi: Bulan {nama_bulan} {year_val}")

    # --- PILIHAN MODE REKAP (TIDAK MERUBAH LOGIKA UTAMA) ---
    st.markdown("### 🛠️ Pilih Mode Laporan")
    mode_rekap = st.selectbox("Mau rekap berdasarkan apa?", ["Kategori (Kantor/Luar/Isbat)", "Nama Petugas (Penghulu)"])

    if mode_rekap == "Kategori (Kantor/Luar/Isbat)":
        if 'Nikah Di' in df.columns:
            df['Nikah Di'] = df['Nikah Di'].str.strip().str.upper()
            list_pilihan = sorted(df['Nikah Di'].unique())
            pilihan = st.selectbox("Pilih Kategori:", list_pilihan)
            df_filtered = df[df['Nikah Di'] == pilihan].copy()
        else:
            st.error("Kolom 'Nikah Di' tidak ditemukan!")
            df_filtered = pd.DataFrame()
    else:
        if 'Nama Penghulu Hadir' in df.columns:
            list_pilihan = sorted(df['Nama Penghulu Hadir'].unique())
            pilihan = st.selectbox("Pilih Nama Petugas:", list_pilihan)
            df_filtered = df[df['Nama Penghulu Hadir'] == pilihan].copy()
        else:
            st.error("Kolom 'Nama Penghulu Hadir' tidak ditemukan!")
            df_filtered = pd.DataFrame()

    # --- PROSES TRANSFORMASI (KODE TETAP FOKUS PADA MAPPING LO) ---
    if not df_filtered.empty:
        df_filtered['No'] = range(1, len(df_filtered) + 1)
        # Logika Gabung Seri Huruf (P) + No Perforasi
        df_filtered['No_Perforasi_Fix'] = df_filtered['No Seri Huruf'] + ". " + df_filtered['No Perforasi']
        
        # Mapping Kolom Lengkap (Sesuai Pesanan: Pendaftaran, Wali, Nikah Di)
        mapping = {
            'No': 'No',
            'No Perforasi': 'No_Perforasi_Fix',
            'No Pemeriksaan': 'No Pemeriksaan',
            'No Aktanikah': 'No Aktanikah',
            'No Pendaftaran': 'No Pendaftaran',
            'Nama Suami': 'Nama Suami',
            'Nama Istri': 'Nama Istri',
            'Tanggal': 'Tanggal Akad',
            'Jam': 'Jam Akad',
            'Kelurahan': 'Nama Kelurahan',
            'Penghulu': 'Nama Penghulu Hadir',
            'Status Wali': 'Status Wali',
            'Nikah Di': 'Nikah Di'
        }
        
        final_df = df_filtered[[v for v in mapping.values()]].copy()
        final_df.columns = list(mapping.keys())

        st.info(f"📊 Ditemukan {len(final_df)} data untuk **{pilihan}**")

        if st.button(f"🚀 Download Laporan {pilihan}"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                workbook = writer.book
                worksheet = writer.sheets['Laporan']

                # Styling (Tetap Seperti Kode Awal)
                fmt_judul = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
                fmt_data = workbook.add_format({'border': 1, 'valign': 'vcenter', 'align': 'left'})

                # Judul Atas
                last_col = chr(64 + len(final_df.columns))
                worksheet.merge_range(f'A1:{last_col}1', f'LAPORAN REKAP {pilihan}', fmt_judul)
                worksheet.merge_range(f'A2:{last_col}2', f'BULAN {nama_bulan} TAHUN {year_val}', fmt_judul)

                # Auto Width & Write String (Anti Error NaN)
                for col_num, col_name in enumerate(final_df.columns):
                    if col_name == 'No':
                        width = 5
                    else:
                        max_len = final_df[col_name].astype(str).map(len).max()
                        width = max(max_len, len(col_name)) + 3
                    worksheet.set_column(col_num, col_num, width)
                    worksheet.write(4, col_num, col_name, fmt_header)
                    for row_num in range(len(final_df)):
                        val = str(final_df.iloc[row_num, col_num])
                        worksheet.write_string(row_num + 5, col_num, val if val != 'nan' else '', fmt_data)

            st.download_button(label=f"📥 Klik Sini Download {pilihan}", data=output.getvalue(), file_name=f"REKAP_{pilihan.replace(' ', '_')}.xlsx")