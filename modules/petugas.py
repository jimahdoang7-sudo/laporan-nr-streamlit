import streamlit as st
import pandas as pd
from io import BytesIO

# --- FUNGSI FORMAT EXCEL (TETAP - JANGAN DIUBAH) ---
def format_excel(writer, final_df, title, month, year):
    workbook = writer.book
    worksheet = writer.sheets['Laporan']
    fmt_judul = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    fmt_data = workbook.add_format({'border': 1, 'valign': 'vcenter', 'align': 'left'})
    
    last_col = chr(64 + len(final_df.columns))
    worksheet.merge_range(f'A1:{last_col}1', f'LAPORAN {title}', fmt_judul)
    worksheet.merge_range(f'A2:{last_col}2', f'BULAN {month} TAHUN {year}', fmt_judul)

    for col_num, col_name in enumerate(final_df.columns):
        width = 5 if col_name == 'No' else max(final_df[col_name].astype(str).map(len).max(), len(col_name)) + 3
        worksheet.set_column(col_num, col_num, width)
        worksheet.write(4, col_num, col_name, fmt_header)
        for row_num in range(len(final_df)):
            val = str(final_df.iloc[row_num, col_num])
            worksheet.write_string(row_num + 5, col_num, val if val != 'NAN' else '', fmt_data)

def render(df, bulan, tahun):
    st.subheader("👨‍💼 Laporan Per Petugas")
    st.markdown("<br>", unsafe_allow_html=True) 

    # --- SMART COLUMN DETECTOR ---
    def get_col(keywords):
        for key in keywords:
            # Cari yang mengandung kata kunci (Case Insensitive)
            for c in df.columns:
                if key.upper() in str(c).upper():
                    return c
        return None

    # Deteksi Kolom Penting (Prioritas Nama agar tidak narik NIP)
    col_petugas = get_col(["NAMA PENGHULU HADIR", "NAMA PENGHULU", "PENGHULU HADIR"])
    col_lokasi = get_col(["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])

    if not col_petugas:
        st.error("❌ Kolom Nama Penghulu tidak ditemukan di file ini!")
        return

    # --- TOMBOL 1: PILIH NAMA PETUGAS ---
    list_petugas = sorted(df[col_petugas].unique())
    pilihan_pet = st.selectbox("📌 1. Pilih Nama Petugas (Penghulu):", list_petugas, key="select_petugas_utama")

    # Filter data milik petugas tersebut
    df_petugas = df[df[col_petugas] == pilihan_pet]

    # --- TOMBOL 2: FILTER KATEGORI ---
    st.write("📂 **2. Pilih Kategori Peristiwa:**")
    pilihan_kat = st.radio(
        label="Pilih Jenis:",
        options=["SEMUA DATA", "KUA / KANTOR", "LUAR KUA / BEDOL", "ISBAT"],
        horizontal=True,
        key="radio_tombol_ke_2"
    )

    # --- LOGIKA FILTER BERLAPIS (FIX DATA NYASAR) ---
    if pilihan_kat == "SEMUA DATA":
        df_f = df_petugas.copy()
    elif pilihan_kat == "KUA / KANTOR":
        # Logika: Ada 'KANTOR' tapi tak ada 'LUAR'
        mask = (df_petugas[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df_petugas[col_lokasi].str.contains("LUAR", na=False))
        df_f = df_petugas[mask].copy()
    elif pilihan_kat == "LUAR KUA / BEDOL":
        df_f = df_petugas[df_petugas[col_lokasi].str.contains("LUAR", na=False)].copy()
    elif pilihan_kat == "ISBAT":
        df_f = df_petugas[df_petugas[col_lokasi].str.contains("ISBAT", na=False)].copy()

    # --- TAMPILKAN TABEL ---
    if not df_f.empty:
        df_f['No_Urut'] = range(1, len(df_f) + 1)
        
        # Gabung Seri & Perforasi dengan pengaman
        c_seri = get_col(["NO SERI HURUF", "SERI HURUF"])
        c_perfo = get_col(["NO PERFORASI", "PERFORASI"])
        
        if c_seri and c_perfo:
            df_f['P_Perfo'] = df_f[c_seri].astype(str) + ". " + df_f[c_perfo].astype(str)
        elif c_perfo:
            df_f['P_Perfo'] = df_f[c_perfo]
        else:
            df_f['P_Perfo'] = ""
        
        # Mapping Kolom sesuai desain lo
        mapping = {
            'No': 'No_Urut', 
            'No Perforasi': 'P_Perfo', 
            'No Pemeriksaan': get_col(["PEMERIKSAAN"]), 
            'No Aktanikah': get_col(["AKTANIKAH", "AKTA NIKAH"]), 
            'No Pendaftaran': get_col(["PENDAFTARAN"]), 
            'Nama Suami': get_col(["NAMA SUAMI"]), 
            'Nama Istri': get_col(["NAMA ISTRI"]), 
            'Tanggal': get_col(["TANGGAL AKAD", "TGL AKAD"]), 
            'Jam': get_col(["JAM AKAD", "WAKTU AKAD"]), 
            'Kelurahan': get_col(["KELURAHAN", "DESA"]), 
            'Penghulu': col_petugas, 
            'Status Wali': get_col(["STATUS WALI"]), 
            'Nikah Di': col_lokasi
        }
        
        # Filter hanya kolom yang ketemu kuncinya
        available_cols = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[available_cols].copy()
        
        # Rename ke judul cantik
        rename_map = {v: k for k, v in mapping.items() if v in available_cols}
        final_df.rename(columns=rename_map, inplace=True)

        st.info(f"📊 Menampilkan **{len(final_df)}** data untuk **{pilihan_pet}** ({pilihan_kat})")
        st.dataframe(final_df, use_container_width=True)

        if st.button(f"🚀 Cetak Excel {pilihan_pet}"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                format_excel(writer, final_df, f"PETUGAS {pilihan_pet} ({pilihan_kat})", bulan, tahun)
            
            st.download_button(
                label=f"📥 Ambil File Excel", 
                data=output.getvalue(), 
                file_name=f"LAPORAN_{pilihan_pet}_{pilihan_kat}.xlsx"
            )
    else:
        st.warning(f"⚠️ Data tidak ditemukan untuk filter ini.")