import streamlit as st
import pandas as pd
from io import BytesIO

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
    st.subheader("📁 Laporan Kategori Utama (K/LK/ISBAT)")
    
    # --- SMART COLUMN DETECTOR ---
    def get_col(keywords):
        for key in keywords:
            for c in df.columns:
                if key.upper() in str(c).upper(): return c
        return None

    col_lokasi = get_col(["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])
    
    if not col_lokasi:
        st.error("❌ Kolom Lokasi tidak ditemukan!")
        return

    # --- PILIHAN KATEGORI MANUAL (AGAR LENGKAP) ---
    st.write("🔍 **Filter Peristiwa:**")
    pilihan = st.radio(
        "Pilih Kategori:",
        ["SEMUA PERISTIWA", "KUA / KANTOR", "LUAR KUA / BEDOL", "ISBAT"],
        horizontal=True
    )
    
    # --- LOGIKA FILTER EKSLUSIF ---
    if pilihan == "SEMUA PERISTIWA":
        df_f = df.copy()
    elif pilihan == "KUA / KANTOR":
        # Logika: Ada 'KANTOR' tapi TIDAK ADA 'LUAR'
        mask = (df[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df[col_lokasi].str.contains("LUAR", na=False))
        df_f = df[mask].copy()
    elif pilihan == "LUAR KUA / BEDOL":
        df_f = df[df[col_lokasi].str.contains("LUAR", na=False)].copy()
    else: # ISBAT
        df_f = df[df[col_lokasi].str.contains("ISBAT", na=False)].copy()

    if not df_f.empty:
        df_f['No_Fix'] = range(1, len(df_f) + 1)
        
        # Gabung Seri & Perforasi
        c_seri = get_col(["NO SERI HURUF", "SERI HURUF"])
        c_perfo = get_col(["NO PERFORASI", "PERFORASI"])
        if c_seri and c_perfo:
            df_f['P_Perforasi'] = df_f[c_seri].astype(str) + ". " + df_f[c_perfo].astype(str)
        else:
            df_f['P_Perforasi'] = df_f[c_perfo] if c_perfo else ""

        mapping = {
            'No': 'No_Fix',
            'No Perforasi': 'P_Perforasi',
            'No Pemeriksaan': get_col(["NO PEMERIKSAAN", "PEMERIKSAAN"]),
            'No Aktanikah': get_col(["AKTANIKAH", "AKTA NIKAH"]),
            'Nama Suami': get_col(["NAMA SUAMI"]),
            'Nama Istri': get_col(["NAMA ISTRI"]),
            'Tanggal': get_col(["TANGGAL AKAD", "TGL AKAD"]),
            'Kelurahan': get_col(["NAMA KELURAHAN", "KELURAHAN", "DESA"]),
            'Penghulu': get_col(["NAMA PENGHULU HADIR", "NAMA PENGHULU"]), # Pasti NAMA, bukan NIP
            'Nikah Di': col_lokasi
        }
        
        available_source_cols = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[available_source_cols].copy()
        rename_map = {v: k for k, v in mapping.items() if v in available_source_cols}
        final_df.rename(columns=rename_map, inplace=True)

        st.info(f"📊 Menampilkan **{len(final_df)}** data untuk kategori: **{pilihan}**")
        st.dataframe(final_df, use_container_width=True)

        if st.button(f"🚀 Cetak Excel {pilihan}"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                format_excel(writer, final_df, pilihan, bulan, tahun)
            st.download_button(label="📥 Download Hasil", data=output.getvalue(), file_name=f"LAPORAN_{pilihan}.xlsx")
    else:
        st.warning(f"⚠️ Tidak ada data ditemukan untuk kategori {pilihan}.")