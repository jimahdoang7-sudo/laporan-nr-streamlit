import streamlit as st
import pandas as pd
from io import BytesIO
from modules.petugas import format_excel # Kita pakai fungsi format yang sudah ada

def render(df, bulan, tahun):
    st.subheader("🌍 Laporan Khusus Peristiwa WNA")
    
    # --- SMART COLUMN DETECTOR ---
    def get_col(keywords):
        for key in keywords:
            # 1. Cari yang persis
            for c in df.columns:
                if key.upper() == str(c).upper().strip():
                    return c
            # 2. Cari yang mengandung kata kunci
            for c in df.columns:
                if key.upper() in str(c).upper():
                    return c
        return None

    # Cari kolom Kewarganegaraan
    col_ws = get_col(["WARGANEGARA SUAMI", "WN SUAMI", "WNA S"])
    col_wi = get_col(["WARGANEGARA ISTRI", "WN ISTRI", "WNA I"])
    
    if not col_ws or not col_wi:
        st.error("❌ Kolom Kewarganegaraan (Suami/Istri) tidak ditemukan!")
        return

    # --- LOGIKA FILTER WNA ---
    # Cari yang salah satu atau keduanya BUKAN WNI
    df_f = df[(df[col_ws] != 'WNI') | (df[col_wi] != 'WNI')].copy()
    
    if not df_f.empty:
        df_f['No_Fix'] = range(1, len(df_f) + 1)
        
        # Mapping Kolom Cantik
        mapping = {
            'No': 'No_Fix',
            'Nama Suami': get_col(["NAMA SUAMI"]),
            'Asal WNA Suami': col_ws,
            'Nama Istri': get_col(["NAMA ISTRI"]),
            'Asal WNA Istri': col_wi,
            'Tanggal': get_col(["TANGGAL AKAD", "TGL AKAD"]),
            'Penghulu': get_col(["NAMA PENGHULU HADIR", "NAMA PENGHULU", "PENGHULU HADIR"]),
            'Nikah Di': get_col(["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])
        }
        
        # Ambil kolom yang tersedia
        available_cols = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[available_cols].copy()
        
        # Rename ke judul cantik
        rename_map = {v: k for k, v in mapping.items() if v in available_cols}
        final_df.rename(columns=rename_map, inplace=True)
        
        st.info(f"📊 Ditemukan **{len(final_df)}** data pengantin WNA.")
        st.dataframe(final_df, use_container_width=True)
        
        # --- FITUR CETAK EXCEL ---
        if st.button("🚀 Cetak Laporan WNA"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Tulis data mulai baris ke-5
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                
                # Gunakan format standar kita
                format_excel(writer, final_df, "DATA PERISTIWA WNA", bulan, tahun)
            
            st.download_button(
                label="📥 Download Excel WNA",
                data=output.getvalue(),
                file_name=f"Laporan_WNA_{bulan}_{tahun}.xlsx",
                mime="application/vnd.ms-excel"
            )
    else:
        st.warning("🌙 Tidak ada data WNA yang ditemukan untuk bulan ini.")