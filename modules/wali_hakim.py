import streamlit as st
import pandas as pd
from io import BytesIO
from modules.petugas import format_excel 

def render(df, bulan, tahun):
    st.subheader("⚖️ Ruang Laporan Wali Hakim")
    
    # --- SMART COLUMN DETECTOR ---
    def get_col(keywords):
        for key in keywords:
            # 1. Cari yang persis dulu
            for c in df.columns:
                if key.upper() == str(c).upper().strip():
                    return c
            # 2. Cari yang mengandung kata kunci
            for c in df.columns:
                if key.upper() in str(c).upper():
                    return c
        return None

    # Kita ambil kolom Status Wali sebagai dasar filter
    col_status_wali = get_col(["STATUS WALI", "STATUS_WALI"])
    # Kolom pendukung lainnya
    col_nama_wali = get_col(["NAMA LENGKAP WALI", "WALI HAKIM", "NAMA WALI"])
    
    if not col_status_wali:
        st.error("❌ Kolom 'Status Wali' tidak ditemukan!")
        return

    # --- LOGIKA FILTER WALI HAKIM ---
    # Sekarang kita cari kata 'HAKIM' di kolom Status Wali
    df_f = df[df[col_status_wali].str.contains("HAKIM", na=False, case=False)].copy()
    
    if not df_f.empty:
        df_f['No_Fix'] = range(1, len(df_f) + 1)
        
        # Mapping kolom untuk tampilan tabel
        mapping = {
            'No': 'No_Fix',
            'Nama Suami': get_col(["NAMA SUAMI"]),
            'Nama Istri': get_col(["NAMA ISTRI"]),
            'Status Wali': col_status_wali,
            'Nama Wali': col_nama_wali,
            'Sebab Wali': get_col(["SEBAB MENJADI WALI", "SEBAB WALI"]),
            'Tanggal': get_col(["TANGGAL AKAD", "TGL AKAD"]),
            'Penghulu': get_col(["NAMA PENGHULU HADIR", "NAMA PENGHULU", "PENGHULU HADIR"]),
            'Nikah Di': get_col(["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])
        }

        # Ambil kolom yang tersedia
        cols_tersedia = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[cols_tersedia].copy()
        
        # Rename ke judul cantik
        current_rename_map = {v: k for k, v in mapping.items() if v in cols_tersedia}
        final_df.rename(columns=current_rename_map, inplace=True)

        st.info(f"✅ Ditemukan **{len(final_df)}** data Wali Hakim (Filter: {col_status_wali})")
        st.dataframe(final_df, use_container_width=True, hide_index=True)

        # --- TOMBOL CETAK ---
        if st.button("🚀 Cetak Wali Hakim"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                format_excel(writer, final_df, "KHUSUS WALI HAKIM", bulan, tahun)
            
            st.download_button(
                label="📥 Download Excel Wali Hakim",
                data=output.getvalue(),
                file_name=f"Laporan_Wali_Hakim_{bulan}_{tahun}.xlsx",
                mime="application/vnd.ms-excel"
            )
    else:
        st.warning(f"🌙 Tidak ada data dengan status 'HAKIM' di kolom '{col_status_wali}'.")
