import streamlit as st
import pandas as pd
from io import BytesIO
from modules.petugas import format_excel 

def render(df, bulan, tahun):
    st.subheader("💰 Laporan PNBP NR")
    
    def get_col(keywords):
        for key in keywords:
            for c in df.columns:
                if key.upper() in str(c).upper().strip():
                    return c
        return None

    # Filter: Ambil yang Luar KUA
    col_lokasi = get_col(["NIKAH DI", "LOKASI"])
    if col_lokasi:
        df_f = df[df[col_lokasi].str.contains("LUAR", na=False)].copy()
    else:
        df_f = df.copy()

    if not df_f.empty:
        df_f['No_Fix'] = range(1, len(df_f) + 1)
        
        # --- MAPPING DENGAN DOUBLE CHECK (BIAR GAK KOSONG) ---
        # Kita cari kolom aslinya dulu di Excel lo
        c_pendaftaran = get_col(["NO PENDAFTARAN", "NOMOR PENDAFTARAN"])
        c_tgl_daftar = get_col(["TANGGAL DAFTAR", "TGL DAFTAR"])
        c_ntpn = get_col(["NO NTPN", "NTPN", "NOMOR NTPN"])
        c_tgl_bayar = get_col(["TANGGAL BAYAR", "TGL BAYAR", "TGL SETOR"])
        # Ini biang keroknya: Gue cari TARIF atau BIAYA kalau "Jumlah" gak ketemu
        c_jumlah = get_col(["JUMLAH YANG DI SETOR", "JUMLAH SETOR", "TARIF", "BIAYA"])
        c_penghulu = get_col(["NAMA PENGHULU HADIR", "NAMA PENGHULU"])

        # Kita susun tabelnya
        final_df = pd.DataFrame()
        final_df['No'] = df_f['No_Fix']
        final_df['No Pendaftaran'] = df_f[c_pendaftaran] if c_pendaftaran else ""
        final_df['Tanggal Daftar'] = df_f[c_tgl_daftar] if c_tgl_daftar else ""
        final_df['Nama Suami'] = df_f[get_col(["NAMA SUAMI"])]
        final_df['Nama Istri'] = df_f[get_col(["NAMA ISTRI"])]
        final_df['Tanggal Akad'] = df_f[get_col(["TANGGAL AKAD"])]
        final_df['Jam'] = df_f[get_col(["JAM AKAD", "JAM"])]
        final_df['No NTPN'] = df_f[c_ntpn] if c_ntpn else "-"
        final_df['Tanggal Bayar'] = df_f[c_tgl_bayar] if c_tgl_bayar else "-"
        
        # ISI KOLOM JUMLAH (Jika di Excel kosong, kita paksa isi 600.000)
        if c_jumlah:
            final_df['Jumlah yang disetor'] = df_f[c_jumlah]
        else:
            final_df['Jumlah yang disetor'] = "Rp. 600.000"
            
        final_df['Nama Penghulu Hadir'] = df_f[c_penghulu] if c_penghulu else "-"

        st.info(f"📊 Menampilkan **{len(final_df)}** data PNBP. (Kolom Jumlah diset otomatis jika data kosong)")
        st.dataframe(final_df, use_container_width=True, hide_index=True)

        if st.button("🚀 Cetak Excel PNBP"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                format_excel(writer, final_df, "PNBP NIKAH DI LUAR KANTOR", bulan, tahun)
            st.download_button(label="📥 Ambil Excel PNBP", data=output.getvalue(), file_name=f"PNBP_{bulan}.xlsx")
    else:
        st.warning("Data Luar Kantor tidak ditemukan.")
