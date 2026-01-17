import streamlit as st
import pandas as pd
from io import BytesIO
import altair as alt

# --- FUNGSI FORMAT EXCEL (TETAP) ---
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

    def get_col(keywords):
        for key in keywords:
            for c in df.columns:
                if key.upper() in str(c).upper(): return c
        return None

    col_petugas = get_col(["NAMA PENGHULU HADIR", "NAMA PENGHULU", "PENGHULU HADIR"])
    col_lokasi = get_col(["NIKAH DI", "TEMPAT NIKAH", "LOKASI"])
    col_daftar = get_col(["NO PENDAFTARAN", "NOMOR PENDAFTARAN", "PENDAFTARAN"])

    if not col_petugas:
        st.error("❌ Kolom Nama Penghulu tidak ditemukan!")
        return

    # --- RINGKASAN & CHART ---
    all_p = sorted(df[col_petugas].unique())
    summary_data = []
    plot_data = []
    
    for p_name in all_p:
        df_p_all = df[df[col_petugas] == p_name]
        c_kan = len(df_p_all[(df_p_all[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df_p_all[col_lokasi].str.contains("LUAR", na=False))])
        c_luar = len(df_p_all[df_p_all[col_lokasi].str.contains("LUAR", na=False)])
        c_isbat = len(df_p_all[df_p_all[col_daftar].str.contains("^IB", na=False, case=False)]) if col_daftar else 0
        
        summary_data.append({"Petugas": p_name, "🏢 K": c_kan, "🚗 LK": c_luar, "⚖️ ISB": c_isbat, "Total": len(df_p_all)})
        plot_data.append({"Petugas": p_name, "Kategori": "Kantor", "Jumlah": c_kan})
        plot_data.append({"Petugas": p_name, "Kategori": "Luar", "Jumlah": c_luar})
        plot_data.append({"Petugas": p_name, "Kategori": "Isbat", "Jumlah": c_isbat})

    c1, c2 = st.columns([1.5, 1])
    with c1:
        df_plot = pd.DataFrame(plot_data)
        chart = alt.Chart(df_plot).mark_bar().encode(
            x=alt.X('Jumlah:Q'), y=alt.Y('Petugas:N', sort='-x'),
            color=alt.Color('Kategori:N', scale=alt.Scale(range=['#9c27b0', '#e91e63', '#ffa000'])),
            tooltip=['Petugas', 'Kategori', 'Jumlah']
        ).properties(height=150)
        st.altair_chart(chart, use_container_width=True)
    
    with c2:
        for row in summary_data:
            st.markdown(f"""
            <div style="border-left: 3px solid #9c27b0; padding-left: 10px; margin-bottom: 5px; background: rgba(255,255,255,0.4); border-radius: 0 10px 10px 0;">
                <span style="color: #7b1fa2; font-size: 0.8rem; font-weight: bold;">{row['Petugas']}</span><br>
                <span style="font-size: 0.75rem;">K: {row['🏢 K']} | LK: {row['🚗 LK']} | ISB: {row['⚖️ ISB']} | <b>Tot: {row['Total']}</b></span>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)

    pilihan_pet = st.selectbox("📌 1. Pilih Nama Petugas:", all_p, key="select_petugas_utama")
    df_petugas = df[df[col_petugas] == pilihan_pet]
    pilihan_kat = st.radio("📂 2. Pilih Kategori:", ["SEMUA DATA", "KUA / KANTOR", "LUAR KUA / BEDOL", "ISBAT"], horizontal=True)

    # --- LOGIKA FILTER ---
    if pilihan_kat == "SEMUA DATA":
        df_f = df_petugas.copy()
    elif pilihan_kat == "KUA / KANTOR":
        mask = (df_petugas[col_lokasi].str.contains("KANTOR|KUA", na=False)) & (~df_petugas[col_lokasi].str.contains("LUAR", na=False))
        df_f = df_petugas[mask].copy()
    elif pilihan_kat == "LUAR KUA / BEDOL":
        df_f = df_petugas[df_petugas[col_lokasi].str.contains("LUAR", na=False)].copy()
    elif pilihan_kat == "ISBAT":
        df_f = df_petugas[df_petugas[col_daftar].str.contains("^IB", na=False, case=False)].copy() if col_daftar else pd.DataFrame()

    # --- TAMPILKAN TABEL ---
    if not df_f.empty:
        df_f['No_Urut'] = range(1, len(df_f) + 1)
        c_seri = get_col(["NO SERI HURUF", "SERI HURUF"])
        c_perfo = get_col(["NO PERFORASI", "PERFORASI"])
        df_f['P_Perfo'] = (df_f[c_seri].astype(str) + ". " + df_f[c_perfo].astype(str)) if (c_seri and c_perfo) else (df_f[c_perfo] if c_perfo else "")

        # --- MAPPING TETAP UTUH ---
        mapping = {
            'No': 'No_Urut', 'No Perforasi': 'P_Perfo', 
            'No Pemeriksaan': get_col(["PEMERIKSAAN"]), 
            'No Aktanikah': get_col(["AKTANIKAH", "AKTA NIKAH"]), 
            'No Pendaftaran': col_daftar, 
            'Nama Suami': get_col(["NAMA SUAMI"]), 'Nama Istri': get_col(["NAMA ISTRI"]), 
            'Tanggal': get_col(["TANGGAL AKAD", "TGL AKAD"]), 
            'Jam': get_col(["JAM AKAD", "WAKTU AKAD"]), 
            'Kelurahan': get_col(["KELURAHAN", "DESA"]), 
            'Penghulu': col_petugas, 
            'Status Wali': get_col(["STATUS WALI"]), 'Nikah Di': col_lokasi
        }
        
        available_cols = [v for v in mapping.values() if v and v in df_f.columns]
        final_df = df_f[available_cols].copy()
        rename_map = {v: k for k, v in mapping.items() if v in available_cols}
        final_df.rename(columns=rename_map, inplace=True)

        st.info(f"📊 Menampilkan **{len(final_df)}** data untuk **{pilihan_pet}**")
        
        # FIX: Sembunyikan indeks 0, 1, 2...
        st.dataframe(final_df, use_container_width=True, hide_index=True)

        if st.button(f"🚀 Cetak Excel {pilihan_pet}"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Laporan', startrow=4)
                format_excel(writer, final_df, f"PETUGAS {pilihan_pet}", bulan, tahun)
            st.download_button(label="📥 Ambil File", data=output.getvalue(), file_name=f"LAPORAN_{pilihan_pet}.xlsx")
    else:
        st.warning("⚠️ Data tidak ditemukan.")
