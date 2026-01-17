import pandas as pd
import os
import glob

ARCHIVE_DIR = "data/archive"

def get_all_data():
    if not os.path.exists(ARCHIVE_DIR):
        os.makedirs(ARCHIVE_DIR)
        return None
        
    all_files = glob.glob(os.path.join(ARCHIVE_DIR, "*.csv")) + \
                glob.glob(os.path.join(ARCHIVE_DIR, "*.xlsx"))
    
    if not all_files:
        return None

    all_data = []
    for filename in all_files:
        # PENGAMAN 1: Abaikan file temporary Excel (~$namafile.xlsx)
        if os.path.basename(filename).startswith("~$"):
            continue
            
        try:
            # PENGAMAN 2: Gunakan engine openpyxl untuk Excel
            if filename.endswith('.csv'):
                df = pd.read_csv(filename, dtype=str)
            else:
                df = pd.read_excel(filename, dtype=str)
            
            if not df.empty:
                df['Nama File'] = os.path.basename(filename)
                all_data.append(df)
        except Exception as e:
            # Jangan biarkan 1 file rusak menghentikan seluruh aplikasi
            st.error(f"Gagal membaca file {os.path.basename(filename)}: {e}")
            continue

    if all_data:
        # PENGAMAN 3: Gabungkan dengan penanganan kolom yang berbeda
        return pd.concat(all_data, axis=0, ignore_index=True, sort=False)
    return None