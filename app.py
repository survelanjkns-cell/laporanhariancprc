import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# 1. Tetapan Tarikh
today_str = datetime.now().strftime('%d %B %Y')
yesterday = datetime.now() - timedelta(days=1)
yesterday_str = yesterday.strftime('%d April %Y') # Anda boleh tukar %B untuk nama bulan penuh

# 2. Fungsi untuk membaca data dari Google Sheets
def load_bkk_data():
    sheet_id = "1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c"
    
    # Baca data utama (Sheet 2026) untuk semakan insiden semalam
    url_main = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=2026"
    df_main = pd.read_csv(url_main)
    
    # Baca data table (Sheet table 2026) - Range AH2:AU
    # Nota: Kita baca keseluruhan dan slice mengikut keperluan column AH:AU
    url_table = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=table%202026"
    df_table_raw = pd.read_csv(url_table)
    
    # Extract columns AH:AU (Indeks column bermula dari 0, sila sesuaikan jika perlu)
    # Anda mungkin perlu menyesuaikan indeks column mengikut kedudukan AH hingga AU
    df_table = df_table_raw.iloc[:, 33:47] 
    
    return df_main, df_table

try:
    df_main, df_table = load_bkk_data()
    
    # Semakan insiden semalam pada column "TKH LAPOR"
    # Anda perlu pastikan format tarikh dalam Google Sheet sepadan
    insiden_semalam = df_main[df_main['TKH LAPOR'] == yesterday_str]
    total_insiden = len(insiden_semalam)

    # --- BAHAGIAN PAPARAN LAPORAN 4.0 ---
    st.markdown("### **4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)**")
    
    # Sub tajuk 4.1
    sub_title = "4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor."
    
    if total_insiden == 0:
        st.write(f"{sub_title} Tiada Insiden dilaporkan pada {yesterday_str}.")
    else:
        st.write(f"{sub_title} Sejumlah {total_insiden} input notifikasi Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK) telah diterima pada {yesterday_str} dengan pecahan mengikut penyakit seperti dalam jadual 4.")

    # Pembersihan Data Table (Remove bracket content)
    # Contoh: "1 (1)" menjadi "1"
    def clean_brackets(val):
        if isinstance(val, str):
            import re
            return re.sub(r'\s*\([^)]*\)', '', val)
        return val

    df_table_cleaned = df_table.applymap(clean_brackets)

    # Papar Table
    st.table(df_table_cleaned)
    
    # Tajuk Bawah Table
    st.markdown("<p style='text-align: center;'>Jadual 4 : Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)</p>", unsafe_url=True)

    # Sumber (Footer)
    st.write(f"*\*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({today_str} @ 10.00 am)*")

except Exception as e:
    st.error(f"Ralat memuatkan data: {e}")
