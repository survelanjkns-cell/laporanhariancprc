import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import pytz
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os
import re
import matplotlib.pyplot as plt

# --- KONSTAN & URL ---
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
# GID untuk tab "GRAF TREND KES MINGGUAN"
GID_GRAF = "1360010996" 
URL_GRAF = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_GRAF}"

TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING',
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# --- FUNGSI PENJANA GRAF (RAJAH 1) ---
def generate_trend_chart(url):
    try:
        # Membaca data dengan timeout untuk mengelakkan ralat 400 yang lama
        df_raw = pd.read_csv(url, header=None)
        
        # Ekstrak Data mengikut Range yang ditetapkan
        # x-axis: Minggu (B2:BB2)
        x_axis = df_raw.iloc[1, 1:54].values   
        # Data 2025: (B6:BB6)
        data_2025 = pd.to_numeric(df_raw.iloc[5, 1:54], errors='coerce') 
        # Data 2026: (B7:BB7)
        data_2026 = pd.to_numeric(df_raw.iloc[6, 1:54], errors='coerce') 
        # Moving Median: (B8:BB8)
        data_median = pd.to_numeric(df_raw.iloc[7, 1:54], errors='coerce')

        plt.figure(figsize=(12, 6))
        
        # Lukis garisan mengikut kod warna imej
        plt.plot(x_axis, data_2025, color='#4285F4', label='2025', linewidth=2.5) # Biru
        plt.plot(x_axis, data_2026, color='#EA4335', label='2026', linewidth=2.5) # Merah
        plt.plot(x_axis, data_median, color='#FBBC05', label='Moving median 4 tahun (2022,2023,2024,2025)', linewidth=2.5) # Kuning

        # Kemasan Estetik
        plt.xticks(x_axis, rotation=90, fontsize=8)
        plt.yticks(range(0, 1501, 250))
        plt.ylim(0, 1250)
        plt.xlim(1, 53)
        plt.grid(axis='y', linestyle='-', alpha=0.3)
        
        # Menghilangkan border atas dan kanan
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        plt.gca().spines['left'].set_visible(False)
        
        # Legend di bawah tengah
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3, frameon=False, fontsize=9)
        
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight', dpi=300)
        plt.close()
        img_stream.seek(0)
        return img_stream
    except Exception as e:
        st.error(f"Gagal memproses data Google Sheet: {e}")
        return None

# --- FUNGSI GENERATE DOCX ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    
    # ... [Bahagian kod header laporan anda kekalkan di sini] ...
    
    # --- MASUKKAN GRAF SELEPAS JADUAL 3 ---
    doc.add_paragraph() 
    chart_img = generate_trend_chart(URL_GRAF)
    
    if chart_img:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_cap = p_cap.add_run("Rajah 1 : Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor")
        run_cap.font.name = 'Arial'
        run_cap.font.size = Pt(10)
        run_cap.bold = True
        
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_img.add_run()
        run_img.add_picture(chart_img, width=Inches(6.2))
    
    # ... [Bahagian kod footer/BKK anda kekalkan di sini] ...

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # Simulasi pemprosesan data (Gantikan dengan logik sebenar anda)
            # doc_out = generate_docx(...)
            
            # Jika ralat 400 masih berlaku, Streamlit akan menangkapnya di sini
            st.success("✅ Laporan berjaya dijana!")
            # st.download_button(...)
            
        except Exception as e:
            st.error(f"Ralat: {e}")
