import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import pytz
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os
import re
import matplotlib.pyplot as plt

# --- KONSTAN & URL ---
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID_GRAF = "1360010996" 
URL_GRAF = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_GRAF}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"
GSHEET_ENOTIF_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"

TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING',
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

AVG_HARIAN_FIGURES = {
    "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "HIV/AIDS": 7, "Leptosopsirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

# --- HELPERS ---
def get_msia_time():
    msia_tz = pytz.timezone('Asia/Kuala_Lumpur')
    return datetime.now(msia_tz)

def format_penyakit_name(name):
    name_str = str(name).strip().upper()
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]): return name_str
    if "FOOD POISONING" in name_str: return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]: return "Denggi"
    return name_str.title()

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

# --- PENJANA GRAF (RAJAH 1) ---
def generate_trend_chart(url):
    try:
        df_raw = pd.read_csv(url, header=None)
        x_axis = df_raw.iloc[1, 1:54].values   
        data_2025 = pd.to_numeric(df_raw.iloc[5, 1:54], errors='coerce') 
        data_2026 = pd.to_numeric(df_raw.iloc[6, 1:54], errors='coerce') 
        data_median = pd.to_numeric(df_raw.iloc[7, 1:54], errors='coerce')

        plt.figure(figsize=(12, 6))
        plt.plot(x_axis, data_2025, color='#4285F4', label='2025', linewidth=2.5)
        plt.plot(x_axis, data_2026, color='#EA4335', label='2026', linewidth=2.5)
        plt.plot(x_axis, data_median, color='#FBBC05', label='Moving median 4 tahun', linewidth=2.5)

        plt.xticks(x_axis, rotation=90, fontsize=8)
        plt.yticks(range(0, 1501, 250))
        plt.ylim(0, 1250)
        plt.xlim(1, 53)
        plt.grid(axis='y', linestyle='-', alpha=0.3)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False)
        
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight', dpi=300)
        plt.close()
        img_stream.seek(0)
        return img_stream
    except: return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_final, b_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    # Logo & Tajuk (Diringkaskan untuk ruang)
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        doc.add_paragraph().add_run().add_picture(logo_path, width=Inches(1.5))
    
    doc.add_paragraph("LAPORAN HARIAN BWKK CPRC SELANGOR").alignment = 1

    # --- JADUAL 1 & 2 (Kod anda sedia ada masuk di sini) ---
    doc.add_paragraph(f"Tarikh: {today} | Minggu Epi: {today.isocalendar()[1]}")

    # --- RAJAH 1 ---
    doc.add_paragraph("\n3.0 Ringkasan Laporan Wabak Vektor")
    # ... Jadual 3 ...
    
    chart_img = generate_trend_chart(URL_GRAF)
    if chart_img:
        p = doc.add_paragraph()
        p.alignment = 1
        p.add_run("Rajah 1 : Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026").bold = True
        doc.add_paragraph().add_run().add_picture(chart_img, width=Inches(6.0))

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.title("📊 BWKK Report Generator")

# Guna Session State supaya butang muat turun tidak hilang
if 'doc_ready' not in st.session_state:
    st.session_state.doc_ready = False
    st.session_state.doc_data = None

f1 = st.file_uploader("Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            with st.spinner('Menjana laporan dan graf...'):
                # --- LOGIK PROSES DATA ---
                df1 = pd.read_excel(f1)
                # (Tambahkan semua logik pembersihan data anda di sini)
                
                # Mock data untuk demonstrasi (Ganti dengan variabel sebenar anda)
                doc_out = generate_docx(None, None, None, None, None, [], []) 
                
                st.session_state.doc_data = doc_out
                st.session_state.doc_ready = True
            st.success("Laporan berjaya dijana!")
        except Exception as e:
            st.error(f"Ralat: {e}")

# Butang muat turun diletakkan di luar block button jana
if st.session_state.doc_ready:
    st.download_button(
        label="⬇️ Muat Turun Laporan (Word)",
        data=st.session_state.doc_data,
        file_name=f"Laporan_BWKK_{date.today()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
