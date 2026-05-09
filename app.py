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
def set_repeat_table_header(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w')))
    trPr.append(tblHeader)

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

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    return f"{target_date.day:02d} {months_ms[target_date.month]} {target_date.year} ({days_ms[target_date.strftime('%A')]})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

# --- PENJANA GRAF (RAJAH 1) ---
def generate_trend_chart(url):
    try:
        df_raw = pd.read_csv(url, header=None)
        x_axis = df_raw.iloc[1, 1:54].values
        d2025 = pd.to_numeric(df_raw.iloc[5, 1:54], errors='coerce')
        d2026 = pd.to_numeric(df_raw.iloc[6, 1:54], errors='coerce')
        dmed = pd.to_numeric(df_raw.iloc[7, 1:54], errors='coerce')

        plt.figure(figsize=(11, 5))
        plt.plot(x_axis, d2025, color='#4285F4', label='2025', linewidth=2)
        plt.plot(x_axis, d2026, color='#EA4335', label='2026', linewidth=2)
        plt.plot(x_axis, dmed, color='#FBBC05', label='Moving Median', linewidth=2)
        plt.xticks(x_axis, rotation=90, fontsize=8)
        plt.grid(axis='y', alpha=0.3)
        plt.legend(loc='lower center', bbox_to_anchor=(0.5, -0.3), ncol=3)
        
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', dpi=200)
        plt.close()
        buf.seek(0)
        return buf
    except: return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table, bkk_details, df_yesterday_list):
    doc = Document()
    now = get_msia_time()
    today = now.date()
    yesterday = today - timedelta(days=1)
    
    # Page Setup
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2)
    
    # 1. Logo
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        p = doc.add_paragraph()
        p.alignment = 1
        p.add_run().add_picture(logo_path, width=Inches(1.6))
        
    # 2. Tajuk
    for txt in ["LAPORAN HARIAN BWKK CPRC SELANGOR", "JABATAN KESIHATAN NEGERI SELANGOR"]:
        p = doc.add_paragraph()
        p.alignment = 1
        run = p.add_run(txt)
        apply_font(run, 11, True)

    doc.add_paragraph(f"Tarikh: {get_malay_date(today)} | Minggu Epi: {get_epi_week(today)}").alignment = 1

    # --- JADUAL 1 (eNotifikasi) ---
    doc.add_paragraph("\n1.0 Ringkasan Laporan Input Enotifikasi").bold = True
    # (Di sini anda perlu masukkan kod pembinaan Jadual 1 anda yang asal)
    # ... Kod jadual 1 ...

    # --- SECTION 3.0 & GRAF ---
    doc.add_paragraph("\n3.0 Ringkasan Laporan Wabak Vektor").bold = True
    # (Di sini anda perlu masukkan kod pembinaan Jadual 3 anda yang asal)
    # ... Kod jadual 3 ...

    # MASUKKAN RAJAH 1
    chart = generate_trend_chart(URL_GRAF)
    if chart:
        doc.add_paragraph().alignment = 1
        p = doc.add_paragraph()
        p.alignment = 1
        run = p.add_run("Rajah 1 : Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026")
        apply_font(run, 10, True)
        doc.add_picture(chart, width=Inches(6))

    # --- SECTION 4.0 (BKK) ---
    doc.add_paragraph("\n4.0 Ringkasan Laporan BKK").bold = True
    # ... Kod jadual 4 ...

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# --- STREAMLIT UI ---
if 'doc_file' not in st.session_state:
    st.session_state.doc_file = None

st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("Upload Excel 1", type=["xlsx", "xls"])
f2 = st.file_uploader("Upload Excel 2", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        # Masukkan logik bacaan excel anda (pd.read_excel)
        # Guna fungsi generate_docx(...)
        st.session_state.doc_file = generate_docx(None, None, None, None, None, [], []) 
        st.success("Laporan Berjaya Dijana!")

if st.session_state.doc_file:
    st.download_button("⬇️ Muat Turun File", st.session_state.doc_file, "Laporan.docx")
