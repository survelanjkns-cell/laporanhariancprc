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

# --- KONSTAN & MAPPING DATA ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING',
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

AVG_HARIAN_FIGURES = {
    "Denggi": 427, "Covid-19": 54, "Hfmd": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "Hiv/Aids": 7, "Leptosopsirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID_GRAF = "757820121"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

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
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]):
        return name_str
    if "FOOD POISONING" in name_str: return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]: return "Denggi"
    return name_str.title()

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def clean_val(val):
    if pd.isna(val) or str(val).strip() in ["", "-", "nan"]: return "-"
    cleaned = re.sub(r'\s*\(.*?\)', '', str(val)).strip()
    return cleaned if cleaned != "" else "-"

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    return f"{target_date.day:02d} {months_ms.get(target_date.month)} {target_date.year} ({days_ms.get(target_date.strftime('%A'))})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def add_table_title(doc, label, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_label = p.add_run(f"{label} : ")
    apply_font(run_label, 11, bold=True)
    run_title = p.add_run(title)
    apply_font(run_title, 11, bold=False)

# --- CHART GENERATOR (FIXED FLOAT/NAN ERROR) ---
def fetch_and_generate_chart():
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&gid={GID_GRAF}"
        df_graf = pd.read_csv(url, header=None)
        
        # Paksa x-axis jadi string dan bersihkan NaN supaya tidak ralat float
        raw_x = df_graf.iloc[1, 1:53].fillna('').astype(str).values
        x_axis = [str(x) for x in raw_x]
        
        # Ambil data Y
        y_2025 = pd.to_numeric(df_graf.iloc[5, 1:53], errors='coerce').fillna(0).values
        y_2026 = pd.to_numeric(df_graf.iloc[6, 1:53], errors='coerce').fillna(0).values
        y_median = pd.to_numeric(df_graf.iloc[7, 1:53], errors='coerce').fillna(0).values

        plt.figure(figsize=(10, 4))
        plt.plot(x_axis, y_2025, label='2025', color='#4285F4', linewidth=2)
        plt.plot(x_axis, y_2026, label='2026', color='#EA4335', linewidth=2)
        plt.plot(x_axis, y_median, label='Moving median 4 tahun (2022,2023,2024,2025)', color='#FBBC05', linewidth=2)
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=3, frameon=False, fontsize=8)
        plt.xticks(rotation=90, fontsize=7)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=300)
        plt.close()
        buf.seek(0)
        return buf
    except Exception as e:
        st.warning(f"Gagal menjana graf: {e}")
        return None

# --- DOCX GENERATOR ---
def generate_docx(matrix, col_totals, wabak_df, vector_df, bkk_table_final, is_bkk_empty, bkk_details, df_yesterday_list, chart_img):
    doc = Document()
    now_msia = get_msia_time()
    today, yesterday = now_msia.date(), now_msia.date() - timedelta(days=1)
    
    # ... (Struktur Header & Ringkasan eNotifikasi anda) ...
    # Bahagian ini dipendekkan untuk fokus pada integrasi graf
    
    # [Contoh kedudukan selepas Jadual 3 Vektor]
    if chart_img:
        doc.add_page_break()
        add_table_title(doc, "Rajah 1", "Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(chart_img, width=Inches(6))

    # Signature Table (Tanpa border)
    doc.add_paragraph()
    sig_table = doc.add_table(rows=8, cols=3)
    tbl = sig_table._tbl
    tblPr = tbl.tblPr if tbl.tblPr else tbl.get_or_add_tblPr()
    tblBorders = parse_xml(r'<w:tblBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>' % nsdecls('w'))
    tblPr.append(tblBorders)
    
    # Isi label tandatangan
    def fill_sig(r, l):
        apply_font(sig_table.rows[r].cells[0].paragraphs[0].add_run(l), 11, False)
        apply_font(sig_table.rows[r].cells[1].paragraphs[0].add_run(":"), 11, False)
    fill_sig(0, "Disediakan"); fill_sig(1, "Jawatan")
    fill_sig(3, "Disemak"); fill_sig(4, "Jawatan")
    fill_sig(6, "Disahkan"); fill_sig(7, "Jawatan")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator")
st.title("📄 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # 1. Process eNotifikasi
            df1 = pd.read_excel(f1)
            df1 = df1[(df1['Notifikasi Status'] != 'Abai Notifikasi') & (df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS))]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            col_totals = matrix.sum()

            # 2. Jana Graf
            chart = fetch_and_generate_chart()
            
            # 3. Jana Dokumen (Sila lengkapkan pemprosesan df2 anda di sini)
            # Placeholder untuk df lain supaya kod boleh run
            doc_out = generate_docx(matrix, col_totals, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), True, [], [], chart)
            
            st.success("✅ Laporan Berjaya!")
            st.download_button("⬇️ Muat Turun", doc_out, f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
