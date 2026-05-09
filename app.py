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

# --- KONSTAN & MAPPING DATA ---
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

# --- CONFIG GOOGLE SHEETS ---
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID_VEKTOR = "0"
GID_GRAF = "68285521" 

URL_VEKTOR = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_VEKTOR}"
URL_GRAF = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_GRAF}"
URL_BKK = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

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
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    return f"{((target_date - start_date).days // 7) + 1}/{target_date.year}"

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
    run_label = p.add_run(f"{label} : ")
    apply_font(run_label, 11, bold=True)
    run_title = p.add_run(title)
    apply_font(run_title, 11, bold=False)

# --- FUNGSI GENERATE GRAF ---
def create_dengue_chart():
    try:
        df_chart = pd.read_csv(URL_GRAF, header=None)
        weeks = df_chart.iloc[1, 1:54].values
        data_2025 = pd.to_numeric(df_chart.iloc[5, 1:54], errors='coerce').fillna(0)
        data_2026 = pd.to_numeric(df_chart.iloc[6, 1:54], errors='coerce').fillna(0)
        data_median = pd.to_numeric(df_chart.iloc[7, 1:54], errors='coerce').fillna(0)

        plt.figure(figsize=(10, 5))
        plt.plot(weeks, data_2025, label='2025', color='#4285F4', linewidth=2)
        plt.plot(weeks, data_2026, label='2026', color='#EA4335', linewidth=2)
        plt.plot(weeks, data_median, label='Moving median 4 tahun', color='#FBBC05', linewidth=2)
        plt.xticks(ticks=range(len(weeks)), labels=weeks, fontsize=8)
        plt.gca().xaxis.set_major_locator(plt.MultipleLocator(2))
        plt.ylim(0, 1250)
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=300)
        img_buf.seek(0)
        plt.close()
        return img_buf
    except: return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today, yesterday = now_msia.date(), now_msia.date() - timedelta(days=1)
    
    # Setup Margins
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)

    # 1. Logo & Tajuk
    titles = ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]
    for t in titles:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(t), 10.5, bold=True)

    # 2. Jadual Info Tarikh (Hijau)
    it = doc.add_table(rows=1, cols=2)
    it.width = Cm(16)
    for i, txt in enumerate([f"Tarikh : {get_malay_date(today)}", f"Minggu Epidemiologi : {get_epi_week(today)}"]):
        cell = it.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 11, bold=True)

    # --- JADUAL 1 (E-Notifikasi) ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    
    # Headers J1
    pkd_short = ['GBK', 'HL', 'HS', 'KLG', 'KL', 'KS', 'PTG', 'SB', 'SPG']
    h_cells = t1.rows[0].cells
    apply_font(h_cells[0].add_run("PENYAKIT"), 8, bold=True)
    for i, pkd in enumerate(pkd_short): apply_font(h_cells[i+1].add_run(pkd), 8, bold=True)
    apply_font(h_cells[-2].add_run("Jumlah"), 8, bold=True)
    apply_font(h_cells[-1].add_run("Avg"), 8, bold=True)

    # Data J1 (NaN Safe)
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx+1].cells
        apply_font(row[0].add_run(format_penyakit_name(penyakit)), 8, bold=True)
        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            val = int(row_data.get(pkd, 0)) if pd.notna(row_data.get(pkd)) else 0
            apply_font(row[c_idx+1].paragraphs[0].add_run(str(val)), 8, bold=True)
        
        total_val = int(row_data.get('Grand Total', 0)) if pd.notna(row_data.get('Grand Total')) else 0
        apply_font(row[-2].paragraphs[0].add_run(str(total_val)), 8, bold=True)
        
        avg_val = int(row_data.get('Average Harian', 0)) if pd.notna(row_data.get('Average Harian')) else 0
        apply_font(row[-1].paragraphs[0].add_run(str(avg_val)), 8, bold=True)

    # --- JADUAL 3 (Vektor) & RAJAH 1 ---
    # (Penambahan Graf Rajah 1 selepas J3)
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    
    # ... (Proses J3 anda) ...
    
    # RAJAH 1
    doc.add_paragraph()
    add_table_title(doc, "Rajah 1", "Carta Kes Mingguan Denggi Didaftar")
    chart_buf = create_dengue_chart()
    if chart_buf:
        doc.add_picture(chart_buf, width=Inches(6))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # S1: Proses Data Notifikasi (NaN Safe)
            df1 = pd.read_excel(f1).fillna(0)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            col_sums = matrix.sum()

            # S2: Wabak (NaN Safe)
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2").fillna(0)
            yesterday = (get_msia_time() - timedelta(days=1)).date()
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            
            # S3: Vektor & BKK (NaN Safe)
            v_data = pd.read_csv(URL_VEKTOR, header=None).fillna(0)
            bkk_df = pd.read_csv(URL_BKK, header=None).fillna(0)

            doc_out = generate_docx(matrix, col_sums, None, v_data, bkk_df, True, [], [])
            st.success("✅ Laporan Berjaya!")
            st.download_button("⬇️ Download", doc_out, f"BWKK_{date.today()}.docx")
            
        except Exception as e:
            st.error(f"Ralat: {e}")
