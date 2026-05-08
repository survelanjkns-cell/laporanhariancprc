import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
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

# Sheet IDs
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID_VEKTOR = "0"
GID_GRAF = "1453268800"  # <--- SILA PASTIKAN GID INI BETUL UNTUK TAB GRAF
GSHEET_URL_VEKTOR = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_VEKTOR}"
GSHEET_URL_GRAF = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_GRAF}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
def disable_no_wrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    noWrap = parse_xml(r'<w:noWrap {} w:val="on"/>'.format(nsdecls('w')))
    tcPr.append(noWrap)

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

def set_cell_paddings(cell, top=None, start=None, bottom=None, end=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = parse_xml(r'<w:tcMar {}/>'.format(nsdecls('w')))
    for margin, value in [('top', top), ('left', start), ('bottom', bottom), ('right', end)]:
        if value is not None:
            node = parse_xml(r'<w:{} {} w:w="{}" w:type="dxa"/>'.format(margin, nsdecls('w'), value))
            tcMar.append(node)
    tcPr.append(tcMar)

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

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
    run_label = p.add_run(f"{label} : ")
    apply_font(run_label, 11, bold=True)
    run_title = p.add_run(title)
    apply_font(run_title, 11, bold=False)

def add_pkd_note(doc):
    p = doc.add_paragraph()
    run = p.add_run("*Nota : GBK, Gombak; HL, Hulu Langat; HS, Hulu Selangor; KLG, Klang; KL, Kuala Langat; KS, Kuala Selangor; PTG, Petaling; SB, Sabak Bernam; SPG, Sepang.")
    apply_font(run, 7, bold=False)

# --- CHART GENERATOR ---
def generate_trend_chart_img():
    try:
        # Load the CSV, specifically targeting the data area (A6:M8 based on screenshot)
        df_chart = pd.read_csv(GSHEET_URL_GRAF, skiprows=5)
        
        # Prepare Data
        weeks = [str(i) for i in range(1, 54)]
        y_2025 = pd.to_numeric(df_chart.iloc[0, 1:54], errors='coerce')
        y_2026 = pd.to_numeric(df_chart.iloc[1, 1:54], errors='coerce')
        y_median = pd.to_numeric(df_chart.iloc[2, 1:54], errors='coerce')

        plt.figure(figsize=(10, 5))
        plt.plot(weeks, y_2025, label='2025', color='#3b82f6', marker='.', markersize=4)
        plt.plot(weeks, y_2026, label='2026', color='#ef4444', linewidth=2)
        plt.plot(weeks, y_median, label='Moving median 4 tahun (2022-2025)', color='#f59e0b', linestyle='-')
        
        plt.title("CARTA KES MINGGUAN DENGGI JKN SELANGOR", fontsize=12, fontweight='bold', pad=15)
        plt.xlabel("Minggu Epidemiologi", fontsize=10)
        plt.ylabel("Jumlah Kes", fontsize=10)
        plt.xticks(weeks[::2], fontsize=8) # Show every 2nd label
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False)
        plt.grid(True, axis='y', linestyle='--', alpha=0.5)
        
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=150)
        img_buf.seek(0)
        plt.close()
        return img_buf
    except Exception as e:
        st.error(f"Gagal menjana graf: {e}")
        return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.0)
    content_width = section.page_width - section.left_margin - section.right_margin

    # Logo & Title
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(logo_path, width=Inches(1.5))

    titles = ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 
              "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 
              "JABATAN KESIHATAN NEGERI SELANGOR"]
    for t in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(t), 10.5, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # Header Green Info
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i==0 else f"Minggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- SECTION 1.0 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    p11 = doc.add_paragraph()
    apply_font(p11.add_run(f"1.1 Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {get_malay_date(yesterday)}."), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    # ... [Logic remains same as your original script for tables] ...
    # (Abbreviated here for brevity, keep your original table construction logic)

    # --- SECTION 3.0 & NEW CHART ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    # [Insert your Jadual 3 logic here]
    
    # INSERTING THE GRAPH
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("3.2 Trend Mingguan Kes Denggi"), 11, bold=True)
    chart_file = generate_trend_chart_img()
    if chart_file:
        doc.add_picture(chart_file, width=Inches(6.0))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p_cap.add_run("Rajah 1: Trend Kes Mingguan Denggi Negeri Selangor sehingga 2026"), 9, bold=True)

    # --- SECTION 4.0 BKK ---
    # [Insert your Jadual 4 logic here]

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="wide")
st.title("📊 BWKK Report Generator 2026")

col1, col2 = st.columns(2)
with col1:
    f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx"])
with col2:
    f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx"])

if f1 and f2:
    if st.button("🚀 Jana Laporan & Sambung Graf"):
        try:
            # Data Processing (Sama seperti kod asal anda)
            now_msia = get_msia_time()
            today = now_msia.date()
            yesterday = today - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y")

            # 1. Notifikasi
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            # 2. Wabak
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            # ... [Wabak processing logic] ...

            # 3. Vektor (From GSheet)
            raw_gs = pd.read_csv(GSHEET_URL_VEKTOR, header=None)
            mask_v = raw_gs.apply(lambda r: r.astype(str).str.contains('Petaling').any(), axis=1)
            start_row = mask_v.idxmax()
            v_data = raw_gs.iloc[start_row : start_row + 10, 13:20]

            # 4. BKK (From GSheet)
            df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
            # ... [BKK processing logic] ...
            
            # Temporary placeholders for missing variables in this snippet:
            wabak_df = pd.DataFrame() 
            bkk_table_final = pd.DataFrame()
            is_bkk_empty = True
            bkk_details = []

            # Final Doc Generation
            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, is_bkk_empty, bkk_details)
            st.success("✅ Laporan Berjaya Dijana bersama Graf!")
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
            
        except Exception as e:
            st.error(f"Ralat utama: {e}")
