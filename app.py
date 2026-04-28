import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os
import re

# --- KONSTAN ---
TEMPLATE_PKDS = ['PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG', 'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 'PKD SABAK BERNAM', 'PKD SEPANG']
# GSheet S3 (Vector) & S4 (BKK Table)
SHEET1_URL = "https://docs.google.com/spreadsheets/d/1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w/export?format=csv&gid=0"
# GID untuk "table 2026" adalah 1342717767
SHEET_BKK_TABLE_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
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

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    day_name = days_ms.get(target_date.strftime("%A"), target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

def clean_parentheses(val):
    """Membuang kandungan di dalam kurungan, contoh: '1 (1)' -> '1'"""
    if pd.isna(val): return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.0)
    section.left_margin = section.right_margin = Cm(2.0)

    # 1. Logo (Pastikan fail logo.png.jpg wujud di GitHub)
    if os.path.exists("logo.png.jpg"):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture("logo.png.jpg", width=Inches(1.8))

    # 2. Tajuk Utama
    for text in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 
                 "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 
                 "JABATAN KESIHATAN NEGERI SELANGOR"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(text), 10.5, bold=True)
        p.paragraph_format.space_after = Pt(0)

    # 3. Jadual Hijau
    doc.add_paragraph()
    it = doc.add_table(rows=1, cols=2)
    it.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i in range(2):
        c = it.cell(0, i)
        set_cell_background(c, "C6E0B4")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {((today - date(2026, 1, 4)).days // 7) + 1}/{today.year}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- SECTION 4.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor pada {get_malay_date(yesterday)}."), 10, bold=False)

    # Membina Jadual 4 dari DataFrame GSheet
    t4 = doc.add_table(rows=bkk_table_df.shape[0] + 1, cols=bkk_table_df.shape[1])
    t4.style = 'Table Grid'
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header Row
    for i, col_name in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        set_cell_background(cell, "BFDFFF" if i < bkk_table_df.shape[1]-2 else "FFFF00")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(col_name).replace(" ", "\n")), 7, bold=True)

    # Data Rows
    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            
            # Jika baris terakhir (JUMLAH), tebalkan & kuning
            is_last_row = (r_idx == bkk_table_df.shape[0] - 1)
            txt_clean = clean_parentheses(val)
            
            run = p.add_run(txt_clean)
            apply_font(run, 7.5 if c_idx == 0 else 8, bold=is_last_row or c_idx == 0)
            
            if is_last_row: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian (1.0)", type="xlsx")
f2 = st.file_uploader("📂 Muat Naik Excel Penyenaraian Wabak (2.0)", type="xlsx")

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # S1 logic
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            # S4 Logic - Extract Range AH:AU (Columns 33 to 46 approx)
            # Memandangkan GSheet CSV extract seluruh sheet, kita perlu filter columns
            df_bkk_full = pd.read_csv(SHEET_BKK_TABLE_URL)
            
            # Mencari index AH hingga AU secara dinamik atau manual
            # Dalam pandas read_csv, AH biasanya adalah column index ke-33
            # Kita guna iloc untuk ambil subset AH2:AU
            bkk_table_data = df_bkk_full.iloc[:, 33:47] # Adjust index jika perlu
            bkk_table_data.columns = bkk_table_data.iloc[0] # Set AH2 as header
            bkk_table_data = bkk_table_data[1:].reset_index(drop=True)
            # Buang baris yang kosong sepenuhnya
            bkk_table_data = bkk_table_data.dropna(how='all', subset=[bkk_table_data.columns[0]])

            doc_out = generate_docx(matrix, col_totals, None, None, bkk_table_data)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
