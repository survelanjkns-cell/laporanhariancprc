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
BKK_DISTRICTS = ['GOMBAK', 'HULU LANGAT', 'HULU SELANGOR', 'KLANG', 'KUALA LANGAT', 'KUALA SELANGOR', 'PETALING', 'SABAK BERNAM', 'SEPANG', 'PK P.KLANG', 'PK KLIA']

# GSheet Links
SHEET_VECTOR_URL = "https://docs.google.com/spreadsheets/d/1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w/export?format=csv&gid=0"
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

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.0)
    section.left_margin = section.right_margin = Cm(1.5) # Lebarkan margin sedikit untuk jadual BKK

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk Utama
    for text in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 
                 "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 
                 "JABATAN KESIHATAN NEGERI SELANGOR"]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), 10.5, bold=True)
        para.paragraph_format.space_after = Pt(0)

    # 3. Jadual Hijau
    doc.add_paragraph()
    it = doc.add_table(rows=1, cols=2)
    it.alignment = WD_TABLE_ALIGNMENT.CENTER
    it.width = Inches(6.0)
    for i in range(2):
        c = it.cell(0, i)
        set_cell_background(c, "C6E0B4")
        set_cell_paddings(c, top=120, bottom=120)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- 1.0 Input Enotifikasi ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {get_malay_date(yesterday)}."), 10, bold=False)

    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS','PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS','PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    # Headers t1
    h1 = t1.rows[0].cells
    for i in range(len(h1)):
        set_cell_background(h1[i], "BFDFFF" if i <= len(TEMPLATE_PKDS) else "FFFF00")
        p = h1[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = "PENYAKIT" if i == 0 else (pkd_map.get(TEMPLATE_PKDS[i-1], TEMPLATE_PKDS[i-1]) if i <= len(TEMPLATE_PKDS) else "Jumlah")
        apply_font(p.add_run(txt), 7, bold=True)

    # Data t1
    for r_idx, (peny, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx+1].cells
        apply_font(row[0].paragraphs[0].add_run(str(peny)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            p = row[c_idx+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val))), 8, bold=True)
            if c_idx == len(row_data)-1: set_cell_background(row[c_idx+1], "FFFFB3")

    # Footer t1
    f1 = t1.rows[-1].cells
    apply_font(f1[0].paragraphs[0].add_run("Jumlah"), 7.5, bold=True)
    set_cell_background(f1[0], "FFFF00")
    for i, val in enumerate(col_sums):
        set_cell_background(f1[i+1], "FFFF00")
        p = f1[i+1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(int(val))), 8, bold=True)

    # --- 4.0 BKK (DENGAN DATA AH:AU) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    # Ambil baris jumlah terakhir
    bkk_total_val = clean_val(bkk_table_df.iloc[-1, -2]) # Kolom JUMLAH
    h41_text = f"4.1 Jadual di bawah menunjukkan pecahan kes BKK. Sejumlah {bkk_total_val} input diterima pada {get_malay_date(yesterday)}."
    apply_font(doc.add_paragraph().add_run(h41_text), 10, bold=False)

    t4 = doc.add_table(rows=bkk_table_df.shape[0] + 1, cols=bkk_table_df.shape[1])
    t4.style = 'Table Grid'
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Sizing tetap supaya sebaris (Justified)
    h4_widths = [Inches(1.5)] + [Inches(0.42)] * (bkk_table_df.shape[1] - 3) + [Inches(0.6), Inches(0.8)]

    # Headers t4
    for i, col in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        cell.width = h4_widths[i] if i < len(h4_widths) else Inches(0.5)
        set_cell_background(cell, "BFDFFF" if i < bkk_table_df.shape[1]-2 else "FFFF00" if i == bkk_table_df.shape[1]-2 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(col).replace(" ", "\n")), 6.5 if 0 < i < bkk_table_df.shape[1]-2 else 7, bold=True)

    # Data t4
    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last_row = (r_idx == bkk_table_df.shape[0] - 1)
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_val(val))
            apply_font(run, 7.5 if c_idx == 0 else 8, bold=is_last_row or c_idx == 0)
            
            if is_last_row: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")
            elif c_idx == bkk_table_df.shape[1]-2: set_cell_background(cells[c_idx], "FFFFB3")
            elif c_idx == bkk_table_df.shape[1]-1: set_cell_background(cells[c_idx], "E2EFDA")

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- UI & DATA LOGIC ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian (1.0)", type="xlsx")
f2 = st.file_uploader("📂 Muat Naik Excel Penyenaraian Wabak (2.0)", type="xlsx")

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap (1.0 - 4.0)"):
        try:
            # S1
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            # S4 (Data dari AH2:AU)
            df_bkk_raw = pd.read_csv(SHEET_BKK_TABLE_URL, header=None)
            # Ambil AH2:AU (Col 33 to 46)
            bkk_extract = df_bkk_raw.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_extract.columns = bkk_extract.iloc[0] # Set Header dari AH2
            bkk_final = bkk_extract[1:].reset_index(drop=True)

            # Jana
            doc_out = generate_docx(matrix, col_totals, None, None, bkk_final)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
