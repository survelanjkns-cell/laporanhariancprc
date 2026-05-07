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

# --- KONSTAN ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# Mapping data "Average Harian" daripada gambar yang dibekalkan
AVG_HARIAN_DATA = {
    "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "HIV/AIDS": 7, "Leptospirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
def get_msia_time():
    msia_tz = pytz.timezone('Asia/Kuala_Lumpur')
    return datetime.now(msia_tz)

def format_penyakit_name(name):
    name_str = str(name).strip().upper()
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]):
        return name_str
    if "FOOD POISONING" in name_str:
        return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]:
        return "Denggi"
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
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": 
        return "-"
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
    day_name = days_ms.get(target_date.strftime("%A"), "")
    month_name = months_ms.get(target_date.month, "")
    return f"{target_date.day:02d} {month_name} {target_date.year} ({day_name})"

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
    p.paragraph_format.space_after = Pt(6)

def add_pkd_note(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note_text = "*Nota : GBK, Gombak; HL, Hulu Langat; HS, Hulu Selangor; KLG, Klang; KL, Kuala Langat; KS, Kuala Selangor; PTG, Petaling; SB, Sabak Bernam; SPG, Sepang."
    run = p.add_run(note_text)
    apply_font(run, 7, bold=False)
    p.paragraph_format.space_after = Pt(12)

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo & Tajuk (Diringkaskan untuk penjimatan ruang)
    titles = [("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5), ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5), ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # --- SECTION 1.0 (Jadual 1 dengan Column Tambahan) ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    total_notifications = int(col_sums['Grand Total'])
    h11 = doc.add_paragraph()
    h11_text = f"1.1 Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)}."
    apply_font(h11.add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    
    # Tambah 3 column di hujung: Jumlah, Average Harian
    pkd_count = len(TEMPLATE_PKDS)
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=pkd_count + 3)
    t1.style = 'Table Grid'
    
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    # Header
    h_cells = t1.rows[0].cells
    apply_font(h_cells[0].paragraphs[0].add_run("PENYAKIT"), 8, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    apply_font(h_cells[pkd_count+1].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(h_cells[pkd_count+1], "FFFF00")
    
    # Column Baru: Average Harian
    apply_font(h_cells[pkd_count+2].paragraphs[0].add_run("Average Harian"), 8, bold=True)
    set_cell_background(h_cells[pkd_count+2], "FFC000") # Orange Terang
    h_cells[pkd_count+2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Data Rows
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        nama_f = format_penyakit_name(penyakit)
        apply_font(row[0].paragraphs[0].add_run(nama_f), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        
        # Isi data PKD
        for c_idx in range(pkd_count):
            val = row_data[TEMPLATE_PKDS[c_idx]]
            cell = row[c_idx+1]
            apply_font(cell.paragraphs[0].add_run(str(int(val))), 8, bold=True)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        # Isi Grand Total
        gt_cell = row[pkd_count+1]
        apply_font(gt_cell.paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(gt_cell, "FFFFB3")
        gt_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Isi Average Harian (Warna Orange)
        avg_cell = row[pkd_count+2]
        avg_val = row_data.get('Average Harian', 0)
        apply_font(avg_cell.paragraphs[0].add_run(str(int(avg_val)) if avg_val != 0 else "0"), 8, bold=True)
        set_cell_background(avg_cell, "FFC000") # Orange Terang
        avg_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Footer (Jumlah)
    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = f_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(str(int(col_sums[pkd]))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    # Grand Total Footer
    apply_font(f_cells[pkd_count+1].paragraphs[0].add_run(str(int(col_sums['Grand Total']))), 8, bold=True)
    set_cell_background(f_cells[pkd_count+1], "FFFF00")
    f_cells[pkd_count+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Average Harian Footer (Biarkan Kosong)
    set_cell_background(f_cells[pkd_count+2], "FFC000")

    doc.add_paragraph()
    add_pkd_note(doc)
    
    # (Bahagian Section 2.0, 3.0, 4.0 dikekalkan seperti asal...)
    # [Kod selebihnya sama seperti script asal anda untuk menjimatkan ruang jawapan]
    # ...

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
            now_msia = get_msia_time()
            today = now_msia.date()
            yesterday = today - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y") 

            # S1 - Notifikasi Harian Processing
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            
            # --- LOGIK TAMBAHAN: Masukkan Figure Average Harian ---
            def get_avg(row_name):
                formatted = format_penyakit_name(row_name)
                return AVG_HARIAN_DATA.get(formatted, 0)
            
            matrix['Average Harian'] = [get_avg(idx) for idx in matrix.index]
            # ------------------------------------------------------
            
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            # [Logik S2, S3, S4 diteruskan seperti asal...]
            # Dummy placeholders untuk fungsi generate_docx (Pastikan anda kekalkan logik asal anda di sini)
            wabak_df = pd.DataFrame() # Ganti dengan logik asal anda
            v_data = pd.DataFrame()   # Ganti dengan logik asal anda
            bkk_table_final = pd.DataFrame() # Ganti dengan logik asal anda
            bkk_details = []
            is_bkk_empty = True

            # (Sila masukkan semula logik pemprosesan S2, S3, S4 anda di sini untuk melengkapkan script)
            
            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, is_bkk_empty, bkk_details)
            st.download_button("⬇️ Muat Turun Laporan Lengkap", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
