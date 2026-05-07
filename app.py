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

# --- KONSTAN ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

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

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    day_name = days_ms.get(target_date.strftime("%A"), "")
    month_name = months_ms.get(target_date.month, "")
    return f"{target_date.day:02d} {month_name} {target_date.year} ({day_name})"

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

# --- LOGIK PEMBETULAN AVERAGE (MAPPING DARI GAMBAR) ---
def get_custom_average(penyakit_name):
    # Data tepat mengikut gambar yang diupload
    mapping = {
        "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
        "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
        "Avian Influenza": 8, "HIV/AIDS": 7, "Leptospirosis": 6,
        "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
        "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
    }
    for key, value in mapping.items():
        if key.lower() == penyakit_name.lower():
            return value
    return 0

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo & Tajuk (Diringkaskan untuk kod ini)
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_head.add_run("LAPORAN HARIAN BWKK SELANGOR\nCPRC JABATAN KESIHATAN NEGERI SELANGOR")
    apply_font(run_t, 10.5, bold=True)

    # 2. Jadual Tarikh Hijau
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh: {get_malay_date(today)}" if i == 0 else f"Minggu Epi: {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    doc.add_paragraph()

    # --- SECTION 1.0 (JADUAL UTAMA DENGAN AVERAGE) ---
    p1 = doc.add_paragraph()
    apply_font(p1.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    # Bina Jadual 1 (+3 cols untuk: 1 Penyakit, 9 PKD, 1 Jumlah, 1 Average)
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    
    # Header Jadual
    h_cells = t1.rows[0].cells
    pkd_headers = ["PENYAKIT", "GBK", "HL", "HS", "KLG", "KL", "KS", "PTG", "SB", "SPG", "Jumlah", "Average Harian"]
    
    for i, txt in enumerate(pkd_headers):
        cell = h_cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 8, bold=True)
        if i <= 9: set_cell_background(cell, "BFDFFF") # Biru
        elif i == 10: set_cell_background(cell, "FFFF00") # Kuning
        else: set_cell_background(cell, "FFC000") # Oren Terang

    # Isi Kandungan Data
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        nama_formatted = format_penyakit_name(penyakit)
        
        # Penyakit
        apply_font(row[0].paragraphs[0].add_run(nama_formatted), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        
        # Data PKD
        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            cell = row[c_idx + 1]
            apply_font(cell.paragraphs[0].add_run(str(int(row_data[pkd]))), 8, bold=True)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Column Jumlah
        cell_total = row[10]
        apply_font(cell_total.paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        cell_total.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell_total, "FFFFB3")

        # Column Average Harian (DARI MAPPING GAMBAR)
        cell_avg = row[11]
        avg_val = get_custom_average(nama_formatted)
        apply_font(cell_avg.paragraphs[0].add_run(str(avg_val)), 8, bold=True)
        cell_avg.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_background(cell_avg, "FFD966") # Oren lembut

    # Footer (Baris Jumlah Bawah)
    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    
    for i in range(len(col_sums)):
        cell = f_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(str(int(col_sums[i]))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Kolum Average Harian di baris "Jumlah" dibiarkan KOSONG
    set_cell_background(f_cells[-1], "FFFFFF")

    # (Nota: Section 2, 3, dan 4 boleh ditambah di bawah mengikut keperluan asal anda)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator (Fixed Average)")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx"])

if f1:
    if st.button("🚀 Jana Laporan"):
        try:
            df1 = pd.read_excel(f1)
            # Filter status dan PKD
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            
            # Bina Matrix
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            
            # Totals untuk footer (PKD + Grand Total)
            col_totals = matrix.sum(axis=0)

            # Jana Dokumen (Wabak/Vektor diletakkan None/Empty untuk contoh ini)
            doc_out = generate_docx(matrix, col_totals, None, None, None, True, [])
            
            st.download_button("⬇️ Muat Turun Laporan .docx", data=doc_out, file_name=f"Laporan_BWKK_Terkini.docx")
            st.success("Laporan berjaya dijana dengan Average Harian yang betul!")
        except Exception as e:
            st.error(f"Ralat: {e}")
