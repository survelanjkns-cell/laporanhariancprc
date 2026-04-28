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

# --- KONSTAN ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"

# --- HELPERS ---
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    day_name = days_ms.get(target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.54)
    section.left_margin = section.right_margin = Cm(3.18)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(2.0))

    # 2. Titles (1.0 & 2.0 - Ringkasan untuk penjimatan ruang)
    # ... (Kod S1 & S2 anda kekal sama seperti sebelumnya) ...

    # --- 3.0 Ringkasan Laporan Wabak Vektor ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    try:
        xx = int(float(vector_df.iloc[-1, 1]) + float(vector_df.iloc[-1, 3]) + float(vector_df.iloc[-1, 5]))
    except: xx = 0

    h31_text = f"Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. Sejumlah {xx} input notifikasi wabak vektor telah diterima pada {yesterday.strftime('%d %B %Y')} dengan pecahan mengikut penyakit seperti dalam jadual 3."
    apply_font(doc.add_paragraph().add_run(h31_text), 10, bold=False)

    # Bina Jadual 3 (SOPHISTICATED LOOK)
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.autofit = True # Memastikan saiz sel seimbang

    # Setup Lebar Kolom secara manual untuk 'Justified Look'
    widths = [Inches(1.5), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7), Inches(0.7)]
    for row in t3.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    # Header Row 1 (Merged)
    h3_row1 = t3.rows[0].cells
    h3_row1[0].text = "DAERAH"
    h3_row1[1].merge(h3_row1[2]).text = "DENGGI"
    h3_row1[3].merge(h3_row1[4]).text = "MALARIA"
    h3_row1[5].merge(h3_row1[6]).text = "CHIKUNGUNYA"
    
    for i in [0, 1, 3, 5]:
        set_cell_background(h3_row1[i], "BFDFFF")
        h3_row1[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h3_row1[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 9, bold=True)

    # Header Row 2 (Sub-headers)
    h3_row2 = t3.rows[1].cells
    for i in range(7):
        if i == 0:
            h3_row2[i].text = "DAERAH"
        else:
            h3_row2[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        
        set_cell_background(h3_row2[i], "BFDFFF")
        h3_row2[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h3_row2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 8, bold=True)

    # Data Rows
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            try:
                display_val = str(int(float(val))) if j > 0 else str(val)
            except:
                display_val = str(val)
            
            row_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = row_cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(display_val)
            
            # Warna & Font
            if i == len(vector_df)-1: # Baris JUMLAH
                set_cell_background(row_cells[j], "FFFF00")
                apply_font(run, 9, bold=True)
            elif j == 0: # Kolom DAERAH
                set_cell_background(row_cells[j], "FCE4D6")
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT # Justify Left untuk daerah
                apply_font(run, 8, bold=True)
            else:
                apply_font(run, 8, bold=True)

    doc.add_paragraph()
    p3_cap = doc.add_paragraph()
    p3_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p3_cap.add_run("Jadual 3 : Senarai Notifikasi Wabak Vektor"), 10, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- UI & LOGIC (Kekal sama dengan pembaikan ralat sebelumnya) ---
# ... (Gunakan kod dari jawapan sebelumnya untuk S1, S2, dan GSheet Processing) ...
