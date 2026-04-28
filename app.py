import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os

# --- CONSTANTS ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# Google Sheet Export Link (CSV Format)
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

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
    
    # Page setup (Normal Margins)
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

    # 1. Logo & Titles (Section 1.0 & 2.0 logic same as before...)
    # [Internal logic for 1.0 and 2.0 remains here - keeping concise for display]
    # ... (Same as previous script) ...

    # --- SECTION 3.0 ---
    doc.add_page_break()
    p30 = doc.add_paragraph()
    apply_font(p30.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    # Calculate XX based on total Harian in Vector DF
    # Assuming Vector DF columns: [Daerah, Denggi_H, Denggi_K, Malaria_H, Malaria_K, Chiku_H, Chiku_K]
    xx = int(vector_df.iloc[-1, 1] + vector_df.iloc[-1, 3] + vector_df.iloc[-1, 5])
    
    p31 = doc.add_paragraph()
    yesterday_str = yesterday.strftime('%d %B %Y')
    h31_text = f"3.1 Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. Sejumlah {xx} input notifikasi wabak vektor telah diterima pada {yesterday_str} dengan pecahan mengikut penyakit seperti dalam jadual 3."
    apply_font(p31.add_run(h31_text), 10, bold=False)

    # Table 3
    t3 = doc.add_table(rows=13, cols=7) # Header + 10 PKD + Jumlah + Subheader
    t3.style = 'Table Grid'
    t3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Table 3 Header Row 1 (Merged Categories)
    h3 = t3.rows[0].cells
    h3[0].text = "DAERAH"
    h3[1].merge(h3[2]).text = "DENGGI"
    h3[3].merge(h3[4]).text = "MALARIA"
    h3[5].merge(h3[6]).text = "CHIKUNGUNYA"
    for i in [0, 1, 3, 5]:
        set_cell_background(h3[i], "BFDFFF")
        p = h3[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 9, bold=True)

    # Table 3 Subheader Row (HARIAN / KUM)
    sub_h = t3.rows[1].cells
    set_cell_background(sub_h[0], "BFDFFF")
    for i in range(1, 7):
        sub_h[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        set_cell_background(sub_h[i], "BFDFFF")
        p = sub_h[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 8, bold=True)

    # Fill Table 3 Data from Google Sheet Extract
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            row_cells[j].text = str(val) if j == 0 else str(int(val))
            
            # Formatting
            cell_p = row_cells[j].paragraphs[0]
            cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if i == len(vector_df) - 1: # Last row (JUMLAH)
                set_cell_background(row_cells[j], "FFFF00")
                apply_font(cell_p.runs[0], 9, bold=True)
            elif j == 0: # First column (DAERAH)
                set_cell_background(row_cells[j], "FCE4D6")
                apply_font(cell_p.runs[0], 8, bold=True)
            else:
                apply_font(cell_p.runs[0], 8, bold=True)

    doc.add_paragraph()
    p3_cap = doc.add_paragraph()
    p3_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p3_cap.add_run("Jadual 3 : Senarai Notifikasi Wabak Vektor"), 10, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

file1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian (1.0)", type="xlsx")
file2 = st.file_uploader("📂 Muat Naik Excel Penyenaraian Wabak (2.0)", type="xlsx")

if file1 and file2:
    if st.button("🚀 Jana Laporan Lengkap (1.0 + 2.0 + 3.0)"):
        try:
            # --- 1.0 & 2.0 Processing ---
            # (Same logic as before...)
            
            # --- 3.0 Google Sheet Extraction ---
            with st.spinner('Menarik data dari Google Sheet...'):
                # Read specific range N21:T32 (Index 13 to 19 in zero-indexed CSV)
                # We read the whole sheet and slice the dataframe
                raw_gsheet = pd.read_csv(GSHEET_URL)
                # Slice N21:T32. Adjusting for 0-indexed pandas and CSV structure:
                # Column N is index 13, T is index 19. Rows 21-32 is 20-31.
                vector_data = raw_gsheet.iloc[19:31, 13:20] 
                vector_data.columns = ['Daerah', 'D_H', 'D_K', 'M_H', 'M_K', 'C_H', 'C_K']

            # Call generator with new vector_data
            doc_out = generate_docx(matrix, col_totals, wabak_df, vector_data)
            st.download_button("⬇️ Muat Turun Laporan Lengkap", data=doc_out, file_name=f"Laporan_BWKK_Penuh_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat Data: {e}")
