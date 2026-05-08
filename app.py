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

# --- KONSTAN & MAPPING ---
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

GSHEET_URL = "https://docs.google.com/spreadsheets/d/1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w/export?format=csv&gid=0"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
def set_vertical_middle(cell):
    """Memaksa sel menjadi align middle secara menegak menggunakan XML"""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = parse_xml(r'<w:vAlign {} w:val="center"/>'.format(nsdecls('w')))
    tcPr.append(vAlign)

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def get_msia_time():
    msia_tz = pytz.timezone('Asia/Kuala_Lumpur')
    return datetime.now(msia_tz)

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    return f"{target_date.day:02d} {months_ms.get(target_date.month)} {target_date.year} ({days_ms.get(target_date.strftime('%A'))})"

def add_table_title(doc, label, title):
    p = doc.add_paragraph()
    run = p.add_run(f"{label} : ")
    apply_font(run, 11, bold=True)
    run2 = p.add_run(title)
    apply_font(run2, 11, bold=False)
    p.paragraph_format.space_after = Pt(6)

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, harian_detail_df, vector_df, bkk_df, bkk_details):
    doc = Document()
    today = get_msia_time().date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # Header Logo & Title (Ringkas)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p.add_run("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)\nJABATAN KESIHATAN NEGERI SELANGOR"), 10.5, bold=True)

    # Jadual Tarikh
    itb = doc.add_table(rows=1, cols=2)
    itb.width = content_width
    for i, txt in enumerate([f"Tarikh : {get_malay_date(today)}", f"Minggu Epi : {((today - date(2026,1,4)).days // 7) + 1}/2026"]):
        c = itb.cell(0, i)
        set_cell_background(c, "C6E0B4")
        set_vertical_middle(c)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 11, bold=True)

    # Jadual 2.1 (Paling Kritikal)
    doc.add_paragraph()
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan Pada {get_malay_date(yesterday)}")
    
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    t21.allow_autofit = False
    col_widths = [Inches(0.4), Inches(1.1), Inches(1.0), Inches(3.3), Inches(0.7)]
    
    headers = ["BIL", "WABAK", "DAERAH", "TEMPAT BERLAKU", "BIL KES (AR)"]
    for i, h in enumerate(headers):
        cell = t21.rows[0].cells[i]
        cell.width = col_widths[i]
        set_vertical_middle(cell)
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(h), 8, bold=True)

    if harian_detail_df.empty:
        row = t21.add_row().cells
        m = row[0].merge(row[4])
        set_vertical_middle(m)
        m.text = "Tiada wabak dilaporkan."
        m.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        for idx, row_data in enumerate(harian_detail_df.values, start=1):
            row_cells = t21.add_row().cells
            vals = [str(idx), str(row_data[0]), str(row_data[1]), str(row_data[2]), ""]
            for c in range(5):
                cell = row_cells[c]
                cell.width = col_widths[c]
                set_vertical_middle(cell)
                p = cell.paragraphs[0]
                p.text = vals[c]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c == 3 else WD_ALIGN_PARAGRAPH.CENTER
                apply_font(p.runs[0], 8, bold=False)

    # Jadual 3 (Vektor) - Dibersihkan dari nan/repeated headers
    doc.add_page_break()
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=2, cols=7)
    t3.style = 'Table Grid'
    
    # Header Row 1
    h3_1 = t3.rows[0].cells
    h3_1[0].merge(t3.rows[1].cells[0]).text = "DAERAH"
    h3_1[1].merge(h3_1[2]).text = "DENGGI"
    h3_1[3].merge(h3_1[4]).text = "MALARIA"
    h3_1[5].merge(h3_1[6]).text = "CHIKUNGUNYA"
    
    for i in [0,1,3,5]:
        set_cell_background(h3_1[i], "BFDFFF")
        set_vertical_middle(h3_1[i])
        p = h3_1[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 10, bold=True)

    # Header Row 2
    h3_2 = t3.rows[1].cells
    for i in range(1, 7):
        set_cell_background(h3_2[i], "BFDFFF")
        set_vertical_middle(h3_2[i])
        h3_2[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        p = h3_2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 9, bold=True)

    # Isi Data Vektor
    if not vector_df.empty:
        for _, row_data in vector_df.iterrows():
            cells = t3.add_row().cells
            for j, val in enumerate(row_data):
                set_vertical_middle(cells[j])
                txt = "0" if pd.isna(val) or str(val).lower() == 'nan' else str(val)
                p = cells[j].paragraphs[0]
                p.text = txt
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
                apply_font(p.runs[0], 9, bold=(j==0))

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT ---
st.set_page_config(page_title="Report Generator", layout="centered")
st.title("📊 BWKK Report Generator (Fixed)")

f1 = st.file_uploader("Upload Harian", type=["xlsx"])
f2 = st.file_uploader("Upload Linelisting", type=["xlsx"])

if f1 and f2:
    if st.button("🚀 Jana Laporan"):
        try:
            today = get_msia_time().date()
            yesterday = today - timedelta(days=1)
            
            # Read data & filter
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            
            harian_detail = df2[df2['Tarikh Isytihar Wabak'] == yesterday][[
                'PENYAKIT', 'DAERAH (HURUF BESAR)', 
                'Tempat Berlaku Wabak\n(Alamat diisi lengkap dengan :- No rumah, nama jalan, nama tempat, daerah dan Negeri)'
            ]].fillna("-")

            # GSheet Vector Data Cleanup
            raw_v = pd.read_csv(GSHEET_URL, header=None)
            # Mencari baris "Petaling" sebagai sauh
            try:
                start_idx = raw_v[raw_v[13].astype(str).str.contains("Petaling")].index[0]
                v_df = raw_v.iloc[start_idx : start_idx + 10, 13:20].reset_index(drop=True)
            except:
                v_df = pd.DataFrame()

            doc_file = generate_docx(None, None, None, harian_detail, v_df, None, [])
            st.download_button("⬇️ Download Word", doc_file, f"Laporan_{today}.docx")
            st.success("Siap! Sila semak fail anda.")
        except Exception as e:
            st.error(f"Ralat: {e}")
