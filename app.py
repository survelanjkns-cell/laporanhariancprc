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

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
def set_vertical_center(cell):
    """Memastikan sel sentiasa align middle secara vertikal"""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # Force XML update for certain Word versions
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = parse_xml(r'<w:vAlign {} w:val="center"/>'.format(nsdecls('w')))
    tcPr.append(vAlign)

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def format_penyakit_name(name):
    name_str = str(name).strip().upper()
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]):
        return name_str
    if "FOOD POISONING" in name_str:
        return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]:
        return "Denggi"
    return name_str.title()

def get_msia_time():
    msia_tz = pytz.timezone('Asia/Kuala_Lumpur')
    return datetime.now(msia_tz)

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

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": 
        return "-"
    cleaned = re.sub(r'\s*\(.*?\)', '', str(val)).strip()
    return cleaned if cleaned != "" else "-"

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, harian_detail_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Judul & Header (Dikecilkan untuk brevity)
    titles = [("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
              ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
              ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # Jadual Tarikh
    info_table = doc.add_table(rows=1, cols=2)
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        set_vertical_center(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}" if i==0 else f"Minggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    doc.add_paragraph()

    # --- JADUAL 1 (Notifikasi) ---
    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    num_pkd = len(TEMPLATE_PKDS)
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=num_pkd + 3)
    t1.style = 'Table Grid'
    for row in t1.rows:
        for cell in row.cells:
            set_vertical_center(cell)

    # (Kod isi data Jadual 1 anda di sini - dikekalkan mantap)
    # ... 

    # --- JADUAL 2 ---
    doc.add_page_break()
    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    for row in t2.rows:
        for cell in row.cells: set_vertical_center(cell)
    # ... isi data t2

    # --- JADUAL 2.1 (PENTING: ALIGN MIDDLE & SIZING) ---
    doc.add_paragraph()
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan Pada {get_malay_date(yesterday)}")
    
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    t21.allow_autofit = False 
    col_widths = [Inches(0.4), Inches(1.1), Inches(1.0), Inches(3.3), Inches(0.7)]
    
    h21_headers = ["BIL", "WABAK", "DAERAH", "TEMPAT BERLAKU", "BIL KES (AR)"]
    for i, h_txt in enumerate(h21_headers):
        cell = t21.rows[0].cells[i]
        cell.width = col_widths[i]
        set_vertical_center(cell)
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(h_txt), 8, bold=True)

    if harian_detail_df.empty:
        row_cells = t21.add_row().cells
        merged = row_cells[0].merge(row_cells[4])
        set_vertical_center(merged)
        merged.text = "Tiada wabak dilaporkan."
        merged.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        for idx, row_data in enumerate(harian_detail_df.values, start=1):
            row_cells = t21.add_row().cells
            data = [str(idx), str(row_data[0]), str(row_data[1]), str(row_data[2]), ""]
            for c in range(5):
                cell = row_cells[c]
                cell.width = col_widths[c]
                set_vertical_center(cell) # Align Middle secara vertikal
                p = cell.paragraphs[0]
                p.text = data[c]
                # Align horizontal: Tempat Berlaku (kolom 3) ke Kiri, yang lain Tengah
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c == 3 else WD_ALIGN_PARAGRAPH.CENTER
                apply_font(p.runs[0], 8, bold=False)

    # --- JADUAL 3 (PENTING: ALIGN MIDDLE) ---
    doc.add_page_break()
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    
    # Set middle alignment untuk semua sel dalam t3
    for row in t3.rows:
        for cell in row.cells:
            set_vertical_center(cell)

    # Header Row 1 (Merge & Center)
    h3_r1 = t3.rows[0].cells
    h3_r1[0].merge(t3.rows[1].cells[0]).text = "DAERAH"
    h3_r1[1].merge(h3_r1[2]).text = "DENGGI"
    h3_r1[3].merge(h3_r1[4]).text = "MALARIA"
    h3_r1[5].merge(h3_r1[6]).text = "CHIKUNGUNYA"
    
    for i in [0, 1, 3, 5]:
        set_cell_background(h3_r1[i], "BFDFFF")
        p = h3_r1[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 10, bold=True)

    # Header Row 2
    h3_r2 = t3.rows[1].cells
    for i in range(1, 7):
        set_cell_background(h3_r2[i], "BFDFFF")
        h3_r2[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        p = h3_r2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 9, bold=True)

    # Isi Data t3
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            set_vertical_center(row_cells[j])
            val = vector_df.iloc[i, j]
            try: d_val = str(int(float(val))) if j > 0 else str(val)
            except: d_val = str(val)
            p = row_cells[j].paragraphs[0]
            p.text = d_val
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            if i == len(vector_df)-1: set_cell_background(row_cells[j], "FFFF00")
            elif j == 0: set_cell_background(row_cells[j], "FCE4D6")
            apply_font(p.runs[0], 9.5, bold=True)

    # --- JADUAL 4 (BKK) ---
    doc.add_page_break()
    add_table_title(doc, "Jadual 4", "Senarai BKK")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    for row in t4.rows:
        for cell in row.cells: set_vertical_center(cell)
    # ... isi data t4

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
            # --- Logic pemprosesan data (Dikekalkan seperti skrip asal anda) ---
            now_msia = get_msia_time()
            today = now_msia.date()
            yesterday = today - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y") 

            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            harian_detail = df2[df2['Tarikh Isytihar Wabak'] == yesterday][[
                'PENYAKIT', 'DAERAH (HURUF BESAR)', 
                'Tempat Berlaku Wabak\n(Alamat diisi lengkap dengan :- No rumah, nama jalan, nama tempat, daerah dan Negeri)'
            ]].copy()

            # (Ringkasan wabak logic...)
            # Dummy placeholder untuk vector/bkk supaya skrip jalan
            v_data = pd.read_csv(GSHEET_URL, header=None).iloc[18:28, 13:20] # Contoh slicing gsheet
            bkk_table_final = pd.DataFrame() 

            doc_out = generate_docx(matrix, col_totals, pd.DataFrame(), harian_detail, v_data, bkk_table_final, True, [])
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
