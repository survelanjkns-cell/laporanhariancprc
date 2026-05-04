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
    # Mapping Hari & Bulan ke Bahasa Melayu
    hari_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    bulan_ms = {
        "January": "Januari", "February": "Februari", "March": "Mac", "April": "April", 
        "May": "Mei", "June": "Jun", "July": "Julai", "August": "Ogos", 
        "September": "September", "October": "Oktober", "November": "November", "December": "Disember"
    }
    
    day_name = hari_ms.get(target_date.strftime("%A"))
    month_name = bulan_ms.get(target_date.strftime("%B"))
    day_num = target_date.strftime("%#d") # %#d untuk Windows, %-d untuk Linux (buang leading zero)
    year = target_date.strftime("%Y")
    
    return f"{day_num} {month_name} {year} ({day_name})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df):
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

    # 2. Tajuk Utama
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)
    ]
    doc.add_paragraph()
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # 3. Jadual Tarikh Hijau
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.width = Inches(5.8) 
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        set_cell_paddings(cell, top=120, bottom=120)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 11, bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECTION 1.0 ---
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    total_notifications = int(col_sums['Grand Total'])
    h11_text = f"Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(doc.add_paragraph().add_run(h11_text), 10, bold=False)

    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS','PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS','PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    h_cells = t1.rows[0].cells
    for i in range(len(h_cells)):
        set_cell_paddings(h_cells[i], top=140, bottom=140)
        h_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    apply_font(h_cells[0].paragraphs[0].add_run("PENYAKIT"), 7.5, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 7, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(h_cells[-1].paragraphs[0].add_run("Jumlah"), 7, bold=True)
    set_cell_background(h_cells[-1], "FFFF00")

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(str(penyakit)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            cell = row[c_idx+1]
            apply_font(cell.paragraphs[0].add_run(str(int(val))), 8, bold=True)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if (c_idx + 1) == (len(row_data)): set_cell_background(cell, "FFFFB3")

    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 7.5, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, val in enumerate(col_sums):
        cell = f_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(str(int(val))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p1_cap = doc.add_paragraph()
    p1_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p1_cap.add_run("Jadual 1 : Senarai Input eNotifikasi"), 10, bold=False)

    # --- SECTION 2.0 & 3.0 (Wabak & Vektor) ---
    # Sila tambah logik t2 & t3 anda di sini jika perlu, pastikan gunakan get_malay_date(yesterday)

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    total_val_bkk = clean_val(bkk_table_df.iloc[-1]['JUMLAH'])
    prefix_41 = "4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor."
    
    if total_val_bkk == "0" or total_val_bkk == "-":
        h41_text = f"{prefix_41} Tiada Insiden dilaporkan pada {get_malay_date(yesterday)}."
    else:
        h41_text = f"{prefix_41} Sejumlah {total_val_bkk} input notifikasi Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK) telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 4."
    
    apply_font(doc.add_paragraph().add_run(h41_text), 10, bold=False)

    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    h4_widths = [Inches(1.5)] + [Inches(0.42)] * (len(bkk_table_df.columns) - 3) + [Inches(0.6), Inches(0.8)]

    for i, col in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        cell.width = h4_widths[i]
        if i < len(bkk_table_df.columns)-2: set_cell_background(cell, "BFDFFF")
        elif i == len(bkk_table_df.columns)-2: set_cell_background(cell, "FFFF00")
        else: set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_size = 6.5 if 0 < i < len(bkk_table_df.columns)-2 else 7
        apply_font(p.add_run(str(col).replace(" ", "\n")), f_size, bold=True)

    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last_row = (r_idx == len(bkk_table_df)-1)
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_val(val))
            apply_font(run, 7.5 if c_idx == 0 else 8, bold=is_last_row or c_idx == 0)
            if is_last_row: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")

    p4_cap = doc.add_paragraph()
    p4_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p4_cap.add_run("Jadual 4 : Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 10, bold=False)

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI (Hanya Bahagian Trigger) ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            yesterday = date.today() - timedelta(days=1)
            # S1
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            # S4 Logic
            df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
            bkk_raw = df_bkk_full.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_raw.columns = bkk_raw.iloc[0]
            bkk_table_final = bkk_raw[1:].reset_index(drop=True)

            # Placeholder untuk S2/S3
            wabak_placeholder = pd.DataFrame() 
            vector_placeholder = pd.DataFrame()

            doc_out = generate_docx(matrix, col_totals, wabak_placeholder, vector_placeholder, bkk_table_final)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")

        except Exception as e:
            st.error(f"Ralat: {e}")
