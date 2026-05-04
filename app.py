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

# GSheet Links
SHEET1_URL = "https://docs.google.com/spreadsheets/d/1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w/export?format=csv&gid=0"
# Tab "table 2026" GID: 1342717767 | Tab "2026" GID: 1352807145
SHEET_BKK_DATA_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1352807145"
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

def clean_parentheses(val):
    """Membuang kandungan dalam kurungan, cth: '2 (1)' -> '2'"""
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    cleaned = re.sub(r'\s*\(.*?\)', '', str(val)).strip()
    return cleaned if cleaned != "" else "-"

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, bkk_count_yesterday):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.0)
    section.left_margin = section.right_margin = Cm(2.0)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk Utama
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)
    ]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # 3. Jadual Tarikh Hijau
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.width = Inches(6.0) 
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

    # --- SECTION 1.0 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    total_notifications = int(col_sums['Grand Total'])
    h11_text = f"Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(doc.add_paragraph().add_run(h11_text), 10, bold=False)

    # ... [Kod Jadual 1, 2, 3 kekal sama seperti logik asal anda] ...
    # (Diringkaskan untuk tumpuan pada Seksyen 4.0)

    # --- SECTION 4.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    sub_41 = "4.1 Jadual di bawah menunjukkan jumlah kejadia insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor."
    if bkk_count_yesterday == 0:
        h41_text = f"{sub_41} Tiada Insiden dilaporkan pada {get_malay_date(yesterday)}."
    else:
        h41_text = f"{sub_41} Sejumlah {bkk_count_yesterday} input notifikasi Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK) telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 4."
    
    apply_font(doc.add_paragraph().add_run(h41_text), 10, bold=False)

    # Membina Jadual 4 (Extract dari AH2:AU)
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Sizing Header
    h4_widths = [Inches(1.5)] + [Inches(0.42)] * (len(bkk_table_df.columns)-3) + [Inches(0.6), Inches(0.8)]
    
    for i, col_name in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        set_cell_background(cell, "BFDFFF" if i < len(bkk_table_df.columns)-2 else "FFFF00" if i == len(bkk_table_df.columns)-2 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(col_name).replace(" ", "\n")), 6.5 if 0 < i < len(bkk_table_df.columns)-2 else 7, bold=True)

    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last_row = (r_idx == len(bkk_table_df)-1)
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            cleaned_txt = clean_parentheses(val)
            run = p.add_run(cleaned_txt)
            apply_font(run, 7.5 if c_idx == 0 else 8, bold=is_last_row or c_idx == 0)
            
            if is_last_row: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")
            elif c_idx == len(bkk_table_df.columns)-2: set_cell_background(cells[c_idx], "FFFFB3")
            elif c_idx == len(bkk_table_df.columns)-1: set_cell_background(cells[c_idx], "E2EFDA")

    p4_cap = doc.add_paragraph()
    p4_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p4_cap.add_run("Jadual 4 : Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 10, bold=False)

    doc.add_paragraph()
    footer = doc.add_paragraph()
    # Tarikh sumber dinamik (Today)
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian ", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak ", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            yesterday = date.today() - timedelta(days=1)
            
            # S1 Logic
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            # S4 Logic Part A (Kira insiden semalam dari sheet "2026")
            df_bkk_raw = pd.read_csv(SHEET_BKK_DATA_URL)
            df_bkk_raw.columns = df_bkk_raw.columns.str.strip().str.upper()
            df_bkk_raw['TKH LAPOR'] = pd.to_datetime(df_bkk_raw['TKH LAPOR'], dayfirst=True).dt.date
            bkk_count_yesterday = len(df_bkk_raw[df_bkk_raw['TKH LAPOR'] == yesterday])

            # S4 Logic Part B (Extract AH2:AU dari sheet "table 2026")
            df_bkk_table_full = pd.read_csv(SHEET_BKK_TABLE_URL, header=None)
            # AH:AU adalah column index 33 hingga 46
            bkk_extract = df_bkk_table_full.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_extract.columns = bkk_extract.iloc[0] # Set Header AH2
            bkk_table_final = bkk_extract[1:].reset_index(drop=True)

            # S2 & S3 logic placeholders (Anda perlu masukkan semula logik asalnya di sini)
            # ...

            doc_out = generate_docx(matrix, col_totals, None, None, bkk_table_final, bkk_count_yesterday)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")

        except Exception as e:
            st.error(f"Ralat: {e}")
