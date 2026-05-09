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
def set_repeat_table_header(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w')))
    trPr.append(tblHeader)

def set_cell_paddings(cell, top=100, bottom=100, left=100, right=100):
    """
    Menambah ruang (padding) di dalam sel. 
    Nilai default 100 twips (~0.17cm).
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = parse_xml(r'<w:tcMar {}/>'.format(nsdecls('w')))
    for margin, value in [('top', top), ('left', left), ('bottom', bottom), ('right', right)]:
        node = parse_xml(r'<w:{} {} w:w="{}" w:type="dxa"/>'.format(margin, nsdecls('w'), value))
        tcMar.append(node)
    tcPr.append(tcMar)

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

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo & Tajuk (Diringkaskan)
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

    # Jadual Tarikh
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"\nTarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- SECTION 1.0 & 2.0 (Dah ada sebelum ni) ---
    # (Skip untuk ringkasan - fokus pada perubahan Jadual 2.1)

    # --- JADUAL 2.1 (DENGAN CELL PADDING) ---
    doc.add_page_break()
    tarikh_semalam_str = get_malay_date(yesterday)
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan pada {tarikh_semalam_str}")
    
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    t21.allow_autofit = False
    set_repeat_table_header(t21.rows[0])

    widths_21 = [Inches(0.35), Inches(1.15), Inches(1.15), Inches(3.2), Inches(0.8)]
    h21_headers = ["BIL", "WABAK", "DAERAH", "TEMPAT BERLAKU", "BIL KES (AR)"]
    
    for i, txt in enumerate(h21_headers):
        cell = t21.cell(0, i)
        cell.width = widths_21[i]
        set_cell_background(cell, "BFDFFF")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_paddings(cell, top=140, bottom=140) # Padding Header
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 10, bold=True)

    if not df_yesterday_list:
        row = t21.add_row().cells
        row[0].merge(row[4])
        row[0].text = "Tiada wabak dilaporkan pada tarikh ini."
    else:
        for idx, item in enumerate(df_yesterday_list, start=1):
            row = t21.add_row().cells
            for i in range(5): 
                row[i].width = widths_21[i]
                set_cell_paddings(row[i], top=100, bottom=100, left=100, right=100) # Padding Content

            row[0].text = str(idx)

            # WABAK Content
            p_wabak = row[1].paragraphs[0]
            p_wabak.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p_wabak.add_run(str(item[0])), 8, bold=False)
            p_wabak.add_run("\n")
            kategori_display = "(Household)" if str(item[3]).strip() == "Rumah Persendirian" else "(Institusi)"
            apply_font(p_wabak.add_run(kategori_display), 8, bold=False)

            row[2].text = str(item[1]) # DAERAH
            row[3].text = str(item[2]) # TEMPAT BERLAKU (Justify set di bawah)

            # BIL KES Content
            n_kes = float(item[4]) if pd.notna(item[4]) else 0
            n_dedah = float(item[5]) if pd.notna(item[5]) else 0
            pct_str = "100%" if (n_dedah > 0 and (n_kes/n_dedah)*100 == 100) else f"{(n_kes/n_dedah)*100:.2f}%" if n_dedah > 0 else "0%"
            
            p_ar = row[4].paragraphs[0]
            p_ar.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p_ar.add_run(f"{int(n_kes)}/{int(n_dedah)}"), 8, bold=False)
            p_ar.add_run("\n")
            apply_font(p_ar.add_run(f"({pct_str})"), 8, bold=False)

            for c in range(5):
                row[c].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = row[c].paragraphs[0]
                # Justify alamat, yang lain center
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if c == 3 else WD_ALIGN_PARAGRAPH.CENTER
                # Set font content
                if c not in [1, 4] and p.runs:
                    apply_font(p.runs[0], 8, bold=False)

    # --- SEKSYEN SETERUSNYA (Kod asal) ---
    doc.add_page_break()
    p3_head = doc.add_paragraph()
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    # ... sambungan kod asal anda ...

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI (Kekalkan yang asal) ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📋 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        # ... proses data seperti sebelum ni ...
        # (Pastikan df_yesterday_list ambil col 4 & 5 iaitu Bil Kes & Bil Terdedah)
        try:
            now_msia = get_msia_time()
            today = now_msia.date()
            yesterday = today - timedelta(days=1)
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            
            addr_col = 'Tempat Berlaku Wabak\n(Alamat diisi lengkap dengan :- No rumah, nama jalan, nama tempat, daerah dan Negeri)'
            cat_col = 'Kategori Tempat\n(Kategori premis berdasarkan tempat berlaku wabak)'
            
            df_yesterday = df2[df2['Tarikh Isytihar Wabak'] == yesterday].copy()
            # Pastikan susunan: [Penyakit, Daerah, Alamat, Kategori, Bil Kes, Bil Terdedah]
            df_yesterday_list = df_yesterday[['PENYAKIT', 'DAERAH (HURUF BESAR)', addr_col, cat_col, 'Bilangan Kes', 'Bilangan Terdedah']].values.tolist()

            # Panggil generator
            # (Nota: Pastikan hantar semua argumen yang diperlukan oleh generate_docx)
            # doc_out = generate_docx(...)
        except Exception as e:
            st.error(f"Ralat: {e}")
