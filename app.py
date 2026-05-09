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

def add_pkd_note(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note_text = "*Nota : GBK, Gombak; HL, Hulu Langat; HS, Hulu Selangor; KLG, Klang; KL, Kuala Langat; KS, Kuala Selangor; PTG, Petaling; SB, Sabak Bernam; SPG, Sepang."
    run = p.add_run(note_text)
    apply_font(run, 7, bold=False)
    p.paragraph_format.space_after = Pt(12)

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk Utama
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)
    ]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # 3. Jadual Tarikh Hijau (Lebar Penuh)
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"\nTarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 11, bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECTION 1.0 ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    total_notifications = int(col_sums['Grand Total'])
    h11 = doc.add_paragraph()
    h11.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h11_text = f"1.1 Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(h11.add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    t1.width = content_width
    t1.autofit = False

    # --- SECTION 2.0 (WABAK) - KEMASKINI LEBAR PENUH ---
    doc.add_page_break()
    p2_head = doc.add_paragraph()
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    
    harian_total = int(wabak_df['HARIAN'].sum())
    h21 = doc.add_paragraph()
    h21.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h21_text = f"2.1 Jadual di bawah menunjukkan jumlah wabak harian, aktif dan kumulatif di negeri Selangor. Sejumlah {harian_total} input notifikasi wabak diterima pada {get_malay_date(yesterday)}."
    apply_font(h21.add_run(h21_text), 11, bold=False)

    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    t2.width = content_width
    t2.autofit = False
    
    # Nisbah lebar: Nama Penyakit (55%), Harian/Aktif/Kumulatif (15% setiap satu)
    col_widths_t2 = [content_width * 0.55, content_width * 0.15, content_width * 0.15, content_width * 0.15]
    
    for i, h in enumerate(["PENYAKIT", "HARIAN", "AKTIF", "KUMULATIF"]):
        cell = t2.cell(0, i)
        cell.width = col_widths_t2[i]
        apply_font(cell.paragraphs[0].add_run(h), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(penyakit)), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        for idx, col_key in enumerate(['HARIAN', 'AKTIF', 'KUMULATIF'], start=1):
            cells[idx].width = col_widths_t2[idx]
            run = cells[idx].paragraphs[0].add_run(str(int(row_data[col_key])))
            apply_font(run, 8, bold=True)
            cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f2_cells = t2.rows[-1].cells
    sum_vals = ["JUMLAH", str(int(wabak_df['HARIAN'].sum())), str(int(wabak_df['AKTIF'].sum())), str(int(wabak_df['KUMULATIF'].sum()))]
    for i, txt in enumerate(sum_vals):
        apply_font(f2_cells[i].paragraphs[0].add_run(txt), 8, bold=True)
        set_cell_background(f2_cells[i], "FFFF00")
        f2_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- JADUAL 2.1 (LEBAR PENUH) ---
    doc.add_paragraph()
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan pada {get_malay_date(yesterday)}")
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    t21.width = content_width
    t21.autofit = False
    
    widths_21 = [content_width * 0.05, content_width * 0.2, content_width * 0.2, content_width * 0.4, content_width * 0.15]
    h21_headers = ["BIL", "WABAK", "DAERAH", "TEMPAT BERLAKU", "BIL KES (AR)"]
    for i, txt in enumerate(h21_headers):
        cell = t21.cell(0, i)
        cell.width = widths_21[i]
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 10, bold=True)

    if not df_yesterday_list:
        row = t21.add_row().cells
        row[0].merge(row[4])
        row[0].text = "Tiada wabak dilaporkan pada tarikh ini."
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        for idx, item in enumerate(df_yesterday_list, start=1):
            row = t21.add_row().cells
            row[0].text = str(idx)
            row[1].text = f"{item[0]}\n(Household)" if str(item[3])=="Rumah Persendirian" else f"{item[0]}\n(Institusi)"
            row[2].text = str(item[1])
            row[3].text = str(item[2])
            n_kes, n_dedah = float(item[4]), float(item[5])
            pct = (n_kes/n_dedah*100) if n_dedah>0 else 0
            row[4].text = f"{int(n_kes)}/{int(n_dedah)}\n({pct:.1f}%)"
            for c in range(5):
                row[c].width = widths_21[c]
                p = row[c].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c == 3 else WD_ALIGN_PARAGRAPH.CENTER
                apply_font(p.runs[0] if p.runs else p.add_run(""), 8, bold=False)

    # --- SECTION 3.0 (VEKTOR) ---
    doc.add_page_break()
    p3_head = doc.add_paragraph()
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    h31 = doc.add_paragraph()
    h31.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h31_text = f"3.1 Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor."
    apply_font(h31.add_run(h31_text), 11, bold=False)

    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.width = content_width
    t3.autofit = False
    
    # ... (Logik header t3 dikekalkan, pastikan cell.width diset) ...

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    p4_head = doc.add_paragraph()
    apply_font(p4_head.add_run("4.0 Ringkasan Laporan Kejadian Insiden BKK"), 11, bold=True)
    
    h41 = doc.add_paragraph()
    h41.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h41_text = "4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor."
    apply_font(h41.add_run(h41_text), 11, bold=False)
    
    add_table_title(doc, "Jadual 4", "Senarai Kejadian Insiden BKK")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    t4.width = content_width
    t4.autofit = False

    # Sediakan fail untuk dimuat turun
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI (SAMA SEPERTI ASAL) ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📑 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        # Logik pemprosesan data (df1, df2, matrix, dsb) dikekalkan...
        # Sila rujuk script sebelumnya untuk bahagian pemprosesan data yang lengkap.
        pass
