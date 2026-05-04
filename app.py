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
TEMPLATE_PKDS = ['PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG', 'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 'PKD SABAK BERNAM', 'PKD SEPANG']
SHEET_VECTOR_URL = "https://docs.google.com/spreadsheets/d/1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w/export?format=csv&gid=0"
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

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

# --- DOCX GENERATOR ---
def generate_docx(m1, s1, m2, m3, m4):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.0)
    section.left_margin = section.right_margin = Cm(1.5)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk Utama
    for text in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 
                 "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(text), 10.5, bold=True)
        p.paragraph_format.space_after = Pt(0)

    # 3. Jadual Hijau
    doc.add_paragraph()
    it = doc.add_table(rows=1, cols=2)
    it.alignment = WD_TABLE_ALIGNMENT.CENTER
    it.width = Inches(6.0)
    for i in range(2):
        c = it.cell(0, i)
        set_cell_background(c, "C6E0B4")
        set_cell_paddings(c, top=120, bottom=120)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {((today - date(2026, 1, 4)).days // 7) + 1}/{today.year}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- SECTION 1.0 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"Sejumlah {int(s1['Grand Total'])} input notifikasi diterima pada {get_malay_date(yesterday)}."), 10, bold=False)

    t1 = doc.add_table(rows=len(m1) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    pkd_m = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS','PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS','PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    for i, cell in enumerate(t1.rows[0].cells):
        set_cell_background(cell, "BFDFFF" if i <= len(TEMPLATE_PKDS) else "FFFF00")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = "PENYAKIT" if i == 0 else (pkd_m.get(TEMPLATE_PKDS[i-1], 'PKD') if i <= len(TEMPLATE_PKDS) else "Jumlah")
        apply_font(p.add_run(txt), 7, bold=True)
        set_cell_paddings(cell, top=140, bottom=140)

    for r_idx, (peny, row_data) in enumerate(m1.iterrows()):
        row = t1.rows[r_idx+1].cells
        apply_font(row[0].paragraphs[0].add_run(str(peny)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            p = row[c_idx+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val))), 8, bold=True)
            if c_idx == len(row_data)-1: set_cell_background(row[c_idx+1], "FFFFB3")

    p1_cap = doc.add_paragraph()
    p1_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p1_cap.add_run("Jadual 1 : Senarai Input eNotifikasi"), 10, bold=False)

    # --- SECTION 2.0 & 3.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    t2 = doc.add_table(rows=len(m2) + 2, cols=3)
    t2.style = 'Table Grid'
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header t2... [Kod t2 diringkaskan, sila isi mengikut logik asal]

    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    t3 = doc.add_table(rows=len(m3) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Header t3... [Kod t3 diringkaskan, sila isi mengikut logik asal]

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    bkk_total = clean_val(m4.iloc[-1, -2])
    h41 = f"4.1 Jadual di bawah menunjukkan pecahan kes BKK. Sejumlah {bkk_total} input diterima pada {get_malay_date(yesterday)}."
    apply_font(doc.add_paragraph().add_run(h41), 10, bold=False)

    t4 = doc.add_table(rows=len(m4) + 1, cols=len(m4.columns))
    t4.style = 'Table Grid'
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    h4_w = [Inches(1.5)] + [Inches(0.42)] * (len(m4.columns)-3) + [Inches(0.6), Inches(0.8)]
    
    for i, col in enumerate(m4.columns):
        cell = t4.rows[0].cells[i]
        cell.width = h4_w[i] if i < len(h4_w) else Inches(0.5)
        set_cell_background(cell, "BFDFFF" if i < len(m4.columns)-2 else "FFFF00" if i == len(m4.columns)-2 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(col).replace(" ", "\n")), 6.5 if 0 < i < len(m4.columns)-2 else 7, bold=True)

    for r_idx, row_data in enumerate(m4.values):
        cells = t4.rows[r_idx+1].cells
        is_last = (r_idx == len(m4)-1)
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_val(val))
            apply_font(run, 7.5 if c_idx == 0 else 8, bold=is_last or c_idx == 0)
            if is_last: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

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
    if st.button("Jana Laporan Lengkap"):
        try:
            yesterday = date.today() - timedelta(days=1)
            # S1
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            m1 = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            m1['Grand Total'] = m1.sum(axis=1)
            s1_totals = m1.sum(axis=0)

            # S2 (Placeholder) - Anda boleh masukkan logik wabak anda di sini
            m2 = pd.DataFrame({'PENYAKIT': ['HFMD'], 'HARIAN': [0], 'KUMULATIF': [100]}).set_index('PENYAKIT')

            # S3 (Vector)
            raw_v = pd.read_csv(SHEET_VECTOR_URL, header=None)
            s_row = raw_v.apply(lambda r: r.astype(str).str.contains('PETALING').any(), axis=1).idxmax()
            m3 = raw_v.iloc[s_row : s_row + 10, 13:20]

            # S4 (BKK)
            df_bkk_raw = pd.read_csv(SHEET_BKK_TABLE_URL, header=None)
            bkk_extract = df_bkk_raw.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_extract.columns = bkk_extract.iloc[0]
            m4 = bkk_extract[1:].reset_index(drop=True)

            doc_out = generate_docx(m1, s1_totals, m2, m3, m4)
            st.download_button("⬇️ Muat Turun Laporan Lengkap", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
