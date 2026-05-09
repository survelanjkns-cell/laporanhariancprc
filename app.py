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

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    return f"{target_date.day:02d} {months_ms.get(target_date.month, '')} {target_date.year} ({days_ms.get(target_date.strftime('%A'), '')})"

def add_table_title(doc, label, title):
    p = doc.add_paragraph()
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

    # Tajuk Utama
    titles = [
        "LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)",
        "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)",
        "JABATAN KESIHATAN NEGERI SELANGOR"
    ]
    for text in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), 10.5, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # Info Box (Green)
    it = doc.add_table(rows=1, cols=2)
    it.width = content_width
    for i in range(2):
        cell = it.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"Minggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- 1.0 Ringkasan eNotifikasi ---
    doc.add_paragraph()
    p1 = doc.add_paragraph()
    apply_font(p1.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    p1sub = doc.add_paragraph()
    apply_font(p1sub.add_run(f"1.1 Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {get_malay_date(yesterday)}."), 11, bold=False)
    
    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    
    # Header T1
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    h_cells = t1.rows[0].cells
    apply_font(h_cells[0].paragraphs[0].add_run("PENYAKIT"), 8, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        apply_font(h_cells[i+1].paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 8, bold=True)
        set_cell_background(h_cells[i+1], "BFDFFF")
    apply_font(h_cells[-2].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(h_cells[-2], "FFFF00")
    apply_font(h_cells[-1].paragraphs[0].add_run("Avg"), 8, bold=True)
    set_cell_background(h_cells[-1], "FFC000")

    # Data T1
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx+1].cells
        apply_font(row[0].paragraphs[0].add_run(format_penyakit_name(penyakit)), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            p = row[c_idx+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(row_data[pkd]))), 8, bold=True)
        
        row[-2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[-2].paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(row[-2], "FFFFB3")
        
        row[-1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[-1].paragraphs[0].add_run(str(int(row_data.get('Average Harian', 0)))), 8, bold=True)
        set_cell_background(row[-1], "FFC000")

    # Footer T1
    f_row = t1.rows[-1].cells
    apply_font(f_row[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_row[0], "FFFF00")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        p = f_row[i+1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(int(col_sums[pkd]))), 8, bold=True)
        set_cell_background(f_row[i+1], "FFFF00")
    f_row[-2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(f_row[-2].paragraphs[0].add_run(str(int(col_sums['Grand Total']))), 8, bold=True)
    set_cell_background(f_row[-2], "FFFF00")

    # --- 2.0 & 2.1 Wabak ---
    doc.add_page_break()
    p2 = doc.add_paragraph()
    apply_font(p2.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    p2sub = doc.add_paragraph()
    apply_font(p2sub.add_run(f"2.1 Jadual di bawah menunjukkan jumlah wabak di negeri Selangor."), 11, bold=False)
    
    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    # (Data T2 sedia ada anda...)
    for i, h in enumerate(["PENYAKIT", "HARIAN", "AKTIF", "KUMULATIF"]):
        cell = t2.cell(0, i)
        set_cell_background(cell, "BFDFFF")
        apply_font(cell.paragraphs[0].add_run(h), 8, bold=True)

    # (Logik data T2 sama seperti skrip asal anda)

    # JADUAL 2.1
    doc.add_paragraph()
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan pada {get_malay_date(yesterday)}")
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    t21.allow_autofit = False
    set_repeat_table_header(t21.rows[0])
    
    widths = [Inches(0.4), Inches(1.2), Inches(1.2), Inches(3.0), Inches(0.8)]
    for i, txt in enumerate(["BIL", "WABAK", "DAERAH", "TEMPAT BERLAKU", "BIL KES (AR)"]):
        cell = t21.cell(0, i)
        cell.width = widths[i]
        set_cell_background(cell, "BFDFFF")
        set_cell_paddings(cell, top=140, bottom=140)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 10, bold=True)

    for idx, item in enumerate(df_yesterday_list, start=1):
        row = t21.add_row().cells
        for i in range(5): 
            row[i].width = widths[i]
            set_cell_paddings(row[i], top=100, bottom=100, left=100, right=100)
            row[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        
        row[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[0].paragraphs[0].add_run(str(idx)), 8, bold=False)
        
        # Wabak + Category
        p_w = row[1].paragraphs[0]
        p_w.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p_w.add_run(str(item[0])), 8, bold=False)
        apply_font(p_w.add_run(f"\n({'(Household)' if str(item[3]).strip() == 'Rumah Persendirian' else '(Institusi)'})"), 8, bold=False)
        
        row[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[2].paragraphs[0].add_run(str(item[1])), 8, bold=False)
        
        row[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        apply_font(row[3].paragraphs[0].add_run(str(item[2])), 8, bold=False)
        
        # AR calculation
        n_k, n_d = float(item[4] or 0), float(item[5] or 0)
        pct = (n_k/n_d*100) if n_d > 0 else 0
        pct_s = "100%" if pct == 100 else f"{pct:.2f}%"
        p_ar = row[4].paragraphs[0]
        p_ar.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p_ar.add_run(f"{int(n_k)}/{int(n_d)}\n({pct_s})"), 8, bold=False)

    # --- 3.0 Vektor ---
    doc.add_page_break()
    p3 = doc.add_paragraph()
    apply_font(p3.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    # (Logik Jadual 3 anda...)

    # --- 4.0 BKK ---
    doc.add_page_break()
    p4 = doc.add_paragraph()
    apply_font(p4.add_run("4.0 Ringkasan Laporan Kejadian BKK"), 11, bold=True)
    # (Logik Jadual 4 anda...)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Generator", layout="centered")
st.title("📋 BWKK Report Generator")

f1 = st.file_uploader("Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            now = get_msia_time()
            today, yesterday = now.date(), now.date() - timedelta(days=1)
            
            # S1
            df1 = pd.read_excel(f1)
            df1 = df1[(df1['Notifikasi Status'] != 'Abai Notifikasi') & (df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS))]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            # S2
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            
            addr_c = 'Tempat Berlaku Wabak\n(Alamat diisi lengkap dengan :- No rumah, nama jalan, nama tempat, daerah dan Negeri)'
            cat_c = 'Kategori Tempat\n(Kategori premis berdasarkan tempat berlaku wabak)'
            df_y = df2[df2['Tarikh Isytihar Wabak'] == yesterday].copy()
            df_y_list = df_y[['PENYAKIT', 'DAERAH (HURUF BESAR)', addr_c, cat_c, 'Bilangan Kes', 'Bilangan Terdedah']].values.tolist()

            # Kumulatif (Placeholder logik anda)
            wabak_df = pd.DataFrame(columns=['HARIAN', 'AKTIF', 'KUMULATIF']) 

            # Jana
            # Nota: Sila pastikan v_data & bkk_table ditarik mengikut skrip asal anda.
            # Saya letakkan placeholder supaya butang muat turun muncul.
            v_data, bkk_table, bkk_details = pd.DataFrame(), pd.DataFrame(), []
            
            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table, True, bkk_details, df_y_list)
            st.success("✅ Laporan dijana!")
            st.download_button("⬇️ Muat Turun", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
