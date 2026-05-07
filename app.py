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

# --- KONSTAN & DATA RUJUKAN ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# Data Purata Harian untuk Jadual 1 (Berdasarkan Gambar 2)
DATA_AVERAGE = {
    "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "HIV/AIDS": 7, "Leptospirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1,
    "Chikungunya": 0, "Viral Encephalitis": 0
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID_VEKTOR = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_VEKTOR}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
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

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    return f"{target_date.day:02d} {months_ms[target_date.month]} {target_date.year} ({days_ms[target_date.strftime('%A')]})"

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(1.27)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo & Header
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(logo_path, width=Inches(1.8))

    titles = [("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5), ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5), ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)]
    for text, size in titles:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(text), size, bold=True)
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. Info Box (Green)
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"Minggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    doc.add_paragraph()

    # --- JADUAL 1 (Input eNotifikasi + Average Harian) ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3) # +3: Penyakit, Jumlah, Average
    t1.style = 'Table Grid'
    
    headers = ["PENYAKIT"] + [pkd.replace("PKD ", "") for pkd in TEMPLATE_PKDS] + ["Jumlah", "Average Harian"]
    for i, h in enumerate(headers):
        cell = t1.cell(0, i)
        apply_font(cell.paragraphs[0].add_run(h), 8, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        color = "BFDFFF" if i <= 9 else ("FFFF00" if i == 10 else "FF8C00")
        set_cell_background(cell, color)

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        cells = t1.rows[r_idx + 1].cells
        nama = format_penyakit_name(penyakit)
        apply_font(cells[0].paragraphs[0].add_run(nama), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        
        for c_idx in range(len(TEMPLATE_PKDS)):
            cells[c_idx+1].paragraphs[0].add_run(str(int(row_data[TEMPLATE_PKDS[c_idx]])))
            cells[c_idx+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Jumlah
        apply_font(cells[-2].paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(cells[-2], "FFFFB3")
        cells[-2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Average Harian (Orange)
        avg = DATA_AVERAGE.get(nama, "-")
        apply_font(cells[-1].paragraphs[0].add_run(str(avg)), 8, bold=True)
        set_cell_background(cells[-1], "FF8C00")
        cells[-1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Row Jumlah Footer
    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        f_cells[i+1].paragraphs[0].add_run(str(int(col_sums[pkd])))
        set_cell_background(f_cells[i+1], "FFFF00")
    f_cells[-2].paragraphs[0].add_run(str(int(col_sums['Grand Total'])))
    set_cell_background(f_cells[-2], "FFFF00")
    set_cell_background(f_cells[-1], "FF8C00") # Biar kosong sesuai permintaan

    doc.add_paragraph()

    # --- JADUAL 2 (Notifikasi Wabak + AKTIF) ---
    p2_head = doc.add_paragraph()
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    h2 = ["PENYAKIT", "HARIAN", "AKTIF", "KUMULATIF"]
    for i, h in enumerate(h2):
        cell = t2.cell(0, i)
        apply_font(cell.paragraphs[0].add_run(h), 9, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(penyakit)), 9, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        cells[1].text = str(int(row_data['HARIAN']))
        cells[2].text = str(int(row_data['AKTIF']))
        cells[3].text = str(int(row_data['KUMULATIF']))
        for c in range(1,4): cells[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Footer Jadual 2
    f2 = t2.rows[-1].cells
    f2[0].text = "JUMLAH"
    f2[1].text = str(int(wabak_df['HARIAN'].sum()))
    f2[2].text = str(int(wabak_df['AKTIF'].sum()))
    f2[3].text = str(int(wabak_df['KUMULATIF'].sum()))
    for c in range(4): 
        set_cell_background(f2[c], "FFFF00")
        f2[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # [Nota: Bahagian Vektor dan BKK boleh ditambah di sini mengikut skrip asal anda]

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI & LOGIC ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            now_msia = get_msia_time()
            today = now_msia.date()
            yesterday = today - timedelta(days=1)

            # S1 - Notifikasi Harian
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix.sum(axis=0)

            # S2 - Wabak (Logik AKTIF & Sheet SELANGOR 2)
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            
            # Parsing tarikh
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df2['Tarikh Sebenar Tamat Wabak'] = pd.to_datetime(df2['Tarikh Sebenar Tamat Wabak '], errors='coerce').dt.date
            df2['Tarikh Wabak Dijangka Tamat'] = pd.to_datetime(df2['Tarikh Wabak Dijangka Tamat'], errors='coerce').dt.date

            # Filter Epi Week 1 (4 Jan 2026)
            df2 = df2[df2['Tarikh Isytihar Wabak'] >= date(2026, 1, 4)]
            
            def group_inf(n): return "ILI/ Influenza" if any(x in str(n).upper() for x in ["INFLUENZA", "ILI"]) else n
            df2['PENYAKIT_NORM'] = df2['PENYAKIT'].apply(group_inf)

            wb_sum = []
            for d in df2['PENYAKIT_NORM'].unique():
                if pd.isna(d): continue
                d_df = df2[df2['PENYAKIT_NORM'] == d]
                
                # Harian & Kumulatif
                h = len(d_df[d_df['Tarikh Isytihar Wabak'] == yesterday])
                k = len(d_df)
                
                # Logik AKTIF
                def check_active(row):
                    tamat = row['Tarikh Sebenar Tamat Wabak']
                    if pd.isna(tamat): tamat = row['Tarikh Wabak Dijangka Tamat']
                    # Aktif jika belum ada tarikh tamat atau tarikh tamat >= hari ini
                    return True if (pd.isna(tamat) or tamat >= today) else False
                
                a = d_df.apply(check_active, axis=1).sum()
                wb_sum.append({'PENYAKIT': d, 'HARIAN': h, 'AKTIF': a, 'KUMULATIF': k})
            
            wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values(by='KUMULATIF', ascending=False)

            # S3 & S4 (Placeholder data kosong jika GSheet tidak diakses)
            # Anda boleh masukkan semula kod read_csv(GSHEET_URL) anda di sini
            v_data = pd.DataFrame() 
            bkk_table = pd.DataFrame()

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table, True, [])
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
            
        except Exception as e:
            st.error(f"Ralat: {e}")
