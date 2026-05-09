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

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- FUNCTIONS ---
def get_msia_time():
    msia_tz = pytz.timezone('Asia/Kuala_Lumpur')
    return datetime.now(msia_tz)

def get_malay_date(target_date):
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    return f"{target_date.day:02d} {months_ms[target_date.month]} {target_date.year} ({days_ms[target_date.strftime('%A')]})"

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def format_penyakit_name(name):
    name_str = str(name).strip().upper()
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]): return name_str
    if "FOOD POISONING" in name_str: return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]: return "Denggi"
    return name_str.title()

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

# --- GENERATOR DOCX ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # Tajuk & Info
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)")
    apply_font(run, 11, bold=True)

    # Jadual Tarikh
    t_info = doc.add_table(rows=1, cols=2)
    t_info.width = content_width
    for i, txt in enumerate([f"Tarikh: {get_malay_date(today)}", f"Minggu Epi: {get_epi_week(today)}"]):
        cell = t_info.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p_info = cell.paragraphs[0]
        p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p_info.add_run(txt), 10, bold=True)

    # --- JADUAL 2 (Yang Bermasalah Sebelum Ni) ---
    doc.add_paragraph("\n2.0 Ringkasan Laporan Notifikasi Wabak")
    h21 = doc.add_paragraph()
    h21.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h21.add_run(f"2.1 Jadual di bawah menunjukkan jumlah wabak harian, aktif dan kumulatif di negeri Selangor pada {get_malay_date(yesterday)}.")

    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    t2.width = content_width
    t2.autofit = False
    
    # Set Lebar Kolum (PENTING: Biar Word nampak cantik)
    widths = [content_width * 0.55, content_width * 0.15, content_width * 0.15, content_width * 0.15]
    for i, h in enumerate(["PENYAKIT", "HARIAN", "AKTIF", "KUMULATIF"]):
        cell = t2.cell(0, i)
        cell.width = widths[i]
        set_cell_background(cell, "BFDFFF")
        p_h = cell.paragraphs[0]
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p_h.add_run(h), 8, bold=True)

    # Isi Data Wabak
    for r_idx, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[r_idx+1].cells
        cells[0].text = str(penyakit)
        cells[1].text = str(int(row_data['HARIAN']))
        cells[2].text = str(int(row_data['AKTIF']))
        cells[3].text = str(int(row_data['KUMULATIF']))
        for c_idx in range(4):
            cells[c_idx].width = widths[c_idx]
            p_c = cells[c_idx].paragraphs[0]
            p_c.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p_c.runs[0], 8, bold=False)

    # Simpan ke Memory
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📑 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        with st.spinner("Sedang memproses data..."):
            try:
                now_msia = get_msia_time()
                today = now_msia.date()
                yesterday = today - timedelta(days=1)
                yesterday_str = yesterday.strftime("%d/%m/%Y")

                # --- PROSES DATA ---
                df1 = pd.read_excel(f1)
                df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
                df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
                matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
                matrix['Grand Total'] = matrix.sum(axis=1)
                col_totals = matrix.sum(axis=0)

                df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
                df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
                
                # Kira Ringkasan Wabak
                wb_sum = []
                for d in df2['PENYAKIT'].unique():
                    if pd.isna(d): continue
                    disease_df = df2[df2['PENYAKIT'] == d]
                    h = len(disease_df[disease_df['Tarikh Isytihar Wabak'] == yesterday])
                    k = len(disease_df)
                    wb_sum.append({'PENYAKIT': d, 'HARIAN': h, 'AKTIF': 0, 'KUMULATIF': k})
                wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT')

                # Dummy Data BKK/Vektor untuk demo
                v_data = pd.DataFrame()
                bkk_table = pd.DataFrame()
                
                # --- JANA FAIL ---
                doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table, True, [], [])
                
                # --- PAPAR BUTANG MUAT TURUN ---
                st.success("✅ Laporan Berjaya Dijana!")
                st.download_button(
                    label="⬇️ Muat Turun Laporan .docx",
                    data=doc_out,
                    file_name=f"Laporan_BWKK_{today}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Ralat semasa memproses: {e}")
