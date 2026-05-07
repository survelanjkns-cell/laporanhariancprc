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

# --- KONSTAN & MAPPING DATA (DARI GAMBAR) ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# Figure diambil terus dari gambar Jadual 1 yang dibekalkan sebelum ini
AVG_HARIAN_FIGURES = {
    "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "HIV/AIDS": 7, "Leptospirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
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
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo & Tajuk Utama
    titles = [("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
              ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
              ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECTION 1.0 (JADUAL 1) ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    num_pkd = len(TEMPLATE_PKDS)
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=num_pkd + 3)
    t1.style = 'Table Grid'
    
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    h_cells = t1.rows[0].cells
    apply_font(h_cells[0].paragraphs[0].add_run("PENYAKIT"), 8, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(h_cells[num_pkd+1].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(h_cells[num_pkd+1], "FFFF00")
    apply_font(h_cells[num_pkd+2].paragraphs[0].add_run("Average Harian"), 8, bold=True)
    set_cell_background(h_cells[num_pkd+2], "FFC000")

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(format_penyakit_name(penyakit)), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            cell = row[c_idx+1]
            apply_font(cell.paragraphs[0].add_run(str(int(row_data[pkd]))), 8, bold=True)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[num_pkd+1].paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(row[num_pkd+1], "FFFFB3")
        row[num_pkd+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[num_pkd+2].paragraphs[0].add_run(str(int(row_data.get('Average Harian', 0)))), 8, bold=True)
        set_cell_background(row[num_pkd+2], "FFC000")
        row[num_pkd+2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = f_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(str(int(col_sums[pkd]))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(f_cells[num_pkd+1].paragraphs[0].add_run(str(int(col_sums['Grand Total']))), 8, bold=True)
    set_cell_background(f_cells[num_pkd+1], "FFFF00")
    f_cells[num_pkd+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_background(f_cells[num_pkd+2], "FFC000")

    # --- SECTION 2.0 (JADUAL 2 - WABAK) ---
    doc.add_page_break()
    p2_head = doc.add_paragraph()
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    
    harian_total = int(wabak_df['HARIAN'].sum())
    h21 = doc.add_paragraph()
    h21_text = f"2.1 Sejumlah {harian_total} input notifikasi wabak diterima pada {get_malay_date(yesterday)}."
    apply_font(h21.add_run(h21_text), 11, bold=False)

    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    t2.allow_autofit = False
    
    # PELARASAN LEBAR COLUMN (DILARASKAN: Penyakit dikurangkan sedikit)
    col_widths = [Inches(2.8), Inches(1.0), Inches(1.0), Inches(1.0)]
    h2_cols = ["PENYAKIT", "HARIAN", "AKTIF", "KUMULATIF"]
    
    for i, h in enumerate(h2_cols):
        cell = t2.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(h), 9, bold=True)

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        cells[0].width = col_widths[0]
        set_cell_background(cells[0], "D9E9FF")
        p0 = cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        apply_font(p0.add_run(str(penyakit)), 9, bold=True)
        vals = [row_data['HARIAN'], row_data['AKTIF'], row_data['KUMULATIF']]
        for idx, val in enumerate(vals, start=1):
            cells[idx].width = col_widths[idx]
            p = cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val))), 9, bold=True)

    f2_cells = t2.rows[-1].cells
    footer_labels = ["JUMLAH", wabak_df['HARIAN'].sum(), wabak_df['AKTIF'].sum(), wabak_df['KUMULATIF'].sum()]
    for i, val in enumerate(footer_labels):
        f2_cells[i].width = col_widths[i]
        set_cell_background(f2_cells[i], "FFFF00")
        p = f2_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(val) if i == 0 else str(int(val))), 9, bold=True)

    # --- SECTION 3.0 (VEKTOR) & 4.0 (BKK) ---
    doc.add_page_break()
    p3_head = doc.add_paragraph()
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    h3_r1 = t3.rows[0].cells
    h3_r1[0].merge(t3.rows[1].cells[0]).text = "DAERAH"
    h3_r1[1].merge(h3_r1[2]).text = "DENGGI"
    h3_r1[3].merge(h3_r1[4]).text = "MALARIA"
    h3_r1[5].merge(h3_r1[6]).text = "CHIKUNGUNYA"
    for i in [0,1,3,5]:
        set_cell_background(h3_r1[i], "BFDFFF")
        p = h3_r1[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 10.5, bold=True)
    h3_r2 = t3.rows[1].cells
    for i in range(1, 7):
        h3_r2[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        set_cell_background(h3_r2[i], "BFDFFF")
        p = h3_r2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 10.5, bold=True)
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            try: display_val = str(int(float(val))) if j > 0 else str(val)
            except: display_val = str(val)
            p = row_cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(display_val)
            if i == len(vector_df)-1: set_cell_background(row_cells[j], "FFFF00")
            elif j == 0: set_cell_background(row_cells[j], "FCE4D6")
            apply_font(run, 10.5, bold=True)

    doc.add_page_break()
    p4_head = doc.add_paragraph()
    apply_font(p4_head.add_run("4.0 Ringkasan Laporan Kejadian Insiden BKK"), 11, bold=True)
    add_table_title(doc, "Jadual 4", "Senarai Kejadian Insiden BKK")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    for i, col in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(col).replace(" ","\n")), 8, bold=True)
    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last = (r_idx == len(bkk_table_df)-1)
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_val(val))
            apply_font(run, 8, bold=is_last or c_idx == 0)
            if is_last: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : e-notifikasi & Laporan Wabak ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)
    for label in ["Petugas :", "Jawatan :", "\n\nKetua Petugas :", "Jawatan :"]:
        apply_font(doc.add_paragraph().add_run(label), 11, bold=False)

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
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df2['Tarikh Sebenar Tamat Wabak'] = pd.to_datetime(df2['Tarikh Sebenar Tamat Wabak '], errors='coerce').dt.date
            df2 = df2[df2['Tarikh Isytihar Wabak'] >= date(2026, 1, 4)]
            df2['PENYAKIT'] = df2['PENYAKIT'].apply(lambda x: "ILI/ Influenza" if any(y in str(x).upper() for y in ["INFLUENZA", "ILI"]) else x)
            wb_sum = []
            for d in df2['PENYAKIT'].dropna().unique():
                disease_df = df2[df2['PENYAKIT'] == d]
                h = len(disease_df[disease_df['Tarikh Isytihar Wabak'] == yesterday])
                k = len(disease_df)
                active = disease_df.apply(lambda r: True if (pd.isna(r['Tarikh Sebenar Tamat Wabak']) or r['Tarikh Sebenar Tamat Wabak'] >= today) else False, axis=1).sum()
                wb_sum.append({'PENYAKIT': d, 'HARIAN': h, 'AKTIF': active, 'KUMULATIF': k})
            wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values(by='KUMULATIF', ascending=False)

            raw_gs = pd.read_csv(GSHEET_URL, header=None)
            start_row = raw_gs.apply(lambda r: r.astype(str).str.contains('Petaling').any(), axis=1).idxmax()
            v_data = raw_gs.iloc[start_row : start_row + 10, 13:20].dropna(subset=[13])

            df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
            insiden = df_bkk_full[df_bkk_full.iloc[:, 2].astype(str).str.contains(yesterday_str)]
            bkk_details = [{'kejadian': r[5], 'alamat': r[8], 'daerah': r[4]} for _, r in insiden.iterrows()]
            is_bkk_empty = len(bkk_details) == 0
            bkk_raw = df_bkk_full.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_raw.columns = bkk_raw.iloc[0]
            bkk_table_final = bkk_raw[1:].rename(columns={'GOMBAK':'GBK','HULU LANGAT':'HL','HULU SELANGOR':'HS','KLANG':'KLG','KUALA LANGAT':'KL','KUALA SELANGOR':'KS','PETALING':'PTG','SABAK BERNAM':'SB','SEPANG':'SPG'})

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, is_bkk_empty, bkk_details)
            st.download_button("⬇️ Muat Turun Laporan Lengkap", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
