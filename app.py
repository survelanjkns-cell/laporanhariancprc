import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os

# --- CONSTANTS ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# Google Sheet Export Link (CSV Format)
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

# --- HELPERS ---
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    day_name = days_ms.get(target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    # Page setup
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(2.0))

    # 2. Header Titles
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 11),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 11),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 11)
    ]
    doc.add_paragraph()
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Green Header Table
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_table.width = Inches(5.8) 
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 10, bold=True)

    # --- SECTION 1.0 ---
    doc.add_paragraph()
    p10 = doc.add_paragraph()
    apply_font(p10.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    total_notifications = int(col_sums['Grand Total'])
    p11 = doc.add_paragraph()
    h11_text = f"Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {yesterday.strftime('%d %B %Y')} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(p11.add_run(h11_text), 10, bold=False)

    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS','PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS','PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}

    h_cells = t1.rows[0].cells
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

    doc.add_paragraph()
    p1_cap = doc.add_paragraph()
    p1_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p1_cap.add_run("Jadual 1 : Senarai Input eNotifikasi"), 10, bold=False)

    # --- SECTION 2.0 ---
    doc.add_page_break()
    p20 = doc.add_paragraph()
    apply_font(p20.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    
    harian_total = int(wabak_df['HARIAN'].sum())
    p21 = doc.add_paragraph()
    h21_text = f"Jadual di bawah menunjukkan jumlah wabak harian dan kumulatif di negeri Selangor. {'Sejumlah ' + str(harian_total) + ' input notifikasi wabak' if harian_total > 0 else 'Tiada wabak dilaporkan'} diterima pada {yesterday.strftime('%d %B %Y')}."
    apply_font(p21.add_run(h21_text), 10, bold=False)

    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for i, h in enumerate(["PENYAKIT", "HARIAN", "KUMULATIF"]):
        cell = t2.cell(0, i)
        apply_font(cell.paragraphs[0].add_run(h), 9, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(penyakit)), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        apply_font(cells[1].paragraphs[0].add_run(str(int(row_data['HARIAN']))), 8, bold=True)
        apply_font(cells[2].paragraphs[0].add_run(str(int(row_data['KUMULATIF']))), 8, bold=True)
        cells[1].paragraphs[0].alignment = cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f2 = t2.rows[-1].cells
    apply_font(f2[0].paragraphs[0].add_run("JUMLAH"), 9, bold=True)
    set_cell_background(f2[0], "FFFF00")
    apply_font(f2[1].paragraphs[0].add_run(str(int(wabak_df['HARIAN'].sum()))), 9, bold=True)
    apply_font(f2[2].paragraphs[0].add_run(str(int(wabak_df['KUMULATIF'].sum()))), 9, bold=True)
    for c in range(3):
        set_cell_background(f2[c], "FFFF00")
        f2[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    p2_cap = doc.add_paragraph()
    p2_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p2_cap.add_run("Jadual 2 : Senarai Notifikasi Wabak"), 10, bold=False)

    # --- SECTION 3.0 ---
    doc.add_page_break()
    p30 = doc.add_paragraph()
    apply_font(p30.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    xx = int(vector_df.iloc[-1, 1] + vector_df.iloc[-1, 3] + vector_df.iloc[-1, 5])
    p31 = doc.add_paragraph()
    h31_text = f"Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. Sejumlah {xx} input notifikasi wabak vektor telah diterima pada {yesterday.strftime('%d %B %Y')} dengan pecahan mengikut penyakit seperti dalam jadual 3."
    apply_font(p31.add_run(h31_text), 10, bold=False)

    t3 = doc.add_table(rows=13, cols=7)
    t3.style = 'Table Grid'
    t3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Row 1 (Merged Headers)
    h3 = t3.rows[0].cells
    apply_font(h3[0].paragraphs[0].add_run("DAERAH"), 9, bold=True)
    h3[1].merge(h3[2]).text = "DENGGI"
    h3[3].merge(h3[4]).text = "MALARIA"
    h3[5].merge(h3[6]).text = "CHIKUNGUNYA"
    for i in [0, 1, 3, 5]:
        set_cell_background(h3[i], "BFDFFF")
        h3[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(h3[i].paragraphs[0].runs[0], 9, bold=True)

    # Row 2 (Sub-headers)
    for i in range(7):
        cell = t3.rows[1].cells[i]
        txt = "DAERAH" if i==0 else ("HARIAN" if i%2!=0 else "KUM")
        apply_font(cell.paragraphs[0].add_run(txt), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Data Rows
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            apply_font(row_cells[j].paragraphs[0].add_run(str(val) if j==0 else str(int(val))), 8, bold=True)
            row_cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == len(vector_df)-1: set_cell_background(row_cells[j], "FFFF00")
            elif j == 0: set_cell_background(row_cells[j], "FCE4D6")

    doc.add_paragraph()
    p3_cap = doc.add_paragraph()
    p3_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p3_cap.add_run("Jadual 3 : Senarai Notifikasi Wabak Vektor"), 10, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

file1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian (1.0)", type="xlsx")
file2 = st.file_uploader("📂 Muat Naik Excel Penyenaraian Wabak (2.0)", type="xlsx")

if file1 and file2:
    if st.button("🚀 Jana Laporan Lengkap (1.0 + 2.0 + 3.0)"):
        try:
            # 1.0 Data
            df1 = pd.read_excel(file1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix.sum(axis=0)

            # 2.0 Data
            df2 = pd.read_excel(file2)
            cutoff = date(2026, 1, 4)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df2 = df2[df2['Tarikh Isytihar Wabak'] >= cutoff]
            def group_influenza(name): return "ILI/INFLUENZA" if any(x in str(name).upper() for x in ["INFLUENZA", "ILI"]) else name
            df2['PENYAKIT'] = df2['PENYAKIT'].apply(group_influenza)
            
            yesterday = date.today() - timedelta(days=1)
            unique_d = df2['PENYAKIT'].unique()
            wabak_summary = []
            for d in unique_d:
                if pd.isna(d): continue
                h = len(df2[(df2['PENYAKIT'] == d) & (df2['Tarikh Isytihar Wabak'] == yesterday)])
                k = len(df2[df2['PENYAKIT'] == d])
                wabak_summary.append({'PENYAKIT': d, 'HARIAN': h, 'KUMULATIF': k})
            wabak_df = pd.DataFrame(wabak_summary).set_index('PENYAKIT').sort_values(by='KUMULATIF', ascending=False)

            # 3.0 Data (Google Sheet)
            raw_gsheet = pd.read_csv(GSHEET_URL)
            vector_data = raw_gsheet.iloc[19:31, 13:20] 

            # GENERATE
            doc_out = generate_docx(matrix, col_totals, wabak_df, vector_data)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat Data: {e}")
