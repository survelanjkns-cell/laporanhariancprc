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
def generate_docx(matrix_df, col_sums, wabak_df):
    doc = Document()
    
    # Global Style
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    today = date.today()
    yesterday = today - timedelta(days=1)
    
    # Margins (Top/Bottom: 2.54cm, Left/Right: 3.18cm)
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
    h11_text = f"1.1 Sejumlah {total_notifications} input notifikasi telah diterima pada {yesterday.strftime('%d %B %Y')} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(p11.add_run(h11_text), 10, bold=False)

    # Table 1
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    pkd_map = {
        'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS',
        'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS',
        'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'
    }

    h_cells = t1.rows[0].cells
    run_penyakit = h_cells[0].paragraphs[0].add_run("PENYAKIT")
    apply_font(run_penyakit, 7.5, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        short_pkd = pkd_map.get(pkd, pkd)
        run_pkd = cell.paragraphs[0].add_run(short_pkd)
        apply_font(run_pkd, 7, bold=True) 
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    run_gt = h_cells[-1].paragraphs[0].add_run("Grand Total")
    apply_font(run_gt, 7, bold=True)
    set_cell_background(h_cells[-1], "FFFF00")

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        run_name = row[0].paragraphs[0].add_run(str(penyakit))
        apply_font(run_name, 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            cell = row[c_idx+1]
            run_val = cell.paragraphs[0].add_run(str(int(val)))
            apply_font(run_val, 8, bold=True)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if (c_idx + 1) == (len(row_data)): set_cell_background(cell, "FFFFB3")

    f_cells = t1.rows[-1].cells
    run_fgt = f_cells[0].paragraphs[0].add_run("Grand Total")
    apply_font(run_fgt, 7.5, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, val in enumerate(col_sums):
        cell = f_cells[i+1]
        run_fval = cell.paragraphs[0].add_run(str(int(val)))
        apply_font(run_fval, 8, bold=True)
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
    yesterday_str = yesterday.strftime('%d %B %Y')
    h21_text = f"2.1 {'Sejumlah ' + str(harian_total) + ' input notifikasi wabak' if harian_total > 0 else 'Tiada wabak dilaporkan'} diterima pada {yesterday_str}."
    apply_font(p21.add_run(h21_text), 10, bold=False)

    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    h2_titles = ["PENYAKIT", "HARIAN", "KUMULATIF"]
    for i, h in enumerate(h2_titles):
        cell = t2.cell(0, i)
        run_h2 = cell.paragraphs[0].add_run(h)
        apply_font(run_h2, 9, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        run_p2 = cells[0].paragraphs[0].add_run(str(penyakit))
        apply_font(run_p2, 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        run_h = cells[1].paragraphs[0].add_run(str(int(row_data['HARIAN'])))
        apply_font(run_h, 8, bold=True)
        run_k = cells[2].paragraphs[0].add_run(str(int(row_data['KUMULATIF'])))
        apply_font(run_k, 8, bold=True)
        cells[1].paragraphs[0].alignment = cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f2 = t2.rows[-1].cells
    run_j = f2[0].paragraphs[0].add_run("JUMLAH")
    apply_font(run_j, 9, bold=True)
    set_cell_background(f2[0], "FFFF00")
    for c in range(1, 3):
        val = wabak_df['HARIAN'].sum() if c == 1 else wabak_df['KUMULATIF'].sum()
        run_fv = f2[c].paragraphs[0].add_run(str(int(val)))
        apply_font(run_fv, 9, bold=True)
        set_cell_background(f2[c], "FFFF00")
        f2[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p_cap.add_run("Jadual 2 : Senarai Notifikasi Wabak"), 10, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

file1 = st.file_uploader("📂 Upload Daily Notification Excel", type="xlsx")
file2 = st.file_uploader("📂 Upload Outbreak Listing Excel", type="xlsx")

if file1 and file2:
    if st.button("🚀 Generate Final Report"):
        try:
            # Section 1.0
            df1 = pd.read_excel(file1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix.sum(axis=0)

            # Section 2.0
            df2 = pd.read_excel(file2)
            cutoff = date(2026, 1, 4)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df2 = df2[df2['Tarikh Isytihar Wabak'] >= cutoff]
            
            # --- COMBINATION LOGIC FOR ILI/INFLUENZA ---
            def group_influenza(disease_name):
                name = str(disease_name).upper()
                if "INFLUENZA" in name or "ILI" in name:
                    return "ILI/INFLUENZA"
                return disease_name

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

            doc_out = generate_docx(matrix, col_totals, wabak_df)
            st.download_button("⬇️ Download Final Word Report", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Error: {e}")
