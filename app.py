import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, Inches
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

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    # Page setup
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(0.5)
    section.top_margin = section.bottom_margin = Inches(0.4)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(2.0))

    # 2. Titles
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 12),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 12),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 12)
    ]
    doc.add_paragraph()
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # 3. Green Header Table
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_table.width = Inches(6.8)
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        run.bold = True
        run.font.size = Pt(11)

    # --- SECTION 1.0 ---
    doc.add_paragraph()
    doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi").bold = True
    total_notifications = int(col_sums['Grand Total'])
    h11_text = f"1.1 Sejumlah {total_notifications} input notifikasi telah diterima pada {yesterday.strftime('%d %B %Y')} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    doc.add_paragraph().add_run(h11_text)

    # Table 1 (Matrix)
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    
    # Headers Table 1
    h_cells = t1.rows[0].cells
    h_cells[0].text = "PENYAKIT"
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        cell.text = pkd.replace("PKD ", "")
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.size = Pt(8)
    h_cells[-1].text = "Grand Total"
    set_cell_background(h_cells[-1], "FFFF00")

    # Data rows Table 1
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        row[0].text = str(penyakit)
        set_cell_background(row[0], "D9E9FF")
        row[0].paragraphs[0].runs[0].font.size = Pt(8)
        for c_idx, val in enumerate(row_data):
            cell = row[c_idx+1]
            cell.text = str(int(val))
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if (c_idx + 1) == (len(row_data)): set_cell_background(cell, "FFFFB3")

    # Footer Table 1
    f_cells = t1.rows[-1].cells
    f_cells[0].text = "Grand Total"
    set_cell_background(f_cells[0], "FFFF00")
    for i, val in enumerate(col_sums):
        cell = f_cells[i+1]
        cell.text = str(int(val))
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- SECTION 2.0 ---
    doc.add_page_break()
    doc.add_paragraph().add_run("2.0 Ringkasan Laporan Notifikasi Wabak").bold = True
    
    harian_total = int(wabak_df['HARIAN'].sum())
    yesterday_str = yesterday.strftime('%d %B %Y')
    
    if harian_total > 0:
        h21_text = f"2.1 Sejumlah {harian_total} input notifikasi wabak telah diterima pada {yesterday_str} dengan pecahan mengikut penyakit seperti dalam jadual 2."
    else:
        h21_text = f"2.1 Tiada wabak dilaporkan diterima pada {yesterday_str}."
    
    doc.add_paragraph().add_run(h21_text)

    # Table 2 (Dynamic Wabak Table)
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Headers Table 2
    h2_titles = ["PENYAKIT", "HARIAN", "KUMULATIF"]
    for i, h in enumerate(h2_titles):
        cell = t2.cell(0, i)
        cell.text = h
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True

    # Data Rows Table 2
    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        cells[0].text = str(penyakit)
        set_cell_background(cells[0], "D9E9FF")
        cells[0].paragraphs[0].runs[0].font.size = Pt(9)
        
        cells[1].text = str(int(row_data['HARIAN']))
        cells[2].text = str(int(row_data['KUMULATIF']))
        
        cells[1].vertical_alignment = cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cells[1].paragraphs[0].alignment = cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Footer Table 2
    f2 = t2.rows[-1].cells
    f2[0].text = "JUMLAH"
    set_cell_background(f2[0], "FFFF00")
    f2[1].text = str(int(wabak_df['HARIAN'].sum()))
    f2[2].text = str(int(wabak_df['KUMULATIF'].sum()))
    for c in range(3):
        set_cell_background(f2[c], "FFFF00")
        f2[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        f2[c].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.add_run("Jadual 2 : Senarai Notifikasi Wabak")

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

st.markdown("### Step 1: Upload Required Files")
file1 = st.file_uploader("📂 Upload Daily Notification Excel (Section 1.0)", type="xlsx")
file2 = st.file_uploader("📂 Upload Outbreak Listing Excel (Section 2.0)", type="xlsx")

if file1 and file2:
    st.markdown("---")
    st.success("Both files uploaded! Ready to generate.")
    
    if st.button("🚀 Generate & Download Report"):
        try:
            # --- Section 1.0 Data Processing ---
            df1 = pd.read_excel(file1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix.sum(axis=0)

            # --- Section 2.0 Data Processing ---
            df2 = pd.read_excel(file2)
            today = date.today()
            yesterday = today - timedelta(days=1)
            
            # Epid Week 1 Start Date (Cutoff)
            cutoff_date = date(2026, 1, 4)
            
            # Process Dates
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            
            # CRITICAL FILTER: Only include outbreaks from 04/01/2026 onwards
            df2 = df2[df2['Tarikh Isytihar Wabak'] >= cutoff_date]
            
            # Extract all unique diseases found in the filtered 2026 data
            unique_diseases = df2['PENYAKIT'].unique()
            wabak_summary = []
            
            for dis in unique_diseases:
                if pd.isna(dis): continue
                # Harian: Count matches for yesterday
                h_count = len(df2[(df2['PENYAKIT'] == dis) & (df2['Tarikh Isytihar Wabak'] == yesterday)])
                # Kumulatif: Count all remaining matches (which are now all >= 04/01/26)
                k_count = len(df2[df2['PENYAKIT'] == dis])
                wabak_summary.append({'PENYAKIT': dis, 'HARIAN': h_count, 'KUMULATIF': k_count})
            
            wabak_df = pd.DataFrame(wabak_summary).set_index('PENYAKIT')
            wabak_df = wabak_df.sort_values(by='KUMULATIF', ascending=False)

            # Generate Docx
            doc_out = generate_docx(matrix, col_totals, wabak_df)
            
            st.download_button(
                label="⬇️ Download Full Laporan_BWKK.docx",
                data=doc_out,
                file_name=f"Laporan_BWKK_{date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error processing data: {e}")
else:
    st.info("Upload both files to enable report generation.")
