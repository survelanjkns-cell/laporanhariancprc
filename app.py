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
    # Standard Epid 1 start for 2026 as per user logic
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {
        "Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu",
        "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"
    }
    day_name = days_ms.get(target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums):
    doc = Document()
    
    # Page setup
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(0.5)
    section.top_margin = section.bottom_margin = Inches(0.4)

    # 1. Logo Handling (3x Bigger as requested)
    # Note: Updated to match your GitHub filename 'logo.png.jpg'
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(3.0))
    else:
        st.warning(f"Logo '{logo_path}' not found in the root directory.")

    # 2. Header Titles (KKM line removed as requested)
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 12),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 12),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 12)
    ]
    
    # Spacer after logo
    doc.add_paragraph()

    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        para.paragraph_format.space_after = Pt(0)

    # Spacing between titles and green box
    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # 3. Green Table Box
    today = date.today()
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

    doc.add_paragraph()

    # 4. Section 1.0
    doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi").bold = True
    yesterday = today - timedelta(days=1)
    total_notifications = int(col_sums['Grand Total'])
    
    # Text uses yesterday's date
    yesterday_str = yesterday.strftime('%d %B %Y')
    h11_text = f"1.1 Sejumlah {total_notifications} input notifikasi telah diterima pada {yesterday_str} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    doc.add_paragraph().add_run(h11_text)

    # 5. The Matrix Table
    num_rows = len(matrix_df) + 2
    num_cols = len(TEMPLATE_PKDS) + 2
    
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = 'Table Grid'
    
    # Header Row
    header_cells = table.rows[0].cells
    header_cells[0].text = "PENYAKIT"
    set_cell_background(header_cells[0], "BFDFFF")
    header_cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    header_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_cells[0].paragraphs[0].runs[0].font.bold = True
    
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = header_cells[i+1]
        cell.text = pkd.replace("PKD ", "") # Shorter names
        set_cell_background(cell, "BFDFFF")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.font.size = Pt(8)
        run.font.bold = True
        
    header_cells[-1].text = "Grand Total"
    set_cell_background(header_cells[-1], "FFFF00")
    header_cells[-1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    header_cells[-1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_cells[-1].paragraphs[0].runs[0].font.bold = True

    # Data Rows (Input is sorted descending)
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = table.rows[r_idx + 1].cells
        
        # Col 0: Penyakit (Left Aligned)
        row[0].text = str(penyakit)
        set_cell_background(row[0], "D9E9FF")
        row[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_name = row[0].paragraphs[0]
        p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_name.runs[0].font.size = Pt(8)
        
        # Values (Center Aligned)
        for c_idx, val in enumerate(row_data):
            cell = row[c_idx + 1]
            cell.text = str(int(val))
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p_val = cell.paragraphs[0]
            p_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_val = p_val.runs[0]
            run_val.font.size = Pt(9)
            
            # Highlight Grand Total column
            if (c_idx + 1) == (num_cols - 1):
                set_cell_background(cell, "FFFFB3")
                run_val.font.bold = True

    # Footer Row: Grand Total
    footer_cells = table.rows[-1].cells
    footer_cells[0].text = "Grand Total"
    set_cell_background(footer_cells[0], "FFFF00")
    footer_cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    footer_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_cells[0].paragraphs[0].runs[0].font.bold = True
    
    for i, val in enumerate(col_sums):
        cell = footer_cells[i+1]
        cell.text = str(int(val))
        set_cell_background(cell, "FFFF00")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_foot = cell.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_foot.runs[0].font.bold = True

    doc.add_paragraph()
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.add_run("Jadual 1 : Senarai Input Enotifikasi")

    # Buffer for download
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT PAGE ---
st.set_page_config(page_title="BWKK Gen", layout="centered")

st.markdown("""
    <style>
    .main { background-color: #f5f5f5; }
    .stButton>button { width: 100%; background-color: #28a745; color: white; }
    </style>
    """, unsafe_allow_ Harris=True)

st.title("📊 BWKK Report Generator")
st.info("Upload the Excel file with columns: 'Diagnosis', 'Pejabat Kesihatan', and 'Notifikasi Status'.")

uploaded_file = st.file_uploader("Choose Excel File", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        # 1. Filter Out Abai Notifikasi
        df = df[df['Notifikasi Status'] != 'Abai Notifikasi']
        
        # 2. Extract only required PKDs
        df = df[df['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
        
        # 3. Create Pivot Matrix
        matrix = pd.crosstab(df['Diagnosis'], df['Pejabat Kesihatan'])
        matrix = matrix.reindex(columns=TEMPLATE_PKDS, fill_value=0)
        
        # 4. Sorting logic
        matrix['Grand Total'] = matrix.sum(axis=1)
        matrix = matrix.sort_values(by='Grand Total', ascending=False)
        
        # 5. Column Totals
        col_totals = matrix.sum(axis=0)

        # 6. Generate Docx
        final_doc = generate_docx(matrix, col_totals)
        
        st.success("Report generated successfully!")
        st.download_button(
            label="⬇️ Download Final Word Report",
            data=final_doc,
            file_name=f"Laporan_BWKK_{date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
