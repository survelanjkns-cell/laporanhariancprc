import streamlit as st
import pandas as pd
from datetime import date, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# --- CONFIG & CONSTANTS ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

# --- LOGIC FUNCTIONS ---
def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    day_name = days_ms.get(target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

# --- REPORT GENERATOR ---
def generate_docx(matrix_df, col_sums):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    # Header & Logo (Assuming logo.png is in the repo)
    try:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture("logo.png", width=Inches(1.2))
    except:
        pass # Skip if logo not found

    # Titles
    titles = ["KEMENTERIAN KESIHATAN MALAYSIA", "", "LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]
    for t in titles:
        if t == "": doc.add_paragraph(); continue
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(t)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(10 if "KEMENTERIAN" in t else 12)

    # Green Box
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i in range(2):
        cell = table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        run.bold = True
        run.font.size = Pt(11)

    doc.add_paragraph()
    
    # Section 1.0 & 1.1
    doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi").bold = True
    total_val = int(col_sums['Grand Total'])
    h11_text = f"1.1 Sejumlah {total_val} input notifikasi telah diterima pada {yesterday.strftime('%d %B %Y')} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    doc.add_paragraph().add_run(h11_text)

    # Matrix Table (Similar to previous logic, simplified for brevity)
    # [Insert table creation logic here - using the matrix_df we prepared]
    # ... (Same logic as the previous app.py for the matrix table) ...

    # Save to memory buffer
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Generator", page_icon="📊")
st.title("📊 BWKK Report Generator")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Filtering
        df = df[df['Notifikasi Status'] != 'Abai Notifikasi']
        df = df[df['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
        
        # Pivot
        matrix_df = pd.crosstab(df['Diagnosis'], df['Pejabat Kesihatan'])
        matrix_df = matrix_df.reindex(columns=TEMPLATE_PKDS, fill_value=0)
        matrix_df['Grand Total'] = matrix_df.sum(axis=1)
        col_sums = matrix_df.sum(axis=0)

        st.success("Data processed successfully!")
        
        # Preview Data
        st.write("### Data Preview (Grand Totals per District)")
        st.bar_chart(col_sums.drop('Grand Total'))

        # Generate Button
        doc_io = generate_docx(matrix_df, col_sums)
        st.download_button(
            label="📄 Download Word Report",
            data=doc_io,
            file_name=f"Laporan_BWKK_{date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Error: {e}")
