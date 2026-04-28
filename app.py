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

# --- KONSTAN ---
TEMPLATE_PKDS = ['GOMBAK', 'HULU LANGAT', 'HULU SELANGOR', 'KLANG', 'KUALA LANGAT', 'KUALA SELANGOR', 'PETALING', 'SABAK BERNAM', 'SEPANG']
BKK_DISTRICTS = TEMPLATE_PKDS + ['PK P.KLANG', 'PK KLIA']

# GSheet Links
SHEET1_URL = "https://docs.google.com/spreadsheets/d/1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w/export?format=csv&gid=0"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1352807145"

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

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_df, bkk_total_yesterday):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    # Page setup
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.0)
    section.left_margin = section.right_margin = Cm(2.0)

    # [Logo, Tajuk Utama, Jadual Hijau, S1, S2, S3 kekal sama seperti kod sebelumnya]
    # (Diringkaskan untuk fokus pada S4)

    # --- SECTION 4.0 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    yesterday_str = yesterday.strftime('%d %B %Y')
    if bkk_total_yesterday > 0:
        h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. Sejumlah {bkk_total_yesterday} input notifikasi Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK) telah diterima pada {yesterday_str} dengan pecahan mengikut penyakit seperti dalam jadual 4."
    else:
        h41_text = f"4.1 Tiada Kejadian Insiden dilaporkan pada {yesterday_str}."
    apply_font(doc.add_paragraph().add_run(h41_text), 10, bold=False)

    # Jadual 4
    t4 = doc.add_table(rows=len(bkk_df) + 2, cols=len(BKK_DISTRICTS) + 3)
    t4.style = 'Table Grid'
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    h_cells = t4.rows[0].cells
    headers = ["INSIDEN/\nBENCANA"] + BKK_DISTRICTS + ["JUMLAH", "DIISYTIHAR\nOLEH CPRC\nKKM"]
    for i, txt in enumerate(headers):
        cell = h_cells[i]
        set_cell_background(cell, "BFDFFF" if i <= len(BKK_DISTRICTS) else "FFFF00" if i == len(BKK_DISTRICTS)+1 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(txt)
        apply_font(run, 7, bold=True)
        set_cell_paddings(cell, top=100, bottom=100)

    # Data Rows
    for r_idx, (insiden, row_data) in enumerate(bkk_df.iterrows()):
        row = t4.rows[r_idx + 1].cells
        # Nama Insiden
        p_name = row[0].paragraphs[0]
        apply_font(p_name.add_run(str(insiden)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        
        # Districts
        for c_idx in range(len(BKK_DISTRICTS)):
            val = row_data[BKK_DISTRICTS[c_idx]]
            p = row[c_idx+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val)) if val > 0 else "-"), 8, bold=False)

        # Jumlah
        p_total = row[len(BKK_DISTRICTS)+1].paragraphs[0]
        p_total.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p_total.add_run(str(int(row_data['JUMLAH']))), 8, bold=True)
        set_cell_background(row[len(BKK_DISTRICTS)+1], "FFFFB3")

        # KKM Declare
        p_kkm = row[len(BKK_DISTRICTS)+2].paragraphs[0]
        p_kkm.alignment = WD_ALIGN_PARAGRAPH.CENTER
        kkm_val = row_data['KKM_DECLARE']
        apply_font(p_kkm.add_run(str(int(kkm_val)) if kkm_val > 0 else "-"), 8, bold=False)
        set_cell_background(row[len(BKK_DISTRICTS)+2], "E2EFDA")

    # Bottom Jumlah Row
    f_cells = t4.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("JUMLAH"), 7.5, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i in range(1, len(headers)):
        col_name = headers[i].replace("\n", " ")
        if "DIISYTIHAR" in col_name: col_name = "KKM_DECLARE"
        
        total_val = bkk_df[col_name].sum() if col_name in bkk_df.columns else 0
        p = f_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(int(total_val))), 8, bold=True)
        set_cell_background(f_cells[i], "FFFF00")

    p4_cap = doc.add_paragraph()
    p4_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p4_cap.add_run("Jadual 4 : Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 10, bold=False)

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({today.strftime('%d %B %Y')} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian (1.0)", type="xlsx")
f2 = st.file_uploader("📂 Muat Naik Excel Penyenaraian Wabak (2.0)", type="xlsx")

if f1 and f2:
    if st.button("🚀 Jana Laporan Penuh (1.0 - 4.0)"):
        try:
            yesterday = date.today() - timedelta(days=1)
            
            # [Proses S1, S2, S3 kekal...]
            # (Contoh ringkas untuk S1 & S2)
            df1 = pd.read_excel(f1)
            # ... (logik crosstab sedia ada)

            # --- PROSES S4 (BKK) ---
            with st.spinner('Menarik data BKK...'):
                df_bkk = pd.read_csv(SHEET_BKK_URL)
                # Cleaning
                df_bkk['TKH LAPOR'] = pd.to_datetime(df_bkk['TKH LAPOR'], dayfirst=True).dt.date
                df_bkk['KEJADIAN'] = df_bkk['KEJADIAN'].str.upper()
                df_bkk['DAERAH'] = df_bkk['DAERAH'].str.upper()
                
                # Kira XX (Total Notifikasi Semalam)
                bkk_total_yesterday = len(df_bkk[df_bkk['TKH LAPOR'] == yesterday])
                
                # Sediakan Matrix (Tapis Tahun 2026 jika perlu, tapi kita ambil semua dalam sheet)
                # Matrix Daerah
                matrix_bkk = pd.crosstab(df_bkk['KEJADIAN'], df_bkk['DAERAH']).reindex(columns=BKK_DISTRICTS, fill_value=0)
                
                # KKM Declare (Lajur N)
                kkm_counts = df_bkk[df_bkk['KKM DECLARE'].notna()].groupby('KEJADIAN').size()
                matrix_bkk['KKM_DECLARE'] = kkm_counts
                matrix_bkk['KKM_DECLARE'] = matrix_bkk['KKM_DECLARE'].fillna(0)
                
                matrix_bkk['JUMLAH'] = matrix_bkk[BKK_DISTRICTS].sum(axis=1)
                matrix_bkk = matrix_bkk.sort_values('JUMLAH', ascending=False)

            # Jana Document
            # doc_out = generate_docx(matrix, col_totals, wabak_df, vector_data, matrix_bkk, bkk_total_yesterday)
            # ...
            st.success("Laporan berjaya dijana!")
            
        except Exception as e:
            st.error(f"Ralat: {e}")
