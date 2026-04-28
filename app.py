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
TEMPLATE_PKDS = ['PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG', 'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 'PKD SABAK BERNAM', 'PKD SEPANG']
BKK_DISTRICTS = ['GOMBAK', 'HULU LANGAT', 'HULU SELANGOR', 'KLANG', 'KUALA LANGAT', 'KUALA SELANGOR', 'PETALING', 'SABAK BERNAM', 'SEPANG', 'PK P.KLANG', 'PK KLIA']

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

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    return target_date.strftime(f"%d %B %Y ({days_ms.get(target_date.strftime('%A'))})")

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_df, bkk_count):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.0)
    section.left_margin = section.right_margin = Cm(2.0)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk Utama
    titles = [("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
              ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
              ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    # 3. Jadual Hijau
    doc.add_paragraph()
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        set_cell_paddings(cell, top=120, bottom=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- S1 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {yesterday.strftime('%d %B %Y')}..."), 10, bold=False)
    
    # [Logik Bina Jadual 1, 2, 3 di sini...]
    # Sila rujuk kod S1, S2, S3 anda yang terdahulu untuk mengisi bahagian ini.

    # --- SECTION 4.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana... Sejumlah {bkk_count} input telah diterima pada {get_malay_date(yesterday)}." if bkk_count > 0 else f"4.1 Tiada Kejadian Insiden dilaporkan pada {get_malay_date(yesterday)}."
    apply_font(doc.add_paragraph().add_run(h41_text), 10, bold=False)

    t4 = doc.add_table(rows=len(bkk_df) + 2, cols=len(BKK_DISTRICTS) + 3)
    t4.style = 'Table Grid'
    headers = ["INSIDEN/\nBENCANA"] + BKK_DISTRICTS + ["JUMLAH", "DIISYTIHAR\nOLEH CPRC\nKKM"]
    for i, txt in enumerate(headers):
        cell = t4.rows[0].cells[i]
        set_cell_background(cell, "BFDFFF" if i <= len(BKK_DISTRICTS) else "FFFF00" if i == len(BKK_DISTRICTS)+1 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 7, bold=True)

    for r_idx, (insiden, row_data) in enumerate(bkk_df.iterrows()):
        row = t4.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(str(insiden)), 7, bold=True)
        for c_idx, dist in enumerate(BKK_DISTRICTS):
            val = row_data[dist]
            row[c_idx+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(row[c_idx+1].paragraphs[0].add_run(str(int(val)) if val > 0 else "-"), 8)
        
        apply_font(row[len(BKK_DISTRICTS)+1].paragraphs[0].add_run(str(int(row_data['JUMLAH']))), 8, bold=True)
        set_cell_background(row[len(BKK_DISTRICTS)+1], "FFFFB3")
        apply_font(row[len(BKK_DISTRICTS)+2].paragraphs[0].add_run(str(int(row_data['KKM_DECLARE']))), 8)

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
            # Proses S1, S2, S3...
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            # Proses S4
            df_bkk_raw = pd.read_csv(SHEET_BKK_URL)
            df_bkk_raw.columns = df_bkk_raw.columns.str.strip().str.upper()
            df_bkk_raw['TKH LAPOR'] = pd.to_datetime(df_bkk_raw['TKH LAPOR'], dayfirst=True).dt.date
            
            bkk_count = len(df_bkk_raw[df_bkk_raw['TKH LAPOR'] == yesterday])
            matrix_bkk = pd.crosstab(df_bkk_raw['KEJADIAN'], df_bkk_raw['DAERAH']).reindex(columns=BKK_DISTRICTS, fill_value=0)
            matrix_bkk['KKM_DECLARE'] = df_bkk_raw[df_bkk_raw['KKM DECLARE'].notna()].groupby('KEJADIAN').size().reindex(matrix_bkk.index, fill_value=0)
            matrix_bkk['JUMLAH'] = matrix_bkk[BKK_DISTRICTS].sum(axis=1)

            # GSheet S3 (Vector)
            raw_v = pd.read_csv(SHEET1_URL, header=None)
            start_row = raw_v.apply(lambda r: r.astype(str).str.contains('PETALING').any(), axis=1).idxmax()
            v_data = raw_v.iloc[start_row : start_row + 10, 13:20]

            # Jana Document
            doc_out = generate_docx(matrix, col_totals, None, v_data, matrix_bkk, bkk_count)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
