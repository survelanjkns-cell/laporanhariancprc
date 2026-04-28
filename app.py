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

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    day_name = days_ms.get(target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_df, bkk_total_yesterday):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Cm(2.54)
    section.left_margin = section.right_margin = Cm(3.18)

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk Utama
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)
    ]
    doc.add_paragraph()
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    # 3. Jadual Tarikh Hijau
    doc.add_paragraph()
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.width = Inches(5.8) 
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        set_cell_paddings(cell, top=120, bottom=120)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 11, bold=True)

    # --- S1, S2, S3 (Sila masukkan logik jadual 1-3 anda di sini) ---
    # ... (Kod S1, S2, S3 diringkaskan untuk fokus S4) ...

    # --- SECTION 4.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Crisis (BKK)"), 11, bold=True)
    
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

    headers = ["INSIDEN/\nBENCANA"] + BKK_DISTRICTS + ["JUMLAH", "DIISYTIHAR\nOLEH CPRC\nKKM"]
    for i, txt in enumerate(headers):
        cell = t4.rows[0].cells[i]
        set_cell_background(cell, "BFDFFF" if i <= len(BKK_DISTRICTS) else "FFFF00" if i == len(BKK_DISTRICTS)+1 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 7, bold=True)
        set_cell_paddings(cell, top=80, bottom=80)

    for r_idx, (insiden, row_data) in enumerate(bkk_df.iterrows()):
        row = t4.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(str(insiden)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        
        for c_idx, dist in enumerate(BKK_DISTRICTS):
            val = row_data[dist]
            p = row[c_idx+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val)) if val > 0 else "-"), 8, bold=False)

        # Total & KKM
        row[len(BKK_DISTRICTS)+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[len(BKK_DISTRICTS)+1].paragraphs[0].add_run(str(int(row_data['JUMLAH']))), 8, bold=True)
        set_cell_background(row[len(BKK_DISTRICTS)+1], "FFFFB3")

        kkm_val = row_data['KKM_DECLARE']
        row[len(BKK_DISTRICTS)+2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[len(BKK_DISTRICTS)+2].paragraphs[0].add_run(str(int(kkm_val)) if kkm_val > 0 else "-"), 8, bold=False)
        set_cell_background(row[len(BKK_DISTRICTS)+2], "E2EFDA")

    # Final Footer
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.LEFT
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
    if st.button("🚀 Jana Laporan Lengkap (1.0 - 4.0)"):
        try:
            yesterday = date.today() - timedelta(days=1)
            
            # [PROSES S1 & S2 & S3 ANDA...]
            # (Pastikan variable matrix, col_totals, wabak_df, vector_data tersedia)
            
            # --- PROSES S4 (BKK) DENGAN AUTO-CLEAN ---
            with st.spinner('Menarik data BKK...'):
                df_bkk_raw = pd.read_csv(SHEET_BKK_URL)
                # Auto-Clean Column Names (Buang ruang kosong & tukar ke uppercase)
                df_bkk_raw.columns = df_bkk_raw.columns.str.strip().str.upper()
                
                # Gunakan nama kolom yang dibersihkan
                col_tkh = 'TKH LAPOR'
                col_kejadian = 'KEJADIAN'
                col_daerah = 'DAERAH'
                col_kkm = 'KKM DECLARE'

                df_bkk_raw[col_tkh] = pd.to_datetime(df_bkk_raw[col_tkh], dayfirst=True).dt.date
                df_bkk_raw[col_kejadian] = df_bkk_raw[col_kejadian].str.strip().str.upper()
                df_bkk_raw[col_daerah] = df_bkk_raw[col_daerah].str.strip().str.upper()
                
                bkk_total_yesterday = len(df_bkk_raw[df_bkk_raw[col_tkh] == yesterday])
                
                # Crosstab
                matrix_bkk = pd.crosstab(df_bkk_raw[col_kejadian], df_bkk_raw[col_daerah]).reindex(columns=BKK_DISTRICTS, fill_value=0)
                
                # KKM Declare
                kkm_counts = df_bkk_raw[df_bkk_raw[col_kkm].notna()].groupby(col_kejadian).size()
                matrix_bkk['KKM_DECLARE'] = kkm_counts
                matrix_bkk['KKM_DECLARE'] = matrix_bkk['KKM_DECLARE'].fillna(0)
                
                matrix_bkk['JUMLAH'] = matrix_bkk[BKK_DISTRICTS].sum(axis=1)
                matrix_bkk = matrix_bkk.sort_values('JUMLAH', ascending=False)

            # Jana (Pastikan semua parameter dimasukkan)
            # doc_out = generate_docx(matrix, col_totals, wabak_df, vector_data, matrix_bkk, bkk_total_yesterday)
            # st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
            st.success("Logik BKK sedia! Sila cantumkan dengan kod utama anda.")

        except Exception as e:
            st.error(f"Ralat: {e}")
