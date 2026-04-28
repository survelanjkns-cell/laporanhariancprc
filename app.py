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
    day_name = days_ms.get(target_date.strftime("%A"), target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

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
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)
    ]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    # 3. Jadual Tarikh Hijau
    doc.add_paragraph()
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.width = Inches(6.0) 
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        set_cell_paddings(cell, top=120, bottom=120)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- SECTION 1.0 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {get_malay_date(yesterday)}."), 10, bold=False)

    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS','PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS','PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    for i, cell in enumerate(t1.rows[0].cells):
        set_cell_paddings(cell, top=140, bottom=140)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_background(cell, "BFDFFF" if i <= len(TEMPLATE_PKDS) else "FFFF00")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = "PENYAKIT" if i == 0 else (pkd_map.get(TEMPLATE_PKDS[i-1], TEMPLATE_PKDS[i-1]) if i <= len(TEMPLATE_PKDS) else "Jumlah")
        apply_font(p.add_run(txt), 7, bold=True)

    for r_idx, (peny, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx+1].cells
        apply_font(row[0].paragraphs[0].add_run(str(peny)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            cell = row[c_idx+1]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val))), 8, bold=True)
            if c_idx == len(row_data)-1: set_cell_background(cell, "FFFFB3")

    for i, val in enumerate(col_sums):
        cell = t1.rows[-1].cells[i+1]
        set_cell_background(cell, "FFFF00")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(int(val))), 8, bold=True)
    apply_font(t1.rows[-1].cells[0].paragraphs[0].add_run("Jumlah"), 7.5, bold=True)
    set_cell_background(t1.rows[-1].cells[0], "FFFF00")

    p1_cap = doc.add_paragraph()
    p1_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p1_cap.add_run("Jadual 1 : Senarai Input eNotifikasi"), 10, bold=False)

    # --- SECTION 2.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    harian_total = int(wabak_df['HARIAN'].sum())
    apply_font(doc.add_paragraph().add_run(f"Sejumlah {harian_total} wabak diterima pada {get_malay_date(yesterday)}."), 10, bold=False)

    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(["PENYAKIT", "HARIAN", "KUMULATIF"]):
        cell = t2.cell(0, i)
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(h), 9, bold=True)

    for i, (peny, row_data) in enumerate(wabak_df.iterrows()):
        row = t2.rows[i+1].cells
        apply_font(row[0].paragraphs[0].add_run(str(peny)), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for j, col in enumerate(['HARIAN', 'KUMULATIF']):
            row[j+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(row[j+1].paragraphs[0].add_run(str(int(row_data[col]))), 8, bold=True)

    f2 = t2.rows[-1].cells
    apply_font(f2[0].paragraphs[0].add_run("JUMLAH"), 9, bold=True)
    apply_font(f2[1].paragraphs[0].add_run(str(int(wabak_df['HARIAN'].sum()))), 9, bold=True)
    apply_font(f2[2].paragraphs[0].add_run(str(int(wabak_df['KUMULATIF'].sum()))), 9, bold=True)
    for c in range(3): 
        set_cell_background(f2[c], "FFFF00")
        f2[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    p2_cap = doc.add_paragraph()
    p2_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p2_cap.add_run("Jadual 2 : Senarai Notifikasi Wabak"), 10, bold=False)

    # --- SECTION 3.0 ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    v_widths = [Inches(2.3)] + [Inches(0.55)] * 6
    
    h3_r1 = t3.rows[0].cells
    h3_r1[0].merge(t3.rows[1].cells[0]).text = "DAERAH"
    h3_r1[1].merge(h3_r1[2]).text = "DENGGI"
    h3_r1[3].merge(h3_r1[4]).text = "MALARIA"
    h3_r1[5].merge(h3_r1[6]).text = "CHIKUNGUNYA"
    for i in [0, 1, 3, 5]:
        set_cell_background(h3_r1[i], "BFDFFF")
        h3_r1[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h3_r1[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 8.5, bold=True)

    h3_r2 = t3.rows[1].cells
    for i in range(1, 7):
        h3_r2[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        set_cell_background(h3_r2[i], "BFDFFF")
        p = h3_r2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 7.5, bold=True)

    for i in range(len(vector_df)):
        row = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            try: d_val = str(int(float(val))) if j > 0 else str(val)
            except: d_val = str(val)
            row[j].width = v_widths[j]
            p = row[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(d_val)
            if i == len(vector_df)-1: set_cell_background(row[j], "FFFF00")
            elif j == 0: set_cell_background(row[j], "FCE4D6")
            apply_font(run, 7.5, bold=True)

    p3_cap = doc.add_paragraph()
    p3_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(p3_cap.add_run("Jadual 3 : Senarai Notifikasi Wabak Vektor"), 10, bold=False)

    # --- SECTION 4.0 ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"Sejumlah {bkk_count} input diterima pada {get_malay_date(yesterday)}."), 10, bold=False)

    t4 = doc.add_table(rows=len(bkk_df) + 2, cols=len(BKK_DISTRICTS) + 3)
    t4.style = 'Table Grid'
    h4_widths = [Inches(1.5)] + [Inches(0.42)] * len(BKK_DISTRICTS) + [Inches(0.6), Inches(0.8)]
    
    h4_headers = ["INSIDEN/\nBENCANA"] + BKK_DISTRICTS + ["JUMLAH", "DIISYTIHAR\nOLEH CPRC\nKKM"]
    for i, txt in enumerate(h4_headers):
        cell = t4.rows[0].cells[i]
        cell.width = h4_widths[i]
        set_cell_background(cell, "BFDFFF" if i <= len(BKK_DISTRICTS) else "FFFF00" if i == len(BKK_DISTRICTS)+1 else "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 6.5 if 0 < i <= len(BKK_DISTRICTS) else 7, bold=True)

    for r_idx, (insiden, row_data) in enumerate(bkk_df.iterrows()):
        row = t4.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(str(insiden)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, dist in enumerate(BKK_DISTRICTS):
            val = row_data[dist]
            row[c_idx+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(row[c_idx+1].paragraphs[0].add_run(str(int(val)) if val > 0 else "-"), 8)
        
        apply_font(row[len(BKK_DISTRICTS)+1].paragraphs[0].add_run(str(int(row_data['JUMLAH']))), 8, bold=True)
        set_cell_background(row[len(BKK_DISTRICTS)+1], "FFFFB3")
        apply_font(row[len(BKK_DISTRICTS)+2].paragraphs[0].add_run(str(int(row_data['KKM_DECLARE']))), 8)
        set_cell_background(row[len(BKK_DISTRICTS)+2], "E2EFDA")

    doc.add_paragraph()
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- UI LOGIC ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian (1.0)", type="xlsx")
f2 = st.file_uploader("📂 Muat Naik Excel Penyenaraian Wabak (2.0)", type="xlsx")

if f1 and f2:
    if st.button("🚀 Jana Laporan Penuh (1.0 - 4.0)"):
        try:
            yesterday = date.today() - timedelta(days=1)
            
            # S1
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            # S2
            df2 = pd.read_excel(f2)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            def group_inf(n): return "ILI/INFLUENZA" if any(x in str(n).upper() for x in ["INFLUENZA", "ILI"]) else n
            df2['PENYAKIT'] = df2['PENYAKIT'].apply(group_inf)
            unique_d = df2['PENYAKIT'].unique()
            wb_sum = []
            for d in unique_d:
                if pd.isna(d): continue
                h = len(df2[(df2['PENYAKIT'] == d) & (df2['Tarikh Isytihar Wabak'] == yesterday)])
                k = len(df2[df2['PENYAKIT'] == d])
                wb_sum.append({'PENYAKIT': d, 'HARIAN': h, 'KUMULATIF': k})
            wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values(by='KUMULATIF', ascending=False)

            # S3
            raw_v = pd.read_csv(SHEET1_URL, header=None)
            s_row = raw_v.apply(lambda r: r.astype(str).str.contains('PETALING').any(), axis=1).idxmax()
            v_data = raw_v.iloc[s_row : s_row + 10, 13:20]

            # S4
            df_bkk = pd.read_csv(SHEET_BKK_URL)
            df_bkk.columns = df_bkk.columns.str.strip().str.upper()
            df_bkk['TKH LAPOR'] = pd.to_datetime(df_bkk['TKH LAPOR'], dayfirst=True).dt.date
            bkk_count = len(df_bkk[df_bkk['TKH LAPOR'] == yesterday])
            matrix_bkk = pd.crosstab(df_bkk['KEJADIAN'], df_bkk['DAERAH']).reindex(columns=[d.split(' ')[-1] if 'PKD' in d else d for d in BKK_DISTRICTS], fill_value=0)
            # Standardize column names for crosstab
            matrix_bkk.columns = BKK_DISTRICTS
            matrix_bkk['KKM_DECLARE'] = df_bkk[df_bkk['KKM DECLARE'].notna()].groupby('KEJADIAN').size().reindex(matrix_bkk.index, fill_value=0)
            matrix_bkk['JUMLAH'] = matrix_bkk[BKK_DISTRICTS].sum(axis=1)

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, matrix_bkk, bkk_count)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
