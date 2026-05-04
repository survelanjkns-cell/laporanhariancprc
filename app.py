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
import re

# --- KONSTAN ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING', 
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
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
    days_ms = {
        "Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", 
        "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"
    }
    months_ms = {
        1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun",
        7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"
    }
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

def add_pkd_note(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note_text = "*Nota : GBK, Gombak; HL, Hulu Langat; HS, Hulu Selangor; KLG, Klang; KL, Kuala Langat; KS, Kuala Selangor; PTG, Petaling; SB, Sabak Bernam; SPG, Sepang."
    run = p.add_run(note_text)
    apply_font(run, 7, bold=False)
    p.paragraph_format.space_after = Pt(12)

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo
    logo_path = "logo.png.jpg" 
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(2.0))

    # 2. Tajuk Utama
    titles = [
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 10.5),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 10.5),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 10.5)
    ]
    for text, size in titles:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # 3. Jadual Tarikh Hijau
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 11, bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECTION 1.0 ---
    p1_head = doc.add_paragraph()
    p1_head.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    total_notifications = int(col_sums['Grand Total'])
    h11 = doc.add_paragraph()
    h11.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h11_text = f"1.1 Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(h11.add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")

    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    t1.width = content_width
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS','PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS','PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    h_cells = t1.rows[0].cells
    for i in range(len(h_cells)):
        h_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        h_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(h_cells[0].paragraphs[0].add_run("PENYAKIT"), 7.5, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 7, bold=True)
        set_cell_background(cell, "BFDFFF")
    apply_font(h_cells[-1].paragraphs[0].add_run("Jumlah"), 7, bold=True)
    set_cell_background(h_cells[-1], "FFFF00")

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        row[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        apply_font(row[0].paragraphs[0].add_run(str(penyakit)), 7, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            cell = row[c_idx+1]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            apply_font(cell.paragraphs[0].add_run(str(int(val))), 8, bold=True)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if (c_idx + 1) == (len(row_data)): set_cell_background(cell, "FFFFB3")

    f_cells = t1.rows[-1].cells
    for i in range(len(f_cells)): f_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 7.5, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, val in enumerate(col_sums):
        cell = f_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(str(int(val))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # NOTA DI BAWAH JADUAL 1
    add_pkd_note(doc)

    # --- SECTION 2.0 ---
    doc.add_page_break()
    p2_head = doc.add_paragraph()
    p2_head.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    
    harian_total = int(wabak_df['HARIAN'].sum())
    h21 = doc.add_paragraph()
    h21.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h21_text = f"2.1 Jadual di bawah menunjukkan jumlah wabak harian dan kumulatif di negeri Selangor. Sejumlah {harian_total} input notifikasi wabak diterima pada {get_malay_date(yesterday)}."
    apply_font(h21.add_run(h21_text), 11, bold=False)

    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")

    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    t2.width = content_width
    for i, h in enumerate(["PENYAKIT", "HARIAN", "KUMULATIF"]):
        cell = t2.cell(0, i)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        apply_font(cell.paragraphs[0].add_run(h), 9, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        for c in range(3): cells[c].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        apply_font(cells[0].paragraphs[0].add_run(str(penyakit)), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        apply_font(cells[1].paragraphs[0].add_run(str(int(row_data['HARIAN']))), 8, bold=True)
        apply_font(cells[2].paragraphs[0].add_run(str(int(row_data['KUMULATIF']))), 8, bold=True)
        cells[1].paragraphs[0].alignment = cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f2_cells = t2.rows[-1].cells
    for c in range(3): f2_cells[c].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    apply_font(f2_cells[0].paragraphs[0].add_run("JUMLAH"), 9, bold=True)
    apply_font(f2_cells[1].paragraphs[0].add_run(str(int(wabak_df['HARIAN'].sum()))), 9, bold=True)
    apply_font(f2_cells[2].paragraphs[0].add_run(str(int(wabak_df['KUMULATIF'].sum()))), 9, bold=True)
    for c in range(3): 
        set_cell_background(f2_cells[c], "FFFF00")
        f2_cells[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- SECTION 3.0 ---
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(24) 

    p3_head = doc.add_paragraph()
    p3_head.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    try: xx_v = int(float(vector_df.iloc[-1, 1]) + float(vector_df.iloc[-1, 3]) + float(vector_df.iloc[-1, 5]))
    except: xx_v = 0

    h31 = doc.add_paragraph()
    h31.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    h31_text = f"3.1 Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. Sejumlah {xx_v} input notifikasi wabak vektor telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 3."
    apply_font(h31.add_run(h31_text), 11, bold=False)

    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")

    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.width = content_width
    col_widths_v = [Inches(1.8), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.65)]

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
        h3_r2[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h3_r2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 7.5, bold=True)

    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            try: display_val = str(int(float(val))) if j > 0 else str(val)
            except: display_val = str(val)
            row_cells[j].width = col_widths_v[j]
            row_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = row_cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(display_val)
            if i == len(vector_df)-1: set_cell_background(row_cells[j], "FFFF00")
            elif j == 0: set_cell_background(row_cells[j], "FCE4D6")
            apply_font(run, 7.5, bold=True)

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    p4_head = doc.add_paragraph()
    p4_head.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p4_head.add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    h41 = doc.add_paragraph()
    h41.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if is_bkk_empty:
        h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. Tiada insiden dilaporkan pada {get_malay_date(yesterday)}."
    else:
        total_val_bkk = clean_val(bkk_table_df.iloc[-1]['JUMLAH'])
        h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. Sejumlah {total_val_bkk} input notifikasi Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK) telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut jenis insiden seperti dalam jadual 4."
    apply_font(h41.add_run(h41_text), 11, bold=False)

    add_table_title(doc, "Jadual 4", "Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)")

    t4 = doc.add_table(rows=bkk_table_df.shape[0] + 1, cols=bkk_table_df.shape[1])
    t4.style = 'Table Grid'
    t4.width = content_width

    for i, col_name in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER 
        if i < bkk_table_df.shape[1]-2: set_cell_background(cell, "BFDFFF")
        elif i == bkk_table_df.shape[1]-2: set_cell_background(cell, "FFFF00")
        else: set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_size = 6.5 if 0 < i < bkk_table_df.shape[1]-2 else 7
        run = p.add_run(str(col_name).replace(" ", "\n"))
        apply_font(run, f_size, bold=True)

    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last_row = (r_idx == bkk_table_df.shape[0] - 1)
        for c_idx, val in enumerate(row_data):
            cells[c_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER 
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_val(val))
            apply_font(run, 7.5 if c_idx == 0 else 8, bold=is_last_row or c_idx == 0)
            if is_last_row: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")
            elif c_idx == bkk_table_df.shape[1]-2: set_cell_background(cells[c_idx], "FFFFB3")
            elif c_idx == bkk_table_df.shape[1]-1: set_cell_background(cells[c_idx], "E2EFDA")

    # NOTA DI BAWAH JADUAL 4
    add_pkd_note(doc)

    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    # --- TANDATANGAN PLAIN TEXT ---
    doc.add_paragraph()
    p_petugas = doc.add_paragraph()
    p_petugas.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p_petugas.add_run("Petugas   :"), 11, bold=False)
    
    p_jawatan1 = doc.add_paragraph()
    p_jawatan1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p_jawatan1.add_run("Jawatan  :"), 11, bold=False)
    p_jawatan1.paragraph_format.space_after = Pt(36) 
    
    p_ketua = doc.add_paragraph()
    p_ketua.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p_ketua.add_run("Ketua Petugas :"), 11, bold=False)
    
    p_jawatan2 = doc.add_paragraph()
    p_jawatan2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p_jawatan2.add_run("Jawatan  :"), 11, bold=False)

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
            yesterday = date.today() - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y") 

            # S1
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix.sum(axis=0)

            # S2
            df2 = pd.read_excel(f2)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df2 = df2[df2['Tarikh Isytihar Wabak'] >= date(2026, 1, 4)]
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
            with st.spinner('Menarik data vektor...'):
                raw_gs = pd.read_csv(GSHEET_URL, header=None)
                mask_v = raw_gs.apply(lambda r: r.astype(str).str.contains('Petaling').any(), axis=1)
                if mask_v.any():
                    start_row = mask_v.idxmax()
                    v_data = raw_gs.iloc[start_row : start_row + 10, 13:20]
                    v_data = v_data[v_data[13].notna() & (v_data[13] != '')]
                else:
                    st.error("Data 'PETALING' tidak dijumpai.")
                    st.stop()

            # S4
            with st.spinner('Menarik data BKK...'):
                df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
                tkh_lapor_col = df_bkk_full.iloc[:, 2].astype(str)
                is_bkk_empty = not tkh_lapor_col.str.contains(yesterday_str).any()
                bkk_raw = df_bkk_full.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
                bkk_raw.columns = bkk_raw.iloc[0]
                bkk_table_final = bkk_raw[1:].reset_index(drop=True)
                bkk_map = {'GOMBAK':'GBK','HULU LANGAT':'HL','HULU SELANGOR':'HS','KLANG':'KLG','KUALA LANGAT':'KL','KUALA SELANGOR':'KS','PETALING':'PTG','SABAK BERNAM':'SB','SEPANG':'SPG','PK P.KLANG':'PK.KLG','PK KLIA':'PK.KLIA'}
                bkk_table_final = bkk_table_final.rename(columns=bkk_map)

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, is_bkk_empty)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")

        except Exception as e:
            st.error(f"Ralat: {e}")
