import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import pytz
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os
import re
import matplotlib.pyplot as plt

# --- KONSTAN & MAPPING DATA ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING',
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

AVG_HARIAN_FIGURES = {
    "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "HIV/AIDS": 7, "Leptosopsirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

# --- CONFIG GOOGLE SHEETS ---
SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID_VEKTOR = "0"
GID_GRAF = "68285521" # Sila sahkan nombor ini di URL tab graf anda

# URL Eksport yang lebih stabil
URL_VEKTOR = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_VEKTOR}"
URL_GRAF = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_GRAF}"
URL_BKK = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
def set_repeat_table_header(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w')))
    trPr.append(tblHeader)

def get_msia_time():
    msia_tz = pytz.timezone('Asia/Kuala_Lumpur')
    return datetime.now(msia_tz)

def format_penyakit_name(name):
    name_str = str(name).strip().upper()
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]):
        return name_str
    if "FOOD POISONING" in name_str:
        return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]:
        return "Denggi"
    return name_str.title()

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
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
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

# --- FUNGSI GENERATE GRAF ---
def create_dengue_chart():
    try:
        # Gunakan URL yang sudah dibersihkan
        df_chart = pd.read_csv(URL_GRAF.strip(), header=None)
        
        # Ekstrak data berdasarkan koordinat:
        # Baris 2 (idx 1), Baris 6 (idx 5), Baris 7 (idx 6), Baris 8 (idx 7)
        weeks = df_chart.iloc[1, 1:54].values
        data_2025 = pd.to_numeric(df_chart.iloc[5, 1:54], errors='coerce').fillna(0)
        data_2026 = pd.to_numeric(df_chart.iloc[6, 1:54], errors='coerce').fillna(0)
        data_median = pd.to_numeric(df_chart.iloc[7, 1:54], errors='coerce').fillna(0)

        plt.figure(figsize=(10, 5))
        
        plt.plot(weeks, data_2025, label='2025', color='#4285F4', linewidth=2.5)
        plt.plot(weeks, data_2026, label='2026', color='#EA4335', linewidth=2.5)
        plt.plot(weeks, data_median, label='Moving median 4 tahun (2022,2023,2024,2025)', color='#FBBC05', linewidth=2.5)

        plt.xticks(ticks=range(len(weeks)), labels=weeks, rotation=0, fontsize=8)
        plt.gca().xaxis.set_major_locator(plt.MultipleLocator(2)) 
        
        plt.yticks(range(0, 1501, 250), fontsize=9)
        plt.ylim(0, 1250)
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False, fontsize=9)
        
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=300)
        img_buf.seek(0)
        plt.close()
        return img_buf
    except Exception as e:
        # Ralat 400 akan dikesan di sini jika GID salah
        st.error(f"Gagal menjana graf: {e}")
        return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo (Pastikan fail wujud)
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.8))

    # 2. Tajuk
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

    # Info Tarikh
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"\nTarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 11, bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- 1.0 Enotifikasi ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    total_notifications = int(col_sums['Grand Total'])
    h11 = doc.add_paragraph()
    h11_text = f"1.1 Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(h11.add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    # Headers
    h_cells = t1.rows[0].cells
    apply_font(h_cells[0].paragraphs[0].add_run("PENYAKIT"), 8, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(h_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(h_cells[len(TEMPLATE_PKDS)+1], "FFFF00")
    apply_font(h_cells[len(TEMPLATE_PKDS)+2].paragraphs[0].add_run("Average Harian"), 8, bold=True)
    set_cell_background(h_cells[len(TEMPLATE_PKDS)+2], "FFC000")

    # Data
    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(format_penyakit_name(penyakit)), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            cell = row[c_idx+1]
            apply_font(cell.paragraphs[0].add_run(str(int(row_data[pkd]))), 8, bold=True)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(row[len(TEMPLATE_PKDS)+1], "FFFFB3")
        row[len(TEMPLATE_PKDS)+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(row[len(TEMPLATE_PKDS)+2].paragraphs[0].add_run(str(int(row_data.get('Average Harian', 0)))), 8, bold=True)
        set_cell_background(row[len(TEMPLATE_PKDS)+2], "FFC000")
        row[len(TEMPLATE_PKDS)+2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Footer
    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = f_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(str(int(col_sums[pkd]))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(f_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run(str(int(col_sums['Grand Total']))), 8, bold=True)
    set_cell_background(f_cells[len(TEMPLATE_PKDS)+1], "FFFF00")
    f_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_cell_background(f_cells[len(TEMPLATE_PKDS)+2], "FFC000")

    doc.add_paragraph()
    add_pkd_note(doc)

    # --- 2.0 Wabak ---
    doc.add_page_break()
    p2_head = doc.add_paragraph()
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    h21 = doc.add_paragraph()
    h21_text = f"2.1 Jadual di bawah menunjukkan jumlah wabak harian, aktif dan kumulatif di negeri Selangor. Sejumlah {int(wabak_df['HARIAN'].sum())} input notifikasi wabak diterima pada {get_malay_date(yesterday)}."
    apply_font(h21.add_run(h21_text), 11, bold=False)

    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    for i, h in enumerate(["PENYAKIT", "HARIAN", "AKTIF", "KUMULATIF"]):
        cell = t2.cell(0, i)
        apply_font(cell.paragraphs[0].add_run(h), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(penyakit)), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        for idx, col_key in enumerate(['HARIAN', 'AKTIF', 'KUMULATIF'], start=1):
            run = cells[idx].paragraphs[0].add_run(str(int(row_data[col_key])))
            apply_font(run, 8, bold=True)
            cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f2_cells = t2.rows[-1].cells
    for i, txt in enumerate(["JUMLAH", str(int(wabak_df['HARIAN'].sum())), str(int(wabak_df['AKTIF'].sum())), str(int(wabak_df['KUMULATIF'].sum()))]):
        apply_font(f2_cells[i].paragraphs[0].add_run(txt), 8, bold=True)
        set_cell_background(f2_cells[i], "FFFF00")
        f2_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Jadual 2.1
    doc.add_paragraph()
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan pada {get_malay_date(yesterday)}")
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    set_repeat_table_header(t21.rows[0])
    h21_headers = ["BIL", "WABAK", "DAERAH", "TEMPAT BERLAKU", "BIL KES (AR)"]
    for i, txt in enumerate(h21_headers):
        cell = t21.cell(0, i)
        set_cell_background(cell, "BFDFFF")
        apply_font(cell.paragraphs[0].add_run(txt), 10, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not df_yesterday_list:
        row = t21.add_row().cells
        row[0].merge(row[4])
        row[0].text = "Tiada wabak dilaporkan pada tarikh ini."
    else:
        for idx, item in enumerate(df_yesterday_list, start=1):
            row = t21.add_row().cells
            row[0].text = str(idx)
            row[1].text = f"{item[0]}\n{'(Household)' if str(item[3]).strip() == 'Rumah Persendirian' else '(Institusi)'}"
            row[2].text = str(item[1])
            row[3].text = str(item[2])
            n_kes, n_dedah = float(item[4]), float(item[5])
            pct = (n_kes/n_dedah*100) if n_dedah > 0 else 0
            row[4].text = f"{int(n_kes)}/{int(n_dedah)}\n({pct:.2f}%)"
            for c in range(5): 
                apply_font(row[c].paragraphs[0].runs[0], 8, bold=False)
                row[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if c != 3 else WD_ALIGN_PARAGRAPH.LEFT

    # --- 3.0 Vektor ---
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.space_before = Pt(24) 
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    try: h3_total = int(float(vector_df.iloc[-1, 1]) + float(vector_df.iloc[-1, 3]) + float(vector_df.iloc[-1, 5]))
    except: h3_total = 0
    h31 = doc.add_paragraph()
    apply_font(h31.add_run(f"3.1 Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. Sejumlah {h3_total} input notifikasi wabak vektor telah diterima pada {get_malay_date(yesterday)}."), 11, bold=False)

    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    h3_r1 = t3.rows[0].cells
    h3_r1[0].merge(t3.rows[1].cells[0]).text = "DAERAH"
    h3_r1[1].merge(h3_r1[2]).text = "DENGGI"
    h3_r1[3].merge(h3_r1[4]).text = "MALARIA"
    h3_r1[5].merge(h3_r1[6]).text = "CHIKUNGUNYA"
    for i in [0, 1, 3, 5]: 
        set_cell_background(h3_r1[i], "BFDFFF")
        apply_font(h3_r1[i].paragraphs[0].add_run(""), 10, bold=True) # Reset and format
        h3_r1[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    h3_r2 = t3.rows[1].cells
    for i in range(1, 7):
        h3_r2[i].text = "HARIAN" if i % 2 != 0 else "KUM"
        set_cell_background(h3_r2[i], "BFDFFF")
        apply_font(h3_r2[i].paragraphs[0].runs[0], 9, bold=True)
        h3_r2[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i in range(len(vector_df)):
        row = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            txt = str(int(float(val))) if j > 0 else str(val).upper()
            run = row[j].paragraphs[0].add_run(txt)
            apply_font(run, 9, bold=True)
            row[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            if i == len(vector_df)-1: set_cell_background(row[j], "FFFF00")
            elif j == 0: set_cell_background(row[j], "FCE4D6")

    # --- RAJAH 1 (CARTA) ---
    doc.add_paragraph()
    p_rajah = doc.add_paragraph()
    run_r = p_rajah.add_run("Rajah 1 : ")
    apply_font(run_r, 11, bold=True)
    run_rt = p_rajah.add_run("Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor")
    apply_font(run_rt, 11, bold=False)

    chart_buf = create_dengue_chart()
    if chart_buf:
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(chart_buf, width=Inches(6.2))
    doc.add_paragraph()

    # --- 4.0 BKK ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden BKK. {'Tiada' if is_bkk_empty else len(bkk_details)} insiden dilaporkan pada {get_malay_date(yesterday)}."
    apply_font(doc.add_paragraph().add_run(h41_text), 11, bold=False)
    
    add_table_title(doc, "Jadual 4", "Senarai Kejadian BKK")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    for i, col in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        set_cell_background(cell, "BFDFFF" if i < len(bkk_table_df.columns)-2 else "FFFF00")
        apply_font(cell.paragraphs[0].add_run(str(col)), 8, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        for c_idx, val in enumerate(row_data):
            run = cells[c_idx].paragraphs[0].add_run(clean_val(val))
            apply_font(run, 8, bold=(r_idx == len(bkk_table_df)-1))
            if r_idx == len(bkk_table_df)-1: set_cell_background(cells[c_idx], "FFFF00")

    doc.add_paragraph()
    add_pkd_note(doc)
    apply_font(doc.add_paragraph().add_run(f"*Sumber : Sistem e-notifikasi ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Notifikasi Harian (Excel)", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Linelisting Wabak (Excel)", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            today = get_msia_time().date()
            yesterday = today - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y")

            # Data S1
            df1 = pd.read_excel(f1)
            df1 = df1[(df1['Notifikasi Status'] != 'Abai Notifikasi') & (df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS))]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            # Data S2
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df_yesterday_list = df2[df2['Tarikh Isytihar Wabak'] == yesterday][['PENYAKIT', 'DAERAH (HURUF BESAR)', 'Tempat Berlaku Wabak\n(Alamat diisi lengkap dengan :- No rumah, nama jalan, nama tempat, daerah dan Negeri)', 'Kategori Tempat\n(Kategori premis berdasarkan tempat berlaku wabak)', 'Bilangan Kes', 'Bilangan Terdedah']].values.tolist()
            
            wb_sum = []
            df2_filt = df2[df2['Tarikh Isytihar Wabak'] >= date(2026, 1, 4)]
            for d in df2_filt['PENYAKIT'].unique():
                if pd.isna(d): continue
                ds_df = df2_filt[df2_filt['PENYAKIT'] == d]
                wb_sum.append({'PENYAKIT': d, 'HARIAN': len(ds_df[ds_df['Tarikh Isytihar Wabak'] == yesterday]), 'AKTIF': len(ds_df), 'KUMULATIF': len(ds_df)})
            wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values(by='KUMULATIF', ascending=False)

            # Data S3 & S4 (G-Sheets)
            v_data = pd.read_csv(URL_VEKTOR, header=None).iloc[18:28, 13:20] # Contoh slicing, sesuaikan jika perlu
            df_bkk = pd.read_csv(URL_BKK, header=None)
            insiden_smlm = df_bkk[df_bkk.iloc[:, 2].astype(str).str.contains(yesterday_str)]
            bkk_table = df_bkk.iloc[1:12, 33:47].fillna("-") # Contoh slicing BKK

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table, (len(insiden_smlm)==0), insiden_smlm, df_yesterday_list)
            st.success("✅ Laporan berjaya dijana!")
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
