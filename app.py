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
    "Denggi": 427, "Covid-19": 54, "Hfmd": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "Hiv/Aids": 7, "Leptosopsirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
# GID untuk "GRAF TREND KES MINGGUAN"
GID_GRAF = "757820121" 
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
GSHEET_GRAF_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_GRAF}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

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
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-" or str(val).lower() == "nan":
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

# --- CHART GENERATOR ---
def fetch_and_generate_chart():
    try:
        # Load data from Google Sheet
        df_graf = pd.read_csv(GSHEET_GRAF_URL, header=None)
        
        # Ranges based on user input:
        # x-axis b2:bb2 -> row 1 (0-indexed), cols 1 to 53
        x_axis = df_graf.iloc[1, 1:53].values
        
        # 2025 line b6:bb6 -> row 5
        y_2025 = pd.to_numeric(df_graf.iloc[5, 1:53], errors='coerce').values
        
        # 2026 line b7:bb7 -> row 6
        y_2026 = pd.to_numeric(df_graf.iloc[6, 1:53], errors='coerce').values
        
        # Median b8:bb8 -> row 7
        y_median = pd.to_numeric(df_graf.iloc[7, 1:53], errors='coerce').values

        # Plotting
        plt.figure(figsize=(10, 5))
        
        # Plot 2025 (Blue)
        plt.plot(x_axis, y_2025, label='2025', color='#4285F4', linewidth=2.5)
        
        # Plot 2026 (Red)
        plt.plot(x_axis, y_2026, label='2026', color='#EA4335', linewidth=2.5)
        
        # Plot Median (Yellow/Orange)
        plt.plot(x_axis, y_median, label='Moving median 4 tahun (2022,2023,2024,2025)', color='#FBBC05', linewidth=2.5)

        # Formatting
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False, fontsize=9)
        plt.xticks(rotation=90, fontsize=8)
        plt.yticks(range(0, 1500, 250))
        plt.grid(False)
        
        # Remove top and right spines
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        
        plt.tight_layout()

        # Save to buffer
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', dpi=300)
        plt.close()
        img_stream.seek(0)
        return img_stream
    except Exception as e:
        st.error(f"Gagal menjana graf: {e}")
        return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list, chart_img):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
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
        if i == 0:
            txt = f"\nTarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)"
        else:
            txt = f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(txt)
        apply_font(run, 11, bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- SECTION 1.0 (eNotifikasi) ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    
    total_notifications = int(col_sums['Grand Total'])
    h11 = doc.add_paragraph()
    h11.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
    h11_text = f"1.1 Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {total_notifications} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(h11.add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    
    # ... (Isian Jadual 1 dikekalkan sama seperti kod asal anda)
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    h_cells = t1.rows[0].cells
    for i in range(len(h_cells)):
        h_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        h_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(h_cells[0].paragraphs[0].add_run("Penyakit"), 8, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
    apply_font(h_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(h_cells[len(TEMPLATE_PKDS)+1], "FFFF00")
    apply_font(h_cells[len(TEMPLATE_PKDS)+2].paragraphs[0].add_run("Average Harian"), 8, bold=True)
    set_cell_background(h_cells[len(TEMPLATE_PKDS)+2], "FFC000")

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        apply_font(row[0].paragraphs[0].add_run(format_penyakit_name(penyakit)), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            cell = row[c_idx+1]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(row_data[pkd]))), 8, bold=True)
        apply_font(row[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(row[len(TEMPLATE_PKDS)+1], "FFFFB3")
        apply_font(row[len(TEMPLATE_PKDS)+2].paragraphs[0].add_run(str(int(row_data.get('Average Harian', 0)))), 8, bold=True)
        set_cell_background(row[len(TEMPLATE_PKDS)+2], "FFC000")

    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = f_cells[i+1]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(cell.paragraphs[0].add_run(str(int(col_sums[pkd]))), 8, bold=True)
        set_cell_background(cell, "FFFF00")
    apply_font(f_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run(str(int(col_sums['Grand Total']))), 8, bold=True)
    set_cell_background(f_cells[len(TEMPLATE_PKDS)+1], "FFFF00")
    set_cell_background(f_cells[len(TEMPLATE_PKDS)+2], "FFC000")

    doc.add_paragraph()
    add_pkd_note(doc)

    # --- SECTION 2.0 (WABAK) ---
    doc.add_page_break()
    p2_head = doc.add_paragraph()
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    harian_total = int(wabak_df['HARIAN'].sum())
    h21 = doc.add_paragraph()
    h21_text = f"2.1 Jadual di bawah menunjukkan jumlah wabak harian, aktif dan kumulatif di negeri Selangor. Sejumlah {harian_total} input notifikasi wabak diterima pada {get_malay_date(yesterday)}."
    apply_font(h21.add_run(h21_text), 11, bold=False)

    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    # ... (Isian Jadual 2 dikekalkan)
    for i, h in enumerate(["Penyakit", "Harian", "Aktif", "Kumulatif"]):
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
    footer_vals = ["Jumlah", str(int(wabak_df['HARIAN'].sum())), str(int(wabak_df['AKTIF'].sum())), str(int(wabak_df['KUMULATIF'].sum()))]
    for i, txt in enumerate(footer_vals):
        apply_font(f2_cells[i].paragraphs[0].add_run(txt), 8, bold=True)
        set_cell_background(f2_cells[i], "FFFF00")
        f2_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Jadual 2.1
    doc.add_paragraph()
    add_table_title(doc, "Jadual 2.1", f"Senarai Wabak Yang Dilaporkan pada {tarikh_semalam_str}")
    t21 = doc.add_table(rows=1, cols=5)
    t21.style = 'Table Grid'
    # ... (Logik Jadual 2.1 dikekalkan)
    h21_headers = ["Bil", "Wabak", "Daerah", "Tempat Berlaku", "Bil Kes (AR)"]
    for i, txt in enumerate(h21_headers):
        cell = t21.cell(0, i)
        set_cell_background(cell, "BFDFFF")
        apply_font(cell.paragraphs[0].add_run(txt), 10, bold=True)

    if not df_yesterday_list:
        row = t21.add_row().cells
        row[0].merge(row[4])
        row[0].text = "Tiada wabak dilaporkan pada tarikh ini."
    else:
        for idx, item in enumerate(df_yesterday_list, start=1):
            row = t21.add_row().cells
            row[0].text = str(idx)
            row[1].text = f"{item[0]}\n({'Household' if str(item[3]).strip() == 'Rumah Persendirian' else 'Institusi'})"
            row[2].text = str(item[1]).title()
            row[3].text = str(item[2])
            n_kes = int(item[4]) if pd.notna(item[4]) else 0
            n_dedah = int(item[5]) if pd.notna(item[5]) else 0
            row[4].text = f"{n_kes}/{n_dedah}"

    # --- SECTION 3.0 (VEKTOR) ---
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.page_break_before = True 
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    xx_v = int(vector_df.iloc[-1, 1]) + int(vector_df.iloc[-1, 3]) + int(vector_df.iloc[-1, 5])
    h31 = doc.add_paragraph()
    h31_text = f"3.1 Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. {xx_v} input notifikasi wabak vektor telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 3."
    apply_font(h31.add_run(h31_text), 11, bold=False)

    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    # ... (Isian Jadual 3 dikekalkan)
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            apply_font(row_cells[j].paragraphs[0].add_run(str(val)), 9, bold=True)

    # --- NEW: RAJAH 1 (GRAF DENGGI) ---
    if chart_img:
        doc.add_paragraph()
        add_table_title(doc, "Rajah 1", "Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor")
        p_chart = doc.add_paragraph()
        p_chart.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_chart = p_chart.add_run()
        run_chart.add_picture(chart_img, width=Inches(6.2))
        doc.add_paragraph()

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    p4_head = doc.add_paragraph()
    apply_font(p4_head.add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    h41 = doc.add_paragraph()
    if is_bkk_empty:
        h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. Tiada insiden dilaporkan pada {get_malay_date(yesterday)}."
    else:
        h41_text = f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. Terdapat {len(bkk_details)} kejadian dilaporkan pada {get_malay_date(yesterday)}."
    apply_font(h41.add_run(h41_text), 11, bold=False)
    
    add_table_title(doc, "Jadual 4", "Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        for c_idx, val in enumerate(row_data):
            apply_font(cells[c_idx].paragraphs[0].add_run(clean_val(val)), 8, bold=False)

    doc.add_paragraph()
    add_pkd_note(doc)
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    # --- TANDATANGAN ---
    sig_table = doc.add_table(rows=8, cols=3)
    # ... (Sama seperti kod asal anda untuk border none)
    fill_sig_row = lambda r, l: (apply_font(sig_table.rows[r].cells[0].paragraphs[0].add_run(l), 11, False), apply_font(sig_table.rows[r].cells[1].paragraphs[0].add_run(":"), 11, False))
    fill_sig_row(0, "Disediakan"); fill_sig_row(1, "Jawatan")
    fill_sig_row(3, "Disemak"); fill_sig_row(4, "Jawatan")
    fill_sig_row(6, "Disahkan"); fill_sig_row(7, "Jawatan")

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📄 BWKK Report Generator")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # ... (Logik pemprosesan data asal dikekalkan)
            # ...
            
            # 1. Jana Chart Denggi
            chart_img = fetch_and_generate_chart()
            
            # 2. Jana Docx dengan chart
            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, (len(bkk_details)==0), bkk_details, df_yesterday_list, chart_img)
            
            st.success("✅ Laporan berjaya dijana!")
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
