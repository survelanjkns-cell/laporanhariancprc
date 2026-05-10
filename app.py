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
GID = "0"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"
GRAF_DATA_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet=GRAF%20TREND%20KES%20MINGGUAN"

# --- HELPERS ---
def capture_dengue_trend_graph():
    """Fetches data from grid (B:BB) and recreates the floating graph."""
    try:
        df_trend = pd.read_csv(GRAF_DATA_URL, header=None)
        labels = df_trend.iloc[:, 0].astype(str).str.strip()
        
        # Search for row indices dynamically
        try:
            idx_2025 = labels[labels == "2025"].index[0]
            idx_2026 = labels[labels == "2026"].index[0]
            idx_median = labels[labels.str.contains("median", case=False, na=False)].index[0]
        except IndexError:
            # Fallback to screenshot positions if labels moved/unreadable
            idx_2025, idx_2026, idx_median = 5, 6, 7 

        weeks = range(1, 54)
        # Data range B:BB corresponds to columns 1 to 53
        y2025 = pd.to_numeric(df_trend.iloc[idx_2025, 1:54], errors='coerce')
        y2026 = pd.to_numeric(df_trend.iloc[idx_2026, 1:54], errors='coerce')
        ymedian = pd.to_numeric(df_trend.iloc[idx_median, 1:54], errors='coerce')

        plt.figure(figsize=(11, 5))
        plt.plot(weeks, y2025, label='2025', color='#4472C4', linewidth=1.5)
        plt.plot(weeks, y2026, label='2026', color='#C00000', linewidth=2.5)
        plt.plot(weeks, ymedian, label='Moving median 4 tahun (2022,2023,2024,2025)', color='#FFC000', linewidth=1.5)

        plt.title('CARTA KES MINGGUAN DENGGI YANG DIDAFTAR BAGI TAHUN 2025-2026\nNEGERI SELANGOR', 
                  fontsize=11, fontweight='bold', pad=15)
        plt.xticks(weeks, fontsize=7)
        plt.yticks(range(0, 1501, 250), fontsize=8)
        plt.grid(axis='y', linestyle='-', alpha=0.3)
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, fontsize=8, frameon=False)
        
        plt.tight_layout()
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', dpi=300)
        img_stream.seek(0)
        plt.close()
        return img_stream
    except Exception as e:
        st.warning(f"Gagal menjana graf: {e}")
        return None

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

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list, dengue_graph):
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

    # Section 1.0, 2.0 remain standard... (Truncated for space, keeping your original structure)
    # [Insert existing Section 1 and 2 logic here]

    # --- SECTION 3.0 (VEKTOR) ---
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.page_break_before = True 
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    # ... [Insert existing Jadual 3 logic here] ...
    # (Assuming Jadual 3 is already rendered in your code)

    # --- RAJAH 1 (GRAPH) ---
    if dengue_graph:
        doc.add_paragraph()
        p_fig = doc.add_paragraph()
        p_fig.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_fig = p_fig.add_run("Rajah 1 : Carta Mingguan Denggi Didaftar Bagi Tahun 2025-2026 Negeri Selangor")
        apply_font(run_fig, 11, bold=True)
        
        doc.add_picture(dengue_graph, width=Inches(6.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

    # --- SECTION 4.0 (BKK) AND SIGNATURES ---
    # ... [Insert existing Section 4 and Signature logic here] ...

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
            now_msia = get_msia_time()
            today = now_msia.date()
            yesterday = today - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y")

            # Data Processing Logic...
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            # ... [Other data processing for Wabak, Vektor, BKK] ...
            # Loading Vector data from GSHEET to pass to generator
            raw_gs = pd.read_csv(GSHEET_URL, header=None)
            mask_v = raw_gs.apply(lambda r: r.astype(str).str.contains('Petaling').any(), axis=1)
            v_data = raw_gs.iloc[mask_v.idxmax() : mask_v.idxmax() + 11, 13:20].dropna(how='all')

            # Process BKK data
            df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
            insiden_semalam = df_bkk_full[df_bkk_full.iloc[:, 2].astype(str).str.contains(yesterday_str)]
            bkk_details = [{'kejadian': r[5], 'alamat': r[8], 'daerah': r[4]} for _, r in insiden_semalam.iterrows()]
            bkk_raw = df_bkk_full.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_raw.columns = bkk_raw.iloc[0]
            bkk_table_final = bkk_raw[1:].reset_index(drop=True).rename(columns={'GOMBAK':'GBK','HULU LANGAT':'HL','HULU SELANGOR':'HS','KLANG':'KLG','KUALA LANGAT':'KL','KUALA SELANGOR':'KS','PETALING':'PTG','SABAK BERNAM':'SB','SEPANG':'SPG'})

            # NEW: Generate Trend Graph
            st.info("Sedang memproses graf trend...")
            trend_graph = capture_dengue_trend_graph()

            # Final Doc Generation
            doc_out = generate_docx(matrix, col_totals, pd.DataFrame(), v_data, bkk_table_final, (len(bkk_details)==0), bkk_details, [], trend_graph)
            
            st.success("✅ Laporan berjaya dijana!")
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
        except Exception as e:
            st.error(f"Ralat Utama: {e}")
