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
    try:
        df_trend = pd.read_csv(GRAF_DATA_URL, header=None)
        labels = df_trend.iloc[:, 0].astype(str).str.strip()
        try:
            idx_2025 = labels[labels == "2025"].index[0]
            idx_2026 = labels[labels == "2026"].index[0]
            idx_median = labels[labels.str.contains("median", case=False, na=False)].index[0]
        except IndexError:
            idx_2025, idx_2026, idx_median = 5, 6, 7 

        weeks = range(1, 54)
        y2025 = pd.to_numeric(df_trend.iloc[idx_2025, 1:54], errors='coerce')
        y2026 = pd.to_numeric(df_trend.iloc[idx_2026, 1:54], errors='coerce')
        ymedian = pd.to_numeric(df_trend.iloc[idx_median, 1:54], errors='coerce')

        plt.figure(figsize=(11, 5))
        plt.plot(weeks, y2025, label='2025', color='#4472C4', linewidth=1.5)
        plt.plot(weeks, y2026, label='2026', color='#C00000', linewidth=2.5)
        plt.plot(weeks, ymedian, label='Moving median 4 tahun (2022,2023,2024,2025)', color='#FFC000', linewidth=1.5)
        plt.title('CARTA KES MINGGUAN DENGGI YANG DIDAFTAR BAGI TAHUN 2025-2026\nNEGERI SELANGOR', fontsize=11, fontweight='bold', pad=15)
        plt.xticks(weeks, fontsize=7)
        plt.grid(axis='y', linestyle='-', alpha=0.2)
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
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]): return name_str
    if "FOOD POISONING" in name_str: return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]: return "Denggi"
    return name_str.title()

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def clean_val(val):
    if pd.isna(val) or str(val).strip() in ["", "-", "nan"]: return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    return f"{((target_date - start_date).days // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    return f"{target_date.day:02d} {months_ms[target_date.month]} {target_date.year} ({days_ms[target_date.strftime('%A')]})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def add_table_title(doc, label, title):
    p = doc.add_paragraph()
    run_label = p.add_run(f"{label} : ")
    apply_font(run_label, 11, bold=True)
    run_title = p.add_run(title)
    apply_font(run_title, 11, bold=False)

def add_pkd_note(doc):
    p = doc.add_paragraph()
    run = p.add_run("*Nota : GBK, Gombak; HL, Hulu Langat; HS, Hulu Selangor; KLG, Klang; KL, Kuala Langat; KS, Kuala Selangor; PTG, Petaling; SB, Sabak Bernam; SPG, Sepang.")
    apply_font(run, 7, bold=False)

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list, trend_graph):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # Logo & Titles
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture(logo_path, width=Inches(1.8))

    for text in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), 10.5, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # Green Header Table
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width 
    for i, txt in enumerate([f"\nTarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)", f"\nMinggu Epidemiologi : {get_epi_week(today)}"]):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 11, bold=True)

    doc.add_paragraph()

    # --- SECTION 1.0 (Summary) ---
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    h11_text = f"1.1 Jadual di bawah menunjukkan jumlah input enotifikasi di negeri Selangor. Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    apply_font(doc.add_paragraph().add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    
    # ... Table 1 logic continues ...
    # (Simplified for brevity, keeping your original coloring/formatting logic)
    
    # --- SECTION 3.0 (VEKTOR) ---
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.page_break_before = True 
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    # ... Jadual 3 Logic ...

    # --- RAJAH 1: GRAPH INSERTION ---
    if trend_graph:
        doc.add_paragraph()
        p_fig = doc.add_paragraph()
        p_fig.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_fig = p_fig.add_run("Rajah 1 : Carta Mingguan Denggi Didaftar Bagi Tahun 2025-2026 Negeri Selangor")
        apply_font(run_fig, 11, bold=True)
        doc.add_picture(trend_graph, width=Inches(6.2))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    # ... BKK Logic ...

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

            # Load Files
            df1 = pd.read_excel(f1)
            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")

            # FIX: DYNAMIC COLUMN SEARCH for "Tempat Berlaku" and "Kategori"
            # This searches for columns that CONTAIN specific keywords to avoid "not in index" errors
            addr_col = [c for c in df2.columns if "Tempat Berlaku Wabak" in str(c) or "Alamat" in str(c)][0]
            cat_col = [c for c in df2.columns if "Kategori Tempat" in str(c)][0]

            # Processing df2 (Wabak)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df_yesterday = df2[df2['Tarikh Isytihar Wabak'] == yesterday].copy()
            df_yesterday_list = df_yesterday[['PENYAKIT', 'DAERAH (HURUF BESAR)', addr_col, cat_col, 'Bilangan Kes', 'Bilangan Terdedah']].values.tolist()

            # ... Rest of Matrix/Wabak Logic ...
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)
            
            # GSheets Data
            v_data = pd.read_csv(GSHEET_URL, header=None) # Simplified vector fetch
            bkk_table_final = pd.read_csv(SHEET_BKK_URL) # Simplified BKK fetch
            
            # Generate Graph
            st.info("Sedang memproses graf trend...")
            trend_img = capture_dengue_trend_graph()

            # Final Doc
            doc_out = generate_docx(matrix, col_totals, pd.DataFrame(), v_data, bkk_table_final, True, [], df_yesterday_list, trend_img)
            st.success("✅ Laporan berjaya dijana!")
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{today}.docx")
            
        except Exception as e:
            st.error(f"Ralat Utama: {e}")
