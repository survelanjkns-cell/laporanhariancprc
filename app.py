import streamlit as st
import pandas as pd
import plotly.graph_objects as go
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
# URLs
SHEET_GRAPH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTDprYai1uaP1L-JP6kuHRZX18AmDHX0ROEzRE37DaCHMo0cNWUvRa8R-65RZAK7XFWI6pb_-X-jF24/pub?gid=1525373641&single=true&output=csv"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

# --- HELPERS ---
def to_float(val):
    """Clean string numbers with commas (e.g. '1,222') and convert to float."""
    try:
        if pd.isna(val) or str(val).strip() == "": return 0.0
        clean_str = str(val).replace(',', '').strip()
        return float(clean_str)
    except:
        return 0.0

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
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    return f"{target_date.day:02d} {months_ms[target_date.month]} {target_date.year} ({days_ms[target_date.strftime('%A')]})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def add_table_title(doc, label, title):
    p = doc.add_paragraph()
    apply_font(p.add_run(f"{label} : "), 11, True)
    apply_font(p.add_run(title), 11, False)

# --- GRAPH IMAGE GENERATOR ---
def get_graph_image():
    try:
        df_graph = pd.read_csv(SHEET_GRAPH_URL, skiprows=1)
        df_graph = df_graph.set_index(df_graph.columns[0])
        df_plot = df_graph.transpose()
        
        fig = go.Figure()
        if '2025' in df_plot.columns:
            fig.add_trace(go.Scatter(x=df_plot.index, y=df_plot['2025'], mode='lines', name='2025', line=dict(color='#4285F4', width=2)))
        if '2026' in df_plot.columns:
            fig.add_trace(go.Scatter(x=df_plot.index, y=df_plot['2026'], mode='lines', name='2026', line=dict(color='#EA4335', width=3)))
        
        median_col = [col for col in df_plot.columns if 'Moving median' in str(col)]
        if median_col:
            fig.add_trace(go.Scatter(x=df_plot.index, y=df_plot[median_col[0]], mode='lines', name='Median (4 thn)', line=dict(color='#FBBC04', width=2)))

        fig.update_layout(plot_bgcolor="white", legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"), margin=dict(l=20, r=20, t=20, b=20))
        img_bytes = fig.to_image(format="png", width=1000, height=500, scale=2)
        return io.BytesIO(img_bytes)
    except:
        return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_final, is_bkk_empty, bkk_details, df_yesterday_list):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # Header section
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        doc.add_paragraph().add_run().add_picture(logo_path, width=Inches(1.8))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    titles = ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 
              "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 
              "JABATAN KESIHATAN NEGERI SELANGOR"]
    for t in titles:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(t), 10.5, True)

    # Date Table
    info_table = doc.add_table(rows=1, cols=2)
    info_table.width = content_width
    for i in range(2):
        set_cell_background(info_table.cell(0, i), "C6E0B4")
        p = info_table.cell(0, i).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"Tarikh : {get_malay_date(today)}" if i==0 else f"Minggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, True)

    # Section 1.0 (eNotifikasi)
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, True)
    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    # ... (Table 1 creation logic - simplified for brevity)

    # Section 2.0 (Wabak)
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, True)
    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    # ... (Table 2 creation logic)

    # Section 3.0 (Vektor)
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, True)
    
    # Numeric fix for comma strings
    denggi_h = to_float(vector_df.iloc[-1, 1])
    malaria_h = to_float(vector_df.iloc[-1, 3])
    chiku_h = to_float(vector_df.iloc[-1, 5])
    xx_v = int(denggi_h + malaria_h + chiku_h)
    
    h31 = doc.add_paragraph()
    apply_font(h31.add_run(f"3.1 Pecahan mengikut penyakit seperti dalam jadual 3. Jumlah harian: {xx_v}."), 11, False)

    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    # ... (Table 3 header logic)
    
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            if j == 0: txt = str(val).title()
            else:
                num = to_float(val)
                txt = f"{int(num):,}" if num > 0 else "-"
            p = row_cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j==0 else WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(txt), 9, True)

    # --- INTEGRATED GRAPH (RAJAH 1) ---
    doc.add_paragraph()
    graph_img = get_graph_image()
    if graph_img:
        p_graph = doc.add_paragraph()
        p_graph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_graph.add_run().add_picture(graph_img, width=Inches(6.2))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(cap.add_run("Rajah 1 : Carta Kes Mingguan Denggi Didaftarkan Bagi Tahun 2025-2026 Negeri Selangor"), 9, True)

    # Section 4.0 (BKK)
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana (BKK)"), 11, True)
    # ... (Table 4 logic)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator")
st.title("📑 BWKK Report Generator")

f1 = st.file_uploader("Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # Data Loading
            df1 = pd.read_excel(f1)
            df1 = df1[(df1['Notifikasi Status'] != 'Abai Notifikasi') & (df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS))]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            # ... (Existing filtering/processing logic for df2)

            raw_gs = pd.read_csv(GSHEET_URL, header=None)
            v_data = raw_gs.iloc[raw_gs.apply(lambda r: 'Petaling' in str(r.values), axis=1).idxmax():, 13:20].dropna(how='all')

            df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
            # ... (Existing BKK logic)

            doc_out = generate_docx(matrix, col_totals, pd.DataFrame(), v_data, pd.DataFrame(), True, [], [])
            st.success("✅ Berjaya!")
            st.download_button("⬇️ Muat Turun", data=doc_out, file_name=f"BWKK_Report.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
