import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta
import pytz
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
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
    "Denggi": 427, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 8, "HIV/AIDS": 7, "Leptosopsirosis": 6,
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
GRAF_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet=GRAF%20TREND%20KES%20MINGGUAN"
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
    if any(x in name_str for x in ["HIV", "AIDS", "HFMD", "COVID-19"]): return name_str
    if "FOOD POISONING" in name_str: return "Keracunan Makanan"
    if name_str in ["DENGUE/DHF", "DENGUE"]: return "Denggi"
    return name_str.title()

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "-": return "-"
    return re.sub(r'\s*\(.*?\)', '', str(val)).strip()

def get_malay_date(target_date):
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    months_ms = {1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun", 7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"}
    return f"{target_date.day:02d} {months_ms[target_date.month]} {target_date.year} ({days_ms[target_date.strftime('%A')]})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

def add_table_title(doc, label, title, is_figure=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_figure else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f"{label} : ")
    apply_font(run, 11, bold=True)
    run2 = p.add_run(title)
    apply_font(run2, 11, bold=False)

# --- GRAF GENERATOR ---
def create_trend_chart(df_graf):
    try:
        weeks = df_graf.iloc[0, 1:54].values
        d_2025 = pd.to_numeric(df_graf.iloc[4, 1:54], errors='coerce').fillna(0)
        d_2026 = pd.to_numeric(df_graf.iloc[5, 1:54], errors='coerce')
        d_median = pd.to_numeric(df_graf.iloc[6, 1:54], errors='coerce').fillna(0)

        plt.figure(figsize=(10, 5))
        plt.plot(weeks, d_2025, color='#4285F4', label='2025', linewidth=2)
        plt.plot(weeks, d_2026, color='#EA4335', label='2026', linewidth=2)
        plt.plot(weeks, d_median, color='#FBBC05', label='Moving median 4 tahun (2022,2023,2024,2025)', linewidth=2)

        plt.xticks(weeks[::2], fontsize=8) 
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().spines['right'].set_visible(False)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=300)
        plt.close()
        buf.seek(0)
        return buf
    except: return None

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list, df_graf):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)

    p_logo = doc.add_paragraph()
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists("logo.png.jpg"): p_logo.add_run().add_picture("logo.png.jpg", width=Inches(1.8))
    
    for t in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(t), 10.5, bold=True)
        para.paragraph_format.space_after = Pt(0)

    doc.add_paragraph(f"\nTarikh: {get_malay_date(today)}")

    # --- SECTION 3 (VEKTOR) ---
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.space_before = Pt(24)
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.allow_autofit = False
    col_w = [Inches(1.8), Inches(0.75), Inches(0.75), Inches(0.75), Inches(0.75), Inches(0.75), Inches(0.75)]
    
    for r_idx in range(len(vector_df)):
        row = t3.rows[r_idx+2].cells
        for c_idx in range(7):
            row[c_idx].width = col_w[c_idx]
            p = row[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(vector_df.iloc[r_idx, c_idx])), 9, bold=True)

    # --- RAJAH 1 (GRAF) ---
    img_buf = create_trend_chart(df_graf)
    if img_buf:
        doc.add_paragraph().paragraph_format.space_before = Pt(18)
        add_table_title(doc, "Rajah 1", "Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor", is_figure=True)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(img_buf, width=Inches(5.8))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
# PERUBAHAN 1: Buang "+ Graf Trend" dari tajuk
st.title("📄 BWKK Report Generator")

# PERUBAHAN 2: Benarkan muat naik fail .xls dan .xlsx
f1 = st.file_uploader("Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # Load fail (pandas menyokong .xls jika xlrd dipasang)
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
            matrix = matrix.sort_values(by='Grand Total', ascending=False)
            col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

            df2 = pd.read_excel(f2, sheet_name="SELANGOR 2")
            today = get_msia_time().date()
            yesterday = today - timedelta(days=1)
            df_yesterday_list = df2[pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date == yesterday].values.tolist()
            
            wabak_df = pd.DataFrame({'HARIAN':[0], 'AKTIF':[0], 'KUMULATIF':[0]}, index=['Denggi'])
            v_data = pd.read_csv(GSHEET_URL, header=None).iloc[10:20, 13:20]
            df_graf = pd.read_csv(GRAF_URL, header=None)

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, pd.DataFrame(), True, [], df_yesterday_list, df_graf)
            
            st.success("✅ Laporan Berjaya Dijana!")
            st.download_button("⬇️ Muat Turun Laporan", doc_out, f"Laporan_BWKK_{today}.docx")
            
        except Exception as e:
            st.error(f"Ralat: {e}")
