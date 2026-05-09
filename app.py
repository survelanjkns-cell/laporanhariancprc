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
GID_DATA_NOTIFIKASI = "0"
# URL untuk Sheet GRAF TREND KES MINGGUAN (Gid perlu diselaraskan jika berbeza, di sini andaikan gid=527022201 berdasarkan susunan sheet umum)
GID_GRAF = "527022201" 

GSHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID_DATA_NOTIFIKASI}"
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

def add_table_title(doc, label, title, is_figure=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if not is_figure else WD_ALIGN_PARAGRAPH.CENTER
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

# --- GRAF GENERATOR ---
def create_trend_chart(df_graf):
    # Mengambil baris mengikut range yang dinyatakan
    # x_axis: b2:bb2 (index 0 dalam CSV tq), 2025: b6:bb6 (index 4), 2026: b7:bb7 (index 5), Median: b8:bb8 (index 6)
    weeks = df_graf.iloc[0, 1:53].values # b2:bb2
    data_2025 = pd.to_numeric(df_graf.iloc[4, 1:53], errors='coerce').values # b6:bb6
    data_2026 = pd.to_numeric(df_graf.iloc[5, 1:53], errors='coerce').values # b7:bb7
    data_median = pd.to_numeric(df_graf.iloc[6, 1:53], errors='coerce').values # b8:bb8

    plt.figure(figsize=(10, 5))
    plt.plot(weeks, data_2025, color='#4285F4', label='2025', linewidth=2)
    plt.plot(weeks, data_2026, color='#EA4335', label='2026', linewidth=2)
    plt.plot(weeks, data_median, color='#FBBC05', label='Moving median 4 tahun (2022,2023,2024,2025)', linewidth=2)

    plt.xticks(weeks, fontsize=8, rotation=0)
    plt.yticks(range(0, 1500, 250), fontsize=9)
    plt.grid(False)
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False, fontsize=9)
    
    # Remove top and right spines
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    plt.tight_layout()
    
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=300)
    plt.close()
    img_stream.seek(0)
    return img_stream

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list, df_graf):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # (Bahagian Logo & Header dikekalkan sama...)
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.8))

    for text, size in [("LAPORAN HARIAN BWKK", 10.5), ("CPRC JKN SELANGOR", 10.5)]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(para.add_run(text), size, bold=True)
        para.paragraph_format.space_after = Pt(0)

    # --- SECTION 1, 2 (Kekal Seperti Skrip Sebelumnya) ---
    # ... [Kod Section 1.0 & 2.0 & Jadual 2.1 dipendekkan untuk fokus pada Section 3.0] ...

    # --- SECTION 3.0 (VEKTOR) ---
    p3_head = doc.add_paragraph()
    p3_head.paragraph_format.space_before = Pt(24)
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    # Jadual 3 (Kekal Sama)
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    # ... [Kod pembinaan Jadual 3 dikekalkan] ...
    # (Kod pembinaan Jadual 3 anda di sini)

    # --- TAMBAHAN: RAJAH 1 (GRAF TREND) ---
    doc.add_paragraph().paragraph_format.space_before = Pt(18)
    img_graf = create_trend_chart(df_graf)
    
    # Judul Rajah
    add_table_title(doc, "Rajah 1", "Carta Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor", is_figure=True)
    
    # Masukkan Gambar
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_img = p_img.add_run()
    run_img.add_picture(img_graf, width=Inches(6.0))

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    # ... [Kod Section 4.0 dikekalkan sama] ...

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Report Generator", layout="centered")
st.title("📄 BWKK Report Generator + Graf Trend")

f1 = st.file_uploader("📥 Muat Naik Excel Notifikasi Harian", type=["xlsx", "xls"])
f2 = st.file_uploader("📥 Muat Naik Excel Linelisting Wabak", type=["xlsx", "xls"])

if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            # Load data as before...
            # Tambahan: Ambil data Graf dari Google Sheet
            df_graf = pd.read_csv(GRAF_URL, header=None) # Load raw tanpa header untuk mapping range tepat
            
            # (Proses data lain seperti matrix, wabak_df, vector_df, bkk_table_final...)
            # ... kod pemprosesan data anda ...

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, (len(bkk_details)==0), bkk_details, df_yesterday_list, df_graf)
            st.success("✅ Laporan berjaya dijana!")
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_Versi_Graf_{date.today()}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
