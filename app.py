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
import urllib.parse
import io
import os
import re
import requests

# Import matplotlib untuk Graf 6.1
try:
    import matplotlib.pyplot as plt
    import numpy as np
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# --- KONSTAN & MAPPING DATA ---
TEMPLATE_PKDS = [
    'PKD GOMBAK', 'PKD HULU LANGAT', 'PKD HULU SELANGOR', 'PKD KLANG',
    'PKD KUALA LANGAT', 'PKD KUALA SELANGOR', 'PKD PETALING',
    'PKD SABAK BERNAM', 'PKD SEPANG'
]

AVG_HARIAN_FIGURES = {
    "Denggi": 426, "COVID-19": 54, "HFMD": 52, "Tuberculosis": 28,
    "Keracunan Makanan": 22, "Measles": 12, "Viral Hepatitis": 9,
    "Avian Influenza": 0, "HIV/AIDS": 7, "Leptospirosis": 6,
    "Dysentery": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
    "Gonorrhoea": 2, "Pertussis": 2, "Malaria": 1, "Mers-Cov": 1
}

SHEET_ID = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GID = "0"

# --- URL GOOGLE SHEET WABAK (LIVE & AUTOMATIC SNAPSHOT) ---
SHEET_ID_WABAK = "1SMu8z0MONnxkduZEaRyVNrEnH7KkvnJ9EjuVxSi3WOY"
GID_RAW = "0"  

RAW_GID_AUDIT = "1442328310" 
GID_AUDIT_YESTERDAY = str(RAW_GID_AUDIT).replace("#", "").replace("gid=", "").strip()

# --- URL GOOGLE SHEET BKK ---
BKK_SPREADSHEET_ID = "1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c"

# --- URL GOOGLE SHEET JEREBU ---
SHEET_ID_JEREBU = "1nVvq4MOi2GiLCaKjg4b3pIlAUHkQnD15BAqmBrcPjwY"
FASILITI_JEREBU = [
    ("KK Kota Damansara", "KK KOTA DAMANSARA"),
    ("KK Shah Alam", "KK SHAH ALAM"),
    ("Hospital Shah Alam", "HOSPITAL SHAH ALAM"),
    ("KK Pandamaran", "KK PANDAMARAN"),
    ("KK Telok Datok", "KK TELOK DATOK"),
    ("KK Kuala Selangor", "KK KUALA SELANGOR")
]

CHART_IMAGE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTDprYai1uaP1L-JP6kuHRZX18AmDHX0ROEzRE37DaCHMo0cNWUvRa8R-65RZAK7XFWI6pb_-X-jF24/pubchart?oid=1681812411&format=image"

# --- HELPER DUAL-ENDPOINT CSV READER ---
def read_gsheet_csv(sheet_id, gid="0", sheet_name=None, header='default', range_val=None):
    sheet_id = str(sheet_id).strip()
    
    if sheet_name:
        encoded_name = urllib.parse.quote(sheet_name)
        gviz_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_name}"
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_name}"
    else:
        gid = str(gid).replace("#", "").replace("gid=", "").strip()
        gviz_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}"
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        
    if range_val:
        export_url += f"&range={range_val}"
    
    last_err = None
    for url in [export_url, gviz_url]:
        try:
            if header is None:
                return pd.read_csv(url, header=None)
            else:
                return pd.read_csv(url)
        except Exception as e:
            last_err = e
            continue
            
    raise Exception(f"Gagal memuat turun data dari Google Sheet (ID: {sheet_id}, Sheet: {sheet_name or gid}). Ralat: {last_err}")

# --- HELPERS FUNGSI & APIMS ---
def parse_apims_pasted_text(raw_text):
    selangor_stations = {}
    stations_map = {
        "SHAH ALAM": "Shah Alam",
        "JOHAN SETIA": "Klang (Johan Setia)",
        "PETALING JAYA": "Petaling Jaya",
        "KLANG": "Klang",
        "KUALA SELANGOR": "Kuala Selangor",
        "BANTING": "Banting",
    }
    cleaned_text = raw_text.replace("\n", " ").replace("\r", " ")
    pattern = r"(Shah Alam|JOHAN SETIA|Petaling Jaya|Kuala Selangor|Banting|Klang)(?:.*?)(?:CA\w+|MCAQM\w+)?\s*(\d{1,3})\*\*"
    matches = re.findall(pattern, cleaned_text, re.IGNORECASE)

    for loc_found, val_str in matches:
        loc_upper = loc_found.strip().upper()
        for key_name, std_name in stations_map.items():
            if key_name in loc_upper:
                if key_name == "KLANG" and "JOHAN SETIA" in loc_upper:
                    continue
                selangor_stations[std_name] = int(val_str)
                break
    return selangor_stations

def format_ipu_narrative(data_stesen, tarikh_hari_ini, tarikh_semalam, epi_week_lepas):
    if not data_stesen:
        return f"Pada {tarikh_hari_ini} jam 9.00 pagi, tiada data stesen pemantauan IPU dikesan. Perincian kes penyakit berkaitan jerebu yang dilaporkan oleh fasiliti sentinel pada {tarikh_semalam} adalah seperti di Jadual 6.1, manakala Rajah 6.1 hingga Rajah 6.3 menunjukkan tren mingguan konjunktivitis, URTI dan asma berbanding bacaan IPU tertinggi sehingga ME{epi_week_lepas}."
        
    unhealthy = []
    moderate = []
    good = []
    
    for loc, val in data_stesen.items():
        if val > 100:
            unhealthy.append(f"{loc} ({val})")
        elif val > 50:
            moderate.append(f"{loc} ({val})")
        else:
            good.append(f"{loc} ({val})")
            
    def join_list(items):
        if not items: return ""
        if len(items) == 1: return items[0]
        return ", ".join(items[:-1]) + " dan " + items[-1]
        
    parts = []
    if unhealthy:
        parts.append(f"tahap Tidak Sihat (101-200) iaitu {join_list(unhealthy)}")
    if moderate:
        parts.append(f"tahap Sederhana (51-100) iaitu {join_list(moderate)}")
    if good:
        parts.append(f"tahap Baik (0-50) iaitu {join_list(good)}")
        
    status_str = ", manakala ".join(parts) if len(parts) > 1 else "".join(parts)
    jumlah_stesen = len(data_stesen)
    
    teks = f"Pada {tarikh_hari_ini} jam 9.00 pagi, kesemua {jumlah_stesen} stesen pemantauan di Selangor merekodkan IPU pada {status_str}. Perincian kes penyakit berkaitan jerebu yang dilaporkan oleh fasiliti sentinel pada {tarikh_semalam} adalah seperti di Jadual 6.1, manakala Rajah 6.1 hingga Rajah 6.3 menunjukkan tren mingguan konjunktivitis, URTI dan asma berbanding bacaan IPU tertinggi sehingga ME{epi_week_lepas}."
    
    return teks

def format_wad_narrative(wad_count):
    if wad_count == 0:
        return " Tiada kemasukkan ke wad atau kematian disyaki berkaitan jerebu dilaporkan. Pemantauan penyakit berkaitan jerebu diteruskan."
    
    if wad_count == 1:
        count_str = "satu (1)"
    else:
        num_word = {
            2: "dua (2)", 3: "tiga (3)", 4: "empat (4)", 5: "lima (5)",
            6: "enam (6)", 7: "tujuh (7)", 8: "lapan (8)", 9: "sembilan (9)"
        }
        word = num_word.get(wad_count, str(wad_count))
        count_str = f"sejumlah {word}"
        
    return f" Terdapat {count_str} kemasukan ke wad disyaki berkaitan jerebu dilaporkan. Pemantauan penyakit berkaitan jerebu diteruskan."

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
    if "MONKEYPOX" in name_str:
        return "Mpox"
    if "DYSENTRY" in name_str or "DYSENTERY" in name_str:
        return "Dysentery"
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

def get_epi_week_last_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date: return "N/A"
    days_diff = (target_date - start_date).days
    current_epi = (days_diff // 7) + 1
    last_epi = current_epi - 1
    if last_epi == 0:
        return f"52/{target_date.year - 1}"
    return f"{last_epi}/{target_date.year}"

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
    p.paragraph_format.space_before = Pt(12)
    run_label = p.add_run(f"{label} : ")
    apply_font(run_label, 11, bold=True)
    run_title = p.add_run(title)
    apply_font(run_title, 11, bold=False)
    p.paragraph_format.space_after = Pt(6)

def add_pkd_note(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.15

    run_main = p.add_run("*Nota :\n")
    apply_font(run_main, 8, bold=True)

    senarai_daerah = [
        "GBK = Gombak", "HL = Hulu Langat", "HS = Hulu Selangor",
        "KLG = Klang", "KL = Kuala Langat", "KS = Kuala Selangor",
        "PTG = Petaling", "SB = Sabak Bernam", "SPG = Sepang"
    ]
    
    teks_daerah = "\n".join(senarai_daerah)
    run_item = p.add_run(teks_daerah)
    apply_font(run_item, 7, bold=False)

def add_bkk_note(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(30) 
    p.paragraph_format.line_spacing = 1.15

    run_main = p.add_run("*Nota :\n")
    apply_font(run_main, 8, bold=True)

    senarai_nota_bkk = [
        "PK PK = Pejabat Kesihatan Pelabuhan Klang",
        "PK KLIA = Pejabat Kesihatan KLIA"
    ]
    
    teks_nota = "\n".join(senarai_nota_bkk)
    run_item = p.add_run(teks_nota)
    apply_font(run_item, 7, bold=False)

def format_bkk_number(val, is_person=False):
    try:
        num = int(float(str(val).strip()))
    except:
        num = 0
        
    if num == 0:
        return "0"
        
    if num == 1 and is_person:
        return "seorang"
    
    num_word = {
        1: "satu (1)", 2: "dua (2)", 3: "tiga (3)", 4: "empat (4)", 5: "lima (5)",
        6: "enam (6)", 7: "tujuh (7)", 8: "lapan (8)", 9: "sembilan (9)"
    }
    
    return num_word.get(num, str(num))

# --- GENERATOR EXCEL AUDIT 3 SHEET ---
def generate_excel_audit(df2_filt, wabak_df, df_yesterday, yesterday_str):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df2_filt.to_excel(writer, sheet_name='Raw_Linelist_Wabak', index=False)
        wabak_df.reset_index().to_excel(writer, sheet_name='Pivot_Senarai_Penyakit', index=False)
        
        if not df_yesterday.empty:
            pivot_harian = pd.crosstab(
                df_yesterday['DAERAH (HURUF BESAR)'],
                df_yesterday['PENYAKIT'],
                margins=True,
                margins_name='Grand Total'
            )
        else:
            pivot_harian = pd.DataFrame({'Mesej': ['Tiada Wabak Baharu']})
            
        pivot_harian.to_excel(writer, sheet_name='Pivot_Wabak_Baharu_Harian')
        
    output.seek(0)
    return output

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, df_yesterday_list, jerebu_data, parsed_apims, hsa_wad_count, df_graf_konj):
    doc = Document()
    now_msia = get_msia_time()
    today = now_msia.date()
    yesterday = today - timedelta(days=1)

    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.8))

    titles = [
        "LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)",
        "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)",
        "JABATAN KESIHATAN NEGERI SELANGOR"
    ]
    for idx, text in enumerate(titles):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        apply_font(run, 10.5, bold=True)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.line_spacing = 1.0
        
        if idx == len(titles) - 1:
            para.paragraph_format.space_after = Pt(6)
        else:
            para.paragraph_format.space_after = Pt(0)

    # --- KOTAK HIJAU ---
    info_table = doc.add_table(rows=2, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False

    col_widths_info = [content_width * 0.5, content_width * 0.5]
    for row in info_table.rows:
        for idx, width in enumerate(col_widths_info):
            row.cells[idx].width = width

    for row in info_table.rows:
        for cell in row.cells:
            set_cell_background(cell, "C6E0B4")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    cell_tarikh = info_table.cell(0, 0)
    p_tarikh = cell_tarikh.paragraphs[0]
    p_tarikh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tarikh.paragraph_format.space_before = Pt(6)
    p_tarikh.paragraph_format.space_after = Pt(2)
    run_tarikh = p_tarikh.add_run(f"Tarikh : {get_malay_date(today)}")
    apply_font(run_tarikh, 11, bold=True)

    cell_epi = info_table.cell(0, 1)
    p_epi = cell_epi.paragraphs[0]
    p_epi.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_epi.paragraph_format.space_before = Pt(6)
    p_epi.paragraph_format.space_after = Pt(2)
    run_epi = p_epi.add_run(f"Minggu Epidemiologi : {get_epi_week(today)}")
    apply_font(run_epi, 11, bold=True)

    cell_nota = info_table.cell(1, 0)
    cell_nota.merge(info_table.cell(1, 1))
    p_nota = cell_nota.paragraphs[0]
    p_nota.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nota.paragraph_format.space_before = Pt(2)
    p_nota.paragraph_format.space_after = Pt(6)
    p_nota.paragraph_format.line_spacing = 1.15  

    ayat_baris1 = f"Data harian adalah berdasarkan input yang direkodkan pada {get_malay_date(yesterday)}\n"
    ayat_baris2 = f"Dimuat turun dan disemak pada {get_malay_date(today)}, jam 8.00 pagi"
    
    run_nota = p_nota.add_run(ayat_baris1 + ayat_baris2)
    apply_font(run_nota, 11, bold=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- 1.0 Ringkasan Laporan Input e-Notifikasi ---
    p1_head = doc.add_paragraph()
    apply_font(p1_head.add_run("1.0 Ringkasan Laporan Input e-Notifikasi"), 11, bold=True)
    
    total_notifications = int(col_sums['Grand Total'])
    h11 = doc.add_paragraph()
    h11.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
    h11_text = f"Jadual di bawah menunjukkan jumlah input e-Notifikasi di negeri Selangor yang direkodkan pada {get_malay_date(yesterday)}. Sejumlah {total_notifications} input notifikasi diterima pada tarikh tersebut dengan pecahan mengikut penyakit seperti dalam Jadual 1.1."
    apply_font(h11.add_run(h11_text), 11, bold=False)

    add_table_title(doc, "Jadual 1.1", "Jumlah Input e-Notifikasi")
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 3)
    t1.style = 'Table Grid'
    t1.width = content_width 
    
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    h_cells = t1.rows[0].cells
    for i in range(len(h_cells)):
        h_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        h_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    apply_font(h_cells[0].paragraphs[0].add_run("Penyakit"), 10, bold=True)
    set_cell_background(h_cells[0], "BFDFFF")
    
    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        apply_font(cell.paragraphs[0].add_run(pkd_map.get(pkd, pkd)), 10, bold=True)
        set_cell_background(cell, "BFDFFF")
    
    apply_font(h_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run("Jumlah"), 10, bold=True)
    set_cell_background(h_cells[len(TEMPLATE_PKDS)+1], "FFFF00")
    apply_font(h_cells[len(TEMPLATE_PKDS)+2].paragraphs[0].add_run("Purata Harian"), 10, bold=True)
    set_cell_background(h_cells[len(TEMPLATE_PKDS)+2], "FFC000")

    for r_idx, (penyakit, row_data) in enumerate(matrix_df.iterrows()):
        row = t1.rows[r_idx + 1].cells
        nama_formatted = format_penyakit_name(penyakit)
        apply_font(row[0].paragraphs[0].add_run(nama_formatted), 8, bold=True)
        set_cell_background(row[0], "D9E9FF")
        row[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        for c_idx, pkd in enumerate(TEMPLATE_PKDS):
            cell = row[c_idx+1]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(row_data[pkd]))), 8, bold=True)
        
        gt_cell = row[len(TEMPLATE_PKDS)+1]
        gt_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        gt_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(gt_cell.paragraphs[0].add_run(str(int(row_data['Grand Total']))), 8, bold=True)
        set_cell_background(gt_cell, "FFFFB3")
        
        avg_cell = row[len(TEMPLATE_PKDS)+2]
        avg_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        avg_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(avg_cell.paragraphs[0].add_run(str(int(row_data.get('Average Harian', 0)))), 8, bold=True)
        set_cell_background(avg_cell, "FFC000")

    f_cells = t1.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f_cells[0], "FFFF00")
    f_cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = f_cells[i+1]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(cell.paragraphs[0].add_run(str(int(col_sums[pkd]))), 8, bold=True)
        set_cell_background(cell, "FFFF00")

    f_cells[len(TEMPLATE_PKDS)+1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    f_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(f_cells[len(TEMPLATE_PKDS)+1].paragraphs[0].add_run(str(int(col_sums['Grand Total']))), 8, bold=True)
    set_cell_background(f_cells[len(TEMPLATE_PKDS)+1], "FFFF00")

    avg_total_cell = f_cells[len(TEMPLATE_PKDS)+2]
    avg_total_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    avg_total_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_font(avg_total_cell.paragraphs[0].add_run("647"), 8, bold=True)
    set_cell_background(avg_total_cell, "FFC000")

    add_pkd_note(doc)
    
    # --- 2.0 Ringkasan Laporan Notifikasi Wabak ---
    p2_head = doc.add_paragraph()
    p2_head.paragraph_format.page_break_before = True 
    p2_head.paragraph_format.space_before = Pt(12)
    apply_font(p2_head.add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    
    harian_total = int(wabak_df['HARIAN'].sum())
    
    h21 = doc.add_paragraph()
    h21.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
    
    if harian_total == 0:
        h21_text = f"Jadual di bawah menunjukkan jumlah wabak harian, aktif dan kumulatif di negeri Selangor. Tiada wabak telah direkodkan pada {get_malay_date(yesterday)}."
    else:
        harian_total_str = format_bkk_number(harian_total, is_person=False)
        h21_text = f"Jadual di bawah menunjukkan jumlah wabak harian, aktif dan kumulatif di negeri Selangor. Sejumlah {harian_total_str} input notifikasi wabak telah direkodkan pada {get_malay_date(yesterday)}."
        
    apply_font(h21.add_run(h21_text), 11, bold=False)

    add_table_title(doc, "Jadual 2.1", "Senarai Notifikasi Wabak")
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=4)
    t2.style = 'Table Grid'
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.autofit = False  
    
    col_widths_t2 = [content_width * 0.48, content_width * 0.1733, content_width * 0.1733, content_width * 0.1733]

    for i, h in enumerate(["Penyakit", "Harian", "Aktif", "Kumulatif"]):
        cell = t2.cell(0, i)
        cell.width = col_widths_t2[i]
        apply_font(cell.paragraphs[0].add_run(h), 8, bold=True)
        set_cell_background(cell, "BFDFFF")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i, (penyakit, row_data) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        for idx_w in range(4):
            cells[idx_w].width = col_widths_t2[idx_w]
            
        apply_font(cells[0].paragraphs[0].add_run(str(penyakit)), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        for idx, col_key in enumerate(['HARIAN', 'AKTIF', 'KUMULATIF'], start=1):
            run = cells[idx].paragraphs[0].add_run(str(int(row_data[col_key])))
            apply_font(run, 8, bold=True)
            cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    f2_cells = t2.rows[-1].cells
    footer_vals = ["Jumlah", str(int(wabak_df['HARIAN'].sum())), str(int(wabak_df['AKTIF'].sum())), str(int(wabak_df['KUMULATIF'].sum()))]
    for i, txt in enumerate(footer_vals):
        f2_cells[i].width = col_widths_t2[i]
        apply_font(f2_cells[i].paragraphs[0].add_run(txt), 8, bold=True)
        set_cell_background(f2_cells[i], "FFFF00")
        f2_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    if df_yesterday_list:
        doc.add_paragraph()
        tarikh_semalam_str = get_malay_date(yesterday)
        add_table_title(doc, "Jadual 2.2", f"Senarai Wabak Yang Dilaporkan pada {tarikh_semalam_str}")
        
        t21 = doc.add_table(rows=1, cols=5)
        t21.style = 'Table Grid'
        t21.width = content_width 
        t21.allow_autofit = False
        set_repeat_table_header(t21.rows[0])

        widths_21 = [content_width * 0.05, content_width * 0.2, content_width * 0.2, content_width * 0.4, content_width * 0.15]
        h21_headers = ["Bil", "Wabak", "Daerah", "Tempat Berlaku", "Kadar Serangan"]
        for i, txt in enumerate(h21_headers):
            cell = t21.cell(0, i)
            cell.width = widths_21[i]
            set_cell_background(cell, "BFDFFF")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(txt), 10, bold=True)

        for idx, item in enumerate(df_yesterday_list, start=1):
            row = t21.add_row().cells
            for i in range(5): 
                row[i].width = widths_21[i]
                row[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            row[0].text = f"{idx}."
            
            p_wabak = row[1].paragraphs[0]
            p_wabak.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_name = p_wabak.add_run(str(item[0]))
            apply_font(run_name, 8, bold=False)
            p_wabak.add_run("\n")
            kategori_display = "(Household)" if str(item[3]).strip() == "Rumah Persendirian" else "(Institusi)"
            run_cat = p_wabak.add_run(kategori_display)
            apply_font(run_cat, 8, bold=False)

            row[2].text = str(item[1]).title() 
            row[3].text = str(item[2]) 

            n_kes = float(item[4]) if pd.notna(item[4]) else 0
            n_dedah = float(item[5]) if pd.notna(item[5]) else 0
            
            if n_dedah > 0:
                calc_pct = (n_kes / n_dedah) * 100
                pct_str = f"{int(calc_pct)}%" if calc_pct % 1 == 0 else f"{calc_pct:.2f}%"
            else:
                pct_str = "0%"

            p_ar = row[4].paragraphs[0]
            p_ar.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_ar_main = p_ar.add_run(f"{int(n_kes)}/{int(n_dedah)}")
            apply_font(run_ar_main, 8, bold=False)
            p_ar.add_run("\n")
            run_ar_pct = p_ar.add_run(f"({pct_str})")
            apply_font(run_ar_pct, 8, bold=False)

            for c in range(5):
                p = row[c].paragraphs[0]
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c == 3 else WD_ALIGN_PARAGRAPH.CENTER
                if p.runs: apply_font(p.runs[0], 8, bold=False)

    doc.add_paragraph()
    doc.add_paragraph()

    # --- 3.0 Ringkasan Laporan Wabak Vektor ---
    p3_head = doc.add_paragraph()
    apply_font(p3_head.add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    
    try: 
        val_sum = float(vector_df.iloc[-1, 1]) + float(vector_df.iloc[-1, 3]) + float(vector_df.iloc[-1, 5])
        xx_v = int(val_sum)
    except: 
        xx_v = 0
    
    if xx_v == 0:
        teks_vektor_awal = "Tiada notifikasi"
    else:
        xx_v_display = format_bkk_number(xx_v, is_person=False)
        teks_vektor_awal = f"Sebanyak {xx_v_display} notifikasi"
        
    h31 = doc.add_paragraph()
    h31.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
    h31_text = f"Jadual di bawah menunjukkan jumlah wabak vektor harian dan kumulatif di negeri Selangor. {teks_vektor_awal} wabak vektor telah direkodkan pada {get_malay_date(yesterday)} dengan pecahan mengikut penyakit seperti dalam Jadual 3.1. Rajah 3.1 pula menunjukkan tren kes mingguan denggi yang didaftarkan dari tahun 2025 hingga kini."
    apply_font(h31.add_run(h31_text), 11, bold=False)

    add_table_title(doc, "Jadual 3.1", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    t3.width = content_width 
    t3.allow_autofit = False  

    col_widths_v = [content_width * 0.25, content_width * 0.125, content_width * 0.125, content_width * 0.125, content_width * 0.125, content_width * 0.125, content_width * 0.125]
    
    h3_r1 = t3.rows[0].cells
    h3_r1[0].merge(t3.rows[1].cells[0]).text = "Daerah"
    h3_r1[1].merge(h3_r1[2]).text = "Denggi"
    h3_r1[3].merge(h3_r1[4]).text = "Malaria"
    h3_r1[5].merge(h3_r1[6]).text = "Chikungunya"
    
    for i in [0, 1, 3, 5]:
        cell = h3_r1[i]
        set_cell_background(cell, "BFDFFF")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        if i == 0: cell.width = col_widths_v[0]
        else: cell.width = col_widths_v[i] + col_widths_v[i+1]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs: apply_font(p.runs[0], 10, bold=True)
        else: apply_font(p.add_run(cell.text), 10, bold=True)

    h3_r2 = t3.rows[1].cells
    for i in range(1, 7):
        h3_r2[i].text = "Harian" if i % 2 != 0 else "Kumulatif"
        h3_r2[i].width = col_widths_v[i]
        set_cell_background(h3_r2[i], "BFDFFF")
        h3_r2[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h3_r2[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.runs[0], 9, bold=True)

    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        is_jumlah_row = str(vector_df.iloc[i, 0]).strip().upper() == "JUMLAH"
        
        for j in range(7):
            val = vector_df.iloc[i, j]
            row_cells[j].width = col_widths_v[j]
            if pd.isna(val) or str(val).lower() == "nan": display_val = "-"
            elif j == 0: 
                display_val = "Jumlah" if is_jumlah_row else str(val).title()
            else:
                try: display_val = f"{int(float(val)):,}"
                except: display_val = str(val)
            
            p = row_cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(display_val)
            
            if is_jumlah_row: 
                set_cell_background(row_cells[j], "FFFF00") 
            elif j == 0: 
                set_cell_background(row_cells[j], "FCE4D6") 
                
            apply_font(run, 9, bold=True)
            row_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    try:
        response = requests.get(CHART_IMAGE_URL)
        if response.status_code == 200:
            doc.add_paragraph()
            
            border_table = doc.add_table(rows=1, cols=1)
            border_table.style = 'Table Grid'
            border_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            border_table.autofit = False
            
            cell_graf = border_table.cell(0, 0)
            cell_graf.width = Inches(6.2)
            cell_graf.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            img_stream = io.BytesIO(response.content)
            p_img = cell_graf.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(6)
            run_img = p_img.add_run()
            run_img.add_picture(img_stream, width=Inches(6.0))
            
            p_rajah_head = doc.add_paragraph()
            p_rajah_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_rajah_head.paragraph_format.space_before = Pt(8)
            p_rajah_head.paragraph_format.space_after = Pt(12)
            
            run_rajah_label = p_rajah_head.add_run("Rajah 3.1 : ")
            apply_font(run_rajah_label, 11, bold=True)
            run_rajah_title = p_rajah_head.add_run("Tren Kes Mingguan Denggi Didaftar Bagi Tahun 2025 - 2026 Negeri Selangor")
            apply_font(run_rajah_title, 11, bold=False)
            
        else:
            st.warning("Gagal memuat turun imej graf.")
    except Exception as e:
        st.error(f"Ralat semasa memproses Rajah 3.1: {e}")

    # --- 4.0 Ringkasan Laporan Kejadian Bencana, Kecemasan dan Krisis (BKK) ---
    doc.add_page_break()
    p4_head = doc.add_paragraph()
    apply_font(p4_head.add_run("4.0 Ringkasan Laporan Kejadian Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)

    h41 = doc.add_paragraph()
    h41.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 

    penutup_text = " Jadual 4.1 di bawah menunjukkan jumlah kejadian bencana, kecemasan dan krisis (BKK) yang telah dilaporkan di negeri Selangor pada tahun 2026 mengikut daerah."

    if is_bkk_empty:
        h41_text = f"Tiada kejadian BKK dilaporkan pada {get_malay_date(yesterday)}.{penutup_text}"
    else:
        count = len(bkk_details)
        count_str = format_bkk_number(count, is_person=False)
        
        h41_text = f"Terdapat {count_str} kejadian BKK yang dilaporkan pada {get_malay_date(yesterday)} iaitu kejadian"
        
        ordinal_words = {1: " pertama ialah ", 2: " Kedua ialah insiden ", 3: " Ketiga ialah insiden ", 4: " Keempat ialah insiden ", 5: " Kelima ialah insiden "}
        
        narrative_parts = []
        for idx, item in enumerate(bkk_details, start=1):
            prefix = ordinal_words.get(idx, f" Insiden ke-{idx} ialah ") if count > 1 else ""
            
            kej_str = str(item['kejadian']).lower()
            alamat_str = str(item['alamat']).strip()
            daerah_str = str(item['daerah']).title()
            
            def clean_to_int(v):
                if pd.isna(v) or str(v).strip() in ["", "-", "nan", "0"]:
                    return 0
                try:
                    return int(float(str(v).strip()))
                except:
                    return 0

            kes_val = clean_to_int(item.get('bil_kes', 0))
            kem_val = clean_to_int(item.get('bil_kematian', 0))
            
            kes_str_formatted = format_bkk_number(kes_val, is_person=True)
            kem_str_formatted = format_bkk_number(kem_val, is_person=True)
            
            if kes_val > 0:
                if kes_val == 1:
                    mangsa_prefix = " Seorang mangsa terlibat"
                else:
                    mangsa_prefix = f" Sejumlah {kes_str_formatted} orang mangsa terlibat"

                if kem_val > 0:
                    if kem_val == 1:
                        status_str = f"{mangsa_prefix} dalam kejadian tersebut dengan satu (1) kematian dilaporkan."
                    else:
                        status_str = f"{mangsa_prefix} dalam kejadian tersebut dengan {kem_str_formatted} kematian dilaporkan."
                else:
                    status_str = f"{mangsa_prefix} dalam kejadian tersebut."
            else:
                if kem_val > 0:
                    if kem_val == 1:
                        status_str = f" satu (1) kematian dilaporkan dalam kejadian tersebut."
                    else:
                        status_str = f" {kem_str_formatted} kematian dilaporkan dalam kejadian tersebut."
                else:
                    status_str = ""
            
            if count == 1:
                full_event_sentence = f" {kej_str} berlaku di {alamat_str}, {daerah_str}.{status_str}"
            else:
                full_event_sentence = f"{prefix}{kej_str} di {alamat_str}, {daerah_str}.{status_str}"
                
            narrative_parts.append(full_event_sentence)
            
        h41_text += "".join(narrative_parts) + penutup_text

    apply_font(h41.add_run(h41_text), 11, bold=False)

    add_table_title(doc, "Jadual 4.1", "Jumlah Kejadian Bencana, Kecemasan dan Krisis (BKK) di Selangor Pada Tahun 2026 Mengikut Daerah")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    t4.width = content_width 
    
    h4_col_count = len(bkk_table_df.columns)
    for i, col in enumerate(bkk_table_df.columns):
        cell = t4.rows[0].cells[i]
        if i < h4_col_count-1: 
            set_cell_background(cell, "BFDFFF")
        else: 
            set_cell_background(cell, "FFFF00") 
            
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        header_text = str(col).strip()
        if header_text.upper() == "INSIDEN/BENCANA": header_text = "Insiden/Bencana"
        apply_font(p.add_run(header_text.replace(" ", "\n")), 8, bold=True)
    
    for r_idx, row_data in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last_row = (r_idx == len(bkk_table_df)-1)
        for c_idx, val in enumerate(row_data):
            cells[c_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_val(val))
            apply_font(run, 8, bold=is_last_row or c_idx == 0)
            if is_last_row: 
                set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: 
                set_cell_background(cells[c_idx], "D9E9FF")
            elif c_idx == bkk_table_df.shape[1]-1: 
                set_cell_background(cells[c_idx], "FFFFB3") 

    add_bkk_note(doc)

    # --- 5.0 Pemantauan Rumor Survelan ---
    p5_head = doc.add_paragraph()
    p5_head.paragraph_format.space_before = Pt(12)
    apply_font(p5_head.add_run("5.0 Pemantauan Rumor Survelan"), 11, bold=True)
    
    p5_space = doc.add_paragraph()
    apply_font(p5_space.add_run(""), 11)

    # --- 6.0 Pemantauan Bilik Gerakan Jerebu CPRC JKNS ---
    doc.add_page_break()
    p6_head = doc.add_paragraph()
    p6_head.paragraph_format.space_before = Pt(12)
    apply_font(p6_head.add_run("6.0 Pemantauan Bilik Gerakan Jerebu CPRC JKNS"), 11, bold=True)

    # --- KOTAK NARRATIVE JEREBU SEBELUM JADUAL 6.1 ---
    p6_intro = doc.add_paragraph()
    p6_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    tarikh_hari_ini_str = get_malay_date(today)
    tarikh_semalam_str = get_malay_date(yesterday)
    minggu_epi_lepas_str = get_epi_week_last_week(today)
    
    naratif_sebelum_61 = format_ipu_narrative(parsed_apims, tarikh_hari_ini_str, tarikh_semalam_str, minggu_epi_lepas_str)
    apply_font(p6_intro.add_run(naratif_sebelum_61), 11, bold=False)
    
    add_table_title(doc, "Jadual 6.1", "Bilangan Kes Penyakit Berkaitan Jerebu yang Dilaporkan oleh Fasiliti Sentinel Jerebu di Selangor")
    
    t6 = doc.add_table(rows=len(jerebu_data) + 3, cols=4)
    t6.style = 'Table Grid'
    t6.width = content_width 
    
    h6_cells = t6.rows[0].cells
    headers_6 = ["Fasiliti Sentinel Jerebu", "Konjunktivitis", "URTI", "Asma"]
    col_widths_6 = [content_width * 0.4, content_width * 0.2, content_width * 0.2, content_width * 0.2]
    
    for i, h in enumerate(headers_6):
        h6_cells[i].width = col_widths_6[i]
        set_cell_background(h6_cells[i], "8FAADC")
        h6_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h6_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
        apply_font(p.add_run(h), 10, bold=True)

    sum_h_konj = sum_h_urti = sum_h_asma = 0
    sum_k_konj = sum_k_urti = sum_k_asma = 0
    
    for r_idx, row_data in enumerate(jerebu_data):
        cells = t6.rows[r_idx + 1].cells
        fasiliti = row_data["Fasiliti"]
        
        cells[0].width = col_widths_6[0]
        p0 = cells[0].paragraphs[0]
        apply_font(p0.add_run(fasiliti), 10, bold=True)
        
        # Pengecualian Khas bagi Hospital Shah Alam (Konjunktivitis dan URTI dihitamkan)
        if fasiliti == "Hospital Shah Alam":
            set_cell_background(cells[1], "808080")
            set_cell_background(cells[2], "808080")
            
            cells[3].width = col_widths_6[3]
            cells[3].text = str(row_data["H_Asma"])
            cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(cells[3].paragraphs[0].runs[0], 10, bold=False)
            
            sum_h_asma += row_data["H_Asma"]
            sum_k_asma += row_data["K_Asma"]
        else:
            vals = [row_data["H_Konj"], row_data["H_URTI"], row_data["H_Asma"]]
            for i, val in enumerate(vals):
                cells[i+1].width = col_widths_6[i+1]
                cells[i+1].text = str(val)
                cells[i+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                apply_font(cells[i+1].paragraphs[0].runs[0], 10, bold=False)
                
            sum_h_konj += row_data["H_Konj"]
            sum_h_urti += row_data["H_URTI"]
            sum_h_asma += row_data["H_Asma"]
            
            sum_k_konj += row_data["K_Konj"]
            sum_k_urti += row_data["K_URTI"]
            sum_k_asma += row_data["K_Asma"]

    f_idx_1 = len(jerebu_data) + 1
    f1_cells = t6.rows[f_idx_1].cells
    f1_cells[0].text = "JUMLAH"
    apply_font(f1_cells[0].paragraphs[0].runs[0], 10, bold=True)
    
    f1_vals = [sum_h_konj, sum_h_urti, sum_h_asma]
    for i, val in enumerate(f1_vals):
        f1_cells[i+1].text = str(val)
        f1_cells[i+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(f1_cells[i+1].paragraphs[0].runs[0], 10, bold=False)
        
    for cell in f1_cells:
        set_cell_background(cell, "FFC000")

    f_idx_2 = len(jerebu_data) + 2
    f2_cells = t6.rows[f_idx_2].cells
    f2_cells[0].text = "KUMULATIF (Bermula 9/9/2026)"
    apply_font(f2_cells[0].paragraphs[0].runs[0], 10, bold=True)
    
    f2_vals = [sum_k_konj, sum_k_urti, sum_k_asma]
    for i, val in enumerate(f2_vals):
        f2_cells[i+1].text = str(val)
        f2_cells[i+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(f2_cells[i+1].paragraphs[0].runs[0], 10, bold=True)
        
    for cell in f2_cells:
        set_cell_background(cell, "FFFF00")

    p_nota_6 = doc.add_paragraph()
    p_nota_6.paragraph_format.space_before = Pt(6)
    p_nota_6.paragraph_format.space_after = Pt(12)
    run_nota_6 = p_nota_6.add_run("*Nota: Klinik Kesihatan Sentinel tidak beroperasi pada hujung minggu/cuti pelepasan am")
    apply_font(run_nota_6, 8, bold=True)
    run_nota_6.italic = True

    # --- NARATIF SELEPAS JADUAL 6.1 ---
    p6_naratif = doc.add_paragraph()
    p6_naratif.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    parts_nar = []
    if sum_h_konj > 0:
        parts_nar.append(f"{sum_h_konj} kes konjunktivitis")
    if sum_h_urti > 0:
        parts_nar.append(f"{sum_h_urti} kes URTI")
    if sum_h_asma > 0:
        parts_nar.append(f"{sum_h_asma} kes asma")
        
    if not parts_nar:
        ayat_dinamik = f"Pada {tarikh_semalam_str}, Tiada kes dilaporkan oleh fasiliti sentinel."
    else:
        if len(parts_nar) == 1:
            str_penyakit = parts_nar[0]
        elif len(parts_nar) == 2:
            str_penyakit = f"{parts_nar[0]} dan {parts_nar[1]}"
        else:
            str_penyakit = f"{parts_nar[0]}, {parts_nar[1]} dan {parts_nar[2]}"
            
        ayat_dinamik = f"Pada {tarikh_semalam_str}, Sebanyak {str_penyakit} dilaporkan oleh fasiliti sentinel."
        
    ayat_akhir = format_wad_narrative(hsa_wad_count)
    
    apply_font(p6_naratif.add_run(ayat_dinamik + ayat_akhir), 11, bold=False)

    # --- TAMBAHAN GRAF KONJUNKTIVITIS (Rajah 6.1) ---
    if HAS_MATPLOTLIB and df_graf_konj is not None and not df_graf_konj.empty:
        try:
            df_graf_konj = df_graf_konj.dropna(axis=1, how='all')
            # Bersihkan tajuk lajur daripada pembatas ruang & newline (\n)
            df_graf_konj.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df_graf_konj.columns]
            
            me_col = df_graf_konj.columns[0]
            ipu_col = df_graf_konj.columns[-1]
            clinics = df_graf_konj.columns[1:-1]
            
            def extract_me_num(val):
                try:
                    return int(re.search(r'\d+', str(val)).group())
                except:
                    return -1
                    
            df_graf_konj['me_num'] = df_graf_konj[me_col].apply(extract_me_num)
            
            last_epi_num = int(get_epi_week_last_week(today).split('/')[0])
            
            df_plot = df_graf_konj[(df_graf_konj['me_num'] > 0) & (df_graf_konj['me_num'] <= last_epi_num)].copy()
            
            if not df_plot.empty:
                fig, ax1 = plt.subplots(figsize=(9, 4.5))
                colors = ['#7030A0', '#C00000', '#92D050', '#8064A2', '#4BACC6', '#F79646']
                
                x_labels = df_plot[me_col].astype(str).tolist()
                x = np.arange(len(x_labels))
                
                for idx, clinic in enumerate(clinics):
                    y_data = pd.to_numeric(df_plot[clinic], errors='coerce').fillna(0)
                    ax1.plot(x, y_data, label=clinic, color=colors[idx % len(colors)], linewidth=2.5)
                    
                ax1.set_ylabel("BILANGAN KES", fontweight='bold', fontsize=9)
                ax1.set_xticks(x)
                ax1.set_xticklabels(x_labels, rotation=90, fontsize=8)
                
                ax2 = ax1.twinx()
                ipu_data = pd.to_numeric(df_plot[ipu_col], errors='coerce').fillna(0)
                ax2.plot(x, ipu_data, label="IPU TERTINGGI", color='red', linestyle='--', linewidth=2.5)
                
                for i, val in enumerate(ipu_data):
                    if val > 0:
                        ax2.annotate(str(int(val)), (x[i], val), textcoords="offset points", xytext=(0,6), ha='center', fontsize=7, fontweight='bold', color='#555555')
                
                ax1.grid(True, axis='y', linestyle='--', alpha=0.5)
                
                max_kes = pd.to_numeric(df_plot[clinics].stack(), errors='coerce').max()
                ax1.set_ylim(0, max(35, max_kes + 5))
                
                max_ipu = ipu_data.max()
                ax2.set_ylim(0, max(250, max_ipu + 50))
                
                lines_1, labels_1 = ax1.get_legend_handles_labels()
                lines_2, labels_2 = ax2.get_legend_handles_labels()
                fig.legend(lines_1 + lines_2, labels_1 + labels_2, loc='lower center', bbox_to_anchor=(0.5, -0.15), ncol=3, frameon=False, fontsize=8)
                
                fig.text(0.5, 0.05, "MINGGU EPID", ha='center', fontweight='bold', fontsize=9)
                
                plt.tight_layout()
                plt.subplots_adjust(bottom=0.25)
                
                img_stream = io.BytesIO()
                plt.savefig(img_stream, format='png', dpi=300, bbox_inches='tight')
                img_stream.seek(0)
                plt.close(fig)
                
                doc.add_paragraph() 
                
                p_graf = doc.add_paragraph()
                p_graf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_graf.add_run().add_picture(img_stream, width=Inches(6.0))
                
                start_me = df_plot['me_num'].iloc[0]
                end_me = df_plot['me_num'].iloc[-1]
                year_str = today.year
                
                p_caption = doc.add_paragraph()
                p_caption.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_caption.paragraph_format.space_before = Pt(6)
                
                run_cap_label = p_caption.add_run("Rajah 6.1: ")
                apply_font(run_cap_label, 11, bold=True)
                run_cap_text = p_caption.add_run(f"Tren Mingguan Kes Konjunktivitis dan Bacaan IPU Tertinggi di Fasiliti Sentinel Jerebu Selangor, ME{start_me}–ME{end_me}/{year_str}")
                apply_font(run_cap_text, 11, bold=False)
                
        except Exception as e:
            print("Ralat graf:", e)


    # --- 7.0 Rumusan oleh Ketua Petugas CPRC Selangor ---
    p7_head = doc.add_paragraph()
    p7_head.paragraph_format.space_before = Pt(12)
    apply_font(p7_head.add_run("7.0 Rumusan oleh Ketua Petugas CPRC Selangor"), 11, bold=True)
    
    p7_space = doc.add_paragraph()
    apply_font(p7_space.add_run(""), 11)

    # --- JADUAL TANDATANGAN ---
    doc.add_paragraph() 
    
    def add_sig_block(doc, title):
        p_main = doc.add_paragraph()
        p_main.paragraph_format.space_before = Pt(12)  
        p_main.paragraph_format.space_after = Pt(2)
        p_main.paragraph_format.line_spacing = 1.0
        p_main.paragraph_format.tab_stops.add_tab_stop(Inches(2.0))
        apply_font(p_main.add_run(f"{title}\t:"), 11, bold=False)
        
        p_jawatan = doc.add_paragraph()
        p_jawatan.paragraph_format.space_before = Pt(2)
        p_jawatan.paragraph_format.space_after = Pt(2)
        p_jawatan.paragraph_format.line_spacing = 1.0
        p_jawatan.paragraph_format.tab_stops.add_tab_stop(Inches(2.0))
        apply_font(p_jawatan.add_run("Jawatan\t:"), 11, bold=False)
        
        p_tarikh = doc.add_paragraph()
        p_tarikh.paragraph_format.space_before = Pt(2)
        p_tarikh.paragraph_format.space_after = Pt(2)
        p_tarikh.paragraph_format.line_spacing = 1.0
        p_tarikh.paragraph_format.tab_stops.add_tab_stop(Inches(2.0))
        apply_font(p_tarikh.add_run("Tarikh\t:"), 11, bold=False)

    add_sig_block(doc, "Disediakan oleh")
    add_sig_block(doc, "Disemak oleh")
    add_sig_block(doc, "Disahkan oleh")

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target


# --- STREAMLIT UI ---
st.set_page_config(
    page_title="LAPORAN CPRC SELANGOR", 
    page_icon="Logo Jab Kesihatan Negeri Selangor.png", 
    layout="centered"
)

# Deep CSS overrides
st.markdown("""
    <style>
    /* Paksa Latar Belakang Dark Mode */
    .stAppViewContainer, .stApp {
        background-color: #0E1117 !important;
        color: #FAFAFA !important;
    }
    
    /* Sembunyikan Header & Footer Streamlit */
    #MainMenu, header, .stAppHeader, [data-testid="stHeader"],
    footer, .stAppFooter, [data-testid="stFooter"], .stAppDeployDropdown {
        visibility: hidden !important;
        display: none !important;
    }
    
    /* TETAPAN KHUSUS TEXTAREA & INPUT BOX (DARK GREY BACKGROUND + WHITE TEXT) */
    .stTextArea textarea,
    div[data-baseweb="textarea"],
    div[data-baseweb="textarea"] > div,
    div[data-baseweb="base-input"],
    div[data-baseweb="input"] {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border-color: #4B5563 !important;
    }

    .stTextArea textarea::placeholder,
    div[data-baseweb="textarea"] textarea::placeholder {
        color: #9CA3AF !important;
        opacity: 1 !important;
    }

    /* TETAPAN KHUSUS KOTAK UPLOAD FILE */
    div[data-testid="stFileUploader"] > section {
        background-color: #1E293B !important;
        border: 1px dashed #64748B !important;
    }
    div[data-testid="stFileUploader"] span, 
    div[data-testid="stFileUploader"] small,
    div[data-testid="stFileUploader"] p {
        color: #E2E8F0 !important;
    }

    /* TETAPAN KHUSUS KOTAK MAKLUMAT (st.info, st.success, st.error, st.warning) */
    div[data-testid="stAlert"],
    div[data-testid="stNotification"],
    .stAlert {
        background-color: #1E293B !important;
        border: 1px solid #334155 !important;
    }
    div[data-testid="stAlert"] p,
    div[data-testid="stNotification"] p,
    .stAlert p {
        color: #F8FAFC !important;
    }

    /* TETAPAN BUTANG (JANA & DOWNLOAD) */
    div.stButton > button, 
    div.stDownloadButton > button,
    button[kind="primary"],
    button[kind="secondary"] {
        background-color: #374151 !important;
        color: #FFFFFF !important;
        border: 1px solid #4B5563 !important;
        font-weight: bold !important;
    }

    div.stButton > button:hover, 
    div.stDownloadButton > button:hover {
        background-color: #4B5563 !important;
        color: #FFFFFF !important;
        border-color: #9CA3AF !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📋 Jana Laporan Harian CPRC Selangor")

now_msia = get_msia_time()
today = now_msia.date()
yesterday = today - timedelta(days=1)
tarikh_guideline = get_malay_date(yesterday)

st.subheader("📁 1. Muat Naik Excel Notifikasi Harian")
st.info(f"**Guideline:** Sila muat turun file notifikasi pada **{tarikh_guideline}** dari sistem eNotifikasi dan muat naik di sini.")
f1 = st.file_uploader("Pilih fail Notifikasi Harian", type=["xlsx", "xls"], label_visibility="collapsed")

st.markdown("---")

st.subheader("📋 2. Tampal (Paste) Data Bacaan APIMS")
st.info("Sila copy teks jadual dari laman web APIMS dan paste ke dalam kotak di bawah bagi tujuan penjanaan naratif IPU (Jadual 6.1).")
raw_apims_input = st.text_area("Tampal data APIMS di sini:", height=150, placeholder="Contoh: Shah Alam 180** JOHAN SETIA 189** Petaling Jaya 175**")

if not HAS_MATPLOTLIB:
    st.sidebar.warning("Modul 'matplotlib' tidak dijumpai. Sila run `pip install matplotlib numpy` pada terminal untuk membolehkan skrip menjana graf Jerebu (Rajah 6.1).")

if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False

if f1:
    if st.session_state.get('last_file') != f1.name:
        st.session_state.report_generated = False
        st.session_state.last_file = f1.name

    if st.button("🚀 Jana Laporan Lengkap"):
        with st.spinner("Sedang memproses data dan memuat turun jadual wabak secara live..."):
            try:
                # 1. PARSE APIMS
                parsed_apims = {}
                if raw_apims_input.strip():
                    parsed_apims = parse_apims_pasted_text(raw_apims_input)

                engine_type = "xlrd" if f1.name.endswith(".xls") else "openpyxl"
                df1 = pd.read_excel(f1, engine=engine_type)
                df1 = df1[df1['Notifikasi Status'] != 'Abai Notifikasi']
                df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
                matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
                matrix['Grand Total'] = matrix.sum(axis=1)
                matrix['Average Harian'] = [AVG_HARIAN_FIGURES.get(format_penyakit_name(idx), 0) for idx in matrix.index]
                matrix = matrix.sort_values(by='Grand Total', ascending=False)
                col_totals = matrix[TEMPLATE_PKDS + ['Grand Total']].sum(axis=0)

                # --- 1. PROSES PEMBACAAN LIVE GOOGLE SHEET WABAK (HARI INI) ---
                df2 = read_gsheet_csv(SHEET_ID_WABAK, GID_RAW)
                df2.columns = df2.columns.str.strip()
                
                df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak'], dayfirst=True, errors='coerce').dt.date
                df2['Tarikh Sebenar Tamat Wabak'] = pd.to_datetime(df2['Tarikh Sebenar Tamat Wabak'], dayfirst=True, errors='coerce').dt.date
                df2['Tarikh Wabak Dijangka Tamat'] = pd.to_datetime(df2['Tarikh Wabak Dijangka Tamat'], dayfirst=True, errors='coerce').dt.date

                addr_col = 'Tempat Berlaku Wabak\n(Alamat diisi lengkap dengan :- No rumah, nama jalan, nama tempat, daerah dan Negeri)'
                cat_col = 'Kategori Tempat\n(Kategori premis berdasarkan tempat berlaku wabak)'
                
                df2_raw_audit = df2.copy()
                df2 = df2.drop_duplicates(subset=['PENYAKIT', 'Tarikh Isytihar Wabak', addr_col], keep='first')

                df_yesterday = df2[df2['Tarikh Isytihar Wabak'] == yesterday].copy()
                df_yesterday_list = df_yesterday[['PENYAKIT', 'DAERAH (HURUF BESAR)', addr_col, cat_col, 'Bilangan Kes', 'Bilangan Terdedah']].values.tolist()

                df2_filt = df2[(df2['Tarikh Isytihar Wabak'] >= date(2026, 1, 3)) & (df2['Tarikh Isytihar Wabak'] <= yesterday)].copy()
                def group_inf(n): return "ILI/ Influenza" if any(x in str(n).upper() for x in ["INFLUENZA", "ILI"]) else n
                
                df2_raw_audit['PENYAKIT_GROUP'] = df2_raw_audit['PENYAKIT'].apply(group_inf)
                df2_filt['PENYAKIT'] = df2_filt['PENYAKIT'].apply(group_inf)
                
                wb_sum = []
                for d in df2_filt['PENYAKIT'].unique():
                    if pd.isna(d): continue
                    disease_df = df2_filt[df2_filt['PENYAKIT'] == d]
                    h = len(disease_df[disease_df['Tarikh Isytihar Wabak'] == yesterday])
                    k = len(disease_df)
                    def check_active(row):
                        tamat = row['Tarikh Sebenar Tamat Wabak'] if pd.notna(row['Tarikh Sebenar Tamat Wabak']) else row['Tarikh Wabak Dijangka Tamat']
                        return True if (pd.isna(tamat) or tamat >= today) else False
                    active_count = disease_df.apply(check_active, axis=1).sum()
                    wb_sum.append({'PENYAKIT': d, 'HARIAN': h, 'AKTIF': active_count, 'KUMULATIF': k})
                
                wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values(by='KUMULATIF', ascending=False)

                # --- 2. PEMBACAAN AUTOMATIK SNAPSHOT SEMALAM ---
                df2_prev = pd.DataFrame()
                has_snapshot = False
                
                try:
                    df2_prev = read_gsheet_csv(SHEET_ID_WABAK, GID_AUDIT_YESTERDAY)
                    df2_prev.columns = df2_prev.columns.str.strip()
                    if not df2_prev.empty and len(df2_prev.columns) > 3:
                        has_snapshot = True
                        df2_prev['Tarikh Isytihar Wabak'] = pd.to_datetime(df2_prev['Tarikh Isytihar Wabak'], dayfirst=True, errors='coerce').dt.date
                        df2_prev = df2_prev.drop_duplicates(subset=['PENYAKIT', 'Tarikh Isytihar Wabak', addr_col], keep='first')
                        df2_prev_filt = df2_prev[(df2_prev['Tarikh Isytihar Wabak'] >= date(2026, 1, 3)) & (df2_prev['Tarikh Isytihar Wabak'] <= yesterday)].copy()
                        df2_prev_filt['PENYAKIT'] = df2_prev_filt['PENYAKIT'].apply(group_inf)
                except Exception:
                    has_snapshot = False

                # --- 3. AUDIT BARIS DEMI BARIS AUTOMATIK ---
                deleted_rows = []
                added_rows = []
                
                if has_snapshot and not df2_prev_filt.empty:
                    def gen_key(df_target):
                        if 'Form Response Edit URL' in df_target.columns:
                            return df_target['Form Response Edit URL'].astype(str).str.strip()
                        else:
                            return (df_target['PENYAKIT'].astype(str) + "_" + 
                                    df_target[addr_col].astype(str) + "_" + 
                                    df_target['Tarikh Isytihar Wabak'].astype(str))

                    df2_prev_filt['UNIQUE_KEY'] = gen_key(df2_prev_filt)
                    df2_filt['UNIQUE_KEY'] = gen_key(df2_filt)
                    
                    keys_prev = set(df2_prev_filt['UNIQUE_KEY'])
                    keys_today = set(df2_filt['UNIQUE_KEY'])
                    
                    deleted_keys = keys_prev - keys_today
                    added_keys = keys_today - keys_prev
                    
                    if deleted_keys:
                        deleted_rows = df2_prev_filt[df2_prev_filt['UNIQUE_KEY'].isin(deleted_keys)][['DAERAH (HURUF BESAR)', 'PENYAKIT', addr_col, 'Tarikh Isytihar Wabak']].values.tolist()
                    if added_keys:
                        added_rows = df2_filt[df2_filt['UNIQUE_KEY'].isin(added_keys)][['DAERAH (HURUF BESAR)', 'PENYAKIT', addr_col, 'Tarikh Isytihar Wabak']].values.tolist()

                # --- PEMPROSESAN DATA GOOGLE SHEET BKK & VECTOR ---
                raw_gs = read_gsheet_csv(SHEET_ID, GID, header=None)
                mask_v = raw_gs.apply(lambda r: r.astype(str).str.contains('Petaling').any(), axis=1)
                v_data = raw_gs.iloc[mask_v.idxmax() : mask_v.idxmax() + 11, 13:20]
                v_data = v_data.dropna(how='all')
                v_data = v_data[~v_data.iloc[:, 0].astype(str).str.lower().str.contains('nan')]
                v_data = v_data[~v_data.iloc[:, 0].astype(str).str.lower().str.contains('dari tarikh|tarikh')]

                df_bkk_raw_data = read_gsheet_csv(BKK_SPREADSHEET_ID, "1352807145", header=None)
                clean_date_series = df_bkk_raw_data.iloc[:, 2].astype(str).str.strip()
                df_bkk_raw_data['datetime_lapor'] = pd.to_datetime(clean_date_series, dayfirst=True, errors='coerce').dt.date
                
                insiden_semalam = df_bkk_raw_data[df_bkk_raw_data['datetime_lapor'] == yesterday]
                
                bkk_details = [{
                    'kejadian': r[5], 
                    'alamat': r[8], 
                    'daerah': r[4],
                    'bil_kes': r[9],      
                    'bil_kematian': r[10]  
                } for _, r in insiden_semalam.iterrows()]
                
                df_bkk_jadual_full = read_gsheet_csv(BKK_SPREADSHEET_ID, "1342717767", header=None)
                
                bkk_raw = df_bkk_jadual_full.iloc[0:, 33:46].dropna(how='all').reset_index(drop=True)
                bkk_raw.columns = bkk_raw.iloc[0]
                
                new_cols = []
                for col in bkk_raw.columns:
                    col_str = str(col).strip()
                    if re.search(r'(moving|median|4 tahun)', col_str, re.IGNORECASE):
                        new_cols.append('Purata Bergerak 4 Tahun (2022,2023,2024,2025)')
                    else:
                        new_cols.append(col_str)
                bkk_raw.columns = new_cols

                bkk_table_final = bkk_raw[1:].reset_index(drop=True).rename(columns={
                    'GOMBAK':'GBK','HULU LANGAT':'HL','HULU SELANGOR':'HS','KLANG':'KLG','KUALA LANGAT':'KL','KUALA SELANGOR':'KS','PETALING':'PTG','SABAK BERNAM':'SB','SEPANG':'SPG',
                    'PK P.KLANG': 'PK PK'
                })

                # --- PEMPROSESAN DATA JEREBU ---
                jerebu_data = []
                kumulatif_start_jerebu = date(2026, 9, 9)
                hsa_wad_count = 0
                
                for display_name, tab_name in FASILITI_JEREBU:
                    try:
                        df_j = read_gsheet_csv(SHEET_ID_JEREBU, sheet_name=tab_name, header=None)
                        df_j['Tarikh_Clean'] = pd.to_datetime(df_j.iloc[:, 1].astype(str).str.strip(), dayfirst=True, errors='coerce').dt.date
                        
                        df_harian = df_j[df_j['Tarikh_Clean'] == yesterday]
                        df_kumu = df_j[(df_j['Tarikh_Clean'] >= kumulatif_start_jerebu) & (df_j['Tarikh_Clean'] <= yesterday)]
                        
                        def safe_sum_jerebu(df_sub, col_idx):
                            if col_idx < len(df_sub.columns):
                                return pd.to_numeric(df_sub.iloc[:, col_idx], errors='coerce').fillna(0).sum()
                            return 0
                        
                        if display_name == "Hospital Shah Alam":
                            hsa_wad_count = int(safe_sum_jerebu(df_harian, 9)) # Column J (index 9) for HSA Wad Kemasukan Asma
                            
                            jerebu_data.append({
                                "Fasiliti": display_name,
                                "H_Konj": 0, "H_URTI": 0,
                                "H_Asma": int(safe_sum_jerebu(df_harian, 5)),
                                "K_Konj": 0, "K_URTI": 0,
                                "K_Asma": int(safe_sum_jerebu(df_kumu, 5))
                            })
                        else:
                            jerebu_data.append({
                                "Fasiliti": display_name,
                                "H_Konj": int(safe_sum_jerebu(df_harian, 5)),  
                                "H_URTI": int(safe_sum_jerebu(df_harian, 9)),  
                                "H_Asma": int(safe_sum_jerebu(df_harian, 13)),  
                                "K_Konj": int(safe_sum_jerebu(df_kumu, 5)),
                                "K_URTI": int(safe_sum_jerebu(df_kumu, 9)),
                                "K_Asma": int(safe_sum_jerebu(df_kumu, 13))
                            })
                    except Exception:
                        jerebu_data.append({
                            "Fasiliti": display_name,
                            "H_Konj": 0, "H_URTI": 0, "H_Asma": 0,
                            "K_Konj": 0, "K_URTI": 0, "K_Asma": 0
                        })
                        
                # --- DAPATKAN DATA GRAF JEREBU (BARIS 4 SEBAGAI NAMA TAJUK LAJUR) ---
                df_graf_konj = None
                if HAS_MATPLOTLIB:
                    try:
                        graf_sheet_id = "1lAOM256C1e7SI8y8EDF0ayc8di-d4yIB5qWjx7IECkI"
                        graf_sheet_name = "GRAF CONJUNCTIVITIS"
                        # Membaca bermula baris 4 (range B4:H)
                        df_graf_konj = read_gsheet_csv(graf_sheet_id, sheet_name=graf_sheet_name, range_val="B4:H")
                    except Exception as e:
                        st.warning(f"Gagal memuat turun data untuk graf jerebu: {e}")

                doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, (len(bkk_details)==0), bkk_details, df_yesterday_list, jerebu_data, parsed_apims, hsa_wad_count, df_graf_konj)
                excel_out = generate_excel_audit(df2_filt, wabak_df, df_yesterday, get_malay_date(yesterday))
                
                file_date = today.strftime("%d.%m.%y")
                
                st.session_state.doc_bytes = doc_out.getvalue()
                st.session_state.excel_bytes = excel_out.getvalue()
                st.session_state.file_name_custom = f"Laporan CPRC Selangor ({file_date}).docx"
                st.session_state.excel_name_custom = f"Audit Data Wabak CPRC ({file_date}).xlsx"
                st.session_state.deleted_rows = deleted_rows
                st.session_state.added_rows = added_rows
                st.session_state.has_snapshot = has_snapshot
                st.session_state.report_generated = True

            except Exception as e:
                st.error(f"Ralat semasa memproses data: {e}")

    # --- PAPARAN HASIL LAPORAN ---
    if st.session_state.get('report_generated', False):
        st.markdown("---")
        st.subheader("🔍 Papan Pengesahan Data (Validation Box - Live Auto-Audit)")
        
        deleted_rows = st.session_state.deleted_rows
        added_rows = st.session_state.added_rows
        has_snapshot = st.session_state.has_snapshot
        
        if deleted_rows:
            st.error(f"⚠️ **AMARAN DISCREPANCY: DIKESAN {len(deleted_rows)} BARIS DATA REKOD WABAK SEMALAM HILANG / DIPADAM DARI GOOGLE SHEET!**")
            st.write("Senarai rekod yang hilang/dipadam:")
            df_del_disp = pd.DataFrame(deleted_rows, columns=['Daerah', 'Penyakit', 'Alamat / Premis', 'Tarikh Isytihar'])
            st.dataframe(df_del_disp, use_container_width=True)
        elif has_snapshot:
            st.success("✅ **STATUS VALIDASI AUTOMATIK:** Semua baris data daripada tab `Audit_Yesterday` sepadan 100% dengan data hari ini tanpa sebarang kehilangan baris rekod.")
        else:
            st.warning("⚠️ **Peringatan Pautan Snapshot:** Tab `Audit_Yesterday` tidak dapat dibaca secara automatik.")
            st.info("💡 **Langkah Semakan:**\n1. Sila pastikan tetapan perkongsian Google Sheet diubah kepada **'Anyone with the link can view'**.\n2. Sila tekan butang **Run (▶)** di Google Apps Script sekali lagi.")

        if added_rows:
            st.info(f"ℹ️ **{len(added_rows)} Rekod Wabak Baharu Dikesan Masuk Hari Ini:**")
            df_add_disp = pd.DataFrame(added_rows, columns=['Daerah', 'Penyakit', 'Alamat / Premis', 'Tarikh Isytihar'])
            st.dataframe(df_add_disp, use_container_width=True)

        st.markdown("### 📥 Muat Turun Hasil Laporan & Data")
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            st.download_button(
                label="📄 Muat Turun Laporan Word (.docx)", 
                data=st.session_state.doc_bytes, 
                file_name=st.session_state.file_name_custom,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        with col_btn2:
            st.download_button(
                label="📊 Muat Turun Data Audit Excel (.xlsx)", 
                data=st.session_state.excel_bytes, 
                file_name=st.session_state.excel_name_custom,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
