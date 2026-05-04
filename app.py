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

def set_cell_paddings(cell, top=None, start=None, bottom=None, end=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = parse_xml(r'<w:tcMar {}/>'.format(nsdecls('w')))
    for margin, value in [('top', top), ('left', start), ('bottom', bottom), ('right', end)]:
        if value is not None:
            node = parse_xml(r'<w:{} {} w:w="{}" w:type="dxa"/>'.format(margin, nsdecls('w'), value))
            tcMar.append(node)
    tcPr.append(tcMar)

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
    day_name = days_ms.get(target_date.strftime("%A"))
    month_name = months_ms.get(target_date.month)
    return f"{target_date.day:02d} {month_name} {target_date.year} ({day_name})"

def apply_font(run, size, bold=True):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.bold = bold

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty):
    doc = Document()
    today = date.today()
    yesterday = today - timedelta(days=1)
    
    # ... (Bahagian atas dokumen kekal sama seperti skrip sebelumnya) ...

    # --- SECTION 4.0 (BKK) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    prefix_41 = "4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor."
    
    # LOGIK BARU: Jika tiada data dalam sheet rujukan (Turus C: TKH LAPOR)
    if is_bkk_empty:
        h41_text = f"{prefix_41} Tiada insiden dilaporkan pada {get_malay_date(yesterday)}."
    else:
        total_val_bkk = clean_val(bkk_table_df.iloc[-1]['JUMLAH'])
        h41_text = f"{prefix_41} Sejumlah {total_val_bkk} input notifikasi Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK) telah diterima pada {get_malay_date(yesterday)} dengan pecahan mengikut jenis insiden seperti dalam jadual 4."
    
    apply_font(doc.add_paragraph().add_run(h41_text), 10, bold=False)

    # ... (Bahagian pembinaan jadual dan tandatangan plain text kekal sama) ...
    # --- TAMBAHAN: BAHAGIAN TANDATANGAN (Plain Text) ---
    doc.add_paragraph() 
    p_petugas = doc.add_paragraph()
    run_p = p_petugas.add_run("Petugas   :")
    apply_font(run_p, 10, bold=False)
    
    p_jawatan1 = doc.add_paragraph()
    run_j1 = p_jawatan1.add_run("Jawatan  :")
    apply_font(run_j1, 10, bold=False)
    p_jawatan1.paragraph_format.space_after = Pt(36) 
    
    p_ketua = doc.add_paragraph()
    run_kp = p_ketua.add_run("Ketua Petugas :")
    apply_font(run_kp, 10, bold=False)
    
    p_jawatan2 = doc.add_paragraph()
    run_j2 = p_jawatan2.add_run("Jawatan  :")
    apply_font(run_j2, 10, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
if f1 and f2:
    if st.button("🚀 Jana Laporan Lengkap"):
        try:
            yesterday = date.today() - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y") # Format tarikh dalam Sheet

            # ... (S1, S2, S3 Logic) ...

            # S4 Logic - Menarik data BKK dan menyemak tarikh
            with st.spinner('Menarik data BKK...'):
                df_bkk_full = pd.read_csv(SHEET_BKK_URL, header=None)
                
                # Check Turus C (indeks 2) untuk tarikh TKH LAPOR
                # Semak jika tarikh semalam wujud dalam turus C
                tkh_lapor_col = df_bkk_full.iloc[:, 2].astype(str)
                is_bkk_empty = not tkh_lapor_col.str.contains(yesterday_str).any()

                bkk_raw = df_bkk_full.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
                bkk_raw.columns = bkk_raw.iloc[0]
                bkk_table_final = bkk_raw[1:].reset_index(drop=True)
                
                # Rename columns
                bkk_map = {'GOMBAK':'GBK','HULU LANGAT':'HL','HULU SELANGOR':'HS','KLANG':'KLG','KUALA LANGAT':'KL','KUALA SELANGOR':'KS','PETALING':'PTG','SABAK BERNAM':'SB','SEPANG':'SPG','PK P.KLANG':'PK.KLG','PK KLIA':'PK.KLIA'}
                bkk_table_final = bkk_table_final.rename(columns=bkk_map)

            doc_out = generate_docx(matrix, col_totals, wabak_df, v_data, bkk_table_final, is_bkk_empty)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_out, file_name=f"Laporan_BWKK_{date.today()}.docx")

        except Exception as e:
            st.error(f"Ralat: {e}")
