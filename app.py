import streamlit as st
import pandas as pd
from datetime import date, timedelta
import io
import os
import re
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
    if pd.isna(val) or str(val).strip() in ["", "-", "nan"]: 
        return "-"
    return str(val).strip()

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

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, is_bkk_empty, bkk_details, yesterday_date):
    doc = Document()
    today = date.today()
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.0)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo
    if os.path.exists("logo.png.jpg"):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.add_run().add_picture("logo.png.jpg", width=Inches(1.5))

    # 2. Tajuk Utama
    for text in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(text), 10.5, bold=True)
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # 3. Jadual Tarikh
    info_t = doc.add_table(rows=1, cols=2)
    info_t.width = content_width
    for i, txt in enumerate([f"Tarikh: {get_malay_date(today)}", f"Minggu Epidemiologi: {((today - date(2026,1,4)).days // 7) + 1}/{today.year}"]):
        cell = info_t.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(txt), 11, bold=True)

    # --- SECTION 1.0 (eNotifikasi) ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    p11 = doc.add_paragraph()
    apply_font(p11.add_run(f"1.1 Sejumlah {int(col_sums['Grand Total'])} notifikasi telah diterima pada {get_malay_date(yesterday_date)}."), 11, bold=False)

    # --- SECTION 2.0 (Wabak - FIXED) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    harian_total = int(wabak_df['HARIAN'].sum())
    p21 = doc.add_paragraph()
    apply_font(p21.add_run(f"2.1 Sejumlah {harian_total} input notifikasi wabak diterima pada {get_malay_date(yesterday_date)}."), 11, bold=False)

    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    hdrs = ["PENYAKIT", "HARIAN", "KUMULATIF"]
    for i, h in enumerate(hdrs):
        cell = t2.cell(0, i)
        set_cell_background(cell, "BFDFFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(h), 9, bold=True)

    for i, (idx, row) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(idx)), 9, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        for j, val in enumerate([row['HARIAN'], row['KUMULATIF']]):
            p = cells[j+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val))), 9, bold=True)

    # Footer JUMLAH yang hilang
    f_cells = t2.rows[-1].cells
    apply_font(f_cells[0].paragraphs[0].add_run("JUMLAH"), 9, bold=True)
    apply_font(f_cells[1].paragraphs[0].add_run(str(harian_total)), 9, bold=True)
    apply_font(f_cells[2].paragraphs[0].add_run(str(int(wabak_df['KUMULATIF'].sum()))), 9, bold=True)
    for c in range(3):
        set_cell_background(f_cells[c], "FFFF00")
        f_cells[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- SECTION 3.0 (Vektor - FIXED) ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    # Header logic (Simplified for safety)
    h3_cells = t3.rows[0].cells
    for i, txt in enumerate(["DAERAH", "DENGGI", "", "MALARIA", "", "CHIKUNGUNYA", ""]):
        if txt: 
            apply_font(h3_cells[i].paragraphs[0].add_run(txt), 9, bold=True)
            set_cell_background(h3_cells[i], "BFDFFF")
    
    # Isi data vektor
    for i in range(len(vector_df)):
        row_cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            txt = str(val) if pd.notna(val) and str(val) != "" else "0"
            p = row_cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
            apply_font(p.add_run(txt), 9, bold=False)

    # --- SECTION 4.0 (Naratif Dinamik) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden BKK"), 11, bold=True)
    h41 = doc.add_paragraph()
    h41.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if is_bkk_empty:
        txt_bkk = f"4.1 Tiada insiden dilaporkan pada {get_malay_date(yesterday_date)}."
    else:
        num_map = {1:"satu (1)", 2:"dua (2)", 3:"tiga (3)", 4:"empat (4)", 5:"lima (5)"}
        insiden_texts = [f"kejadian {d['kejadian'].lower()} di {d['alamat']}, {d['daerah']}" for d in bkk_details]
        detail_str = (", ".join(insiden_texts[:-1]) + " dan " + insiden_texts[-1]) if len(insiden_texts) > 1 else insiden_texts[0]
        txt_bkk = f"4.1 Terdapat {num_map.get(len(bkk_details), len(bkk_details))} kejadian dilaporkan pada {get_malay_date(yesterday_date)} iaitu {detail_str}."
    
    apply_font(h41.add_run(txt_bkk), 11, bold=False)

    # Footer Simpan
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- STREAMLIT UI ---
st.set_page_config(page_title="BWKK Generator", layout="centered")
st.title("📊 BWKK Report Generator (V2026)")

f1 = st.file_uploader("📂 Muat Naik Excel Notifikasi Harian", type=["xlsx"])
f2 = st.file_uploader("📂 Muat Naik Excel Linelisting Wabak", type=["xlsx"])

if f1 and f2:
    if st.button("🚀 Jana Laporan"):
        try:
            yesterday = date.today() - timedelta(days=1)
            yesterday_str = yesterday.strftime("%d/%m/%Y")

            # Process S1 & S2
            df1 = pd.read_excel(f1)
            df1 = df1[df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS)]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            col_totals = matrix.sum(axis=0)

            df2 = pd.read_excel(f2)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            wb_sum = []
            for d in df2['PENYAKIT'].unique():
                if pd.isna(d): continue
                h = len(df2[(df2['PENYAKIT'] == d) & (df2['Tarikh Isytihar Wabak'] == yesterday)])
                k = len(df2[df2['PENYAKIT'] == d])
                wb_sum.append({'PENYAKIT': d, 'HARIAN': h, 'KUMULATIF': k})
            wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values('KUMULATIF', ascending=False)

            # S3 - Vektor (Google Sheet)
            raw_gs = pd.read_csv(GSHEET_URL, header=None)
            # Cari baris yang ada perkataan "Petaling" untuk pivot point
            mask = raw_gs.apply(lambda r: r.astype(str).str.contains('PETALING', case=False).any(), axis=1)
            if mask.any():
                idx = mask.idxmax()
                vector_data = raw_gs.iloc[idx : idx+10, 13:20] # Kolum N hingga T
            else:
                vector_data = pd.DataFrame([["Tiada Data"]*7])

            # S4 - BKK (Google Sheet)
            df_bkk = pd.read_csv(SHEET_BKK_URL, header=None)
            # Column C index 2
            insiden_rows = df_bkk[df_bkk.iloc[:, 2].astype(str).str.contains(yesterday_str)]
            bkk_details = []
            for _, r in insiden_rows.iterrows():
                bkk_details.append({'kejadian': str(r[5]), 'alamat': str(r[8]), 'daerah': str(r[4])})
            
            # Data untuk jadual statik BKK
            bkk_table = df_bkk.iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_table.columns = bkk_table.iloc[0]
            bkk_table = bkk_table[1:]

            # Generate
            doc_file = generate_docx(matrix, col_totals, wabak_df, vector_data, bkk_table, len(bkk_details)==0, bkk_details, yesterday)
            st.download_button("⬇️ Muat Turun Laporan", data=doc_file, file_name=f"Laporan_BWKK_{date.today()}.docx")
            st.success("Laporan berjaya dijana!")

        except Exception as e:
            st.error(f"Ralat: {e}")
