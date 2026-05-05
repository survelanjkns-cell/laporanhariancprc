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

# GSheet Links
SHEET_ID_VECTOR = "1bjyNcntm-I6nRaIVkVdJqJRAzn5r2tYFfjUAN0emv9w"
GSHEET_VECTOR_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_VECTOR}/export?format=csv&gid=0"

# GSheet BKK
SHEET_BKK_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1352807145"
SHEET_BKK_TABLE_URL = "https://docs.google.com/spreadsheets/d/1Fp6IORRfdWSJCTC8vqSSoQz6RpCpNXHzO6jj0tHEf2c/export?format=csv&gid=1342717767"

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
    run_label = p.add_run(f"{label} : ")
    apply_font(run_label, 11, bold=True)
    run_title = p.add_run(title)
    apply_font(run_title, 11, bold=False)
    p.paragraph_format.space_after = Pt(6)

def add_pkd_note(doc):
    p = doc.add_paragraph()
    run = p.add_run("*Nota : GBK, Gombak; HL, Hulu Langat; HS, Hulu Selangor; KLG, Klang; KL, Kuala Langat; KS, Kuala Selangor; PTG, Petaling; SB, Sabak Bernam; SPG, Sepang.")
    apply_font(run, 7, bold=False)

# --- FUNGSI GENERATE AYAT BKK DINAMIK ---
def generate_bkk_narrative(bkk_incidents, target_date_str):
    num_map = {1: "satu (1)", 2: "dua (2)", 3: "tiga (3)", 4: "empat (4)", 5: "lima (5)"}
    total = len(bkk_incidents)
    total_word = num_map.get(total, f"{total} ({total})")
    
    prefix = "4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. "
    
    if total == 1:
        row = bkk_incidents.iloc[0]
        return f"{prefix}Terdapat satu (1) kejadian dilaporkan pada {target_date_str} iaitu kejadian {str(row['KEJADIAN']).lower()} di alamat {row['ALAMAT KEJADIAN']}, {row['DAERAH']}."
    
    summary_parts = []
    grouped = bkk_incidents.groupby('KEJADIAN')
    for incident_type, group in grouped:
        count = len(group)
        details = [f"{r['ALAMAT KEJADIAN']}, {r['DAERAH']}" for _, r in group.iterrows()]
        summary_parts.append(f"{count} kejadian {incident_type.lower()} di " + " dan ".join(details))
    
    return f"{prefix}Sejumlah {total_word} iaitu " + " manakala ".join(summary_parts) + "."

# --- DOCX GENERATOR ---
def generate_docx(matrix_df, col_sums, wabak_df, vector_df, bkk_table_df, bkk_incidents, yesterday_date):
    doc = Document()
    today = date.today()
    yesterday_str = get_malay_date(yesterday_date)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(2.54)
    content_width = section.page_width - section.left_margin - section.right_margin

    # 1. Logo & Tajuk
    logo_path = "logo.png.jpg"
    if os.path.exists(logo_path):
        doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[-1].add_run().add_picture(logo_path, width=Inches(1.8))

    for t in ["LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", "JABATAN KESIHATAN NEGERI SELANGOR"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(t), 10.5, bold=True)
        p.paragraph_format.space_after = Pt(0)

    # 3. Jadual Tarikh Hijau
    doc.add_paragraph()
    it = doc.add_table(rows=1, cols=2)
    it.width = content_width
    for i in range(2):
        c = it.cell(0, i)
        set_cell_background(c, "C6E0B4")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        txt = f"\nTarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        apply_font(p.add_run(txt), 11, bold=True)

    # --- 1.0 (Size 8) ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("1.0 Ringkasan Laporan Input Enotifikasi"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"1.1 Sejumlah {int(col_sums['Grand Total'])} input notifikasi telah diterima pada {yesterday_str}."), 11, bold=False)
    add_table_title(doc, "Jadual 1", "Senarai Input eNotifikasi")
    
    t1 = doc.add_table(rows=len(matrix_df) + 2, cols=len(TEMPLATE_PKDS) + 2)
    t1.style = 'Table Grid'
    pkd_map = {'PKD GOMBAK': 'GBK', 'PKD HULU LANGAT': 'HL', 'PKD HULU SELANGOR': 'HS', 'PKD KLANG': 'KLG', 'PKD KUALA LANGAT': 'KL', 'PKD KUALA SELANGOR': 'KS', 'PKD PETALING': 'PTG', 'PKD SABAK BERNAM': 'SB', 'PKD SEPANG': 'SPG'}
    
    h1 = t1.rows[0].cells
    apply_font(h1[0].paragraphs[0].add_run("PENYAKIT"), 8, bold=True)
    set_cell_background(h1[0], "BFDFFF")
    for i, pkd in enumerate(TEMPLATE_PKDS):
        apply_font(h1[i+1].paragraphs[0].add_run(pkd_map.get(pkd)), 8, bold=True)
        set_cell_background(h1[i+1], "BFDFFF")
    apply_font(h1[-1].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(h1[-1], "FFFF00")

    for r_idx, (peny, row_data) in enumerate(matrix_df.iterrows()):
        cells = t1.rows[r_idx + 1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(peny)), 8, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        for c_idx, val in enumerate(row_data):
            p = cells[c_idx+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(val))), 8, bold=True)
            if c_idx == len(row_data)-1: set_cell_background(cells[c_idx+1], "FFFFB3")

    f1 = t1.rows[-1].cells
    apply_font(f1[0].paragraphs[0].add_run("Jumlah"), 8, bold=True)
    set_cell_background(f1[0], "FFFF00")
    for i, val in enumerate(col_sums):
        p = f1[i+1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(int(val))), 8, bold=True)
        set_cell_background(f1[i+1], "FFFF00")
    
    doc.add_paragraph()
    add_pkd_note(doc)

    # --- 2.0 (Size 9) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("2.0 Ringkasan Laporan Notifikasi Wabak"), 11, bold=True)
    apply_font(doc.add_paragraph().add_run(f"2.1 Sejumlah {int(wabak_df['HARIAN'].sum())} input notifikasi wabak diterima pada {yesterday_str}."), 11, bold=False)
    add_table_title(doc, "Jadual 2", "Senarai Notifikasi Wabak")
    
    t2 = doc.add_table(rows=len(wabak_df) + 2, cols=3)
    t2.style = 'Table Grid'
    headers_2 = ["PENYAKIT", "HARIAN", "KUMULATIF"]
    for i, h in enumerate(headers_2):
        c = t2.cell(0, i)
        set_cell_background(c, "BFDFFF")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(h), 9, bold=True)

    for i, (peny, row) in enumerate(wabak_df.iterrows()):
        cells = t2.rows[i+1].cells
        apply_font(cells[0].paragraphs[0].add_run(str(peny)), 9, bold=True)
        set_cell_background(cells[0], "D9E9FF")
        for j, v in enumerate([row['HARIAN'], row['KUMULATIF']]):
            p = cells[j+1].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_font(p.add_run(str(int(v))), 9, bold=True)

    f2 = t2.rows[-1].cells
    apply_font(f2[0].paragraphs[0].add_run("JUMLAH"), 9, bold=True)
    for j, v in enumerate([wabak_df['HARIAN'].sum(), wabak_df['KUMULATIF'].sum()]):
        p = f2[j+1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(p.add_run(str(int(v))), 9, bold=True)
    for c in range(3): set_cell_background(f2[c], "FFFF00")

    # --- 3.0 (Size 10.5) ---
    doc.add_paragraph()
    apply_font(doc.add_paragraph().add_run("3.0 Ringkasan Laporan Wabak Vektor"), 11, bold=True)
    add_table_title(doc, "Jadual 3", "Senarai Notifikasi Wabak Vektor")
    t3 = doc.add_table(rows=len(vector_df) + 2, cols=7)
    t3.style = 'Table Grid'
    # Header logic diringkaskan
    h3_r1 = t3.rows[0].cells
    h3_r1[0].merge(t3.rows[1].cells[0]).text = "DAERAH"
    for i, txt in zip([1, 3, 5], ["DENGGI", "MALARIA", "CHIKUNGUNYA"]):
        h3_r1[i].merge(h3_r1[i+1]).text = txt
        set_cell_background(h3_r1[i], "BFDFFF")
        h3_r1[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_font(h3_r1[i].paragraphs[0].runs[0], 10.5)
    set_cell_background(h3_r1[0], "BFDFFF")

    for i in range(len(vector_df)):
        cells = t3.rows[i+2].cells
        for j in range(7):
            val = vector_df.iloc[i, j]
            p = cells[j].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
            apply_font(p.add_run(str(int(float(val))) if j > 0 else str(val)), 10.5)
            if i == len(vector_df)-1: set_cell_background(cells[j], "FFFF00")
            elif j == 0: set_cell_background(cells[j], "FCE4D6")

    # --- 4.0 (Naratif Dinamik) ---
    doc.add_page_break()
    apply_font(doc.add_paragraph().add_run("4.0 Ringkasan Laporan Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)"), 11, bold=True)
    
    narrative = generate_bkk_narrative(bkk_incidents, yesterday_str) if not bkk_incidents.empty else f"4.1 Jadual di bawah menunjukkan jumlah kejadian insiden bencana, kecemasan dan krisis (BKK) di negeri Selangor. Tiada insiden dilaporkan pada {yesterday_str}."
    p_nar = doc.add_paragraph()
    p_nar.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    apply_font(p_nar.add_run(narrative), 11, bold=False)

    add_table_title(doc, "Jadual 4", "Senarai Kejadian Insiden Bencana, Kecemasan dan Krisis (BKK)")
    t4 = doc.add_table(rows=len(bkk_table_df) + 1, cols=len(bkk_table_df.columns))
    t4.style = 'Table Grid'
    for i, col in enumerate(bkk_table_df.columns):
        c = t4.rows[0].cells[i]
        set_cell_background(c, "BFDFFF" if i < len(bkk_table_df.columns)-2 else "FFFF00" if i == len(bkk_table_df.columns)-2 else "C6E0B4")
        apply_font(c.paragraphs[0].add_run(str(col).replace(" ", "\n")), 8)
    
    for r_idx, row in enumerate(bkk_table_df.values):
        cells = t4.rows[r_idx+1].cells
        is_last = (r_idx == len(bkk_table_df)-1)
        for c_idx, v in enumerate(row):
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            apply_font(p.add_run(clean_val(v)), 8, bold=is_last or c_idx == 0)
            if is_last: set_cell_background(cells[c_idx], "FFFF00")
            elif c_idx == 0: set_cell_background(cells[c_idx], "D9E9FF")

    doc.add_paragraph()
    add_pkd_note(doc)
    footer = doc.add_paragraph()
    apply_font(footer.add_run(f"\n*Sumber : Sistem e-notifikasi, Laporan Wabak KKM dimuat turun pada ({get_malay_date(today)} @ 10.00 am)"), 9, bold=False)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- UI STREAMLIT ---
st.set_page_config(page_title="BWKK Generator", layout="centered")
st.title("📊 BWKK Report Generator")

f1 = st.file_uploader("📂 Notifikasi Harian (Excel)", type=["xlsx"])
f2 = st.file_uploader("📂 Linelisting Wabak (Excel)", type=["xlsx"])

if f1 and f2:
    if st.button("🚀 Jana Laporan"):
        try:
            yesterday = date.today() - timedelta(days=1)
            yesterday_fmt = yesterday.strftime("%d/%m/%Y")

            # Data S1
            df1 = pd.read_excel(f1)
            df1 = df1[(df1['Notifikasi Status'] != 'Abai Notifikasi') & (df1['Pejabat Kesihatan'].isin(TEMPLATE_PKDS))]
            matrix = pd.crosstab(df1['Diagnosis'], df1['Pejabat Kesihatan']).reindex(columns=TEMPLATE_PKDS, fill_value=0)
            matrix['Grand Total'] = matrix.sum(axis=1)
            matrix = matrix.sort_values('Grand Total', ascending=False)
            
            # Data S2
            df2 = pd.read_excel(f2)
            df2['Tarikh Isytihar Wabak'] = pd.to_datetime(df2['Tarikh Isytihar Wabak']).dt.date
            df2['PENYAKIT'] = df2['PENYAKIT'].apply(lambda x: "ILI/INFLUENZA" if any(s in str(x).upper() for s in ["ILI", "INFLUENZA"]) else x)
            wb_sum = [{'PENYAKIT': d, 'HARIAN': len(df2[(df2['PENYAKIT'] == d) & (df2['Tarikh Isytihar Wabak'] == yesterday)]), 'KUMULATIF': len(df2[df2['PENYAKIT'] == d])} for d in df2['PENYAKIT'].unique() if not pd.isna(d)]
            wabak_df = pd.DataFrame(wb_sum).set_index('PENYAKIT').sort_values('KUMULATIF', ascending=False)

            # Data S3
            raw_v = pd.read_csv(GSHEET_VECTOR_URL, header=None)
            idx = raw_v[raw_v.apply(lambda r: r.astype(str).str.contains('Petaling').any(), axis=1)].index[0]
            v_data = raw_v.iloc[idx : idx + 10, 13:20]

            # Data S4
            df_bkk_full = pd.read_csv(SHEET_BKK_URL)
            df_bkk_full.columns = df_bkk_full.columns.str.strip()
            # Filter insiden semalam
            bkk_incidents = df_bkk_full[df_bkk_full['TKH LAPOR'].astype(str).str.contains(yesterday_fmt)]
            
            # Table S4
            bkk_t_raw = pd.read_csv(SHEET_BKK_TABLE_URL, header=None).iloc[1:, 33:47].dropna(how='all').reset_index(drop=True)
            bkk_t_raw.columns = bkk_t_raw.iloc[0]
            bkk_table_final = bkk_t_raw[1:].rename(columns={'GOMBAK':'GBK','HULU LANGAT':'HL','HULU SELANGOR':'HS','KLANG':'KLG','KUALA LANGAT':'KL','KUALA SELANGOR':'KS','PETALING':'PTG','SABAK BERNAM':'SB','SEPANG':'SPG'})

            doc_out = generate_docx(matrix, matrix.sum(), wabak_df, v_data, bkk_table_final, bkk_incidents, yesterday)
            st.download_button("⬇️ Muat Turun Docx", data=doc_out, file_name=f"Laporan_BWKK_{yesterday}.docx")
        except Exception as e:
            st.error(f"Ralat: {e}")
