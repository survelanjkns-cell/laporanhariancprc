import os
import pandas as pd
from datetime import date, timedelta
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- HELPER FUNCTIONS FOR WORD ---

def get_epi_week(target_date):
    """Calculates Epi Week based on Start Date: Jan 4, 2026"""
    start_date = date(2026, 1, 4)
    if target_date < start_date:
        return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def get_malay_date(target_date):
    """Formats date and day into Malay."""
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    day_name = days_ms.get(target_date.strftime("%A"))
    return target_date.strftime(f"%d %B %Y ({day_name})")

def set_cell_background(cell, hex_color):
    """Helper to set background color of a table cell (Hex Color)."""
    # Create XML element for shading
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color))
    # Inject it into the cell's table cell properties (tcPr)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def style_heading_text(p, text, size=12, bold=True, center=True):
    """Standard styling for main headings."""
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(0)

# --- DATA PROCESSING LOGIC ---

def process_excel_data(excel_path):
    """
    Reads, filters, and pivots the uploaded Excel file.
    Expects specific headers: 'Diagnosis', 'Pejabat Kesihatan', 'Notifikasi Status'.
    """
    df = pd.read_excel(excel_path)
    
    # 1. Filter: Exclude "Abai Notifikasi"
    df = df[df['Notifikasi Status'] != 'Abai Notifikasi']
    
    # 2. Pivot: Matrix Table (Penyakit vs PKD)
    matrix_df = pd.crosstab(df['Diagnosis'], df['Pejabat Kesihatan'], rownames=['PENYAKIT'])
    
    # 3. Handle Rows and Columns Grand Totals
    matrix_df['Grand Total'] = matrix_df.sum(axis=1) # Row Sums
    col_sums = matrix_df.sum(axis=0) # Column Sums (Series)
    
    return matrix_df, col_sums

# --- REPORT GENERATION LOGIC ---

def create_full_bwkk_report(output_path, matrix_data=None, col_sums=None):
    doc = Document()
    
    # Page Setup (Narrow margins)
    section = doc.sections[0]
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    
    # Date Calculations
    today = date.today()
    yesterday = today - timedelta(days=1)

    # ==========================================
    # --- PART 1: PAGE HEADER (Original Layout) ---
    # ==========================================
    # Logo Placeholder
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_para.add_run("[JKNS LOGO PLACEHOLDER]").font.size = Pt(14)
    
    # Title Text
    titles = [
        ("KEMENTERIAN KESIHATAN MALAYSIA", 10),
        ("", 0),
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)", 12),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 12),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 12)
    ]
    for text, size in titles:
        if text == "": doc.add_paragraph(); continue
        style_heading_text(doc.add_paragraph(), text, size)

    doc.add_paragraph() # Spacer

    # Green Info Box (Table)
    info_table = doc.add_table(rows=1, cols=2)
    info_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_table.style = 'Table Grid'
    # Fill cells and content
    for i in range(2):
        cell = info_table.cell(0, i)
        set_cell_background(cell, "C6E0B4") # Light Green
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        content = f"Tarikh : {get_malay_date(today)}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(content)
        run.bold = True
        run.font.size = Pt(11)

    doc.add_paragraph() # Spacer

    # ==========================================
    # --- PART 2: SECTION 1.0 (New Image 1) ---
    # ==========================================
    
    # Heading 1.0
    h10 = doc.add_paragraph()
    h10_run = h10.add_run("1.0 Ringkasan Laporan Input Enotifikasi")
    h10_run.bold = True
    h10_run.font.size = Pt(12)
    h10.paragraph_format.space_before = Pt(18)
    
    # Calculate totals for text
    grand_total_str = "0" if col_sums is None else f"{int(col_sums['Grand Total']):,}"
    
    # Section 1.1 Text
    h11 = doc.add_paragraph()
    yesterday_text = get_malay_date(yesterday)
    yesterday_date_only = yesterday.strftime("%d %B %Y")
    
    h11_text = f"1.1 Sejumlah {grand_total_str} input notifikasi telah diterima pada {yesterday_date_only} dengan pecahan mengikut penyakit seperti dalam jadual 1."
    h11_run = h11.add_run(h11_text)
    h11_run.font.size = Pt(11)
    h11.paragraph_format.space_before = Pt(12)
    h11.paragraph_format.space_after = Pt(12)

    # --- PART 3: THE MATRIX TABLE ---
    
    if matrix_data is None:
        # Placeholder if no data provided
        doc.add_paragraph("[Matrix Table Will Be Generated Here Upon Upload]").alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        # Define Color Hex Codes
        COLOR_BLUE_HEAD = "BFDFFF" # Header Blue
        COLOR_BLUE_COL1 = "D9E9FF" # Column 1 light Blue
        COLOR_YELLOW_HEAD = "FFFF00" # Grand Total Yellow Header
        COLOR_YELLOW_COL = "FFFFB3" # Column 1/Row 1 Yellow

        # Prepare Table Structure
        # matrix_data.index = Diseases (PENYAKIT)
        # matrix_data.columns = PKDs + 'Grand Total'
        
        num_penyakit = len(matrix_data)
        num_pkd_cols = len(matrix_data.columns) # Columns from Pivot (includes Grand Total)
        total_rows = num_penyakit + 2 # Header Row + Penyakit Rows + Column Grand Total Row
        total_cols = num_pkd_cols + 1 # Penyakit Col + PKD Cols

        # Create Table
        matrix_table = doc.add_table(rows=total_rows, cols=total_cols)
        matrix_table.style = 'Table Grid'
        matrix_table.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Column widths estimation (Penyakit 2", PKD 0.5")
        widths = [Inches(2.0)] + [Inches(0.55)] * num_pkd_cols
        for cell in matrix_table.columns[0].cells: cell.width = widths[0]
        # Skip the disease column and apply width to subsequent cols
        for col_idx in range(1, total_cols):
            for cell in matrix_table.columns[col_idx].cells:
                cell.width = widths[1]

        # 1. --- ROW 1: HEADERS ---
        # Cell [0,0] = PENYAKIT (Blue)
        cell_dx = matrix_table.cell(0, 0)
        set_cell_background(cell_dx, COLOR_BLUE_HEAD)
        p_dx = cell_dx.paragraphs[0]
        p_dx.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_dx.add_run("PENYAKIT").bold = True

        # Fill PKD Headers (Blue)
        pkd_list = matrix_data.columns.tolist()[:-1] # All columns except 'Grand Total'
        for col_idx, pkd_name in enumerate(pkd_list):
            cell = matrix_table.cell(0, col_idx + 1)
            set_cell_background(cell, COLOR_BLUE_HEAD)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(pkd_name).font.size = Pt(8) # Small font to fit

        # Cell [0, TotalCols-1] = Grand Total (Yellow)
        cell_gt_head = matrix_table.cell(0, total_cols - 1)
        set_cell_background(cell_gt_head, COLOR_YELLOW_HEAD)
        p_gt_head = cell_gt_head.paragraphs[0]
        p_gt_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_gt_head.add_run("Grand Total").bold = True

        # 2. --- PENYAKIT DATA ROWS ---
        penyakit_list = matrix_data.index.tolist()
        for row_idx, penyakit_name in enumerate(penyakit_list):
            table_row_idx = row_idx + 1
            
            # Fill Column 1 (Blue): Penyakit Name
            cell_name = matrix_table.cell(table_row_idx, 0)
            set_cell_background(cell_name, COLOR_BLUE_COL1)
            p_name = cell_name.paragraphs[0]
            p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_name = p_name.add_run(penyakit_name)
            run_name.font.size = Pt(9)
            
            # Fill Data Cells (Center)
            row_data = matrix_data.iloc[row_idx]
            for col_idx, value in enumerate(row_data.tolist()):
                current_table_col = col_idx + 1
                cell_val = matrix_table.cell(table_row_idx, current_table_col)
                # Coloring final Column (Grand Total) Yellow
                if current_table_col == total_cols - 1:
                    set_cell_background(cell_val, COLOR_YELLOW_COL)
                
                p_val = cell_val.paragraphs[0]
                p_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
                val_run = p_val.add_run(str(int(value)))
                val_run.bold = True
                val_run.font.size = Pt(9)

        # 3. --- FINAL ROW: COLUMN SUMS (Grand Total) ---
        table_final_row = total_rows - 1
        
        # Cell [FinalRow, 0] = "Grand Total" (Yellow)
        cell_colsum_head = matrix_table.cell(table_final_row, 0)
        set_cell_background(cell_colsum_head, COLOR_YELLOW_HEAD)
        p_cs_head = cell_colsum_head.paragraphs[0]
        p_cs_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cs_head.add_run("Grand Total").bold = True
        
        # Fill Column Sums (Yellow)
        for col_idx, col_sum_value in enumerate(col_sums.tolist()):
            current_table_col = col_idx + 1
            cell_cs_val = matrix_table.cell(table_final_row, current_table_col)
            set_cell_background(cell_cs_val, COLOR_YELLOW_HEAD)
            p_cs_val = cell_cs_val.paragraphs[0]
            p_cs_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cs_val_run = p_cs_val.add_run(str(int(col_sum_value)))
            cs_val_run.bold = True
            cs_val_run.font.size = Pt(9)

        # --- JADUAL 1 Caption ---
        p_caption = doc.add_paragraph()
        p_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption_run = p_caption.add_run("Jadual 1 : Senarai Input Enotifikasi")
        caption_run.font.size = Pt(12)
        p_caption.paragraph_format.space_before = Pt(12)

    doc.save(output_path)

# --- FLASK ROUTES (Unchanged interface, updated handler) ---

@app.route('/')
def index():
    return '''
    <!doctype html>
    <title>BWKK Report Generator</title>
    <style>
        body { font-family: sans-serif; text-align: center; padding: 50px; background: #f4f4f4; }
        .box { background: white; padding: 30px; border-radius: 8px; display: inline-block; box-shadow: 0 2px 10px rgba(0,0,0,0.1); width: 400px;}
        input[type="file"] { margin: 20px 0; padding: 10px; border: 1px solid #ccc; width: 100%; border-radius: 4px;}
        button { background: #007bff; color: white; border: none; padding: 12px 24px; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%;}
        button:hover { background: #0056b3; }
        .help { color: #666; font-size: 0.9em; text-align: left; margin-top: 15px;}
    </style>
    <div class="box">
        <h2>BWKK Report Generator</h2>
        <form method="post" action="/generate" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx" required><br>
            <button type="submit">Upload Excel & Generate BWKK Report</button>
        </form>
        <div class="help">
        <p><b>Required Excel Structure:</b></p>
        <p>Your uploaded .xlsx file MUST contain exact column headers: "Diagnosis", "Pejabat Kesihatan", and "Notifikasi Status".</p>
        </div>
    </div>
    '''

@app.route('/generate', methods=['POST'])
def generate():
    if 'file' not in request.files: return "No file uploaded", 400
    file = request.files['file']
    if file.filename == '': return "No file selected", 400
    if not file.filename.endswith('.xlsx'): return "Please upload an .xlsx Excel file", 400

    # Save Uploaded Excel
    excel_path = os.path.join(UPLOAD_FOLDER, f"input_{file.filename}")
    file.save(excel_path)

    try:
        # 1. PROCESS DATA
        matrix_data, col_sums = process_excel_data(excel_path)

        # 2. GENERATE FULL REPORT
        report_name = f"BWKK_Report_{date.today()}.docx"
        report_path = os.path.join(UPLOAD_FOLDER, report_name)
        create_full_bwkk_report(report_path, matrix_data, col_sums)

        return send_file(report_path, as_attachment=True)
    except KeyError as e:
        return f"Error: Missing expected column header in Excel: {e}", 400
    except Exception as e:
        return f"An unexpected error occurred during processing: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)
