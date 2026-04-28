import os
from datetime import date
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- REPORT GENERATION LOGIC ---

def get_epi_week(target_date):
    start_date = date(2026, 1, 4)
    if target_date < start_date:
        return "N/A"
    days_diff = (target_date - start_date).days
    return f"{(days_diff // 7) + 1}/{target_date.year}"

def set_cell_background(cell, fill_color):
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), fill_color))
    cell._tc.get_or_add_tcPr().append(shading_elm)

def create_docx(output_path, logo_path="logo.png"):
    doc = Document()
    
    # Page Setup
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.5)
    
    # Logo
    if os.path.exists(logo_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(logo_path, width=Inches(1.2))
    
    # Header Text
    headers = [
        ("KEMENTERIAN KESIHATAN MALAYSIA", 10),
        ("", 0),
        ("LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEmasan, KRISIS (BWKK)", 12),
        ("PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)", 12),
        ("JABATAN KESIHATAN NEGERI SELANGOR", 12)
    ]
    
    for text, size in headers:
        if text == "":
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    # Green Info Box
    today = date.today()
    days_ms = {"Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu", "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"}
    date_str = today.strftime(f"%d %B %Y ({days_ms.get(today.strftime('%A'))})")
    
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.style = 'Table Grid'
    
    for i in range(2):
        cell = table.cell(0, i)
        set_cell_background(cell, "C6E0B4")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        content = f"Tarikh : {date_str}\n(Sehingga jam 10.00 pagi)" if i == 0 else f"\nMinggu Epidemiologi : {get_epi_week(today)}"
        run = p.add_run(content)
        run.bold = True
        run.font.size = Pt(11)

    doc.save(output_path)

# --- FLASK ROUTES ---

@app.route('/')
def index():
    return '''
    <!doctype html>
    <title>BWKK Report Generator</title>
    <style>
        body { font-family: sans-serif; text-align: center; padding: 50px; background: #f4f4f4; }
        .box { background: white; padding: 30px; border-radius: 8px; display: inline-block; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        input { margin: 20px 0; }
        button { background: #28a745; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }
    </style>
    <div class="box">
        <h2>BWKK Report Generator</h2>
        <form method="post" action="/generate" enctype="multipart/form-data">
            <input type="file" name="file" required><br>
            <button type="submit">Upload & Generate Report</button>
        </form>
    </div>
    '''

@app.route('/generate', methods=['POST'])
def generate():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    # Save the uploaded file (we'll use it later to parse data)
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Generate the report
    report_name = "Laporan_BWKK_Generated.docx"
    report_path = os.path.join(UPLOAD_FOLDER, report_name)
    create_docx(report_path)

    return send_file(report_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
