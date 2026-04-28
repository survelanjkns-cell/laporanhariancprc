import os
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def get_epi_week(target_date):
    """Calculates Epi Week based on Start Date: Jan 4, 2026"""
    start_date = date(2026, 1, 4)
    if target_date < start_date:
        return "N/A"
    days_diff = (target_date - start_date).days
    week_num = (days_diff // 7) + 1
    return f"{week_num}/{target_date.year}"

def set_cell_background(cell, fill_color):
    """Helper to set background color of a table cell (Hex Color)"""
    shading_elm_1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), fill_color))
    cell._tc.get_or_add_tcPr().append(shading_elm_1)

def generate_word_report(output_filename, logo_path="logo.png"):
    doc = Document()

    # --- 1. Set Margins ---
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # --- 2. Logo Section ---
    if os.path.exists(logo_path):
        # Adding logo to a centered paragraph
        logo_para = doc.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_logo = logo_para.add_run()
        run_logo.add_picture(logo_path, width=Inches(1.5)) 
    else:
        print(f"Warning: {logo_path} not found. Skipping logo.")

    # --- 3. Title Section ---
    titles = [
        "KEMENTERIAN KESIHATAN MALAYSIA", # Added sub-logo text
        "", # Spacer
        "LAPORAN HARIAN KEJADIAN BENCANA, WABAK, KECEMASAN, KRISIS (BWKK)",
        "PUSAT KESIAPSIAGAAN DAN TINDAKCEPAT KRISIS (CPRC)",
        "JABATAN KESIHATAN NEGERI SELANGOR"
    ]

    for title in titles:
        if title == "":
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(title)
        run.bold = True
        # Slightly smaller font for Ministry name as per standard JKN layouts
        size = 10 if title == "KEMENTERIAN KESIHATAN MALAYSIA" else 12
        run.font.size = Pt(size)
        run.font.name = 'Arial'
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph() # Spacer

    # --- 4. Date & Epi Week Box (Table) ---
    today = date.today()
    days_ms = {
        "Monday": "Isnin", "Tuesday": "Selasa", "Wednesday": "Rabu",
        "Thursday": "Khamis", "Friday": "Jumaat", "Saturday": "Sabtu", "Sunday": "Ahad"
    }
    day_name = days_ms.get(today.strftime("%A"))
    date_str = today.strftime(f"%d %B %Y ({day_name})")
    epi_week = get_epi_week(today)

    # Create a 1x2 table
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.style = 'Table Grid'
    
    # Set background color (Light Green - C6E0B4)
    bg_color = "C6E0B4"
    
    # Cell 1: Date and Fixed Time
    cell_left = table.cell(0, 0)
    set_cell_background(cell_left, bg_color)
    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p_left.add_run(f"Tarikh : {date_str}\n(Sehingga jam 10.00 pagi)")
    run1.bold = True
    run1.font.size = Pt(11)

    # Cell 2: Epi Week
    cell_right = table.cell(0, 1)
    set_cell_background(cell_right, bg_color)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p_right.add_run(f"\nMinggu Epidemiologi : {epi_week}")
    run2.bold = True
    run2.font.size = Pt(11)

    # Save
    doc.save(output_filename)
    print(f"Word report saved as {output_filename}")

if __name__ == "__main__":
    generate_word_report("Laporan_BWKK_with_Logo.docx")
