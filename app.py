# ... (Kod Jadual 3.1 sebelum ini) ...
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

    doc.add_paragraph()

    # --- TAMBAH RAJAH 3.1 (Tren Denggi) ---
    try:
        # PENTING: Jangan jadikan baris ini sebagai ulasan (comment)
        CHART_IMAGE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTDprYai1uaP1L-JP6kuHRZX18AmDHX0ROEzRE37DaCHMo0cNWUvRa8R-65RZAK7XFWI6pb_-X-jF24/pubchart?oid=1681812411&format=image"
        
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
            print("Gagal memuat turun imej graf.")
    except Exception as e:
        print(f"Ralat semasa memproses Rajah 3.1: {e}")
    # ... (Sambungan kod Bahagian 4.0) ...
