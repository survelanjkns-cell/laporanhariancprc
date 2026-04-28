# Map for shorter PKD names to prevent vertical wrapping
    pkd_map = {
        'PKD GOMBAK': 'GOMBAK',
        'PKD HULU LANGAT': 'H.LANGAT',
        'PKD HULU SELANGOR': 'H.SEL',
        'PKD KLANG': 'KLANG',
        'PKD KUALA LANGAT': 'K.LANGAT',
        'PKD KUALA SELANGOR': 'K.SEL',
        'PKD PETALING': 'PETALING',
        'PKD SABAK BERNAM': 'S.BERNAM',
        'PKD SEPANG': 'SEPANG'
    }

    for i, pkd in enumerate(TEMPLATE_PKDS):
        cell = h_cells[i+1]
        short_name = pkd_map.get(pkd, pkd)
        run_pkd = cell.paragraphs[0].add_run(short_name)
        
        # Reduced font size slightly to 6.5pt for headers to ensure they fit
        apply_font(run_pkd, 6.5, bold=True) 
        
        set_cell_background(cell, "BFDFFF")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Prevent the cell from wrapping text vertically
        cell.paragraphs[0].paragraph_format.keep_together = True
