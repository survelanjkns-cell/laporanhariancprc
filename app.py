def format_ipu_narrative(data_stesen, tarikh_semalam, epi_week_lepas):
    if not data_stesen:
        return f"Pada {tarikh_semalam}, tiada rekod bacaan IPU yang tertinggi dikesan untuk stesen pemantauan di Selangor. Perincian kes penyakit berkaitan jerebu yang dilaporkan oleh fasiliti sentinel pada tarikh tersebut adalah seperti di Jadual 6.1, manakala Rajah 6.1 hingga Rajah 6.3 menunjukkan tren mingguan konjunktivitis, URTI dan asma berbanding bacaan IPU tertinggi sehingga ME{epi_week_lepas}."
        
    # Susun mengikut bacaan IPU dari tinggi ke rendah
    sorted_stesen = sorted(data_stesen.items(), key=lambda item: item[1], reverse=True)
    
    first_loc, first_val = sorted_stesen[0]
    
    if len(sorted_stesen) > 1:
        other_items = [f"{loc} ({val})" for loc, val in sorted_stesen[1:]]
        if len(other_items) == 1:
            other_str = other_items[0]
        else:
            other_str = ", ".join(other_items[:-1]) + " dan " + other_items[-1]
            
        teks_ipu = f"direkodkan di {first_loc} ({first_val}), diikuti {other_str}"
    else:
        teks_ipu = f"direkodkan di {first_loc} ({first_val})"
        
    teks = f"Pada {tarikh_semalam}, bacaan Indeks Pencemar Udara (IPU) tertinggi di Selangor {teks_ipu}. Perincian kes penyakit berkaitan jerebu yang dilaporkan oleh fasiliti sentinel pada tarikh tersebut adalah seperti di Jadual 6.1, manakala Rajah 6.1 hingga Rajah 6.3 menunjukkan tren mingguan konjunktivitis, URTI dan asma berbanding bacaan IPU tertinggi sehingga ME{epi_week_lepas}."
    
    return teks
