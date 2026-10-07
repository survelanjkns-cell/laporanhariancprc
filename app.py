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
    "Dysentry": 5, "Syphilis": 5, "Typhoid/Paratyphoid": 5,
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
def read_gsheet_csv(sheet_id, gid="0", sheet_name=None, header='default'):
    sheet_id = str(sheet_id).strip()
    
    if sheet_name:
        encoded_name = urllib.parse.quote(sheet_name)
        gviz_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_name}"
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_name}"
    else:
        gid = str(gid).replace("#", "").replace("gid=", "").strip()
        gviz_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}"
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    
    last_err = None
    for url in [gviz_url, export_url]:
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

def format_ipu_narrative(data_stesen, tarikh_hari_ini, tarikh_semalam, epi_week):
    if not data_stesen:
        return f"Pada {tarikh_hari_ini} jam 9.00 pagi, tiada data stesen pemantauan IPU dikesan. Perincian kes penyakit berkaitan jerebu yang dilaporkan oleh fasiliti sentinel pada {tarikh_semalam} adalah seperti di Jadual 6.1, manakala Rajah 6.1 hingga Rajah 6.3 menunjukkan tren mingguan konjunktivitis, URTI dan asma berbanding bacaan IPU tertinggi sehingga ME{epi_week}."
        
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
    
    teks = f"Pada {tarikh_hari_ini} jam 9.00 pagi, kesemua {jumlah_stesen} stesen pemantauan di Selangor merekodkan IPU pada {status_str}. Perincian kes penyakit berkaitan jerebu yang dilaporkan oleh fasiliti sentinel pada {tarikh_semalam} adalah seperti di Jadual 6.1, manakala Rajah 6.1 hingga Rajah 6.3 menunjukkan tren mingguan konjunktivitis, URTI dan asma berbanding bacaan IPU tertinggi sehingga ME{epi_week}."
    
    return teks

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
