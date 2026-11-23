from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import datetime
import re
import os
import json
import base64
import requests

cf_bp = Blueprint('cf_bp', __name__, template_folder='templates')

# నెలల పేర్ల మ్యాపింగ్
MONTH_MAP = {
    1: "June-2026", 2: "July-2026", 3: "August-2026", 4: "September-2026",
    5: "October-2026", 6: "November-2026", 7: "December-2026", 8: "January-2027",
    9: "February-2027", 10: "March-2027", 11: "April-2027", 12: "May-2027"
}

# 📋 CF వారీగా కేటాయించిన బ్యాంకుల మ్యాపింగ్
CF_BANK_MAPPING = {
    "NARESH": ["BANK OF BARODA", "BANK OF INDIA"],
    "SATHISH": ["KDCCB", "PUNJAB NATIONAL BANK"],
    "NAVEEN": ["CANARA BANK", "CENTRAL BANK OF INDIA"],
    "SRIKANTH": ["12903_STATE BANK OF INDIA", "20140_STATE BANK OF INDIA", "20366_STATE BANK OF INDIA"],
    "DEVADAS": ["SIRCILLA TELANGANA GRAMEENA BANK", "THANGALLAPALLY TELANGANA GRAMEENA BANK"],
    "NAGARJUNA": ["UNION BANK OF INDIA"],
    "KANAKARAJU": ["INDIAN BANK"],
    "KAMALAKAR": ["INDIAN BANK"],
    "MARKANDEYULU": ["BOINPARI UNION BANK", "NIMMAPALLY"]
}

# 🏦 FILE NAME ఆధారంగా డిఫాల్ట్ / ప్రైమరీ IFSC CODE మ్యాపింగ్
FILE_NAME_IFSC_MAPPING = {
    "12903_STATE BANK OF INDIA": "SBIN0012903",
    "20140_STATE BANK OF INDIA": "SBIN0020140",
    "20366_STATE BANK OF INDIA": "SBIN0020366",
    "BANK OF BARODA": "BARB0SIRCIL",
    "BANK OF INDIA": "BKID0005745",
    "BOINPALLI UNION BANK": "UBIN0807486",
    "BOINPALLE UNION BANK": "UBIN0807486",
    "CANARA BANK": "CNRB0004170",
    "CENTRAL BANK OF INDIA": "CBIN0285183",
    "INDIAN BANK": "IDIB000S289",
    "KDCCB": "TSAB0020033",
    "NIMMAPALLY": "TGRB0000384",
    "PUNJAB NATIONAL BANK": "PUNB0589900",
    "SIRCILLA TELANGANA GRAMEENA BANK": "TGRB0000215",
    "THANGALLAPALLY TELANGANA GRAMEENA BANK": "TGRB0000086",
    "UNION BANK OF INDIA": "UBIN0805289"
}

# 🔑 CF పాస్‌వర్డ్స్ మ్యాపింగ్
CF_PASSWORDS = {
    "NARESH": "naresh@2026",
    "SATHISH": "sathish@2026",
    "NAVEEN": "naveen@2026",
    "SRIKANTH": "srikanth@2026",
    "DEVADAS": "devadas@2026",
    "NAGARJUNA": "nagarjuna@2026",
    "KANAKARAJU": "kanakaraju@2026",
    "KAMALAKAR": "kamalakar@2026",
    "MARKANDEYULU": "markandeyulu@2026"
}

# 🏛️ సిరిసిల్ల మున్సిపాలిటీ వార్డుల వారీగా CF మ్యాపింగ్
WARD_CF_MAPPING = {
    20: "DEVADAS",
    17: "KANAKARAJU", 18: "KANAKARAJU", 19: "KANAKARAJU", 25: "KANAKARAJU", 26: "KANAKARAJU",
    4: "SRIKANTH", 15: "SRIKANTH", 16: "SRIKANTH",
    21: "NAGARJUNA", 22: "NAGARJUNA", 23: "NAGARJUNA", 24: "NAGARJUNA",
    1: "SATHISH", 2: "SATHISH", 27: "SATHISH", 28: "SATHISH", 32: "SATHISH", 33: "SATHISH",
    29: "KAMALAKAR", 30: "KAMALAKAR", 31: "KAMALAKAR",
    3: "NAVEEN", 5: "NAVEEN", 6: "NAVEEN", 7: "NAVEEN", 8: "NAVEEN",
    9: "NAVEEN", 10: "NAVEEN", 11: "NAVEEN", 12: "NAVEEN", 13: "NAVEEN", 14: "NAVEEN"
}

# 🔍 సిరిసిల్ల మున్సిపాలిటీ మాస్టర్ డేటా ఆధారంగా కాన్ఫ్లిక్ట్‌లను కూడా పూర్తిగా నివారించే ఖచ్చితమైన రేంజ్ మ్యాపింగ్
MUNICIPALITY_RANGES = [
    # Ward 1
    ("2-5", 1, 11, 1), ("2-5", 95, 149, 1), ("2-6", 1, 142, 1), ("2-7", 1, 140, 1), ("2-8", 1, 65, 1), ("2-4", 84, 139, 1),
    # Ward 2
    ("3-1", 1, 140, 2), ("3-2", 1, 139, 2), ("3-3", 1, 140, 2), ("3-4", 1, 119, 2), ("3-5", 1, 31, 2),
    # Ward 3
    ("3-5", 32, 122, 3), ("3-6", 1, 164, 3), ("3-7", 1, 98, 3),
    # Ward 4
    ("3-7", 99, 136, 4), ("3-8", 1, 100, 4), ("4-1", 1, 136, 4), ("4-2", 1, 130, 4),
    # Ward 5
    ("4-3", 1, 132, 5), ("4-4", 1, 132, 5), ("4-5", 1, 74, 5),
    # Ward 6
    ("4-5", 75, 150, 6), ("4-6", 1, 164, 6), ("4-7", 1, 123, 6), ("4-8", 1, 69, 6),
    # Ward 7
    ("5-1", 85, 146, 7), ("5-2", 1, 131, 7), ("5-3", 1, 125, 7), ("5-4", 1, 103, 7), ("5-5", 1, 139, 7), ("5-6", 1, 131, 7), ("5-7", 1, 138, 7),
    # Ward 8
    ("4-8", 70, 115, 8), ("5-1", 1, 84, 8), ("5-8", 1, 149, 8), ("6-1", 1, 161, 8), ("6-2", 1, 132, 8),
    # Ward 9
    ("6-3", 1, 117, 9), ("6-4", 1, 129, 9), ("6-5", 1, 97, 9),
    # Ward 10
    ("6-5", 98, 139, 10), ("6-6", 1, 129, 10), ("6-7", 1, 91, 10), ("6-8", 1, 95, 10),
    # Ward 11
    ("7-1", 1, 139, 11), ("7-2", 1, 107, 11),
    # Ward 12
    ("7-2", 108, 127, 12), ("7-3", 1, 133, 12), ("7-4", 1, 138, 12), ("7-5", 1, 125, 12),
    # Ward 13
    ("7-6", 1, 131, 13), ("7-7", 1, 146, 13), ("7-8", 1, 140, 13), ("8-1", 1, 150, 13), ("8-2", 1, 135, 13), ("8-3", 1, 69, 13),
    # Ward 14
    ("8-3", 70, 118, 14), ("8-4", 1, 129, 14), ("8-5", 1, 137, 14), ("8-8", 63, 99, 14),
    # Ward 15
    ("8-8", 11, 62, 15), ("9-3", 1, 125, 15), ("9-4", 1, 119, 15),
    # Ward 16
    ("8-6", 1, 130, 16), ("8-7", 1, 145, 16), ("8-8", 1, 10, 16), ("8-8", 100, 101, 16),
    # Ward 17
    ("9-1", 1, 150, 17), ("9-2", 1, 159, 17),
    # Ward 18
    ("9-5", 1, 75, 18), ("9-6", 1, 63, 18), ("10-1", 79, 144, 18), ("10-2", 1, 92, 18), ("10-3", 100, 135, 18), ("10-4", 1, 117, 18),
    # Ward 19
    ("10-3", 1, 99, 19), ("10-5", 1, 120, 19), ("10-6", 1, 141, 19), ("10-7", 1, 34, 19), ("9-8", 102, 139, 19),
    # Ward 20
    ("9-6", 64, 125, 20), ("9-7", 1, 107, 20), ("10-1", 1, 78, 20), ("9-8", 1, 101, 20), ("9-8", 140, 166, 20),
    # Ward 21
    ("10-7", 35, 133, 21), ("10-8", 1, 142, 21), ("11-1", 1, 156, 21),
    # Ward 22
    ("12-2", 32, 147, 22), ("12-3", 1, 111, 22), ("12-4", 1, 73, 22), ("12-5", 1, 112, 22),
    # Ward 23
    ("11-2", 1, 133, 23), ("11-3", 1, 148, 23), ("11-4", 1, 133, 23),
    # Ward 24
    ("11-8", 33, 132, 24), ("12-1", 1, 152, 24), ("12-5", 113, 142, 24), ("12-6", 1, 54, 24), ("12-2", 1, 31, 24), ("12-6", 95, 132, 24),
    # Ward 25
    ("11-5", 1, 125, 25), ("11-6", 1, 127, 25), ("11-7", 1, 162, 25), ("11-8", 1, 32, 25),
    # Ward 26
    ("12-6", 55, 94, 26), ("12-6", 133, 149, 26), ("12-7", 1, 149, 26), ("12-8", 1, 64, 26),
    # Ward 27
    ("12-8", 65, 170, 27), ("12-9", 111, 169, 27), ("12-10", 1, 169, 27),
    # Ward 28
    ("12-11", 1, 181, 28), ("1-1", 1, 82, 28),
    # Ward 29
    ("1-1", 83, 120, 29), ("1-2", 1, 146, 29), ("1-3", 1, 74, 29), ("1-6", 1, 14, 29), ("1-6", 19, 22, 29),
    # Ward 30
    ("1-5", 19, 134, 30), ("1-6", 23, 85, 30), ("1-7", 1, 148, 30), ("1-8", 1, 25, 30),
    # Ward 31
    ("1-3", 75, 161, 31), ("1-4", 1, 169, 31), ("1-6", 15, 18, 31), ("1-5", 1, 18, 31),
    # Ward 32
    ("2-1", 93, 122, 32), ("2-5", 12, 94, 32), ("2-3", 1, 18, 32), ("2-3", 122, 137, 32), ("2-2", 1, 128, 32), ("2-4", 1, 83, 32), ("12-9", 1, 110, 32),
    # Ward 33
    ("1-8", 26, 119, 33), ("2-1", 1, 76, 33), ("2-3", 19, 121, 33), ("2-1", 77, 92, 33)
]

# 🌍 ఏరియా / విలేజ్ వారీగా CF మ్యాపింగ్
AREA_CF_MAPPING = {
    "NARESH": ["THANGALLA PALLY","THANGALLAPALLY","RAMNAGAR","THANGALLAPALLI","PADMANAGAR","LINGAMWADA","OPPOSITE RADHESHYAM SATH SANGH LINGAM WADA","PADMNAGAR","PADMA NAGAR","TEXTILEPARK","TEXTILE PARK"],
    "KANAKARAJU": ["THADUR", "MANDAPALLI"],
    "MARKANDEYULU": [
        "ANANTHAPALLE", "BHOOPATHI NAGAR", "BOINPALLE", "BURUGUPALLY",
        "CHANDRAMPET", "DUNDRAPALLY", "JYOTHI NAGAR", "MUSTIPALLI",
        "NIMMAPALLI", "RAGUDU", "RAJEEV NAGAR", "TETTEKUNTA",
        "THADAGONDA", "VELJIPURAM"
    ]
}

# 🧹 H.No లోని "H.NO.", "H.NO", "HNO", "D.NO", స్పేస్‌లు మొదలైన వాటిని క్లీన్ చేసే హెల్పర్ ఫంక్షన్
def clean_hno_for_matching(hno_val):
    if hno_val is None:
        return ""
    s = str(hno_val).strip().upper()
    if not s or s in ["NAN", "NONE", "NULL", "-"]:
        return ""

    # ముందుగా ఉండే H.NO., H.NO, HNO, D.NO., DOOR NO, HOUSE NO, NO. వంటి ప్రిఫిక్స్‌లను తొలగించడం
    s = re.sub(
        r'^\s*(?:H\.?\s*N\.?\s*O\.?|HOUSE\s*NO\.?|DOOR\s*NO\.?|D\.?\s*NO\.?|UNIT\s*H\.?\s*NO\.?|RESIDENTIAL\s*H\.?\s*NO\.?|NO\.?)\s*[:\-\.]*\s*',
        '',
        s,
        flags=re.IGNORECASE
    )

    # మధ్యలో పొరపాటున "H.NO" లేదా "H.NO." చేరినా తొలగించడం
    s = re.sub(r'\bH\.?\s*N\.?\s*O\.?\b', '', s, flags=re.IGNORECASE)

    # హైఫన్లు, స్లాష్‌లు, స్పేస్‌లను క్రమబద్ధీకరించడం
    s = re.sub(r'\s*-\s*', '-', s)
    s = re.sub(r'\s*/\s*', '/', s)
    s = re.sub(r'\s+', ' ', s).strip(' -/')
    return s

# 🧹 ఏరియా / విలేజ్ పేరును క్లీన్ చేసే హెల్పర్ ఫంక్షన్
def clean_area_for_matching(area_val):
    if area_val is None:
        return ""
    s = str(area_val).strip().upper()
    if not s or s in ["NAN", "NONE", "NULL", "-"]:
        return ""

    s = re.sub(r'\bH\.?\s*N\.?\s*O\.?\b', ' ', s, flags=re.IGNORECASE)
    s = re.sub(r'[^A-Z0-9\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# 🔎 క్లీన్ చేసిన H.No నుండి (block_prefix, door_num) వేరు చేసే హెల్పర్ ఫంక్షన్
def extract_block_and_door(cleaned_hno):
    if not cleaned_hno:
        return None, None

    # ఉదాహరణ: "3-6-45", "10-3-110A", "1-4-120/2" -> block_prefix="3-6", door_num=45
    m = re.search(r'\b(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*(\d+)', cleaned_hno)
    if m:
        block_prefix = f"{int(m.group(1))}-{int(m.group(2))}"
        door_num = int(m.group(3))
        return block_prefix, door_num

    # కేవలం 2 భాగాలు ఉంటే (ఉదా: "3-6")
    m2 = re.search(r'\b(\d{1,2})\s*-\s*(\d{1,2})\b', cleaned_hno)
    if m2:
        block_prefix = f"{int(m2.group(1))}-{int(m2.group(2))}"
        return block_prefix, None

    return None, None

# 🏦 FILE NAME ఆధారంగా IFSC కోడ్ కనుగొనే హెల్పర్ ఫంక్షన్
def get_ifsc_for_file_name(file_name, records_list=None):
    clean_file = re.sub(r'\s+', ' ', str(file_name or '').strip().upper())
    if not clean_file or clean_file in ['NAN', 'NONE']:
        return ""

    if records_list:
        ifsc_counts = {}
        for row in records_list:
            r_file = re.sub(r'\s+', ' ', str(row.get('FILE NAME', '') or '').strip().upper())
            r_ifsc = str(row.get('IFSC code', '') or row.get('IFSC CODE', '') or '').strip().upper()
            if r_file == clean_file and re.match(r'^[A-Z]{4}0[A-Z0-9]{6}$', r_ifsc):
                ifsc_counts[r_ifsc] = ifsc_counts.get(r_ifsc, 0) + 1
        if ifsc_counts:
            return max(ifsc_counts, key=ifsc_counts.get)

    if clean_file in FILE_NAME_IFSC_MAPPING:
        return FILE_NAME_IFSC_MAPPING[clean_file]

    normalized_variant = clean_file.replace("BOINPALLE", "BOINPALLI")
    if normalized_variant in FILE_NAME_IFSC_MAPPING:
        return FILE_NAME_IFSC_MAPPING[normalized_variant]

    return ""

# 🔐 ఆధార్ నంబర్‌ను క్లీన్ చేసే హెల్పర్ ఫంక్షన్
def normalize_aadhar(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r'\D', '', s)

# 📍 షీట్‌లో ఖచ్చితమైన రో నంబర్‌ను (Aadhaar ఆధారంగా) కనుగొనే హెల్పర్ ఫంక్షన్
def find_row_by_aadhar(worksheet, aadhar_col_idx, target_aadhar_clean):
    if aadhar_col_idx == -1 or not target_aadhar_clean:
        return None, []
    all_aadhars = worksheet.col_values(aadhar_col_idx + 1)
    for idx in range(1, len(all_aadhars)):
        cell_clean = normalize_aadhar(all_aadhars[idx])
        if cell_clean and cell_clean == target_aadhar_clean:
            sheet_row_num = idx + 1
            return sheet_row_num, all_aadhars
    return None, all_aadhars

# 📏 షీట్‌లో చివరి నిజమైన డేటా రో నంబర్ కనుగొనే హెల్పర్ ఫంక్షన్
def get_last_non_empty_row(worksheet):
    all_values = worksheet.get_all_values()
    last_row = 1
    for r_idx, row_vals in enumerate(all_values, start=1):
        if any(str(cell).strip() != "" for cell in row_vals):
            last_row = r_idx
    return last_row

# 🛟 Update చేసే ముందు పాత డేటాను Backup షీట్‌లో సేవ్ చేసే హెల్పర్ ఫంక్షన్
def backup_row_before_update(spreadsheet, headers, old_row_values, target_row_num, cf_name, aadhar_clean):
    backup_sheet_name = "Update_Backup"
    try:
        ws_backup = spreadsheet.worksheet(backup_sheet_name)
    except Exception:
        ws_backup = spreadsheet.add_worksheet(title=backup_sheet_name, rows=1000, cols=max(len(headers) + 4, 15))
        backup_headers = ["Backup_Timestamp", "Updated_By_CF", "Verified_Aadhaar", "Source_Row_Snapshot"] + headers
        ws_backup.append_row(backup_headers)

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    padded_old = list(old_row_values[:len(headers)])
    while len(padded_old) < len(headers):
        padded_old.append("")

    backup_row = [timestamp, cf_name, aadhar_clean, f"Row_{target_row_num}"] + padded_old
    ws_backup.append_row(backup_row)

# 🧠 డేటాను క్లీన్ చేసి, Unit H.No (లేదా Residential H.No) & Area ఆధారంగా ఖచ్చితమైన CF ని ఎంచుకునే ఫంక్షన్
def analyze_cf_name(unit_hno, unit_area, residential_hno=""):
    cleaned_area = clean_area_for_matching(unit_area)

    # ముందుగా Unit H.No క్లీన్ చేసి చెక్ చేస్తాం; అది ఖాళీగా ఉంటే Residential H.No క్లీన్ చేసి చెక్ చేస్తాం
    hno_candidates = [
        clean_hno_for_matching(unit_hno),
        clean_hno_for_matching(residential_hno)
    ]

    for cleaned_hno in hno_candidates:
        if not cleaned_hno:
            continue

        block_prefix, door_num = extract_block_and_door(cleaned_hno)
        if not block_prefix:
            continue

        # 1. బ్లాక్ ప్రిఫిక్స్ మరియు డోర్ నంబర్ రేంజ్ ఆధారంగా ఖచ్చితమైన వార్డ్ CF కనుగొనడం
        if door_num is not None:
            for prefix, min_d, max_d, ward_no in MUNICIPALITY_RANGES:
                if prefix == block_prefix and min_d <= door_num <= max_d:
                    if ward_no in WARD_CF_MAPPING:
                        return WARD_CF_MAPPING[ward_no]

        # 2. డోర్ నంబర్ రేంజ్ దొరకకపోతే బ్లాక్ ప్రిఫిక్స్ ఆధారంగా వార్డ్ CF కనుగొనడం
        for prefix, min_d, max_d, ward_no in MUNICIPALITY_RANGES:
            if prefix == block_prefix:
                if ward_no in WARD_CF_MAPPING:
                    return WARD_CF_MAPPING[ward_no]

    # 3. క్లీన్ చేసిన ఏరియా / విలేజ్ ఆధారంగా CF మ్యాచింగ్
    if cleaned_area:
        for cf, areas in AREA_CF_MAPPING.items():
            for a in areas:
                cleaned_mapped_area = clean_area_for_matching(a)
                if cleaned_mapped_area and cleaned_mapped_area in cleaned_area:
                    return cf

    return "MARKANDEYULU"

# 🛠️ డేట్ ఫార్మాటింగ్ హెల్పర్ ఫంక్షన్
def format_date_val(val):
    if not val or pd.isna(val):
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime('%Y-%m-%d')
    val_str = str(val).strip()
    if not val_str or val_str.upper() in ['NAN', 'NONE', 'NAT']:
        return ""
    if ' ' in val_str:
        val_str = val_str.split(' ')[0]
    try:
        dt = pd.to_datetime(val_str, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')
    except:
        pass
    return val_str

# 1. CF లాగిన్ / నేమ్ సెలెక్షన్ పేజీ
@cf_bp.route('/cf/login', methods=['GET', 'POST'])
def cf_login():
    if request.method == 'POST':
        data = request.get_json() if request.is_json else request.form
        cf_name = str(data.get('cf_name', '')).strip().upper()
        cf_password = str(data.get('cf_password', '')).strip()

        if not cf_name or not cf_password:
            return jsonify({"status": "error", "message": "దయచేసి పేరు మరియు పాస్‌వర్డ్ రెండు ఎంటర్ చేయండి!"})

        if cf_name in CF_PASSWORDS and CF_PASSWORDS[cf_name] == cf_password:
            session['cf_logged_in'] = True
            session['cf_name'] = cf_name
            return jsonify({"status": "success", "message": "లాగిన్ విజయవంతమైంది!"})
        else:
            return jsonify({"status": "error", "message": "తప్పు పేరు లేదా పాస్‌వర్డ్ ఇవ్వబడింది!"})

    return render_template('cf_login.html')

# 2. CF డాష్‌బోర్డ్
@cf_bp.route('/cf/dashboard')
def cf_dashboard():
    if not session.get('cf_logged_in'):
        return redirect(url_for('cf_bp.cf_login'))

    cf_name = session.get('cf_name')
    return render_template('cf_dashboard.html', cf_name=cf_name, months=MONTH_MAP)

# 3. విధానం 1: పేరు / ఏరియా ప్రకారం రిపోర్ట్ డౌన్‌లోడ్ రౌట్
@cf_bp.route('/cf/download-report', methods=['POST'])
def cf_download_report_by_name():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    try:
        from app import get_sheets_client, SPREADSHEET_NAME

        cf_name = session.get('cf_name')
        current_month = int(request.form.get('month_number', 1))

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)

        df_trans = pd.DataFrame(spreadsheet.worksheet('Transactions_Year1').get_all_records())
        df_thrift = pd.DataFrame(spreadsheet.worksheet('NEW THRIFT DATA').get_all_records())

        df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0)
        df_trans['MONTH_NUMBER'] = pd.to_numeric(df_trans['MONTH_NUMBER'], errors='coerce').fillna(0)

        df_trans_month = df_trans[df_trans['MONTH_NUMBER'] == current_month].copy()
        df_trans_month = df_trans_month.drop_duplicates(subset=['MONTH_NUMBER', 'RD ACCOUNT_1 (Beneficiary Contribution)'], keep='first')

        df_thrift['RD 1_Clean'] = df_thrift['RD 1'].astype(str).str.strip()
        df_trans_month['RD1_Clean'] = df_trans_month['RD ACCOUNT_1 (Beneficiary Contribution)'].astype(str).str.strip()

        df_cross = pd.merge(df_thrift, df_trans_month[['RD1_Clean', 'Govt_Share', 'MONTH_NUMBER']], left_on='RD 1_Clean', right_on='RD1_Clean', how='left')
        df_cross['Govt_Share'] = df_cross['Govt_Share'].fillna(0)

        cf_col = 'NAME OF CF' if 'NAME OF CF' in df_cross.columns else 'CF NAME'
        group_df = df_cross[df_cross[cf_col].astype(str).str.strip().str.upper() == cf_name].copy()

        if group_df.empty:
            return "మీ పరిధిలో ఎలాంటి డేటా లభించలేదు.", 404

        return generate_excel_response(group_df, cf_name, current_month, "Name_Based")

    except Exception as e:
        return f"రిపోర్ట్ జనరేట్ చేయడంలో లోపం: {str(e)}"

# 4. విధానం 2: కేటాయించిన బ్యాంకుల (Allotted Bank) ప్రకారం రిపోర్ట్ డౌన్‌లోడ్ రౌట్
@cf_bp.route('/cf/download-report-allotted-bank', methods=['POST'])
def cf_download_report_by_bank():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    try:
        from app import get_sheets_client, SPREADSHEET_NAME

        cf_name = session.get('cf_name')
        current_month = int(request.form.get('month_number', 1))

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)

        df_trans = pd.DataFrame(spreadsheet.worksheet('Transactions_Year1').get_all_records())
        df_thrift = pd.DataFrame(spreadsheet.worksheet('NEW THRIFT DATA').get_all_records())

        df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0)
        df_trans['MONTH_NUMBER'] = pd.to_numeric(df_trans['MONTH_NUMBER'], errors='coerce').fillna(0)

        df_trans_month = df_trans[df_trans['MONTH_NUMBER'] == current_month].copy()
        df_trans_month = df_trans_month.drop_duplicates(subset=['MONTH_NUMBER', 'RD ACCOUNT_1 (Beneficiary Contribution)'], keep='first')

        df_thrift['RD 1_Clean'] = df_thrift['RD 1'].astype(str).str.strip()
        df_trans_month['RD1_Clean'] = df_trans_month['RD ACCOUNT_1 (Beneficiary Contribution)'].astype(str).str.strip()

        bank_col = 'FILE NAME' if 'FILE NAME' in df_thrift.columns else ('File Name' if 'File Name' in df_thrift.columns else 'Bank Name')

        df_cross = pd.merge(
            df_thrift,
            df_trans_month[['RD1_Clean', 'Govt_Share', 'MONTH_NUMBER']],
            left_on='RD 1_Clean',
            right_on='RD1_Clean',
            how='left'
        )
        df_cross['Govt_Share'] = df_cross['Govt_Share'].fillna(0)
        allowed_banks = [b.strip().upper() for b in CF_BANK_MAPPING.get(cf_name, [])]

        def match_bank(row):
            b_name = str(row.get(bank_col, '')).strip().upper()
            if not b_name or b_name == 'NAN' or b_name == 'NONE':
                return False
            for allowed in allowed_banks:
                if allowed == b_name:
                    return True
            return False

        group_df = df_cross[df_cross.apply(match_bank, axis=1)].copy()

        if group_df.empty:
            return f"మీ పరిధిలోని బ్యాంకులకు ({allowed_banks}) సంబంధించిన ఎలాంటి డేటా లభించలేదు.", 404

        return generate_excel_response(group_df, cf_name, current_month, "Bank_Allotted")

    except Exception as e:
        return f"రిపోర్ట్ జనరేట్ చేయడంలో లోపం: {str(e)}"

# కామన్ ఎక్సెల్ జనరేషన్ హెల్పర్ ఫంక్షన్
def generate_excel_response(group_df, cf_name, current_month, report_type):
    account_columns = ['Bank Account Number', 'RD 1']
    for col in account_columns:
        if col in group_df.columns:
            group_df[col] = "'" + group_df[col].astype(str)

    df_active = group_df[group_df['Govt_Share'] > 0].copy()
    df_missing = group_df[group_df['Govt_Share'] == 0].copy()

    wb = openpyxl.Workbook()

    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    blue_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    red_fill = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    green_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')

    # --- Sheet 1: Active ---
    ws1 = wb.active
    ws1.title = 'Sent_To_Bank_List'
    ws1.append(['BANK_SNO', 'Name Of Benificiary', 'Aadhar Number', 'Bank Account Number', 'RD 1', 'Contact Number', 'FILE NAME', 'Govt Share Amount', 'Status', 'NAME OF CF'])
    for col_num in range(1, 11):
        ws1.cell(row=1, column=col_num).font = header_font
        ws1.cell(row=1, column=col_num).fill = blue_fill

    for _, r in df_active.iterrows():
        ws1.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Aadhar Number', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), r.get('Contact Number', ''), r.get('FILE NAME', ''), r.get('Govt_Share', 0),
            f"{MONTH_MAP[current_month]} Govt_Share Released",
            r.get('NAME OF CF', '')
        ])

    # --- Sheet 2: Missing ---
    ws2 = wb.create_sheet(title='Missing_or_Zero_List')
    ws2.append(['BANK_SNO', 'Name Of Benificiary', 'Aadhar Number', 'Bank Account Number', 'RD 1', 'Contact Number', 'FILE NAME', 'Status', 'NAME OF CF'])
    for col_num in range(1, 10):
        ws2.cell(row=1, column=col_num).font = header_font
        ws2.cell(row=1, column=col_num).fill = red_fill

    for _, r in df_missing.iterrows():
        ws2.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Aadhar Number', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), r.get('Contact Number', ''), r.get('FILE NAME', ''),
            f"Zero or missing transactions for {MONTH_MAP[current_month]}",
            r.get('NAME OF CF', '')
        ])

    # --- Sheet 3: All ---
    ws3 = wb.create_sheet(title='All_Beneficiaries_Master')
    ws3.append(['BANK_SNO', 'Name Of Benificiary', 'Aadhar Number', 'Bank Account Number', 'RD 1', 'Contact Number', 'FILE NAME', 'Govt_Share'])
    for col_num in range(1, 9):
        ws3.cell(row=1, column=col_num).font = header_font
        ws3.cell(row=1, column=col_num).fill = green_fill

    for _, r in group_df.iterrows():
        ws3.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Aadhar Number', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), r.get('Contact Number', ''), r.get('FILE NAME', ''), r.get('Govt_Share', 0)
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    response = make_response(output.read())
    response.headers["Content-Disposition"] = f"attachment; filename={cf_name}_{report_type}_Month_{current_month}_Report.xlsx"
    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return response

@cf_bp.route('/cf/logout', methods=['POST'])
def cf_logout():
    session.pop('cf_logged_in', None)
    session.pop('cf_name', None)
    return jsonify({"status": "success", "message": "లాగౌట్ అయ్యారు."})

# 5. ఆధార్ నంబర్ ద్వారా లబ్ధిదారుని సెర్చ్ చేయడం / కొత్త ఎంట్రీ కోసం డేటా సిద్ధం చేయడం
@cf_bp.route('/cf/search-beneficiary', methods=['POST'])
def cf_search_beneficiary():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    try:
        from app import get_sheets_client, SPREADSHEET_NAME

        cf_name = session.get('cf_name')
        data = request.get_json() if request.is_json else request.form
        raw_aadhar = str(data.get('aadhar_number', '')).strip()
        aadhar_clean = normalize_aadhar(raw_aadhar)
        selected_file_name = str(data.get('file_name', '')).strip()

        if not aadhar_clean:
            return jsonify({"status": "error", "message": "దయచేసి సరైన ఆధార్ నంబర్ ఎంటర్ చేయండి!"})

        if not selected_file_name:
            return jsonify({"status": "error", "message": "దయచేసి ముందుగా FILE NAME ఎంచుకోండి!"})

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)

        email_id = ""
        try:
            ws_email = spreadsheet.worksheet('Bank_Email_Ids')
            email_records = ws_email.get_all_records()
            for row in email_records:
                if str(row.get('FILE_NAME') or row.get('FILE NAME', '')).strip().upper() == selected_file_name.upper():
                    email_id = str(row.get('OFFICIAL_EMAIL', '')).strip()
                    break
        except Exception:
            pass

        worksheet_new = spreadsheet.worksheet('NEW THRIFT DATA')
        records_new = worksheet_new.get_all_records()
        columns = worksheet_new.row_values(1)

        auto_ifsc = get_ifsc_for_file_name(selected_file_name, records_new)

        matched_record = None
        source_sheet = None

        for row in records_new:
            if normalize_aadhar(row.get('Aadhar Number', '')) == aadhar_clean:
                matched_record = row
                source_sheet = 'new'
                break

        if not matched_record:
            try:
                worksheet_old = spreadsheet.worksheet('Old Thrift Data')
                records_old = worksheet_old.get_all_records()
                for row in records_old:
                    if normalize_aadhar(row.get('Aadhar Number', '')) == aadhar_clean:
                        matched_record = row
                        source_sheet = 'old'
                        break
            except Exception:
                pass

        if matched_record:
            cleaned_matched = {}
            for k, v in matched_record.items():
                if any(date_keyword in k.lower() for date_keyword in ['dob', 'date of birth', 'date']):
                    cleaned_matched[k] = format_date_val(v)
                else:
                    cleaned_matched[k] = v
            matched_record = cleaned_matched

        max_sno = 0
        max_bank_sno = 0

        for row in records_new:
            s_val = str(row.get('SNO', '0')).strip()
            if s_val.isdigit():
                max_sno = max(max_sno, int(s_val))
            if str(row.get('FILE NAME', '')).strip().upper() == selected_file_name.upper():
                b_sno = str(row.get('BANK_SNO', '0')).strip()
                if b_sno.isdigit():
                    max_bank_sno = max(max_bank_sno, int(b_sno))

        next_sno = max_sno + 1
        next_bank_sno = max_bank_sno + 1

        unique_professions = sorted(list(set([str(r.get('Profession', '')).strip() for r in records_new if r.get('Profession')])))
        unique_villages = sorted(list(set([str(r.get('Village', '')).strip() for r in records_new if r.get('Village')])))
        unique_areas = sorted(list(set([str(r.get('Area', '')).strip() for r in records_new if r.get('Area')])))
        unique_unit_owners = sorted(list(set([str(r.get('Unit Owner', '')).strip() for r in records_new if r.get('Unit Owner')])))
        unique_unit_hns = sorted(list(set([str(r.get('H.NO. Unit H.No', '')).strip() for r in records_new if r.get('H.NO. Unit H.No')])))
        unique_relations = sorted(list(set([str(r.get('Relation', '')).strip() for r in records_new if r.get('Relation')])))

        bank_name = selected_file_name

        if matched_record and source_sheet == 'new':
            current_ifsc = str(matched_record.get('IFSC code', '')).strip()
            if not current_ifsc and auto_ifsc:
                matched_record['IFSC code'] = auto_ifsc

            return jsonify({
                "status": "existing",
                "message": "రికార్డు కనుగొనబడింది - ఆధార్ ఆధారంగా సురక్షితంగా అప్‌డేట్ (Update) అవుతుంది",
                "data": matched_record,
                "auto_ifsc": auto_ifsc,
                "columns": columns,
                "dropdown_data": {
                    "professions": unique_professions,
                    "villages": unique_villages,
                    "areas": unique_areas,
                    "unit_owners": unique_unit_owners,
                    "unit_hns": unique_unit_hns,
                    "relations": unique_relations
                }
            })
        else:
            base_data = matched_record.copy() if matched_record else {}
            base_data['Aadhar Number'] = aadhar_clean
            base_data['SNO'] = next_sno
            base_data['BANK_SNO'] = next_bank_sno
            base_data['FILE NAME'] = selected_file_name
            base_data['Bank Name'] = bank_name
            if auto_ifsc:
                base_data['IFSC code'] = auto_ifsc
            base_data['Caste'] = 'BC'
            base_data['Profession'] = 'POWER LOOM WEAVER'
            base_data['DISTRICT NAME'] = 'RAJANNA SIRCILLA'
            base_data['MANDAL NAME'] = 'SIRCILLA'
            base_data['NAME OF CF'] = cf_name
            base_data['EMAIL ID'] = email_id
            base_data['BANK SENT'] = 'NOT SENT'
            base_data['8% Of Monthly Wages'] = 1200
            base_data['MONTHLY WAGES'] = int(1200 * 100 / 8)

            return jsonify({
                "status": "new",
                "message": "ఇది కొత్త ఎంట్రీ (New Entry) కింద చివరి డేటా రో కింద సురక్షితంగా సేవ్ అవుతుంది.",
                "data": base_data,
                "auto_ifsc": auto_ifsc,
                "columns": columns,
                "dropdown_data": {
                    "professions": unique_professions,
                    "villages": unique_villages,
                    "areas": unique_areas,
                    "unit_owners": unique_unit_owners,
                    "unit_hns": unique_unit_hns,
                    "relations": unique_relations
                }
            })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# బ్రౌజర్ లో డ్రాప్-డౌన్ కోసం Bank_Email_Ids నుండి FILE NAME లను మరియు వాటి IFSC కోడ్‌లను పంపే రౌట్
@cf_bp.route('/cf/get-bank-files', methods=['GET'])
def cf_get_bank_files():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
    try:
        from app import get_sheets_client, SPREADSHEET_NAME
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        ws = spreadsheet.worksheet('Bank_Email_Ids')
        records = ws.get_all_records()

        records_new = []
        try:
            records_new = spreadsheet.worksheet('NEW THRIFT DATA').get_all_records()
        except Exception:
            pass

        file_names = []
        file_ifsc_map = {}

        for r in records:
            f_name = str(r.get('FILE_NAME') or r.get('FILE NAME') or '').strip()
            if f_name:
                file_names.append(f_name)
                file_ifsc_map[f_name] = get_ifsc_for_file_name(f_name, records_new)

        return jsonify({
            "status": "success",
            "file_names": file_names,
            "file_ifsc_map": file_ifsc_map
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e), "file_names": [], "file_ifsc_map": {}})

# 6. row_index వాడకుండా ఆధార్ ఆధారంగా డేటాను సురక్షితంగా Update లేదా Insert చేయడం
@cf_bp.route('/cf/save-beneficiary', methods=['POST'])
def cf_save_beneficiary():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    try:
        from app import get_sheets_client, SPREADSHEET_NAME

        cf_name = session.get('cf_name')
        data = request.get_json() if request.is_json else request.form

        requested_action = str(data.get('action_type', '')).strip().lower()
        form_data = data.get('form_data') or {}

        if not isinstance(form_data, dict):
            return jsonify({"status": "error", "message": "సరైన ఫారమ్ డేటా అందలేదు!"})

        form_aadhar_clean = normalize_aadhar(form_data.get('Aadhar Number', ''))
        if not form_aadhar_clean:
            return jsonify({"status": "error", "message": "Aadhar Number తప్పనిసరి! ఖాళీగా సేవ్ చేయలేము."})

        form_data['Aadhar Number'] = form_aadhar_clean

        # 🔍 యూనిట్ H.No (లేదా రెసిడెన్షియల్ H.No) మరియు ఏరియా డేటాను క్లీన్ చేసి NAME OF CF ని ఖచ్చితంగా అనలైజ్ చేయడం
        unit_hno = form_data.get('H.NO. Unit H.No', '')
        residential_hno = form_data.get('H.NO. Residentional H.No', '')
        unit_area = form_data.get('Area', '')
        analyzed_cf = analyze_cf_name(unit_hno, unit_area, residential_hno)
        form_data['NAME OF CF'] = analyzed_cf

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet('NEW THRIFT DATA')
        headers = worksheet.row_values(1)

        if not headers:
            return jsonify({"status": "error", "message": "షీట్ హెడర్స్ (Row 1) కనుగొనబడలేదు!"})

        aadhar_col_idx = headers.index('Aadhar Number') if 'Aadhar Number' in headers else -1
        if aadhar_col_idx == -1:
            return jsonify({"status": "error", "message": "షీట్‌లో 'Aadhar Number' కాలమ్ కనుగొనబడలేదు!"})

        records_new = worksheet.get_all_records()
        selected_file = str(form_data.get('FILE NAME', '')).strip()
        if selected_file:
            auto_ifsc = get_ifsc_for_file_name(selected_file, records_new)
            if auto_ifsc:
                form_data['IFSC code'] = auto_ifsc

        matched_row_num, all_aadhars = find_row_by_aadhar(worksheet, aadhar_col_idx, form_aadhar_clean)

        if requested_action == 'insert':
            if matched_row_num is not None:
                return jsonify({
                    "status": "error",
                    "message": f"ఈ ఆధార్ నంబర్ ఇప్పటికే Row {matched_row_num} లో ఉంది! కొత్త ఎంట్రీ (Insert) చేయలేము. దయచేసి Update ఎంచుకోండి."
                })
            effective_action = 'insert'
        elif requested_action == 'update':
            if matched_row_num is None:
                return jsonify({
                    "status": "error",
                    "message": "అప్‌డేట్ చేయడానికి ఈ ఆధార్ నంబర్ షీట్‌లో కనుగొనబడలేదు! ఇది కొత్త ఎంట్రీ అయితే Insert ఎంచుకోండి."
                })
            effective_action = 'update'
        else:
            effective_action = 'update' if matched_row_num is not None else 'insert'

        # ==================== UPDATE LOGIC ====================
        if effective_action == 'update':
            live_cell_aadhar = normalize_aadhar(all_aadhars[matched_row_num - 1])
            if live_cell_aadhar != form_aadhar_clean:
                return jsonify({
                    "status": "error",
                    "message": "భద్రతా తనిఖీ విఫలమైంది: ఆధార్ నంబర్ సరిపోలలేదు. ఎవరి డేటా overwrite అవ్వకుండా సేవ్ ఆపివేయబడింది!"
                })

            existing_row_values = worksheet.row_values(matched_row_num)
            backup_row_before_update(spreadsheet, headers, existing_row_values, matched_row_num, cf_name, form_aadhar_clean)

            updated_row = list(existing_row_values[:len(headers)])
            while len(updated_row) < len(headers):
                updated_row.append("")

            for col_name, val in form_data.items():
                if col_name in headers:
                    c_idx = headers.index(col_name)
                    updated_row[c_idx] = "" if val is None else val

            updated_row[aadhar_col_idx] = form_aadhar_clean

            end_col_letter = get_column_letter(len(headers))
            range_notation = f"A{matched_row_num}:{end_col_letter}{matched_row_num}"
            worksheet.update(range_notation, [updated_row], value_input_option='USER_ENTERED')

            return jsonify({
                "status": "success",
                "message": f"డేటా సురక్షితంగా అప్‌డేట్ చేయబడింది (IFSC ఆటో-సెట్ & Backup సేవ్ చేయబడింది)! ఆటో-అనలైజ్ చేయబడిన CF: {analyzed_cf}"
            })

        # ==================== INSERT LOGIC ====================
        elif effective_action == 'insert':
            recheck_row, _ = find_row_by_aadhar(worksheet, aadhar_col_idx, form_aadhar_clean)
            if recheck_row is not None:
                return jsonify({
                    "status": "error",
                    "message": f"ఈ ఆధార్ నంబర్ ఇప్పటికే Row {recheck_row} లో ఉంది! డూప్లికేట్ ఎంట్రీ నివారించబడింది."
                })

            max_sno = 0
            if 'SNO' in headers:
                sno_col_idx = headers.index('SNO') + 1
                all_snos = worksheet.col_values(sno_col_idx)[1:]
                max_sno = max([int(str(s).strip()) for s in all_snos if str(s).strip().isdigit()] or [0])

            form_data['SNO'] = max_sno + 1

            new_row = []
            for header in headers:
                val = form_data.get(header, '')
                new_row.append("" if val is None else val)

            last_non_empty_row = get_last_non_empty_row(worksheet)
            next_row_num = last_non_empty_row + 1

            end_col_letter = get_column_letter(len(headers))
            range_notation = f"A{next_row_num}:{end_col_letter}{next_row_num}"
            worksheet.update(range_notation, [new_row], value_input_option='USER_ENTERED')

            return jsonify({
                "status": "success",
                "message": f"కొత్త రికార్డు చివరి రో లో సురక్షితంగా సేవ్ చేయబడింది (IFSC ఆటో-సెట్)! ఆటో-అనలైజ్ చేయబడిన CF: {analyzed_cf}"
            })

        return jsonify({"status": "error", "message": "సరికైన యాక్షన్ రాలేదు!"})

    except Exception as e:
        return jsonify({"status": "error", "message": f"సేవ్ చేయడంలో లోపం: {str(e)}"})


@cf_bp.route('/cf/ai-ocr', methods=['POST'])
def cf_ai_ocr():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    api_key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if not api_key:
        return jsonify({"status": "fallback", "message": "AI API Key లేదు, లోకల్ OCR వాడుతోంది."})

    try:
        data = request.get_json() or {}
        image_base64 = data.get("image_base64", "")
        scan_type = data.get("scan_type", "Aadhar Number")

        if not image_base64:
            return jsonify({"status": "error", "message": "ఇమేజ్ అందలేదు."})

        model_name = os.environ.get("QWEN_VL_MODEL", "qwen/qwen2.5-vl-72b-instruct:free")

        prompt_text = f"""
You are an expert Indian government document & handwritten form data extractor for Telangana Handlooms & Textiles.
Analyze this image carefully (it may contain printed text, Telugu + English mix, or handwritten notes/forms/Aadhaar/Bank Passbook).
Focus on scan context: "{scan_type}".

Extract all visible fields accurately and return ONLY a valid JSON object with these exact keys (use empty string "" if not visible):
{{
  "number": "primary 12-digit ID OR 10-digit Mobile OR Bank Account Number based on {scan_type} (digits only, no spaces)",
  "id_number": "12-digit beneficiary ID digits only",
  "name": "Beneficiary full name in ENGLISH CAPITAL LETTERS",
  "father_name": "Father or Husband name (C/O, S/O, W/O, D/O) in ENGLISH CAPITAL LETTERS",
  "dob": "Date of birth in DD/MM/YYYY format (or YYYY if only year is given)",
  "house_no": "Residential House Number (e.g. 3-6-45) in UPPERCASE",
  "village": "Village or Town name in ENGLISH CAPITAL LETTERS",
  "mandal": "Mandal name in ENGLISH CAPITAL LETTERS",
  "district": "District name in ENGLISH CAPITAL LETTERS (e.g. RAJANNA SIRCILLA)",
  "mobile": "10-digit mobile number digits only",
  "account_number": "Bank account number digits only",
  "ifsc": "11-character IFSC code in UPPERCASE",
  "unit_hno": "Working loom/unit house number if visible",
  "unit_owner": "Unit owner name in ENGLISH CAPITAL LETTERS if visible",
  "nominee_name": "Nominee name in ENGLISH CAPITAL LETTERS if visible",
  "nominee_id": "12-digit nominee ID digits only if visible",
  "nominee_dob": "Nominee DOB in DD/MM/YYYY or YYYY if visible"
}}
Return ONLY raw JSON. No markdown, no explanation.
"""

        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            json={
                "model": model_name,
                "temperature": 0.0,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt_text},
                            {"type": "image_url", "image_url": {"url": image_base64}}
                        ]
                    }
                ]
            },
            timeout=25
        )

        if resp.status_code != 200:
            return jsonify({"status": "fallback", "message": f"AI API Status {resp.status_code}"})

        res_json = resp.json()
        raw_content = res_json["choices"][0]["message"]["content"].strip()
        raw_content = re.sub(r'^```(?:json)?\s*', '', raw_content)
        raw_content = re.sub(r'\s*```$', '', raw_content).strip()

        extracted = json.loads(raw_content)
        return jsonify({"status": "success", "extracted": extracted})

    except Exception as e:
        return jsonify({"status": "fallback", "message": str(e)})
