from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import io
import re

cf_bp = Blueprint('cf_bp', __name__, template_folder='templates')

MONTH_MAP = {
    1: "June-2026", 2: "July-2026", 3: "August-2026", 4: "September-2026",
    5: "October-2026", 6: "November-2026", 7: "December-2026", 8: "January-2027",
    9: "February-2027", 10: "March-2027", 11: "April-2027", 12: "May-2027"
}

CF_BANK_MAPPING = {
    "NARESH": ["BANK OF BARODA", "BANK OF INDIA"],
    "SATHISH": ["KDCCB", "PUNJAB NATIONAL BANK"],
    "NAVEEN": ["CANARA BANK", "CENTRAL BANK OF INDIA"],
    "SRIKANTH": ["12903_STATE BANK OF INDIA", "20140_STATE BANK OF INDIA", "20366_STATE BANK OF INDIA"],
    "DEVADAS": ["SIRCILLA TELANGANA GRAMEENA BANK", "THANGALLAPALLY TELANGANA GRAMEENA BANK"],
    "NAGARJUNA": ["UNION BANK OF INDIA"],
    "KANAKARAJU": ["INDIAN BANK"],
    "KAMAL": ["INDIAN BANK"],
    "MARKANDEYULU": ["BOINPARI UNION BANK", "NIMMAPALLY"]
}

CF_PASSWORDS = {
    "NARESH": "naresh@2026", "SATHISH": "sathish@2026", "NAVEEN": "naveen@2026",
    "SRIKANTH": "srikanth@2026", "DEVADAS": "devadas@2026", "NAGARJUNA": "nagarjuna@2026",
    "KANAKARAJU": "kanakaraju@2026", "KAMAL": "kamal@2026", "MARKANDEYULU": "markandeyulu@2026"
}

def clean_col(col_name):
    """కాలమ్ పేరులోని %, స్పేసులు, స్పెషల్ క్యారెక్టర్లను క్లీన్ చేస్తుంది"""
    return re.sub(r'[^a-zA-Z0-9_]', '_', str(col_name)).lower().strip('_')

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

@cf_bp.route('/cf/dashboard')
def cf_dashboard():
    if not session.get('cf_logged_in'):
        return redirect(url_for('cf_bp.cf_login'))
    cf_name = session.get('cf_name')
    return render_template('cf_dashboard.html', cf_name=cf_name, months=MONTH_MAP)

@cf_bp.route('/cf/logout', methods=['POST'])
def cf_logout():
    session.pop('cf_logged_in', None)
    session.pop('cf_name', None)
    return jsonify({"status": "success", "message": "లాగౌట్ అయ్యారు."})


# 🚀 1. GSheet నుండి Supabase కి వన్-టైమ్ మైగ్రేషన్
@cf_bp.route('/cf/sync-to-supabase', methods=['GET'])
def sync_to_supabase():
    try:
        from app import get_sheets_client, SPREADSHEET_NAME, supabase
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet('NEW THRIFT DATA')
        records = worksheet.get_all_records()
        
        for row in records:
            clean_row = {clean_col(k): (v if pd.notna(v) else "") for k, v in row.items()}
            # aadharnumber ప్రాథమిక కీ గా ఉండేలా చూసుకోవడం
            if 'aadharnumber' not in clean_row or not clean_row['aadharnumber']:
                clean_row['aadharnumber'] = str(row.get('Aadhar Number', '')).strip()
                
            if clean_row['aadharnumber']:
                supabase.table('thrift_beneficiaries').upsert(clean_row, on_conflict='aadharnumber').execute()
            
        return "సక్సెస్! గూగుల్ షీట్ డేటా అంతా Supabase లోకి సింక్ అయింది!"
    except Exception as e:
        return f"లోపం: {str(e)}"


# 🚀 2. Supabase నుండి ఆధార్ ద్వారా సెర్చ్ (డ్రాప్‌డౌన్‌లు మరియు ఒరిజినల్ లాజిక్‌తో)
@cf_bp.route('/cf/search-beneficiary', methods=['POST'])
def cf_search_beneficiary():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
        
    try:
        from app import supabase
        cf_name = session.get('cf_name')
        data = request.get_json() if request.is_json else request.form
        aadhar_number = str(data.get('aadhar_number', '')).strip()
        selected_file_name = str(data.get('file_name', '')).strip()
        
        if not aadhar_number or not selected_file_name:
            return jsonify({"status": "error", "message": "దయచేసి ఆధార్ నంబర్ మరియు ఫైల్ పేరు ఎంటర్ చేయండి!"})

        # ఈమెయిల్ పొందడం
        email_id = ""
        try:
            email_res = supabase.table('bank_email_ids').select("*").execute()
            for row in (email_res.data or []):
                f_val = str(row.get('file_name') or row.get('FILE_NAME') or row.get('FILE NAME') or '').strip().upper()
                if f_val == selected_file_name.upper():
                    email_id = str(row.get('official_email') or row.get('OFFICIAL_EMAIL') or '').strip()
                    break
        except Exception:
            pass

        res_new = supabase.table('thrift_beneficiaries').select("*").execute()
        records_new = res_new.data if res_new.data else []
        
        matched_record = None
        source_sheet = None
        
        for row in records_new:
            row_aadhar = str(row.get('aadharnumber') or row.get('Aadhar Number') or '').strip()
            if row_aadhar == aadhar_number:
                matched_record = row
                source_sheet = 'new'
                break

        if not matched_record:
            try:
                res_old = supabase.table('old_thrift_data').select("*").execute()
                for row in (res_old.data or []):
                    row_aadhar = str(row.get('aadharnumber') or row.get('Aadhar Number') or '').strip()
                    if row_aadhar == aadhar_number:
                        matched_record = row
                        source_sheet = 'old'
                        break
            except Exception:
                pass

        columns = list(records_new[0].keys()) if records_new else [
            'sno', 'bank_sno', 'name_of_benificiary', 'father_name', 'date_of_birth', 'age', 
            'aadharnumber', 'caste', 'profession', 'h_no_residentional_h_no', 'village_name', 
            'mandal_name', 'district_name', 'contact_number', 'unit_owner', 'h_no_unit_h_no', 
            'area', 'village', 'relation', 'bank_account_number', 'bank_name', 'rd_1', 
            'monthly_wages', '8_of_monthly_wages', 'name_of_cf', 'email_id', 'file_name', 'bank_sent'
        ]

        # SNO గణన
        max_bank_sno = 0
        for row in records_new:
            f_name = str(row.get('file_name') or row.get('FILE NAME') or '').strip().upper()
            if f_name == selected_file_name.upper():
                b_sno = str(row.get('bank_sno') or row.get('BANK_SNO', '0')).strip()
                if b_sno.isdigit():
                    max_bank_sno = max(max_bank_sno, int(b_sno))
        next_bank_sno = max_bank_sno + 1

        # డ్రాప్‌డౌన్ ఆప్షన్ల సమాచారం
        unique_professions = sorted(list(set([str(r.get('profession', '')).strip() for r in records_new if r.get('profession')])))
        unique_villages = sorted(list(set([str(r.get('village', '')).strip() for r in records_new if r.get('village')])))
        unique_areas = sorted(list(set([str(r.get('area', '')).strip() for r in records_new if r.get('area')])))
        unique_unit_owners = sorted(list(set([str(r.get('unit_owner', '')).strip() for r in records_new if r.get('unit_owner')])))
        unique_unit_hns = sorted(list(set([str(r.get('h_no_unit_h_no', '')).strip() for r in records_new if r.get('h_no_unit_h_no')])))
        unique_relations = sorted(list(set([str(r.get('relation', '')).strip() for r in records_new if r.get('relation')])))

        dropdown_data = {
            "professions": unique_professions,
            "villages": unique_villages,
            "areas": unique_areas,
            "unit_owners": unique_unit_owners,
            "unit_hns": unique_unit_hns,
            "relations": unique_relations
        }

        if matched_record and source_sheet == 'new':
            return jsonify({
                "status": "existing",
                "message": "రికార్డు కనుగొనబడింది - ఇది అప్‌డేట్ (Update) అవుతుంది",
                "data": matched_record,
                "columns": columns,
                "dropdown_data": dropdown_data
            })
        else:
            base_data = matched_record.copy() if matched_record else {}
            base_data['aadharnumber'] = aadhar_number
            base_data['bank_sno'] = next_bank_sno
            base_data['file_name'] = selected_file_name
            base_data['bank_name'] = selected_file_name
            base_data['caste'] = 'BC'
            base_data['profession'] = 'POWER LOOM WEAVER'
            base_data['district_name'] = 'RAJANNA SIRCILLA'
            base_data['mandal_name'] = 'SIRCILLA'
            base_data['name_of_cf'] = cf_name
            base_data['email_id'] = email_id
            base_data['bank_sent'] = 'NOT SENT'
            base_data['8_of_monthly_wages'] = 1200
            base_data['monthly_wages'] = 15000

            return jsonify({
                "status": "new",
                "message": "ఇది కొత్త ఎంట్రీ (New Entry) కింద సేవ్ అవుతుంది. వివరాలను పరిశీలించండి.",
                "data": base_data,
                "columns": columns,
                "dropdown_data": dropdown_data
            })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@cf_bp.route('/cf/get-bank-files', methods=['GET'])
def cf_get_bank_files():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
    try:
        from app import supabase
        res = supabase.table('bank_email_ids').select("*").execute()
        file_names = []
        for r in (res.data if res.data else []):
            f_name = r.get('file_name') or r.get('file name') or r.get('FILE NAME') or ''
            if f_name: file_names.append(str(f_name).strip())
        return jsonify({"status": "success", "file_names": list(set(file_names))})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e), "file_names": []})


# 🚀 3. Supabase లో నేరుగా సేవ్ చేయడం
@cf_bp.route('/cf/save-beneficiary', methods=['POST'])
def cf_save_beneficiary():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
    try:
        from app import supabase
        cf_name = session.get('cf_name')
        data = request.get_json() if request.is_json else request.form
        action_type = data.get('action_type')
        form_data = data.get('form_data', {})
        
        clean_form_data = {clean_col(k): v for k, v in form_data.items()}
        
        if action_type == 'insert':
            res = supabase.table('thrift_beneficiaries').select("sno").execute()
            all_snos = [int(r['sno']) for r in (res.data if res.data else []) if str(r.get('sno', '')).isdigit()]
            clean_form_data['sno'] = max(all_snos) + 1 if all_snos else 1
            clean_form_data['name_of_cf'] = cf_name

        supabase.table('thrift_beneficiaries').upsert(clean_form_data, on_conflict='aadharnumber').execute()
        return jsonify({"status": "success", "message": "డేటా Supabase లో విజయవంతంగా సేవ్ చేయబడింది!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# 🚀 4. గూగుల్ షీట్‌కు ఒరిజినల్ కాలమ్ పేర్లతో బల్క్ సింక్
@cf_bp.route('/cf/bulk-sync-to-gsheet', methods=['POST'])
def bulk_sync_to_gsheet():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
    try:
        from app import get_sheets_client, SPREADSHEET_NAME, supabase
        response = supabase.table('thrift_beneficiaries').select("*").execute()
        all_records = response.data
        if not all_records:
            return jsonify({"status": "error", "message": "సింక్ చేయడానికి డేటా లేదు!"})
            
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet('NEW THRIFT DATA')
        
        # గూగుల్ షీట్ ఒరిజినల్ హెడర్లు దెబ్బతినకుండా ఉంచడం
        original_headers = worksheet.row_values(1)
        
        rows = [original_headers]
        for r in all_records:
            row_data = []
            for orig_h in original_headers:
                c_h = clean_col(orig_h)
                row_data.append(r.get(c_h, r.get(orig_h, '')))
            rows.append(row_data)
            
        worksheet.clear()
        worksheet.update(rows)
        return jsonify({"status": "success", "message": "గూగుల్ షీట్‌కు బల్క్ సింక్ విజయవంతంగా పూర్తయింది!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
