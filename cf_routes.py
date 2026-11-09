from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import io

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


# 🚀 1. GSheet నుండి Supabase కి వన్-టైమ్ మైగ్రేషన్ రౌట్
@cf_bp.route('/cf/sync-to-supabase', methods=['GET'])
def sync_to_supabase():
    try:
        from app import get_sheets_client, SPREADSHEET_NAME, supabase
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet('NEW THRIFT DATA')
        records = worksheet.get_all_records()
        
        for row in records:
            clean_row = {k: (v if pd.notna(v) else "") for k, v in row.items()}
            supabase.table('thrift_beneficiaries').upsert(clean_row, on_conflict='Aadhar Number').execute()
            
        return "సక్సెస్! గూగుల్ షీట్ డేటా అంతా Supabase లోకి సింక్ అయింది!"
    except Exception as e:
        return f"లోపం: {str(e)}"


# 🚀 2. Supabase నుండి ఆధార్ నంబర్ ద్వారా సెర్చ్ చేయడం (సూపర్ ఫాస్ట్)
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
        
        if not aadhar_number:
            return jsonify({"status": "error", "message": "దయచేసి ఆధార్ నంబర్ ఎంటర్ చేయండి!"})
        if not selected_file_name:
            return jsonify({"status": "error", "message": "దయచేసి ముందుగా FILE NAME ఎంచుకోండి!"})
            
        # Supabase నుండి ఈమెయిల్ ఐడి పొందడం
        email_id = ""
        try:
            email_res = supabase.table('bank_email_ids').select("OFFICIAL_EMAIL").eq('FILE NAME', selected_file_name).execute()
            if email_res.data:
                email_id = email_res.data[0].get('OFFICIAL_EMAIL', '')
        except Exception:
            pass

        # Supabase నుండి న్యూ త్రిఫ్ట్ డేటా తెచ్చుకోవడం
        res_new = supabase.table('thrift_beneficiaries').select("*").execute()
        records_new = res_new.data if res_new.data else []
        
        # కాలమ్స్ జాబితా కోసం మొదటి రికార్డ్ కీస్ తీసుకోవడం
        columns = list(records_new[0].keys()) if records_new else [
            'SNO', 'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
            'Aadhar Number', 'Caste', 'Profession', 'H.NO. Residentional H.No', 'VILLAGE NAME', 
            'MANDAL NAME', 'DISTRICT NAME', 'Contact Number', 'Unit Owner', 'H.NO. Unit H.No', 
            'Area', 'Village', 'Relation', 'Bank Account Number', 'Bank Name', 'RD 1', 
            'MONTHLY WAGES', '8% Of Monthly Wages', 'NAME OF CF', 'EMAIL ID', 'FILE NAME', 'BANK SENT'
        ]

        matched_record = None
        source_sheet = None
        
        for row in records_new:
            if str(row.get('Aadhar Number', '')).strip() == aadhar_number:
                matched_record = row
                source_sheet = 'new'
                break
                
        if not matched_record:
            try:
                res_old = supabase.table('old_thrift_data').select("*").eq('Aadhar Number', aadhar_number).execute()
                if res_old.data:
                    matched_record = res_old.data[0]
                    source_sheet = 'old'
            except Exception:
                pass

        # BANK_SNO గరిష్ట విలువ + 1 లెక్కించడం
        max_bank_sno = 0
        for row in records_new:
            if str(row.get('FILE NAME', '')).strip().upper() == selected_file_name.upper():
                b_sno = str(row.get('BANK_SNO', '0')).strip()
                if b_sno.isdigit():
                    max_bank_sno = max(max_bank_sno, int(b_sno))
        next_bank_sno = max_bank_sno + 1

        # సజెషన్స్ కోసం యూనిక్ లిస్ట్స్
        unique_professions = sorted(list(set([str(r.get('Profession', '')).strip() for r in records_new if r.get('Profession')])))
        unique_villages = sorted(list(set([str(r.get('Village', '')).strip() for r in records_new if r.get('Village')])))
        unique_areas = sorted(list(set([str(r.get('Area', '')).strip() for r in records_new if r.get('Area')])))
        unique_unit_owners = sorted(list(set([str(r.get('Unit Owner', '')).strip() for r in records_new if r.get('Unit Owner')])))
        unique_unit_hns = sorted(list(set([str(r.get('H.NO. Unit H.No', '')).strip() for r in records_new if r.get('H.NO. Unit H.No')])))
        unique_relations = sorted(list(set([str(r.get('Relation', '')).strip() for r in records_new if r.get('Relation')])))

        if matched_record and source_sheet == 'new':
            return jsonify({
                "status": "existing",
                "message": "రికార్డు కనుగొనబడింది - ఇది అప్‌డేట్ (Update) అవుతుంది",
                "data": matched_record,
                "columns": columns,
                "dropdown_data": {
                    "professions": unique_professions, "villages": unique_villages,
                    "areas": unique_areas, "unit_owners": unique_unit_owners,
                    "unit_hns": unique_unit_hns, "relations": unique_relations
                }
            })
        else:
            base_data = matched_record.copy() if matched_record else {}
            base_data['Aadhar Number'] = aadhar_number
            base_data['BANK_SNO'] = next_bank_sno
            base_data['FILE NAME'] = selected_file_name
            base_data['Bank Name'] = selected_file_name
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
                "message": "ఇది కొత్త ఎంట్రీ (New Entry) కింద సేవ్ అవుతుంది. వివరాలను పరిశీలించండి.",
                "data": base_data,
                "columns": columns,
                "dropdown_data": {
                    "professions": unique_professions, "villages": unique_villages,
                    "areas": unique_areas, "unit_owners": unique_unit_owners,
                    "unit_hns": unique_unit_hns, "relations": unique_relations
                }
            })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@cf_bp.route('/cf/get-bank-files', methods=['GET'])
def cf_get_bank_files():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
    try:
        from app import supabase
        res = supabase.table('bank_email_ids').select("FILE NAME").execute()
        file_names = [str(r.get('FILE NAME', '')).strip() for r in res.data if r.get('FILE NAME')]
        return jsonify({"status": "success", "file_names": file_names})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e), "file_names": []})


# 🚀 3. Supabase లో నేరుగా డేటా సేవ్ / అప్‌డేట్ చేయడం (సెస్కన్లలో అయిపోతుంది)
@cf_bp.route('/cf/save-beneficiary', methods=['POST'])
def cf_save_beneficiary():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
        
    try:
        from app import supabase
        cf_name = session.get('cf_name')
        data = request.get_json() if request.is_json else request.form
        
        action_type = data.get('action_type') # 'update' లేదా 'insert'
        form_data = data.get('form_data')
        
        if action_type == 'insert':
            # SNO ఆటో-ఇంక్రిమెంట్ లెక్కించడం
            res = supabase.table('thrift_beneficiaries').select("SNO").execute()
            all_snos = [int(r['SNO']) for r in res.data if str(r.get('SNO', '')).isdigit()]
            form_data['SNO'] = max(all_snos) + 1 if all_snos else 1
            form_data['NAME OF CF'] = cf_name

        # Supabase UPSERT (ఉంటే అప్‌డేట్ లేదంటే ఇన్సర్ట్ అవుతుంది)
        supabase.table('thrift_beneficiaries').upsert(form_data, on_conflict='Aadhar Number').execute()
        
        return jsonify({"status": "success", "message": "డేటా Supabase లో విజయవంతంగా సేవ్ చేయబడింది!"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# 🚀 4. ఎండింగ్ లో గూగుల్ షీట్‌కి బల్క్ సింక్ చేసే రౌట్
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
        
        headers = list(all_records[0].keys())
        rows = [headers]
        for r in all_records:
            rows.append([r.get(h, '') for h in headers])
            
        worksheet.clear()
        worksheet.update(rows)
        
        return jsonify({"status": "success", "message": "గూగుల్ షీట్‌కు బల్క్ సింక్ విజయవంతంగా పూర్తయింది!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
