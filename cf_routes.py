from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import io

cf_bp = Blueprint('cf_bp', __name__, template_folder='templates')

# నెలల పేర్ల మ్యాపింగ్
MONTH_MAP = {
    1: "June-2026", 2: "July-2026", 3: "August-2026", 4: "September-2026",
    5: "October-2026", 6: "November-2026", 7: "December-2026", 8: "January-2027",
    9: "February-2027", 10: "March-2027", 11: "April-2027", 12: "May-2027"
}

# 📋 CF వారీగా కేటాయించిన బ్యాంకుల మ్యాపింగ్ (CF Allotment Mapping)
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

# 🔑 CF పాస్‌వర్డ్స్ మ్యాపింగ్ (సెక్యూర్ లాగిన్ కోసం)
CF_PASSWORDS = {
    "NARESH": "naresh@2026",
    "SATHISH": "sathish@2026",
    "NAVEEN": "naveen@2026",
    "SRIKANTH": "srikanth@2026",
    "DEVADAS": "devadas@2026",
    "NAGARJUNA": "nagarjuna@2026",
    "KANAKARAJU": "kanakaraju@2026",
    "KAMAL": "kamal@2026",
    "MARKANDEYULU": "markandeyulu@2026"
}

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
    
   #account_cols = ['Bank Account Number', 'RD 1']  # మీ షీట్‌లోని అకౌంట్ కాలమ్ పేర్లు ఇక్కడ ఇవ్వండి
    #for col in account_cols:
    #    if col in group_df.columns:        # నంబర్ చివర వచ్చే .0 లను తొలగించి, ముందుగా ' చేర్చడానికి
    #        group_df[col] = "'" + group_df[col].astype(str).str.replace(r'\.0$', '', regex=True)

    account_columns = ['Bank Account Number', 'RD 1']
    for col in account_columns:
        if col in group_df.columns:        # ముందుగా వాల్యూస్‌ని స్ట్రింగ్‌గా మార్చి, వాటి ముందు "'" కలుపుతాము
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
    ws1.append(['BANK_SNO', 'Name Of Benificiary', 'Aadhar Number', 'Bank Account Number', 'RD 1','Contact Number','FILE NAME', 'Govt Share Amount', 'Status','NAME OF CF'])
    for col_num in range(1, 11):
        ws1.cell(row=1, column=col_num).font = header_font
        ws1.cell(row=1, column=col_num).fill = blue_fill
        
    for _, r in df_active.iterrows():
        ws1.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Aadhar Number', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), r.get('Contact Number',''),r.get('FILE NAME',''), r.get('Govt_Share', 0),
            f"{MONTH_MAP[current_month]} Govt_Share Released",
            r.get('NAME OF CF','')
        ])
        
    # --- Sheet 2: Missing ---
    ws2 = wb.create_sheet(title='Missing_or_Zero_List' )
    ws2.append(['BANK_SNO', 'Name Of Benificiary', 'Aadhar Number', 'Bank Account Number', 'RD 1','Contact Number','FILE NAME', 'Status','NAME OF CF'])
    for col_num in range(1, 10):
        ws2.cell(row=1, column=col_num).font = header_font
        ws2.cell(row=1, column=col_num).fill = red_fill
        
    for _, r in df_missing.iterrows():
        ws2.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Aadhar Number', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''),r.get('Contact Number',''),r.get('FILE NAME',''),
            f"Zero or missing transactions for {MONTH_MAP[current_month]}",
            r.get('NAME OF CF','')
        ])
        
    # --- Sheet 3: All ---
    ws3 = wb.create_sheet(title='All_Beneficiaries_Master')
    ws3.append(['BANK_SNO', 'Name Of Benificiary', 'Aadhar Number', 'Bank Account Number', 'RD 1','Contact Number','FILE NAME', 'Govt_Share'])
    for col_num in range(1, 9):
        ws3.cell(row=1, column=col_num).font = header_font
        ws3.cell(row=1, column=col_num).fill = green_fill
        
    for _, r in group_df.iterrows():
        ws3.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Aadhar Number', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''),r.get('Contact Number',''),r.get('FILE NAME',''), r.get('Govt_Share', 0)
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
        aadhar_number = str(data.get('aadhar_number', '')).strip()
        selected_file_name = str(data.get('file_name', '')).strip()
        
        if not aadhar_number:
            return jsonify({"status": "error", "message": "దయచేసి ఆధార్ నంబర్ ఎంటర్ చేయండి!"})
            
        if not selected_file_name:
            return jsonify({"status": "error", "message": "దయచేసి ముందుగా FILE NAME ఎంచుకోండి!"})
            
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        # 'Bank_Email_Ids' నుండి ఈమెయిల్ ఐడిని పొందడం
        email_id = ""
        try:
            ws_email = spreadsheet.worksheet('Bank_Email_Ids')
            email_records = ws_email.get_all_records()
            for row in email_records:
                if str(row.get('FILE NAME', '')).strip().upper() == selected_file_name.upper():
                    email_id = str(row.get('OFFICIAL_EMAIL', '')).strip()
                    break
        except Exception:
            pass

        worksheet_new = spreadsheet.worksheet('NEW THRIFT DATA')
        records_new = worksheet_new.get_all_records()
        columns = worksheet_new.row_values(1)
        
        matched_record = None
        row_index = None
        source_sheet = None
        
        for idx, row in enumerate(records_new, start=2):
            if str(row.get('Aadhar Number', '')).strip() == aadhar_number:
                matched_record = row
                row_index = idx
                source_sheet = 'new'
                break
                
        if not matched_record:
            try:
                worksheet_old = spreadsheet.worksheet('Old Thrift Data')
                records_old = worksheet_old.get_all_records()
                for row in records_old:
                    if str(row.get('Aadhar Number', '')).strip() == aadhar_number:
                        matched_record = row
                        source_sheet = 'old'
                        break
            except Exception:
                pass

        # BANK_SNO గణన: యూజర్ ఎంచుకున్న FILE NAME కి మ్యాచ్ అయ్యే రికార్డ్స్‌లోని 'BANK_SNO' లలో గరిష్ట విలువ + 1
        max_bank_sno = 0
        for row in records_new:
            if str(row.get('FILE NAME', '')).strip().upper() == selected_file_name.upper():
                b_sno = str(row.get('BANK_SNO', '0')).strip()
                if b_sno.isdigit():
                    max_bank_sno = max(max_bank_sno, int(b_sno))
        next_bank_sno = max_bank_sno + 1

        # వివిధ కాలమ్స్‌కోసం యూనిక్ లిస్ట్స్ సేకరించడం (Suggestions / Dropdowns కోసం)
        unique_professions = sorted(list(set([str(r.get('Profession', '')).strip() for r in records_new if r.get('Profession')])))
        unique_villages = sorted(list(set([str(r.get('Village', '')).strip() for r in records_new if r.get('Village')])))
        unique_areas = sorted(list(set([str(r.get('Area', '')).strip() for r in records_new if r.get('Area')])))
        unique_unit_owners = sorted(list(set([str(r.get('Unit Owner', '')).strip() for r in records_new if r.get('Unit Owner')])))
        unique_unit_hns = sorted(list(set([str(r.get('H.NO. Unit H.No', '')).strip() for r in records_new if r.get('H.NO. Unit H.No')])))
        unique_relations = sorted(list(set([str(r.get('Relation', '')).strip() for r in records_new if r.get('Relation')])))

        # బ్యాంక్ నేమ్ కూడా FILE NAME అవుతుంది
        bank_name = selected_file_name

        if matched_record and source_sheet == 'new':
            return jsonify({
                "status": "existing",
                "message": "రికార్డు కనుగొనబడింది - ఇది అప్‌డేట్ (Update) అవుతుంది",
                "data": matched_record,
                "row_index": row_index,
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
            # కొత్త రికార్డు లేదా ఓల్డ్ డేటా నుండి లోడ్ అయినట్లయితే డిఫాల్ట్ వాల్యూస్ సెట్ చేయబడతాయి
            base_data = matched_record.copy() if matched_record else {}
            base_data['Aadhar Number'] = aadhar_number
            base_data['BANK_SNO'] = next_bank_sno
            base_data['FILE NAME'] = selected_file_name
            base_data['Bank Name'] = bank_name
            base_data['Caste'] = 'BC'
            base_data['Profession'] = 'POWER LOOM WEAVER'
            base_data['DISTRICT NAME'] = 'RAJANNA SIRCILLA'
            base_data['MANDAL NAME'] = 'SIRCILLA' # వీవర్లు ఎక్కువగా ఉన్న మండలం డిఫాల్ట్
            base_data['NAME OF CF'] = cf_name
            base_data['EMAIL ID'] = email_id
            base_data['BANK SENT'] = 'NOT SENT'
            base_data['8% Of Monthly Wages'] = 1200
            base_data['MONTHLY WAGES'] = int(1200 * 100 / 8) # 15000

            return jsonify({
                "status": "new",
                "message": "ఇది కొత్త ఎంట్రీ (New Entry) కింద సేవ్ అవుతుంది. దయచేసి వివరాలను పరిశీలించి అవసరమైన మార్పులు చేయండి.",
                "data": base_data,
                "row_index": '',
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


# బ్రౌజర్ లో డ్రాప్-డౌన్ కోసం Bank_Email_Ids నుండి FILE NAME లను పంపే రౌట్
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
        file_names = [str(r.get('FILE NAME', '')).strip() for r in records if r.get('FILE NAME')]
        return jsonify({"status": "success", "file_names": file_names})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e), "file_names": []})


# 6. డేటాను అప్‌డేట్ లేదా కొత్తది సేవ్ చేయడం (Update or Insert Record)
@cf_bp.route('/cf/save-beneficiary', methods=['POST'])
def cf_save_beneficiary():
    if not session.get('cf_logged_in'):
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})
        
    try:
        from app import get_sheets_client, SPREADSHEET_NAME
        
        cf_name = session.get('cf_name')
        data = request.get_json() if request.is_json else request.form
        
        action_type = data.get('action_type') # 'update' లేదా 'insert'
        row_index = data.get('row_index')
        form_data = data.get('form_data') # డిక్షనరీ రూపంలో ఫీల్డ్స్
        
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        worksheet = spreadsheet.worksheet('NEW THRIFT DATA')
        headers = worksheet.row_values(1)
        
        if action_type == 'update' and row_index:
            # పాత రో ని అప్‌డేట్ చేయడం
            for col_name, val in form_data.items():
                if col_name in headers:
                    col_idx = headers.index(col_name) + 1
                    worksheet.update_cell(int(row_index), col_idx, val)
            return jsonify({"status": "success", "message": "డేటా విజయవంతంగా అప్‌డేట్ చేయబడింది!"})
            
        elif action_type == 'insert':
            # కొత్త రికార్డును ఇన్‌సర్ట్ చేయడం
            new_row = []
            max_sno = 0
            
            # SNO ఆటో-ఇంక్రిమెంట్ కోసం
            try:
                sno_col_idx = headers.index('SNO') + 1
                all_snos = worksheet.col_values(sno_col_idx)[1:]
                max_sno = max([int(s) for s in all_snos if str(s).isdigit()] or [0])
            except:
                max_sno = len(worksheet.get_all_values())
                
            form_data['SNO'] = max_sno + 1
            form_data['NAME OF CF'] = cf_name # లాగిన్ అయిన CF పేరు ఆటోమేటిక్ గా సెట్ అవుతుంది
            
            for header in headers:
                new_row.append(form_data.get(header, ''))
                
            worksheet.append_row(new_row)
            return jsonify({"status": "success", "message": "కొత్త రికార్డు విజయవంతంగా సేవ్ చేయబడింది!"})
            
        return jsonify({"status": "error", "message": "సరికైన యాక్షన్ రాలేదు!"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})
