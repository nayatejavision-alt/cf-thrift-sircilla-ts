import os
import json
from flask import Flask, request, jsonify, render_template, session
from flask_cors import CORS
from google.oauth2 import service_account
import google.auth
import google.auth.transport.requests
import gspread
from datetime import timedelta
import csv
from flask import make_response
import os
from werkzeug.utils import secure_filename
from admin_routes import admin_bp
from cf_routes import cf_bp
import io
import pandas as pd
from flask import send_file

app = Flask(__name__)
app.secret_key = "thrift_scheme_secret_key"
app.permanent_session_lifetime = timedelta(days=7) # ఉదాహరణకు 7 రోజుల పాటు లాగిన్ పోదు! 
CORS(app)

SPREADSHEET_NAME = "NOT_BANK_SENT_FORM"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.register_blueprint(admin_bp)
app.register_blueprint(cf_bp)
# Google Sheets API Connection Setup Function

def get_sheets_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    credentials_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    
    if credentials_json_str:
        info = json.loads(credentials_json_str)
        credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    else:
        cred_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        if cred_path and os.path.exists(cred_path):
            credentials = service_account.Credentials.from_service_account_file(cred_path, scopes=scopes)
        else:
            credentials, _ = google.auth.default(scopes=scopes)
            
    # టోకెన్ చెల్లుబాటు పర్యవేక్షణ
    request_session = google.auth.transport.requests.Request()
    if not credentials.valid:
        credentials.refresh(request_session)
            
    client = gspread.authorize(credentials)
    return client

# 1. హోమ్ పేజీ రౌట్ (బ్యాంక్ లిస్ట్‌ను డ్రాప్‌డౌన్ కోసం పంపుతుంది)
@app.route("/")
def home():
    bank_list = []
    try:
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        # BANK_LOGINS ట్యాబ్ నుండి బ్యాంకుల వివరాలు తీసుకోవడం
        login_sheet = spreadsheet.worksheet("BANK_LOGINS")
        logins = login_sheet.get_all_records()
        
        for row in logins:
            b_name = (
                row.get("Bank Name")
                or row.get("BANK NAME")
                or row.get("Bank_Name")
            )
            if b_name:
                cleaned_name = str(b_name).strip()
                if cleaned_name and cleaned_name not in bank_list:
                    bank_list.append(cleaned_name)

        print("--- డ్రాప్‌డౌన్ కోసం విజయవంతంగా లోడ్ అయిన బ్యాంకులు ---")
        print(bank_list)

    except Exception as e:
        print("Error fetching bank list:", e)
        bank_list = []

    return render_template("index.html", banks=bank_list)

# 2. ఆధార్ శోధన API రౌట్
@app.route('/api/search', methods=['POST'])
def search_beneficiary():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': 'చెల్లుబాటు అయ్యే వివరాలు పంపలేదు.'}), 400

        input_id = str(data.get('aadhar', '')).strip().replace('-', '').replace(' ', '')

        if not input_id or len(input_id) != 12 or not input_id.isdigit():
            return jsonify({'success': False, 'message': 'దయచేసి సరైన 12 అంకెల సంఖ్యను ఎంటర్ చేయండి.'}), 400

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
        records = target_sheet.get_all_records()

        matched_record = None
        for row in records:
            raw_val = str(row.get('Aadhar Number') or row.get('ADHAR') or row.get('AADHAR') or '').split('.')[0].strip()
            if raw_val == input_id:
                matched_record = row
                break

        if matched_record:
            masked_id = f"XXXX-XXXX-{input_id[-4:]}"

            result = {
                'sno': matched_record.get('BANK_SNO'),
                'masked_aadhar': masked_id,
                'name': matched_record.get('Name Of Benificiary') or matched_record.get('NAME'),
                'father_name': matched_record.get('Father Name') or matched_record.get('FATHER NAME'),
                'village': matched_record.get('VILLAGE NAME') or matched_record.get('VILLAGE'),
                'mandal': matched_record.get('MANDAL NAME') or matched_record.get('MANDAL'),
                'bank_name': matched_record.get('Bank Name') or matched_record.get('BANK NAME'),
                'rd1_status': matched_record.get('RD 1'),
                'rd2_status': matched_record.get('RD 2'),
                'wages': matched_record.get('MONTHLY WAGES')
            }
            return jsonify({'success': True, 'data': result})
        else:
            return jsonify({'success': False, 'message': 'సదరు సంఖ్యతో ఎలాంటి వివరాలు లభించలేదు.'})

    except Exception as e:
        print(f"Error details: {repr(e)}")
        return jsonify({'success': False, 'message': f'సర్వర్ లోపం సంభవించింది: {str(e)}'}), 500

# 3. బ్యాంక్ లాగిన్ API
@app.route("/api/login", methods=["POST"])
def bank_login():
    try:
        data = request.get_json()
        bank_name = str(data.get("bank_name", "")).strip()
        pin = str(data.get("pin", "")).strip()

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        login_sheet = spreadsheet.worksheet("BANK_LOGINS")
        logins = login_sheet.get_all_records()

        authenticated = False
        for row in logins:
            sheet_bank = str(row.get("Bank Name") or row.get("BANK NAME") or "").strip()
            sheet_pin = str(row.get("PIN") or row.get("Pin") or "").strip()

            if sheet_bank.lower() == bank_name.lower() and sheet_pin == pin:
                authenticated = True
                break

        if authenticated:
            session["bank_user"] = bank_name
            session.permanent = True
            return jsonify({"status": "success", "message": "లాగిన్ విజయవంతమైంది!"})
        else:
            return jsonify({"status": "error", "message": "తప్పు PIN లేదా బ్యాంక్ పేరు!"})

    except Exception as e:
        return jsonify({"status": "error", "message": f"లాగిన్ లోపం: {str(e)}"})

# 4. లాగిన్ అయిన బ్యాంకు డేటాను తీసుకునే API
@app.route("/api/bank-data", methods=["GET"])
def get_bank_data():
    try:
        bank_user = session.get("bank_user")
        if not bank_user:
            return jsonify({"status": "error", "message": "దయచేసి ముందుగా లాగిన్ అవ్వండి!"})

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        # 1. మాస్టర్ డేటా తీసుకోవడం
        target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
        master_records = target_sheet.get_all_records()

        # 2. ట్రాన్సాక్షన్ డేటా (Transactions_Year1) తీసుకోవడం
        try:
            trans_sheet = spreadsheet.worksheet('Transactions_Year1')
            trans_records = trans_sheet.get_all_records()
        except Exception:
            trans_records = []

        current_month = 1  # ప్రస్తుత నెల సంఖ్య (month_number = 1)
        bank_user_clean = bank_user.strip().lower()

        # 💡 ఆప్టిమైజేషన్: ట్రాన్సాక్షన్స్ కోసం సెట్ క్రియేట్ చేయడం
        completed_transactions = set()
        for t in trans_records:
            t_sno = str(t.get("BANK_SNO") or t.get("Bank SNO") or "").strip()
            t_month = str(t.get("month_number") or t.get("MONTH") or t.get("MONTH_NUMBER") or "").strip()
            t_filename = str(t.get("FILE NAME") or t.get("File Name") or "").strip().lower()

            if t_month == str(current_month) and (bank_user_clean in t_filename or t_filename in bank_user_clean):
                bal_val = t.get("Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)") or t.get("Bank Closing Balance") or ""
                balance_map[t_sno] = bal_val
            
            try:
                govt_share = float(t.get("Govt_Share") or t.get("GOVT_SHARE") or 0)
            except ValueError:
                govt_share = 0

            # నెల, govt_share > 0 మరియు బ్యాంక్ పేరు మ్యాచ్ అవ్వడం (సబ్-స్ట్రింగ్ మ్యాచ్)
            if t_month == str(current_month) and govt_share > 0 and (bank_user_clean in t_filename or t_filename in bank_user_clean):
                completed_transactions.add(t_sno)

        # 3. మాస్టర్ డేటాను ఫిల్టర్ చేయడం
        filtered_data = []
        for row in master_records:
            file_name = str(row.get("FILE NAME") or row.get("File Name") or "").strip()
            
            if bank_user_clean in file_name.lower():
                bank_sno = str(row.get("BANK_SNO") or "").strip()

                # ట్రాన్సాక్షన్ షీట్‌లో బ్యాలెన్స్ ఉంటే దానిని రో లోకి సెట్ చేయడం
                if bank_sno in balance_map:
                    row['Bank Closing Balance'] = balance_map[bank_sno]
                
                # ట్రాన్సాక్షన్ షీట్‌లో ఈ BANK_SNO ఉందో లేదో చెక్ చేయడం
                is_completed = bank_sno in completed_transactions
                
                row['is_completed'] = is_completed
                filtered_data.append(row)

        return jsonify({"status": "success", "bank": bank_user, "data": filtered_data})

    except Exception as e:
        print(f"Get Bank Data Error: {repr(e)}")
        return jsonify({"status": "error", "message": f"డేటా లోడ్ అవ్వడంలో లోపం: {str(e)}"})



        
@app.route("/api/download-csv", methods=["GET"])
def download_csv():
    bank_user = session.get("bank_user")
    if not bank_user:
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    target_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
    records = target_sheet.get_all_records()

    # ఆ బ్యాంక్ డేటాను ఫిల్టర్ చేయడం
    filtered_data = [
        row
        for row in records
        if bank_user.lower()
        in str(row.get("FILE NAME") or row.get("File Name") or "").lower()
    ]

    # CSV రెస్పాన్స్ తయారు చేయడం
    output = []
    output.append(["SNO", "Name", "Account No", "File Name"])

    for idx, row in enumerate(filtered_data):
        output.append(
            [
                idx + 1,
                row.get("Name Of Benificiary") or row.get("NAME") or "",
                row.get("Bank Account Number") or row.get("ACCOUNT") or "",
                row.get("FILE NAME") or "",
            ]
        )

    response = make_response()
    writer = csv.writer(response.stream)
    for row in output:
        writer.writerow(row)

    response.headers["Content-Disposition"] = (
        f"attachment; filename={bank_user}_beneficiaries.csv"
    )
    response.headers["Content-Type"] = "text/csv"
    return response
    

@app.route("/api/update-rd-accounts", methods=["POST"])
def update_rd_accounts():
    try:
        bank_user = session.get("bank_user")
        if not bank_user:
            return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

        data = request.get_json()
        bank_sno = str(data.get("bank_sno")).strip()
        rd1 = str(data.get("rd1", "")).strip()
        rd2 = str(data.get("rd2", "")).strip()
        bank_closing_balance_str = str(data.get("bank_closing_balance", "")).strip()
        
        # బ్యాలెన్స్‌ను నంబర్‌గా మార్చడం (లెక్కల కోసం)
        try:
            bank_closing_balance = float(bank_closing_balance_str) if bank_closing_balance_str else 0.0
        except ValueError:
            bank_closing_balance = 0.0

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        # 1. 'NEW THRIFT DATA' (Master) లో RD 1 మరియు RD 2 అప్‌డేట్ చేయడం
        target_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        master_records = target_sheet.get_all_records()
        
        master_row_index = None
        beneficiary_name = ""
        beneficiary_acc = ""
        monthly_wages = 0.0
        
        for idx, row in enumerate(master_records):
            row_bank_sno = str(row.get("BANK_SNO") or row.get("Bank SNO") or "").strip()
            file_name = str(row.get("FILE NAME") or row.get("File Name") or "").strip()
            
            if row_bank_sno == bank_sno and bank_user.lower() in file_name.lower():
                master_row_index = idx + 2  # gspread row index + header row
                beneficiary_name = str(row.get("Name Of Benificiary") or row.get("NAME") or "").strip()
                beneficiary_acc = str(row.get("Bank Account Number") or row.get("ACCOUNT") or "").strip()
                
                # '8% Of Monthly Wages' లేదా డ్యూ లెక్కించడానికి మాస్టర్ డేటా నుండి విలువ తీసుకోవడం
                # (మీ షీట్‌లో ఆ కాలమ్ పేరు ఎలా ఉందో దాని ప్రకారం ఇక్కడ తీసుకోవచ్చు, ఉదాహరణకు 'Monthly Wages' లేదా డైరెక్ట్ డ్యూ)
                try:
                    wages_val = row.get("8% Of Monthly Wages") or row.get("MONTHLY WAGES") or 0
                    monthly_wages = float(wages_val)
                except ValueError:
                    monthly_wages = 0.0
                break
                
        if master_row_index:
            header_row = target_sheet.row_values(1)
            rd1_col, rd2_col = None, None
            for c_idx, header in enumerate(header_row):
                h_cleaned = header.strip().upper()
                if h_cleaned in ["RD 1", "RD_1"]:
                    rd1_col = c_idx + 1
                elif h_cleaned in ["RD 2", "RD_2"]:
                    rd2_col = c_idx + 1
                    
            if rd1_col:
                target_sheet.update_cell(master_row_index, rd1_col, rd1)
            if rd2_col:
                target_sheet.update_cell(master_row_index, rd2_col, rd2)

        # మాస్టర్ నుండి డ్యూ (Due = 8% of Monthly Wages) లెక్కించడం (లేదా మాస్టర్ రో నుండి నేరుగా తెచ్చుకోవడం)
        # ఒకవేళ మీ మాస్టర్ షీట్‌లో డ్యూ నేరుగా ఉంటే ఆ కాలమ్ వాల్యూ వాడొచ్చు. ఉదాహరణకు 8% వేతనం:
        due_amount = monthly_wages  # ఒకవేళ వేతనాలు లేకపోతే డిఫాల్ట్ డ్యూ లేదా మాస్టర్ నుండి తెచ్చుకోవచ్చు

        # 2. 'Transactions_Year1' షీట్‌లో డిపెండెన్సీ ఫార్ములాలతో అప్‌డేట్ చేయడం
        trans_sheet = spreadsheet.worksheet("Transactions_Year1")
        trans_records = trans_sheet.get_all_records()

        current_month = 1
        trans_row_index = None
        
        for idx, row in enumerate(trans_records):
            t_sno = str(row.get("BANK_SNO") or row.get("Bank SNO") or "").strip()
            t_month = str(row.get("MONTH_NUMBER") or row.get("month_number") or "").strip()
            t_filename = str(row.get("FILE NAME") or row.get("File Name") or "").strip().lower()
            
            if t_sno == bank_sno and t_month == str(current_month) and bank_user.lower() in t_filename:
                trans_row_index = idx + 2
                # ఒకవేళ ఆల్రెడీ ఆ రో లో Due ఉంటే దానిని కూడా తీసుకోవచ్చు
                try:
                    existing_due = float(row.get("(C) ఈ నెల డ్యూ(Due)") or 0)
                    if existing_due > 0:
                        due_amount = existing_due
                except ValueError:
                    pass
                break

        trans_headers = trans_sheet.row_values(1)
        col_map = {h.strip(): c_idx + 1 for c_idx, h in enumerate(trans_headers)}

        # మీరు చెప్పిన ఫార్ములాల ప్రకారం లెక్కలు కట్టడం:
        previous_arrears = 0.0  # (B) మొదటి నెల కాబట్టి 0
        actual_paid = bank_closing_balance  # (D) మొదటి నెల కాబట్టి Closing Balance సమానం
        
        # Govt_Share = min(Due, Actual Paid)
        govt_share = min(due_amount, actual_paid)
        
        # Remaining Arrears = max(0, Previous Arrears + Due - Actual Paid)
        remaining_arrears = max(0.0, previous_arrears + due_amount - actual_paid)
        
        # Govt RD Closing Bal (H) = Govt_Share
        govt_rd_closing_bal = govt_share
        
        # Govt RD Status (I) = "MATCHED" ఒకవేళ Arrears <= 0 లేదంటే "PENDING"
        govt_rd_status = "MATCHED" if remaining_arrears <= 0 else "PENDING"

        # కాలమ్ పేర్లు మీ షీట్ ప్రకారం:
        rd1_col_name = "RD ACCOUNT_1 (Beneficiary Contribution)"
        balance_col_name = "Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)"
        arrears_b_name = "(B) పాత బాకీ(Previous Arrears)"
        due_c_name = "(C) ఈ నెల డ్యూ(Due)"
        actual_paid_d_name = "(D) Actual Paid"
        govt_share_name = "Govt_Share"
        remaining_f_name = "(F) మిగిలిన బాకీ(Remaining Arrears)"
        govt_closing_h_name = "Govt RD Closing Bal (H)"
        govt_status_i_name = "Govt RD Status (I)"

        if trans_row_index:
            # ఒకవేళ ఆల్రెడీ ఎంట్రీ ఉంటే అన్ని డిపెండెంట్ కాలమ్స్‌ని అప్‌డేట్ చేయడం
            updates = [
                (rd1_col_name, rd1),
                (balance_col_name, bank_closing_balance_str),
                (arrears_b_name, previous_arrears),
                (due_c_name, due_amount),
                (actual_paid_d_name, actual_paid),
                (govt_share_name, govt_share),
                (remaining_f_name, remaining_arrears),
                (govt_closing_h_name, govt_rd_closing_bal),
                (govt_status_i_name, govt_rd_status)
            ]
            for col_name, val in updates:
                if col_name in col_map:
                    trans_sheet.update_cell(trans_row_index, col_map[col_name], val)
        else:
            # కొత్త ఎంట్రీ అయితే డిక్షనరీ ద్వారా అన్ని కాలమ్స్‌ని ఫిల్ చేసి అపెండ్ చేయడం
            row_data_dict = {h.strip(): "" for h in trans_headers}
            
            for h in trans_headers:
                h_clean = h.strip()
                h_upper = h_clean.upper()
                
                if h_clean in ["MONTH_NUMBER", "month_number"]:
                    row_data_dict[h_clean] = current_month
                elif h_clean in ["BANK_SNO", "Bank SNO"]:
                    row_data_dict[h_clean] = bank_sno
                elif "NAME OF BENIFICIARY" in h_upper or h_clean == "Name Of Benificiary":
                    row_data_dict[h_clean] = beneficiary_name
                elif "BANK ACCOUNT NUMBER" in h_upper or h_clean == "Bank Account Number":
                    row_data_dict[h_clean] = beneficiary_acc
                elif h_clean == rd1_col_name:
                    row_data_dict[h_clean] = rd1
                elif h_clean == balance_col_name:
                    row_data_dict[h_clean] = bank_closing_balance_str
                elif h_clean == arrears_b_name:
                    row_data_dict[h_clean] = previous_arrears
                elif h_clean == due_c_name:
                    row_data_dict[h_clean] = due_amount
                elif h_clean == actual_paid_d_name:
                    row_data_dict[h_clean] = actual_paid
                elif h_clean == govt_share_name:
                    row_data_dict[h_clean] = govt_share
                elif h_clean == remaining_f_name:
                    row_data_dict[h_clean] = remaining_arrears
                elif h_clean == govt_closing_h_name:
                    row_data_dict[h_clean] = govt_rd_closing_bal
                elif h_clean == govt_status_i_name:
                    row_data_dict[h_clean] = govt_rd_status
                elif h_clean == "FILE NAME":
                    row_data_dict[h_clean] = bank_user

            new_row_data = [row_data_dict.get(h.strip(), "") for h in trans_headers]
            trans_sheet.append_row(new_row_data)

        return jsonify({"status": "success", "message": "డేటా మరియు ఫార్ములా లెక్కలు విజయవంతంగా అప్‌డేట్ చేయబడ్డాయి!"})

    except Exception as e:
        print(f"Update Error Exception: {repr(e)}")
        return jsonify({"status": "error", "message": f"లోపం సంభవించింది: {str(e)}"})


@app.route("/api/upload-file", methods=["POST"])
def upload_file():
    bank_user = session.get("bank_user")
    if not bank_user:
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    if "file" not in request.files:
        return jsonify({"status": "error", "message": "ఫైల్ ఏదీ ఎంచుకోలేదు."})

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"status": "error", "message": "ఫైల్ పేరు ఖాళీగా ఉంది."})

    if file:
        filename = secure_filename(file.filename)
        # బ్యాంక్ పేరుతో ఫైల్ సేవ్ అవ్వడానికి
        safe_bank_name = bank_user.replace(" ", "_")
        final_filename = f"{safe_bank_name}_{filename}"
        file.save(os.path.join(UPLOAD_FOLDER, final_filename))

        return jsonify(
            {
                "status": "success",
                "message": "ఫైల్ విజయవంతంగా అప్‌లోడ్ చేయబడింది!",
            }
        )



@app.route("/api/download-excel", methods=["GET"])
def download_excel():
    bank_user = session.get("bank_user")
    if not bank_user:
        return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    try:
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        # మాస్టర్ డేటా తీసుకోవడం మరియు FILE NAME తో ఫిల్టర్ చేయడం
        target_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        master_records = target_sheet.get_all_records()
        df_master = pd.DataFrame(master_records)
        
        # కేవలం లాగిన్ అయిన బ్యాంకుకు సంబంధించిన రికార్డ్స్ మాత్రమే ఫిల్టర్ చేయడం (case-insensitive సురక్షితంగా ఉండటానికి)
        df_master = df_master[df_master['FILE NAME'].astype(str).str.strip().str.lower() == bank_user.strip().lower()].copy()

        # ట్రాన్సాక్షన్ డేటా తీసుకోవడం
        try:
            trans_sheet = spreadsheet.worksheet("Transactions_Year1")
            trans_records = trans_sheet.get_all_records()
            df_trans = pd.DataFrame(trans_records)
            df_trans = df_trans[df_trans['FILE NAME'].astype(str).str.strip().str.lower() == bank_user.strip().lower()].copy()
        except Exception:
            df_trans = pd.DataFrame()

        # ట్రాన్సాక్షన్స్ నుండి బ్యాలెన్స్‌లను మ్యాప్ చేయడానికి ఒక డిక్షనరీ తయారు చేయడం (iterrows వాడాలి)
        balance_map = {}
        if not df_trans.empty:
            for _, t in df_trans.iterrows():
                t_sno = str(t.get("BANK_SNO") or t.get("Bank SNO") or "").strip()
                t_month = str(t.get("MONTH_NUMBER") or t.get("month_number") or "").strip()
                t_balance = t.get("Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)") or t.get("Bank Closing Balance") or ""
                
                if t_month == "1" and t_sno:
                    balance_map[t_sno] = t_balance

        filtered_data = []
        # ఇక్కడ iterrows() వాడటం వల్ల ప్రతి రో (row) సరిగ్గా వస్తుంది
        for idx, (_, row) in enumerate(df_master.iterrows()):
            file_name = str(row.get("FILE NAME") or row.get("File Name") or "").strip()
            bank_sno = str(row.get("BANK_SNO") or row.get("Bank SNO") or "").strip()
            
            # ట్రాన్సాక్షన్ షీట్ నుండి లేదా మాస్టర్ నుండి బ్యాలెన్స్ తీసుకోవడం
            closing_balance = balance_map.get(bank_sno, row.get("Bank Closing Balance", ""))

            filtered_data.append({
                "BANK_SNO": bank_sno or idx + 1,
                "Name Of Benificiary": row.get("Name Of Benificiary") or row.get("NAME") or "",
                "Bank Account Number": row.get("Bank Account Number") or row.get("ACCOUNT") or "",
                "RD ACCOUNT_1": row.get("RD 1") or row.get("RD_1") or "",
                "RD ACCOUNT_2": row.get("RD 2") or row.get("RD_2") or "",
                "Bank Closing Balance": closing_balance,
                "FILE NAME": file_name
            })

        if not filtered_data:
            return jsonify({"status": "error", "message": "డెటా ఏదీ అందుబాటులో లేదు."})

        df = pd.DataFrame(filtered_data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Pending_Entries')
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{bank_user.replace(' ', '_')}_Thrift_Data.xlsx"
        )
    except Exception as e:
        print(f"Excel Download Error: {repr(e)}")
        return jsonify({"status": "error", "message": f"డౌన్‌లోడ్ చేయడంలో లోపం: {str(e)}"})
        
@app.route("/api/bulk-upload", methods=["POST"])
def bulk_upload():
    try:
        bank_user = session.get("bank_user")
        if not bank_user:
            return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

        if "file" not in request.files:
            return jsonify({"status": "error", "message": "ఫైల్ ఏదీ ఎంచుకోలేదు."})

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"status": "error", "message": "ఫైల్ పేరు ఖాళీగా ఉంది."})

        # పాండాస్ ద్వారా అప్‌లోడ్ అయిన ఎక్సెల్ ఫైల్‌ను చదవడం
        df = pd.read_excel(file)
        
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        target_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        master_records = target_sheet.get_all_records()
        master_headers = target_sheet.row_values(1)
        master_col_map = {h.strip().upper(): c_idx + 1 for c_idx, h in enumerate(master_headers)}

        trans_sheet = spreadsheet.worksheet("Transactions_Year1")
        trans_records = trans_sheet.get_all_records()
        trans_headers = trans_sheet.row_values(1)
        trans_col_map = {h.strip(): c_idx + 1 for c_idx, h in enumerate(trans_headers)}

        current_month = 1
        updated_count = 0

        # ఎక్సెల్ ఫైల్‌లోని ప్రతీ రో (Row) ని లూప్ తిప్పుతూ అప్‌డేట్ చేయడం
        for _, row in df.iterrows():
            bank_sno = str(row.get("BANK_SNO") or "").strip()
            rd1 = str(row.get("RD ACCOUNT_1") or "").strip()
            rd2 = str(row.get("RD ACCOUNT_2") or "").strip()
            balance_str = str(row.get("Bank Closing Balance") or "").strip()

            if not bank_sno:
                continue

            try:
                bank_closing_balance = float(balance_str) if balance_str else 0.0
            except ValueError:
                bank_closing_balance = 0.0

            # 1. Master Sheet లో అప్‌డేట్ చేయడానికి రో ఇండెక్స్ కనుగొనడం
            master_row_idx = None
            monthly_wages = 0.0
            beneficiary_name = ""
            beneficiary_acc = ""

            for idx, m_row in enumerate(master_records):
                m_sno = str(m_row.get("BANK_SNO") or "").strip()
                m_file = str(m_row.get("FILE NAME") or m_row.get("File Name") or "").strip()
                if m_sno == bank_sno and bank_user.lower() in m_file.lower():
                    master_row_idx = idx + 2
                    beneficiary_name = str(m_row.get("Name Of Benificiary") or m_row.get("NAME") or "").strip()
                    beneficiary_acc = str(m_row.get("Bank Account Number") or m_row.get("ACCOUNT") or "").strip()
                    try:
                        monthly_wages = float(m_row.get("8% Of Monthly Wages") or m_row.get("MONTHLY WAGES") or 0)
                    except ValueError:
                        monthly_wages = 0.0
                    break

            if master_row_idx:
                if "RD 1" in master_col_map:
                    target_sheet.update_cell(master_row_idx, master_col_map["RD 1"], rd1)
                elif "RD_1" in master_col_map:
                    target_sheet.update_cell(master_row_idx, master_col_map["RD_1"], rd1)
                    
                if "RD 2" in master_col_map:
                    target_sheet.update_cell(master_row_idx, master_col_map["RD 2"], rd2)
                elif "RD_2" in master_col_map:
                    target_sheet.update_cell(master_row_idx, master_col_map["RD_2"], rd2)

            # 2. Transactions Sheet లో లెక్కలు కట్టి అప్‌డేట్ చేయడం
            due_amount = monthly_wages
            previous_arrears = 0.0
            actual_paid = bank_closing_balance
            govt_share = min(due_amount, actual_paid)
            remaining_arrears = max(0.0, previous_arrears + due_amount - actual_paid)
            govt_rd_closing_bal = govt_share
            govt_rd_status = "MATCHED" if remaining_arrears <= 0 else "PENDING"

            trans_row_idx = None
            for idx, t_row in enumerate(trans_records):
                t_sno = str(t_row.get("BANK_SNO") or t_row.get("Bank SNO") or "").strip()
                t_month = str(t_row.get("MONTH_NUMBER") or t_row.get("month_number") or "").strip()
                t_filename = str(t_row.get("FILE NAME") or t_row.get("File Name") or "").strip().lower()
                if t_sno == bank_sno and t_month == str(current_month) and bank_user.lower() in t_filename:
                    trans_row_idx = idx + 2
                    break

            rd1_col_name = "RD ACCOUNT_1 (Beneficiary Contribution)"
            balance_col_name = "Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)"
            arrears_b_name = "(B) పాత బాకీ(Previous Arrears)"
            due_c_name = "(C) ఈ నెల డ్యూ(Due)"
            actual_paid_d_name = "(D) Actual Paid"
            govt_share_name = "Govt_Share"
            remaining_f_name = "(F) మిగిలిన బాకీ(Remaining Arrears)"
            govt_closing_h_name = "Govt RD Closing Bal (H)"
            govt_status_i_name = "Govt RD Status (I)"

            if trans_row_idx:
                updates = [
                    (rd1_col_name, rd1),
                    (balance_col_name, balance_str),
                    (due_c_name, due_amount),
                    (actual_paid_d_name, actual_paid),
                    (govt_share_name, govt_share),
                    (remaining_f_name, remaining_arrears),
                    (govt_closing_h_name, govt_rd_closing_bal),
                    (govt_status_i_name, govt_rd_status)
                ]
                for col_name, val in updates:
                    if col_name in trans_col_map:
                        trans_sheet.update_cell(trans_row_idx, trans_col_map[col_name], val)
            else:
                row_data_dict = {h.strip(): "" for h in trans_headers}
                for h in trans_headers:
                    h_clean = h.strip()
                    if h_clean in ["MONTH_NUMBER", "month_number"]:
                        row_data_dict[h_clean] = current_month
                    elif h_clean in ["BANK_SNO", "Bank SNO"]:
                        row_data_dict[h_clean] = bank_sno
                    elif h_clean == rd1_col_name:
                        row_data_dict[h_clean] = rd1
                    elif h_clean == balance_col_name:
                        row_data_dict[h_clean] = balance_str
                    elif h_clean == due_c_name:
                        row_data_dict[h_clean] = due_amount
                    elif h_clean == actual_paid_d_name:
                        row_data_dict[h_clean] = actual_paid
                    elif h_clean == govt_share_name:
                        row_data_dict[h_clean] = govt_share
                    elif h_clean == remaining_f_name:
                        row_data_dict[h_clean] = remaining_arrears
                    elif h_clean == govt_closing_h_name:
                        row_data_dict[h_clean] = govt_rd_closing_bal
                    elif h_clean == govt_status_i_name:
                        row_data_dict[h_clean] = govt_rd_status
                    elif h_clean == "FILE NAME":
                        row_data_dict[h_clean] = bank_user

                new_row_data = [row_data_dict.get(h.strip(), "") for h in trans_headers]
                trans_sheet.append_row(new_row_data)

            updated_count += 1

        return jsonify({"status": "success", "message": f"మొత్తం {updated_count} లబ్ధిదారుల వివరాలు విజయవంతంగా బల్క్ అప్‌డేట్ చేయబడ్డాయి! 🎉"})

    except Exception as e:
        print(f"Bulk Upload Error: {repr(e)}")
        return jsonify({"status": "error", "message": f"బల్క్ అప్‌డేట్‌లో లోపం: {str(e)}"})


# 5. లాగౌట్ API
@app.route("/api/logout", methods=["POST"])
def bank_logout():
    session.pop("bank_user", None)
    return jsonify({"status": "success", "message": "లాగౌట్ అయ్యారు."})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
