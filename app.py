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

app = Flask(__name__)
app.secret_key = "thrift_scheme_secret_key"
app.permanent_session_lifetime = timedelta(days=7) # ఉదాహరణకు 7 రోజుల పాటు లాగిన్ పోదు!
CORS(app)

SPREADSHEET_NAME = "NOT_BANK_SENT_FORM"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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

        # 2. ట్రాన్సాక్షన్ డేటా (Transactions_Year1) తీసుకోవడం (డూప్లికేట్ చెక్ చేయడానికి)
        try:
            trans_sheet = spreadsheet.worksheet('Transactions_Year1')
            trans_records = trans_sheet.get_all_records()
        except Exception:
            trans_records = [] # ఒకవేళ షీట్ లేకపోతే ఎంప్టీ లిస్ట్

        current_month = 1  # ప్రస్తుత నెల సంఖ్య (month_number = 1)

        filtered_data = []
        for row in master_records:
            file_name = str(row.get("FILE NAME") or row.get("File Name") or "").strip()
            if bank_user.lower() in file_name.lower():
                bank_sno = str(row.get("BANK_SNO") or "").strip()
                
                # Transactions_Year1 లో ఈ bank_sno మరియు month_number (1) ఉందో లేదో చెక్ చేయడం
                is_completed = False
                for t in trans_records:
                    t_sno = str(t.get("BANK_SNO") or t.get("Bank SNO") or "").strip()
                    t_month = str(t.get("month_number") or t.get("MONTH") or "").strip()
                    if t_sno == bank_sno and t_month == str(current_month):
                        is_completed = True
                        break
                
                # రో డేటాకు ఈ స్టేటస్ జోడించడం
                row['is_completed'] = is_completed
                filtered_data.append(row)

        return jsonify({"status": "success", "bank": bank_user, "data": filtered_data})

    except Exception as e:
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
        rd1 = f"'{data.get('rd1')}'"
        rd2 = f"'{data.get('rd2')}'"

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        target_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        
        records = target_sheet.get_all_records()
        
        row_index = None
        for idx, row in enumerate(records):
            # షీట్‌లోని BANK_SNO మరియు FILE NAME తీసుకోవడం
            row_bank_sno = str(row.get("BANK_SNO") or row.get("Bank SNO") or "").strip()
            file_name = str(row.get("FILE NAME") or row.get("File Name") or "").strip()
            
            # డబుల్ చెకింగ్: BANK_SNO మరియు బ్యాంక్ పేరు రెండూ మ్యాచ్ అవ్వాలి!
            if row_bank_sno == bank_sno and bank_user.lower() in file_name.lower():
                row_index = idx + 2  # gspread row index + header row
                break
                
        if row_index:
            header_row = target_sheet.row_values(1)
            rd1_col = None
            rd2_col = None
            
            for c_idx, header in enumerate(header_row):
                h_cleaned = header.strip().upper()
                if h_cleaned in ["RD 1", "RD_1"]:
                    rd1_col = c_idx + 1
                elif h_cleaned in ["RD 2", "RD_2"]:
                    rd2_col = c_idx + 1
                    
            if rd1_col:
                target_sheet.update_cell(row_index, rd1_col, rd1)
            if rd2_col:
                target_sheet.update_cell(row_index, rd2_col, rd2)
                
            return jsonify({"status": "success", "message": "RD ఖాతాల వివరాలు అప్‌డేట్ అయ్యాయి!"})
        else:
            return jsonify({"status": "error", "message": "సంబంధిత రికార్డు కనుగొనబడలేదు లేదా మీకు ఈ రికార్డును మార్చే అనుమతి లేదు."})

    except Exception as e:
        print(f"Update RD Error: {repr(e)}")
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

# 5. లాగౌట్ API
@app.route("/api/logout", methods=["POST"])
def bank_logout():
    session.pop("bank_user", None)
    return jsonify({"status": "success", "message": "లాగౌట్ అయ్యారు."})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
