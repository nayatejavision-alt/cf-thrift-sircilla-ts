import os
import json
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from google.oauth2 import service_account
import google.auth
import google.auth.transport.requests
import gspread

app = Flask(__name__)
CORS(app)

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

# 1. హోమ్ పేజీ రౌట్
@app.route("/")
def home():
    try:
        # BANK_LOGINS ట్యాబ్ నుండి బ్యాంకుల వివరాలు తీసుకోవడం
        login_sheet = sheet.worksheet("BANK_LOGINS")
        logins = login_sheet.get_all_records()
        # షీట్ నుండి వచ్చిన డేటాని కమాండ్ ప్రాంప్ట్‌లో ప్రింట్ చేయడం
        print("--- గూగుల్ షీట్ నుండి వచ్చిన డేటా ---")
        print(logins)
        bank_list = []
        for row in logins:
            # ఏ నేమ్‌తో ఉందో చెక్ చేయడం
            b_name = (
                row.get("Bank Name")
                or row.get("BANK NAME")
                or row.get("Bank_Name")
            )
            if b_name:
                bank_list.append(str(b_name).strip())

        print("--- డ్రాప్‌డౌన్ కోసం వచ్చిన బ్యాంకులు ---")
        print(bank_list)

        # డ్రాప్‌డౌన్ కోసం కేవలం బ్యాంక్ పేర్ల జాబితా (List) తయారు చేయడం
        bank_list = [row.get("Bank Name","Bank_Name").strip() for row in logins if row.get("Bank Name","Bank_Name")]
    except Exception as e:
        print("Error fetching bank list:", e)
        bank_list = []

    # బ్యాంకుల లిస్ట్‌ను index.html ఫైల్‌కి పంపడం
    return render_template("index.html", banks=bank_list)
# 2. శోధన API రౌట్
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

        # 1. గూగుల్ డ్రైవ్‌లోని ఫైల్ పేరుతో స్ప్రెడ్‌షీట్‌ను ఓపెన్ చేయడం
        spreadsheet = client.open('NOT_BANK_SENT_FORM')

        # 2. అందులోని 'NEW THRIFT DATA' అనే నిర్దిష్టమైన షీట్‌ను ఎంచుకోవడం
        sheet = spreadsheet.worksheet('NEW THRIFT DATA')

        # 3. ఆ షీట్ నుండి రికార్డులను పొందడం
        records = sheet.get_all_records()

        matched_record = None
        for row in records:
            raw_val = str(row.get('Aadhar Number') or row.get('ADHAR') or row.get('AADHAR') or '').split('.')[0].strip()
            if raw_val == input_id:
                matched_record = row
                break

        if matched_record:
            masked_id = f"XXXX-XXXX-{input_id[-4:]}"

            result = {
                'sno': matched_record.get('SNO'),
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
