import os
import json
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from google.oauth2 import service_account
import google.auth
import gspread

app = Flask(__name__)
CORS(app)

def get_sheets_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # 1. Render లో ఉన్న GOOGLE_CREDENTIALS_JSON Environment Variable ని చెక్ చేయడం
    credentials_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    
    if credentials_json_str:
        # JSON స్ట్రింగ్ ని డిక్షనరీగా మార్చి సర్వీస్ అకౌంట్ క్రెడెన్షియల్స్ పొందడం
        info = json.loads(credentials_json_str)
        credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    else:
        # క్రెడెన్షియల్ ఫైల్ పాత్ ఉందో లేదో చెక్ చేయడం (Local Development కోసం)
        cred_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        if cred_path and os.path.exists(cred_path):
            credentials = service_account.Credentials.from_service_account_file(cred_path, scopes=scopes)
        else:
            # మిగతా డిఫాల్ట్ ఆథెంటికేషన్
            credentials, _ = google.auth.default(scopes=scopes)
            
    client = gspread.authorize(credentials)
    return client

# 1. హోమ్ పేజీ రౌట్ (HTML సెర్చ్ ఫారమ్ చూపిస్తుంది)
@app.route('/')
def home():
    return render_template('index.html')

# 2. శోధన API రౌట్
@app.route('/api/search', methods=['POST'])
def search_beneficiary():
    data = request.get_json()
    input_aadhar = str(data.get('aadhar', '')).strip().replace('-', '').replace(' ', '')

    if not input_aadhar or len(input_aadhar) != 12 or not input_aadhar.isdigit():
        return jsonify({'success': False, 'message': 'దయచేసి సరైన 12 అంకెల సంఖ్యను ఎంటర్ చేయండి.'}), 400

    try:
        client = get_sheets_client()
        sheet = client.open('NEW THRIFT DATA').sheet1 
        records = sheet.get_all_records()

        matched_record = None
        for row in records:
            raw_val = str(row.get('Aadhar Number') or row.get('ADHAR') or '').split('.')[0].strip()
            if raw_val == input_aadhar:
                matched_record = row
                break

        if matched_record:
            masked_aadhar = f"XXXX-XXXX-{input_aadhar[-4:]}"

            result = {
                'sno': matched_record.get('SNO'),
                'masked_aadhar': masked_aadhar,
                'name': matched_record.get('Name Of Benificiary'),
                'father_name': matched_record.get('Father Name'),
                'village': matched_record.get('VILLAGE NAME'),
                'mandal': matched_record.get('MANDAL NAME'),
                'bank_name': matched_record.get('Bank Name'),
                'rd1_status': matched_record.get('RD 1'),
                'rd2_status': matched_record.get('RD 2'),
                'wages': matched_record.get('MONTHLY WAGES')
            }
            return jsonify({'success': True, 'data': result})
        else:
            return jsonify({'success': False, 'message': 'సదరు సంఖ్యతో ఎలాంటి వివరాలు లభించలేదు.'})

    except Exception as e:
        return jsonify({'success': False, 'message': f'సర్వర్ లోపం సంభవించింది: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
