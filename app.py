import os
import json
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from google.oauth2 import service_account
import google.auth
import gspread

app = Flask(__name__)
CORS(app)

def home():
    return render_template('index.html')

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
            
    client = gspread.authorize(credentials)
    return client

@app.route('/api/search', methods=['POST'])
def search_beneficiary():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': 'చెల్లుబాటు అయ్యే వివరాలు పంపలేదు.'}), 400

        input_aadhar = str(data.get('aadhar', '')).strip().replace('-', '').replace(' ', '')

        if not input_aadhar or len(input_aadhar) != 12 or not input_aadhar.isdigit():
            return jsonify({'success': False, 'message': 'దయచేసి సరైన 12 అంకెల సంఖ్యను ఎంటర్ చేయండి.'}), 400

        client = get_sheets_client()
        # షీట్ ఓపెన్ చేయడం
        spreadsheet = client.open('NEW THRIFT DATA')
        sheet = spreadsheet.sheet1
        
        # గూగుల్ షీట్ నుండి రికార్డులు పొందడం
        records = sheet.get_all_records()

        matched_record = None
        for row in records:
            # కాలమ్ పేర్లు మీ గూగుల్ షీట్ ఆధారంగా చెక్ చేసుకోండి
            raw_val = str(row.get('Aadhar Number') or row.get('ADHAR') or row.get('AADHAR') or '').split('.')[0].strip()
            if raw_val == input_aadhar:
                matched_record = row
                break

        if matched_record:
            masked_aadhar = f"XXXX-XXXX-{input_aadhar[-4:]}"

            result = {
                'sno': matched_record.get('SNO'),
                'masked_aadhar': masked_aadhar,
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
        # ఎర్రర్ వివరాలను సరిగ్గా ప్రింట్ చేయడం
        print(f"Error details: {repr(e)}")
        return jsonify({'success': False, 'message': f'సర్వర్ లోపం సంభవించింది: {str(e)}'}), 500
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
