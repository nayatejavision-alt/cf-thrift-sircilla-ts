import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import google.auth
import gspread

app = Flask(__name__)
CORS(app)  # బ్లాగ్ నుండి API కి రిక్వెస్ట్‌లు అనుమతించడానికి

def get_sheets_client():
    # Application Default Credentials (ADC) / Workload Identity ఉపయోగించి కనెక్ట్ అవ్వడం
    credentials, project = google.auth.default(
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    client = gspread.authorize(credentials)
    return client

@app.route('/api/search', methods=['POST'])
def search_beneficiary():
    data = request.get_json()
    input_aadhar = str(data.get('aadhar', '')).strip().replace('-', '').replace(' ', '')

    # ఆధార్ సంఖ్య వ్యాలిడేషన్ (12 అంకెలు ఉండాలి)
    if not input_aadhar or len(input_aadhar) != 12 or not input_aadhar.isdigit():
        return jsonify({'success': False, 'message': 'దయచేసి సరైన 12 అంకెల ఆధార్ సంఖ్యను ఎంటర్ చేయండి.'}), 400

    try:
        client = get_sheets_client()
        # మీ Google Sheet పేరు పక్కాగా 'NEW THRIFT DATA' ఉందో లేదో సరిచూసుకోండి
        sheet = client.open('NEW THRIFT DATA').sheet1 
        records = sheet.get_all_records()

        matched_record = None
        for row in records:
            # షీట్‌లోని 'Aadhar Number' లేదా 'ADHAR' కాలమ్‌ను తనిఖీ చేయడం
            raw_val = str(row.get('Aadhar Number') or row.get('ADHAR') or '').split('.')[0].strip()
            if raw_val == input_aadhar:
                matched_record = row
                break

        if matched_record:
            # ఆధార్ నంబర్‌ను మాస్క్ చేయడం (ప్రైవసీ రక్షణ)
            masked_aadhar = f"XXXX-XXXX-{input_aadhar[-4:]}"

            # బ్లాగ్‌లో చూపించడానికి సురక్షితమైన వివరాలు
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
            return jsonify({'success': False, 'message': 'సదరు ఆధార్ సంఖ్యతో ఎలాంటి వివరాలు లభించలేదు.'}), 444

    except Exception as e:
        return jsonify({'success': False, 'message': f'సర్వర్ లోపం సంభవించింది: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
