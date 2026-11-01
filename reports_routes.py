import os
import openpyxl
from datetime import datetime
import pandas as pd
from flask import Blueprint, jsonify, request, send_file, session, render_template
import traceback

reports_bp = Blueprint('reports_bp', __name__)

# టెంప్లేట్ ఫైల్ ఉండే ఫోల్డర్ (assets/templates లేదా templates ఫోల్డర్)
TEMPLATE_DIR = 'assets/templates'
OUTPUT_DIR = 'assets/outputs'
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)

@reports_bp.route('/reports-dashboard')
def reports_dashboard():
    return render_template('reports_dashboard.html')

@reports_bp.route('/api/generate-thrift-report', methods=['POST'])
def generate_thrift_report():
    try:
        data = request.get_json() if request.is_json else request.form
        form_type = data.get('form_type', '').strip().upper() # FORM_III_A, FORM_III_B, FORM_V
        
        if not form_type:
            return jsonify({'status': 'error', 'message': 'దయచేసి రిపోర్ట్ రకాన్ని (Form Type) ఎంచుకోండి!'}), 400

        # గూగుల్ షీట్ నుండి డేటా లోడ్ చేయడం
        from app import get_sheets_client, SPREADSHEET_NAME
        
        client = get_sheets_client()
        sh = client.open(SPREADSHEET_NAME)
        ws_thrift = sh.worksheet("NEW THRIFT DATA")
        df_thrift = pd.DataFrame(ws_thrift.get_all_records())

        if df_thrift.empty:
            return jsonify({'status': 'error', 'message': 'గూగుల్ షీట్‌లో డేటా అందుబాటులో లేదు!'}), 400

        current_month = datetime.now().strftime("%B_%Y")
        template_filename = "Forms III, IV & V under Powerloom Thrift Scheme.xlsx"
        template_path = os.path.join(TEMPLATE_DIR, template_filename)

        if not os.path.exists(template_path):
            return jsonify({'status': 'error', 'message': f'టెంప్లేట్ ఫైల్ కనుగొనబడలేదు: {template_path}'}), 404

        wb = openpyxl.load_workbook(template_path)
        output_filename = ""

        # ==========================================
        # 1. FORM III-A (Powerloom Weavers Only)
        # ==========================================
        if form_type == 'FORM_III_A':
            df_weavers = df_thrift[df_thrift['Profession'].astype(str).str.strip().str.upper() == 'POWER LOOM WEAVER'].copy()
            df_weavers = df_weavers.reset_index(drop=True)

            df_weavers['Father_Husband_Name'] = df_weavers['Father Name'].fillna('')
            df_weavers['Full_Address'] = (
                df_weavers['H.NO. Residentional H.No'].fillna('').astype(str) + ", " +
                df_weavers['VILLAGE NAME'].fillna('').astype(str) + ", " +
                df_weavers['MANDAL NAME'].fillna('').astype(str) + ", " +
                df_weavers['DISTRICT NAME'].fillna('').astype(str)
            )
            df_weavers['SB_Account_Details'] = (
                "A/c: " + df_weavers['Bank Account Number'].fillna('').astype(str) +
                ", Bank: " + df_weavers['Bank Name'].fillna('').astype(str) +
                ", IFSC: " + df_weavers['IFSC code'].fillna('').astype(str)
            )
            df_weavers['Nominee_Details'] = (
                "Name: " + df_weavers['Nominee Name'].fillna('').astype(str) +
                ", Rel: " + df_weavers['Relation'].fillna('').astype(str)
            )
            df_weavers['Date_of_Enrollment'] = "01/08/2026"

            report_df = pd.DataFrame({
                'Sl. No.': range(1, len(df_weavers) + 1),
                'Name of the Powerloom Weaver': df_weavers['Name Of Benificiary'],
                'Whether belongs to MACS/SSI/ Individual': 'Individual',
                'If Yes, Indicate Name of MAC/SSI': '',
                'Geo Tagged Loom No.': df_weavers['GEO TAG NUMBER'],
                'Full Address': df_weavers['Full_Address'],
                "Father's/Husband's Name": df_weavers['Father_Husband_Name'],
                'Date of Enrollment': df_weavers['Date_of_Enrollment'],
                'Aadhar No.': df_weavers['Aadhar Number'],
                'Age as per Aadhar': df_weavers['Age'],
                'Caste': df_weavers['Caste'],
                'Mobile No.': df_weavers['Contact Number'],
                'Amount of Contribution (8% of wages)': df_weavers['8% Of Monthly Wages'],
                'Amount of 8% contribution required from Govt': df_weavers['8% Of Monthly Wages'],
                'S/B Account details': df_weavers['SB_Account_Details'],
                'Nominee details': df_weavers['Nominee_Details']
            })

            target_sheet = next((s for s in wb.sheetnames if "IIIA Revised" in s or "Form IIIA" in s), wb.sheetnames[0])
            ws = wb[target_sheet]
            start_row = 7
            output_filename = f"FORM_III_A_Powerloom_Weavers_Report_{current_month}.xlsx"

        # ==========================================
        # 2. FORM III-B (Ancillary Workers)
        # ==========================================
        elif form_type == 'FORM_III_B':
            df_ancillary = df_thrift[df_thrift['Profession'].astype(str).str.strip().str.upper() != 'POWER LOOM WEAVER'].copy()
            df_ancillary = df_ancillary.reset_index(drop=True)

            df_ancillary['Father_Husband_Name'] = df_ancillary['Father Name'].fillna('')
            df_ancillary['Full_Address'] = (
                df_ancillary['H.NO. Residentional H.No'].fillna('').astype(str) + ", " +
                df_ancillary['VILLAGE NAME'].fillna('').astype(str) + ", " +
                df_ancillary['MANDAL NAME'].fillna('').astype(str) + ", " +
                df_ancillary['DISTRICT NAME'].fillna('').astype(str)
            )
            df_ancillary['SB_Account_Details'] = (
                "A/c: " + df_ancillary['Bank Account Number'].fillna('').astype(str) +
                ", Bank: " + df_ancillary['Bank Name'].fillna('').astype(str) +
                ", IFSC: " + df_ancillary['IFSC code'].fillna('').astype(str)
            )
            df_ancillary['Nominee_Details'] = (
                "Name: " + df_ancillary['Nominee Name'].fillna('').astype(str) +
                ", Rel: " + df_ancillary['Relation'].fillna('').astype(str)
            )
            df_ancillary['Date_of_Enrollment'] = "01/08/2026"

            report_df = pd.DataFrame({
                'Sl. No.': range(1, len(df_ancillary) + 1),
                'Name of the Ancillary Worker': df_ancillary['Name Of Benificiary'],
                'Category of Ancillary Work': df_ancillary['Profession'],
                'Geo Tagged Loom No.': df_ancillary['GEO TAG NUMBER'],
                'Full Address': df_ancillary['Full_Address'],
                "Father's/Husband's Name": df_ancillary['Father_Husband_Name'],
                'Date of Enrollment': df_ancillary['Date_of_Enrollment'],
                'Aadhar No.': df_ancillary['Aadhar Number'],
                'Age as per Aadhar': df_ancillary['Age'],
                'Caste': df_ancillary['Caste'],
                'Mobile No.': df_ancillary['Contact Number'],
                'Amount of Contribution (8% of wages)': df_ancillary['8% Of Monthly Wages'],
                'Amount of 8% contribution required from Govt': df_ancillary['8% Of Monthly Wages'],
                'S/B Account details': df_ancillary['SB_Account_Details'],
                'Nominee details': df_ancillary['Nominee_Details']
            })

            target_sheet = next((s for s in wb.sheetnames if "IIIB Revised" in s or "Form IIIB" in s), wb.sheetnames[0])
            ws = wb[target_sheet]
            start_row = 7
            output_filename = f"FORM_III_B_Ancillary_Workers_Report_{current_month}.xlsx"

        # ==========================================
        # 3. FORM V (All Beneficiaries Total Register)
        # ==========================================
        elif form_type == 'FORM_V':
            df_v = df_thrift.copy().reset_index(drop=True)
            df_v['Village_Mandal'] = (
                df_v['VILLAGE NAME'].fillna('').astype(str) + ", " +
                df_v['MANDAL NAME'].fillna('').astype(str)
            )
            df_v['Date_of_Enrollment'] = "01/08/2026"

            report_df = pd.DataFrame({
                'Sl. No.': range(1, len(df_v) + 1),
                'Name of the beneficiary': df_v['Name Of Benificiary'],
                'Profession/Activity': df_v['Profession'],
                'Father / Husband Name': df_v['Father Name'].fillna(''),
                'Age': df_v['Age'],
                'Aadhar No.': df_v['Aadhar Number'],
                'Address - H.No. & Street': df_v['H.NO. Residentional H.No'].fillna(''),
                'Address - Village & Mandal': df_v['Village_Mandal'],
                'Address - District': df_v['DISTRICT NAME'].fillna(''),
                'Geo-tag Number': df_v['GEO TAG NUMBER'].fillna(''),
                'Mobile No.': df_v['Contact Number'].fillna(''),
                'Date of Enrollment into the Scheme': df_v['Date_of_Enrollment'],
                'SB A/c No.': df_v['Bank Account Number'].fillna(''),
                'Name of the Bank & Branch Location': df_v['Bank Name'].fillna(''),
                'IFSC Code': df_v['IFSC code'].fillna(''),
                'RD A/c No.1': df_v['RD 1'].fillna(''),
                'RD A/c No.2': df_v['RD 2'].fillna(''),
                'Declared wages earned per month': df_v['MONTHLY WAGES'].fillna(''),
                'Weaver contribution 8%': df_v['8% Of Monthly Wages'].fillna(''),
                'Govt. contribution 8%': df_v['8% Of Monthly Wages'].fillna('')
            })

            target_sheet = next((s for s in wb.sheetnames if "Form-V" in s or "Form V" in s), wb.sheetnames[0])
            ws = wb[target_sheet]
            start_row = 11
            output_filename = f"FORM_V_Register_Report_{current_month}.xlsx"
        else:
            return jsonify({'status': 'error', 'message': 'తప్పు ఫారమ్ రకం ఎంచుకోబడింది!'}), 400

        # ఎక్సెల్ గడులను నింపడం
        for row_idx, row in enumerate(report_df.itertuples(index=False)):
            current_r = start_row + row_idx
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=current_r, column=col_idx, value=value)

        output_path = os.path.join(OUTPUT_DIR, output_filename)
        wb.save(output_path)

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        error_msg = traceback.format_exc()
        print("ERROR:", error_msg)
        return jsonify({'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}), 500
