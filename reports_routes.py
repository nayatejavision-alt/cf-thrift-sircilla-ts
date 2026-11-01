import os
import openpyxl
from datetime import datetime
import pandas as pd
from flask import Blueprint, jsonify, request, send_file, session, render_template
import traceback

reports_bp = Blueprint('reports_bp', __name__)

TEMPLATE_DIR = 'assets/templates'
OUTPUT_DIR = 'assets/outputs'
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)

@reports_bp.route('/reports-dashboard')
def reports_dashboard():
    return render_template('reports_dashboard.html')

def safe_set_cell_value(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            top_left_cell.value = value
            return
    cell.value = value

@reports_bp.route('/api/generate-thrift-report', methods=['POST'])
def generate_thrift_report():
    try:
        data = request.get_json() if request.is_json else request.form
        form_type = data.get('form_type', '').strip().upper() 
        target_month = int(data.get('month', 1))
        
        if not form_type:
            return jsonify({'status': 'error', 'message': 'దయచేసి రిపోర్ట్ రకాన్ని ఎంచుకోండి!'}), 400

        from app import get_sheets_client, SPREADSHEET_NAME, MONTH_MAPPING
        
        client = get_sheets_client()
        sh = client.open(SPREADSHEET_NAME)
        
        # 1. త్రిఫ్ట్ మాస్టర్ డేటా లోడ్ చేయడం (Form III-A & III-B దీనిపైనే ఆధారపడతాయి)
        ws_thrift = sh.worksheet("NEW THRIFT DATA")
        df_thrift = pd.DataFrame(ws_thrift.get_all_records())

        if df_thrift.empty:
            return jsonify({'status': 'error', 'message': 'గూగుల్ షీట్‌లో త్రిఫ్ట్ డేటా అందుబాటులో లేదు!'}), 400

        selected_month_name, selected_year = MONTH_MAPPING.get(target_month, ('June', '2026'))
        current_month_str = f"{selected_month_name}_{selected_year}"
        
        template_filename = "Forms III, IV & V under Powerloom Thrift Scheme.xlsx"
        template_path = os.path.join(TEMPLATE_DIR, template_filename)

        if not os.path.exists(template_path):
            return jsonify({'status': 'error', 'message': f'టెంప్లేట్ ఫైల్ కనుగొనబడలేదు: {template_path}'}), 404

        wb = openpyxl.load_workbook(template_path)
        output_filename = ""
        report_df = pd.DataFrame()

        # అకౌంట్ నంబర్లు సైంటిఫిక్ నోటేషన్ కాకుండా టెక్స్ట్ ఫార్మాట్‌లో ఉంచే హెల్పర్
        def format_acc(val):
            if val is None:
                return ''
            val_str = str(val).split('.')[0].strip().strip("'")
            if val_str and val_str != 'nan' and val_str != '0':
                return f"'{val_str}"
            return ''

        # ==========================================
        # 1. FORM III-A (Powerloom Weavers Only - నేరుగా df_thrift వాడకం)
        # ==========================================
        if form_type == 'FORM_III_A':
            df_weavers = df_thrift[df_thrift['Profession'].astype(str).str.strip().str.upper() == 'POWER LOOM WEAVER'].copy()
            df_weavers = df_weavers.reset_index(drop=True)

            df_weavers['Father_Husband_Name'] = df_weavers.get('Father Name', '').fillna('')
            df_weavers['Full_Address'] = (
                df_weavers.get('H.NO. Residentional H.No', '').fillna('').astype(str) + ", " +
                df_weavers.get('VILLAGE NAME', '').fillna('').astype(str) + ", " +
                df_weavers.get('MANDAL NAME', '').fillna('').astype(str) + ", " +
                df_weavers.get('DISTRICT NAME', '').fillna('').astype(str)
            )
            df_weavers['SB_Account_Details'] = (
                "A/c: " + df_weavers.get('Bank Account Number', '').fillna('').astype(str) +
                ", Bank: " + df_weavers.get('Bank Name', '').fillna('').astype(str) +
                ", IFSC: " + df_weavers.get('IFSC code', '').fillna('').astype(str)
            )
            df_weavers['Nominee_Details'] = (
                "Name: " + df_weavers.get('Nominee Name', '').fillna('').astype(str) +
                ", Rel: " + df_weavers.get('Relation', '').fillna('').astype(str)
            )
            df_weavers['Date_of_Enrollment'] = "01/06/2026"

            report_df = pd.DataFrame({
                'Sl. No.': range(1, len(df_weavers) + 1),
                'Name of the Powerloom Weaver': df_weavers.get('Name Of Benificiary', ''),
                'Whether belongs to MACS/SSI/ Individual': 'Individual',
                'If Yes, Indicate Name of MAC/SSI': '',
                'Geo Tagged Loom No.': df_weavers.get('GEO TAG NUMBER', ''),
                'Full Address': df_weavers['Full_Address'],
                "Father's/Husband's Name": df_weavers['Father_Husband_Name'],
                'Date of Enrollment': df_weavers['Date_of_Enrollment'],
                'Aadhar No.': df_weavers.get('Aadhar Number', '').apply(format_acc),
                'Age as per Aadhar': df_weavers.get('Age', ''),
                'Caste': df_weavers.get('Caste', 'BC'),
                'Mobile No.': df_weavers.get('Contact Number', '').apply(format_acc),
                'Amount of Contribution (8% of wages)': df_weavers.get('8% Of Monthly Wages', ''),
                'Amount of 8% contribution required from Govt': df_weavers.get('8% Of Monthly Wages', ''),
                'S/B Account details': df_weavers['SB_Account_Details'],
                'Nominee details': df_weavers['Nominee_Details']
            })

            target_sheet = next((s for s in wb.sheetnames if "IIIA Revised" in s or "Form IIIA" in s), wb.sheetnames[0])
            ws = wb[target_sheet]
            start_row = 7
            output_filename = f"FORM_III_A_Powerloom_Weavers_Report_{current_month_str}.xlsx"

        # ==========================================
        # 2. FORM III-B (Ancillary Workers - నేరుగా df_thrift వాడకం)
        # ==========================================
        elif form_type == 'FORM_III_B':
            df_ancillary = df_thrift[df_thrift['Profession'].astype(str).str.strip().str.upper() != 'POWER LOOM WEAVER'].copy()
            df_ancillary = df_ancillary.reset_index(drop=True)

            df_ancillary['Father_Husband_Name'] = df_ancillary.get('Father Name', '').fillna('')
            df_ancillary['Full_Address'] = (
                df_ancillary.get('H.NO. Residentional H.No', '').fillna('').astype(str) + ", " +
                df_ancillary.get('VILLAGE NAME', '').fillna('').astype(str) + ", " +
                df_ancillary.get('MANDAL NAME', '').fillna('').astype(str) + ", " +
                df_ancillary.get('DISTRICT NAME', '').fillna('').astype(str)
            )
            df_ancillary['SB_Account_Details'] = (
                "A/c: " + df_ancillary.get('Bank Account Number', '').fillna('').astype(str) +
                ", Bank: " + df_ancillary.get('Bank Name', '').fillna('').astype(str) +
                ", IFSC: " + df_ancillary.get('IFSC code', '').fillna('').astype(str)
            )
            df_ancillary['Nominee_Details'] = (
                "Name: " + df_ancillary.get('Nominee Name', '').fillna('').astype(str) +
                ", Rel: " + df_ancillary.get('Relation', '').fillna('').astype(str)
            )
            df_ancillary['Date_of_Enrollment'] = "01/06/2026"

            report_df = pd.DataFrame({
                'Sl. No.': range(1, len(df_ancillary) + 1),
                'Name of the Ancillary Worker': df_ancillary.get('Name Of Benificiary', ''),
                'Category of Ancillary Work': df_ancillary.get('Profession', ''),
                'Geo Tagged Loom No.': df_ancillary.get('GEO TAG NUMBER', ''),
                'Full Address': df_ancillary['Full_Address'],
                "Father's/Husband's Name": df_ancillary['Father_Husband_Name'],
                'Date of Enrollment': df_ancillary['Date_of_Enrollment'],
                'Aadhar No.': df_ancillary.get('Aadhar Number', '').apply(format_acc),
                'Age as per Aadhar': df_ancillary.get('Age', ''),
                'Caste': df_ancillary.get('Caste', 'BC'),
                'Mobile No.': df_ancillary.get('Contact Number', '').apply(format_acc),
                'Amount of Contribution (8% of wages)': df_ancillary.get('8% Of Monthly Wages', ''),
                'Amount of 8% contribution required from Govt': df_ancillary.get('8% Of Monthly Wages', ''),
                'S/B Account details': df_ancillary['SB_Account_Details'],
                'Nominee details': df_ancillary['Nominee_Details']
            })

            target_sheet = next((s for s in wb.sheetnames if "IIIB Revised" in s or "Form IIIB" in s), wb.sheetnames[0])
            ws = wb[target_sheet]
            start_row = 7
            output_filename = f"FORM_III_B_Ancillary_Workers_Report_{current_month_str}.xlsx"

        # ==========================================
        # 3. FORM V (All Beneficiaries Total Register - కేవలం దీనికి మాత్రమే df_filtered / ట్రాన్సాక్షన్ ఫిల్టర్ వర్తిస్తుంది)
        # ==========================================
        elif form_type == 'FORM_V':
            try:
                ws_trans = sh.worksheet("Transactions_Year1")
                df_trans = pd.DataFrame(ws_trans.get_all_records())
            except Exception:
                df_trans = pd.DataFrame()

            if not df_trans.empty:
                df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0)
                df_trans['MONTH_NUMBER'] = pd.to_numeric(df_trans['MONTH_NUMBER'], errors='coerce')
                
                df_month_trans = df_trans[(df_trans['MONTH_NUMBER'] == target_month) & (df_trans['Govt_Share'] > 0)]
                
                df_filtered = pd.merge(
                    df_month_trans,
                    df_thrift,
                    left_on='RD ACCOUNT_1 (Beneficiary Contribution)',
                    right_on='RD 1',
                    how='inner',
                    suffixes=('', '_THRIFT')
                )
                df_filtered = df_filtered.drop_duplicates(subset=['RD 1']).reset_index(drop=True)
            else:
                df_filtered = pd.DataFrame()

            # ఒకవేళ ఆ నెలలో ట్రాన్సాక్షన్ డేటా లేకపోతే మొత్తం మాస్టర్ డేటా తీసుకోవడానికి లేదా ఖాళీగా ఉంచడానికి
            if df_filtered.empty:
                df_filtered = df_thrift.copy()

            df_v = df_filtered.copy().reset_index(drop=True)
            
            if not df_v.empty:
                df_v['Village_Mandal'] = (
                    df_v.get('VILLAGE NAME', '').fillna('').astype(str) + ", " +
                    df_v.get('MANDAL NAME', '').fillna('').astype(str)
                )
                df_v['Date_of_Enrollment'] = "01/06/2026"

                report_df = pd.DataFrame({
                    'Sl. No.': range(1, len(df_v) + 1),
                    'Name of the beneficiary': df_v.get('Name Of Benificiary', ''),
                    'Profession/Activity': df_v.get('Profession', ''),
                    'Father / Husband Name': df_v.get('Father Name', '').fillna(''),
                    'Age': df_v.get('Age', ''),
                    'Aadhar No.': df_v.get('Aadhar Number', '').apply(format_acc),
                    'Address - H.No. & Street': df_v.get('H.NO. Residentional H.No', '').fillna(''),
                    'Address - Village & Mandal': df_v['Village_Mandal'],
                    'Address - District': df_v.get('DISTRICT NAME', '').fillna(''),
                    'Geo-tag Number': df_v.get('GEO TAG NUMBER', '').fillna(''),
                    'Mobile No.': df_v.get('Contact Number', '').apply(format_acc),
                    'Date of Enrollment into the Scheme': df_v['Date_of_Enrollment'],
                    'SB A/c No.': df_v.get('Bank Account Number', '').apply(format_acc),
                    'Name of the Bank & Branch Location': df_v.get('Bank Name', '').fillna(''),
                    'IFSC Code': df_v.get('IFSC code', '').fillna(''),
                    'RD A/c No.1': df_v.get('RD 1', '').apply(format_acc),
                    'RD A/c No.2': df_v.get('RD 2', '').apply(format_acc),
                    'Declared wages earned per month': df_v.get('MONTHLY WAGES', '').fillna(''),
                    'Weaver contribution 8%': df_v.get('8% Of Monthly Wages', '').fillna(''),
                    'Govt. contribution 8%': df_v.get('8% Of Monthly Wages', '').fillna('')
                })

            target_sheet = next((s for s in wb.sheetnames if "Form-V" in s or "Form V" in s), wb.sheetnames[0])
            ws = wb[target_sheet]
            start_row = 11
            output_filename = f"FORM_V_Register_Report_{current_month_str}.xlsx"
        else:
            return jsonify({'status': 'error', 'message': 'తప్పు ఫారమ్ రకం ఎంచుకోబడింది!'}), 400

        if report_df.empty:
            return jsonify({'status': 'error', 'message': 'రిపోర్ట్ చేయడానికి ఎలాంటి రికార్డులు కనుగొనబడలేదు!'}), 400

        # ఎక్సెల్ గడులను సురక్షితంగా నింపడం
        for row_idx, row in enumerate(report_df.itertuples(index=False)):
            current_r = start_row + row_idx
            for col_idx, value in enumerate(row, start=1):
                safe_set_cell_value(ws, current_r, col_idx, value)

        output_path = os.path.join(OUTPUT_DIR, output_filename)
        wb.save(output_path)

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        error_msg = traceback.format_exc()
        print("ERROR:", error_msg)
        return jsonify({'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}), 500
