import os
import gc
import openpyxl
from datetime import datetime
import pandas as pd
from flask import Blueprint, jsonify, request, send_file, render_template
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

def format_acc(val):
    if val is None:
        return ''
    val_str = str(val).split('.')[0].strip().strip("'")
    if val_str and val_str != 'nan' and val_str != '0':
        return f"'{val_str}"
    return ''

# ==========================================
# 1. FORM III-A (Powerloom Weavers Report)
# ==========================================
@reports_bp.route('/api/generate-form-iiia', methods=['POST'])
def generate_form_iiia():
    try:
        data = request.get_json() if request.is_json else request.form
        target_month = int(data.get('month', 1))
        
        from app import get_sheets_client, SPREADSHEET_NAME, MONTH_MAPPING
        client = get_sheets_client()
        sh = client.open(SPREADSHEET_NAME)
        
        ws_thrift = sh.worksheet("NEW THRIFT DATA")
        thrift_data = ws_thrift.get_all_records()
        if not thrift_data:
            return jsonify({'status': 'error', 'message': 'గూగుల్ షీట్‌లో త్రిఫ్ట్ డేటా అందుబాటులో లేదు!'}), 400
            
        df_thrift = pd.DataFrame(thrift_data)
        del thrift_data
        gc.collect()

        selected_month_name, selected_year = MONTH_MAPPING.get(target_month, ('June', '2026'))
        current_month_str = f"{selected_month_name}_{selected_year}"
        
        template_path = os.path.join(TEMPLATE_DIR, "Forms III, IV & V under Powerloom Thrift Scheme.xlsx")
        wb = openpyxl.load_workbook(template_path)

        profession_col = 'Profession' if 'Profession' in df_thrift.columns else df_thrift.columns[0]
        df_weavers = df_thrift[df_thrift[profession_col].astype(str).str.strip().str.upper() == 'POWER LOOM WEAVER'].copy()
        df_weavers = df_weavers.reset_index(drop=True)

        father_col = 'Father Name' if 'Father Name' in df_weavers.columns else ''
        hno_col = 'H.NO. Residentional H.No' if 'H.NO. Residentional H.No' in df_weavers.columns else ''
        vill_col = 'VILLAGE NAME' if 'VILLAGE NAME' in df_weavers.columns else ''
        mandal_col = 'MANDAL NAME' if 'MANDAL NAME' in df_weavers.columns else ''
        dist_col = 'DISTRICT NAME' if 'DISTRICT NAME' in df_weavers.columns else ''
        acc_col = 'Bank Account Number' if 'Bank Account Number' in df_weavers.columns else ''
        bank_col = 'Bank Name' if 'Bank Name' in df_weavers.columns else ''
        ifsc_col = 'IFSC code' if 'IFSC code' in df_weavers.columns else ''
        nom_col = 'Nominee Name' if 'Nominee Name' in df_weavers.columns else ''
        rel_col = 'Relation' if 'Relation' in df_weavers.columns else ''

        df_weavers['Father_Husband_Name'] = df_weavers[father_col].fillna('') if father_col else ''
        df_weavers['Full_Address'] = (
            (df_weavers[hno_col].fillna('').astype(str) + ", ") if hno_col else ""
        ) + (
            (df_weavers[vill_col].fillna('').astype(str) + ", ") if vill_col else ""
        ) + (
            (df_weavers[mandal_col].fillna('').astype(str) + ", ") if mandal_col else ""
        ) + (
            df_weavers[dist_col].fillna('').astype(str) if dist_col else ""
        )

        df_weavers['SB_Account_Details'] = (
            "A/c: " + df_weavers[acc_col].fillna('').astype(str) +
            ", Bank: " + df_weavers[bank_col].fillna('').astype(str) +
            ", IFSC: " + df_weavers[ifsc_col].fillna('').astype(str) if (acc_col and bank_col and ifsc_col) else ''
        )
        df_weavers['Nominee_Details'] = (
            "Name: " + df_weavers[nom_col].fillna('').astype(str) +
            ", Rel: " + df_weavers[rel_col].fillna('').astype(str) if (nom_col and rel_col) else ''
        )
        df_weavers['Date_of_Enrollment'] = "01/06/2026"

        name_col = 'Name Of Benificiary' if 'Name Of Benificiary' in df_weavers.columns else df_weavers.columns[1]
        geo_col = 'GEO TAG NUMBER' if 'GEO TAG NUMBER' in df_weavers.columns else ''
        aadhar_col = 'Aadhar Number' if 'Aadhar Number' in df_weavers.columns else ''
        age_col = 'Age' if 'Age' in df_weavers.columns else ''
        caste_col = 'Caste' if 'Caste' in df_weavers.columns else ''
        mobile_col = 'Contact Number' if 'Contact Number' in df_weavers.columns else ''
        wages_col = '8% Of Monthly Wages' if '8% Of Monthly Wages' in df_weavers.columns else ''

        report_df = pd.DataFrame({
            'Sl. No.': range(1, len(df_weavers) + 1),
            'Name of the Powerloom Weaver': df_weavers[name_col],
            'Whether belongs to MACS/SSI/ Individual': 'Individual',
            'If Yes, Indicate Name of MAC/SSI': '',
            'Geo Tagged Loom No.': df_weavers[geo_col].fillna('') if geo_col else '',
            'Full Address': df_weavers['Full_Address'],
            "Father's/Husband's Name": df_weavers['Father_Husband_Name'],
            'Date of Enrollment': df_weavers['Date_of_Enrollment'],
            'Aadhar No.': df_weavers[aadhar_col].apply(format_acc) if aadhar_col else '',
            'Age as per Aadhar': df_weavers[age_col] if age_col else '',
            'Caste': df_weavers[caste_col].fillna('BC') if caste_col else 'BC',
            'Mobile No.': df_weavers[mobile_col].apply(format_acc) if mobile_col else '',
            'Amount of Contribution (8% of wages)': df_weavers[wages_col] if wages_col else '',
            'Amount of 8% contribution required from Govt': df_weavers[wages_col] if wages_col else '',
            'S/B Account details': df_weavers['SB_Account_Details'],
            'Nominee details': df_weavers['Nominee_Details']
        })

        if report_df.empty:
            return jsonify({'status': 'error', 'message': 'రిపోర్ట్ చేయడానికి ఎలాంటి రికార్డులు కనుగొనబడలేదు!'}), 400

        target_sheet = next((s for s in wb.sheetnames if "IIIA Revised" in s or "Form IIIA" in s), wb.sheetnames[0])
        ws = wb[target_sheet]
        start_row = 7
        output_filename = f"FORM_III_A_Powerloom_Weavers_Report_{current_month_str}.xlsx"

        for row_idx, row in enumerate(report_df.itertuples(index=False)):
            current_r = start_row + row_idx
            for col_idx, value in enumerate(row, start=1):
                safe_set_cell_value(ws, current_r, col_idx, value)

        output_path = os.path.join(OUTPUT_DIR, output_filename)
        wb.save(output_path)
        
        del df_thrift, df_weavers
        gc.collect()

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        print("ERROR IN FORM_III_A:", traceback.format_exc())
        return jsonify({'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}), 500


# ==========================================
# 2. FORM III-B (Ancillary Workers Report)
# ==========================================
@reports_bp.route('/api/generate-form-iiib', methods=['POST'])
def generate_form_iiib():
    try:
        data = request.get_json() if request.is_json else request.form
        target_month = int(data.get('month', 1))
        
        from app import get_sheets_client, SPREADSHEET_NAME, MONTH_MAPPING
        client = get_sheets_client()
        sh = client.open(SPREADSHEET_NAME)
        
        ws_thrift = sh.worksheet("NEW THRIFT DATA")
        thrift_data = ws_thrift.get_all_records()
        if not thrift_data:
            return jsonify({'status': 'error', 'message': 'గూగుల్ షీట్‌లో త్రిఫ్ట్ డేటా అందుబాటులో లేదు!'}), 400
            
        df_thrift = pd.DataFrame(thrift_data)
        del thrift_data
        gc.collect()

        selected_month_name, selected_year = MONTH_MAPPING.get(target_month, ('June', '2026'))
        current_month_str = f"{selected_month_name}_{selected_year}"
        
        template_path = os.path.join(TEMPLATE_DIR, "Forms III, IV & V under Powerloom Thrift Scheme.xlsx")
        wb = openpyxl.load_workbook(template_path)

        profession_col = 'Profession' if 'Profession' in df_thrift.columns else df_thrift.columns[0]
        df_ancillary = df_thrift[df_thrift[profession_col].astype(str).str.strip().str.upper() != 'POWER LOOM WEAVER'].copy()
        df_ancillary = df_ancillary.reset_index(drop=True)

        father_col = 'Father Name' if 'Father Name' in df_ancillary.columns else ''
        hno_col = 'H.NO. Residentional H.No' if 'H.NO. Residentional H.No' in df_ancillary.columns else ''
        vill_col = 'VILLAGE NAME' if 'VILLAGE NAME' in df_ancillary.columns else ''
        mandal_col = 'MANDAL NAME' if 'MANDAL NAME' in df_ancillary.columns else ''
        dist_col = 'DISTRICT NAME' if 'DISTRICT NAME' in df_ancillary.columns else ''
        acc_col = 'Bank Account Number' if 'Bank Account Number' in df_ancillary.columns else ''
        bank_col = 'Bank Name' if 'Bank Name' in df_ancillary.columns else ''
        ifsc_col = 'IFSC code' if 'IFSC code' in df_ancillary.columns else ''
        nom_col = 'Nominee Name' if 'Nominee Name' in df_ancillary.columns else ''
        rel_col = 'Relation' if 'Relation' in df_ancillary.columns else ''

        df_ancillary['Father_Husband_Name'] = df_ancillary[father_col].fillna('') if father_col else ''
        df_ancillary['Full_Address'] = (
            (df_ancillary[hno_col].fillna('').astype(str) + ", ") if hno_col else ""
        ) + (
            (df_ancillary[vill_col].fillna('').astype(str) + ", ") if vill_col else ""
        ) + (
            (df_ancillary[mandal_col].fillna('').astype(str) + ", ") if mandal_col else ""
        ) + (
            df_ancillary[dist_col].fillna('').astype(str) if dist_col else ""
        )

        df_ancillary['SB_Account_Details'] = (
            "A/c: " + df_ancillary[acc_col].fillna('').astype(str) +
            ", Bank: " + df_ancillary[bank_col].fillna('').astype(str) +
            ", IFSC: " + df_ancillary[ifsc_col].fillna('').astype(str) if (acc_col and bank_col and ifsc_col) else ''
        )
        df_ancillary['Nominee_Details'] = (
            "Name: " + df_ancillary[nom_col].fillna('').astype(str) +
            ", Rel: " + df_ancillary[rel_col].fillna('').astype(str) if (nom_col and rel_col) else ''
        )
        df_ancillary['Date_of_Enrollment'] = "01/06/2026"

        name_col = 'Name Of Benificiary' if 'Name Of Benificiary' in df_ancillary.columns else df_ancillary.columns[1]
        geo_col = 'GEO TAG NUMBER' if 'GEO TAG NUMBER' in df_ancillary.columns else ''
        aadhar_col = 'Aadhar Number' if 'Aadhar Number' in df_ancillary.columns else ''
        age_col = 'Age' if 'Age' in df_ancillary.columns else ''
        caste_col = 'Caste' if 'Caste' in df_ancillary.columns else ''
        mobile_col = 'Contact Number' if 'Contact Number' in df_ancillary.columns else ''
        wages_col = '8% Of Monthly Wages' if '8% Of Monthly Wages' in df_ancillary.columns else ''

        report_df = pd.DataFrame({
            'Sl. No.': range(1, len(df_ancillary) + 1),
            'Name of the Ancillary Worker': df_ancillary[name_col],
            'Category of Ancillary Work': df_ancillary[profession_col],
            'Geo Tagged Loom No.': df_ancillary[geo_col].fillna('') if geo_col else '',
            'Full Address': df_ancillary['Full_Address'],
            "Father's/Husband's Name": df_ancillary['Father_Husband_Name'],
            'Date of Enrollment': df_ancillary['Date_of_Enrollment'],
            'Aadhar No.': df_ancillary[aadhar_col].apply(format_acc) if aadhar_col else '',
            'Age as per Aadhar': df_ancillary[age_col] if age_col else '',
            'Caste': df_ancillary[caste_col].fillna('BC') if caste_col else 'BC',
            'Mobile No.': df_ancillary[mobile_col].apply(format_acc) if mobile_col else '',
            'Amount of Contribution (8% of wages)': df_ancillary[wages_col] if wages_col else '',
            'Amount of 8% contribution required from Govt': df_ancillary[wages_col] if wages_col else '',
            'S/B Account details': df_ancillary['SB_Account_Details'],
            'Nominee details': df_ancillary['Nominee_Details']
        })

        if report_df.empty:
            return jsonify({'status': 'error', 'message': 'రిపోర్ట్ చేయడానికి ఎలాంటి రికార్డులు కనుగొనబడలేదు!'}), 400

        target_sheet = next((s for s in wb.sheetnames if "IIIB Revised" in s or "Form IIIB" in s), wb.sheetnames[0])
        ws = wb[target_sheet]
        start_row = 7
        output_filename = f"FORM_III_B_Ancillary_Workers_Report_{current_month_str}.xlsx"

        for row_idx, row in enumerate(report_df.itertuples(index=False)):
            current_r = start_row + row_idx
            for col_idx, value in enumerate(row, start=1):
                safe_set_cell_value(ws, current_r, col_idx, value)

        output_path = os.path.join(OUTPUT_DIR, output_filename)
        wb.save(output_path)
        
        del df_thrift, df_ancillary
        gc.collect()

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        print("ERROR IN FORM_III_B:", traceback.format_exc())
        return jsonify({'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}), 500


# ==========================================
# 3. FORM V (All Beneficiaries Total Register)
# ==========================================
@reports_bp.route('/api/generate-form-v', methods=['POST'])
def generate_form_v():
    try:
        data = request.get_json() if request.is_json else request.form
        target_month = int(data.get('month', 1))
        
        from app import get_sheets_client, SPREADSHEET_NAME, MONTH_MAPPING
        client = get_sheets_client()
        sh = client.open(SPREADSHEET_NAME)
        
        ws_thrift = sh.worksheet("NEW THRIFT DATA")
        thrift_data = ws_thrift.get_all_records()
        if not thrift_data:
            return jsonify({'status': 'error', 'message': 'గూగుల్ షీట్‌లో త్రిఫ్ట్ డేటా అందుబాటులో లేదు!'}), 400
            
        df_thrift = pd.DataFrame(thrift_data)
        del thrift_data
        gc.collect()

        selected_month_name, selected_year = MONTH_MAPPING.get(target_month, ('June', '2026'))
        current_month_str = f"{selected_month_name}_{selected_year}"
        
        template_path = os.path.join(TEMPLATE_DIR, "Forms III, IV & V under Powerloom Thrift Scheme.xlsx")
        wb = openpyxl.load_workbook(template_path)

        try:
            ws_trans = sh.worksheet("Transactions_Year1")
            trans_data = ws_trans.get_all_records()
            df_trans = pd.DataFrame(trans_data) if trans_data else pd.DataFrame()
            del trans_data
        except Exception:
            df_trans = pd.DataFrame()

        df_filtered = pd.DataFrame()
        if not df_trans.empty and 'Govt_Share' in df_trans.columns and 'MONTH_NUMBER' in df_trans.columns:
            df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0)
            df_trans['MONTH_NUMBER'] = pd.to_numeric(df_trans['MONTH_NUMBER'], errors='coerce')
            
            df_month_trans = df_trans[(df_trans['MONTH_NUMBER'] == target_month) & (df_trans['Govt_Share'] > 0)]
            del df_trans
            gc.collect()
            
            if not df_month_trans.empty and 'RD ACCOUNT_1 (Beneficiary Contribution)' in df_month_trans.columns and 'RD 1' in df_thrift.columns:
                df_filtered = pd.merge(
                    df_month_trans,
                    df_thrift,
                    left_on='RD ACCOUNT_1 (Beneficiary Contribution)',
                    right_on='RD 1',
                    how='inner',
                    suffixes=('', '_THRIFT')
                )
                df_filtered = df_filtered.drop_duplicates(subset=['RD 1']).reset_index(drop=True)
            del df_month_trans
            gc.collect()

        if df_filtered.empty:
            df_filtered = df_thrift.copy()

        df_v = df_filtered.copy().reset_index(drop=True)
        del df_filtered
        gc.collect()
        
        vill_col = 'VILLAGE NAME' if 'VILLAGE NAME' in df_v.columns else ''
        mandal_col = 'MANDAL NAME' if 'MANDAL NAME' in df_v.columns else ''
        df_v['Village_Mandal'] = (
            df_v[vill_col].fillna('').astype(str) + ", " + df_v[mandal_col].fillna('').astype(str) if (vill_col and mandal_col) else ''
        )
        df_v['Date_of_Enrollment'] = "01/06/2026"

        name_col = 'Name Of Benificiary' if 'Name Of Benificiary' in df_v.columns else df_v.columns[1]
        prof_col = 'Profession' if 'Profession' in df_v.columns else ''
        father_col = 'Father Name' if 'Father Name' in df_v.columns else ''
        age_col = 'Age' if 'Age' in df_v.columns else ''
        aadhar_col = 'Aadhar Number' if 'Aadhar Number' in df_v.columns else ''
        hno_col = 'H.NO. Residentional H.No' if 'H.NO. Residentional H.No' in df_v.columns else ''
        dist_col = 'DISTRICT NAME' if 'DISTRICT NAME' in df_v.columns else ''
        geo_col = 'GEO TAG NUMBER' if 'GEO TAG NUMBER' in df_v.columns else ''
        mobile_col = 'Contact Number' if 'Contact Number' in df_v.columns else ''
        acc_col = 'Bank Account Number' if 'Bank Account Number' in df_v.columns else ''
        bank_col = 'Bank Name' if 'Bank Name' in df_v.columns else ''
        ifsc_col = 'IFSC code' if 'IFSC code' in df_v.columns else ''
        rd1_col = 'RD 1' if 'RD 1' in df_v.columns else ''
        rd2_col = 'RD 2' if 'RD 2' in df_v.columns else ''
        wages_col = 'MONTHLY WAGES' if 'MONTHLY WAGES' in df_v.columns else ''
        percent_wages = '8% Of Monthly Wages' if '8% Of Monthly Wages' in df_v.columns else ''

        report_df = pd.DataFrame({
            'Sl. No.': range(1, len(df_v) + 1),
            'Name of the beneficiary': df_v[name_col] if name_col else '',
            'Profession/Activity': df_v[prof_col].fillna('') if prof_col else '',
            'Father / Husband Name': df_v[father_col].fillna('') if father_col else '',
            'Age': df_v[age_col] if age_col else '',
            'Aadhar No.': df_v[aadhar_col].apply(format_acc) if aadhar_col else '',
            'Address - H.No. & Street': df_v[hno_col].fillna('') if hno_col else '',
            'Address - Village & Mandal': df_v['Village_Mandal'],
            'Address - District': df_v[dist_col].fillna('') if dist_col else '',
            'Geo-tag Number': df_v[geo_col].fillna('') if geo_col else '',
            'Mobile No.': df_v[mobile_col].apply(format_acc) if mobile_col else '',
            'Date of Enrollment into the Scheme': df_v['Date_of_Enrollment'],
            'SB A/c No.': df_v[acc_col].apply(format_acc) if acc_col else '',
            'Name of the Bank & Branch Location': df_v[bank_col].fillna('') if bank_col else '',
            'IFSC Code': df_v[ifsc_col].fillna('') if ifsc_col else '',
            'RD A/c No.1': df_v[rd1_col].apply(format_acc) if rd1_col else '',
            'RD A/c No.2': df_v[rd2_col].apply(format_acc) if rd2_col else '',
            'Declared wages earned per month': df_v[wages_col].fillna('') if wages_col else '',
            'Weaver contribution 8%': df_v[percent_wages].fillna('') if percent_wages else '',
            'Govt. contribution 8%': df_v[percent_wages].fillna('') if percent_wages else ''
        })

        if report_df.empty:
            return jsonify({'status': 'error', 'message': 'రిపోర్ట్ చేయడానికి ఎలాంటి రికార్డులు కనుగొనబడలేదు!'}), 400

        target_sheet = next((s for s in wb.sheetnames if "Form-V" in s or "Form V" in s), wb.sheetnames[0])
        ws = wb[target_sheet]
        start_row = 11
        output_filename = f"FORM_V_Register_Report_{current_month_str}.xlsx"

        for row_idx, row in enumerate(report_df.itertuples(index=False)):
            current_r = start_row + row_idx
            for col_idx, value in enumerate(row, start=1):
                safe_set_cell_value(ws, current_r, col_idx, value)

        output_path = os.path.join(OUTPUT_DIR, output_filename)
        wb.save(output_path)
        
        del df_thrift, df_v
        gc.collect()

        return send_file(output_path, as_attachment=True, download_name=output_filename)

    except Exception as e:
        print("ERROR IN FORM_V:", traceback.format_exc())
        return jsonify({'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}), 500
