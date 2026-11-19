import io
import time
from flask import Blueprint, request, send_file, flash, redirect, url_for
from google.oauth2.service_account import Credentials
import gspread
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import os

pd.set_option('future.no_silent_downcasting', True)

sheets_bp = Blueprint('sheets_bp', __name__)

SPREADSHEET_NAME = 'NOT_BANK_SENT_FORM'

month_map = {
    1: "June-2026", 2: "July-2026", 3: "August-2026", 4: "September-2026",
    5: "October-2026", 6: "November-2026", 7: "December-2026", 8: "January-2027",
    9: "February-2027", 10: "March-2027", 11: "April-2027", 12: "May-2027"
}

cf_sheet_ids = {
    'MARKANDEYULU': '1KG80e8sA2LjqGVJjYfRic7qcaaW40S0mC5HSi0sgfB4',
    'NARESH': '1TjXkA-nWhvqFHBa6LvQODMlpsT7EDuo2hdLOqSKHtdo',
    'SATHISH': '19tEtNCes1DoNkc6UT1jE-xYbGJ_0VqANbhwqGB8gYXA',
    'NAVEEN': '1Lgcfoa3oLlwTStoQp_tfReK3CwNCKqHsxN1Vw3f20V4',
    'SRIKANTH': '1ki-Qepnq0--qO5emIWxBvUI24MrTu8bj0Wt1b8ogSj4',
    'NAGARJUNA': '1AgAndBvngKLGLpvYWDeItgmDszSeD8A5JeIhQleqwN0',
    'DEVADAS': '1w2C0lEu2XN7V7ldsWi5geBs8H2ZB4atZbpkRrBFw1io',
    'KAMALAKAR': '13IZ-SzdsCxSqWrvkvae41BRZlm45lVZa-PCtpO9MqHw',
    'KANAKARAJU': '1Rx91pAUm3CqFXlaavpbnY6LhuWOfNIkkH94l7FtRVIQ',
    'TEXTILE PARK': '1Y1YuE7adT1crmNSpiwIl3rAXDaX4Ag2orqhPXzWQKVU',
}

# కలిపిన కొత్త కాలమ్స్ ('Residential Address', 'Unit Address') ప్రకారం అప్‌డేట్ చేసిన లిస్ట్‌లు
active_columns = [
    'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
    'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
    'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
    'Govt_Share', 'Contact Number', 'Unit Address', 'BANK_REMARKS', 
    'NAME OF CF', 'FILE NAME_x', '8% Of Monthly Wages',
]
rename_active_map = {
    'BANK_SNO': 'S.No', 'Name Of Benificiary': 'Beneficiary Name',
    'Father Name': 'Father Name', 'Date OF Birth': 'Date Of Birth', 'Age': 'Age',
    'Aadhar Number': 'Aadhaar No', 'Caste': 'Caste', 'Profession': 'Profession',
    'Residential Address': 'Residential Address', 'Bank Account Number': 'Account No',
    'RD 1': 'RD Account - 1', 'RD ACCOUNT_2 (Govt. Contribution)': 'RD Account - 2',
    'FILE NAME_x': 'Bank Name', 'Govt_Share': 'Govt_Share',
    'Contact Number': 'Phone Number', 'Status': 'Status', '8% Of Monthly Wages': 'Amount',
    'Unit Address': 'Unit Address', 'BANK_REMARKS': 'Bank Remarks', 'NAME OF CF': 'Name Of CF'
}

missing_columns = [
    'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
    'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
    'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
    '8% Of Monthly Wages', 'Contact Number', 'Unit Address', 'BANK_REMARKS', 
    'NAME OF CF', 'FILE NAME_x',
]
rename_missing_map = {
    'BANK_SNO': 'S.No', 'Name Of Benificiary': 'Beneficiary Name',
    'Father Name': 'Father Name', 'Date OF Birth': 'Date Of Birth', 'Age': 'Age',
    'Aadhar Number': 'Aadhaar No', 'Caste': 'Caste', 'Profession': 'Profession',
    'Residential Address': 'Residential Address', 'Bank Account Number': 'Account No',
    'RD 1': 'RD Account - 1', 'RD ACCOUNT_2 (Govt. Contribution)': 'RD Account - 2',
    'FILE NAME_x': 'Bank Name', '8% Of Monthly Wages': 'Amount',
    'Contact Number': 'Phone Number', 'Status': 'Status',
    'Unit Address': 'Unit Address', 'BANK_REMARKS': 'Bank Remarks', 'NAME OF CF': 'Name Of CF'
}

all_columns = [
    'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
    'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
    'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
    'Govt_Share', 'Contact Number', 'Unit Address', 'BANK_REMARKS', 
    'NAME OF CF', 'FILE NAME_x', '8% Of Monthly Wages',
]
rename_all_map = {
    'BANK_SNO': 'S.No', 'Name Of Benificiary': 'Beneficiary Name',
    'Father Name': 'Father Name', 'Date OF Birth': 'Date Of Birth', 'Age': 'Age',
    'Aadhar Number': 'Aadhaar No', 'Caste': 'Caste', 'Profession': 'Profession',
    'Residential Address': 'Residential Address', 'Bank Account Number': 'Account No',
    'RD 1': 'RD Account - 1', 'RD ACCOUNT_2 (Govt. Contribution)': 'RD Account - 2',
    'FILE NAME_x': 'Bank Name', 'Govt_Share': 'Govt Share Amount',
    'Contact Number': 'Phone Number', 'Status': 'Status', '8% Of Monthly Wages': 'Amount',
    'Unit Address': 'Unit Address', 'BANK_REMARKS': 'Bank Remarks', 'NAME OF CF': 'Name Of CF'
}

not_sent_columns = [
    'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
    'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
    'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
    'Contact Number', 'Unit Address', 'BANK_REMARKS', 'NAME OF CF', 
    'FILE NAME_x', '8% Of Monthly Wages', 'BANK SENT',
]
rename_not_sent_map = {
    'BANK_SNO': 'S.No', 'Name Of Benificiary': 'Beneficiary Name',
    'Father Name': 'Father Name', 'Date OF Birth': 'Date Of Birth', 'Age': 'Age',
    'Aadhar Number': 'Aadhaar No', 'Caste': 'Caste', 'Profession': 'Profession',
    'Residential Address': 'Residential Address', 'Bank Account Number': 'Account No',
    'RD 1': 'RD Account - 1', 'RD ACCOUNT_2 (Govt. Contribution)': 'RD Account - 2',
    'FILE NAME_x': 'Bank Name', 'Contact Number': 'Phone Number', 'Status': 'Status',
    '8% Of Monthly Wages': 'Amount', 'Unit Address': 'Unit Address',
    'BANK_REMARKS': 'Bank Remarks', 'NAME OF CF': 'Name Of CF', 'BANK SENT': 'Bank Sent Status'
}

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_gspread_client():
    credentials_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if credentials_json_str:
        import json
        info = json.loads(credentials_json_str)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        cred_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json')
        creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
    return gspread.authorize(creds)

def combine_address_fields(df, cols):
    """ఇచ్చిన కాలమ్స్‌లోని విలువలను ఖాళీలు లేకుండా కామా(,) తో కలిపి ఒకే స్ట్రింగ్‌గా ఇస్తుంది."""
    def join_row(row):
        parts = []
        for col in cols:
            val = str(row.get(col, '')).strip()
            if val and val.lower() not in ['nan', 'none', 'null']:
                parts.append(val)
        return ', '.join(parts)
    return df.apply(join_row, axis=1)

def fetch_and_process_data(current_month):
    gc = get_gspread_client()
    sh = gc.open(SPREADSHEET_NAME)

    ws_trans = sh.worksheet('Transactions_Year1')
    df_trans = pd.DataFrame(ws_trans.get_all_records())

    ws_thrift = sh.worksheet('NEW THRIFT DATA')
    df_thrift = pd.DataFrame(ws_thrift.get_all_records())

    # BANK SENT కాలమ్ ఉందో లేదో తనిఖీ చేసి ఫ్లాగ్ సెట్ చేయడం
    if 'BANK SENT' in df_thrift.columns:
        df_thrift['Is_Not_Sent'] = df_thrift['BANK SENT'].astype(str).str.strip().str.upper() == 'NOT SENT'
    else:
        df_thrift['Is_Not_Sent'] = False

    df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0)
    df_trans['MONTH_NUMBER'] = df_trans['MONTH_NUMBER'].astype(str).str.strip()
    df_trans['MONTH_NUMBER'] = pd.to_numeric(df_trans['MONTH_NUMBER'], errors='coerce').fillna(0).astype(int)

    df_trans_month = df_trans[df_trans['MONTH_NUMBER'] == int(current_month)].copy()
    df_trans_month = df_trans_month.drop_duplicates(
        subset=['MONTH_NUMBER', 'RD ACCOUNT_1 (Beneficiary Contribution)'], keep='first'
    )

    df_thrift['RD 1_Clean'] = df_thrift['RD 1'].astype(str).str.strip()
    df_trans_month['RD1_Clean'] = df_trans_month['RD ACCOUNT_1 (Beneficiary Contribution)'].astype(str).str.strip()

    df_cross = pd.merge(
        df_thrift,
        df_trans_month[['RD1_Clean', 'Govt_Share', 'MONTH_NUMBER', 'FILE NAME']],
        left_on='RD 1_Clean', right_on='RD1_Clean', how='left'
    )

    df_cross['Govt_Share'] = df_cross['Govt_Share'].fillna(0)
    month_label = month_map.get(current_month, f"Month-{current_month}")
    
    df_cross['Status'] = df_cross.apply(
        lambda r: 'Not Sent to Bank' if r.get('Is_Not_Sent', False)
        else (f'{month_label} Govt_Share Released' if r['Govt_Share'] > 0
              else f'Zero or missing transactions reported for {month_label}'),
        axis=1
    )

    # 1. Residential Address కాలమ్ తయారీ
    df_cross['Residential Address'] = combine_address_fields(
        df_cross, ['H.NO. Residentional H.No', 'VILLAGE NAME', 'MANDAL NAME']
    )

    # 2. Unit Address కాలమ్ తయారీ
    df_cross['Unit Address'] = combine_address_fields(
        df_cross, ['Unit Owner', 'H.NO. Unit H.No', 'Area', 'Village']
    )

    # బ్యాంకుకు పంపనివారిని (Is_Not_Sent == True) విడిగా వేరు చేసి,
    # మిగిలిన ప్రధాన డేటాను (df_sent_cross) బ్యాంకుకు పంపినవారి కోసం ఉంచుతాము
    df_not_sent_all = df_cross[df_cross['Is_Not_Sent']].copy()
    df_sent_cross = df_cross[~df_cross['Is_Not_Sent']].copy()

    return gc, df_sent_cross, df_not_sent_all, month_label

def generate_cf_excel_in_memory(clean_cf_name, group_df, df_sent_cross, df_not_sent_group, current_month, month_label):
    wb = openpyxl.Workbook()
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    blue_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    red_fill = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    green_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
    orange_fill = PatternFill(start_color='D97706', end_color='D97706', fill_type='solid')

    df_active_cf = group_df[group_df['Govt_Share'] > 0].copy()
    df_missing_cf = group_df[group_df['Govt_Share'] == 0].copy()

    # SHEET 1: ACTIVE (19 కాలమ్స్ -> A1:S1)
    ws1 = wb.active
    ws1.title = 'Sent_To_Bank_List'
    ws1.merge_cells('A1:S1')
    ws1['A1'] = f'{clean_cf_name} - MONTH {current_month} BANK SENT ({len(df_active_cf)} Records)'
    ws1['A1'].font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    ws1['A1'].fill = blue_fill
    ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws1.append([])
    headers_active = [
        'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
        'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
        'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
        'Bank / Branch Name', 'Govt Share Amount', 'Contact Number', 
        'Unit Address', 'BANK_REMARKS', 'NAME OF CF', 'Status'
    ]
    ws1.append(headers_active)
    for col_num in range(1, len(headers_active) + 1):
        c = ws1.cell(row=3, column=col_num)
        c.font = header_font
        c.fill = blue_fill

    r_idx1 = 4
    for _, r in df_active_cf.iterrows():
        ws1.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Father Name', ''),
            r.get('Date OF Birth', ''), r.get('Age', ''), r.get('Aadhar Number', ''),
            r.get('Caste', ''), r.get('Profession', ''), r.get('Residential Address', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), 
            r.get('RD ACCOUNT_2 (Govt. Contribution)', ''),
            r.get('FILE NAME_x', r.get('FILE NAME', '')), r.get('Govt_Share', 0),
            r.get('Contact Number', ''), r.get('Unit Address', ''), 
            r.get('BANK_REMARKS', ''), r.get('NAME OF CF', ''), 
            f'{month_label} Govt_Share Released'
        ])
        ws1.cell(row=r_idx1, column=14).number_format = '₹#,##0.00'
        r_idx1 += 1

    # SHEET 2: MISSING (18 కాలమ్స్ -> A1:R1)
    ws2 = wb.create_sheet(title='Missing_or_Zero_List')
    ws2.merge_cells('A1:R1')
    ws2['A1'] = f'{clean_cf_name} - MONTH {current_month} MISSING / ZERO SHARE ({len(df_missing_cf)} Records)'
    ws2['A1'].font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    ws2['A1'].fill = red_fill
    ws2['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws2.append([])
    headers_missing = [
        'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
        'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
        'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
        'Bank / Branch Name', 'Contact Number', 'Unit Address', 
        'BANK_REMARKS', 'NAME OF CF', 'Status'
    ]
    ws2.append(headers_missing)
    for col_num in range(1, len(headers_missing) + 1):
        c = ws2.cell(row=3, column=col_num)
        c.font = header_font
        c.fill = red_fill

    for _, r in df_missing_cf.iterrows():
        ws2.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Father Name', ''),
            r.get('Date OF Birth', ''), r.get('Age', ''), r.get('Aadhar Number', ''),
            r.get('Caste', ''), r.get('Profession', ''), r.get('Residential Address', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), 
            r.get('RD ACCOUNT_2 (Govt. Contribution)', ''),
            r.get('FILE NAME_x', r.get('FILE NAME', 'Not Assigned')), r.get('Contact Number', ''),
            r.get('Unit Address', ''), r.get('BANK_REMARKS', ''), 
            r.get('NAME OF CF', ''), f'Zero or missing transactions reported for {month_label}'
        ])

    # SHEET 3: ALL BENEFICIARIES SENT TO BANK (19 కాలమ్స్ -> A1:S1)
    ws3 = wb.create_sheet(title='All_Beneficiaries_Master')
    ws3.merge_cells('A1:S1')
    ws3['A1'] = f'PROJECT BANK SENT BENEFICIARIES MASTER DATA - MONTH {current_month} ({len(df_sent_cross)} Records)'
    ws3['A1'].font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    ws3['A1'].fill = green_fill
    ws3['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws3.append([])
    headers_all = [
        'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
        'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
        'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
        'Bank / Branch Name', 'Govt Share Amount', 'Contact Number', 
        'Unit Address', 'BANK_REMARKS', 'NAME OF CF', 'Status'
    ]
    ws3.append(headers_all)
    for col_num in range(1, len(headers_all) + 1):
        c = ws3.cell(row=3, column=col_num)
        c.font = header_font
        c.fill = green_fill

    r_idx3 = 4
    for _, r in df_sent_cross.iterrows():
        ws3.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Father Name', ''),
            r.get('Date OF Birth', ''), r.get('Age', ''), r.get('Aadhar Number', ''),
            r.get('Caste', ''), r.get('Profession', ''), r.get('Residential Address', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), 
            r.get('RD ACCOUNT_2 (Govt. Contribution)', ''),
            r.get('FILE NAME_x', r.get('FILE NAME', '')), r.get('Govt_Share', 0),
            r.get('Contact Number', ''), r.get('Unit Address', ''), 
            r.get('BANK_REMARKS', ''), r.get('NAME OF CF', ''), r.get('Status', '')
        ])
        ws3.cell(row=r_idx3, column=14).number_format = '₹#,##0.00'
        r_idx3 += 1

    # SHEET 4: NOT SENT TO BANK LIST (సపరేట్ వర్క్ షీట్ టాబ్)
    ws4 = wb.create_sheet(title='Not_Sent_To_Bank_List')
    ws4.merge_cells('A1:S1')
    ws4['A1'] = f'{clean_cf_name} - NOT SENT TO BANK LIST ({len(df_not_sent_group)} Records)'
    ws4['A1'].font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    ws4['A1'].fill = orange_fill
    ws4['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws4.append([])
    headers_not_sent = [
        'BANK_SNO', 'Name Of Benificiary', 'Father Name', 'Date OF Birth', 'Age', 
        'Aadhar Number', 'Caste', 'Profession', 'Residential Address', 
        'Bank Account Number', 'RD 1', 'RD ACCOUNT_2 (Govt. Contribution)', 
        'Bank / Branch Name', 'Contact Number', 'Unit Address', 
        'BANK_REMARKS', 'NAME OF CF', 'BANK SENT', 'Status'
    ]
    ws4.append(headers_not_sent)
    for col_num in range(1, len(headers_not_sent) + 1):
        c = ws4.cell(row=3, column=col_num)
        c.font = header_font
        c.fill = orange_fill

    for _, r in df_not_sent_group.iterrows():
        ws4.append([
            r.get('BANK_SNO', ''), r.get('Name Of Benificiary', ''), r.get('Father Name', ''),
            r.get('Date OF Birth', ''), r.get('Age', ''), r.get('Aadhar Number', ''),
            r.get('Caste', ''), r.get('Profession', ''), r.get('Residential Address', ''),
            r.get('Bank Account Number', ''), r.get('RD 1', ''), 
            r.get('RD ACCOUNT_2 (Govt. Contribution)', ''),
            r.get('FILE NAME_x', r.get('FILE NAME', 'Not Assigned')), r.get('Contact Number', ''),
            r.get('Unit Address', ''), r.get('BANK_REMARKS', ''), 
            r.get('NAME OF CF', ''), r.get('BANK SENT', 'NOT SENT'), 'Not Sent to Bank'
        ])

    for ws in [ws1, ws2, ws3, ws4]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 16)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@sheets_bp.route('/update_sheets', methods=['POST'])
def update_sheets():
    current_month = int(request.form.get('month', 1))
    try:
        gc, df_sent_cross, df_not_sent_all, _ = fetch_and_process_data(current_month)
        df_all_master_clean = df_sent_cross.fillna('').infer_objects(copy=False)
        df_not_sent_clean_all = df_not_sent_all.fillna('').infer_objects(copy=False)

        cf_col = 'NAME OF CF' if 'NAME OF CF' in df_sent_cross.columns else 'CF NAME'
        grouped_cf = df_sent_cross.groupby(cf_col)
        grouped_not_sent_cf = df_not_sent_all.groupby(cf_col) if not df_not_sent_all.empty and cf_col in df_not_sent_all.columns else {}

        updated_count = 0
        for cf_name, group_df in grouped_cf:
            clean_cf_name = str(cf_name).strip()
            if clean_cf_name not in cf_sheet_ids:
                continue

            df_active_clean = group_df[group_df['Govt_Share'] > 0].fillna('').infer_objects(copy=False)
            df_missing_clean = group_df[group_df['Govt_Share'] == 0].fillna('').infer_objects(copy=False)

            if isinstance(grouped_not_sent_cf, dict) or grouped_not_sent_cf.empty:
                df_cf_not_sent = pd.DataFrame()
            else:
                try:
                    df_cf_not_sent = grouped_not_sent_cf.get_group(cf_name).fillna('').infer_objects(copy=False)
                except KeyError:
                    df_cf_not_sent = pd.DataFrame()

            sheet_id = cf_sheet_ids[clean_cf_name]
            for attempt in range(3):
                try:
                    spreadsheet = gc.open_by_key(sheet_id)
                    worksheets = spreadsheet.worksheets()

                    # TAB 1: Sent_To_Bank_List
                    sh1 = worksheets[0] if len(worksheets) > 0 else spreadsheet.add_worksheet(title="Sent_To_Bank_List", rows="1000", cols="30")
                    sh1.clear()
                    cols1 = [c for c in active_columns + ['Status'] if c in df_active_clean.columns]
                    df_act_sel = df_active_clean[cols1].rename(columns=rename_active_map)
                    sh1.update(range_name='A1', values=[df_act_sel.columns.values.tolist()] + df_act_sel.values.tolist())

                    # TAB 2: Missing_or_Zero_List
                    sh2 = worksheets[1] if len(worksheets) > 1 else spreadsheet.add_worksheet(title="Missing_or_Zero_List", rows="1000", cols="30")
                    sh2.clear()
                    cols2 = [c for c in missing_columns + ['Status'] if c in df_missing_clean.columns]
                    df_mis_sel = df_missing_clean[cols2].rename(columns=rename_missing_map)
                    sh2.update(range_name='A1', values=[df_mis_sel.columns.values.tolist()] + df_mis_sel.values.tolist())

                    # TAB 3: All_Beneficiaries_Master (Bank Sent Only)
                    sh3 = worksheets[2] if len(worksheets) > 2 else spreadsheet.add_worksheet(title="All_Beneficiaries_Master", rows="2000", cols="30")
                    sh3.clear()
                    cols3 = [c for c in all_columns + ['Status'] if c in df_all_master_clean.columns]
                    df_all_sel = df_all_master_clean[cols3].rename(columns=rename_all_map)
                    sh3.update(range_name='A1', values=[df_all_sel.columns.values.tolist()] + df_all_sel.values.tolist())

                    # TAB 4: Not_Sent_To_Bank_List (సపరేట్ టాబ్)
                    try:
                        sh4 = spreadsheet.worksheet("Not_Sent_To_Bank_List")
                    except gspread.WorksheetNotFound:
                        sh4 = spreadsheet.add_worksheet(title="Not_Sent_To_Bank_List", rows="1000", cols="30")
                    sh4.clear()
                    if not df_cf_not_sent.empty:
                        cols4 = [c for c in not_sent_columns + ['Status'] if c in df_cf_not_sent.columns]
                        df_ns_sel = df_cf_not_sent[cols4].rename(columns=rename_not_sent_map)
                        sh4.update(range_name='A1', values=[df_ns_sel.columns.values.tolist()] + df_ns_sel.values.tolist())
                    else:
                        sh4.update(range_name='A1', values=[["No Not-Sent records for this CF"]])

                    updated_count += 1
                    break
                except Exception as e:
                    if attempt == 2:
                        raise e
                    time.sleep(5)
            time.sleep(1.5)

        flash(f'✅ విజయవంతంగా Month {current_month} కోసం {updated_count} CF Google Sheets అప్‌డేట్ అయ్యాయి (Not Sent లిస్ట్ విడి టాబ్‌లో చేర్చబడింది)!', 'success')
    except Exception as e:
        flash(f'❌ Error ఏర్పడింది: {str(e)}', 'danger')

    return redirect(url_for('home'))

@sheets_bp.route('/download_excel', methods=['POST'])
def download_excel():
    current_month = int(request.form.get('month', 1))
    selected_cf = request.form.get('cf_name')

    try:
        _, df_sent_cross, df_not_sent_all, month_label = fetch_and_process_data(current_month)
        cf_col = 'NAME OF CF' if 'NAME OF CF' in df_sent_cross.columns else 'CF NAME'
        
        group_df = df_sent_cross[df_sent_cross[cf_col].astype(str).str.strip() == selected_cf].copy()
        not_sent_group_df = df_not_sent_all[df_not_sent_all[cf_col].astype(str).str.strip() == selected_cf].copy() if not df_not_sent_all.empty and cf_col in df_not_sent_all.columns else pd.DataFrame()

        excel_stream = generate_cf_excel_in_memory(selected_cf, group_df, df_sent_cross, not_sent_group_df, current_month, month_label)
        filename = f"{selected_cf}_Month_{current_month}_Report.xlsx"

        return send_file(
            excel_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        flash(f'❌ Excel డౌన్‌లోడ్‌లో సమస్య ఏర్పడింది: {str(e)}', 'danger')
        return redirect(url_for('home'))
