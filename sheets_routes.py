import gc
import io
import json
import logging
import math
import os
import threading
import time

from flask import Blueprint, flash, jsonify, redirect, request, send_file, url_for
from google.oauth2.service_account import Credentials
import gspread
from gspread.exceptions import APIError
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

pd.set_option('future.no_silent_downcasting', True)

logger = logging.getLogger(__name__)
sheets_bp = Blueprint('sheets_bp', __name__)

SPREADSHEET_NAME = 'NOT_BANK_SENT_FORM'

_update_lock = threading.Lock()
_update_state = {
    "running": False,
    "month": None,
    "completed_cfs": 0,
    "total_cfs": 0,
    "current_cf": "",
    "last_message": "",
    "last_error": "",
}

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


def build_sample_record_df(clean_cf_name: str, status_text: str, template_df: pd.DataFrame = None) -> pd.DataFrame:
    """ఏదైనా CF లేదా టాబ్‌లో రికార్డులు లేనప్పుడు AppSheet సజావుగా పనిచేయడానికి SAMPLE / NO RECORD రో ఇస్తుంది."""
    sample_row = {
        'BANK_SNO': 'SAMPLE',
        'Name Of Benificiary': 'SAMPLE / NO RECORD',
        'Father Name': 'SAMPLE / NO RECORD',
        'Date OF Birth': '',
        'Age': 0,
        'Aadhar Number': '',
        'Caste': '',
        'Profession': '',
        'Residential Address': '',
        'Bank Account Number': '',
        'RD 1': '',
        'RD ACCOUNT_2 (Govt. Contribution)': '',
        'Govt_Share': 0,
        'Contact Number': '',
        'Unit Address': '',
        'BANK_REMARKS': 'SAMPLE / NO RECORD',
        'NAME OF CF': clean_cf_name,
        'FILE NAME_x': 'Not Assigned',
        '8% Of Monthly Wages': 0,
        'BANK SENT': 'SAMPLE / NO RECORD',
        'Status': status_text,
    }
    if template_df is not None and not template_df.empty:
        for col in template_df.columns:
            if col not in sample_row:
                sample_row[col] = ''
        return pd.DataFrame([sample_row])[list(template_df.columns)]
    return pd.DataFrame([sample_row])


def redirect_home():
    for endpoint in ('bank_bp.home', 'home'):
        try:
            return redirect(url_for(endpoint))
        except Exception:
            continue
    return redirect('/')


def get_gspread_client():
    credentials_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if credentials_json_str:
        info = json.loads(credentials_json_str)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        cred_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json')
        creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
    return gspread.authorize(creds)


def gspread_call_with_retry(func, *args, max_retries=5, base_delay=2.0, **kwargs):
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except APIError as e:
            status_code = getattr(getattr(e, 'response', None), 'status_code', None)
            if status_code in (429, 500, 502, 503, 504) and attempt < max_retries - 1:
                sleep_time = base_delay * (2 ** attempt)
                logger.warning(f"Sheets API status {status_code}. Retrying in {sleep_time:.1f}s...")
                time.sleep(sleep_time)
            else:
                raise
        except Exception:
            if attempt < max_retries - 1:
                sleep_time = base_delay * (2 ** attempt)
                time.sleep(sleep_time)
            else:
                raise


def _clean_scalar(val):
    if val is None:
        return ''
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return ''
        return val
    try:
        if pd.isna(val):
            return ''
    except (TypeError, ValueError):
        pass
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        fval = float(val)
        return '' if (math.isnan(fval) or math.isinf(fval)) else fval
    if isinstance(val, (np.bool_,)):
        return bool(val)
    return val


def df_to_sheet_values(df: pd.DataFrame):
    if df is None or df.empty:
        return []
    headers = [str(c) for c in df.columns.tolist()]
    rows = [[_clean_scalar(v) for v in row] for row in df.itertuples(index=False, name=None)]
    return [headers] + rows


def write_worksheet_safely(worksheet, values_2d):
    if not values_2d:
        values_2d = [['No records found']]

    needed_rows = len(values_2d) + 2
    needed_cols = max((len(r) for r in values_2d), default=1) + 1
    if worksheet.row_count < needed_rows or worksheet.col_count < needed_cols:
        gspread_call_with_retry(
            worksheet.resize,
            rows=max(worksheet.row_count, needed_rows),
            cols=max(worksheet.col_count, needed_cols)
        )

    if worksheet.row_count > len(values_2d):
        end_col_letter = get_column_letter(max(worksheet.col_count, needed_cols))
        gspread_call_with_retry(worksheet.batch_clear, [f"A1:{end_col_letter}{worksheet.row_count}"])

    gspread_call_with_retry(
        worksheet.update,
        range_name='A1',
        values=values_2d,
        value_input_option='USER_ENTERED'
    )


def combine_address_fields(df: pd.DataFrame, cols):
    existing_cols = [c for c in cols if c in df.columns]
    if not existing_cols:
        return pd.Series([''] * len(df), index=df.index)

    def join_row(row):
        parts = []
        for col in existing_cols:
            raw = row.get(col, '')
            if pd.isna(raw):
                continue
            val = str(raw).strip()
            if val and val.lower() not in ('nan', 'none', 'null', 'nat'):
                parts.append(val)
        return ', '.join(parts)

    return df.apply(join_row, axis=1)


def fetch_and_process_data(current_month: int):
    gspread_client = get_gspread_client()
    sh = gspread_call_with_retry(gspread_client.open, SPREADSHEET_NAME)

    ws_trans = gspread_call_with_retry(sh.worksheet, 'Transactions_Year1')
    trans_records = gspread_call_with_retry(ws_trans.get_all_records)
    df_trans = pd.DataFrame(trans_records)
    del trans_records

    ws_thrift = gspread_call_with_retry(sh.worksheet, 'NEW THRIFT DATA')
    thrift_records = gspread_call_with_retry(ws_thrift.get_all_records)
    df_thrift = pd.DataFrame(thrift_records)
    del thrift_records
    gc.collect()

    if df_thrift.empty:
        raise ValueError("'NEW THRIFT DATA' షీట్‌లో డేటా లేదు.")

    if 'BANK SENT' in df_thrift.columns:
        df_thrift['Is_Not_Sent'] = df_thrift['BANK SENT'].astype(str).str.strip().str.upper() == 'NOT SENT'
    else:
        df_thrift['Is_Not_Sent'] = False

    if df_trans.empty:
        df_trans_month = pd.DataFrame(columns=['RD1_Clean', 'Govt_Share', 'MONTH_NUMBER', 'FILE NAME'])
    else:
        if 'Govt_Share' not in df_trans.columns:
            df_trans['Govt_Share'] = 0.0
        else:
            df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0.0)

        if 'MONTH_NUMBER' not in df_trans.columns:
            df_trans['MONTH_NUMBER'] = 0
        else:
            df_trans['MONTH_NUMBER'] = pd.to_numeric(
                df_trans['MONTH_NUMBER'].astype(str).str.strip(), errors='coerce'
            ).fillna(0).astype(int)

        rd1_trans_col = 'RD ACCOUNT_1 (Beneficiary Contribution)'
        if rd1_trans_col not in df_trans.columns:
            raise KeyError(f"Transactions_Year1 లో '{rd1_trans_col}' కాలమ్ కనిపించలేదు.")

        if 'FILE NAME' not in df_trans.columns:
            df_trans['FILE NAME'] = ''

        df_trans_month = df_trans[df_trans['MONTH_NUMBER'] == int(current_month)].copy()
        df_trans_month['RD1_Clean'] = df_trans_month[rd1_trans_col].astype(str).str.strip()
        df_trans_month = df_trans_month[df_trans_month['RD1_Clean'] != '']
        df_trans_month = df_trans_month.drop_duplicates(
            subset=['MONTH_NUMBER', 'RD1_Clean'], keep='first'
        )

    del df_trans
    gc.collect()

    if 'RD 1' not in df_thrift.columns:
        raise KeyError("NEW THRIFT DATA లో 'RD 1' కాలమ్ కనిపించలేదు.")

    df_thrift['RD 1_Clean'] = df_thrift['RD 1'].astype(str).str.strip()

    right_cols = [c for c in ['RD1_Clean', 'Govt_Share', 'MONTH_NUMBER', 'FILE NAME'] if c in df_trans_month.columns]
    df_cross = pd.merge(
        df_thrift,
        df_trans_month[right_cols] if not df_trans_month.empty else pd.DataFrame(columns=['RD1_Clean', 'Govt_Share', 'FILE NAME']),
        left_on='RD 1_Clean',
        right_on='RD1_Clean',
        how='left'
    )

    del df_thrift, df_trans_month
    gc.collect()

    if 'Govt_Share' not in df_cross.columns:
        df_cross['Govt_Share'] = 0.0
    else:
        df_cross['Govt_Share'] = pd.to_numeric(df_cross['Govt_Share'], errors='coerce').fillna(0.0)

    month_label = month_map.get(current_month, f"Month-{current_month}")

    df_cross['Status'] = df_cross.apply(
        lambda r: 'Not Sent to Bank' if bool(r.get('Is_Not_Sent', False))
        else (f'{month_label} Govt_Share Released' if float(r.get('Govt_Share', 0.0)) > 0
              else f'Zero or missing transactions reported for {month_label}'),
        axis=1
    )

    df_cross['Residential Address'] = combine_address_fields(
        df_cross, ['H.NO. Residentional H.No', 'VILLAGE NAME', 'MANDAL NAME']
    )
    df_cross['Unit Address'] = combine_address_fields(
        df_cross, ['Unit Owner', 'H.NO. Unit H.No', 'Area', 'Village']
    )

    df_not_sent_all = df_cross[df_cross['Is_Not_Sent']].copy()
    df_sent_cross = df_cross[~df_cross['Is_Not_Sent']].copy()

    del df_cross
    gc.collect()

    return gspread_client, df_sent_cross, df_not_sent_all, month_label


def _update_single_cf_spreadsheet(gspread_client, clean_cf_name: str, sheet_id: str,
                                  df_active_clean: pd.DataFrame, df_missing_clean: pd.DataFrame,
                                  master_values: list, df_cf_not_sent: pd.DataFrame):
    """ఒక్క CF స్ప్రెడ్‌షీట్‌లోని అన్ని టాబ్‌లను సురక్షితంగా అప్‌డేట్ చేస్తుంది."""
    spreadsheet = gspread_call_with_retry(gspread_client.open_by_key, sheet_id)
    worksheets = gspread_call_with_retry(spreadsheet.worksheets)

    sh1 = worksheets[0] if len(worksheets) > 0 else gspread_call_with_retry(
        spreadsheet.add_worksheet, title="Sent_To_Bank_List", rows="1000", cols="25"
    )
    cols1 = [c for c in active_columns + ['Status'] if c in df_active_clean.columns]
    write_worksheet_safely(sh1, df_to_sheet_values(df_active_clean[cols1].rename(columns=rename_active_map)))
    time.sleep(0.5)

    sh2 = worksheets[1] if len(worksheets) > 1 else gspread_call_with_retry(
        spreadsheet.add_worksheet, title="Missing_or_Zero_List", rows="1000", cols="25"
    )
    cols2 = [c for c in missing_columns + ['Status'] if c in df_missing_clean.columns]
    write_worksheet_safely(sh2, df_to_sheet_values(df_missing_clean[cols2].rename(columns=rename_missing_map)))
    time.sleep(0.5)

    sh3 = worksheets[2] if len(worksheets) > 2 else gspread_call_with_retry(
        spreadsheet.add_worksheet, title="All_Beneficiaries_Master", rows="2000", cols="25"
    )
    write_worksheet_safely(sh3, master_values)
    time.sleep(0.5)

    try:
        sh4 = gspread_call_with_retry(spreadsheet.worksheet, "Not_Sent_To_Bank_List")
    except gspread.WorksheetNotFound:
        sh4 = gspread_call_with_retry(
            spreadsheet.add_worksheet, title="Not_Sent_To_Bank_List", rows="1000", cols="25"
        )

    cols4 = [c for c in not_sent_columns + ['Status'] if c in df_cf_not_sent.columns]
    write_worksheet_safely(sh4, df_to_sheet_values(df_cf_not_sent[cols4].rename(columns=rename_not_sent_map)))


def _run_all_cfs_update(current_month: int, target_cf: str = 'ALL', track_progress: bool = False):
    gspread_client, df_sent_cross, df_not_sent_all, month_label = fetch_and_process_data(current_month)

    cf_col = 'NAME OF CF' if 'NAME OF CF' in df_sent_cross.columns else ('CF NAME' if 'CF NAME' in df_sent_cross.columns else None)
    if not cf_col:
        raise KeyError("డేటాలో 'NAME OF CF' లేదా 'CF NAME' కాలమ్ కనిపించలేదు.")

    df_all_master_clean = df_sent_cross.fillna('').infer_objects(copy=False)
    cols3 = [c for c in all_columns + ['Status'] if c in df_all_master_clean.columns]
    master_values = df_to_sheet_values(df_all_master_clean[cols3].rename(columns=rename_all_map))
    del df_all_master_clean
    gc.collect()

    grouped_cf = df_sent_cross.groupby(cf_col)
    eligible_groups = [
        (str(cf_name).strip(), cf_name, group_df)
        for cf_name, group_df in grouped_cf
        if str(cf_name).strip() in cf_sheet_ids
        and (target_cf in ('', 'ALL', 'ALL CFS') or str(cf_name).strip().upper() == target_cf)
    ]

    # డేటాలో అసలు రికార్డులు లేని CF షీట్స్ కూడా SAMPLE / NO RECORD తో అప్‌డేట్ అయ్యేలా చేర్చడం
    existing_clean_names = {name for name, _, _ in eligible_groups}
    for mapped_cf_name in cf_sheet_ids.keys():
        if mapped_cf_name not in existing_clean_names:
            if target_cf in ('', 'ALL', 'ALL CFS') or mapped_cf_name.upper() == target_cf:
                eligible_groups.append((mapped_cf_name, mapped_cf_name, pd.DataFrame(columns=df_sent_cross.columns)))

    if track_progress:
        _update_state["total_cfs"] = len(eligible_groups)
        _update_state["completed_cfs"] = 0

    has_not_sent_cf = (
        not df_not_sent_all.empty
        and cf_col in df_not_sent_all.columns
    )
    grouped_not_sent_cf = df_not_sent_all.groupby(cf_col) if has_not_sent_cf else None

    updated_count = 0
    failed_cfs = []

    for clean_cf_name, raw_cf_name, group_df in eligible_groups:
        if track_progress:
            _update_state["current_cf"] = clean_cf_name

        df_active_raw = group_df[group_df['Govt_Share'] > 0] if (not group_df.empty and 'Govt_Share' in group_df.columns) else pd.DataFrame()
        df_missing_raw = group_df[group_df['Govt_Share'] == 0] if (not group_df.empty and 'Govt_Share' in group_df.columns) else pd.DataFrame()

        if df_active_raw.empty:
            df_active_clean = build_sample_record_df(
                clean_cf_name, f'SAMPLE / NO RECORD ({month_label})', df_active_raw if not df_active_raw.empty else group_df
            ).fillna('')
        else:
            df_active_clean = df_active_raw.fillna('').infer_objects(copy=False)

        if df_missing_raw.empty:
            df_missing_clean = build_sample_record_df(
                clean_cf_name, f'SAMPLE / NO RECORD ({month_label})', df_missing_raw if not df_missing_raw.empty else group_df
            ).fillna('')
        else:
            df_missing_clean = df_missing_raw.fillna('').infer_objects(copy=False)

        if grouped_not_sent_cf is not None and raw_cf_name in grouped_not_sent_cf.groups:
            df_cf_not_sent = grouped_not_sent_cf.get_group(raw_cf_name).fillna('').infer_objects(copy=False)
        else:
            df_cf_not_sent = build_sample_record_df(
                clean_cf_name, 'SAMPLE / NO RECORD', df_not_sent_all if not df_not_sent_all.empty else group_df
            ).fillna('')

        sheet_id = cf_sheet_ids[clean_cf_name]
        try:
            _update_single_cf_spreadsheet(
                gspread_client=gspread_client,
                clean_cf_name=clean_cf_name,
                sheet_id=sheet_id,
                df_active_clean=df_active_clean,
                df_missing_clean=df_missing_clean,
                master_values=master_values,
                df_cf_not_sent=df_cf_not_sent,
            )
            updated_count += 1
            if track_progress:
                _update_state["completed_cfs"] = updated_count
        except Exception as cf_err:
            logger.exception(f"Failed updating sheet for CF '{clean_cf_name}': {cf_err}")
            failed_cfs.append(f"{clean_cf_name} ({cf_err})")

        del df_active_clean, df_missing_clean, df_cf_not_sent
        gc.collect()
        time.sleep(1.2)

    del master_values, df_sent_cross, df_not_sent_all
    gc.collect()

    if failed_cfs and updated_count == 0:
        raise RuntimeError("అన్ని CF షీట్స్ అప్‌డేట్ చేయడంలో విఫలమైంది: " + "; ".join(failed_cfs))

    msg = f"✅ విజయవంతంగా Month {current_month} కోసం {updated_count} CF Google Sheets అప్‌డేట్ అయ్యాయి!"
    if failed_cfs:
        msg += f" (కొన్ని షీట్స్‌లో సమస్య: {', '.join(failed_cfs)})"
    return msg, bool(failed_cfs)


def _background_all_cfs_worker(current_month: int):
    try:
        msg, _ = _run_all_cfs_update(current_month=current_month, target_cf='ALL', track_progress=True)
        _update_state["last_message"] = msg
        _update_state["last_error"] = ""
        logger.info(msg)
    except Exception as exc:
        logger.exception("Background update for ALL CFs failed")
        _update_state["last_error"] = str(exc)
    finally:
        _update_state["running"] = False
        _update_state["current_cf"] = ""
        gc.collect()


def generate_cf_excel_in_memory(clean_cf_name, group_df, df_sent_cross, df_not_sent_group, current_month, month_label):
    wb = openpyxl.Workbook()
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    blue_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    red_fill = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    green_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
    orange_fill = PatternFill(start_color='D97706', end_color='D97706', fill_type='solid')

    df_active_cf = group_df[group_df['Govt_Share'] > 0] if (not group_df.empty and 'Govt_Share' in group_df.columns) else pd.DataFrame()
    df_missing_cf = group_df[group_df['Govt_Share'] == 0] if (not group_df.empty and 'Govt_Share' in group_df.columns) else pd.DataFrame()

    if df_active_cf.empty:
        df_active_cf = build_sample_record_df(clean_cf_name, f'SAMPLE / NO RECORD ({month_label})')
    if df_missing_cf.empty:
        df_missing_cf = build_sample_record_df(clean_cf_name, f'SAMPLE / NO RECORD ({month_label})')
    if df_not_sent_group is None or df_not_sent_group.empty:
        df_not_sent_group = build_sample_record_df(clean_cf_name, 'SAMPLE / NO RECORD')

    # SHEET 1: ACTIVE
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
        is_sample = str(r.get('Name Of Benificiary', '')).startswith('SAMPLE / NO RECORD')
        ws1.append([
            _clean_scalar(r.get('BANK_SNO', '')), _clean_scalar(r.get('Name Of Benificiary', '')),
            _clean_scalar(r.get('Father Name', '')), _clean_scalar(r.get('Date OF Birth', '')),
            _clean_scalar(r.get('Age', '')), _clean_scalar(r.get('Aadhar Number', '')),
            _clean_scalar(r.get('Caste', '')), _clean_scalar(r.get('Profession', '')),
            _clean_scalar(r.get('Residential Address', '')), _clean_scalar(r.get('Bank Account Number', '')),
            _clean_scalar(r.get('RD 1', '')), _clean_scalar(r.get('RD ACCOUNT_2 (Govt. Contribution)', '')),
            _clean_scalar(r.get('FILE NAME_x', r.get('FILE NAME', ''))),
            0 if is_sample else _clean_scalar(r.get('Govt_Share', 0)),
            _clean_scalar(r.get('Contact Number', '')), _clean_scalar(r.get('Unit Address', '')),
            _clean_scalar(r.get('BANK_REMARKS', '')), _clean_scalar(r.get('NAME OF CF', '')),
            'SAMPLE / NO RECORD' if is_sample else f'{month_label} Govt_Share Released'
        ])
        ws1.cell(row=r_idx1, column=14).number_format = '₹#,##0.00'
        r_idx1 += 1

    # SHEET 2: MISSING
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
        is_sample = str(r.get('Name Of Benificiary', '')).startswith('SAMPLE / NO RECORD')
        ws2.append([
            _clean_scalar(r.get('BANK_SNO', '')), _clean_scalar(r.get('Name Of Benificiary', '')),
            _clean_scalar(r.get('Father Name', '')), _clean_scalar(r.get('Date OF Birth', '')),
            _clean_scalar(r.get('Age', '')), _clean_scalar(r.get('Aadhar Number', '')),
            _clean_scalar(r.get('Caste', '')), _clean_scalar(r.get('Profession', '')),
            _clean_scalar(r.get('Residential Address', '')), _clean_scalar(r.get('Bank Account Number', '')),
            _clean_scalar(r.get('RD 1', '')), _clean_scalar(r.get('RD ACCOUNT_2 (Govt. Contribution)', '')),
            _clean_scalar(r.get('FILE NAME_x', r.get('FILE NAME', 'Not Assigned'))),
            _clean_scalar(r.get('Contact Number', '')), _clean_scalar(r.get('Unit Address', '')),
            _clean_scalar(r.get('BANK_REMARKS', '')), _clean_scalar(r.get('NAME OF CF', '')),
            'SAMPLE / NO RECORD' if is_sample else f'Zero or missing transactions reported for {month_label}'
        ])

    # SHEET 3: MASTER
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
            _clean_scalar(r.get('BANK_SNO', '')), _clean_scalar(r.get('Name Of Benificiary', '')),
            _clean_scalar(r.get('Father Name', '')), _clean_scalar(r.get('Date OF Birth', '')),
            _clean_scalar(r.get('Age', '')), _clean_scalar(r.get('Aadhar Number', '')),
            _clean_scalar(r.get('Caste', '')), _clean_scalar(r.get('Profession', '')),
            _clean_scalar(r.get('Residential Address', '')), _clean_scalar(r.get('Bank Account Number', '')),
            _clean_scalar(r.get('RD 1', '')), _clean_scalar(r.get('RD ACCOUNT_2 (Govt. Contribution)', '')),
            _clean_scalar(r.get('FILE NAME_x', r.get('FILE NAME', ''))), _clean_scalar(r.get('Govt_Share', 0)),
            _clean_scalar(r.get('Contact Number', '')), _clean_scalar(r.get('Unit Address', '')),
            _clean_scalar(r.get('BANK_REMARKS', '')), _clean_scalar(r.get('NAME OF CF', '')),
            _clean_scalar(r.get('Status', ''))
        ])
        ws3.cell(row=r_idx3, column=14).number_format = '₹#,##0.00'
        r_idx3 += 1

    # SHEET 4: NOT SENT
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
        is_sample = str(r.get('Name Of Benificiary', '')).startswith('SAMPLE / NO RECORD')
        ws4.append([
            _clean_scalar(r.get('BANK_SNO', '')), _clean_scalar(r.get('Name Of Benificiary', '')),
            _clean_scalar(r.get('Father Name', '')), _clean_scalar(r.get('Date OF Birth', '')),
            _clean_scalar(r.get('Age', '')), _clean_scalar(r.get('Aadhar Number', '')),
            _clean_scalar(r.get('Caste', '')), _clean_scalar(r.get('Profession', '')),
            _clean_scalar(r.get('Residential Address', '')), _clean_scalar(r.get('Bank Account Number', '')),
            _clean_scalar(r.get('RD 1', '')), _clean_scalar(r.get('RD ACCOUNT_2 (Govt. Contribution)', '')),
            _clean_scalar(r.get('FILE NAME_x', r.get('FILE NAME', 'Not Assigned'))),
            _clean_scalar(r.get('Contact Number', '')), _clean_scalar(r.get('Unit Address', '')),
            _clean_scalar(r.get('BANK_REMARKS', '')), _clean_scalar(r.get('NAME OF CF', '')),
            _clean_scalar(r.get('BANK SENT', 'SAMPLE / NO RECORD' if is_sample else 'NOT SENT')),
            'SAMPLE / NO RECORD' if is_sample else 'Not Sent to Bank'
        ])

    for ws in [ws1, ws2, ws3, ws4]:
        for col in ws.columns:
            sample_cells = list(col)[:100]
            max_len = max((len(str(cell.value or '')) for cell in sample_cells), default=10)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 16), 50)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@sheets_bp.route('/update_status', methods=['GET'])
def update_status():
    return jsonify(_update_state)


@sheets_bp.route('/update_sheets', methods=['POST'])
def update_sheets():
    try:
        current_month = int(request.form.get('month', 1))
    except (TypeError, ValueError):
        current_month = 1

    target_cf = (request.form.get('cf_name') or '').strip().upper()
    is_all_cfs = target_cf in ('', 'ALL', 'ALL CFS')

    if is_all_cfs:
        with _update_lock:
            if _update_state["running"]:
                flash('⏳ అన్ని CF Google Sheets అప్‌డేట్ ఇప్పటికే బ్యాక్‌గ్రౌండ్‌లో నడుస్తోంది.', 'info')
                return redirect_home()

            _update_state["running"] = True
            _update_state["month"] = current_month
            _update_state["completed_cfs"] = 0
            _update_state["total_cfs"] = len(cf_sheet_ids)
            _update_state["current_cf"] = "Starting..."
            _update_state["last_error"] = ""

        worker = threading.Thread(
            target=_background_all_cfs_worker,
            args=(current_month,),
            daemon=True
        )
        worker.start()
        flash(f'🚀 Month {current_month} కోసం ALL CFs అప్‌డేట్ బ్యాక్‌గ్రౌండ్‌లో ప్రారంభమైంది. Worker timeout రాకుండా అన్ని షీట్స్ అప్‌డేట్ అవుతాయి!', 'success')
        return redirect_home()

    try:
        msg, has_warnings = _run_all_cfs_update(current_month=current_month, target_cf=target_cf, track_progress=False)
        flash(msg, 'warning' if has_warnings else 'success')
    except Exception as e:
        logger.exception("Error in /update_sheets")
        flash(f'❌ Error ఏర్పడింది: {str(e)}', 'danger')

    return redirect_home()


@sheets_bp.route('/download_excel', methods=['POST'])
def download_excel():
    try:
        current_month = int(request.form.get('month', 1))
    except (TypeError, ValueError):
        current_month = 1

    selected_cf = (request.form.get('cf_name') or '').strip()
    if not selected_cf:
        flash('❌ దయచేసి CF పేరు ఎంచుకోండి.', 'danger')
        return redirect_home()

    try:
        _, df_sent_cross, df_not_sent_all, month_label = fetch_and_process_data(current_month)
        cf_col = 'NAME OF CF' if 'NAME OF CF' in df_sent_cross.columns else ('CF NAME' if 'CF NAME' in df_sent_cross.columns else None)

        if cf_col:
            group_df = df_sent_cross[df_sent_cross[cf_col].astype(str).str.strip() == selected_cf].copy()
            if not df_not_sent_all.empty and cf_col in df_not_sent_all.columns:
                not_sent_group_df = df_not_sent_all[df_not_sent_all[cf_col].astype(str).str.strip() == selected_cf].copy()
            else:
                not_sent_group_df = pd.DataFrame()
        else:
            group_df = pd.DataFrame()
            not_sent_group_df = pd.DataFrame()

        excel_stream = generate_cf_excel_in_memory(
            selected_cf, group_df, df_sent_cross, not_sent_group_df, current_month, month_label
        )
        safe_cf_filename = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in selected_cf).strip()
        filename = f"{safe_cf_filename}_Month_{current_month}_Report.xlsx"

        return send_file(
            excel_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logger.exception("Error in /download_excel")
        flash(f'❌ Excel డౌన్‌లోడ్‌లో సమస్య ఏర్పడింది: {str(e)}', 'danger')
        return redirect_home()
