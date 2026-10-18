import io
import sqlite3
from flask import (
    Blueprint,
    jsonify,
    make_response,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
import openpyxl
from openpyxl.styles import Font, PatternFill
import pandas as pd
from rapidfuzz import fuzz

mismatch_bp = Blueprint('mismatch_bp', __name__, template_folder='templates')

DB_PATH = 'thrift_audit.db'


# 1. Database Setup (Added bank_sno and contact_no)
def init_db():
  conn = sqlite3.connect(DB_PATH)
  cursor = conn.cursor()

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS flagged_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id TEXT,
            original_name TEXT,
            matched_name TEXT,
            account_no TEXT,
            rd1_acc TEXT,
            bank_name TEXT,
            bank_sno TEXT,
            contact_no TEXT,
            amount REAL,
            match_score REAL,
            flag_type TEXT,
            status TEXT DEFAULT 'PENDING',
            action_by TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
  conn.commit()
  conn.close()


init_db()


# 2. Dashboard View with Bank List Filter
@mismatch_bp.route('/audit/dashboard')
def audit_dashboard():
  if not session.get('admin_logged_in') and not session.get('cf_logged_in'):
    return redirect('/admin/login')

  selected_bank = request.args.get('bank_name', '')

  conn = sqlite3.connect(DB_PATH)

  bank_list = []
  try:
    df_banks = pd.read_sql_query(
        'SELECT DISTINCT bank_name FROM flagged_records WHERE bank_name IS NOT'
        " NULL AND bank_name != ''",
        conn,
    )
    bank_list = df_banks['bank_name'].tolist()
  except:
    pass

  if not bank_list:
    try:
      from app import SPREADSHEET_NAME, get_sheets_client

      client = get_sheets_client()
      spreadsheet = client.open(SPREADSHEET_NAME)
      df_thrift = pd.DataFrame(
          spreadsheet.worksheet('NEW THRIFT DATA').get_all_records()
      )

      bank_col = (
          'FILE NAME' if 'FILE NAME' in df_thrift.columns else 'Bank Name'
      )
      if bank_col in df_thrift.columns:
        bank_list = (
            df_thrift[bank_col]
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )
        bank_list = [
            b for b in bank_list if b and b != 'nan' and b != 'None'
        ]
        bank_list.sort()
    except Exception as e:
      print('Bank list fetch error from sheet:', e)

  query = 'SELECT * FROM flagged_records'
  params = []
  if selected_bank:
    query += ' WHERE bank_name = ?'
    params.append(selected_bank)

  df_flags = pd.read_sql_query(query, conn, params=params)
  conn.close()

  total_scanned = len(df_flags)
  exact_duplicates = len(df_flags[df_flags['flag_type'] == 'EXACT_DUPLICATE'])
  similar_mismatches = len(
      df_flags[df_flags['flag_type'] == 'SIMILAR_MISMATCH']
  )
  pending_records = len(df_flags[df_flags['status'] == 'PENDING'])
  total_amount_at_risk = df_flags[df_flags['status'] == 'PENDING'][
      'amount'
  ].sum()

  records = df_flags.to_dict(orient='records')

  return render_template(
      'mismatch_dashboard.html',
      records=records,
      bank_list=bank_list,
      selected_bank=selected_bank,
      total_scanned=total_scanned,
      exact_duplicates=exact_duplicates,
      similar_mismatches=similar_mismatches,
      pending_records=pending_records,
      total_amount_at_risk=total_amount_at_risk,
  )


# 3. Bank-wise Audit Scanning API
@mismatch_bp.route('/api/scan-mismatches', methods=['POST'])
def scan_mismatches():
  try:
    data = request.get_json() or {}
    target_bank = data.get('bank_name', '')

    from app import SPREADSHEET_NAME, get_sheets_client

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    df_thrift = pd.DataFrame(
        spreadsheet.worksheet('NEW THRIFT DATA').get_all_records()
    )

    if df_thrift.empty:
      return jsonify({'status': 'error', 'message': 'డేటా ఖాళీగా ఉంది!'})

    name_col = (
        'Name Of Benificiary'
        if 'Name Of Benificiary' in df_thrift.columns
        else 'BENEFICIARY_NAME'
    )
    acc_col = (
        'Bank Account Number'
        if 'Bank Account Number' in df_thrift.columns
        else 'ACCOUNT_NO'
    )
    amt_col = (
        '8% Of Monthly Wages'
        if '8% Of Monthly Wages' in df_thrift.columns
        else 'AMOUNT'
    )
    bank_col = (
        'FILE NAME' if 'FILE NAME' in df_thrift.columns else 'Bank Name'
    )
    bank_sno_col = (
        'BANK_SNO' if 'BANK_SNO' in df_thrift.columns else 'Bank SNO'
    )
    contact_col = (
        'Contact Number'
        if 'Contact Number' in df_thrift.columns
        else 'CONTACT_NUMBER'
    )

    if target_bank:
      df_thrift = df_thrift[
          df_thrift[bank_col].astype(str).str.strip() == target_bank.strip()
      ]

    if df_thrift.empty:
      return jsonify({
          'status': 'error',
          'message': f'సెలెక్ట్ చేసిన బ్యాంక్ ({target_bank}) కి డేటా దొరకలేదు!',
      })

    df_thrift[name_col] = (
        df_thrift[name_col].astype(str).str.strip().str.upper()
    )
    df_thrift = df_thrift[df_thrift[name_col] != '']

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    if target_bank:
      cursor.execute(
          "DELETE FROM flagged_records WHERE bank_name = ? AND status ="
          " 'PENDING'",
          (target_bank,),
      )
    else:
      cursor.execute("DELETE FROM flagged_records WHERE status = 'PENDING'")

    flagged_count = 0

    duplicates = df_thrift[
        df_thrift.duplicated(subset=[name_col], keep=False)
        | df_thrift.duplicated(subset=[acc_col], keep=False)
    ]

    for _, row in duplicates.iterrows():
      name = row[name_col]
      amt = float(row.get(amt_col, 0) or 0)
      acc = str(row.get(acc_col, ''))
      rd1 = str(row.get('RD 1', ''))
      bank = str(row.get(bank_col, ''))
      bank_sno = str(row.get(bank_sno_col, ''))
      contact = str(row.get(contact_col, ''))

      cursor.execute(
          """
                INSERT INTO flagged_records 
                (upload_id, original_name, matched_name, account_no, rd1_acc, bank_name, bank_sno, contact_no, amount, match_score, flag_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
          (
              'SCAN_BATCH_1',
              name,
              name,
              acc,
              rd1,
              bank,
              bank_sno,
              contact,
              amt,
              100.0,
              'EXACT_DUPLICATE',
          ),
      )
      flagged_count += 1

    records_list = df_thrift.to_dict(orient='records')
    n = len(records_list)

    for i in range(n):
      for j in range(i + 1, n):
        r1 = records_list[i]
        r2 = records_list[j]

        name1 = str(r1[name_col])
        name2 = str(r2[name_col])

        if name1 == name2:
          continue

        score = fuzz.ratio(name1, name2)
        if 80 <= score < 100:
          amt = float(r1.get(amt_col, 0) or 0)
          acc = str(r1.get(acc_col, ''))
          rd1 = str(r1.get('RD 1', ''))
          bank = str(r1.get(bank_col, ''))
          bank_sno = str(r1.get(bank_sno_col, ''))
          contact = str(r1.get(contact_col, ''))

          cursor.execute(
              """
                    INSERT INTO flagged_records 
                    (upload_id, original_name, matched_name, account_no, rd1_acc, bank_name, bank_sno, contact_no, amount, match_score, flag_type)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
              (
                  'SCAN_BATCH_1',
                  name1,
                  name2,
                  acc,
                  rd1,
                  bank,
                  bank_sno,
                  contact,
                  amt,
                  score,
                  'SIMILAR_MISMATCH',
              ),
          )
          flagged_count += 1

    conn.commit()
    conn.close()

    return jsonify({
        'status': 'success',
        'message': f'స్కానింగ్ పూర్తయింది! {flagged_count} వివరాలు కనుగొనబడ్డాయి.',
    })

  except Exception as e:
    import traceback

    traceback.print_exc()
    return jsonify(
        {'status': 'error', 'message': f'స్కానింగ్ లో లోపం: {str(e)}'}
    )


# 4. Fetch Details of Specific Beneficiary Mismatches
@mismatch_bp.route('/api/get-beneficiary-matches', methods=['POST'])
def get_beneficiary_matches():
  data = request.get_json() or {}
  name = data.get('name', '')
  bank = data.get('bank_name', '')

  conn = sqlite3.connect(DB_PATH)
  cursor = conn.cursor()

  query = """
        SELECT * FROM flagged_records 
        WHERE (original_name = ? OR matched_name = ?)
    """
  params = [name, name]

  if bank:
    query += ' AND bank_name = ?'
    params.append(bank)

  df = pd.read_sql_query(query, conn, params=params)
  conn.close()

  return jsonify({'status': 'success', 'records': df.to_dict(orient='records')})


# 5. Action Resolution API
@mismatch_bp.route('/api/resolve-mismatch', methods=['POST'])
def resolve_mismatch():
  data = request.get_json()
  record_id = data.get('id')
  action = data.get('action')
  action_by = session.get('admin_name', 'Admin')

  conn = sqlite3.connect(DB_PATH)
  cursor = conn.cursor()
  cursor.execute(
      """
        UPDATE flagged_records 
        SET status = ?, action_by = ? 
        WHERE id = ?
    """,
      (action, action_by, record_id),
  )
  conn.commit()
  conn.close()

  return jsonify(
      {'status': 'success', 'message': f'స్టేటస్ {action} గా మార్చబడింది!'}
  )


# 6. Sync to Google Sheets API
@mismatch_bp.route('/api/sync-to-sheets', methods=['POST'])
def sync_to_sheets():
  try:
    from app import SPREADSHEET_NAME, get_sheets_client

    conn = sqlite3.connect(DB_PATH)
    df_flags = pd.read_sql_query('SELECT * FROM flagged_records', conn)
    conn.close()

    if df_flags.empty:
      return jsonify(
          {'status': 'error', 'message': 'సింక్ చేయడానికి ఎలాంటి డేటా లేదు!'}
      )

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    try:
      ws = spreadsheet.worksheet('AUDIT_FLAGGED_MISMATCHES')
    except:
      ws = spreadsheet.add_worksheet(
          title='AUDIT_FLAGGED_MISMATCHES', rows='1000', cols='15'
      )

    ws.clear()
    data = [df_flags.columns.values.tolist()] + df_flags.values.tolist()
    ws.update(range_name='A1', values=data)

    return jsonify(
        {'status': 'success', 'message': 'డేటా గూగుల్ షీట్‌కు సింక్ అయ్యింది!'}
    )

  except Exception as e:
    return jsonify(
        {'status': 'error', 'message': f'సింక్ చేయడంలో లోపం: {str(e)}'}
    )
