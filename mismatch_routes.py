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


# 1. SQLite డేటాబేస్ సెటప్ (In-built)
def init_db():
  conn = sqlite3.connect(DB_PATH)
  cursor = conn.cursor()

  # ఆడిట్ టేబుల్
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS flagged_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id TEXT,
            original_name TEXT,
            matched_name TEXT,
            account_no TEXT,
            rd1_acc TEXT,
            bank_name TEXT,
            amount REAL,
            match_score REAL,
            flag_type TEXT, -- 'EXACT_DUPLICATE' or 'SIMILAR_MISMATCH'
            status TEXT DEFAULT 'PENDING', -- 'PENDING', 'APPROVED', 'REJECTED'
            action_by TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
  conn.commit()
  conn.close()


# యాప్ స్టార్ట్ అయినప్పుడు DB ఇనిషియలైజ్ అవ్వడం
init_db()


# 2. డాష్‌బోర్డ్ వ్యూ రౌట్
@mismatch_bp.route('/audit/dashboard')
def audit_dashboard():
  if not session.get('admin_logged_in') and not session.get('cf_logged_in'):
    return redirect('/admin/login')

  conn = sqlite3.connect(DB_PATH)
  df_flags = pd.read_sql_query('SELECT * FROM flagged_records', conn)
  conn.close()

  # Widgets Data
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
      total_scanned=total_scanned,
      exact_duplicates=exact_duplicates,
      similar_mismatches=similar_mismatches,
      pending_records=pending_records,
      total_amount_at_risk=total_amount_at_risk,
  )


# 3. గూగుల్ షీట్ స్కాన్ చేసి ఫజ్జీ మ్యాచింగ్ రన్ చేసే API
@mismatch_bp.route('/api/scan-mismatches', methods=['POST'])
def scan_mismatches():
  try:
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

    names = df_thrift[name_col].astype(str).str.strip().tolist()

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM flagged_records WHERE status = 'PENDING'"
    )  # పాత పెండింగ్ రికార్డులను క్లియర్ చేయడం

    flagged_count = 0

    # Fuzzy Matching & Duplicate Scan
    for i in range(len(names)):
      for j in range(i + 1, len(names)):
        name1 = names[i].upper()
        name2 = names[j].upper()

        if not name1 or not name2:
          continue

        # Similarity Score కాలిక్యులేషన్ (80% నుండి 100%)
        score = fuzz.ratio(name1, name2)

        if score >= 80:
          flag_type = 'EXACT_DUPLICATE' if score == 100 else 'SIMILAR_MISMATCH'
          row_data = df_thrift.iloc[i]

          amt = float(row_data.get(amt_col, 0) or 0)
          acc = str(row_data.get(acc_col, ''))
          rd1 = str(row_data.get('RD 1', ''))
          bank = str(
              row_data.get('FILE NAME') or row_data.get('Bank Name') or ''
          )

          cursor.execute(
              """
                        INSERT INTO flagged_records 
                        (upload_id, original_name, matched_name, account_no, rd1_acc, bank_name, amount, match_score, flag_type)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
              (
                  'SCAN_BATCH_1',
                  names[i],
                  names[j],
                  acc,
                  rd1,
                  bank,
                  amt,
                  score,
                  flag_type,
              ),
          )
          flagged_count += 1

    conn.commit()
    conn.close()

    return jsonify({
        'status': 'success',
        'message': f'స్కానింగ్ పూర్తయింది! {flagged_count} అనుమానాస్పద రికార్డులు గుర్తించబడ్డాయి.',
    })

  except Exception as e:
        return jsonify({"status": "error", "message": f"స్కానింగ్ లో లోపం: {str(e)}"})

# 4. ఆమోదం / రిజెక్షన్ API (Action Handle)
@mismatch_bp.route('/api/resolve-mismatch', methods=['POST'])
def resolve_mismatch():
    data = request.get_json()
    record_id = data.get('id')
    action = data.get('action') # 'APPROVED' or 'REJECTED'
    action_by = session.get('admin_name', 'Admin')

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE flagged_records 
        SET status = ?, action_by = ? 
        WHERE id = ?
    """, (action, action_by, record_id))
    conn.commit()
    conn.close()

    return jsonify({"status": "success", "message": f"స్టేటస్ {action} గా నవీకరించబడింది!"})

# 5. ఫలితాలను గూగుల్ షీట్‌లోకి సింక్ చేసే API (Sync to Google Sheets)
@mismatch_bp.route('/api/sync-to-sheets', methods=['POST'])
def sync_to_sheets():
    try:
        from app import get_sheets_client, SPREADSHEET_NAME
        
        conn = sqlite3.connect(DB_PATH)
        df_flags = pd.read_sql_query("SELECT * FROM flagged_records", conn)
        conn.close()

        if df_flags.empty:
            return jsonify({"status": "error", "message": "సింక్ చేయడానికి ఎలాంటి డేటా లేదు!"})

        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)

        # గూగుల్ షీట్ లో కొత్త 'AUDIT_FLAGGED_MISMATCHES' ట్యాబ్ సృష్టించడం లేదా ఓపెన్ చేయడం
        try:
            ws = spreadsheet.worksheet('AUDIT_FLAGGED_MISMATCHES')
        except:
            ws = spreadsheet.add_worksheet(title='AUDIT_FLAGGED_MISMATCHES', rows="1000", cols="15")

        ws.clear()
        data = [df_flags.columns.values.tolist()] + df_flags.values.tolist()
        ws.update(range_name='A1', values=data)

        return jsonify({"status": "success", "message": "డేటా గూగుల్ షీట్‌లోకి పర్ఫెక్ట్‌గా సింక్ అయ్యింది!"})

    except Exception as e:
        return jsonify({"status": "error", "message": f"సింక్ చేయడంలో లోపం: {str(e)}"})
