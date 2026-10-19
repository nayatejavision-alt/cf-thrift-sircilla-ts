import io
import sqlite3
import pandas as pd
from flask import Blueprint, render_template, request, jsonify, session, redirect, send_file
from rapidfuzz import fuzz

mismatch_bp = Blueprint('mismatch_bp', __name__, template_folder='templates')

DB_PATH = 'thrift_audit.db'

# 1. Database Setup with bank_sno & contact_no
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS flagged_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id TEXT,
            bank_sno TEXT,
            original_name TEXT,
            matched_name TEXT,
            account_no TEXT,
            rd1_acc TEXT,
            bank_name TEXT,
            contact_no TEXT,
            amount REAL,
            match_score REAL,
            flag_type TEXT,
            status TEXT DEFAULT 'PENDING',
            action_by TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    try:
        cursor.execute("ALTER TABLE flagged_records ADD COLUMN bank_sno TEXT")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE flagged_records ADD COLUMN contact_no TEXT")
    except:
        pass
        
    conn.commit()
    conn.close()

init_db()

# 2. Dashboard View with Bank List Filter & Audit Sheet Check
@mismatch_bp.route('/audit/dashboard')
def audit_dashboard():
    if not session.get('bank_user'):
        return redirect('/')

    selected_bank = session.get('bank_user')

    conn = sqlite3.connect(DB_PATH)
    bank_list = [selected_bank] if selected_bank else []

    # మొదటిసారి లోడ్ అయ్యేటప్పుడు AUDIT_FLAGGED_MISMATCHES షీట్ నుండి డేటా ఉంటే లోకల్ DB లోకి తెచ్చుకోవడం
    if selected_bank:
        try:
            from app import get_sheets_client, SPREADSHEET_NAME
            client = get_sheets_client()
            spreadsheet = client.open(SPREADSHEET_NAME)
            ws_audit = spreadsheet.worksheet('AUDIT_FLAGGED_MISMATCHES')
            audit_records = ws_audit.get_all_records()
            df_audit = pd.DataFrame(audit_records)
            
            if not df_audit.empty and 'bank_name' in df_audit.columns:
                bank_audit_data = df_audit[df_audit['bank_name'].astype(str).str.strip() == selected_bank.strip()]
                if not bank_audit_data.empty:
                    # లోకల్ DB లో పాతవి ఉన్నా లేకపోయినా షీట్ డేటాతో సింక్ ఉంచడం
                    cursor = conn.cursor()
                    cursor.execute("SELECT COUNT(*) FROM flagged_records WHERE bank_name = ?", (selected_bank,))
                    count = cursor.fetchone()[0]
                    if count == 0:
                        for _, row in bank_audit_data.iterrows():
                            cursor.execute("""
                                INSERT INTO flagged_records 
                                (upload_id, bank_sno, original_name, matched_name, account_no, rd1_acc, bank_name, contact_no, amount, match_score, flag_type, status, action_by)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                str(row.get('upload_id', 'SCAN_BATCH_1')),
                                str(row.get('bank_sno', '')),
                                str(row.get('original_name', '')),
                                str(row.get('matched_name', '')),
                                str(row.get('account_no', '')),
                                str(row.get('rd1_acc', '')),
                                str(row.get('bank_name', '')),
                                str(row.get('contact_no', '')),
                                float(row.get('amount', 0) or 0),
                                float(row.get('match_score', 0) or 0),
                                str(row.get('flag_type', '')),
                                str(row.get('status', 'PENDING')),
                                str(row.get('action_by', ''))
                            ))
                        conn.commit()
        except Exception as e:
            print("Audit sheet load check error:", e)

    query = "SELECT * FROM flagged_records"
    params = []
    if selected_bank:
        query += " WHERE bank_name = ?"
        params.append(selected_bank)

    df_flags = pd.read_sql_query(query, conn, params=params)
    conn.close()

    total_scanned = len(df_flags)
    exact_duplicates = len(df_flags[df_flags['flag_type'] == 'EXACT_DUPLICATE'])
    similar_mismatches = len(df_flags[df_flags['flag_type'] == 'SIMILAR_MISMATCH'])
    pending_records = len(df_flags[df_flags['status'] == 'PENDING'])
    total_amount_at_risk = df_flags[df_flags['status'] == 'PENDING']['amount'].sum()

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
        total_amount_at_risk=total_amount_at_risk
    )

# 3. Bank-wise Audit Scanning API
@mismatch_bp.route('/api/scan-mismatches', methods=['POST'])
def scan_mismatches():
    try:
        data = request.get_json() or {}
        target_bank = data.get('bank_name', '')
        
        from app import get_sheets_client, SPREADSHEET_NAME
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        df_thrift = pd.DataFrame(spreadsheet.worksheet('NEW THRIFT DATA').get_all_records())
        
        if df_thrift.empty:
            return jsonify({"status": "error", "message": "డేటా ఖాళీగా ఉంది!"})
            
        name_col = 'Name Of Benificiary' if 'Name Of Benificiary' in df_thrift.columns else 'BENEFICIARY_NAME'
        acc_col = 'Bank Account Number' if 'Bank Account Number' in df_thrift.columns else 'ACCOUNT_NO'
        amt_col = '8% Of Monthly Wages' if '8% Of Monthly Wages' in df_thrift.columns else 'AMOUNT'
        bank_col = 'FILE NAME' if 'FILE NAME' in df_thrift.columns else 'Bank Name'
        
        if target_bank:
            df_thrift = df_thrift[df_thrift[bank_col].astype(str).str.strip() == target_bank.strip()]
            
        if df_thrift.empty:
            return jsonify({"status": "error", "message": f"సెలెక్ట్ చేసిన బ్యాంక్ ({target_bank}) కి డేటా దొరకలేదు!"})
            
        df_thrift[name_col] = df_thrift[name_col].astype(str).str.strip().str.upper()
        df_thrift = df_thrift[df_thrift[name_col] != '']

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        if target_bank:
            cursor.execute("DELETE FROM flagged_records WHERE bank_name = ? AND status = 'PENDING'", (target_bank,))
        else:
            cursor.execute("DELETE FROM flagged_records WHERE status = 'PENDING'")
            
        flagged_count = 0
        
        # Exact Duplicate Search
        duplicates = df_thrift[
            df_thrift.duplicated(subset=[name_col], keep=False) | 
            df_thrift.duplicated(subset=[acc_col], keep=False)
        ]
        
        for _, row in duplicates.iterrows():
            name = row[name_col]
            bank_sno = str(row.get('BANK_SNO', ''))
            contact = str(row.get('Contact Number', ''))
            amt = float(row.get(amt_col, 0) or 0)
            acc = str(row.get(acc_col, ''))
            rd1 = str(row.get('RD 1', ''))
            bank = str(row.get(bank_col, ''))
            
            cursor.execute("""
                INSERT INTO flagged_records 
                (upload_id, bank_sno, original_name, matched_name, account_no, rd1_acc, bank_name, contact_no, amount, match_score, flag_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, ('SCAN_BATCH_1', bank_sno, name, name, acc, rd1, bank, contact, amt, 100.0, 'EXACT_DUPLICATE'))
            flagged_count += 1
            
        # Fuzzy Matching
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
                    bank_sno = str(r1.get('BANK_SNO', ''))
                    contact = str(r1.get('Contact Number', ''))
                    amt = float(r1.get(amt_col, 0) or 0)
                    acc = str(r1.get(acc_col, ''))
                    rd1 = str(r1.get('RD 1', ''))
                    bank = str(r1.get(bank_col, ''))
                    
                    cursor.execute("""
                        INSERT INTO flagged_records 
                        (upload_id, bank_sno, original_name, matched_name, account_no, rd1_acc, bank_name, contact_no, amount, match_score, flag_type)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, ('SCAN_BATCH_1', bank_sno, name1, name2, acc, rd1, bank, contact, amt, score, 'SIMILAR_MISMATCH'))
                    flagged_count += 1
                    
        cursor.execute("""
            DELETE FROM flagged_records 
            WHERE id NOT IN (
                SELECT MIN(id) 
                FROM flagged_records 
                GROUP BY bank_name, bank_sno, original_name, matched_name, flag_type
            ) AND status = 'PENDING'
        """)

        if target_bank:
            cursor.execute("SELECT COUNT(*) FROM flagged_records WHERE bank_name = ? AND status = 'PENDING'", (target_bank,))
        else:
            cursor.execute("SELECT COUNT(*) FROM flagged_records WHERE status = 'PENDING'")
        
        final_count = cursor.fetchone()[0]
        
        conn.commit()
        conn.close()
        
        return jsonify({"status": "success", "message": f"స్కానింగ్ పూర్తయింది! {final_count} రికార్డులు నివేదించబడ్డాయి."})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"status": "error", "message": f"స్కానింగ్ లో లోపం: {str(e)}"})

# 4. Fetch Details of Specific Beneficiary Mismatches (Without Duplicates)
@mismatch_bp.route('/api/get-beneficiary-matches', methods=['POST'])
def get_beneficiary_matches():
    data = request.get_json() or {}
    name = data.get('name', '')
    bank = data.get('bank_name', '')
    
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT DISTINCT bank_sno, original_name, matched_name, account_no, rd1_acc, contact_no, amount, flag_type, match_score 
        FROM flagged_records 
        WHERE (original_name = ? OR matched_name = ?)
    """
    params = [name, name]
    
    if bank:
        query += " AND bank_name = ?"
        params.append(bank)
        
    df = pd.read_sql_query(query, conn, params=params)
    if not df.empty:
        df = df.drop_duplicates(subset=['bank_sno', 'original_name', 'account_no', 'flag_type'])
    conn.close()
    
    return jsonify({"status": "success", "records": df.to_dict(orient='records')})

# 5. Action Resolution API with Selective Google Sheet Update
@mismatch_bp.route('/api/resolve-mismatch', methods=['POST'])
def resolve_mismatch():
    data = request.get_json()
    record_id = data.get('id')
    action = data.get('action')
    action_by = session.get('admin_name', 'Admin')
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # ఈ రికార్డ్ ఏ బ్యాంకుకు చెందిందో కనుగొనడం
    cursor.execute("SELECT bank_name FROM flagged_records WHERE id = ?", (record_id,))
    row = cursor.fetchone()
    target_bank = row[0] if row else ''

    cursor.execute("""
        UPDATE flagged_records 
        SET status = ?, action_by = ? 
        WHERE id = ?
    """, (action, action_by, record_id))
    conn.commit()
    conn.close()
    
    # సెలెక్టివ్ అప్‌డేట్: కేవలం ఈ బ్యాంకు డేటాను మాత్రమే గూగుల్ షీట్‌లో అప్‌డేట్ చేయడం
    if target_bank:
        try:
            from app import get_sheets_client, SPREADSHEET_NAME
            client = get_sheets_client()
            spreadsheet = client.open(SPREADSHEET_NAME)
            
            try:
                ws = spreadsheet.worksheet('AUDIT_FLAGGED_MISMATCHES')
                existing_data = ws.get_all_records()
                df_existing = pd.DataFrame(existing_data)
            except:
                ws = spreadsheet.add_worksheet(title='AUDIT_FLAGGED_MISMATCHES', rows="1000", cols="15")
                df_existing = pd.DataFrame()

            conn = sqlite3.connect(DB_PATH)
            df_target = pd.read_sql_query("SELECT * FROM flagged_records WHERE bank_name = ?", conn, params=[target_bank])
            conn.close()

            if not df_existing.empty and 'bank_name' in df_existing.columns:
                df_other_banks = df_existing[df_existing['bank_name'].astype(str).str.strip() != target_bank.strip()]
                df_final = pd.concat([df_other_banks, df_target], ignore_index=True)
            else:
                df_final = df_target

            ws.clear()
            sheet_data = [df_final.columns.values.tolist()] + df_final.values.tolist()
            ws.update(range_name='A1', values=sheet_data)
        except Exception as e:
            print("Selective sheet sync error:", e)
    
    return jsonify({"status": "success", "message": f"స్టేటస్ {action} గా మార్చబడింది!"})

# 6. Full Sync to Google Sheets API
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
        
        try:
            ws = spreadsheet.worksheet('AUDIT_FLAGGED_MISMATCHES')
        except:
            ws = spreadsheet.add_worksheet(title='AUDIT_FLAGGED_MISMATCHES', rows="1000", cols="15")
            
        ws.clear()
        data = [df_flags.columns.values.tolist()] + df_flags.values.tolist()
        ws.update(range_name='A1', values=data)
        
        return jsonify({"status": "success", "message": "డేటా గూగుల్ షీట్‌కు సింక్ అయ్యింది!"})
        
    except Exception as e:
        return jsonify({"status": "error", "message": f"సింక్ చేయడంలో లోపం: {str(e)}"})

# 7. ఆడిట్ డేటాను ఎక్సెల్ (.xlsx) గా డౌన్‌లోడ్ చేసే API
@mismatch_bp.route('/api/download-audit-excel')
def download_audit_excel():
    try:
        conn = sqlite3.connect(DB_PATH)
        selected_bank = request.args.get('bank_name', '')
        
        query = "SELECT * FROM flagged_records"
        params = []
        if selected_bank:
            query += " WHERE bank_name = ?"
            params.append(selected_bank)

        df = pd.read_sql_query(query, conn, params=params)
        conn.close()

        if df.empty:
            return jsonify({"status": "error", "message": "డౌన్‌లోడ్ చేయడానికి ఎలాంటి డేటా లేదు!"})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Audit_Mismatches')
        output.seek(0)

        filename = f"Thrift_Audit_Mismatches_{selected_bank or 'All_Banks'}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"status": "error", "message": f"డౌన్‌లోడ్ చేయడంలో లోపం: {str(e)}"})
