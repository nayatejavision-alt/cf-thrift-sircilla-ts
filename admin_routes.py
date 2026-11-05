from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response
import pandas as pd
import io

admin_bp = Blueprint('admin_bp', __name__, template_folder='templates')

@admin_bp.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        data = request.get_json() if request.is_json else request.form
        username = data.get('username', '').strip()
        password = data.get('password', '').strip()
        
        if username == "admin" and password == "thrift@admin2026":
            session['admin_logged_in'] = True
            session['admin_user'] = username
            return jsonify({"status": "success", "message": "అడ్మిన్ లాగిన్ విజయవంతమైంది!"})
        else:
            return jsonify({"status": "error", "message": "తప్పు యూజర్‌నేమ్ లేదా పాస్‌వర్డ్!"})
            
    return render_template('admin_login.html')

@admin_bp.route('/admin/logout', methods=['POST'])
def admin_logout():
    session.pop('admin_logged_in', None)
    session.pop('admin_user', None)
    return jsonify({"status": "success", "message": "లాగౌట్ అయ్యారు."})

@admin_bp.route('/admin/dashboard')
def admin_dashboard():
    from app import get_sheets_client, SPREADSHEET_NAME
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin_bp.admin_login'))
        
    try:
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        master_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        trans_sheet = spreadsheet.worksheet("Transactions_Year1")
        
        master_records = master_sheet.get_all_records()
        trans_records = trans_sheet.get_all_records()
        
        # 1. డేటాను Pandas DataFrames లోకి లోడ్ చేయడం
        df_master = pd.DataFrame(master_records)
        df_trans = pd.DataFrame(trans_records)
        
        if df_master.empty:
            return render_template('admin_dashboard.html', total=0, completed=0, pending=0, 
                                   batch_summary={}, bank_summary={}, batch_bank_summary={}, 
                                   completed_batches_count=0, total_batches_count=0)

        # కాలమ్ పేర్లను గుర్తించి స్టాండర్డైజ్ చేయడం
        bank_col = 'FILE NAME' if 'FILE NAME' in df_master.columns else ('File Name' if 'File Name' in df_master.columns else df_master.columns[0])
        rd_col = 'RD ACCOUNT_1 (Beneficiary Contribution)' if 'RD ACCOUNT_1 (Beneficiary Contribution)' in df_master.columns else ('RD 1' if 'RD 1' in df_master.columns else None)
        batch_col_m = 'BATCH_NUMBER' if 'BATCH_NUMBER' in df_master.columns else ('BATCH NO' if 'BATCH NO' in df_master.columns else 'Batch')
        
        df_master['Bank_Name'] = df_master[bank_col].astype(str).str.strip().replace('', 'Other')
        df_master['RD_Account'] = df_master[rd_col].astype(str).str.strip() if rd_col else ''
        df_master['Batch_No'] = df_master[batch_col_m].astype(str).str.strip().replace('', '1')

        # 2. ట్రాన్సాక్షన్ డేటాను ప్రాసెస్ చేయడం
        t_rd_col = 'RD ACCOUNT_1 (Beneficiary Contribution)' if 'RD ACCOUNT_1 (Beneficiary Contribution)' in df_trans.columns else ('RD 1' if 'RD 1' in df_trans.columns else None)
        t_batch_col = 'BATCH_NUMBER' if 'BATCH_NUMBER' in df_trans.columns else ('BATCH NO' if 'BATCH NO' in df_trans.columns else 'Batch')
        
        if not df_trans.empty and t_rd_col and t_batch_col:
            df_trans['RD_Account'] = df_trans[t_rd_col].astype(str).str.strip()
            df_trans['Batch_No'] = df_trans[t_batch_col].astype(str).str.strip()
            
            if 'Govt_Share' in df_trans.columns:
                df_trans['Govt_Share'] = pd.to_numeric(df_trans['Govt_Share'], errors='coerce').fillna(0)
            else:
                df_trans['Govt_Share'] = 0

            # కంప్లీట్ అయిన రికార్డులను యూనివర్శల్‌గా ఫిల్టర్ చేయడం
            completed_trans = df_trans[df_trans['Govt_Share'] > 0][['RD_Account', 'Batch_No']].drop_duplicates()
            completed_trans['Is_Completed'] = True
            
            # మాస్టర్ డేటాతో మెర్జ్ చేయడం (Merge Master with Transactions)
            df_merged = pd.merge(df_master, completed_trans, on=['RD_Account', 'Batch_No'], how='left')
        else:
            df_merged = df_master.copy()
            df_merged['Is_Completed'] = False

        df_merged['Is_Completed'] = df_merged['Is_Completed'].fillna(False)
        df_merged['Status_Val'] = df_merged['Is_Completed'].astype(int)

        total_beneficiaries = len(df_merged)
        total_completed_all = int(df_merged['Status_Val'].sum())
        total_pending_all = total_beneficiaries - total_completed_all

        # 3. Pandas Groupby ద్వారా బ్యాచ్ వారీగా సమ్మరీ
        batch_grouped = df_merged.groupby('Batch_No').agg(
            total=('RD_Account', 'count'),
            completed=('Status_Val', 'sum')
        ).reset_index()
        batch_grouped['pending'] = batch_grouped['total'] - batch_grouped['completed']
        
        batch_summary = {}
        for _, row in batch_grouped.iterrows():
            batch_summary[str(row['Batch_No'])] = {
                "total": int(row['total']),
                "completed": int(row['completed']),
                "pending": int(row['pending'])
            }

        # 4. Pandas Groupby ద్వారా బ్యాంక్ వారీగా సమ్మరీ
        bank_grouped = df_merged.groupby('Bank_Name').agg(
            total=('RD_Account', 'count'),
            completed=('Status_Val', 'sum')
        ).reset_index()
        bank_grouped['pending'] = bank_grouped['total'] - bank_grouped['completed']
        
        bank_summary = {}
        for _, row in bank_grouped.iterrows():
            bank_summary[str(row['Bank_Name'])] = {
                "total": int(row['total']),
                "completed": int(row['completed']),
                "pending": int(row['pending'])
            }

        # 5. Pandas Groupby ద్వారా బ్యాచ్ + బ్యాంక్ మ్యాట్రిక్స్ సమ్మరీ
        batch_bank_grouped = df_merged.groupby(['Batch_No', 'Bank_Name']).agg(
            total=('RD_Account', 'count'),
            completed=('Status_Val', 'sum')
        ).reset_index()
        batch_bank_grouped['pending'] = batch_bank_grouped['total'] - batch_bank_grouped['completed']

        batch_bank_summary = {}
        for _, row in batch_bank_grouped.iterrows():
            key = (str(row['Batch_No']), str(row['Bank_Name']))
            batch_bank_summary[key] = {
                "total": int(row['total']),
                "completed": int(row['completed']),
                "pending": int(row['pending'])
            }

        completed_batches_count = sum(1 for b, stats in batch_summary.items() if stats["pending"] == 0 and stats["total"] > 0)
        total_batches_count = len(batch_summary)

        return render_template('admin_dashboard.html', 
                               total=total_beneficiaries, 
                               completed=total_completed_all, 
                               pending=total_pending_all, 
                               batch_summary=batch_summary,
                               bank_summary=bank_summary,
                               batch_bank_summary=batch_bank_summary,
                               completed_batches_count=completed_batches_count,
                               total_batches_count=total_batches_count)
                               
    except Exception as e:
        print(f"Admin Dashboard Error with Pandas: {repr(e)}")
        return f"లోపం సంభవించింది: {str(e)}"
