from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response
import csv
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
        
        # 1. కంప్లీటెడ్ ట్రాన్సాక్షన్స్ గుర్తించడం (సూచిక: (RD_ACCOUNT, BATCH_NUMBER))
        completed_batch_set = set()
        for row in trans_records:
            t_batch = str(row.get("BATCH_NUMBER") or row.get("BATCH NO") or "").strip()
            t_rd1 = str(row.get("RD ACCOUNT_1 (Beneficiary Contribution)") or row.get("RD 1") or "").strip()
            try:
                govt_share = float(row.get("Govt_Share") or 0)
            except ValueError:
                govt_share = 0.0
                
            if govt_share > 0:
                completed_batch_set.add((t_rd1, t_batch))
                
        total_beneficiaries = len(master_records)
        
        # 2. డైనమిక్ బ్యాచ్ సమ్మరీ (Batch-wise Summary)
        batch_summary = {}
        # 3. బ్యాంక్ వారీగా సమ్మరీ (Bank-wise Summary)
        bank_summary = {}
        # 4. బ్యాచ్ & బ్యాంక్ కంబైన్డ్ సమ్మరీ (Batch + Bank Matrix)
        batch_bank_summary = {}

        total_completed_all = 0
        total_pending_all = 0

        for row in master_records:
            bank_name = str(row.get("FILE NAME") or row.get("File Name") or "Other").strip()
            bank_rd1 = str(row.get("RD ACCOUNT_1 (Beneficiary Contribution)") or row.get("RD 1") or "").strip()
            batch_no = str(row.get("BATCH_NUMBER") or row.get("BATCH NO") or "1").strip()
            
            # చెక్: ఈ RD మరియు ఈ బ్యాచ్ నెంబర్‌తో ట్రాన్సాక్షన్ పూర్తయిందా?
            is_completed = (bank_rd1, batch_no) in completed_batch_set
            
            if is_completed:
                total_completed_all += 1
            else:
                total_pending_all += 1

            # Batch Summary ని డైనమిక్‌గా బిల్డ్ చేయడం
            if batch_no not in batch_summary:
                batch_summary[batch_no] = {"total": 0, "completed": 0, "pending": 0}
            batch_summary[batch_no]["total"] += 1
            if is_completed:
                batch_summary[batch_no]["completed"] += 1
            else:
                batch_summary[batch_no]["pending"] += 1

            # Bank Summary బిల్డ్ చేయడం
            if bank_name not in bank_summary:
                bank_summary[bank_name] = {"total": 0, "completed": 0, "pending": 0}
            bank_summary[bank_name]["total"] += 1
            if is_completed:
                bank_summary[bank_name]["completed"] += 1
            else:
                bank_summary[bank_name]["pending"] += 1

            # Batch + Bank Combined Summary
            key = (batch_no, bank_name)
            if key not in batch_bank_summary:
                batch_bank_summary[key] = {"total": 0, "completed": 0, "pending": 0}
            batch_bank_summary[key]["total"] += 1
            if is_completed:
                batch_bank_summary[key]["completed"] += 1
            else:
                batch_bank_summary[key]["pending"] += 1

        # బ్యాచ్ కంప్లీషన్ లెక్కించడం (ఒక బ్యాచ్ లో పెండింగ్ లేకపోతే అది పూర్తిగా Completed Batch అవుతుంది)
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
        print(f"Admin Dashboard Error: {repr(e)}")
        return f"లోపం సంభవించింది: {str(e)}"
