from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify
import pandas as pd
# మీ ప్రధాన అప్ నుండి గూగుల్ షీట్స్ క్లైంట్ మరియు కాన్స్టెంట్స్ ఇంపోర్ట్ చేసుకోవాలి
#from app import get_sheets_client, SPREADSHEET_NAME 

admin_bp = Blueprint('admin_bp', __name__, template_folder='templates')

# 1. అడ్మిన్ లాగిన్ పేజీ / రౌట్
@admin_bp.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        data = request.get_json() if request.is_json else request.form
        username = data.get('username', '').strip()
        password = data.get('password', '').strip()
        
        # ఇక్కడ మీరు అడ్మిన్ యూజర్‌నేమ్ మరియు పాస్‌వర్డ్ సెట్ చేసుకోవచ్చు
        if username == "admin" and password == "thrift@admin2026":
            session['admin_logged_in'] = True
            session['admin_user'] = username
            return jsonify({"status": "success", "message": "అడ్మిన్ లాగిన్ విజయవంతమైంది!"})
        else:
            return jsonify({"status": "error", "message": "తప్పు యూజర్‌నేమ్ లేదా పాస్‌వర్డ్!"})
            
    return render_template('admin_login.html')

# 2. అడ్మిన్ లాగౌట్
@admin_bp.route('/admin/logout', methods=['POST'])
def admin_logout():
    session.pop('admin_logged_in', None)
    session.pop('admin_user', None)
    return jsonify({"status": "success", "message": "లాగౌట్ అయ్యారు."})

# 3. ఓవరాల్ డాష్‌బోర్డ్ & బ్యాంక్-వైజ్ రిపోర్ట్స్
@admin_bp.route('/admin/dashboard')
def admin_dashboard():
    from app import get_sheets_client, SPREADSHEET_NAME
    # అడ్మిన్ లాగిన్ అయ్యారా లేదా అని చెక్ చేయడం
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin_bp.admin_login'))
        
    try:
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        # మాస్టర్ డేటా మరియు ట్రాన్సాక్షన్ డేటా తెచ్చుకోవడం
        master_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        trans_sheet = spreadsheet.worksheet("Transactions_Year1")
        
        master_records = master_sheet.get_all_records()
        trans_records = trans_sheet.get_all_records()
        
        total_beneficiaries = len(master_records)
        
        # కంప్లీటెడ్ ట్రాన్సాక్షన్స్ సెట్ (Govt_Share > 0 ఉన్నవి)
        completed_set = set()
        for row in trans_records:
            t_sno = str(row.get("BANK_SNO") or row.get("Bank SNO") or "").strip()
            t_month = str(row.get("MONTH_NUMBER") or row.get("month_number") or "").strip()
            try:
                govt_share = float(row.get("Govt_Share") or 0)
            except ValueError:
                govt_share = 0.0
                
            if t_month == "1" and govt_share > 0:
                completed_set.add(t_sno)
                
        completed_count = len(completed_set)
        pending_count = total_beneficiaries - completed_count
        
        # బ్యాంక్ వారీగా సమ్మరీ తయారు చేయడం (Bank-wise Summary)
        bank_summary = {}
        for row in master_records:
            bank_name = str(row.get("FILE NAME") or row.get("File Name") or "Other").strip()
            bank_sno = str(row.get("BANK_SNO") or row.get("Bank SNO") or "").strip()
            
            if bank_name not in bank_summary:
                bank_summary[bank_name] = {"total": 0, "completed": 0, "pending": 0}
                
            bank_summary[bank_name]["total"] += 1
            if bank_sno in completed_set:
                bank_summary[bank_name]["completed"] += 1
            else:
                bank_summary[bank_name]["pending"] += 1
                
        return render_template('admin_dashboard.html', 
                               total=total_beneficiaries, 
                               completed=completed_count, 
                               pending=pending_count, 
                               bank_summary=bank_summary)
                               
    except Exception as e:
        print(f"Admin Dashboard Error: {repr(e)}")
        return f"లోపం సంభవించింది: {str(e)}"
