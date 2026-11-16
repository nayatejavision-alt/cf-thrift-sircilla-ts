import io  # ఫైల్‌ను రామ్‌లో (RAM) ఆబ్జెక్ట్‌గా సేవ్ చేయడానికి
import os
from flask import Flask, render_template, request, redirect, url_for, flash, Blueprint, jsonify, session, send_file
from src.parser import parse_term_deposit_file, parse_prn_file
import pandas as pd

rd_bp = Blueprint('rd_bp', __name__, template_folder='templates')

UPLOAD_FOLDER = 'temp_file.txt'

@rd_bp.route('/rd/index')
def index():
    return render_template('rd_view.html')

@rd_bp.route('/rd/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'GET':
        return redirect(url_for('rd_bp.index'))
    
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    
    if file:
        file.save(UPLOAD_FOLDER)
        df = parse_term_deposit_file(UPLOAD_FOLDER)
        
        if not df.empty:
            total_balance = f"{df['Balance'].sum():,.2f}"
            
            summary = df.groupby('Section')['Balance'].agg(['count', 'sum']).reset_index()
            summary.columns = ['సెక్షన్', 'ఖాతాల సంఖ్య', 'మొత్తం బ్యాలెన్స్ (₹)']
            summary['మొత్తం బ్యాలెన్స్ (₹)'] = summary['మొత్తం బ్యాలెన్స్ (₹)'].map('{:,.2f}'.format)
            summary_table = summary.to_html(classes='table table-striped', index=False)
            
            data_table = df.to_html(classes='table table-striped', index=False)
            
            return render_template('rd_dashboard.html', 
                                   total_balance=total_balance, 
                                   summary_table=summary_table, 
                                   data_table=data_table)
        
    return "డెటా ప్రాసెస్ చేయడంలో లోపం ఏర్పడింది.", 400

@rd_bp.route('/compare', methods=['GET', 'POST'])
def compare_files():
    if request.method == 'GET':
        return render_template('rd_view.html')
        
    if 'file1' not in request.files or 'file2' not in request.files:
        flash('దయచేసి రెండు ఫైళ్లను సెలెక్ట్ చేయండి!')
        return redirect(request.url)
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    if file1.filename == '' or file2.filename == '':
        flash('ఫైల్ సెలెక్ట్ చేయబడలేదు!')
        return redirect(request.url)
    
    upload_folder = '/tmp'
    os.makedirs(upload_folder, exist_ok=True)
    
    path1 = os.path.join(upload_folder, file1.filename)
    path2 = os.path.join(upload_folder, file2.filename)
    
    file1.save(path1)
    file2.save(path2)

    session['file1_name'] = file1.filename
    session['file2_name'] = file2.filename
    
    try:
        df1 = parse_term_deposit_file(path1)
        df2 = parse_prn_file(path2)
        
        # --- ముఖ్యమైన లాజిక్ (Data Cleaning & Validation) ---
        # 1. అకౌంట్ నంబర్లను స్ట్రింగ్ ఫార్మాట్‌కు మార్చడం (లీడింగ్ జీరోలు పోకుండా ఉండటానికి)
        if 'Account_No' in df1.columns:
            df1['Account_No'] = df1['Account_No'].astype(str).str.strip()
        if 'RD1 Acc Num' in df2.columns:
            df2['RD1 Acc Num'] = df2['RD1 Acc Num'].astype(str).str.strip()
            
            # 2. 'ACCOUNT CLOSED', 'TEMP ADDRESS' వంటి టెక్స్ట్ ఉన్న రోస్‌ను వేరు చేయడం (Data Integrity)
            # కేవలం డిజిట్లు (నెంబర్లు) ఉన్న అకౌంట్ నంబర్లను మాత్రమే మ్యాచింగ్ కోసం ఉంచుతాం
            is_numeric_acc = df2['RD1 Acc Num'].str.match(r'^\d+$', na=False)
            df2_valid = df2[is_numeric_acc].copy()
            df2_invalid = df2[~is_numeric_acc].copy()  # టెక్స్ట్/స్టేటస్ ఉన్న ప్రత్యేక రోస్
        else:
            df2_valid = df2.copy()
            df2_invalid = pd.DataFrame()

        # 3. రెండు డేటాఫ్రేమ్‌లను Account_No మరియు RD1 Acc Num ఆధారంగా మెర్జ్ చేయడం (Outer Join)
        merged_df = pd.merge(df1, df2_valid, left_on='Account_No', right_on='RD1 Acc Num', how='outer', indicator=True)
        
        matched_data = merged_df[merged_df['_merge'] == 'both']
        unmatched_data = merged_df[merged_df['_merge'] != 'both']
        
        # డ్యాష్‌బోర్డ్‌లో ప్రదర్శించడానికి హెచ్‌టిఎంఎల్ టేబుల్స్ సిద్ధం చేయడం
        matched_table = matched_data.head(100).to_html(classes='table table-striped', index=False)
        
        return render_template('rd_dashboard.html', 
                               matched_count=len(matched_data), 
                               unmatched_count=len(unmatched_data),
                               invalid_count=len(df2_invalid),
                               tables=[matched_table])
                           
    except Exception as e:
        flash(f'ఫైల్స్ ప్రాసెస్ చేస్తున్నప్పుడు ఎర్రర్ వచ్చింది: {str(e)}')
        return redirect(request.url)


@rd_bp.route('/download/<type>')
def download_excel(type):
    path1 = os.path.join('/tmp', session.get('file1_name', ''))
    path2 = os.path.join('/tmp', session.get('file2_name', ''))
    
    if not os.path.exists(path1) or not os.path.exists(path2):
        flash("ఫైల్స్ లభించలేదు, దయచేసి మళ్ళీ అప్‌లోడ్ చేయండి.")
        return redirect(url_for('rd_bp.index'))

    df1 = parse_term_deposit_file(path1)
    df2 = parse_prn_file(path2)
    
    # సేమ్ డేటా క్లీనింగ్ లాజిక్ డౌన్‌లోడ్ సమయంలోనూ వర్తింపజేయాలి
    if 'Account_No' in df1.columns:
        df1['Account_No'] = df1['Account_No'].astype(str).str.strip()
    if 'RD1 Acc Num' in df2.columns:
        df2['RD1 Acc Num'] = df2['RD1 Acc Num'].astype(str).str.strip()
        is_numeric_acc = df2['RD1 Acc Num'].str.match(r'^\d+$', na=False)
        df2_valid = df2[is_numeric_acc].copy()
    else:
        df2_valid = df2.copy()

    merged_df = pd.merge(df1, df2_valid, left_on='Account_No', right_on='RD1 Acc Num', how='outer', indicator=True)

    # యూజర్ క్లిక్ చేసిన టైప్ ఆధారంగా ఎక్సెల్ ఫైల్ జనరేట్ చేయడం
    if type == 'matched':
        export_df = merged_df[merged_df['_merge'] == 'both'].drop(columns=['_merge'])
        filename = "Matched_Records.xlsx"
    elif type == 'unmatched':
        export_df = merged_df[merged_df['_merge'] != 'both'].drop(columns=['_merge'])
        filename = "Unmatched_Records.xlsx"
    else:
        export_df = merged_df.drop(columns=['_merge'])
        filename = "Full_Merged_Data.xlsx"

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Data')
    output.seek(0)

    return send_file(
        output,
        download_name=filename,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
