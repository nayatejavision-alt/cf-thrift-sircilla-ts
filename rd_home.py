import os
from flask import Flask, render_template, request, redirect, url_for, flash
from flask import Blueprint, render_template, request, jsonify, session, redirect, send_file
from src.parser import parse_term_deposit_file, parse_prn_file
import pandas as pd

rd_bp = Blueprint('rd_bp', __name__, template_folder='templates')

UPLOAD_FOLDER = 'temp_file.txt'

@rd_bp.route('/rd/index')
def index():
    return render_template('rd_view.html')

@rd_bp.route('/rd/upload', methods=['GET', 'POST'])
def upload_file():
    # ఒకవేళ గెట్ రిక్వెస్ట్ వస్తే ఇండెక్స్ పేజీకి రీడైరెక్ట్ చేయవచ్చు
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
    # 1. User upload chesina files unnaayo ledo check cheyadam
    if 'file1' not in request.files or 'file2' not in request.files:
        flash('Dayachesi rendu files select cheyandi!')
        return redirect(request.url)
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    if file1.filename == '' or file2.filename == '':
        flash('Asalu file select avvaledu!')
        return redirect(request.url)
    
    # 2. Upload aina files ni temporary folder lo save cheyadam
    upload_folder = '/tmp'  # Mee project lo unna uploads folder path kuda vadukovachu
    os.makedirs(upload_folder, exist_ok=True)
    
    path1 = os.path.join(upload_folder, file1.filename)
    path2 = os.path.join(upload_folder, file2.filename)
    
    file1.save(path1)
    file2.save(path2)
    
    try:
        # 3. Parser functions dwara rendu files ni DataFrames ga convert cheyadam
        # Modati file kosam mee daggara unna function vadandi (eg: parse_term_deposit_file)
        df1 = parse_term_deposit_file(path1)
        
        # Rendava file (.prn) kosam mana kotha parser function
        df2 = parse_prn_file(path2)
        
        # 4. Account numbers adharanga rendu DataFrames ni merge cheyadam
        # Note: df1 lo account column peru 'Account_No' anukundam, df2 lo 'RD1 Acc Num' ani untundi.
        # Oka vela column names veru ga unte ikkada mari chusukovalsi untundi.
        merged_df = pd.merge(df1, df2, left_on='Account_No', right_on='RD1 Acc Num', how='outer', indicator=True)
        
        # 5. Match aina (both) mariyu match avvani data ni separate cheyadam
        matched_data = merged_df[merged_df['_merge'] == 'both']
        unmatched_data = merged_df[merged_df['_merge'] != 'both']
        
        # 6. Results ni render_template dwara dashboard ki pampadam
        return render_template('rd_dashboard.html', 
                               matched_count=len(matched_data), 
                               unmatched_count=len(unmatched_data),
                               tables=[matched_data.to_html(classes='data', header="true")])
                           
    except Exception as e:
        flash(f'Files process chesthunappudu error vachindi: {str(e)}')
        return redirect(request.url)
