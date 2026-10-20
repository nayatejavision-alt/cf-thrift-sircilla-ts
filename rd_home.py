import os
from flask import Flask, render_template, request, redirect, url_for
from flask import Blueprint, render_template, request, jsonify, session, redirect, send_file
from src.parser import parse_term_deposit_file

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
