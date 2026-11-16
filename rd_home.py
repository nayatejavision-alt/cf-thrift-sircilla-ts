import io
import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, session, send_file
import pandas as pd
import openpyxl
import pickle  # పైభాగంలో ఇంపోర్ట్ చేసుకోండి
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from src.parser import parse_term_deposit_file


rd_bp = Blueprint('rd_bp', __name__, template_folder='templates')

UPLOAD_FOLDER = 'temp_file.txt'

# పర్ఫెక్ట్ PRN పార్సర్ లాజిక్ (Fixed-width slicing based on header positions)
def parse_custom_prn_file(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
    
    # SNO, Name, Aadhar, Contact, Bank Acc, Bank Name, RD1 Acc, CIF NO
    colspecs = [(0, 4), (4, 47), (47, 61), (61, 75), (75, 90), (90, 124), (124, 147), (147, None)]
    
    data = []
    for line in lines[1:]:
        line_clean = line.rstrip('\r\n')
        if not line_clean.strip():
            continue
        row = [line_clean[s:e].strip() for s, e in colspecs]
        
        # భద్రతా నిబంధనల ప్రకారం ఆధార్ నంబర్లను మాస్క్ చేయడం (Redaction)
        if len(row) > 2 and row[2]:
            row[2] = "[Aadhaar Redacted]"
        data.append(row)
        
    columns = ['SNO', 'Name Of Beneficiary', 'Aadhar Number', 'Contact Number', 'Bank Acc Number', 'Name As Per Bank Records', 'RD1 Acc Num', 'CIF NO']
    return pd.DataFrame(data, columns=columns)


@rd_bp.route('/rd/index')
def index():
    return render_template('rd_view.html')




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
        df2 = parse_custom_prn_file(path2)
              
        merged_df = pd.merge(df1, df2, left_on='Account_No', right_on='RD1 Acc Num', how='outer', indicator=True)
        
        # ❌ పాత పద్ధతి (కుకీ సైజ్ పెరిగిపోతుంది):
        # session['merged_data_cache'] = merged_df.to_json(...)
        
        # ✅ కొత్త పద్ధతి: డేటాను /tmp లో పికిల్ ఫైల్‌గా సేవ్ చేయడం (కుకీ ఓవర్‌లోడ్ ఉండదు)
        cache_path = os.path.join(upload_folder, 'merged_cache.pkl')
        merged_df.to_pickle(cache_path)
        
        return render_template('rd_dashboard.html', 
                               matched_count=len(merged_df[merged_df['_merge'] == 'both']), 
                               unmatched_count=len(merged_df[merged_df['_merge'] != 'both']),
                               tables=[merged_df.head(100).to_html(classes='table table-striped', index=False)])
                           
    except Exception as e:
        flash(f'ఫైల్స్ ప్రాసెస్ చేయడంలో లోపం ఏర్పడింది: {str(e)}')
        return redirect(request.url)


@rd_bp.route('/download/<type>')
def download_excel(type):
    cache_path = os.path.join('/tmp', 'merged_cache.pkl')
    
    if os.path.exists(cache_path):
        merged_df = pd.read_pickle(cache_path)
    else:
        # ఫైల్ మిస్ అయితే మళ్లీ ప్రాసెస్ చేయడం
        path1 = os.path.join('/tmp', session.get('file1_name', ''))
        path2 = os.path.join('/tmp', session.get('file2_name', ''))
        
        if not os.path.exists(path1) or not os.path.exists(path2):
            flash("ఫైల్స్ లభించలేదు, దయచేసి మళ్ళీ అప్‌లోడ్ చేయండి.")
            return redirect(url_for('rd_bp.index'))

        df1 = parse_term_deposit_file(path1)
        df2 = parse_custom_prn_file(path2)
        merged_df = pd.merge(df1, df2, left_on='Account_No', right_on='RD1 Acc Num', how='outer', indicator=True)

    if type == 'matched':
        export_df = merged_df[merged_df['_merge'] == 'both'].drop(columns=['_merge'])
        filename = "Matched_Records.xlsx"
    elif type == 'unmatched':
        export_df = merged_df[merged_df['_merge'] != 'both'].drop(columns=['_merge'])
        filename = "Unmatched_Records.xlsx"
    else:
        export_df = merged_df.drop(columns=['_merge'])
        filename = "Full_Merged_Data.xlsx"

    # మిగతా ఎక్సెల్ జనరేషన్ కోడ్ అలాగే ఉంచండి...
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report Data"

    # గ్రిడ్ లైన్స్ ఎనేబుల్ చేయడం
    ws.views.sheetView[0].showGridLines = True

    # స్టైల్స్ డిఫైన్ చేయడం
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )

    # హెడర్స్ రాసేటప్పుడు _merge మినహాయించి మిగతా కాలమ్స్ తీసుకోవడం
    columns = list(export_df.columns)
    ws.append(columns)
    
    # హెడర్ రో ఫార్మాటింగ్
    for col_num in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[1].height = 25

    # డేటా రోస్ అపెండ్ చేయడం
    for r_idx, row_data in enumerate(export_df.values, start=2):
        ws.append(list(row_data))
        ws.row_dimensions[r_idx].height = 20
        for c_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="left", vertical="center")
            # అకౌంట్ నంబర్లు సైంటిఫిక్ నోటేషన్‌లోకి మారకుండా టెక్స్ట్ ఫార్మాట్
            cell.number_format = '@'

    # కాలమ్ వెడల్పును ఆటో-ఫిట్ చేయడం
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws.freeze_panes = 'A2'

    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name=filename,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
