from flask import Blueprint, render_template, request, session, redirect, url_for, jsonify, make_response, send_file
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


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
        rd_col = 'RD 1' if 'RD 1' in df_master.columns else ('RD 1' if 'RD 1' in df_master.columns else None)
        batch_col_m = 'BATCH_NUMBER' if 'BATCH_NUMBER' in df_master.columns else ('BATCH NO' if 'BATCH NO' in df_master.columns else 'Batch')
        
        df_master['Bank_Name'] = df_master[bank_col].astype(str).str.strip().replace('', 'Other')
        df_master['RD_Account'] = df_master[rd_col].astype(str).str.strip() if rd_col else ''
        # df_master['Batch_No'] = df_master[batch_col_m].astype(str).str.strip().replace('', '1')

        # 2. ట్రాన్సాక్షన్ డేటాను ప్రాసెస్ చేయడం
        t_rd_col = 'RD ACCOUNT_1 (Beneficiary Contribution)' if 'RD ACCOUNT_1 (Beneficiary Contribution)' in df_trans.columns else ('RD ACCOUNT_1 (Beneficiary Contribution)' if 'RD ACCOUNT_1 (Beneficiary Contribution)' in df_trans.columns else None)
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
            df_merged = pd.merge(df_master, completed_trans, on=['RD_Account'], how='left')
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




@admin_bp.route('/admin/export_excel')
def export_admin_excel():
    if not session.get('admin_logged_in'):
        return redirect(url_for('admin_bp.admin_login'))

    try:
        from app import get_sheets_client, SPREADSHEET_NAME
        client = get_sheets_client()
        spreadsheet = client.open(SPREADSHEET_NAME)
        
        master_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
        trans_sheet = spreadsheet.worksheet("Transactions_Year1")
        
        df_master = pd.DataFrame(master_sheet.get_all_records())
        df_trans = pd.DataFrame(trans_sheet.get_all_records())
        
        if df_master.empty:
            return "డేటా అందుబాటులో లేదు."

        bank_col = 'FILE NAME' if 'FILE NAME' in df_master.columns else ('File Name' if 'File Name' in df_master.columns else df_master.columns[0])
        rd_col = 'RD 1' if 'RD 1' in df_master.columns else ('RD 1' if 'RD 1' in df_master.columns else None)
        
        df_master['Bank_Name'] = df_master[bank_col].astype(str).str.strip().replace('', 'Other')
        df_master['RD_Account'] = df_master[rd_col].astype(str).str.strip() if rd_col else ''
        
        t_rd_col = 'RD ACCOUNT_1 (Beneficiary Contribution)' if 'RD ACCOUNT_1 (Beneficiary Contribution)' in df_trans.columns else None
        t_batch_col = 'BATCH_NUMBER' if 'BATCH_NUMBER' in df_trans.columns else ('BATCH NO' if 'BATCH NO' in df_trans.columns else None)
        
        if not df_trans.empty and t_rd_col:
            df_trans['RD_Account'] = df_trans[t_rd_col].astype(str).str.strip()
            df_trans['Govt_Share'] = pd.to_numeric(df_trans.get('Govt_Share', 0), errors='coerce').fillna(0)
            
            if t_batch_col:
                completed_trans = df_trans[df_trans['Govt_Share'] > 0][['RD_Account', t_batch_col]].drop_duplicates()
                completed_trans.rename(columns={t_batch_col: 'Batch_No'}, inplace=True)
            else:
                completed_trans = df_trans[df_trans['Govt_Share'] > 0][['RD_Account']].drop_duplicates()
                completed_trans['Batch_No'] = '1'

            completed_trans['Is_Completed'] = True
            df_merged = pd.merge(df_master, completed_trans, on=['RD_Account'], how='left')
        else:
            df_merged = df_master.copy()
            df_merged['Is_Completed'] = False
            df_merged['Batch_No'] = '1'

        df_merged['Is_Completed'] = df_merged['Is_Completed'].fillna(False)
        df_merged['Batch_No'] = df_merged['Batch_No'].fillna('1').astype(str)
        df_merged['Status_Val'] = df_merged['Is_Completed'].astype(int)

        # 1. Batch Aggregates
        batch_df = df_merged.groupby('Batch_No').agg(
            total=('RD_Account', 'count'),
            completed=('Status_Val', 'sum')
        ).reset_index()
        batch_df['pending'] = batch_df['total'] - batch_df['completed']

        # 2. Bank Aggregates
        bank_df = df_merged.groupby('Bank_Name').agg(
            total=('RD_Account', 'count'),
            completed=('Status_Val', 'sum')
        ).reset_index()
        bank_df['pending'] = bank_df['total'] - bank_df['completed']

        # 3. Batch + Bank Detailed Matrix
        batch_bank_df = df_merged.groupby(['Batch_No', 'Bank_Name']).agg(
            total=('RD_Account', 'count'),
            completed=('Status_Val', 'sum')
        ).reset_index()
        batch_bank_df['pending'] = batch_bank_df['total'] - batch_bank_df['completed']

        # Openpyxl Excel Workbook creation with Styling
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Thrift Summary"
        ws.views.sheetView[0].showGridLines = True

        # Styles
        f_title = Font(name="Calibri", size=15, bold=True, color="1F4E79")
        f_sec = Font(name="Calibri", size=12, bold=True, color="1F4E79")
        f_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        f_bold = Font(name="Calibri", size=11, bold=True)
        fill_navy = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                             top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

        # Heading
        ws['A1'] = "TS Powerloom Weavers Thrift Fund Scheme - Summary Report"
        ws['A1'].font = f_title

        r = 3
        # --- Table 1: Batch-wise Summary ---
        ws.cell(row=r, column=1, value="1. Batch-wise Summary").font = f_sec
        r += 1
        b_headers = ["Batch No", "Total", "Completed", "Pending", "Completion Rate"]
        for col_idx, h in enumerate(b_headers, 1):
            cell = ws.cell(row=r, column=col_idx, value=h)
            cell.font, cell.fill = f_hdr, fill_navy
            cell.alignment = Alignment(horizontal="center")
        
        r += 1
        start_b_row = r
        for _, row in batch_df.iterrows():
            ws.cell(row=r, column=1, value=str(row['Batch_No']))
            ws.cell(row=r, column=2, value=int(row['total'])).number_format = '#,##0'
            ws.cell(row=r, column=3, value=int(row['completed'])).number_format = '#,##0'
            ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = '#,##0'
            ws.cell(row=r, column=5, value=f"=C{r}/B{r}").number_format = '0.0%'
            for c in range(1, 6): ws.cell(row=r, column=c).border = border_thin
            r += 1

        # Total Row Batch
        ws.cell(row=r, column=1, value="Total").font = f_bold
        ws.cell(row=r, column=2, value=f"=SUM(B{start_b_row}:B{r-1})").font = f_bold
        ws.cell(row=r, column=3, value=f"=SUM(C{start_b_row}:C{r-1})").font = f_bold
        ws.cell(row=r, column=4, value=f"=SUM(D{start_b_row}:D{r-1})").font = f_bold
        ws.cell(row=r, column=5, value=f"=C{r}/B{r}").font = f_bold
        for c in range(1, 6): ws.cell(row=r, column=c).fill = fill_sub
        
        r += 3

        # --- Table 2: Bank-wise Summary ---
        ws.cell(row=r, column=1, value="2. Bank-wise Summary").font = f_sec
        r += 1
        bank_headers = ["Bank Name", "Total", "Completed", "Pending", "Completion Rate"]
        for col_idx, h in enumerate(bank_headers, 1):
            cell = ws.cell(row=r, column=col_idx, value=h)
            cell.font, cell.fill = f_hdr, fill_navy
            cell.alignment = Alignment(horizontal="center")
        
        r += 1
        start_bank_row = r
        for _, row in bank_df.iterrows():
            ws.cell(row=r, column=1, value=str(row['Bank_Name']))
            ws.cell(row=r, column=2, value=int(row['total'])).number_format = '#,##0'
            ws.cell(row=r, column=3, value=int(row['completed'])).number_format = '#,##0'
            ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = '#,##0'
            ws.cell(row=r, column=5, value=f"=C{r}/B{r}").number_format = '0.0%'
            for c in range(1, 6): ws.cell(row=r, column=c).border = border_thin
            r += 1

        # Total Row Bank
        ws.cell(row=r, column=1, value="Total").font = f_bold
        ws.cell(row=r, column=2, value=f"=SUM(B{start_bank_row}:B{r-1})").font = f_bold
        ws.cell(row=r, column=3, value=f"=SUM(C{start_bank_row}:C{r-1})").font = f_bold
        ws.cell(row=r, column=4, value=f"=SUM(D{start_bank_row}:D{r-1})").font = f_bold
        ws.cell(row=r, column=5, value=f"=C{r}/B{r}").font = f_bold
        for c in range(1, 6): ws.cell(row=r, column=c).fill = fill_sub

        r += 3

        # --- Table 3: Batch + Bank Detailed Matrix ---
        ws.cell(row=r, column=1, value="3. Batch & Bank Detailed Breakdown").font = f_sec
        r += 1
        matrix_headers = ["Batch No", "Bank Name", "Total", "Completed", "Pending"]
        for col_idx, h in enumerate(matrix_headers, 1):
            cell = ws.cell(row=r, column=col_idx, value=h)
            cell.font, cell.fill = f_hdr, fill_navy
            cell.alignment = Alignment(horizontal="center")
        
        r += 1
        for _, row in batch_bank_df.iterrows():
            ws.cell(row=r, column=1, value=str(row['Batch_No']))
            ws.cell(row=r, column=2, value=str(row['Bank_Name']))
            ws.cell(row=r, column=3, value=int(row['total'])).number_format = '#,##0'
            ws.cell(row=r, column=4, value=int(row['completed'])).number_format = '#,##0'
            ws.cell(row=r, column=5, value=f"=C{r}-D{r}").number_format = '#,##0'
            for c in range(1, 6): ws.cell(row=r, column=c).border = border_thin
            r += 1

        # Auto column width adjustment
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

        # Save to buffer and download response
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(output, download_name="Thrift_Admin_Summary.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return f"Excel ఎగుమతిలో లోపం సంభవించింది: {str(e)}"
