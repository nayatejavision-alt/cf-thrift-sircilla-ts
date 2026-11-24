import io
from flask import Blueprint, jsonify, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from datetime import datetime

import os
import re

excel_tools_bp = Blueprint('excel_tools_bp', __name__)


@excel_tools_bp.route('/process', methods=['POST'])
def process_join():
    join_type = request.form.get('join_type')

    try:
        if join_type == 'single_file':
            if 'file_single' not in request.files:
                return 'దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.', 400
            file = request.files['file_single']
            sheet1 = request.form.get('sheet1', '').strip()
            col1 = request.form.get('col1', '').strip()
            sheet2 = request.form.get('sheet2', '').strip()
            col2 = request.form.get('col2', '').strip()

            df1 = pd.read_excel(file, sheet_name=sheet1 if sheet1 else 0)
            df2 = pd.read_excel(file, sheet_name=sheet2 if sheet2 else 1)

        elif join_type == 'two_files':
            if 'file1' not in request.files or 'file2' not in request.files:
                return 'దయచేసి రెండు ఫైల్స్ అప్‌లోడ్ చేయండి.', 400
            file1 = request.files['file1']
            sheet1_f1 = request.form.get('sheet1_f1', '').strip()
            col1 = request.form.get('col1_f1', '').strip()

            file2 = request.files['file2']
            sheet2_f2 = request.form.get('sheet2_f2', '').strip()
            col2 = request.form.get('col2_f2', '').strip()

            df1 = pd.read_excel(file1, sheet_name=sheet1_f1 if sheet1_f1 else 0)
            df2 = pd.read_excel(file2, sheet_name=sheet2_f2 if sheet2_f2 else 0)
        else:
            return 'సరైన జాయిన్ ఆప్షన్ ఎంచుకోండి.', 400

        df1.columns = [str(c).strip() for c in df1.columns]
        df2.columns = [str(c).strip() for c in df2.columns]

        if col1 not in df1.columns or col2 not in df2.columns:
            return (
                f"లోపం: '{col1}' లేదా '{col2}' కాలమ్స్ ఎక్సెల్ ఫైల్‌లో కనుగొనబడలేదు!",
                400,
            )

        def clean_acc(val):
            if pd.isna(val):
                return ''
            val_str = str(val).strip()
            if val_str.endswith('.0'):
                val_str = val_str[:-2]
            return val_str.replace("'", '').strip()

        df1['Clean_Key'] = df1[col1].apply(clean_acc)
        df2['Clean_Key'] = df2[col2].apply(clean_acc)

        df1_subset = df1.drop_duplicates(subset=['Clean_Key']).copy()
        df2_subset = df2.drop_duplicates(subset=['Clean_Key']).copy()

        merged_df = pd.merge(
            df1_subset,
            df2_subset,
            on='Clean_Key',
            how='left',
            suffixes=('_left', '_right'),
        )
        merged_df = merged_df.drop(columns=['Clean_Key'])

        output_file = 'Merged_RD_Accounts_Output.xlsx'
        merged_df.to_excel(output_file, index=False)

        return send_file(output_file, as_attachment=True, download_name=output_file)

    except Exception as e:
        return f'ఫైల్ ప్రాసెసింగ్‌లో లోపం సంభవించింది: {str(e)}', 500


@excel_tools_bp.route('/process-reconcile', methods=['POST'])
def process_reconcile():
    try:
        file = request.files['reconcile_file']
        sheet_name = request.form.get('rec_sheet_name', '').strip()
        rd_col = request.form.get('rec_rd_col', 'Rd').strip()
        ben_col = request.form.get('rec_ben_col', 'Ben').strip()
        amt_col = request.form.get('rec_amt_col', 'Amount').strip()

        df = pd.read_excel(file, sheet_name=sheet_name if sheet_name else 0)
        df.columns = [str(c).strip() for c in df.columns]

        df_clean = df[[rd_col, ben_col, amt_col]].dropna(
            subset=[rd_col, ben_col]
        ).copy()
        df_clean['Ben_clean'] = df_clean[ben_col].astype(str).str.strip()

        def format_rd(val):
            try:
                return str(int(float(val)))
            except Exception:
                return str(val).strip()

        df_clean['Rd_formatted'] = df_clean[rd_col].apply(format_rd)

        unique_bens = []
        seen = set()
        for ben in df_clean['Ben_clean']:
            if ben not in seen:
                seen.add(ben)
                unique_bens.append(ben)

        max_rds = df_clean.groupby('Ben_clean')['Rd_formatted'].count().max()

        rows_data = []
        for idx, ben in enumerate(unique_bens, 1):
            ben_rows = df_clean[df_clean['Ben_clean'] == ben]
            row_dict = {'S.No': idx, 'Name Of Beneficiary': ben}

            rds = ben_rows['Rd_formatted'].tolist()
            amounts = ben_rows[amt_col].tolist()

            for i in range(len(rds)):
                acc_num = i + 1
                row_dict[f'RD Account {acc_num}'] = str(rds[i])
                row_dict[f'Closing Balance {acc_num}'] = (
                    amounts[i] if pd.notnull(amounts[i]) else 0.0
                )

            rows_data.append(row_dict)

        df_result = pd.DataFrame(rows_data)

        all_cols = ['S.No', 'Name Of Beneficiary']
        for i in range(1, max_rds + 1):
            all_cols.append(f'RD Account {i}')
            all_cols.append(f'Closing Balance {i}')

        all_cols.append('Total RDs')
        all_cols.append('Total Closing Balance (₹)')

        df_result = df_result.reindex(columns=all_cols)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Reconciled RD Accounts'
        ws.views.sheetView[0].showGridLines = True

        HEADER_FILL = PatternFill(
            start_color='1B365D', end_color='1B365D', fill_type='solid'
        )
        SUBHEADER_FILL = PatternFill(
            start_color='2C4D75', end_color='2C4D75', fill_type='solid'
        )
        HIGHLIGHT_FILL = PatternFill(
            start_color='134E5E', end_color='134E5E', fill_type='solid'
        )
        ZEBRA_FILL = PatternFill(
            start_color='F4F7FA', end_color='F4F7FA', fill_type='solid'
        )

        max_col_letter = get_column_letter(len(all_cols))
        ws.merge_cells(f'A1:{max_col_letter}1')
        title_cell = ws['A1']
        title_cell.value = (
            'DEPARTMENT OF HANDLOOMS & TEXTILE - THRIFT SCHEME RD ACCOUNTS STATEMENT'
        )
        title_cell.font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        title_cell.fill = HEADER_FILL
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 35

        ws.row_dimensions[2].height = 28
        for col_idx, h_text in enumerate(all_cols, 1):
            cell = ws.cell(row=2, column=col_idx, value=h_text)
            cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
            cell.fill = (
                HIGHLIGHT_FILL
                if col_idx >= len(all_cols) - 1
                else SUBHEADER_FILL
            )
            cell.alignment = Alignment(
                horizontal='center', vertical='center', wrap_text=True
            )

        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9'),
        )

        for r_idx, row in df_result.iterrows():
            row_num = r_idx + 3
            ws.row_dimensions[row_num].height = 20

            ws.cell(row=row_num, column=1, value=r_idx + 1).alignment = Alignment(
                horizontal='center', vertical='center'
            )
            ws.cell(
                row=row_num, column=2, value=row['Name Of Beneficiary']
            ).alignment = Alignment(horizontal='left', vertical='center')

            for i in range(1, max_rds + 1):
                rd_val = row.get(f'RD Account {i}', '')
                bal_val = row.get(f'Closing Balance {i}', 0.0)

                c_rd = ws.cell(
                    row=row_num,
                    column=2 * i + 1,
                    value=(
                        str(rd_val)
                        if pd.notnull(rd_val) and str(rd_val) != ''
                        else '-'
                    ),
                )
                c_rd.alignment = Alignment(horizontal='center', vertical='center')
                c_rd.number_format = '@'

                c_bal = ws.cell(
                    row=row_num,
                    column=2 * i + 2,
                    value=bal_val if pd.notnull(bal_val) else 0.0,
                )
                c_bal.alignment = Alignment(horizontal='right', vertical='center')
                c_bal.number_format = '₹#,##0.00'

            tot_rd_col = len(all_cols) - 1
            tot_bal_col = len(all_cols)

            rd_cols_letters = [
                get_column_letter(2 * i + 1) for i in range(1, max_rds + 1)
            ]
            bal_cols_letters = [
                get_column_letter(2 * i + 2) for i in range(1, max_rds + 1)
            ]

            rd_range_str = ','.join(
                [f'{col}{row_num}' for col in rd_cols_letters]
            )
            rd_range_range = (
                f'{rd_cols_letters[0]}{row_num}:{rd_cols_letters[-1]}{row_num}'
            )

            c_tot_rd = ws.cell(row=row_num, column=tot_rd_col)
            c_tot_rd.value = (
                f'=COUNTA({rd_range_str})-COUNTIF({rd_range_range}, "-")'
            )
            c_tot_rd.alignment = Alignment(horizontal='center', vertical='center')
            c_tot_rd.font = Font(name='Calibri', size=10, bold=True)

            c_tot_bal = ws.cell(row=row_num, column=tot_bal_col)
            c_tot_bal.value = '=' + '+'.join(
                [f'{col}{row_num}' for col in bal_cols_letters]
            )
            c_tot_bal.number_format = '₹#,##0.00'
            c_tot_bal.alignment = Alignment(horizontal='right', vertical='center')
            c_tot_bal.font = Font(name='Calibri', size=10, bold=True)

            for col_c in range(1, len(all_cols) + 1):
                c_item = ws.cell(row=row_num, column=col_c)
                c_item.border = thin_border
                if r_idx % 2 == 1:
                    c_item.fill = ZEBRA_FILL

        total_row = len(df_result) + 3
        ws.merge_cells(
            start_row=total_row,
            start_column=1,
            end_row=total_row,
            end_column=2,
        )
        ws.cell(
            total_row, 1, value='GRAND TOTAL / మొత్తం వివరాలు'
        ).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(total_row, 1).font = Font(name='Calibri', size=10, bold=True)

        for c in range(3, len(all_cols) + 1):
            col_ltr = get_column_letter(c)
            cell_tot = ws.cell(total_row, c)
            if c == tot_rd_col:
                cell_tot.value = f'=SUM({col_ltr}3:{col_ltr}{total_row-1})'
                cell_tot.font = Font(name='Calibri', size=10, bold=True)
            elif c == tot_bal_col:
                cell_tot.value = f'=SUM({col_ltr}3:{col_ltr}{total_row-1})'
                cell_tot.number_format = '₹#,##0.00'
                cell_tot.font = Font(name='Calibri', size=10, bold=True)
            elif c % 2 == 0:
                cell_tot.value = f'=SUM({col_ltr}3:{col_ltr}{total_row-1})'
                cell_tot.number_format = '₹#,##0.00'
            else:
                cell_tot.value = '-'

        ws.freeze_panes = 'C3'

        output_filename = 'RD_Accounts_Reconciled_Enhanced.xlsx'
        wb.save(output_filename)
        return send_file(
            output_filename, as_attachment=True, download_name=output_filename
        )

    except Exception as e:
        return f'Error processing reconciliation file: {str(e)}', 400



def _clean_account_value(val):
    """అకౌంట్ నంబర్లలోని .0, అపోస్ట్రోఫీలు, స్పేస్‌లు, NaN లను సురక్షితంగా క్లీన్ చేస్తుంది."""
    if pd.isna(val):
        return ''
    s = str(val).strip()
    if s.lower() == 'nan':
        return ''
    if s.endswith('.0'):
        s = s[:-2]
    return s.replace("'", '').strip()


def _sanitize_filename_part(name: str) -> str:
    """ఫైల్ పేరులో అనవసర క్యారెక్టర్లు లేకుండా క్లీన్ చేస్తుంది."""
    cleaned = re.sub(r'[^\w\-_.]', '_', name.strip())
    return re.sub(r'_+', '_', cleaned).strip('_')


def _read_sheet_as_str(file_storage, sheet_name: str, fallback_idx: int = 0) -> pd.DataFrame:
    """ఫైల్ నుండి నిర్దిష్ట షీట్‌ను (లేదా fallback index) స్ట్రింగ్‌లుగా చదివి కాలమ్ పేర్లను క్లీన్ చేస్తుంది."""
    file_storage.stream.seek(0)
    excel_file = pd.ExcelFile(file_storage)
    sheet_names = excel_file.sheet_names

    if sheet_name and sheet_name in sheet_names:
        target_sheet = sheet_name
    else:
        safe_idx = fallback_idx if fallback_idx < len(sheet_names) else 0
        target_sheet = sheet_names[safe_idx]

    df = pd.read_excel(excel_file, sheet_name=target_sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


@excel_tools_bp.route('/process-master-match', methods=['POST'])
def process_master_match():
    """ఒకే ఫైల్ లేదా రెండు వేర్వేరు ఫైల్స్ నుండి Master (acc 1, acc 2) vs Data (acc) మ్యాచ్ చేసి Matched & Unmatched రిపోర్ట్ ఇస్తుంది."""
    try:
        mm_mode = request.form.get('mm_mode', 'single_file').strip()
        master_col1 = request.form.get('mm_master_col1', '').strip()
        master_col2 = request.form.get('mm_master_col2', '').strip()
        data_col = request.form.get('mm_data_col', '').strip()
        case_insensitive = request.form.get('mm_case_insensitive') == 'on'
        custom_prefix = request.form.get('mm_output_prefix', '').strip()

        if not master_col1 or not data_col:
            return 'దయచేసి Master Column 1 మరియు Data Match Column ఎంచుకోండి.', 400

        if mm_mode == 'single_file':
            if 'mm_file_single' not in request.files:
                return 'దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.', 400
            file_single = request.files['mm_file_single']
            if not file_single or file_single.filename == '':
                return 'ఫైల్ ఎంచుకోలేదు.', 400

            master_sheet = request.form.get('mm_single_master_sheet', '').strip()
            data_sheet = request.form.get('mm_single_data_sheet', '').strip()

            master_df = _read_sheet_as_str(file_single, master_sheet, fallback_idx=0)
            data_df = _read_sheet_as_str(file_single, data_sheet, fallback_idx=1)
            source_label = os.path.splitext(file_single.filename)[0]

        elif mm_mode == 'two_files':
            if 'mm_file_master' not in request.files or 'mm_file_data' not in request.files:
                return 'దయచేసి Master ఫైల్ మరియు Data ఫైల్ — రెండూ అప్‌లోడ్ చేయండి.', 400
            file_master = request.files['mm_file_master']
            file_data = request.files['mm_file_data']

            if not file_master or file_master.filename == '' or not file_data or file_data.filename == '':
                return 'రెండు ఎక్సెల్ ఫైల్స్ సరిగ్గా ఎంచుకోండి.', 400

            master_sheet = request.form.get('mm_two_master_sheet', '').strip()
            data_sheet = request.form.get('mm_two_data_sheet', '').strip()

            master_df = _read_sheet_as_str(file_master, master_sheet, fallback_idx=0)
            data_df = _read_sheet_as_str(file_data, data_sheet, fallback_idx=0)

            m_base = os.path.splitext(file_master.filename)[0]
            d_base = os.path.splitext(file_data.filename)[0]
            source_label = f'{d_base}_vs_{m_base}'

        else:
            return 'సరైన మోడ్ (Single File / Two Files) ఎంచుకోండి.', 400

        if master_col1 not in master_df.columns:
            return f"Master షీట్‌లో '{master_col1}' కాలమ్ కనుగొనబడలేదు.", 400
        if master_col2 and master_col2 not in master_df.columns:
            return f"Master షీట్‌లో '{master_col2}' కాలమ్ కనుగొనబడలేదు.", 400
        if data_col not in data_df.columns:
            return f"Data షీట్‌లో '{data_col}' కాలమ్ కనుగొనబడలేదు.", 400

        def normalize_key(val):
            cleaned = _clean_account_value(val)
            return cleaned.upper() if case_insensitive else cleaned

        acc1_series = master_df[master_col1].dropna().map(normalize_key)
        master_accounts = set(x for x in acc1_series if x != '')

        if master_col2:
            acc2_series = master_df[master_col2].dropna().map(normalize_key)
            master_accounts.update(x for x in acc2_series if x != '')

        clean_keys = data_df[data_col].map(normalize_key)
        is_matched = clean_keys.isin(master_accounts) & (clean_keys != '')

        matched_df = data_df[is_matched].copy()
        unmatched_df = data_df[~is_matched].copy()

        # డైనమిక్ ఔట్‌పుట్ ఫైల్ నేమ్
        safe_base = _sanitize_filename_part(
            custom_prefix if custom_prefix else f'Master_Match_{source_label}'
        )
        if not safe_base:
            safe_base = 'Master_Match_Report'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        dynamic_filename = f'{safe_base}_{timestamp}.xlsx'

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            matched_df.to_excel(writer, sheet_name='Matched', index=False)
            unmatched_df.to_excel(writer, sheet_name='Unmatched', index=False)
        output_buffer.seek(0)

        return send_file(
            output_buffer,
            as_attachment=True,
            download_name=dynamic_filename,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'Master Match ప్రాసెసింగ్‌లో లోపం సంభవించింది: {str(e)}', 500

@excel_tools_bp.route('/api/parse-excel-headers', methods=['POST'])
def parse_excel_headers():
    """అప్‌లోడ్ చేసిన ఎక్సెల్ ఫైల్ నుండి షీట్ పేర్లు మరియు కాలమ్స్ ఎక్స్‌ట్రాక్ట్ చేసే API"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'ఫైల్ లభించలేదు.'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'ఫైల్ ఎంచుకోలేదు.'}), 400

        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names

        df_first = pd.read_excel(excel_file, sheet_name=sheet_names[0], nrows=1)
        columns = [str(col).strip() for col in df_first.columns]

        return jsonify({
            'status': 'success',
            'sheets': sheet_names,
            'columns': columns,
        })
    except Exception as e:
        return jsonify(
            {'status': 'error', 'message': f'ఫైల్ రీడ్ చేయడంలో లోపం: {str(e)}'}
        ), 500


@excel_tools_bp.route('/api/get-sheet-columns', methods=['POST'])
def get_sheet_columns():
    """ఎంచుకున్న నిర్దిష్ట షీట్ లోని కాలమ్స్ ఎక్స్‌ట్రాక్ట్ చేసే API"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'ఫైల్ లభించలేదు.'}), 400

        file = request.files['file']
        sheet_name = request.form.get('sheet_name', '').strip()

        excel_file = pd.ExcelFile(file)
        if sheet_name not in excel_file.sheet_names:
            sheet_name = excel_file.sheet_names[0]

        df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=1)
        columns = [str(col).strip() for col in df.columns]

        return jsonify({'status': 'success', 'columns': columns})
    except Exception as e:
        return jsonify(
            {'status': 'error', 'message': f'కాలమ్స్ లోడ్ చేయడంలో లోపం: {str(e)}'}
        ), 500
