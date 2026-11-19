from flask import Blueprint, jsonify, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

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
