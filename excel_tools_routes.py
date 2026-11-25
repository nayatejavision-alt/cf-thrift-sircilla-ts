from datetime import datetime
import io
import os
import re
from flask import Blueprint, jsonify, request, send_file
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

excel_tools_bp = Blueprint('excel_tools_bp', __name__)


def _clean_account_value(val):
    """అకౌంట్ నంబర్లలోని .0, అపోస్ట్రోఫీలు, స్పేస్‌లు, NaN లను సురక్షితంగా క్లీన్ చేస్తుంది."""
    if pd.isna(val):
        return ''
    s = str(val).strip()
    if s.lower() in ('nan', 'none', 'null', '-'):
        return ''
    if s.endswith('.0'):
        s = s[:-2]
    return s.replace("'", '').strip()


def _sanitize_filename_part(name: str) -> str:
    """ఫైల్ పేరులో అనవసర క్యారెక్టర్లు లేకుండా క్లీన్ చేస్తుంది."""
    cleaned = re.sub(r'[^\w\-_.]', '_', (name or '').strip())
    return re.sub(r'_+', '_', cleaned).strip('_')


def _read_sheet_as_str(
    file_storage, sheet_name: str, fallback_idx: int = 0
) -> pd.DataFrame:
    """నిర్దిష్ట షీట్‌ను స్ట్రింగ్‌లుగా చదివి కాలమ్ పేర్లను క్లీన్ చేస్తుంది."""
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


def _style_openpyxl_sheet(
    ws, title_text: str, text_cols=None, currency_cols=None
):
    """ఎక్సెల్ షీట్‌కు ప్రొఫెషనల్ హెడర్, బోర్డర్స్, ఫార్మాట్స్ మరియు కాలమ్ విడ్త్ సెట్ చేస్తుంది."""
    text_cols = set(text_cols or [])
    currency_cols = set(currency_cols or [])

    ws.views.sheetView[0].showGridLines = True
    max_col = max(ws.max_column, 1)
    max_row = ws.max_row
    max_col_letter = get_column_letter(max_col)

    HEADER_FILL = PatternFill(
        start_color='1B365D', end_color='1B365D', fill_type='solid'
    )
    SUBHEADER_FILL = PatternFill(
        start_color='2C4D75', end_color='2C4D75', fill_type='solid'
    )
    ZEBRA_FILL = PatternFill(
        start_color='F8FAFC', end_color='F8FAFC', fill_type='solid'
    )

    thin_side = Side(style='thin', color='CBD5E1')
    cell_border = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    # Insert Title Row
    ws.insert_rows(1)
    ws.merge_cells(f'A1:{max_col_letter}1')
    title_cell = ws['A1']
    title_cell.value = title_text
    title_cell.font = Font(name='Calibri', size=13, bold=True, color='FFFFFF')
    title_cell.fill = HEADER_FILL
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 32

    # Header Row (Row 2)
    ws.row_dimensions[2].height = 26
    headers_map = {}
    for col_idx in range(1, max_col + 1):
        h_cell = ws.cell(row=2, column=col_idx)
        headers_map[col_idx] = str(h_cell.value or '').strip()
        h_cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        h_cell.fill = SUBHEADER_FILL
        h_cell.alignment = Alignment(
            horizontal='center', vertical='center', wrap_text=True
        )
        h_cell.border = cell_border

    # Data Rows (Row 3+)
    for r_idx in range(3, max_row + 2):
        ws.row_dimensions[r_idx].height = 20
        is_even = r_idx % 2 == 0
        for c_idx in range(1, max_col + 1):
            c = ws.cell(row=r_idx, column=c_idx)
            h_name = headers_map.get(c_idx, '')
            c.border = cell_border
            c.font = Font(name='Calibri', size=10, color='0F172A')
            if is_even:
                c.fill = ZEBRA_FILL

            if h_name in text_cols:
                c.value = (
                    ''
                    if c.value is None or str(c.value).strip().lower() == 'nan'
                    else str(c.value).strip()
                )
                c.number_format = '@'
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif h_name in currency_cols:
                try:
                    c.value = float(c.value or 0.0)
                except Exception:
                    c.value = 0.0
                c.number_format = '₹#,##0.00'
                c.alignment = Alignment(horizontal='right', vertical='center')
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')

    # Auto-fit column widths safely
    for c_idx in range(1, max_col + 1):
        col_letter = get_column_letter(c_idx)
        max_len = 12
        for r_idx in range(2, max_row + 2):
            val = ws.cell(row=r_idx, column=c_idx).value
            if val is not None:
                val_len = len(str(val))
                if val_len > max_len:
                    max_len = val_len
        ws.column_dimensions[col_letter].width = min(max_len + 4, 44)

    ws.freeze_panes = 'A3'


# ==============================================================================
# STANDALONE TEXT (.TXT) & PRN (.PRN) PARSER CONVERTERS TO EXCEL
# ==============================================================================
def _parse_term_deposit_text_bytes(raw_bytes: bytes) -> pd.DataFrame:
    """Term Deposit / RD / FD టెక్స్ట్ బైట్స్‌ను పార్స్ చేస్తుంది; regex మ్యాచ్ కాకపోతే generic text parser వాడుతుంది."""
    text = raw_bytes.decode('cp1252', errors='ignore')
    lines = text.splitlines()

    current_sec = 'Unknown'
    clean_rows = []

    for line in lines:
        if (
            'Fixed deposits -' in line
            or 'Short term deposits -' in line
            or 'Recurring Deposits' in line
        ):
            current_sec = line.strip()

        match = re.match(
            r'^\s*(\d{1,4})\s+(\d+)\s+(\d+)\s+(.*?)\s+(\d+\.\d+)\s+([A-Z]+)\s+([\d,]+\.\d+)\s+(\d+)\s+(\d{2}/\d{2}/\d{4})\s+([\d,]+\.\d+)',
            line,
        )
        if match:
            (
                sno,
                cust_no,
                acc_no,
                acc_name,
                int_rate,
                status,
                balance,
                tenor,
                mat_date,
                avail_int,
            ) = match.groups()
            clean_rows.append({
                'Section': current_sec,
                'SNO': int(sno),
                'Customer_No': _clean_account_value(cust_no),
                'Account_No': _clean_account_value(acc_no),
                'Account_Name': acc_name.strip(),
                'Int_Rate': float(int_rate),
                'Status': status,
                'Balance': float(balance.replace(',', '')),
                'Tenor': tenor,
                'Mat_Date': mat_date,
                'Avail_Int': float(avail_int.replace(',', '')),
            })

    if clean_rows:
        return pd.DataFrame(clean_rows)

    # Generic fallback for other .txt files (pipe/tab/comma/multi-space separated)
    non_empty = [ln.strip() for ln in lines if ln.strip()]
    if not non_empty:
        return pd.DataFrame(columns=['Raw_Line'])

    sep = None
    for candidate in ('|', '\t', ','):
        if candidate in non_empty[0]:
            sep = candidate
            break

    if sep:
        parsed = [
            [cell.strip() for cell in ln.split(sep)] for ln in non_empty
        ]
    else:
        parsed = [re.split(r'\s{2,}', ln) for ln in non_empty]

    max_cols = max(len(r) for r in parsed)
    headers = [f'Column_{i+1}' for i in range(max_cols)]
    padded = [r + [''] * (max_cols - len(r)) for r in parsed]
    return pd.DataFrame(padded, columns=headers)


def _parse_prn_bytes(raw_bytes: bytes, mode: str = 'auto') -> pd.DataFrame:
    """PRN ఫైల్‌ను Fixed-width (colspecs) మరియు Regex fallback తో సురక్షితంగా పార్స్ చేస్తుంది."""
    text = raw_bytes.decode('utf-8', errors='replace')
    lines = text.splitlines()

    # 1. Fixed-width parsing (SNO, Name, Aadhar, Contact, Bank Acc, Bank Name, RD1 Acc, CIF NO)
    colspecs = [
        (0, 4),
        (4, 47),
        (47, 61),
        (61, 75),
        (75, 90),
        (90, 124),
        (124, 147),
        (147, None),
    ]
    fw_rows = []
    for line in lines[1:]:
        line_clean = line.rstrip('\r\n')
        if not line_clean.strip():
            continue
        row = [line_clean[s:e].strip() for s, e in colspecs]
        # ఆధార్ కాలమ్ (index 2) లో 12 అంకెల సంఖ్య ఉంటే రెడ్ యాక్ట్ చేయడం
        if len(row) > 2 and row[2]:
            digits_only = re.sub(r'\D', '', row[2])
            if len(digits_only) == 12:
                row[2] = '[Aadhaar Redacted]'
        row[4] = _clean_account_value(row[4])
        row[6] = _clean_account_value(row[6])
        row[7] = _clean_account_value(row[7])
        fw_rows.append(row)

    fw_columns = [
        'SNO',
        'Name Of Beneficiary',
        'Aadhar Number',
        'Contact Number',
        'Bank Acc Number',
        'Name As Per Bank Records',
        'RD1 Acc Num',
        'CIF NO',
    ]
    fw_df = pd.DataFrame(fw_rows, columns=fw_columns)

    # Check if fixed-width produced meaningful RD1 / Bank Acc data
    has_fw_data = False
    if not fw_df.empty:
        non_empty_acc = (
            (fw_df['Bank Acc Number'] != '') | (fw_df['RD1 Acc Num'] != '')
        ).sum()
        if non_empty_acc > 0:
            has_fw_data = True

    if mode == 'fixed_width' or (mode == 'auto' and has_fw_data):
        return fw_df

    # 2. Regex fallback parser
    cp1252_lines = raw_bytes.decode('cp1252', errors='ignore').splitlines()
    parsed_data = []
    for line in cp1252_lines:
        line_str = line.rstrip('\r\n')
        if (
            'SNO' in line_str
            or 'Contact Number' in line_str
            or not line_str.strip()
        ):
            continue

        sno_match = re.match(r'^(\d+)(.*)$', line_str.strip())
        if not sno_match:
            continue

        sno = sno_match.group(1)
        rem_str = sno_match.group(2).strip()

        parts = rem_str.split()
        rd1_val = '0'
        if parts:
            if parts[-1].isdigit() and len(parts[-1]) >= 8:
                rd1_val = parts[-1]
                rem_str = ' '.join(parts[:-1])
            elif (
                (
                    len(parts) >= 2
                    and parts[-1]
                    in ['CLOSED', 'ADDRESS', 'NUMBER', 'BALANCE']
                )
                or 'ACCOUNT' in rem_str
                or 'TEMP' in rem_str
            ):
                rd1_val = '0'

        pattern = r'^(.*?)\s+(\d{9,11})\s+(\d{9,11})(?:\s*(\d{9,11}))?(.*)$'
        match = re.match(pattern, rem_str)

        if match:
            ben_name, contact, bank_acc, cif, rest_bank = match.groups()
            parsed_data.append({
                'SNO': int(sno),
                'Name Of Beneficiary': ben_name.strip(),
                'Contact Number': _clean_account_value(contact),
                'Bank Acc Number': _clean_account_value(bank_acc),
                'CIF NO': _clean_account_value(cif) if cif else '',
                'NAME AS PER BANK RECORDS': (
                    rest_bank.strip() if rest_bank else ''
                ),
                'RD1 Acc Num': _clean_account_value(rd1_val),
            })
        else:
            parsed_data.append({
                'SNO': int(sno),
                'Name Of Beneficiary': rem_str,
                'Contact Number': '',
                'Bank Acc Number': '',
                'CIF NO': '',
                'NAME AS PER BANK RECORDS': '',
                'RD1 Acc Num': _clean_account_value(rd1_val),
            })

    return pd.DataFrame(parsed_data)


@excel_tools_bp.route('/process-txt-to-excel', methods=['POST'])
def process_txt_to_excel():
    """Text (.txt) ఫైల్‌ను పార్స్ చేసి ఫార్మాట్ చేసిన ఎక్సెల్ (.xlsx) ఫైల్‌గా డౌన్‌లోడ్ చేస్తుంది."""
    try:
        if 'txt_file' not in request.files:
            return 'దయచేసి .txt ఫైల్ అప్‌లోడ్ చేయండి.', 400
        file = request.files['txt_file']
        if not file or file.filename == '':
            return 'ఫైల్ ఎంచుకోలేదు.', 400

        custom_prefix = request.form.get('txt_output_prefix', '').strip()
        raw_bytes = file.read()
        df = _parse_term_deposit_text_bytes(raw_bytes)

        if df.empty:
            return 'టెక్స్ట్ ఫైల్‌లో ప్రాసెస్ చేయడానికి డేటా లభించలేదు.', 400

        text_cols = [
            c
            for c in ('Customer_No', 'Account_No', 'Tenor', 'Mat_Date')
            if c in df.columns
        ]
        currency_cols = [
            c for c in ('Balance', 'Avail_Int') if c in df.columns
        ]

        safe_base = _sanitize_filename_part(
            custom_prefix
            if custom_prefix
            else f'TXT_Converted_{os.path.splitext(file.filename)[0]}'
        )
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = f'{safe_base or "Text_To_Excel"}_{timestamp}.xlsx'

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Text_Data', index=False)
            ws = writer.book['Text_Data']
            _style_openpyxl_sheet(
                ws,
                f'TEXT TO EXCEL STATEMENT ({file.filename})',
                text_cols=text_cols,
                currency_cols=currency_cols,
            )
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'TXT టు Excel మార్పులో లోపం సంభవించింది: {str(e)}', 500


@excel_tools_bp.route('/process-prn-to-excel', methods=['POST'])
def process_prn_to_excel():
    """PRN (.prn) ఫైల్‌ను పార్స్ చేసి ఫార్మాట్ చేసిన ఎక్సెల్ (.xlsx) ఫైల్‌గా డౌన్‌లోడ్ చేస్తుంది."""
    try:
        if 'prn_file' not in request.files:
            return 'దయచేసి .prn ఫైల్ అప్‌లోడ్ చేయండి.', 400
        file = request.files['prn_file']
        if not file or file.filename == '':
            return 'ఫైల్ ఎంచుకోలేదు.', 400

        parse_mode = request.form.get('prn_parse_mode', 'auto').strip()
        custom_prefix = request.form.get('prn_output_prefix', '').strip()

        raw_bytes = file.read()
        df = _parse_prn_bytes(raw_bytes, mode=parse_mode)

        if df.empty:
            return 'PRN ఫైల్‌లో ప్రాసెస్ చేయడానికి డేటా లభించలేదు.', 400

        text_cols = [
            c
            for c in (
                'Aadhar Number',
                'Contact Number',
                'Bank Acc Number',
                'RD1 Acc Num',
                'CIF NO',
            )
            if c in df.columns
        ]

        safe_base = _sanitize_filename_part(
            custom_prefix
            if custom_prefix
            else f'PRN_Converted_{os.path.splitext(file.filename)[0]}'
        )
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = f'{safe_base or "PRN_To_Excel"}_{timestamp}.xlsx'

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='PRN_Data', index=False)
            ws = writer.book['PRN_Data']
            _style_openpyxl_sheet(
                ws,
                f'PRN TO EXCEL STATEMENT ({file.filename})',
                text_cols=text_cols,
            )
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'PRN టు Excel మార్పులో లోపం సంభవించింది: {str(e)}', 500


# ==============================================================================
# EXISTING EXCEL TOOLS (JOIN, RECONCILE, MASTER MATCH, DUPLICATES, CLEAN FORMAT)
# ==============================================================================
@excel_tools_bp.route('/process', methods=['POST'])
def process_join():
    """Excel Left Join Automation (In-Memory)."""
    join_type = request.form.get('join_type')
    custom_prefix = request.form.get('join_output_prefix', '').strip()

    try:
        if join_type == 'single_file':
            if 'file_single' not in request.files:
                return 'దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.', 400
            file = request.files['file_single']
            if not file or file.filename == '':
                return 'ఫైల్ ఎంచుకోలేదు.', 400
            sheet1 = request.form.get('sheet1', '').strip()
            col1 = request.form.get('col1', '').strip()
            sheet2 = request.form.get('sheet2', '').strip()
            col2 = request.form.get('col2', '').strip()

            df1 = _read_sheet_as_str(file, sheet1, fallback_idx=0)
            df2 = _read_sheet_as_str(file, sheet2, fallback_idx=1)

        elif join_type == 'two_files':
            if 'file1' not in request.files or 'file2' not in request.files:
                return 'దయచేసి రెండు ఫైల్స్ అప్‌లోడ్ చేయండి.', 400
            file1 = request.files['file1']
            file2 = request.files['file2']
            if not file1.filename or not file2.filename:
                return 'రెండు ఫైల్స్ సరిగ్గా ఎంచుకోండి.', 400

            sheet1_f1 = request.form.get('sheet1_f1', '').strip()
            col1 = request.form.get('col1_f1', '').strip()
            sheet2_f2 = request.form.get('sheet2_f2', '').strip()
            col2 = request.form.get('col2_f2', '').strip()

            df1 = _read_sheet_as_str(file1, sheet1_f1, fallback_idx=0)
            df2 = _read_sheet_as_str(file2, sheet2_f2, fallback_idx=0)
        else:
            return 'సరైన జాయిన్ ఆప్షన్ ఎంచుకోండి.', 400

        if col1 not in df1.columns or col2 not in df2.columns:
            return (
                f"లోపం: '{col1}' లేదా '{col2}' కాలమ్స్ ఎక్సెల్ ఫైల్‌లో కనుగొనబడలేదు!",
                400,
            )

        df1 = df1.copy()
        df2 = df2.copy()
        df1['_Clean_Key'] = df1[col1].apply(_clean_account_value)
        df2['_Clean_Key'] = df2[col2].apply(_clean_account_value)

        df1_subset = df1.drop_duplicates(subset=['_Clean_Key']).copy()
        df2_subset = df2.drop_duplicates(subset=['_Clean_Key']).copy()

        merged_df = pd.merge(
            df1_subset,
            df2_subset,
            on='_Clean_Key',
            how='left',
            suffixes=('_left', '_right'),
        ).drop(columns=['_Clean_Key'])

        safe_base = _sanitize_filename_part(
            custom_prefix if custom_prefix else 'Merged_RD_Accounts_Output'
        )
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = f'{safe_base or "Merged_Output"}_{timestamp}.xlsx'

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name='Left_Joined_Data', index=False)
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'ఫైల్ ప్రాసెసింగ్‌లో లోపం సంభవించింది: {str(e)}', 500


@excel_tools_bp.route('/process-reconcile', methods=['POST'])
def process_reconcile():
    """RD ఖాతాల రీకన్సిలియేషన్ & ఫార్మాటర్ టూల్ (In-Memory)."""
    try:
        if 'reconcile_file' not in request.files:
            return 'దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.', 400
        file = request.files['reconcile_file']
        if not file or file.filename == '':
            return 'ఫైల్ ఎంచుకోలేదు.', 400

        sheet_name = request.form.get('rec_sheet_name', '').strip()
        rd_col = request.form.get('rec_rd_col', 'Rd').strip()
        ben_col = request.form.get('rec_ben_col', 'Ben').strip()
        amt_col = request.form.get('rec_amt_col', 'Amount').strip()

        file.stream.seek(0)
        excel_file = pd.ExcelFile(file)
        target_sheet = (
            sheet_name
            if (sheet_name and sheet_name in excel_file.sheet_names)
            else excel_file.sheet_names[0]
        )
        df = pd.read_excel(excel_file, sheet_name=target_sheet)
        df.columns = [str(c).strip() for c in df.columns]

        for req_col in (rd_col, ben_col, amt_col):
            if req_col not in df.columns:
                return f"'{req_col}' కాలమ్ షీట్‌లో కనుగొనబడలేదు.", 400

        df_clean = df[[rd_col, ben_col, amt_col]].dropna(
            subset=[rd_col, ben_col]
        ).copy()
        df_clean['Ben_clean'] = df_clean[ben_col].astype(str).str.strip()
        df_clean = df_clean[df_clean['Ben_clean'] != '']

        def format_rd(val):
            cleaned = _clean_account_value(val)
            try:
                if cleaned.isdigit():
                    return str(int(cleaned))
                return str(int(float(cleaned)))
            except Exception:
                return cleaned

        df_clean['Rd_formatted'] = df_clean[rd_col].apply(format_rd)
        df_clean = df_clean[df_clean['Rd_formatted'] != '']

        if df_clean.empty:
            return 'ప్రాసెస్ చేయడానికి చెల్లుబాటు అయ్యే RD/Beneficiary డేటా లేదు.', 400

        unique_bens = list(dict.fromkeys(df_clean['Ben_clean'].tolist()))
        max_rds = int(df_clean.groupby('Ben_clean')['Rd_formatted'].count().max())

        rows_data = []
        for idx, ben in enumerate(unique_bens, 1):
            ben_rows = df_clean[df_clean['Ben_clean'] == ben]
            row_dict = {'S.No': idx, 'Name Of Beneficiary': ben}

            rds = ben_rows['Rd_formatted'].tolist()
            amounts = ben_rows[amt_col].tolist()

            for i in range(len(rds)):
                acc_num = i + 1
                row_dict[f'RD Account {acc_num}'] = str(rds[i])
                amt_num = pd.to_numeric(amounts[i], errors='coerce')
                row_dict[f'Closing Balance {acc_num}'] = (
                    float(amt_num) if pd.notnull(amt_num) else 0.0
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
        TOTAL_FILL = PatternFill(
            start_color='E2E8F0', end_color='E2E8F0', fill_type='solid'
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
                    value=float(bal_val) if pd.notnull(bal_val) else 0.0,
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
            rd_range_span = (
                f'{rd_cols_letters[0]}{row_num}:{rd_cols_letters[-1]}{row_num}'
            )

            c_tot_rd = ws.cell(row=row_num, column=tot_rd_col)
            c_tot_rd.value = (
                f'=COUNTA({rd_range_str})-COUNTIF({rd_range_span}, "-")'
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
        ws.row_dimensions[total_row].height = 24
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
            cell_tot.fill = TOTAL_FILL
            cell_tot.border = thin_border
            if c == tot_rd_col:
                cell_tot.value = f'=SUM({col_ltr}3:{col_ltr}{total_row-1})'
                cell_tot.font = Font(name='Calibri', size=10, bold=True)
                cell_tot.alignment = Alignment(
                    horizontal='center', vertical='center'
                )
            elif c == tot_bal_col:
                cell_tot.value = f'=SUM({col_ltr}3:{col_ltr}{total_row-1})'
                cell_tot.number_format = '₹#,##0.00'
                cell_tot.font = Font(name='Calibri', size=10, bold=True)
                cell_tot.alignment = Alignment(
                    horizontal='right', vertical='center'
                )
            elif c % 2 == 0:
                cell_tot.value = f'=SUM({col_ltr}3:{col_ltr}{total_row-1})'
                cell_tot.number_format = '₹#,##0.00'
                cell_tot.alignment = Alignment(
                    horizontal='right', vertical='center'
                )
            else:
                cell_tot.value = '-'
                cell_tot.alignment = Alignment(
                    horizontal='center', vertical='center'
                )

        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 30
        for c in range(3, len(all_cols) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

        ws.freeze_panes = 'C3'

        output_filename = (
            f'RD_Accounts_Reconciled_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=output_filename,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'Error processing reconciliation file: {str(e)}', 400


@excel_tools_bp.route('/process-master-match', methods=['POST'])
def process_master_match():
    """Master (acc 1, acc 2) vs Data (acc) మ్యాచ్ చేసి Matched & Unmatched రిపోర్ట్ ఇస్తుంది."""
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

            master_df = _read_sheet_as_str(
                file_single, master_sheet, fallback_idx=0
            )
            data_df = _read_sheet_as_str(file_single, data_sheet, fallback_idx=1)
            source_label = os.path.splitext(file_single.filename)[0]

        elif mm_mode == 'two_files':
            if (
                'mm_file_master' not in request.files
                or 'mm_file_data' not in request.files
            ):
                return 'దయచేసి Master ఫైల్ మరియు Data ఫైల్ — రెండూ అప్‌లోడ్ చేయండి.', 400
            file_master = request.files['mm_file_master']
            file_data = request.files['mm_file_data']

            if (
                not file_master
                or file_master.filename == ''
                or not file_data
                or file_data.filename == ''
            ):
                return 'రెండు ఎక్సెల్ ఫైల్స్ సరిగ్గా ఎంచుకోండి.', 400

            master_sheet = request.form.get('mm_two_master_sheet', '').strip()
            data_sheet = request.form.get('mm_two_data_sheet', '').strip()

            master_df = _read_sheet_as_str(
                file_master, master_sheet, fallback_idx=0
            )
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

        safe_base = _sanitize_filename_part(
            custom_prefix
            if custom_prefix
            else f'Master_Match_{source_label}'
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


@excel_tools_bp.route('/process-duplicates', methods=['POST'])
def process_duplicates():
    """ఎంచుకున్న కాలమ్ ఆధారంగా Duplicate మరియు Unique రికార్డులను వేరు చేసి రిపోర్ట్ ఇస్తుంది."""
    try:
        if 'dup_file' not in request.files:
            return 'దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.', 400
        file = request.files['dup_file']
        if not file or file.filename == '':
            return 'ఫైల్ ఎంచుకోలేదు.', 400

        sheet_name = request.form.get('dup_sheet', '').strip()
        check_col = request.form.get('dup_col', '').strip()
        keep_option = request.form.get('dup_keep', 'first').strip()

        df = _read_sheet_as_str(file, sheet_name, fallback_idx=0)
        if not check_col or check_col not in df.columns:
            return 'దయచేసి సరైన చెక్ కాలమ్ ఎంచుకోండి.', 400

        df = df.copy()
        df['_Norm_Key'] = (
            df[check_col].apply(_clean_account_value).str.upper()
        )

        counts = df['_Norm_Key'].value_counts()
        dup_keys = set(counts[counts > 1].index) - {''}
        is_duplicate = df['_Norm_Key'].isin(dup_keys)

        duplicates_df = df[is_duplicate].drop(columns=['_Norm_Key']).copy()

        if keep_option == 'none':
            unique_df = df[~is_duplicate].drop(columns=['_Norm_Key']).copy()
        else:
            seen = set()
            keep_indices = []
            for idx, k in enumerate(df['_Norm_Key']):
                if k == '' or k not in seen:
                    if k != '':
                        seen.add(k)
                    keep_indices.append(idx)
            unique_df = df.iloc[keep_indices].drop(columns=['_Norm_Key']).copy()

        safe_base = _sanitize_filename_part(
            os.path.splitext(file.filename)[0]
        )
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = f'Duplicates_Check_{safe_base or "Report"}_{timestamp}.xlsx'

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            unique_df.to_excel(writer, sheet_name='Unique_Records', index=False)
            duplicates_df.to_excel(
                writer, sheet_name='Duplicate_Records', index=False
            )
            wb = writer.book
            if 'Unique_Records' in wb.sheetnames:
                _style_openpyxl_sheet(
                    wb['Unique_Records'],
                    f'UNIQUE RECORDS ({check_col}) - COUNT: {len(unique_df)}',
                )
            if 'Duplicate_Records' in wb.sheetnames:
                _style_openpyxl_sheet(
                    wb['Duplicate_Records'],
                    f'DUPLICATE RECORDS ({check_col}) - COUNT: {len(duplicates_df)}',
                )
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'Duplicate చెక్ ప్రాసెసింగ్‌లో లోపం: {str(e)}', 500


@excel_tools_bp.route('/process-clean-format', methods=['POST'])
def process_clean_format():
    """అకౌంట్ నంబర్లను Text (@) గా మరియు అమౌంట్‌లను Currency (₹) గా క్లీన్ చేసి ఎక్సెల్ ఇస్తుంది."""
    try:
        if 'clean_file' not in request.files:
            return 'దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.', 400
        file = request.files['clean_file']
        if not file or file.filename == '':
            return 'ఫైల్ ఎంచుకోలేదు.', 400

        sheet_name = request.form.get('clean_sheet', '').strip()
        acc_cols_raw = request.form.getlist('clean_acc_cols')
        amt_cols_raw = request.form.getlist('clean_amt_cols')
        drop_empty = request.form.get('clean_drop_empty') == 'on'

        df = _read_sheet_as_str(file, sheet_name, fallback_idx=0)
        df = df.copy()

        for col in acc_cols_raw:
            if col in df.columns:
                df[col] = df[col].apply(_clean_account_value)

        numeric_amt_cols = []
        for col in amt_cols_raw:
            if col in df.columns:
                cleaned_amt = (
                    df[col]
                    .astype(str)
                    .str.replace('₹', '', regex=False)
                    .str.replace(',', '', regex=False)
                    .str.strip()
                )
                df[col] = pd.to_numeric(cleaned_amt, errors='coerce').fillna(0.0)
                numeric_amt_cols.append(col)

        if drop_empty:
            df = df.dropna(how='all')
            non_empty_mask = df.apply(
                lambda r: any(str(v).strip() != '' for v in r), axis=1
            )
            df = df[non_empty_mask]

        safe_base = _sanitize_filename_part(
            os.path.splitext(file.filename)[0]
        )
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out_name = f'Formatted_Data_{safe_base or "Sheet"}_{timestamp}.xlsx'

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
            wb = writer.book
            ws = wb['Cleaned_Data']
            _style_openpyxl_sheet(
                ws,
                'STANDARDIZED & FORMATTED OFFICE DATA',
                text_cols=acc_cols_raw,
                currency_cols=numeric_amt_cols,
            )
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
        )

    except Exception as e:
        return f'డేటా ఫార్మాటింగ్‌లో లోపం: {str(e)}', 500


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
        return (
            jsonify({
                'status': 'error',
                'message': f'ఫైల్ రీడ్ చేయడంలో లోపం: {str(e)}',
            }),
            500,
        )


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
        return (
            jsonify({
                'status': 'error',
                'message': f'కాలమ్స్ లోడ్ చేయడంలో లోపం: {str(e)}',
            }),
            500,
        )


# ==============================================================================
# GENERALIZED & SAFE PARSER + DEDUPLICATOR EXCEL TOOL (/process-safe-parser)
# ==============================================================================
def _safe_clean_digits(val, expected_len=None, allow_status=False):
    """
    నంబర్లను స్ట్రింగ్‌లుగా క్లీన్ చేస్తుంది.
    సైంటిఫిక్ నోటేషన్ (3.89674E+11) లేదా .0 వస్తే సాధారణ అంకెలుగా మారుస్తుంది.
    స్టేటస్ టెక్స్ట్ ఉంటే (ACCOUNT CLOSED, TEMP ADDRESS మొదలైనవి) వేరుగా రిటర్న్ చేస్తుంది.
    """
    if val is None or pd.isna(val):
        return None, None

    s = str(val).strip()
    if s == "" or s.lower() in ("nan", "none", "null", "-", "nat"):
        return None, None

    # Scientific notation (ఉదా: 3.89674E+11)
    if re.match(r"^[+-]?\d+(\.\d+)?[eE][+-]?\d+$", s):
        try:
            s = f"{int(float(s))}"
        except Exception:
            pass

    # Trailing .0
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]

    s_clean = s.replace("'", "").strip()
    digits_only = re.sub(r"\D", "", s_clean)

    status_keywords = (
        "CLOSED", "TEMP", "BALANCE", "INOPERATIVE", "WRONG", "HOLD", "FREEZE", "MISSING", "NA"
    )
    if allow_status and (
        len(digits_only) < 6 or any(k in s_clean.upper() for k in status_keywords)
    ):
        return None, s_clean.upper()

    if not digits_only:
        return None, s_clean

    flag = None
    if expected_len and len(digits_only) != expected_len:
        flag = f"LEN_{len(digits_only)}_INSTEAD_OF_{expected_len}"

    return digits_only, flag


def _safe_normalize_name(name):
    """పేర్లలోని అదనపు స్పేస్‌లు, పక్కపక్కనే రిపీట్ అయ్యే పదాలను సరిచేస్తుంది."""
    if name is None or pd.isna(name):
        return ""
    s = re.sub(r"\s+", " ", str(name)).strip()
    parts = s.split(" ")
    cleaned_parts = []
    for p in parts:
        if not cleaned_parts or cleaned_parts[-1].lower() != p.lower():
            cleaned_parts.append(p)
    return " ".join(cleaned_parts)


def _parse_safe_text_lines(raw_text: str):
    """SNO తో మొదలయ్యే టెక్స్ట్/రిపోర్ట్ లైన్లను పార్స్ చేస్తుంది."""
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    records = []
    unmatched_lines = []

    
    # SNO తర్వాత స్పేస్ ఉన్నా లేకపోయినా పేరును వేరు చేయడానికి \s*(?P<ben_name>[A-Za-z].+?) వాడాము 
    pattern = re.compile(
      r"^(?P<sno>\d+)\s*"
      r"(?P<ben_name>[A-Za-z].+?)\s+"
      r"(?P<id12>\d{11,12})\s+"
      r"(?P<phone>\d{10,12})\s+"
      r"(?P<bank_acc>\d{6,20})\s+"
      r"(?P<bank_name>.+?)\s+"
      r"(?P<rd1>\d{6,20}|ACCOUNT CLOSED|TEMP ADDRESS|NO BALANCE|INOPERATIVE ACCOUNT|WRONG ACCOUNT NUMBER)\s+"
      r"(?P<cif>\d{6,20})$",
      re.IGNORECASE,
    )

    for line in lines:
        if line.upper().startswith("SNO ") and "CIF" in line.upper():
            continue
        m = pattern.match(line)
        if m:
            d = m.groupdict()
            ben_clean = _safe_normalize_name(d["ben_name"])
            bank_name_clean = _safe_normalize_name(d["bank_name"])

            phone_clean, phone_flag = _safe_clean_digits(d["phone"], expected_len=10)
            acc_clean, acc_flag = _safe_clean_digits(d["bank_acc"])
            rd1_acc, rd1_status = _safe_clean_digits(d["rd1"], allow_status=True)
            cif_clean, cif_flag = _safe_clean_digits(d["cif"])

            flags = []
            if phone_flag:
                flags.append(f"PHONE:{phone_flag}")
            if acc_flag:
                flags.append(f"BANK_ACC:{acc_flag}")
            if cif_flag:
                flags.append(f"CIF:{cif_flag}")

            records.append({
                "SNO": int(d["sno"]),
                "Beneficiary_Name": ben_clean,
                "Name_As_Per_Bank": bank_name_clean,
                "Name_Match": ben_clean.lower() == bank_name_clean.lower(),
                "ID_Number": "[ID Redacted]",
                "Contact_Number": phone_clean or "",
                "Bank_Acc_Number": acc_clean or "",
                "RD1_Acc_Number": rd1_acc or "",
                "RD1_Status": rd1_status if not rd1_acc else "ACTIVE",
                "CIF_Number": cif_clean or "",
                "Data_Flags": ", ".join(flags) if flags else "OK",
            })
        else:
            unmatched_lines.append({"Raw_Line": line})

    return pd.DataFrame(records), pd.DataFrame(unmatched_lines)


def _parse_safe_tabular_df(df_in: pd.DataFrame):
    """Excel లేదా CSV డేటాఫ్రేమ్‌ను సురక్షితంగా క్లీన్ చేసి స్టాండర్డ్ ఫార్మాట్‌లోకి మారుస్తుంది."""
    df_in = df_in.copy()
    df_in.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df_in.columns]

    def find_col(candidates):
        lower_map = {c.lower(): c for c in df_in.columns}
        # Exact match first
        for cand in candidates:
            if cand.lower() in lower_map:
                return lower_map[cand.lower()]
        # Substring match
        for c in df_in.columns:
            cl = c.lower()
            if any(cand.lower() in cl for cand in candidates):
                return c
        return None

    col_sno = find_col(["sno", "s.no", "sl.no", "bank_sno", "serial"])
    col_ben = find_col(["benificiary", "beneficiary", "ben name", "name"])
    col_bank_name = find_col(["as per bank", "bank name"])
    col_phone = find_col(["contact", "phone", "mobile"])
    col_bank_acc = find_col(["bank acc", "bank account", "account num", "acc no"])
    col_rd1 = find_col(["rd1", "rd 1", "rd_1", "rd account", "rd acc"])
    col_cif = find_col(["cif"])

    records = []
    for idx, row in df_in.iterrows():
        sno_raw = row[col_sno] if col_sno else (idx + 1)
        sno_digits, _ = _safe_clean_digits(sno_raw)
        sno_val = int(sno_digits) if sno_digits else (idx + 1)

        ben_clean = _safe_normalize_name(row[col_ben]) if col_ben else ""
        bank_name_clean = _safe_normalize_name(row[col_bank_name]) if col_bank_name else ben_clean

        phone_clean, phone_flag = _safe_clean_digits(row[col_phone], expected_len=10) if col_phone else (None, None)
        acc_clean, acc_flag = _safe_clean_digits(row[col_bank_acc]) if col_bank_acc else (None, None)
        rd1_acc, rd1_status = _safe_clean_digits(row[col_rd1], allow_status=True) if col_rd1 else (None, "MISSING")
        cif_clean, cif_flag = _safe_clean_digits(row[col_cif]) if col_cif else (None, None)

        # Skip completely empty rows
        if not any([ben_clean, acc_clean, rd1_acc, cif_clean]):
            continue

        flags = []
        if phone_flag:
            flags.append(f"PHONE:{phone_flag}")
        if acc_flag:
            flags.append(f"BANK_ACC:{acc_flag}")
        if cif_flag:
            flags.append(f"CIF:{cif_flag}")

        records.append({
            "SNO": sno_val,
            "Beneficiary_Name": ben_clean,
            "Name_As_Per_Bank": bank_name_clean,
            "Name_Match": ben_clean.lower() == bank_name_clean.lower() if (ben_clean and bank_name_clean) else True,
            "Contact_Number": phone_clean or "",
            "Bank_Acc_Number": acc_clean or "",
            "RD1_Acc_Number": rd1_acc or "",
            "RD1_Status": rd1_status if not rd1_acc else "ACTIVE",
            "CIF_Number": cif_clean or "",
            "Data_Flags": ", ".join(flags) if flags else "OK",
        })

    return pd.DataFrame(records), pd.DataFrame(columns=["Raw_Line"])


def _deduplicate_safe_df(df: pd.DataFrame) -> pd.DataFrame:
    """యాక్టివ్ RD1 అకౌంట్‌కు ప్రాధాన్యత ఇస్తూ డూప్లికేట్ రికార్డులను తొలగిస్తుంది."""
    if df.empty:
        return df

    work = df.copy()
    work["_has_rd1"] = (work["RD1_Acc_Number"].astype(str).str.strip() != "").astype(int)
    work = work.sort_values(by=["_has_rd1", "SNO"], ascending=[False, True])

    has_rd1 = work[work["_has_rd1"] == 1].drop_duplicates(subset=["RD1_Acc_Number"], keep="first")
    no_rd1 = work[work["_has_rd1"] == 0].drop_duplicates(
        subset=["Bank_Acc_Number", "Beneficiary_Name"], keep="first"
    )

    merged = pd.concat([has_rd1, no_rd1], ignore_index=True)
    merged = merged.drop(columns=["_has_rd1"]).sort_values("SNO").reset_index(drop=True)
    return merged


@excel_tools_bp.route("/process-safe-parser", methods=["POST"])
def process_safe_parser():
    """
    Generalized & Safe Parser Tool:
    .xlsx, .csv, .txt, .prn ఫైళ్లను క్లీన్ చేసి, డూప్లికేట్స్ తొలగించి,
    Cleaned_Data, Status_And_Flags, Unmatched_Lines షీట్లతో ఎక్సెల్ ఫైల్ ఇస్తుంది.
    """
    try:
        if "safe_file" not in request.files:
            return "దయచేసి ఫైల్ అప్‌లోడ్ చేయండి.", 400
        file = request.files["safe_file"]
        if not file or file.filename == "":
            return "ఫైల్ ఎంచుకోలేదు.", 400

        ext = os.path.splitext(file.filename)[1].lower()
        sheet_name = request.form.get("safe_sheet_name", "").strip()
        custom_prefix = request.form.get("safe_output_prefix", "").strip()

        raw_bytes = file.read()

        if ext in (".txt", ".prn"):
            text_content = raw_bytes.decode("utf-8", errors="ignore")
            if not text_content.strip():
                text_content = raw_bytes.decode("cp1252", errors="ignore")
            parsed_df, unmatched_df = _parse_safe_text_lines(text_content)
        elif ext == ".csv":
            df_csv = pd.read_csv(io.BytesIO(raw_bytes), dtype=str)
            parsed_df, unmatched_df = _parse_safe_tabular_df(df_csv)
        else:
            excel_file = pd.ExcelFile(io.BytesIO(raw_bytes))
            target_sheet = (
                sheet_name
                if (sheet_name and sheet_name in excel_file.sheet_names)
                else excel_file.sheet_names[0]
            )
            df_xl = pd.read_excel(excel_file, sheet_name=target_sheet, dtype=str)
            parsed_df, unmatched_df = _parse_safe_tabular_df(df_xl)

        if parsed_df.empty and unmatched_df.empty:
            return "ఫైల్‌లో ప్రాసెస్ చేయడానికి డేటా లభించలేదు.", 400

        cleaned_df = _deduplicate_safe_df(parsed_df)
        flagged_df = (
            parsed_df[
                (parsed_df["RD1_Status"] != "ACTIVE")
                | (parsed_df["Data_Flags"] != "OK")
                | (~parsed_df["Name_Match"])
            ].copy()
            if not parsed_df.empty
            else pd.DataFrame()
        )

        text_cols = [
            c for c in ("ID_Number", "Contact_Number", "Bank_Acc_Number", "RD1_Acc_Number", "CIF_Number")
            if c in cleaned_df.columns
        ]

        safe_base = _sanitize_filename_part(
            custom_prefix if custom_prefix else f"Safe_Parsed_{os.path.splitext(file.filename)[0]}"
        )
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"{safe_base or 'Safe_Parsed_Report'}_{timestamp}.xlsx"

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            cleaned_df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
            flagged_df.to_excel(writer, sheet_name="Status_And_Flags", index=False)
            unmatched_df.to_excel(writer, sheet_name="Unmatched_Lines", index=False)

            wb = writer.book
            _style_openpyxl_sheet(
                wb["Cleaned_Data"],
                f"CLEANED & DEDUPLICATED DATA ({len(cleaned_df)} Records)",
                text_cols=text_cols,
            )
            _style_openpyxl_sheet(
                wb["Status_And_Flags"],
                f"STATUS / NAME MISMATCH / ANOMALY FLAGS ({len(flagged_df)} Records)",
                text_cols=text_cols,
            )
            _style_openpyxl_sheet(
                wb["Unmatched_Lines"],
                f"UNMATCHED / RAW LINES ({len(unmatched_df)} Lines)",
            )
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return f"Safe Parser ప్రాసెసింగ్‌లో లోపం సంభవించింది: {str(e)}", 500
