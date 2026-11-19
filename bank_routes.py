import csv
import io
import math
from flask import Blueprint, jsonify, make_response, render_template, request, send_file, session
import gspread
import pandas as pd

bank_bp = Blueprint('bank_bp', __name__)


@bank_bp.route('/')
def home():
  from app import SPREADSHEET_NAME, get_sheets_client
  from sheets_routes import cf_sheet_ids, month_map

  bank_list = []
  try:
    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    login_sheet = spreadsheet.worksheet('BANK_LOGINS')
    logins = login_sheet.get_all_records()

    for row in logins:
      b_name = (
          row.get('Bank Name') or row.get('BANK NAME') or row.get('Bank_Name')
      )
      if b_name:
        cleaned_name = str(b_name).strip()
        if cleaned_name and cleaned_name not in bank_list:
          bank_list.append(cleaned_name)

  except Exception as e:
    print('Error fetching bank list:', e)
    bank_list = []

  cf_list = list(cf_sheet_ids.keys())
  return render_template(
      'index.html', banks=bank_list, months=month_map, cf_list=cf_list
  )


@bank_bp.route('/api/search', methods=['POST'])
def search_beneficiary():
  try:
    from app import SPREADSHEET_NAME, get_sheets_client

    data = request.get_json()
    if not data:
      return (
          jsonify({'success': False, 'message': 'సరైన డేటా అందలేదు.'}),
          400,
      )

    input_id = (
        str(data.get('aadhar', '')).strip().replace('-', '').replace(' ', '')
    )

    if not input_id or len(input_id) != 12 or not input_id.isdigit():
      return (
          jsonify({
              'success': False,
              'message': 'దయచేసి సరైన 12 అంకెల సంఖ్యను ఎంటర్ చేయండి.',
          }),
          400,
      )

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
    records = target_sheet.get_all_records()

    matched_record = None
    for row in records:
      raw_val = (
          str(
              row.get('Aadhar Number')
              or row.get('ADHAR')
              or row.get('AADHAR')
              or ''
          )
          .split('.')[0]
          .strip()
      )
      if raw_val == input_id:
        matched_record = row
        break

    if matched_record:
      masked_id = f'XXXX-XXXX-{input_id[-4:]}'
      result = {
          'sno': matched_record.get('BANK_SNO'),
          'masked_aadhar': masked_id,
          'name': (
              matched_record.get('Name Of Benificiary')
              or matched_record.get('NAME')
          ),
          'father_name': (
              matched_record.get('Father Name')
              or matched_record.get('FATHER NAME')
          ),
          'village': (
              matched_record.get('VILLAGE NAME') or matched_record.get('VILLAGE')
          ),
          'mandal': (
              matched_record.get('MANDAL NAME') or matched_record.get('MANDAL')
          ),
          'bank_name': (
              matched_record.get('Bank Name') or matched_record.get('BANK NAME')
          ),
          'rd1_status': matched_record.get('RD 1'),
          'rd2_status': matched_record.get('RD 2'),
          'wages': matched_record.get('MONTHLY WAGES'),
      }
      return jsonify({'success': True, 'data': result})
    else:
      return jsonify(
          {'success': False, 'message': 'సదరు సంఖ్యతో ఎలాంటి వివరాలు లభించలేదు.'}
      )

  except Exception as e:
    return (
        jsonify(
            {'success': False, 'message': f'సర్వర్ లోపం సంభవించింది: {str(e)}'}
        ),
        500,
    )


@bank_bp.route('/api/login', methods=['POST'])
def bank_login():
  try:
    from app import SPREADSHEET_NAME, get_sheets_client

    data = request.get_json()
    bank_name = str(data.get('bank_name', '')).strip()
    pin = str(data.get('pin', '')).strip()

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    login_sheet = spreadsheet.worksheet('BANK_LOGINS')
    logins = login_sheet.get_all_records()

    authenticated = False
    for row in logins:
      sheet_bank = str(
          row.get('Bank Name') or row.get('BANK NAME') or ''
      ).strip()
      sheet_pin = str(row.get('PIN') or row.get('Pin') or '').strip()

      if sheet_bank.lower() == bank_name.lower() and sheet_pin == pin:
        authenticated = True
        break

    if authenticated:
      session['bank_user'] = bank_name
      session.permanent = True
      return jsonify({'status': 'success', 'message': 'లాగిన్ విజయవంతమైంది!'})
    else:
      return jsonify(
          {'status': 'error', 'message': 'తప్పు PIN లేదా బ్యాంక్ పేరు!'}
      )

  except Exception as e:
    return jsonify({'status': 'error', 'message': f'లాగిన్ లోపం: {str(e)}'})


@bank_bp.route('/api/bank-data', methods=['GET'])
def get_bank_data():
  try:
    from app import SPREADSHEET_NAME, get_sheets_client

    bank_user = session.get('bank_user')
    if not bank_user:
      return jsonify(
          {'status': 'error', 'message': 'దయచేసి ముందుగా లాగిన్ అవ్వండి!'}
      )

    current_month = str(
        request.args.get('month', session.get('current_month', '1'))
    ).strip()
    batch_number = str(
        request.args.get('batch', session.get('batch_number', '1'))
    ).strip()

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
    master_records = target_sheet.get_all_records()

    # --- బ్యాంకుకు పంపని (NOT SENT) రికార్డులను తొలగించుటకు ఫిల్టర్ ---
    master_records = [
        row
        for row in master_records
        if str(row.get('BANK SENT', '')).strip().upper() != 'NOT SENT'
    ]

    try:
      trans_sheet = spreadsheet.worksheet('Transactions_Year1')
      trans_records = trans_sheet.get_all_records()
    except Exception:
      trans_records = []

    bank_user_clean = bank_user.strip().lower()

    balance_map = {}
    completed_transactions = set()
    already_covered_in_this_month = set()

    for t in trans_records:
      t_sno = str(t.get('BANK_SNO') or t.get('Bank SNO') or '').strip()
      t_month = str(
          t.get('MONTH_NUMBER') or t.get('month_number') or t.get('MONTH') or ''
      ).strip()
      t_batch = str(
          t.get('BATCH_NUMBER')
          or t.get('BATCH')
          or t.get('batch_number')
          or ''
      ).strip()
      t_filename = (
          str(t.get('FILE NAME') or t.get('File Name') or '').strip().lower()
      )

      if t_month == current_month and t_filename == bank_user_clean:
        bal_val = (
            t.get(
                'Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న'
                ' బ్యాలెన్స్)'
            )
            or t.get('Bank Closing Balance')
            or ''
        )
        if t_sno:
          balance_map[t_sno] = bal_val

        try:
          govt_share_val = float(
              t.get('Govt_Share') or t.get('GOVT_SHARE') or 0
          )
        except ValueError:
          govt_share_val = 0

        if t_batch != batch_number and govt_share_val > 0:
          if t_sno:
            already_covered_in_this_month.add(t_sno)

        if t_batch == batch_number and govt_share_val > 0:
          if t_sno:
            completed_transactions.add(t_sno)

    filtered_data = []
    for row in master_records:
      file_name = (
          str(row.get('FILE NAME') or row.get('File Name') or '')
          .strip()
          .lower()
      )

      if file_name == bank_user_clean:
        bank_sno = str(row.get('BANK_SNO') or '').strip()

        if bank_sno in already_covered_in_this_month:
          continue

        if bank_sno in balance_map:
          row['Bank Closing Balance'] = balance_map[bank_sno]

        is_completed = bank_sno in completed_transactions
        row['is_completed'] = is_completed
        filtered_data.append(row)

    return jsonify({
        'status': 'success',
        'bank': bank_user,
        'month': current_month,
        'batch': batch_number,
        'data': filtered_data,
    })

  except Exception as e:
    return jsonify(
        {'status': 'error', 'message': f'డేటా లోడ్ అవ్వడంలో లోపం: {str(e)}'}
    )


@bank_bp.route('/api/download-csv', methods=['GET'])
def download_csv():
  from app import SPREADSHEET_NAME, get_sheets_client

  bank_user = session.get('bank_user')
  if not bank_user:
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  client = get_sheets_client()
  spreadsheet = client.open(SPREADSHEET_NAME)
  target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
  records = target_sheet.get_all_records()

  # --- బ్యాంకుకు పంపని (NOT SENT) రికార్డులను తొలగించుట ---
  records = [
      row
      for row in records
      if str(row.get('BANK SENT', '')).strip().upper() != 'NOT SENT'
  ]

  bank_user_clean = bank_user.strip().lower()
  filtered_data = [
      row
      for row in records
      if str(row.get('FILE NAME') or row.get('File Name') or '')
      .strip()
      .lower()
      == bank_user_clean
  ]

  output = []
  output.append(['SNO', 'Name', 'Account No', 'File Name'])

  for idx, row in enumerate(filtered_data):
    output.append([
        idx + 1,
        row.get('Name Of Benificiary') or row.get('NAME') or '',
        row.get('Bank Account Number') or row.get('ACCOUNT') or '',
        row.get('FILE NAME') or '',
    ])

  response = make_response()
  writer = csv.writer(response.stream)
  for row in output:
    writer.writerow(row)

  response.headers['Content-Disposition'] = (
      f'attachment; filename={bank_user}_beneficiaries.csv'
  )
  response.headers['Content-Type'] = 'text/csv'
  return response


@bank_bp.route('/api/update-rd-accounts', methods=['POST'])
def update_rd_accounts():
  try:
    from app import SPREADSHEET_NAME, get_sheets_client

    bank_user = session.get('bank_user')
    if not bank_user:
      return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

    data = request.get_json()
    bank_sno = str(data.get('bank_sno')).strip()
    rd1 = str(data.get('rd1', '')).strip()
    rd2 = str(data.get('rd2', '')).strip()
    bank_closing_balance_str = str(data.get('bank_closing_balance', '')).strip()

    current_month = str(
        data.get('month', session.get('current_month', '1'))
    ).strip()
    batch_number = str(
        data.get('batch', session.get('batch_number', '1'))
    ).strip()

    try:
      bank_closing_balance = (
          float(bank_closing_balance_str) if bank_closing_balance_str else 0.0
      )
    except ValueError:
      bank_closing_balance = 0.0

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
    master_records = target_sheet.get_all_records()

    master_row_index = None
    beneficiary_name = ''
    beneficiary_acc = ''
    monthly_wages = 0.0
    bank_user_clean = bank_user.strip().lower()

    for idx, row in enumerate(master_records):
      row_bank_sno = str(row.get('BANK_SNO') or row.get('Bank SNO') or '').strip()
      file_name = (
          str(row.get('FILE NAME') or row.get('File Name') or '')
          .strip()
          .lower()
      )

      if row_bank_sno == bank_sno and file_name == bank_user_clean:
        master_row_index = idx + 2
        beneficiary_name = str(
            row.get('Name Of Benificiary') or row.get('NAME') or ''
        ).strip()
        beneficiary_acc = str(
            row.get('Bank Account Number') or row.get('ACCOUNT') or ''
        ).strip()

        try:
          wages_val = (
              row.get('8% Of Monthly Wages') or row.get('MONTHLYWAGES') or 0
          )
          monthly_wages = float(wages_val)
        except ValueError:
          monthly_wages = 0.0
        break

    if master_row_index:
      header_row = target_sheet.row_values(1)
      rd1_col, rd2_col = None, None
      for c_idx, header in enumerate(header_row):
        h_cleaned = header.strip().upper()
        if h_cleaned in ['RD 1', 'RD_1']:
          rd1_col = c_idx + 1
        elif h_cleaned in ['RD 2', 'RD_2']:
          rd2_col = c_idx + 1

      if rd1_col:
        target_sheet.update_cell(master_row_index, rd1_col, rd1)
      if rd2_col:
        target_sheet.update_cell(master_row_index, rd2_col, rd2)

    due_amount = monthly_wages

    trans_sheet = spreadsheet.worksheet('Transactions_Year1')
    trans_records = trans_sheet.get_all_records()

    trans_row_index = None

    for idx, row in enumerate(trans_records):
      t_sno = str(row.get('BANK_SNO') or row.get('Bank SNO') or '').strip()
      t_month = str(
          row.get('MONTH_NUMBER') or row.get('month_number') or ''
      ).strip()
      t_batch = str(
          row.get('BATCH_NUMBER') or row.get('batch_number') or ''
      ).strip()
      t_filename = (
          str(row.get('FILE NAME') or row.get('File Name') or '').strip().lower()
      )

      if (
          t_sno == bank_sno
          and t_month == current_month
          and t_batch == batch_number
          and t_filename == bank_user_clean
      ):
        trans_row_index = idx + 2
        try:
          existing_due = float(row.get('(C) ఈ నెల డ్యూ(Due)') or 0)
          if existing_due > 0:
            due_amount = existing_due
        except ValueError:
          pass
        break

    trans_headers = trans_sheet.row_values(1)
    col_map = {h.strip().lower(): c_idx + 1 for c_idx, h in enumerate(trans_headers)}

    balance_col_idx = None
    for h_clean, c_idx in col_map.items():
      if 'bank closing balance' in h_clean or 'బ్యాలెన్స్' in h_clean:
        balance_col_idx = c_idx
        break

    standard_col_map = {
        h.strip(): c_idx + 1 for c_idx, h in enumerate(trans_headers)
    }

    previous_arrears = 0.0
    actual_paid = bank_closing_balance

    govt_share = min(due_amount, actual_paid)
    remaining_arrears = max(0.0, previous_arrears + due_amount - actual_paid)
    govt_rd_closing_bal = govt_share
    govt_status_val = 'MATCHED' if remaining_arrears <= 0 else 'PENDING'

    rd1_col_name = 'RD ACCOUNT_1 (Beneficiary Contribution)'
    arrears_b_name = '(B) పాత బాకీ(Previous Arrears)'
    due_c_name = '(C) ఈ నెల డ్యూ(Due)'
    actual_paid_d_name = '(D) Actual Paid'
    govt_share_name = 'Govt_Share'
    remaining_f_name = '(F) మిగిలిన బాకీ(Remaining Arrears)'
    govt_closing_h_name = 'Govt RD Closing Bal (H)'
    govt_status_i_name = 'Govt RD Status (I)'

    if trans_row_index:
      if balance_col_idx:
        trans_sheet.update_cell(
            trans_row_index, balance_col_idx, bank_closing_balance_str
        )

      updates = [
          (rd1_col_name, rd1),
          (arrears_b_name, previous_arrears),
          (due_c_name, due_amount),
          (actual_paid_d_name, actual_paid),
          (govt_share_name, govt_share),
          (remaining_f_name, remaining_arrears),
          (govt_closing_h_name, govt_rd_closing_bal),
          (govt_status_i_name, govt_status_val),
      ]
      for col_name, val in updates:
        if col_name in standard_col_map:
          trans_sheet.update_cell(
              trans_row_index, standard_col_map[col_name], val
          )
    else:
      row_data_dict = {h.strip(): '' for h in trans_headers}

      for h in trans_headers:
        h_clean = h.strip()
        h_upper = h_clean.upper()

        if h_clean in ['MONTH_NUMBER', 'month_number']:
          row_data_dict[h_clean] = current_month
        elif h_clean in ['BATCH_NUMBER', 'batch_number']:
          row_data_dict[h_clean] = batch_number
        elif h_clean in ['BANK_SNO', 'Bank SNO']:
          row_data_dict[h_clean] = bank_sno
        elif (
            'NAME OF BENIFICIARY' in h_upper
            or h_clean == 'Name Of Benificiary'
        ):
          row_data_dict[h_clean] = beneficiary_name
        elif (
            'BANK ACCOUNT NUMBER' in h_upper
            or h_clean == 'Bank Account Number'
        ):
          row_data_dict[h_clean] = beneficiary_acc
        elif h_clean == rd1_col_name:
          row_data_dict[h_clean] = rd1
        elif (
            'bank closing balance' in h_clean.lower()
            or 'బ్యాలెన్స్' in h_clean
        ):
          row_data_dict[h_clean] = bank_closing_balance_str
        elif h_clean == arrears_b_name:
          row_data_dict[h_clean] = previous_arrears
        elif h_clean == due_c_name:
          row_data_dict[h_clean] = due_amount
        elif h_clean == actual_paid_d_name:
          row_data_dict[h_clean] = actual_paid
        elif h_clean == govt_share_name:
          row_data_dict[h_clean] = govt_share
        elif h_clean == remaining_f_name:
          row_data_dict[h_clean] = remaining_arrears
        elif h_clean == govt_closing_h_name:
          row_data_dict[h_clean] = govt_rd_closing_bal
        elif h_clean == govt_status_i_name:
          row_data_dict[h_clean] = govt_status_val
        elif h_clean == 'FILE NAME':
          row_data_dict[h_clean] = bank_user

      new_row_data = [row_data_dict.get(h.strip(), '') for h in trans_headers]
      trans_sheet.append_row(new_row_data)

    return jsonify({
        'status': 'success',
        'message': 'డేటా మరియు లెక్కలు విజయవంతంగా అప్‌డేట్ చేయబడ్డాయి!',
    })

  except Exception as e:
    return jsonify(
        {'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}
    )


@bank_bp.route('/api/download-excel', methods=['GET'])
def download_excel():
  from app import SPREADSHEET_NAME, format_as_excel_text, get_sheets_client

  bank_user = session.get('bank_user')
  if not bank_user:
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  try:
    current_month = str(
        request.args.get('month', session.get('current_month', '1'))
    ).strip()

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
    master_records = target_sheet.get_all_records()
    df_master = pd.DataFrame(master_records)

    # --- బ్యాంకుకు పంపని (NOT SENT) రికార్డులను తొలగించుట ---
    if 'BANK SENT' in df_master.columns:
      df_master = df_master[
          df_master['BANK SENT'].astype(str).str.strip().str.upper()
          != 'NOT SENT'
      ].copy()

    bank_user_clean = bank_user.strip().lower()
    df_master = df_master[
        df_master['FILE NAME'].astype(str).str.strip().str.lower()
        == bank_user_clean
    ].copy()

    try:
      trans_sheet = spreadsheet.worksheet('Transactions_Year1')
      trans_records = trans_sheet.get_all_records()
      df_trans = pd.DataFrame(trans_records)
      if not df_trans.empty:
        df_trans = df_trans[
            df_trans['FILE NAME'].astype(str).str.strip().str.lower()
            == bank_user_clean
        ].copy()
    except Exception:
      df_trans = pd.DataFrame()

    month_trans_records = []
    batch_wise_data = {}
    all_month_data = []

    if not df_trans.empty:
      for _, t in df_trans.iterrows():
        t_sno = str(t.get('BANK_SNO') or t.get('Bank SNO') or '').strip()
        t_month = str(
            t.get('MONTH_NUMBER') or t.get('month_number') or ''
        ).strip()
        t_batch = str(
            t.get('BATCH_NUMBER') or t.get('batch_number') or ''
        ).strip()

        if t_month == current_month:
          month_trans_records.append((t_sno, t_batch, t))

    sno_to_batch = {}
    sno_to_balance = {}
    for t_sno, t_batch, t in month_trans_records:
      t_balance = (
          t.get(
              'Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న'
              ' బ్యాలెన్స్)'
          )
          or t.get('Bank Closing Balance')
          or ''
      )
      if t_sno:
        sno_to_batch[t_sno] = t_batch
        sno_to_balance[t_sno] = t_balance

    for idx, (_, row) in enumerate(df_master.iterrows()):
      file_name = str(
          row.get('FILE NAME') or row.get('File Name') or ''
      ).strip()
      bank_sno = str(row.get('BANK_SNO') or row.get('Bank SNO') or '').strip()

      raw_batch = sno_to_batch.get(bank_sno, '')
      assigned_batch = raw_batch if raw_batch != '' else 'Unassigned'

      closing_balance = sno_to_balance.get(
          bank_sno, row.get('Bank Closing Balance', '')
      )

      acc_num = format_as_excel_text(
          row.get('Bank Account Number') or row.get('ACCOUNT')
      )
      rd_1 = format_as_excel_text(row.get('RD 1') or row.get('RD_1'))
      rd_2 = format_as_excel_text(row.get('RD 2') or row.get('RD_2'))

      record_item = {
          'BANK_SNO': bank_sno or idx + 1,
          'MONTH_NUMBER': current_month,
          'BATCH_NUMBER': assigned_batch,
          'Name Of Benificiary': (
              row.get('Name Of Benificiary') or row.get('NAME') or ''
          ),
          'Contact Number': (
              row.get('Contact Number') or row.get('Contact_Number') or ''
          ),
          'Bank Account Number': acc_num,
          'RD ACCOUNT_1': rd_1,
          'RD ACCOUNT_2': rd_2,
          'Bank Closing Balance': closing_balance,
          'FILE NAME': file_name,
          'NAME OF CF': (row.get('NAME OF CF') or row.get('NAME_OF_CF') or ''),
      }

      all_month_data.append(record_item)

      if assigned_batch not in batch_wise_data:
        batch_wise_data[assigned_batch] = []
      batch_wise_data[assigned_batch].append(record_item)

    if not all_month_data:
      return jsonify({'status': 'error', 'message': 'డేటా ఏదీ అందుబాటులో లేదు.'})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
      sorted_batches = sorted(
          batch_wise_data.keys(), key=lambda x: (x == 'Unassigned', x)
      )

      for batch_key in sorted_batches:
        df_batch = pd.DataFrame(batch_wise_data[batch_key])
        sheet_name = (
            f'Batch_{batch_key}'
            if batch_key != 'Unassigned'
            else 'Unassigned_Batch'
        )
        df_batch.to_excel(writer, index=False, sheet_name=sheet_name)

      df_all = pd.DataFrame(all_month_data)
      df_all.to_excel(writer, index=False, sheet_name='All_Data')

    output.seek(0)
    filename = f"{bank_user.replace(' ', '_')}_Month_{current_month}_All_Batches_With_CF_NAME.xlsx"

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )
  except Exception as e:
    return jsonify(
        {'status': 'error', 'message': f'డౌన్‌లోడ్ చేయడంలో లోపం: {str(e)}'}
    )


@bank_bp.route('/api/bulk-upload', methods=['POST'])
def bulk_upload():
  try:
    from app import SPREADSHEET_NAME, get_sheets_client

    bank_user = session.get('bank_user')
    if not bank_user:
      return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

    if 'file' not in request.files:
      return jsonify({'status': 'error', 'message': 'ఫైల్ ఏదీ ఎంచుకోలేదు.'})

    file = request.files['file']
    if file.filename == '':
      return jsonify({'status': 'error', 'message': 'ఫైల్ పేరు ఖాళీగా ఉంది.'})

    if file.filename.endswith('.csv'):
      df = pd.read_csv(file)
    else:
      df = pd.read_excel(file)

    current_month = str(
        request.form.get('month', session.get('current_month', '1'))
    ).strip()
    batch_number = str(
        request.form.get('batch', session.get('batch_number', '1'))
    ).strip()

    def clean_sno(val):
      if val is None:
        return ''
      val_str = str(val).strip()
      if not val_str or val_str.lower() == 'nan':
        return ''
      if val_str.endswith('.0'):
        val_str = val_str[:-2]
      return val_str.strip()

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
    master_records = target_sheet.get_all_records()
    master_headers = target_sheet.row_values(1)
    master_col_map = {
        h.strip().upper(): c_idx + 1 for c_idx, h in enumerate(master_headers)
    }

    trans_sheet = spreadsheet.worksheet('Transactions_Year1')
    trans_records = trans_sheet.get_all_records()
    trans_headers = trans_sheet.row_values(1)
    trans_col_map = {
        h.strip(): c_idx + 1 for c_idx, h in enumerate(trans_headers)
    }

    bank_user_clean = bank_user.strip().lower()

    already_covered_snos = set()
    for t_row in trans_records:
      t_sno = clean_sno(t_row.get('BANK_SNO') or t_row.get('Bank SNO'))
      t_month = str(
          t_row.get('MONTH_NUMBER') or t_row.get('month_number') or ''
      ).strip()
      t_batch = str(
          t_row.get('BATCH_NUMBER') or t_row.get('batch_number') or ''
      ).strip()
      t_filename = (
          str(t_row.get('FILE NAME') or t_row.get('File Name') or '')
          .strip()
          .lower()
      )

      if (
          t_month == current_month
          and t_batch != batch_number
          and t_filename == bank_user_clean
          and t_sno
      ):
        try:
          if float(t_row.get('Govt_Share') or t_row.get('GOVT_SHARE') or 0) > 0:
            already_covered_snos.add(t_sno)
        except ValueError:
          pass

    updated_count = 0
    skipped_beneficiaries = []

    master_cell_updates = []
    trans_cell_updates = []
    new_trans_rows = []

    master_row_dict = {}
    for idx, m_row in enumerate(master_records):
      m_sno = clean_sno(m_row.get('BANK_SNO') or m_row.get('Bank SNO'))
      m_file = (
          str(m_row.get('FILE NAME') or m_row.get('File Name') or '')
          .strip()
          .lower()
      )

      if m_sno and m_file == bank_user_clean:
        try:
          wages_val = float(
              m_row.get('8% Of Monthly Wages')
              or m_row.get('MONTHLY WAGES')
              or 0
          )
        except (ValueError, TypeError):
          wages_val = 0.0

        master_row_dict[m_sno] = {
            'row_idx': idx + 2,
            'name': str(
                m_row.get('Name Of Benificiary') or m_row.get('NAME') or ''
            ).strip(),
            'acc': str(
                m_row.get('Bank Account Number') or m_row.get('ACCOUNT') or ''
            ).strip(),
            'wages': wages_val,
        }

    trans_row_dict = {}
    for idx, t_row in enumerate(trans_records):
      t_sno = clean_sno(t_row.get('BANK_SNO') or t_row.get('Bank SNO'))
      t_month = str(
          t_row.get('MONTH_NUMBER') or t_row.get('month_number') or ''
      ).strip()
      t_batch = str(
          t_row.get('BATCH_NUMBER') or t_row.get('batch_number') or ''
      ).strip()
      t_filename = (
          str(t_row.get('FILE NAME') or t_row.get('File Name') or '')
          .strip()
          .lower()
      )

      if (
          t_sno
          and t_month == current_month
          and t_batch == batch_number
          and t_filename == bank_user_clean
      ):
        trans_row_dict[t_sno] = idx + 2

    rd1_col = master_col_map.get('RD 1') or master_col_map.get('RD_1')
    rd2_col = master_col_map.get('RD 2') or master_col_map.get('RD_2')

    df.columns = [str(c).strip() for c in df.columns]

    def find_col(possible_keys):
      for col in df.columns:
        if col.upper() in [k.upper() for k in possible_keys]:
          return col
      return None

    sno_col_key = find_col(['BANK_SNO', 'SNO', 'Bank SNO'])
    rd1_col_key = find_col(['RD ACCOUNT_1', 'RD 1', 'RD_1', 'RD Account 1'])
    rd2_col_key = find_col(
        ['RD ACCOUNT_2', 'RD 2', 'RD_2', 'RD Account 2', 'RD-2']
    )
    bal_col_key = find_col([
        'Bank Closing Balance',
        'Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)',
        'Closing Balance',
    ])

    for _, row in df.iterrows():
      bank_sno = clean_sno(row[sno_col_key]) if sno_col_key else ''
      rd1 = (
          str(row[rd1_col_key]).strip().strip("'")
          if rd1_col_key and pd.notnull(row[rd1_col_key])
          else ''
      )
      rd2 = (
          str(row[rd2_col_key]).strip().strip("'")
          if rd2_col_key and pd.notnull(row[rd2_col_key])
          else ''
      )
      raw_balance = row[bal_col_key] if bal_col_key else ''

      if not bank_sno or bank_sno not in master_row_dict:
        continue

      if bank_sno in already_covered_snos:
        skipped_beneficiaries.append(bank_sno)
        continue

      try:
        if (
            pd.isna(raw_balance)
            or raw_balance is None
            or str(raw_balance).strip() == ''
            or str(raw_balance).strip().lower() == 'nan'
        ):
          bank_closing_balance = 0.0
          balance_str = '0.0'
        else:
          bank_closing_balance = float(raw_balance)
          if pd.isna(bank_closing_balance) or math.isinf(bank_closing_balance):
            bank_closing_balance = 0.0
          balance_str = str(bank_closing_balance)
      except (ValueError, TypeError):
        bank_closing_balance = 0.0
        balance_str = '0.0'

      m_info = master_row_dict[bank_sno]
      master_row_idx = m_info['row_idx']
      monthly_wages = m_info['wages']
      beneficiary_name = m_info['name']
      beneficiary_acc = m_info['acc']

      if rd1_col:
        master_cell_updates.append(
            gspread.Cell(master_row_idx, rd1_col, f"'{rd1}" if rd1 else '')
        )
      if rd2_col:
        master_cell_updates.append(
            gspread.Cell(master_row_idx, rd2_col, f"'{rd2}" if rd2 else '')
        )

      due_amount = monthly_wages
      previous_arrears = 0.0
      actual_paid = bank_closing_balance
      govt_share = min(due_amount, actual_paid)
      remaining_arrears = max(0.0, previous_arrears + due_amount - actual_paid)
      govt_rd_closing_bal = govt_share
      govt_rd_status = 'MATCHED' if remaining_arrears <= 0 else 'PENDING'

      rd1_col_name = 'RD ACCOUNT_1 (Beneficiary Contribution)'
      balance_col_name = (
          'Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)'
      )
      due_c_name = '(C) ఈ నెల డ్యూ(Due)'
      actual_paid_d_name = '(D) Actual Paid'
      govt_share_name = 'Govt_Share'
      remaining_f_name = '(F) మిగిలిన బాకీ(Remaining Arrears)'
      govt_closing_h_name = 'Govt RD Closing Bal (H)'
      govt_status_i_name = 'Govt RD Status (I)'

      if bank_sno in trans_row_dict:
        t_row_idx = trans_row_dict[bank_sno]
        t_updates_map = {
            rd1_col_name: f"'{rd1}" if rd1 else '',
            balance_col_name: balance_str,
            due_c_name: due_amount,
            actual_paid_d_name: actual_paid,
            govt_share_name: govt_share,
            remaining_f_name: remaining_arrears,
            govt_closing_h_name: govt_rd_closing_bal,
            govt_status_i_name: govt_rd_status,
        }
        for c_name, val in t_updates_map.items():
          if c_name in trans_col_map:
            trans_cell_updates.append(
                gspread.Cell(t_row_idx, trans_col_map[c_name], val)
            )
      else:
        row_data_dict = {h.strip(): '' for h in trans_headers}
        for h in trans_headers:
          h_clean = h.strip()
          if h_clean in ['MONTH_NUMBER', 'month_number']:
            row_data_dict[h_clean] = current_month
          elif h_clean in ['BATCH_NUMBER', 'batch_number']:
            row_data_dict[h_clean] = batch_number
          elif h_clean in ['BANK_SNO', 'Bank SNO']:
            row_data_dict[h_clean] = bank_sno
          elif h_clean == 'Name Of Benificiary':
            row_data_dict[h_clean] = beneficiary_name
          elif h_clean == 'Bank Account Number':
            row_data_dict[h_clean] = beneficiary_acc
          elif h_clean == rd1_col_name:
            row_data_dict[h_clean] = f"'{rd1}" if rd1 else ''
          elif h_clean == balance_col_name:
            row_data_dict[h_clean] = balance_str
          elif h_clean == due_c_name:
            row_data_dict[h_clean] = due_amount
          elif h_clean == actual_paid_d_name:
            row_data_dict[h_clean] = actual_paid
          elif h_clean == govt_share_name:
            row_data_dict[h_clean] = govt_share
          elif h_clean == remaining_f_name:
            row_data_dict[h_clean] = remaining_arrears
          elif h_clean == govt_closing_h_name:
            row_data_dict[h_clean] = govt_rd_closing_bal
          elif h_clean == govt_status_i_name:
            row_data_dict[h_clean] = govt_rd_status
          elif h_clean == 'FILE NAME':
            row_data_dict[h_clean] = bank_user

        new_trans_rows.append(
            [row_data_dict.get(h.strip(), '') for h in trans_headers]
        )

      updated_count += 1

    chunk_size = 500
    if master_cell_updates:
      for i in range(0, len(master_cell_updates), chunk_size):
        target_sheet.update_cells(master_cell_updates[i : i + chunk_size])

    if trans_cell_updates:
      for i in range(0, len(trans_cell_updates), chunk_size):
        trans_sheet.update_cells(trans_cell_updates[i : i + chunk_size])

    if new_trans_rows:
      trans_sheet.append_rows(new_trans_rows)

    return jsonify({
        'status': 'success',
        'updated_count': updated_count,
        'skipped_count': len(skipped_beneficiaries),
        'skipped_snos': skipped_beneficiaries,
        'message': (
            f'మొత్తం {updated_count} లబ్ధిదారుల వివరాలు విజయవంతంగా అప్‌డేట్'
            ' చేయబడ్డాయి! 🎉'
        ),
    })

  except Exception as e:
    return (
        jsonify(
            {'status': 'error', 'message': f'బల్క్ అప్‌డేట్‌లో లోపం: {str(e)}'}
        ),
        500,
    )


@bank_bp.route('/api/logout', methods=['POST'])
def api_logout():
  session.pop('bank_user', None)
  session.pop('current_month', None)
  session.pop('batch_number', None)
  return jsonify({'status': 'success', 'message': 'లాగౌట్ అయ్యారు.'})
