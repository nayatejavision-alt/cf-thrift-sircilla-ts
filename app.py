import io
import csv
from datetime import timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import imaplib
import json
import os
import time
import math
from admin_routes import admin_bp
from cf_routes import cf_bp
from flask import (
    Flask,
    jsonify,
    make_response,
    render_template,
    request,
    send_file,
    session,
)
from google.oauth2 import service_account
from flask_cors import CORS
from google.auth import default
from google.oauth2.service_account import Credentials
import gspread
from num2words import num2words
from mismatch_routes import mismatch_bp
from reports_routes import reports_bp
import pandas as pd
from rd_home import rd_bp
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'thrift_scheme_secret_key'
app.permanent_session_lifetime = timedelta(days=7)
CORS(app)

SPREADSHEET_NAME = 'NOT_BANK_SENT_FORM'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Gmail Credentials
SENDER_EMAIL = 'cf.thrift.sircillats@gmail.com'
APP_PASSWORD = 'hkcuhdvqrdquvubn'

# Month Mapping
MONTH_MAPPING = {
    1: ('June', '2026'),
    2: ('July', '2026'),
    3: ('August', '2026'),
    4: ('September', '2026'),
    5: ('October', '2026'),
    6: ('November', '2026'),
    7: ('December', '2026'),
    8: ('January', '2027'),
    9: ('February', '2027'),
    10: ('March', '2027'),
    11: ('April', '2027'),
    12: ('May', '2027'),
}

app.register_blueprint(admin_bp)
app.register_blueprint(cf_bp)
app.register_blueprint(mismatch_bp)
app.register_blueprint(rd_bp)
app.register_blueprint(reports_bp)


def get_sheets_client():
  scopes = [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
  ]
  credentials_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')

  if credentials_json_str:
    info = json.loads(credentials_json_str)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=scopes
    )
  else:
    cred_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
    if cred_path and os.path.exists(cred_path):
      creds = service_account.Credentials.from_service_account_file(
          cred_path, scopes=scopes
      )
    else:
      creds, _ = default(scopes=scopes)

  return gspread.authorize(creds)


def format_as_excel_text(val):
  if val is None:
    return ''
  val_str = str(val).strip().strip("'")
  if val_str:
    return f"'{val_str}"
  return ''


def get_bank_list_from_sheet():
  try:
    client = get_sheets_client()
    sh = client.open(SPREADSHEET_NAME)
    ws_email_map = sh.worksheet('Bank_Email_Ids')
    df_email_map = pd.DataFrame(ws_email_map.get_all_records())
    banks = (
        df_email_map['FILE NAME'].dropna().astype(str).str.strip().unique().tolist()
    )
    return [b for b in banks if b and b != '0']
  except Exception as e:
    print(f'Error fetching bank email list: {e}')
    return []


@app.route('/')
def home():
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

  return render_template('index.html', banks=bank_list)


@app.route('/draft-generator')
def draft_generator_page():
  banks = get_bank_list_from_sheet()
  return render_template('draft_generator.html', banks=banks)


@app.route('/api/get-bank-doc-details', methods=['GET'])
def get_bank_doc_details():
  try:
    bank_name = request.args.get('bank_name', '').strip()
    client = get_sheets_client()
    sh = client.open(SPREADSHEET_NAME)
    ws_doc = sh.worksheet('PREPARE_DOC_RD_TRANSACTIONS')
    records = ws_doc.get_all_records()

    default_data = {
        'Cheque_No': '',
        'Cheque_Date': '',
    }
    for row in records:
      if str(row.get('FILE_NAME', '')).strip().lower() == bank_name.lower():
        default_data = {
            'Cheque_No': row.get('Cheque_No', ''),
            'Cheque_Date': row.get('Cheque_Date', ''),
        }
        break
    return jsonify({'status': 'success', 'data': default_data})
  except Exception as e:
    return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/generate-draft', methods=['POST'])
def generate_draft_api():
  try:
    data = request.get_json() if request.is_json else request.form

    target_month = int(data.get('month', 1))
    target_batch = int(data.get('batch', 1))
    selected_bank = str(data.get('bank_name', '')).strip()

    user_cheque_no = str(data.get('cheque_no', '')).strip()
    user_cheque_date = str(data.get('cheque_date', '')).strip()

    if not selected_bank:
      return (
          jsonify(
              {'status': 'error', 'message': 'దయచేసి బ్యాంకును ఎంచుకోండి!'}
          ),
          400,
      )

    selected_month_name, selected_year = MONTH_MAPPING.get(
        target_month, ('June', '2026')
    )

    client = get_sheets_client()
    sh = client.open(SPREADSHEET_NAME)

    ws_trans = sh.worksheet('Transactions_Year1')
    df_trans = pd.DataFrame(ws_trans.get_all_records())

    ws_thrift = sh.worksheet('NEW THRIFT DATA')
    df_thrift = pd.DataFrame(ws_thrift.get_all_records())

    ws_email_map = sh.worksheet('Bank_Email_Ids')
    df_email_map = pd.DataFrame(ws_email_map.get_all_records())
    email_dict = dict(
        zip(
            df_email_map['FILE NAME'].astype(str).str.strip(),
            df_email_map['OFFICIAL_EMAIL'].astype(str).str.strip(),
        )
    )

    try:
      ws_doc_sheet = sh.worksheet('PREPARE_DOC_RD_TRANSACTIONS')
      doc_records = ws_doc_sheet.get_all_records()
    except Exception:
      doc_records = []

    bank_doc_info = {}
    for r in doc_records:
      if str(r.get('FILE_NAME', '')).strip().lower() == selected_bank.lower():
        bank_doc_info = r
        break

    cheque_no = user_cheque_no or str(bank_doc_info.get('Cheque_No', ''))
    cheque_date = user_cheque_date or str(bank_doc_info.get('Cheque_Date', ''))

    df_trans['Govt_Share'] = pd.to_numeric(
        df_trans['Govt_Share'], errors='coerce'
    ).fillna(0)
    df_trans['MONTH_NUMBER'] = pd.to_numeric(
        df_trans['MONTH_NUMBER'], errors='coerce'
    )
    df_trans['BATCH_NUMBER'] = (
        pd.to_numeric(df_trans['BATCH_NUMBER'], errors='coerce').fillna(0)
        if 'BATCH_NUMBER' in df_trans.columns
        else 0
    )

    bank_thrift_df = df_thrift[
        df_thrift['FILE NAME'].astype(str).str.strip() == selected_bank
    ].copy()

    bank_month_df = pd.merge(
        bank_thrift_df,
        df_trans[df_trans['MONTH_NUMBER'] == target_month],
        left_on='RD 1',
        right_on='RD ACCOUNT_1 (Beneficiary Contribution)',
        how='left',
        suffixes=('', '_TRANS'),
    )

    bank_month_df['Govt_Share'] = pd.to_numeric(
        bank_month_df['Govt_Share'], errors='coerce'
    ).fillna(0)
    bank_month_df['BATCH_NUMBER'] = (
        pd.to_numeric(bank_month_df['BATCH_NUMBER'], errors='coerce').fillna(0)
        if 'BATCH_NUMBER' in bank_month_df.columns
        else 0
    )

    active_df = bank_month_df[
        (bank_month_df['Govt_Share'] > 0)
        & (bank_month_df['BATCH_NUMBER'] == target_batch)
    ]
    active_df = active_df.drop_duplicates(subset=['RD 1'], keep='first')

    pending_df = bank_month_df[
        (bank_month_df['Govt_Share'] == 0)
        | (bank_month_df['BATCH_NUMBER'] == 0)
    ]

    record_count = len(active_df)
    pending_count = len(pending_df)
    total_govt_share = active_df['Govt_Share'].sum()

    cheque_amount = str(total_govt_share)

    bank_email = email_dict.get(selected_bank, '')
    if not bank_email or '@' not in bank_email:
      return jsonify({
          'status': 'error',
          'message': f'{selected_bank} కి అధికారిక ఈమెయిల్ ఐడి లేదు!',
      })

    target_folder_path = (
        f'/tmp/Bank_Wise_Excel_Files_Month_{target_month}_Batch_{target_batch}'
    )
    if not os.path.exists(target_folder_path):
      os.makedirs(target_folder_path)

    local_excel_path = os.path.join(
        target_folder_path,
        f'{selected_bank}_Month_{target_month}_Batch_{target_batch}.xlsx',
    )

    with pd.ExcelWriter(local_excel_path, engine='openpyxl') as writer:
      export_df = pd.DataFrame()
      export_df['BENEFICIARY SNO'] = active_df.get(
          'BANK_SNO', active_df.get('SNO', '')
      )
      export_df['Name Of Benificiary'] = active_df['Name Of Benificiary']
      export_df['Bank Account Number'] = active_df[
          'Bank Account Number'
      ].apply(format_as_excel_text)
      export_df['RD ACCOUNT_1 (Beneficiary Contribution)'] = active_df[
          'RD ACCOUNT_1 (Beneficiary Contribution)'
      ].apply(format_as_excel_text)
      export_df['RD-1 (Beneficiary) Contribution'] = active_df['Govt_Share']

      if 'RD 2' in active_df.columns:
        export_df['RD ACCOUNT_2 (Govt Contribution)'] = active_df['RD 2'].apply(
            format_as_excel_text
        )
      elif 'RD ACCOUNT_2' in active_df.columns:
        export_df['RD ACCOUNT_2 (Govt Contribution)'] = active_df[
            'RD ACCOUNT_2'
        ].apply(format_as_excel_text)
      elif 'RD 2_THRIFT' in active_df.columns:
        export_df['RD ACCOUNT_2 (Govt Contribution)'] = active_df[
            'RD 2_THRIFT'
        ].apply(format_as_excel_text)
      else:
        export_df['RD ACCOUNT_2 (Govt Contribution)'] = ''

      export_df['RD-2 (Govt_Share) Contribution'] = active_df['Govt_Share']
      export_df['Contact Number'] = (
          active_df['Contact Number'].apply(format_as_excel_text)
          if 'Contact Number' in active_df.columns
          else ''
      )
      export_df.to_excel(writer, sheet_name='Eligible_Credit_List', index=False)

      if pending_count > 0:
        p_export_df = pd.DataFrame()
        p_export_df['SNO'] = pending_df.get(
            'BANK_SNO', pending_df.get('SNO', '')
        )
        p_export_df['Name Of Benificiary'] = pending_df['Name Of Benificiary']
        p_export_df['Bank Account Number'] = pending_df[
            'Bank Account Number'
        ].apply(format_as_excel_text)
        p_export_df['RD ACCOUNT_1'] = pending_df[
            'RD ACCOUNT_1 (Beneficiary Contribution)'
        ].apply(format_as_excel_text)
        p_export_df['Contact Number'] = (
            pending_df['Contact Number'].apply(format_as_excel_text)
            if 'Contact Number' in pending_df.columns
            else ''
        )
        p_export_df['Status'] = (
            'RD-1 Account Pending / Unassigned Batch / Zero Transaction'
        )
        p_export_df.to_excel(
            writer, sheet_name='Pending_RD1_List', index=False
        )

    letter_path = None
    doc_template_path = 'RD2_CREDIT_SANCTION_TEMPLATE.docx'
    if os.path.exists(doc_template_path):
      import docx

      doc = docx.Document(doc_template_path)
      try:
        amt_words = (
            num2words(float(cheque_amount or total_govt_share), lang='en_IN')
            .title()
            + ' Only'
        )
      except:
        amt_words = f'Rupees {cheque_amount or total_govt_share} Only'

      replacements = {
          '{{Bank_Name}}': str(
              bank_doc_info.get('Bank_Name', selected_bank)
          ),
          '{{Location}}': str(bank_doc_info.get('Location', 'Sircilla')),
          '{{Current_Date}}': time.strftime('%d/%m/%Y'),
          '{{Total_Count_Of_Records}}': str(
              bank_doc_info.get('Total_Count_Of_Records', len(bank_month_df))
          ),
          '{{Cheque_No}}': cheque_no,
          '{{Cheque_Date}}': cheque_date,
          '{{Cheque_Amount}}': cheque_amount or str(total_govt_share),
          '{{Amount_In_Words}}': amt_words,
          '{{Rd_Total_Count_for_month_number}}': str(record_count),
          '{{RD_Start_from_for_month_number}}': '1',
          '{{RD_End_With_for_month_number}}': str(record_count),
          '{{This_Office_Previous_Dates}}': str(
              bank_doc_info.get(
                  'This_Office_Previous_Dates',
                  '04/07/2026 , 15/07/2026',
              )
          ),
      }

      for p in doc.paragraphs:
        for k, v in replacements.items():
          if k in p.text:
            p.text = p.text.replace(k, str(v))

      for table in doc.tables:
        for row in table.rows:
          for cell in row.cells:
            for k, v in replacements.items():
              if k in cell.text:
                cell.text = cell.text.replace(k, str(v))

      letter_path = os.path.join(
          target_folder_path, f'Official_Letter_{selected_bank}.docx'
      )
      doc.save(letter_path)

    subject = f'Official Letter & Beneficiary Excel List for TSPWTFS & SS RD2 Credit - {selected_month_name} {selected_year} (Batch-{target_batch}) - Rc.No. 176/2025-RSCL - {selected_bank}'

    pending_notice_html = ''
    if pending_count > 0:
      pending_notice_html = f"""
            <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 12px; margin-top: 15px; border-radius: 4px; color: #856404;">
                <p style="margin:0; font-weight: bold;">⚠️ Pending / Unassigned Accounts Notice:</p>
                <p style="margin:5px 0 0 0;">There are <b>{pending_count} beneficiaries</b> whose accounts are either pending, unassigned to this batch, or have zero transactions.</p>
            </div>
            """

    cc_recipients_html = """
        <br>
        <p style="font-size: 13px; color: #555; line-height: 1.4;">
            <b>Copy submitted to / కాపీ సమర్పించడమైనది:</b><br>
            1. <b>The Lead District Manager (LDM), Union Bank of India, Sircilla</b> — తగు సమాచారము మరియు తగిన తదుపరి చర్య నిమిత్తము పంపడమైనది.<br>
            2. <b>Branch Manager, Union Bank of India (Nodal Branch)</b> — ఆర్‌బిఐ నిబంధనల ప్రకారం పాజిటివ్ పే సిస్టమ్ (PPS) పాటించుట కొరకు పంపడమైనది.<br>
            3. <b>Assistant Director, Handlooms & Textiles, Rajanna Sircilla</b> — కార్యాలయ రికార్డులు మరియు పర్యవేక్షణ నిమిత్తము.
        </p>
        """

    OFFICIAL_SIGNATURE_HTML = """
        <br><br>
        <hr style="border: 0; border-top: 1px solid #ccc;">
        <p style="font-family: Arial, sans-serif; font-size: 13px; color: #333;">
            <b>ఇట్లు / Regards,</b><br>
            <b>Office of the Assistant Director</b><br>
            Handlooms & Textiles, Rajanna Sircilla District.<br>
            Government of Telangana.<br>
            📧 Email: cf.thrift.sircillats@gmail.com | 📞 Phone: +91-9700389438
        </p>
        """

    body_html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.5;">
            <p>Dear <b>{selected_bank} Branch Manager</b>,</p>
            <p>Please find attached the official letter along with the Beneficiary Excel List for crediting the Government 8% matching share into RD2 Accounts for <b>{selected_month_name} {selected_year} (Batch-{target_batch})</b>.</p>
            <p><b>Summary Details:</b> Active: {record_count} | Amount: ₹{total_govt_share:,.2f} | Cheque No: {cheque_no}</p>
            {pending_notice_html}
            <p>Kindly process the credits strictly adhering to the <b>RBI Positive Pay System (PPS)</b> guidelines.</p>
            <p style="color: #444; font-weight: bold; margin-top: 15px;">ఈ విషయము తగు సమాచారము మరియు తగిన తదుపరి చర్య నిమిత్తము పంపడమైనది.</p>
            {cc_recipients_html}
            {OFFICIAL_SIGNATURE_HTML}
        </body>
        </html>
        """

    message = MIMEMultipart()
    message['To'] = bank_email
    message['Cc'] = ', '.join([
        'ldmsircilla@unionbankofindia.bank.in',
        'ubin0827444@unionbankofindia.bank.in',
        'rajannahandtex@gmail.com',
    ])
    message['From'] = SENDER_EMAIL
    message['Subject'] = subject
    message.attach(MIMEText(body_html, 'html'))

    if os.path.exists(local_excel_path):
      excel_filename = os.path.basename(local_excel_path)
      with open(local_excel_path, 'rb') as f:
        part1 = MIMEBase('application', 'octet-stream')
        part1.set_payload(f.read())
      encoders.encode_base64(part1)
      part1.add_header('Content-Disposition', f'attachment; filename="{excel_filename}"')
      message.attach(part1)

    if letter_path and os.path.exists(letter_path):
      letter_filename = os.path.basename(letter_path)
      with open(letter_path, 'rb') as f:
        part2 = MIMEBase('application', 'octet-stream')
        part2.set_payload(f.read())
