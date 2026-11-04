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
import openpyxl
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
    current_app,
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


# ✅ ఇదివరకు సగ్గా పనిచేసిన (ఫైల్ 7) సేఫ్ ఆథెంటికేషన్ పద్ధతి (ఎర్రర్స్ రావు)
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


# ==========================================
# GMAIL DRAFT GENERATOR ROUTES
# ==========================================

# ==========================================
# GMAIL DRAFT GENERATOR ROUTES (THREADED)
# ==========================================

@app.route('/draft-generator')
def draft_generator_page():
  banks = get_bank_list_from_sheet()
  return render_template('draft_generator.html', banks=banks)


@app.route('/api/get-bank-doc-details', methods=['GET'])
def get_bank_doc_details():
  """ఫ్రంట్ ఎండ్ నుండి బ్యాంక్ పేరు బట్టి PREPARE_DOC_RD_TRANSACTIONS నుండి డిఫాల్ట్ విలువలు పంపడానికి"""
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


def background_draft_task(data, app_context):
  """బ్యాక్‌గ్రౌండ్‌లో హెవీ ప్రాసెసింగ్ (Excel, Word, IMAP) చేయడానికి థ్రెడ్ ఫంక్షన్"""
  with app_context:
    try:
      target_month = int(data.get('month', 1))
      target_batch = int(data.get('batch', 1))
      selected_bank = str(data.get('bank_name', '')).strip()

      user_cheque_no = str(data.get('cheque_no', '')).strip()
      user_cheque_date = str(data.get('cheque_date', '')).strip()

      if not selected_bank:
        print("Error: Bank name is missing in background task.")
        return

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
          suffixes=('', '_TRANS')
      )

      bank_month_df['Govt_Share'] = pd.to_numeric(bank_month_df['Govt_Share'], errors='coerce').fillna(0)
      bank_month_df['BATCH_NUMBER'] = (
          pd.to_numeric(bank_month_df['BATCH_NUMBER'], errors='coerce').fillna(0)
          if 'BATCH_NUMBER' in bank_month_df.columns
          else 0
      )

      if 'BANK_REMARKS' in bank_month_df.columns:
        bank_month_df['BANK_REMARKS_CLEAN'] = bank_month_df['BANK_REMARKS'].astype(str).str.strip().str.upper()
      else:
        bank_month_df['BANK_REMARKS_CLEAN'] = ''

      died_mask = bank_month_df['BANK_REMARKS_CLEAN'].str.contains('DIED', na=False)

      active_df = bank_month_df[
          (bank_month_df['Govt_Share'] > 0)
          & (bank_month_df['BATCH_NUMBER'] == target_batch)
          & (~died_mask)
      ]
      active_df = active_df.drop_duplicates(subset=['RD 1'], keep='first')

      pending_df = bank_month_df[
          (bank_month_df['Govt_Share'] == 0)
          | (bank_month_df['BATCH_NUMBER'] == 0)
          | (died_mask)
      ]

      record_count = len(active_df)
      pending_count = len(pending_df)
      total_govt_share = active_df['Govt_Share'].sum()
      cheque_amount = str(total_govt_share)

      bank_email = email_dict.get(selected_bank, '')
      if not bank_email or '@' not in bank_email:
        print(f'Error: {selected_bank} కి అధికారిక ఈమెయిల్ ఐడి లేదు!')
        return

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
          p_export_df['Status'] = pending_df['BANK_REMARKS_CLEAN'].apply(
              lambda x: 'Beneficiary Died (మరణించినవారు)' if 'DIED' in str(x) else 'RD-1 Account Pending / Unassigned Batch / Zero Transaction'
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
            '{{Current_Date}}': cheque_date,
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
                <p style="margin:0; font-weight: bold;">⚠️ Pending / Excluded Accounts Notice:</p>
                <p style="margin:5px 0 0 0;">There are <b>{pending_count} beneficiaries</b> whose accounts are pending, unassigned, or marked as expired/died.</p>
            </div>
            """

      cc_recipients_html = """
        <br>
        <p style="font-size: 13px; color: #555; line-height: 1.4;">
            <b>Copy submitted to / కాపీ సమర్పించడమైనది:</b><br>
            1. <b>The Lead District Manager (LDM), Union Bank of India, Sircilla</b> — తగు సమాచారము మరియు తగిన తదుపరి చర్య నిమిత్తము పంపడమైనది.<br>
            2. <b>Branch Manager, Union Bank of India (Nodal Branch)</b> — పాజిటివ్ పే సిస్టమ్ (PPS) పాటించుట కొరకు పంపడమైనది.<br>
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
        encoders.encode_base64(part2)
        part2.add_header('Content-Disposition', f'attachment; filename="{letter_filename}"')
        message.attach(part2)

      mail = imaplib.IMAP4_SSL('imap.gmail.com')
      mail.login(SENDER_EMAIL, APP_PASSWORD)
      mail.append(
          '"[Gmail]/Drafts"',
          '',
          imaplib.Time2Internaldate(time.time()),
          message.as_bytes(),
      )
      mail.logout()

      try:
        try:
          ws_log = sh.worksheet('Bank_Transactions')
        except:
          ws_log = sh.add_worksheet(
              title='Bank_Transactions', rows=100, cols=10
          )
          ws_log.append_row([
              'Timestamp',
              'Bank Name',
              'Month',
              'Batch',
              'Cheque No',
              'Cheque Date',
              'Amount',
              'Beneficiaries Count',
          ])

        ws_log.append_row([
            time.strftime('%Y-%m-%d %H:%M:%S'),
            selected_bank,
            f'{selected_month_name} {selected_year}',
            target_batch,
            cheque_no,
            cheque_date,
            total_govt_share,
            record_count,
        ])
      except Exception as log_err:
        print(f'Logging Error: {log_err}')

      print(f'Background Draft generated successfully for {selected_bank}!')

    except Exception as e:
      print(f'Background Draft Generation Error: {str(e)}')


@app.route('/api/generate-draft', methods=['POST'])
def generate_draft_api():
  import threading
  try:
    data = request.get_json() if request.is_json else request.form
    selected_bank = str(data.get('bank_name', '')).strip()

    if not selected_bank:
      return (
          jsonify(
              {'status': 'error', 'message': 'దయచేసి బ్యాంకును ఎంచుకోండి!'}
          ),
          400,
      )

    # Flask యాప్ కాంటెక్స్ట్ కోల్పోకుండా ఉండటానికి current_app._get_current_object() వాడాలి
    app_context = current_app._get_current_object().app_context() if hasattr(current_app, '_get_current_object') else app.app_context()

    # బ్యాక్‌గ్రౌండ్ థ్రెడ్‌ని స్టార్ట్ చేయడం (టైమ్‌అవుట్ సమస్య రాదు)
    thread = threading.Thread(target=background_draft_task, args=(data, app_context))
    thread.daemon = True
    thread.start()

    return jsonify({
        'status': 'success',
        'message': (
            'ఇమెయిల్ డ్రాఫ్ట్ బ్యాక్‌గ్రౌండ్‌లో తయారవుతోంది. దయచేసి కొన్ని క్షణాల తరువాత మీ జిమెయిల్ డ్రాఫ్ట్స్ చెక్ చేయండి!'
        ),
    })

  except Exception as e:
    return jsonify(
        {'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}
    ), 500

@app.route('/api/search', methods=['POST'])
def search_beneficiary():
  try:
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


@app.route('/api/login', methods=['POST'])
def bank_login():
  try:
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
    return jsonify({"status": "error", "message": f"లాగిన్ లోపం: {str(e)}"})


@app.route('/api/bank-data', methods=['GET'])
def get_bank_data():
  try:
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


@app.route('/api/download-csv', methods=['GET'])
def download_csv():
  bank_user = session.get('bank_user')
  if not bank_user:
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  client = get_sheets_client()
  spreadsheet = client.open(SPREADSHEET_NAME)
  target_sheet = spreadsheet.worksheet('NEW THRIFT DATA')
  records = target_sheet.get_all_records()

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


@app.route('/api/update-rd-accounts', methods=['POST'])
def update_rd_accounts():
  try:
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
    col_map = {}
    for c_idx, h in enumerate(trans_headers):
      h_clean = h.strip().lower()
      col_map[h_clean] = c_idx + 1

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


@app.route('/api/download-excel', methods=['GET'])
def download_excel():
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

    def format_as_excel_text(val):
      if val is None:
        return ''
      val_str = str(val).strip().strip("'")
      if val_str:
        return f"'{val_str}"
      return ''

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
          'Bank Account Number': acc_num,
          'RD ACCOUNT_1': rd_1,
          'RD ACCOUNT_2': rd_2,
          'Bank Closing Balance': closing_balance,
          'FILE NAME': file_name,
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
    filename = (
        f"{bank_user.replace(' ', '_')}_Month_{current_month}_All_Batches.xlsx"
    )

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


@app.route("/api/bulk-upload", methods=["POST"])
def bulk_upload():
  try:
    bank_user = session.get("bank_user")
    if not bank_user:
      return jsonify({"status": "error", "message": "దయచేసి లాగిన్ అవ్వండి!"})

    if "file" not in request.files:
      return jsonify({"status": "error", "message": "ఫైల్ ఏదీ ఎంచుకోలేదు."})

    file = request.files["file"]
    if file.filename == "":
      return jsonify({"status": "error", "message": "ఫైల్ పేరు ఖాళీగా ఉంది."})

    df = pd.read_excel(file)

    current_month = str(
        request.form.get("month", session.get("current_month", "1"))
    ).strip()
    batch_number = str(
        request.form.get("batch", session.get("batch_number", "1"))
    ).strip()

    # --- పర్ఫెక్ట్ మాచింగ్ కోసం SNO క్లీనింగ్ ఫంక్షన్ ---
    def clean_sno(val):
      if val is None:
        return ""
      val_str = str(val).strip()
      if not val_str or val_str.lower() == "nan":
        return ""
      if val_str.endswith(".0"):
        val_str = val_str[:-2]
      return val_str.strip()

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    target_sheet = spreadsheet.worksheet("NEW THRIFT DATA")
    master_records = target_sheet.get_all_records()
    master_headers = target_sheet.row_values(1)
    master_col_map = {
        h.strip().upper(): c_idx + 1 for c_idx, h in enumerate(master_headers)
    }

    trans_sheet = spreadsheet.worksheet("Transactions_Year1")
    trans_records = trans_sheet.get_all_records()
    trans_headers = trans_sheet.row_values(1)
    trans_col_map = {
        h.strip(): c_idx + 1 for c_idx, h in enumerate(trans_headers)
    }

    bank_user_clean = bank_user.strip().lower()

    # 1. Transactions నుండి already_covered_snos ని సరిగ్గా క్లీన్ చేసి సేకరించడం
    already_covered_snos = set()
    for t_row in trans_records:
      t_sno = clean_sno(t_row.get("BANK_SNO") or t_row.get("Bank SNO"))
      t_month = str(
          t_row.get("MONTH_NUMBER") or t_row.get("month_number") or ""
      ).strip()
      t_batch = str(
          t_row.get("BATCH_NUMBER") or t_row.get("batch_number") or ""
      ).strip()
      t_filename = (
          str(t_row.get("FILE NAME") or t_row.get("File Name") or "")
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
          if (
              float(
                  t_row.get("Govt_Share") or t_row.get("GOVT_SHARE") or 0
              )
              > 0
          ):
            already_covered_snos.add(t_sno)
        except ValueError:
          pass

    updated_count = 0
    skipped_beneficiaries = []

    master_cell_updates = []
    trans_cell_updates = []
    new_trans_rows = []

    # 2. Master Records ని క్లీన్డ్ SNO తో మ్యాప్ చేయడం
    master_row_dict = {}
    for idx, m_row in enumerate(master_records):
      m_sno = clean_sno(m_row.get("BANK_SNO") or m_row.get("Bank SNO"))
      m_file = (
          str(m_row.get("FILE NAME") or m_row.get("File Name") or "")
          .strip()
          .lower()
      )

      if m_sno and m_file == bank_user_clean:
        try:
          wages_val = float(
              m_row.get("8% Of Monthly Wages")
              or m_row.get("MONTHLY WAGES")
              or 0
          )
        except (ValueError, TypeError):
          wages_val = 0.0

        master_row_dict[m_sno] = {
            "row_idx": idx + 2,
            "name": str(
                m_row.get("Name Of Benificiary") or m_row.get("NAME") or ""
            ).strip(),
            "acc": str(
                m_row.get("Bank Account Number") or m_row.get("ACCOUNT") or ""
            ).strip(),
            "wages": wages_val,
        }

    # 3. Transactions Row Dict ని క్లీన్డ్ SNO తో మ్యాప్ చేయడం
    trans_row_dict = {}
    for idx, t_row in enumerate(trans_records):
      t_sno = clean_sno(t_row.get("BANK_SNO") or t_row.get("Bank SNO"))
      t_month = str(
          t_row.get("MONTH_NUMBER") or t_row.get("month_number") or ""
      ).strip()
      t_batch = str(
          t_row.get("BATCH_NUMBER") or t_row.get("batch_number") or ""
      ).strip()
      t_filename = (
          str(t_row.get("FILE NAME") or t_row.get("File Name") or "")
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

    rd1_col = master_col_map.get("RD 1") or master_col_map.get("RD_1")
    rd2_col = master_col_map.get("RD 2") or master_col_map.get("RD_2")

    # 4. ఎక్సెల్ డేటాఫ్రేమ్ (df) లోని SNO లను క్లీన్ చేసి చెక్ చేయడం
    for _, row in df.iterrows():
      bank_sno = clean_sno(row.get("BANK_SNO"))
      rd1 = str(row.get("RD ACCOUNT_1") or row.get("RD 1") or "").strip()
      rd2 = str(row.get("RD ACCOUNT_2") or row.get("RD 2") or "").strip()
      
      raw_balance = (
          row.get("Bank Closing Balance")
          or row.get("Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)")
          or ""
      )

      if not bank_sno or bank_sno not in master_row_dict:
        continue

      if bank_sno in already_covered_snos:
        skipped_beneficiaries.append(bank_sno)
        continue

      try:
        if (
            pd.isna(raw_balance)
            or raw_balance is None
            or str(raw_balance).strip() == ""
            or str(raw_balance).strip().lower() == "nan"
        ):
          bank_closing_balance = 0.0
          balance_str = "0.0"
        else:
          bank_closing_balance = float(raw_balance)
          if pd.isna(bank_closing_balance) or math.isinf(bank_closing_balance):
            bank_closing_balance = 0.0
          balance_str = str(bank_closing_balance)
      except (ValueError, TypeError):
        bank_closing_balance = 0.0
        balance_str = "0.0"

      m_info = master_row_dict[bank_sno]
      master_row_idx = m_info["row_idx"]
      monthly_wages = m_info["wages"]
      beneficiary_name = m_info["name"]
      beneficiary_acc = m_info["acc"]

      if rd1_col:
        master_cell_updates.append(gspread.Cell(master_row_idx, rd1_col, rd1))
      if rd2_col:
        master_cell_updates.append(gspread.Cell(master_row_idx, rd2_col, rd2))

      due_amount = monthly_wages
      previous_arrears = 0.0
      actual_paid = bank_closing_balance
      govt_share = min(due_amount, actual_paid)
      remaining_arrears = max(0.0, previous_arrears + due_amount - actual_paid)
      govt_rd_closing_bal = govt_share
      govt_rd_status = "MATCHED" if remaining_arrears <= 0 else "PENDING"

      rd1_col_name = "RD ACCOUNT_1 (Beneficiary Contribution)"
      balance_col_name = (
          "Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)"
      )
      due_c_name = "(C) ఈ నెల డ్యూ(Due)"
      actual_paid_d_name = "(D) Actual Paid"
      govt_share_name = "Govt_Share"
      remaining_f_name = "(F) మిగిలిన బాకీ(Remaining Arrears)"
      govt_closing_h_name = "Govt RD Closing Bal (H)"
      govt_status_i_name = "Govt RD Status (I)"

      if bank_sno in trans_row_dict:
        t_row_idx = trans_row_dict[bank_sno]
        t_updates_map = {
            rd1_col_name: rd1,
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
        row_data_dict = {h.strip(): "" for h in trans_headers}
        for h in trans_headers:
          h_clean = h.strip()
          if h_clean in ["MONTH_NUMBER", "month_number"]:
            row_data_dict[h_clean] = current_month
          elif h_clean in ["BATCH_NUMBER", "batch_number"]:
            row_data_dict[h_clean] = batch_number
          elif h_clean in ["BANK_SNO", "Bank SNO"]:
            row_data_dict[h_clean] = bank_sno
          elif h_clean == "Name Of Benificiary":
            row_data_dict[h_clean] = beneficiary_name
          elif h_clean == "Bank Account Number":
            row_data_dict[h_clean] = beneficiary_acc
          elif h_clean == rd1_col_name:
            row_data_dict[h_clean] = rd1
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
          elif h_clean == "FILE NAME":
            row_data_dict[h_clean] = bank_user

        new_trans_rows.append(
            [row_data_dict.get(h.strip(), "") for h in trans_headers]
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

    message = (
        f"మొత్తం {updated_count} లబ్ధిదారుల వివరాలు విజయవంతంగా అప్‌డేట్"
        " చేయబడ్డాయి! 🎉"
    )
    if skipped_beneficiaries:
      message += (
          " (గమనిక: SNOs"
          f" {', '.join(map(str, skipped_beneficiaries))} ఈ నెలలో ఇప్పటికే"
          " వేరే బ్యాచ్‌లో ఉన్నందున స్కిప్ చేయబడ్డాయి)."
      )

    return jsonify({"status": "success", "message": message})

  except Exception as e:
    return jsonify(
        {"status": "error", "message": f"బల్క్ అప్‌డేట్‌లో లోపం: {str(e)}"}
    )


def format_as_excel_text(val):
      if val is None:
        return ''
      val_str = str(val).strip().strip("'")
      if val_str:
        return f"'{val_str}"
      return ''


@app.route('/process', methods=['POST'])
def process_join():
    join_type = request.form.get('join_type')
    
    try:
        if join_type == 'single_file':
            if 'file_single' not in request.files:
                return "దయచేసి ఎక్సెల్ ఫైల్ అప్‌లోడ్ చేయండి.", 400
            file = request.files['file_single']
            sheet1 = request.form.get('sheet1', '').strip()
            col1 = request.form.get('col1', '').strip()
            sheet2 = request.form.get('sheet2', '').strip()
            col2 = request.form.get('col2', '').strip()
            
            df1 = pd.read_excel(file, sheet_name=sheet1 if sheet1 else 0)
            df2 = pd.read_excel(file, sheet_name=sheet2 if sheet2 else 1)
            
        elif join_type == 'two_files':
            if 'file1' not in request.files or 'file2' not in request.files:
                return "దయచేసి రెండు ఫైల్స్ అప్‌లోడ్ చేయండి.", 400
            file1 = request.files['file1']
            sheet1_f1 = request.form.get('sheet1_f1', '').strip()
            col1 = request.form.get('col1_f1', '').strip()
            
            file2 = request.files['file2']
            sheet2_f2 = request.form.get('sheet2_f2', '').strip()
            col2 = request.form.get('col2_f2', '').strip()
            
            df1 = pd.read_excel(file1, sheet_name=sheet1_f1 if sheet1_f1 else 0)
            df2 = pd.read_excel(file2, sheet_name=sheet2_f2 if sheet2_f2 else 0)
        else:
            return "సరైన జాయిన్ ఆప్షన్ ఎంచుకోండి.", 400

        # కాలమ్ పేర్లను ట్రిమ్ చేయడం
        df1.columns = [str(c).strip() for c in df1.columns]
        df2.columns = [str(c).strip() for c in df2.columns]
        
        if col1 not in df1.columns or col2 not in df2.columns:
            return f"లోపం: '{col1}' లేదా '{col2}' కాలమ్స్ ఎక్సెల్ ఫైల్‌లో కనుగొనబడలేదు!", 400
        
        # కీ ని క్లీన్ చేయడానికి ఫంక్షన్
        def clean_acc(val):
            if pd.isna(val):
                return ""
            val_str = str(val).strip()
            if val_str.endswith(".0"):
                val_str = val_str[:-2]
            return val_str.replace("'", "").strip()
        
        df1['Clean_Key'] = df1[col1].apply(clean_acc)
        df2['Clean_Key'] = df2[col2].apply(clean_acc)
        
        # Left Join చేయడం
        merged_df = pd.merge(df1, df2, on='Clean_Key', how='left', suffixes=('_left', '_right'))
        merged_df = merged_df.drop(columns=['Clean_Key'])
        
        output_file = "Merged_RD_Accounts_Output.xlsx"
        merged_df.to_excel(output_file, index=False)
        
        return send_file(output_file, as_attachment=True, download_name=output_file)
        
    except Exception as e:
        return f"ఫైల్ ప్రాసెసింగ్‌లో లోపం సంభవించింది: {str(e)}", 500

@app.route('/process-reconcile', methods=['POST'])
def process_reconcile():
    try:
        # ఫ్రంట్ ఎండ్ నుండి వచ్చే ఫైల్ మరియు ఇన్‌పుట్‌లను తీసుకోవడం
        file = request.files['reconcile_file']
        sheet_name = request.form.get('rec_sheet_name', '').strip()
        rd_col = request.form.get('rec_rd_col', 'Rd').strip()
        ben_col = request.form.get('rec_ben_col', 'Ben').strip()
        amt_col = request.form.get('rec_amt_col', 'Amount').strip()
        
        # ఎక్సెల్ ఫైల్ చదవడం
        df = pd.read_excel(file, sheet_name=sheet_name if sheet_name else 0)
        
        # కాలమ్ పేర్లను ట్రిమ్ చేయడం
        df.columns = [str(c).strip() for c in df.columns]
        
        # ఇచ్చిన కాలమ్స్ ఆధారంగా డేటాను క్లీన్ చేయడం
        df_clean = df[[rd_col, ben_col, amt_col]].dropna(subset=[rd_col, ben_col]).copy()
        df_clean['Ben_clean'] = df_clean[ben_col].astype(str).str.strip()
        
        # RD ఖాతా నంబర్‌ను టెక్స్ట్ (String) ఫార్మాట్‌లో మార్చడం
        def format_rd(val):
            try:
                return str(int(float(val)))
            except:
                return str(val).strip()
                
        df_clean['Rd_formatted'] = df_clean[rd_col].apply(format_rd)
        
        # ఒకే లబ్ధిదారుని పేరుతో ఉన్న అన్ని అకౌంట్లను అమర్చడం
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
            row_dict = {"S.No": idx, "Name Of Beneficiary": ben}
            
            rds = ben_rows['Rd_formatted'].tolist()
            amounts = ben_rows[amt_col].tolist()
            
            for i in range(len(rds)):
                acc_num = i + 1
                row_dict[f"RD Account {acc_num}"] = f"'{rds[i]}"
                row_dict[f"Closing Balance {acc_num}"] = amounts[i] if pd.notnull(amounts[i]) else 0.0
                
            rows_data.append(row_dict)
            
        df_result = pd.DataFrame(rows_data)
        
        # కాలమ్స్ ఆర్డర్ సర్దుబాటు చేయడం
        all_cols = ["S.No", "Name Of Beneficiary"]
        for i in range(1, max_rds + 1):
            all_cols.append(f"RD Account {i}")
            all_cols.append(f"Closing Balance {i}")
            
        df_result = df_result.reindex(columns=all_cols)
        
        # openpyxl ద్వారా ప్రొఫెషనల్ డిజైన్‌తో ఎక్సెల్ ఫైల్ సృష్టించడం
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reconciled RD Accounts"
        
        # Title Header
        max_col_letter = get_column_letter(len(all_cols))
        ws.merge_cells(f"A1:{max_col_letter}1")
        title_cell = ws["A1"]
        title_cell.value = "BANK - RECURRING DEPOSIT ACCOUNTS STATEMENT"
        title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32
        
        # Table Headers
        headers = list(df_result.columns)
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num, value=header)
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
        ws.row_dimensions[3].height = 25
        
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
        
        # డేటా రాయడం
        for row_idx, row in df_result.iterrows():
            row_num = row_idx + 4
            current_row_data = row.tolist()
            
            for col_num, val in enumerate(current_row_data, 1):
                cell = ws.cell(row=row_num, column=col_num)
                cell.value = "" if (pd.isna(val) or val is None) else val
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                
                if row_idx % 2 == 1:
                    cell.fill = zebra_fill
                    
                header_name = headers[col_num - 1]
                if header_name == "S.No":
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif header_name == "Name Of Beneficiary":
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif "RD Account" in header_name:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.number_format = '@'
                elif "Closing Balance" in header_name:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00'
                        
            ws.row_dimensions[row_num].height = 20
            
        # Column Widths
        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 32
        
        output_filename = "RD_Accounts_Reconciled_Output.xlsx"
        wb.save(output_filename)
        
        return send_file(output_filename, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        return f"Error processing reconciliation file: {str(e)}", 400

@app.route('/api/logout', methods=['POST'])
def api_logout():
  session.pop('bank_user', None)
  session.pop('current_month', None)
  session.pop('batch_number', None)
  return jsonify({'status': 'success', 'message': 'లాగౌట్ అయ్యారు.'})


if __name__ == '__main__':
  port = int(os.environ.get('PORT', 8080))
  app.run(host='0.0.0.0', port=port)
