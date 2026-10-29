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
        #'Cheque_Amount': '',
    }
    for row in records:
      if str(row.get('FILE_NAME', '')).strip().lower() == bank_name.lower():
        default_data = {
            'Cheque_No': row.get('Cheque_No', ''),
            'Cheque_Date': row.get('Cheque_Date', ''),
            #'Cheque_Amount': row.get('Cheque_Amount', ''),
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

    # ఫ్రంట్ ఎండ్ నుండి ఎడిట్ చేసి పంపిన చెక్ వివరాలు (లేకపోతే ఎంప్టీ)
    user_cheque_no = str(data.get('cheque_no', '')).strip()
    user_cheque_date = str(data.get('cheque_date', '')).strip()
    # user_cheque_amount = str(data.get('cheque_amount', '')).strip()

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

    # PREPARE_DOC_RD_TRANSACTIONS నుండి డిఫాల్ట్ డాక్యుమెంట్ వివరాలు తెచ్చుకోవడం
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

    # యూజర్ ఫ్రంట్ ఎండ్ నుండి పంపితే అవే వాడాలి,లేకపోతే షీట్ డిఫాల్ట్ వాల్యూస్ తీసుకోవాలి
    cheque_no = user_cheque_no or str(bank_doc_info.get('Cheque_No', ''))
    cheque_date = user_cheque_date or str(bank_doc_info.get('Cheque_Date', ''))
    #cheque_amount = user_cheque_amount or str(
    #    bank_doc_info.get('Cheque_Amount', '')
    # )

    df_thrift = df_thrift.drop_duplicates(subset=['RD 1'], keep='first')
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

    df_merged_full = pd.merge(
        df_trans[df_trans['MONTH_NUMBER'] == target_month],
        df_thrift,
        left_on='RD ACCOUNT_1 (Beneficiary Contribution)',
        right_on='RD 1',
        how='inner',
        suffixes=('', '_THRIFT'),
    )

    # 1. ముందుగా ఆ బ్యాంకుకి చెందిన మొత్తం త్రిఫ్ట్ బెనెఫిషరీలను తీసుకోవాలి
    bank_thrift_df = df_thrift[
        df_thrift['FILE NAME'].astype(str).str.strip() == selected_bank
    ].copy()

    # డూప్లికేట్స్ తొలగించడం
    bank_thrift_df = bank_thrift_df.drop_duplicates(subset=['RD 1'], keep='first')

    # 2. ఆ నెల ట్రాన్సాక్షన్ డేటాతో 'left' merge చేయడం (దీనివల్ల అసలు ఎంట్రీ లేనివారు కూడా మిస్ కాకుండా వస్తారు)
    bank_month_df = pd.merge(
        bank_thrift_df,
        df_trans[df_trans['MONTH_NUMBER'] == target_month],
        left_on='RD 1',
        right_on='RD ACCOUNT_1 (Beneficiary Contribution)',
        how='left',
        suffixes=('', '_TRANS')
    )

    # 3. నంబర్ ఫార్మాటింగ్ క్లీన్ చేయడం
    bank_month_df['Govt_Share'] = pd.to_numeric(bank_month_df['Govt_Share'], errors='coerce').fillna(0)
    bank_month_df['BATCH_NUMBER'] = (
        pd.to_numeric(bank_month_df['BATCH_NUMBER'], errors='coerce').fillna(0)
        if 'BATCH_NUMBER' in bank_month_df.columns
        else 0
    )

    # 4. Active DF (ప్రస్తుత బ్యాచ్‌లో సక్సెస్‌ఫుల్‌గా ఉన్నవి)
    active_df = bank_month_df[
        (bank_month_df['Govt_Share'] > 0)
        & (bank_month_df['BATCH_NUMBER'] == target_batch)
    ]

    # 5. Pending DF (మీరు కోరుకున్న లాజిక్: టోటల్ త్రిఫ్ట్ నుండి ఆక్టివ్ మరియు ఇతర బ్యాచ్‌లను మినహాయించడం)
    pending_df = bank_month_df[
        (bank_month_df['Govt_Share'] == 0)
        | (bank_month_df['BATCH_NUMBER'] == 0)
        | (
            (bank_month_df['BATCH_NUMBER'] != target_batch)
            & (bank_month_df['Govt_Share'] > 0)
        )
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

    # 1. ఎక్సెల్ ఫైల్ తయారీ
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

    # 2. ఆఫీస్ లెటర్ డాక్యుమెంట్ తయారీ (Word Doc)
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

    # మెయిల్ బాడీ తయారీ
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
            
            1. <b>The Lead District Manager (LDM), Union Bank of India, Sircilla</b> — తగు సమాచారము మరియు తగిన తదుపరి చర్య నిమిత్తము పంపడమైనది (For information and necessary action).<br>
            2. <b>Branch Manager, Union Bank of India (Nodal Branch - Positive Pay System compliance)</b> — తగు సమాచారము నిమిత్తము మరియు ఆర్‌బిఐ నిబంధనల ప్రకారం పాజిటివ్ పే సిస్టమ్ (PPS) పాటించుట కొరకు పంపడమైనది.<br>
            3. <b>Assistant Director, Handlooms & Textiles, Rajanna Sircilla</b> — కార్యాలయ రికార్డులు మరియు పర్యవేక్షణ నిమిత్తము (For information & record).
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

# 1. ఎక్సెల్ ఫైల్ అటాచ్‌మెంట్ (ఎక్స్‌టెన్షన్‌తో సహా)
    if os.path.exists(local_excel_path):
      excel_filename = os.path.basename(local_excel_path)  # ఉదా: INDIAN_BANK_Month_1_Batch_1.xlsx
      with open(local_excel_path, 'rb') as f:
        part1 = MIMEBase('application', 'octet-stream')
        part1.set_payload(f.read())
      encoders.encode_base64(part1)
      # ఇక్కడ డబుల్ కోట్స్ కచ్చితంగా ఉండాలి
      part1.add_header('Content-Disposition', f'attachment; filename="{excel_filename}"')
      message.attach(part1)

    # 2. ఆఫీస్ లెటర్ డాక్యుమెంట్ అటాచ్‌మెంట్ (ఎక్స్‌టెన్షన్‌తో సహా - .docx)
    if letter_path and os.path.exists(letter_path):
      letter_filename = os.path.basename(letter_path)  # ఉదా: Official_Letter_INDIAN_BANK.docx
      with open(letter_path, 'rb') as f:
        part2 = MIMEBase('application', 'octet-stream')
        part2.set_payload(f.read())
      encoders.encode_base64(part2)
      # ఇక్కడ కూడా డబుల్ కోట్స్ కచ్చితంగా ఉండాలి
      part2.add_header('Content-Disposition', f'attachment; filename="{letter_filename}"')
      message.attach(part2)
    # జిమెయిల్ డ్రాఫ్ట్‌లో సేవ్ చేయడం
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(SENDER_EMAIL, APP_PASSWORD)
    mail.append(
        '"[Gmail]/Drafts"',
        '',
        imaplib.Time2Internaldate(time.time()),
        message.as_bytes(),
    )
    mail.logout()

    # 3. ట్రాన్సాక్షన్ సక్సెస్ అయిన తర్వాత 'Bank_Transactions' షీట్‌లో లాగ్ చేయడం
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

    return jsonify({
        'status': 'success',
        'message': f'{selected_bank} కోసం అధికారిక లేఖ మరియు ఎక్సెల్ ఫైల్‌తో జిమెయిల్ డ్రాఫ్ట్ విజయవంతంగా తయారైంది!',
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


@app.route('/api/bulk-upload', methods=['POST'])
def bulk_upload():
  try:
    bank_user = session.get('bank_user')
    if not bank_user:
      return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

    if 'file' not in request.files:
      return jsonify({'status': 'error', 'message': 'ఫైల్ ఏదీ ఎంచుకోలేదు.'})

    file = request.files['file']
    if file.filename == '':
      return jsonify({'status': 'error', 'message': 'ఫైల్ పేరు ఖాళీగా ఉంది.'})

    df = pd.read_excel(file)

    current_month = str(
        request.form.get('month', session.get('current_month', '1'))
    ).strip()
    batch_number = str(
        request.form.get('batch', session.get('batch_number', '1'))
    ).strip()

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
      t_sno = str(
          t_row.get('BANK_SNO') or t_row.get('Bank SNO') or ''
      ).strip()
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
          if (
              float(t_row.get('Govt_Share') or t_row.get('GOVT_SHARE') or 0) > 0
          ):
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
      m_sno = str(m_row.get('BANK_SNO') or '').strip()
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
      t_sno = str(t_row.get('BANK_SNO') or t_row.get('Bank SNO') or '').strip()
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

    for _, row in df.iterrows():
      bank_sno = str(row.get('BANK_SNO') or '').strip()
      rd1 = str(row.get('RD ACCOUNT_1') or '').strip()
      rd2 = str(row.get('RD ACCOUNT_2') or '').strip()
      balance_str = str(row.get('Bank Closing Balance') or '').strip()

      if not bank_sno or bank_sno not in master_row_dict:
        continue

      if bank_sno in already_covered_snos:
        skipped_beneficiaries.append(bank_sno)
        continue

      try:
        bank_closing_balance = float(balance_str) if balance_str else 0.0
      except ValueError:
        bank_closing_balance = 0.0

      m_info = master_row_dict[bank_sno]
      master_row_idx = m_info['row_idx']
      monthly_wages = m_info['wages']
      beneficiary_name = m_info['name']
      beneficiary_acc = m_info['acc']

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
      govt_rd_status = 'MATCHED' if remaining_arrears <= 0 else 'PENDING'

      rd1_col_name = 'RD ACCOUNT_1 (Beneficiary Contribution)'
      balance_col_name = (
          'Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న బ్యాలెన్స్)'
      )
      arrears_b_name = '(B) పాత బాకీ(Previous Arrears)'
      due_c_name = '(C) ఈ నెల డ్యూ(Due)'
      actual_paid_d_name = '(D) Actual Paid'
      govt_share_name = 'Govt_Share'
      remaining_f_name = '(F) మిగిలిన బాకీ(Remaining Arrears)'
      govt_closing_h_name = 'Govt RD Closing Bal (H)'
      govt_status_i_name = 'Govt RD Status (I)'

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
          elif h_clean == 'FILE NAME':
            row_data_dict[h_clean] = bank_user

        new_trans_rows.append(
            [row_data_dict.get(h.strip(), '') for h in trans_headers]
        )

      updated_count += 1

    if master_cell_updates:
      target_sheet.update_cells(master_cell_updates)
    if trans_cell_updates:
      trans_sheet.update_cells(trans_cell_updates)
    if new_trans_rows:
      trans_sheet.append_rows(new_trans_rows)

    message = f'మొత్తం {updated_count} లబ్ధిదారుల వివరాలు విజయవంతంగా అప్‌డేట్ చేయబడ్డాయి! 🎉'
    if skipped_beneficiaries:
      message += f" (గమనిక: SNOs {', '.join(skipped_beneficiaries)} ఈ నెలలో ఇప్పటికే వేరే బ్యాచ్‌లో ఉన్నందున స్కిప్ చేయబడ్డాయి)."

    return jsonify({'status': 'success', 'message': message})

  except Exception as e:
    return jsonify(
        {'status': 'error', 'message': f'బల్క్ అప్‌డేట్‌లో లోపం: {str(e)}'}
    )

def format_as_excel_text(val):
      if val is None:
        return ''
      val_str = str(val).strip().strip("'")
      if val_str:
        return f"'{val_str}"
      return ''


@app.route('/api/logout', methods=['POST'])
def api_logout():
  session.pop('bank_user', None)
  session.pop('current_month', None)
  session.pop('batch_number', None)
  return jsonify({'status': 'success', 'message': 'లాగౌట్ అయ్యారు.'})


if __name__ == '__main__':
  port = int(os.environ.get('PORT', 8080))
  app.run(host='0.0.0.0', port=port)
