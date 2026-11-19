import os
from datetime import timedelta
import json
import pandas as pd
import gspread
from flask import Flask
from flask_cors import CORS
from google.oauth2 import service_account
from google.auth import default

from admin_routes import admin_bp
from cf_routes import cf_bp
from mismatch_routes import mismatch_bp
from reports_routes import reports_bp
from rd_home import rd_bp
from sheets_routes import sheets_bp, month_map, cf_sheet_ids

# కొత్తగా విభజించిన మూడు బ్లూప్రింట్స్
from bank_routes import bank_bp
from draft_routes import draft_bp
from excel_tools_routes import excel_tools_bp

app = Flask(__name__)
app.secret_key = 'thrift_scheme_secret_key'
app.permanent_session_lifetime = timedelta(days=7)
CORS(app)

SPREADSHEET_NAME = 'NOT_BANK_SENT_FORM'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SENDER_EMAIL = 'cf.thrift.sircillats@gmail.com'
APP_PASSWORD = 'hkcuhdvqrdquvubn'

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

PROCESS_STEPS_CONFIG = {
    '1': {
        'step_name': '1. లబ్ధిదారుల ప్రాథమిక జాబితా తయారీ',
        'template': 'MASTER_LIST_TEMPLATE.docx',
        'sheet_name': 'Master_Beneficiary_List',
        'subject_prefix': 'Submission of Verified Beneficiary Master List for TSPWTFS & SS',
        'body_desc': 'forwarding the verified Master Beneficiary List under Telangana State Powerloom Weavers Thrift Fund Scheme (TSPWTFS & SS)'
    },
    '2': {
        'step_name': '2. కొత్త RD ఖాతాల ప్రారంభం కోసం అభ్యర్థన',
        'template': 'NEW_ACCOUNT_REQUEST_TEMPLATE.docx',
        'sheet_name': 'New_Account_Opening_List',
        'subject_prefix': 'Request for Opening New Thrift RD Accounts (RD-1 & RD-2)',
        'body_desc': 'requesting you to open new Thrift Recurring Deposit (RD-1 & RD-2) accounts for the enclosed eligible weavers'
    },
    '3': {
        'step_name': '3. బ్యాంకు నుండి కొత్త RD ఖాతా నంబర్ల స్వీకరణ',
        'template': 'ACCOUNT_MAPPING_TEMPLATE.docx',
        'sheet_name': 'Account_Mapping_Status',
        'subject_prefix': 'Follow-up / Receipt of Newly Opened Thrift RD Account Numbers',
        'body_desc': 'regarding the status and mapping of newly opened Thrift RD account numbers'
    },
    '4': {
        'step_name': '4. డేటా మ్యాపింగ్ మరియు మాస్టర్ షీట్ అప్‌డేట్',
        'template': 'MAPPING_UPDATE_TEMPLATE.docx',
        'sheet_name': 'Updated_Master_Mapping',
        'subject_prefix': 'Updated Master Database & Account Mapping Confirmation',
        'body_desc': 'sharing the updated Master Database after mapping the RD account numbers'
    },
    '5': {
        'step_name': '5. కార్మికుడి వాటా (8%) బ్యాంక్ స్టేట్‌మెంట్ల సేకరణ',
        'template': 'SCROLL_COLLECTION_TEMPLATE.docx',
        'sheet_name': 'Monthly_8Percent_Scroll_Req',
        'subject_prefix': 'Request for Monthly 8% Beneficiary Contribution Bank Statement / Scroll',
        'body_desc': 'requesting the monthly bank scroll / transaction statement showing the 8% weaver share deposits into RD-1 accounts'
    },
    '6': {
        'step_name': '6. రీకన్సిలియేషన్ మరియు ప్రభుత్వ వాటా (8%) గణన',
        'template': 'RECONCILIATION_TEMPLATE.docx',
        'sheet_name': 'Reconciliation_Statement',
        'subject_prefix': 'Monthly Reconciliation & 8% Govt Matching Share Calculation Statement',
        'body_desc': 'forwarding the reconciled statement of 8% weaver contributions and corresponding 8% Government matching share calculation'
    },
    '7': {
        'step_name': '7. ప్రభుత్వ వాటా నిధుల మంజూరు మరియు బదిలీ',
        'template': 'FUND_SANCTION_TEMPLATE.docx',
        'sheet_name': 'Fund_Sanction_Details',
        'subject_prefix': 'Intimation of Sanction & Fund Transfer for Govt Matching Share',
        'body_desc': 'intimating the sanction and transfer of Government matching share funds'
    },
    '8': {
        'step_name': '8. లబ్ధిదారుల వారీగా కేటాయింపు జాబితా పంపడం (Credit Advice)',
        'template': 'RD2_CREDIT_SANCTION_TEMPLATE.docx',
        'sheet_name': 'Eligible_Credit_List',
        'subject_prefix': 'Official Letter & Beneficiary Excel List for TSPWTFS & SS RD2 Credit',
        'body_desc': 'crediting the Government matching share into RD2 Accounts'
    },
    '9': {
        'step_name': '9. ప్రభుత్వ వాటా జమ మరియు కన్ఫర్మేషన్',
        'template': 'CREDIT_CONFIRMATION_TEMPLATE.docx',
        'sheet_name': 'Credit_Confirmation_Req',
        'subject_prefix': 'Request for Credit Confirmation Statement & Failed Transactions Report',
        'body_desc': 'requesting the Credit Confirmation Statement (with UTR/Transaction IDs) and list of any failed/returned credits'
    },
    '10': {
        'step_name': '10. లోపాల సవరణ మరియు ఉత్తర ప్రత్యుత్తరాలు (Rectification)',
        'template': 'RECTIFICATION_TEMPLATE.docx',
        'sheet_name': 'Rectification_List',
        'subject_prefix': 'Rectification of Bank Account Errors, Dormant Accounts & IFSC Updates',
        'body_desc': 'requesting rectification of account discrepancies, dormant accounts, or IFSC updates as per the attached list'
    },
    '11': {
        'step_name': '11. ఎగ్జిక్యూటివ్ సమ్మరీ మరియు ఆడిట్ రిపోర్ట్స్',
        'template': 'AUDIT_SUMMARY_TEMPLATE.docx',
        'sheet_name': 'Executive_Audit_Summary',
        'subject_prefix': 'Periodic Executive Summary & Audit Reconciliation Report',
        'body_desc': 'submitting the consolidated Executive Summary and Audit Reconciliation Report'
    },
    '12': {
        'step_name': '12. మెచ్యూరిటీ, NOC జారీ మరియు ఖాతా క్లోజర్',
        'template': 'MATURITY_NOC_TEMPLATE.docx',
        'sheet_name': 'Maturity_NOC_List',
        'subject_prefix': 'Issuance of NOC for Maturity & Closure of Thrift RD Accounts',
        'body_desc': 'issuing No Objection Certificate (NOC) for maturity closure and release of funds to the beneficiaries savings accounts'
    }
}


def format_as_excel_text(val):
  if val is None or (isinstance(val, float) and pd.isna(val)):
    return ''
  val_str = str(val).strip().strip("'")
  if val_str.endswith('.0'):
    val_str = val_str[:-2]
  if val_str:
    return f"'{val_str}"
  return ''


def get_sheets_client():
  scopes = [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
  ]
  credentials_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
  if credentials_json_str:
    info = json.loads(credentials_json_str)
    creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
  else:
    cred_path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
    if cred_path and os.path.exists(cred_path):
      creds = service_account.Credentials.from_service_account_file(cred_path, scopes=scopes)
    else:
      creds, _ = default(scopes=scopes)
  return gspread.authorize(creds)


def get_bank_list_from_sheet():
  try:
    client = get_sheets_client()
    sh = client.open(SPREADSHEET_NAME)
    ws_email_map = sh.worksheet('Bank_Email_Ids')
    df_email_map = pd.DataFrame(ws_email_map.get_all_records())
    banks = df_email_map['FILE_NAME'].dropna().astype(str).str.strip().unique().tolist()
    return [b for b in banks if b and b != '0']
  except Exception as e:
    print(f'Error fetching bank email list: {e}')
    return []


# Blueprints Register (app డిఫైన్ అయిన తర్వాతే రిజిస్టర్ చేయాలి)
app.register_blueprint(admin_bp)
app.register_blueprint(cf_bp)
app.register_blueprint(mismatch_bp)
app.register_blueprint(rd_bp)
app.register_blueprint(reports_bp)
app.register_blueprint(sheets_bp)
app.register_blueprint(bank_bp)
app.register_blueprint(draft_bp)
app.register_blueprint(excel_tools_bp)


if __name__ == '__main__':
  port = int(os.environ.get('PORT', 8080))
  app.run(host='0.0.0.0', port=port)
