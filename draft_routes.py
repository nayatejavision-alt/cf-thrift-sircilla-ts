import imaplib
import os
import threading
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from flask import Blueprint, current_app, jsonify, render_template, request
from num2words import num2words
import pandas as pd

draft_bp = Blueprint('draft_bp', __name__)


@draft_bp.route('/draft-generator')
def draft_generator_page():
  from app import get_bank_list_from_sheet

  banks = get_bank_list_from_sheet()
  return render_template('draft_generator.html', banks=banks)


@draft_bp.route('/api/get-bank-doc-details', methods=['GET'])
def get_bank_doc_details():
  """ఫ్రంట్ ఎండ్ నుండి బ్యాంక్ పేరు బట్టి PREPARE_DOC_RD_TRANSACTIONS నుండి డిఫాల్ట్ విలువలు పంపడానికి"""
  try:
    from app import get_sheets_client, SPREADSHEET_NAME

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
  """బ్యాక్‌గ్రౌండ్‌లో 1 నుండి 12 దశలకు అనుగుణంగా (Excel, Word, IMAP) ప్రాసెసింగ్"""
  from app import (
      APP_PASSWORD,
      MONTH_MAPPING,
      PROCESS_STEPS_CONFIG,
      SENDER_EMAIL,
      SPREADSHEET_NAME,
      get_sheets_client,
  )

  def format_text(val):
    if pd.isna(val) or val is None:
      return ''
    s = str(val).strip()
    if s.endswith('.0'):
      s = s[:-2]
    return s

  def build_beneficiary_export_df(df_in, include_status_col=False):
    if df_in is None or df_in.empty:
      cols = [
          'SNO',
          'Name Of Benificiary',
          'Bank Account Number',
          'RD-1 Account No',
          'RD-2 Account No',
          'Contact Number',
          'Nominee Name',
          'Relation',
          'Nominee DOB',
          'BANK SENT',
          'Remarks',
      ]
      if include_status_col:
        cols.append('Status')
      return pd.DataFrame(columns=cols)

    out = pd.DataFrame()
    out['SNO'] = df_in.get('BANK_SNO', df_in.get('SNO', ''))
    out['Name Of Benificiary'] = df_in.get('Name Of Benificiary', '')
    out['Bank Account Number'] = df_in.get('Bank Account Number', '').apply(
        format_text
    )
    out['RD-1 Account No'] = df_in.get('RD 1', '').apply(format_text)
    out['RD-2 Account No'] = df_in.get('RD 2', '').apply(format_text)
    out['Contact Number'] = (
        df_in.get('Contact Number', '').apply(format_text)
        if 'Contact Number' in df_in.columns
        else ''
    )
    out['Nominee Name'] = df_in.get('Nominee Name', '').apply(format_text)
    out['Relation'] = df_in.get('Relation', '').apply(format_text)
    out['Nominee DOB'] = df_in.get('Nominee DOB', '').apply(format_text)
    out['BANK SENT'] = df_in.get('BANK SENT', '')
    out['Remarks'] = df_in.get('BANK_REMARKS', '')
    if include_status_col:
      out['Status'] = df_in.get('Status', '')
    return out

  with app_context:
    try:
      process_step = str(data.get('process_step', '8')).strip()
      target_month = int(data.get('month', 1))
      target_batch = int(data.get('batch', 1))
      selected_bank = str(data.get('bank_name', '')).strip()

      user_cheque_no = str(data.get('cheque_no', '')).strip()
      user_cheque_date = str(data.get('cheque_date', '')).strip()
      user_cheque_amount = str(data.get('cheque_amount', '')).strip()

      if not selected_bank:
        print('Error: Bank name is missing in background task.')
        return

      step_cfg = PROCESS_STEPS_CONFIG.get(
          process_step, PROCESS_STEPS_CONFIG['8']
      )
      selected_month_name, selected_year = MONTH_MAPPING.get(
          target_month, ('June', '2026')
      )

      client = get_sheets_client()
      sh = client.open(SPREADSHEET_NAME)

      ws_trans = sh.worksheet('Transactions_Year1')
      df_trans = pd.DataFrame(ws_trans.get_all_records())

      ws_thrift = sh.worksheet('NEW THRIFT DATA')
      df_thrift_all = pd.DataFrame(ws_thrift.get_all_records())

      # BANK SENT ఆధారంగా NOT SENT వారిని మరియు ఇప్పటికే పంపినవారిని గుర్తించడం
      if 'BANK SENT' in df_thrift_all.columns:
        not_sent_mask_all = (
            df_thrift_all['BANK SENT'].astype(str).str.strip().str.upper()
            == 'NOT SENT'
        )
      else:
        not_sent_mask_all = pd.Series([False] * len(df_thrift_all))

      df_thrift_not_sent_all = df_thrift_all[not_sent_mask_all].copy()
      df_thrift_sent_all = df_thrift_all[~not_sent_mask_all].copy()

      ws_email_map = sh.worksheet('Bank_Email_Ids')
      df_email_map = pd.DataFrame(ws_email_map.get_all_records())
      email_dict = dict(
          zip(
              df_email_map['FILE_NAME'].astype(str).str.strip(),
              df_email_map['OFFICIAL_EMAIL'].astype(str).str.strip(),
          )
      )

      bank_email = email_dict.get(selected_bank, '')
      if not bank_email or '@' not in bank_email:
        print(f'Error: {selected_bank} కి అధికారిక ఈమెయిల్ ఐడి లేదు!')
        return

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

      # బ్యాంకు వారీగా Sent మరియు Not Sent డేటా వేరు చేయడం
      bank_sent_df = df_thrift_sent_all[
          df_thrift_sent_all['FILE NAME'].astype(str).str.strip()
          == selected_bank
      ].copy()
      bank_not_sent_df = df_thrift_not_sent_all[
          df_thrift_not_sent_all['FILE NAME'].astype(str).str.strip()
          == selected_bank
      ].copy()

      for sub_df in [bank_sent_df, bank_not_sent_df]:
        if 'BANK_REMARKS' in sub_df.columns:
          sub_df['BANK_REMARKS_CLEAN'] = (
              sub_df['BANK_REMARKS'].astype(str).str.strip().str.upper()
          )
        else:
          sub_df['BANK_REMARKS_CLEAN'] = ''

      target_folder_path = f'/tmp/Step_{process_step}_Bank_Files_Month_{target_month}_Batch_{target_batch}'
      os.makedirs(target_folder_path, exist_ok=True)

      local_excel_path = os.path.join(
          target_folder_path,
          f'{selected_bank}_Step{process_step}_{step_cfg["sheet_name"]}_M{target_month}_B{target_batch}.xlsx',
      )

      record_count = 0
      pending_count = 0
      total_govt_share = 0.0

      # ========================================================
      # దశల వారీగా (Step 1 to 12) Excel File తయారీ లాజిక్
      # ========================================================
      with pd.ExcelWriter(local_excel_path, engine='openpyxl') as writer:
        # --- దశ 1 & దశ 2 (ప్రారంభ దశలు: NOT SENT వారు Current_Eligible_List లో, Sent వారు Already_Sent_To_Bank షీట్‌లో) ---
        if process_step in ['1', '2']:
          if process_step == '2' and not bank_not_sent_df.empty and 'RD 1' in bank_not_sent_df.columns:
            rd1_col = bank_not_sent_df['RD 1'].astype(str).str.strip()
            pending_rd_df = bank_not_sent_df[
                (rd1_col == '') | (rd1_col == '0') | (rd1_col.str.lower() == 'nan')
            ].copy()
            current_eligible_df = (
                pending_rd_df if len(pending_rd_df) > 0 else bank_not_sent_df.copy()
            )
          else:
            current_eligible_df = bank_not_sent_df.copy()

          record_count = len(current_eligible_df)

          df_eligible_out = build_beneficiary_export_df(current_eligible_df)
          df_eligible_out.to_excel(
              writer, sheet_name='Current_Eligible_List', index=False
          )

          df_already_sent_out = build_beneficiary_export_df(bank_sent_df)
          df_already_sent_out.to_excel(
              writer, sheet_name='Already_Sent_To_Bank', index=False
          )

        # --- దశ 6, 7, 8, 9 (నెలవారీ లావాదేవీలు / Govt Share / Credit Advice) ---
        elif process_step in ['6', '7', '8', '9']:
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

          bank_month_df = pd.merge(
              bank_sent_df,
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
              pd.to_numeric(bank_month_df['BATCH_NUMBER'], errors='coerce')
              .fillna(0)
              if 'BATCH_NUMBER' in bank_month_df.columns
              else 0
          )

          if 'BANK_REMARKS' in bank_month_df.columns:
            bank_month_df['BANK_REMARKS_CLEAN'] = (
                bank_month_df['BANK_REMARKS'].astype(str).str.strip().str.upper()
            )
          else:
            bank_month_df['BANK_REMARKS_CLEAN'] = ''

          died_mask = bank_month_df['BANK_REMARKS_CLEAN'].str.contains(
              'DIED', na=False
          )

          active_df = bank_month_df[
              (bank_month_df['Govt_Share'] > 0)
              & (bank_month_df['BATCH_NUMBER'] == target_batch)
              & (~died_mask)
          ].drop_duplicates(subset=['RD 1'], keep='first')

          pending_df = bank_month_df[
              (bank_month_df['Govt_Share'] == 0)
              | (bank_month_df['BATCH_NUMBER'] == 0)
              | (died_mask)
          ]

          record_count = len(active_df)
          pending_count = len(pending_df)
          total_govt_share = float(active_df['Govt_Share'].sum())

          export_df = pd.DataFrame()
          export_df['BENEFICIARY SNO'] = active_df.get(
              'BANK_SNO', active_df.get('SNO', '')
          )
          export_df['Name Of Benificiary'] = active_df['Name Of Benificiary']
          export_df['Bank Account Number'] = active_df[
              'Bank Account Number'
          ].apply(format_text)
          export_df['RD ACCOUNT_1 (Beneficiary Contribution)'] = active_df[
              'RD ACCOUNT_1 (Beneficiary Contribution)'
          ].apply(format_text)
          export_df['RD-1 (Beneficiary) Contribution'] = active_df['Govt_Share']

          if 'RD 2' in active_df.columns:
            export_df['RD ACCOUNT_2 (Govt Contribution)'] = active_df[
                'RD 2'
            ].apply(format_text)
          elif 'RD ACCOUNT_2' in active_df.columns:
            export_df['RD ACCOUNT_2 (Govt Contribution)'] = active_df[
                'RD ACCOUNT_2'
            ].apply(format_text)
          elif 'RD 2_THRIFT' in active_df.columns:
            export_df['RD ACCOUNT_2 (Govt Contribution)'] = active_df[
                'RD 2_THRIFT'
            ].apply(format_text)
          else:
            export_df['RD ACCOUNT_2 (Govt Contribution)'] = ''

          export_df['RD-2 (Govt_Share) Contribution'] = active_df['Govt_Share']
          export_df['Nominee Name'] = active_df.get('Nominee Name', '').apply(
              format_text
          )
          export_df['Relation'] = active_df.get('Relation', '').apply(
              format_text
          )
          export_df['Nominee DOB'] = active_df.get('Nominee DOB', '').apply(
              format_text
          )
          export_df['Contact Number'] = (
              active_df['Contact Number'].apply(format_text)
              if 'Contact Number' in active_df.columns
              else ''
          )
          export_df.to_excel(writer, sheet_name=step_cfg['sheet_name'], index=False)

          if pending_count > 0:
            p_export_df = pd.DataFrame()
            p_export_df['SNO'] = pending_df.get(
                'BANK_SNO', pending_df.get('SNO', '')
            )
            p_export_df['Name Of Benificiary'] = pending_df[
                'Name Of Benificiary'
            ]
            p_export_df['Bank Account Number'] = pending_df[
                'Bank Account Number'
            ].apply(format_text)
            p_export_df['RD ACCOUNT_1'] = pending_df[
                'RD ACCOUNT_1 (Beneficiary Contribution)'
            ].apply(format_text)
            p_export_df['Contact Number'] = (
                pending_df['Contact Number'].apply(format_text)
                if 'Contact Number' in pending_df.columns
                else ''
            )
            p_export_df['Nominee Name'] = pending_df.get(
                'Nominee Name', ''
            ).apply(format_text)
            p_export_df['Relation'] = pending_df.get('Relation', '').apply(
                format_text
            )
            p_export_df['Nominee DOB'] = pending_df.get('Nominee DOB', '').apply(
                format_text
            )
            p_export_df['Status'] = pending_df['BANK_REMARKS_CLEAN'].apply(
                lambda x: (
                    'Beneficiary Expired'
                    if 'DIED' in str(x)
                    else (
                        'RD-1 Account Pending / Unassigned Batch / Zero'
                        ' Transaction'
                    )
                )
            )
            p_export_df.to_excel(
                writer, sheet_name='Pending_RD1_List', index=False
            )

          if not bank_not_sent_df.empty:
            build_beneficiary_export_df(bank_not_sent_df).to_excel(
                writer, sheet_name='Not_Sent_To_Bank', index=False
            )

        # --- దశ 3 & దశ 10 మరియు మిగిలిన దశలు ---
        elif process_step == '10':
          remarks_mask = bank_sent_df['BANK_REMARKS_CLEAN'] != ''
          rect_df = (
              bank_sent_df[remarks_mask].copy()
              if remarks_mask.sum() > 0
              else bank_sent_df.copy()
          )
          record_count = len(rect_df)

          export_df = pd.DataFrame()
          export_df['SNO'] = rect_df.get('BANK_SNO', rect_df.get('SNO', ''))
          export_df['Name Of Benificiary'] = rect_df.get(
              'Name Of Benificiary', ''
          )
          export_df['Savings Account No'] = rect_df.get(
              'Bank Account Number', ''
          ).apply(format_text)
          export_df['RD-1 Account No'] = rect_df.get('RD 1', '').apply(
              format_text
          )
          export_df['RD-2 Account No'] = rect_df.get('RD 2', '').apply(
              format_text
          )
          export_df['Issue / Remarks'] = rect_df.get('BANK_REMARKS', '')
          export_df['Bank Action / Updated Details'] = ''
          export_df.to_excel(writer, sheet_name=step_cfg['sheet_name'], index=False)

          if not bank_not_sent_df.empty:
            build_beneficiary_export_df(bank_not_sent_df).to_excel(
                writer, sheet_name='Not_Sent_To_Bank', index=False
            )

        else:
          record_count = len(bank_sent_df)
          export_df = build_beneficiary_export_df(bank_sent_df)
          export_df.to_excel(writer, sheet_name=step_cfg['sheet_name'], index=False)

          if not bank_not_sent_df.empty:
            build_beneficiary_export_df(bank_not_sent_df).to_excel(
                writer, sheet_name='Not_Sent_To_Bank', index=False
            )

      cheque_amount = user_cheque_amount or str(total_govt_share)

      # ========================================================
      # Word Letter తయారీ (దశకు తగిన Template వాడుతూ)
      # ========================================================
      letter_path = None
      doc_template_path = step_cfg['template']
      if not os.path.exists(doc_template_path):
        doc_template_path = 'RD2_CREDIT_SANCTION_TEMPLATE.docx'

      if os.path.exists(doc_template_path):
        import docx

        doc = docx.Document(doc_template_path)
        try:
          amt_val = float(cheque_amount) if cheque_amount else 0.0
          amt_words = num2words(amt_val, lang='en_IN').title() + ' Only'
        except Exception:
          amt_words = f'Rupees {cheque_amount} Only'

        replacements = {
            '{{Bank_Name}}': str(bank_doc_info.get('Bank_Name', selected_bank)),
            '{{Location}}': str(bank_doc_info.get('Location', 'Sircilla')),
            '{{Current_Date}}': cheque_date or time.strftime('%d/%m/%Y'),
            '{{Total_Count_Of_Records}}': str(record_count),
            '{{Cheque_No}}': cheque_no,
            '{{Cheque_Date}}': cheque_date,
            '{{Cheque_Amount}}': cheque_amount,
            '{{Amount_In_Words}}': amt_words,
            '{{Rd_Total_Count_for_month_number}}': str(record_count),
            '{{RD_Start_from_for_month_number}}': '1',
            '{{RD_End_With_for_month_number}}': str(record_count),
            '{{Month_Name}}': f'{selected_month_name} {selected_year}',
            '{{Batch_Number}}': str(target_batch),
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
            target_folder_path,
            f'Official_Letter_Step{process_step}_{selected_bank}.docx',
        )
        doc.save(letter_path)

      # ========================================================
      # Gmail Draft తయారీ (IMAP)
      # ========================================================
      subject = f"{step_cfg['subject_prefix']} - {selected_month_name} {selected_year} (Batch-{target_batch}) - Rc.No. 176/2025-RSCL - {selected_bank}"

      pending_notice_html = ''
      if pending_count > 0:
        pending_notice_html = f"""
            <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 12px; margin-top: 15px; border-radius: 4px; color: #856404;">
                <p style="margin:0; font-weight: bold;">⚠️ Pending / Excluded Accounts Notice:</p>
                <p style="margin:5px 0 0 0;">There are <b>{pending_count} beneficiaries</b> whose accounts are pending, unassigned, or marked as expired.</p>
            </div>
            """

      if process_step in ['6', '7', '8'] and (total_govt_share > 0 or cheque_no):
        cheque_info_html = f'<p><b>Summary Details:</b> Beneficiaries: {record_count} | Amount: ₹{total_govt_share:,.2f} | Cheque No: {cheque_no}</p>'
      else:
        cheque_info_html = f"<p><b>Summary Details:</b> Total Beneficiaries in List: {record_count} | Stage: {step_cfg['step_name']}</p>"

      cc_recipients_html = """
        <br>
        <p style="font-size: 13px; color: #555; line-height: 1.4;">
            <b>Copy submitted to / కాపీ సమర్పించడమైనది:</b><br>
            1. <b>Branch Manager, Union Bank of India (Nodal Branch)</b> — పాజిటివ్ పే సిస్టమ్ (PPS) / సమన్వయము కొరకు పంపడమైనది.<br>
            2. <b>Assistant Director, Handlooms & Textiles, Rajanna Sircilla</b> — కార్యాలయ రికార్డులు మరియు పర్యవేక్షణ నిమిత్తము.
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
            <p>Please find attached the official letter along with the Beneficiary Excel List for <b>{step_cfg['body_desc']}</b> for <b>{selected_month_name} {selected_year} (Batch-{target_batch})</b>.</p>
            {cheque_info_html}
            {pending_notice_html}
            <p>Kindly process the same strictly adhering to the scheme guidelines and <b>RBI Positive Pay System (PPS)</b> norms wherever applicable.</p>
            <p style="color: #444; font-weight: bold; margin-top: 15px;">ఈ విషయము తగు సమాచారము మరియు తగిన తదుపరి చర్య నిమిత్తము పంపడమైనది.</p>
            {cc_recipients_html}
            {OFFICIAL_SIGNATURE_HTML}
        </body>
        </html>
        """

      message = MIMEMultipart()
      message['To'] = bank_email
      message['Cc'] = ', '.join([
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
        part1.add_header(
            'Content-Disposition', f'attachment; filename="{excel_filename}"'
        )
        message.attach(part1)

      if letter_path and os.path.exists(letter_path):
        letter_filename = os.path.basename(letter_path)
        with open(letter_path, 'rb') as f:
          part2 = MIMEBase('application', 'octet-stream')
          part2.set_payload(f.read())
        encoders.encode_base64(part2)
        part2.add_header(
            'Content-Disposition', f'attachment; filename="{letter_filename}"'
        )
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
        except Exception:
          ws_log = sh.add_worksheet(
              title='Bank_Transactions', rows=100, cols=12
          )
          ws_log.append_row([
              'Timestamp',
              'Process Step',
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
            step_cfg['step_name'],
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

      print(
          f'Background Draft generated successfully for {selected_bank} (Step'
          f' {process_step})!'
      )

    except Exception as e:
      print(f'Background Draft Generation Error: {str(e)}')


@draft_bp.route('/api/generate-draft', methods=['POST'])
def generate_draft_api():
  try:
    from app import PROCESS_STEPS_CONFIG

    data = request.get_json() if request.is_json else request.form
    selected_bank = str(data.get('bank_name', '')).strip()
    process_step = str(data.get('process_step', '8')).strip()

    if not selected_bank:
      return (
          jsonify(
              {'status': 'error', 'message': 'దయచేసి బ్యాంకును ఎంచుకోండి!'}
          ),
          400,
      )

    step_label = PROCESS_STEPS_CONFIG.get(process_step, {}).get(
        'step_name', f'దశ {process_step}'
    )
    app_context = current_app._get_current_object().app_context()

    thread = threading.Thread(
        target=background_draft_task, args=(data, app_context)
    )
    thread.daemon = True
    thread.start()

    return jsonify({
        'status': 'success',
        'message': (
            f'"{step_label}" కోసం {selected_bank} ఇమెయిల్ డ్రాఫ్ట్'
            ' బ్యాక్‌గ్రౌండ్‌లో తయారవుతోంది. దయచేసి కొన్ని క్షణాల తరువాత మీ'
            ' జిమెయిల్ డ్రాఫ్ట్స్ చెక్ చేయండి!'
        ),
    })

  except Exception as e:
    return (
        jsonify(
            {'status': 'error', 'message': f'లోపం సంభవించింది: {str(e)}'}
        ),
        500,
    )
