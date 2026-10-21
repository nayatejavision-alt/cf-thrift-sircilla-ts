import re
import pandas as pd
import numpy as np


def parse_term_deposit_file(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
        
    current_sec = "Unknown"
    clean_rows = []
    
    for line in lines:
        if 'Fixed deposits -' in line or 'Short term deposits -' in line or 'Recurring Deposits' in line:
            current_sec = line.strip()
            
        match = re.match(r'^\s*(\d{1,4})\s+(\d+)\s+(\d+)\s+(.*?)\s+(\d+\.\d+)\s+([A-Z]+)\s+([\d,]+\.\d+)\s+(\d+)\s+(\d{2}/\d{2}/\d{4})\s+([\d,]+\.\d+)', line)
        if match:
            sno, cust_no, acc_no, acc_name, int_rate, status, balance, tenor, mat_date, avail_int = match.groups()
            clean_rows.append({
                'Section': current_sec,
                'SNO': int(sno),
                'Customer_No': cust_no,
                'Account_No': acc_no,
                'Account_Name': acc_name.strip(),
                'Int_Rate': float(int_rate),
                'Status': status,
                'Balance': float(balance.replace(',', '')),
                'Tenor': tenor,
                'Mat_Date': mat_date,
                'Avail_Int': float(avail_int.replace(',', ''))
            })
            
    return pd.DataFrame(clean_rows)


def clean_numeric_field(val):
    try:
        # ఒకవేళ విలువ స్ట్రింగ్ అయితే, ముందుగా వైట్ స్పేస్‌లను తొలగించాలి
        if isinstance(val, str):
            val = val.strip()
            # "ACCOUNT CLOSED" లేదా ఇతర టెక్స్ట్ ఉంటే దాբై నన్ (NaN) గా మార్చండి
            if not val.replace('.', '', 1).isdigit():
                return np.nan
        
        # ఫ్లోట్ లేదా ఇంటర్‌గా మార్చడానికి ప్రయత్నించండి
        num = float(val)
        
        # 0 లేదా నెగెటివ్ విలువలు ఉంటే వాటిని హ్యాండిల్ చేయడం
        if num <= 0:
            return np.nan
            
        return int(num) # లేదా అవసరాన్ని బట్టి float
        
    except (ValueError, TypeError):
        # ఏవైనా ఎర్రర్స్ వస్తే ప్రోగ్రామ్ ఆగిపోకుండా నన్ (NaN) రిటర్న్ చేస్తుంది
        return np.nan




def parse_prn_file(prn_file_path):
  parsed_data = []

  with open(prn_file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

  for line in lines:
    line_str = line.rstrip('\n')

    # హెడర్ లేదా ఖాళీ లైన్లను ఇగ్నోర్ చేయడానికి
    if 'SNO' in line_str or 'Contact Number' in line_str or not line_str.strip():
      continue

    # 1. SNO ఎక్స్‌ట్రాక్షన్
    sno_match = re.match(r'^\s*(\d+)\s+', line_str)
    if not sno_match:
      continue

    sno = sno_match.group(1)
    rem_str = line_str[sno_match.end() :]

    # 2. Regex ప్యాటర్న్
    pattern = (
        r'^(.*?)\s+(\d{9,11})\s+(\d{9,11})(?:\s+(\d{9,11}))?\s*(.*?)\s+([A-Z0-9\s-]+)$'
    )
    match = re.match(pattern, rem_str)

    if match:
      ben_name, contact, bank_acc, cif, bank_name, rd1_acc = match.groups()

      cif_val = cif if cif else ''
      bank_name_val = bank_name.strip() if bank_name else ''
      rd1_val = rd1_acc.strip() if rd1_acc else ''

      # CIF మిస్ అయ్యి, బ్యాంక్ నేమ్ ప్లేస్ లో స్టేటస్ పడితే అడ్జస్ట్ చేసే లాజిక్
      if not cif_val and bank_name_val in ['ACCOUNT', 'WRONG']:
        rd1_val = bank_name_val + ' ' + rd1_val
        bank_name_val = ''

      parsed_data.append({
          'SNO': int(sno),
          'Name Of Beneficiary': ben_name.strip(),
          'Contact Number': contact,
          'Bank Acc Number': bank_acc,
          'CIF NO': cif_val,
          'NAME AS PER BANK RECORDS': bank_name_val,
          'RD1 Acc Num': rd1_val,
      })

  # డేటాఫ్రేమ్ తయారు చేయడం
  df = pd.DataFrame(parsed_data)

  # కాలమ్ పేర్లలోని అవాంఛిత స్పేస్‌లను తొలగించడం
  df.columns = df.columns.str.strip()

  # ----------------------------------------------------
  # కొత్త లాజిక్: RD1 Acc Num లో పేర్లు+నంబర్ ఉంటే నంబర్ ఉంచి పేరు తీసేయడం, కేవలం పేరు/టెక్స్ట్ ఉంటే 0 పెట్టడం
  # ----------------------------------------------------


  def clean_rd1_account(val):
    val_str = str(val).strip()
    # నంబర్లను వెతకడం
    numbers = re.findall(r'\d+', val_str)
    if numbers:
      # నంబర్ దొరికితే దానిని రిటర్న్ చేయడం (ఎక్కువ డిజిట్లు ఉన్న అకౌంట్ నంబర్‌ను ఎంచుకుంటుంది)
      return max(numbers, key=len) if len(numbers) > 1 else numbers[0]
    else:
      # కేవలం పేరు లేదా టెక్స్ట్ ఉంటే '0' ఉంచడం
      return '0'


  if 'RD1 Acc Num' in df.columns:
    df['RD1 Acc Num'] = df['RD1 Acc Num'].apply(clean_rd1_account)

  if 'Bank Acc Number' in df.columns:
    df['Bank Acc Number'] = (
        df['Bank Acc Number']
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )

  if 'Name Of Beneficiary' in df.columns:
    df['Name Of Beneficiary'] = df['Name Of Beneficiary'].astype(str).str.strip()


  return df
