import re
import pandas as pd
import numpy as np


def parse_term_deposit_file(file_path):
    with open(file_path, 'r', encoding='cp1252', errors='ignore') as f:
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
  
  with open(prn_file_path, 'r', encoding='cp1252') as f:
    lines = f.readlines()

  for line in lines:
    line_str = line.rstrip('\n')

    # హెడర్ లేదా ఖాళీ లైన్లను ఇగ్నోర్ చేయడానికి
    if (
        'SNO' in line_str
        or 'Contact Number' in line_str
        or not line_str.strip()
    ):
      continue

    # 1. SNO ఎక్స్‌ట్రాక్షన్ (స్పేస్ లేకపోయినా మొదటి అంకెలను SNO గా తీసుకోవడానికి)
    sno_match = re.match(r'^(\d+)(.*)$', line_str.strip())
    if not sno_match:
      continue

    sno = sno_match.group(1)
    rem_str = sno_match.group(2).strip()

    # 2. చివరన ఉన్న RD1 అకౌంట్ నంబర్ లేదా స్టేటస్ ను వేరు చేయడం (చివరి భాగం)
    # ఉదాహరణకు చివరిలో 10 అంకెల నంబర్ లేదా 'ACCOUNT CLOSED' లాంటివి ఉంటాయి
    parts = rem_str.split()
    rd1_val = '0'
    if parts:
      # చివరి పదం అంకెల్లో ఉంటే అది RD1 నంబర్
      if parts[-1].isdigit() and len(parts[-1]) >= 8:
        rd1_val = parts[-1]
        rem_str = ' '.join(
            parts[:-1]
        )  # మిగతా స్ట్రింగ్‌ను పైభాగం ప్రాసెసింగ్‌కి ఉంచడం
      elif (
          len(parts) >= 2
          and parts[-1] in ['CLOSED', 'ADDRESS', 'NUMBER', 'BALANCE']
          or 'ACCOUNT' in rem_str
          or 'TEMP' in rem_str
      ):
        # స్టేటస్ లాంటివి ఉంటే వాటిని హ్యాండిల్ చేయడం
        rd1_val = '0'

    # 3. మిగతా మధ్యలో ఉన్న నంబర్లు (Contact, Bank Acc, CIF) వెతకడానికి Regex
    # ఇది అతుక్కుపోయిన CIF నంబర్లను కూడా విడదీస్తుంది
    pattern = r'^(.*?)\s+(\d{9,11})\s+(\d{9,11})(?:\s*(\d{9,11}))?(.*)$'
    match = re.match(pattern, rem_str)

    if match:
      ben_name, contact, bank_acc, cif, rest_bank = match.groups()

      cif_val = cif if cif else ''
      bank_name_val = rest_bank.strip() if rest_bank else ''

      # ఒకవేళ bank_name_val లో నంబర్లు ఉంటే వాటిని శుభ్రం చేయడం
      # (కొన్నిసార్లు CIF నంబర్ rest_bank లోకి వెళ్తే దాన్ని సరిచేయవచ్చు)

      parsed_data.append({
          'SNO': int(sno),
          'Name Of Beneficiary': ben_name.strip(),
          'Contact Number': contact,
          'Bank Acc Number': bank_acc,
          'CIF NO': cif_val,
          'NAME AS PER BANK RECORDS': bank_name_val,
          'RD1 Acc Num': rd1_val,
      })
    else:
      # ఒకవేళ రెగ్యులర్ ప్యాటర్న్ మ్యాచ్ కాకపోతే ఫాల్‌బ్యాక్ లాజిక్
      parsed_data.append({
          'SNO': int(sno),
          'Name Of Beneficiary': rem_str,
          'Contact Number': '',
          'Bank Acc Number': '',
          'CIF NO': '',
          'NAME AS PER BANK RECORDS': '',
          'RD1 Acc Num': rd1_val,
      })

  df = pd.DataFrame(parsed_data)
return df

