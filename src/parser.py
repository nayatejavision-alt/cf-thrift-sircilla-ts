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

def parse_prn_file(file_path):
    """
    .prn లేదా టెక్స్ట్ ఫైల్‌ను చదివి, డేటాను క్లీన్ చేసి డేటాఫ్రేమ్‌ను రిటర్న్ చేస్తుంది.
    """
    try:
        # 1. ఇక్కడ encoding='latin-1' ని జోడించండి
        df = pd.read_csv(file_path, sep=r'\s+', engine='python', encoding='latin-1')
    except Exception as e:
        # 2. ఇక్కడ కూడా encoding='latin-1' ని జోడించండి
        df = pd.read_fwf(file_path, encoding='latin-1')
    
    # కాలమ్ పేర్లలోని అవాంఛిత స్పేస్‌లను తొలగించడం
    df.columns = df.columns.str.strip()

    df['RD1 Acc Num'] = df['RD1 Acc Num'].apply(clean_numeric_field)
    # NaN ఉన్న రోస్ అన్నింటినీ తొలగించి కొత్త డేటాఫ్రేమ్‌కు అసైన్ చేయడ
    df_cleaned = df.dropna(subset=['RD1 Acc Num']).copy()


       return df_cleaned


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

# నమూనా డేటాఫ్రేమ్ పై అప్లై చేయడం
# df['Account_No'] = df['Account_No'].apply(clean_numeric_field)
