import re
import pandas as pd


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
        # వైట్‌స్పేస్ (Spaces/Tabs) ఆధారంగా డేటాను చదవడం కోసం
        df = pd.read_csv(file_path, sep=r'\s+', engine='python')
    except Exception as e:
        # ఒకవేళ ఫిక్స్డ్ విడ్త్ ఫార్మాట్ అయితే read_fwf వాడవచ్చు
        df = pd.read_fwf(file_path)
    
    # కాలమ్ పేర్లలోని అవాంఛిత స్పేస్‌లను తొలగించడం
    df.columns = df.columns.str.strip()
    
    # RD1 Acc Num మరియు Bank Acc Number కాలమ్స్‌ని క్లీన్ చేయడం
    if 'RD1 Acc Num' in df.columns:
        df['RD1 Acc Num'] = df['RD1 Acc Num'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
    if 'Bank Acc Number' in df.columns:
        df['Bank Acc Number'] = df['Bank Acc Number'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
    if 'Name Of Beneficiary' in df.columns:
        df['Name Of Beneficiary'] = df['Name Of Beneficiary'].astype(str).str.strip()
        
    return df
