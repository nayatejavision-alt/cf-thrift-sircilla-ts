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
        # 1. ఇక్కడ encoding='latin-1' ని జోడించండి
        df = pd.read_csv(file_path, sep=r'\s+', engine='python', encoding='latin-1')
    except Exception as e:
        # 2. ఇక్కడ కూడా encoding='latin-1' ని జోడించండి
        df = pd.read_fwf(file_path, encoding='latin-1')
    
    # కాలమ్ పేర్లలోని అవాంఛిత స్పేస్‌లను తొలగించడం
    df.columns = df.columns.str.strip()

    # ఉదాహరణకు మీ డేటాఫ్రేమ్ 'df' అనుకుందా
    # 1. కేవలం నంబర్లు (డిజిట్స్) ఉన్నవి మాత్రమే వాలిడ్ డేటాగా తీసుకోవడం
    valid_df = df[df['RD1 Acc Num'].astype(str).str.match('^\d+$')]

    # 2. నంబర్లు కాకుండా టెక్స్ట్/ఎర్రర్లు ఉన్న రోస్‌ని వేరు చేయడం
    error_df = df[~df['RD1 Acc Num'].astype(str).str.match('^\d+$')]

    # 3. ఆ ఎర్రర్ డేటాని ప్రత్యేకంగా ఎక్సెల్ ఫైల్‌గా సేవ్/డౌన్‌లోడ్ చేసుకోవడం
    error_df.to_excel('Invalid_Account_Numbers.xlsx', index=False)
    
    # RD1 Acc Num మరియు Bank Acc Number కాలమ్స్‌ని క్లీన్ చేయడం
    if 'RD1 Acc Num' in valid_df.columns:
        valid_df['RD1 Acc Num'] = valid_df['RD1 Acc Num'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
    if 'Bank Acc Number' in valid_df.columns:
        valid_df['Bank Acc Number'] = valid_df['Bank Acc Number'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        
    if 'Name Of Beneficiary' in valid_df.columns:
        valid_df['Name Of Beneficiary'] = valid_df['Name Of Beneficiary'].astype(str).str.strip()
        
    return valid_df
