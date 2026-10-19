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
