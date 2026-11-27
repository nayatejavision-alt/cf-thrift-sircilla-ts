import base64
import datetime
import io
import json
import os
import re
import time
from urllib.parse import quote
from flask import Blueprint, jsonify, make_response, redirect, render_template, request, session, url_for
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import requests

cf_bp = Blueprint('cf_bp', __name__, template_folder='templates')

# ⚡ ఇన్-మెమరీ TTL క్యాష్ (రిఫ్రెష్ ఆప్షన్ & షార్ట్ TTL తో సురక్షితం)
_CACHE = {}
CACHE_TTL_SECONDS = 15


def _cache_get(key):
  entry = _CACHE.get(key)
  if not entry:
    return None
  expires_at, val = entry
  if time.monotonic() > expires_at:
    _CACHE.pop(key, None)
    return None
  return val


def _cache_set(key, val, ttl=CACHE_TTL_SECONDS):
  _CACHE[key] = (time.monotonic() + ttl, val)


def _cache_invalidate(prefix=''):
  if not prefix:
    _CACHE.clear()
    return
  for k in [k for k in _CACHE if k.startswith(prefix)]:
    _CACHE.pop(k, None)


# నెలల పేర్ల మ్యాపింగ్
MONTH_MAP = {
    1: 'June-2026',
    2: 'July-2026',
    3: 'August-2026',
    4: 'September-2026',
    5: 'October-2026',
    6: 'November-2026',
    7: 'December-2026',
    8: 'January-2027',
    9: 'February-2027',
    10: 'March-2027',
    11: 'April-2027',
    12: 'May-2027',
}

# 📋 CF వారీగా కేటాయించిన బ్యాంకుల మ్యాపింగ్
CF_BANK_MAPPING = {
    'NARESH': ['BANK OF BARODA', 'BANK OF INDIA'],
    'SATHISH': ['KDCCB', 'PUNJAB NATIONAL BANK'],
    'NAVEEN': ['CANARA BANK', 'CENTRAL BANK OF INDIA'],
    'SRIKANTH': [
        '12903_STATE BANK OF INDIA',
        '20140_STATE BANK OF INDIA',
        '20366_STATE BANK OF INDIA',
    ],
    'DEVADAS': [
        'SIRCILLA TELANGANA GRAMEENA BANK',
        'THANGALLAPALLY TELANGANA GRAMEENA BANK',
    ],
    'NAGARJUNA': ['UNION BANK OF INDIA'],
    'KANAKARAJU': ['INDIAN BANK'],
    'KAMALAKAR': ['INDIAN BANK'],
    'MARKANDEYULU': [
        'BOINPARI UNION BANK',
        'BOINPALLI UNION BANK',
        'BOINPALLE UNION BANK',
        'NIMMAPALLY',
    ],
}

# 🏦 FILE NAME ఆధారంగా డిఫాల్ట్ / ప్రైమరీ IFSC CODE మ్యాపింగ్
FILE_NAME_IFSC_MAPPING = {
    '12903_STATE BANK OF INDIA': 'SBIN0012903',
    '20140_STATE BANK OF INDIA': 'SBIN0020140',
    '20366_STATE BANK OF INDIA': 'SBIN0020366',
    'BANK OF BARODA': 'BARB0SIRCIL',
    'BANK OF INDIA': 'BKID0005745',
    'BOINPALLI UNION BANK': 'UBIN0807486',
    'BOINPALLE UNION BANK': 'UBIN0807486',
    'BOINPARI UNION BANK': 'UBIN0807486',
    'CANARA BANK': 'CNRB0004170',
    'CENTRAL BANK OF INDIA': 'CBIN0285183',
    'INDIAN BANK': 'IDIB000S289',
    'KDCCB': 'TSAB0020033',
    'NIMMAPALLY': 'TGRB0000384',
    'PUNJAB NATIONAL BANK': 'PUNB0589900',
    'SIRCILLA TELANGANA GRAMEENA BANK': 'TGRB0000215',
    'THANGALLAPALLY TELANGANA GRAMEENA BANK': 'TGRB0000086',
    'UNION BANK OF INDIA': 'UBIN0805289',
}

# 🔑 CF పాస్‌వర్డ్స్ మ్యాపింగ్
CF_PASSWORDS = {
    'NARESH': 'naresh@2026',
    'SATHISH': 'sathish@2026',
    'NAVEEN': 'naveen@2026',
    'SRIKANTH': 'srikanth@2026',
    'DEVADAS': 'devadas@2026',
    'NAGARJUNA': 'nagarjuna@2026',
    'KANAKARAJU': 'kanakaraju@2026',
    'KAMALAKAR': 'kamalakar@2026',
    'MARKANDEYULU': 'markandeyulu@2026',
}

# 🏛️ సిరిసిల్ల మున్సిపాలిటీ వార్డుల వారీగా CF మ్యాపింగ్
WARD_CF_MAPPING = {
    20: 'DEVADAS',
    17: 'KANAKARAJU',
    18: 'KANAKARAJU',
    19: 'KANAKARAJU',
    25: 'KANAKARAJU',
    26: 'KANAKARAJU',
    4: 'SRIKANTH',
    15: 'SRIKANTH',
    16: 'SRIKANTH',
    21: 'NAGARJUNA',
    22: 'NAGARJUNA',
    23: 'NAGARJUNA',
    24: 'NAGARJUNA',
    1: 'SATHISH',
    2: 'SATHISH',
    27: 'SATHISH',
    28: 'SATHISH',
    32: 'SATHISH',
    33: 'SATHISH',
    29: 'KAMALAKAR',
    30: 'KAMALAKAR',
    31: 'KAMALAKAR',
    3: 'NAVEEN',
    5: 'NAVEEN',
    6: 'NAVEEN',
    7: 'NAVEEN',
    8: 'NAVEEN',
    9: 'NAVEEN',
    10: 'NAVEEN',
    11: 'NAVEEN',
    12: 'NAVEEN',
    13: 'NAVEEN',
    14: 'NAVEEN',
}

# 🔍 సిరిసిల్ల మున్సిపాలిటీ మాస్టర్ డేటా రేంజ్ మ్యాపింగ్
MUNICIPALITY_RANGES = [
    ('2-5', 1, 11, 1),
    ('2-5', 95, 149, 1),
    ('2-6', 1, 142, 1),
    ('2-7', 1, 140, 1),
    ('2-8', 1, 65, 1),
    ('2-4', 84, 139, 1),
    ('3-1', 1, 140, 2),
    ('3-2', 1, 139, 2),
    ('3-3', 1, 140, 2),
    ('3-4', 1, 119, 2),
    ('3-5', 1, 31, 2),
    ('3-5', 32, 122, 3),
    ('3-6', 1, 164, 3),
    ('3-7', 1, 98, 3),
    ('3-7', 99, 136, 4),
    ('3-8', 1, 100, 4),
    ('4-1', 1, 136, 4),
    ('4-2', 1, 130, 4),
    ('4-3', 1, 132, 5),
    ('4-4', 1, 132, 5),
    ('4-5', 1, 74, 5),
    ('4-5', 75, 150, 6),
    ('4-6', 1, 164, 6),
    ('4-7', 1, 123, 6),
    ('4-8', 1, 69, 6),
    ('5-1', 85, 146, 7),
    ('5-2', 1, 131, 7),
    ('5-3', 1, 125, 7),
    ('5-4', 1, 103, 7),
    ('5-5', 1, 139, 7),
    ('5-6', 1, 131, 7),
    ('5-7', 1, 138, 7),
    ('4-8', 70, 115, 8),
    ('5-1', 1, 84, 8),
    ('5-8', 1, 149, 8),
    ('6-1', 1, 161, 8),
    ('6-2', 1, 132, 8),
    ('6-3', 1, 117, 9),
    ('6-4', 1, 129, 9),
    ('6-5', 1, 97, 9),
    ('6-5', 98, 139, 10),
    ('6-6', 1, 129, 10),
    ('6-7', 1, 91, 10),
    ('6-8', 1, 95, 10),
    ('7-1', 1, 139, 11),
    ('7-2', 1, 107, 11),
    ('7-2', 108, 127, 12),
    ('7-3', 1, 133, 12),
    ('7-4', 1, 138, 12),
    ('7-5', 1, 125, 12),
    ('7-6', 1, 131, 13),
    ('7-7', 1, 146, 13),
    ('7-8', 1, 140, 13),
    ('8-1', 1, 150, 13),
    ('8-2', 1, 135, 13),
    ('8-3', 1, 69, 13),
    ('8-3', 70, 118, 14),
    ('8-4', 1, 129, 14),
    ('8-5', 1, 137, 14),
    ('8-8', 63, 99, 14),
    ('8-8', 11, 62, 15),
    ('9-3', 1, 125, 15),
    ('9-4', 1, 119, 15),
    ('8-6', 1, 130, 16),
    ('8-7', 1, 145, 16),
    ('8-8', 1, 10, 16),
    ('8-8', 100, 101, 16),
    ('9-1', 1, 150, 17),
    ('9-2', 1, 159, 17),
    ('9-5', 1, 75, 18),
    ('9-6', 1, 63, 18),
    ('10-1', 79, 144, 18),
    ('10-2', 1, 92, 18),
    ('10-3', 100, 135, 18),
    ('10-4', 1, 117, 18),
    ('10-3', 1, 99, 19),
    ('10-5', 1, 120, 19),
    ('10-6', 1, 141, 19),
    ('10-7', 1, 34, 19),
    ('9-8', 102, 139, 19),
    ('9-6', 64, 125, 20),
    ('9-7', 1, 107, 20),
    ('10-1', 1, 78, 20),
    ('9-8', 1, 101, 20),
    ('9-8', 140, 166, 20),
    ('10-7', 35, 133, 21),
    ('10-8', 1, 142, 21),
    ('11-1', 1, 156, 21),
    ('12-2', 32, 147, 22),
    ('12-3', 1, 111, 22),
    ('12-4', 1, 73, 22),
    ('12-5', 1, 112, 22),
    ('11-2', 1, 133, 23),
    ('11-3', 1, 148, 23),
    ('11-4', 1, 133, 23),
    ('11-8', 33, 132, 24),
    ('12-1', 1, 152, 24),
    ('12-5', 113, 142, 24),
    ('12-6', 1, 54, 24),
    ('12-2', 1, 31, 24),
    ('12-6', 95, 132, 24),
    ('11-5', 1, 125, 25),
    ('11-6', 1, 127, 25),
    ('11-7', 1, 162, 25),
    ('11-8', 1, 32, 25),
    ('12-6', 55, 94, 26),
    ('12-6', 133, 149, 26),
    ('12-7', 1, 149, 26),
    ('12-8', 1, 64, 26),
    ('12-8', 65, 170, 27),
    ('12-9', 111, 169, 27),
    ('12-10', 1, 169, 27),
    ('12-11', 1, 181, 28),
    ('1-1', 1, 82, 28),
    ('1-1', 83, 120, 29),
    ('1-2', 1, 146, 29),
    ('1-3', 1, 74, 29),
    ('1-6', 1, 14, 29),
    ('1-6', 19, 22, 29),
    ('1-5', 19, 134, 30),
    ('1-6', 23, 85, 30),
    ('1-7', 1, 148, 30),
    ('1-8', 1, 25, 30),
    ('1-3', 75, 161, 31),
    ('1-4', 1, 169, 31),
    ('1-6', 15, 18, 31),
    ('1-5', 1, 18, 31),
    ('2-1', 93, 122, 32),
    ('2-5', 12, 94, 32),
    ('2-3', 1, 18, 32),
    ('2-3', 122, 137, 32),
    ('2-2', 1, 128, 32),
    ('2-4', 1, 83, 32),
    ('12-9', 1, 110, 32),
    ('1-8', 26, 119, 33),
    ('2-1', 1, 76, 33),
    ('2-3', 19, 121, 33),
    ('2-1', 77, 92, 33),
]

# 🌍 ఏరియా / విలేజ్ వారీగా CF మ్యాపింగ్
AREA_CF_MAPPING = {
    'NARESH': [
        'THANGALLA PALLY',
        'THANGALLAPALLY',
        'RAMNAGAR',
        'THANGALLAPALLI',
        'PADMANAGAR',
        'LINGAMWADA',
        'OPPOSITE RADHESHYAM SATH SANGH LINGAM WADA',
        'PADMNAGAR',
        'PADMA NAGAR',
        'TEXTILEPARK',
        'TEXTILE PARK',
    ],
    'KANAKARAJU': ['THADUR', 'MANDAPALLI'],
    'MARKANDEYULU': [
        'ANANTHAPALLE',
        'BHOOPATHI NAGAR',
        'BOINPALLE',
        'BURUGUPALLY',
        'CHANDRAMPET',
        'DUNDRAPALLY',
        'JYOTHI NAGAR',
        'MUSTIPALLI',
        'NIMMAPALLI',
        'RAGUDU',
        'RAJEEV NAGAR',
        'TETTEKUNTA',
        'THADAGONDA',
        'VELJIPURAM',
    ],
}


# ==============================================================================
# HELPER & DATA CLEANING FUNCTIONS
# ==============================================================================
def clean_str_cell(val):
  """సెల్ విలువలోని .0, అపోస్ట్రోఫీలు, NaN లను సురక్షితంగా క్లీన్ చేస్తుంది."""
  if val is None or (isinstance(val, float) and pd.isna(val)):
    return ''
  s = str(val).strip()
  if not s or s.lower() in ('nan', 'none', 'null', '-'):
    return ''
  if s.endswith('.0'):
    s = s[:-2]
  return s.replace("'", '').strip()


def format_account_text(val):
  """
  బ్యాంకు అకౌంట్ నంబర్ (ముఖ్యంగా 0 తో మొదలయ్యే Union Bank & 16 అంకెల PNB ఖాతాలు)
  Google Sheets / Excel లో టెక్స్ట్‌గా సేవ్ కావడానికి ముందు సింగిల్ కోట్ (') జోడిస్తుంది.
  ఇప్పటికే ముందు ' ఉంటే మళ్ళీ కలపకుండా క్లీన్‌గా ఉంచుతుంది.
  """
  s = clean_str_cell(val)
  if not s:
    return ''
  return f"'{s}"


def clean_indian_mobile_10(raw_val):
  """10 అంకెల భారతీయ మొబైల్ నంబర్‌ను తిరిగి ఇస్తుంది (లేదా ఖాళీ స్ట్రింగ్)."""
  digits = re.sub(r'\D', '', clean_str_cell(raw_val))
  if len(digits) == 12 and digits.startswith('91'):
    digits = digits[2:]
  elif len(digits) == 11 and digits.startswith('0'):
    digits = digits[1:]
  if len(digits) == 10 and digits[0] in ('6', '7', '8', '9'):
    return digits
  return ''


def _norm_col_key(text):
  """కాలమ్ పేర్లను పోల్చడానికి spaces, underscores, dots, hyphens తొలగిస్తుంది."""
  s = re.sub(r'[^a-z0-9]+', ' ', str(text or '').strip().lower())
  return re.sub(r'\s+', ' ', s).strip()


def find_header_match(headers, candidates, allow_partial=True):
  """
    హెడర్ జాబితాలో సరిపోయే కాలమ్‌ను సురక్షితంగా కనుగొంటుంది.
    pd.Index వచ్చినా Ambiguous truth value ఎర్రర్ రాకుండా list గా మారుస్తుంది.
    """
  if headers is None or candidates is None:
    return None

  if isinstance(headers, pd.Index):
    headers_list = headers.tolist()
  else:
    headers_list = list(headers)

  if isinstance(candidates, pd.Index):
    candidates_list = candidates.tolist()
  else:
    candidates_list = list(candidates)

  if len(headers_list) == 0 or len(candidates_list) == 0:
    return None

  exact_map = {}
  norm_map = {}
  for h in headers_list:
    h_str = str(h).strip()
    if not h_str or h_str.lower() == 'nan':
      continue
    hl = h_str.lower()
    if hl not in exact_map:
      exact_map[hl] = h_str
    hn = _norm_col_key(h_str)
    if hn and hn not in norm_map:
      norm_map[hn] = h_str

  for c in candidates_list:
    c_str = str(c).strip()
    if not c_str:
      continue
    cl = c_str.lower()
    if cl in exact_map:
      return exact_map[cl]

  for c in candidates_list:
    c_str = str(c).strip()
    if not c_str:
      continue
    cn = _norm_col_key(c_str)
    if cn and cn in norm_map:
      return norm_map[cn]

  if not allow_partial:
    return None

  for c in candidates_list:
    c_str = str(c).strip()
    if not c_str:
      continue
    cn = _norm_col_key(c_str)
    if len(cn) < 4:
      continue
    pattern = re.compile(
        rf'(?:^||_|\s){re.escape(cn)}(?:$||_|\s)', re.IGNORECASE
    )
    for h_norm, h_orig in norm_map.items():
      if pattern.search(h_norm):
        return h_orig

  return None


def clean_hno_for_matching(hno_val):
  if hno_val is None:
    return ''
  s = str(hno_val).strip().upper()
  if not s or s in ['NAN', 'NONE', 'NULL', '-']:
    return ''
  s = re.sub(
      r'^\s*(?:H\.?\s*N\.?\s*O\.?|HOUSE\s*NO\.?|DOOR\s*NO\.?|D\.?\s*NO\.?|UNIT\s*H\.?\s*NO\.?|RESIDENTIAL\s*H\.?\s*NO\.?|NO\.?)\s*[:\-\.]*\s*',
      '',
      s,
      flags=re.IGNORECASE,
  )
  s = re.sub(r'H\.?\s*N\.?\s*O\.?', '', s, flags=re.IGNORECASE)
  s = re.sub(r'\s*-\s*', '-', s)
  s = re.sub(r'\s*/\s*', '/', s)
  s = re.sub(r'\s+', ' ', s).strip(' -/')
  return s


def clean_area_for_matching(area_val):
  if area_val is None:
    return ''
  s = str(area_val).strip().upper()
  if not s or s in ['NAN', 'NONE', 'NULL', '-']:
    return ''
  s = re.sub(r'H\.?\s*N\.?\s*O\.?', ' ', s, flags=re.IGNORECASE)
  s = re.sub(r'[^A-Z0-9\s]', ' ', s)
  s = re.sub(r'\s+', ' ', s).strip()
  return s


def extract_block_and_door(cleaned_hno):
  if not cleaned_hno:
    return None, None
  m = re.search(r'(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*(\d+)', cleaned_hno)
  if m:
    return f'{int(m.group(1))}-{int(m.group(2))}', int(m.group(3))
  m2 = re.search(r'(\d{1,2})\s*-\s*(\d{1,2})', cleaned_hno)
  if m2:
    return f'{int(m2.group(1))}-{int(m2.group(2))}', None
  return None, None


def _clean_ifsc(val):
  """11 అక్షరాల చెల్లుబాటు అయ్యే IFSC కోడ్‌ను క్లీన్ చేసి తిరిగి ఇస్తుంది (లేదా '')."""
  s = re.sub(r'[^A-Z0-9]', '', clean_str_cell(val).upper())
  if re.match(r'^[A-Z]{4}0[A-Z0-9]{6}$', s):
    return s
  return ''


def get_ifsc_for_file_name(file_name, records_list=None):
  """FILE NAME ఆధారంగా ముందుగా మాస్టర్ మ్యాపింగ్, ఆ తర్వాత షీట్ డేటా నుండి IFSC ఇస్తుంది."""
  clean_file = re.sub(r'\s+', ' ', str(file_name or '').strip().upper())
  if not clean_file or clean_file in ['NAN', 'NONE']:
    return ''

  if clean_file in FILE_NAME_IFSC_MAPPING:
    return FILE_NAME_IFSC_MAPPING[clean_file]

  normalized_variant = re.sub(r'BOINPARI|BOINPALLE', 'BOINPALLI', clean_file)
  if normalized_variant in FILE_NAME_IFSC_MAPPING:
    return FILE_NAME_IFSC_MAPPING[normalized_variant]

  if records_list:
    ifsc_counts = {}
    for row in records_list:
      r_file = re.sub(
          r'\s+', ' ', str(row.get('FILE NAME', '') or '').strip().upper()
      )
      r_ifsc = _clean_ifsc(
          row.get('IFSC code', '') or row.get('IFSC CODE', '') or ''
      )
      if r_file == clean_file and r_ifsc:
        ifsc_counts[r_ifsc] = ifsc_counts.get(r_ifsc, 0) + 1
    if ifsc_counts:
      return max(ifsc_counts, key=ifsc_counts.get)

  return ''


def _build_ifsc_to_file_map(records_new=None):
  """IFSC కోడ్ నుండి FILE NAME ను కనుగొనడానికి రివర్స్ మ్యాపింగ్ సిద్ధం చేస్తుంది."""
  ifsc_to_file = {}

  for f_name, ifsc_code in FILE_NAME_IFSC_MAPPING.items():
    c_ifsc = _clean_ifsc(ifsc_code)
    c_file = re.sub(r'\s+', ' ', str(f_name).strip().upper())
    if c_ifsc and c_file and c_ifsc not in ifsc_to_file:
      ifsc_to_file[c_ifsc] = c_file

  if records_new:
    for r in records_new:
      r_file = re.sub(
          r'\s+',
          ' ',
          clean_str_cell(
              r.get('FILE NAME') or r.get('Bank Name') or ''
          ).upper(),
      )
      r_ifsc = _clean_ifsc(r.get('IFSC code') or r.get('IFSC CODE') or '')
      if r_ifsc and r_file and r_ifsc not in ifsc_to_file:
        ifsc_to_file[r_ifsc] = r_file

  return ifsc_to_file


def _resolve_bank_file_and_ifsc(
    row_dict,
    records_new,
    email_map,
    cf_name='',
    default_file_name='',
    allow_single_cf_bank=True,
    ifsc_to_file_map=None,
):
  """
    రో లో ఉన్న IFSC code, Bank Name, FILE NAME లేదా CF మ్యాపింగ్ ఆధారంగా
    కచ్చితమైన FILE NAME, Bank Name, IFSC code, EMAIL ID లను నిర్ణయిస్తుంది.
    """
  if ifsc_to_file_map is None:
    ifsc_to_file_map = _build_ifsc_to_file_map(records_new)

  raw_file = re.sub(
      r'\s+',
      ' ',
      clean_str_cell(
          row_dict.get('FILE NAME')
          or row_dict.get('File Name')
          or default_file_name
          or ''
      ).upper(),
  )
  raw_bank = re.sub(
      r'\s+',
      ' ',
      clean_str_cell(
          row_dict.get('Bank Name') or row_dict.get('BANK NAME') or ''
      ).upper(),
  )
  raw_ifsc = _clean_ifsc(
      row_dict.get('IFSC code') or row_dict.get('IFSC CODE') or ''
  )

  resolved_file = ''

  # 1. ముందుగా IFSC కోడ్ ఆధారంగా FILE NAME కనుగొనడం
  if raw_ifsc and raw_ifsc in ifsc_to_file_map:
    resolved_file = ifsc_to_file_map[raw_ifsc]

  # 2. FILE NAME నేరుగా మ్యాపింగ్‌లో ఉంటే దాన్ని వాడటం
  if not resolved_file and raw_file:
    norm_file = re.sub(r'BOINPARI|BOINPALLE', 'BOINPALLI', raw_file)
    if raw_file in FILE_NAME_IFSC_MAPPING:
      resolved_file = raw_file
    elif norm_file in FILE_NAME_IFSC_MAPPING:
      resolved_file = norm_file
    else:
      resolved_file = raw_file

  # 3. Bank Name ఆధారంగా FILE NAME మ్యాచ్ చేయడం
  if not resolved_file and raw_bank:
    norm_bank = re.sub(r'BOINPARI|BOINPALLE', 'BOINPALLI', raw_bank)
    if raw_bank in FILE_NAME_IFSC_MAPPING:
      resolved_file = raw_bank
    elif norm_bank in FILE_NAME_IFSC_MAPPING:
      resolved_file = norm_bank
    else:
      for k in FILE_NAME_IFSC_MAPPING:
        if raw_bank == k or norm_bank == k:
          resolved_file = k
          break
      if not resolved_file:
        for k in FILE_NAME_IFSC_MAPPING:
          if raw_bank in k or norm_bank in k:
            resolved_file = k
            break

  # 4. ఇంకా ఖాళీగా ఉంటే CF కు కేటాయించిన బ్యాంకులలో ఒకే బ్యాంక్ ఉంటే దాన్ని తీసుకోవడం
  if not resolved_file and allow_single_cf_bank and cf_name:
    cf_banks = [
        re.sub(r'\s+', ' ', str(b).strip().upper())
        for b in CF_BANK_MAPPING.get(str(cf_name).strip().upper(), [])
    ]
    if len(cf_banks) == 1:
      resolved_file = cf_banks[0]

  resolved_ifsc = raw_ifsc
  if not resolved_ifsc and resolved_file:
    resolved_ifsc = _clean_ifsc(
        get_ifsc_for_file_name(resolved_file, records_new)
    )

  if raw_ifsc and raw_ifsc in ifsc_to_file_map:
    resolved_file = ifsc_to_file_map[raw_ifsc]

  resolved_bank = raw_bank if raw_bank else resolved_file

  resolved_email = clean_str_cell(
      row_dict.get('EMAIL ID') or row_dict.get('Email ID') or ''
  )
  if not resolved_email and resolved_file:
    resolved_email = email_map.get(resolved_file, '')

  return resolved_file, resolved_bank, resolved_ifsc, resolved_email


def normalize_aadhar(val):
  if val is None or (isinstance(val, float) and pd.isna(val)):
    return ''
  s = str(val).strip()
  if s.endswith('.0'):
    s = s[:-2]
  return re.sub(r'\D', '', s)


def find_row_by_aadhar(worksheet, aadhar_col_idx, target_aadhar_clean):
  if aadhar_col_idx == -1 or not target_aadhar_clean:
    return None, []
  all_aadhars = worksheet.col_values(aadhar_col_idx + 1)
  for idx in range(1, len(all_aadhars)):
    cell_clean = normalize_aadhar(all_aadhars[idx])
    if cell_clean and cell_clean == target_aadhar_clean:
      return idx + 1, all_aadhars
  return None, all_aadhars


def get_last_non_empty_row(worksheet):
  all_values = worksheet.get_all_values()
  last_row = 1
  for r_idx, row_vals in enumerate(all_values, start=1):
    if any(str(cell).strip() != '' for cell in row_vals):
      last_row = r_idx
  return last_row


def backup_row_before_update(
    spreadsheet, headers, old_row_values, target_row_num, cf_name, aadhar_clean
):
  backup_sheet_name = 'Update_Backup'
  try:
    ws_backup = spreadsheet.worksheet(backup_sheet_name)
  except Exception:
    ws_backup = spreadsheet.add_worksheet(
        title=backup_sheet_name, rows=1000, cols=max(len(headers) + 4, 15)
    )
    backup_headers = [
        'Backup_Timestamp',
        'Updated_By_CF',
        'Verified_Aadhaar',
        'Source_Row_Snapshot',
    ] + headers
    ws_backup.append_row(backup_headers)

  timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
  padded_old = list(old_row_values[: len(headers)])
  while len(padded_old) < len(headers):
    padded_old.append('')

  backup_row = [
      timestamp,
      cf_name,
      aadhar_clean,
      f'Row_{target_row_num}',
  ] + padded_old
  ws_backup.append_row(backup_row)


def analyze_cf_name(unit_hno, unit_area, residential_hno=''):
  cleaned_area = clean_area_for_matching(unit_area)
  hno_candidates = [
      clean_hno_for_matching(unit_hno),
      clean_hno_for_matching(residential_hno),
  ]

  for cleaned_hno in hno_candidates:
    if not cleaned_hno:
      continue
    block_prefix, door_num = extract_block_and_door(cleaned_hno)
    if not block_prefix:
      continue

    if door_num is not None:
      for prefix, min_d, max_d, ward_no in MUNICIPALITY_RANGES:
        if prefix == block_prefix and min_d <= door_num <= max_d:
          if ward_no in WARD_CF_MAPPING:
            return WARD_CF_MAPPING[ward_no]

    for prefix, min_d, max_d, ward_no in MUNICIPALITY_RANGES:
      if prefix == block_prefix and ward_no in WARD_CF_MAPPING:
        return WARD_CF_MAPPING[ward_no]

  if cleaned_area:
    for cf, areas in AREA_CF_MAPPING.items():
      for a in areas:
        cleaned_mapped_area = clean_area_for_matching(a)
        if cleaned_mapped_area and cleaned_mapped_area in cleaned_area:
          return cf

  return 'MARKANDEYULU'


def _clean_row_before_cf_mapping(mapped_dict):
  """CF మ్యాపింగ్‌కు ముందు ఇంటి నంబర్లు, ఏరియా డేటాను క్లీన్ చేస్తుంది."""
  cleaned_unit_hno = clean_hno_for_matching(
      mapped_dict.get('H.NO. Unit H.No', '')
  )
  cleaned_res_hno = clean_hno_for_matching(
      mapped_dict.get('H.NO. Residentional H.No', '')
  )
  cleaned_area = clean_area_for_matching(mapped_dict.get('Area', ''))
  return cleaned_unit_hno, cleaned_res_hno, cleaned_area


def format_date_val(val):
  if not val or pd.isna(val):
    return ''
  if isinstance(val, (datetime.date, datetime.datetime)):
    return val.strftime('%Y-%m-%d')
  val_str = str(val).strip()
  if not val_str or val_str.upper() in ['NAN', 'NONE', 'NAT']:
    return ''
  if ' ' in val_str:
    val_str = val_str.split(' ')[0]
  try:
    dt = pd.to_datetime(val_str, errors='coerce')
    if pd.notna(dt):
      return dt.strftime('%Y-%m-%d')
  except Exception:
    pass
  return val_str


def _cf_scope_details(cf_name):
  """లాగిన్ అయిన CF కు చెందిన బ్యాంకులు, వార్డులు, ఏరియాలను తిరిగి ఇస్తుంది."""
  cf_upper = str(cf_name or '').strip().upper()
  banks = CF_BANK_MAPPING.get(cf_upper, [])
  wards = sorted(
      [w for w, c in WARD_CF_MAPPING.items() if str(c).upper() == cf_upper]
  )
  areas = AREA_CF_MAPPING.get(cf_upper, [])
  return banks, wards, areas


# ==============================================================================
# 1. CF LOGIN, DASHBOARD & LOGOUT
# ==============================================================================
@cf_bp.route('/cf/login', methods=['GET', 'POST'])
def cf_login():
  if request.method == 'POST':
    data = request.get_json() if request.is_json else request.form
    cf_name = str(data.get('cf_name', '')).strip().upper()
    cf_password = str(data.get('cf_password', '')).strip()

    if not cf_name or not cf_password:
      return jsonify({
          'status': 'error',
          'message': 'దయచేసి పేరు మరియు పాస్‌వర్డ్ రెండు ఎంటర్ చేయండి!',
      })

    if cf_name in CF_PASSWORDS and CF_PASSWORDS[cf_name] == cf_password:
      session['cf_logged_in'] = True
      session['cf_name'] = cf_name
      return jsonify({'status': 'success', 'message': 'లాగిన్ విజయవంతమైంది!'})
    return jsonify({
        'status': 'error',
        'message': 'తప్పు పేరు లేదా పాస్‌వర్డ్ ఇవ్వబడింది!',
    })

  return render_template('cf_login.html')


@cf_bp.route('/cf/dashboard')
def cf_dashboard():
  if not session.get('cf_logged_in'):
    return redirect(url_for('cf_bp.cf_login'))

  cf_name = session.get('cf_name')
  allotted_banks, allotted_wards, allotted_areas = _cf_scope_details(cf_name)
  return render_template(
      'cf_dashboard.html',
      cf_name=cf_name,
      months=MONTH_MAP,
      allotted_banks=allotted_banks,
      allotted_wards=allotted_wards,
      allotted_areas=allotted_areas,
  )


@cf_bp.route('/cf/logout', methods=['POST'])
def cf_logout():
  session.pop('cf_logged_in', None)
  session.pop('cf_name', None)
  return jsonify({'status': 'success', 'message': 'లాగౌట్ అయ్యారు.'})


# ==============================================================================
# 2. CF BENEFICIARIES LIST, INLINE EDIT & TELUGU WHATSAPP/SMS TOOL
# ==============================================================================
@cf_bp.route('/cf/my-beneficiaries', methods=['GET'])
def cf_my_beneficiaries():
  """లాగిన్ అయిన CF కు సంబంధించిన లబ్ధిదారుల వివరాలను ఇస్తుంది (?refresh=1 తో క్యాష్ క్లియర్ అవుతుంది)."""
  if not session.get('cf_logged_in'):
    return (
        jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'}),
        401,
    )

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    cf_name = str(session.get('cf_name') or '').strip().upper()
    current_month = str(request.args.get('month_number', '1')).strip()
    bank_filter = str(request.args.get('bank_name', 'ALL')).strip().upper()
    force_refresh = str(request.args.get('refresh', '')).strip() in (
        '1',
        'true',
        'yes',
    )

    cache_key = f'ben:{cf_name}:{current_month}:{bank_filter}'
    if force_refresh:
      _cache_invalidate(f'ben:{cf_name}:')
    else:
      cached = _cache_get(cache_key)
      if cached is not None:
        return jsonify({'status': 'success', 'data': cached, 'cached': True})

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    ws_thrift = spreadsheet.worksheet('NEW THRIFT DATA')
    all_rows = ws_thrift.get_all_values()
    if len(all_rows) < 2:
      return jsonify({'status': 'success', 'data': []})

    headers = [str(h).strip() for h in all_rows[0]]
    idx_map = {h: i for i, h in enumerate(headers)}

    col_cf = find_header_match(headers, ['name of cf', 'cf name', 'cf_name'])
    col_name = find_header_match(
        headers, ['name of benificiary', 'name of beneficiary', 'beneficiary']
    )
    col_phone = find_header_match(
        headers, ['contact number', 'mobile', 'phone', 'contact']
    )
    col_file_bank = find_header_match(
        headers, ['file name', 'bank name', 'bank']
    )
    col_acc = find_header_match(
        headers, ['bank account number', 'account number', 'account']
    )
    col_rd1 = find_header_match(headers, ['rd 1', 'rd_1', 'rd1'])
    col_rd2 = find_header_match(headers, ['rd 2', 'rd_2', 'rd2'])
    col_bal = find_header_match(
        headers, ['bank closing balance', 'rd 1 closing balance', 'balance']
    )
    col_sno = find_header_match(headers, ['bank_sno', 'sno'])

    bal_by_sno = {}
    try:
      ws_trans = spreadsheet.worksheet('Transactions_Year1')
      for t in ws_trans.get_all_records():
        t_m = str(
            t.get('MONTH_NUMBER') or t.get('month_number') or ''
        ).strip()
        t_sno = clean_str_cell(t.get('BANK_SNO') or t.get('Bank SNO'))
        if t_m == current_month and t_sno:
          b_val = clean_str_cell(
              t.get(
                  'Bank Closing Balance (బ్యాంక్ స్టేట్‌మెంట్‌లో ఉన్న'
                  ' బ్యాలెన్స్)'
              )
              or t.get('Bank Closing Balance')
          )
          if b_val:
            bal_by_sno[t_sno] = b_val
    except Exception:
      pass

    results = []
    for r_idx, row in enumerate(all_rows[1:], start=2):

      def cell_v(col_key):
        if not col_key or col_key not in idx_map:
          return ''
        pos = idx_map[col_key]
        return clean_str_cell(row[pos]) if pos < len(row) else ''

      row_cf = cell_v(col_cf).upper()
      if col_cf and row_cf != cf_name:
        continue

      b_name = cell_v(col_file_bank)
      if bank_filter and bank_filter != 'ALL':
        if b_name.upper() != bank_filter:
          continue

      sno_v = cell_v(col_sno)
      phone_v = cell_v(col_phone)
      rd1_v = cell_v(col_rd1)
      rd2_v = cell_v(col_rd2)
      bal_v = bal_by_sno.get(sno_v) or cell_v(col_bal)

      valid_phone = bool(clean_indian_mobile_10(phone_v))
      valid_rd1 = len(re.sub(r'\D', '', rd1_v)) >= 6
      is_complete = bool(valid_phone and valid_rd1)

      results.append({
          'row_index': r_idx,
          'sno': sno_v or str(len(results) + 1),
          'masked_id': f'SNO #{sno_v}' if sno_v else f'Row {r_idx}',
          'name': cell_v(col_name),
          'phone': phone_v,
          'bank_name': b_name,
          'bank_acc': cell_v(col_acc),
          'rd1': rd1_v,
          'rd2': rd2_v,
          'balance': bal_v,
          'is_complete': is_complete,
      })

    _cache_set(cache_key, results)
    return jsonify({'status': 'success', 'data': results, 'cached': False})

  except Exception as e:
    return (
        jsonify({'status': 'error', 'message': f'జాబితా లోడ్ లోపం: {str(e)}'}),
        500,
    )


@cf_bp.route('/cf/quick-update-beneficiary', methods=['POST'])
def cf_quick_update_beneficiary():
  """CF డ్యాష్‌బోర్డ్ టేబుల్ నుండి లబ్ధిదారుడి పేరు, ఫోన్, బ్యాంకు ఖాతా, RD1, RD2 తక్షణమే అప్‌డేట్ చేస్తుంది."""
  if not session.get('cf_logged_in'):
    return (
        jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'}),
        401,
    )

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    cf_name = str(session.get('cf_name') or '').strip().upper()
    payload = request.get_json() if request.is_json else request.form
    row_index = int(payload.get('row_index', 0))
    if row_index < 2:
      return jsonify({'status': 'error', 'message': 'చెల్లని రో నంబర్.'})

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    ws = spreadsheet.worksheet('NEW THRIFT DATA')
    headers = [str(h).strip() for h in ws.row_values(1)]

    col_cf = find_header_match(headers, ['name of cf', 'cf name', 'cf_name'])
    if col_cf:
      cf_col_idx = headers.index(col_cf) + 1
      owner_cf = clean_str_cell(ws.cell(row_index, cf_col_idx).value).upper()
      if owner_cf != cf_name:
        return jsonify({
            'status': 'error',
            'message': 'ఈ రికార్డును మార్చే అనుమతి లేదు.',
        })

    existing_vals = ws.row_values(row_index)
    aadhar_col_idx = (
        headers.index('Aadhar Number') if 'Aadhar Number' in headers else -1
    )
    aadhar_clean = (
        normalize_aadhar(existing_vals[aadhar_col_idx])
        if aadhar_col_idx != -1 and aadhar_col_idx < len(existing_vals)
        else ''
    )
    backup_row_before_update(
        spreadsheet, headers, existing_vals, row_index, cf_name, aadhar_clean
    )

    col_map_fields = {
        'name': find_header_match(
            headers,
            ['name of benificiary', 'name of beneficiary', 'beneficiary'],
        ),
        'phone': find_header_match(
            headers, ['contact number', 'mobile', 'phone', 'contact']
        ),
        'bank_acc': find_header_match(
            headers, ['bank account number', 'account number', 'account']
        ),
        'rd1': find_header_match(headers, ['rd 1', 'rd_1', 'rd1']),
        'rd2': find_header_match(headers, ['rd 2', 'rd_2', 'rd2']),
    }

    for key, col_name in col_map_fields.items():
      if key in payload and col_name:
        c_idx = headers.index(col_name) + 1
        if key in ('bank_acc', 'rd1', 'rd2'):
          clean_v = format_account_text(payload.get(key))
        else:
          clean_v = clean_str_cell(payload.get(key))
        ws.update_cell(row_index, c_idx, clean_v)

    _cache_invalidate(f'ben:{cf_name}:')
    return jsonify(
        {'status': 'success', 'message': 'వివరాలు అప్‌డేట్ అయ్యాయి!'}
    )

  except Exception as e:
    return jsonify({'status': 'error', 'message': f'అప్‌డేట్ లోపం: {str(e)}'})


@cf_bp.route('/cf/generate-whatsapp-links-excel', methods=['POST'])
def cf_generate_whatsapp_links_excel():
  """అప్‌లోడ్ చేసిన ఎక్సెల్/CSV నుండి తెలుగు WhatsApp & SMS లింకుల ఎక్సెల్ జనరేట్ చేస్తుంది."""
  if not session.get('cf_logged_in'):
    return (
        jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'}),
        401,
    )

  try:
    if 'wa_file' not in request.files:
      return 'దయచేసి ఫైల్ అప్‌లోడ్ చేయండి.', 400
    f = request.files['wa_file']
    if not f or not f.filename:
      return 'ఫైల్ ఎంచుకోలేదు.', 400

    sheet_name = str(request.form.get('wa_sheet', '')).strip()
    month_label = str(request.form.get('wa_month_label', '')).strip()
    custom_tpl = str(request.form.get('wa_custom_template', '')).strip()

    col_phone = str(request.form.get('wa_col_phone', '')).strip()
    col_name = str(request.form.get('wa_col_name', '')).strip()
    col_bank = str(request.form.get('wa_col_bank', '')).strip()
    col_acc = str(request.form.get('wa_col_acc', '')).strip()
    col_rd1 = str(request.form.get('wa_col_rd1', '')).strip()
    col_rd2 = str(request.form.get('wa_col_rd2', '')).strip()

    if f.filename.lower().endswith('.csv'):
      df = pd.read_csv(f, dtype=str)
    else:
      xl = pd.ExcelFile(f)
      target_sheet = (
          sheet_name if sheet_name in xl.sheet_names else xl.sheet_names[0]
      )
      df = pd.read_excel(xl, sheet_name=target_sheet, dtype=str)

    df.columns = [str(c).strip() for c in df.columns]
    if col_phone not in df.columns:
      return f"'{col_phone}' కాలమ్ ఫైల్‌లో కనుగొనబడలేదు.", 400

    valid_rows = []
    invalid_rows = []

    for idx, row in df.iterrows():
      m10 = clean_indian_mobile_10(row.get(col_phone, ''))
      b_name = (
          clean_str_cell(row.get(col_name, '')) if col_name in df.columns else ''
      )
      b_bank = (
          clean_str_cell(row.get(col_bank, '')) if col_bank in df.columns else ''
      )
      b_acc = (
          clean_str_cell(row.get(col_acc, '')) if col_acc in df.columns else ''
      )
      r1 = clean_str_cell(row.get(col_rd1, '')) if col_rd1 in df.columns else ''
      r2 = clean_str_cell(row.get(col_rd2, '')) if col_rd2 in df.columns else ''

      if custom_tpl:
        msg = (
            custom_tpl.replace('{name}', b_name or 'లబ్ధిదారులారా')
            .replace('{bank}', b_bank or '-')
            .replace('{acc}', b_acc or '-')
            .replace('{rd1}', r1 or '-')
            .replace('{rd2}', r2 or '-')
            .replace('{balance}', r2 or '-')
            .replace('{month}', month_label or '')
        )
      else:
        m_txt = f' ({month_label})' if month_label else ''
        msg = (
            f"నమస్కారం {b_name or 'లబ్ధిదారులారా'} గారు,\n"
            f'చేనేత మరియు జౌళి శాఖ - త్రిఫ్టు పథకం{m_txt} మీ వివరాలు:\n'
            f'• బ్యాంక్: {b_bank or "-"}\n'
            f'• ఖాతా సంఖ్య: {b_acc or "-"}\n'
            f'• RD-1 ఖాతా: {r1 or "-"}\n'
            f'• RD-2 ఖాతా: {r2 or "-"}\n'
            'ఏవైనా లోపాలు/మార్పులు ఉంటే వెంటనే మీ CF గారికి తెలియజేయండి.'
        )

      encoded_msg = quote(msg, safe='')
      if m10:
        valid_rows.append({
            'S.No': len(valid_rows) + 1,
            'Beneficiary_Name': b_name,
            'Mobile_10_Digit': m10,
            'Bank_Name': b_bank,
            'Bank_Account': format_account_text(b_acc),
            'RD1_Account': format_account_text(r1),
            'RD2_Account': format_account_text(r2),
            'Telugu_Message': msg,
            'WhatsApp_Link': (
                f'https://api.whatsapp.com/send?phone=91{m10}&text={encoded_msg}'
            ),
            'SMS_Link': f'sms:+91{m10}?body={encoded_msg}',
        })
      else:
        invalid_rows.append({
            'Row_No': idx + 2,
            'Beneficiary_Name': b_name,
            'Raw_Phone': clean_str_cell(row.get(col_phone, '')),
            'RD1_Account': format_account_text(r1),
        })

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
      df_valid = pd.DataFrame(valid_rows)
      df_invalid = pd.DataFrame(invalid_rows)
      df_valid.to_excel(writer, sheet_name='WhatsApp_SMS_Links', index=False)
      df_invalid.to_excel(
          writer, sheet_name='Invalid_Missing_Phone', index=False
      )

      if not df_valid.empty:
        ws = writer.book['WhatsApp_SMS_Links']
        wa_c = list(df_valid.columns).index('WhatsApp_Link') + 1
        sms_c = list(df_valid.columns).index('SMS_Link') + 1
        for r_i in range(2, len(df_valid) + 2):
          c1 = ws.cell(r_i, wa_c)
          c1.hyperlink = str(c1.value)
          c1.font = Font(color='0563C1', underline='single')
          c2 = ws.cell(r_i, sms_c)
          c2.hyperlink = str(c2.value)
          c2.font = Font(color='0563C1', underline='single')
    buf.seek(0)

    cf_lbl = session.get('cf_name', 'CF')
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(
        buf,
        as_attachment=True,
        download_name=f'{cf_lbl}_Telugu_WhatsApp_SMS_Links_{ts}.xlsx',
        mimetype=(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ),
    )
  except Exception as e:
    return f'WhatsApp/SMS లింక్ జనరేషన్ లోపం: {str(e)}', 500


# ==============================================================================
# 3. MONTHLY REPORTS DOWNLOAD (NAME-BASED & BANK-ALLOTTED)
# ==============================================================================
def _build_monthly_cross_df(spreadsheet, current_month):
  df_trans = pd.DataFrame(
      spreadsheet.worksheet('Transactions_Year1').get_all_records()
  )
  df_thrift = pd.DataFrame(
      spreadsheet.worksheet('NEW THRIFT DATA').get_all_records()
  )

  df_trans['Govt_Share'] = pd.to_numeric(
      df_trans['Govt_Share'], errors='coerce'
  ).fillna(0)
  df_trans['MONTH_NUMBER'] = pd.to_numeric(
      df_trans['MONTH_NUMBER'], errors='coerce'
  ).fillna(0)

  df_trans_month = df_trans[df_trans['MONTH_NUMBER'] == current_month].copy()
  df_trans_month = df_trans_month.drop_duplicates(
      subset=['MONTH_NUMBER', 'RD ACCOUNT_1 (Beneficiary Contribution)'],
      keep='first',
  )

  df_thrift['RD 1_Clean'] = df_thrift['RD 1'].apply(clean_str_cell)
  df_trans_month['RD1_Clean'] = df_trans_month[
      'RD ACCOUNT_1 (Beneficiary Contribution)'
  ].apply(clean_str_cell)

  df_cross = pd.merge(
      df_thrift,
      df_trans_month[['RD1_Clean', 'Govt_Share', 'MONTH_NUMBER']],
      left_on='RD 1_Clean',
      right_on='RD1_Clean',
      how='left',
  )
  df_cross['Govt_Share'] = df_cross['Govt_Share'].fillna(0)
  return df_cross


@cf_bp.route('/cf/download-report', methods=['POST'])
def cf_download_report_by_name():
  if not session.get('cf_logged_in'):
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    cf_name = session.get('cf_name')
    current_month = int(request.form.get('month_number', 1))

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    df_cross = _build_monthly_cross_df(spreadsheet, current_month)

    cf_col = 'NAME OF CF' if 'NAME OF CF' in df_cross.columns else 'CF NAME'
    group_df = df_cross[
        df_cross[cf_col].astype(str).str.strip().str.upper() == cf_name
    ].copy()

    if group_df.empty:
      return 'మీ పరిధిలో ఎలాంటి డేటా లభించలేదు.', 404

    return generate_excel_response(
        group_df, cf_name, current_month, 'Name_Based'
    )

  except Exception as e:
    return f'రిపోర్ట్ జనరేట్ చేయడంలో లోపం: {str(e)}'


@cf_bp.route('/cf/download-report-allotted-bank', methods=['POST'])
def cf_download_report_by_bank():
  if not session.get('cf_logged_in'):
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    cf_name = session.get('cf_name')
    current_month = int(request.form.get('month_number', 1))

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    df_cross = _build_monthly_cross_df(spreadsheet, current_month)

    bank_col = (
        'FILE NAME'
        if 'FILE NAME' in df_cross.columns
        else ('File Name' if 'File Name' in df_cross.columns else 'Bank Name')
    )
    allowed_banks = [
        b.strip().upper() for b in CF_BANK_MAPPING.get(cf_name, [])
    ]

    def match_bank(row):
      b_name = str(row.get(bank_col, '')).strip().upper()
      if not b_name or b_name in ('NAN', 'NONE'):
        return False
      return b_name in allowed_banks

    group_df = df_cross[df_cross.apply(match_bank, axis=1)].copy()

    if group_df.empty:
      return (
          f'మీ పరిధిలోని బ్యాంకులకు ({allowed_banks}) సంబంధించిన ఎలాంటి డేటా'
          ' లభించలేదు.',
          404,
      )

    return generate_excel_response(
        group_df, cf_name, current_month, 'Bank_Allotted'
    )

  except Exception as e:
    return f'రిపోర్ట్ జనరేట్ చేయడంలో లోపం: {str(e)}'


def generate_excel_response(group_df, cf_name, current_month, report_type):
  # డబుల్ సింగిల్ కోట్స్ ('') రాకుండా format_account_text తో సురక్షితంగా ఫార్మాట్ చేయడం
  for col in ['Bank Account Number', 'RD 1', 'RD 2']:
    if col in group_df.columns:
      group_df[col] = group_df[col].apply(format_account_text)

  df_active = group_df[group_df['Govt_Share'] > 0].copy()
  df_missing = group_df[group_df['Govt_Share'] == 0].copy()

  wb = openpyxl.Workbook()
  header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
  blue_fill = PatternFill(
      start_color='1F4E78', end_color='1F4E78', fill_type='solid'
  )
  red_fill = PatternFill(
      start_color='C00000', end_color='C00000', fill_type='solid'
  )
  green_fill = PatternFill(
      start_color='008000', end_color='008000', fill_type='solid'
  )

  ws1 = wb.active
  ws1.title = 'Sent_To_Bank_List'
  ws1.append([
      'BANK_SNO',
      'Name Of Benificiary',
      'Aadhar Number',
      'Bank Account Number',
      'RD 1',
      'Contact Number',
      'FILE NAME',
      'Govt Share Amount',
      'Status',
      'NAME OF CF',
  ])
  for col_num in range(1, 11):
    ws1.cell(row=1, column=col_num).font = header_font
    ws1.cell(row=1, column=col_num).fill = blue_fill
  for _, r in df_active.iterrows():
    ws1.append([
        r.get('BANK_SNO', ''),
        r.get('Name Of Benificiary', ''),
        r.get('Aadhar Number', ''),
        r.get('Bank Account Number', ''),
        r.get('RD 1', ''),
        r.get('Contact Number', ''),
        r.get('FILE NAME', ''),
        r.get('Govt_Share', 0),
        f'{MONTH_MAP[current_month]} Govt_Share Released',
        r.get('NAME OF CF', ''),
    ])

  ws2 = wb.create_sheet(title='Missing_or_Zero_List')
  ws2.append([
      'BANK_SNO',
      'Name Of Benificiary',
      'Aadhar Number',
      'Bank Account Number',
      'RD 1',
      'Contact Number',
      'FILE NAME',
      'Status',
      'NAME OF CF',
  ])
  for col_num in range(1, 10):
    ws2.cell(row=1, column=col_num).font = header_font
    ws2.cell(row=1, column=col_num).fill = red_fill
  for _, r in df_missing.iterrows():
    ws2.append([
        r.get('BANK_SNO', ''),
        r.get('Name Of Benificiary', ''),
        r.get('Aadhar Number', ''),
        r.get('Bank Account Number', ''),
        r.get('RD 1', ''),
        r.get('Contact Number', ''),
        r.get('FILE NAME', ''),
        f'Zero or missing transactions for {MONTH_MAP[current_month]}',
        r.get('NAME OF CF', ''),
    ])

  ws3 = wb.create_sheet(title='All_Beneficiaries_Master')
  ws3.append([
      'BANK_SNO',
      'Name Of Benificiary',
      'Aadhar Number',
      'Bank Account Number',
      'RD 1',
      'Contact Number',
      'FILE NAME',
      'Govt_Share',
  ])
  for col_num in range(1, 9):
    ws3.cell(row=1, column=col_num).font = header_font
    ws3.cell(row=1, column=col_num).fill = green_fill
  for _, r in group_df.iterrows():
    ws3.append([
        r.get('BANK_SNO', ''),
        r.get('Name Of Benificiary', ''),
        r.get('Aadhar Number', ''),
        r.get('Bank Account Number', ''),
        r.get('RD 1', ''),
        r.get('Contact Number', ''),
        r.get('FILE NAME', ''),
        r.get('Govt_Share', 0),
    ])

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)

  response = make_response(output.read())
  response.headers['Content-Disposition'] = (
      f'attachment; filename={cf_name}_{report_type}_Month_{current_month}_Report.xlsx'
  )
  response.headers['Content-Type'] = (
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  )
  return response


# ==============================================================================
# 4. SEARCH, BANK FILES, SAVE & AI OCR
# ==============================================================================
@cf_bp.route('/cf/search-beneficiary', methods=['POST'])
def cf_search_beneficiary():
  if not session.get('cf_logged_in'):
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    cf_name = session.get('cf_name')
    data = request.get_json() if request.is_json else request.form
    raw_aadhar = str(data.get('aadhar_number', '')).strip()
    aadhar_clean = normalize_aadhar(raw_aadhar)
    selected_file_name = str(data.get('file_name', '')).strip()

    if len(aadhar_clean) != 12:
      return jsonify({
          'status': 'error',
          'message': 'దయచేసి సరైన 12 అంకెల నంబర్ ఎంటర్ చేయండి!',
      })
    if not selected_file_name:
      return jsonify({
          'status': 'error',
          'message': 'దయచేసి ముందుగా FILE NAME ఎంచుకోండి!',
      })

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)

    email_id = ''
    try:
      ws_email = spreadsheet.worksheet('Bank_Email_Ids')
      for row in ws_email.get_all_records():
        if (
            str(row.get('FILE_NAME') or row.get('FILE NAME', ''))
            .strip()
            .upper()
            == selected_file_name.upper()
        ):
          email_id = str(row.get('OFFICIAL_EMAIL', '')).strip()
          break
    except Exception:
      pass

    worksheet_new = spreadsheet.worksheet('NEW THRIFT DATA')
    records_new = worksheet_new.get_all_records()
    columns = worksheet_new.row_values(1)
    auto_ifsc = get_ifsc_for_file_name(selected_file_name, records_new)

    matched_record = None
    source_sheet = None

    for row in records_new:
      if normalize_aadhar(row.get('Aadhar Number', '')) == aadhar_clean:
        matched_record = row
        source_sheet = 'new'
        break

    if not matched_record:
      try:
        worksheet_old = spreadsheet.worksheet('Old Thrift Data')
        for row in worksheet_old.get_all_records():
          if normalize_aadhar(row.get('Aadhar Number', '')) == aadhar_clean:
            matched_record = row
            source_sheet = 'old'
            break
      except Exception:
        pass

    if matched_record:
      cleaned_matched = {}
      for k, v in matched_record.items():
        if any(dk in k.lower() for dk in ['dob', 'date of birth', 'date']):
          cleaned_matched[k] = format_date_val(v)
        elif k in ('Bank Account Number', 'RD 1', 'RD 2'):
          cleaned_matched[k] = clean_str_cell(v)
        else:
          cleaned_matched[k] = v
      matched_record = cleaned_matched

    max_sno = 0
    max_bank_sno = 0
    for row in records_new:
      s_val = str(row.get('SNO', '0')).strip()
      if s_val.isdigit():
        max_sno = max(max_sno, int(s_val))
      if (
          str(row.get('FILE NAME', '')).strip().upper()
          == selected_file_name.upper()
      ):
        b_sno = str(row.get('BANK_SNO', '0')).strip()
        if b_sno.isdigit():
          max_bank_sno = max(max_bank_sno, int(b_sno))

    unique_professions = sorted(
        list({
            str(r.get('Profession', '')).strip()
            for r in records_new
            if r.get('Profession')
        })
    )
    unique_villages = sorted(
        list({
            str(r.get('Village', '')).strip()
            for r in records_new
            if r.get('Village')
        })
    )
    unique_areas = sorted(
        list({
            str(r.get('Area', '')).strip() for r in records_new if r.get('Area')
        })
    )
    unique_unit_owners = sorted(
        list({
            str(r.get('Unit Owner', '')).strip()
            for r in records_new
            if r.get('Unit Owner')
        })
    )
    unique_unit_hns = sorted(
        list({
            str(r.get('H.NO. Unit H.No', '')).strip()
            for r in records_new
            if r.get('H.NO. Unit H.No')
        })
    )
    unique_relations = sorted(
        list({
            str(r.get('Relation', '')).strip()
            for r in records_new
            if r.get('Relation')
        })
    )

    dropdown_payload = {
        'professions': unique_professions,
        'villages': unique_villages,
        'areas': unique_areas,
        'unit_owners': unique_unit_owners,
        'unit_hns': unique_unit_hns,
        'relations': unique_relations,
    }

    if matched_record and source_sheet == 'new':
      if not str(matched_record.get('IFSC code', '')).strip() and auto_ifsc:
        matched_record['IFSC code'] = auto_ifsc
      return jsonify({
          'status': 'existing',
          'message': (
              'రికార్డు కనుగొనబడింది - ఆధార్ ఆధారంగా సురక్షితంగా అప్‌డేట్'
              ' (Update) అవుతుంది'
          ),
          'data': matched_record,
          'auto_ifsc': auto_ifsc,
          'columns': columns,
          'dropdown_data': dropdown_payload,
      })

    base_data = matched_record.copy() if matched_record else {}
    base_data['Aadhar Number'] = aadhar_clean
    base_data['SNO'] = max_sno + 1
    base_data['BANK_SNO'] = max_bank_sno + 1
    base_data['FILE NAME'] = selected_file_name
    base_data['Bank Name'] = selected_file_name
    if auto_ifsc:
      base_data['IFSC code'] = auto_ifsc
    base_data['Caste'] = 'BC'
    base_data['Profession'] = 'POWER LOOM WEAVER'
    base_data['DISTRICT NAME'] = 'RAJANNA SIRCILLA'
    base_data['MANDAL NAME'] = 'SIRCILLA'
    base_data['NAME OF CF'] = cf_name
    base_data['EMAIL ID'] = email_id
    base_data['BANK SENT'] = 'NOT SENT'
    base_data['8% Of Monthly Wages'] = 1200
    base_data['MONTHLY WAGES'] = int(1200 * 100 / 8)

    return jsonify({
        'status': 'new',
        'message': (
            'ఇది కొత్త ఎంట్రీ (New Entry) కింద చివరి డేటా రో కింద సురక్షితంగా సేవ్'
            ' అవుతుంది.'
        ),
        'data': base_data,
        'auto_ifsc': auto_ifsc,
        'columns': columns,
        'dropdown_data': dropdown_payload,
    })

  except Exception as e:
    return jsonify({'status': 'error', 'message': str(e)})


@cf_bp.route('/cf/get-bank-files', methods=['GET'])
def cf_get_bank_files():
  if not session.get('cf_logged_in'):
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})
  try:
    cached = _cache_get('bank_files')
    if cached is not None:
      return jsonify({'status': 'success', **cached})

    from app import get_sheets_client, SPREADSHEET_NAME

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    ws = spreadsheet.worksheet('Bank_Email_Ids')
    records = ws.get_all_records()

    records_new = []
    try:
      records_new = spreadsheet.worksheet('NEW THRIFT DATA').get_all_records()
    except Exception:
      pass

    file_names = []
    file_ifsc_map = {}
    for r in records:
      f_name = str(r.get('FILE_NAME') or r.get('FILE NAME') or '').strip()
      if f_name:
        file_names.append(f_name)
        file_ifsc_map[f_name] = get_ifsc_for_file_name(f_name, records_new)

    payload = {'file_names': file_names, 'file_ifsc_map': file_ifsc_map}
    _cache_set('bank_files', payload, ttl=120)
    return jsonify({'status': 'success', **payload})
  except Exception as e:
    return jsonify({
        'status': 'error',
        'message': str(e),
        'file_names': [],
        'file_ifsc_map': {},
    })


@cf_bp.route('/cf/save-beneficiary', methods=['POST'])
def cf_save_beneficiary():
  if not session.get('cf_logged_in'):
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    cf_name = session.get('cf_name')
    data = request.get_json() if request.is_json else request.form

    requested_action = str(data.get('action_type', '')).strip().lower()
    form_data = data.get('form_data') or {}

    if not isinstance(form_data, dict):
      return jsonify(
          {'status': 'error', 'message': 'సరైన ఫారమ్ డేటా అందలేదు!'}
      )

    form_aadhar_clean = normalize_aadhar(form_data.get('Aadhar Number', ''))
    if len(form_aadhar_clean) != 12:
      return jsonify({
          'status': 'error',
          'message': 'సరైన 12 అంకెల Aadhar Number తప్పనిసరి!',
      })

    form_data['Aadhar Number'] = form_aadhar_clean

    # బ్యాంకు ఖాతా నంబర్, RD 1, RD 2 లను టెక్స్ట్‌గా సేవ్ చేయడానికి ముందు సింగిల్ కోట్ (') జోడించడం
    for acc_col in ('Bank Account Number', 'RD 1', 'RD 2'):
      if acc_col in form_data:
        form_data[acc_col] = format_account_text(form_data.get(acc_col, ''))

    unit_hno = form_data.get('H.NO. Unit H.No', '')
    residential_hno = form_data.get('H.NO. Residentional H.No', '')
    unit_area = form_data.get('Area', '')
    analyzed_cf = analyze_cf_name(unit_hno, unit_area, residential_hno)
    form_data['NAME OF CF'] = analyzed_cf

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    worksheet = spreadsheet.worksheet('NEW THRIFT DATA')
    headers = worksheet.row_values(1)

    if not headers:
      return jsonify(
          {'status': 'error', 'message': 'షీట్ హెడర్స్ (Row 1) కనుగొనబడలేదు!'}
      )

    aadhar_col_idx = (
        headers.index('Aadhar Number') if 'Aadhar Number' in headers else -1
    )
    if aadhar_col_idx == -1:
      return jsonify({
          'status': 'error',
          'message': "షీట్‌లో 'Aadhar Number' కాలమ్ కనుగొనబడలేదు!",
      })

    selected_file = str(form_data.get('FILE NAME', '')).strip()
    if selected_file:
      auto_ifsc = get_ifsc_for_file_name(selected_file)
      if not auto_ifsc:
        auto_ifsc = get_ifsc_for_file_name(
            selected_file, worksheet.get_all_records()
        )
      if auto_ifsc:
        form_data['IFSC code'] = auto_ifsc

    matched_row_num, all_aadhars = find_row_by_aadhar(
        worksheet, aadhar_col_idx, form_aadhar_clean
    )

    if requested_action == 'insert':
      if matched_row_num is not None:
        return jsonify({
            'status': 'error',
            'message': f'ఈ ఆధార్ నంబర్ ఇప్పటికే Row {matched_row_num} లో ఉంది!',
        })
      effective_action = 'insert'
    elif requested_action == 'update':
      if matched_row_num is None:
        return jsonify({
            'status': 'error',
            'message': 'అప్‌డేట్ చేయడానికి ఈ ఆధార్ నంబర్ షీట్‌లో కనుగొనబడలేదు!',
        })
      effective_action = 'update'
    else:
      effective_action = 'update' if matched_row_num is not None else 'insert'

    if effective_action == 'update':
      live_cell_aadhar = normalize_aadhar(all_aadhars[matched_row_num - 1])
      if live_cell_aadhar != form_aadhar_clean:
        return jsonify({
            'status': 'error',
            'message': 'భద్రతా తనిఖీ విఫలమైంది: ఆధార్ నంబర్ సరిపోలలేదు.',
        })

      existing_row_values = worksheet.row_values(matched_row_num)
      backup_row_before_update(
          spreadsheet,
          headers,
          existing_row_values,
          matched_row_num,
          cf_name,
          form_aadhar_clean,
      )

      updated_row = list(existing_row_values[: len(headers)])
      while len(updated_row) < len(headers):
        updated_row.append('')

      for col_name, val in form_data.items():
        if col_name in headers:
          c_idx = headers.index(col_name)
          if col_name in ('Bank Account Number', 'RD 1', 'RD 2'):
            updated_row[c_idx] = format_account_text(val)
          else:
            updated_row[c_idx] = '' if val is None else val

      # పాత రో లో ఉన్న అకౌంట్ నంబర్లకు కూడా సింగిల్ కోట్ (') ఉండేలా నిర్ధారించడం
      for acc_col in ('Bank Account Number', 'RD 1', 'RD 2'):
        if acc_col in headers:
          ac_idx = headers.index(acc_col)
          updated_row[ac_idx] = format_account_text(updated_row[ac_idx])

      updated_row[aadhar_col_idx] = form_aadhar_clean
      end_col_letter = get_column_letter(len(headers))
      worksheet.update(
          f'A{matched_row_num}:{end_col_letter}{matched_row_num}',
          [updated_row],
          value_input_option='USER_ENTERED',
      )

      _cache_invalidate('ben:')
      _cache_invalidate('bank_files')
      return jsonify({
          'status': 'success',
          'message': (
              f'డేటా సురక్షితంగా అప్‌డేట్ చేయబడింది! CF: {analyzed_cf}'
          ),
      })

    elif effective_action == 'insert':
      recheck_row, _ = find_row_by_aadhar(
          worksheet, aadhar_col_idx, form_aadhar_clean
      )
      if recheck_row is not None:
        return jsonify({
            'status': 'error',
            'message': f'ఈ ఆధార్ నంబర్ ఇప్పటికే Row {recheck_row} లో ఉంది!',
        })

      max_sno = 0
      if 'SNO' in headers:
        sno_col_idx = headers.index('SNO') + 1
        all_snos = worksheet.col_values(sno_col_idx)[1:]
        max_sno = max(
            [
                int(str(s).strip())
                for s in all_snos
                if str(s).strip().isdigit()
            ]
            or [0]
        )

      form_data['SNO'] = max_sno + 1
      new_row = []
      for h in headers:
        raw_v = form_data.get(h, '')
        if h in ('Bank Account Number', 'RD 1', 'RD 2'):
          new_row.append(format_account_text(raw_v))
        else:
          new_row.append('' if raw_v is None else raw_v)

      next_row_num = get_last_non_empty_row(worksheet) + 1
      end_col_letter = get_column_letter(len(headers))
      worksheet.update(
          f'A{next_row_num}:{end_col_letter}{next_row_num}',
          [new_row],
          value_input_option='USER_ENTERED',
      )

      _cache_invalidate('ben:')
      _cache_invalidate('bank_files')
      return jsonify({
          'status': 'success',
          'message': f'కొత్త రికార్డు సేవ్ చేయబడింది! CF: {analyzed_cf}',
      })

    return jsonify({'status': 'error', 'message': 'సరికైన యాక్షన్ రాలేదు!'})

  except Exception as e:
    return jsonify({'status': 'error', 'message': f'సేవ్ చేయడంలో లోపం: {str(e)}'})


@cf_bp.route('/cf/ai-ocr', methods=['POST'])
def cf_ai_ocr():
  if not session.get('cf_logged_in'):
    return jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'})

  api_key = os.environ.get('OPENROUTER_API_KEY', '').strip()
  if not api_key:
    return jsonify({
        'status': 'fallback',
        'message': 'AI API Key లేదు, లోకల్ OCR వాడుతోంది.',
    })

  try:
    data = request.get_json() or {}
    image_base64 = data.get('image_base64', '')
    scan_type = data.get('scan_type', 'Aadhar Number')

    if not image_base64:
      return jsonify({'status': 'error', 'message': 'ఇమేజ్ అందలేదు.'})

    model_name = os.environ.get(
        'QWEN_VL_MODEL', 'qwen/qwen2.5-vl-72b-instruct:free'
    )
    prompt_text = f"""
You are an expert Indian government document & handwritten form data extractor for Telangana Handlooms & Textiles.
Analyze this image carefully and return ONLY a valid JSON object with keys:
"number", "id_number", "name", "father_name", "dob", "house_no", "village", "mandal", "district", "mobile", "account_number", "ifsc", "unit_hno", "unit_owner", "nominee_name", "nominee_id", "nominee_dob".
Scan context: "{scan_type}". Use "" if not visible. Return ONLY raw JSON.
"""
    resp = requests.post(
        'https://openrouter.ai/api/v1/chat/completions',
        headers={
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json',
        },
        json={
            'model': model_name,
            'temperature': 0.0,
            'messages': [{
                'role': 'user',
                'content': [
                    {'type': 'text', 'text': prompt_text},
                    {'type': 'image_url', 'image_url': {'url': image_base64}},
                ],
            }],
        },
        timeout=25,
    )

    if resp.status_code != 200:
      return jsonify({
          'status': 'fallback',
          'message': f'AI API Status {resp.status_code}',
      })

    raw_content = resp.json()['choices'][0]['message']['content'].strip()
    raw_content = re.sub(r'^```(?:json)?\s*', '', raw_content)
    raw_content = re.sub(r'\s*```$', '', raw_content).strip()
    return jsonify({'status': 'success', 'extracted': json.loads(raw_content)})

  except Exception as e:
    return jsonify({'status': 'fallback', 'message': str(e)})


# ==============================================================================
# 5. BULK UPLOAD TO NEW THRIFT DATA (WITH TEXT ACCOUNT FORMATTING)
# ==============================================================================
def _build_existing_records_excel_b64(existing_items, headers):
  """ఇన్సర్ట్ చేయకుండా ఆపివేసిన (ఇదివరకు ఉన్న) రికార్డ్‌లను Excel బైట్స్/base64 గా మారుస్తుంది."""
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = 'Already_Existing_Skipped'

  export_cols = ['Skip_Reason', 'Existing_Sheet_Row'] + [
      h for h in headers if h
  ]
  header_font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
  header_fill = PatternFill(
      start_color='991B1B', end_color='991B1B', fill_type='solid'
  )
  ws.append(export_cols)
  for c_idx in range(1, len(export_cols) + 1):
    cell = ws.cell(row=1, column=c_idx)
    cell.font = header_font
    cell.fill = header_fill

  for item in existing_items:
    rec = item.get('record', {})
    row_out = [
        item.get('reason', 'Already exists in NEW THRIFT DATA'),
        item.get('existing_row', ''),
    ]
    for h in headers:
      if h:
        if h in ('Bank Account Number', 'RD 1', 'RD 2'):
          row_out.append(format_account_text(rec.get(h, '')))
        else:
          row_out.append(clean_str_cell(rec.get(h, '')))
    ws.append(row_out)

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return base64.b64encode(buf.read()).decode('utf-8')


@cf_bp.route('/cf/bulk-upload-thrift', methods=['POST'])
def cf_bulk_upload_thrift():
  """NEW THRIFT DATA లో Aadhar Number కీ కాలమ్ తో బల్క్ ఇన్సర్ట్ / అప్‌డేట్ చేస్తుంది."""
  if not session.get('cf_logged_in'):
    return (
        jsonify({'status': 'error', 'message': 'దయచేసి లాగిన్ అవ్వండి!'}),
        401,
    )

  try:
    from app import get_sheets_client, SPREADSHEET_NAME

    if 'bulk_file' not in request.files:
      return jsonify({
          'status': 'error',
          'message': 'దయచేసి Excel లేదా CSV ఫైల్ అప్‌లోడ్ చేయండి!',
      })

    f = request.files['bulk_file']
    if not f or not f.filename:
      return jsonify({'status': 'error', 'message': 'ఫైల్ ఎంచుకోలేదు!'})

    sheet_name = str(request.form.get('bulk_sheet', '')).strip()
    upload_mode = (
        str(request.form.get('upload_mode', 'insert_only')).strip().lower()
    )
    default_file_name = str(request.form.get('default_file_name', '')).strip()
    cf_name = str(session.get('cf_name') or '').strip().upper()

    if f.filename.lower().endswith('.csv'):
      df = pd.read_csv(f, dtype=str)
    else:
      xl = pd.ExcelFile(f)
      target_sheet = (
          sheet_name if sheet_name in xl.sheet_names else xl.sheet_names[0]
      )
      df = pd.read_excel(xl, sheet_name=target_sheet, dtype=str)

    if df.empty:
      return jsonify(
          {'status': 'error', 'message': 'అప్‌లోడ్ చేసిన ఫైల్‌లో డేటా లేదు!'}
      )

    # కాలమ్ పేర్లను సురక్షితంగా list గా మార్చడం
    clean_cols = []
    seen_cols = {}
    for idx_c, c in enumerate(df.columns.tolist()):
      c_name = str(c).strip() if c is not None else f'Col_{idx_c}'
      if not c_name or c_name.lower() == 'nan':
        c_name = f'Unnamed_{idx_c}'
      if c_name in seen_cols:
        seen_cols[c_name] += 1
        c_name = f'{c_name}_{seen_cols[c_name]}'
      else:
        seen_cols[c_name] = 0
      clean_cols.append(c_name)
    df.columns = clean_cols
    df_col_list = list(df.columns)

    aadhar_src_col = find_header_match(
        df_col_list,
        ['aadhar number', 'aadhaar number', 'aadhar', 'aadhaar', 'uid'],
        allow_partial=False,
    )
    if not aadhar_src_col:
      return jsonify({
          'status': 'error',
          'message': "ఫైల్‌లో 'Aadhar Number' కాలమ్ కనుగొనబడలేదు!",
      })

    client = get_sheets_client()
    spreadsheet = client.open(SPREADSHEET_NAME)
    ws = spreadsheet.worksheet('NEW THRIFT DATA')

    existing_all_rows = ws.get_all_values()
    if not existing_all_rows:
      return jsonify(
          {'status': 'error', 'message': "'NEW THRIFT DATA' షీట్ ఖాళీగా ఉంది!"}
      )

    headers = [str(h).strip() for h in existing_all_rows[0]]
    if 'Aadhar Number' not in headers:
      return jsonify({
          'status': 'error',
          'message': "'NEW THRIFT DATA' షీట్‌లో 'Aadhar Number' కాలమ్ లేదు!",
      })

    records_new = [
        {
            headers[i]: (r_vals[i] if i < len(r_vals) else '')
            for i in range(len(headers))
        }
        for r_vals in existing_all_rows[1:]
    ]

    aadhar_col_idx = headers.index('Aadhar Number')
    ifsc_to_file_map = _build_ifsc_to_file_map(records_new)

    email_map = {}
    try:
      ws_email = spreadsheet.worksheet('Bank_Email_Ids')
      for er in ws_email.get_all_records():
        ef = re.sub(
            r'\s+',
            ' ',
            clean_str_cell(
                er.get('FILE_NAME') or er.get('FILE NAME') or ''
            ).upper(),
        )
        em = clean_str_cell(er.get('OFFICIAL_EMAIL') or '')
        if ef and em:
          email_map[ef] = em
    except Exception:
      pass

    existing_aadhar_to_row = {}
    last_non_empty_row = 1
    for r_idx, r_vals in enumerate(existing_all_rows[1:], start=2):
      if any(str(cell).strip() != '' for cell in r_vals):
        last_non_empty_row = r_idx
      if aadhar_col_idx < len(r_vals):
        a_norm = normalize_aadhar(r_vals[aadhar_col_idx])
        if len(a_norm) == 12 and a_norm not in existing_aadhar_to_row:
          existing_aadhar_to_row[a_norm] = r_idx

    max_sno = 0
    if 'SNO' in headers:
      sno_idx = headers.index('SNO')
      for r_vals in existing_all_rows[1:]:
        if sno_idx < len(r_vals):
          s_txt = clean_str_cell(r_vals[sno_idx])
          if s_txt.isdigit():
            max_sno = max(max_sno, int(s_txt))

    max_bank_sno_by_file = {}
    max_bank_sno_global = 0
    if 'BANK_SNO' in headers:
      bsno_idx = headers.index('BANK_SNO')
      fname_idx = headers.index('FILE NAME') if 'FILE NAME' in headers else -1
      for r_vals in existing_all_rows[1:]:
        if bsno_idx < len(r_vals):
          b_txt = clean_str_cell(r_vals[bsno_idx])
          if b_txt.isdigit():
            b_int = int(b_txt)
            max_bank_sno_global = max(max_bank_sno_global, b_int)
            if fname_idx != -1 and fname_idx < len(r_vals):
              f_key = re.sub(
                  r'\s+', ' ', clean_str_cell(r_vals[fname_idx]).upper()
              )
              if f_key:
                max_bank_sno_by_file[f_key] = max(
                    max_bank_sno_by_file.get(f_key, 0), b_int
                )

    seen_in_file = set()
    invalid_rows = []
    existing_skipped_items = []
    to_insert_rows = []
    to_update_items = []

    for idx, row in df.iterrows():
      excel_row_no = idx + 2
      raw_a = row.get(aadhar_src_col, '')
      a_clean = normalize_aadhar(raw_a)

      if len(a_clean) != 12:
        invalid_rows.append({
            'row_no': excel_row_no,
            'raw_aadhar': clean_str_cell(raw_a),
            'name': clean_str_cell(row.get('Name Of Benificiary', '')),
            'reason': '12 అంకెల ఆధార్ నంబర్ లేదు',
        })
        continue

      if a_clean in seen_in_file:
        existing_skipped_items.append({
            'reason': f'Uploaded File Duplicate (Row {excel_row_no})',
            'existing_row': 'In Upload File',
            'record': {
                'Aadhar Number': a_clean,
                'Name Of Benificiary': clean_str_cell(
                    row.get('Name Of Benificiary', '')
                ),
            },
        })
        continue
      seen_in_file.add(a_clean)

      mapped_dict = {}
      for h in headers:
        matched_src = find_header_match(df_col_list, [h], allow_partial=False)
        if not matched_src:
          matched_src = find_header_match(df_col_list, [h], allow_partial=True)
        mapped_dict[h] = (
            clean_str_cell(row.get(matched_src, '')) if matched_src else ''
        )

      mapped_dict['Aadhar Number'] = a_clean

      # బ్యాంకు ఖాతా నంబర్, RD 1, RD 2 లకు ముందు సింగిల్ కోట్ (') జోడించడం
      for acc_col in ('Bank Account Number', 'RD 1', 'RD 2'):
        if acc_col in mapped_dict and mapped_dict[acc_col]:
          mapped_dict[acc_col] = format_account_text(mapped_dict[acc_col])

      clean_uhno, clean_rhno, clean_area = _clean_row_before_cf_mapping(
          mapped_dict
      )
      analyzed_cf = analyze_cf_name(clean_uhno, clean_area, clean_rhno)
      mapped_dict['NAME OF CF'] = analyzed_cf or cf_name

      res_file, res_bank, res_ifsc, res_email = _resolve_bank_file_and_ifsc(
          mapped_dict,
          records_new,
          email_map,
          cf_name=mapped_dict['NAME OF CF'],
          default_file_name=default_file_name,
          allow_single_cf_bank=True,
          ifsc_to_file_map=ifsc_to_file_map,
      )
      if res_file:
        mapped_dict['FILE NAME'] = res_file
      if res_bank:
        mapped_dict['Bank Name'] = res_bank
      if res_ifsc:
        mapped_dict['IFSC code'] = res_ifsc
      if res_email:
        mapped_dict['EMAIL ID'] = res_email

      if not mapped_dict.get('Caste'):
        mapped_dict['Caste'] = 'BC'
      if not mapped_dict.get('Profession'):
        mapped_dict['Profession'] = 'POWER LOOM WEAVER'
      if not mapped_dict.get('DISTRICT NAME'):
        mapped_dict['DISTRICT NAME'] = 'RAJANNA SIRCILLA'
      if not mapped_dict.get('MANDAL NAME'):
        mapped_dict['MANDAL NAME'] = 'SIRCILLA'
      if not mapped_dict.get('BANK SENT'):
        mapped_dict['BANK SENT'] = 'NOT SENT'

      if a_clean in existing_aadhar_to_row:
        sheet_r = existing_aadhar_to_row[a_clean]
        old_r_vals = (
            existing_all_rows[sheet_r - 1]
            if (sheet_r - 1) < len(existing_all_rows)
            else []
        )
        old_rec = {
            h: (old_r_vals[i] if i < len(old_r_vals) else '')
            for i, h in enumerate(headers)
        }
        old_rec['Aadhar Number'] = a_clean

        if upload_mode == 'update_existing':
          to_update_items.append({
              'sheet_row': sheet_r,
              'aadhar': a_clean,
              'old_vals': old_r_vals,
              'new_dict': mapped_dict,
          })
        else:
          existing_skipped_items.append({
              'reason': f'Already exists in NEW THRIFT DATA (Row {sheet_r})',
              'existing_row': sheet_r,
              'record': old_rec,
          })
      else:
        max_sno += 1
        mapped_dict['SNO'] = str(max_sno)

        if 'BANK_SNO' in headers and not clean_str_cell(
            mapped_dict.get('BANK_SNO', '')
        ):
          if res_file:
            next_bsno = max_bank_sno_by_file.get(res_file, 0) + 1
            max_bank_sno_by_file[res_file] = next_bsno
            max_bank_sno_global = max(max_bank_sno_global, next_bsno)
            mapped_dict['BANK_SNO'] = str(next_bsno)
          else:
            max_bank_sno_global += 1
            mapped_dict['BANK_SNO'] = str(max_bank_sno_global)

        new_row_list = []
        for h in headers:
          val_h = mapped_dict.get(h, '')
          if h in ('Bank Account Number', 'RD 1', 'RD 2'):
            new_row_list.append(format_account_text(val_h))
          else:
            new_row_list.append(val_h)
        to_insert_rows.append(new_row_list)

    updated_count = 0
    for upd in to_update_items:
      r_num = upd['sheet_row']
      a_val = upd['aadhar']
      old_vals = upd['old_vals']
      n_dict = upd['new_dict']

      backup_row_before_update(
          spreadsheet, headers, old_vals, r_num, cf_name, a_val
      )

      merged = list(old_vals[: len(headers)])
      while len(merged) < len(headers):
        merged.append('')

      for idx_h, h_name in enumerate(headers):
        if h_name == 'Aadhar Number':
          merged[idx_h] = a_val
          continue
        if h_name in ('Bank Account Number', 'RD 1', 'RD 2'):
          incoming_acc = clean_str_cell(n_dict.get(h_name, ''))
          if incoming_acc != '':
            merged[idx_h] = format_account_text(incoming_acc)
          elif merged[idx_h]:
            merged[idx_h] = format_account_text(merged[idx_h])
          continue
        incoming_v = clean_str_cell(n_dict.get(h_name, ''))
        if incoming_v != '':
          merged[idx_h] = incoming_v

      end_col_letter = get_column_letter(len(headers))
      ws.update(
          f'A{r_num}:{end_col_letter}{r_num}',
          [merged],
          value_input_option='USER_ENTERED',
      )
      updated_count += 1

    inserted_count = 0
    if to_insert_rows:
      start_row = last_non_empty_row + 1
      end_col_letter = get_column_letter(len(headers))
      end_row = start_row + len(to_insert_rows) - 1
      ws.update(
          f'A{start_row}:{end_col_letter}{end_row}',
          to_insert_rows,
          value_input_option='USER_ENTERED',
      )
      inserted_count = len(to_insert_rows)

    _cache_invalidate('ben:')
    _cache_invalidate('bank_files')

    existing_b64 = (
        _build_existing_records_excel_b64(existing_skipped_items, headers)
        if existing_skipped_items
        else ''
    )

    msg = (
        f'పూర్తయింది! కొత్తవి Insert: {inserted_count} | '
        f'Update: {updated_count} | '
        f'ఇదివరకు ఉన్నవి (Skipped): {len(existing_skipped_items)} | '
        f'చెల్లనివి: {len(invalid_rows)}'
    )

    return jsonify({
        'status': 'success',
        'message': msg,
        'inserted_count': inserted_count,
        'updated_count': updated_count,
        'existing_skipped_count': len(existing_skipped_items),
        'invalid_count': len(invalid_rows),
        'existing_excel_b64': existing_b64,
        'existing_filename': (
            f"Existing_Records_Skipped_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        ),
    })

  except Exception as e:
    return (
        jsonify({'status': 'error', 'message': f'బల్క్ ప్రాసెస్ లోపం: {str(e)}'}),
        500,
    )
