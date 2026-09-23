# utils.py
import re
import pandas as pd


def safe_str(val) -> str:
  if pd.isna(val):
    return ""
  return str(val).strip()


def safe_int(val) -> int:
  if pd.isna(val):
    return 0
  return int(val)


def safe_float(val) -> float:
  if pd.isna(val):
    return 0.0
  return float(val)


def format_columns(*values) -> str:
  """自動過濾空白/NaN，用空格組合並自動清除前後多餘空白"""
  parts = [safe_str(v) for v in values if safe_str(v)]
  return " ".join(parts)


def validate_and_extract_period(date_str: str):
  """驗證資料期間，支援 ~ 或是 至 隔開的日期格式，並取出起訖的 yyyymmdd"""
  if pd.isna(date_str):
    return False, "資料期間欄位為空", None, None

  date_str = str(date_str).strip()

  pattern_tilde = r"(\d{4})/(\d{2})/(\d{2})~(\d{4})/(\d{2})/(\d{2})"
  match = re.search(pattern_tilde, date_str)

  if not match:
    pattern_zhi = r"(\d{4})/(\d{2})/(\d{2})\s*至\s*(\d{4})/(\d{2})/(\d{2})"
    match = re.search(pattern_zhi, date_str)

  if not match:
    return False, f"日期格式不符 (找不到起訖區間)，實際: {date_str}", None, None

  start_y, start_m, start_d, end_y, end_m, end_d = match.groups()
  return True, "格式正確", f"{start_y}{start_m}{start_d}", f"{end_y}{end_m}{end_d}"