# utils.py
"""通用的資料清洗、驗證與 Excel 讀取輔助函式。

此模組集中放置所有比對規則模組（compare/*.py）共用的邏輯，
例如：安全型別轉換、字串組合、資料期間格式驗證，
以及 Excel 中繼資料（表頭上方標籤）驗證與資料表格讀取。
"""
from dataclasses import dataclass
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


def safe_str(val: Any) -> str:
  """將任意值安全轉換為去除前後空白的字串，若為 NaN/None 則回傳空字串。"""
  if pd.isna(val):
    return ""
  return str(val).strip()


def safe_int(val: Any) -> int:
  """將任意值安全轉換為 int，若為 NaN/None 則回傳 0。"""
  if pd.isna(val):
    return 0
  return int(val)


def safe_float(val: Any) -> float:
  """將任意值安全轉換為 float，若為 NaN/None 則回傳 0.0。"""
  if pd.isna(val):
    return 0.0
  return float(val)


def format_columns(*values: Any) -> str:
  """自動過濾空白/NaN，用空格組合並自動清除前後多餘空白。"""
  parts = [safe_str(v) for v in values if safe_str(v)]
  return " ".join(parts)


def validate_and_extract_period(
    date_str: Any,
) -> Tuple[bool, str, Optional[str], Optional[str]]:
  """驗證資料期間，支援 ~ 或是 至 隔開的日期格式，並取出起訖的 yyyymmdd。

  Args:
    date_str: 原始儲存格內容，例如 "2024/01/01~2024/01/31"。

  Returns:
    (是否驗證通過, 說明訊息, 起始日 yyyymmdd 或 None, 結束日 yyyymmdd 或 None)
  """
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


# ---------------------------------------------------------------------------
# 以下為 Excel 讀取與結構驗證共用工具
# 供 compare/*.py 內的 process_Old / process_New 等解析函式共用，
# 避免每個比對規則都重複撰寫相同的讀檔、驗證邏輯。
# ---------------------------------------------------------------------------


@dataclass
class MetadataField:
  """描述 Excel 工作表中需要驗證的中繼標籤座標。

  Attributes:
    col_name: 預期出現在儲存格內容中的關鍵字（例如 "分公司"、"排名"）。
    x: 欄索引（column，從 0 開始）。
    y: 列索引（row，從 0 開始）。
  """

  col_name: str
  x: int
  y: int


def read_all_sheets_raw(file_path: str) -> Dict[str, pd.DataFrame]:
  """以不帶 header 的方式讀取 Excel 檔案內所有工作表。

  用於後續檢查工作表上方的中繼資訊（標籤、儲存格座標）。

  Raises:
    ValueError: 當檔案內容為空或無法讀取任何工作表時拋出。
  """
  all_sheets_raw = pd.read_excel(file_path, sheet_name=None, header=None)
  if not all_sheets_raw:
    raise ValueError("錯誤：Excel 檔案內容為空或無法讀取。")
  return all_sheets_raw


def get_cell_value(df_raw: pd.DataFrame, x: int, y: int) -> str:
  """安全取得指定座標 (x=欄, y=列) 的儲存格內容，超出範圍時回傳空字串。"""
  if y < df_raw.shape[0] and x < df_raw.shape[1]:
    return str(df_raw.iloc[y, x])
  return ""


def validate_sheet_structure(
    df_raw: pd.DataFrame,
    sheet_name: str,
    fields: List[MetadataField],
) -> None:
  """依序驗證工作表指定座標是否包含必要的標籤/欄位文字。

  Args:
    df_raw: 以 header=None 讀取的原始工作表 DataFrame。
    sheet_name: 工作表名稱（用於錯誤訊息顯示）。
    fields: 需要驗證的座標與關鍵字清單。

  Raises:
    ValueError: 當列數不足，或任一座標內容不含指定關鍵字時拋出。
  """
  min_required_rows = max(field.y for field in fields) + 1
  if df_raw.shape[0] < min_required_rows:
    raise ValueError(f"工作表【{sheet_name}】的列數不足，不符合報表格式。")

  for field in fields:
    cell_value = get_cell_value(df_raw, field.x, field.y)
    if field.col_name not in cell_value:
      raise ValueError(
          f"工作表【{sheet_name}】結構錯誤（座標 x={field.x}, y={field.y} 的內容"
          f" '{cell_value}' 不包含必要的標籤/欄位 '{field.col_name}'）。"
      )


def read_data_table(file_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
  """讀取工作表下方的正式資料表格（略過上方中繼資訊區塊），並清除全空列。

  Args:
    file_path: Excel 檔案路徑。
    sheet_name: 要讀取的工作表名稱。
    header_row: 表頭所在的列索引（0 起算），會傳給 pandas 的 header 參數。

  Returns:
    已清除全空列、所有欄位皆以字串型態讀入的資料表格 DataFrame。
  """
  df_table = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype=str)
  return df_table.dropna(how="all")
