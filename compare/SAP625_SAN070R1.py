from typing import Any, Dict, List
import pandas as pd
import re
from dataclasses import dataclass, asdict

from company_helper import get_company_id

# 🌟 這裡設定要顯示在介面下拉選單中的名稱
RULE_NAME = "SAP625/SAN070R1比對"

# 1. 定義你要儲存的物件結構
@dataclass
class DealerRecord:
  sheet_name: str  # 來源工作表名稱 (例如: "工作表2")
  row_Index:int  # 該筆資料在原始 Excel 表格的列索引 (從 0 開始計算)
  company_branch: str  # 分公司名稱 (從 Excel B4 / iloc[3, 1] 抓取)
  department: str  # 部門名稱 (從 Excel B5 / iloc[4, 1] 抓取)
  start_date: str  # 資料期間-起始日 (從 I5 抓取並解析後格式化為 yyyymmdd)
  end_date: str  # 資料期間-結束日 (從 I5 抓取並解析後格式化為 yyyymmdd)
  ranking: int  # 排名 (對應 Excel 表格「排名」欄位)
  level: str  # 級數 (對應 Excel 表格「級數」欄位)
  dealer_name: str  # 經銷商 (對應 Excel 表格「經銷商」欄位，如 10097 生華)
  sales_person: str  # 業務員 (對應 Excel 表格「業務員」欄位，如 00942 范誌宏)
  sales_amount: float  # 銷售金額 (對應 Excel 表格「銷售金額」欄位，數值型態)
  sales_count: int = 0  # 銷售筆數 (對應 Excel 表格「銷售筆數」欄位，數值型態，預設為 0)
  

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
    
    # 1. 第一種嘗試：尋找標準的 yyyy/MM/dd~yyyy/MM/dd 格式
    pattern_tilde = r"(\d{4})/(\d{2})/(\d{2})~(\d{4})/(\d{2})/(\d{2})"
    match = re.search(pattern_tilde, date_str)
    
    # 2. 如果第一種沒找到，嘗試第二種：包含「至」的格式 (例如 2024/01/01 至 2024/01/31)
    if not match:
        pattern_zhi = r"(\d{4})/(\d{2})/(\d{2})\s*至\s*(\d{4})/(\d{2})/(\d{2})"
        match = re.search(pattern_zhi, date_str)

    # 3. 如果兩種都找不到，回報錯誤
    if not match:
        return (
            False,
            f"日期格式不符 (找不到起訖區間)，實際: {date_str}",
            None,
            None,
        )

    start_y, start_m, start_d, end_y, end_m, end_d = match.groups()
    start_yyyymmdd = f"{start_y}{start_m}{start_d}"
    end_yyyymmdd = f"{end_y}{end_m}{end_d}"

    return True, "格式正確", start_yyyymmdd, end_yyyymmdd


def process_Old(file_path: str) -> list[DealerRecord]:
  # 1. 以不帶 header 方式讀取所有分頁，用來檢查上方中繼標籤與儲存格
  all_sheets_raw = pd.read_excel(file_path, sheet_name=None, header=None)
  if not all_sheets_raw:
    raise ValueError("錯誤：Excel 檔案內容為空或無法讀取。")

  all_records = []
  required_columns = ["排名", "級數", "經銷商", "業務員", "銷售數", "銷售金額"]

  # 2. 開始跑 Worksheets (工作表) 迴圈
  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
      continue  # 跳過 form 分頁

    print(f"--- 正在檢核工作表: {sheet_name} ---")

    # 確保該分頁有足夠的列數可以抓取上方中繼資訊
    if df_raw.shape[0] < 6:
      raise ValueError(
          f"工作表【{sheet_name}】的列數不足，不符合 SAP 報表格式。"
      )

    # 3. 驗證上方固定標籤與抓取格子值
    # A4 欄位標籤 (index [3, 0]), 實際值在隔壁或同列
    branch_label = str(df_raw.iloc[3, 0])
    dept_label = str(df_raw.iloc[4, 0])

    if "分公司" not in branch_label or "部門" not in dept_label:
      raise ValueError(
          f"工作表【{sheet_name}】的表頭結構錯誤（找不到分公司或部門標籤）。"
      )

    # 抓取中繼數值：分公司 (3, 1)、部門 (4, 1)、資料期間 (4, 8) [即 I5]
    company_branch = df_raw.iloc[3, 1]
    department = df_raw.iloc[4, 1]
    date_period_raw = df_raw.iloc[4, 8]
    company_name = df_raw.iloc[1, 0]

    # 檢查中繼欄位是否為空
    if (
        pd.isna(company_branch)
        or pd.isna(department)
        or pd.isna(date_period_raw)
    ):
      raise ValueError(f"工作表【{sheet_name}】上方有必要中繼資訊為空！")

    # 4. 嚴格驗證「資料期間」格式並轉換出起訖 yyyymmdd
    is_valid, msg, start_date, end_date = validate_and_extract_period(
        date_period_raw
    )
    if not is_valid:
      raise ValueError(f"工作表【{sheet_name}】{msg}")

    print(
        f" ✓ 中繼驗證通過 | 分公司: {company_branch} | 部門: {department} |"
        f" 期間: {start_date}~{end_date}"
    )

    # 5. 讀取該分頁下方的正式表格 (指定 header=6 略過上方雜訊)
    df_table = pd.read_excel(file_path, sheet_name=sheet_name, header=6,dtype=str)
    df_table = df_table.dropna(how="all")  # 清理全空列

    # 6. 檢核正式表格是否包含必要的欄位名稱
    for col in required_columns:
      if col not in df_table.columns:
        raise ValueError(
            f"工作表【{sheet_name}】缺少必要的資料欄位【{col}】，請檢查表頭是否在第"
            " 7 行。"
        )

    # 7. 逐行將資料打包成 DealerRecord 物件並加入清單
    for idx, row in df_table.iterrows():
      if row.iloc[2:6].isna().all():
        continue

      record = DealerRecord(
          sheet_name=sheet_name,
          row_Index=idx,
          company_branch=str(company_branch),
          department=str(department),
          start_date=start_date,
          end_date=end_date,
          ranking=str("" if pd.isna(row["排名"]) else str(row["排名"])),
          level=str(row["級數"]) if not pd.isna(row["級數"]) else "",
          dealer_name=str(format_columns(row.iloc[2],row.iloc[3])),
          sales_person=str(get_company_id(company_name) + str(row.iloc[4]) + " " + str(row.iloc[5])),  # 假設業務員資訊在第 5、6 欄
          sales_amount=safe_float(row.iloc[7]),
          sales_count=safe_int(row.iloc[6]),
      )
      all_records.append(record)
  return all_records

# --- 測試範例 ---
# text = "2024/01/01~2024/01/31"
# is_valid, msg, start_date, end_date = validate_and_extract_period(text)
# print(is_valid, msg)
# print(f"開始日期: {start_date}")  # 輸出: 20240101
# print(f"結束日期: {end_date}")    # 輸出: 20240131


def process_New(file_path: str) -> list[DealerRecord]:
  # 1. 以不帶 header 方式讀取所有分頁，用來檢查上方中繼標籤與儲存格
  all_sheets_raw = pd.read_excel(file_path, sheet_name=None, header=None)
  if not all_sheets_raw:
    raise ValueError("錯誤：Excel 檔案內容為空或無法讀取。")

  all_records = []
  required_columns = ["分公司", "部門", "業務員", "經銷商", "銷售金額", "銷售數", "銷售金額"]

  # 2. 開始跑 Worksheets (工作表) 迴圈
  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
      continue  # 跳過 form 分頁

    print(f"--- 正在檢核工作表: {sheet_name} ---")

    # 確保該分頁有足夠的列數可以抓取上方中繼資訊
    if df_raw.shape[0] < 6:
      raise ValueError(
          f"工作表【{sheet_name}】的列數不足，不符合 SAN070R1 報表格式。"
      )

    # 3. 驗證上方固定標籤與抓取格子值 (根據你的截圖位置)
    # A4 欄位標籤 (index [3, 0]), 實際值在隔壁或同列
    level_label = str(df_raw.iloc[7,4])

    if "級數" not in level_label:
      raise ValueError(
          f"工作表【{sheet_name}】的表頭結構錯誤（找不到級數標籤）。"
      )

    # 抓取中繼數值：資料期間 (3, 0)
    date_period_raw = df_raw.iloc[3, 0]

    # 檢查中繼欄位是否為空
    if (
       pd.isna(date_period_raw)
    ):
      raise ValueError(f"工作表【{sheet_name}】上方有必要中繼資訊為空！")

    # 4. 嚴格驗證「資料期間」格式並轉換出起訖 yyyymmdd
    is_valid, msg, start_date, end_date = validate_and_extract_period(
        date_period_raw
    )
    if not is_valid:
      raise ValueError(f"工作表【{sheet_name}】{msg}")

    print(
        f" ✓ 期間: {start_date}~{end_date}"
    )

    # 5. 讀取該分頁下方的正式表格 (指定 header=6 略過上方雜訊)
    df_table = pd.read_excel(file_path, sheet_name=sheet_name, header=10,dtype=str)
    df_table = df_table.dropna(how="all")  # 清理全空列

    # 6. 檢核正式表格是否包含必要的欄位名稱
    for col in required_columns:
      if col not in df_table.columns:
        raise ValueError(
            f"工作表【{sheet_name}】缺少必要的資料欄位【{col}】，請檢查表頭是否在第"
            " 9 行。"
        )

    # 7. 逐行將資料打包成 DealerRecord 物件並加入清單
    for idx, row in df_table.iterrows():
      if row.iloc[0:4].isna().all():
        continue

      record = DealerRecord(
          sheet_name=sheet_name,
          row_Index=idx,
          company_branch=safe_str(row.iloc[0]),
          department=safe_str(row.iloc[1]),
          start_date=start_date,
          end_date=end_date,
          ranking=0,  # 新檔沒有排名欄位，先設為 0
          level=safe_str(row.iloc[4]),
          dealer_name=safe_str(row["經銷商"]),
          sales_person=safe_str(row["業務員"]), 
          sales_amount=safe_float(row.iloc[6]),
          sales_count=safe_int(row.iloc[5]),
      )
      all_records.append(record)

  return all_records

# --- 測試範例 ---
# text = "2024/01/01~2024/01/31"
# is_valid, msg, start_date, end_date = validate_and_extract_period(text)
# print(is_valid, msg)
# print(f"開始日期: {start_date}")  # 輸出: 20240101
# print(f"結束日期: {end_date}")    # 輸出: 20240131


def compare(old_file_path: str, new_file_path: str):
  """比對邏輯主函式：接收檔案路徑，內部自動處理多工作表與單工作表"""
  print(f"--- 正在執行：{RULE_NAME} ---")

  try:
    # 1. 透過 process_Old 取得整理好的舊檔 DataFrame（含多分頁合併）
    df_old = process_Old(old_file_path)
    
    # 2. 處理新檔（假設新檔是單一工作表，同樣需要指定 header=6）
    df_new = process_New(new_file_path)

    print (f"成功讀取檔案！舊列數: {len(df_old)} | 新列數: {len(df_new)}")

    if not df_old or not df_new:
      return "比對失敗：舊專案或新專案的表格內容為空。", [], []

    # 3. 接下來就可以直接套用剛剛的全欄位嚴格比對！
    all_columns = [
        'company_branch', 'department', 'start_date', 'end_date', 
        'level', 'dealer_name', 'sales_person', 'sales_amount', 'sales_count'
    ]

    # 如果 df_old 和 df_new 是物件清單，必須先轉成 DataFrame：
    df_old_pd = pd.DataFrame([vars(r) for r in df_old])
    df_new_pd = pd.DataFrame([vars(r) for r in df_new])

    merged_df = pd.merge(df_old_pd, df_new_pd, on=all_columns, how='outer', indicator=True)

    missing_in_new = merged_df[merged_df['_merge'] == 'left_only']
    missing_in_old = merged_df[merged_df['_merge'] == 'right_only']
    matched_rows = merged_df[merged_df['_merge'] == 'both']

    # 4. 產出訊息與回傳清單
    result_msg = f"""--- 全欄位嚴格比對結果 ---
    完全一致的筆數: {len(matched_rows)}
    舊檔有但新檔沒有的筆數: {len(missing_in_new)}
    新檔有但舊檔沒有的筆數: {len(missing_in_old)}"""

    return result_msg, missing_in_new.to_dict(orient='records'), missing_in_old.to_dict(orient='records')

  except Exception as e:
    return f"讀取或比對過程發生錯誤: {e}"