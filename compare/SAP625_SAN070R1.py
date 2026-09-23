from typing import Any, Dict, List
import pandas as pd
import re
from dataclasses import dataclass, asdict

from company_helper import get_company_id
from utils import format_columns, safe_float, safe_int, safe_str, validate_and_extract_period

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
  
def process_Old(file_path: str) -> list[DealerRecord]:
    # 1. 以不帶 header 方式讀取所有分頁，用來檢查上方中繼標籤與儲存格
    all_sheets_raw = pd.read_excel(file_path, sheet_name=None, header=None)
    if not all_sheets_raw:
        raise ValueError("錯誤：Excel 檔案內容為空或無法讀取。")

    all_records = []

    # 定義要比對的欄位/標籤物件清單 {colName, x, y}
    # 註：x 代表行索引 (欄)，y 代表列索引 (列)
    metadata_check_fields = [
        {"colName": "分公司", "x": 0, "y": 3},   # A4 儲存格
        {"colName": "部門", "x": 0, "y": 4},     # A5 儲存格
        {"colName": "排名", "x": 0, "y": 6},     # 表格欄位：排名 (第 7 行)
        {"colName": "級數", "x": 1, "y": 6},     # 表格欄位：級數
        {"colName": "經銷商", "x": 2, "y": 6},   # 表格欄位：經銷商
        {"colName": "業務員", "x": 4, "y": 6},   # 表格欄位：業務員
        {"colName": "銷售數", "x": 6, "y": 6},   # 表格欄位：銷售數
        {"colName": "銷售金額", "x": 7, "y": 6}  # 表格欄位：銷售金額
    ]

    # 2. 開始跑 Worksheets (工作表) 迴圈
    for sheet_name, df_raw in all_sheets_raw.items():
        if sheet_name == "form":
            continue  # 跳過 form 分頁

        print(f"--- 正在檢核工作表: {sheet_name} ---")

        # 確保該分頁有足夠的列數可以抓取上方中繼資訊與表頭

        # 自動計算 metadata_check_fields 中需要的最小列數 (最大 y 索引 + 1)
        min_required_rows = max(field["y"] for field in metadata_check_fields) + 1
        if df_raw.shape[0] < min_required_rows:
            raise ValueError(
                f"工作表【{sheet_name}】的列數不足，不符合 SAP 報表格式。"
            )

        # 3. 透過物件清單的 (x, y) 座標動態驗證內容是否包含指定的 colName
        for field in metadata_check_fields:
            col_name = field["colName"]
            x = field["x"]
            y = field["y"]
            
            cell_value = str(df_raw.iloc[y, x]) if y < df_raw.shape[0] and x < df_raw.shape[1] else ""

            if col_name not in cell_value:
                raise ValueError(
                    f"工作表【{sheet_name}】結構錯誤（座標 x={x}, y={y} 的內容"
                    f" '{cell_value}' 不包含必要的標籤/欄位 '{col_name}'）。"
                )

        # 4. 抓取中繼數值：分公司 (3, 1)、部門 (4, 1)、資料期間 (4, 8) [即 I5]
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

        # 5. 嚴格驗證「資料期間」格式並轉換出起訖 yyyymmdd
        is_valid, msg, start_date, end_date = validate_and_extract_period(
            date_period_raw
        )
        if not is_valid:
            raise ValueError(f"工作表【{sheet_name}】{msg}")

        print(
            f" ✓ 中繼驗證通過 | 分公司: {company_branch} | 部門: {department} |"
            f" 期間: {start_date}~{end_date}"
        )

        # 6. 讀取該分頁下方的正式表格 (指定 header=6 略過上方雜訊)
        df_table = pd.read_excel(file_path, sheet_name=sheet_name, header=6, dtype=str)
        df_table = df_table.dropna(how="all")  # 清理全空列

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
                dealer_name=str(format_columns(row.iloc[2], row.iloc[3])),
                sales_person=str(get_company_id(company_name) + str(row.iloc[4]) + " " + str(row.iloc[5])), 
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
  # 定義要比對的欄位/標籤物件清單 {colName, x, y}
  metadata_check_fields = [
      {"colName": "級數", "x": 4, "y": 7},     # 級數標籤 (E8 儲存格)
      {"colName": "分公司", "x": 0, "y": 10},   # 表格欄位：分公司 (第 11 行)
      {"colName": "部門", "x": 1, "y": 10},     # 表格欄位：部門
      {"colName": "業務員", "x": 2, "y": 10},   # 表格欄位：業務員
      {"colName": "經銷商", "x": 3, "y": 10},   # 表格欄位：經銷商
      {"colName": "銷售數", "x": 5, "y": 10},   # 表格欄位：銷售數
      {"colName": "銷售金額", "x": 6, "y": 10}  # 表格欄位：銷售金額
  ]
  # 自動計算該報表所需的最小列數 (最大 y 索引 + 1)
  min_required_rows = max(field["y"] for field in metadata_check_fields) + 1
  # 2. 開始跑 Worksheets (工作表) 迴圈
  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
        continue  # 跳過 form 分頁
    print(f"--- 正在檢核工作表: {sheet_name} ---")
    # 確保該分頁有足夠的列數可以抓取上方中繼資訊與表頭
    if df_raw.shape[0] < min_required_rows:
        raise ValueError(
            f"工作表【{sheet_name}】的列數不足，不符合 SAN070R1 報表格式。"
        )
    # 3. 透過物件清單的 (x, y) 座標統一進行結構驗證
    for field in metadata_check_fields:
        col_name = field["colName"]
        x = field["x"]
        y = field["y"]
        
        cell_value = str(df_raw.iloc[y, x]) if y < df_raw.shape[0] and x < df_raw.shape[1] else ""
        if col_name not in cell_value:
            raise ValueError(
                f"工作表【{sheet_name}】結構錯誤（座標 x={x}, y={y} 的內容"
                f" '{cell_value}' 不包含必要的標籤/欄位 '{col_name}'）。"
            )
    # 4. 抓取中繼數值：資料期間 (7, 5)
    date_period_raw = df_raw.iloc[7, 5]
    # 5. 統一判斷中繼欄位是否為空白
    if pd.isna(date_period_raw):
        raise ValueError(f"工作表【{sheet_name}】上方有必要中繼資訊為空！")
    # 6. 若不是空白，開始驗證資料期間格式
    is_valid, msg, start_date, end_date = validate_and_extract_period(date_period_raw)
    if not is_valid:
        raise ValueError(f"工作表【{sheet_name}】{msg}")
    # 7. 若期間正確，顯示中繼驗證通過（格式與風格統一）
    print(
        f" ✓ 中繼驗證通過 | 期間: {start_date}~{end_date}"
    )
  # 5. 讀取該分頁下方的正式表格 (指定 header=6 略過上方雜訊)
  df_table = pd.read_excel(file_path, sheet_name=sheet_name, header=10,dtype=str)
  df_table = df_table.dropna(how="all")  # 清理全空列
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
    """比對邏輯主函式：只抓出單邊未匹配的純淨資料"""
    print(f"--- 正在執行：{RULE_NAME} ---")

    try:
        df_old = process_Old(old_file_path)
        df_new = process_New(new_file_path)

        if not df_old or not df_new:
            return "比對失敗：舊專案或新專案的表格內容為空。", [], []

        df_old_pd = pd.DataFrame([vars(r) for r in df_old])
        df_new_pd = pd.DataFrame([vars(r) for r in df_new])

        # 定義雙方的身分識別 Key（用來判斷是不是同一筆）
        all_columns = [
        'company_branch', 'department', 'start_date', 'end_date', 
        'level', 'dealer_name', 'sales_person', 'sales_amount', 'sales_count'
        ]

        # --- 1. 找出「舊檔有，但新檔沒有」的資料 ---
        # 以 df_old 為主體，去對照 df_new 的 Key
        merged_left = pd.merge(df_old_pd, df_new_pd[all_columns], on=all_columns, how='left', indicator=True)
        missing_in_new_df = merged_left[merged_left['_merge'] == 'left_only'].drop(columns=['_merge'])

        # --- 2. 找出「新檔有，但舊檔沒有」的資料 ---
        # 以 df_new 為主體，去對照 df_old 的 Key
        merged_right = pd.merge(df_new_pd, df_old_pd[all_columns], on=all_columns, how='left', indicator=True)
        missing_in_old_df = merged_right[merged_right['_merge'] == 'left_only'].drop(columns=['_merge'])

        # 3. 執行 merge，找出兩邊完全一致的資料
        merged_both = pd.merge(
            df_old_pd, 
            df_new_pd[all_columns], 
            on=all_columns, 
            how='left',       # 因為只要兩邊都有的，用 inner 也可以，或者維持 how='left' 搭配過濾
            indicator=True
        )
        matched_df = merged_both[merged_both['_merge'] == 'both'].drop(columns=['_merge'])

        # 4. 產出統計訊息
        result_msg = f"""--- 獨立差異比對結果 ---
        舊檔獨有（新檔缺漏）筆數: {len(missing_in_new_df)}
        新檔獨有（舊檔多出）筆數: {len(missing_in_old_df)}
        兩邊完全一致筆數: {len(matched_df)}"""

        return (
            result_msg, 
            missing_in_new_df.to_dict(orient='records'), 
            missing_in_old_df.to_dict(orient='records')
        )

    except Exception as e:
        return f"讀取或比對過程發生錯誤: {e}", [], []