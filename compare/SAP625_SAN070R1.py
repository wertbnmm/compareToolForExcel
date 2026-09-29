"""SAP625（舊系統）與 SAN070R1（新系統）經銷商業績報表比對規則。

此模組負責：
  1. 解析 SAP625 (process_Old) 與 SAN070R1 (process_New) 兩種 Excel 報表格式，
     並將每一列資料轉換為統一的 DealerRecord 物件。
  2. 透過 compare() 對外提供統一介面，同時從「資料列」與「經銷商」兩個維度
     分析新舊報表差異，並將差異明細匯出為 Markdown 報表。

Excel 讀取、中繼資料驗證等共用邏輯已抽出至 utils.py，
本模組只保留「SAP625/SAN070R1 專屬」的欄位座標設定與組裝邏輯。
"""
from dataclasses import dataclass
from typing import Any, List, Tuple

import pandas as pd

from company_helper import get_company_id
from resultExport import export_records_report
from utils import (
    MetadataField,
    format_columns,
    read_all_sheets_raw,
    read_data_table,
    safe_float,
    safe_int,
    safe_str,
    validate_and_extract_period,
    validate_sheet_structure,
)

# 🌟 這裡設定要顯示在介面下拉選單中的名稱
RULE_NAME = "SAP625/SAN070R1比對"


# 1. 定義你要儲存的物件結構
@dataclass
class DealerRecord:
  """單一經銷商業績資料列（舊/新兩種報表解析後皆會統一成此結構）。"""

  sheet_name: str  # 來源工作表名稱 (例如: "工作表2")
  row_Index: int  # 該筆資料在原始 Excel 表格的列索引 (從 0 開始計算)
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


# SAP625（舊檔）上方中繼標籤與表頭座標定義
# x 代表欄索引，y 代表列索引
_OLD_METADATA_FIELDS: List[MetadataField] = [
    MetadataField("分公司", x=0, y=3),  # A4 儲存格
    MetadataField("部門", x=0, y=4),  # A5 儲存格
    MetadataField("排名", x=0, y=6),  # 表格欄位：排名 (第 7 行)
    MetadataField("級數", x=1, y=6),  # 表格欄位：級數
    MetadataField("經銷商", x=2, y=6),  # 表格欄位：經銷商
    MetadataField("業務員", x=4, y=6),  # 表格欄位：業務員
    MetadataField("銷售數", x=6, y=6),  # 表格欄位：銷售數
    MetadataField("銷售金額", x=7, y=6),  # 表格欄位：銷售金額
]
_OLD_TABLE_HEADER_ROW = 6  # 正式資料表格表頭所在列 (略過上方雜訊)

# SAN070R1（新檔）上方中繼標籤與表頭座標定義
_NEW_METADATA_FIELDS: List[MetadataField] = [
    MetadataField("級數", x=4, y=7),  # 級數標籤 (E8 儲存格)
    MetadataField("分公司", x=0, y=10),  # 表格欄位：分公司 (第 11 行)
    MetadataField("部門", x=1, y=10),  # 表格欄位：部門
    MetadataField("業務員", x=2, y=10),  # 表格欄位：業務員
    MetadataField("經銷商", x=3, y=10),  # 表格欄位：經銷商
    MetadataField("銷售數", x=5, y=10),  # 表格欄位：銷售數
    MetadataField("銷售金額", x=6, y=10),  # 表格欄位：銷售金額
]
_NEW_TABLE_HEADER_ROW = 10  # 正式資料表格表頭所在列 (略過上方雜訊)


def _extract_period_or_raise(date_period_raw: Any, sheet_name: str) -> Tuple[str, str]:
  """驗證並取出資料期間的起訖日 (yyyymmdd)，驗證失敗時拋出 ValueError。"""
  is_valid, msg, start_date, end_date = validate_and_extract_period(date_period_raw)
  if not is_valid:
    raise ValueError(f"工作表【{sheet_name}】{msg}")
  return start_date, end_date


def process_Old(file_path: str) -> List[DealerRecord]:
  """解析 SAP625（舊系統）匯出的 Excel 報表，回傳 DealerRecord 清單。

  Args:
    file_path: SAP625 報表的 Excel 檔案路徑。

  Returns:
    解析後的經銷商業績資料清單。

  Raises:
    ValueError: 當工作表結構不符合預期格式時拋出。
  """
  # 1. 以不帶 header 方式讀取所有分頁，用來檢查上方中繼標籤與儲存格
  all_sheets_raw = read_all_sheets_raw(file_path)
  all_records: List[DealerRecord] = []

  # 2. 開始跑 Worksheets (工作表) 迴圈
  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
      continue  # 跳過 form 分頁

    print(f"--- 正在檢核工作表: {sheet_name} ---")

    # 3. 透過共用工具驗證上方中繼標籤與表頭欄位是否存在
    validate_sheet_structure(df_raw, sheet_name, _OLD_METADATA_FIELDS)

    # 4. 抓取中繼數值：分公司 (3, 1)、部門 (4, 1)、資料期間 (4, 8) [即 I5]
    company_branch = df_raw.iloc[3, 1]
    department = df_raw.iloc[4, 1]
    date_period_raw = df_raw.iloc[4, 8]
    company_name = df_raw.iloc[1, 0]

    # 檢查中繼欄位是否為空
    if pd.isna(company_branch) or pd.isna(department) or pd.isna(date_period_raw):
      raise ValueError(f"工作表【{sheet_name}】上方有必要中繼資訊為空！")

    # 5. 嚴格驗證「資料期間」格式並轉換出起訖 yyyymmdd
    start_date, end_date = _extract_period_or_raise(date_period_raw, sheet_name)

    print(
        f" ✓ 中繼驗證通過 | 分公司: {company_branch} | 部門: {department} |"
        f" 期間: {start_date}~{end_date}"
    )

    # 6. 讀取該分頁下方的正式表格 (略過上方雜訊)
    df_table = read_data_table(file_path, sheet_name, _OLD_TABLE_HEADER_ROW)

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


def process_New(file_path: str) -> List[DealerRecord]:
  """解析 SAN070R1（新系統）匯出的 Excel 報表，回傳 DealerRecord 清單。

  Args:
    file_path: SAN070R1 報表的 Excel 檔案路徑。

  Returns:
    解析後的經銷商業績資料清單。

  Raises:
    ValueError: 當工作表結構不符合預期格式時拋出。
  """
  # 1. 以不帶 header 方式讀取所有分頁，用來檢查上方中繼標籤與儲存格
  all_sheets_raw = read_all_sheets_raw(file_path)
  all_records: List[DealerRecord] = []

  # 2. 開始跑 Worksheets (工作表) 迴圈
  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
      continue  # 跳過 form 分頁

    print(f"--- 正在檢核工作表: {sheet_name} ---")

    # 3. 透過共用工具統一進行結構驗證
    validate_sheet_structure(df_raw, sheet_name, _NEW_METADATA_FIELDS)

    # 4. 抓取中繼數值：資料期間 (7, 5)
    date_period_raw = df_raw.iloc[7, 5]

    # 5. 統一判斷中繼欄位是否為空白
    if pd.isna(date_period_raw):
      raise ValueError(f"工作表【{sheet_name}】上方有必要中繼資訊為空！")

    # 6. 若不是空白，開始驗證資料期間格式
    start_date, end_date = _extract_period_or_raise(date_period_raw, sheet_name)

    # 7. 若期間正確，顯示中繼驗證通過（格式與風格統一）
    print(f" ✓ 中繼驗證通過 | 期間: {start_date}~{end_date}")

    # 8. 讀取該分頁下方的正式表格 (略過上方雜訊)
    df_table = read_data_table(file_path, sheet_name, _NEW_TABLE_HEADER_ROW)

    # 9. 逐行將資料打包成 DealerRecord 物件並加入清單
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


def _clean_dealer_names(dealer_name_col: pd.Series) -> set:
  """將經銷商名稱欄位清洗成不含空值/空字串的集合，供集合運算比較差異。"""
  cleaned = dealer_name_col.dropna().astype(str).str.strip()
  return set(cleaned[cleaned != ""])


def _export_diff(diff_df: pd.DataFrame, title: str, filename: str) -> None:
  """將差異結果 DataFrame 轉成 list[dict] 後匯出成 Markdown 報表。

  export_records_report 只接受 list[dict]；若直接傳入 DataFrame，
  records_to_markdown 內的 `if not records:` 會對 DataFrame 做布林判斷，
  拋出 "The truth value of a DataFrame is ambiguous" 錯誤，故統一經由本函式轉換。
  """
  export_records_report(diff_df.to_dict(orient="records"), title=title, filename=filename)


def compare(old_file_path: str, new_file_path: str) -> str:
  """比對邏輯主函式：同時從「資料列」與「經銷商」兩個維度分析新舊報表差異。

  比對範圍包含：
    1. 資料列維度：舊檔獨有、新檔獨有、兩邊完全一致的資料列筆數。
    2. 經銷商維度：舊/新檔各自獨有的經銷商，以及兩邊都有但銷售金額或
       銷售數量不同的經銷商。

  比對完成後會直接呼叫 export_records_report 將差異明細匯出為 Markdown
  報表（存放於 compareResult 資料夾），並回傳純文字的統計結果供 UI 顯示。

  Args:
    old_file_path: SAP625（舊系統）Excel 檔案路徑。
    new_file_path: SAN070R1（新系統）Excel 檔案路徑。

  Returns:
    比對結果的統計文字說明；若讀取/比對過程發生例外，則回傳錯誤訊息字串。
  """
  print(f"--- 正在執行：{RULE_NAME} ---")

  try:
    df_old = process_Old(old_file_path)
    df_new = process_New(new_file_path)

    if not df_old or not df_new:
      return "比對失敗：舊專案或新專案的表格內容為空。"

    df_old_pd = pd.DataFrame([vars(r) for r in df_old])
    df_new_pd = pd.DataFrame([vars(r) for r in df_new])

    # --- 1. 資料列維度比較：定義雙方的身分識別 Key（用來判斷是不是同一筆） ---
    key_columns = [
        "company_branch", "department", "start_date", "end_date",
        "level", "dealer_name", "sales_person", "sales_amount", "sales_count",
    ]

    # 以 df_old 為主體，找出「舊檔有，但新檔沒有」的資料
    merged_left = pd.merge(df_old_pd, df_new_pd[key_columns], on=key_columns, how="left", indicator=True)
    missing_in_new_df = merged_left[merged_left["_merge"] == "left_only"].drop(columns=["_merge"])

    # 以 df_new 為主體，找出「新檔有，但舊檔沒有」的資料
    merged_right = pd.merge(df_new_pd, df_old_pd[key_columns], on=key_columns, how="left", indicator=True)
    missing_in_old_df = merged_right[merged_right["_merge"] == "left_only"].drop(columns=["_merge"])

    # 找出兩邊完全一致的資料（用來計算「兩邊完全一致筆數」）
    merged_both = pd.merge(df_old_pd, df_new_pd[key_columns], on=key_columns, how="left", indicator=True)
    matched_df = merged_both[merged_both["_merge"] == "both"].drop(columns=["_merge"])

    # --- 2. 經銷商維度比較：只看「經銷商」這個欄位本身有無出現在雙方報表 ---
    old_dealers = _clean_dealer_names(df_old_pd["dealer_name"])
    new_dealers = _clean_dealer_names(df_new_pd["dealer_name"])

    dealers_only_in_old = sorted(old_dealers - new_dealers)  # 舊的有，新的沒有（少掉的）
    dealers_only_in_new = sorted(new_dealers - old_dealers)  # 新的有，舊的沒有（新增的）
    dealers_both = sorted(old_dealers & new_dealers)  # 兩邊都有的共同經銷商

    df_dealers_old = pd.DataFrame({"經銷商名稱": dealers_only_in_old})
    df_dealers_new = pd.DataFrame({"經銷商名稱": dealers_only_in_new})

    # --- 3. 經銷商銷售變異比較：同一經銷商在新舊檔的銷售金額/數量是否不同 ---
    dealer_merged = pd.merge(
        df_old_pd[["dealer_name", "sales_amount", "sales_count"]],
        df_new_pd[["dealer_name", "sales_amount", "sales_count"]],
        on="dealer_name",
        suffixes=("_old", "_new"),
        how="inner",
    )
    dealer_variance_df = dealer_merged[
        (dealer_merged["sales_amount_old"] != dealer_merged["sales_amount_new"])
        | (dealer_merged["sales_count_old"] != dealer_merged["sales_count_new"])
    ]

    # --- 4. 產出統計訊息（涵蓋資料列與經銷商兩種維度） ---
    result_msg = f"""--- 獨立差異比對結果 ---
    【資料列維度比較】
    - 舊檔獨有（新檔缺漏）資料筆數: {len(missing_in_new_df)} 筆
    - 新檔獨有（舊檔多出）資料筆數: {len(missing_in_old_df)} 筆
    - 兩邊完全一致資料筆數: {len(matched_df)} 筆
    ----------------------------
    【經銷商維度比較】
    - 舊檔有、新檔沒有的經銷商: {len(dealers_only_in_old)} 家
    - 新檔有、舊檔沒有的經銷商: {len(dealers_only_in_new)} 家
    - 兩邊都有的共同經銷商: {len(dealers_both)} 家
    - 銷售金額/數量有變異的經銷商: {len(dealer_variance_df)} 家"""

    # --- 5. 匯出報表 ---
    # export_records_report 只接受 list[dict]，DataFrame 必須先用
    # _export_diff() 統一轉換，避免 records_to_markdown 內的
    # `if not records:` 對 DataFrame 做布林判斷而拋出
    # "The truth value of a DataFrame is ambiguous" 錯誤。
    _export_diff(
        missing_in_old_df,
        "SAP625 獨有資料清單",
        f"差異明細_SAP625(舊)系統獨有資料",
    )
    _export_diff(
        missing_in_new_df,
        "SAN070R1 獨有資料清單",
        f"差異明細_SAN070R1(新)系統獨有資料",
    )
    _export_diff(
        df_dealers_old,
        "SAP625 獨有經銷商",
        f"差異明細_SAP625(舊)系統獨有經銷商",
    )
    _export_diff(
        df_dealers_new,
        "SAN070R1 獨有經銷商",
        f"差異明細_SAN070R1(新)系統獨有經銷商",
    )
    _export_diff(
        dealer_variance_df,
        "SAN070R1與SAP625 經銷商銷售數量/價格比較",
        f"差異明細_SAN070R1與SAP625經銷商銷售金額數量變異",
    )
    return result_msg
  except Exception as e:
    return f"讀取或比對過程發生錯誤: {e}"
