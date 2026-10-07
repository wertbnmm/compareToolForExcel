"""SAP162（舊系統）與 SAN080R1（新系統）撤票收回金額比對規則。

此模組負責：
  1. 解析 SAP162 與 SAN080R1 報表，依 empno 彙總 inward_amt（撤票收回金額）。
  2. 比對同 empno 金額是否一致，並自動標示「僅存在舊系統」、「僅存在新系統」或「金額不一致」。
  3. 於差異報表結尾固定加入「總數」列，並匯出為 Markdown 報表。
"""
from dataclasses import dataclass
from typing import Any, Dict, List, Tuple

import pandas as pd

from excel_comparator.models.user.user_query_request import UserQueryRequest
from excel_comparator.shared.company_helper import get_company_id
from excel_comparator.shared.report_export import export_records_report
from excel_comparator.services.user_services import UserService
from excel_comparator.shared.excel_utils import (
    MetadataField,
    read_all_sheets_raw,
    read_data_table,
    safe_float,
    safe_str,
    validate_sheet_structure,
)

# 🌟 介面下拉選單中的名稱
RULE_NAME = "SAP162/SAN080R1撤票收回金額比對"
user_service = UserService()


@dataclass
class DealerRecord:
  """單一業務員撤票收回金額統計資料紀錄。"""
  emp_id: str  # 員工編號
  sales_person: str  # 業務員
  inward_amt: float  # 撤票收回金額 (inward_amt)


# SAP162（舊檔）上方中繼標籤與表頭座標定義
_OLD_METADATA_FIELDS: List[MetadataField] = [
    MetadataField("auto_no", x=0, y=0),       # A1 儲存格
    MetadataField("apply_date", x=1, y=0),    # B1 儲存格
    MetadataField("cust_no", x=2, y=0),       # C1 儲存格
    MetadataField("apply_kind", x=3, y=0),    # D1 儲存格
    MetadataField("apply_remark", x=4, y=0),  # E1 儲存格
    MetadataField("emp_id", x=5, y=0),        # F1 儲存格
    MetadataField("comp_no", x=6, y=0),       # G1 儲存格
    MetadataField("dept_no", x=7, y=0),       # H1 儲存格
    MetadataField("bank_no", x=8, y=0),       # I1 儲存格
    MetadataField("bank_acc", x=9, y=0),      # J1 儲存格
    MetadataField("note_no", x=10, y=0),      # K1 儲存格
    MetadataField("lift_date", x=11, y=0),    # L1 儲存格
    MetadataField("inward_amt", x=12, y=0),   # M1 儲存格
    MetadataField("rem_amt", x=13, y=0),      # N1 儲存格
    MetadataField("sub_amt", x=15, y=0),      # P1 儲存格
    MetadataField("remark", x=16, y=0),       # Q1 儲存格
    MetadataField("keyin_emp", x=17, y=0),    # R1 儲存格
    MetadataField("keyin_date", x=18, y=0),   # S1 儲存格
    MetadataField("cust_sname", x=19, y=0),   # T1 儲存格
    MetadataField("emp_cname", x=20, y=0),    # U1 儲存格
    MetadataField("bank_sname", x=21, y=0),   # V1 儲存格
    MetadataField("bank_no3", x=22, y=0),     # W1 儲存格
    MetadataField("bank_acc3", x=23, y=0),    # X1 儲存格
    MetadataField("pay_date", x=24, y=0),     # Y1 儲存格
    MetadataField("pick_yn", x=25, y=0),      # Z1 儲存格
    MetadataField("snap_no", x=26, y=0),      # AA1 儲存格
    MetadataField("pick_date", x=27, y=0),    # AB1 儲存格
    MetadataField("hifi_pick", x=28, y=0),    # AC1 儲存格
    MetadataField("cancel_date", x=29, y=0),  # AD1 儲存格
    MetadataField("cash_amt", x=30, y=0),     # AE1 儲存格
    MetadataField("write_yn", x=31, y=0),     # AF1 儲存格
    MetadataField("harc110_pick", x=32, y=0), # AG1 儲存格
    MetadataField("harc120_pick", x=33, y=0), # AH1 儲存格
]
_OLD_TABLE_HEADER_ROW = 0


def _decode_legacy_text(value: Any) -> str:
  """將 XLS 無 CODEPAGE 記錄造成的 Latin-1 字串還原為繁中字串。"""
  text = safe_str(value)
  if not text:
    return ""
  try:
    return text.encode("latin1").decode("cp950").strip()
  except (UnicodeEncodeError, UnicodeDecodeError):
    return text


def _get_cell_text(df_raw: pd.DataFrame, row_index: int, col_index: int) -> str:
  """安全取得指定儲存格文字，欄位不存在時視為空白。"""
  if df_raw.shape[0] <= row_index or df_raw.shape[1] <= col_index:
    return ""
  return safe_str(df_raw.iloc[row_index, col_index])


# SAN080R1（新檔）上方中繼標籤與表頭座標定義
_NEW_METADATA_FIELDS: List[MetadataField] = [
    MetadataField("分公司", x=0, y=4),        # A5 儲存格
    MetadataField("部門", x=1, y=4),          # B5 儲存格
    MetadataField("業代", x=2, y=4),          # C5 儲存格
    MetadataField("比較收票總計", x=3, y=4),  # D5 儲存格
    MetadataField("現有店數", x=4, y=4),      # E5 儲存格
    MetadataField("協議店數", x=5, y=4),      # F5 儲存格
    MetadataField("收票合計", x=12, y=4),     # M5 儲存格
    MetadataField("撤票金額", x=13, y=4),     # N5 儲存格
    MetadataField("撤票收回金額", x=14, y=4), # O5 儲存格
    MetadataField("總計", x=15, y=4),         # P5 儲存格
    MetadataField("達成進度比", x=16, y=4),   # Q5 儲存格
]
_NEW_TABLE_HEADER_ROW = 6


def process_Old(file_path: str) -> List[DealerRecord]:
  """解析已依查詢條件整理好的 SAP162 匯出檔，依員工編號彙總 inward_amt。"""
  all_sheets_raw = read_all_sheets_raw(file_path)
  raw_rows = []

  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
      continue

    print(f"--- 正在檢核 SAP162 工作表: {sheet_name} ---")
    validate_sheet_structure(df_raw, sheet_name, _OLD_METADATA_FIELDS)
    df_table = read_data_table(file_path, sheet_name, _OLD_TABLE_HEADER_ROW)

    for _, row in df_table.iterrows():
      emp_id = safe_str(row["emp_id"])
      sales_person = _decode_legacy_text(row["emp_cname"]) or emp_id
      inward_amt = safe_float(row["inward_amt"])
      raw_rows.append({
          "emp_id": emp_id,
          "sales_person": sales_person,
          "inward_amt": inward_amt,
      })

  if not raw_rows:
    return []

  df_temp = pd.DataFrame(raw_rows)
  df_grouped = (
      df_temp.groupby("emp_id", as_index=False)
      .agg(sales_person=("sales_person", "first"), inward_amt=("inward_amt", "sum"))
  )

  return [
      DealerRecord(
          emp_id=r["emp_id"],
          sales_person=r["sales_person"],
          inward_amt=r["inward_amt"],
      )
      for _, r in df_grouped.iterrows()
  ]


def _query_employee_id_map(names: List[str], compid: str) -> Dict[str, str]:
  """一次依新系統業務員名稱與公司別查回 empno 對照表。"""
  print(f"--- 查詢新系統業務員：公司 {compid}，共 {len(names)} 人 ---")
  users = user_service.query_user_info_tsql(
      UserQueryRequest(compid=compid, multiple_name=sorted(set(names)))
  )
  employee_ids: Dict[str, str] = {}

  for user in users:
    name = safe_str(user.usernm)
    emp_id = safe_str(user.empno)
    if name in names and emp_id:
      employee_ids[name] = emp_id

  missing_names = [name for name in names if name not in employee_ids]
  for name in missing_names:
    print(f"[警告] 新系統業務員【{name}】查無對應 emp，將保留於差異結果。")

  return employee_ids


def process_New(file_path: str) -> List[DealerRecord]:
  """解析 SAN080R1 匯出的 Excel 報表，依 empno 彙總撤票收回金額。"""
  all_sheets_raw = read_all_sheets_raw(file_path)
  raw_rows = []

  for sheet_name, df_raw in all_sheets_raw.items():
    if sheet_name == "form":
      continue

    print(f"--- 正在檢核 SAN080R1 工作表: {sheet_name} ---")
    validate_sheet_structure(df_raw, sheet_name, _NEW_METADATA_FIELDS)
    company_name = df_raw.iloc[1, 6]
    company_id = get_company_id(company_name)
    if not company_id:
      print(
          f"工作表【{sheet_name}】找不到對應公司：{company_name or '空白'}"
      )
      return []
    print(
        f"[公司驗證通過] 公司: {company_name} | 公司代碼: {company_id}"
    )
    df_table = read_data_table(file_path, sheet_name, _NEW_TABLE_HEADER_ROW)

    for _, row in df_table.iterrows():
      sales_person = safe_str(row.iloc[2])
      if not sales_person or sales_person == "nan":
        continue
      raw_rows.append({
          "sales_person": sales_person,
          "compid": company_id,
          "inward_amt": safe_float(row.iloc[14]),
      })

  if not raw_rows:
    print("SAN080R1 沒有可查詢的業務員資料，略過使用者查詢。")
    return []

  df_temp = pd.DataFrame(raw_rows)
  employee_ids: Dict[Tuple[str, str], str] = {}
  for compid, company_rows in df_temp.groupby("compid"):
    names = sorted(set(company_rows["sales_person"]))
    for name, emp_id in _query_employee_id_map(names, compid).items():
      employee_ids[(compid, name)] = emp_id

  df_temp["emp_id"] = [
      employee_ids.get((row["compid"], row["sales_person"]))
      for _, row in df_temp.iterrows()
  ]
  missing_mask = df_temp["emp_id"].isna()
  df_temp.loc[missing_mask, "emp_id"] = df_temp.loc[missing_mask, "sales_person"].map(
      lambda name: f"查無對應emp:{name}"
  )
  df_temp.loc[missing_mask, "sales_person"] = df_temp.loc[missing_mask, "sales_person"].map(
      lambda name: f"{name}（查無對應emp）"
  )
  df_grouped = (
      df_temp.groupby("emp_id", as_index=False)
      .agg(sales_person=("sales_person", "first"), inward_amt=("inward_amt", "sum"))
  )
  return [
      DealerRecord(
          emp_id=row["emp_id"],
          sales_person=row["sales_person"],
          inward_amt=row["inward_amt"],
      )
      for _, row in df_grouped.iterrows()
  ]


def _export_diff(diff_df: pd.DataFrame, title: str, filename: str) -> None:
  """將差異結果 DataFrame 轉成 list[dict] 後匯出成 Markdown 報表。"""
  export_records_report(diff_df.to_dict(orient="records"), title=title, filename=filename)


def compare(old_file_path: str, new_file_path: str) -> str:
  """比對邏輯主函式：

  1. 判斷業務員是「僅存在舊系統」、「僅存在新系統」還是「金額不一致」。
  2. 自動補上名為「總數」的彙總行與總筆數計算。
  """
  print(f"--- 正在執行：{RULE_NAME} ---")

  try:
    df_old = process_Old(old_file_path)
    df_new = process_New(new_file_path)

    if not df_old or not df_new:
      return "比對失敗：SAP162 或 SAN080R1 的表格內容為空。"

    df_old_pd = pd.DataFrame([vars(r) for r in df_old])
    df_new_pd = pd.DataFrame([vars(r) for r in df_new])

    # 1. 依員工編號 outer join，保留兩邊的歸屬狀況
    merged_df = pd.merge(
        df_old_pd,
        df_new_pd,
        on="emp_id",
        suffixes=("_sap162", "_san080r1"),
        how="outer"
    )

    # 2. 定義狀態分類函數：判斷是缺人還是金額不對
    def classify_status(row):
      if pd.isna(row["inward_amt_sap162"]):
        return "僅存在新系統(舊系統無此人)"
      elif pd.isna(row["inward_amt_san080r1"]):
        return "僅存在舊系統(新系統無此人)"
      elif row["inward_amt_sap162"] != row["inward_amt_san080r1"]:
        return "兩邊皆有但金額不一致"
      return "相符"

    merged_df["status"] = merged_df.apply(classify_status, axis=1)

    # 3. 篩選出所有不相符（有差異）的項目
    variance_df = merged_df[merged_df["status"] != "相符"].copy()
    variance_df["sales_person_sap162"] = variance_df["sales_person_sap162"].fillna(
        "舊系統無對應業務員"
    )
    variance_df["sales_person_san080r1"] = variance_df["sales_person_san080r1"].fillna(
        "新系統無對應業務員"
    )
    variance_df[["inward_amt_sap162", "inward_amt_san080r1"]] = variance_df[
        ["inward_amt_sap162", "inward_amt_san080r1"]
    ].fillna(0)

    # 計算整體總金額
    total_old_sum = df_old_pd["inward_amt"].sum()
    total_new_sum = df_new_pd["inward_amt"].sum()

    # 4. 在差異明細中固定加入一筆 emp / sales_person 叫做 "總數" 的紀錄列
    summary_row = pd.DataFrame([{
        "emp_id": "總數",
        "sales_person_sap162": "",
        "sales_person_san080r1": "",
        "inward_amt_sap162": total_old_sum,
        "inward_amt_san080r1": total_new_sum,
        "status": f"差異總筆數: {len(variance_df)} 人"
    }])
    
    # 組合最終要匯出的報表資料（明細 + 總數列）
    export_df = pd.concat([variance_df, summary_row], ignore_index=True)

    # 5. 產出執行結果訊息
    missing_in_old = len(variance_df[variance_df["status"].str.contains("舊系統無此人")])
    missing_in_new = len(variance_df[variance_df["status"].str.contains("新系統無此人")])
    amount_mismatch = len(variance_df[variance_df["status"] == "兩邊皆有但金額不一致"])

    result_msg = f"""--- SAP162 與 SAN080R1 撤票收回金額比對結果 ---
【差異分類統計】
- 僅存在新系統（舊系統無此人）: {missing_in_old} 人
- 僅存在舊系統（新系統無此人）: {missing_in_new} 人
- 兩邊皆有但金額不一致: {amount_mismatch} 人
- 異常總人數: {len(variance_df)} 人

【總金額比對】
- SAP162 撤票收回總金額: {total_old_sum:,.2f}
- SAN080R1 撤票收回總金額: {total_new_sum:,.2f}
--------------------------------------------"""

    # 6. 匯出帶有「總數」列的差異報表
    _export_diff(
        export_df,
        "SAP162與SAN080R1 業務員撤票收回金額差異明細（含總數）",
        "差異明細_SAP162與SAN080R1_撤票收回金額不一致",
    )

    return result_msg

  except Exception as e:
    return f"讀取或比對過程發生錯誤: {e}"