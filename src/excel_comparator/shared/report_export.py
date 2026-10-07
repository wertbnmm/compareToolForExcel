"""比對結果匯出模組：將資料清單轉換為 Markdown 報表並寫入檔案。"""
from datetime import datetime
import os
import sys
from typing import Any, Dict, List

import pandas as pd

def records_to_markdown(records, title: str) -> str:
  """將字典清單或 DataFrame 轉換為 Markdown 表格。"""
  if isinstance(records, pd.DataFrame):
    if records.empty:
      return f"### {title}\n\n（無資料）\n"
    records = records.to_dict(orient="records")

  if not records:
    return f"### {title}\n\n（無資料）\n"

  raw_date = records[0].get("start_date", "")
  date_suffix = ""
  if len(str(raw_date)) == 8:
    date_suffix = f" ({raw_date[:4]}/{raw_date[4:6]}/{raw_date[6:]})"

  full_title = f"{title}{date_suffix}"
  columns = list(records[0].keys())
  header = "| " + " | ".join(columns) + " |"
  separator = "| " + " | ".join(["---"] * len(columns)) + " |"
  rows = []
  for r in records:
    row_values = [str(r.get(col, "")).replace("\n", " ") for col in columns]
    rows.append("| " + " | ".join(row_values) + " |")

  md_table = f"### {full_title}\n" + "\n".join([header, separator] + rows) + "\n"
  return md_table


def get_output_dir() -> str:
    """取得比對結果輸出資料夾路徑。"""
    if getattr(sys, "frozen", False):
        project_root = os.path.dirname(sys.executable)
    else:
        project_root = os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "..", "..")
        )
    return os.path.join(project_root, "output")


def export_records_report(
    records: List[Dict[str, Any]], title: str, filename: str = "import_result.md"
) -> None:
    """將結果匯出至 output 資料夾，檔名會自動附加 yyyyMMddHHmmss。"""
    output_dir = get_output_dir()
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    base_name, _ = os.path.splitext(filename)
    final_filename = f"{base_name}_{timestamp}.md"
    file_path = os.path.join(output_dir, final_filename)
    md_content = records_to_markdown(records, title=title)

    with open(file_path, "w", encoding="utf-8-sig") as f:
        f.write("# 資料匯入報表\n\n")
        f.write(md_content)
        
    print(f"✅ 報表已成功匯出至：{file_path}")