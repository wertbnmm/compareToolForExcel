from datetime import datetime
import os

def records_to_markdown(records: list[dict], title: str) -> str:
    """將字典清單轉換為 Markdown 表格，並自動將日期轉換為 yyyy/mm/dd 附加在標題後方"""
    if not records:
        return f"### {title}\n\n（無資料）\n"
    
    # 1. 自動從第一筆資料抓取 start_date (例如 '20240101') 轉成 '2024/01/01'
    raw_date = records[0].get('start_date', '')
    date_suffix = ""
    if len(raw_date) == 8:
        date_suffix = f" ({raw_date[:4]}/{raw_date[4:6]}/{raw_date[6:]})"
    
    full_title = f"{title}{date_suffix}"
    
    # 2. 從字典的 key 動態取得所有欄位名稱
    columns = list(records[0].keys())
    
    # 3. 組合表頭與分隔線
    header = "| " + " | ".join(columns) + " |"
    separator = "| " + " | ".join(["---"] * len(columns)) + " |"
    
    # 4. 組合每一列的資料
    rows = []
    for r in records:
        row_values = [str(r.get(col, "")).replace("\n", " ") for col in columns]
        rows.append("| " + " | ".join(row_values) + " |")
        
    # 5. 組裝成完整的 Markdown 格式
    md_table = f"### {full_title}\n" + "\n".join([header, separator] + rows) + "\n"
    return md_table


def export_records_report(records: list[dict], title: str, filename: str = "import_result.md"):
    """接收資料清單，自動建立 compareResult 資料夾並匯出為 Markdown 檔案，檔名會自動附加 yyyyMMddHHmmss"""
    output_dir = "compareResult"
    
    # 如果資料夾不存在就自動建立
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 取得當前時間並格式化為 yyyyMMddHHmmss (例如: 20260922181107)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    
    # 2. 處理檔名：分離主檔名與副檔名，確保時間戳記加在正確位置，並以 .md 結尾
    base_name, _ = os.path.splitext(filename)
    final_filename = f"{base_name}_{timestamp}.md"
    
    file_path = os.path.join(output_dir, final_filename)
    
    # 3. 轉換成 MD 表格
    md_content = records_to_markdown(records, title=title)
    
    # 4. 寫入檔案
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(f"# 資料匯入報表\n\n")
        f.write(md_content)
        
    print(f"✅ 報表已成功匯出至：{file_path}")