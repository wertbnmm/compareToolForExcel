# 專案規格與架構

## 1. 專案概要

Excel Comparator 是 Windows 桌面工具，使用 CustomTkinter 提供 GUI、pandas 讀取 Excel，依選定規則比較舊系統與新系統報表，最後將差異輸出為 Markdown。

## 2. 使用者流程

1. 啟動程式。
2. 選擇或拖曳舊系統 Excel。
3. 選擇或拖曳新系統 Excel。
4. 選擇比對模式。
5. 選擇 `SIT` 或 `UAT`。
6. 按下「開始執行比對」。
7. GUI 顯示比對摘要，差異報表輸出至 `output\`。
8. 按下「開啟輸出資料夾」查看結果；資料夾尚未產生時會顯示提示。

## 3. 比對規則

| 規則 | 舊系統 | 新系統 | 主要比對鍵 |
|---|---|---|---|
| SAP625/SAN070R1 | SAP625 | SAN070R1 | 經銷商、部門、期間及銷售欄位 |
| SAP162/SAN080R1 | SAP162 | SAN080R1 | `empno` |

每個規則必須提供：

```python
RULE_NAME = "介面顯示名稱"


def compare(old_file_path: str, new_file_path: str) -> str:
    """執行比對並回傳 GUI 顯示的摘要。"""
```

`comparison_registry.py` 會掃描 `comparisons\*.py`，找到 `RULE_NAME` 與 `compare()` 後加入 GUI 下拉選單。

## 4. SAP162/SAN080R1 規格

- SAP162 舊系統直接使用 Excel 的 `emp_id`。
- SAN080R1 新系統先取得公司名稱，透過 `get_company_id()` 轉換為 `compid`。
- 公司不存在時停止解析並顯示工作表名稱與公司名稱。
- 業務員名稱去重後一次放入 `UserQueryRequest.multiple_name`。
- 查詢條件同時使用公司代碼與業務員名稱。
- 使用者資料回傳後以 `empno` 比對，不使用 `userid`。
- 查不到使用者時保留差異資料並顯示警告。
- 差異報表最後加入「總數」列，金額欄位顯示總和，業務員欄位保持空白。

## 5. 輸出規格

- 目錄：專案根目錄 `output\`
- 格式：Markdown
- 檔名：`原始檔名_yyyyMMddHHmmss.md`
- 每次輸出建立新檔，不覆蓋舊報表。

## 6. 系統架構

```text
GUI
  ↓
Application / Comparison Registry
  ↓
Comparisons
  ├─ Excel shared utilities
  ├─ Company helper
  ├─ Report exporter
  └─ UserService（SAP162/SAN080R1）
       ↓
     UserRepository
       ↓
     BaseRepository
       ↓
     DatabaseManager
       ↓
     SQL Server
```

## 7. 目錄結構

```text
compareToolForExcel\
├─ src\excel_comparator\
│  ├─ main.py
│  ├─ gui.py
│  ├─ application\comparison_registry.py
│  ├─ comparisons\
│  ├─ database\db_manager.py
│  ├─ models\user\
│  ├─ repositories\
│  ├─ services\
│  └─ shared\
├─ config\
│  ├─ appsettings.SIT.json
│  └─ appsettings.UAT.json
├─ docs\
├─ output\
├─ ExcelComparator.spec
├─ requirements.txt
├─ run.ps1
└─ DEVELOPMENT.md
```

| 層級 | 責任 |
|---|---|
| GUI | 檔案選擇、環境選擇、顯示訊息、呼叫 `compare()` |
| Application | 掃描並註冊比對規則 |
| Comparisons | 特定報表解析與業務比對 |
| Services | 封裝業務服務 |
| Repositories | 執行預存程序或原生 T-SQL、轉換 DTO |
| Database | 讀取設定、建立 pyodbc 連線 |
| Shared | Excel、公司代碼與報表輸出共用功能 |
| Models | Request 與 Response DTO |
