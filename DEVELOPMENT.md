# Excel Comparator 開發文件

## 1. 文件目的

本文件說明 Excel Comparator 的功能規格、程式架構、資料流程、開發方式、打包方式與後續維護規則。

適用對象：

- 維護既有比對規則的開發人員
- 新增 Excel 比對規則的開發人員
- 維護部署設定與資料庫連線的工程人員

本專案是 Windows 桌面程式，使用 CustomTkinter 提供介面，使用 pandas 讀取 Excel，並將比對結果輸出為 Markdown。

---

## 2. 功能規格

### 2.1 使用者操作流程

1. 啟動程式。
2. 選擇舊系統 Excel 檔案，或將檔案拖曳至舊系統區域。
3. 選擇新系統 Excel 檔案，或將檔案拖曳至新系統區域。
4. 從下拉選單選擇比對規則。
5. 按下「開始執行比對」。
6. 程式解析兩份 Excel、執行規則、在畫面顯示結果。
7. 差異報表寫入 `output\`。
8. 按下「開啟輸出資料夾」可以開啟報表位置。若尚未產生資料，畫面會提示先執行比對。

### 2.2 比對規則介面

每個比對模組必須提供：

```python
RULE_NAME = "介面顯示名稱"


def compare(old_file_path: str, new_file_path: str) -> str:
    """執行比對並回傳供 GUI 顯示的摘要文字。"""
```

`RULE_NAME` 會顯示在 GUI 下拉選單中。`compare()` 是規則註冊器與 GUI 共用的統一入口。

### 2.3 目前提供的規則

| 規則 | 舊系統 | 新系統 | 比對鍵 | 輸出內容 |
|---|---|---|---|---|
| SAP625/SAN070R1 | SAP625 | SAN070R1 | 經銷商、部門、期間及銷售欄位 | 資料列差異、經銷商差異、金額/數量變異 |
| SAP162/SAN080R1 | SAP162 | SAN080R1 | `empno` | 撤票收回金額差異及總金額 |

### 2.4 SAP162/SAN080R1 特殊規格

- 舊系統直接使用 SAP162 的 `emp_id`。
- 新系統先從報表取得業務員名稱與公司名稱。
- 公司名稱必須經 `get_company_id()` 轉成 `compid`；找不到公司時停止解析並顯示錯誤。
- 業務員名稱去重後，透過一個 `UserQueryRequest` 的 `multiple_name` 批次查詢。
- 查詢條件同時包含名稱與公司代碼，避免查到其他公司的同名人員。
- 查詢結果以 `empno` 作為比對鍵，不使用 `userid`。
- 查不到對應人員時保留差異資料，並顯示「查無對應 emp」警告。
- 差異報表最後固定加入「總數」列；業務員欄位保持空白，金額欄位顯示總和。

### 2.5 輸出規格

- 輸出目錄：專案根目錄的 `output\`
- 檔案格式：Markdown
- 檔名格式：`原始檔名_yyyyMMddHHmmss.md`
- 每次匯出建立新檔，不覆蓋前一次結果。
- 匯出內容包含報表標題、欄位表頭、差異資料與必要的摘要資訊。

---

## 3. 專案結構

```text
compareToolForExcel\
├─ src\
│  └─ excel_comparator\
│     ├─ __main__.py                 # python -m excel_comparator 入口
│     ├─ main.py                     # 建立並啟動 GUI
│     ├─ gui.py                      # GUI、檔案選擇、事件處理
│     ├─ application\
│     │  └─ comparison_registry.py   # 掃描與載入比對規則
│     ├─ comparisons\
│     │  ├─ SAP162_SAN080R1.py
│     │  └─ SAP625_SAN070R1.py
│     ├─ database\
│     │  └─ db_manager.py            # 設定檔與資料庫連線
│     ├─ models\
│     │  └─ user\
│     │     ├─ user_query_request.py
│     │     └─ user_query_response.py
│     ├─ repositories\
│     │  ├─ base_repository.py       # 預存程序共用執行器
│     │  └─ user_repository.py
│     ├─ services\
│     │  └─ user_services.py
│     └─ shared\
│        ├─ company_helper.py
│        ├─ enums.py
│        ├─ excel_utils.py
│        └─ report_export.py
├─ config\
│  └─ appsettings.json               # 外部資料庫設定
├─ output\                           # 執行後產生，已加入忽略清單
├─ ExcelComparator.spec              # PyInstaller 設定
├─ requirements.txt
├─ run.ps1
└─ DEVELOPMENT.md
```

### 3.1 分層責任

| 層級 | 責任 |
|---|---|
| GUI | 接收使用者輸入、顯示訊息、呼叫統一的 `compare()` |
| Application | 找到並註冊可用的比對規則 |
| Comparisons | 解析特定報表格式、執行業務比對、匯出結果 |
| Services | 封裝業務服務，例如使用者查詢 |
| Repositories | 呼叫預存程序並轉換查詢結果 |
| Database | 讀取設定並建立 pyodbc 連線 |
| Shared | Excel 讀取、公司代碼、輸出報表等共用功能 |
| Models | 定義查詢 Request 與 Response DTO |

---

## 4. 執行環境與啟動

### 4.1 安裝依賴

在專案根目錄執行：

```powershell
pip install -r requirements.txt
```

必要環境：

- Python
- SQL Server ODBC Driver 17
- 可連線至設定檔中資料庫的網路環境
- Windows 桌面環境

### 4.2 啟動程式

建議使用：

```powershell
.\run.ps1
```

指定環境：

```powershell
.\run.ps1 -Environment SIT
.\run.ps1 -Environment UAT
```

直接啟動：

```powershell
$env:PYTHONPATH = "src"
$env:EXCEL_COMPARATOR_ENV = "SIT"
python -m excel_comparator
```

`run.ps1` 會設定 `PYTHONPATH`，並透過 `PYTHONDONTWRITEBYTECODE` 避免產生新的 `.pyc`。

### 4.3 資料庫設定

開發環境讀取：

```text
config\appsettings.json
```

指定 `SIT` 或 `UAT` 時讀取：

```text
config\appsettings.SIT.json
config\appsettings.UAT.json
```

指定環境但檔案不存在時，程式會直接報錯，不會退回共用設定檔，避免誤連到其他環境。

打包後優先讀取：

```text
ExcelComparator.exe
config\
└─ appsettings.json
```

若外部設定不存在，程式才會嘗試讀取 PyInstaller 暫存目錄中的內嵌設定。

正式環境應優先使用外部設定檔，因為資料庫連線資訊可能包含密碼，不應將真實密碼硬編在 exe。若需要提供預設設定，應提供不含敏感資訊的範本檔。

---

## 5. 新增比對規則

### 5.1 建立檔案

在 `src\excel_comparator\comparisons\` 建立新檔案，例如：

```text
SAP999_SAN999R1.py
```

### 5.2 定義資料模型與欄位規格

使用 `MetadataField` 定義需要驗證的 Excel 標籤與座標：

```python
from dataclasses import dataclass
from typing import List

from excel_comparator.shared.excel_utils import (
    MetadataField,
    read_all_sheets_raw,
    read_data_table,
    safe_float,
    safe_str,
    validate_sheet_structure,
)


@dataclass
class Record:
    key: str
    amount: float


_NEW_METADATA_FIELDS: List[MetadataField] = [
    MetadataField("業務員", x=0, y=4),
    MetadataField("金額", x=1, y=4),
]
```

欄位座標使用 0 起算：

- `x`：欄索引
- `y`：列索引

### 5.3 實作解析函式

建議將兩種報表分成 `process_Old()` 與 `process_New()`，每個函式只負責：

1. 讀取工作表。
2. 跳過不含資料的特殊工作表，例如 `form`。
3. 驗證中繼資訊與表頭。
4. 讀取正式資料區。
5. 轉換成統一的資料模型。

```python
def process_Old(file_path: str) -> List[Record]:
    sheets = read_all_sheets_raw(file_path)
    records: List[Record] = []

    for sheet_name, raw_df in sheets.items():
        validate_sheet_structure(raw_df, sheet_name, _OLD_METADATA_FIELDS)
        table = read_data_table(file_path, sheet_name, header_row=6)

        for _, row in table.iterrows():
            key = safe_str(row.iloc[0])
            if not key:
                continue
            records.append(Record(key=key, amount=safe_float(row.iloc[1])))

    return records
```

### 5.4 實作統一入口

```python
RULE_NAME = "SAP999/SAN999R1 比對"


def compare(old_file_path: str, new_file_path: str) -> str:
    old_records = process_Old(old_file_path)
    new_records = process_New(new_file_path)

    # 依規則建立 DataFrame，計算差異並匯出報表
    export_records_report(
        records,
        title="SAP999/SAN999R1 差異明細",
        filename="差異明細_SAP999_SAN999R1",
    )
    return "比對完成"
```

### 5.5 打包新規則

`ExcelComparator.spec` 會在建置時掃描：

```text
src\excel_comparator\comparisons\*.py
```

新增規則後重新執行打包即可：

```powershell
pyinstaller ExcelComparator.spec
```

不需要手動在 spec 內逐一新增模組名稱。

---

## 6. 使用者查詢開發規範

### 6.1 查詢流程

```text
Comparison Rule
    ↓
UserService
    ↓
UserRepository
    ↓
BaseRepository
    ↓
DatabaseManager
    ↓
USP_Portal_QueryUserInfo
```

比對規則只使用：

```python
user_service.query_user_info(request)
```

不應在比對規則直接建立 `DatabaseManager`。

目前 SAP162 規則在模組層級建立一個 `UserService` 實例，查詢時重用：

```python
user_service = UserService()
```

### 6.2 批次名稱查詢

名稱必須去重後放入 `multiple_name`：

```python
request = UserQueryRequest(
    compid=company_id,
    multiple_name=sorted(set(names)),
)
users = user_service.query_user_info(request)
```

`UserQueryRequest.to_tuple()` 會將名稱轉換成 SQL Server `dbo.STRARRAY` 所需的 tuple list。

目前同時保留兩種查詢入口：

- `query_user_info()`：呼叫既有預存程序。
- `query_user_info_tsql()`：在同一個 T-SQL batch 建立 `@MULTIPLE_NAME` table variable，將去重後的名稱以參數化方式插入，再查詢 `TVF_HPMUSER_GETALL()`。

原生 T-SQL 查詢不可將 Python 的 `list[str]` 直接綁定為 `@MULTIPLE_NAME`。必須為每個名稱建立 `(?)` placeholder，並將名稱參數依序放在 `params`，最後再放入 `compid`。這樣可以保留批次查詢並避免字串拼接造成 SQL Injection。

### 6.3 比對鍵規則

使用者查詢回傳 DTO 後，應使用 `empno` 建立比對字典：

```python
employee_ids = {
    user.usernm: user.empno
    for user in users
    if user.usernm and user.empno
}
```

不可在 SAP162 規則改用 `userid` 取代 `empno`，除非需求規格明確變更。

---

## 7. 共用工具維護規則

### 7.1 `excel_utils.py`

- `safe_str()`：處理空值並轉成清除前後空白的字串。
- `safe_int()`：處理空值並轉成整數。
- `safe_float()`：處理空值並轉成浮點數。
- `format_columns()`：組合多個欄位並忽略空值。
- `validate_and_extract_period()`：驗證期間格式並回傳起訖日期。
- `MetadataField`：定義欄位標籤及座標。
- `validate_sheet_structure()`：驗證工作表格式。
- `read_all_sheets_raw()`：以無表頭方式讀取工作表。
- `read_data_table()`：讀取正式資料區。

新增共用方法前，先確認是否能延伸現有方法，避免每個規則重複處理空值或 Excel 讀取。

### 7.2 `company_helper.py`

公司查詢支援：

- 公司代碼，例如 `HD`
- 公司簡稱，例如 `和泰興業`
- 公司完整名稱，例如 `和泰興業股份有限公司`

新增公司時，同時更新 `COMPANY_MAPPING` 的 `short_name` 與 `full_name`。

### 7.3 `report_export.py`

匯出方法接收字典清單，並自動：

1. 建立 `output\`。
2. 加上時間戳記。
3. 轉換為 Markdown。
4. 以 UTF-8 with BOM 寫入檔案。

DataFrame 應先使用：

```python
records = dataframe.to_dict(orient="records")
```

再呼叫 `export_records_report()`。

---

## 8. 錯誤處理與訊息規範

- Excel 格式錯誤：使用 `validate_sheet_structure()` 拋出 `ValueError`。
- 公司查不到：顯示工作表名稱與原始公司名稱，停止該次新系統解析。
- 使用者查不到：保留差異資料，顯示 `[警告]`，不可靜默丟棄。
- 資料庫錯誤：`BaseRepository` 顯示預存程序、連線名稱與錯誤原因後重新拋出。
- GUI 執行錯誤：由 GUI 顯示錯誤訊息，不應讓程式無提示結束。
- 輸出目錄不存在：開啟資料夾按鈕只提示使用者先完成比對，不應因點擊按鈕自動產生空資料夾。

錯誤訊息應包含：

- 發生階段
- 工作表名稱或檔案名稱
- 可辨識的欄位或公司名稱
- 使用者可以採取的修正方式

---

## 9. 測試與驗證

目前專案沒有正式的 pytest/Karma 測試套件，至少應執行以下檢查：

### 9.1 語法檢查

```powershell
$env:PYTHONDONTWRITEBYTECODE = "1"
python -m compileall -q src
```

### 9.2 差異檢查

```powershell
git diff --check
```

### 9.3 功能驗證

1. 啟動 GUI。
2. 確認兩種比對規則都出現在下拉選單。
3. 確認舊檔、新檔選取與拖曳都能使用。
4. 確認缺少 `output\` 時，開啟資料夾按鈕顯示提示。
5. 使用代表性 Excel 執行 SAP625/SAN070R1。
6. 使用代表性 Excel 執行 SAP162/SAN080R1。
7. 確認報表產生於 `output\`，且差異內容、總數列與錯誤訊息正確。
8. 使用 PyInstaller 打包後，確認所有規則都出現在下拉選單。
9. 確認執行檔旁的 `config\appsettings.json` 可覆蓋預設設定。

---

## 10. 打包與部署

### 10.1 打包

```powershell
pip install -r requirements.txt
pyinstaller ExcelComparator.spec
```

打包時會：

- 使用 `src\excel_comparator\main.py` 作為入口。
- 將 `comparisons` 下的所有 Python 模組加入 `hiddenimports`。
- 將 `comparisons` 下的 Python 檔案放入打包後的 `excel_comparator\comparisons\` 資料夾，供註冊器掃描。
- 將 `config\` 放入打包內容。

### 10.2 部署內容

正式部署至少需要：

```text
ExcelComparator.exe
config\
└─ appsettings.json
```

若使用 one-file 模式，仍建議在 exe 旁放置外部 `config\appsettings.json`，讓不同環境可以不重新打包就更換資料庫設定。

---

## 11. 後續維護方法

### 11.1 修改既有規則

1. 先確認 Excel 實際欄位位置與需求規格。
2. 只修改對應規則檔案，不要把規則專屬邏輯放入共用工具。
3. 若欄位讀取方式可被兩個以上規則使用，再考慮放入 `excel_utils.py`。
4. 同步更新本文件的規格表與規則說明。
5. 執行語法檢查、差異檢查與代表性 Excel 驗證。

### 11.2 新增共用功能

1. 搜尋現有方法，確認沒有重複功能。
2. 定義清楚的參數、回傳值與例外行為。
3. 撰寫簡短 docstring，說明用途與特殊限制。
4. 更新本文件的共用工具章節。
5. 確認既有兩個比對規則均未改變行為。

### 11.3 修改資料庫設定

- 不要將帳密寫入 Python 原始碼。
- 不要把真實設定檔提交到公開版本控制。
- 變更預存程序參數順序時，必須同步修改 `UserQueryRequest.to_tuple()`。
- 變更 DTO 欄位時，確認 `HpmUserDto.from_dict()` 仍能正確映射欄位。

### 11.4 註解與文件規範

- 模組用途放在檔案最上方的 module docstring。
- 類別與公開方法使用 docstring。
- 只為非直覺的商業規則、Excel 座標或相容性原因加註解。
- 不使用「新增這段」、「第 1 步、第 2 步」等變更歷史式註解。
- 不在註解重複描述明顯的程式碼。
- 規格變更、部署方式與維護方式更新在本文件，不堆在程式註解中。

---

## 12. 已知限制

- Excel 欄位位置是依固定報表格式定義，報表版型變更時必須更新 `MetadataField`。
- 使用者查詢需要可用的 SQL Server、ODBC Driver 與正確的外部設定檔。
- GUI 使用 Windows 檔案總管開啟輸出目錄，因此不以跨平台為目標。
- 動態載入比對規則需要重新打包才能讓發布版本包含新增模組。
- `output\` 只保存報表，不作為輸入資料來源。

---

## 13. 維護檢查清單

每次修改完成後確認：

- [ ] 沒有新增未使用的 import。
- [ ] 沒有刪除仍被 GUI、註冊器或其他模組呼叫的方法。
- [ ] 新增或修改的比對規則仍提供 `RULE_NAME` 與 `compare()`。
- [ ] Excel 座標與表頭列索引已核對。
- [ ] 錯誤訊息包含檔案/工作表/欄位等必要上下文。
- [ ] `UserQueryRequest.multiple_name` 已去重。
- [ ] SAP162 的比對鍵仍是 `empno`。
- [ ] 外部 `config\appsettings.json` 不含不應提交的敏感資訊。
- [ ] `python -m compileall -q src` 通過。
- [ ] `git diff --check` 通過。
- [ ] 打包後下拉選單仍能載入所有比對規則。
