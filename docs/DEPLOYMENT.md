# 部署與環境設定

## 1. 依賴

```powershell
pip install -r requirements.txt
```

必要環境：

- Windows
- Python
- SQL Server ODBC Driver 17
- 可連線到 SIT/UAT 資料庫的網路

## 2. 啟動

一般啟動：

```powershell
.\run.ps1
```

GUI 啟動後從「執行環境」下拉選單切換 SIT/UAT，UI 預設選擇 SIT。

直接執行：

```powershell
$env:PYTHONPATH = "src"
python -m excel_comparator
```

## 3. 設定檔

```text
config\
├─ appsettings.SIT.json
└─ appsettings.UAT.json
```

每份設定必須包含：

```json
{
  "DebugMode": true,
  "ConnectionStrings": {
    "DefaultConnection": "...",
    "WorkflowConnection": "...",
    "DKAUConnection": "..."
  }
}
```

環境選擇與設定檔對應：

| 環境 | 設定檔 |
|---|---|
| SIT | `appsettings.SIT.json` |
| UAT | `appsettings.UAT.json` |

打包後只讀取 PyInstaller 內嵌設定，不讀取 exe 外部設定。缺少內嵌設定時直接報錯，不會退回其他環境。

設定檔包含資料庫密碼，禁止提交公開版本控制，打包檔也應限制存取權限。

## 4. PyInstaller

```powershell
pyinstaller ExcelComparator.spec
```

spec 會在編譯時：

- 將 `comparisons\*.py` 加入 `hiddenimports`。
- 將 `comparisons\*.py` 放入打包後的 `excel_comparator\comparisons\`，供註冊器掃描。
- 將 `appsettings.SIT.json` 與 `appsettings.UAT.json` 放入內嵌 `config\`。

新增比對規則或修改設定後，必須重新執行打包。

## 5. 設定檔載入規則

執行環境由 GUI 選擇後設定。`DatabaseManager` 在 frozen 模式下只從 PyInstaller 的 `_MEIPASS` 讀取對應設定：

```text
SIT → config\appsettings.SIT.json
UAT → config\appsettings.UAT.json
```

程式碼或外部環境不應提供另一份同名設定來覆蓋內嵌內容。
