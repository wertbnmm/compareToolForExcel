# Excel Comparator 文件索引

本專案文件已依使用情境拆分，避免所有規格與維護內容集中在單一檔案。

## 文件導覽

| 文件 | 用途 |
|---|---|
| [專案規格與架構](docs/PROJECT-SPEC.md) | 功能規格、比對規則、資料流程與目錄結構 |
| [開發指南](docs/DEVELOPMENT-GUIDE.md) | 新增比對規則、Excel 解析、UserService 與共用工具 |
| [部署與環境設定](docs/DEPLOYMENT.md) | SIT/UAT 切換、appsettings、PyInstaller 與啟動 |
| [維護與驗證](docs/MAINTENANCE.md) | 錯誤處理、測試、維護流程與完成檢查表 |

## 快速啟動

安裝依賴：

```powershell
pip install -r requirements.txt
```

啟動 GUI：

```powershell
.\run.ps1
```

GUI 啟動後，可在「執行環境」下拉選擇 `SIT` 或 `UAT`。

## 快速打包

```powershell
pyinstaller ExcelComparator.spec
```

打包內容與設定檔規則請參考[部署與環境設定](docs/DEPLOYMENT.md)。

## 重要原則

- 比對規則放在 `src\excel_comparator\comparisons\`。
- 每個比對規則提供 `RULE_NAME` 與 `compare(old_file_path, new_file_path)`。
- SAP162/SAN080R1 使用 `empno` 作為比對鍵。
- 使用者查詢名稱必須去重後放入 `UserQueryRequest.multiple_name`。
- 打包後只使用內嵌的 SIT/UAT `appsettings`，不讀取外部設定。
- 變更完成後至少執行 Python 語法檢查與 `git diff --check`。
