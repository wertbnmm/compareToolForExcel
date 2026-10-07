# 維護與驗證

## 1. 錯誤處理

| 情況 | 處理方式 |
|---|---|
| Excel 格式錯誤 | `validate_sheet_structure()` 拋出 `ValueError` |
| 公司不存在 | 顯示工作表與公司名稱，停止解析 |
| 使用者查不到 | 保留差異資料並顯示警告 |
| SQL/資料庫錯誤 | 顯示預存程序或 T-SQL、連線與錯誤原因後重新拋出 |
| 輸出資料夾不存在 | 提示先執行比對，不因點擊按鈕建立空資料夾 |
| 內嵌設定不存在 | 直接報錯，不改讀外部或其他環境設定 |

錯誤訊息應包含發生階段、檔案/工作表、欄位或公司名稱，以及可採取的處理方式。

## 2. 語法與差異檢查

```powershell
$env:PYTHONDONTWRITEBYTECODE = "1"
python -c "from pathlib import Path; [compile(p.read_text(encoding='utf-8'), str(p), 'exec') for p in Path('src').rglob('*.py')]; print('syntax ok')"
git diff --check
```

## 3. 功能驗證

1. 啟動 GUI，確認兩個比對規則出現在下拉選單。
2. 確認舊檔、新檔選擇與拖曳都可使用。
3. 確認環境下拉可以選擇 SIT/UAT。
4. 確認缺少 `output\` 時，開啟資料夾按鈕顯示提示。
5. 使用代表性檔案執行 SAP625/SAN070R1。
6. 使用代表性檔案執行 SAP162/SAN080R1。
7. 確認 SAP162 使用 `empno`，且名稱批次查詢可取得資料。
8. 確認報表產生於 `output\`，總數列與差異內容正確。
9. 打包後確認所有比對規則仍可載入。
10. 確認打包後 SIT/UAT 設定切換使用內嵌檔案。

## 4. 修改既有規則

1. 先核對 Excel 實際欄位位置與需求規格。
2. 優先只修改對應的 comparison module。
3. 共用邏輯只有在兩個以上規則需要時才放入 `shared`。
4. 同步更新 `docs\PROJECT-SPEC.md` 或 `docs\DEVELOPMENT-GUIDE.md`。
5. 執行語法、差異與代表性 Excel 驗證。

## 5. 修改資料庫查詢

- 不要在 Python 中拼接使用者輸入形成 SQL。
- `multiple_name` 必須去重並使用 parameter binding。
- 修改 `UserQueryRequest.to_tuple()` 時確認預存程序參數順序。
- 修改 DTO 欄位時確認 `HpmUserDto.from_dict()` 映射正常。
- 變更 `TVF_HPMUSER_GETALL()` 條件時，分別驗證公司代碼與姓名條件。

## 6. 完成檢查表

- [ ] 沒有新增未使用的 import。
- [ ] 沒有刪除仍被 GUI、註冊器或其他模組呼叫的方法。
- [ ] 比對規則仍提供 `RULE_NAME` 與 `compare()`。
- [ ] Excel 座標與表頭索引已核對。
- [ ] 錯誤訊息包含必要上下文。
- [ ] `multiple_name` 已去重。
- [ ] SAP162 比對鍵仍是 `empno`。
- [ ] SIT/UAT 設定檔必要欄位完整。
- [ ] 沒有把真實密碼寫入公開文件或程式碼。
- [ ] Python 語法檢查通過。
- [ ] `git diff --check` 通過。
- [ ] 打包後下拉選單可載入所有比對規則。
