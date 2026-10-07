# 開發指南

## 1. 新增比對規則

在下列目錄建立新檔案：

```text
src\excel_comparator\comparisons\SAP999_SAN999R1.py
```

### 1.1 定義資料模型與欄位

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


_METADATA_FIELDS: List[MetadataField] = [
    MetadataField("業務員", x=0, y=4),
    MetadataField("金額", x=1, y=4),
]
```

`MetadataField` 的 `x` 是欄索引，`y` 是列索引，均從 0 開始。

### 1.2 分離舊檔與新檔解析

建議分成 `process_Old()` 與 `process_New()`，每個函式依序：

1. 讀取所有工作表。
2. 跳過特殊工作表，例如 `form`。
3. 驗證中繼標籤與表頭。
4. 讀取正式資料表。
5. 轉換為統一的資料模型。

```python
def process_Old(file_path: str) -> List[Record]:
    sheets = read_all_sheets_raw(file_path)
    records: List[Record] = []

    for sheet_name, raw_df in sheets.items():
        validate_sheet_structure(raw_df, sheet_name, _METADATA_FIELDS)
        table = read_data_table(file_path, sheet_name, header_row=6)

        for _, row in table.iterrows():
            key = safe_str(row.iloc[0])
            if key:
                records.append(
                    Record(key=key, amount=safe_float(row.iloc[1]))
                )

    return records
```

### 1.3 實作統一入口

```python
RULE_NAME = "SAP999/SAN999R1 比對"


def compare(old_file_path: str, new_file_path: str) -> str:
    old_records = process_Old(old_file_path)
    new_records = process_New(new_file_path)

    # 建立差異資料並轉成 list[dict]
    export_records_report(
        records,
        title="SAP999/SAN999R1 差異明細",
        filename="差異明細_SAP999_SAN999R1",
    )
    return "比對完成"
```

新增檔案後不需要修改註冊器；重新打包時 `ExcelComparator.spec` 會自動收集 `comparisons\*.py`。

## 2. 使用者查詢

查詢流程：

```text
Comparison Rule
  → UserService
  → UserRepository
  → BaseRepository
  → DatabaseManager
  → SQL Server
```

比對規則只呼叫：

```python
user_service.query_user_info_tsql(request)
```

不要在比對規則直接建立 `DatabaseManager`。

### 2.1 批次名稱查詢

```python
request = UserQueryRequest(
    compid=company_id,
    multiple_name=sorted(set(names)),
)
users = user_service.query_user_info_tsql(request)
```

原生 T-SQL 方法會：

1. 建立 `@MULTIPLE_NAME` table variable。
2. 使用 `(?)` placeholder 逐筆參數化插入名稱。
3. 使用 `COMPID` 與 `USERNM` 過濾 `TVF_HPMUSER_GETALL()`。
4. 轉換為 `HpmUserDto`。

不可將 Python list 直接拼接進 SQL。

### 2.2 SAP162 比對鍵

```python
employee_ids = {
    user.usernm: user.empno
    for user in users
    if user.usernm and user.empno
}
```

SAP162 固定使用 `empno`，不可改用 `userid`，除非需求規格明確變更。

## 3. 共用工具

### `excel_utils.py`

- `safe_str()`：空值轉空字串並清除前後空白。
- `safe_int()`、`safe_float()`：安全數值轉換。
- `format_columns()`：組合多欄並忽略空值。
- `validate_and_extract_period()`：驗證期間並取得起訖日期。
- `read_all_sheets_raw()`：無表頭讀取所有工作表。
- `validate_sheet_structure()`：驗證標籤與欄位座標。
- `read_data_table()`：讀取正式資料區。

### `company_helper.py`

公司查詢支援公司代碼、簡稱與完整名稱。新增公司時要同步維護 `COMPANY_MAPPING`。

### `report_export.py`

DataFrame 必須先轉成字典清單：

```python
export_records_report(
    dataframe.to_dict(orient="records"),
    title="差異明細",
    filename="差異明細",
)
```

## 4. 註解規範

- 模組用途寫在 module docstring。
- 公開類別與方法使用 docstring。
- 只註解非直覺的商業規則、Excel 座標或相容性原因。
- 不使用「新增這段」或變更歷史式註解。
- 不重複解釋一眼可讀懂的程式碼。
- 規格、部署與維護流程放在 `docs\`，不要堆在程式註解中。
