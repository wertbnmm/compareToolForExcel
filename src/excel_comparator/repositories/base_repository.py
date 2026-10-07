import collections.abc
from excel_comparator.shared.enums import ConnectionName
from excel_comparator.database.db_manager import DatabaseManager


class BaseRepository:

  def __init__(self, db_manager: DatabaseManager):
    self.db_manager = db_manager

  def _execute_sp_to_dict_list(
      self, conn_name: ConnectionName, sp_name: str, params: tuple = None
  ) -> list:
    """共用：執行預存程序並自動將結果轉為 Dict 列表，處理連線開關與例外"""
    conn = None
    cursor = None
    try:
      conn = self.db_manager.get_connection(conn_name)
      cursor = conn.cursor()

      if params:
        placeholders = ", ".join(["?"] * len(params))
        sp_sql = f"{{CALL {sp_name} ({placeholders})}}"
        cursor.execute(sp_sql, params)
      else:
        sp_sql = f"{{CALL {sp_name}}}"
        cursor.execute(sp_sql)

      rows = []
      if cursor.description:
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
          # 將每一列轉為 {欄位名稱: value} 字典
          rows.append(dict(zip(columns, row)))

      # 只有當 DebugMode 為 true 時，才額外印出成功回傳結果與參數
      if self.db_manager.is_debug_mode():
        print("\n" + "=" * 60)
        print(f" 🔍 [DEBUG 執行成功] 預存程序: {sp_name}")
        print(f" • 來源連線: {conn_name.name}")
        print(f" • 帶入參數 (Params): {params}")

        # 判斷回傳內容是否為可迭代物件（迴圈），若是則跑迴圈印出每一筆結果
        if isinstance(rows, collections.abc.Iterable) and not isinstance(
            rows, (str, bytes, dict)
        ):
          print(f" • 回傳筆數: {len(rows)}")

      conn.commit()
      return rows

    except Exception as e:
      # 發生例外時，無論 DebugMode 為何，全面印出詳細錯誤細節與帶入參數
      print("\n" + "=" * 50)
      print(f" ❌ [錯誤細節] 執行預存程序 {sp_name} 失敗！")
      print(f" • 來源連線: {conn_name.name}")
      print(f" • 帶入參數: {params}")
      print(f" • 錯誤原因: {e}")
      print("=" * 50 + "\n")

      if conn:
        conn.rollback()
      raise e

    finally:
      if cursor:
        cursor.close()
      if conn:
        conn.close()
  def _execute_tsql_to_dict_list(
      self, conn_name: ConnectionName, tsql: str, params: tuple = None
  ) -> list:
    """共用：執行原生 T-SQL 語法並自動將結果轉為 Dict 列表，處理連線開關與例外"""
    conn = None
    cursor = None
    try:
      conn = self.db_manager.get_connection(conn_name)
      cursor = conn.cursor()

      print(f"--- 開始執行 T-SQL：{conn_name.name} ---")

      # 執行原生 T-SQL（支援帶入參數以防止 SQL Injection）
      if params:
        cursor.execute(tsql, params)
      else:
        cursor.execute(tsql)

      rows = []
      # 檢查是否有回傳的資料行結構（適用於 SELECT 等會回傳結果的語法）
      if cursor.description:
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
          # 將每一列轉為 {欄位名稱: value} 字典
          rows.append(dict(zip(columns, row)))

      # 只有當 DebugMode 為 true 時，才額外印出成功回傳結果與 SQL 語法
      if self.db_manager.is_debug_mode():
        print("\n" + "=" * 60)
        print(f" 🔍 [DEBUG 執行成功] 原生 T-SQL")
        print(f" • 來源連線: {conn_name.name}")
        print(f" • 執行語法: {tsql}")
        print(f" • 帶入參數 (Params): {params}")

        if isinstance(rows, collections.abc.Iterable) and not isinstance(
            rows, (str, bytes, dict)
        ):
          print(f" • 回傳筆數: {len(rows)}")

      conn.commit()
      return rows

    except Exception as e:
      # 發生例外時，全面印出詳細錯誤細節、SQL 語法與帶入參數
      print("\n" + "=" * 50)
      print(f" ❌ [錯誤細節] 執行 T-SQL 失敗！")
      print(f" • 來源連線: {conn_name.name}")
      print(f" • 執行語法: {tsql}")
      print(f" • 帶入參數: {params}")
      print(f" • 錯誤原因: {e}")
      print("=" * 50 + "\n")

      if conn:
        conn.rollback()
      raise e

    finally:
      if cursor:
        cursor.close()
      if conn:
        conn.close()