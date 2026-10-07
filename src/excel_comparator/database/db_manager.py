import json
import os
import sys
from pathlib import Path
import pyodbc
from excel_comparator.shared.enums import ConnectionName

class DatabaseManager:

  def __init__(self, config_path="config/appsettings.json"):
    environment = os.getenv("EXCEL_COMPARATOR_ENV", "").strip().upper()
    if environment and environment not in {"SIT", "UAT"}:
      raise ValueError(
          f"不支援的執行環境：{environment}，只允許 SIT 或 UAT。"
      )

    config_path_obj = Path(config_path)
    if environment:
      config_path_obj = config_path_obj.with_name(
          f"{config_path_obj.stem}.{environment}{config_path_obj.suffix}"
      )

    if getattr(sys, "frozen", False):
      external_config = Path(sys.executable).parent / config_path_obj
      bundled_root = getattr(sys, "_MEIPASS", None)
      bundled_config = (
          Path(bundled_root) / config_path_obj
          if bundled_root
          else external_config
      )
      config_candidates = (external_config, bundled_config)
    else:
      base_path = Path(__file__).resolve().parent.parent.parent.parent
      config_candidates = (base_path / config_path_obj,)

    self.config_file = next(
        (path for path in config_candidates if path.is_file()),
        config_candidates[0],
    )

    try:
      with open(self.config_file, "r", encoding="utf-8") as f:
        self.config = json.load(f)
    except (OSError, json.JSONDecodeError):
      raise FileNotFoundError(f"找不到或無法讀取設定檔: {self.config_file}")

  def is_debug_mode(self) -> bool:
    """檢查 appsettings.json 是否開啟 Debug 模式

    (預設為 False，若設定檔未填寫則關閉)
    """
    return self.config.get("DebugMode", False)

  def _convert_to_pyodbc_string(self, ado_net_conn_str: str) -> str:
    pyodbc_str = (
        ado_net_conn_str.replace("Data Source=", "SERVER=")
        .replace("Initial Catalog=", "DATABASE=")
        .replace("User ID=", "UID=")
        .replace("Password=", "PWD=")
        .replace("Encrypt=False", "Encrypt=no")
        .replace("TrustServerCertificate=True", "TrustServerCertificate=yes")
    )
    if "DRIVER=" not in pyodbc_str.upper():
      pyodbc_str = "DRIVER={ODBC Driver 17 for SQL Server};" + pyodbc_str
    return pyodbc_str

  def get_connection(self, conn_name: ConnectionName):
    key_str = conn_name.value
    if key_str not in self.config["ConnectionStrings"]:
      err_msg = f"在設定檔中找不到對應的連線 Key: {key_str}"
      raise ValueError(err_msg)

    raw_conn_str = self.config["ConnectionStrings"][key_str]
    pyodbc_conn_str = self._convert_to_pyodbc_string(raw_conn_str)

    return pyodbc.connect(pyodbc_conn_str)