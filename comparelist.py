from dataclasses import dataclass
import glob
import importlib
import os
from typing import Any

from altair import Dict
from narwhals import List

def get_comparison_methods():
  """自動掃描 compare 資料夾，讀取每個模組的 RULE_NAME 與 compare 方法"""
  methods = {}

  current_dir = os.path.dirname(__file__)
  compare_dir = os.path.join(current_dir, "compare")

  if not os.path.exists(compare_dir):
    return methods

  pattern = os.path.join(compare_dir, "*.py")
  for file_path in glob.glob(pattern):
    filename = os.path.basename(file_path)

    if filename.startswith("__"):
      continue

    module_name = filename[:-3]

    try:
      # 動態載入模組
      module = importlib.import_module(f"compare.{module_name}")

      # 檢查是否有統一的 compare 方法
      if hasattr(module, "compare") and callable(module.compare):
        # 取得模組內的 RULE_NAME，若沒寫就用檔名代替
        display_name = getattr(module, "RULE_NAME", module_name)

        # 以「顯示名稱」作為 Key，對應到該比對函式
        methods[display_name] = module.compare
    except Exception as e:
      print(f"載入模組 {module_name} 失敗: {e}")

  return methods