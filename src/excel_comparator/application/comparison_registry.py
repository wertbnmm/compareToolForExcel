import glob
import importlib
import os
from typing import Callable, Dict


def get_comparison_methods() -> Dict[str, Callable]:
  """自動掃描 comparisons 資料夾，讀取每個模組的 RULE_NAME 與 compare 方法。

  Returns:
    以「顯示名稱」為 key、對應比對函式 (compare) 為 value 的字典，
    可直接提供給 UI 下拉選單使用。
  """
  methods: Dict[str, Callable] = {}

  current_dir = os.path.dirname(__file__)
  compare_dir = os.path.join(os.path.dirname(current_dir), "comparisons")

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
      module = importlib.import_module(
          f"excel_comparator.comparisons.{module_name}"
      )

      # 檢查是否有統一的 compare 方法
      if hasattr(module, "compare") and callable(module.compare):
        # 取得模組內的 RULE_NAME，若沒寫就用檔名代替
        display_name = getattr(module, "RULE_NAME", module_name)

        # 以「顯示名稱」作為 Key，對應到該比對函式
        methods[display_name] = module.compare
    except Exception as e:
      print(f"載入模組 {module_name} 失敗: {e}")

  return methods