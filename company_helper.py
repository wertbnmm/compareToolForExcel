# 公司的對照資料庫
from typing import Optional

COMPANY_MAPPING = {
    "HA": {
        "short_name": "和泰器材",
        "full_name": "和泰冷凍空調器材股份有限公司",
    },
    "HD": {"short_name": "和泰興業", "full_name": "和泰興業股份有限公司"},
    "HK": {"short_name": "柏原和泰", "full_name": "台灣柏原和泰股份有限公司"},
    "HL": {"short_name": "興昊物流", "full_name": "興昊物流股份有限公司"},
    "HS": {
        "short_name": "和泰服務行銷",
        "full_name": "和泰服務行銷股份有限公司",
    },
    "LA": {"short_name": "隆昊搬家", "full_name": "隆昊搬家貨運有限公司"},
    "LB": {"short_name": "和泰運輸", "full_name": "和泰運輸有限公司"},
    "LC": {"short_name": "和泰起重", "full_name": "和泰起重工程有限公司"},
    "LD": {"short_name": "隆和起重", "full_name": "隆和起重工程有限公司"},
    "LE": {"short_name": "隆昊起重", "full_name": "隆昊起重工程有限公司"},
}


def get_company_id(name_input: str) -> Optional[str]:
  """根據輸入的公司名稱（簡稱或全銜），回傳對應的 Company ID (例如: 'HD')

  如果找不到則回傳 None。
  """
  if not name_input:
    return None

  # 將輸入文字去除前後空白
  clean_input = str(name_input).strip()

  for comp_id, info in COMPANY_MAPPING.items():
    # 支援透過 Company ID 直接查詢 (例如輸入 "HD")
    if clean_input.upper() == comp_id:
      return comp_id

    # 支援透過簡稱查詢 (例如輸入 "和泰興業")
    if info["short_name"] in clean_input:
      return comp_id

    # 支援透過全銜查詢 (例如輸入 "和泰興業股份有限公司")
    if info["full_name"] in clean_input:
      return comp_id

  return None  # 找不到對應的公司


# --- 測試範例 ---
if __name__ == "__main__":
  test_inputs = [
      "和泰興業",
      "和泰興業股份有限公司",
      "HD",
      "台灣柏原和泰股份有限公司",
      "未知公司",
  ]

  for name in test_inputs:
    comp_id = get_company_id(name)
    print(f"輸入: '{name}' ==> 查詢到的 Company ID: {comp_id}")