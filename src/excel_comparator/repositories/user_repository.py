# repositories/user_repository.py
from excel_comparator.shared.enums import ConnectionName
from excel_comparator.models.user.user_query_request import UserQueryRequest
from excel_comparator.models.user.user_query_response import HpmUserDto
from excel_comparator.repositories.base_repository import BaseRepository


class UserRepository(BaseRepository):

  def query_user_info(
      self, conn_name: ConnectionName, request: UserQueryRequest
  ) -> list:
    # 1. 呼叫共用方法執行 SP 取得字典列表
    raw_rows = self._execute_sp_to_dict_list(
        conn_name=conn_name,
        sp_name="USP_Portal_QueryUserInfo",
        params=request.to_tuple(),
    )

    # 2. 對應成 DTO 回傳
    return [HpmUserDto.from_dict(r) for r in raw_rows]

  def query_user_info_tsql(
      self, conn_name: ConnectionName, request: UserQueryRequest
  ) -> list:
    names = sorted({
        str(name).strip()
        for name in request.multiple_name
        if name is not None and str(name).strip()
    })
    if not names or not request.compid:
      print(
          "[警告] 使用者 T-SQL 查詢缺少公司代碼或業務員名稱，略過資料庫查詢。"
      )
      return []

    # 確保這裡是用 "?" 而不是 "(?)"，避免產生雙層括號
    placeholders = ", ".join("(?)" for _ in names)

    tsql = f"""
    SELECT *
    FROM dbo.TVF_HPMUSER_GETALL() AS U
    WHERE U.COMPID = ?
      AND U.USERNM IN ({placeholders})
    """

    params = (request.compid,) + tuple(names)

    print("--- 執行的 T-SQL ---")
    print(tsql)
    print("--- 傳入的參數總數 ---", len(params))

    raw_rows = self._execute_tsql_to_dict_list(
        conn_name=conn_name,
        tsql=tsql,
        params=params,
    )

    return [HpmUserDto.from_dict(r) for r in raw_rows]