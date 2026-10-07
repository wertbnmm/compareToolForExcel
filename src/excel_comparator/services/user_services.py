from excel_comparator.shared.enums import ConnectionName
from excel_comparator.database.db_manager import DatabaseManager
from excel_comparator.models.user.user_query_request import UserQueryRequest
from excel_comparator.repositories.user_repository import UserRepository


class UserService:

  connectionName = ConnectionName.DEFAULT

  def __init__(self):
    self.user_repository = UserRepository(DatabaseManager())

  def query_user_info(
      self, request: UserQueryRequest
  ) -> list:
    """查詢使用者資訊

    Repository 已經負責執行 SP 並轉型成 DTO，
    因此 Service 層直接呼叫並回傳（若有額外商業邏輯可在這邊處理）。
    """
    return self.user_repository.query_user_info(self.connectionName, request)

  def query_user_info_tsql(self, request: UserQueryRequest) -> list:
    """以原生 T-SQL 查詢使用者資訊。"""
    return self.user_repository.query_user_info_tsql(
        self.connectionName, request
    )