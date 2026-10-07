from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class UserQueryRequest:
  result_count: Optional[int] = None
  userid: Optional[str] = None
  userid_fuzzy: Optional[str] = None
  empno: Optional[str] = None
  usernm: Optional[str] = None
  usernm_fuzzy: Optional[str] = None
  branchid: Optional[str] = None
  deptid_fuzzy: Optional[str] = None
  deptsimpnm_fuzzy: Optional[str] = None
  deptfullnm_fuzzy: Optional[str] = None
  userkind: Optional[str] = None
  echotitlecd: Optional[str] = None
  slzonecd: Optional[str] = None
  valid_yn: Optional[bool] = None
  salerepflag_yn: Optional[bool] = None
  compid: Optional[str] = None
  multiple_userid: List[str] = field(
      default_factory=list
  )  # 對應 dbo.STRARRAY 的 ITEM
  multiple_name: List[str] = field(
      default_factory=list
  )  # 對應 dbo.STRARRAY 的 ITEM

  def to_tuple(self) -> Tuple:
    """將 Request 物件轉換為 pyodbc 執行 SP 所需的參數 Tuple

    對應 SQL Server 的 dbo.STRARRAY (ITEM VARCHAR(2000))，
    pyodbc 需要傳入 List of Tuples，例如 [('ID1',), ('ID2',)]
    """
    tvp_param = (
        [(str(uid),) for uid in self.multiple_userid]
        if self.multiple_userid
        else []
    )
    multiple_name_param = (
        [(str(name),) for name in self.multiple_name]
        if self.multiple_name
        else []
    )

    return (
        self.result_count,
        self.userid,
        self.userid_fuzzy,
        self.empno,
        self.usernm,
        self.usernm_fuzzy,
        self.branchid,
        self.deptid_fuzzy,
        self.deptsimpnm_fuzzy,
        self.deptfullnm_fuzzy,
        self.userkind,
        self.echotitlecd,
        self.slzonecd,
        self.valid_yn,
        self.salerepflag_yn,
        self.compid,
        tvp_param,  # 對應 @MULTIPLE_USERID dbo.STRARRAY (ITEM)
        multiple_name_param,  # 對應 @MULTIPLE_NAME dbo.STRARRAY (ITEM)
    )