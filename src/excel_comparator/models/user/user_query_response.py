from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


@dataclass
class HpmUserDto:
  """單筆使用者資料 DTO (對應 C# HPMUSERDto)"""

  userid: Optional[str] = None
  userpw: Optional[str] = None
  userkind: Optional[str] = None
  usernm: Optional[str] = None
  userengnm: Optional[str] = None
  empno: Optional[str] = None
  deptid: Optional[str] = None
  jobtitleid: Optional[str] = None
  jobfuncid: Optional[str] = None
  echotitlecd: Optional[str] = None
  branchid: Optional[str] = None
  director: Optional[str] = None
  empsts: Optional[str] = None
  empctg: Optional[str] = None
  techteam: Optional[str] = None
  watteam: Optional[str] = None
  emailaddr: Optional[str] = None
  mobilno: Optional[str] = None
  hiredt: Optional[datetime] = None
  resigndt: Optional[datetime] = None
  effdt: Optional[datetime] = None
  expdt: Optional[datetime] = None
  remarks: Optional[str] = None
  pwchangedt: Optional[datetime] = None
  pwerrcnt: Optional[int] = None
  crtuserid: Optional[str] = None
  mtuserid: Optional[str] = None
  defaultpw: Optional[str] = None
  ipaddr: Optional[str] = None
  sex: Optional[str] = None
  birthday: Optional[datetime] = None
  marritalsts: Optional[str] = None
  zipcode1: Optional[str] = None
  zipcode2: Optional[str] = None
  addr: Optional[str] = None
  telarea: Optional[str] = None
  telno: Optional[str] = None
  educationlvl: Optional[str] = None
  schnm: Optional[str] = None
  major: Optional[str] = None
  salerepmark: Optional[str] = None
  extno: Optional[str] = None
  salerepflag: Optional[str] = None
  hrmobileno: Optional[str] = None
  ipadid: Optional[str] = None
  slzonecd: Optional[str] = None
  jobtitlegrpcd: Optional[str] = None
  hremailaddr: Optional[str] = None
  brand: Optional[str] = None
  lastlogindt: Optional[datetime] = None
  usercarkindcd: Optional[str] = None
  usercarbrandcd: Optional[str] = None
  compid: Optional[str] = None
  crtdt: Optional[datetime] = None
  crtpgmid: Optional[str] = None
  mtdt: Optional[datetime] = None
  mtpgmid: Optional[str] = None
  comp_name: Optional[str] = None
  branchsimpnm: Optional[str] = None
  deptsimpnm: Optional[str] = None
  deptfullnm: Optional[str] = None
  echotitlenm: Optional[str] = None
  slzonecdnm: Optional[str] = None
  valid_yn: Optional[bool] = None
  salerepflag_yn: Optional[bool] = None

  @classmethod
  def from_dict(cls, data: dict):
    """將 SQL 撈出的 dict 自動對應至 DTO 屬性 (忽略大小寫差異)"""
    lower_data = {k.lower(): v for k, v in data.items()}
    field_names = {f.name for f in cls.__dataclass_fields__.values()}
    filtered_data = {
        k: v for k, v in lower_data.items() if k in field_names
    }
    return cls(**filtered_data)


@dataclass
class UserQueryResponse:
  """統一的使用者查詢回應包裝 (Response)"""

  success: bool = True
  message: str = "Success"
  data: List[HpmUserDto] = field(default_factory=list)
  total_count: int = 0