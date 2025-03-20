from __future__ import annotations

from typing import Any, List

from pydantic import BaseModel

class Data(BaseModel):
    platformName: str
    type: str
    malwareStatus: str
    id: str
    name: str
    platformId: str
    creationTime: str
    backupId: str
    sessionId: Any
    allowedOperations: List[str]
    backupFileId: str


class Pagination(BaseModel):
    total: int
    count: int
    skip: int
    limit: int


class RestorePointResponse(BaseModel):
    data: List[Data]
    pagination: Pagination
    

