from pydantic import BaseModel, Field
from typing import Optional

class RestoreSessionRequest(BaseModel):
    date_time: str = Field(alias="dateTime")
    show_all_versions: bool = Field(alias="showAllVersions", default=True)
    show_deleted: bool = Field(alias="showDeleted", default=True)
    type_restore: str = Field(alias="type", default="Vex")

class Link(BaseModel):
    href: str

class Links(BaseModel):
    self: Link
    organization: Link
    restoreSessionEvents: Link

class RestoreSessionResponse(BaseModel):
    id: str
    name: str
    organization: str
    type: str
    creationTime: str
    endTime: str
    state: str
    result: str
    initiatedBy: str
    scopeName: Optional[str] = None
    eTag: int
    _links: Links


class SearchRequest(BaseModel):
    query: str
