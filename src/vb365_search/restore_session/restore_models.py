from typing import Optional
from pydantic import BaseModel, Field

class RestoreSessionRequest(BaseModel):
    date_time: Optional[str] = Field(alias="dateTime", default=None)
    show_all_versions: bool = Field(alias="showAllVersions", default=True)
    show_deleted: bool = Field(alias="showDeleted", default=True)
    type_restore: str = Field(alias="type", default="Vex")



class Href(BaseModel):
    href: str

class _Links(BaseModel):
    property1: Href
    property2: Href

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
    details: str
    scopeName: str
    clientHost: str
    reason: str
    eTag: int
    _links: _Links


class SearchRequest(BaseModel):
    query: str
