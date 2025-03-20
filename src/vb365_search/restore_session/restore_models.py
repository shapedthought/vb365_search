from pydantic import BaseModel, Field
from typing import Any, Dict, Optional, List


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


class RestoreSessionsResponse(BaseModel):
    offset: int
    limit: int
    set_id: str = Field(alias="setId")
    results: List[RestoreSessionResponse]
    _links: Links


class Items(BaseModel):
    id: str
    itemRestoreDetailsType: str
    title: str
    error: str
    name: str
    itemType: str
    status: str
    path: str
    warnings: List[str]
    mailboxEmail: str
    mailboxIsArchive: bool
    mailboxIsPublic: bool
    createdMailboxItemsCount: int
    mergedMailboxItemsCount: int
    failedMailboxItemsCount: int
    skippedMailboxItemsCount: int
    cannotContinueMailboxError: str
    childItems: List[Dict[str, Any]]
    oneDriveWebId: str
    oneDriveSiteId: str
    oneDriveUrl: str


class RestoreStatisticsResponse(BaseModel):
    totalItemsCount: int
    restoredItemsCount: int
    failedItemsCount: int
    skippedItemsCount: int
    warnings: List[str]
    errors: List[str]
    items: List[Items]
