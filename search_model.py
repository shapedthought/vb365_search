from __future__ import annotations

from typing import List

from pydantic import BaseModel, Field


class Attachment(BaseModel):
    name: str
    sizeBytes: int


class Property1(BaseModel):
    href: str


class Property2(BaseModel):
    href: str


class _Links(BaseModel):
    property1: Property1
    property2: Property2


class Property11(BaseModel):
    uri: str
    method: str


class Property21(BaseModel):
    uri: str
    method: str


class _Actions(BaseModel):
    property1: Property11
    property2: Property21


class Result(BaseModel):
    mailboxId: str
    attachments: List[Attachment]
    organizer: str
    attendees: str
    startTime: str
    endTime: str
    location: str
    subject: str
    recurrencePatternFormat: str
    recurring: bool
    itemClass: str
    _links: _Links
    _actions: _Actions
    id: str
    name: str
    address: str
    businessPhone: str
    company: str
    displayAs: str
    email: str
    fax: str
    fileAs: str
    fullName: str
    homePhone: str
    imAddress: str
    jobTitle: str
    mobile: str
    webPage: str
    from_: str = Field(..., alias="from")
    postedOn: str
    importance: str
    cc: str
    bcc: str
    to: str
    sent: str
    received: str
    reminder: bool
    duration: int
    entryType: str
    date: str
    status: str
    percentComplete: int
    startDate: str
    dueDate: str
    owner: str


class Property12(BaseModel):
    href: str


class Property22(BaseModel):
    href: str


class _Links1(BaseModel):
    property1: Property12
    property2: Property22


class SearchResponse(BaseModel):
    offset: int
    limit: int
    setId: str
    results: List[Result]
    _links: _Links1
