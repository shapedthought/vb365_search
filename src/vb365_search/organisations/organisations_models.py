from __future__ import annotations

from typing import Any, Dict, List

from pydantic import BaseModel


class Self(BaseModel):
    href: str


class _Links(BaseModel):
    self: Self


class ExchangeOnlineSettings(BaseModel):
    useApplicationOnlyAuth: bool
    account: str
    grantAdminAccess: bool
    useMfa: bool
    applicationId: str
    applicationCertificateThumbprint: str


class SharePointOnlineSettings(BaseModel):
    useApplicationOnlyAuth: bool
    officeOrganizationName: str
    sharePointSaveAllWebParts: bool
    account: str
    grantAdminAccess: bool
    useMfa: bool
    applicationId: str
    applicationCertificateThumbprint: str


class Self1(BaseModel):
    href: str


class Jobs(BaseModel):
    href: str


class Groups(BaseModel):
    href: str


class Users(BaseModel):
    href: str


class Sites(BaseModel):
    href: str


class Teams(BaseModel):
    href: str


class UsedRepositories(BaseModel):
    href: str


class RbacRoles(BaseModel):
    href: str


class _Links1(BaseModel):
    self: Self1
    jobs: Jobs
    groups: Groups
    users: Users
    sites: Sites
    teams: Teams
    usedRepositories: UsedRepositories
    rbacRoles: RbacRoles


class Result(BaseModel):
    isTeamsOnline: bool
    isTeamsChatsOnline: bool
    exchangeOnlineSettings: ExchangeOnlineSettings
    sharePointOnlineSettings: SharePointOnlineSettings
    isExchangeOnline: bool
    isSharePointOnline: bool
    type: str
    region: str
    id: str
    name: str
    officeName: str
    msid: str
    backedUpOrganizationId: str
    _links: _Links1
    _actions: Dict[str, Any]


class OrganisationResponseModel(BaseModel):
    offset: int
    limit: int
    _links: _Links
    results: List[Result]
