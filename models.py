from pydantic import BaseModel, Field, field_validator

from typing import Optional


class MicrosoftConfig(BaseModel):
    tenant_name: str = Field(default="")
    tenant_id: str = Field(default="")
    application_id: str = Field(default="")


class Vb365Config(BaseModel):
    api_address: str = Field(default="")
    username: str = Field(default="")
    password: str = Field(default="")
    version: str = Field(default="v8")


class Configuration(BaseModel):
    microsoft: MicrosoftConfig
    vb365: Vb365Config


class DeviceRequest(BaseModel):
    client_id: str
    scope: str


class DeviceResponse(BaseModel):
    user_code: str
    device_code: str
    verification_uri: str
    expires_in: int
    interval: int
    message: str


class Assertion(BaseModel):
    grant_type: str
    client_id: str
    device_code: str


class AssertionResponse(BaseModel):
    token_type: str
    scope: str
    expires_in: int
    access_token: str
    refresh_token: str


class VBLoginRequest(BaseModel):
    grant_type: str
    assertion: AssertionResponse


class VBLoginResponse(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str
    expires_in: int
    userName: str
    issued: str = Field(alias=".issued")
    expires: str = Field(alias=".expires")


class AuthHeaders(BaseModel):
    accept: str = Field(alias="Accept", default="application/json")
    content_type: str = Field(alias="Content-Type", default="application/json")
    authorization: str = Field(alias="Authorization")

    @field_validator("authorization")
    def check_authorization(cls, value):
        if not value.startswith("Bearer "):
            raise ValueError("Authorization must be a Bearer token")


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
