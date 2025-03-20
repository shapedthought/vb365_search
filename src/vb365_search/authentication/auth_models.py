from enum import StrEnum
from typing import Optional
from pydantic import BaseModel, Field, HttpUrl


class DeviceCodeResponse(BaseModel):
    user_code: str
    device_code: str
    verification_uri: HttpUrl
    expires_in: int
    interval: int


class TokenResponse(BaseModel):
    access_token: str
    refresh_token: Optional[str] = None
    id_token: Optional[str] = None
    expires_in: int
    token_type: str


class VeeamTokenResponse(BaseModel):
    access_token: str
    refresh_token: str
    expires_in: int
    token_type: str


class AuthConfig(BaseModel):
    tenant_name: str
    client_id: str
    veeam_api_url: HttpUrl


class AuthHeaders(BaseModel):
    accept: str = Field(alias="Accept", default="application/json")
    content_type: str = Field(alias="Content-Type", default="application/json")
    authorization: str = Field(alias="Authorization")


class M365Permissions(StrEnum):
    DIRECTORY_READ_ALL = "Directory.Read.All"
    DIRECTORY_READ_WRITE_ALL = "Directory.ReadWrite.All"
    DIRECTORY_ACCESS_AS_USER_ALL = "Directory.AccessAsUser.All"
    USER_READ = "User.Read"
    USER_READ_WRITE = "User.ReadWrite"
    USER_READ_ALL = "User.Read.All"
    USER_READ_WRITE_ALL = "User.ReadWrite.All"
    OFFLINE_ACCESS = "offline_access"


class DeviceCodeRequestData(BaseModel):
    client_id: str
    scope: str


class TokenRequestData(BaseModel):
    grant_type: str
    client_id: str
    device_code: str


class VeeamTokenData(BaseModel):
    grant_type: str
    client_id: str
    assertion: str
    disable_antiforgery_token: bool
