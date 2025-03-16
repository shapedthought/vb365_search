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