from pydantic import BaseModel, Field


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
    client_id: str
    assertion: AssertionResponse


class VBLoginResponse(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str
    expires_in: int
    userName: str
    issued: str = Field(alias=".issued")
    expires: str = Field(alias=".expires")


# class AuthHeaders(BaseModel):
#     accept: str = Field(alias="Accept", default="application/json")
#     content_type: str = Field(alias="Content-Type", default="application/json")
#     authorization: str = Field(alias="Authorization")

#     @field_validator("authorization")
#     def check_authorization(cls, value: str):
#         if not value.startswith("Bearer "):
#             raise ValueError("Authorization must be a Bearer token")
