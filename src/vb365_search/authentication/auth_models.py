from enum import StrEnum
from typing import Optional
from pydantic import BaseModel, Field, HttpUrl


class DeviceCodeResponse(BaseModel):
    """
    Response model from the OAuth device code authorization flow.

    Contains the information needed for the user to complete the device authentication.
    """

    user_code: str = Field(
        description="Code that user needs to enter on the verification page"
    )
    device_code: str = Field(
        description="Code used by the application to poll for authentication status"
    )
    verification_uri: HttpUrl = Field(
        description="URL where the user enters the user_code"
    )
    expires_in: int = Field(description="Time in seconds until the device code expires")
    interval: int = Field(description="Recommended polling interval in seconds")


class TokenResponse(BaseModel):
    """
    Standard OAuth token response model.

    Contains the tokens received after successful authentication.
    """

    access_token: str = Field(description="Token used to access protected resources")
    refresh_token: Optional[str] = Field(
        None, description="Token used to obtain new access tokens"
    )
    id_token: Optional[str] = Field(
        None,
        description="JWT containing user identity information (for OpenID Connect)",
    )
    expires_in: int = Field(
        description="Time in seconds until the access token expires"
    )
    token_type: str = Field(description="Type of token, typically 'Bearer'")


class VeeamTokenResponse(BaseModel):
    """
    Veeam-specific OAuth token response model.

    Contains the tokens received after authenticating with Veeam's authentication system.
    """

    access_token: str = Field(description="Token used to access Veeam API resources")
    refresh_token: str = Field(description="Token used to renew the access token")
    expires_in: int = Field(
        description="Time in seconds until the access token expires"
    )
    token_type: str = Field(description="Type of token, typically 'Bearer'")


class AuthConfig(BaseModel):
    """
    Authentication configuration model.

    Contains the necessary parameters to establish authentication with services.
    """

    tenant_name: str = Field(
        description="Name of the tenant for multi-tenant applications"
    )
    client_id: str = Field(description="Application identifier")
    veeam_api_url: HttpUrl = Field(description="Base URL for the Veeam API")


class AuthHeaders(BaseModel):
    """
    HTTP Authentication headers for Veeam Backup for Microsoft 365.

    Used to authenticate API requests after a successful OAuth flow.
    """

    accept: str = Field(
        alias="Accept",
        default="application/json",
        description="Accepted response format",
    )
    """Authorization header containing token type and access token"""
    content_type: str = Field(
        alias="Content-Type",
        default="application/json",
        description="Request body format",
    )
    authorization: str = Field(
        alias="Authorization", description="Authorization header containing the token"
    )


class StandardAuthHeaders(AuthHeaders):
    """
    Standard HTTP Authentication headers.

    Used to authenticate API requests after a successful OAuth flow.
    """

    pass


class M365Permissions(StrEnum):
    """
    Enumeration of Microsoft 365 API permissions.

    Used to specify the required permissions when requesting access tokens.
    """

    DIRECTORY_READ_ALL = "Directory.Read.All"
    DIRECTORY_READ_WRITE_ALL = "Directory.ReadWrite.All"
    DIRECTORY_ACCESS_AS_USER_ALL = "Directory.AccessAsUser.All"
    USER_READ = "User.Read"
    USER_READ_WRITE = "User.ReadWrite"
    USER_READ_ALL = "User.Read.All"
    USER_READ_WRITE_ALL = "User.ReadWrite.All"
    OFFLINE_ACCESS = "offline_access"


class DeviceCodeRequestData(BaseModel):
    """
    Request model for initiating a device code flow.

    Contains the parameters needed to request a device code.
    """

    client_id: str = Field(description="Application identifier")
    scope: str = Field(description="Space-separated list of requested permissions")


class TokenRequestData(BaseModel):
    """
    Request model for token exchange using device code.

    Contains the parameters needed to exchange a device code for tokens.
    """

    grant_type: str = Field(
        description="Type of grant, typically 'urn:ietf:params:oauth:grant-type:device_code'"
    )
    client_id: str = Field(description="Application identifier")
    device_code: str = Field(
        description="Device code received from the device code response"
    )


class VeeamTokenData(BaseModel):
    """
    Request model for obtaining Veeam-specific tokens.

    Contains the parameters needed for token requests to Veeam's authentication system.
    """

    grant_type: str = Field(description="Type of grant")
    client_id: str = Field(description="Application identifier")
    assertion: str = Field(description="Token assertion or JWT")
    disable_antiforgery_token: bool = Field(
        description="Whether to disable anti-forgery token validation"
    )


class VeeamPassWordRequest(BaseModel):
    """
    Request model for obtaining Veeam-specific tokens.

    Contains the parameters needed for token requests to Veeam's authentication system.
    """

    grant_type: str = Field(description="Type of grant")
    username: str = Field(description="username")
    password: str = Field(description="password")
    disable_antiforgery_token: bool = Field(
        description="Whether to disable anti-forgery token validation", default=True
    )
