from pydantic import BaseModel
import requests
import json
import time
from typing import List, Dict, Any, Optional
from requests.exceptions import RequestException

from vb365_search.authentication.auth_models import AuthHeaders

from .auth_models import (
    DeviceCodeRequestData,
    DeviceCodeResponse,
    M365Permissions,
    TokenRequestData,
    VeeamTokenData,
    VeeamTokenResponse,
    AuthConfig,
)


class AuthenticateModern:
    def __init__(self, config: AuthConfig, verify: bool = True):
        self.config = config
        self.verify = verify
        self.device_code_url = f"https://login.microsoftonline.com/{self.config.tenant_name}/oauth2/v2.0/devicecode"
        self.token_url = f"https://login.microsoftonline.com/{self.config.tenant_name}/oauth2/v2.0/token"
        self.default_timeout = 30  # Default timeout for requests in seconds

    def _make_request(
        self,
        url: str,
        data: BaseModel,
    ) -> Dict[str, Any]:
        """
        Make an HTTP POST request with error handling and proper timeouts
        """
        try:
            response = requests.post(
                url,
                data=data.model_dump(),
                verify=self.verify,
                timeout=self.default_timeout,
            )
            response.raise_for_status()
            return response.json()
        except RequestException as e:
            raise ValueError(f"API request failed: {str(e)}") from e

    def _get_device_code(self, permissions_str: str) -> DeviceCodeResponse:
        """Get device code from Microsoft Identity platform"""
        device_code_data = DeviceCodeRequestData(
            client_id=self.config.client_id,
            scope=permissions_str,
        )

        device_code_json = self._make_request(self.device_code_url, device_code_data)
        return DeviceCodeResponse.model_validate(device_code_json)

    def _poll_for_token(self, device_code: DeviceCodeResponse) -> Dict[str, Any]:
        """Poll for token after user authentication"""
        token_data = TokenRequestData(
            grant_type="urn:ietf:params:oauth:grant-type:device_code",
            client_id=self.config.client_id,
            device_code=device_code.device_code,
        )

        max_attempts = 30

        for attempt in range(max_attempts):
            try:
                token_json = self._make_request(self.token_url, token_data)

                if (
                    "error" in token_json
                    and token_json["error"] == "authorization_pending"
                ):
                    print(
                        f"Waiting for authentication... (attempt {attempt + 1}/{max_attempts})"
                    )
                    time.sleep(device_code.interval)
                else:
                    return token_json
            except ValueError:
                # If request fails, wait and try again
                time.sleep(device_code.interval)

        raise ValueError("Maximum polling attempts reached. Authentication timed out.")

    def _get_veeam_token(self, token_json: Dict[str, Any]) -> VeeamTokenResponse:
        """Get Veeam token using the Microsoft token assertion"""
        veeam_token_url = f"{self.config.veeam_api_url}/token"
        veeam_token_data = VeeamTokenData(
            grant_type="urn:ietf:params:oauth:grant-type:jwt-bearer",
            client_id=self.config.tenant_name,
            assertion=json.dumps(token_json),
            disable_antiforgery_token=True,
        )

        veeam_tokens = self._make_request(veeam_token_url, veeam_token_data)

        return VeeamTokenResponse(
            access_token=veeam_tokens["access_token"],
            refresh_token=veeam_tokens["refresh_token"],
            expires_in=veeam_tokens["expires_in"],
            token_type=veeam_tokens["token_type"],
        )

    def authenticate(
        self, permissions: Optional[List[M365Permissions]] = None
    ) -> VeeamTokenResponse:
        """
        Authenticate to Veeam Backup for Microsoft 365 using modern app-only authentication.

        This method initiates a device code flow authentication process with Microsoft Identity Platform,
        requiring the user to manually authenticate using a web browser, and then exchanges the
        resulting token for a Veeam-specific access token.

        Args:
            permissions: Optional list of Microsoft 365 permissions to request.
                         If None, default permissions will be used.

        Returns:
            VeeamTokenResponse object containing the access token, refresh token,
            token expiration information, and token type.

        Raises:
            ValueError: If authentication fails or times out
        """
        if permissions is None:
            permissions = AuthenticateModern.create_default_permissions()

        AuthenticateModern.verify_required_permissions(permissions)

        permissions_str = AuthenticateModern.create_default_permissions_str(permissions)

        # Step 1: Obtain device code
        device_code = self._get_device_code(permissions_str)
        print(f"Please visit: {device_code.verification_uri}")
        print(f"Enter the code: {device_code.user_code}")

        # Step 2: User authenticates with the code (manual step)
        print("After authenticating, press Enter to continue...")
        input()

        # Step 3: Obtain assertion with access token
        token_json = self._poll_for_token(device_code)

        if "access_token" not in token_json:
            raise ValueError(
                "Failed to obtain access token from Microsoft Identity platform"
            )

        # Step 4: Get Veeam token
        self.veeam_token_response = self._get_veeam_token(token_json)
        return self.veeam_token_response

    def authenticate_return_headers(
        self,
        permissions: Optional[List[M365Permissions]] = None,
        file_name: Optional[str] = None,
    ) -> AuthHeaders:
        """
        Authenticate and return authorization headers for Veeam Backup for Microsoft 365 API calls.

        This method extends the authenticate method by formatting the resulting token into HTTP
        authorization headers and optionally saving them to a file for future use.

        Args:
            permissions: Optional list of Microsoft 365 permissions to request.
                         If None, default permissions will be used.
            file_name: Optional path to save the authentication headers as JSON.
                       If provided, headers will be saved to this file.

        Returns:
            AuthHeaders object containing the formatted authorization header

        Raises:
            ValueError: If authentication fails or if file_name is None
            IOError: If there's an error writing to the specified file
        """

        veeam_token_model = self.authenticate(permissions)
        auth_headers = AuthHeaders(
            Authorization=f"{veeam_token_model.token_type} {veeam_token_model.access_token}"
        )

        if file_name:
            with open(file_name, "w") as file:
                json.dump(auth_headers.model_dump(), file)
        else:
            raise ValueError("File path required to save headers")

        return auth_headers

    @staticmethod
    def create_default_permissions_str(permissions: List[M365Permissions]) -> str:
        """
        Converts a list of Microsoft 365 permission enums to a space-separated string.

        Args:
            permissions: List of M365Permissions enum values

        Returns:
            A space-separated string of permissions suitable for OAuth requests
        """
        return " ".join(permissions)

    @staticmethod
    def create_default_permissions() -> List[M365Permissions]:
        """
        Creates the default list of Microsoft 365 permissions required for Veeam Backup authentication.

        Returns:
            List of default M365Permissions enum values including directory access,
            user read/write permissions, and offline access
        """
        return [
            M365Permissions.DIRECTORY_ACCESS_AS_USER_ALL,
            M365Permissions.USER_READ_WRITE_ALL,
            M365Permissions.OFFLINE_ACCESS,
        ]

    @staticmethod
    def auth_headers_from_file(file_path: str) -> AuthHeaders:
        """
        Loads authentication headers from a saved JSON file.

        Args:
            file_path: Path to the JSON file containing authentication headers

        Returns:
            AuthHeaders object with loaded authentication information

        Raises:
            FileNotFoundError: If the specified file path does not exist
            JSONDecodeError: If the file contains invalid JSON
        """
        with open(file_path, "r") as file:
            auth_headers = json.load(file)

        return AuthHeaders(**auth_headers)

    @staticmethod
    def verify_required_permissions(permissions: List[M365Permissions]) -> None:
        """
        Verify that the permissions list contains exactly one user permission,
        one directory permission, one offline access permission, and a total of three permissions.

        Args:
            permissions: List of M365Permissions to check

        Raises:
            ValueError: If the required permissions are missing or if there are extra permissions
        """
        # Count types of permissions
        directory_permissions = [p for p in permissions if p.startswith("Directory.")]
        user_permissions = [p for p in permissions if p.startswith("User.")]
        offline_access = [p for p in permissions if p == M365Permissions.OFFLINE_ACCESS]

        # Check counts
        errors: List[str] = []

        if len(directory_permissions) != 1:
            errors.append(
                f"Expected exactly one Directory permission, found {len(directory_permissions)}"
            )

        if len(user_permissions) != 1:
            errors.append(
                f"Expected exactly one User permission, found {len(user_permissions)}"
            )

        if len(offline_access) != 1:
            errors.append(
                f"Expected exactly one {M365Permissions.OFFLINE_ACCESS} permission, found {len(offline_access)}"
            )

        if len(permissions) != 3:
            errors.append(
                f"Expected exactly 3 total permissions, found {len(permissions)}"
            )

        if errors:
            error_message = "\n".join(errors)
            raise ValueError(
                f"Permission validation failed:\n{error_message}\n\nRequired permissions must include exactly one Directory permission, one User permission, and the {M365Permissions.OFFLINE_ACCESS} permission."
            )
