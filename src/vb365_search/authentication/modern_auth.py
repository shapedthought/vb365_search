import requests
import json
import time
from enum import StrEnum
from typing import List

from .auth_models import DeviceCodeResponse, VeeamTokenResponse, AuthConfig


class M365Permissions(StrEnum):
    DIRECTORY_READ_ALL = "Directory.Read.All"
    DIRECTORY_READ_WRITE_ALL = "Directory.ReadWrite.All"
    DIRECTORY_ACCESS_AS_USER_ALL = "Directory.AccessAsUser.All"
    USER_READ = "User.Read"
    USER_READ_WRITE = "User.ReadWrite"
    USER_READ_ALL = "User.Read.All"
    USER_READ_WRITE_ALL = "User.ReadWrite.All"
    OFFLINE_ACCESS = "offline_access"

class AuthenticateModern:
    
    def __init__(self, config: AuthConfig):
        self.config = config
        self.device_code_url = f"https://login.microsoftonline.com/{self.config.tenant_name}/oauth2/v2.0/devicecode"
        self.token_url = f"https://login.microsoftonline.com/{self.config.tenant_name}/oauth2/v2.0/token"

    def authenticate_veeam_backup_m365(self, permissions: List[M365Permissions], verify: bool = False) -> VeeamTokenResponse:
        """
        Authenticate to Veeam Backup for Microsoft 365 using modern app-only authentication
        """
        
        permissions_str = AuthenticateModern.create_basic_permisions_str(permissions)
        
        # Step 1: Obtain device code
        device_code_data = {
            "client_id": self.config.client_id,
            "scope": permissions_str
        }
        
        device_code_response = requests.post(self.device_code_url, data=device_code_data)
        device_code_response.raise_for_status()
        
        device_code_json = device_code_response.json()
        device_code = DeviceCodeResponse.model_validate(device_code_json)
        
        print(f"Please visit: {device_code.verification_uri}")
        print(f"Enter the code: {device_code.user_code}")
        
        # Step 2: User authenticates with the code (manual step)
        print("After authenticating, press Enter to continue...")
        input()
        
        # Step 3: Obtain assertion with access token
        
        token_data = {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": self.config.client_id,
            "device_code": device_code.device_code
        }
        
        # Poll for token (may need to wait for user to complete authentication)
        max_attempts = 30
        token_json = {}
        
        for attempt in range(max_attempts):
            token_response = requests.post(self.token_url, data=token_data)
            token_response.raise_for_status()
            
            token_json = token_response.json()
            
            if 'error' in token_json and token_json['error'] == 'authorization_pending':
                print(f"Waiting for authentication... (attempt {attempt+1}/{max_attempts})")
                time.sleep(device_code.interval)
            else:
                break
        
        if 'access_token' not in token_json:
            raise ValueError("Failed to obtain access token from Microsoft Identity platform")
              
        # Step 4: Log in to Veeam Backup for Microsoft 365 REST API
        veeam_token_url = f"{self.config.veeam_api_url}/token"
        veeam_token_data: dict[str, str | bool] = {
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "client_id": self.config.tenant_name,
            "assertion": json.dumps(token_json),
            "disable_antiforgery_token": True
        }
        
        veeam_response = requests.post(veeam_token_url, data=veeam_token_data, verify=verify)
        veeam_response.raise_for_status()
        veeam_tokens = veeam_response.json()
        
        
        self.veeam_token_response = VeeamTokenResponse(
            access_token=veeam_tokens['access_token'],
            refresh_token=veeam_tokens['refresh_token'],
            expires_in=veeam_tokens['expires_in'],
            token_type=veeam_tokens['token_type'],
        )
        
        return self.veeam_token_response

    @staticmethod
    def create_basic_permisions_str(permissions: List[M365Permissions]) -> str:
        return " ".join(permissions)

    @staticmethod
    def create_default_permissions() -> List[M365Permissions]:
        return ([
            M365Permissions.DIRECTORY_ACCESS_AS_USER_ALL,
            M365Permissions.USER_READ_WRITE_ALL,
            M365Permissions.OFFLINE_ACCESS
        ])