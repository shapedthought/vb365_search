import requests
import json
import time

from auth_models import DeviceCodeResponse, VeeamTokenResponse, AuthConfig

class AuthenticateModern:
    
    def __init__(self, config: AuthConfig):
        self.config = config
        self.device_code_url = f"https://login.microsoftonline.com/{self.config.tenant_name}/oauth2/v2.0/devicecode"
        self.token_url = f"https://login.microsoftonline.com/{self.config.tenant_name}/oauth2/v2.0/token"

    def authenticate_veeam_backup_o365(self) -> VeeamTokenResponse:
        """
        Authenticate to Veeam Backup for Microsoft 365 using modern app-only authentication
        """
        # Step 1: Obtain device code
        device_code_data = {
            "client_id": self.config.client_id,
            "scope": "Directory.AccessAsUser.All User.ReadWrite.All offline_access"
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
        veeam_token_url = f"{self.config.veeam_api_url}/v8/token"
        veeam_token_data = {
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "client_id": self.config.tenant_name,
            "assertion": json.dumps(token_json)
        }
        
        veeam_response = requests.post(veeam_token_url, data=veeam_token_data, verify=False)
        veeam_response.raise_for_status()
        veeam_tokens = veeam_response.json()
        
        self.veeam_token_response = VeeamTokenResponse.model_validate(veeam_tokens)
        
        return self.veeam_token_response


