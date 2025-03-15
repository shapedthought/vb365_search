import requests

from restore_models import RestoreSessionRequest, RestoreSessionResponse
from ..authentication.auth_models import AuthConfig, AuthHeaders

class RestoreSession:
    def __init__(self, config: AuthConfig, auth_headers: AuthHeaders):
        self.config = config
        self.auth_headers = auth_headers
        self.restore_session_url = f"https://{self.config.veeam_api_url}/RestoreSessions"

    def create_restore_session(self, restore_session_request: RestoreSessionRequest, verify: bool = False) -> RestoreSessionResponse:
        """
        Create a restore session
        """
        try:
            restore_response = requests.post(
                self.restore_session_url,
                json=restore_session_request.model_dump(),
                headers=self.auth_headers.model_dump(),
                verify=verify,
            )
            restore_response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Restore session failed: {e}")
            raise

        return RestoreSessionResponse(**restore_response.json())