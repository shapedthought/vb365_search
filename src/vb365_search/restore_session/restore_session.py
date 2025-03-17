import requests

from vb365_search.restore_session.restore_models import RestoreSessionRequest, RestoreSessionResponse
from vb365_search.authentication.auth_models import AuthConfig, AuthHeaders

class RestoreSession:
    def __init__(self, config: AuthConfig, auth_headers: AuthHeaders):
        self.config = config
        self.auth_headers = auth_headers
        self.restore_session_url = f"{self.config.veeam_api_url}/Organization/Explore"

    def create_restore_session(self, restore_session_request: RestoreSessionRequest, verify: bool = True) -> RestoreSessionResponse:
        """
        Create a restore session
        """
        try:
            restore_response = requests.post(
                self.restore_session_url,
                json=restore_session_request.model_dump(by_alias=True),
                headers=self.auth_headers.model_dump(),
                verify=verify,
            )
            restore_response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Restore session failed: {e}")
            raise
        
        restore_json = restore_response.json()
        print(restore_json)
        return RestoreSessionResponse(**restore_json)
