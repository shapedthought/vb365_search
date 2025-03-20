import requests
from typing import Any, Dict, Optional, Type, TypeVar, Union

from vb365_search.restore_session.restore_models import (
    RestoreSessionRequest,
    RestoreSessionResponse,
    RestoreSessionsResponse,
    RestoreStatisticsResponse,
)
from vb365_search.authentication.auth_models import AuthConfig, AuthHeaders
from vb365_search.utils.helpers import load_json, save_json

T = TypeVar("T")


class RestoreSession:
    def __init__(
        self, config: AuthConfig, auth_headers: AuthHeaders, verify: bool = True
    ):
        self.config = config
        self.auth_headers = auth_headers
        self.verify = verify
        self.explore_session_url = f"{self.config.veeam_api_url}/Organization/Explore"
        self.restore_session_url = f"{self.config.veeam_api_url}/RestoreSessions"

    def _make_request(
        self,
        method: str,
        url: str,
        response_model: Optional[Type[T]] = None,
        json_data: Optional[Dict[str, Any]] = None,
        error_message: str = "Request failed",
    ) -> Union[T, None]:
        """
        Make an HTTP request with error handling

        Args:
            method: HTTP method (get, post, etc.)
            url: URL to make the request to
            response_model: Pydantic model to parse the response into
            json_data: JSON data to send in the request
            error_message: Error message to display if the request fails

        Returns:
            The parsed response or None for requests without responses
        """
        try:
            response = requests.request(
                method=method,
                url=url,
                headers=self.auth_headers.model_dump(),
                json=json_data,
                verify=self.verify,
            )
            response.raise_for_status()

            if response_model and response.status_code != 204:
                response_json = response.json()
                save_json(response_json, "response.json")
                return response_model(**response_json)
            return None

        except requests.exceptions.RequestException as e:
            print(f"{error_message}: {e}")
            raise

    def create_restore_session(
        self, restore_session_request: RestoreSessionRequest, save: bool = False
    ) -> RestoreSessionResponse:
        """Create a restore session"""
        json_data = restore_session_request.model_dump(by_alias=True)
        response = self._make_request(
            method="post",
            url=self.explore_session_url,
            response_model=RestoreSessionResponse,
            json_data=json_data,
            error_message="Restore session creation failed",
        )
        # We know response will be RestoreSessionResponse based on the type parameter
        cast_response = RestoreSessionResponse.model_validate(response)

        if save:
            save_json(cast_response.model_dump(), "restore_session_response.json")

        return cast_response

    def get_restore_session(self, restore_session_id: str) -> RestoreSessionResponse:
        """
        Get a restore session

        Args:
            restore_session_id: Restore session ID

        Returns:
            RestoreSessionResponse object
        """

        response = self._make_request(
            method="get",
            url=f"{self.explore_session_url}/{restore_session_id}",
            response_model=RestoreSessionResponse,
            error_message=f"Get restore session {restore_session_id} failed",
        )
        return RestoreSessionResponse.model_validate(response)

    def get_all_restore_sessions(self) -> RestoreSessionsResponse:
        """
        Get all restore sessions

        Returns:
            RestoreSessionsResponse object
        """
        respone = self._make_request(
            method="get",
            url=self.restore_session_url,
            response_model=RestoreSessionsResponse,
            error_message="Get all restore sessions failed",
        )
        return RestoreSessionsResponse.model_validate(respone)

    def stop_restore_session(self, restore_session_id: str) -> None:
        """
        Stop a restore session

        Args:
            restore_session_id: Restore session ID
        """
        self._make_request(
            method="post",
            url=f"{self.restore_session_url}/{restore_session_id}/Stop",
            error_message=f"Stop restore session {restore_session_id} failed",
        )

    def stop_all_restore_sessions(self) -> None:
        """Stop all restore sessions"""
        all_restore_sessions = self.get_all_restore_sessions().results
        for restore_session in all_restore_sessions:
            self.stop_restore_session(restore_session.id)

    def get_restore_statistics(
        self, restore_session_id: str
    ) -> RestoreStatisticsResponse:
        """
        Get restore statistics

        Args:
            restore_session_id: Restore session ID

        Returns:
            Restore statistics
        """
        response = self._make_request(
            method="get",
            url=f"{self.restore_session_url}/{restore_session_id}/Statistics",
            response_model=RestoreStatisticsResponse,
            error_message=f"Get restore statistics for session {restore_session_id} failed",
        )
        return RestoreStatisticsResponse.model_validate(response)

    @staticmethod
    def session_from_file(file_path: str) -> RestoreSessionResponse:
        """Load a restore session from a file"""
        data = load_json(file_path)
        response = RestoreSessionResponse(**data)
        return response
