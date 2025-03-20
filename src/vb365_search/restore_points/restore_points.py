
from typing import List
from vb365_search.authentication.auth_models import StandardAuthHeaders
from .restore_points_models import RestorePointResponse
from vb365_search.models.models import Configuration
import requests

class RestorePoints():
    def __init__(self, config: Configuration, auth_headers: StandardAuthHeaders, verify: bool = True):
        self.auth_headers = auth_headers
        self.config = config
        self.verify = verify
        self.restore_points_url = f"https://{self.config.vb365.api_address}:4443/v8/RestorePoints"
        
    def get_restore_points(self) -> RestorePointResponse:
        response = requests.get(
            url=self.restore_points_url,
            headers=self.auth_headers.model_dump(),
            verify=self.verify
        )
        response.raise_for_status()
        
        response_json = response.json()
        restore_points = RestorePointResponse(**response_json)
        return restore_points
    
    def get_restore_point_dates(self) -> List[str]:
        restore_points = self.get_restore_points()
        restore_point_dates = [restore_point.creationTime for restore_point in restore_points.data]
        return restore_point_dates