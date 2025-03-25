from vb365_search.authentication.auth_models import StandardAuthHeaders
from vb365_search.models.models import Configuration
from .organisations_models import OrganisationResponseModel
import requests


class Organisation:
    def __init__(
        self,
        config: Configuration,
        auth_headers: StandardAuthHeaders,
        verify: bool = True,
    ):
        self.auth_headers = auth_headers
        self.config = config
        self.verify = verify
        self.organisations_url = f"https://{self.config.vb365.api_address}:4443/{self.config.vb365.version}/Organizations"

    def get_organisations(self) -> OrganisationResponseModel:
        response = None
        try:
            response = requests.get(
                url=self.organisations_url,
                headers=self.auth_headers.model_dump(),
                verify=self.verify,
            )
            response.raise_for_status()
        except requests.exceptions.HTTPError as e:
            raise requests.exceptions.HTTPError(
                f"Failed to fetch organisations. Status code: {response.status_code if response else 'N/A'}, Response: {response.text if response else 'No response'}"
            ) from e

        response_json = response.json()
        organisations = OrganisationResponseModel(**response_json)
        return organisations
