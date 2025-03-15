import requests
import urllib3
import webbrowser
import tomllib
import toml
import json
import fire
from datetime import datetime, timezone
from pathlib import Path
import contextlib
import logging

from src.vb365_search.restore_session.restore_models import RestoreSessionResponse, SearchRequest

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from src.vb365_search.authentication.modern_auth import AuthenticateModern, AuthConfig, AuthHeaders

from src.vb365_search.models.models import (
    Configuration,
    MicrosoftConfig,
    Vb365Config,
)

from search_model import SearchResponse

# Setup at the top of your file
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def save_json(data, filename) -> None:
    output_path = Path(filename)
    with open(output_path, "w") as f:
        json.dump(data, f, indent=4)
    return output_path


def get_config() -> Configuration:
    try:
        with open("configuration.toml", mode="rb") as fp:
            config = Configuration(**tomllib.load(fp))
        return config
    except (FileNotFoundError, tomllib.TOMLDecodeError) as e:
        print(f"Configuration error: {e}")
        raise


@contextlib.contextmanager
def vb365_session():
    with open("auth_headers.json", mode="r") as fp:
        auth_headers = AuthHeaders(**json.load(fp))
    session = requests.Session()
    session.headers.update(auth_headers.model_dump())
    session.verify = False
    try:
        yield session
    finally:
        session.close()


def create_config_template():
    """
    Create a configuration template file
    """
    config = Configuration(
        microsoft=MicrosoftConfig(),
        vb365=Vb365Config(),
    )

    with open("configuration.toml", mode="w") as fp:
        toml.dump(config.model_dump(), fp)


def login():
    """
    Login to Veeam Backup for Microsoft Office 365
    """
    config = get_config()

    logger.info("Logging in...")
    
    auth_config = AuthConfig(
        tenant_name=config.microsoft.tenant_name,
        client_id=config.microsoft.application_id,
        veeam_api_url=f"https://{config.vb365.api_address}:4443/{config.vb365.version}",
    )
    
    auth_modern = AuthenticateModern(auth_config)
    
    try:
        veeam_token_model = auth_modern.authenticate_veeam_backup_o365(verify=False)
    except Exception as e:
        print(f"Login failed: {e}")
        raise
    
    auth_headers = AuthHeaders(
        authorization
        =f"{veeam_token_model.token_type} {veeam_token_model.access_token}"
    )
    
    save_json(veeam_token_model.model_dump(), "auth_response.json")
    save_json(auth_headers.model_dump(), "auth_headers.json")
    


def search(term: str, print_results: bool = False, limit: int = 30):
    with open("auth_headers.json", mode="rb") as fp:
        auth_headers = AuthHeaders(**json.load(fp))

    with open("restore.json", mode="r") as fp:
        restore_model = RestoreSessionResponse(**json.load(fp))

    config = get_config()

    ex_search_url = f"https://{config.vb365.api_address}:4443/{config.vb365.version}/RestoreSessions/{restore_model.id}/organization/mailboxes/search?limit={limit}"

    search_body = SearchRequest(term=term)

    logger.info(f"Searching for {term}...")

    try:
        search_res = requests.post(
            ex_search_url,
            json=search_body.model_dump(),
            headers=auth_headers.model_dump(),
            verify=False,
        )
    except requests.exceptions.RequestException as e:
        print(f"Search failed: {e}")
        raise

    search_model = SearchResponse(**search_res.json())

    if print_results:
        for i in search_model.results:
            print(f"Subject: {i.subject}")
            print(f"Received: {i.received}")
            print(f"From: {i.from_}")
            print(f"Sent: {i.to}")
            print("")

    logger.info(f"Search complete! {len(search_model.results)} items found.")

    dt = datetime.now(timezone.utc)
    dt_str = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    save_str = f"search-{dt_str}.json"

    # After saving results
    output_path = save_json(search_model.model_dump(), save_str)
    logger.info(f"Search results saved to {output_path}")

    return search_model  # Return the model for potential further use


def logout():
    try:
        with open("auth_headers.json", mode="r") as fp:
            auth_headers = AuthHeaders(**json.load(fp))

        with open("restore.json", mode="r") as fp:
            restore_model = RestoreSessionResponse(**json.load(fp))

        config = get_config()

        logout_url = f"https://{config.vb365.api_address}:4443/{config.vb365.version}/RestoreSessions/{restore_model.id}/Stop"

        session = requests.Session()
        logout_res = session.post(
            logout_url, headers=auth_headers.model_dump(), verify=False
        )
        logout_res.raise_for_status()
        print("Log out successful!")
    except requests.exceptions.RequestException as e:
        print(f"Logout failed: {e}")
    finally:
        session.close()


def main():
    fire.Fire(
        {
            "login": login,
            "search": search,
            "template": create_config_template,
            "logout": logout,
        }
    )


if __name__ == "__main__":
    main()
