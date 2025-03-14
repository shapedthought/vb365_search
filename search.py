import requests
import urllib3
import webbrowser
import tomllib
import toml
import pprint
import json
import fire
import time
import pyperclip as pc
from halo import Halo
from datetime import datetime, timezone
from pathlib import Path
import contextlib
import logging

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from models import (
    DeviceRequest,
    DeviceResponse,
    Assertion,
    AssertionResponse,
    VBLoginRequest,
    Configuration,
    VBLoginResponse,
    AuthHeaders,
    RestoreSessionRequest,
    RestoreSessionResponse,
    SearchRequest,
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

    print(config.microsoft.application_id)
    logger.info("Logging in...")

    application_id = config.microsoft.application_id
    tenant_id = config.microsoft.tenant_id
    tenant = config.microsoft.tenant_name
    version = config.vb365.version

    vb_address = config.vb365.api_address
    vb_base_url = f"https://{vb_address}:4443/{version}/"

    ms_login = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"

    device_body = DeviceRequest(
        client_id=application_id,
        scope="scope=Directory.AccessAsUser.All User.ReadWrite.All offline_access",
    )

    try:
        device_response = requests.post(
            ms_login, data=device_body.model_dump(), verify=False
        )
    except requests.exceptions.RequestException as e:
        print(f"Login failed: {e}")
        raise

    device_response_model = DeviceResponse(**device_response.json())

    user_code = device_response_model.user_code
    device_code = device_response_model.device_code

    pc.copy(user_code)
    print(
        f"User code {user_code} copied to clipboard. Please paste into web browser which will open when you continue ({ms_login})."
    )
    logger.info(f"User code {user_code} copied to clipboard.")
    input("Press Enter to continue...")
    webbrowser.open(device_response_model.verification_uri)

    # pause until user has logged in
    input("Once you have logged in, please press enter to continue...")

    spinner = Halo(text="Completing login...", spinner="arc")
    spinner.start()

    ms_token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    token_body = Assertion(
        grant_type="'urn:ietf:params:oauth:grant-type:device_code'",
        client_id=application_id,
        device_code=device_code,
    )

    try:
        api_res = requests.post(
            ms_token_url, data=token_body.model_dump(), verify=False
        )
    except requests.exceptions.RequestException as e:
        print(f"Login failed: {e}")
        raise

    api_res_json = AssertionResponse(**api_res.json())

    veeam_login_url = f"{vb_base_url}token"

    veeam_login_body = VBLoginRequest(
        grant_type=f"urn:ietf: params:oauth: grant - type:jwt - bearer & client_id ={tenant}",
        assertion=api_res_json.model_dump(),
    )

    # Pause
    time.sleep(3)

    try:
        vb365_res = requests.post(
            veeam_login_url, data=veeam_login_body.model_dump(), verify=False
        )
    except requests.exceptions.RequestException as e:
        print(f"Login failed: {e}")
        raise

    vb365_res_json = VBLoginResponse(**vb365_res.json())

    auth_headers = AuthHeaders(authorization=f"Bearer {vb365_res_json.access_token}")

    # Pause
    time.sleep(3)

    save_json(auth_headers.model_dump(), "auth_headers.json")

    spinner.succeed("Login complete!\n")
    logger.info("Login complete! Creating Restore Session...")

    restore_url = f"{vb_base_url}Organization/Explore"

    body = RestoreSessionRequest()

    # Pause
    time.sleep(3)

    try:
        restore_res = requests.post(
            restore_url,
            headers=auth_headers.model_dump(),
            json=body.model_dump(),
            verify=False,
        )
    except requests.exceptions.RequestException as e:
        print(f"Restore Session creation failed: {e}")
        raise

    restore_model = RestoreSessionResponse(**restore_res.json())

    restore_id = restore_model.id

    save_json(restore_model.model_dump(), "restore.json")

    logger.info(f"Restore Session created! ID: {restore_id}")


def search(term: str, print_results: bool = False, limit: int = 30):
    with open("auth_headers.json", mode="rb") as fp:
        auth_headers = AuthHeaders(**json.load(fp))

    with open("restore.json", mode="r") as fp:
        restore_model = RestoreSessionResponse(**json.load(fp))

    config = get_config()

    ex_search_url = f"https://{config.vb365.api_address}:4443/{config.vb365.version}/RestoreSessions/{restore_model.id}/organization/mailboxes/search?limit={limit}"

    search_body = SearchRequest(term=term)

    logger.info(f"Searching for {term}...")
    spinner = Halo(text="Searching...", spinner="dots")
    spinner.start()

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

    spinner.stop()

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
