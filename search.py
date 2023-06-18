import requests
import urllib3
import webbrowser
import tomllib
import pprint
import json
import fire
import time
import pyperclip as pc
from halo import Halo
from datetime import datetime
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def save_json(data, filename):
    with open(filename, 'w') as f:
        json.dump(data, f, indent=4)

def get_config():
    with open("configuration.toml", mode="rb") as fp:
        config = tomllib.load(fp)
    return config

def login():

    config = get_config()

    print(config['microsoft']['application_id'])

    # client_id = config['microsoft']['client_id']
    application_id = config['microsoft']['application_id']
    tenant_id = config['microsoft']['tenant_id']
    user_id = config['microsoft']['user_id']
    user_tenant = f"{user_id}.{tenant_id}"

    vb_address = config['vb365']['api_address']
    vb_base_url = veeam_login_url = f"https://{vb_address}:4443/v7/"

    ms_login = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"

    device_body = {
        "client_id": application_id,
        'scope': [f'api://{application_id}/access_as_user openid profile offline_access']
    }

    device_response = requests.post(ms_login, data=device_body, verify=False)
    device_response.raise_for_status()
    device_response_json = device_response.json()
    user_code = device_response_json['user_code']
    device_code = device_response_json['device_code']

    pc.copy(user_code)
    print(f"User code {user_code} copied to clipboard. Please paste into webroswer which will open when you continue.")
    input("Press Enter to continue...")
    webbrowser.open(device_response_json['verification_uri'])

    # pause until user has logged in
    input("Once you have logged in, please press enter to continue...")

    ms_token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_body = {
        'grant_type': 'urn:ietf:params:oauth:grant-type:device_code',
        'client_id': application_id,
        'device_code': device_code
    }

    api_res = requests.post(ms_token_url, data=token_body, verify=False)
    api_res.raise_for_status()

    api_res_json = api_res.json()
    access_token = api_res_json['access_token']

    veeam_login_url = f"{vb_base_url}token"
    veeam_login_body = {
        'grant_type': 'operator',
        'client_id': user_tenant,
        'assertion': access_token
    }

    # Pause
    time.sleep(3)

    vb365_res = requests.post(veeam_login_url, data=veeam_login_body, verify=False)
    vb365_res.raise_for_status()

    vb365_res_json = vb365_res.json()

    restore_headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Bearer " + vb365_res_json["access_token"]
        }

    # Standard login
    username = config['vb365']['username']
    password = config['vb365']['password']

    standard_body = {
        "grant_type": "password",
        "username": username,
        "password": password
        }

    # Pause
    time.sleep(3)
    
    standard_res = requests.post(veeam_login_url, data=standard_body, verify=False)
    standard_res.raise_for_status()
    standard_json = standard_res.json()

    standard_accee_token = standard_json["access_token"]
    standard_headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Bearer " + standard_accee_token
    }

    save_json(restore_headers, "restore_headers.json")
    save_json(standard_headers, "standard_headers.json")

    print("Logins successful! Creating Restore Session...")

    restore_url = "https://192.168.0.219:4443/v7/Organization/Explore"

    
    dt = datetime.utcnow()
    dt_str = dt.strftime('%Y-%m-%dT%H:%M:%SZ')

    body = {
        "dateTime": dt_str,
        "showAllVersions": True,
        "showDeleted": True,
        "type": "Vex"
    }

    # Pause
    time.sleep(3)

    restore_res = requests.post(restore_url, headers=restore_headers, json=body, verify=False)

    restore_json = restore_res.json()

    restore_id = restore_json['id']

    save_json(restore_json, "restore.json")

    print(f"Restore Session created! ID: {restore_id}")


def search(term: str, print_results: bool = False, limit: int = 30):
    with open("restore_headers.json", mode="rb") as fp:
        restore_headers = json.load(fp)

    with open("restore.json", mode="r") as fp:
        restore_json = json.load(fp)

    restore_id = restore_json['id']

    config = get_config()
    vb_address = config['vb365']['api_address']

    ex_search_url = f"https://{vb_address}:4443/v7/RestoreSessions/{restore_id}/organization/mailboxes/search?limit={limit}"

    search_body = {
        "query": term
    }

    print("Search term: " + term)
    spinner = Halo(text='Searching...', spinner='dots')
    spinner.start()
    search_res = requests.post(ex_search_url, json=search_body, headers=restore_headers, verify=False)
    spinner.stop()
    
    search_res.raise_for_status()

    search_json = search_res.json()

    # dict_keys(['subject', 'itemClass', '_links', '_actions', 'id', 'from', 'cc', 'bcc', 'to', 'sent', 'received', 'reminder', 'importance'])
    if print_results:
        for i in search_json['results']:
            print(f"Subject: {i['subject']}")
            print(f"Received: {i['received']}")
            print(f"From: {i['from']}")
            print(f"Sent: {i['to']}")
            print("")

    print(f"Search complete! {len(search_json['results'])} items found.")

    dt = datetime.utcnow()
    dt_str = dt.strftime('%Y-%m-%dT%H:%M:%SZ')
    save_str = f"search-{dt_str}.json"

    save_json(search_json, save_str)

    print(f"Search results saved to {save_str}")

def logout():
    #https://localhost:4443/v7/RestoreSessions/{restoreSessionId}/Stop
    with open("restore_headers.json", mode="rb") as fp:
        restore_headers = json.load(fp)

    with open("restore.json", mode="r") as fp:
        restore_json = json.load(fp)

    restore_id = restore_json['id']

    config = get_config()
    vb_address = config['vb365']['api_address']

    logout_url = f"https://{vb_address}:4443/v7/RestoreSessions/{restore_id}/Stop"

    logout_res = requests.post(logout_url, headers=restore_headers, verify=False)

    print(logout_res)

    print("Log out succesful!")

def main():
    fire.Fire({
        'login': login,
        'search': search,
        'logout': logout
  })

if __name__ == "__main__":
    main()