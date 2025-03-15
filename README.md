# VB365 Search

A Python library for searching in Veeam Backup for Microsoft 365 environments.

## Features

- Search Exchange mailboxes in Veeam Backup for Microsoft 365
- Extensible framework for adding additional search types (SharePoint, OneDrive, etc.)
- Command-line interface for easy use
- Configurable via TOML configuration files
- Modern authentication with Microsoft Identity Platform

## Installation

```bash
pip install vb365-search
```

## Configuration

Create a configuration file using the template command:

```bash
vb365 template
```

This will create a `configuration.toml` file with the following structure:

```toml
[microsoft]
tenant_name = ""
tenant_id = ""
application_id = ""

[vb365]
api_address = ""
username = ""
password = ""
version = "v8"
```

### Configuration Options

#### Microsoft Section

- `tenant_name`: Your Microsoft 365 tenant name (e.g., "contoso.onmicrosoft.com")
- `tenant_id`: Your Microsoft 365 tenant ID
- `application_id`: The Application (client) ID of the Azure AD application

#### Veeam Backup for Microsoft 365 Section

- `api_address`: The address of the Veeam Backup for Microsoft 365 server
- `username`: The username for the Veeam Backup for Microsoft 365 server (optional)
- `password`: The password for the Veeam Backup for Microsoft 365 server (optional)
- `version`: The API version to use (default: "v8")

## Usage

### Command-Line Interface

#### Login

```bash
vb365 login
```

This will authenticate with Microsoft Identity Platform and create a restore session.

#### Search

```bash
vb365 search --term "subject: Test" --search-type exchange --limit 100 --print-results
```

Options:
- `--term`: The search term or query
- `--search-type`: The type of search to perform (exchange, sharepoint, onedrive)
- `--limit`: The maximum number of results to return (default: 30)
- `--print-results`: Whether to print results to the console (default: False)
- `--config-path`: Path to the configuration file (optional)

#### Logout

```bash
vb365 logout
```

This will stop the restore session.

### Python API

```python
from vb365_search.authentication.modern_auth import AuthenticateModern
from vb365_search import ExchangeSearch
from vb365_search.config import load_config, auth_from_config
from vb365_search.utils.helpers import load_json, save_json, headers_from_veeam_token_response, auth_from_config
from vb365_search.restore_session.restore_models import RestoreSessionRequest, RestoreSessionResponse, RestoreSession
from vb365_search.search.exchange import ExchangeItemsInMailboxesSearch

# Load configuration
config = load_config("configuration.toml")

# Create Authentication object
auth_config = auth_from_config(config)

# Authenticate using modern auth, returns a VeeamTokenResponse
auth_modern = AuthenticateModern(auth_config)

try:
    veeam_token_response = auth_modern.authenticate_veeam_backup_o365()
except Exception as e:
    print(f"Login failed: {e}")
    raise

# Create authentication headers from the VeeamTokenResponse
auth_headers = headers_from_veeam_token_response(veeam_token_response)

# Create authentication headers from the VeeamTokenResponse
# Defaults to date_time: None, show_all_versions: True, show_deleted: True, type_restore: Vex
# Note that each restore session is locked to a specific type
restore_session_request = RestoreSessionRequest()

# Create restore session object, having this independent allows for multiple sessions to be created
restore_session = RestoreSession(
    config=config,
    auth_headers=auth_headers
)

# Then create the restore session
restore_session_response = restore_session.create_restore_session(
    restore_session_request=restore_session_request,
    verify=False
)

# Then create the Exchange Items Mailboxes Search object
search = ExchangeItemsInMailboxesSearch(
    config=config,
    auth_headers=auth_headers,
    restore_session_id=restore_session_response.id
)

# Execute search
results = search.search("subject: Test", limit=100)

# Print results
pprint(results)

# Save results to file
save_json(results.model_dump())
```

## Search Query Syntax

The search query syntax follows the Veeam Backup for Microsoft 365 search syntax. See the [Veeam documentation](https://helpcenter.veeam.com/docs/vbo365/rest/search.html) for details.

Examples:
- `subject: Test` - Search for emails with "Test" in the subject
- `from: user@example.com` - Search for emails from a specific sender
- `received: 2023-01-01..2023-01-31` - Search for emails received in January 2023
- `has:attachments` - Search for emails with attachments

## Development

### Project Structure

```
vb365_search/
├── pyproject.toml        # Modern Python packaging
├── setup.py              # For backward compatibility
├── src/
│   └── vb365_search/
│       ├── __init__.py   # Package version and exports
│       ├── cli.py        # CLI entry points using Fire
│       ├── config.py     # Configuration handling
│       ├── authentication/
│       │   ├── __init__.py
│       │   ├── auth_models.py
│       │   └── modern_auth.py
│       ├── models/
│       │   ├── __init__.py
│       │   └── models.py
│       ├── restore_session/
│       │   ├── __init__.py
│       │   ├── restore_models.py
│       │   └── restore_session.py
│       ├── search/
│       │   ├── __init__.py
│       │   ├── base.py           # Base search class
│       │   ├── exchange.py       # Exchange-specific search
│       │   ├── sharepoint.py     # SharePoint search
│       │   ├── onedrive.py       # OneDrive search
│       │   └── models.py         # Search models
│       └── utils/
│           ├── __init__.py
│           └── helpers.py
└── tests/
    ├── __init__.py
    ├── test_authentication/
    ├── test_restore_session/
    └── test_search/
        └── test_exchange.py
```

### Running Tests

```bash
pytest
```

## License

MIT

## Credits

This project is based on the Veeam Backup for Microsoft 365 REST API.
