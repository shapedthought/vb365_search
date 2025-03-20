"""
vb365_search - A Python library for searching in Veeam Backup for Microsoft 365 environments
"""

__version__ = "0.1.0"

from .authentication.modern_auth import AuthenticateModern
from .authentication.auth_models import AuthConfig, AuthHeaders
from .models.models import Configuration
from .search import (
    BaseSearch,
    ExchangeItemsInMailboxesSearch,
    SharePointSearch,
    OneDriveSearch,
)

__all__ = [
    "AuthenticateModern",
    "AuthConfig",
    "AuthHeaders",
    "Configuration",
    "BaseSearch",
    "ExchangeItemsInMailboxesSearch",
    "SharePointSearch",
    "OneDriveSearch",
]
