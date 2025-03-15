"""
Search module for vb365_search
"""

from .base import BaseSearch
from .exchange import ExchangeItemsInMailboxesSearch
from .sharepoint import SharePointSearch
from .onedrive import OneDriveSearch
from .models import (
    SearchRequest,
    ExchangeItemsInMailboxes,
    ExchangeItemsInMailboxesResponse,
    SharePointSearchRequest,
    SharePointSearchResponse,
    OneDriveSearchRequest,
    OneDriveSearchResponse,
)

__all__ = [
    "BaseSearch",
    "ExchangeItemsInMailboxesSearch",
    "SharePointSearch",
    "OneDriveSearch",
    "SearchRequest",
    "ExchangeItemsInMailboxes",
    "ExchangeItemsInMailboxesResponse",
    "SharePointSearchRequest",
    "SharePointSearchResponse",
    "OneDriveSearchRequest",
    "OneDriveSearchResponse",
]
