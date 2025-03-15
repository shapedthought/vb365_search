"""
Search module for vb365_search
"""

from .base import BaseSearch
from .exchange import ExchangeSearch
from .sharepoint import SharePointSearch
from .onedrive import OneDriveSearch
from .models import (
    SearchRequest,
    ExchangeSearchRequest,
    ExchangeSearchResponse,
    SharePointSearchRequest,
    SharePointSearchResponse,
    OneDriveSearchRequest,
    OneDriveSearchResponse,
)

__all__ = [
    "BaseSearch",
    "ExchangeSearch",
    "SharePointSearch",
    "OneDriveSearch",
    "SearchRequest",
    "ExchangeSearchRequest",
    "ExchangeSearchResponse",
    "SharePointSearchRequest",
    "SharePointSearchResponse",
    "OneDriveSearchRequest",
    "OneDriveSearchResponse",
]
