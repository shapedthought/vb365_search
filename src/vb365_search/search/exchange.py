"""
Exchange search implementation for vb365_search
"""

import requests
import logging
from typing import Any, Dict, List

from vb365_search.authentication.auth_models import AuthHeaders
from vb365_search.models.models import Configuration

from .base import BaseSearch
from .models import ExchangeItemsInMailboxes, ExchangeItemsInMailboxesResponse

logger = logging.getLogger(__name__)

class ExchangeItemsInMailboxesSearch(BaseSearch):
    """
    Exchange-specific search implementation
    
    This class implements the search functionality for Exchange mailboxes
    in Veeam Backup for Microsoft 365.
    """
    
    def __init__(self, config: Configuration, auth_headers: AuthHeaders, restore_session_id: str):
        """
        Initialize the Exchange search
        
        Args:
            config: Configuration dictionary with API endpoints and other settings
            auth_headers: Authentication headers for API requests
            restore_session_id: ID of the restore session
        """
        super().__init__(config, auth_headers, restore_session_id)
        self.api_address = config.vb365.api_address
        self.api_version = config.vb365.version
        
        if not self.api_address:
            raise ValueError("API address not found in configuration")
    
    def search(self, term: str, limit: int = 30, **kwargs: Any) -> ExchangeItemsInMailboxesResponse:
        """
        Search Exchange mailboxes
        
        Args:
            term: Search term or query
            limit: Maximum number of results to return
            **kwargs: Additional search parameters
            
        Returns:
            ExchangeSearchResponse object with search results
        """
        self.last_search_term = term
        
        search_url = (
            f"https://{self.api_address}:4443/{self.api_version}/RestoreSessions/"
            f"{self.restore_session_id}/organization/mailboxes/search?limit={limit}"
        )
        
        search_body = ExchangeItemsInMailboxes(term=term)
        
        logger.info(f"Searching Exchange mailboxes for '{term}'...")
        
        try:
            response = requests.post(
                search_url,
                json=search_body.model_dump(),
                headers=self.auth_headers.model_dump(),
                verify=False,
            )
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            logger.error(f"Exchange search failed: {e}")
            raise
        
        self.results = response.json()
        search_response = ExchangeItemsInMailboxesResponse(**self.results)
        
        logger.info(f"Search complete! {len(search_response.results)} items found.")
        
        return search_response
        
    def get_results(self) -> List[Dict[str, Any]]:
        """
        Get the results of the search in a standardized format
        
        Returns:
            List of result dictionaries with standardized keys
        """
        if not self.results:
            return []
        
        search_response = ExchangeItemsInMailboxesResponse(**self.results)
        
        standardized_results: List[Dict[str, Any]] = []
        
        # Convert to a standardized format
        for result in search_response.results:
            standardized_results.append({
                "subject": result.subject,
                "received": result.received,
                "from": result.from_,
                "to": result.to,
                "cc": result.cc,
                "bcc": result.bcc,
                "importance": result.importance,
                "has_attachments": len(result.attachments) > 0,
                "attachments": [a.name for a in result.attachments],
                "id": result.id,
            })
        
        return standardized_results
