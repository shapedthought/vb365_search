"""
Base search class for vb365_search
"""

from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional
from datetime import datetime, timezone
from pathlib import Path
import json
import logging

from vb365_search.authentication.auth_models import AuthHeaders
from vb365_search.models.models import Configuration

logger = logging.getLogger(__name__)

class BaseSearch(ABC):
    """
    Base class for all search types in Veeam Backup for Microsoft 365
    
    This abstract class defines the interface that all search implementations
    must follow. Specific search types (Exchange, SharePoint, OneDrive, etc.)
    should inherit from this class and implement the required methods.
    """
    
    def __init__(self, config: Configuration, auth_headers: AuthHeaders, restore_session_id: str):
        """
        Initialize the base search class
        
        Args:
            config: Configuration object
            auth_headers: Authentication object
            restore_session_id: ID of the restore session
        """
        self.config = config
        self.auth_headers = auth_headers
        self.restore_session_id = restore_session_id
        self.results = None
        self.last_search_term = None
    
    @abstractmethod
    def search(self, term: str, limit: int = 30, **kwargs) -> Any:
        """
        Execute a search with the given term
        
        Args:
            term: Search term or query
            limit: Maximum number of results to return
            **kwargs: Additional search parameters
            
        Returns:
            Search results object
        """
        pass
    
    @abstractmethod
    def get_results(self) -> List[Dict[str, Any]]:
        """
        Get the results of the search in a standardized format
        
        Returns:
            List of result dictionaries
        """
        pass
    
    def save_results(self, filename: Optional[str] = None) -> Path:
        """
        Save search results to a JSON file
        
        Args:
            filename: Optional filename, defaults to search-{timestamp}.json
            
        Returns:
            Path to the saved file
        """
        if not self.results:
            raise ValueError("No search results to save")
        
        if not filename:
            dt = datetime.now(timezone.utc)
            dt_str = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            filename = f"search-{dt_str}.json"
        
        output_path = Path(filename)
        with open(output_path, "w") as f:
            json.dump(self.results, f, indent=4)
        
        logger.info(f"Search results saved to {output_path}")
        return output_path
    
    def print_results(self) -> None:
        """
        Print search results to the console
        """
        if not self.results:
            print("No search results to display")
            return
        
        results = self.get_results()
        for result in results:
            for key, value in result.items():
                print(f"{key}: {value}")
            print("")
