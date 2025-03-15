"""
Command-line interface for vb365_search
"""

import logging
from typing import Optional

import fire

from .config import load_config, create_config_template, get_config_path
from .authentication.modern_auth import AuthenticateModern
from .authentication.auth_models import AuthHeaders
from .restore_session.restore_models import RestoreSessionRequest, RestoreSessionResponse
from .restore_session.restore_session import RestoreSession
from .search.exchange import ExchangeItemsInMailboxesSearch, ExchangeSearch
from .search.sharepoint import SharePointSearch
from .search.onedrive import OneDriveSearch
from .utils.helpers import save_json, load_json, auth_from_config

# Setup logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class CLI:
    """
    Command-line interface for vb365_search
    """
    
    def template(self, output_path: Optional[str] = None) -> None:
        """
        Create a configuration template file
        
        Args:
            output_path: Path to save the template, defaults to 'configuration.toml'
        """
        path = create_config_template(output_path)
        logger.info(f"Configuration template created at {path}")
    
    def login(self, config_path: Optional[str] = None) -> None:
        """
        Login to Veeam Backup for Microsoft Office 365
        
        Args:
            config_path: Path to the configuration file, defaults to auto-detection
        """
        if config_path is None:
            try:
                config_path = get_config_path()
            except FileNotFoundError:
                logger.error("Configuration file not found. Use 'template' to create one.")
                return
        
        config = load_config(config_path)
        
        logger.info("Logging in...")
        
        auth_config = auth_from_config(config)
        
        auth_modern = AuthenticateModern(auth_config)
        
        try:
            veeam_token_model = auth_modern.authenticate_veeam_backup_o365()
        except Exception as e:
            logger.error(f"Login failed: {e}")
            raise
        
        auth_headers = AuthHeaders(
            authorization=f"{veeam_token_model.token_type} {veeam_token_model.access_token}"
        )
        
        # Save authentication response and headers
        save_json(veeam_token_model.model_dump(), "auth_response.json")
        save_json(auth_headers.model_dump(), "auth_headers.json")
        
        # Create restore session
        restore_session = RestoreSession(auth_config, auth_headers)
        restore_request = RestoreSessionRequest()
        
        try:
            restore_response = restore_session.create_restore_session(restore_request)
            save_json(restore_response.model_dump(), "restore.json")
            logger.info(f"Restore session created with ID: {restore_response.id}")
        except Exception as e:
            logger.error(f"Failed to create restore session: {e}")
            raise
    
    def search(
        self,
        term: str,
        search_type: str = "exchange",
        limit: int = 30,
        print_results: bool = False,
        config_path: Optional[str] = None,
    ) -> None:
        """
        Search Veeam Backup for Microsoft Office 365
        
        Args:
            term: Search term or query
            search_type: Type of search (exchange, sharepoint, onedrive)
            limit: Maximum number of results to return
            print_results: Whether to print results to the console
            config_path: Path to the configuration file, defaults to auto-detection
        """
        # Load auth headers and restore session
        try:
            auth_headers_dict = load_json("auth_headers.json")
            restore_model = RestoreSessionResponse(**load_json("restore.json"))
        except FileNotFoundError:
            logger.error("Authentication files not found. Please login first.")
            return
        
        # Load configuration
        if config_path is None:
            try:
                config_path = get_config_path()
            except FileNotFoundError:
                logger.error("Configuration file not found. Use 'template' to create one.")
                return
        
        config = load_config(config_path)
        
        # Create search object based on search type
        if search_type.lower() == "exchange":
            search_obj = ExchangeItemsInMailboxesSearch(
                config.model_dump(),
                auth_headers_dict,
                restore_model.id,
            )
        elif search_type.lower() == "sharepoint":
            search_obj = SharePointSearch(
                config.model_dump(),
                auth_headers_dict,
                restore_model.id,
            )
        elif search_type.lower() == "onedrive":
            search_obj = OneDriveSearch(
                config.model_dump(),
                auth_headers_dict,
                restore_model.id,
            )
        else:
            logger.error(f"Unknown search type: {search_type}")
            return
        
        # Execute search
        try:
            search_response = search_obj.search(term, limit=limit)
            
            if print_results:
                search_obj.print_results()
            
            # Save results
            output_path = search_obj.save_results()
            logger.info(f"Search results saved to {output_path}")
            
            return search_response
        except Exception as e:
            logger.error(f"Search failed: {e}")
            raise
    
    def logout(self, config_path: Optional[str] = None) -> None:
        """
        Logout from Veeam Backup for Microsoft Office 365
        
        Args:
            config_path: Path to the configuration file, defaults to auto-detection
        """
        try:
            auth_headers_dict = load_json("auth_headers.json")
            restore_model = RestoreSessionResponse(**load_json("restore.json"))
        except FileNotFoundError:
            logger.error("Authentication files not found. Please login first.")
            return
        
        # Load configuration
        if config_path is None:
            try:
                config_path = get_config_path()
            except FileNotFoundError:
                logger.error("Configuration file not found. Use 'template' to create one.")
                return
        
        config = load_config(config_path)
        
        # Stop restore session
        logout_url = (
            f"https://{config.vb365.api_address}:4443/{config.vb365.version}/"
            f"RestoreSessions/{restore_model.id}/Stop"
        )
        
        import requests
        
        try:
            session = requests.Session()
            logout_res = session.post(
                logout_url, headers=auth_headers_dict, verify=False
            )
            logout_res.raise_for_status()
            logger.info("Logout successful!")
        except requests.exceptions.RequestException as e:
            logger.error(f"Logout failed: {e}")
        finally:
            session.close()


def main():
    """
    Main entry point for the CLI
    """
    fire.Fire(CLI)


if __name__ == "__main__":
    main()
