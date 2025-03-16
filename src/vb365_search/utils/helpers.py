"""
Helper functions for vb365_search
"""

import json
import logging
from pathlib import Path
from typing import Any, Dict, Union

from pydantic import HttpUrl

from vb365_search.authentication.auth_models import AuthConfig, AuthHeaders, VeeamTokenResponse
from vb365_search.models.models import Configuration

logger = logging.getLogger(__name__)

def save_json(data: Dict[str, Any], filename: Union[str, Path]) -> Path:
    """
    Save data to a JSON file
    
    Args:
        data: Data to save
        filename: Filename to save to
        
    Returns:
        Path to the saved file
    """
    output_path = Path(filename)
    with open(output_path, "w") as f:
        json.dump(data, f, indent=4)
    
    logger.info(f"Data saved to {output_path}")
    return output_path


def load_json(filename: Union[str, Path]) -> Dict[str, Any]:
    """
    Load data from a JSON file
    
    Args:
        filename: Filename to load from
        
    Returns:
        Loaded data
        
    Raises:
        FileNotFoundError: If the file is not found
        json.JSONDecodeError: If the file is not valid JSON
    """
    input_path = Path(filename)
    with open(input_path, "r") as f:
        data = json.load(f)
    
    return data


def ensure_directory_exists(path: Union[str, Path]) -> Path:
    """
    Ensure that a directory exists, creating it if necessary
    
    Args:
        path: Path to the directory
        
    Returns:
        Path to the directory
    """
    directory = Path(path)
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def format_size(size_bytes: int) -> str:
    """
    Format a size in bytes to a human-readable string
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        Human-readable size string
    """
    if size_bytes < 1024:
        return f"{size_bytes} B"
    
    size_kb = size_bytes / 1024
    if size_kb < 1024:
        return f"{size_kb:.2f} KB"
    
    size_mb = size_kb / 1024
    if size_mb < 1024:
        return f"{size_mb:.2f} MB"
    
    size_gb = size_mb / 1024
    return f"{size_gb:.2f} GB"

def auth_from_config(config: Configuration) -> AuthConfig:
    """
    Create an AuthConfig object from a Configuration object
    """
    return AuthConfig(
        client_id=config.microsoft.application_id,
        tenant_name=config.microsoft.tenant_name,
        veeam_api_url=HttpUrl(f"https://{config.vb365.api_address}:4443/{config.vb365.version}")
    )
    
def headers_from_veeam_token_response(veeam_token_model: VeeamTokenResponse) -> AuthHeaders:
    """
    Create an AuthHeaders object from a VeeamTokenResponse object
    """
    auth_headers = AuthHeaders(
            Authorization=f"{veeam_token_model.token_type} {veeam_token_model.access_token}"
        )
    return auth_headers