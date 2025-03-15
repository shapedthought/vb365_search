"""
Configuration handling for vb365_search
"""

import os
import tomllib
import toml
from pathlib import Path
from typing import Optional, Union

from .models.models import Configuration

def load_config(config_path: Optional[Union[str, Path]] = None) -> Configuration:
    """
    Load configuration from a TOML file
    
    Args:
        config_path: Path to the configuration file, defaults to 'configuration.toml'
        
    Returns:
        Configuration object
        
    Raises:
        FileNotFoundError: If the configuration file is not found
        tomllib.TOMLDecodeError: If the configuration file is not valid TOML
    """
    if config_path is None:
        config_path = "configuration.toml"
    
    config_path = Path(config_path)
    
    try:
        with open(config_path, mode="rb") as fp:
            config_dict = tomllib.load(fp)
        return Configuration(**config_dict)
    except (FileNotFoundError, tomllib.TOMLDecodeError) as e:
        raise ValueError(f"Configuration error: {e}")

def create_config_template(output_path: Optional[Union[str, Path]] = None) -> Path:
    """
    Create a configuration template file
    
    Args:
        output_path: Path to save the template, defaults to 'configuration.toml'
        
    Returns:
        Path to the created template file
    """
    from .models.models import MicrosoftConfig, Vb365Config
    
    if output_path is None:
        output_path = "configuration.toml"
    
    output_path = Path(output_path)
    
    config = Configuration(
        microsoft=MicrosoftConfig(),
        vb365=Vb365Config(),
    )
    
    with open(output_path, mode="w") as fp:
        toml.dump(config.model_dump(), fp)
    
    return output_path


def get_config_path() -> Path:
    """
    Get the path to the configuration file
    
    This function checks for the configuration file in the following locations:
    1. Current working directory
    2. User's home directory
    3. XDG_CONFIG_HOME directory (Linux/macOS)
    4. AppData directory (Windows)
    
    Returns:
        Path to the configuration file
        
    Raises:
        FileNotFoundError: If the configuration file is not found
    """
    # Check current working directory
    cwd_config = Path("configuration.toml")
    if cwd_config.exists():
        return cwd_config
    
    # Check user's home directory
    home_config = Path.home() / "configuration.toml"
    if home_config.exists():
        return home_config
    
    # Check XDG_CONFIG_HOME directory (Linux/macOS)
    xdg_config_home = os.environ.get("XDG_CONFIG_HOME")
    if xdg_config_home:
        xdg_config = Path(xdg_config_home) / "vb365_search" / "configuration.toml"
        if xdg_config.exists():
            return xdg_config
    
    # Check .config directory in user's home (Linux/macOS)
    dot_config = Path.home() / ".config" / "vb365_search" / "configuration.toml"
    if dot_config.exists():
        return dot_config
    
    # Check AppData directory (Windows)
    appdata = os.environ.get("APPDATA")
    if appdata:
        appdata_config = Path(appdata) / "vb365_search" / "configuration.toml"
        if appdata_config.exists():
            return appdata_config
    
    raise FileNotFoundError("Configuration file not found")
