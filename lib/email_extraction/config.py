"""
Email extraction utility library for processing .msg files and extracting attachments.
"""

import os
import json
from typing import Dict, List, Optional


def load_or_create_config(
    config_path: str,
    defaults: Optional[Dict[str, str]] = None
) -> Dict:
    """
    Load configuration from JSON file or create with defaults if it doesn't exist.
    
    Args:
        config_path: Path to the configuration JSON file
        defaults: Dictionary of default configuration values to add if missing
        
    Returns:
        Dictionary containing configuration values
    """
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
    else:
        config = {}

    if defaults:
        modified = False
        for key, value in defaults.items():
            if key not in config:
                config[key] = value
                modified = True

        if modified:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)

    return config
