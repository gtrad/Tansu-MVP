"""
Settings management for Tansu.
Stores user preferences in a JSON file.
"""

import json
import os
from typing import Any, Optional
from database import get_app_dir


SETTINGS_FILE = "tansu_settings.json"

# Default settings
DEFAULTS = {
    "check_for_updates": True,  # Whether to check for updates on startup
    "first_run_complete": False,  # Whether first-run dialog has been shown
    "anonymous_id": None,  # Random ID for anonymous analytics (generated on first run)
}


def _get_settings_path() -> str:
    """Get the full path to the settings file."""
    return os.path.join(get_app_dir(), SETTINGS_FILE)


def load_settings() -> dict:
    """Load settings from file, returning defaults for missing values."""
    settings = DEFAULTS.copy()

    try:
        path = _get_settings_path()
        if os.path.exists(path):
            with open(path, 'r') as f:
                saved = json.load(f)
                settings.update(saved)
    except Exception:
        pass

    return settings


def save_settings(settings: dict):
    """Save settings to file."""
    try:
        path = _get_settings_path()
        with open(path, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        print(f"Warning: Could not save settings: {e}")


def get_setting(key: str, default: Any = None) -> Any:
    """Get a single setting value."""
    settings = load_settings()
    return settings.get(key, default if default is not None else DEFAULTS.get(key))


def set_setting(key: str, value: Any):
    """Set a single setting value."""
    settings = load_settings()
    settings[key] = value
    save_settings(settings)


def is_first_run() -> bool:
    """Check if this is the first time the app is running."""
    return not get_setting("first_run_complete", False)


def mark_first_run_complete():
    """Mark that the first-run dialog has been shown."""
    set_setting("first_run_complete", True)


def get_anonymous_id() -> str:
    """Get or create an anonymous ID for this installation."""
    import uuid

    anon_id = get_setting("anonymous_id")
    if not anon_id:
        anon_id = str(uuid.uuid4())
        set_setting("anonymous_id", anon_id)

    return anon_id
