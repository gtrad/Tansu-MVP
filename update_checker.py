"""
Update checker for Tansu.
Checks GitHub Releases for new versions.
"""

import urllib.request
import json
import logging
import threading
from typing import Optional, Callable
import platform

from version import __version__, GITHUB_REPO

logger = logging.getLogger(__name__)

GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def parse_version(version_str: str) -> tuple:
    """Parse version string like '1.2.3' into tuple (1, 2, 3) for comparison."""
    # Remove 'v' prefix if present
    version_str = version_str.lstrip('v')
    try:
        return tuple(int(x) for x in version_str.split('.'))
    except ValueError:
        return (0, 0, 0)


def check_for_update() -> Optional[dict]:
    """
    Check GitHub for a newer version.

    Returns:
        dict with 'version', 'url', 'notes' if update available
        None if no update or error
    """
    try:
        # Build request with User-Agent (required by GitHub API)
        req = urllib.request.Request(
            GITHUB_API_URL,
            headers={
                'User-Agent': f'Tansu/{__version__} ({platform.system()})',
                'Accept': 'application/vnd.github.v3+json'
            }
        )

        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode('utf-8'))

        latest_version = data.get('tag_name', '').lstrip('v')
        current_version = __version__

        # Compare versions
        if parse_version(latest_version) > parse_version(current_version):
            return {
                'version': latest_version,
                'url': data.get('html_url', ''),
                'download_url': _get_download_url(data),
                'notes': data.get('body', '')[:500]  # Truncate release notes
            }

        return None

    except Exception as e:
        logger.debug(f"Update check failed: {e}")
        return None


def _get_download_url(release_data: dict) -> str:
    """Extract the appropriate download URL for the current platform."""
    assets = release_data.get('assets', [])
    system = platform.system().lower()

    for asset in assets:
        name = asset.get('name', '').lower()

        if system == 'darwin' and ('.app' in name or '.dmg' in name or 'mac' in name):
            return asset.get('browser_download_url', '')
        elif system == 'windows' and ('.exe' in name or 'win' in name):
            return asset.get('browser_download_url', '')

    # Fallback to release page
    return release_data.get('html_url', '')


def check_for_update_async(callback: Callable[[Optional[dict]], None]):
    """
    Check for updates in a background thread.

    Args:
        callback: Function to call with the result (None or update info dict)
    """
    def _check():
        result = check_for_update()
        callback(result)

    thread = threading.Thread(target=_check, daemon=True)
    thread.start()


# For testing
if __name__ == "__main__":
    print(f"Current version: {__version__}")
    print(f"Checking for updates from: {GITHUB_API_URL}")

    result = check_for_update()
    if result:
        print(f"Update available: {result['version']}")
        print(f"Download: {result['url']}")
    else:
        print("No update available (or couldn't check)")
