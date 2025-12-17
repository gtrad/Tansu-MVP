#!/usr/bin/env python3
"""
Variable Tracker Tray/Menu Bar App Launcher
Detects the OS and runs the appropriate tray application.

- Windows: System tray icon using pystray
- Mac: Menu bar app using rumps
"""

import platform
import sys


def main():
    system = platform.system()

    if system == "Windows":
        try:
            from tray_app_windows import main as windows_main
            windows_main()
        except ImportError as e:
            print(f"Error: Missing dependencies for Windows tray app: {e}")
            print("Install them with: pip install pystray Pillow pywin32")
            sys.exit(1)

    elif system == "Darwin":
        try:
            from menubar_app import main as mac_main
            mac_main()
        except ImportError as e:
            print(f"Error: Missing dependencies for Mac menu bar app: {e}")
            print("Install rumps with: pip install rumps")
            sys.exit(1)

    else:
        print(f"Tray app not supported on {system}")
        print("You can still use the main app: python app.py")
        sys.exit(1)


if __name__ == "__main__":
    main()
