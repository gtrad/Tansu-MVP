"""
Tansu Launcher - Entry point for the bundled application.
Launches both the main GUI and the menubar app.
"""

import sys
import os
import subprocess
import multiprocessing


def get_app_dir():
    """Get the directory containing app files."""
    if getattr(sys, 'frozen', False):
        # Running as bundled app
        return sys._MEIPASS
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))


def run_menubar():
    """Run the menubar app in a separate process."""
    app_dir = get_app_dir()

    # Import and run menubar directly (in subprocess)
    sys.path.insert(0, app_dir)
    os.chdir(app_dir)

    # Import menubar module
    import menubar_app
    menubar_app.main()


def run_main_app():
    """Run the main tkinter app."""
    app_dir = get_app_dir()
    sys.path.insert(0, app_dir)
    os.chdir(app_dir)

    import app
    app.main_gui_only()


def main():
    """Main entry point - launches both apps."""
    # Start menubar in a separate process
    menubar_process = multiprocessing.Process(target=run_menubar)
    menubar_process.daemon = True
    menubar_process.start()

    # Run main app in this process
    run_main_app()

    # Clean up menubar when main app closes
    if menubar_process.is_alive():
        menubar_process.terminate()


if __name__ == "__main__":
    multiprocessing.freeze_support()  # Required for PyInstaller
    main()
