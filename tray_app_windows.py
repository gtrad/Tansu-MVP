"""
Variable Tracker System Tray App for Windows
Lives in the system tray for quick variable insertion into Word.
Uses pystray for cross-platform system tray support.
Also starts the API server for Word VBA macro integration.
"""

import logging
import threading
import uuid
from typing import Optional

# Configure logging
logging.basicConfig(level=logging.INFO)

# Tray icon imports
try:
    import pystray
    from PIL import Image, ImageDraw, ImageFont
    HAS_PYSTRAY = True
except ImportError:
    HAS_PYSTRAY = False
    logging.warning("pystray/Pillow not available - tray app disabled")

# Word integration
try:
    from word_windows import WordIntegration, HAS_WIN32
except ImportError:
    HAS_WIN32 = False

from database import VariableDatabase
from api_server import start_api_server, stop_api_server


def create_tray_icon_image(size=64):
    """Create a simple 'VT' icon for the system tray."""
    # Create image with dark background
    image = Image.new('RGB', (size, size), color=(45, 45, 45))
    draw = ImageDraw.Draw(image)

    # Try to use a font, fall back to default
    try:
        font = ImageFont.truetype("arial.ttf", int(size * 0.5))
    except:
        font = ImageFont.load_default()

    # Draw "VT" text in green (matching Excel green from the app)
    text = "VT"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) // 2
    y = (size - text_height) // 2 - bbox[1]
    draw.text((x, y), text, fill=(33, 115, 70), font=font)

    return image


class WindowsTrayApp:
    """System tray application for Windows."""

    def __init__(self):
        if not HAS_PYSTRAY:
            raise RuntimeError("pystray and Pillow are required for the tray app")

        self.db = VariableDatabase()
        self.word = WordIntegration() if HAS_WIN32 else None
        self.icon = None
        self._running = True

    def _check_word_available(self) -> bool:
        """Check if Word integration is available and a document is open."""
        if not self.word:
            return False
        try:
            doc = self.word.get_active_document()
            return doc is not None
        except:
            return False

    def _show_notification(self, title: str, message: str):
        """Show a system notification."""
        if self.icon:
            try:
                self.icon.notify(message, title)
            except:
                pass  # Notifications may not be supported on all systems

    def _insert_variable(self, var: dict, as_field: bool = True, with_unit: bool = False):
        """Insert a variable into Word."""
        if not self._check_word_available():
            self._show_notification(
                "Variable Tracker",
                "Please open a Word document first."
            )
            return

        # Determine the value to insert
        value_to_insert = var['value']
        if with_unit and var.get('unit'):
            value_to_insert = f"{var['value']} {var['unit']}"

        try:
            if as_field:
                success = self.word.insert_variable(var['name'], value_to_insert)
            else:
                # Insert as plain text
                doc = self.word.get_active_document()
                if doc:
                    word_app = self.word._get_word_app()
                    word_app.Selection.TypeText(value_to_insert)
                    success = True
                else:
                    success = False

            if success:
                # Record usage in database (only for field insertions)
                if as_field:
                    try:
                        doc = self.word.get_active_document()
                        if doc:
                            guid = self.word.get_document_guid(doc)
                            if not guid:
                                guid = str(uuid.uuid4())
                                self.word.set_document_guid(guid, doc)

                            name = doc.Name
                            try:
                                path = doc.FullName
                            except:
                                path = name

                            doc_id = self.db.register_document(
                                guid=guid,
                                name=name,
                                path=path,
                                doc_type="word"
                            )
                            self.db.record_usage(var['id'], doc_id, with_unit=with_unit)
                    except Exception as e:
                        logging.error(f"Error recording usage: {e}")

                insert_type = "field" if as_field else "text"
                unit_info = " with unit" if with_unit else ""
                self._show_notification(
                    "Variable Tracker",
                    f"'{var['name']}' inserted as {insert_type}{unit_info}"
                )
            else:
                self._show_notification(
                    "Variable Tracker",
                    f"Failed to insert '{var['name']}'"
                )
        except Exception as e:
            logging.error(f"Error inserting variable: {e}")
            self._show_notification(
                "Variable Tracker",
                f"Error: {str(e)}"
            )

    def _create_insert_callback(self, var: dict, as_field: bool, with_unit: bool):
        """Create a callback function for menu item clicks."""
        def callback(icon, item):
            # Run in a separate thread to avoid blocking the menu
            thread = threading.Thread(
                target=self._insert_variable,
                args=(var, as_field, with_unit)
            )
            thread.start()
        return callback

    def _build_menu(self):
        """Build the system tray menu."""
        menu_items = []

        # Get all variables from database
        variables = self.db.get_all_variables()

        if variables:
            for var in variables:
                display_text = f"{var['name']}: {var['value']}"
                if var.get('unit'):
                    display_text += f" {var['unit']}"

                # Build submenu items
                submenu_items = [
                    pystray.MenuItem(
                        "Insert as Field (updatable)",
                        self._create_insert_callback(var, as_field=True, with_unit=False)
                    ),
                    pystray.MenuItem(
                        "Insert as Text",
                        self._create_insert_callback(var, as_field=False, with_unit=False)
                    ),
                ]

                # Add "with unit" options if variable has a unit
                if var.get('unit'):
                    submenu_items.extend([
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem(
                            "Insert as Field with Unit",
                            self._create_insert_callback(var, as_field=True, with_unit=True)
                        ),
                        pystray.MenuItem(
                            "Insert as Text with Unit",
                            self._create_insert_callback(var, as_field=False, with_unit=True)
                        ),
                    ])

                menu_items.append(
                    pystray.MenuItem(display_text, pystray.Menu(*submenu_items))
                )

            menu_items.append(pystray.Menu.SEPARATOR)
        else:
            menu_items.append(
                pystray.MenuItem("No variables yet", None, enabled=False)
            )
            menu_items.append(pystray.Menu.SEPARATOR)

        # Add utility options
        menu_items.extend([
            pystray.MenuItem("Refresh Variables", self._refresh_menu),
            pystray.MenuItem("Open Variable Tracker", self._open_main_app),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._quit_app),
        ])

        return pystray.Menu(*menu_items)

    def _refresh_menu(self, icon, item):
        """Refresh the menu with latest variables."""
        icon.menu = self._build_menu()
        self._show_notification("Variable Tracker", "Variable list refreshed")

    def _open_main_app(self, icon, item):
        """Open the main Variable Tracker app."""
        import subprocess
        import os
        import sys

        app_dir = os.path.dirname(os.path.abspath(__file__))
        python_exe = sys.executable

        # Start the main app
        subprocess.Popen([python_exe, os.path.join(app_dir, 'app.py')])

    def _quit_app(self, icon, item):
        """Quit the tray app."""
        self._running = False
        icon.stop()

    def run(self):
        """Run the system tray application."""
        image = create_tray_icon_image()
        menu = self._build_menu()

        self.icon = pystray.Icon(
            "Variable Tracker",
            image,
            "Variable Tracker - Click to insert variables",
            menu
        )

        # Start periodic refresh in background
        def refresh_loop():
            import time
            while self._running:
                time.sleep(5)  # Refresh every 5 seconds
                if self._running and self.icon:
                    try:
                        self.icon.menu = self._build_menu()
                    except:
                        pass

        refresh_thread = threading.Thread(target=refresh_loop, daemon=True)
        refresh_thread.start()

        self.icon.run()


def main():
    if not HAS_PYSTRAY:
        print("Error: pystray and Pillow are required.")
        print("Install them with: pip install pystray Pillow")
        return

    if not HAS_WIN32:
        print("Warning: pywin32 not available - Word integration disabled")
        print("Install it with: pip install pywin32")

    # Start the API server for Word VBA macro integration
    print("Starting Tansu API server on http://127.0.0.1:5050")
    start_api_server()

    try:
        app = WindowsTrayApp()
        app.run()
    finally:
        stop_api_server()


if __name__ == "__main__":
    main()
