#!/usr/bin/env python3
"""
Variable Tracker Menu Bar App
Lives in the Mac menu bar for quick variable insertion into Word.
"""

import rumps
import subprocess
import logging
import uuid
from database import VariableDatabase

# Configure logging
logging.basicConfig(level=logging.INFO)

# Property name for storing GUID in Word document
GUID_PROPERTY_NAME = "VariableTrackerGUID"


def run_applescript(script: str) -> str:
    """Run an AppleScript and return the output."""
    result = subprocess.run(
        ['osascript', '-e', script],
        capture_output=True,
        text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr}")
    return result.stdout.strip()


def insert_variable_into_word(var_name: str, var_value: str, as_field: bool = True) -> bool:
    """Insert a variable into Word at the current cursor position.

    If as_field=True, inserts a DOCVARIABLE field that can be updated later.
    If as_field=False, inserts plain text.
    """
    try:
        # Escape for AppleScript
        name_escaped = var_name.replace('"', '\\"')
        value_escaped = var_value.replace('"', '\\"')

        if as_field:
            # Insert as DOCVARIABLE field + set the document variable
            script = f'''
tell application "Microsoft Word"
    activate
    set doc to active document

    -- Set the document variable (delete first if exists, then create and set value separately)
    try
        delete variable "{name_escaped}" of doc
    end try
    make new variable at doc with properties {{name:"{name_escaped}"}}
    set variable value of variable "{name_escaped}" of doc to "{value_escaped}"

    -- Insert the DOCVARIABLE field at cursor
    set theSelection to selection
    make new field at text object of theSelection with properties {{field type:field doc variable, field text:"{name_escaped}"}}
end tell

-- Press F9 to update/refresh the field display
tell application "System Events"
    delay 0.2
    key code 101
end tell

return "success"
'''
        else:
            # Insert as plain text
            script = f'''
tell application "Microsoft Word"
    type text selection text "{value_escaped}"
    return "success"
end tell
'''
        run_applescript(script)
        return True

    except Exception as e:
        logging.error(f"Error inserting variable: {e}")
        return False


def check_word_document_open() -> bool:
    """Check if Word has a document open."""
    script = '''
tell application "System Events"
    set isRunning to (exists process "Microsoft Word")
end tell
if isRunning then
    tell application "Microsoft Word"
        if (count of documents) > 0 then
            return "true"
        end if
    end tell
end if
return "false"
'''
    try:
        result = run_applescript(script)
        return result == "true"
    except:
        return False


def get_active_document_info() -> dict:
    """Get info about the active Word document, including persistent GUID."""
    # First, get or create GUID
    guid_script = f'''
tell application "Microsoft Word"
    set doc to active document
    try
        set existingGuid to value of custom document property "{GUID_PROPERTY_NAME}" of doc
        if existingGuid is not "" then
            return existingGuid
        end if
    end try
    return ""
end tell
'''
    try:
        existing_guid = run_applescript(guid_script)
    except:
        existing_guid = ""

    # If no GUID, create one
    if not existing_guid:
        new_guid = str(uuid.uuid4())
        set_guid_script = f'''
tell application "Microsoft Word"
    set doc to active document
    try
        make new custom document property at doc with properties {{name:"{GUID_PROPERTY_NAME}", value:"{new_guid}"}}
    end try
end tell
'''
        try:
            run_applescript(set_guid_script)
            existing_guid = new_guid
        except:
            pass

    # Get name and path (handle unsaved documents)
    info_script = '''
tell application "Microsoft Word"
    set doc to active document
    set docName to name of doc
    try
        set docPath to full name of doc
    on error
        set docPath to docName
    end try
    set isSaved to saved of doc
    return docName & "||" & docPath & "||" & isSaved
end tell
'''
    try:
        result = run_applescript(info_script)
        parts = result.split("||")
        name = parts[0] if len(parts) > 0 else "Unknown"
        path = parts[1] if len(parts) > 1 else name
        is_saved = parts[2] == "true" if len(parts) > 2 else False

        # For unsaved documents, use GUID as path identifier
        if not is_saved or path == name:
            path = f"unsaved:{existing_guid}"

        return {
            'guid': existing_guid,
            'name': name,
            'path': path,
            'is_saved': is_saved
        }
    except:
        return None


class VariableTrackerMenuBar(rumps.App):
    def __init__(self):
        super().__init__("VT", quit_button=None)
        self.db = VariableDatabase()
        self._last_var_count = 0
        self._last_var_hash = ""
        self.build_menu()

        # Set up a timer to check for database changes every 2 seconds
        self.timer = rumps.Timer(self._check_for_updates, 2)
        self.timer.start()

    def _check_for_updates(self, _):
        """Check if variables have changed and rebuild menu if needed."""
        try:
            variables = self.db.get_all_variables()
            # Create a simple hash of variable data
            var_hash = str([(v['id'], v['name'], v['value']) for v in variables])
            if var_hash != self._last_var_hash:
                self._last_var_hash = var_hash
                self.build_menu()
        except:
            pass

    def build_menu(self):
        """Build the menu with current variables."""
        self.menu.clear()

        # Get all variables from database
        variables = self.db.get_all_variables()

        if variables:
            # Add submenu for each variable with options
            for var in variables:
                display_text = f"{var['name']}: {var['value']}"
                if var.get('unit'):
                    display_text += f" {var['unit']}"

                # Create submenu with insert options
                submenu = rumps.MenuItem(display_text)
                submenu.add(rumps.MenuItem(
                    "Insert as Field (updatable)",
                    callback=lambda sender, v=var: self.insert_variable(v, as_field=True, with_unit=False)
                ))
                submenu.add(rumps.MenuItem(
                    "Insert as Text",
                    callback=lambda sender, v=var: self.insert_variable(v, as_field=False, with_unit=False)
                ))
                # Add "with unit" options if variable has a unit
                if var.get('unit'):
                    submenu.add(rumps.separator)
                    submenu.add(rumps.MenuItem(
                        "Insert as Field with Unit",
                        callback=lambda sender, v=var: self.insert_variable(v, as_field=True, with_unit=True)
                    ))
                    submenu.add(rumps.MenuItem(
                        "Insert as Text with Unit",
                        callback=lambda sender, v=var: self.insert_variable(v, as_field=False, with_unit=True)
                    ))
                self.menu.add(submenu)

            self.menu.add(rumps.separator)
        else:
            self.menu.add(rumps.MenuItem("No variables yet", callback=None))
            self.menu.add(rumps.separator)

        # Add refresh and quit options
        self.menu.add(rumps.MenuItem("Refresh Variables", callback=self.refresh_menu))
        self.menu.add(rumps.MenuItem("Open Variable Tracker", callback=self.open_main_app))
        self.menu.add(rumps.separator)
        self.menu.add(rumps.MenuItem("Quit", callback=self.quit_app))

    def insert_variable(self, var, as_field: bool = False, with_unit: bool = False):
        """Insert a variable into Word."""
        if not check_word_document_open():
            rumps.notification(
                title="Variable Tracker",
                subtitle="No Word Document",
                message="Please open a Word document first."
            )
            return

        # Determine the value to insert
        value_to_insert = var['value']
        if with_unit and var.get('unit'):
            value_to_insert = f"{var['value']} {var['unit']}"

        success = insert_variable_into_word(var['name'], value_to_insert, as_field=as_field)

        if success:
            # Record usage in database (only for field insertions)
            if as_field:
                try:
                    doc_info = get_active_document_info()
                    if doc_info and doc_info.get('guid'):
                        # Register/get document using GUID (survives rename/move)
                        doc_id = self.db.register_document(
                            guid=doc_info['guid'],
                            name=doc_info['name'],
                            path=doc_info['path'],
                            doc_type="word"
                        )
                        self.db.record_usage(var['id'], doc_id, with_unit=with_unit)
                except Exception as e:
                    logging.error(f"Error recording usage: {e}")

            insert_type = "field" if as_field else "text"
            unit_info = " with unit" if with_unit else ""
            rumps.notification(
                title="Variable Tracker",
                subtitle="Inserted",
                message=f"'{var['name']}' inserted as {insert_type}{unit_info}"
            )
        else:
            rumps.notification(
                title="Variable Tracker",
                subtitle="Error",
                message=f"Failed to insert '{var['name']}'"
            )

    def refresh_menu(self, _):
        """Refresh the menu with latest variables."""
        self.build_menu()
        rumps.notification(
            title="Variable Tracker",
            subtitle="Refreshed",
            message="Variable list updated"
        )

    def open_main_app(self, _):
        """Open the main Variable Tracker app."""
        import os
        app_dir = os.path.dirname(os.path.abspath(__file__))
        subprocess.Popen(['python', os.path.join(app_dir, 'run.py')])

    def quit_app(self, _):
        """Quit the menu bar app."""
        rumps.quit_application()


def main():
    VariableTrackerMenuBar().run()


if __name__ == "__main__":
    main()
