"""
Word integration module using AppleScript for macOS.
Handles inserting variables, scanning documents, and updating values.

Note: This module only works on macOS with Microsoft Word installed.
"""

import logging
import subprocess
import re
import uuid
from typing import Optional
from dataclasses import dataclass

# Check if we're on macOS and osascript is available
HAS_APPLESCRIPT = False
try:
    result = subprocess.run(
        ['which', 'osascript'],
        capture_output=True,
        text=True
    )
    if result.returncode == 0:
        HAS_APPLESCRIPT = True
except Exception:
    pass

if not HAS_APPLESCRIPT:
    logging.warning("AppleScript not available - Word integration disabled")


# Custom property name for our tracking GUID
GUID_PROPERTY_NAME = "VariableTrackerGUID"


@dataclass
class DocumentInfo:
    """Information about a Word document."""
    guid: str
    name: str
    path: str
    variables: list[str]  # List of variable names found


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


class WordIntegration:
    """Handles all Word AppleScript automation on macOS."""

    def __init__(self):
        if not HAS_APPLESCRIPT:
            raise RuntimeError("AppleScript is required for Word integration on macOS")

    def get_active_document(self):
        """Check if Word is running and has a document open. Returns True/False."""
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
        except Exception as e:
            logging.error(f"Error checking for active document: {e}")
            return False

    # -------------------------
    # GUID Management
    # -------------------------

    def get_document_guid(self, doc=None) -> Optional[str]:
        """Get the tracking GUID from active document's custom properties."""
        if not self.get_active_document():
            return None

        script = f'''
tell application "Microsoft Word"
    try
        set propValue to value of custom document property "{GUID_PROPERTY_NAME}" of active document
        return propValue
    on error
        return ""
    end try
end tell
'''
        try:
            result = run_applescript(script)
            return result if result else None
        except Exception as e:
            logging.error(f"Error reading document GUID: {e}")
            return None

    def set_document_guid(self, guid: str, doc=None) -> bool:
        """Set tracking GUID as custom document property."""
        if not self.get_active_document():
            return False

        # Escape any quotes in the guid
        guid_escaped = guid.replace('"', '\\"')

        script = f'''
tell application "Microsoft Word"
    set doc to active document
    try
        set value of custom document property "{GUID_PROPERTY_NAME}" of doc to "{guid_escaped}"
    on error
        make new custom document property at doc with properties {{name:"{GUID_PROPERTY_NAME}", value:"{guid_escaped}"}}
    end try
    return "true"
end tell
'''
        try:
            run_applescript(script)
            return True
        except Exception as e:
            logging.error(f"Error setting document GUID: {e}")
            return False

    # -------------------------
    # Variable Operations
    # -------------------------

    def insert_variable(self, var_name: str, var_value: str, doc=None) -> bool:
        """
        Insert a DOCVARIABLE field at the current cursor position.
        Also sets the document variable value.
        """
        if not self.get_active_document():
            return False

        try:
            # First set the document variable
            self._set_doc_variable(var_name, var_value)

            # Escape variable name for AppleScript
            var_name_escaped = var_name.replace('"', '\\"')

            # Insert DOCVARIABLE field at cursor
            # Word must be activated first for this to work reliably
            script = f'''
tell application "Microsoft Word"
    activate
    set theSelection to selection
    make new field at text object of theSelection with properties {{field type:field doc variable, field text:"{var_name_escaped}"}}
end tell

-- Press F9 to update/refresh the field display
tell application "System Events"
    delay 0.2
    key code 101
end tell

return "true"
'''
            run_applescript(script)
            return True
        except Exception as e:
            logging.error(f"Error inserting variable: {e}")
            return False

    def _set_doc_variable(self, name: str, value: str):
        """Set a document variable value."""
        # Escape for AppleScript
        name_escaped = name.replace('"', '\\"')
        value_escaped = value.replace('"', '\\"')

        script = f'''
tell application "Microsoft Word"
    set doc to active document
    try
        delete variable "{name_escaped}" of doc
    end try
    make new variable at doc with properties {{name:"{name_escaped}"}}
    set variable value of variable "{name_escaped}" of doc to "{value_escaped}"
end tell
'''
        try:
            run_applescript(script)
        except Exception as e:
            logging.error(f"Error setting document variable: {e}")
            raise

    def get_doc_variable_value(self, var_name: str, doc=None) -> Optional[str]:
        """Get the current value of a document variable."""
        if not self.get_active_document():
            return None

        var_name_escaped = var_name.replace('"', '\\"')

        script = f'''
tell application "Microsoft Word"
    try
        set varValue to variable value of variable "{var_name_escaped}" of active document
        return varValue
    on error
        return ""
    end try
end tell
'''
        try:
            result = run_applescript(script)
            return result if result else None
        except Exception as e:
            logging.error(f"Error getting document variable: {e}")
            return None

    def get_document_variables(self, doc=None) -> dict[str, str]:
        """Get all document variables as a dict of name -> value."""
        if not self.get_active_document():
            return {}

        script = '''
tell application "Microsoft Word"
    set doc to active document
    set varList to {}
    set allVars to variables of doc
    repeat with v in allVars
        set varName to name of v
        set varValue to variable value of v
        set end of varList to varName & "|||" & varValue
    end repeat
    set AppleScript's text item delimiters to "~~~"
    return varList as text
end tell
'''
        try:
            result = run_applescript(script)
            if not result:
                return {}

            variables = {}
            for item in result.split("~~~"):
                if "|||" in item:
                    parts = item.split("|||", 1)
                    if len(parts) == 2:
                        variables[parts[0]] = parts[1]
            return variables
        except Exception as e:
            logging.error(f"Error getting document variables: {e}")
            return {}

    # -------------------------
    # Scanning
    # -------------------------

    def scan_document(self, doc=None) -> DocumentInfo:
        """
        Scan a document to find all DOCVARIABLE fields.
        Returns document info including list of variable names used.
        """
        if not self.get_active_document():
            raise ValueError("No active document")

        # Get or create GUID
        guid = self.get_document_guid()
        if guid is None:
            guid = str(uuid.uuid4())
            self.set_document_guid(guid)

        # Get document name and path
        script = '''
tell application "Microsoft Word"
    set docName to name of active document
    set docPath to full name of active document
    return docName & "||" & docPath
end tell
'''
        try:
            result = run_applescript(script)
            parts = result.split("||")
            name = parts[0] if len(parts) > 0 else "Unknown"
            path = parts[1] if len(parts) > 1 else name
        except Exception as e:
            logging.error(f"Error getting document info: {e}")
            name = "Unknown"
            path = "Unknown"

        # Find all DOCVARIABLE fields
        variable_names = []

        script = '''
tell application "Microsoft Word"
    set doc to active document
    set varNames to {}
    set allFields to fields of doc
    repeat with f in allFields
        if field type of f is field doc variable then
            set fieldCode to field code of f
            set codeText to content of fieldCode
            set end of varNames to codeText
        end if
    end repeat
    return varNames as text
end tell
'''
        try:
            result = run_applescript(script)
            if result:
                # Parse field codes to extract variable names
                # Field codes look like: " DOCVARIABLE widget_a_cost \* MERGEFORMAT "
                # Use findall to get all DOCVARIABLE names
                matches = re.findall(r'DOCVARIABLE\s+(\S+)', result)
                for var_name in matches:
                    var_name = var_name.strip('"')
                    if var_name and var_name not in variable_names:
                        variable_names.append(var_name)
        except Exception as e:
            logging.error(f"Error scanning fields: {e}")

        return DocumentInfo(
            guid=guid,
            name=name,
            path=path,
            variables=variable_names
        )

    # -------------------------
    # Updating
    # -------------------------

    def update_variables(self, variables: dict[str, str], doc=None) -> list[str]:
        """
        Update document variables with new values from database.

        Args:
            variables: Dict mapping variable name to new value
            doc: Document to update (uses active if None)

        Returns:
            List of variable names that were updated
        """
        if not self.get_active_document():
            return []

        updated = []

        try:
            # Update document variables
            for name, value in variables.items():
                current = self.get_doc_variable_value(name)
                if current is not None and current != value:
                    self._set_doc_variable(name, value)
                    updated.append(name)
                elif current is None:
                    # Variable doesn't exist in doc yet, add it
                    self._set_doc_variable(name, value)

            # Refresh all fields to show new values
            if updated:
                script = '''
tell application "Microsoft Word"
    activate
    set doc to active document

    -- Update each field individually
    set fieldCount to count of fields of doc
    repeat with i from 1 to fieldCount
        set f to field i of doc
        try
            update field f
        end try
    end repeat
end tell
'''
                run_applescript(script)

            return updated
        except Exception as e:
            logging.error(f"Error updating variables: {e}")
            return updated

    def get_stale_variables(self, db_variables: dict[str, str], doc=None) -> dict[str, tuple[str, str]]:
        """
        Find variables in document that don't match database values.

        Args:
            db_variables: Dict mapping variable name to database value
            doc: Document to check

        Returns:
            Dict mapping variable name to (doc_value, db_value) for stale variables
        """
        if not self.get_active_document():
            return {}

        stale = {}

        try:
            for name, db_value in db_variables.items():
                doc_value = self.get_doc_variable_value(name)
                if doc_value is not None and doc_value != db_value:
                    stale[name] = (doc_value, db_value)
        except Exception as e:
            logging.error(f"Error checking stale variables: {e}")

        return stale
