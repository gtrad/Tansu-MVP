"""
Word integration module using COM automation.
Handles inserting variables, scanning documents, and updating values.

Note: This module requires pywin32 and only works on Windows.
"""

import logging
from typing import Optional
from dataclasses import dataclass

# COM imports - will only work on Windows
try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    logging.warning("pywin32 not available - Word integration disabled")


# Custom property name for our tracking GUID
GUID_PROPERTY_NAME = "VariableTrackerGUID"


@dataclass
class DocumentInfo:
    """Information about a Word document."""
    guid: str
    name: str
    path: str
    variables: list[str]  # List of variable names found


class WordIntegration:
    """Handles all Word COM automation."""
    
    def __init__(self):
        if not HAS_WIN32:
            raise RuntimeError("pywin32 is required for Word integration")
        self._word = None

    def _get_word_app(self):
        """Get or create Word application instance."""
        if self._word is None:
            try:
                # Try to connect to existing Word instance
                self._word = win32com.client.GetActiveObject("Word.Application")
            except:
                # No Word running, create new instance
                self._word = win32com.client.Dispatch("Word.Application")
                self._word.Visible = True
        return self._word

    def get_active_document(self):
        """Get the currently active Word document."""
        word = self._get_word_app()
        if word.Documents.Count == 0:
            return None
        return word.ActiveDocument

    # -------------------------
    # GUID Management
    # -------------------------

    def get_document_guid(self, doc=None) -> Optional[str]:
        """Get the tracking GUID from a document, or None if not set."""
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            return None
        
        # Check custom document properties
        try:
            props = doc.CustomDocumentProperties
            for i in range(1, props.Count + 1):
                prop = props.Item(i)
                if prop.Name == GUID_PROPERTY_NAME:
                    return prop.Value
        except Exception as e:
            logging.error(f"Error reading document properties: {e}")
        
        return None

    def set_document_guid(self, guid: str, doc=None) -> bool:
        """Set the tracking GUID on a document."""
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            return False
        
        try:
            props = doc.CustomDocumentProperties
            
            # Check if already exists
            existing = False
            for i in range(1, props.Count + 1):
                prop = props.Item(i)
                if prop.Name == GUID_PROPERTY_NAME:
                    prop.Value = guid
                    existing = True
                    break
            
            if not existing:
                # Add new property (msoPropertyTypeString = 4)
                props.Add(GUID_PROPERTY_NAME, False, 4, guid)
            
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
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            return False
        
        try:
            # Set the document variable
            self._set_doc_variable(doc, var_name, var_value)
            
            # Insert a DOCVARIABLE field at cursor position
            selection = self._get_word_app().Selection
            
            # wdFieldDocVariable = 64
            field = selection.Fields.Add(
                Range=selection.Range,
                Type=64,  # wdFieldDocVariable
                Text=var_name,
                PreserveFormatting=True
            )
            
            # Update the field to show the value
            field.Update()
            
            return True
        except Exception as e:
            logging.error(f"Error inserting variable: {e}")
            return False

    def _set_doc_variable(self, doc, name: str, value: str):
        """Set a document variable value."""
        try:
            # Check if variable exists
            found = False
            for i in range(1, doc.Variables.Count + 1):
                var = doc.Variables.Item(i)
                if var.Name == name:
                    var.Value = value
                    found = True
                    break
            
            if not found:
                doc.Variables.Add(name, value)
        except Exception as e:
            logging.error(f"Error setting document variable: {e}")
            raise

    def get_doc_variable_value(self, var_name: str, doc=None) -> Optional[str]:
        """Get the current value of a document variable."""
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            return None
        
        try:
            for i in range(1, doc.Variables.Count + 1):
                var = doc.Variables.Item(i)
                if var.Name == var_name:
                    return var.Value
        except Exception as e:
            logging.error(f"Error getting document variable: {e}")
        
        return None

    # -------------------------
    # Scanning
    # -------------------------

    def scan_document(self, doc=None) -> DocumentInfo:
        """
        Scan a document to find all DOCVARIABLE fields.
        Returns document info including list of variable names used.
        """
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            raise ValueError("No active document")
        
        # Get or create GUID
        guid = self.get_document_guid(doc)
        if guid is None:
            import uuid
            guid = str(uuid.uuid4())
            self.set_document_guid(guid, doc)
        
        # Get document name and path
        name = doc.Name
        try:
            path = doc.FullName
        except:
            path = name  # Document not yet saved
        
        # Find all DOCVARIABLE fields
        variable_names = []
        
        # wdFieldDocVariable = 64
        for field in doc.Fields:
            if field.Type == 64:  # DOCVARIABLE
                # Extract variable name from field code
                # Field code looks like: " DOCVARIABLE  VarName  \* MERGEFORMAT "
                code = field.Code.Text.strip()
                parts = code.split()
                if len(parts) >= 2:
                    var_name = parts[1].strip('"')
                    if var_name not in variable_names:
                        variable_names.append(var_name)
        
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
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            return []
        
        updated = []
        
        try:
            # Update document variables
            for name, value in variables.items():
                current = self.get_doc_variable_value(name, doc)
                if current is not None and current != value:
                    self._set_doc_variable(doc, name, value)
                    updated.append(name)
                elif current is None:
                    # Variable doesn't exist in doc yet, add it
                    self._set_doc_variable(doc, name, value)
            
            # Refresh all fields to show new values
            if updated:
                doc.Fields.Update()
            
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
        if doc is None:
            doc = self.get_active_document()
        if doc is None:
            return {}
        
        stale = {}
        
        try:
            for i in range(1, doc.Variables.Count + 1):
                var = doc.Variables.Item(i)
                name = var.Name
                doc_value = var.Value
                
                if name in db_variables:
                    db_value = db_variables[name]
                    if doc_value != db_value:
                        stale[name] = (doc_value, db_value)
        except Exception as e:
            logging.error(f"Error checking stale variables: {e}")
        
        return stale
