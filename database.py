"""
Database module for variable storage and document tracking.
Uses SQLite for local persistence.
"""

import sqlite3
import sys
import os
from pathlib import Path
from datetime import datetime
from typing import Optional
import uuid


def get_app_dir():
    """Get the application directory, handling frozen (PyInstaller) apps."""
    if getattr(sys, 'frozen', False):
        # Running as compiled app
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))


def get_db_path(filename="variables.db"):
    """Get the full path to the database file."""
    return os.path.join(get_app_dir(), filename)


class VariableDatabase:
    def __init__(self, db_path: str = None):
        if db_path is None:
            db_path = get_db_path()
        self.db_path = db_path
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        """Initialize database schema."""
        conn = self._get_connection()
        cursor = conn.cursor()

        # Variables table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS variables (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                value TEXT NOT NULL,
                unit TEXT,
                description TEXT,
                excel_file TEXT,
                excel_sheet TEXT,
                excel_cell TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Documents table - tracks documents we've seen
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                guid TEXT UNIQUE NOT NULL,
                name TEXT,
                path TEXT,
                doc_type TEXT DEFAULT 'word',
                first_seen TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_scanned TIMESTAMP
            )
        """)

        # Usage table - which variables are in which documents
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS usage (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                variable_id INTEGER NOT NULL,
                document_id INTEGER NOT NULL,
                with_unit INTEGER DEFAULT 0,
                last_verified TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (variable_id) REFERENCES variables(id) ON DELETE CASCADE,
                FOREIGN KEY (document_id) REFERENCES documents(id) ON DELETE CASCADE,
                UNIQUE(variable_id, document_id)
            )
        """)

        # Add with_unit column if it doesn't exist (for existing databases)
        try:
            cursor.execute("ALTER TABLE usage ADD COLUMN with_unit INTEGER DEFAULT 0")
        except sqlite3.OperationalError:
            pass  # Column already exists

        # Add Excel link columns if they don't exist (for existing databases)
        for col in ['excel_file', 'excel_sheet', 'excel_cell']:
            try:
                cursor.execute(f"ALTER TABLE variables ADD COLUMN {col} TEXT")
            except sqlite3.OperationalError:
                pass  # Column already exists

        # Excel files table - tracks Excel files by GUID
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS excel_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                guid TEXT UNIQUE NOT NULL,
                name TEXT,
                path TEXT,
                first_seen TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_synced TIMESTAMP
            )
        """)

        # Add excel_file_id column to variables if it doesn't exist
        try:
            cursor.execute("ALTER TABLE variables ADD COLUMN excel_file_id INTEGER REFERENCES excel_files(id)")
        except sqlite3.OperationalError:
            pass  # Column already exists

        # Excel ranges table - saved ranges for batch syncing
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS excel_ranges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                start_cell TEXT NOT NULL,
                excel_file_id INTEGER REFERENCES excel_files(id),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_synced TIMESTAMP
            )
        """)

        # Add excel_file_id column to excel_ranges if it doesn't exist
        try:
            cursor.execute("ALTER TABLE excel_ranges ADD COLUMN excel_file_id INTEGER REFERENCES excel_files(id)")
        except sqlite3.OperationalError:
            pass  # Column already exists

        conn.commit()
        conn.close()

    # -------------------------
    # Variable CRUD operations
    # -------------------------

    def add_variable(self, name: str, value: str, unit: str = "", description: str = "") -> int:
        """Add a new variable. Returns the new variable ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO variables (name, value, unit, description) VALUES (?, ?, ?, ?)",
            (name, value, unit, description)
        )
        conn.commit()
        var_id = cursor.lastrowid
        conn.close()
        return var_id

    def update_variable(self, var_id: int, name: str = None, value: str = None,
                        unit: str = None, description: str = None,
                        excel_file: str = None, excel_sheet: str = None,
                        excel_cell: str = None) -> bool:
        """Update an existing variable. Only updates provided fields."""
        conn = self._get_connection()
        cursor = conn.cursor()

        updates = []
        params = []

        if name is not None:
            updates.append("name = ?")
            params.append(name)
        if value is not None:
            updates.append("value = ?")
            params.append(value)
        if unit is not None:
            updates.append("unit = ?")
            params.append(unit)
        if description is not None:
            updates.append("description = ?")
            params.append(description)
        if excel_file is not None:
            updates.append("excel_file = ?")
            params.append(excel_file if excel_file else None)
        if excel_sheet is not None:
            updates.append("excel_sheet = ?")
            params.append(excel_sheet if excel_sheet else None)
        if excel_cell is not None:
            updates.append("excel_cell = ?")
            params.append(excel_cell if excel_cell else None)

        if not updates:
            return False

        updates.append("updated_at = CURRENT_TIMESTAMP")
        params.append(var_id)

        cursor.execute(
            f"UPDATE variables SET {', '.join(updates)} WHERE id = ?",
            params
        )
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        return success

    def delete_variable(self, var_id: int) -> bool:
        """Delete a variable by ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM variables WHERE id = ?", (var_id,))
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        return success

    def get_variable(self, var_id: int) -> Optional[dict]:
        """Get a single variable by ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM variables WHERE id = ?", (var_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_variable_by_name(self, name: str) -> Optional[dict]:
        """Get a single variable by name."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM variables WHERE name = ?", (name,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_all_variables(self) -> list[dict]:
        """Get all variables."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM variables ORDER BY name")
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_variables_with_excel_links(self) -> list[dict]:
        """Get all variables that have Excel cell links."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM variables
            WHERE excel_file IS NOT NULL AND excel_file != ''
            ORDER BY name
        """)
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    # -------------------------
    # Document operations
    # -------------------------

    def register_document(self, guid: str, name: str, path: str, doc_type: str = "word") -> int:
        """Register a new document or return existing ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # Check if already exists
        cursor.execute("SELECT id FROM documents WHERE guid = ?", (guid,))
        row = cursor.fetchone()
        
        if row:
            # Update name/path in case they changed
            cursor.execute(
                "UPDATE documents SET name = ?, path = ? WHERE guid = ?",
                (name, path, guid)
            )
            conn.commit()
            doc_id = row['id']
        else:
            cursor.execute(
                "INSERT INTO documents (guid, name, path, doc_type) VALUES (?, ?, ?, ?)",
                (guid, name, path, doc_type)
            )
            conn.commit()
            doc_id = cursor.lastrowid
        
        conn.close()
        return doc_id

    def get_document_by_guid(self, guid: str) -> Optional[dict]:
        """Get a document by its GUID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM documents WHERE guid = ?", (guid,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def update_document_scanned(self, doc_id: int):
        """Update the last_scanned timestamp for a document."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE documents SET last_scanned = CURRENT_TIMESTAMP WHERE id = ?",
            (doc_id,)
        )
        conn.commit()
        conn.close()

    # -------------------------
    # Usage tracking
    # -------------------------

    def record_usage(self, variable_id: int, document_id: int, with_unit: bool = False):
        """Record that a variable is used in a document."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO usage (variable_id, document_id, with_unit, last_verified)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(variable_id, document_id)
            DO UPDATE SET with_unit = ?, last_verified = CURRENT_TIMESTAMP
        """, (variable_id, document_id, 1 if with_unit else 0, 1 if with_unit else 0))
        conn.commit()
        conn.close()

    def clear_usage_for_document(self, document_id: int):
        """Clear all usage records for a document (before re-scanning)."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM usage WHERE document_id = ?", (document_id,))
        conn.commit()
        conn.close()

    def get_variable_usage(self, variable_id: int) -> list[dict]:
        """Get all documents that use a specific variable."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT d.*, u.last_verified
            FROM documents d
            JOIN usage u ON d.id = u.document_id
            WHERE u.variable_id = ?
            ORDER BY d.name
        """, (variable_id,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_document_variables(self, document_id: int) -> list[dict]:
        """Get all variables used in a specific document."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT v.*, u.last_verified, u.with_unit
            FROM variables v
            JOIN usage u ON v.id = u.variable_id
            WHERE u.document_id = ?
            ORDER BY v.name
        """, (document_id,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_usage_with_unit(self, variable_name: str, document_guid: str) -> Optional[bool]:
        """Get the with_unit flag for a variable in a document."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT u.with_unit
            FROM usage u
            JOIN variables v ON v.id = u.variable_id
            JOIN documents d ON d.id = u.document_id
            WHERE v.name = ? AND d.guid = ?
        """, (variable_name, document_guid))
        row = cursor.fetchone()
        conn.close()
        return bool(row['with_unit']) if row else None

    def get_all_documents(self) -> list[dict]:
        """Get all tracked documents."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM documents ORDER BY name")
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def delete_document(self, doc_id: int):
        """Delete a document and its usage records."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM usage WHERE document_id = ?", (doc_id,))
        cursor.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
        conn.commit()
        conn.close()

    @staticmethod
    def generate_guid() -> str:
        """Generate a new GUID for a document."""
        return str(uuid.uuid4())

    # -------------------------
    # Excel File operations
    # -------------------------

    def register_excel_file(self, guid: str, name: str, path: str) -> int:
        """Register a new Excel file or return existing ID."""
        conn = self._get_connection()
        cursor = conn.cursor()

        # Check if already exists
        cursor.execute("SELECT id FROM excel_files WHERE guid = ?", (guid,))
        row = cursor.fetchone()

        if row:
            # Update name/path in case they changed
            cursor.execute(
                "UPDATE excel_files SET name = ?, path = ?, last_synced = CURRENT_TIMESTAMP WHERE guid = ?",
                (name, path, guid)
            )
            conn.commit()
            file_id = row['id']
        else:
            cursor.execute(
                "INSERT INTO excel_files (guid, name, path) VALUES (?, ?, ?)",
                (guid, name, path)
            )
            conn.commit()
            file_id = cursor.lastrowid

        conn.close()
        return file_id

    def get_excel_file_by_guid(self, guid: str) -> Optional[dict]:
        """Get an Excel file by its GUID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM excel_files WHERE guid = ?", (guid,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_excel_file_by_id(self, file_id: int) -> Optional[dict]:
        """Get an Excel file by its ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM excel_files WHERE id = ?", (file_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_all_excel_files(self) -> list[dict]:
        """Get all tracked Excel files."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM excel_files ORDER BY name")
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def update_excel_file_path(self, guid: str, new_path: str, new_name: str = None):
        """Update the path (and optionally name) for an Excel file."""
        conn = self._get_connection()
        cursor = conn.cursor()
        if new_name:
            cursor.execute(
                "UPDATE excel_files SET path = ?, name = ? WHERE guid = ?",
                (new_path, new_name, guid)
            )
        else:
            cursor.execute(
                "UPDATE excel_files SET path = ? WHERE guid = ?",
                (new_path, guid)
            )
        conn.commit()
        conn.close()

    def delete_excel_file(self, file_id: int):
        """Delete an Excel file record."""
        conn = self._get_connection()
        cursor = conn.cursor()
        # Clear references in variables
        cursor.execute("UPDATE variables SET excel_file_id = NULL WHERE excel_file_id = ?", (file_id,))
        # Clear references in excel_ranges
        cursor.execute("UPDATE excel_ranges SET excel_file_id = NULL WHERE excel_file_id = ?", (file_id,))
        # Delete the file record
        cursor.execute("DELETE FROM excel_files WHERE id = ?", (file_id,))
        conn.commit()
        conn.close()

    def link_variable_to_excel_file(self, var_id: int, excel_file_id: int):
        """Link a variable to an Excel file by ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE variables SET excel_file_id = ? WHERE id = ?",
            (excel_file_id, var_id)
        )
        conn.commit()
        conn.close()

    def get_variables_by_excel_file(self, excel_file_id: int) -> list[dict]:
        """Get all variables linked to a specific Excel file."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT * FROM variables WHERE excel_file_id = ? ORDER BY name",
            (excel_file_id,)
        )
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    # -------------------------
    # Excel Range operations
    # -------------------------

    def add_excel_range(self, name: str, file_path: str, sheet_name: str, start_cell: str) -> int:
        """Add a saved Excel range. Returns the new range ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO excel_ranges (name, file_path, sheet_name, start_cell) VALUES (?, ?, ?, ?)",
            (name, file_path, sheet_name, start_cell)
        )
        conn.commit()
        range_id = cursor.lastrowid
        conn.close()
        return range_id

    def get_all_excel_ranges(self) -> list[dict]:
        """Get all saved Excel ranges."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM excel_ranges ORDER BY name")
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_excel_range(self, range_id: int) -> Optional[dict]:
        """Get a single Excel range by ID."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM excel_ranges WHERE id = ?", (range_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def update_excel_range_synced(self, range_id: int):
        """Update the last_synced timestamp for a range."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE excel_ranges SET last_synced = CURRENT_TIMESTAMP WHERE id = ?",
            (range_id,)
        )
        conn.commit()
        conn.close()

    def delete_excel_range(self, range_id: int) -> bool:
        """Delete a saved Excel range."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM excel_ranges WHERE id = ?", (range_id,))
        conn.commit()
        success = cursor.rowcount > 0
        conn.close()
        return success
