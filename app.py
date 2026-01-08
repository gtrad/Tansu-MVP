"""
Variable Tracker - Main GUI Application
Uses CustomTkinter for modern look and feel.
"""

import customtkinter as ctk
from tkinter import messagebox, Menu
import logging
import os
import platform
import subprocess
import uuid
import webbrowser
from typing import Optional

from database import VariableDatabase
from docx_updater import update_docx_variables, get_docx_variables
from excel_reader import (
    validate_excel_link, sync_variables_from_excel, validate_excel_range,
    read_range_as_variables, read_sheet_preview, get_sheet_names,
    get_or_create_excel_guid
)
from version import __version__, __app_name__
from settings import is_first_run, mark_first_run_complete, get_setting, set_setting
from update_checker import check_for_update_async
from api_server import start_api_server, stop_api_server

# Try to import Word integration (Windows and macOS)
try:
    from word_integration import WordIntegration, HAS_WORD
except ImportError:
    HAS_WORD = False
    WordIntegration = None

# Configure appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Configure logging
logging.basicConfig(level=logging.INFO)


# -------------------------
# Dialog Classes
# -------------------------

class VariableDialog(ctk.CTkToplevel):
    """Dialog for adding/editing a variable."""

    def __init__(self, parent, title: str, variable: dict = None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x400")
        self.resizable(False, False)

        self.result = None
        self.variable = variable or {}

        self.transient(parent)
        self.grab_set()

        self._create_widgets()
        self._populate_fields()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self.name_entry.focus_set()

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(main_frame, text="Name:", anchor="w").pack(fill="x", pady=(0, 5))
        self.name_entry = ctk.CTkEntry(main_frame, width=350)
        self.name_entry.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(main_frame, text="Value:", anchor="w").pack(fill="x", pady=(0, 5))
        self.value_entry = ctk.CTkEntry(main_frame, width=350)
        self.value_entry.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(main_frame, text="Unit (optional):", anchor="w").pack(fill="x", pady=(0, 5))
        self.unit_entry = ctk.CTkEntry(main_frame, width=350)
        self.unit_entry.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(main_frame, text="Description (optional):", anchor="w").pack(fill="x", pady=(0, 5))
        self.desc_entry = ctk.CTkEntry(main_frame, width=350)
        self.desc_entry.pack(fill="x", pady=(0, 15))

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="gray",
                      command=self.destroy).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="Save", width=100,
                      command=self._save).pack(side="right")

    def _populate_fields(self):
        if self.variable:
            self.name_entry.insert(0, self.variable.get('name', ''))
            self.value_entry.insert(0, self.variable.get('value', ''))
            self.unit_entry.insert(0, self.variable.get('unit', ''))
            self.desc_entry.insert(0, self.variable.get('description', ''))

    def _save(self):
        name = self.name_entry.get().strip()
        value = self.value_entry.get().strip()

        if not name:
            messagebox.showerror("Error", "Name is required", parent=self)
            return
        if not value:
            messagebox.showerror("Error", "Value is required", parent=self)
            return

        self.result = {
            'name': name,
            'value': value,
            'unit': self.unit_entry.get().strip(),
            'description': self.desc_entry.get().strip()
        }
        self.destroy()


class ImportDialog(ctk.CTkToplevel):
    """Dialog for importing variables from pasted Excel/table data."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Import Variables")
        self.geometry("700x550")
        self.resizable(True, True)

        self.result = None  # List of variables to import

        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

        self.paste_text.focus_set()

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Instructions
        ctk.CTkLabel(main_frame,
                     text="Paste table data from Excel (columns: Name, Value, Unit, Description)",
                     anchor="w", font=("", 13, "bold")).pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(main_frame,
                     text="Tab-separated. Name and Value are required. Unit and Description are optional.",
                     anchor="w", text_color="gray", font=("", 11)).pack(fill="x", pady=(0, 10))

        # Paste area
        self.paste_text = ctk.CTkTextbox(main_frame, height=150)
        self.paste_text.pack(fill="x", pady=(0, 10))

        # Parse button
        ctk.CTkButton(main_frame, text="Parse Data", width=120,
                      command=self._parse_data).pack(anchor="w", pady=(0, 15))

        # Preview area
        ctk.CTkLabel(main_frame, text="Preview:", anchor="w", font=("", 12, "bold")).pack(fill="x", pady=(0, 5))

        self.preview_frame = ctk.CTkScrollableFrame(main_frame, height=200)
        self.preview_frame.pack(fill="both", expand=True, pady=(0, 10))

        # Status
        self.status_label = ctk.CTkLabel(main_frame, text="Paste your data and click 'Parse Data'",
                                          text_color="gray", anchor="w")
        self.status_label.pack(fill="x", pady=(0, 10))

        # Buttons
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="gray",
                      command=self.destroy).pack(side="right", padx=(10, 0))
        self.import_btn = ctk.CTkButton(btn_frame, text="Import", width=100,
                                         command=self._import, state="disabled")
        self.import_btn.pack(side="right")

        self.parsed_variables = []

    def _parse_data(self):
        """Parse the pasted text into variables."""
        # Clear preview
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        text = self.paste_text.get("1.0", "end").strip()
        if not text:
            self.status_label.configure(text="No data to parse", text_color="orange")
            return

        self.parsed_variables = []
        lines = text.split('\n')
        errors = []

        for i, line in enumerate(lines, 1):
            line = line.strip()
            if not line:
                continue

            # Split by tab (Excel copy) or multiple spaces
            if '\t' in line:
                parts = line.split('\t')
            else:
                # Fall back to splitting by 2+ spaces
                import re
                parts = re.split(r'\s{2,}', line)

            if len(parts) < 2:
                errors.append(f"Line {i}: Need at least Name and Value")
                continue

            name = parts[0].strip()
            value = parts[1].strip()
            unit = parts[2].strip() if len(parts) > 2 else ""
            description = parts[3].strip() if len(parts) > 3 else ""

            # Validate name (must be valid for Word DOCVARIABLE)
            if not name:
                errors.append(f"Line {i}: Name is empty")
                continue

            # Replace spaces with underscores for Word compatibility
            name = name.replace(' ', '_')

            if not value:
                errors.append(f"Line {i}: Value is empty for '{name}'")
                continue

            self.parsed_variables.append({
                'name': name,
                'value': value,
                'unit': unit,
                'description': description
            })

        # Show preview
        if self.parsed_variables:
            # Header row
            header = ctk.CTkFrame(self.preview_frame)
            header.pack(fill="x", pady=(0, 5))
            header.grid_columnconfigure((0, 1, 2, 3), weight=1)

            ctk.CTkLabel(header, text="Name", font=("", 11, "bold")).grid(row=0, column=0, sticky="w", padx=5)
            ctk.CTkLabel(header, text="Value", font=("", 11, "bold")).grid(row=0, column=1, sticky="w", padx=5)
            ctk.CTkLabel(header, text="Unit", font=("", 11, "bold")).grid(row=0, column=2, sticky="w", padx=5)
            ctk.CTkLabel(header, text="Description", font=("", 11, "bold")).grid(row=0, column=3, sticky="w", padx=5)

            for var in self.parsed_variables:
                row = ctk.CTkFrame(self.preview_frame, fg_color=("gray90", "gray20"))
                row.pack(fill="x", pady=1)
                row.grid_columnconfigure((0, 1, 2, 3), weight=1)

                ctk.CTkLabel(row, text=var['name'], anchor="w").grid(row=0, column=0, sticky="w", padx=5, pady=3)
                ctk.CTkLabel(row, text=var['value'], anchor="w").grid(row=0, column=1, sticky="w", padx=5, pady=3)
                ctk.CTkLabel(row, text=var['unit'] or "-", anchor="w", text_color="gray").grid(row=0, column=2, sticky="w", padx=5, pady=3)
                desc_text = var['description'][:30] + "..." if len(var['description']) > 30 else var['description'] or "-"
                ctk.CTkLabel(row, text=desc_text, anchor="w", text_color="gray").grid(row=0, column=3, sticky="w", padx=5, pady=3)

            self.import_btn.configure(state="normal")
            status = f"Found {len(self.parsed_variables)} variable(s)"
            if errors:
                status += f" ({len(errors)} error(s))"
            self.status_label.configure(text=status, text_color="green" if not errors else "orange")
        else:
            self.import_btn.configure(state="disabled")
            self.status_label.configure(text=f"No valid variables found. {len(errors)} error(s)", text_color="red")

        # Show errors if any
        if errors:
            for err in errors[:3]:
                err_label = ctk.CTkLabel(self.preview_frame, text=err, text_color="red", anchor="w")
                err_label.pack(fill="x", pady=1)
            if len(errors) > 3:
                ctk.CTkLabel(self.preview_frame, text=f"... and {len(errors) - 3} more errors",
                             text_color="red", anchor="w").pack(fill="x")

    def _import(self):
        """Import the parsed variables."""
        if self.parsed_variables:
            self.result = self.parsed_variables
            self.destroy()


class UsageDialog(ctk.CTkToplevel):
    """Dialog showing where a variable is used."""

    def __init__(self, parent, variable_name: str, documents: list[dict]):
        super().__init__(parent)
        self.title(f"Usage: {variable_name}")
        self.geometry("500x400")

        self.transient(parent)
        self.grab_set()

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(main_frame, text=f"'{variable_name}' is used in {len(documents)} document(s):",
                     anchor="w", font=("", 14, "bold")).pack(fill="x", pady=(0, 15))

        scroll_frame = ctk.CTkScrollableFrame(main_frame, width=440, height=280)
        scroll_frame.pack(fill="both", expand=True)

        if documents:
            for doc in documents:
                doc_frame = ctk.CTkFrame(scroll_frame)
                doc_frame.pack(fill="x", pady=5, padx=5)

                ctk.CTkLabel(doc_frame, text=doc.get('name', 'Unknown'),
                             font=("", 12, "bold"), anchor="w").pack(fill="x", padx=10, pady=(10, 0))
                path_display = doc.get('path', '')
                if path_display.startswith('unsaved:'):
                    path_display = "(unsaved document)"
                ctk.CTkLabel(doc_frame, text=path_display,
                             font=("", 10), text_color="gray", anchor="w").pack(fill="x", padx=10, pady=(0, 10))
        else:
            ctk.CTkLabel(scroll_frame, text="Not used in any tracked documents.",
                         text_color="gray").pack(pady=20)

        ctk.CTkLabel(main_frame, text="Deleted files are automatically removed from this list.",
                     font=("", 10), text_color="gray").pack(pady=(10, 0))

        ctk.CTkButton(main_frame, text="Close", width=100,
                      command=self.destroy).pack(pady=(10, 0))


class LinkExcelDialog(ctk.CTkToplevel):
    """Dialog for linking a variable to an Excel cell."""

    def __init__(self, parent, variable: dict):
        super().__init__(parent)
        self.title(f"Link to Excel: {variable['name']}")
        self.geometry("550x350")
        self.resizable(False, False)

        self.result = None
        self.variable = variable

        self.transient(parent)
        self.grab_set()

        self._create_widgets()
        self._populate_fields()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Current value display
        ctk.CTkLabel(main_frame, text=f"Variable: {self.variable['name']}",
                     font=("", 14, "bold"), anchor="w").pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(main_frame, text=f"Current value: {self.variable['value']}",
                     text_color="gray", anchor="w").pack(fill="x", pady=(0, 15))

        # Excel file path
        ctk.CTkLabel(main_frame, text="Excel File:", anchor="w").pack(fill="x", pady=(0, 5))
        file_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        file_frame.pack(fill="x", pady=(0, 15))
        self.file_entry = ctk.CTkEntry(file_frame, width=400)
        self.file_entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(file_frame, text="Browse", width=80,
                      command=self._browse_file).pack(side="right", padx=(10, 0))

        # Sheet name
        ctk.CTkLabel(main_frame, text="Sheet Name:", anchor="w").pack(fill="x", pady=(0, 5))
        self.sheet_entry = ctk.CTkEntry(main_frame, width=500)
        self.sheet_entry.pack(fill="x", pady=(0, 15))

        # Cell reference
        ctk.CTkLabel(main_frame, text="Cell Reference (e.g., A1, B5):", anchor="w").pack(fill="x", pady=(0, 5))
        self.cell_entry = ctk.CTkEntry(main_frame, width=500)
        self.cell_entry.pack(fill="x", pady=(0, 10))

        # Test/status
        test_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        test_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(test_frame, text="Test Link", width=100,
                      command=self._test_link).pack(side="left")
        self.status_label = ctk.CTkLabel(test_frame, text="", anchor="w")
        self.status_label.pack(side="left", padx=(15, 0))

        # Buttons
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkButton(btn_frame, text="Remove Link", width=100, fg_color="darkred",
                      command=self._remove_link).pack(side="left")
        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="gray",
                      command=self.destroy).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="Save", width=100,
                      command=self._save).pack(side="right")

    def _populate_fields(self):
        if self.variable.get('excel_file'):
            self.file_entry.insert(0, self.variable['excel_file'])
        if self.variable.get('excel_sheet'):
            self.sheet_entry.insert(0, self.variable['excel_sheet'])
        if self.variable.get('excel_cell'):
            self.cell_entry.insert(0, self.variable['excel_cell'])

    def _browse_file(self):
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if file_path:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)

    def _test_link(self):
        file_path = self.file_entry.get().strip()
        sheet_name = self.sheet_entry.get().strip()
        cell_ref = self.cell_entry.get().strip().upper()

        if not all([file_path, sheet_name, cell_ref]):
            self.status_label.configure(text="Fill in all fields first", text_color="orange")
            return

        is_valid, message = validate_excel_link(file_path, sheet_name, cell_ref)
        self.status_label.configure(
            text=message,
            text_color="green" if is_valid else "red"
        )

    def _save(self):
        file_path = self.file_entry.get().strip()
        sheet_name = self.sheet_entry.get().strip()
        cell_ref = self.cell_entry.get().strip().upper()

        if file_path and (not sheet_name or not cell_ref):
            messagebox.showerror("Error", "Please fill in Sheet Name and Cell Reference", parent=self)
            return

        self.result = {
            'excel_file': file_path,
            'excel_sheet': sheet_name,
            'excel_cell': cell_ref
        }
        self.destroy()

    def _remove_link(self):
        self.result = {
            'excel_file': '',
            'excel_sheet': '',
            'excel_cell': ''
        }
        self.destroy()


class ImportRangeDialog(ctk.CTkToplevel):
    """Dialog for importing variables from an Excel range with visual preview."""

    def __init__(self, parent, save_callback=None):
        super().__init__(parent)
        self.title("Import from Excel")
        self.geometry("1000x750")
        self.resizable(True, True)

        self.result = None
        self.save_result = None  # For saving the range config
        self.selected_cell = None
        self.cell_buttons = {}
        self.sheet_data = []
        self.current_file = None
        self.current_sheet = None
        self.save_callback = save_callback  # Callback for saving range

        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # Top section: File and sheet selection
        top_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        top_frame.pack(fill="x", pady=(0, 10))

        # File selection
        ctk.CTkLabel(top_frame, text="Excel File:", anchor="w").pack(side="left")
        self.file_entry = ctk.CTkEntry(top_frame, width=300)
        self.file_entry.pack(side="left", padx=(10, 5))
        ctk.CTkButton(top_frame, text="Browse", width=70,
                      command=self._browse_file).pack(side="left", padx=(0, 15))

        # Sheet dropdown
        ctk.CTkLabel(top_frame, text="Sheet:", anchor="w").pack(side="left")
        self.sheet_var = ctk.StringVar(value="")
        self.sheet_dropdown = ctk.CTkOptionMenu(top_frame, variable=self.sheet_var,
                                                 values=[""], width=120,
                                                 command=self._on_sheet_change)
        self.sheet_dropdown.pack(side="left", padx=(10, 0))

        # Action buttons on the right
        self.save_btn = ctk.CTkButton(top_frame, text="Save Range", width=100,
                                       fg_color="#217346", command=self._save_range, state="disabled")
        self.save_btn.pack(side="right", padx=(10, 0))
        self.import_btn = ctk.CTkButton(top_frame, text="Import", width=80,
                                         command=self._import, state="disabled")
        self.import_btn.pack(side="right")

        # Instructions
        ctk.CTkLabel(main_frame,
                     text="Click a cell to set the starting point. Data reads: Name | Value | Unit (3 columns, down until empty).",
                     anchor="w", text_color="gray", font=("", 11)).pack(fill="x", pady=(0, 10))

        # Spreadsheet preview grid
        grid_container = ctk.CTkFrame(main_frame)
        grid_container.pack(fill="both", expand=True, pady=(0, 10))

        # Column headers (A, B, C, etc.)
        self.header_frame = ctk.CTkFrame(grid_container, fg_color="transparent")
        self.header_frame.pack(fill="x")

        # Scrollable grid area
        self.grid_scroll = ctk.CTkScrollableFrame(grid_container, height=350)
        self.grid_scroll.pack(fill="both", expand=True)

        # Selection info and preview
        info_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        info_frame.pack(fill="x", pady=(0, 10))

        self.selection_label = ctk.CTkLabel(info_frame, text="Selected: None", anchor="w", font=("", 12, "bold"))
        self.selection_label.pack(side="left")

        self.preview_label = ctk.CTkLabel(info_frame, text="", anchor="w", text_color="gray")
        self.preview_label.pack(side="left", padx=(20, 0))

        # Variables preview (what will be imported)
        ctk.CTkLabel(main_frame, text="Variables to import:", anchor="w", font=("", 12, "bold")).pack(fill="x", pady=(0, 5))
        self.vars_preview = ctk.CTkScrollableFrame(main_frame, height=120)
        self.vars_preview.pack(fill="x", pady=(0, 10))

        self.vars_status = ctk.CTkLabel(main_frame, text="Select a starting cell to preview variables", text_color="gray", anchor="w")
        self.vars_status.pack(fill="x", pady=(0, 10))

        # Bottom buttons
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="gray",
                      command=self.destroy).pack(side="right")

        self.loaded_variables = []

    def _browse_file(self):
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if file_path:
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, file_path)
            self._load_file(file_path)

    def _load_file(self, file_path):
        """Load the Excel file and populate sheet dropdown."""
        try:
            sheets = get_sheet_names(file_path)
            self.sheet_dropdown.configure(values=sheets)
            if sheets:
                self.sheet_var.set(sheets[0])
                self.current_file = file_path
                self._load_sheet_preview(file_path, sheets[0])
        except Exception as e:
            self.vars_status.configure(text=f"Error loading file: {e}", text_color="red")

    def _on_sheet_change(self, sheet_name):
        """Handle sheet selection change."""
        if self.current_file and sheet_name:
            self._load_sheet_preview(self.current_file, sheet_name)

    def _load_sheet_preview(self, file_path, sheet_name):
        """Load and display sheet preview grid."""
        try:
            self.sheet_data = read_sheet_preview(file_path, sheet_name, max_rows=50, max_cols=10)
            self.current_sheet = sheet_name
            self._build_grid()
            self.selected_cell = None
            self.selection_label.configure(text="Selected: None")
            self.preview_label.configure(text="")
            self._clear_vars_preview()
            self.import_btn.configure(state="disabled")
        except Exception as e:
            self.vars_status.configure(text=f"Error loading sheet: {e}", text_color="red")

    def _build_grid(self):
        """Build the spreadsheet-like grid using tkinter Labels for speed."""
        import tkinter as tk

        # Clear existing
        for widget in self.header_frame.winfo_children():
            widget.destroy()
        for widget in self.grid_scroll.winfo_children():
            widget.destroy()
        self.cell_labels = {}

        if not self.sheet_data:
            return

        num_cols = len(self.sheet_data[0]) if self.sheet_data else 0

        # Column headers
        ctk.CTkLabel(self.header_frame, text="", width=40).pack(side="left")
        for col_idx in range(num_cols):
            col_letter = chr(65 + col_idx)
            lbl = ctk.CTkLabel(self.header_frame, text=col_letter, width=100, font=("", 11, "bold"))
            lbl.pack(side="left", padx=1)

        # Data rows - use tk.Label for speed
        bg_color = "#2b2b2b"  # Dark background to match theme
        for row_idx, row_data in enumerate(self.sheet_data):
            row_frame = tk.Frame(self.grid_scroll, bg=bg_color)
            row_frame.pack(fill="x", pady=1)

            # Row number
            tk.Label(row_frame, text=str(row_idx + 1), width=4, font=("", 10),
                    fg="gray", bg=bg_color).pack(side="left")

            for col_idx, cell_value in enumerate(row_data):
                col_letter = chr(65 + col_idx)
                cell_ref = f"{col_letter}{row_idx + 1}"

                # Truncate long values
                display_val = str(cell_value)[:12] if cell_value else ""

                lbl = tk.Label(
                    row_frame,
                    text=display_val,
                    width=12,
                    height=1,
                    font=("", 10),
                    bg="#3d3d3d",
                    fg="white",
                    relief="flat",
                    cursor="hand2"
                )
                lbl.pack(side="left", padx=1)

                # Bind click event
                lbl.bind("<Button-1>", lambda e, r=row_idx, c=col_idx, ref=cell_ref: self._on_cell_click(r, c, ref))

                self.cell_labels[(row_idx, col_idx)] = lbl

    def _on_cell_click(self, row_idx, col_idx, cell_ref):
        """Handle cell selection."""
        # Reset all cells to default color
        for (r, c), lbl in self.cell_labels.items():
            lbl.configure(bg="#3d3d3d")

        self.selected_cell = (row_idx, col_idx, cell_ref)

        # First get the variables to know exactly which rows to highlight
        if self.current_file and self.current_sheet:
            is_valid, message, variables = validate_excel_range(self.current_file, self.current_sheet, cell_ref)

            if is_valid and variables:
                # Highlight only the rows that will be imported
                rows_to_highlight = [v['row'] - 1 for v in variables]  # Convert to 0-indexed

                for r in rows_to_highlight:
                    for c_offset in range(3):
                        c = col_idx + c_offset
                        if (r, c) in self.cell_labels:
                            if c_offset == 0:
                                self.cell_labels[(r, c)].configure(bg="#2E8B57")  # Green for Name
                            elif c_offset == 1:
                                self.cell_labels[(r, c)].configure(bg="#4682B4")  # Blue for Value
                            else:
                                self.cell_labels[(r, c)].configure(bg="#8B668B")  # Purple for Unit

                self.loaded_variables = variables
                self._show_vars_preview(variables)
                self.vars_status.configure(text=f"Found {len(variables)} variable(s) to import", text_color="green")
                self.import_btn.configure(state="normal")
                self.save_btn.configure(state="normal")
            else:
                self._clear_vars_preview()
                self.vars_status.configure(text=message if message else "No variables found", text_color="orange")
                self.import_btn.configure(state="disabled")
                self.save_btn.configure(state="disabled")
                self.loaded_variables = []

        # Update selection info
        cell_value = self.sheet_data[row_idx][col_idx] if col_idx < len(self.sheet_data[row_idx]) else ""
        self.selection_label.configure(text=f"Selected: {cell_ref}")
        self.preview_label.configure(text=f"Value: {cell_value}" if cell_value else "(empty)")

    def _show_vars_preview(self, variables):
        """Show variables in the preview area."""
        self._clear_vars_preview()

        # Header
        header = ctk.CTkFrame(self.vars_preview)
        header.pack(fill="x", pady=(0, 3))
        ctk.CTkLabel(header, text="Name", width=150, font=("", 10, "bold"), anchor="w").pack(side="left", padx=5)
        ctk.CTkLabel(header, text="Value", width=150, font=("", 10, "bold"), anchor="w").pack(side="left", padx=5)
        ctk.CTkLabel(header, text="Unit", width=80, font=("", 10, "bold"), anchor="w").pack(side="left", padx=5)

        for var in variables[:10]:
            row = ctk.CTkFrame(self.vars_preview, fg_color=("gray90", "gray20"))
            row.pack(fill="x", pady=1)
            ctk.CTkLabel(row, text=var['name'], width=150, anchor="w", font=("", 10)).pack(side="left", padx=5, pady=2)
            ctk.CTkLabel(row, text=var['value'], width=150, anchor="w", font=("", 10)).pack(side="left", padx=5, pady=2)
            ctk.CTkLabel(row, text=var['unit'] or "-", width=80, anchor="w", font=("", 10), text_color="gray").pack(side="left", padx=5, pady=2)

        if len(variables) > 10:
            ctk.CTkLabel(self.vars_preview, text=f"... and {len(variables) - 10} more",
                        text_color="gray", font=("", 10)).pack(anchor="w", padx=5)

    def _clear_vars_preview(self):
        """Clear the variables preview."""
        for widget in self.vars_preview.winfo_children():
            widget.destroy()
        self.vars_status.configure(text="Select a starting cell to preview variables", text_color="gray")

    def _save_range(self):
        """Save the Excel range configuration for future syncing."""
        if not self.current_file or not self.current_sheet or not self.selected_cell:
            return

        cell_ref = self.selected_cell[2]

        # Ask for a name for this range
        name_dialog = ctk.CTkInputDialog(
            text=f"Enter a name for this Excel range:\n({self.current_sheet} starting at {cell_ref})",
            title="Save Excel Range"
        )
        range_name = name_dialog.get_input()

        if not range_name:
            return

        self.save_result = {
            'name': range_name.strip(),
            'file_path': self.current_file,
            'sheet_name': self.current_sheet,
            'start_cell': cell_ref,
            'variables': self.loaded_variables
        }
        self.destroy()

    def _import(self):
        """Import the loaded variables."""
        if self.loaded_variables:
            self.result = self.loaded_variables
            self.destroy()


# -------------------------
# First Run & Update Dialogs
# -------------------------

class FirstRunDialog(ctk.CTkToplevel):
    """Dialog shown on first launch to get user consent for update checks."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title(f"Welcome to {__app_name__}")
        self.geometry("450x280")
        self.resizable(False, False)

        self.result = None  # Will be True if user accepts, False otherwise

        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=25, pady=20)

        ctk.CTkLabel(
            main_frame,
            text=f"Welcome to {__app_name__}!",
            font=("", 20, "bold")
        ).pack(pady=(0, 15))

        ctk.CTkLabel(
            main_frame,
            text="Thank you for using Tansu. To help improve the app,\nwe can check for updates when you start the app.",
            justify="center"
        ).pack(pady=(0, 10))

        ctk.CTkLabel(
            main_frame,
            text="This only checks GitHub for new releases.\nNo personal data is collected.",
            text_color="gray",
            justify="center"
        ).pack(pady=(0, 20))

        self.update_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            main_frame,
            text="Check for updates on startup",
            variable=self.update_var
        ).pack(pady=(0, 20))

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(
            btn_frame,
            text="Get Started",
            width=120,
            command=self._on_accept
        ).pack(side="right")

    def _on_accept(self):
        self.result = self.update_var.get()
        self.destroy()


class UpdateAvailableDialog(ctk.CTkToplevel):
    """Dialog shown when a new version is available."""

    def __init__(self, parent, update_info: dict):
        super().__init__(parent)
        self.title("Update Available")
        self.geometry("400x200")
        self.resizable(False, False)

        self.update_info = update_info

        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=25, pady=20)

        ctk.CTkLabel(
            main_frame,
            text="A new version is available!",
            font=("", 16, "bold")
        ).pack(pady=(0, 10))

        ctk.CTkLabel(
            main_frame,
            text=f"Version {self.update_info['version']} is now available.\nYou have version {__version__}.",
            justify="center"
        ).pack(pady=(0, 20))

        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(
            btn_frame,
            text="Download",
            width=100,
            command=self._on_download
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame,
            text="Later",
            width=80,
            fg_color="gray",
            command=self.destroy
        ).pack(side="left")

    def _on_download(self):
        url = self.update_info.get('download_url') or self.update_info.get('url')
        if url:
            webbrowser.open(url)
        self.destroy()


class FeedbackDialog(ctk.CTkToplevel):
    """Dialog for submitting feedback/issues."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Send Feedback")
        self.geometry("500x450")
        self.resizable(False, False)

        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=25, pady=20)

        ctk.CTkLabel(
            main_frame,
            text="Send Feedback",
            font=("", 18, "bold")
        ).pack(pady=(0, 5))

        ctk.CTkLabel(
            main_frame,
            text=f"Tansu v{__version__} (Beta)",
            text_color="gray"
        ).pack(pady=(0, 15))

        # Feedback type
        type_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        type_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(type_frame, text="Type:", width=60, anchor="w").pack(side="left")
        self.feedback_type = ctk.CTkComboBox(
            type_frame,
            values=["Bug Report", "Feature Request", "General Feedback", "Question"],
            width=200
        )
        self.feedback_type.pack(side="left")
        self.feedback_type.set("General Feedback")

        # Description
        ctk.CTkLabel(main_frame, text="Description:", anchor="w").pack(fill="x", pady=(10, 5))
        self.description_text = ctk.CTkTextbox(main_frame, height=150)
        self.description_text.pack(fill="x", pady=(0, 10))

        # Email (optional)
        email_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        email_frame.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(email_frame, text="Email (optional):", width=110, anchor="w").pack(side="left")
        self.email_entry = ctk.CTkEntry(email_frame, width=250, placeholder_text="For follow-up")
        self.email_entry.pack(side="left")

        # Buttons
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        ctk.CTkButton(
            btn_frame,
            text="Submit on GitHub",
            width=140,
            command=self._submit_github
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame,
            text="Copy to Clipboard",
            width=130,
            fg_color="gray",
            command=self._copy_to_clipboard
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame,
            text="Cancel",
            width=80,
            fg_color="gray",
            command=self.destroy
        ).pack(side="left")

    def _get_feedback_text(self) -> str:
        """Format the feedback for submission."""
        from version import GITHUB_REPO
        feedback_type = self.feedback_type.get()
        description = self.description_text.get("1.0", "end-1c").strip()
        email = self.email_entry.get().strip()

        text = f"**Type:** {feedback_type}\n\n"
        text += f"**Description:**\n{description}\n\n"
        text += f"**App Version:** {__version__}\n"
        text += f"**OS:** {platform.system()} {platform.release()}\n"
        if email:
            text += f"**Contact:** {email}\n"

        return text

    def _submit_github(self):
        """Open GitHub issues page with pre-filled content."""
        from version import GITHUB_REPO
        import urllib.parse

        feedback_type = self.feedback_type.get()
        description = self.description_text.get("1.0", "end-1c").strip()

        if not description:
            messagebox.showwarning("Missing Description", "Please enter a description.")
            return

        # Create GitHub issue URL with pre-filled body
        title = urllib.parse.quote(f"[{feedback_type}] ")
        body = urllib.parse.quote(self._get_feedback_text())

        url = f"https://github.com/{GITHUB_REPO}/issues/new?title={title}&body={body}"
        webbrowser.open(url)
        self.destroy()

    def _copy_to_clipboard(self):
        """Copy feedback to clipboard."""
        description = self.description_text.get("1.0", "end-1c").strip()

        if not description:
            messagebox.showwarning("Missing Description", "Please enter a description.")
            return

        text = self._get_feedback_text()
        self.clipboard_clear()
        self.clipboard_append(text)

        messagebox.showinfo("Copied", "Feedback copied to clipboard!")
        self.destroy()


# -------------------------
# Quick Insert Popup (Global Hotkey)
# -------------------------

def run_applescript(script: str) -> str:
    """Run an AppleScript and return the output (Mac only)."""
    result = subprocess.run(
        ['osascript', '-e', script],
        capture_output=True,
        text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr}")
    return result.stdout.strip()


def check_word_document_open() -> bool:
    """Check if Word has a document open."""
    if platform.system() == "Darwin":
        # Mac: Use AppleScript
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
    elif platform.system() == "Windows":
        # Windows: Use COM
        try:
            import win32com.client
            word = win32com.client.GetObject(Class="Word.Application")
            return word.Documents.Count > 0
        except:
            return False
    return False


def insert_variable_into_word(var_name: str, var_value: str, as_field: bool = True) -> bool:
    """Insert a variable into Word at the current cursor position."""
    if platform.system() == "Darwin":
        return _insert_variable_mac(var_name, var_value, as_field)
    elif platform.system() == "Windows":
        return _insert_variable_windows(var_name, var_value, as_field)
    return False


def _insert_variable_mac(var_name: str, var_value: str, as_field: bool) -> bool:
    """Insert variable on Mac using AppleScript."""
    try:
        name_escaped = var_name.replace('"', '\\"')
        value_escaped = var_value.replace('"', '\\"')

        if as_field:
            script = f'''
tell application "Microsoft Word"
    activate
    set doc to active document
    try
        delete variable "{name_escaped}" of doc
    end try
    make new variable at doc with properties {{name:"{name_escaped}"}}
    set variable value of variable "{name_escaped}" of doc to "{value_escaped}"
    set theSelection to selection
    make new field at text object of theSelection with properties {{field type:field doc variable, field text:"{name_escaped}"}}
end tell
tell application "System Events"
    delay 0.2
    key code 101
end tell
return "success"
'''
        else:
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


def _insert_variable_windows(var_name: str, var_value: str, as_field: bool) -> bool:
    """Insert variable on Windows using COM."""
    try:
        import win32com.client
        word = win32com.client.GetObject(Class="Word.Application")
        doc = word.ActiveDocument

        if as_field:
            # Set document variable
            try:
                doc.Variables(var_name).Delete()
            except:
                pass
            doc.Variables.Add(var_name, var_value)

            # Insert DOCVARIABLE field at cursor
            selection = word.Selection
            selection.Fields.Add(
                Range=selection.Range,
                Type=-1,  # wdFieldEmpty
                Text=f'DOCVARIABLE "{var_name}"',
                PreserveFormatting=True
            )

            # Update the field
            selection.Fields.Update()
        else:
            # Insert as plain text
            word.Selection.TypeText(var_value)

        return True
    except Exception as e:
        logging.error(f"Error inserting variable: {e}")
        return False


class QuickInsertPopup(ctk.CTkToplevel):
    """Popup for quickly inserting variables via hotkey."""

    def __init__(self, parent, variables: list):
        super().__init__(parent)
        self.title("Insert Variable")
        self.geometry("350x400")
        self.resizable(False, False)

        # Make it float on top (but don't bring parent to front)
        self.attributes("-topmost", True)

        self.variables = variables
        self.filtered_vars = variables.copy()
        self.selected_index = 0
        self.result = None

        self._create_widgets()

        # Center on screen
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - self.winfo_width()) // 2
        y = (screen_height - self.winfo_height()) // 3
        self.geometry(f"+{x}+{y}")

        # Force focus to this window and search entry
        self.lift()
        self.focus_force()
        self.search_entry.focus_set()

        # Bind keys
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", self._on_enter)
        self.bind("<Up>", self._on_up)
        self.bind("<Down>", self._on_down)

    def _create_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # Title
        hotkey_name = "Option+Space" if platform.system() == "Darwin" else "Alt+Space"
        ctk.CTkLabel(
            main_frame,
            text=f"Insert Variable ({hotkey_name})",
            font=("", 14, "bold")
        ).pack(pady=(0, 10))

        # Search entry
        self.search_entry = ctk.CTkEntry(
            main_frame,
            placeholder_text="Search variables...",
            width=300
        )
        self.search_entry.pack(fill="x", pady=(0, 10))
        self.search_entry.bind("<KeyRelease>", self._on_search)

        # Variables list
        self.list_frame = ctk.CTkScrollableFrame(main_frame, height=250)
        self.list_frame.pack(fill="both", expand=True, pady=(0, 10))

        self._update_list()

        # Insert options
        options_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        options_frame.pack(fill="x")

        ctk.CTkButton(
            options_frame,
            text="Insert as Field",
            width=150,
            command=lambda: self._insert(as_field=True)
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            options_frame,
            text="Insert as Text",
            width=120,
            fg_color="gray",
            command=lambda: self._insert(as_field=False)
        ).pack(side="left")

    def _update_list(self):
        # Clear existing
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        if not self.filtered_vars:
            ctk.CTkLabel(
                self.list_frame,
                text="No variables found",
                text_color="gray"
            ).pack(pady=20)
            return

        for i, var in enumerate(self.filtered_vars):
            is_selected = (i == self.selected_index)
            bg_color = "#1f538d" if is_selected else "transparent"

            row = ctk.CTkFrame(self.list_frame, fg_color=bg_color, corner_radius=5)
            row.pack(fill="x", pady=2)

            display = f"{var['name']}: {var['value']}"
            if var.get('unit'):
                display += f" {var['unit']}"

            label = ctk.CTkLabel(row, text=display, anchor="w")
            label.pack(fill="x", padx=10, pady=5)

            # Single click to select, double click to insert
            row.bind("<Button-1>", lambda e, idx=i: self._select_item(idx))
            label.bind("<Button-1>", lambda e, idx=i: self._select_item(idx))
            row.bind("<Double-Button-1>", lambda e, idx=i: self._select_and_insert(idx))
            label.bind("<Double-Button-1>", lambda e, idx=i: self._select_and_insert(idx))

    def _on_search(self, event):
        query = self.search_entry.get().lower()
        if query:
            self.filtered_vars = [
                v for v in self.variables
                if query in v['name'].lower() or query in str(v['value']).lower()
            ]
        else:
            self.filtered_vars = self.variables.copy()

        self.selected_index = 0
        self._update_list()

    def _on_up(self, event):
        if self.filtered_vars and self.selected_index > 0:
            self.selected_index -= 1
            self._update_list()

    def _on_down(self, event):
        if self.filtered_vars and self.selected_index < len(self.filtered_vars) - 1:
            self.selected_index += 1
            self._update_list()

    def _on_enter(self, event):
        self._insert(as_field=True)

    def _select_item(self, idx):
        """Select an item without inserting."""
        self.selected_index = idx
        self._update_list()

    def _select_and_insert(self, idx):
        self.selected_index = idx
        self._insert(as_field=True)

    def _insert(self, as_field: bool):
        if not self.filtered_vars:
            self.withdraw()
            self.destroy()
            return

        var = self.filtered_vars[self.selected_index]

        if not check_word_document_open():
            messagebox.showwarning(
                "No Word Document",
                "Please open a Word document first."
            )
            self.withdraw()
            self.destroy()
            return

        # Hide the popup immediately
        self.withdraw()

        # Insert the variable
        success = insert_variable_into_word(var['name'], var['value'], as_field=as_field)
        if not success:
            messagebox.showerror("Error", f"Failed to insert '{var['name']}'")

        # Now destroy
        self.destroy()


# -------------------------
# Main Application Window
# -------------------------

class VariableTrackerApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self.title(f"{__app_name__} v{__version__}")
        self.geometry("725x600")
        self.minsize(600, 400)

        # Set window icon
        self._set_icon()

        self.db = VariableDatabase()

        self.word: Optional[WordIntegration] = None
        if HAS_WORD and WordIntegration:
            try:
                self.word = WordIntegration()
            except Exception as e:
                logging.warning(f"Could not initialize Word integration: {e}")

        self._create_widgets()
        self._refresh_variable_list()

        self.attributes("-topmost", False)
        self.update()

        # Show first-run dialog if needed
        if is_first_run():
            self.after(100, self._show_first_run_dialog)
        elif get_setting("check_for_updates"):
            # Check for updates in background
            self.after(500, self._check_for_updates)

        # Start global hotkey listener (Option+Space)
        self._start_hotkey_listener()
        self._quick_insert_popup = None

        # Start API server for Word add-in
        start_api_server()

    def _set_icon(self):
        """Set the application window icon."""
        import os
        from database import get_app_dir

        # Find icon file
        app_dir = get_app_dir()
        icon_path = os.path.join(app_dir, "icon.png")

        # Try different locations
        if not os.path.exists(icon_path):
            icon_path = os.path.join(os.path.dirname(__file__), "icon.png")

        if os.path.exists(icon_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(icon_path)
                photo = ImageTk.PhotoImage(img)
                self.wm_iconphoto(True, photo)
                self._icon_photo = photo  # Keep reference to prevent garbage collection
            except Exception as e:
                logging.debug(f"Could not set icon: {e}")

    def _show_first_run_dialog(self):
        """Show the first-run welcome dialog."""
        dialog = FirstRunDialog(self)
        self.wait_window(dialog)

        # Save user's preference
        if dialog.result is not None:
            set_setting("check_for_updates", dialog.result)
            mark_first_run_complete()

            # Check for updates if user opted in
            if dialog.result:
                self.after(100, self._check_for_updates)

    def _check_for_updates(self):
        """Check for updates in the background."""
        def on_update_result(update_info):
            if update_info:
                # Schedule dialog on main thread
                self.after(0, lambda: self._show_update_dialog(update_info))

        check_for_update_async(on_update_result)

    def _show_update_dialog(self, update_info: dict):
        """Show the update available dialog."""
        UpdateAvailableDialog(self, update_info)

    def _create_widgets(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(toolbar, text="+ Add", width=70,
                      command=self._add_variable).pack(side="left", padx=(0, 5))
        ctk.CTkButton(toolbar, text="From Excel", width=85, fg_color="#217346",
                      command=self._import_from_excel).pack(side="left", padx=(0, 5))
        ctk.CTkButton(toolbar, text="Edit", width=60,
                      command=self._edit_variable).pack(side="left", padx=(0, 5))
        ctk.CTkButton(toolbar, text="Delete", width=70, fg_color="darkred",
                      command=self._delete_variable).pack(side="left", padx=(0, 5))

        ctk.CTkFrame(toolbar, width=2, height=30, fg_color="gray").pack(side="left", padx=10)

        ctk.CTkButton(toolbar, text="Sync Excel", width=90, fg_color="#217346",
                      command=self._sync_excel).pack(side="left", padx=(0, 5))

        ctk.CTkFrame(toolbar, width=2, height=30, fg_color="gray").pack(side="left", padx=10)

        word_state = "normal" if self.word else "disabled"

        ctk.CTkButton(toolbar, text="Update Open", width=90,
                      command=self._update_document, state=word_state).pack(side="left", padx=(0, 5))
        ctk.CTkButton(toolbar, text="Update All", width=90,
                      command=self._update_all_files).pack(side="left", padx=(0, 5))
        ctk.CTkButton(toolbar, text="Scan", width=70,
                      command=self._scan_document, state=word_state).pack(side="left", padx=(0, 5))

        content = ctk.CTkFrame(self)
        content.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)

        list_frame = ctk.CTkFrame(content)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(list_frame, text="Variables", font=("", 16, "bold")).grid(
            row=0, column=0, sticky="w", padx=15, pady=10)

        search_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, sticky="e", padx=15, pady=10)

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._refresh_variable_list())
        ctk.CTkEntry(search_frame, placeholder_text="Search...", width=150,
                     textvariable=self.search_var).pack(side="right")

        self.var_scroll = ctk.CTkScrollableFrame(list_frame)
        self.var_scroll.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.var_scroll.grid_columnconfigure(0, weight=1)

        # Status bar with BETA badge and feedback button
        status_frame = ctk.CTkFrame(self, fg_color="transparent")
        status_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 10))
        status_frame.grid_columnconfigure(0, weight=1)

        self.status_var = ctk.StringVar(value="Ready")
        status_bar = ctk.CTkLabel(status_frame, textvariable=self.status_var,
                                   anchor="w", text_color="gray")
        status_bar.grid(row=0, column=0, sticky="w")

        # Beta badge and feedback button (right side)
        beta_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        beta_frame.grid(row=0, column=1, sticky="e")

        beta_label = ctk.CTkLabel(
            beta_frame,
            text="BETA",
            font=("", 11, "bold"),
            text_color="#c026d3",
            fg_color="#3a3a3a",
            corner_radius=4,
            padx=6,
            pady=2
        )
        beta_label.pack(side="left", padx=(0, 8))

        feedback_btn = ctk.CTkButton(
            beta_frame,
            text="Feedback",
            width=70,
            height=24,
            font=("", 11),
            fg_color="#555555",
            hover_color="#666666",
            command=self._show_feedback_dialog
        )
        feedback_btn.pack(side="left")

    def _show_feedback_dialog(self):
        """Show the feedback dialog."""
        FeedbackDialog(self)

    def _refresh_variable_list(self):
        """Refresh the list of variables displayed."""
        for widget in self.var_scroll.winfo_children():
            widget.destroy()

        variables = self.db.get_all_variables()

        search = self.search_var.get().lower()
        if search:
            variables = [v for v in variables if search in v['name'].lower()
                        or search in v.get('value', '').lower()
                        or search in v.get('description', '').lower()]

        self.var_widgets = {}
        for var in variables:
            frame = ctk.CTkFrame(self.var_scroll)
            frame.pack(fill="x", pady=2)
            frame.grid_columnconfigure(1, weight=1)

            check_var = ctk.BooleanVar()
            check = ctk.CTkCheckBox(frame, text="", variable=check_var, width=20)
            check.grid(row=0, column=0, rowspan=2, padx=(10, 5), pady=10)

            # Show Excel link indicator if linked
            name_text = var['name']
            if var.get('excel_file'):
                name_text += "  [Excel]"
            name_label = ctk.CTkLabel(frame, text=name_text, font=("", 13, "bold"), anchor="w")
            name_label.grid(row=0, column=1, sticky="w", padx=5, pady=(10, 0))

            value_text = var['value']
            if var.get('unit'):
                value_text += f" {var['unit']}"
            value_label = ctk.CTkLabel(frame, text=value_text, text_color="gray", anchor="w")
            value_label.grid(row=1, column=1, sticky="w", padx=5, pady=(0, 10))

            usage_btn = ctk.CTkButton(frame, text="Usage", width=60, height=25,
                                       command=lambda v=var: self._show_usage(v))
            usage_btn.grid(row=0, column=2, rowspan=2, padx=10, pady=10)

            self.var_widgets[var['id']] = {
                'frame': frame,
                'check_var': check_var,
                'variable': var
            }

        self.status_var.set(f"{len(variables)} variable(s)")

    def _get_selected_variable(self) -> Optional[dict]:
        """Get the currently selected variable."""
        for data in self.var_widgets.values():
            if data['check_var'].get():
                return data['variable']
        return None

    def _add_variable(self):
        dialog = VariableDialog(self, "Add Variable")
        self.wait_window(dialog)

        if dialog.result:
            try:
                self.db.add_variable(**dialog.result)
                self._refresh_variable_list()
                self.status_var.set(f"Added variable: {dialog.result['name']}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not add variable: {e}")

    def _register_excel_file_with_guid(self, file_path: str) -> Optional[int]:
        """Register an Excel file with GUID tracking. Returns the excel_file_id."""
        try:
            # Get or create GUID in the Excel file
            guid = get_or_create_excel_guid(file_path)
            if not guid:
                logging.warning(f"Could not create GUID for {file_path}")
                return None

            # Register in database
            name = os.path.basename(file_path)
            excel_file_id = self.db.register_excel_file(guid, name, file_path)
            return excel_file_id
        except Exception as e:
            logging.warning(f"Could not register Excel file: {e}")
            return None

    def _import_from_excel(self):
        """Import variables from an Excel range."""
        dialog = ImportRangeDialog(self)
        self.wait_window(dialog)

        # Handle saving the range
        if dialog.save_result:
            save_data = dialog.save_result
            try:
                # Register the Excel file with GUID tracking
                excel_file_id = self._register_excel_file_with_guid(save_data['file_path'])

                # Save the range configuration
                self.db.add_excel_range(
                    name=save_data['name'],
                    file_path=save_data['file_path'],
                    sheet_name=save_data['sheet_name'],
                    start_cell=save_data['start_cell']
                )

                # Also import the variables
                added, updated, errors = self._do_import_variables(
                    save_data['variables'],
                    excel_file_id=excel_file_id,
                    excel_file_path=save_data['file_path'],
                    excel_sheet=save_data['sheet_name']
                )

                self._refresh_variable_list()

                msg = [f"Saved range '{save_data['name']}'"]
                if added:
                    msg.append(f"Added {added} variable(s)")
                if updated:
                    msg.append(f"Updated {updated} variable(s)")

                self.status_var.set(", ".join(msg))
                messagebox.showinfo("Range Saved", "\n".join(msg) + "\n\nUse 'Sync Excel' to refresh from this range later.")

            except Exception as e:
                messagebox.showerror("Error", f"Could not save range: {e}")
            return

        # Handle regular import
        if dialog.result:
            added, updated, errors = self._do_import_variables(dialog.result)

            self._refresh_variable_list()

            # Show result
            msg = []
            if added:
                msg.append(f"Added {added} new variable(s)")
            if updated:
                msg.append(f"Updated {updated} existing variable(s)")
            if errors:
                msg.append(f"{len(errors)} error(s)")

            self.status_var.set(", ".join(msg))

            if errors:
                messagebox.showwarning("Import Complete",
                    "\n".join(msg) + "\n\nErrors:\n" + "\n".join(errors[:5]))
            else:
                messagebox.showinfo("Import Complete", "\n".join(msg))

    def _do_import_variables(self, variables: list[dict], excel_file_id: int = None,
                              excel_file_path: str = None, excel_sheet: str = None) -> tuple[int, int, list]:
        """Import a list of variables. Returns (added, updated, errors).

        If excel_file_id is provided, links variables to that Excel file for GUID tracking.
        """
        added = 0
        updated = 0
        errors = []

        for var_data in variables:
            try:
                # Check if variable already exists
                existing = self.db.get_variable_by_name(var_data['name'])
                if existing:
                    # Update existing variable
                    self.db.update_variable(existing['id'],
                                            value=var_data['value'],
                                            unit=var_data.get('unit', ''))
                    # Link to Excel file if provided
                    if excel_file_id:
                        self.db.link_variable_to_excel_file(existing['id'], excel_file_id)
                    updated += 1
                else:
                    # Add new variable
                    var_id = self.db.add_variable(
                        name=var_data['name'],
                        value=var_data['value'],
                        unit=var_data.get('unit', ''),
                        description=''
                    )
                    # Link to Excel file if provided
                    if excel_file_id:
                        self.db.link_variable_to_excel_file(var_id, excel_file_id)
                    added += 1
            except Exception as e:
                errors.append(f"{var_data['name']}: {e}")

        return added, updated, errors

    def _edit_variable(self):
        var = self._get_selected_variable()
        if not var:
            messagebox.showwarning("No Selection", "Please select a variable to edit.")
            return

        dialog = VariableDialog(self, "Edit Variable", var)
        self.wait_window(dialog)

        if dialog.result:
            try:
                self.db.update_variable(var['id'], **dialog.result)
                self._refresh_variable_list()
                self.status_var.set(f"Updated variable: {dialog.result['name']}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not update variable: {e}")

    def _delete_variable(self):
        var = self._get_selected_variable()
        if not var:
            messagebox.showwarning("No Selection", "Please select a variable to delete.")
            return

        if messagebox.askyesno("Confirm Delete",
                               f"Delete variable '{var['name']}'?\n\n"
                               "This will not remove it from documents where it's already inserted."):
            self.db.delete_variable(var['id'])
            self._refresh_variable_list()
            self.status_var.set(f"Deleted variable: {var['name']}")

    def _link_excel(self):
        """Link selected variable to an Excel cell."""
        var = self._get_selected_variable()
        if not var:
            messagebox.showwarning("No Selection", "Please select a variable to link.")
            return

        dialog = LinkExcelDialog(self, var)
        self.wait_window(dialog)

        if dialog.result is not None:
            try:
                self.db.update_variable(
                    var['id'],
                    excel_file=dialog.result['excel_file'],
                    excel_sheet=dialog.result['excel_sheet'],
                    excel_cell=dialog.result['excel_cell']
                )
                self._refresh_variable_list()
                if dialog.result['excel_file']:
                    self.status_var.set(f"Linked {var['name']} to Excel")
                else:
                    self.status_var.set(f"Removed Excel link for {var['name']}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not update link: {e}")

    def _resolve_excel_file(self, stored_path: str, excel_file_id: int = None) -> Optional[str]:
        """
        Resolve an Excel file path, prompting user to locate if missing.
        Returns the valid path or None if user cancels.
        """
        from tkinter import filedialog
        from excel_reader import get_excel_guid, get_or_create_excel_guid

        # Check if file exists at stored path
        if os.path.exists(stored_path):
            return stored_path

        # File not found - prompt user to locate it
        filename = os.path.basename(stored_path)
        result = messagebox.askyesno(
            "Excel File Not Found",
            f"Cannot find Excel file:\n{filename}\n\nWould you like to locate it?"
        )

        if not result:
            return None

        # Get expected GUID if we have a file_id
        expected_guid = None
        if excel_file_id:
            excel_file = self.db.get_excel_file_by_id(excel_file_id)
            if excel_file:
                expected_guid = excel_file.get('guid')

        # Show file picker
        new_path = filedialog.askopenfilename(
            title=f"Locate {filename}",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialfile=filename
        )

        if not new_path:
            return None

        # Verify GUID if we have one
        if expected_guid:
            actual_guid = get_excel_guid(new_path)
            if actual_guid and actual_guid != expected_guid:
                # GUIDs don't match - this is a different file
                if not messagebox.askyesno(
                    "Different File",
                    f"The selected file appears to be different from the original.\n\n"
                    f"Use this file anyway?"
                ):
                    return None

            # Update the stored path
            self.db.update_excel_file_path(expected_guid, new_path, os.path.basename(new_path))

        return new_path

    def _sync_excel(self):
        """Sync all variables with Excel links and saved ranges."""
        linked_vars = self.db.get_variables_with_excel_links()
        saved_ranges = self.db.get_all_excel_ranges()

        if not linked_vars and not saved_ranges:
            messagebox.showinfo("No Links", "No Excel links or saved ranges found.\n\nUse 'From Excel' to import and save a range, or\nselect a variable and click 'Link' to connect it to an Excel cell.")
            return

        all_changes = {}
        range_changes = []
        skipped_files = set()  # Track files user chose to skip

        # Get changes from individual cell links
        if linked_vars:
            # Group variables by Excel file to avoid multiple prompts for same file
            file_paths = {}
            for var in linked_vars:
                path = var.get('excel_file')
                if path:
                    if path not in file_paths:
                        file_paths[path] = []
                    file_paths[path].append(var)

            # Resolve each file path
            resolved_paths = {}
            for path in file_paths:
                if path in skipped_files:
                    continue
                resolved = self._resolve_excel_file(path, file_paths[path][0].get('excel_file_id'))
                if resolved:
                    resolved_paths[path] = resolved
                else:
                    skipped_files.add(path)

            # Update variables with resolved paths for syncing
            vars_to_sync = []
            for var in linked_vars:
                path = var.get('excel_file')
                if path in resolved_paths:
                    var_copy = dict(var)
                    var_copy['excel_file'] = resolved_paths[path]
                    vars_to_sync.append(var_copy)

            if vars_to_sync:
                changes = sync_variables_from_excel(vars_to_sync)
                for var_id, (old_val, new_val) in changes.items():
                    var = next((v for v in linked_vars if v['id'] == var_id), None)
                    if var:
                        all_changes[var_id] = (var['name'], old_val, new_val)

        # Get changes from saved ranges
        for saved_range in saved_ranges:
            file_path = saved_range['file_path']

            # Skip if user already chose to skip this file
            if file_path in skipped_files:
                continue

            # Resolve file path
            resolved_path = self._resolve_excel_file(file_path, saved_range.get('excel_file_id'))
            if not resolved_path:
                skipped_files.add(file_path)
                continue

            try:
                is_valid, message, variables = validate_excel_range(
                    resolved_path,
                    saved_range['sheet_name'],
                    saved_range['start_cell']
                )
                if is_valid:
                    for var_data in variables:
                        existing = self.db.get_variable_by_name(var_data['name'])
                        if existing:
                            old_val = existing.get('value', '')
                            new_val = var_data['value']
                            if old_val != new_val:
                                range_changes.append({
                                    'var_id': existing['id'],
                                    'name': var_data['name'],
                                    'old_val': old_val,
                                    'new_val': new_val,
                                    'range_name': saved_range['name']
                                })
            except Exception as e:
                logging.warning(f"Error syncing range '{saved_range['name']}': {e}")

        # Combine all changes
        total_changes = len(all_changes) + len(range_changes)

        if total_changes == 0:
            checked = len(linked_vars) + sum(1 for _ in saved_ranges)
            self.status_var.set("All Excel-linked variables are up to date")
            msg = f"Checked {len(linked_vars)} linked variable(s)"
            if saved_ranges:
                msg += f" and {len(saved_ranges)} saved range(s)"
            msg += ".\n\nAll values match Excel."
            messagebox.showinfo("Up to Date", msg)
            return

        # Show confirmation
        msg = f"Found {total_changes} value(s) to update:\n\n"

        # Show cell-linked changes
        shown = 0
        for var_id, (name, old_val, new_val) in list(all_changes.items())[:3]:
            msg += f"  {name}: {old_val} -> {new_val}\n"
            shown += 1

        # Show range changes
        for rc in range_changes[:max(0, 5 - shown)]:
            msg += f"  {rc['name']}: {rc['old_val']} -> {rc['new_val']} (from {rc['range_name']})\n"
            shown += 1

        if total_changes > 5:
            msg += f"  ... and {total_changes - 5} more\n"
        msg += "\nUpdate these values?"

        if not messagebox.askyesno("Confirm Sync", msg):
            return

        # Apply updates
        updated = 0

        # Update cell-linked variables
        for var_id, (name, old_val, new_val) in all_changes.items():
            try:
                self.db.update_variable(var_id, value=new_val)
                updated += 1
            except Exception as e:
                logging.warning(f"Error updating variable {var_id}: {e}")

        # Update range variables
        for rc in range_changes:
            try:
                self.db.update_variable(rc['var_id'], value=rc['new_val'])
                updated += 1
            except Exception as e:
                logging.warning(f"Error updating variable {rc['name']}: {e}")

        # Update last_synced for ranges
        for saved_range in saved_ranges:
            self.db.update_excel_range_synced(saved_range['id'])

        self._refresh_variable_list()
        self.status_var.set(f"Synced {updated} variable(s) from Excel")
        messagebox.showinfo("Sync Complete", f"Updated {updated} variable(s) from Excel.")

    def _show_usage(self, variable: dict):
        documents = self.db.get_variable_usage(variable['id'])

        # Auto-cleanup: remove documents that no longer exist on disk
        valid_documents = []
        for doc in documents:
            path = doc.get('path', '')
            # Skip unsaved documents check - they may still be open
            if path.startswith('unsaved:'):
                valid_documents.append(doc)
            else:
                # Convert Mac path format (Macintosh HD:Users:...) to POSIX
                if path.startswith('Macintosh HD:'):
                    posix_path = '/' + path.replace('Macintosh HD:', '').replace(':', '/')
                else:
                    posix_path = path

                import os
                if os.path.exists(posix_path):
                    valid_documents.append(doc)
                else:
                    # File doesn't exist, remove from database
                    self.db.delete_document(doc['id'])

        UsageDialog(self, variable['name'], valid_documents)

    def _scan_document(self):
        if not self.word:
            messagebox.showerror("Error", "Word integration not available")
            return

        doc = self.word.get_active_document()
        if not doc:
            messagebox.showwarning("No Document", "Please open a Word document first, then click Scan Document.")
            return

        try:
            doc_info = self.word.scan_document()

            doc_id = self.db.register_document(
                guid=doc_info.guid,
                name=doc_info.name,
                path=doc_info.path,
                doc_type="word"
            )

            self.db.clear_usage_for_document(doc_id)

            found_count = 0
            for var_name in doc_info.variables:
                var = self.db.get_variable_by_name(var_name)
                if var:
                    self.db.record_usage(var['id'], doc_id)
                    found_count += 1

            self.db.update_document_scanned(doc_id)

            self.status_var.set(
                f"Scanned '{doc_info.name}': found {len(doc_info.variables)} variable(s), "
                f"{found_count} tracked"
            )

            untracked = [v for v in doc_info.variables
                        if not self.db.get_variable_by_name(v)]
            if untracked:
                messagebox.showinfo("Scan Complete",
                    f"Found {len(doc_info.variables)} variable(s).\n\n"
                    f"Untracked variables:\n" + "\n".join(f"  - {v}" for v in untracked))
            else:
                messagebox.showinfo("Scan Complete",
                    f"Found {len(doc_info.variables)} variable(s), all tracked.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to scan document: {e}")

    def _update_document(self):
        if not self.word:
            messagebox.showerror("Error", "Word integration not available")
            return

        doc = self.word.get_active_document()
        if not doc:
            messagebox.showwarning("No Document", "Please open a Word document first, then click Update Document.")
            return

        try:
            doc_info = self.word.scan_document()
            doc_id = self.db.register_document(
                guid=doc_info.guid,
                name=doc_info.name,
                path=doc_info.path,
                doc_type="word"
            )

            # Build values dict - respect the with_unit flag from when variable was inserted
            all_vars = self.db.get_all_variables()
            db_values = {}
            for v in all_vars:
                var_name = v['name']
                # Check if this variable was inserted with unit
                with_unit = self.db.get_usage_with_unit(var_name, doc_info.guid)

                if with_unit and v.get('unit'):
                    db_values[var_name] = f"{v['value']} {v['unit']}"
                else:
                    db_values[var_name] = v['value']

            # Update usage records (preserve existing with_unit flags)
            for var_name in doc_info.variables:
                var = self.db.get_variable_by_name(var_name)
                if var:
                    # Keep existing with_unit flag
                    existing_with_unit = self.db.get_usage_with_unit(var_name, doc_info.guid)
                    self.db.record_usage(var['id'], doc_id, with_unit=existing_with_unit or False)

            stale = self.word.get_stale_variables(db_values)

            if not stale:
                self.status_var.set("Document is up to date")
                messagebox.showinfo("Up to Date", "All variables are current.")
                return

            msg = "The following variables will be updated:\n\n"
            for name, (old, new) in stale.items():
                msg += f"  - {name}: {old} -> {new}\n"

            if messagebox.askyesno("Confirm Update", msg):
                updated = self.word.update_variables(db_values)
                self.status_var.set(f"Updated {len(updated)} variable(s)")
                messagebox.showinfo("Updated", f"Updated {len(updated)} variable(s)")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update document: {e}")

    def _update_all_files(self):
        """Update all tracked .docx files using direct XML manipulation."""
        import os

        # Get all tracked documents
        documents = self.db.get_all_documents()

        if not documents:
            messagebox.showinfo("No Documents", "No tracked documents found.\n\nUse 'Scan Document' to track documents.")
            return

        # Filter to only .docx files and convert paths
        docx_files = []
        for doc in documents:
            path = doc.get('path', '')
            if not path or path.startswith('unsaved:'):
                continue

            # Convert Mac path format to POSIX
            if path.startswith('Macintosh HD:'):
                posix_path = '/' + path.replace('Macintosh HD:', '').replace(':', '/')
            else:
                posix_path = path

            if posix_path.lower().endswith('.docx') and os.path.exists(posix_path):
                docx_files.append({
                    'doc': doc,
                    'path': posix_path
                })

        if not docx_files:
            messagebox.showinfo("No Files", "No .docx files found that exist on disk.\n\nFiles may have been moved or deleted.")
            return

        # Build variable values dict
        all_vars = self.db.get_all_variables()

        # Track what will be updated
        files_to_update = []
        for file_info in docx_files:
            posix_path = file_info['path']
            doc = file_info['doc']

            try:
                # Get current values in the file
                current_vars = get_docx_variables(posix_path)

                # Build new values respecting with_unit flags
                new_values = {}
                changes = []

                for v in all_vars:
                    var_name = v['name']
                    if var_name not in current_vars:
                        continue  # Variable not in this document

                    # Check if variable was inserted with unit
                    with_unit = self.db.get_usage_with_unit(var_name, doc.get('guid', ''))

                    if with_unit and v.get('unit'):
                        new_value = f"{v['value']} {v['unit']}"
                    else:
                        new_value = v['value']

                    new_values[var_name] = new_value

                    # Check if it's different
                    old_value = current_vars.get(var_name, '')
                    if old_value != new_value:
                        changes.append((var_name, old_value, new_value))

                if changes:
                    files_to_update.append({
                        'path': posix_path,
                        'name': doc.get('name', os.path.basename(posix_path)),
                        'values': new_values,
                        'changes': changes
                    })

            except Exception as e:
                logging.warning(f"Error checking {posix_path}: {e}")

        if not files_to_update:
            self.status_var.set("All files are up to date")
            messagebox.showinfo("Up to Date", f"Checked {len(docx_files)} file(s).\n\nAll variables are current.")
            return

        # Show confirmation
        msg = f"Found {len(files_to_update)} file(s) with changes:\n\n"
        for f in files_to_update[:5]:  # Show first 5
            msg += f"• {f['name']}: {len(f['changes'])} change(s)\n"
        if len(files_to_update) > 5:
            msg += f"... and {len(files_to_update) - 5} more\n"
        msg += "\nUpdate all files? (Backups will be created)"

        if not messagebox.askyesno("Confirm Update All", msg):
            return

        # Perform updates
        updated_count = 0
        errors = []

        for f in files_to_update:
            try:
                update_docx_variables(f['path'], f['values'], backup=True)
                updated_count += 1
            except Exception as e:
                errors.append(f"{f['name']}: {e}")

        # Report results
        if errors:
            self.status_var.set(f"Updated {updated_count} file(s), {len(errors)} error(s)")
            error_msg = f"Updated {updated_count} file(s).\n\nErrors:\n" + "\n".join(errors[:5])
            messagebox.showwarning("Partial Success", error_msg)
        else:
            self.status_var.set(f"Updated {updated_count} file(s)")
            messagebox.showinfo("Update Complete", f"Successfully updated {updated_count} file(s).\n\nBackup files (.bak) were created.")

    def _start_hotkey_listener(self):
        """Start the global hotkey listener for Alt+Space (Option+Space on Mac)."""
        self._event_monitor = None
        self._listener = None

        if platform.system() == "Darwin":
            self._start_hotkey_listener_mac()
        else:
            self._start_hotkey_listener_pynput()

    def _start_hotkey_listener_mac(self):
        """macOS-specific hotkey - delay start to avoid crash in bundled app."""
        # Check/request Accessibility permission first
        if not self._check_accessibility_permission():
            logging.warning("Accessibility permission not granted - hotkey disabled")
            return

        # Delay the actual listener start until after GUI is fully initialized
        # This avoids the crash that happens when pynput accesses keyboard APIs too early
        self.after(2000, self._setup_mac_hotkey_delayed)

    def _setup_mac_hotkey_delayed(self):
        """Set up macOS hotkey using Quartz CGEventTap directly."""
        import threading

        try:
            from Quartz import (
                CGEventTapCreate, CGEventTapEnable, CFMachPortCreateRunLoopSource,
                CFRunLoopGetCurrent, CFRunLoopAddSource, CFRunLoopRun,
                kCGSessionEventTap, kCGHeadInsertEventTap, kCGEventTapOptionListenOnly,
                CGEventMaskBit, kCGEventKeyDown,
                CGEventGetIntegerValueField, CGEventGetFlags,
                kCGKeyboardEventKeycode, kCFRunLoopDefaultMode
            )

            app_ref = self

            def hotkey_callback(proxy, event_type, event, refcon):
                try:
                    keycode = CGEventGetIntegerValueField(event, kCGKeyboardEventKeycode)
                    flags = CGEventGetFlags(event)
                    # Option/Alt flag = 0x80000, Space keycode = 49
                    if keycode == 49 and (flags & 0x80000):
                        app_ref.after(0, app_ref._show_quick_insert)
                except Exception:
                    pass
                return event

            def run_event_tap():
                tap = CGEventTapCreate(
                    kCGSessionEventTap,
                    kCGHeadInsertEventTap,
                    kCGEventTapOptionListenOnly,
                    CGEventMaskBit(kCGEventKeyDown),
                    hotkey_callback,
                    None
                )
                if tap is None:
                    logging.warning("Failed to create event tap - Input Monitoring permission may be needed")
                    return

                run_loop_source = CFMachPortCreateRunLoopSource(None, tap, 0)
                CFRunLoopAddSource(CFRunLoopGetCurrent(), run_loop_source, kCFRunLoopDefaultMode)
                CGEventTapEnable(tap, True)
                logging.info("Global hotkey (Option+Space) registered")
                CFRunLoopRun()

            self._hotkey_thread = threading.Thread(target=run_event_tap, daemon=True)
            self._hotkey_thread.start()

            # Show Input Monitoring instructions on first launch
            self._show_input_monitoring_instructions_once()

        except Exception as e:
            logging.warning(f"Could not start macOS hotkey: {e}")
            self._show_input_monitoring_instructions()

    def _show_input_monitoring_instructions_once(self):
        """Show Input Monitoring instructions on first launch only."""
        import os

        # Use a marker file in user's home directory
        marker_file = os.path.expanduser('~/.tansu_input_monitoring_shown')

        if os.path.exists(marker_file):
            return

        self._show_input_monitoring_instructions()

        # Create marker file so we don't show again
        try:
            with open(marker_file, 'w') as f:
                f.write('1')
        except Exception:
            pass

    def _show_input_monitoring_instructions(self):
        """Show instructions for enabling Input Monitoring permission."""
        # Bring the app to the front
        self.lift()
        self.focus_force()

        msg = (
            "For the Option+Space hotkey to work, Tansu needs permissions.\n\n"
            "To enable:\n"
            "1. Open System Settings > Privacy & Security\n"
            "2. Add Tansu to both Accessibility AND Input Monitoring\n"
            "3. Restart Tansu\n\n"
            "Alternatively, use the 'Update Open' button to insert variables."
        )
        messagebox.showinfo("Hotkey Setup", msg, parent=self)

    def _check_accessibility_permission(self) -> bool:
        """Check if Accessibility permission is granted, prompt if not."""
        try:
            import objc
            from ApplicationServices import AXIsProcessTrustedWithOptions
            from Foundation import NSDictionary

            # kAXTrustedCheckOptionPrompt = True will show the system dialog
            options = NSDictionary.dictionaryWithObject_forKey_(
                True,  # Prompt user if not trusted
                "AXTrustedCheckOptionPrompt"
            )
            trusted = AXIsProcessTrustedWithOptions(options)
            return trusted
        except Exception as e:
            logging.warning(f"Could not check accessibility permission: {e}")
            # Fall back to assuming it's granted
            return True

    def _start_hotkey_listener_pynput(self):
        """Windows/Linux hotkey using pynput."""
        try:
            from pynput import keyboard

            def on_hotkey():
                # Schedule on main thread
                self.after(0, self._show_quick_insert)

            # Alt+Space hotkey
            hotkey = keyboard.HotKey(
                keyboard.HotKey.parse('<alt>+<space>'),
                on_hotkey
            )

            def on_press(key):
                hotkey.press(self._listener.canonical(key))

            def on_release(key):
                hotkey.release(self._listener.canonical(key))

            self._listener = keyboard.Listener(on_press=on_press, on_release=on_release)
            self._listener.start()

            logging.info("Global hotkey (Alt+Space) registered via pynput")

        except Exception as e:
            logging.warning(f"Could not start hotkey listener: {e}")

    def _stop_hotkey_listener(self):
        """Stop the global hotkey listener and clean up."""
        # macOS uses thread with CFRunLoop, Windows uses pynput listener
        if platform.system() == "Darwin":
            # The daemon thread will be stopped automatically when app exits
            pass
        elif self._listener:
            try:
                self._listener.stop()
                self._listener = None
            except Exception as e:
                logging.warning(f"Error stopping listener: {e}")

    def _show_quick_insert(self):
        """Show the quick insert popup."""
        # Don't open multiple popups
        if self._quick_insert_popup and self._quick_insert_popup.winfo_exists():
            self._quick_insert_popup.focus_set()
            return

        variables = self.db.get_all_variables()
        if not variables:
            messagebox.showinfo("No Variables", "Add some variables first.")
            return

        self._quick_insert_popup = QuickInsertPopup(self, variables)

    def destroy(self):
        """Clean up resources before destroying the window."""
        self._stop_hotkey_listener()
        stop_api_server()
        super().destroy()


# -------------------------
# Main Entry Point
# -------------------------

def main():
    """Run the main application."""
    app = VariableTrackerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
