# Tansu - Variable Tracker for Word Documents

Tansu is a desktop application that helps you manage variables across Microsoft Word documents. Import values from Excel, track them in a central database, and keep all your Word documents in sync.

![License: CC BY-NC 4.0](https://img.shields.io/badge/License-CC%20BY--NC%204.0-lightgrey.svg)

## Features

- **Variable Management** - Create, edit, and organize variables with names, values, and units
- **Excel Integration** - Import variables directly from Excel spreadsheets, with live sync support
- **Word DOCVARIABLE Fields** - Insert variables as updatable fields in Word documents
- **Batch Updates** - Update all linked Word documents with a single click
- **Cross-Platform** - Works on macOS and Windows
- **Global Hotkey** - Press Option+Space (Mac) or Alt+Space (Windows) to quickly insert variables

## Installation

### Prerequisites

- Python 3.8 or higher
- Microsoft Word (for document integration)
- Microsoft Excel (optional, for Excel import features)

### Quick Start

1. Clone the repository:
   ```bash
   git clone https://github.com/gtrad/Tansu-MVP.git
   cd Tansu-MVP
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the app:
   ```bash
   python app.py
   ```

### Building Standalone Apps

#### macOS
```bash
./build_mac.sh
# Creates: dist/Tansu.app
```

#### Windows
```bash
build_windows.bat
# Creates: dist\Tansu\Tansu.exe
```

## Usage

### Adding Variables

1. Click **+ Add** to create a new variable
2. Enter a name (e.g., `project_budget`), value, and optional unit
3. The variable is now available for insertion into Word documents

### Importing from Excel

1. Click **From Excel** to open the Excel range picker
2. Select your Excel file and sheet
3. Click and drag to select a range of cells
4. Click **Import** to add variables, or **Save Range** to save for future syncing

### Inserting into Word

**Option 1: From the main app**
1. Open a Word document
2. Place your cursor where you want the variable
3. In Tansu, select a variable and click **Update Open**

**Option 2: Global Hotkey (Recommended)**
Press **Option+Space** (Mac) or **Alt+Space** (Windows) while in Word to open the quick insert popup.

> **macOS Note:** The first time you use the hotkey, you may need to grant Tansu **Input Monitoring** permission:
> 1. Open **System Settings > Privacy & Security > Input Monitoring**
> 2. Click the **+** button and add Tansu from your Applications folder
> 3. Restart Tansu for the hotkey to work

**Option 3: Word Ribbon Button (Windows)**
See `word_addin/INSTALL_WORD_ADDIN.txt` for instructions on adding an "Insert Variable" button directly to Word's ribbon.

### Updating Documents

- **Update Open** - Updates the currently active Word document
- **Update All** - Updates all .docx files that contain your variables
- **Sync Excel** - Refreshes variable values from linked Excel files

## How It Works

Tansu uses Microsoft Word's **DOCVARIABLE** fields to create updatable placeholders in your documents. When you insert a variable:

1. A document variable is created in Word with your value
2. A DOCVARIABLE field is inserted at the cursor position
3. The field displays the current value but can be updated later

When values change, simply click "Update" in Tansu to refresh all fields across your documents.

### Document Tracking

Each document gets a unique GUID stored as a custom document property. This GUID persists if the file is renamed, moved, or copied, allowing consistent tracking.

## Project Structure

```
tansu/
├── app.py              # Main GUI application
├── database.py         # SQLite database management
├── word_integration.py # Word integration (cross-platform wrapper)
├── word_mac.py         # macOS Word integration (AppleScript)
├── word_windows.py     # Windows Word integration (pywin32)
├── excel_reader.py     # Excel file reading
├── docx_updater.py     # Batch .docx file updates
├── menubar_app.py      # macOS menu bar app (rumps)
├── tray_app_windows.py # Windows system tray app (pystray)
├── api_server.py       # HTTP API for Word VBA integration
├── launcher.py         # macOS launcher (GUI + menu bar)
└── word_addin/         # VBA code for Word ribbon integration
```

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International License** (CC BY-NC 4.0).

You are free to:
- **Share** - copy and redistribute the material
- **Adapt** - remix, transform, and build upon the material

Under these terms:
- **Attribution** - You must give appropriate credit
- **NonCommercial** - You may not use the material for commercial purposes

See [LICENSE](LICENSE) for details.

For commercial licensing inquiries, please contact the project maintainers.

## Acknowledgments

- Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for the modern UI
- Excel reading powered by [openpyxl](https://openpyxl.readthedocs.io/)
- Word document manipulation via [python-docx](https://python-docx.readthedocs.io/)
