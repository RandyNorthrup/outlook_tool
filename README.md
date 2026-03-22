# Outlook PST/OST Tool

A standalone Windows GUI tool for reading and extracting data from Microsoft Outlook PST and OST files — **without requiring Microsoft Outlook to be installed**.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

## Features

- **Browse** the complete folder hierarchy of PST/OST files
- **View** emails, contacts, calendar items, tasks, and notes
- **Export Calendar** to ICS (iCalendar) or CSV
- **Export Contacts** to VCF (vCard) or CSV
- **Export Emails** to individual EML files or CSV
- **Extract Attachments** from any folder or individual message
- **Search** across all items by subject, body, or sender
- **View MAPI Properties** for advanced inspection
- **Sortable columns** in the item list
- **Detail view** with tabbed content, properties, and attachment management

## Download

Grab the latest standalone `.exe` from the [Releases](../../releases) page — no Python installation required.

## Running from Source

### Prerequisites

- Python 3.10 or later
- `libpff-python` (pypff)

### Install

```bash
pip install libpff-python
```

Or use the included setup script:

```bash
install.bat
```

### Run

```bash
python main.py
```

Or double-click `run.bat`.

## Building the Executable

The GitHub Actions workflow automatically builds a standalone `.exe` on every release tag push. To build locally:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name OutlookTool main.py
```

The output will be in the `dist/` folder.

## Project Structure

| File | Description |
|---|---|
| `main.py` | Entry point |
| `app.py` | GUI application (tkinter) |
| `pst_handler.py` | PST/OST file reader using pypff |
| `exporters.py` | Export functions (ICS, VCF, EML, CSV) |
| `requirements.txt` | Python dependencies |
| `install.bat` | Windows setup script |
| `run.bat` | Windows launch script |

## Usage

1. **Open** a `.pst` or `.ost` file via `File > Open PST/OST...` or the toolbar
2. **Browse** folders in the left panel
3. **Select** a folder to load its items in the item list
4. **Click** an item to preview it; double-click for the full detail view
5. **Export** using the Export menu or toolbar buttons
6. **Search** with `Ctrl+F` or the toolbar Search button

## License

This project is licensed under the [MIT License](LICENSE).

## Author

**Randy Northrup**
