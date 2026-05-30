# Outlook PST/OST Tool

A standalone GUI tool for reading and extracting data from Microsoft Outlook PST and OST
files — **without requiring Microsoft Outlook to be installed**.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## Features

- **Browse** the complete folder hierarchy of PST/OST files
- **View** emails, contacts, calendar items, tasks, and notes — HTML-only emails (common
  in OST files) are rendered as readable text instead of raw markup
- **Check items** with a checkbox column and **Save Checked Items As** `txt`, `csv`,
  `html`, `eml`, or `pdf` — attachments are embedded where the format allows (EML MIME
  parts, HTML data-URIs, PDF embedded files) and otherwise saved to a per-item folder
- **Export a folder + all its subfolders** in any of those formats, recreating the folder
  tree on disk and including every item type
- **Extract attachments** — a single attachment, all from a message, or a whole folder
- **Export Calendar** to ICS (RFC-5545) or CSV with real start/end times and location
- **Export Contacts** to VCF (vCard) or CSV
- **Search** across all items by subject, body, or sender
- **View MAPI Properties** for advanced inspection
- **Sortable columns** and a tabbed detail view

## Download

Grab the latest standalone `.exe` from the [Releases](../../releases) page — no Python
installation required.

## Running from Source

### Prerequisites

- Python 3.10 or later
- The dependencies in `requirements.txt`:
  - `libpff-python` (pypff) — reads PST/OST files
  - `xhtml2pdf` — pure-Python HTML→PDF rendering (no native GTK dependency)
  - `pikepdf` — embeds attachments inside exported PDFs

### Install

Use the setup script for your platform — it creates a local `.venv` and installs the
dependencies:

```bash
# Windows
install.bat

# macOS / Linux
./install.sh
```

Or install manually into your own environment:

```bash
pip install -r requirements.txt
```

### Run

```bash
# Windows
run.bat

# macOS / Linux
./run.sh
```

Or run directly with `python main.py` (Windows) / `python3 main.py` (macOS/Linux).

## Usage

1. **Open** a `.pst` or `.ost` file via `File > Open PST/OST...` or the toolbar
2. **Browse** folders on the left; click a folder to load its items
3. **Click** an item to preview it (with its attachments); double-click for full detail
4. **Tick the checkboxes** of the items you want, then **Save Checked** and pick a format
   (`txt` / `csv` / `html` / `eml` / `pdf`) — use *Check All* / *Clear Checks* as needed
5. **Export Folder** to export the selected folder and its subfolders in a chosen format
6. **Export** calendar/contacts/emails in bulk from the **Export** menu
7. **Search** with `Ctrl+F`

## Testing

The test suite lives in `tests/`:

```bash
# Fast unit tests (no PST/OST file required)
pytest tests/test_units.py

# End-to-end verification against real files in temp/ (PST + OST)
python tests/verify_exports.py

# Headless GUI smoke test (needs a virtual display on Linux)
xvfb-run -a python tests/smoke_gui.py
```

`verify_exports.py` exercises the property extraction, all five per-item export formats,
recursive folder export, and ICS validity against the actual data, and exits non-zero on
any failure.

## Linting

The project is configured for strict linting with [ruff](https://docs.astral.sh/ruff/)
(see `pyproject.toml`):

```bash
ruff check .
```

## Building the Executable

`xhtml2pdf` and `pikepdf` ship as standard wheels (no native GTK runtime needed), so the
standalone build stays a single file. PyInstaller produces a native binary for whichever
OS you run it on (`.exe` on Windows, a Unix executable on macOS/Linux):

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
| `pst_handler.py` | PST/OST reader using pypff (MAPI property + named-property extraction) |
| `exporters.py` | Per-item renderers and export engines (txt/csv/html/eml/pdf, ICS, VCF) |
| `requirements.txt` | Python dependencies |
| `pyproject.toml` | Project metadata + ruff/pytest configuration |
| `tests/` | Unit tests, end-to-end verification, and GUI smoke test |
| `install.bat` / `run.bat` | Windows setup/launch scripts |
| `install.sh` / `run.sh` | macOS/Linux setup/launch scripts |

## License

This project is licensed under the [MIT License](LICENSE).

## Author

**Randy Northrup**
