"""Outlook PST/OST Tool - main GUI application.

A single-window tkinter application for reading and extracting data from Outlook
PST and OST files without requiring Microsoft Outlook to be installed.
"""

import contextlib
import os
import platform
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Add the script's directory to path so imports work from any cwd.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import exporters
from pst_handler import (
    PYPFF_AVAILABLE,
    AttachmentInfo,
    FolderInfo,
    MessageItem,
    PstReader,
    html_to_text,
)

# --------------------------------------------------------------------------- #
# Constants (no magic numbers/strings scattered through the UI code)
# --------------------------------------------------------------------------- #

APP_TITLE = "Outlook PST/OST Tool"
APP_AUTHOR = "Randy Northrup"

MAIN_WINDOW_SIZE = "1280x820"
MAIN_WINDOW_MIN = (900, 620)
DETAIL_WINDOW_SIZE = "780x600"
SEARCH_WINDOW_SIZE = "860x560"
STATS_WINDOW_SIZE = "680x420"
ABOUT_WINDOW_SIZE = "400x460"
FORMAT_DIALOG_PAD = 16

SPLASH_WIDTH = 420
SPLASH_HEIGHT = 210
SPLASH_DELAY_MS = 1500
SPLASH_BG = "#2b579a"
SPLASH_FG = "#ffffff"
SPLASH_FG_SUB = "#b0c4de"
SPLASH_FG_HINT = "#d0d8e8"

# Item-list columns.
COL_CHECK = "check"
COL_SUBJECT = "subject"
COL_FROM = "from"
COL_DATE = "date"
COL_TYPE = "type"
COL_ATT = "att"
ITEM_COLUMNS = (COL_CHECK, COL_SUBJECT, COL_FROM, COL_DATE, COL_TYPE, COL_ATT)

CHECK_ON = "☑"   # ballot box with check
CHECK_OFF = "☐"  # empty ballot box

# Column widths.
WIDTH_CHECK = 36
WIDTH_SUBJECT = 320
WIDTH_FROM = 200
WIDTH_DATE = 150
WIDTH_TYPE = 80
WIDTH_ATT = 44
WIDTH_ATT_NAME = 420
WIDTH_ATT_SIZE = 110
MINWIDTH_SUBJECT = 100
MINWIDTH_FROM = 80
MINWIDTH_DATE = 80
MINWIDTH_TYPE = 60
MINWIDTH_SMALL = 30

PANE_WEIGHT_FOLDERS = 1
PANE_WEIGHT_CONTENT = 3
PANE_WEIGHT_LIST = 3
PANE_WEIGHT_PREVIEW = 2
ATTACH_TREE_HEIGHT = 4

# Fonts — pick families that exist on each platform (Tk falls back if missing).
_OS = platform.system()
if _OS == "Windows":
    UI_FONT, MONO_FONT = "Segoe UI", "Consolas"
elif _OS == "Darwin":
    UI_FONT, MONO_FONT = "Helvetica Neue", "Menlo"
else:  # Linux / other
    UI_FONT, MONO_FONT = "DejaVu Sans", "DejaVu Sans Mono"

FONT_PREVIEW_SUBJECT = (UI_FONT, 12, "bold")
FONT_DETAIL_TITLE = (UI_FONT, 13, "bold")
FONT_ABOUT_TITLE = (UI_FONT, 14, "bold")
FONT_SECTION = (UI_FONT, 10, "bold")
FONT_BODY = (MONO_FONT, 10)          # raw / monospaced (MAPI properties view)
FONT_READING = (UI_FONT, 11)         # readable proportional text (message bodies)
FONT_TREE = (UI_FONT, 10)
FONT_TREE_HEADING = (UI_FONT, 10, "bold")
FONT_SPLASH_TITLE = (UI_FONT, 18, "bold")
FONT_SPLASH_SUB = (UI_FONT, 10)

# Colors (subtle, Outlook-inspired accent).
COLOR_META = "#5b6470"
COLOR_ACCENT = "#2b579a"
COLOR_ROW_ODD = "#ffffff"
COLOR_ROW_EVEN = "#f3f6fb"
COLOR_ROW_CHECKED = "#d9e8fb"
COLOR_HEADER_BG = "#2b579a"
COLOR_HEADER_FG = "#ffffff"

# Treeview row presentation.
TREE_ROW_HEIGHT = 24
TAG_ODD = "oddrow"
TAG_EVEN = "evenrow"
TAG_CHECKED = "checkedrow"

# Preferred ttk theme per platform (falls back to whatever is available).
PREFERRED_THEME = {"Windows": "vista", "Darwin": "aqua"}.get(_OS, "clam")

# Folder-type display indicators.
FOLDER_TYPE_INDICATORS = {
    "email": "[Mail]",
    "contacts": "[Contacts]",
    "calendar": "[Calendar]",
    "tasks": "[Tasks]",
    "notes": "[Notes]",
    "other": "",
}

# Export type metadata (item_type -> human label).
TYPE_LABELS = {
    "calendar": "calendar items",
    "contact": "contacts",
    "email": "emails",
}

NO_SUBJECT = "(No Subject)"

# Format chooser options: (value, description).
FORMAT_CHOICES = (
    ("eml", "EML  -  standard e-mail file (attachments embedded as MIME parts)"),
    ("pdf", "PDF  -  printable document (attachments embedded inside the PDF)"),
    ("html", "HTML  -  web page (inline images + attachments embedded)"),
    ("txt", "TXT  -  plain text (attachments saved in a side folder)"),
    ("csv", "CSV  -  one spreadsheet row per item (+ attachments folder)"),
)
DEFAULT_FORMAT = "eml"

INSTALL_HINT = (
    "pypff (libpff-python) is required but not installed.\n\n"
    "Install with one of:\n"
    "  pip install libpff-python\n"
    "  conda install -c conda-forge libpff-python\n\n"
    "See https://github.com/libyal/libpff for details."
)


class OutlookToolApp:
    """Main GUI application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(MAIN_WINDOW_SIZE)
        self.root.minsize(*MAIN_WINDOW_MIN)

        self.reader: PstReader | None = None
        self.current_folder: FolderInfo | None = None
        self.current_messages: list[MessageItem] = []
        self.folder_map: dict[str, FolderInfo] = {}
        self.message_map: dict[str, MessageItem] = {}
        self.checked_items: set = set()
        self.preview_attachments: list[AttachmentInfo] = []
        self.preview_att_map: dict[str, AttachmentInfo] = {}
        self._sort_reverse: dict[str, bool] = {}

        self._setup_style()
        self._build_menu()
        self._build_header()
        self._build_toolbar()
        self._build_main_layout()
        self._build_statusbar()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Control-o>", lambda _e: self._open_file())
        self.root.bind("<Control-f>", lambda _e: self._show_search())
        self.root.bind("<Control-s>", lambda _e: self._save_selected())

        if not PYPFF_AVAILABLE:
            self._set_status(
                "WARNING: pypff not installed. Install with: pip install libpff-python")

    # ======================== Theme / Header ========================

    def _setup_style(self) -> None:
        """Apply a clean, consistent ttk theme and presentation tweaks."""
        style = ttk.Style()
        if PREFERRED_THEME in style.theme_names():
            with contextlib.suppress(tk.TclError):
                style.theme_use(PREFERRED_THEME)
        style.configure("Treeview", rowheight=TREE_ROW_HEIGHT, font=FONT_TREE)
        style.configure("Treeview.Heading", font=FONT_TREE_HEADING, padding=(4, 4))
        style.configure("TButton", padding=(8, 4))

    def _build_header(self) -> None:
        """A slim Outlook-blue title strip showing the open file.

        Uses classic tk widgets so the background color renders on every
        platform (native ttk themes ignore frame/label background colors).
        """
        header = tk.Frame(self.root, bg=COLOR_HEADER_BG)
        header.pack(fill=tk.X)
        tk.Label(header, text=APP_TITLE, bg=COLOR_HEADER_BG, fg=COLOR_HEADER_FG,
                 font=(UI_FONT, 12, "bold")).pack(side=tk.LEFT, padx=10, pady=6)
        self.header_file_label = tk.Label(
            header, text="No file open", bg=COLOR_HEADER_BG, fg=COLOR_HEADER_FG,
            font=(UI_FONT, 10))
        self.header_file_label.pack(side=tk.RIGHT, padx=10)

    # ======================== Menu Bar ========================

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open PST/OST...", command=self._open_file,
                              accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Close File", command=self._close_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)

        export_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Export", menu=export_menu)

        cal_menu = tk.Menu(export_menu, tearoff=0)
        export_menu.add_cascade(label="Calendar", menu=cal_menu)
        cal_menu.add_command(label="Export All to ICS...",
                             command=lambda: self._export_typed("calendar", "ics"))
        cal_menu.add_command(label="Export All to CSV...",
                             command=lambda: self._export_typed("calendar", "csv"))

        con_menu = tk.Menu(export_menu, tearoff=0)
        export_menu.add_cascade(label="Contacts", menu=con_menu)
        con_menu.add_command(label="Export All to VCF...",
                             command=lambda: self._export_typed("contact", "vcf"))
        con_menu.add_command(label="Export All to CSV...",
                             command=lambda: self._export_typed("contact", "csv"))

        email_menu = tk.Menu(export_menu, tearoff=0)
        export_menu.add_cascade(label="Emails", menu=email_menu)
        email_menu.add_command(label="Export All to EML files...",
                               command=lambda: self._export_typed("email", "eml"))
        email_menu.add_command(label="Export All to CSV...",
                               command=lambda: self._export_typed("email", "csv"))

        export_menu.add_separator()
        export_menu.add_command(
            label="Save Checked Items As...   (txt/csv/html/eml/pdf)",
            command=self._save_selected, accelerator="Ctrl+S")
        export_menu.add_command(label="Export Folder + Subfolders As...",
                                command=self._export_folder_tree)
        export_menu.add_separator()
        export_menu.add_command(label="Extract Attachments (Current Folder)...",
                                command=self._export_attachments)
        export_menu.add_separator()
        export_menu.add_command(label="Export Current Folder to CSV...",
                                command=self._export_folder_csv)
        export_menu.add_command(label="Export Entire File to CSV...",
                                command=self._export_all_csv)

        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Search...", command=self._show_search,
                               accelerator="Ctrl+F")
        tools_menu.add_separator()
        tools_menu.add_command(label="File Properties...",
                               command=self._show_file_properties)
        tools_menu.add_command(label="Folder Statistics...",
                               command=self._show_folder_stats)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)

    # ======================== Toolbar ========================

    def _build_toolbar(self) -> None:
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=2, pady=(2, 0))

        def sep() -> None:
            ttk.Separator(toolbar, orient=tk.VERTICAL).pack(
                side=tk.LEFT, fill=tk.Y, padx=4, pady=2)

        ttk.Button(toolbar, text="Open File",
                   command=self._open_file).pack(side=tk.LEFT, padx=2)
        sep()
        ttk.Button(toolbar, text="Export Calendar",
                   command=lambda: self._export_typed("calendar", "ics")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Export Contacts",
                   command=lambda: self._export_typed("contact", "vcf")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Export Emails",
                   command=lambda: self._export_typed("email", "eml")).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Attachments",
                   command=self._export_attachments).pack(side=tk.LEFT, padx=2)
        sep()
        ttk.Button(toolbar, text="Save Checked",
                   command=self._save_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Export Folder",
                   command=self._export_folder_tree).pack(side=tk.LEFT, padx=2)
        sep()
        ttk.Button(toolbar, text="Check All",
                   command=lambda: self._set_all_checks(True)).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Clear Checks",
                   command=lambda: self._set_all_checks(False)).pack(
            side=tk.LEFT, padx=2)
        sep()
        ttk.Button(toolbar, text="Search",
                   command=self._show_search).pack(side=tk.LEFT, padx=2)

    # ======================== Main Layout ========================

    def _build_main_layout(self) -> None:
        self.main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        self._build_folder_pane()
        self._build_content_pane()

    def _build_folder_pane(self) -> None:
        left_frame = ttk.LabelFrame(self.main_pane, text="Folders", padding=4)
        self.main_pane.add(left_frame, weight=PANE_WEIGHT_FOLDERS)

        self.folder_tree = ttk.Treeview(left_frame, show="tree", selectmode="browse")
        folder_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL,
                                       command=self.folder_tree.yview)
        self.folder_tree.configure(yscrollcommand=folder_scroll.set)
        self.folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        folder_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.bind("<<TreeviewSelect>>", self._on_folder_select)

        self.folder_context_menu = tk.Menu(self.folder_tree, tearoff=0)
        self.folder_context_menu.add_command(
            label="Export Folder + Subfolders As...  (txt/csv/html/eml/pdf)",
            command=self._export_folder_tree)
        self.folder_context_menu.add_command(
            label="Export This Folder to CSV...", command=self._export_folder_csv)
        self.folder_tree.bind("<Button-3>", self._show_folder_context_menu)

    def _build_content_pane(self) -> None:
        right_pane = ttk.PanedWindow(self.main_pane, orient=tk.VERTICAL)
        self.main_pane.add(right_pane, weight=PANE_WEIGHT_CONTENT)

        self._build_item_list(right_pane)
        self._build_preview(right_pane)

    def _build_item_list(self, parent: ttk.PanedWindow) -> None:
        list_frame = ttk.LabelFrame(
            parent, text="Items  (click a row to preview; tick the box to select "
                         "for actions)", padding=4)
        parent.add(list_frame, weight=PANE_WEIGHT_LIST)

        self.item_list = ttk.Treeview(list_frame, columns=ITEM_COLUMNS,
                                       show="headings", selectmode="browse")
        self.item_list.heading(COL_CHECK, text=CHECK_OFF,
                                command=self._toggle_all_checks)
        self.item_list.heading(COL_SUBJECT, text="Subject",
                                command=lambda: self._sort_column(COL_SUBJECT))
        self.item_list.heading(COL_FROM, text="From / Name",
                                command=lambda: self._sort_column(COL_FROM))
        self.item_list.heading(COL_DATE, text="Date",
                                command=lambda: self._sort_column(COL_DATE))
        self.item_list.heading(COL_TYPE, text="Type",
                                command=lambda: self._sort_column(COL_TYPE))
        self.item_list.heading(COL_ATT, text="Att",
                                command=lambda: self._sort_column(COL_ATT))

        self.item_list.column(COL_CHECK, width=WIDTH_CHECK, minwidth=WIDTH_CHECK,
                              anchor=tk.CENTER, stretch=False)
        self.item_list.column(COL_SUBJECT, width=WIDTH_SUBJECT,
                              minwidth=MINWIDTH_SUBJECT)
        self.item_list.column(COL_FROM, width=WIDTH_FROM, minwidth=MINWIDTH_FROM)
        self.item_list.column(COL_DATE, width=WIDTH_DATE, minwidth=MINWIDTH_DATE)
        self.item_list.column(COL_TYPE, width=WIDTH_TYPE, minwidth=MINWIDTH_TYPE)
        self.item_list.column(COL_ATT, width=WIDTH_ATT, minwidth=MINWIDTH_SMALL,
                              anchor=tk.CENTER)

        list_ys = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                command=self.item_list.yview)
        list_xs = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL,
                                command=self.item_list.xview)
        self.item_list.configure(yscrollcommand=list_ys.set,
                                 xscrollcommand=list_xs.set)
        self.item_list.grid(row=0, column=0, sticky="nsew")
        list_ys.grid(row=0, column=1, sticky="ns")
        list_xs.grid(row=1, column=0, sticky="ew")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # Zebra striping + a clear highlight for checked rows.
        self.item_list.tag_configure(TAG_ODD, background=COLOR_ROW_ODD)
        self.item_list.tag_configure(TAG_EVEN, background=COLOR_ROW_EVEN)
        self.item_list.tag_configure(TAG_CHECKED, background=COLOR_ROW_CHECKED)

        self.item_list.bind("<<TreeviewSelect>>", self._on_item_select)
        self.item_list.bind("<Button-1>", self._on_item_click, add="+")
        self.item_list.bind("<Double-1>", self._on_item_double_click)
        self.item_list.bind("<space>", self._on_item_space)

        self.item_context_menu = tk.Menu(self.item_list, tearoff=0)
        self.item_context_menu.add_command(
            label="Save Checked Items As...  (txt/csv/html/eml/pdf)",
            command=self._save_selected)
        self.item_context_menu.add_command(label="View Details",
                                            command=self._context_view_details)
        self.item_context_menu.add_command(label="Save Attachments...",
                                            command=self._context_save_attachments)
        self.item_context_menu.add_separator()
        self.item_context_menu.add_command(label="Check / Uncheck Row",
                                            command=self._context_toggle_check)
        self.item_context_menu.add_command(label="Copy Subject",
                                            command=self._context_copy_subject)
        self.item_list.bind("<Button-3>", self._show_item_context_menu)

    def _build_preview(self, parent: ttk.PanedWindow) -> None:
        preview_frame = ttk.LabelFrame(parent, text="Preview", padding=4)
        parent.add(preview_frame, weight=PANE_WEIGHT_PREVIEW)

        header = ttk.Frame(preview_frame)
        header.pack(fill=tk.X, padx=2, pady=2)
        self.preview_subject = ttk.Label(header, text="",
                                         font=FONT_PREVIEW_SUBJECT,
                                         foreground=COLOR_ACCENT)
        self.preview_subject.pack(anchor=tk.W)
        self.preview_meta = ttk.Label(header, text="", foreground=COLOR_META)
        self.preview_meta.pack(anchor=tk.W)

        ttk.Separator(preview_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=2)

        # Attachments sub-panel (single + all save).
        self.preview_att_frame = ttk.LabelFrame(preview_frame, text="Attachments",
                                                padding=4)
        self.att_tree = ttk.Treeview(self.preview_att_frame, columns=("name", "size"),
                                     show="headings", selectmode="browse",
                                     height=ATTACH_TREE_HEIGHT)
        self.att_tree.heading("name", text="Filename")
        self.att_tree.heading("size", text="Size")
        self.att_tree.column("name", width=WIDTH_ATT_NAME)
        self.att_tree.column("size", width=WIDTH_ATT_SIZE, anchor=tk.E)
        att_scroll = ttk.Scrollbar(self.preview_att_frame, orient=tk.VERTICAL,
                                   command=self.att_tree.yview)
        self.att_tree.configure(yscrollcommand=att_scroll.set)
        self.att_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        att_scroll.pack(side=tk.LEFT, fill=tk.Y)
        att_btns = ttk.Frame(self.preview_att_frame)
        att_btns.pack(side=tk.LEFT, fill=tk.Y, padx=4)
        ttk.Button(att_btns, text="Save", command=self._save_one_attachment).pack(
            fill=tk.X, pady=1)
        ttk.Button(att_btns, text="Save All...",
                   command=self._save_all_preview_attachments).pack(fill=tk.X, pady=1)
        self.att_tree.bind("<Double-1>", lambda _e: self._save_one_attachment())
        # packed/forgotten dynamically by _show_preview

        self.preview_text = scrolledtext.ScrolledText(
            preview_frame, wrap=tk.WORD, state=tk.DISABLED, font=FONT_READING,
            relief=tk.FLAT, padx=8, pady=6)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    # ======================== Status Bar ========================

    def _build_statusbar(self) -> None:
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_label = ttk.Label(
            status_frame, text="Ready - Open a PST or OST file to begin")
        self.status_label.pack(side=tk.LEFT, padx=5, pady=2)

        self.checked_count_label = ttk.Label(status_frame, text="Checked: 0")
        self.checked_count_label.pack(side=tk.RIGHT, padx=5, pady=2)
        self.item_count_label = ttk.Label(status_frame, text="Items: 0")
        self.item_count_label.pack(side=tk.RIGHT, padx=5, pady=2)

    def _set_status(self, text: str) -> None:
        self.status_label.config(text=text)
        self.root.update_idletasks()

    # ======================== File Operations ========================

    def _ensure_reader(self) -> bool:
        if self.reader is not None:
            return True
        if not PYPFF_AVAILABLE:
            messagebox.showerror("Missing Dependency", INSTALL_HINT)
            return False
        try:
            self.reader = PstReader()
            return True
        except ImportError as exc:
            messagebox.showerror("Missing Dependency", str(exc))
            return False

    def _open_file(self) -> None:
        if not self._ensure_reader():
            return

        filepath = filedialog.askopenfilename(
            title="Open PST/OST File",
            filetypes=[
                ("Outlook Data Files", "*.pst *.ost"),
                ("PST Files", "*.pst"),
                ("OST Files", "*.ost"),
                ("All Files", "*.*"),
            ],
        )
        if not filepath:
            return

        self._set_status(f"Opening {os.path.basename(filepath)}...")
        self._clear_all()

        def load() -> None:
            try:
                if self.reader is None:
                    return
                self.reader.open(filepath)
                folders = self.reader.get_folder_tree()
                name = os.path.basename(filepath)
                self.root.after(0, lambda: self._populate_folder_tree(folders))
                self.root.after(0, lambda: self.header_file_label.config(text=name))
                self.root.after(0, lambda: self.root.title(f"{name} - {APP_TITLE}"))
                self.root.after(0, lambda: self._set_status(f"Loaded: {name}"))
            except Exception as exc:
                err = str(exc)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Failed to open file:\n{err}"))
                self.root.after(0, lambda: self._set_status("Error opening file"))

        threading.Thread(target=load, daemon=True).start()

    def _close_file(self) -> None:
        if self.reader:
            self.reader.close()
        self._clear_all()
        self._set_status("Ready - Open a PST or OST file to begin")

    def _clear_all(self) -> None:
        self.folder_tree.delete(*self.folder_tree.get_children())
        self.item_list.delete(*self.item_list.get_children())
        self._clear_preview()
        self.folder_map.clear()
        self.message_map.clear()
        self.checked_items.clear()
        self.current_folder = None
        self.current_messages = []
        self.item_count_label.config(text="Items: 0")
        self._update_checked_count()
        self.header_file_label.config(text="No file open")
        self.root.title(APP_TITLE)

    def _clear_preview(self) -> None:
        self.preview_subject.config(text="")
        self.preview_meta.config(text="")
        self._populate_preview_attachments([])
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.config(state=tk.DISABLED)

    def _check_file_open(self) -> bool:
        if not self.reader or not self.reader.is_open:
            messagebox.showwarning("No File", "Please open a PST or OST file first.")
            return False
        return True

    # ======================== Folder Tree ========================

    def _populate_folder_tree(self, folders: list[FolderInfo]) -> None:
        self.folder_tree.delete(*self.folder_tree.get_children())
        self.folder_map.clear()

        path_to_id: dict[str, str] = {}
        for folder in folders:
            parent_path = "/".join(folder.path.split("/")[:-1])
            parent_id = path_to_id.get(parent_path, "")
            indicator = FOLDER_TYPE_INDICATORS.get(folder.folder_type, "")
            if indicator:
                display = f"{folder.name} ({folder.message_count}) {indicator}"
            else:
                display = f"{folder.name} ({folder.message_count})"
            tree_id = self.folder_tree.insert(parent_id, tk.END, text=display,
                                              open=True)
            path_to_id[folder.path] = tree_id
            self.folder_map[tree_id] = folder

    def _show_folder_context_menu(self, event: tk.Event) -> None:
        try:
            row = self.folder_tree.identify_row(event.y)
            if row:
                self.folder_tree.selection_set(row)
            self.folder_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.folder_context_menu.grab_release()

    # ======================== Item List & checkboxes ========================

    def _on_folder_select(self, _event: tk.Event) -> None:
        selection = self.folder_tree.selection()
        if not selection:
            return
        folder_info = self.folder_map.get(selection[0])
        if not folder_info:
            return

        self.current_folder = folder_info
        self._set_status(f"Loading {folder_info.name}...")

        def load() -> None:
            try:
                if self.reader is None:
                    return
                messages = self.reader.get_messages(folder_info)
                self.root.after(0, lambda: self._populate_item_list(messages))
                self.root.after(0, lambda: self._set_status(
                    f"{folder_info.name} - {len(messages)} items"))
            except Exception as exc:
                err = str(exc)
                self.root.after(0, lambda: self._set_status(
                    f"Error loading folder: {err}"))

        threading.Thread(target=load, daemon=True).start()

    def _populate_item_list(self, messages: list[MessageItem]) -> None:
        self.item_list.delete(*self.item_list.get_children())
        self.message_map.clear()
        self.checked_items.clear()
        self.current_messages = messages
        self._clear_preview()

        for msg in messages:
            date_str = msg.date.strftime("%Y-%m-%d %H:%M") if msg.date else ""
            from_name = msg.sender_name or msg.properties.get("display_name", "")
            att_str = "Yes" if msg.has_attachments else ""
            tree_id = self.item_list.insert("", tk.END, values=(
                CHECK_OFF,
                msg.subject or NO_SUBJECT,
                from_name,
                date_str,
                msg.item_type.capitalize(),
                att_str,
            ))
            self.message_map[tree_id] = msg

        self.item_count_label.config(text=f"Items: {len(messages)}")
        self._update_checked_count()
        self._refresh_check_header()
        self._restripe()

    def _restripe(self) -> None:
        """Reapply zebra striping and the checked-row highlight to all rows."""
        for idx, row in enumerate(self.item_list.get_children()):
            if row in self.checked_items:
                tag = TAG_CHECKED
            else:
                tag = TAG_EVEN if idx % 2 else TAG_ODD
            self.item_list.item(row, tags=(tag,))

    def _on_item_click(self, event: tk.Event) -> str | None:
        """Toggle the checkbox when the check column cell is clicked."""
        if self.item_list.identify_region(event.x, event.y) != "cell":
            return None
        if self.item_list.identify_column(event.x) != f"#{1}":
            return None
        row = self.item_list.identify_row(event.y)
        if row:
            self._toggle_check(row)
        return "break"  # don't change row selection for checkbox clicks

    def _on_item_space(self, _event: tk.Event) -> str:
        for row in self.item_list.selection():
            self._toggle_check(row)
        return "break"

    def _toggle_check(self, row: str) -> None:
        if row in self.checked_items:
            self.checked_items.discard(row)
            self.item_list.set(row, COL_CHECK, CHECK_OFF)
            idx = self.item_list.index(row)
            self.item_list.item(row, tags=(TAG_EVEN if idx % 2 else TAG_ODD,))
        else:
            self.checked_items.add(row)
            self.item_list.set(row, COL_CHECK, CHECK_ON)
            self.item_list.item(row, tags=(TAG_CHECKED,))
        self._update_checked_count()
        self._refresh_check_header()

    def _set_all_checks(self, checked: bool) -> None:
        self.checked_items.clear()
        for row in self.item_list.get_children():
            self.item_list.set(row, COL_CHECK, CHECK_ON if checked else CHECK_OFF)
            if checked:
                self.checked_items.add(row)
        self._update_checked_count()
        self._refresh_check_header()
        self._restripe()

    def _toggle_all_checks(self) -> None:
        rows = self.item_list.get_children()
        self._set_all_checks(not (rows and len(self.checked_items) == len(rows)))

    def _refresh_check_header(self) -> None:
        rows = self.item_list.get_children()
        all_on = bool(rows) and len(self.checked_items) == len(rows)
        self.item_list.heading(COL_CHECK, text=CHECK_ON if all_on else CHECK_OFF)

    def _update_checked_count(self) -> None:
        self.checked_count_label.config(text=f"Checked: {len(self.checked_items)}")

    def _on_item_select(self, _event: tk.Event) -> None:
        selection = self.item_list.selection()
        if not selection:
            return
        msg = self.message_map.get(selection[0])
        if msg:
            self._show_preview(msg)

    def _on_item_double_click(self, _event: tk.Event) -> None:
        selection = self.item_list.selection()
        if not selection:
            return
        msg = self.message_map.get(selection[0])
        if msg:
            self._show_detail_window(msg)

    # ======================== Context Menu ========================

    def _show_item_context_menu(self, event: tk.Event) -> None:
        try:
            row = self.item_list.identify_row(event.y)
            if row:
                self.item_list.selection_set(row)
            self.item_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.item_context_menu.grab_release()

    def _get_selected_message(self) -> MessageItem | None:
        selection = self.item_list.selection()
        if selection:
            return self.message_map.get(selection[0])
        return None

    def _get_action_messages(self) -> list[MessageItem]:
        """Messages to act on: checked items, else the previewed row."""
        if self.checked_items:
            return [self.message_map[i] for i in self.checked_items
                    if i in self.message_map]
        single = self._get_selected_message()
        return [single] if single else []

    def _context_view_details(self) -> None:
        msg = self._get_selected_message()
        if msg:
            self._show_detail_window(msg)

    def _context_toggle_check(self) -> None:
        for row in self.item_list.selection():
            self._toggle_check(row)

    def _context_save_attachments(self) -> None:
        msg = self._get_selected_message()
        if not msg or not msg.has_attachments:
            messagebox.showinfo("No Attachments", "This item has no attachments.")
            return
        dirpath = filedialog.askdirectory(title="Save Attachments To")
        if not dirpath or self.reader is None:
            return
        count = exporters.export_attachments(self.reader, [msg], dirpath)
        messagebox.showinfo("Done", f"Saved {count} attachment(s).")

    def _context_copy_subject(self) -> None:
        msg = self._get_selected_message()
        if msg and msg.subject:
            self.root.clipboard_clear()
            self.root.clipboard_append(msg.subject)

    # ======================== Preview Pane ========================

    def _show_preview(self, msg: MessageItem) -> None:
        self.preview_subject.config(text=msg.subject or NO_SUBJECT)
        self.preview_meta.config(text=self._build_meta(msg))

        body = msg.body_text or html_to_text(msg.body_html)
        if not body and msg.item_type == "contact":
            body = self._build_contact_card(msg)
        if not body:
            body = "(No content available)"

        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", body)
        self.preview_text.config(state=tk.DISABLED)

        self._load_preview_attachments(msg)

    @staticmethod
    def _build_meta(msg: MessageItem) -> str:
        parts: list[str] = []
        props = msg.properties
        if msg.item_type == "email":
            if msg.sender_name or msg.sender_email:
                who = msg.sender_name or msg.sender_email
                parts.append(f"From: {who}")
            if msg.recipients:
                parts.append(f"To: {msg.recipients}")
            if msg.date:
                parts.append(f"Date: {msg.date.strftime('%Y-%m-%d %H:%M:%S')}")
            if msg.has_attachments:
                parts.append(f"Attachments: {msg.attachment_count}")
        elif msg.item_type == "contact":
            email = exporters.resolve_contact_email(msg)
            if email:
                parts.append(f"Email: {email}")
            if props.get("company_name"):
                parts.append(f"Company: {props['company_name']}")
            if props.get("business_phone"):
                parts.append(f"Phone: {props['business_phone']}")
        elif msg.item_type == "calendar":
            if props.get("start_date"):
                parts.append(f"Start: {props['start_date']}")
            if props.get("end_date"):
                parts.append(f"End: {props['end_date']}")
            if props.get("location"):
                parts.append(f"Location: {props['location']}")
        else:
            if msg.date:
                parts.append(f"Date: {msg.date}")
            if msg.message_class:
                parts.append(f"Class: {msg.message_class}")
        return "  |  ".join(parts)

    @staticmethod
    def _build_contact_card(msg: MessageItem) -> str:
        props = msg.properties
        fields = [
            ("display_name", "Name"),
            ("given_name", "First Name"),
            ("surname", "Last Name"),
            ("business_phone", "Business Phone"),
            ("home_phone", "Home Phone"),
            ("mobile_phone", "Mobile"),
            ("company_name", "Company"),
            ("title", "Job Title"),
            ("department_name", "Department"),
            ("office_location", "Office"),
            ("postal_address", "Address"),
            ("nickname", "Nickname"),
        ]
        lines = []
        email = exporters.resolve_contact_email(msg)
        if email:
            lines.append(f"Email: {email}")
        for key, label in fields:
            val = props.get(key, "")
            if val:
                lines.append(f"{label}: {val}")
        return "\n".join(lines) if lines else "(No contact details available)"

    # ---- preview attachments ----

    def _load_preview_attachments(self, msg: MessageItem) -> None:
        attachments: list[AttachmentInfo] = []
        if self.reader is not None and msg.has_attachments:
            try:
                attachments = self.reader.get_attachments(msg)
            except Exception:
                attachments = []
        self._populate_preview_attachments(attachments)

    def _populate_preview_attachments(self, attachments: list[AttachmentInfo]) -> None:
        self.preview_attachments = attachments
        self.preview_att_map.clear()
        if hasattr(self, "att_tree"):
            self.att_tree.delete(*self.att_tree.get_children())
        if not attachments:
            if self.preview_att_frame.winfo_manager():
                self.preview_att_frame.pack_forget()
            return
        for att in attachments:
            row = self.att_tree.insert("", tk.END, values=(
                att.name, exporters.format_size(att.size)))
            self.preview_att_map[row] = att
        if not self.preview_att_frame.winfo_manager():
            self.preview_att_frame.pack(fill=tk.X, padx=2, pady=(0, 2),
                                        before=self.preview_text)

    def _save_one_attachment(self) -> None:
        sel = self.att_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select an attachment to save.")
            return
        att = self.preview_att_map.get(sel[0])
        if att is None or self.reader is None:
            return
        path = filedialog.asksaveasfilename(title="Save Attachment",
                                            initialfile=att.name)
        if not path:
            return
        if self.reader.save_attachment(att, path):
            messagebox.showinfo("Saved", f"Saved to:\n{path}")
        else:
            messagebox.showerror("Error", "Failed to save this attachment.")

    def _save_all_preview_attachments(self) -> None:
        if not self.preview_attachments or self.reader is None:
            messagebox.showinfo("No Attachments", "There are no attachments to save.")
            return
        dirpath = filedialog.askdirectory(title="Save All Attachments To")
        if not dirpath:
            return
        msg = self._get_selected_message()
        if msg is None:
            return
        count = exporters.export_attachments(self.reader, [msg], dirpath)
        messagebox.showinfo("Done", f"Saved {count} attachment(s).")

    # ======================== Detail Window ========================

    def _show_detail_window(self, msg: MessageItem) -> None:
        win = tk.Toplevel(self.root)
        win.title(msg.subject or NO_SUBJECT)
        win.geometry(DETAIL_WINDOW_SIZE)
        win.transient(self.root)

        header = ttk.Frame(win, padding=10)
        header.pack(fill=tk.X)
        ttk.Label(header, text=msg.subject or NO_SUBJECT,
                  font=FONT_DETAIL_TITLE).pack(anchor=tk.W)

        if msg.sender_name or msg.sender_email:
            from_str = msg.sender_name or ""
            if msg.sender_email:
                from_str += f" <{msg.sender_email}>" if from_str else msg.sender_email
            ttk.Label(header, text=f"From: {from_str}").pack(anchor=tk.W)
        if msg.recipients:
            ttk.Label(header, text=f"To: {msg.recipients}").pack(anchor=tk.W)
        if msg.cc:
            ttk.Label(header, text=f"CC: {msg.cc}").pack(anchor=tk.W)
        if msg.date:
            ttk.Label(header, text=f"Date: {msg.date}").pack(anchor=tk.W)
        ttk.Label(header, text=f"Type: {msg.item_type.capitalize()}  |  "
                  f"Class: {msg.message_class or 'N/A'}",
                  foreground=COLOR_META).pack(anchor=tk.W)

        ttk.Separator(win, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self._build_detail_content_tab(notebook, msg)
        self._build_detail_props_tab(notebook, msg)
        if msg.has_attachments and self.reader:
            self._build_detail_attachments_tab(notebook, msg)

    def _build_detail_content_tab(self, notebook: ttk.Notebook,
                                  msg: MessageItem) -> None:
        body_frame = ttk.Frame(notebook)
        notebook.add(body_frame, text="Content")
        body_text = scrolledtext.ScrolledText(body_frame, wrap=tk.WORD,
                                              font=FONT_READING, relief=tk.FLAT,
                                              padx=8, pady=6)
        body_text.pack(fill=tk.BOTH, expand=True)
        content = msg.body_text or html_to_text(msg.body_html)
        if not content and msg.item_type == "contact":
            content = self._build_contact_card(msg)
        body_text.insert("1.0", content or "(No content)")
        body_text.config(state=tk.DISABLED)

    @staticmethod
    def _build_detail_props_tab(notebook: ttk.Notebook, msg: MessageItem) -> None:
        props_frame = ttk.Frame(notebook)
        notebook.add(props_frame, text="MAPI Properties")
        props_text = scrolledtext.ScrolledText(props_frame, wrap=tk.WORD,
                                               font=FONT_BODY)
        props_text.pack(fill=tk.BOTH, expand=True)
        max_value_len = 300
        lines = []
        for key in sorted(msg.properties.keys()):
            if key.startswith("_"):
                continue
            val = str(msg.properties[key])
            if len(val) > max_value_len:
                val = val[:max_value_len] + "..."
            lines.append(f"{key}: {val}")
        props_text.insert("1.0", "\n".join(lines) if lines else "(No properties)")
        props_text.config(state=tk.DISABLED)

    def _build_detail_attachments_tab(self, notebook: ttk.Notebook,
                                      msg: MessageItem) -> None:
        att_frame = ttk.Frame(notebook)
        notebook.add(att_frame, text=f"Attachments ({msg.attachment_count})")

        att_tree = ttk.Treeview(att_frame, columns=("name", "size"),
                                show="headings", selectmode="browse")
        att_tree.heading("name", text="Filename")
        att_tree.heading("size", text="Size")
        att_tree.column("name", width=WIDTH_ATT_NAME)
        att_tree.column("size", width=WIDTH_ATT_SIZE)
        att_tree.pack(fill=tk.BOTH, expand=True)

        attachments: list[AttachmentInfo] = []
        if self.reader is not None:
            try:
                attachments = self.reader.get_attachments(msg)
                for att in attachments:
                    att_tree.insert("", tk.END, values=(
                        att.name, exporters.format_size(att.size)))
            except Exception:
                pass

        btn_frame = ttk.Frame(att_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        def save_selected() -> None:
            sel = att_tree.selection()
            if not sel:
                messagebox.showinfo("Select", "Select an attachment first.")
                return
            idx = att_tree.index(sel[0])
            if idx < len(attachments) and self.reader is not None:
                att = attachments[idx]
                path = filedialog.asksaveasfilename(title="Save Attachment",
                                                    initialfile=att.name)
                if path and self.reader.save_attachment(att, path):
                    messagebox.showinfo("Saved", f"Saved to:\n{path}")
                elif path:
                    messagebox.showerror("Error", "Failed to save attachment.")

        def save_all() -> None:
            dirpath = filedialog.askdirectory(title="Save All Attachments To")
            if dirpath and self.reader is not None:
                count = exporters.export_attachments(self.reader, [msg], dirpath)
                messagebox.showinfo("Done", f"Saved {count} attachment(s).")

        ttk.Button(btn_frame, text="Save Selected", command=save_selected).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save All", command=save_all).pack(
            side=tk.LEFT, padx=5)

    # ======================== Export Helpers ========================

    def _choose_format(self, title: str = "Choose Export Format") -> str | None:
        """Modal dialog returning one of txt/csv/html/eml/pdf, or None."""
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=FORMAT_DIALOG_PAD)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Export format:", font=FONT_SECTION).pack(
            anchor=tk.W, pady=(0, 8))

        fmt_var = tk.StringVar(value=DEFAULT_FORMAT)
        for value, label in FORMAT_CHOICES:
            ttk.Radiobutton(frame, text=label, value=value,
                            variable=fmt_var).pack(anchor=tk.W, pady=1)

        result: dict[str, str | None] = {"fmt": None}

        def ok() -> None:
            result["fmt"] = fmt_var.get()
            dlg.destroy()

        btns = ttk.Frame(frame)
        btns.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(btns, text="Export", command=ok).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(
            side=tk.RIGHT, padx=6)
        dlg.bind("<Return>", lambda _e: ok())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.root.wait_window(dlg)
        return result["fmt"]

    def _ask_destination(self, fmt: str, default_name: str) -> str | None:
        if fmt == "csv":
            return filedialog.asksaveasfilename(
                title="Save CSV As", defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                initialfile=f"{default_name}.csv") or None
        return filedialog.askdirectory(
            title=f"Choose Output Folder for {fmt.upper()} Files") or None

    def _run_export(self, work, busy_msg: str) -> None:
        """Run work(progress) in a worker thread with status updates."""
        self._set_status(busy_msg)

        def progress(done: int, total: int, _label: str) -> None:
            txt = f"{busy_msg}  {done}/{total}" if total else f"{busy_msg}  {done}"
            self.root.after(0, lambda: self._set_status(txt))

        def runner() -> None:
            try:
                count = work(progress)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Export Complete", f"Exported {count} item(s)."))
                self.root.after(0, lambda: self._set_status(
                    f"Export complete - {count} item(s)"))
            except Exception as exc:
                err = str(exc)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Export failed:\n{err}"))
                self.root.after(0, lambda: self._set_status("Export error"))

        threading.Thread(target=runner, daemon=True).start()

    def _save_selected(self) -> None:
        if not self._check_file_open():
            return
        messages = self._get_action_messages()
        if not messages:
            messagebox.showinfo(
                "Nothing Selected",
                "Tick the checkbox on one or more items (or select a row), "
                "then try again.")
            return

        fmt = self._choose_format(f"Save {len(messages)} Item(s)")
        if not fmt:
            return
        dest = self._ask_destination(fmt, "selected_items")
        if not dest:
            return
        reader = self.reader
        if reader is None:
            return

        def work(progress):
            return exporters.export_selected(reader, messages, fmt, dest, progress)

        self._run_export(work, f"Exporting {len(messages)} item(s) to {fmt.upper()}...")

    def _export_folder_tree(self) -> None:
        if not self._check_file_open():
            return
        if not self.current_folder:
            messagebox.showinfo("No Folder", "Select a folder in the tree first.")
            return

        folder = self.current_folder
        fmt = self._choose_format(f"Export '{folder.name}' (+ subfolders)")
        if not fmt:
            return
        out_dir = filedialog.askdirectory(
            title="Choose Output Folder for the Exported Tree")
        if not out_dir:
            return
        reader = self.reader
        if reader is None:
            return

        def work(progress):
            return exporters.export_folder_tree(reader, folder, fmt, out_dir, progress)

        self._run_export(
            work, f"Exporting '{folder.name}' (+subfolders) to {fmt.upper()}...")

    def _export_typed(self, item_type: str, fmt: str) -> None:
        if not self._check_file_open():
            return
        label = TYPE_LABELS.get(item_type, "items")
        self._set_status(f"Collecting {label}...")

        def do_export() -> None:
            try:
                if self.reader is None:
                    return
                items = self.reader.get_all_items_by_type(item_type)
                if not items:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "No Items", f"No {label} found in this file."))
                    self.root.after(0, lambda: self._set_status("Ready"))
                    return
                self.root.after(0, lambda: self._save_typed(items, item_type, fmt,
                                                            label))
            except Exception as exc:
                err = str(exc)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Export failed:\n{err}"))
                self.root.after(0, lambda: self._set_status("Export error"))

        threading.Thread(target=do_export, daemon=True).start()

    def _save_typed(self, items: list[MessageItem], item_type: str, fmt: str,
                    label: str) -> None:
        count = 0
        if item_type == "calendar" and fmt == "ics":
            path = filedialog.asksaveasfilename(
                title="Export Calendar", defaultextension=".ics",
                filetypes=[("iCalendar", "*.ics"), ("All", "*.*")],
                initialfile="calendar.ics")
            if path:
                count = exporters.export_calendar_to_ics(items, path)
        elif item_type == "calendar" and fmt == "csv":
            path = filedialog.asksaveasfilename(
                title="Export Calendar", defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                initialfile="calendar.csv")
            if path:
                count = exporters.export_calendar_to_csv(items, path)
        elif item_type == "contact" and fmt == "vcf":
            path = filedialog.asksaveasfilename(
                title="Export Contacts", defaultextension=".vcf",
                filetypes=[("vCard", "*.vcf"), ("All", "*.*")],
                initialfile="contacts.vcf")
            if path:
                count = exporters.export_contacts_to_vcf(items, path)
        elif item_type == "contact" and fmt == "csv":
            path = filedialog.asksaveasfilename(
                title="Export Contacts", defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                initialfile="contacts.csv")
            if path:
                count = exporters.export_contacts_to_csv(items, path)
        elif item_type == "email" and fmt == "eml":
            dirpath = filedialog.askdirectory(
                title="Select Output Directory for EML Files")
            if dirpath:
                count = exporters.export_emails_to_eml(items, dirpath, self.reader)
        elif item_type == "email" and fmt == "csv":
            path = filedialog.asksaveasfilename(
                title="Export Emails", defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                initialfile="emails.csv")
            if path:
                count = exporters.export_emails_to_csv(items, path)

        if count > 0:
            messagebox.showinfo("Export Complete", f"Exported {count} {label}.")
        self._set_status("Ready")

    def _export_attachments(self) -> None:
        if not self._check_file_open():
            return
        if not self.current_messages:
            messagebox.showinfo("No Items", "Select a folder with items first.")
            return
        dirpath = filedialog.askdirectory(title="Save Attachments To")
        if not dirpath:
            return
        messages = self.current_messages
        reader = self.reader
        if reader is None:
            return

        def work(_progress):
            return exporters.export_attachments(reader, messages, dirpath)

        self._run_export(work, "Extracting attachments...")

    def _export_folder_csv(self) -> None:
        if not self._check_file_open():
            return
        if not self.current_messages:
            messagebox.showinfo("No Items", "Select a folder with items first.")
            return
        path = filedialog.asksaveasfilename(
            title="Export Folder to CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All", "*.*")],
            initialfile="folder_export.csv")
        if path:
            count = exporters.export_all_items_to_csv(self.current_messages, path)
            messagebox.showinfo("Done", f"Exported {count} items to CSV.")

    def _export_all_csv(self) -> None:
        if not self._check_file_open():
            return
        path = filedialog.asksaveasfilename(
            title="Export All to CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All", "*.*")],
            initialfile="all_items.csv")
        if not path:
            return
        reader = self.reader
        if reader is None:
            return

        def work(_progress):
            all_items: list[MessageItem] = []
            for folder in reader.get_folder_tree():
                all_items.extend(reader.get_messages(folder))
            return exporters.export_all_items_to_csv(all_items, path)

        self._run_export(work, "Exporting all items...")

    # ======================== Tools ========================

    def _show_search(self) -> None:
        if not self._check_file_open():
            return
        SearchWindow(self)

    def _show_file_properties(self) -> None:
        if not self._check_file_open() or self.reader is None:
            return
        filepath = self.reader.filepath
        if filepath is None:
            return
        try:
            file_size = os.path.getsize(filepath)
        except OSError:
            file_size = 0

        folders = self.reader.get_folder_tree()
        total_items = sum(f.message_count for f in folders)
        ext = os.path.splitext(filepath)[1].lower()
        ftype = ("PST (Personal Storage Table)" if ext == ".pst"
                 else "OST (Offline Storage Table)" if ext == ".ost" else "Unknown")
        info = (
            f"Filename:  {os.path.basename(filepath)}\n"
            f"Path:      {filepath}\n"
            f"Size:      {exporters.format_size(file_size)}\n"
            f"Type:      {ftype}\n\n"
            f"Folders:       {len(folders)}\n"
            f"Total Items:   {total_items}\n"
        )
        messagebox.showinfo("File Properties", info)

    def _show_folder_stats(self) -> None:
        if not self._check_file_open() or self.reader is None:
            return
        folders = self.reader.get_folder_tree()

        win = tk.Toplevel(self.root)
        win.title("Folder Statistics")
        win.geometry(STATS_WINDOW_SIZE)
        win.transient(self.root)

        cols = ("folder", "type", "items", "subfolders")
        headings = ("Folder Path", "Type", "Items", "Subfolders")
        tree = ttk.Treeview(win, columns=cols, show="headings", selectmode="browse")
        for col, text in zip(cols, headings, strict=True):
            tree.heading(col, text=text)
        tree.column("folder", width=350)
        tree.column("type", width=100)
        tree.column("items", width=80, anchor=tk.E)
        tree.column("subfolders", width=80, anchor=tk.E)
        scroll = ttk.Scrollbar(win, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        for folder in folders:
            tree.insert("", tk.END, values=(
                folder.path, folder.folder_type.capitalize(),
                folder.message_count, folder.subfolder_count))

    def _show_about(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("About Outlook Tool")
        win.geometry(ABOUT_WINDOW_SIZE)
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=APP_TITLE, font=FONT_ABOUT_TITLE).pack(pady=(0, 5))
        ttk.Label(frame, text=f"Created by {APP_AUTHOR}",
                  font=FONT_SPLASH_SUB).pack(pady=(0, 15))
        about_text = (
            "A standalone tool for reading and extracting data from\n"
            "Microsoft Outlook PST and OST files.\n"
            "Does NOT require Outlook to be installed.\n\n"
            "Features:\n"
            "  - Browse folder structure\n"
            "  - View emails, contacts, calendar items\n"
            "  - Check items and Save As txt / csv / html / eml / pdf\n"
            "  - Export a folder + subfolders in any format\n"
            "  - Embed or extract attachments (single or all)\n"
            "  - Export calendar to ICS / CSV, contacts to VCF / CSV\n"
            "  - Search across all items\n"
            "  - View MAPI properties"
        )
        ttk.Label(frame, text=about_text, justify=tk.LEFT).pack(anchor=tk.W)
        ttk.Button(frame, text="OK", command=win.destroy).pack(pady=(15, 0))

    # ======================== Utilities ========================

    def _sort_column(self, col: str) -> None:
        reverse = self._sort_reverse.get(col, False)
        rows = [(self.item_list.set(iid, col), iid)
                for iid in self.item_list.get_children()]
        rows.sort(reverse=reverse)
        for index, (_, iid) in enumerate(rows):
            self.item_list.move(iid, "", index)
        self._sort_reverse[col] = not reverse
        self._restripe()

    def _on_close(self) -> None:
        if self.reader:
            self.reader.close()
        self.root.destroy()


class SearchWindow:
    """Search dialog that scans all folders for matching items."""

    def __init__(self, app: OutlookToolApp) -> None:
        self.app = app
        self.result_messages: dict[str, MessageItem] = {}

        win = tk.Toplevel(app.root)
        self.win = win
        win.title("Search")
        win.geometry(SEARCH_WINDOW_SIZE)
        win.transient(app.root)

        search_frame = ttk.Frame(win, padding=5)
        search_frame.pack(fill=tk.X)
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        entry = ttk.Entry(search_frame, textvariable=self.search_var, width=50)
        entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(search_frame, text="Search", command=self._do_search).pack(
            side=tk.LEFT, padx=5)

        opts = ttk.Frame(win, padding=(5, 0))
        opts.pack(fill=tk.X)
        self.search_subject = tk.BooleanVar(value=True)
        self.search_body = tk.BooleanVar(value=True)
        self.search_sender = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Subject", variable=self.search_subject).pack(
            side=tk.LEFT, padx=5)
        ttk.Checkbutton(opts, text="Body", variable=self.search_body).pack(
            side=tk.LEFT, padx=5)
        ttk.Checkbutton(opts, text="Sender / Name",
                        variable=self.search_sender).pack(side=tk.LEFT, padx=5)

        results = ttk.Frame(win, padding=5)
        results.pack(fill=tk.BOTH, expand=True)
        cols = ("subject", "from", "date", "type", "folder")
        self.tree = ttk.Treeview(results, columns=cols, show="headings",
                                 selectmode="browse")
        for col, text, width in (
            ("subject", "Subject", 250), ("from", "From / Name", 150),
            ("date", "Date", 130), ("type", "Type", 70), ("folder", "Folder", 200),
        ):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width)
        scroll = ttk.Scrollbar(results, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", self._on_double)

        self.status = ttk.Label(win, text="", padding=5)
        self.status.pack(fill=tk.X)
        entry.bind("<Return>", lambda _e: self._do_search())
        entry.focus_set()

    def _do_search(self) -> None:
        query = self.search_var.get().strip().lower()
        if not query or self.app.reader is None:
            return
        self.tree.delete(*self.tree.get_children())
        self.result_messages.clear()
        self.status.config(text="Searching...")
        self.win.update_idletasks()

        reader = self.app.reader
        opts = (self.search_subject.get(), self.search_body.get(),
                self.search_sender.get())

        def worker() -> None:
            found = self._scan(reader, query, opts)
            self.win.after(0, lambda: self._show(found))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _scan(reader: PstReader, query: str, opts) -> list[MessageItem]:
        want_subject, want_body, want_sender = opts
        found: list[MessageItem] = []
        for folder in reader.get_folder_tree():
            try:
                for msg in reader.get_messages(folder):
                    if want_subject and msg.subject and query in msg.subject.lower():
                        found.append(msg)
                        continue
                    if want_body and msg.body_text and query in msg.body_text.lower():
                        found.append(msg)
                        continue
                    if want_sender:
                        sender = (msg.sender_name or "").lower()
                        display = str(msg.properties.get("display_name", "")).lower()
                        if query in sender or query in display:
                            found.append(msg)
            except Exception:
                pass
        return found

    def _show(self, found: list[MessageItem]) -> None:
        for msg in found:
            date_str = msg.date.strftime("%Y-%m-%d %H:%M") if msg.date else ""
            from_name = msg.sender_name or msg.properties.get("display_name", "")
            row = self.tree.insert("", tk.END, values=(
                msg.subject or NO_SUBJECT, from_name, date_str,
                msg.item_type.capitalize(), msg.folder_path))
            self.result_messages[row] = msg
        self.status.config(text=f"Found {len(found)} items")

    def _on_double(self, _event: tk.Event) -> None:
        sel = self.tree.selection()
        if sel:
            msg = self.result_messages.get(sel[0])
            if msg:
                self.app._show_detail_window(msg)


def _show_splash(root: tk.Tk) -> tk.Toplevel:
    """Show a simple splash screen while the app loads."""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    x = (splash.winfo_screenwidth() - SPLASH_WIDTH) // 2
    y = (splash.winfo_screenheight() - SPLASH_HEIGHT) // 2
    splash.geometry(f"{SPLASH_WIDTH}x{SPLASH_HEIGHT}+{x}+{y}")
    splash.configure(bg=SPLASH_BG)
    tk.Label(splash, text=APP_TITLE, font=FONT_SPLASH_TITLE, fg=SPLASH_FG,
             bg=SPLASH_BG).pack(expand=True, pady=(40, 5))
    tk.Label(splash, text=f"by {APP_AUTHOR}", font=FONT_SPLASH_SUB, fg=SPLASH_FG_SUB,
             bg=SPLASH_BG).pack()
    tk.Label(splash, text="Loading...", font=FONT_SPLASH_SUB, fg=SPLASH_FG_HINT,
             bg=SPLASH_BG).pack(expand=True, pady=(5, 40))
    splash.update()
    return splash


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    splash = _show_splash(root)
    OutlookToolApp(root)

    def finish_loading() -> None:
        splash.destroy()
        root.deiconify()

    root.after(SPLASH_DELAY_MS, finish_loading)
    root.mainloop()


if __name__ == "__main__":
    main()
