"""
Outlook PST/OST Tool - Main GUI Application.
A Windows GUI tool for reading and extracting data from Outlook PST and OST files
without requiring Microsoft Outlook to be installed.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import List, Optional

# Add the script's directory to path so imports work from any cwd
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pst_handler import PstReader, FolderInfo, MessageItem, PYPFF_AVAILABLE
import exporters


class OutlookToolApp:
    """Main GUI application."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Outlook PST/OST Tool")
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)

        self.reader: Optional[PstReader] = None
        self.current_folder: Optional[FolderInfo] = None
        self.current_messages: List[MessageItem] = []
        self.folder_map = {}   # tree_id -> FolderInfo
        self.message_map = {}  # tree_id -> MessageItem
        self._sort_reverse = {}  # column -> bool (tracks sort direction)

        self._build_menu()
        self._build_toolbar()
        self._build_main_layout()
        self._build_statusbar()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self._open_file())
        self.root.bind('<Control-f>', lambda e: self._show_search())

        # Try to initialize PstReader on startup to check pypff availability
        if not PYPFF_AVAILABLE:
            self._set_status(
                "WARNING: pypff not installed. Install with: pip install libpff-python")

    # ======================== Menu Bar ========================

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open PST/OST...", command=self._open_file,
                              accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Close File", command=self._close_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)

        # Export
        export_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Export", menu=export_menu)

        cal_menu = tk.Menu(export_menu, tearoff=0)
        export_menu.add_cascade(label="Calendar", menu=cal_menu)
        cal_menu.add_command(label="Export All to ICS...",
                             command=lambda: self._export_typed('calendar', 'ics'))
        cal_menu.add_command(label="Export All to CSV...",
                             command=lambda: self._export_typed('calendar', 'csv'))

        con_menu = tk.Menu(export_menu, tearoff=0)
        export_menu.add_cascade(label="Contacts", menu=con_menu)
        con_menu.add_command(label="Export All to VCF...",
                             command=lambda: self._export_typed('contact', 'vcf'))
        con_menu.add_command(label="Export All to CSV...",
                             command=lambda: self._export_typed('contact', 'csv'))

        email_menu = tk.Menu(export_menu, tearoff=0)
        export_menu.add_cascade(label="Emails", menu=email_menu)
        email_menu.add_command(label="Export All to EML files...",
                               command=lambda: self._export_typed('email', 'eml'))
        email_menu.add_command(label="Export All to CSV...",
                               command=lambda: self._export_typed('email', 'csv'))

        export_menu.add_separator()
        export_menu.add_command(label="Extract Attachments (Current Folder)...",
                                command=self._export_attachments)
        export_menu.add_separator()
        export_menu.add_command(label="Export Current Folder to CSV...",
                                command=self._export_folder_csv)
        export_menu.add_command(label="Export Entire File to CSV...",
                                command=self._export_all_csv)

        # Tools
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Search...", command=self._show_search,
                               accelerator="Ctrl+F")
        tools_menu.add_separator()
        tools_menu.add_command(label="File Properties...",
                               command=self._show_file_properties)
        tools_menu.add_command(label="Folder Statistics...",
                               command=self._show_folder_stats)

        # Help
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)

    # ======================== Toolbar ========================

    def _build_toolbar(self):
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=2, pady=(2, 0))

        ttk.Button(toolbar, text="Open File",
                   command=self._open_file).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=2)
        ttk.Button(toolbar, text="Export Calendar",
                   command=lambda: self._export_typed('calendar', 'ics')).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Export Contacts",
                   command=lambda: self._export_typed('contact', 'vcf')).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Export Emails",
                   command=lambda: self._export_typed('email', 'eml')).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Attachments",
                   command=self._export_attachments).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=2)
        ttk.Button(toolbar, text="Search",
                   command=self._show_search).pack(side=tk.LEFT, padx=2)

    # ======================== Main Layout ========================

    def _build_main_layout(self):
        # Horizontal split: folder tree | (item list + preview)
        self.main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # --- Left: Folder Tree ---
        left_frame = ttk.LabelFrame(self.main_pane, text="Folders", padding=4)
        self.main_pane.add(left_frame, weight=1)

        self.folder_tree = ttk.Treeview(left_frame, show='tree', selectmode='browse')
        folder_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL,
                                       command=self.folder_tree.yview)
        self.folder_tree.configure(yscrollcommand=folder_scroll.set)
        self.folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        folder_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.bind('<<TreeviewSelect>>', self._on_folder_select)

        # --- Right: item list + preview (vertical split) ---
        right_pane = ttk.PanedWindow(self.main_pane, orient=tk.VERTICAL)
        self.main_pane.add(right_pane, weight=3)

        # Item list
        list_frame = ttk.LabelFrame(right_pane, text="Items", padding=4)
        right_pane.add(list_frame, weight=2)

        columns = ('subject', 'from', 'date', 'type', 'att')
        self.item_list = ttk.Treeview(list_frame, columns=columns,
                                       show='headings', selectmode='browse')
        self.item_list.heading('subject', text='Subject',
                               command=lambda: self._sort_column('subject'))
        self.item_list.heading('from', text='From / Name',
                               command=lambda: self._sort_column('from'))
        self.item_list.heading('date', text='Date',
                               command=lambda: self._sort_column('date'))
        self.item_list.heading('type', text='Type',
                               command=lambda: self._sort_column('type'))
        self.item_list.heading('att', text='Att',
                               command=lambda: self._sort_column('att'))

        self.item_list.column('subject', width=300, minwidth=100)
        self.item_list.column('from', width=200, minwidth=80)
        self.item_list.column('date', width=150, minwidth=80)
        self.item_list.column('type', width=80, minwidth=60)
        self.item_list.column('att', width=40, minwidth=30, anchor=tk.CENTER)

        list_ys = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                 command=self.item_list.yview)
        list_xs = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL,
                                 command=self.item_list.xview)
        self.item_list.configure(yscrollcommand=list_ys.set,
                                  xscrollcommand=list_xs.set)
        self.item_list.grid(row=0, column=0, sticky='nsew')
        list_ys.grid(row=0, column=1, sticky='ns')
        list_xs.grid(row=1, column=0, sticky='ew')
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.item_list.bind('<<TreeviewSelect>>', self._on_item_select)
        self.item_list.bind('<Double-1>', self._on_item_double_click)

        # Context menu for item list
        self.item_context_menu = tk.Menu(self.item_list, tearoff=0)
        self.item_context_menu.add_command(label="View Details",
                                            command=self._context_view_details)
        self.item_context_menu.add_command(label="Save Attachments...",
                                            command=self._context_save_attachments)
        self.item_context_menu.add_separator()
        self.item_context_menu.add_command(label="Copy Subject",
                                            command=self._context_copy_subject)
        self.item_list.bind('<Button-3>', self._show_item_context_menu)

        # Preview pane
        preview_frame = ttk.LabelFrame(right_pane, text="Preview", padding=4)
        right_pane.add(preview_frame, weight=1)

        self.preview_header = ttk.Frame(preview_frame)
        self.preview_header.pack(fill=tk.X, padx=2, pady=2)

        self.preview_subject = ttk.Label(self.preview_header, text="",
                                          font=('Segoe UI', 11, 'bold'))
        self.preview_subject.pack(anchor=tk.W)
        self.preview_meta = ttk.Label(self.preview_header, text="",
                                       foreground='gray')
        self.preview_meta.pack(anchor=tk.W)

        ttk.Separator(preview_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=2)

        self.preview_text = scrolledtext.ScrolledText(
            preview_frame, wrap=tk.WORD, state=tk.DISABLED,
            font=('Consolas', 10))
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    # ======================== Status Bar ========================

    def _build_statusbar(self):
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_label = ttk.Label(status_frame,
                                       text="Ready - Open a PST or OST file to begin")
        self.status_label.pack(side=tk.LEFT, padx=5, pady=2)

        self.item_count_label = ttk.Label(status_frame, text="Items: 0")
        self.item_count_label.pack(side=tk.RIGHT, padx=5, pady=2)

    def _set_status(self, text: str):
        self.status_label.config(text=text)
        self.root.update_idletasks()

    # ======================== File Operations ========================

    def _ensure_reader(self) -> bool:
        """Ensure PstReader is initialized."""
        if self.reader is not None:
            return True
        if not PYPFF_AVAILABLE:
            messagebox.showerror(
                "Missing Dependency",
                "pypff (libpff-python) is required but not installed.\n\n"
                "Install with one of:\n"
                "  pip install libpff-python\n"
                "  conda install -c conda-forge libpff-python\n\n"
                "See https://github.com/libyal/libpff for details."
            )
            return False
        try:
            self.reader = PstReader()
            return True
        except ImportError as e:
            messagebox.showerror("Missing Dependency", str(e))
            return False

    def _open_file(self):
        if not self._ensure_reader():
            return

        filepath = filedialog.askopenfilename(
            title="Open PST/OST File",
            filetypes=[
                ("Outlook Data Files", "*.pst *.ost"),
                ("PST Files", "*.pst"),
                ("OST Files", "*.ost"),
                ("All Files", "*.*"),
            ]
        )
        if not filepath:
            return

        self._set_status(f"Opening {os.path.basename(filepath)}...")
        self._clear_all()

        def load():
            try:
                if self.reader is None:
                    return
                self.reader.open(filepath)
                folders = self.reader.get_folder_tree()
                self.root.after(0, lambda: self._populate_folder_tree(folders))
                self.root.after(0, lambda: self._set_status(
                    f"Loaded: {os.path.basename(filepath)}"))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Failed to open file:\n{err_msg}"))
                self.root.after(0, lambda: self._set_status("Error opening file"))

        threading.Thread(target=load, daemon=True).start()

    def _close_file(self):
        if self.reader:
            self.reader.close()
        self._clear_all()
        self._set_status("Ready - Open a PST or OST file to begin")

    def _clear_all(self):
        self.folder_tree.delete(*self.folder_tree.get_children())
        self.item_list.delete(*self.item_list.get_children())
        self._clear_preview()
        self.folder_map.clear()
        self.message_map.clear()
        self.current_folder = None
        self.current_messages = []
        self.item_count_label.config(text="Items: 0")

    def _clear_preview(self):
        self.preview_subject.config(text="")
        self.preview_meta.config(text="")
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.config(state=tk.DISABLED)

    def _check_file_open(self) -> bool:
        if not self.reader or not self.reader.is_open:
            messagebox.showwarning("No File", "Please open a PST or OST file first.")
            return False
        return True

    # ======================== Folder Tree ========================

    def _populate_folder_tree(self, folders: List[FolderInfo]):
        self.folder_tree.delete(*self.folder_tree.get_children())
        self.folder_map.clear()

        path_to_id = {}
        type_indicators = {
            'email': '[Mail]',
            'contacts': '[Contacts]',
            'calendar': '[Calendar]',
            'tasks': '[Tasks]',
            'notes': '[Notes]',
            'other': '',
        }

        for folder in folders:
            parent_path = '/'.join(folder.path.split('/')[:-1])
            parent_id = path_to_id.get(parent_path, '')

            indicator = type_indicators.get(folder.folder_type, '')
            if indicator:
                display = f"{folder.name} ({folder.message_count}) {indicator}"
            else:
                display = f"{folder.name} ({folder.message_count})"

            tree_id = self.folder_tree.insert(parent_id, tk.END, text=display,
                                               open=True)
            path_to_id[folder.path] = tree_id
            self.folder_map[tree_id] = folder

    # ======================== Event Handlers ========================

    def _on_folder_select(self, event):
        selection = self.folder_tree.selection()
        if not selection:
            return

        folder_info = self.folder_map.get(selection[0])
        if not folder_info:
            return

        self.current_folder = folder_info
        self._set_status(f"Loading {folder_info.name}...")

        def load():
            try:
                if self.reader is None:
                    return
                messages = self.reader.get_messages(folder_info)
                self.root.after(0, lambda: self._populate_item_list(messages))
                self.root.after(0, lambda: self._set_status(
                    f"{folder_info.name} - {len(messages)} items"))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: self._set_status(
                    f"Error loading folder: {err_msg}"))

        threading.Thread(target=load, daemon=True).start()

    def _populate_item_list(self, messages: List[MessageItem]):
        self.item_list.delete(*self.item_list.get_children())
        self.message_map.clear()
        self.current_messages = messages
        self._clear_preview()

        for msg in messages:
            date_str = msg.date.strftime('%Y-%m-%d %H:%M') if msg.date else ''
            from_name = msg.sender_name or msg.properties.get('display_name', '')
            att_str = 'Yes' if msg.has_attachments else ''

            tree_id = self.item_list.insert('', tk.END, values=(
                msg.subject or '(No Subject)',
                from_name,
                date_str,
                msg.item_type.capitalize(),
                att_str,
            ))
            self.message_map[tree_id] = msg

        self.item_count_label.config(text=f"Items: {len(messages)}")

    def _on_item_select(self, event):
        selection = self.item_list.selection()
        if not selection:
            return
        msg = self.message_map.get(selection[0])
        if msg:
            self._show_preview(msg)

    def _on_item_double_click(self, event):
        selection = self.item_list.selection()
        if not selection:
            return
        msg = self.message_map.get(selection[0])
        if msg:
            self._show_detail_window(msg)

    # ======================== Context Menu ========================

    def _show_item_context_menu(self, event):
        try:
            self.item_list.selection_set(
                self.item_list.identify_row(event.y))
            self.item_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.item_context_menu.grab_release()

    def _get_selected_message(self) -> Optional[MessageItem]:
        selection = self.item_list.selection()
        if selection:
            return self.message_map.get(selection[0])
        return None

    def _context_view_details(self):
        msg = self._get_selected_message()
        if msg:
            self._show_detail_window(msg)

    def _context_save_attachments(self):
        msg = self._get_selected_message()
        if not msg or not msg.has_attachments:
            messagebox.showinfo("No Attachments", "This item has no attachments.")
            return

        dirpath = filedialog.askdirectory(title="Save Attachments To")
        if not dirpath:
            return

        if self.reader is not None:
            count = exporters.export_attachments(self.reader, [msg], dirpath)
            messagebox.showinfo("Done", f"Saved {count} attachment(s).")

    def _context_copy_subject(self):
        msg = self._get_selected_message()
        if msg and msg.subject:
            self.root.clipboard_clear()
            self.root.clipboard_append(msg.subject)

    # ======================== Preview Pane ========================

    def _show_preview(self, msg: MessageItem):
        self.preview_subject.config(text=msg.subject or '(No Subject)')

        meta_parts = []
        if msg.item_type == 'email':
            if msg.sender_name:
                meta_parts.append(f"From: {msg.sender_name}")
            if msg.recipients:
                meta_parts.append(f"To: {msg.recipients}")
            if msg.date:
                meta_parts.append(f"Date: {msg.date.strftime('%Y-%m-%d %H:%M:%S')}")
            if msg.has_attachments:
                meta_parts.append(f"Attachments: {msg.attachment_count}")
        elif msg.item_type == 'contact':
            props = msg.properties
            if props.get('smtp_address'):
                meta_parts.append(f"Email: {props['smtp_address']}")
            if props.get('company_name'):
                meta_parts.append(f"Company: {props['company_name']}")
            if props.get('business_phone'):
                meta_parts.append(f"Phone: {props['business_phone']}")
        elif msg.item_type == 'calendar':
            props = msg.properties
            if props.get('start_date'):
                meta_parts.append(f"Start: {props['start_date']}")
            if props.get('end_date'):
                meta_parts.append(f"End: {props['end_date']}")
            if props.get('location'):
                meta_parts.append(f"Location: {props['location']}")
        else:
            if msg.date:
                meta_parts.append(f"Date: {msg.date}")
            if msg.message_class:
                meta_parts.append(f"Class: {msg.message_class}")

        self.preview_meta.config(text="  |  ".join(meta_parts))

        # Build body content
        body = msg.body_text or msg.body_html or ''
        if not body and msg.item_type == 'contact':
            body = self._build_contact_card(msg)
        if not body:
            body = '(No content available)'

        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert('1.0', body)
        self.preview_text.config(state=tk.DISABLED)

    @staticmethod
    def _build_contact_card(msg: MessageItem) -> str:
        """Build a readable text card for a contact item."""
        props = msg.properties
        lines = []
        fields = [
            ('display_name', 'Name'),
            ('given_name', 'First Name'),
            ('surname', 'Last Name'),
            ('smtp_address', 'Email'),
            ('business_phone', 'Business Phone'),
            ('home_phone', 'Home Phone'),
            ('mobile_phone', 'Mobile'),
            ('company_name', 'Company'),
            ('title', 'Job Title'),
            ('department_name', 'Department'),
            ('office_location', 'Office'),
            ('postal_address', 'Address'),
            ('nickname', 'Nickname'),
        ]
        for key, label in fields:
            val = props.get(key, '')
            if val:
                lines.append(f"{label}: {val}")
        return '\n'.join(lines) if lines else '(No contact details available)'

    # ======================== Detail Window ========================

    def _show_detail_window(self, msg: MessageItem):
        win = tk.Toplevel(self.root)
        win.title(msg.subject or '(No Subject)')
        win.geometry("750x550")
        win.transient(self.root)

        # Header
        header = ttk.Frame(win, padding=10)
        header.pack(fill=tk.X)

        ttk.Label(header, text=msg.subject or '(No Subject)',
                  font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W)

        if msg.sender_name or msg.sender_email:
            from_str = msg.sender_name or ''
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
                  foreground='gray').pack(anchor=tk.W)

        ttk.Separator(win, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        # Notebook with tabs
        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Content tab
        body_frame = ttk.Frame(notebook)
        notebook.add(body_frame, text="Content")
        body_text = scrolledtext.ScrolledText(body_frame, wrap=tk.WORD,
                                               font=('Consolas', 10))
        body_text.pack(fill=tk.BOTH, expand=True)
        content = msg.body_text or msg.body_html or ''
        if not content and msg.item_type == 'contact':
            content = self._build_contact_card(msg)
        body_text.insert('1.0', content or '(No content)')
        body_text.config(state=tk.DISABLED)

        # Properties tab
        props_frame = ttk.Frame(notebook)
        notebook.add(props_frame, text="MAPI Properties")
        props_text = scrolledtext.ScrolledText(props_frame, wrap=tk.WORD,
                                                font=('Consolas', 10))
        props_text.pack(fill=tk.BOTH, expand=True)
        props_lines = []
        for key in sorted(msg.properties.keys()):
            val = str(msg.properties[key])
            if len(val) > 300:
                val = val[:300] + '...'
            props_lines.append(f"{key}: {val}")
        props_text.insert('1.0', '\n'.join(props_lines) if props_lines else '(No properties)')
        props_text.config(state=tk.DISABLED)

        # Attachments tab
        if msg.has_attachments and self.reader:
            att_frame = ttk.Frame(notebook)
            notebook.add(att_frame, text=f"Attachments ({msg.attachment_count})")

            att_tree = ttk.Treeview(att_frame, columns=('name', 'size'),
                                    show='headings', selectmode='browse')
            att_tree.heading('name', text='Filename')
            att_tree.heading('size', text='Size')
            att_tree.column('name', width=400)
            att_tree.column('size', width=100)
            att_tree.pack(fill=tk.BOTH, expand=True)

            attachments = []
            try:
                attachments = self.reader.get_attachments(msg)
                for att in attachments:
                    att_tree.insert('', tk.END, values=(
                        att.name, self._format_size(att.size)))
            except Exception:
                pass

            btn_frame = ttk.Frame(att_frame)
            btn_frame.pack(fill=tk.X, pady=5)

            def save_selected():
                sel = att_tree.selection()
                if not sel:
                    messagebox.showinfo("Select", "Select an attachment first.")
                    return
                idx = att_tree.index(sel[0])
                if idx < len(attachments) and self.reader is not None:
                    att = attachments[idx]
                    filepath = filedialog.asksaveasfilename(
                        title="Save Attachment", initialfile=att.name)
                    if filepath:
                        if self.reader.save_attachment(att, filepath):
                            messagebox.showinfo("Saved", f"Saved to:\n{filepath}")
                        else:
                            messagebox.showerror("Error", "Failed to save attachment.")

            def save_all():
                dirpath = filedialog.askdirectory(title="Save All Attachments To")
                if not dirpath:
                    return
                if self.reader is not None:
                    count = exporters.export_attachments(self.reader, [msg], dirpath)
                    messagebox.showinfo("Done", f"Saved {count} attachment(s).")

            ttk.Button(btn_frame, text="Save Selected", command=save_selected).pack(
                side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Save All", command=save_all).pack(
                side=tk.LEFT, padx=5)

    # ======================== Export Functions ========================

    def _export_typed(self, item_type: str, fmt: str):
        """Export all items of a given type in a given format."""
        if not self._check_file_open():
            return

        type_labels = {
            'calendar': 'calendar items',
            'contact': 'contacts',
            'email': 'emails',
        }
        label = type_labels.get(item_type, 'items')
        self._set_status(f"Collecting {label}...")

        def do_export():
            try:
                if self.reader is None:
                    return
                items = self.reader.get_all_items_by_type(item_type)
                if not items:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "No Items", f"No {label} found in this file."))
                    self.root.after(0, lambda: self._set_status("Ready"))
                    return

                def save():
                    count = 0
                    if item_type == 'calendar' and fmt == 'ics':
                        path = filedialog.asksaveasfilename(
                            title="Export Calendar", defaultextension=".ics",
                            filetypes=[("iCalendar", "*.ics"), ("All", "*.*")],
                            initialfile="calendar.ics")
                        if path:
                            count = exporters.export_calendar_to_ics(items, path)
                    elif item_type == 'calendar' and fmt == 'csv':
                        path = filedialog.asksaveasfilename(
                            title="Export Calendar", defaultextension=".csv",
                            filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                            initialfile="calendar.csv")
                        if path:
                            count = exporters.export_calendar_to_csv(items, path)
                    elif item_type == 'contact' and fmt == 'vcf':
                        path = filedialog.asksaveasfilename(
                            title="Export Contacts", defaultextension=".vcf",
                            filetypes=[("vCard", "*.vcf"), ("All", "*.*")],
                            initialfile="contacts.vcf")
                        if path:
                            count = exporters.export_contacts_to_vcf(items, path)
                    elif item_type == 'contact' and fmt == 'csv':
                        path = filedialog.asksaveasfilename(
                            title="Export Contacts", defaultextension=".csv",
                            filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                            initialfile="contacts.csv")
                        if path:
                            count = exporters.export_contacts_to_csv(items, path)
                    elif item_type == 'email' and fmt == 'eml':
                        dirpath = filedialog.askdirectory(
                            title="Select Output Directory for EML Files")
                        if dirpath:
                            count = exporters.export_emails_to_eml(items, dirpath)
                    elif item_type == 'email' and fmt == 'csv':
                        path = filedialog.asksaveasfilename(
                            title="Export Emails", defaultextension=".csv",
                            filetypes=[("CSV", "*.csv"), ("All", "*.*")],
                            initialfile="emails.csv")
                        if path:
                            count = exporters.export_emails_to_csv(items, path)

                    if count > 0:
                        messagebox.showinfo("Export Complete",
                                            f"Exported {count} {label}.")
                    self._set_status("Ready")

                self.root.after(0, save)
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Export failed:\n{err_msg}"))
                self.root.after(0, lambda: self._set_status("Export error"))

        threading.Thread(target=do_export, daemon=True).start()

    def _export_attachments(self):
        if not self._check_file_open():
            return
        if not self.current_messages:
            messagebox.showinfo("No Items", "Select a folder with items first.")
            return

        dirpath = filedialog.askdirectory(title="Save Attachments To")
        if not dirpath:
            return

        self._set_status("Extracting attachments...")

        def do_export():
            try:
                if self.reader is None:
                    return
                count = exporters.export_attachments(
                    self.reader, self.current_messages, dirpath)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Done", f"Extracted {count} attachment(s)."))
                self.root.after(0, lambda: self._set_status("Ready"))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Export failed:\n{err_msg}"))
                self.root.after(0, lambda: self._set_status("Export error"))

        threading.Thread(target=do_export, daemon=True).start()

    def _export_folder_csv(self):
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

    def _export_all_csv(self):
        if not self._check_file_open():
            return

        path = filedialog.asksaveasfilename(
            title="Export All to CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All", "*.*")],
            initialfile="all_items.csv")
        if not path:
            return

        self._set_status("Exporting all items...")

        def do_export():
            try:
                if self.reader is None:
                    return
                all_items = []
                folders = self.reader.get_folder_tree()
                for folder in folders:
                    items = self.reader.get_messages(folder)
                    all_items.extend(items)

                count = exporters.export_all_items_to_csv(all_items, path)
                self.root.after(0, lambda: messagebox.showinfo(
                    "Done", f"Exported {count} items to CSV."))
                self.root.after(0, lambda: self._set_status("Ready"))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror(
                    "Error", f"Export failed:\n{err_msg}"))
                self.root.after(0, lambda: self._set_status("Export error"))

        threading.Thread(target=do_export, daemon=True).start()

    # ======================== Tools ========================

    def _show_search(self):
        if not self._check_file_open():
            return

        win = tk.Toplevel(self.root)
        win.title("Search")
        win.geometry("850x550")
        win.transient(self.root)

        # Search bar
        search_frame = ttk.Frame(win, padding=5)
        search_frame.pack(fill=tk.X)

        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=50)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Options
        opts_frame = ttk.Frame(win, padding=(5, 0))
        opts_frame.pack(fill=tk.X)

        search_subject = tk.BooleanVar(value=True)
        search_body = tk.BooleanVar(value=True)
        search_sender = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts_frame, text="Subject",
                         variable=search_subject).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(opts_frame, text="Body",
                         variable=search_body).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(opts_frame, text="Sender / Name",
                         variable=search_sender).pack(side=tk.LEFT, padx=5)

        # Results
        results_frame = ttk.Frame(win, padding=5)
        results_frame.pack(fill=tk.BOTH, expand=True)

        cols = ('subject', 'from', 'date', 'type', 'folder')
        results_tree = ttk.Treeview(results_frame, columns=cols, show='headings',
                                     selectmode='browse')
        results_tree.heading('subject', text='Subject')
        results_tree.heading('from', text='From / Name')
        results_tree.heading('date', text='Date')
        results_tree.heading('type', text='Type')
        results_tree.heading('folder', text='Folder')
        results_tree.column('subject', width=250)
        results_tree.column('from', width=150)
        results_tree.column('date', width=130)
        results_tree.column('type', width=70)
        results_tree.column('folder', width=200)

        rs = ttk.Scrollbar(results_frame, orient=tk.VERTICAL,
                            command=results_tree.yview)
        results_tree.configure(yscrollcommand=rs.set)
        results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rs.pack(side=tk.RIGHT, fill=tk.Y)

        status_lbl = ttk.Label(win, text="", padding=5)
        status_lbl.pack(fill=tk.X)

        result_messages = {}

        def do_search():
            query = search_var.get().strip().lower()
            if not query:
                return

            results_tree.delete(*results_tree.get_children())
            result_messages.clear()
            status_lbl.config(text="Searching...")
            win.update_idletasks()

            def search_thread():
                found = []
                if self.reader is None:
                    return
                folders = self.reader.get_folder_tree()
                for folder in folders:
                    try:
                        messages = self.reader.get_messages(folder)
                        for msg in messages:
                            match = False
                            if search_subject.get() and msg.subject and \
                                    query in msg.subject.lower():
                                match = True
                            if not match and search_body.get() and msg.body_text and \
                                    query in msg.body_text.lower():
                                match = True
                            if not match and search_sender.get():
                                sender = (msg.sender_name or '').lower()
                                display = str(
                                    msg.properties.get('display_name', '')).lower()
                                if query in sender or query in display:
                                    match = True
                            if match:
                                found.append(msg)
                    except Exception:
                        pass

                def show():
                    for msg in found:
                        d = msg.date.strftime('%Y-%m-%d %H:%M') if msg.date else ''
                        fn = msg.sender_name or msg.properties.get('display_name', '')
                        tid = results_tree.insert('', tk.END, values=(
                            msg.subject or '(No Subject)',
                            fn, d,
                            msg.item_type.capitalize(),
                            msg.folder_path,
                        ))
                        result_messages[tid] = msg
                    status_lbl.config(text=f"Found {len(found)} items")

                win.after(0, show)

            threading.Thread(target=search_thread, daemon=True).start()

        def on_dbl(event):
            sel = results_tree.selection()
            if sel:
                msg = result_messages.get(sel[0])
                if msg:
                    self._show_detail_window(msg)

        results_tree.bind('<Double-1>', on_dbl)

        search_btn = ttk.Button(search_frame, text="Search", command=do_search)
        search_btn.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<Return>', lambda e: do_search())
        search_entry.focus_set()

    def _show_file_properties(self):
        if not self._check_file_open():
            return
        assert self.reader is not None

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
        ftype = 'PST (Personal Storage Table)' if ext == '.pst' else \
                'OST (Offline Storage Table)' if ext == '.ost' else 'Unknown'

        info = (
            f"Filename:  {os.path.basename(filepath)}\n"
            f"Path:      {filepath}\n"
            f"Size:      {self._format_size(file_size)}\n"
            f"Type:      {ftype}\n"
            f"\n"
            f"Folders:       {len(folders)}\n"
            f"Total Items:   {total_items}\n"
        )
        messagebox.showinfo("File Properties", info)

    def _show_folder_stats(self):
        if not self._check_file_open():
            return
        assert self.reader is not None

        folders = self.reader.get_folder_tree()

        win = tk.Toplevel(self.root)
        win.title("Folder Statistics")
        win.geometry("650x400")
        win.transient(self.root)

        cols = ('folder', 'type', 'items', 'subfolders')
        tree = ttk.Treeview(win, columns=cols, show='headings', selectmode='browse')
        tree.heading('folder', text='Folder Path')
        tree.heading('type', text='Type')
        tree.heading('items', text='Items')
        tree.heading('subfolders', text='Subfolders')
        tree.column('folder', width=350)
        tree.column('type', width=100)
        tree.column('items', width=80, anchor=tk.E)
        tree.column('subfolders', width=80, anchor=tk.E)

        scroll = ttk.Scrollbar(win, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        for folder in folders:
            tree.insert('', tk.END, values=(
                folder.path,
                folder.folder_type.capitalize(),
                folder.message_count,
                folder.subfolder_count,
            ))

    def _show_about(self):
        win = tk.Toplevel(self.root)
        win.title("About Outlook Tool")
        win.geometry("380x400")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Outlook PST/OST Tool",
                  font=('Segoe UI', 14, 'bold')).pack(pady=(0, 5))
        ttk.Label(frame, text="Created by Randy Northrup",
                  font=('Segoe UI', 10)).pack(pady=(0, 15))

        about_text = (
            "A standalone tool for reading and extracting data from\n"
            "Microsoft Outlook PST and OST files.\n"
            "Does NOT require Outlook to be installed.\n\n"
            "Features:\n"
            "  - Browse folder structure\n"
            "  - View emails, contacts, calendar items\n"
            "  - Export calendar to ICS / CSV\n"
            "  - Export contacts to VCF / CSV\n"
            "  - Export emails to EML / CSV\n"
            "  - Extract file attachments\n"
            "  - Search across all items\n"
            "  - View MAPI properties"
        )
        ttk.Label(frame, text=about_text, justify=tk.LEFT).pack(anchor=tk.W)

        ttk.Button(frame, text="OK", command=win.destroy).pack(pady=(15, 0))

    # ======================== Utilities ========================

    def _sort_column(self, col: str):
        """Sort item list by a column, toggling direction."""
        reverse = self._sort_reverse.get(col, False)
        items = [(self.item_list.set(iid, col), iid)
                 for iid in self.item_list.get_children()]
        items.sort(reverse=reverse)
        for index, (_, iid) in enumerate(items):
            self.item_list.move(iid, '', index)
        self._sort_reverse[col] = not reverse

    @staticmethod
    def _format_size(size_bytes: int) -> str:
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 ** 2:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 ** 3:
            return f"{size_bytes / 1024 ** 2:.1f} MB"
        else:
            return f"{size_bytes / 1024 ** 3:.2f} GB"

    def _on_close(self):
        if self.reader:
            self.reader.close()
        self.root.destroy()


def _show_splash(root: tk.Tk) -> tk.Toplevel:
    """Show a simple splash screen while the app loads."""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)

    width, height = 400, 200
    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2
    splash.geometry(f"{width}x{height}+{x}+{y}")

    splash.configure(bg='#2b579a')

    tk.Label(
        splash, text="Outlook PST/OST Tool",
        font=('Segoe UI', 18, 'bold'), fg='white', bg='#2b579a'
    ).pack(expand=True, pady=(40, 5))

    tk.Label(
        splash, text="by Randy Northrup",
        font=('Segoe UI', 10), fg='#b0c4de', bg='#2b579a'
    ).pack()

    tk.Label(
        splash, text="Loading...",
        font=('Segoe UI', 10), fg='#d0d8e8', bg='#2b579a'
    ).pack(expand=True, pady=(5, 40))

    splash.update()
    return splash


def main():
    root = tk.Tk()
    root.withdraw()

    splash = _show_splash(root)

    OutlookToolApp(root)

    def finish_loading():
        splash.destroy()
        root.deiconify()

    root.after(1500, finish_loading)
    root.mainloop()


if __name__ == '__main__':
    main()
