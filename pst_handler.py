"""
PST/OST file handler using pypff (libpff).
Reads Outlook PST and OST files without requiring Microsoft Outlook.
"""

import os
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any

try:
    import pypff
    PYPFF_AVAILABLE = True
except ImportError:
    pypff = None
    PYPFF_AVAILABLE = False


# MAPI Property IDs (high 16 bits of the property tag)
MAPI_PROPS = {
    # Message properties
    0x001A: 'message_class',
    0x0037: 'subject',
    0x1000: 'body',
    0x1013: 'body_html',
    0x0E04: 'display_to',
    0x0E03: 'display_cc',
    0x0E02: 'display_bcc',
    0x0C1A: 'sender_name',
    0x0C1F: 'sender_email',
    0x0065: 'sent_representing_email',
    0x0039: 'client_submit_time',
    0x0E06: 'message_delivery_time',
    0x3007: 'creation_time',
    0x3008: 'last_modification_time',
    0x0E07: 'message_flags',
    0x0E08: 'message_size',
    0x0070: 'conversation_topic',

    # Contact properties
    0x3001: 'display_name',
    0x3A06: 'given_name',
    0x3A11: 'surname',
    0x3A16: 'company_name',
    0x3A17: 'title',
    0x3A08: 'business_phone',
    0x3A09: 'home_phone',
    0x3A1C: 'mobile_phone',
    0x3A15: 'postal_address',
    0x39FE: 'smtp_address',
    0x3A00: 'account',
    0x3A4F: 'nickname',
    0x3A05: 'generation',
    0x3A18: 'department_name',
    0x3A19: 'office_location',
    0x3A24: 'business_address_city',
    0x3A25: 'business_address_state',
    0x3A26: 'business_address_postal_code',
    0x3A27: 'business_address_country',
    0x3A28: 'business_address_street',
    0x3A29: 'home_address_city',
    0x3A2A: 'home_address_state',
    0x3A2B: 'home_address_postal_code',
    0x3A2C: 'home_address_country',
    0x3A2D: 'home_address_street',
    0x3A44: 'middle_name',
    0x3A45: 'display_name_prefix',

    # Calendar properties
    0x0060: 'start_date',
    0x0061: 'end_date',
}


@dataclass
class FolderInfo:
    """Represents a folder in the PST/OST file."""
    name: str
    path: str
    message_count: int
    subfolder_count: int
    folder_type: str  # 'email', 'contacts', 'calendar', 'tasks', 'notes', 'other'
    _pypff_folder: Any = field(default=None, repr=False)


@dataclass
class MessageItem:
    """Represents any item (email, contact, calendar, etc.) in the PST/OST."""
    item_type: str  # 'email', 'contact', 'calendar', 'task', 'note', 'other'
    subject: str = ''
    sender_name: str = ''
    sender_email: str = ''
    recipients: str = ''
    cc: str = ''
    bcc: str = ''
    date: Optional[datetime] = None
    body_text: str = ''
    body_html: str = ''
    has_attachments: bool = False
    attachment_count: int = 0
    message_class: str = ''
    folder_path: str = ''
    properties: Dict[str, Any] = field(default_factory=dict)
    _pypff_message: Any = field(default=None, repr=False)


@dataclass
class AttachmentInfo:
    """Represents an attachment."""
    name: str
    size: int
    _pypff_attachment: Any = field(default=None, repr=False)


def detect_folder_type(folder_name: str) -> str:
    """Detect folder type based on folder name."""
    name_lower = (folder_name or '').lower()
    if any(kw in name_lower for kw in ['inbox', 'sent', 'draft', 'outbox', 'junk', 'deleted', 'archive']):
        return 'email'
    elif 'contact' in name_lower or 'address' in name_lower:
        return 'contacts'
    elif 'calendar' in name_lower or 'appointment' in name_lower:
        return 'calendar'
    elif 'task' in name_lower or 'todo' in name_lower:
        return 'tasks'
    elif 'note' in name_lower or 'journal' in name_lower:
        return 'notes'
    return 'other'


def detect_message_type(message_class: str, folder_type: str) -> str:
    """Detect message type from message class or folder type."""
    if message_class:
        mc = message_class.lower()
        if 'contact' in mc:
            return 'contact'
        elif 'appointment' in mc or 'calendar' in mc or 'meeting' in mc or 'schedule' in mc:
            return 'calendar'
        elif 'task' in mc:
            return 'task'
        elif 'stickynote' in mc:
            return 'note'
        elif 'note' in mc:
            return 'email'
    # Fall back to folder type
    type_map = {
        'contacts': 'contact',
        'calendar': 'calendar',
        'tasks': 'task',
        'notes': 'note',
    }
    return type_map.get(folder_type, 'email')


class PstReader:
    """Reads PST/OST files using pypff (libpff)."""

    def __init__(self):
        if not PYPFF_AVAILABLE:
            raise ImportError(
                "pypff (libpff-python) is required but not installed.\n\n"
                "Install options:\n"
                "  pip install libpff-python\n"
                "  conda install -c conda-forge libpff-python\n\n"
                "If pip fails on Windows, try conda or see:\n"
                "  https://github.com/libyal/libpff"
            )
        self._file = None
        self._filepath = None

    def open(self, filepath: str) -> None:
        """Open a PST or OST file."""
        self.close()
        assert pypff is not None
        self._file = pypff.file()
        self._file.open(filepath)
        self._filepath = filepath

    def close(self) -> None:
        """Close the current file."""
        if self._file is not None:
            try:
                self._file.close()
            except Exception:
                pass
            self._file = None
            self._filepath = None

    @property
    def is_open(self) -> bool:
        return self._file is not None

    @property
    def filepath(self) -> Optional[str]:
        return self._filepath

    @property
    def filename(self) -> Optional[str]:
        return os.path.basename(self._filepath) if self._filepath else None

    def get_folder_tree(self) -> List[FolderInfo]:
        """Get the complete folder hierarchy."""
        if not self.is_open or self._file is None:
            return []
        root = self._file.get_root_folder()
        folders = []
        self._build_folder_tree(root, '', folders)
        return folders

    def _build_folder_tree(self, folder, parent_path: str, folders: List[FolderInfo]) -> None:
        """Recursively build the folder tree."""
        try:
            name = folder.name or '(Root)'
            current_path = f"{parent_path}/{name}" if parent_path else name

            msg_count = 0
            sub_count = 0
            try:
                msg_count = folder.number_of_sub_messages
            except Exception:
                pass
            try:
                sub_count = folder.number_of_sub_folders
            except Exception:
                pass

            folder_type = detect_folder_type(name)
            info = FolderInfo(
                name=name,
                path=current_path,
                message_count=msg_count,
                subfolder_count=sub_count,
                folder_type=folder_type,
                _pypff_folder=folder,
            )
            folders.append(info)

            for i in range(sub_count):
                try:
                    subfolder = folder.get_sub_folder(i)
                    self._build_folder_tree(subfolder, current_path, folders)
                except Exception:
                    pass
        except Exception:
            pass

    def get_messages(self, folder_info: FolderInfo) -> List[MessageItem]:
        """Get all messages/items in a folder."""
        messages = []
        folder = folder_info._pypff_folder
        if folder is None:
            return messages

        try:
            count = folder.number_of_sub_messages
        except Exception:
            return messages

        for i in range(count):
            try:
                msg = folder.get_sub_message(i)
                item = self._parse_message(msg, folder_info)
                messages.append(item)
            except Exception:
                pass

        return messages

    @staticmethod
    def _safe_get(obj, attr, default=''):
        """Safely get an attribute from a pypff object."""
        try:
            val = getattr(obj, attr, None)
            return val if val is not None else default
        except Exception:
            return default

    def _parse_message(self, msg, folder_info: FolderInfo) -> MessageItem:
        """Parse a pypff message into a MessageItem."""
        subject = self._safe_get(msg, 'subject', '')
        sender_name = self._safe_get(msg, 'sender_name', '')
        body_text = self._safe_get(msg, 'plain_text_body', '')
        body_html = self._safe_get(msg, 'html_body', '')

        delivery_time = None
        try:
            delivery_time = msg.delivery_time
        except Exception:
            try:
                delivery_time = msg.creation_time
            except Exception:
                pass

        attachment_count = 0
        try:
            attachment_count = msg.number_of_attachments
        except Exception:
            pass

        # Extract MAPI properties from record sets
        properties = {}
        message_class = ''
        try:
            properties = self._get_mapi_properties(msg)
            message_class = properties.get('message_class', '')
        except Exception:
            pass

        item_type = detect_message_type(message_class, folder_info.folder_type)

        recipients = properties.get('display_to', '')
        cc = properties.get('display_cc', '')
        bcc = properties.get('display_bcc', '')
        sender_email = properties.get('sender_email', '') or properties.get('sent_representing_email', '')

        return MessageItem(
            item_type=item_type,
            subject=subject or properties.get('subject', ''),
            sender_name=sender_name or properties.get('sender_name', ''),
            sender_email=sender_email,
            recipients=recipients,
            cc=cc,
            bcc=bcc,
            date=delivery_time,
            body_text=body_text or properties.get('body', ''),
            body_html=body_html or properties.get('body_html', ''),
            has_attachments=attachment_count > 0,
            attachment_count=attachment_count,
            message_class=message_class,
            folder_path=folder_info.path,
            properties=properties,
            _pypff_message=msg,
        )

    def _get_mapi_properties(self, msg) -> Dict[str, Any]:
        """Extract MAPI properties from message record sets."""
        properties = {}
        try:
            num_sets = msg.number_of_record_sets
        except Exception:
            return properties

        for i in range(num_sets):
            try:
                record_set = msg.get_record_set(i)
                num_entries = record_set.number_of_entries
                for j in range(num_entries):
                    try:
                        entry = record_set.get_entry(j)
                        entry_type = entry.entry_type
                        # Property ID is the high 16 bits of the tag
                        prop_id = (entry_type >> 16) & 0xFFFF

                        if prop_id in MAPI_PROPS:
                            prop_name = MAPI_PROPS[prop_id]
                            try:
                                value = entry.data_as_string
                            except Exception:
                                try:
                                    value = entry.data
                                except Exception:
                                    value = None
                            if value is not None:
                                properties[prop_name] = value
                    except Exception:
                        continue
            except Exception:
                continue

        return properties

    def get_attachments(self, message_item: MessageItem) -> List[AttachmentInfo]:
        """Get attachment info for a message."""
        attachments = []
        msg = message_item._pypff_message
        if msg is None:
            return attachments

        try:
            count = msg.number_of_attachments
        except Exception:
            return attachments

        for i in range(count):
            try:
                att = msg.get_attachment(i)
                name = self._safe_get(att, 'name', f'attachment_{i}')
                if not name:
                    name = f'attachment_{i}'
                size = 0
                try:
                    size = att.size
                except Exception:
                    pass
                attachments.append(AttachmentInfo(
                    name=name,
                    size=size,
                    _pypff_attachment=att,
                ))
            except Exception:
                pass

        return attachments

    def save_attachment(self, attachment: AttachmentInfo, output_path: str) -> bool:
        """Save an attachment to disk."""
        try:
            att = attachment._pypff_attachment
            data = att.read_buffer(att.size)
            with open(output_path, 'wb') as f:
                f.write(data)
            return True
        except Exception:
            return False

    def get_all_items_by_type(self, item_type: str) -> List[MessageItem]:
        """Get all items of a specific type across all folders."""
        items = []
        folders = self.get_folder_tree()
        for folder in folders:
            try:
                messages = self.get_messages(folder)
                for msg in messages:
                    if msg.item_type == item_type:
                        items.append(msg)
            except Exception:
                pass
        return items
