"""
PST/OST file handler using pypff (libpff).
Reads Outlook PST and OST files without requiring Microsoft Outlook.
"""

import contextlib
import html as _html_lib
import os
import re
from collections.abc import Iterator
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

try:
    import pypff
    PYPFF_AVAILABLE = True
except ImportError:
    pypff = None
    PYPFF_AVAILABLE = False


# MAPI value types (low 16 bits of a property tag)
VT_I2 = 0x0002
VT_I4 = 0x0003
VT_R4 = 0x0004
VT_R8 = 0x0005
VT_BOOL = 0x000B
VT_STRING8 = 0x001E
VT_UNICODE = 0x001F
VT_SYSTIME = 0x0040


# Standard MAPI property IDs (these ARE what pypff's entry.entry_type returns directly).
MAPI_PROPS = {
    # Message properties
    0x001A: 'message_class',
    0x0037: 'subject',
    0x0042: 'sent_representing_name',
    0x0070: 'conversation_topic',
    0x0C1A: 'sender_name',
    0x0C1F: 'sender_email',
    0x0065: 'sent_representing_email',
    0x0E04: 'display_to',
    0x0E03: 'display_cc',
    0x0E02: 'display_bcc',
    0x0039: 'client_submit_time',
    0x0E06: 'message_delivery_time',
    0x3007: 'creation_time',
    0x3008: 'last_modification_time',
    0x0E07: 'message_flags',
    0x0E08: 'message_size',
    # NB: the message body (0x1000/0x1013) and transport headers (0x007D) are read
    # via pypff's plain_text_body/html_body accessors instead, to avoid duplicating
    # large strings into every item's properties dict.
    0x0060: 'std_start_date',   # PidTagStartDate
    0x0061: 'std_end_date',     # PidTagEndDate

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
}

# Named MAPI properties (assigned in the 0x8000+ range, per store). The IDs below are
# the canonical assignment Outlook uses for the PSETID_Appointment property set and are
# stable across Outlook-generated PST/OST stores. Resolution falls back across several
# equivalent slots (see _resolve_calendar_dates) for robustness.
NAMED_PROPS = {
    0x8004: 'appt_start',     # PidLidAppointmentStartWhole
    0x8005: 'appt_end',       # PidLidAppointmentEndWhole
    0x8006: 'common_start',   # PidLidCommonStart
    0x8007: 'common_end',     # PidLidCommonEnd
    0x80B2: 'clip_start',     # PidLidClipStart
    0x80B3: 'clip_end',       # PidLidClipEnd
    0x802A: 'location',       # PidLidLocation
    0x803A: 'duration',       # PidLidAppointmentDuration (minutes)
    0x8117: 'tz_desc',        # PidLidTimeZoneDescription
    0x8105: 'all_day_event',  # PidLidAppointmentSubType (all-day flag)
}

# Outlook stores "no date" appointment slots as 4501-01-01; reject anything in
# that far-future range when resolving calendar start/end times.
OUTLOOK_NULL_YEAR = 4000

# Attachment record-set property IDs.
ATT_LONG_FILENAME = 0x3707
ATT_FILENAME = 0x3704
ATT_MIME_TAG = 0x370E
ATT_CONTENT_ID = 0x3712
ATT_RENDER_POSITION = 0x370B


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
    date: datetime | None = None
    body_text: str = ''
    body_html: str = ''
    has_attachments: bool = False
    attachment_count: int = 0
    message_class: str = ''
    folder_path: str = ''
    properties: dict[str, Any] = field(default_factory=dict)
    _pypff_message: Any = field(default=None, repr=False)


@dataclass
class AttachmentInfo:
    """Represents an attachment."""
    name: str
    size: int
    mime_type: str = 'application/octet-stream'
    content_id: str = ''
    is_inline: bool = False
    _pypff_attachment: Any = field(default=None, repr=False)


_EMAIL_FOLDER_KEYWORDS = (
    'inbox', 'sent', 'draft', 'outbox', 'junk', 'deleted', 'archive')


def detect_folder_type(folder_name: str) -> str:
    """Detect folder type based on folder name."""
    name_lower = (folder_name or '').lower()
    if any(kw in name_lower for kw in _EMAIL_FOLDER_KEYWORDS):
        return 'email'
    if 'contact' in name_lower or 'address' in name_lower:
        return 'contacts'
    if 'calendar' in name_lower or 'appointment' in name_lower:
        return 'calendar'
    if 'task' in name_lower or 'todo' in name_lower:
        return 'tasks'
    if 'note' in name_lower or 'journal' in name_lower:
        return 'notes'
    return 'other'


def detect_message_type(message_class: str, folder_type: str) -> str:
    """Detect message type from message class or folder type."""
    if message_class:
        mc = message_class.lower()
        if 'contact' in mc:
            return 'contact'
        if 'appointment' in mc or 'calendar' in mc or 'meeting' in mc or 'schedule' in mc:
            return 'calendar'
        if 'task' in mc:
            return 'task'
        if 'stickynote' in mc:
            return 'note'
        if 'note' in mc:
            return 'email'
    # Fall back to folder type
    type_map = {
        'contacts': 'contact',
        'calendar': 'calendar',
        'tasks': 'task',
        'notes': 'note',
    }
    return type_map.get(folder_type, 'email')


def _collect_email_candidate(value: str, out: list[str]) -> None:
    """Append ``value`` to ``out`` if it looks like a bare e-mail address."""
    cand = value.strip().strip('<>')
    if (cand and ' ' not in cand and '@' in cand
            and '.' in cand.rsplit('@', 1)[-1] and cand not in out):
        out.append(cand)


def _to_text(val: Any) -> str:
    """Decode bytes (pypff returns bodies as bytes) or stringify to text."""
    if val is None:
        return ''
    if isinstance(val, bytes):
        for enc in ('utf-8', 'cp1252', 'latin-1'):
            try:
                return val.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return val.decode('utf-8', errors='replace')
    return str(val)


# Regexes for converting an HTML body to readable plain text (used by the
# preview/detail views and by text/CSV exports). Many modern messages — OST in
# particular — carry only an HTML body, so without this the UI would show markup.
_RE_HTML_DROP = re.compile(r'<(script|style|head)\b[^>]*>.*?</\1>', re.I | re.S)
_RE_HTML_BREAK = re.compile(r'<br\b[^>]*>', re.I)
_RE_HTML_BLOCK = re.compile(
    r'</(p|div|tr|h[1-6]|li|ul|ol|table|blockquote)\s*>', re.I)
_RE_HTML_TAG = re.compile(r'<[^>]+>')
_RE_INLINE_WS = re.compile(r'[ \t\f\v]+')
_RE_EXTRA_BLANKLINES = re.compile(r'\n\s*\n\s*\n+')


def html_to_text(html_str: str) -> str:
    """Convert an HTML body to readable plain text (no tags, entities decoded)."""
    if not html_str:
        return ''
    text = _RE_HTML_DROP.sub(' ', html_str)
    text = _RE_HTML_BREAK.sub('\n', text)
    text = _RE_HTML_BLOCK.sub('\n', text)
    text = _RE_HTML_TAG.sub('', text)
    text = _html_lib.unescape(text)
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = _RE_INLINE_WS.sub(' ', text)
    text = '\n'.join(line.strip() for line in text.split('\n'))
    text = _RE_EXTRA_BLANKLINES.sub('\n\n', text)
    return text.strip()


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
        self._folder_tree: list[FolderInfo] | None = None

    def open(self, filepath: str) -> None:
        """Open a PST or OST file."""
        self.close()
        assert pypff is not None
        self._file = pypff.file()
        self._file.open(filepath)
        self._filepath = filepath
        self._folder_tree = None

    def close(self) -> None:
        """Close the current file."""
        if self._file is not None:
            with contextlib.suppress(Exception):
                self._file.close()
            self._file = None
            self._filepath = None
            self._folder_tree = None

    @property
    def is_open(self) -> bool:
        return self._file is not None

    @property
    def filepath(self) -> str | None:
        return self._filepath

    @property
    def filename(self) -> str | None:
        return os.path.basename(self._filepath) if self._filepath else None

    def get_folder_tree(self) -> list[FolderInfo]:
        """Get the complete folder hierarchy (cached while the file stays open)."""
        if not self.is_open or self._file is None:
            return []
        if self._folder_tree is None:
            root = self._file.get_root_folder()
            folders: list[FolderInfo] = []
            self._build_folder_tree(root, '', folders)
            self._folder_tree = folders
        return self._folder_tree

    def _build_folder_tree(self, folder, parent_path: str, folders: list[FolderInfo]) -> None:
        """Recursively build the folder tree."""
        try:
            name = folder.name or '(Root)'
            current_path = f"{parent_path}/{name}" if parent_path else name

            msg_count = 0
            sub_count = 0
            with contextlib.suppress(Exception):
                msg_count = folder.number_of_sub_messages
            with contextlib.suppress(Exception):
                sub_count = folder.number_of_sub_folders

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

    def get_messages(self, folder_info: FolderInfo) -> list[MessageItem]:
        """Get all messages/items in a folder (non-recursive)."""
        messages: list[MessageItem] = []
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

    def iter_messages_recursive(
        self, folder_info: FolderInfo
    ) -> Iterator[tuple[str, MessageItem]]:
        """Yield (folder_path, MessageItem) for a folder and all of its descendants."""
        base = folder_info._pypff_folder
        if base is None:
            return

        def walk(folder, path: str, ftype: str):
            fi = FolderInfo(
                name=path.rsplit('/', maxsplit=1)[-1], path=path, message_count=0,
                subfolder_count=0, folder_type=ftype, _pypff_folder=folder,
            )
            for msg in self.get_messages(fi):
                yield path, msg
            try:
                sub_count = folder.number_of_sub_folders
            except Exception:
                sub_count = 0
            for i in range(sub_count):
                try:
                    sf = folder.get_sub_folder(i)
                    nm = sf.name or f'folder_{i}'
                    yield from walk(sf, f"{path}/{nm}", detect_folder_type(nm))
                except Exception:
                    pass

        yield from walk(base, folder_info.path, folder_info.folder_type)

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
        subject = _to_text(self._safe_get(msg, 'subject', ''))
        sender_name = _to_text(self._safe_get(msg, 'sender_name', ''))
        body_text = _to_text(self._safe_get(msg, 'plain_text_body', ''))
        body_html = _to_text(self._safe_get(msg, 'html_body', ''))

        delivery_time = None
        try:
            delivery_time = msg.delivery_time
        except Exception:
            with contextlib.suppress(Exception):
                delivery_time = msg.creation_time

        attachment_count = 0
        with contextlib.suppress(Exception):
            attachment_count = msg.number_of_attachments

        # Extract MAPI properties from record sets (standard + named).
        properties: dict[str, Any] = {}
        message_class = ''
        try:
            properties = self._get_mapi_properties(msg)
            message_class = _to_text(properties.get('message_class', ''))
        except Exception:
            pass

        item_type = detect_message_type(message_class, folder_info.folder_type)

        # Resolve calendar start/end/location for calendar items.
        if item_type == 'calendar':
            start, end = self._resolve_calendar_dates(properties, delivery_time)
            if start is not None:
                properties['start_date'] = start
            if end is not None:
                properties['end_date'] = end

        recipients = _to_text(properties.get('display_to', ''))
        cc = _to_text(properties.get('display_cc', ''))
        bcc = _to_text(properties.get('display_bcc', ''))
        sender_email = (_to_text(properties.get('sender_email', ''))
                        or _to_text(properties.get('sent_representing_email', '')))

        return MessageItem(
            item_type=item_type,
            subject=subject or _to_text(properties.get('subject', '')),
            sender_name=(sender_name
                         or _to_text(properties.get('sent_representing_name', ''))),
            sender_email=sender_email,
            recipients=recipients,
            cc=cc,
            bcc=bcc,
            date=delivery_time,
            body_text=body_text,
            body_html=body_html,
            has_attachments=attachment_count > 0,
            attachment_count=attachment_count,
            message_class=message_class,
            folder_path=folder_info.path,
            properties=properties,
            _pypff_message=msg,
        )

    @staticmethod
    def _read_entry_value(entry) -> Any:
        """Read a record-set entry's value according to its MAPI value type."""
        try:
            vt = entry.value_type
        except Exception:
            vt = None
        try:
            if vt in (VT_UNICODE, VT_STRING8):
                return entry.data_as_string
            if vt == VT_SYSTIME:
                return entry.data_as_datetime
            if vt in (VT_I2, VT_I4):
                return entry.data_as_integer
            if vt == VT_BOOL:
                return entry.data_as_boolean
            if vt in (VT_R4, VT_R8):
                return entry.data_as_floating_point
        except Exception:
            pass
        # Fallbacks for unknown/binary types.
        try:
            return entry.data_as_string
        except Exception:
            return None

    def _get_mapi_properties(self, msg) -> dict[str, Any]:
        """Extract MAPI properties (standard + named) from message record sets.

        Only mapped properties are fully decoded. Unmapped entries are skipped
        unless they are strings, which are scanned cheaply for contact e-mail
        addresses (those live in store-specific named props that vary per file).
        """
        properties: dict[str, Any] = {}
        candidates: list[str] = []
        try:
            num_sets = msg.number_of_record_sets
        except Exception:
            num_sets = 0

        for i in range(num_sets):
            try:
                record_set = msg.get_record_set(i)
                num_entries = record_set.number_of_entries
            except Exception:
                continue
            for j in range(num_entries):
                try:
                    entry = record_set.get_entry(j)
                    prop_id = entry.entry_type  # already the 16-bit property id
                except Exception:
                    continue

                name = MAPI_PROPS.get(prop_id) or NAMED_PROPS.get(prop_id)
                if name is not None:
                    if name in properties:
                        continue
                    value = self._read_entry_value(entry)
                    if value is not None and value != '':
                        properties[name] = value
                        if isinstance(value, str):
                            _collect_email_candidate(value, candidates)
                else:
                    # Unmapped: only strings can carry a contact e-mail; avoid the
                    # cost of decoding datetimes/ints we would otherwise discard.
                    try:
                        if entry.value_type not in (VT_UNICODE, VT_STRING8):
                            continue
                        sval = entry.data_as_string
                    except Exception:
                        continue
                    if sval:
                        _collect_email_candidate(sval, candidates)

        if candidates:
            properties['_email_candidates'] = candidates
        return properties

    @staticmethod
    def _resolve_calendar_dates(properties: dict[str, Any], fallback: datetime | None):
        """Resolve appointment start/end from named props, with robust fallbacks."""
        def pick(*keys):
            for k in keys:
                v = properties.get(k)
                # 4501-01-01 is Outlook's "no date" sentinel; reject it.
                if isinstance(v, datetime) and v.year < OUTLOOK_NULL_YEAR:
                    return v
            return None

        start = pick('appt_start', 'common_start', 'clip_start', 'std_start_date')
        end = pick('appt_end', 'common_end', 'clip_end', 'std_end_date')

        if start is None:
            start = fallback
        if end is None and start is not None:
            dur = properties.get('duration')
            end = (start + timedelta(minutes=dur)
                   if isinstance(dur, int) and dur > 0 else start)
        return start, end

    def get_attachments(self, message_item: MessageItem) -> list[AttachmentInfo]:
        """Get attachment info for a message (name/MIME/content-id resolved from props)."""
        attachments: list[AttachmentInfo] = []
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
            except Exception:
                continue

            meta = self._get_attachment_props(att)
            name = (meta.get(ATT_LONG_FILENAME) or meta.get(ATT_FILENAME)
                    or f'attachment_{i}')
            mime = meta.get(ATT_MIME_TAG) or 'application/octet-stream'
            content_id = (meta.get(ATT_CONTENT_ID) or '').strip('<>')
            render_pos = meta.get(ATT_RENDER_POSITION)
            is_inline = bool(content_id) and render_pos not in (None, -1)

            size = 0
            with contextlib.suppress(Exception):
                size = att.size

            attachments.append(AttachmentInfo(
                name=_to_text(name),
                size=size,
                mime_type=_to_text(mime),
                content_id=_to_text(content_id),
                is_inline=is_inline,
                _pypff_attachment=att,
            ))

        return attachments

    def _get_attachment_props(self, att) -> dict[int, Any]:
        """Read selected attachment record-set properties keyed by property id."""
        props: dict[int, Any] = {}
        wanted = {ATT_LONG_FILENAME, ATT_FILENAME, ATT_MIME_TAG,
                  ATT_CONTENT_ID, ATT_RENDER_POSITION}
        try:
            num_sets = att.number_of_record_sets
        except Exception:
            return props
        for i in range(num_sets):
            try:
                rs = att.get_record_set(i)
                num = rs.number_of_entries
            except Exception:
                continue
            for j in range(num):
                try:
                    entry = rs.get_entry(j)
                    pid = entry.entry_type
                except Exception:
                    continue
                if pid in wanted and pid not in props:
                    props[pid] = self._read_entry_value(entry)
        return props

    def read_attachment_bytes(self, attachment: AttachmentInfo) -> bytes | None:
        """Read the raw bytes of an attachment. Returns None if unreadable."""
        att = attachment._pypff_attachment
        if att is None:
            return None
        try:
            size = att.size
        except Exception:
            return None
        if size is None or size <= 0:
            return b''
        try:
            return att.read_buffer(size)
        except Exception:
            # e.g. message-type attachments or corrupt local descriptors
            return None

    def save_attachment(self, attachment: AttachmentInfo, output_path: str) -> bool:
        """Save an attachment to disk."""
        data = self.read_attachment_bytes(attachment)
        if data is None:
            return False
        try:
            with open(output_path, 'wb') as f:
                f.write(data)
            return True
        except Exception:
            return False

    def get_all_items_by_type(self, item_type: str) -> list[MessageItem]:
        """Get all items of a specific type across all folders."""
        items: list[MessageItem] = []
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
