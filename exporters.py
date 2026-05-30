"""
Export functionality for PST/OST data.

Per-item renderers (txt / html / eml / pdf) plus engines for exporting a
selection of items or an entire folder subtree in a chosen format, with
attachments embedded where the format allows (eml MIME parts, html data-URIs,
pdf embedded files) and otherwise written to a per-item ``*_attachments`` folder.
Also exports calendar (ICS/CSV) and contacts (VCF/CSV).
"""

import base64
import contextlib
import csv
import html as _htmlmod
import logging
import os
import re
from collections.abc import Callable
from datetime import datetime, timezone
from email.message import EmailMessage
from email.utils import format_datetime

from pst_handler import MessageItem, PstReader, html_to_text

# xhtml2pdf is noisy; keep it quiet.
logging.getLogger('xhtml2pdf').setLevel(logging.ERROR)
logging.getLogger('xhtml2pdf.w3c').setLevel(logging.ERROR)

# Cap for embedding an attachment as a base64 data-URI in HTML/PDF (bytes).
MAX_EMBED_BYTES = 25 * 1024 * 1024

ProgressFn = Callable[[int, int, str], None] | None


# ===================== small helpers =====================

def _safe_str(value, default=''):
    if value is None:
        return default
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d %H:%M:%S')
    return str(value)


def _sanitize_filename(name: str) -> str:
    invalid_chars = '<>:"/\\|?*\x00'
    name = name or ''
    for c in invalid_chars:
        name = name.replace(c, '_')
    name = name.replace('\r', ' ').replace('\n', ' ').strip(' .')
    return name[:150] if name else 'unnamed'


def _unique_filepath(filepath: str) -> str:
    if not os.path.exists(filepath):
        return filepath
    base, ext = os.path.splitext(filepath)
    counter = 1
    while os.path.exists(filepath):
        filepath = f"{base}_{counter}{ext}"
        counter += 1
    return filepath


_BYTES_PER_KB = 1024
_BYTES_PER_MB = 1024 ** 2
_BYTES_PER_GB = 1024 ** 3


def format_size(size_bytes: int) -> str:
    """Human-readable byte size (e.g. ``1.5 KB``)."""
    size_bytes = size_bytes or 0
    if size_bytes < _BYTES_PER_KB:
        return f"{size_bytes} B"
    if size_bytes < _BYTES_PER_MB:
        return f"{size_bytes / _BYTES_PER_KB:.1f} KB"
    if size_bytes < _BYTES_PER_GB:
        return f"{size_bytes / _BYTES_PER_MB:.1f} MB"
    return f"{size_bytes / _BYTES_PER_GB:.2f} GB"


def _esc(text) -> str:
    return _htmlmod.escape(_safe_str(text))


def _hdr(value) -> str:
    """Sanitize a string for use as an RFC-822 header value (no CR/LF)."""
    return re.sub(r'[\r\n\t]+', ' ', _safe_str(value)).strip()


def _item_basename(item: MessageItem) -> str:
    """A stable, readable base filename for an item (no extension)."""
    date_str = item.date.strftime('%Y%m%d_%H%M%S') if item.date else 'nodate'
    subject = _sanitize_filename(item.subject or f'{item.item_type}_item')
    return f"{date_str}_{subject}"


def resolve_contact_email(item: MessageItem) -> str:
    """Best contact e-mail, excluding the mailbox owner's own address."""
    props = item.properties
    owner = {props.get('sender_email'), props.get('sent_representing_email'),
             item.sender_email}
    owner = {o for o in owner if o}
    for cand in props.get('_email_candidates', []):
        if cand not in owner:
            return cand
    smtp = props.get('smtp_address')
    if smtp and smtp not in owner:
        return smtp
    return ''


# ===================== item -> header rows =====================

def _header_rows(item: MessageItem) -> list[tuple[str, str]]:
    """Label/value pairs describing an item, tailored to its type."""
    rows: list[tuple[str, str]] = []
    if item.item_type == 'calendar':
        rows.append(('Subject', item.subject))
        rows.append(('Start', _safe_str(item.properties.get('start_date'))))
        rows.append(('End', _safe_str(item.properties.get('end_date'))))
        rows.append(('Location', _safe_str(item.properties.get('location'))))
        rows.append(('Organizer', item.sender_name or item.sender_email))
    elif item.item_type == 'contact':
        p = item.properties
        rows.append(('Name', p.get('display_name') or item.subject))
        rows.append(('Email', resolve_contact_email(item)))
        rows.append(('Business Phone', p.get('business_phone', '')))
        rows.append(('Home Phone', p.get('home_phone', '')))
        rows.append(('Mobile', p.get('mobile_phone', '')))
        rows.append(('Company', p.get('company_name', '')))
        rows.append(('Title', p.get('title', '')))
        rows.append(('Address', p.get('postal_address', '')))
    else:
        from_str = item.sender_name or ''
        if item.sender_email:
            from_str = f"{from_str} <{item.sender_email}>".strip()
        rows.append(('From', from_str))
        rows.append(('To', item.recipients))
        rows.append(('Cc', item.cc))
        rows.append(('Subject', item.subject))
        rows.append(('Date', _safe_str(item.date)))
    return [(k, v) for k, v in rows if v]


def _body_text_for(item: MessageItem) -> str:
    if item.body_text:
        return item.body_text
    if item.body_html:
        return html_to_text(item.body_html)
    return ''


# ===================== TXT =====================

def render_item_txt(item: MessageItem) -> str:
    lines = [f"{k}: {v}" for k, v in _header_rows(item)]
    if item.message_class:
        lines.append(f"Class: {item.message_class}")
    if item.has_attachments:
        lines.append(f"Attachments: {item.attachment_count}")
    lines.append('')
    lines.append('-' * 70)
    lines.append('')
    lines.append(_body_text_for(item))
    return '\n'.join(lines)


# ===================== HTML =====================

_HTML_STYLE = (
    "body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#111;"
    "margin:24px;}"
    "table.hdr{border-collapse:collapse;margin-bottom:12px;}"
    "table.hdr td{padding:2px 8px;vertical-align:top;}"
    "table.hdr td.k{font-weight:bold;color:#444;white-space:nowrap;}"
    "h1{font-size:18px;margin:0 0 4px 0;}"
    ".meta{color:#666;font-size:12px;}"
    "hr{border:none;border-top:1px solid #ccc;margin:12px 0;}"
    ".body{margin-top:12px;}"
    "pre{white-space:pre-wrap;font-family:Consolas,monospace;font-size:12px;}"
)


def _embed_inline_images(body_html: str, attachments, reader: PstReader) -> str:
    """Replace src="cid:..." references with base64 data-URIs."""
    cid_map = {}
    for att in attachments:
        if att.content_id:
            data = reader.read_attachment_bytes(att)
            if data is not None and len(data) <= MAX_EMBED_BYTES:
                b64 = base64.b64encode(data).decode('ascii')
                cid_map[att.content_id.lower()] = f"data:{att.mime_type};base64,{b64}"

    if not cid_map:
        return body_html

    def repl(m):
        cid = m.group(2).strip().strip('<>').lower()
        uri = cid_map.get(cid)
        return f'src={m.group(1)}{uri}{m.group(1)}' if uri else m.group(0)

    return re.sub(r'src=(["\'])cid:([^"\']+)\1', repl, body_html, flags=re.I)


def render_item_html(item: MessageItem, reader: PstReader | None = None,
                     embed: bool = True, embed_attachments: bool = True) -> str:
    """Render an item as a standalone HTML document.

    ``embed`` inlines ``cid:`` images as data-URIs. ``embed_attachments`` adds
    non-inline attachments as base64 download links; the PDF path disables it
    because pikepdf embeds those files into the PDF instead.
    """
    attachments = []
    if reader is not None and item.has_attachments:
        try:
            attachments = reader.get_attachments(item)
        except Exception:
            attachments = []

    header_html = '<table class="hdr">' + ''.join(
        f'<tr><td class="k">{_esc(k)}</td><td>{_esc(v)}</td></tr>'
        for k, v in _header_rows(item)) + '</table>'

    if item.body_html:
        body = item.body_html
        if embed and reader is not None and attachments:
            body = _embed_inline_images(body, attachments, reader)
        body_section = f'<div class="body">{body}</div>'
    else:
        body_section = f'<div class="body"><pre>{_esc(_body_text_for(item))}</pre></div>'

    att_section = ''
    non_inline = [a for a in attachments if not a.is_inline]
    if non_inline:
        items_html = []
        for att in non_inline:
            label = f"{_esc(att.name)} ({format_size(att.size)})"
            if embed_attachments and reader is not None:
                data = reader.read_attachment_bytes(att)
                if data is not None and len(data) <= MAX_EMBED_BYTES:
                    b64 = base64.b64encode(data).decode('ascii')
                    items_html.append(
                        f'<li><a download="{_esc(att.name)}" '
                        f'href="data:{att.mime_type};base64,{b64}">{label}</a></li>')
                    continue
            items_html.append(f'<li>{label}</li>')
        att_section = ('<hr><h3>Attachments</h3><ul>' + ''.join(items_html) + '</ul>')

    title = _esc(item.subject or f'({item.item_type})')
    return (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        f'<style>{_HTML_STYLE}</style><title>{title}</title></head><body>'
        f'<h1>{title}</h1>{header_html}<hr>{body_section}{att_section}'
        '</body></html>'
    )


# ===================== EML =====================

def build_eml_bytes(item: MessageItem, reader: PstReader | None = None) -> bytes:
    """Build an RFC-822 .eml with attachments embedded as MIME parts."""
    msg = EmailMessage()
    from_str = item.sender_name or ''
    if item.sender_email:
        from_str = (f"{from_str} <{item.sender_email}>".strip()
                    if from_str else item.sender_email)
    if from_str:
        msg['From'] = _hdr(from_str)
    if item.recipients:
        msg['To'] = _hdr(item.recipients)
    if item.cc:
        msg['Cc'] = _hdr(item.cc)
    if item.subject:
        msg['Subject'] = _hdr(item.subject)
    if item.date:
        with contextlib.suppress(Exception):
            msg['Date'] = format_datetime(item.date)
    if item.message_class:
        msg['X-Outlook-Message-Class'] = _hdr(item.message_class)
    if item.folder_path:
        msg['X-Outlook-Folder'] = _hdr(item.folder_path)

    text = item.body_text or _body_text_for(item) or ''
    msg.set_content(text if text else '(no text body)')
    if item.body_html:
        with contextlib.suppress(Exception):
            msg.add_alternative(item.body_html, subtype='html')

    if reader is not None and item.has_attachments:
        try:
            attachments = reader.get_attachments(item)
        except Exception:
            attachments = []
        for att in attachments:
            try:
                data = reader.read_attachment_bytes(att)
                if data is None:
                    continue
                mime = att.mime_type or 'application/octet-stream'
                if '/' in mime:
                    maintype, subtype = mime.split('/', 1)
                else:
                    maintype, subtype = 'application', 'octet-stream'
                # add_attachment expects an EmailMessage for message/multipart
                # types; attach raw bytes as a generic binary part instead.
                if maintype in ('message', 'multipart'):
                    maintype, subtype = 'application', 'octet-stream'
                msg.add_attachment(data, maintype=maintype, subtype=subtype,
                                   filename=att.name or 'attachment')
            except Exception:
                continue

    return msg.as_bytes()


# ===================== PDF =====================

def _html_to_pdf(html_str: str, out_path: str) -> bool:
    """Render an HTML string to a PDF file via xhtml2pdf. Returns success."""
    from xhtml2pdf import pisa
    try:
        with open(out_path, 'wb') as f:
            status = pisa.CreatePDF(src=html_str, dest=f, encoding='utf-8')
        # status is a pisaDocument whose .err counts rendering errors.
        err = getattr(status, 'err', 1)
        return not err and os.path.getsize(out_path) > 0
    except Exception:
        return False


def _embed_pdf_attachments(pdf_path: str, reader: PstReader,
                           item: MessageItem) -> list[str]:
    """Embed attachment files inside the PDF. Returns names that could not embed."""
    failed: list[str] = []
    if not item.has_attachments:
        return failed
    try:
        import pikepdf
    except Exception:
        return [a.name for a in reader.get_attachments(item)]

    try:
        attachments = reader.get_attachments(item)
    except Exception:
        return failed
    payload = []
    for att in attachments:
        data = reader.read_attachment_bytes(att)
        if data is None:
            failed.append(att.name)
            continue
        payload.append((att.name or 'attachment', data, att.mime_type))
    if not payload:
        return failed

    try:
        pdf = pikepdf.open(pdf_path, allow_overwriting_input=True)
        used = set()
        for name, data, mime in payload:
            key = name or 'attachment'
            n = key
            i = 1
            while n in used or n in pdf.attachments:
                stem, ext = os.path.splitext(key)
                n = f"{stem}_{i}{ext}"
                i += 1
            used.add(n)
            spec = pikepdf.AttachedFileSpec(
                pdf, data,
                description='',
                filename=n,
                mime_type=mime or 'application/octet-stream',
                creation_date='',
                mod_date='',
            )
            pdf.attachments[n] = spec
        pdf.save()
        pdf.close()
    except Exception:
        return [name for name, _, _ in payload]
    return failed


def render_item_pdf(item: MessageItem, reader: PstReader | None,
                    out_path: str) -> bool:
    """Render an item to PDF (embedding attachments). Always produces a file."""
    # Inline images are embedded for fidelity; non-inline attachments are left to
    # pikepdf (below) rather than bloating the HTML with base64 data-URIs.
    html_str = render_item_html(item, reader=reader, embed=True,
                                embed_attachments=False)
    ok = _html_to_pdf(html_str, out_path)
    if not ok:
        # Fallback: minimal, always-valid HTML from plain text.
        header = ''.join(f'<tr><td class="k">{_esc(k)}</td><td>{_esc(v)}</td></tr>'
                         for k, v in _header_rows(item))
        safe = (
            '<!DOCTYPE html><html><head><meta charset="utf-8">'
            f'<style>{_HTML_STYLE}</style></head><body>'
            f'<h1>{_esc(item.subject or item.item_type)}</h1>'
            f'<table class="hdr">{header}</table><hr>'
            f'<pre>{_esc(_body_text_for(item))}</pre></body></html>'
        )
        ok = _html_to_pdf(safe, out_path)
    if ok and reader is not None:
        _embed_pdf_attachments(out_path, reader, item)
    return ok


# ===================== per-email attachment folder (txt/csv) =====================

def _dump_attachments_folder(reader: PstReader, item: MessageItem,
                             folder: str) -> int:
    """Write an item's attachments into ``folder``. Returns count written."""
    if not item.has_attachments:
        return 0
    count = 0
    try:
        attachments = reader.get_attachments(item)
    except Exception:
        return 0
    if not attachments:
        return 0
    os.makedirs(folder, exist_ok=True)
    for att in attachments:
        name = _sanitize_filename(att.name or 'attachment')
        path = _unique_filepath(os.path.join(folder, name))
        if reader.save_attachment(att, path):
            count += 1
    return count


# ===================== unified single-item writer =====================

def write_item(item: MessageItem, reader: PstReader | None, fmt: str,
               out_dir: str) -> str:
    """Write one item in ``fmt`` into ``out_dir``. Returns the file path."""
    base = _item_basename(item)
    if fmt == 'txt':
        content = render_item_txt(item)  # build before opening the file
        path = _unique_filepath(os.path.join(out_dir, base + '.txt'))
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        if reader is not None and item.has_attachments:
            _dump_attachments_folder(reader, item,
                                     os.path.join(out_dir, base + '_attachments'))
        return path
    if fmt == 'html':
        content = render_item_html(item, reader=reader, embed=True)
        path = _unique_filepath(os.path.join(out_dir, base + '.html'))
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        return path
    if fmt == 'eml':
        data = build_eml_bytes(item, reader)
        path = _unique_filepath(os.path.join(out_dir, base + '.eml'))
        with open(path, 'wb') as f:
            f.write(data)
        return path
    if fmt == 'pdf':
        path = _unique_filepath(os.path.join(out_dir, base + '.pdf'))
        render_item_pdf(item, reader, path)
        return path
    raise ValueError(f"Unsupported per-file format: {fmt}")


# ===================== CSV (generic) =====================

_CSV_HEADER = ['Type', 'Subject', 'From / Name', 'To', 'Date', 'Location',
               'Folder', 'Attachments', 'Body Preview']


def _csv_row(item: MessageItem) -> list[str]:
    if item.item_type == 'contact':
        from_name = item.properties.get('display_name', '') or item.subject
        to = resolve_contact_email(item)
    else:
        from_name = item.sender_name or item.properties.get('display_name', '')
        to = item.recipients
    att_names = ''
    if item.has_attachments:
        att_names = f"{item.attachment_count} attachment(s)"
    return [
        item.item_type,
        _safe_str(item.subject),
        _safe_str(from_name),
        _safe_str(to),
        _safe_str(item.date or item.properties.get('start_date')),
        _safe_str(item.properties.get('location', '')),
        item.folder_path,
        att_names,
        _body_text_for(item)[:300],
    ]


def write_items_csv(items: list[MessageItem], out_path: str,
                    reader: PstReader | None = None) -> int:
    """Write items to a CSV; optionally extract attachments alongside."""
    count = 0
    with open(out_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(_CSV_HEADER)
        for item in items:
            writer.writerow(_csv_row(item))
            count += 1
    if reader is not None and any(i.has_attachments for i in items):
        att_root = os.path.splitext(out_path)[0] + '_attachments'
        for item in items:
            if item.has_attachments:
                _dump_attachments_folder(
                    reader, item, os.path.join(att_root, _item_basename(item)))
    return count


# ===================== high-level: export selected items =====================

def export_selected(reader: PstReader, items: list[MessageItem], fmt: str,
                    dest: str, progress: ProgressFn = None) -> int:
    """Export a list of items in ``fmt``.

    For ``csv`` ``dest`` is a file path; for txt/html/eml/pdf it is a directory.
    Attachments are embedded for eml/html/pdf and written to a per-item folder
    for txt/csv. Returns the number of items exported.
    """
    fmt = fmt.lower()
    if not items:
        return 0
    if fmt == 'csv':
        return write_items_csv(items, dest, reader=reader)

    os.makedirs(dest, exist_ok=True)
    total = len(items)
    count = 0
    for i, item in enumerate(items, 1):
        try:
            write_item(item, reader, fmt, dest)
            count += 1
        except Exception:
            pass
        if progress:
            progress(i, total, item.subject or item.item_type)
    return count


# ===================== high-level: export a folder subtree =====================

def export_folder_tree(reader: PstReader, folder_info, fmt: str, out_dir: str,
                       progress: ProgressFn = None) -> int:
    """Export a folder and all of its subfolders, recreating the tree on disk.

    All item types are exported in ``fmt``. For ``csv`` each folder gets one CSV;
    for txt/html/eml/pdf each item becomes a file in its mirrored folder.
    Returns the total number of items exported.
    """
    fmt = fmt.lower()
    root_name = _sanitize_filename(folder_info.name or 'export')
    base_out = os.path.join(out_dir, root_name)
    os.makedirs(base_out, exist_ok=True)

    # Map a source folder path -> on-disk directory, mirroring the hierarchy.
    root_path = folder_info.path
    csv_buckets = {}  # dir -> list[item] (csv mode)
    count = 0

    def dest_dir_for(folder_path: str) -> str:
        rel = folder_path[len(root_path):].lstrip('/') if \
            folder_path.startswith(root_path) else folder_path
        parts = [_sanitize_filename(p) for p in rel.split('/') if p]
        d = os.path.join(base_out, *parts) if parts else base_out
        os.makedirs(d, exist_ok=True)
        return d

    for folder_path, item in reader.iter_messages_recursive(folder_info):
        d = dest_dir_for(folder_path)
        try:
            if fmt == 'csv':
                csv_buckets.setdefault(d, []).append(item)
            else:
                write_item(item, reader, fmt, d)
            count += 1
        except Exception:
            pass
        if progress and count % 25 == 0:
            progress(count, 0, folder_path)

    if fmt == 'csv':
        for d, bucket in csv_buckets.items():
            name = os.path.basename(d.rstrip('/')) or root_name
            write_items_csv(bucket, os.path.join(d, f'{name}.csv'), reader=reader)

    return count


# ===================== Calendar Export =====================

# RFC 5545 content-line folding: lines are limited to 75 octets; continuation
# lines begin with one space, leaving 74 octets of payload.
ICS_LINE_OCTET_LIMIT = 75
ICS_CONTINUATION_OCTET_LIMIT = 74
# UTF-8 continuation bytes match 0b10xxxxxx, i.e. (byte & 0xC0) == 0x80.
_UTF8_CONT_MASK = 0xC0
_UTF8_CONT_VALUE = 0x80


def _escape_ics(text: str) -> str:
    text = _safe_str(text)
    text = text.replace('\\', '\\\\').replace(';', '\\;').replace(',', '\\,')
    return text.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\\n')


def _fold_ics(line: str) -> str:
    """Fold a content line to <=75 octets per RFC 5545."""
    raw = line.encode('utf-8')
    if len(raw) <= ICS_LINE_OCTET_LIMIT:
        return line
    chunks = []
    start = 0
    limit = ICS_LINE_OCTET_LIMIT
    while start < len(raw):
        end = min(start + limit, len(raw))
        # avoid splitting a multibyte char
        while end < len(raw) and (raw[end] & _UTF8_CONT_MASK) == _UTF8_CONT_VALUE:
            end -= 1
        chunks.append(raw[start:end])
        start = end
        limit = ICS_CONTINUATION_OCTET_LIMIT
    return '\r\n '.join(c.decode('utf-8') for c in chunks)


def _ics_dt(dt) -> str | None:
    if not isinstance(dt, datetime):
        return None
    if dt.tzinfo is not None:
        return dt.astimezone().strftime('%Y%m%dT%H%M%SZ')
    return dt.strftime('%Y%m%dT%H%M%S')


def export_calendar_to_ics(items: list[MessageItem], output_path: str) -> int:
    """Export calendar items to a RFC-5545 iCalendar file. Returns count."""
    lines = ['BEGIN:VCALENDAR', 'VERSION:2.0',
             'PRODID:-//OutlookTool//PST Export//EN', 'CALSCALE:GREGORIAN',
             'METHOD:PUBLISH']
    stamp = datetime.now(tz=timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    count = 0
    for idx, item in enumerate(items):
        if item.item_type != 'calendar':
            continue
        start = (item.properties.get('start_date') or item.date
                 or item.properties.get('creation_time')
                 or item.properties.get('last_modification_time'))
        ds = _ics_dt(start)
        if not ds:
            # An event with no resolvable start cannot be valid iCalendar; skip it.
            continue
        end = item.properties.get('end_date')
        de = _ics_dt(end)
        ev = ['BEGIN:VEVENT',
              f'UID:outlooktool-{idx}-{abs(hash((item.subject, str(start))))}@local',
              f'DTSTAMP:{stamp}', f'DTSTART:{ds}']
        if de:
            ev.append(f'DTEND:{de}')
        if item.subject:
            ev.append(f'SUMMARY:{_escape_ics(item.subject)}')
        location = item.properties.get('location', '')
        if location:
            ev.append(f'LOCATION:{_escape_ics(location)}')
        if item.body_text:
            ev.append(f'DESCRIPTION:{_escape_ics(item.body_text)}')
        if item.sender_name:
            email = item.sender_email or 'unknown@local'
            ev.append(f'ORGANIZER;CN={_escape_ics(item.sender_name)}:mailto:{email}')
        ev.append('END:VEVENT')
        lines.extend(_fold_ics(line) for line in ev)
        count += 1
    lines.append('END:VCALENDAR')
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        f.write('\r\n'.join(lines) + '\r\n')
    return count


def export_calendar_to_csv(items: list[MessageItem], output_path: str) -> int:
    count = 0
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Subject', 'Start', 'End', 'Location', 'Organizer',
                         'Description', 'Folder'])
        for item in items:
            if item.item_type != 'calendar':
                continue
            writer.writerow([
                _safe_str(item.subject),
                _safe_str(item.properties.get('start_date') or item.date),
                _safe_str(item.properties.get('end_date')),
                _safe_str(item.properties.get('location', '')),
                _safe_str(item.sender_name),
                _body_text_for(item)[:500],
                item.folder_path,
            ])
            count += 1
    return count


# ===================== Contacts Export =====================

def _vcard_escape(text: str) -> str:
    text = _safe_str(text)
    return text.replace('\\', '\\\\').replace(';', '\\;').replace(',', '\\,') \
               .replace('\r\n', '\\n').replace('\n', '\\n')


def export_contacts_to_vcf(items: list[MessageItem], output_path: str) -> int:
    count = 0
    vcards = []
    for item in items:
        if item.item_type != 'contact':
            continue
        p = item.properties
        lines = ['BEGIN:VCARD', 'VERSION:3.0']
        surname = _vcard_escape(p.get('surname', ''))
        given = _vcard_escape(p.get('given_name', ''))
        display = _vcard_escape(p.get('display_name', '') or item.subject)
        if surname or given:
            lines.append(f'N:{surname};{given};;;')
        if display:
            lines.append(f'FN:{display}')
        email = resolve_contact_email(item)
        if email:
            lines.append(f'EMAIL;TYPE=INTERNET:{email}')
        for key, label in (('business_phone', 'WORK'), ('home_phone', 'HOME'),
                           ('mobile_phone', 'CELL')):
            if p.get(key):
                lines.append(f'TEL;TYPE={label}:{_vcard_escape(p[key])}')
        if p.get('company_name'):
            lines.append(f'ORG:{_vcard_escape(p["company_name"])}')
        if p.get('title'):
            lines.append(f'TITLE:{_vcard_escape(p["title"])}')
        if p.get('postal_address'):
            lines.append(f'ADR;TYPE=WORK:;;{_vcard_escape(p["postal_address"])};;;;')
        if p.get('nickname'):
            lines.append(f'NICKNAME:{_vcard_escape(p["nickname"])}')
        lines.append('END:VCARD')
        vcards.append('\r\n'.join(lines))
        count += 1
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\r\n'.join(vcards) + ('\r\n' if vcards else ''))
    return count


def export_contacts_to_csv(items: list[MessageItem], output_path: str) -> int:
    count = 0
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Display Name', 'First Name', 'Last Name', 'Email',
                         'Business Phone', 'Home Phone', 'Mobile Phone',
                         'Company', 'Job Title', 'Department', 'Office',
                         'Address'])
        for item in items:
            if item.item_type != 'contact':
                continue
            p = item.properties
            writer.writerow([
                _safe_str(p.get('display_name', '') or item.subject),
                _safe_str(p.get('given_name', '')),
                _safe_str(p.get('surname', '')),
                resolve_contact_email(item),
                _safe_str(p.get('business_phone', '')),
                _safe_str(p.get('home_phone', '')),
                _safe_str(p.get('mobile_phone', '')),
                _safe_str(p.get('company_name', '')),
                _safe_str(p.get('title', '')),
                _safe_str(p.get('department_name', '')),
                _safe_str(p.get('office_location', '')),
                _safe_str(p.get('postal_address', '')),
            ])
            count += 1
    return count


# ===================== Email Export (whole-type, back-compat) =====================

def export_emails_to_eml(items: list[MessageItem], output_dir: str,
                         reader: PstReader | None = None) -> int:
    os.makedirs(output_dir, exist_ok=True)
    count = 0
    for item in items:
        if item.item_type != 'email':
            continue
        try:
            write_item(item, reader, 'eml', output_dir)
            count += 1
        except Exception:
            pass
    return count


def export_emails_to_csv(items: list[MessageItem], output_path: str) -> int:
    emails = [i for i in items if i.item_type == 'email']
    return write_items_csv(emails, output_path)


# ===================== Attachment Export =====================

def export_attachments(reader: PstReader, items: list[MessageItem],
                       output_dir: str) -> int:
    os.makedirs(output_dir, exist_ok=True)
    count = 0
    for item in items:
        if not item.has_attachments:
            continue
        try:
            attachments = reader.get_attachments(item)
        except Exception:
            continue
        for att in attachments:
            filename = _sanitize_filename(att.name or f'attachment_{count}')
            filepath = _unique_filepath(os.path.join(output_dir, filename))
            if reader.save_attachment(att, filepath):
                count += 1
    return count


# ===================== General Export =====================

def export_all_items_to_csv(items: list[MessageItem], output_path: str) -> int:
    return write_items_csv(items, output_path)
