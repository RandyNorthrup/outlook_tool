"""
Export functionality for PST/OST data.
Supports exporting calendar (ICS/CSV), contacts (VCF/CSV),
emails (EML/CSV), and attachments.
"""

import csv
import os
from datetime import datetime
from typing import List

from pst_handler import MessageItem, PstReader


def _safe_str(value, default=''):
    """Convert value to string safely."""
    if value is None:
        return default
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d %H:%M:%S')
    return str(value)


def _sanitize_filename(name: str) -> str:
    """Sanitize a string for use as a filename."""
    invalid_chars = '<>:"/\\|?*\x00'
    for c in invalid_chars:
        name = name.replace(c, '_')
    # Remove leading/trailing whitespace and dots
    name = name.strip(' .')
    return name[:200] if name else 'unnamed'


def _unique_filepath(filepath: str) -> str:
    """Ensure filepath is unique by appending counter if needed."""
    if not os.path.exists(filepath):
        return filepath
    base, ext = os.path.splitext(filepath)
    counter = 1
    while os.path.exists(filepath):
        filepath = f"{base}_{counter}{ext}"
        counter += 1
    return filepath


def _escape_ics(text: str) -> str:
    """Escape text for iCalendar format."""
    text = text.replace('\\', '\\\\')
    text = text.replace(';', '\\;')
    text = text.replace(',', '\\,')
    text = text.replace('\n', '\\n')
    text = text.replace('\r', '')
    return text


def _format_ics_datetime(dt: datetime) -> str:
    """Format datetime for iCalendar."""
    if dt.tzinfo:
        return dt.strftime('%Y%m%dT%H%M%SZ')
    return dt.strftime('%Y%m%dT%H%M%S')


# ===================== Calendar Export =====================

def export_calendar_to_ics(items: List[MessageItem], output_path: str) -> int:
    """Export calendar items to ICS (iCalendar) format. Returns count exported."""
    lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//OutlookTool//EN',
    ]

    count = 0
    for idx, item in enumerate(items):
        if item.item_type != 'calendar':
            continue

        lines.append('BEGIN:VEVENT')
        lines.append(f'UID:outlooktool-{idx}-{hash(item.subject or "")}@local')

        if item.subject:
            lines.append(f'SUMMARY:{_escape_ics(item.subject)}')

        start = item.properties.get('start_date') or item.date
        end = item.properties.get('end_date')

        if start:
            if isinstance(start, datetime):
                lines.append(f'DTSTART:{_format_ics_datetime(start)}')
            else:
                lines.append(f'DTSTART:{start}')

        if end:
            if isinstance(end, datetime):
                lines.append(f'DTEND:{_format_ics_datetime(end)}')
            else:
                lines.append(f'DTEND:{end}')

        location = item.properties.get('location', '')
        if location:
            lines.append(f'LOCATION:{_escape_ics(str(location))}')

        if item.body_text:
            lines.append(f'DESCRIPTION:{_escape_ics(item.body_text)}')

        if item.sender_name:
            email = item.sender_email or 'unknown'
            lines.append(f'ORGANIZER;CN={_escape_ics(item.sender_name)}:mailto:{email}')

        lines.append('END:VEVENT')
        count += 1

    lines.append('END:VCALENDAR')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\r\n'.join(lines))

    return count


def export_calendar_to_csv(items: List[MessageItem], output_path: str) -> int:
    """Export calendar items to CSV. Returns count exported."""
    count = 0
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Subject', 'Start Date', 'End Date', 'Location',
                         'Organizer', 'Description', 'Folder'])

        for item in items:
            if item.item_type != 'calendar':
                continue

            start = item.properties.get('start_date') or item.date
            end = item.properties.get('end_date', '')
            location = item.properties.get('location', '')

            writer.writerow([
                _safe_str(item.subject),
                _safe_str(start),
                _safe_str(end),
                _safe_str(location),
                _safe_str(item.sender_name),
                (_safe_str(item.body_text))[:500],
                item.folder_path,
            ])
            count += 1

    return count


# ===================== Contacts Export =====================

def export_contacts_to_vcf(items: List[MessageItem], output_path: str) -> int:
    """Export contacts to vCard (VCF) format. Returns count exported."""
    count = 0
    vcards = []

    for item in items:
        if item.item_type != 'contact':
            continue

        props = item.properties
        lines = ['BEGIN:VCARD', 'VERSION:3.0']

        surname = _safe_str(props.get('surname', ''))
        given = _safe_str(props.get('given_name', ''))
        display = _safe_str(props.get('display_name', '')) or _safe_str(item.subject)

        if surname or given:
            lines.append(f'N:{surname};{given};;;')

        if display:
            lines.append(f'FN:{display}')

        email = _safe_str(props.get('smtp_address', '')) or _safe_str(item.sender_email)
        if email:
            lines.append(f'EMAIL;TYPE=INTERNET:{email}')

        biz_phone = _safe_str(props.get('business_phone', ''))
        home_phone = _safe_str(props.get('home_phone', ''))
        mobile = _safe_str(props.get('mobile_phone', ''))

        if biz_phone:
            lines.append(f'TEL;TYPE=WORK:{biz_phone}')
        if home_phone:
            lines.append(f'TEL;TYPE=HOME:{home_phone}')
        if mobile:
            lines.append(f'TEL;TYPE=CELL:{mobile}')

        company = _safe_str(props.get('company_name', ''))
        title = _safe_str(props.get('title', ''))
        if company:
            lines.append(f'ORG:{company}')
        if title:
            lines.append(f'TITLE:{title}')

        # Business address
        biz_street = _safe_str(props.get('business_address_street', ''))
        biz_city = _safe_str(props.get('business_address_city', ''))
        biz_state = _safe_str(props.get('business_address_state', ''))
        biz_zip = _safe_str(props.get('business_address_postal_code', ''))
        biz_country = _safe_str(props.get('business_address_country', ''))

        if any([biz_street, biz_city, biz_state, biz_zip, biz_country]):
            lines.append(f'ADR;TYPE=WORK:;;{biz_street};{biz_city};{biz_state};{biz_zip};{biz_country}')

        # Home address
        home_street = _safe_str(props.get('home_address_street', ''))
        home_city = _safe_str(props.get('home_address_city', ''))
        home_state = _safe_str(props.get('home_address_state', ''))
        home_zip = _safe_str(props.get('home_address_postal_code', ''))
        home_country = _safe_str(props.get('home_address_country', ''))

        if any([home_street, home_city, home_state, home_zip, home_country]):
            lines.append(f'ADR;TYPE=HOME:;;{home_street};{home_city};{home_state};{home_zip};{home_country}')

        nickname = _safe_str(props.get('nickname', ''))
        if nickname:
            lines.append(f'NICKNAME:{nickname}')

        dept = _safe_str(props.get('department_name', ''))
        if dept:
            lines.append(f'X-DEPARTMENT:{dept}')

        lines.append('END:VCARD')
        vcards.append('\r\n'.join(lines))
        count += 1

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\r\n'.join(vcards))

    return count


def export_contacts_to_csv(items: List[MessageItem], output_path: str) -> int:
    """Export contacts to CSV. Returns count exported."""
    count = 0
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Display Name', 'First Name', 'Last Name', 'Email',
            'Business Phone', 'Home Phone', 'Mobile Phone',
            'Company', 'Job Title', 'Department', 'Office Location',
            'Address',
        ])

        for item in items:
            if item.item_type != 'contact':
                continue

            props = item.properties
            writer.writerow([
                _safe_str(props.get('display_name', '')) or _safe_str(item.subject),
                _safe_str(props.get('given_name', '')),
                _safe_str(props.get('surname', '')),
                _safe_str(props.get('smtp_address', '')) or _safe_str(item.sender_email),
                _safe_str(props.get('business_phone', '')),
                _safe_str(props.get('home_phone', '')),
                _safe_str(props.get('mobile_phone', '')),
                _safe_str(props.get('company_name', '')),
                _safe_str(props.get('title', '')),
                _safe_str(props.get('department_name', '')),
                _safe_str(props.get('office_location', '')),
                _safe_str(props.get('postal_address', '')),
            ])
            count += 1

    return count


# ===================== Email Export =====================

def export_emails_to_eml(items: List[MessageItem], output_dir: str) -> int:
    """Export emails as individual EML files. Returns count exported."""
    os.makedirs(output_dir, exist_ok=True)
    count = 0

    for item in items:
        if item.item_type != 'email':
            continue

        subject = _sanitize_filename(item.subject or 'no_subject')
        date_str = item.date.strftime('%Y%m%d_%H%M%S') if item.date else 'nodate'
        filename = f"{date_str}_{subject}.eml"
        filepath = _unique_filepath(os.path.join(output_dir, filename))

        lines = []
        if item.sender_name or item.sender_email:
            if item.sender_name and item.sender_email:
                lines.append(f'From: {item.sender_name} <{item.sender_email}>')
            else:
                lines.append(f'From: {item.sender_email or item.sender_name}')

        if item.recipients:
            lines.append(f'To: {item.recipients}')
        if item.cc:
            lines.append(f'Cc: {item.cc}')
        if item.subject:
            lines.append(f'Subject: {item.subject}')
        if item.date:
            lines.append(f'Date: {item.date.strftime("%a, %d %b %Y %H:%M:%S %z")}')

        lines.append('MIME-Version: 1.0')

        if item.body_html:
            lines.append('Content-Type: text/html; charset=utf-8')
            lines.append('')
            lines.append(item.body_html)
        else:
            lines.append('Content-Type: text/plain; charset=utf-8')
            lines.append('')
            lines.append(item.body_text or '')

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\r\n'.join(lines))

        count += 1

    return count


def export_emails_to_csv(items: List[MessageItem], output_path: str) -> int:
    """Export emails to CSV. Returns count exported."""
    count = 0
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Subject', 'From', 'To', 'CC', 'Date',
                         'Has Attachments', 'Body Preview', 'Folder'])

        for item in items:
            if item.item_type != 'email':
                continue

            body_preview = (_safe_str(item.body_text))[:200]
            writer.writerow([
                _safe_str(item.subject),
                _safe_str(item.sender_name),
                _safe_str(item.recipients),
                _safe_str(item.cc),
                _safe_str(item.date),
                'Yes' if item.has_attachments else 'No',
                body_preview,
                item.folder_path,
            ])
            count += 1

    return count


# ===================== Attachment Export =====================

def export_attachments(reader: PstReader, items: List[MessageItem], output_dir: str) -> int:
    """Export all attachments from given messages. Returns count exported."""
    os.makedirs(output_dir, exist_ok=True)
    count = 0

    for item in items:
        if not item.has_attachments:
            continue

        attachments = reader.get_attachments(item)
        for att in attachments:
            filename = _sanitize_filename(att.name or f'attachment_{count}')
            filepath = _unique_filepath(os.path.join(output_dir, filename))

            if reader.save_attachment(att, filepath):
                count += 1

    return count


# ===================== General Export =====================

def export_all_items_to_csv(items: List[MessageItem], output_path: str) -> int:
    """Export all items to a single CSV regardless of type. Returns count exported."""
    count = 0
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Type', 'Subject', 'From/Name', 'Date', 'Folder', 'Body Preview'])

        for item in items:
            body_preview = (_safe_str(item.body_text))[:200]
            display_name = _safe_str(item.sender_name) or _safe_str(item.properties.get('display_name', ''))
            writer.writerow([
                item.item_type,
                _safe_str(item.subject),
                display_name,
                _safe_str(item.date),
                item.folder_path,
                body_preview,
            ])
            count += 1

    return count
