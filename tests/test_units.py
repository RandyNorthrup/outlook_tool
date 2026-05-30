"""Fast unit tests for the pure (non-pypff) logic.

These run without a PST/OST file or libpff, so they are safe for CI.
Run with:  pytest tests/test_units.py
"""
import email
import os
import sys
from datetime import datetime
from email import policy

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import exporters  # noqa: E402
from pst_handler import MessageItem, html_to_text  # noqa: E402


def _email(**kw) -> MessageItem:
    base = {"item_type": "email", "subject": "Hi", "sender_name": "A",
            "sender_email": "a@x.com", "recipients": "b@y.com"}
    base.update(kw)
    return MessageItem(**base)


# ---- html_to_text (the OST "shows raw HTML" fix) ----

def test_html_to_text_strips_markup_and_decodes_entities():
    html = ('<!DOCTYPE html><html><head><style>p{color:red}</style></head>'
            '<body><p>Hello<br>World &amp; friends</p>'
            '<script>alert(1)</script></body></html>')
    text = html_to_text(html)
    assert "Hello" in text
    assert "World & friends" in text
    assert "<" not in text and ">" not in text
    assert "color:red" not in text  # <style> content removed
    assert "alert" not in text       # <script> content removed


def test_html_to_text_empty():
    assert html_to_text("") == ""
    assert html_to_text(None) == ""


def test_body_text_for_uses_html_when_no_plaintext():
    item = _email(body_text="", body_html="<p>Only&nbsp;HTML</p>")
    out = exporters._body_text_for(item)
    assert "Only" in out and "HTML" in out
    assert "<p>" not in out


# ---- format_size ----

def test_format_size_boundaries():
    assert exporters.format_size(0) == "0 B"
    assert exporters.format_size(1023) == "1023 B"
    assert exporters.format_size(1024) == "1.0 KB"
    assert exporters.format_size(1024 * 1024) == "1.0 MB"
    assert exporters.format_size(1024 ** 3) == "1.00 GB"


# ---- contact email owner-exclusion ----

def test_resolve_contact_email_excludes_owner():
    item = MessageItem(
        item_type="contact", subject="Jane",
        sender_email="owner@store.com",
        properties={
            "_email_candidates": ["owner@store.com", "jane@real.com"],
            "sender_email": "owner@store.com",
        })
    assert exporters.resolve_contact_email(item) == "jane@real.com"


def test_resolve_contact_email_none_when_only_owner():
    item = MessageItem(
        item_type="contact", subject="X", sender_email="owner@store.com",
        properties={"_email_candidates": ["owner@store.com"],
                    "sender_email": "owner@store.com"})
    assert exporters.resolve_contact_email(item) == ""


# ---- EML building + header sanitization ----

def test_build_eml_sanitizes_crlf_headers_and_parses():
    item = _email(subject="Line1\r\nLine2", recipients="b@y.com\r\nc@z.com",
                  body_text="hello body")
    raw = exporters.build_eml_bytes(item, reader=None)
    msg = email.message_from_bytes(raw, policy=policy.default)
    subject = str(msg["Subject"])
    assert "\n" not in subject and "\r" not in subject
    assert subject == "Line1 Line2"
    assert "hello body" in msg.get_content()


def test_build_eml_message_attachment_does_not_crash():
    # message/rfc822 attachments must be coerced, not crash EmailMessage.
    class FakeReader:
        def get_attachments(self, _item):
            from pst_handler import AttachmentInfo
            return [AttachmentInfo(name="orig.eml", size=3,
                                   mime_type="message/rfc822")]

        def read_attachment_bytes(self, _att):
            return b"abc"

    item = _email(has_attachments=True, attachment_count=1)
    raw = exporters.build_eml_bytes(item, reader=FakeReader())
    msg = email.message_from_bytes(raw)
    parts = [p.get_content_type() for p in msg.walk()]
    assert "application/octet-stream" in parts


# ---- ICS export ----

def _cal(subject, start, end=None, location=""):
    props = {}
    if start is not None:
        props["start_date"] = start
    if end is not None:
        props["end_date"] = end
    if location:
        props["location"] = location
    return MessageItem(item_type="calendar", subject=subject, properties=props,
                       date=start)


def test_ics_export_valid_and_skips_dateless(tmp_path):
    items = [
        _cal("Meeting", datetime(2009, 10, 3, 11, 20),
             datetime(2009, 10, 3, 12, 20), "Room 1"),
        _cal("No date", None),  # must be skipped (invalid without DTSTART)
    ]
    out = tmp_path / "cal.ics"
    n = exporters.export_calendar_to_ics(items, str(out))
    text = out.read_text(encoding="utf-8")
    assert n == 1
    assert text.count("BEGIN:VEVENT") == 1
    assert text.startswith("BEGIN:VCALENDAR")
    assert text.rstrip().endswith("END:VCALENDAR")
    block = text.split("BEGIN:VEVENT")[1]
    assert "DTSTART:" in block and "DTSTAMP:" in block and "SUMMARY:" in block
    assert "LOCATION:Room 1" in block


def test_ics_folding_under_75_octets():
    long_summary = "X" * 200
    item = _cal(long_summary, datetime(2020, 1, 1, 9, 0))
    import tempfile
    with tempfile.NamedTemporaryFile("w+", suffix=".ics", delete=False) as fh:
        path = fh.name
    exporters.export_calendar_to_ics([item], path)
    # newline="" keeps the real CRLF terminators the writer emitted.
    with open(path, encoding="utf-8", newline="") as fh:
        for line in fh.read().split("\r\n"):
            assert len(line.encode("utf-8")) <= 75, repr(line)
    os.unlink(path)


def test_ics_escaping():
    item = _cal("a;b,c\\d", datetime(2020, 1, 1, 9, 0))
    import tempfile
    with tempfile.NamedTemporaryFile("w+", suffix=".ics", delete=False) as fh:
        path = fh.name
    exporters.export_calendar_to_ics([item], path)
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    assert "a\\;b\\,c\\\\d" in text
    os.unlink(path)


# ---- filename sanitization ----

def test_item_basename_sanitizes():
    item = _email(subject='bad/name:with*chars?', date=datetime(2020, 5, 1, 8, 30))
    base = exporters._item_basename(item)
    for ch in '<>:"/\\|?*':
        assert ch not in base
    assert base.startswith("20200501_083000_")
