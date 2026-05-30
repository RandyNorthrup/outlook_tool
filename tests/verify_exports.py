"""
End-to-end verification harness for the Outlook PST/OST Tool.

Runs against the real files in ../temp and exercises:
  1. MAPI property fix (emails To/sender, contacts, calendar start/end/location)
  2. export_selected -> txt/csv/html/eml/pdf with attachments
  3. export_folder_tree -> recreated tree, all item types, counts
  4. calendar ICS validity
  5. full recursive no-crash pass over both files

Exit code is non-zero if any check fails.
"""
import email
import os
import sys
import tempfile
from email import policy

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import exporters  # noqa: E402
from pst_handler import PstReader  # noqa: E402

PST = os.path.join(ROOT, 'temp', 'my emails.pst')
OST = os.path.join(ROOT, 'temp', 'randy.northrup@outlook.com.ost')

_failures = []
_passes = []


def check(name, cond, detail=''):
    if cond:
        _passes.append(name)
        print(f"  PASS  {name}" + (f"  ({detail})" if detail else ''))
    else:
        _failures.append(name)
        print(f"  FAIL  {name}" + (f"  ({detail})" if detail else ''))


def folders_by_name(reader):
    return {f.name: f for f in reader.get_folder_tree()}


def test_property_fix(reader):
    print("\n[1] MAPI property extraction")
    fb = folders_by_name(reader)

    # Emails: recipients + sender_email populated
    emails = []
    for fname in ('misha', 'Sent Items', 'Inbox'):
        if fname in fb:
            emails += [m for m in reader.get_messages(fb[fname])
                       if m.item_type == 'email']
    with_to = sum(1 for m in emails if m.recipients)
    with_sender = sum(1 for m in emails if m.sender_email)
    check("emails have recipients", with_to > 0, f"{with_to}/{len(emails)}")
    check("emails have sender_email", with_sender > 0, f"{with_sender}/{len(emails)}")

    # Contacts: phone or email present on many; owner address NOT used as contact email
    if 'Contacts' in fb:
        contacts = [m for m in reader.get_messages(fb['Contacts'])
                    if m.item_type == 'contact']
        with_phone = sum(1 for c in contacts
                         if c.properties.get('business_phone')
                         or c.properties.get('home_phone'))
        with_email = sum(1 for c in contacts if exporters.resolve_contact_email(c))
        check("contacts have phone numbers", with_phone > 0,
              f"{with_phone}/{len(contacts)}")
        check("contacts have resolved email", with_email > 0,
              f"{with_email}/{len(contacts)}")
        # owner contamination guard
        owner_as_email = sum(
            1 for c in contacts
            if exporters.resolve_contact_email(c) == c.properties.get('sender_email')
            and c.properties.get('sender_email'))
        check("contact email != owner address", owner_as_email == 0,
              f"{owner_as_email} contaminated")

    # Calendar start/end/location
    if 'Calendar' in fb:
        cals = [m for m in reader.get_messages(fb['Calendar'])
                if m.item_type == 'calendar']
        with_start = sum(1 for c in cals if c.properties.get('start_date'))
        with_loc = sum(1 for c in cals if c.properties.get('location'))
        check("calendar items have start_date", with_start > len(cals) * 0.8,
              f"{with_start}/{len(cals)}")
        check("calendar items have location (some)", with_loc > 0,
              f"{with_loc}/{len(cals)}")
        # spot-check
        spot = next((c for c in cals if c.subject.strip() == 'Saturday Sitting'),
                    None)
        if spot:
            s = str(spot.properties.get('start_date'))
            e = str(spot.properties.get('end_date'))
            loc = spot.properties.get('location')
            check("Saturday Sitting start", s.startswith('2009-10-03 11:20'), s)
            check("Saturday Sitting end", e.startswith('2009-10-03 18:20'), e)
            check("Saturday Sitting location", loc == 'Home Zendo', str(loc))


def test_export_selected(reader, workdir):
    print("\n[2] export_selected -> all 5 formats (with attachments)")
    fb = folders_by_name(reader)
    # gather some emails, prioritising ones with attachments
    pool = []
    for fname in ('misha', 'Inbox', 'Sent Items', '2010'):
        if fname in fb:
            pool += [m for m in reader.get_messages(fb[fname])
                     if m.item_type == 'email']
    with_att = [m for m in pool if m.has_attachments][:5]
    sample = (with_att + pool[:5])[:8]
    check("have email sample to export", len(sample) > 0, f"{len(sample)} items")
    check("sample includes attachments", len(with_att) > 0,
          f"{len(with_att)} with attachments")

    # TXT (per-email attachment folder)
    d = os.path.join(workdir, 'txt')
    n = exporters.export_selected(reader, sample, 'txt', d)
    txts = [f for f in os.listdir(d) if f.endswith('.txt')]
    att_dirs = [f for f in os.listdir(d) if f.endswith('_attachments')]
    check("txt: files written", n == len(sample) and len(txts) == n,
          f"{len(txts)} files")
    check("txt: per-email attachment folder", len(att_dirs) > 0 if with_att else True,
          f"{len(att_dirs)} att folders")

    # HTML (data-URI embed)
    d = os.path.join(workdir, 'html')
    exporters.export_selected(reader, sample, 'html', d)
    htmls = [f for f in os.listdir(d) if f.endswith('.html')]
    check("html: files written", len(htmls) == len(sample), f"{len(htmls)} files")
    if with_att:
        any_data_uri = False
        for f in htmls:
            with open(os.path.join(d, f), encoding='utf-8') as fh:
                txt = fh.read()
            if 'data:' in txt and 'base64,' in txt:
                any_data_uri = True
                break
        check("html: attachment embedded as data-URI", any_data_uri)

    # EML (MIME embed + parseable)
    d = os.path.join(workdir, 'eml')
    exporters.export_selected(reader, sample, 'eml', d)
    emls = [f for f in os.listdir(d) if f.endswith('.eml')]
    check("eml: files written", len(emls) == len(sample), f"{len(emls)} files")
    parsed_ok = 0
    att_parts = 0
    for f in emls:
        with open(os.path.join(d, f), 'rb') as fh:
            msg = email.message_from_binary_file(fh, policy=policy.default)
        parsed_ok += 1
        for part in msg.walk():
            if part.get_content_disposition() == 'attachment':
                att_parts += 1
    check("eml: all parse via email module", parsed_ok == len(emls))
    check("eml: attachments embedded as MIME parts", att_parts > 0 if with_att else True,
          f"{att_parts} parts")

    # CSV (tabular + attachments folder)
    p = os.path.join(workdir, 'selected.csv')
    n = exporters.export_selected(reader, sample, 'csv', p)
    check("csv: file written with rows", os.path.exists(p) and n == len(sample),
          f"{n} rows")

    # PDF (embed via pikepdf)
    d = os.path.join(workdir, 'pdf')
    exporters.export_selected(reader, sample, 'pdf', d)
    pdfs = [f for f in os.listdir(d) if f.endswith('.pdf')]
    check("pdf: files written", len(pdfs) == len(sample), f"{len(pdfs)} files")
    import pikepdf
    valid = 0
    embedded_any = False
    for f in pdfs:
        try:
            with pikepdf.open(os.path.join(d, f)) as pdf:
                valid += 1
                if len(pdf.attachments):
                    embedded_any = True
        except Exception:
            pass
    check("pdf: all valid PDFs", valid == len(pdfs), f"{valid}/{len(pdfs)}")
    check("pdf: attachments embedded", embedded_any if with_att else True)


def test_folder_tree(reader, workdir):
    print("\n[3] export_folder_tree -> mirrored tree, all item types")
    fb = folders_by_name(reader)
    target = fb.get('Top of Personal Folders') or fb.get('IPM_SUBTREE')
    if target is None:
        # pick the folder with the most descendant messages
        target = max(reader.get_folder_tree(), key=lambda f: f.message_count)
    print(f"  target folder: {target.path}")

    # count items via recursion, by type
    by_type = {}
    total = 0
    for _path, item in reader.iter_messages_recursive(target):
        by_type[item.item_type] = by_type.get(item.item_type, 0) + 1
        total += 1
    print(f"  recursive items: {total}  types={by_type}")

    d = os.path.join(workdir, 'tree_eml')
    n = exporters.export_folder_tree(reader, target, 'eml', d)
    # count files on disk
    files = sum(len([x for x in fs if x.endswith('.eml')])
                for _r, _ds, fs in os.walk(d))
    check("tree(eml): item count matches", n == total, f"{n} exported / {total}")
    check("tree(eml): files on disk match", files == total, f"{files} files")
    # verify subfolder dirs were created (mirrored tree)
    subdirs = [r for r, _ds, _fs in os.walk(d)]
    check("tree(eml): recreates subfolders", len(subdirs) > 1, f"{len(subdirs)} dirs")
    check("tree(eml): includes non-email items",
          (by_type.get('calendar', 0) + by_type.get('contact', 0)) > 0,
          f"cal={by_type.get('calendar',0)} con={by_type.get('contact',0)}")

    # pdf sample on a smaller subfolder to keep it fast
    small = None
    for f in reader.get_folder_tree():
        if 0 < f.message_count <= 60 and f.folder_type == 'email':
            small = f
            break
    if small:
        d2 = os.path.join(workdir, 'tree_pdf')
        n2 = exporters.export_folder_tree(reader, small, 'pdf', d2)
        pdfs = sum(len([x for x in fs if x.endswith('.pdf')])
                   for _r, _ds, fs in os.walk(d2))
        check("tree(pdf): produces pdfs", pdfs == n2 and pdfs > 0,
              f"{pdfs} pdfs from {small.name}")


def test_calendar_ics(reader, workdir):
    print("\n[4] calendar ICS validity")
    cals = reader.get_all_items_by_type('calendar')
    if not cals:
        print("  (no calendar items)")
        return
    p = os.path.join(workdir, 'calendar.ics')
    n = exporters.export_calendar_to_ics(cals, p)
    with open(p, encoding='utf-8') as fh:
        text = fh.read()
    nvev = text.count('BEGIN:VEVENT')
    check("ics: event count matches", nvev == n, f"{nvev} VEVENTs")
    check("ics: has VCALENDAR wrapper",
          text.startswith('BEGIN:VCALENDAR') and text.rstrip().endswith('END:VCALENDAR'))
    # every event has DTSTART + DTSTAMP + SUMMARY
    blocks = text.split('BEGIN:VEVENT')[1:]
    good = sum(1 for b in blocks if 'DTSTART:' in b and 'DTSTAMP:' in b)
    with_end = sum(1 for b in blocks if 'DTEND:' in b)
    check("ics: every event has DTSTART+DTSTAMP", good == len(blocks),
          f"{good}/{len(blocks)}")
    check("ics: events have DTEND", with_end > len(blocks) * 0.8,
          f"{with_end}/{len(blocks)}")
    # try icalendar parse if available
    try:
        from icalendar import Calendar
        Calendar.from_ical(text)
        check("ics: parses with icalendar lib", True)
    except ImportError:
        print("  (icalendar lib not installed; skipped strict parse)")
    except Exception as e:
        check("ics: parses with icalendar lib", False, str(e)[:60])

    p2 = os.path.join(workdir, 'calendar.csv')
    n2 = exporters.export_calendar_to_csv(cals, p2)
    check("calendar csv written", os.path.exists(p2) and n2 == n, f"{n2} rows")


def test_no_crash(path, workdir, label):
    print(f"\n[5] full recursive no-crash pass: {label}")
    reader = PstReader()
    reader.open(path)
    folders = reader.get_folder_tree()
    total_msgs = 0
    att_read = 0
    att_fail = 0
    errors = 0
    for f in folders:
        try:
            for m in reader.get_messages(f):
                total_msgs += 1
                if m.has_attachments:
                    try:
                        for a in reader.get_attachments(m):
                            data = reader.read_attachment_bytes(a)
                            if data is None:
                                att_fail += 1
                            else:
                                att_read += 1
                    except Exception:
                        errors += 1
        except Exception:
            errors += 1
    print(f"  messages={total_msgs} att_read={att_read} att_unreadable={att_fail} "
          f"folder_errors={errors}")
    check(f"{label}: walked all folders without exception", errors == 0,
          f"{errors} errors")
    check(f"{label}: parsed messages", total_msgs > 0, f"{total_msgs} msgs")
    reader.close()


def main():
    print("=" * 72)
    print("Outlook PST/OST Tool — export verification")
    print("=" * 72)

    workdir = tempfile.mkdtemp(prefix='outlooktool_verify_')
    print("workdir:", workdir)

    # PST: the rich file
    print("\n##### FILE: my emails.pst #####")
    reader = PstReader()
    reader.open(PST)
    test_property_fix(reader)
    test_export_selected(reader, os.path.join(workdir, 'pst'))
    test_folder_tree(reader, os.path.join(workdir, 'pst'))
    test_calendar_ics(reader, os.path.join(workdir, 'pst'))
    reader.close()

    # No-crash pass over BOTH files
    test_no_crash(PST, workdir, "PST")
    test_no_crash(OST, workdir, "OST")

    print("\n" + "=" * 72)
    print(f"RESULT: {len(_passes)} passed, {len(_failures)} failed")
    if _failures:
        print("FAILURES:")
        for f in _failures:
            print("  -", f)
    print("=" * 72)
    return 1 if _failures else 0


if __name__ == '__main__':
    sys.exit(main())
