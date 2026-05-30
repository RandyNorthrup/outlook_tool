"""Headless GUI smoke test (run under a virtual X display, e.g. xvfb-run).

Drives the real OutlookToolApp code paths to verify:
  * the single window builds and renders without error,
  * checkbox action-selection works,
  * the preview shows readable text for an OST HTML-only email (no raw markup),
  * the preview attachments panel populates,
and saves a screenshot if a grabber is available.
"""
import os
import sys
import tkinter as tk

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import app as appmod  # noqa: E402
from pst_handler import PstReader  # noqa: E402

PST = os.path.join(ROOT, "temp", "my emails.pst")
OST = os.path.join(ROOT, "temp", "randy.northrup@outlook.com.ost")
SHOT = os.path.join(HERE, "smoke_screenshot.png")

failures = []


def check(name, cond, detail=""):
    tag = "PASS" if cond else "FAIL"
    if not cond:
        failures.append(name)
    print(f"  {tag}  {name}" + (f"  ({detail})" if detail else ""))


def folder_by_name(reader, name):
    for f in reader.get_folder_tree():
        if f.name == name:
            return f
    return None


def main():
    root = tk.Tk()
    root.geometry("1280x820")
    application = appmod.OutlookToolApp(root)

    # ---- OST: HTML-only email must preview as readable text, not markup ----
    ost = PstReader()
    ost.open(OST)
    application.reader = ost
    inbox = folder_by_name(ost, "Inbox")
    html_email = None
    if inbox:
        for m in ost.get_messages(inbox):
            if m.item_type == "email" and m.body_html and not m.body_text:
                html_email = m
                break
    if html_email is not None:
        application._show_preview(html_email)
        root.update()
        shown = application.preview_text.get("1.0", tk.END)
        check("OST html-only email previews without markup",
              "<html" not in shown.lower() and "<style" not in shown.lower()
              and "<div" not in shown.lower(),
              f"{len(shown)} chars shown")
        check("OST preview has readable content", len(shown.strip()) > 20)
    else:
        check("found OST html-only email", False)
    ost.close()

    # ---- PST: folder load, checkboxes, preview + attachments ----
    pst = PstReader()
    pst.open(PST)
    application.reader = pst
    folders = pst.get_folder_tree()
    application._populate_folder_tree(folders)
    application.header_file_label.config(text=os.path.basename(PST))
    root.title(f"{os.path.basename(PST)} - app")
    root.update()
    check("folder tree populated", len(application.folder_map) > 0,
          f"{len(application.folder_map)} folders")
    check("header shows open file",
          application.header_file_label.cget("text") == os.path.basename(PST))

    misha = folder_by_name(pst, "misha")
    messages = pst.get_messages(misha)
    application.current_folder = misha
    application._populate_item_list(messages)
    root.update()
    rows = application.item_list.get_children()
    check("item list populated", len(rows) == len(messages), f"{len(rows)} rows")

    # check the first two rows
    for row in rows[:2]:
        application._toggle_check(row)
    root.update()
    check("checkbox selection works", len(application.checked_items) == 2,
          f"checked={len(application.checked_items)}")
    check("checked items resolve to messages",
          len(application._get_action_messages()) == 2)

    # select an email that has an attachment -> preview + attachments panel
    att_row = None
    for row in rows:
        msg = application.message_map[row]
        if msg.has_attachments:
            att_row = row
            break
    if att_row is not None:
        application.item_list.selection_set(att_row)
        application._show_preview(application.message_map[att_row])
        root.update()
        check("attachments panel shows entries",
              len(application.att_tree.get_children()) > 0,
              f"{len(application.att_tree.get_children())} attachments")
        check("attachments panel is visible",
              bool(application.preview_att_frame.winfo_manager()))

    # ---- screenshot (best effort) ----
    saved = False
    try:
        from PIL import ImageGrab
        root.update_idletasks()
        root.update()
        x, y = root.winfo_rootx(), root.winfo_rooty()
        w, h = root.winfo_width(), root.winfo_height()
        img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
        img.save(SHOT)
        saved = os.path.exists(SHOT) and os.path.getsize(SHOT) > 0
    except Exception as exc:
        print(f"  (screenshot unavailable: {exc})")
    check("screenshot saved", saved, SHOT if saved else "skipped")

    pst.close()
    root.destroy()

    print(f"\nSMOKE RESULT: {'OK' if not failures else 'FAILURES: ' + ', '.join(failures)}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
