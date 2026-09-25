import win32com.client
import os
import io
import tempfile
from collections import defaultdict
from datetime import datetime, timedelta

# --- Config ---
LIST_FILE = "receivedList.txt"
LOG_FILE = "scan_results.txt"

# --- Load search terms ---
if not os.path.exists(LIST_FILE):
    print(f"Could not find {LIST_FILE} in the current folder.")
    input("Press Enter to exit...")
    exit(1)

with open(LIST_FILE, "r", encoding="utf-8") as f:
    SEARCH_TERMS = [line.strip() for line in f if line.strip()]

if not SEARCH_TERMS:
    print(f"{LIST_FILE} is empty.")
    input("Press Enter to exit...")
    exit(1)

# --- Ask for time range ---
print("\nChoose a time range to scan:")
print("  (1) Past 3 months")
print("  (2) Past 6 months")
print("  (3) Past 12 months")
print("  (4) All time")
choice = input("Selection [2]: ").strip() or "2"

now = datetime.now()
if choice == "1":
    cutoff = now - timedelta(days=90)
    range_label = "Past 3 months"
elif choice == "2":
    cutoff = now - timedelta(days=182)
    range_label = "Past 6 months"
elif choice == "3":
    cutoff = now - timedelta(days=365)
    range_label = "Past 12 months"
elif choice == "4":
    cutoff = None
    range_label = "All time"
else:
    print("Invalid choice. Defaulting to past 6 months.")
    cutoff = now - timedelta(days=182)
    range_label = "Past 6 months"

print(f"\nTime range:   {range_label}")
if cutoff:
    print(f"Cutoff date:  {cutoff.strftime('%Y-%m-%d')}")
print(f"Search terms: {len(SEARCH_TERMS)} loaded from {LIST_FILE}")
for t in SEARCH_TERMS:
    print(f"  - {t}")
print()

# --- Connect to Outlook ---
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

matches_by_email = defaultdict(lambda: {"terms": set(), "files": set()})
found_terms = set()
scanned_count = 0
skipped_count = 0

def read_attachment_bytes(attachment):
    try:
        suffix = os.path.splitext(attachment.FileName)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = tmp.name
        attachment.SaveAsFile(tmp_path)
        with open(tmp_path, "rb") as f:
            data = f.read()
        os.remove(tmp_path)
        return data
    except Exception as e:
        print(f"  [!] Could not read {attachment.FileName}: {e}")
        return None

def extract_text_from_bytes(data, filename):
    ext = os.path.splitext(filename)[1].lower()
    try:
        if ext in (".txt", ".csv", ".log", ".json", ".xml", ".html", ".md"):
            return data.decode("utf-8", errors="ignore")
        elif ext == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(data))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        elif ext == ".docx":
            from docx import Document
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs)
        elif ext == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(data), data_only=True)
            parts = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    parts.append(" ".join(str(c) for c in row if c is not None))
            return "\n".join(parts)
        else:
            return data.decode("latin-1", errors="ignore")
    except Exception as e:
        print(f"  [!] Could not parse {filename}: {e}")
        return ""

def find_terms_in_text(text):
    if not text:
        return set()
    lowered = text.lower()
    return {term for term in SEARCH_TERMS if term.lower() in lowered}

def get_folders(folder):
    """Yield folder and all subfolders, recursively."""
    yield folder
    for sub in folder.Folders:
        yield from get_folders(sub)

folders_to_scan = list(get_folders(inbox))
print(f"Scanning Inbox + all subfolders ({len(folders_to_scan)} folder(s))...\n")

# --- Walk through emails ---
for folder in folders_to_scan:
    messages = folder.Items
    try:
        # Sort newest-first so we can break early once we pass the cutoff
        messages.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    msg_count = messages.Count
    if msg_count == 0:
        continue

    folder_scanned = 0
    folder_skipped = 0
    hit_cutoff = False

    for message in messages:
        # --- TIME FILTER (cheap check, happens before touching attachments) ---
        if cutoff is not None:
            try:
                received = message.ReceivedTime
                # ReceivedTime comes back as a pywintypes.datetime
                if received.year < 1900:
                    # Some items have bogus dates; treat as unknown and skip
                    skipped_count += 1
                    folder_skipped += 1
                    continue
                # Convert to naive datetime for comparison
                received_dt = datetime(received.year, received.month, received.day,
                                       received.hour, received.minute, received.second)
                if received_dt < cutoff:
                    # Since sorted newest-first, everything after this is older too
                    hit_cutoff = True
                    break
            except Exception:
                # If we can't read the date, skip the email rather than scan it
                skipped_count += 1
                folder_skipped += 1
                continue

        # --- Skip emails with no attachments (also cheap) ---
        try:
            if message.Attachments.Count == 0:
                continue
        except Exception:
            continue

        # --- We only reach here for in-range emails with attachments ---
        scanned_count += 1
        folder_scanned += 1

        email_terms = set()
        email_files = set()

        for attachment in message.Attachments:
            filename = attachment.FileName
            data = read_attachment_bytes(attachment)
            if data is None:
                continue

            content = extract_text_from_bytes(data, filename)
            hits = find_terms_in_text(content)

            if hits:
                email_terms |= hits
                email_files.add(filename)
                print(f"  🔎 {filename} → {', '.join(sorted(hits))}")

        if email_terms:
            try:
                key = message.EntryID
            except Exception:
                key = f"{message.Subject}|{message.ReceivedTime}"

            entry = matches_by_email[key]
            entry["subject"] = message.Subject
            entry["sender"] = message.SenderName
            entry["received"] = message.ReceivedTime
            entry["folder"] = folder.Name
            entry["terms"] |= email_terms
            entry["files"] |= email_files
            found_terms |= email_terms

        if folder_scanned % 50 == 0:
            print(f"  ...scanned {folder_scanned} email(s) in [{folder.Name}]")

    if folder_scanned or folder_skipped:
        print(f"[{folder.Name}] scanned {folder_scanned}, skipped {folder_skipped}"
              + (" (hit cutoff, stopped early)" if hit_cutoff else ""))

# --- Bucket results ---
multi_term_emails = [e for e in matches_by_email.values() if len(e["terms"]) >= 2]
single_term_emails = [e for e in matches_by_email.values() if len(e["terms"]) == 1]
missing_terms = [t for t in SEARCH_TERMS if t not in found_terms]

# --- Write report ---
lines = []
def w(s=""):
    lines.append(s)
    print(s)

w()
w("=" * 60)
w(f"SEARCH SUMMARY — {len(SEARCH_TERMS)} term(s) searched")
w("=" * 60)
w(f"Time range:                  {range_label}")
if cutoff:
    w(f"Cutoff date:                 {cutoff.strftime('%Y-%m-%d')}")
w(f"Emails scanned (in range):   {scanned_count}")
w(f"Emails skipped (out of range / no date): {skipped_count}")
w(f"Emails with 2+ terms:        {len(multi_term_emails)}")
w(f"Emails with exactly 1 term:  {len(single_term_emails)}")
w(f"Terms never found:           {len(missing_terms)}")
w()

w("=" * 60)
w("SECTION 1: MULTIPLE TERMS IN THE SAME EMAIL")
w("=" * 60)
if not multi_term_emails:
    w("(none)")
else:
    for e in sorted(multi_term_emails, key=lambda x: x["received"], reverse=True):
        w(f"\n📧 {e['subject']}")
        w(f"   From:     {e['sender']}")
        w(f"   Received: {e['received']}")
        w(f"   Folder:   {e['folder']}")
        w(f"   Terms:    {', '.join(sorted(e['terms']))}")
        w(f"   Files:    {', '.join(sorted(e['files']))}")

w()
w("=" * 60)
w("SECTION 2: SINGLE TERM MATCHES")
w("=" * 60)
if not single_term_emails:
    w("(none)")
else:
    for e in sorted(single_term_emails, key=lambda x: x["received"], reverse=True):
        w(f"\n📧 {e['subject']}")
        w(f"   From:     {e['sender']}")
        w(f"   Received: {e['received']}")
        w(f"   Folder:   {e['folder']}")
        w(f"   Term:     {', '.join(sorted(e['terms']))}")
        w(f"   Files:    {', '.join(sorted(e['files']))}")

w()
w("=" * 60)
w("SECTION 3: TERMS NEVER FOUND")
w("=" * 60)
if not missing_terms:
    w("(none — every term was found at least once)")
else:
    for t in missing_terms:
        w(f"  ❌ {t}")

# ============================================================
# SECTION 4: GROUPED BY CO-OCCURRENCE (lettered buckets)
# ============================================================
w()
w("=" * 60)
w("SECTION 4: GROUPED BY CO-OCCURRENCE")
w("=" * 60)
w()

# Build a mapping: frozenset of terms -> list of email entries with that exact term set
groups = defaultdict(list)
for e in matches_by_email.values():
    key = frozenset(e["terms"])
    groups[key].append(e)

# Sort groups so larger term-sets come first (most informative first),
# then by most recent email within the group.
def group_sort_key(item):
    terms, emails = item
    newest = max(e["received"] for e in emails)
    return (-len(terms), newest)

ordered_groups = sorted(groups.items(), key=group_sort_key)

def files_for_group(emails):
    files = set()
    for e in emails:
        files |= e["files"]
    return sorted(files)

letter = ord("A")
for terms, emails in ordered_groups:
    files = files_for_group(emails)
    file_label = ", ".join(files) if files else "(unknown attachment)"
    w(f"{chr(letter)}) {file_label}")
    for t in sorted(terms):
        w(f"   {t}")
    w()
    letter += 1

# Final bucket: terms that never showed up anywhere
if missing_terms:
    w(f"{chr(letter)}) N/A")
    for t in sorted(missing_terms):
        w(f"   {t}")
    w()
    letter += 1

with open(LOG_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"\nReport saved to: {LOG_FILE}")
input("\nPress Enter to exit...")
