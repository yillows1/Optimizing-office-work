import win32com.client
import os
import io
import tempfile
import re
import difflib
from collections import defaultdict
from datetime import datetime, timedelta

# --- Config ---
LIST_FILE = "receivedList.txt"
LOG_FILE = "scan_results.txt"

# Fuzzy matching settings
FUZZY_ENABLED = True       # set to False to disable near-miss detection
FUZZY_THRESHOLD = 0.85    # 0.0 = anything, 1.0 = exact. 0.85 ~ "one char off"
FUZZY_MIN_LEN = 6          # skip fuzzy matching for terms shorter than this

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
# fuzzy_by_email: EntryID -> { filename -> set of (term, guessed_word, ratio) }
fuzzy_by_email = defaultdict(lambda: defaultdict(set))
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

WORD_RE = re.compile(r"[A-Za-z0-9]+")

def find_near_misses(text, unmatched_terms):
    """
    For each unmatched term (long enough to be fuzzy-matched), look for
    a word in `text` that is close enough to be considered a likely typo.
    Returns {term: [(guessed_word, ratio), ...]}.
    """
    result = {}
    if not text or not unmatched_terms:
        return result

    words = set(WORD_RE.findall(text))
    if not words:
        return result
    words_lower = {w.lower(): w for w in words}

    for term in unmatched_terms:
        if len(term) < FUZZY_MIN_LEN:
            continue
        term_lower = term.lower()
        candidates = [
            (w_orig, w_low) for w_low, w_orig in words_lower.items()
            if abs(len(w_low) - len(term_lower)) <= 2
        ]
        hits = []
        for w_orig, w_low in candidates:
            ratio = difflib.SequenceMatcher(None, term_lower, w_low).ratio()
            if ratio >= FUZZY_THRESHOLD:
                hits.append((w_orig, ratio))
        if hits:
            hits.sort(key=lambda x: -x[1])
            result[term] = hits[:3]
    return result

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
        # --- TIME FILTER ---
        if cutoff is not None:
            try:
                received = message.ReceivedTime
                if received.year < 1900:
                    skipped_count += 1
                    folder_skipped += 1
                    continue
                received_dt = datetime(received.year, received.month, received.day,
                                       received.hour, received.minute, received.second)
                if received_dt < cutoff:
                    hit_cutoff = True
                    break
            except Exception:
                skipped_count += 1
                folder_skipped += 1
                continue

        try:
            if message.Attachments.Count == 0:
                continue
        except Exception:
            continue

        scanned_count += 1
        folder_scanned += 1

        email_terms = set()
        email_files = set()
        # filename -> content text, so we don't re-read attachments for fuzzy pass
        file_contents = {}
        # filename -> set of terms found exactly in that specific file
        file_found_terms = defaultdict(set)

        for attachment in message.Attachments:
            filename = attachment.FileName
            data = read_attachment_bytes(attachment)
            if data is None:
                continue

            content = extract_text_from_bytes(data, filename)
            file_contents[filename] = content
            hits = find_terms_in_text(content)

            if hits:
                email_terms |= hits
                email_files.add(filename)
                file_found_terms[filename] |= hits
                print(f"  🔎 {filename} → {', '.join(sorted(hits))}")

        # --- Fuzzy pass: only for terms NOT already found in that specific file ---
        if FUZZY_ENABLED and email_files:
            try:
                entry_key = message.EntryID
            except Exception:
                entry_key = f"{message.Subject}|{message.ReceivedTime}"

            for filename in email_files:
                content = file_contents.get(filename, "")
                # Only fuzzy-match terms NOT already found in THIS file
                unmatched_for_file = [
                    t for t in SEARCH_TERMS
                    if t not in file_found_terms.get(filename, set())
                ]
                near = find_near_misses(content, unmatched_for_file)
                if near:
                    for term, hits in near.items():
                        for guess_word, ratio in hits:
                            fuzzy_by_email[entry_key][filename].add(
                                (term, guess_word, round(ratio, 3))
                            )

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

fuzzy_terms_anywhere = set()
for fmap in fuzzy_by_email.values():
    for hint_set in fmap.values():
        for term, _, _ in hint_set:
            fuzzy_terms_anywhere.add(term)

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
if FUZZY_ENABLED:
    w(f"Fuzzy matching:              ON (threshold {FUZZY_THRESHOLD}, min length {FUZZY_MIN_LEN})")
else:
    w(f"Fuzzy matching:              OFF")
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

# frozenset of exact terms -> list of (EntryID, email entry)
groups = defaultdict(list)
for entry_id, e in matches_by_email.items():
    key = frozenset(e["terms"])
    groups[key].append((entry_id, e))

def group_sort_key(item):
    terms, emails = item
    newest = max(e["received"] for _, e in emails)
    return (-len(terms), newest)

ordered_groups = sorted(groups.items(), key=group_sort_key)

def files_for_group(emails):
    files = set()
    for _, e in emails:
        files |= e["files"]
    return sorted(files)

letter = ord("A")
for terms, emails in ordered_groups:
    files = files_for_group(emails)
    file_label = ", ".join(files) if files else "(unknown attachment)"
    w(f"{chr(letter)}) {file_label}")
    for t in sorted(terms):
        w(f"   {t}")

    # Lowercase set of exact terms in this group, for quick lookup
    terms_lower = {t.lower() for t in terms}

    # Near-miss suggestions scoped to this group's files only
    suggestion_map = defaultdict(set)
    for entry_id, _ in emails:
        fmap = fuzzy_by_email.get(entry_id)
        if not fmap:
            continue
        for filename, hint_set in fmap.items():
            if filename in files:
                for term, guess, ratio in hint_set:
                    # Skip if this search term is already an exact match in this group
                    if term.lower() in terms_lower:
                        continue
                    # Skip if the guessed word is itself an exact match in this group
                    if guess.lower() in terms_lower:
                        continue
                    suggestion_map[filename].add((term, guess, ratio))

    if suggestion_map:
        w()
        w("   ⚠ Possible near-miss (needs human check):")
        for filename in sorted(suggestion_map):
            for term, guess, ratio in sorted(suggestion_map[filename]):
                w(f"     {filename}: '{guess}' ~ {term}  (similarity {ratio})")
    w()
    letter += 1

# Final bucket: terms never found exactly AND never got a near-miss hint
truly_missing = [t for t in SEARCH_TERMS
                 if t not in found_terms and t not in fuzzy_terms_anywhere]

if truly_missing:
    w(f"{chr(letter)}) N/A")
    for t in sorted(truly_missing):
        w(f"   {t}")
    w()
    letter += 1

with open(LOG_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"\nReport saved to: {LOG_FILE}")
input("\nPress Enter to exit...")