import win32com.client
import os
import io

# --- Ask the user for input ---
SEARCH_TERM = input("Enter the search term: ").strip()
FOLDER_CHOICE = input("Search (1) Inbox only, (2) Inbox + subfolders [1]: ").strip() or "1"
LOG_FILE = "scan_results.txt"  # optional: write matches here instead of saving attachments

if not SEARCH_TERM:
    print("No search term provided. Exiting.")
    exit(1)

print(f"\nSearching for: '{SEARCH_TERM}'\n")

# --- Connect to Outlook ---
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

matches = []

def search_text(text, term):
    if not text:
        return False
    return term.lower() in text.lower()

def read_attachment_bytes(attachment):
    """
    Read attachment content into memory.
    Uses attachment.SaveAsFile to a temp path only if the in-memory read fails.
    """
    try:
        # Outlook exposes attachment.PropertyAccessor for binary streams
        # but the simplest reliable way is SaveAsFile to a temp file
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(attachment.FileName)[1]) as tmp:
            tmp_path = tmp.name
        attachment.SaveAsFile(tmp_path)
        with open(tmp_path, "rb") as f:
            data = f.read()
        os.remove(tmp_path)  # delete immediately
        return data
    except Exception as e:
        print(f"  [!] Could not read {attachment.FileName}: {e}")
        return None

def extract_text_from_bytes(data, filename):
    """Extract readable text from raw bytes based on file extension."""
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
            text_parts = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    text_parts.append(" ".join(str(c) for c in row if c is not None))
            return "\n".join(text_parts)
        
        else:
            # Fallback: raw byte scan for .doc, .xls, unknown types
            return data.decode("latin-1", errors="ignore")
    
    except Exception as e:
        print(f"  [!] Could not parse {filename}: {e}")
        return ""

def get_folders(folder, recurse):
    yield folder
    if recurse:
        for sub in folder.Folders:
            yield from get_folders(sub, recurse)

# --- Determine which folders to scan ---
folders_to_scan = list(get_folders(inbox, recurse=(FOLDER_CHOICE == "2")))
print(f"Scanning {len(folders_to_scan)} folder(s)...\n")

# --- Walk through folders and messages ---
for folder in folders_to_scan:
    messages = folder.Items
    try:
        messages.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    
    msg_count = messages.Count
    if msg_count == 0:
        continue
    
    print(f"[Folder: {folder.Name}] {msg_count} message(s)")
    
    for i, message in enumerate(messages, start=1):
        try:
            if message.Attachments.Count == 0:
                continue
        except Exception:
            continue
        
        for attachment in message.Attachments:
            filename = attachment.FileName
            data = read_attachment_bytes(attachment)
            if data is None:
                continue
            
            content = extract_text_from_bytes(data, filename)
            
            if search_text(content, SEARCH_TERM):
                matches.append({
                    "subject": message.Subject,
                    "sender": message.SenderName,
                    "received": message.ReceivedTime,
                    "file": filename,
                    "folder": folder.Name
                })
                print(f"  ✅ MATCH in {filename} (from: {message.Subject})")
            
            # temp file is already deleted inside read_attachment_bytes
        
        if i % 50 == 0:
            print(f"  ...checked {i}/{msg_count}")

# --- Final report ---
print("\n" + "=" * 50)
print(f"Search term: '{SEARCH_TERM}'")
print(f"Total matches: {len(matches)}")
print("=" * 50)

if not matches:
    print("No matches found.")
else:
    with open(LOG_FILE, "w", encoding="utf-8") as log:
        log.write(f"Search term: {SEARCH_TERM}\n")
        log.write(f"Matches: {len(matches)}\n\n")
        for m in matches:
            line = (
                f"File: {m['file']}\n"
                f"  Subject:  {m['subject']}\n"
                f"  From:     {m['sender']}\n"
                f"  Received: {m['received']}\n"
                f"  Folder:   {m['folder']}\n\n"
            )
            print(line)
            log.write(line)
    print(f"Results also written to: {LOG_FILE}")

input("\nPress Enter to exit...")
