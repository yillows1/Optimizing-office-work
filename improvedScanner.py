"""
improvedScanner.py — Email Database sync for Scanner Packing Lists.

Only ingests emails whose Subject contains "DELIVERY TO" (case-insensitive).
Walks Inbox + all subfolders (Sent / Drafts / Deleted are never visited).
Parses the RITM#########.pdf attachment (ServiceNow catalog task export)
into a normalized set of SQLite tables so later scans read from the DB
instead of re-walking Outlook.
"""
from __future__ import annotations

import io
import json
import os
import re
import sqlite3
import tempfile
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable

import win32com.client

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DEFAULT_DB_NAME = "Email_Delivery.db"
DELIVERY_SUBJECT_TOKEN = "DELIVERY TO"

RITM_RE = re.compile(r"\bRITM[\s:#-]*\d+\b", re.IGNORECASE)

_KIND_DELIVERY = "delivery"
_KIND_PICKUP = "pickup"

_STREET_SUFFIX_RE = re.compile(
    r"\b\d+\s+[A-Za-z0-9 .']+?\s+"
    r"(?:St|Street|Rd|Road|Ave|Avenue|Blvd|Boulevard|Dr|Drive|Ln|Lane|"
    r"Ct|Court|Pkwy|Parkway|Way|Hwy|Highway|Pike|Ter|Terrace|Cir|Circle|"
    r"Pl|Place|Sq|Square|Trl|Trail|Fwy|Freeway)\b",
    re.IGNORECASE,
)
_ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")
_CITY_STATE_RE = re.compile(r"^[^,]+\d[^,]*, [^,]+, [A-Z]{2}\b")


def classify_subject(subject: str) -> tuple[str, str] | None:
    """
    Returns kind ∈ {'delivery','pickup'} or None if the email should be ignored.
    """
    upper = (subject or "").upper()
    if "DELIVERY TO" in upper:
        return _KIND_DELIVERY

    match = RITM_RE.search(subject or "")
    if not match:
        return None
    tail = subject[match.end():].strip(" -:,")
    if not tail:
        return None

    has_suffix = bool(_STREET_SUFFIX_RE.search(tail))
    has_zip = bool(_ZIP_RE.search(tail))
    has_city_state = bool(_CITY_STATE_RE.match(tail))

    if has_suffix and has_zip:
        return _KIND_PICKUP
    if has_suffix or has_zip:
        return _KIND_PICKUP
    if has_city_state:
        return _KIND_PICKUP
    return None

# SerialNumber lives inside the "AssetClass:…-SerialNumber:…-…" summary line.
_SERIAL_RE = re.compile(r"SerialNumber:\s*([^-|]+?)\s*(?:-|\||$)", re.IGNORECASE)
_ASSET_CLASS_RE = re.compile(r"AssetClass:\s*([^-|]+?)\s*(?:-|\||$)", re.IGNORECASE)
_JOB_TYPE_RE = re.compile(r"JobType:\s*([^-|]*?)\s*(?:-|\||$)", re.IGNORECASE)
_QUANTITY_RE = re.compile(r"Quantity:\s*(\d+)", re.IGNORECASE)
_DESC_RE = re.compile(
    r"Description:\s*(.*?)(?:\s*Line Item Count|\s*$)", re.IGNORECASE | re.DOTALL
)

_WEIGHT_KEYS = ("ApproximateCableWeight",)
_GROSS_RE = re.compile(r"Gross(?:Weight)?:\s*(\d+)", re.IGNORECASE)
_TARE_RE = re.compile(r"Tare(?:Weight)?:\s*(\d+)", re.IGNORECASE)
_NET_RE = re.compile(r"Net(?:Weight)?:\s*(\d+)", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Helper text detection — skip "More information", table separators, etc.
# ---------------------------------------------------------------------------
_HELPER_TEXT_PATTERNS = [
    r"^more information$",
    r"^please specify",
    r"^does this",
    r"^alternate",
    r"^site type",
    r"^is valid",
    r"^pick.up instructions?",
    r"^contact",
    r"^preferred",
    r"^employee number",
    r"^configuration item",
    r"^watch list",
    r"^approval",
    r"^assignment group",
    r"^assigned to",
    r"^due date",
    r"^integration task",
    r"^follow up",
    r"^kpi relief",
    r"^short description",
    r"^description:",
    r"^relevant article",
    r"^variables$",
    r"^requester",
    r"^project information",
    r"^project name",
    r"^disposal reason",
    r"^security case",
    r"^hazardous material",
    r"^project number",
    r"^po number",
    r"^equipment location",
    r"^general information",
    r"^line of business",
    r"^service requested",
    r"^other$",
    r"^asset class",
    r"^data upload",
    r"^adr equipment",
    r"^line item count",
    r"^financial location",
    r"^shipping and transport",
    r"^packaging acknowledgement",
    r"^additional information",
    r"^comments$",
    r"^upload attachment",
    r"^vendor$",
    r"^cod required",
    r"^activities$",
    r"^additional comments",
    r"^work notes",
    r"^activity$",
    r"^task =",
    r"^appointment",
    r"^approvers$",
    r"^group approvals",
    r"^parent =",
    r"^task slas",
    r"^affected user",
    r"^catalog item",
    r"^related attachments",
    r"^table name",
    r"^table sys id",
    r"^file name",
    r"^content type",
    r"^no records to display",
    r"^this may be due to security",
    r"^attachment",
    r"^sent:",
    r"^changed:",
    r"^system$",
    r"^historical notes",
    r"^catalog task",
    r"^request task",
    r"^number$",
    r"^item$",
    r"^opened by",
    r"^requested for",
    r"^request item",
    r"^state$",
    r"^follow up",
]

_HELPER_RE = re.compile("|".join(_HELPER_TEXT_PATTERNS), re.IGNORECASE)

_HELPER_WORDS = {
    "more information", "pick-up address", "pick-up instructions",
    "site type", "alternate address", "is valid address",
    "equipment type", "equipment location", "general information",
    "line of business", "service requested", "asset class",
    "data upload", "financial location", "shipping and transport",
    "additional information", "comments", "upload attachment",
    "vendor", "cod required", "activities", "work notes",
    "task slas", "affected user", "catalog item", "related attachments",
    "table name", "table sys id", "file name", "content type",
    "historical notes", "catalog task", "request task", "number",
    "item", "opened by", "requested for", "request item", "state",
    "assignment group", "assigned to", "due date", "integration task",
    "kpi relief", "short description", "relevant article", "variables",
    "requester", "project information", "project name", "disposal reason",
    "security case", "hazardous material", "project number", "po number",
    "general information", "line of business", "service requested",
    "asset class", "data upload", "financial location",
    "shipping and transport", "packaging acknowledgement",
    "additional information", "comments", "upload attachment",
    "vendor", "cod required",
}


def _is_helper_text(line: str) -> bool:
    """Return True if line is just helper/label text, not actual data."""
    stripped = line.strip()
    if not stripped:
        return True
    # Skip if it matches helper patterns
    if _HELPER_RE.search(stripped):
        return True
    # Skip table separators like "| - |" or "|  |"
    if re.fullmatch(r"\|?\s*[-:]?\s*\|?", stripped):
        return True
    # Skip short label-like phrases
    if len(stripped.split()) <= 3 and stripped.lower() in _HELPER_WORDS:
        return True
    return False


def _first_line_after(text: str, label: str, skip_helpers: bool = True) -> str:
    lines = (text or "").splitlines()
    label_lower = label.casefold()
    for index, line in enumerate(lines):
        if label_lower in line.casefold():
            # Try the same line first (e.g. "Location   205 Andrews Rd …")
            remainder = re.sub(
                rf".*?{re.escape(label)}\s*:?\s*", "", line, count=1,
                flags=re.IGNORECASE,
            ).strip()
            if remainder and not (skip_helpers and _is_helper_text(remainder)):
                return remainder
            # Look ahead further (up to 12 lines) and skip helper text
            for following in lines[index + 1:index + 13]:
                if following.strip() and not (skip_helpers and _is_helper_text(following)):
                    return following.strip()
    return ""


def _extract_address_from_pickup(text: str) -> str:
    """
    Extract address from Pick-up Address section, with fallback to Location.
    Returns the best address found.
    """
    # Try Pick-up Address first
    addr = _first_line_after(text, "Pick-up Address")
    if addr:
        return addr
    # Fallback to Location field
    addr = _first_line_after(text, "Location")
    if addr:
        return addr
    # Last resort: try Pick-up Instructions
    addr = _first_line_after(text, "Pick-up Instructions")
    return addr


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------
_SCHEMA = """
CREATE TABLE IF NOT EXISTS Email_Delivery (
    Email_ID          INTEGER PRIMARY KEY AUTOINCREMENT,
    Entry_ID          TEXT NOT NULL,
    Store_ID          TEXT,
    Folder            TEXT NOT NULL,
    Subject           TEXT NOT NULL,
    Sender            TEXT,
    Sender_Email      TEXT,
    Received_Date     TEXT NOT NULL,
    Address           TEXT,
    Type              TEXT NOT NULL DEFAULT 'delivery',
    Status            TEXT NOT NULL DEFAULT 'pending',
    UNIQUE (Entry_ID, Store_ID)
);
CREATE INDEX IF NOT EXISTS idx_delivery_received ON Email_Delivery(Received_Date);
CREATE INDEX IF NOT EXISTS idx_delivery_type     ON Email_Delivery(Type);

CREATE TABLE IF NOT EXISTS Email_Attachment (
    Attachment_ID     INTEGER PRIMARY KEY AUTOINCREMENT,
    Email_ID          INTEGER NOT NULL REFERENCES Email_Delivery(Email_ID) ON DELETE CASCADE,
    Filename          TEXT NOT NULL,
    UNIQUE (Email_ID, Filename)
);

CREATE TABLE IF NOT EXISTS Email_Ritm (
    Email_Ritm_ID     INTEGER PRIMARY KEY AUTOINCREMENT,
    Email_ID          INTEGER NOT NULL REFERENCES Email_Delivery(Email_ID) ON DELETE CASCADE,
    Ritm              TEXT NOT NULL,
    Attachment_Filename TEXT,
    UNIQUE (Email_ID, Ritm)
);
CREATE INDEX IF NOT EXISTS idx_ritm ON Email_Ritm(Ritm);

CREATE TABLE IF NOT EXISTS Email_Item (
    Email_Item_ID     INTEGER PRIMARY KEY AUTOINCREMENT,
    Email_Ritm_ID     INTEGER NOT NULL REFERENCES Email_Ritm(Email_Ritm_ID) ON DELETE CASCADE,
    Item_ID           TEXT NOT NULL,
    Receiving_Date    TEXT NOT NULL,
    Ritm              TEXT NOT NULL,
    Asset_Class       TEXT,
    Job_Type          TEXT,
    Equipment_Type    TEXT,
    Quantity          INTEGER,
    Description       TEXT,
    Gross             INTEGER,
    Tare              INTEGER,
    Net               INTEGER,
    UNIQUE (Email_Ritm_ID, Item_ID)
);
CREATE INDEX IF NOT EXISTS idx_item_lookup ON Email_Item(Item_ID, Receiving_Date);
CREATE INDEX IF NOT EXISTS idx_item_ritm   ON Email_Item(Ritm);
"""

_SCHEMA_LOCK = threading.Lock()
_SCHEMA_READY: dict[str, bool] = {}


def _resolve_db_path(db_path: str | os.PathLike | None) -> Path:
    if db_path:
        return Path(db_path)
    return Path(__file__).resolve().parent / DEFAULT_DB_NAME


def _connect(db_path: str | os.PathLike | None) -> sqlite3.Connection:
    path = _resolve_db_path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=15.0, isolation_level=None)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    key = str(path.resolve()).casefold()
    if not _SCHEMA_READY.get(key):
        with _SCHEMA_LOCK:
            if not _SCHEMA_READY.get(key):
                conn.executescript(_SCHEMA)
                _SCHEMA_READY[key] = True
    return conn


def get_last_received_date(db_path: str | os.PathLike | None) -> str | None:
    conn = _connect(db_path)
    try:
        row = conn.execute(
            "SELECT MAX(Received_Date) FROM Email_Delivery"
        ).fetchone()
        return row[0] if row and row[0] else None
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Attachment text extraction
# ---------------------------------------------------------------------------
def _read_attachment_bytes(attachment) -> bytes | None:
    try:
        suffix = os.path.splitext(attachment.FileName)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = tmp.name
        attachment.SaveAsFile(tmp_path)
        with open(tmp_path, "rb") as f:
            data = f.read()
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        return data
    except Exception as exc:
        print(f"  [!] Could not read {attachment.FileName}: {exc}")
        return None


def _extract_pdf_text(data: bytes) -> str:
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        print(f"  [!] Could not parse PDF: {exc}")
        return ""


# ---------------------------------------------------------------------------
# Field extraction from the RITM PDF
# ---------------------------------------------------------------------------
def find_ritms_in_text(text: str) -> list[str]:
    return sorted({
        re.sub(r"[^A-Za-z0-9]", "", match).upper()
        for match in RITM_RE.findall(text or "")
    })


def _split_asset_class_lines(text: str) -> list[str]:
    """Return joined lines that contain 'AssetClass:…SerialNumber:…'."""
    if not text:
        return []
    out: list[str] = []
    buffer = ""
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            if buffer:
                out.append(buffer)
                buffer = ""
            continue
        buffer = f"{buffer} {stripped}".strip() if buffer else stripped
        if "SerialNumber:" in buffer:
            out.append(buffer)
            buffer = ""
    if buffer and "SerialNumber:" in buffer:
        out.append(buffer)
    return out


def _match_or_none(pattern: re.Pattern, text: str) -> str | None:
    match = pattern.search(text)
    return match.group(1).strip() if match else None


def _match_int(pattern: re.Pattern, text: str) -> int | None:
    match = pattern.search(text)
    return int(match.group(1)) if match else None


def parse_ritm_pdf(text: str, filename: str) -> dict:
    """
    Return a dict with RITM, Location, Location_Address, and Items list.
    Items derive from the "AssetClass:…-SerialNumber:…-…" summary line.
    """
    ritms = find_ritms_in_text(text)
    primary_ritm = ritms[0] if ritms else ""

    location = _extract_address_from_pickup(text)

    equipment_type = _first_line_after(text, "Equipment Type")

    items: list[dict] = []
    for asset_line in _split_asset_class_lines(text):
        serial_match = _SERIAL_RE.search(asset_line)
        if not serial_match:
            continue
        serial = serial_match.group(1).strip()
        if not serial:
            continue

        asset_class = _match_or_none(_ASSET_CLASS_RE, asset_line)
        job_type = _match_or_none(_JOB_TYPE_RE, asset_line)
        quantity_match = _QUANTITY_RE.search(asset_line)
        quantity = int(quantity_match.group(1)) if quantity_match else None
        desc_match = _DESC_RE.search(asset_line)
        description = desc_match.group(1).strip() if desc_match else ""

        gross = _match_int(_GROSS_RE, asset_line)
        tare = _match_int(_TARE_RE, asset_line)
        net = _match_int(_NET_RE, asset_line)

        items.append({
            "item_id": serial,
            "asset_class": asset_class,
            "job_type": job_type,
            "equipment_type": equipment_type,
            "quantity": quantity,
            "description": description,
            "gross": gross,
            "tare": tare,
            "net": net,
        })

    return {
        "ritm": primary_ritm,
        "location": location,
        "items": items,
        "source_filename": filename,
    }


# ---------------------------------------------------------------------------
# Outlook traversal
# ---------------------------------------------------------------------------
def _walk_inbox_only(namespace) -> list:
    """Yield every folder under Inbox (Inbox itself + all descendants)."""
    try:
        root = namespace.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        return []
    out = []
    stack = [root]
    while stack:
        folder = stack.pop()
        out.append(folder)
        try:
            for sub in folder.Folders:
                stack.append(sub)
        except Exception:
            continue
    return out


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def sync_delivery_emails(
    db_path: str | os.PathLike | None = None,
    force_full: bool = False,
    progress_cb: Callable[[str], None] | None = None,
) -> dict:
    """
    Walk the Inbox tree, ingest any "DELIVERY TO" emails newer than
    MAX(Received_Date) in the Email DB, parse their RITM PDFs, and persist.
    Returns a summary dict.
    """
    def emit(msg: str) -> None:
        print(msg)
        if progress_cb:
            try:
                progress_cb(msg)
            except Exception:
                pass

    watermark = None if force_full else get_last_received_date(db_path)
    if watermark:
        emit(f"Watermark (latest in DB): {watermark}")
        try:
            cutoff = datetime.fromisoformat(watermark)
        except ValueError:
            cutoff = None
    else:
        emit("First run (or force full): scanning entire Inbox tree.")
        cutoff = None

    try:
        namespace = win32com.client.Dispatch(
            "Outlook.Application"
        ).GetNamespace("MAPI")
    except Exception as exc:
        raise RuntimeError(f"Could not open Outlook: {exc}") from exc

    folders = _walk_inbox_only(namespace)
    emit(f"Scanning {len(folders)} folder(s) under Inbox (Sent excluded).")

    candidates: list[dict] = []
    for folder in folders:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            count = items.Count
        except Exception:
            continue
        if not count:
            continue
        for message in items:
            try:
                received = message.ReceivedTime
            except Exception:
                continue
            if received is None or getattr(received, "year", 0) < 1900:
                continue
            received_dt = datetime(
                received.year, received.month, received.day,
                received.hour, received.minute, received.second,
            )
            if cutoff and received_dt <= cutoff:
                break  # sorted newest-first
            try:
                subject = str(message.Subject or "")
            except Exception:
                subject = ""
            kind = classify_subject(subject)
            if kind is None:
                continue
            candidates.append({
                "message": message,
                "folder": folder,
                "received_dt": received_dt,
                "subject": subject,
                "kind": kind,
            })

    emit(f"Matched {len(candidates)} delivery/pickup email(s) to ingest.")

    inserted = 0
    skipped = 0
    for index, entry in enumerate(candidates, start=1):
        message = entry["message"]
        folder = entry["folder"]
        try:
            ok = _ingest_one_email(
                message, folder, db_path,
                kind=entry["kind"],
            )
        except Exception as exc:
            emit(f"  [!] Failed on '{entry['subject']}': {exc}")
            ok = False
        if ok:
            inserted += 1
        else:
            skipped += 1
        if index % 25 == 0:
            emit(f"  …ingested {index}/{len(candidates)}")

    total = _count_delivery_rows(db_path)
    emit(f"Done. Inserted: {inserted}, skipped: {skipped}, total in DB: {total}")
    return {
        "inserted": inserted,
        "skipped": skipped,
        "total": total,
        "watermark": get_last_received_date(db_path),
    }


def _ingest_one_email(
    message, folder, db_path, kind: str = "delivery",
) -> bool:
    try:
        entry_id = str(message.EntryID or "")
    except Exception:
        entry_id = ""
    try:
        store_id = str(folder.StoreID or "")
    except Exception:
        store_id = ""
    if not entry_id:
        return False

    subject = str(message.Subject or "")
    try:
        sender_name = str(message.SenderName or "")
    except Exception:
        sender_name = ""
    try:
        sender_email = str(message.SenderEmailAddress or "")
    except Exception:
        sender_email = ""

    received = message.ReceivedTime
    received_iso = datetime(
        received.year, received.month, received.day,
        received.hour, received.minute, received.second,
    ).isoformat(sep=" ")

    attachments: list[tuple[str, bytes]] = []
    try:
        for att in message.Attachments:
            filename = str(att.FileName or "")
            data = _read_attachment_bytes(att)
            if data is not None:
                attachments.append((filename, data))
    except Exception:
        pass

    ritms_in_subject = find_ritms_in_text(subject)

    parsed_ritms: list[dict] = []
    seen_ritms: set[str] = set()
    for filename, data in attachments:
        if not filename.lower().endswith(".pdf"):
            continue
        text = _extract_pdf_text(data)
        if not text:
            continue
        parsed = parse_ritm_pdf(text, filename)
        ritm = parsed["ritm"]
        if not ritm or ritm in seen_ritms:
            continue
        seen_ritms.add(ritm)
        parsed_ritms.append(parsed)

    # If subject mentioned a RITM but no matching PDF, still record the RITM
    # with a NULL attachment so later lookups can find "delivery w/o PDF".
    for ritm in ritms_in_subject:
        if ritm not in seen_ritms:
            seen_ritms.add(ritm)
            parsed_ritms.append({
                "ritm": ritm,
                "location": "",
                "items": [],
                "source_filename": None,
            })

    if not parsed_ritms and not attachments:
        return False

    conn = _connect(db_path)
    try:
        conn.execute("BEGIN")
        try:
            cur = conn.execute(
                "INSERT OR IGNORE INTO Email_Delivery "
                "(Entry_ID, Store_ID, Folder, Subject, Sender, Sender_Email, "
                " Received_Date, Address, Type, Status) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    entry_id, store_id, folder.Name, subject,
                    sender_name, sender_email, received_iso,
                    parsed_ritms[0]["location"] if parsed_ritms else "",
                    kind, "pending",
                ),
            )
            if cur.rowcount == 0:
                conn.execute("COMMIT")
                return False  # already in DB (UNIQUE on Entry_ID, Store_ID)
            email_pk = cur.lastrowid

            for parsed in parsed_ritms:
                att_name = parsed["source_filename"]
                cur = conn.execute(
                    "INSERT OR IGNORE INTO Email_Ritm "
                    "(Email_ID, Ritm, Attachment_Filename) VALUES (?, ?, ?)",
                    (email_pk, parsed["ritm"], att_name),
                )
                if cur.rowcount == 0:
                    row = conn.execute(
                        "SELECT Email_Ritm_ID FROM Email_Ritm "
                        "WHERE Email_ID = ? AND Ritm = ?",
                        (email_pk, parsed["ritm"]),
                    ).fetchone()
                    if row is None:
                        continue
                    ritm_pk = row[0]
                else:
                    ritm_pk = cur.lastrowid

                for item in parsed["items"]:
                    conn.execute(
                        "INSERT OR IGNORE INTO Email_Item "
                        "(Email_Ritm_ID, Item_ID, Receiving_Date, Ritm, "
                        " Asset_Class, Job_Type, Equipment_Type, Quantity, "
                        " Description, Gross, Tare, Net) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (
                            ritm_pk, item["item_id"], received_iso,
                            parsed["ritm"],
                            item["asset_class"], item["job_type"],
                            item["equipment_type"], item["quantity"],
                            item["description"],
                            item["gross"], item["tare"], item["net"],
                        ),
                    )
            conn.execute("COMMIT")
            return True
        except Exception:
            try:
                conn.execute("ROLLBACK")
            except Exception:
                pass
            raise
    finally:
        conn.close()


def _count_delivery_rows(db_path) -> int:
    conn = _connect(db_path)
    try:
        row = conn.execute("SELECT COUNT(*) FROM Email_Delivery").fetchone()
        return int(row[0]) if row else 0
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Read-side API used by PackingListApp
# ---------------------------------------------------------------------------
def fetch_delivery_emails(
    db_path: str | os.PathLike | None = None,
    since: str | None = None,
    ritm_filter: str | None = None,
    kind_filter: str | None = None,
) -> list[dict]:
    """
    Return every ingested delivery email as a list of dicts shaped close to
    the old ExactMatch structure so PackingListApp can reuse the pipeline.
    `since` is ISO date/time inclusive. `ritm_filter` limits to one RITM.
    """
    conn = _connect(db_path)
    try:
        sql = (
            "SELECT d.Email_ID, d.Entry_ID, d.Store_ID, d.Folder, d.Subject, "
            "       d.Sender, d.Sender_Email, d.Received_Date, "
            "       d.Address, d.Type, d.Status, "
            "       r.Ritm, r.Attachment_Filename "
            "FROM Email_Delivery d "
            "LEFT JOIN Email_Ritm r ON r.Email_ID = d.Email_ID "
        )
        clauses: list[str] = []
        params: list[Any] = []
        if since:
            clauses.append("d.Received_Date >= ?")
            params.append(since)
        if ritm_filter:
            clauses.append("r.Ritm = ?")
            params.append(ritm_filter)
        if kind_filter:
            clauses.append("d.Type = ?")
            params.append(kind_filter)
        if clauses:
            sql += "WHERE " + " AND ".join(clauses) + " "
        sql += "ORDER BY d.Received_Date DESC, d.Email_ID DESC"

        rows = conn.execute(sql, params).fetchall()

        items_by_email_ritm = _load_items_for_email_ritms(conn)

        out: list[dict] = []
        for row in rows:
            (
                email_id, entry_id, store_id, folder, subject,
                sender, sender_email, received, address,
                type_value, status_value, ritm, attachment_filename,
            ) = row
            out.append({
                "term": ritm or "",
                "ritm": ritm or "",
                "filename": attachment_filename or "",
                "folder": folder or "",
                "subject": subject or "",
                "received": received or "",
                "sender": sender or "",
                "sender_email": sender_email or "",
                "entry_id": entry_id or "",
                "store_id": store_id or "",
                "ritms": [ritm] if ritm else [],
                "subject_ritms": [ritm] if ritm else [],
                "location": address or "",
                "location_address": address or "",
                "type": type_value or "",
                "status": status_value or "",
                "items": items_by_email_ritm.get((email_id, ritm or ""), []),
            })
        return out
    finally:
        conn.close()


def _load_items_for_email_ritms(conn) -> dict[tuple[int, str], list[dict]]:
    rows = conn.execute(
        "SELECT r.Email_ID, i.Ritm, i.Item_ID, i.Receiving_Date, "
        "       i.Asset_Class, i.Job_Type, i.Equipment_Type, i.Quantity, "
        "       i.Description, i.Gross, i.Tare, i.Net "
        "FROM Email_Item i "
        "JOIN Email_Ritm r ON r.Email_Ritm_ID = i.Email_Ritm_ID"
    ).fetchall()
    out: dict[tuple[int, str], list[dict]] = {}
    for (
        email_id, ritm, item_id, received, asset_class, job_type,
        equipment_type, quantity, description, gross, tare, net,
    ) in rows:
        out.setdefault((email_id, ritm), []).append({
            "description": item_id,
            "item_description": description or "",
            "type": asset_class or equipment_type or "",
            "gross": gross or 0,
            "tare": tare or 0,
            "net": net if net is not None else max((gross or 0) - (tare or 0), 0),
            "quantity": quantity or 1,
            "job_type": job_type or "",
            "item_id": item_id,
            "receiving_date": received,
        })
    return out


def find_items_by_ids(
    db_path: str | os.PathLike | None = None,
    item_ids: list[str] | None = None,
    received_date: str | None = None,
) -> list[dict]:
    """
    The efficient 1-to-many item lookup. Uses idx_item_lookup so an item ID
    that repeats across RITMs is disambiguated by Receiving_Date.
    """
    if not item_ids:
        return []
    conn = _connect(db_path)
    try:
        placeholders = ",".join("?" * len(item_ids))
        sql = (
            "SELECT i.Item_ID, i.Receiving_Date, i.Ritm, i.Asset_Class, "
            "       i.Job_Type, i.Equipment_Type, i.Quantity, i.Description, "
            "       i.Gross, i.Tare, i.Net, "
            "       d.Entry_ID, d.Store_ID, d.Folder, d.Subject, "
            "       r.Attachment_Filename "
            "FROM Email_Item i "
            "JOIN Email_Ritm r ON r.Email_Ritm_ID = i.Email_Ritm_ID "
            "JOIN Email_Delivery d ON d.Email_ID = r.Email_ID "
            f"WHERE i.Item_ID IN ({placeholders}) "
        )
        params: list[Any] = list(item_ids)
        if received_date:
            sql += "AND i.Receiving_Date = ? "
            params.append(received_date)
        rows = conn.execute(sql, params).fetchall()
        out: list[dict] = []
        for row in rows:
            out.append({
                "item_id": row[0], "receiving_date": row[1], "ritm": row[2],
                "asset_class": row[3], "job_type": row[4],
                "equipment_type": row[5], "quantity": row[6],
                "description": row[7], "gross": row[8], "tare": row[9],
                "net": row[10], "entry_id": row[11], "store_id": row[12],
                "folder": row[13], "subject": row[14],
                "attachment_filename": row[15],
            })
        return out
    finally:
        conn.close()


def get_ritm_context(
    db_path: str | os.PathLike | None = None,
    ritm: str | None = None,
) -> dict | None:
    if not ritm:
        return None
    conn = _connect(db_path)
    try:
        row = conn.execute(
            "SELECT d.Email_ID, d.Entry_ID, d.Store_ID, d.Folder, d.Subject, "
            "       d.Sender, d.Sender_Email, d.Received_Date, "
            "       d.Address, r.Attachment_Filename "
            "FROM Email_Ritm r "
            "JOIN Email_Delivery d ON d.Email_ID = r.Email_ID "
            "WHERE r.Ritm = ? "
            "ORDER BY d.Received_Date DESC LIMIT 1",
            (ritm,),
        ).fetchone()
        if row is None:
            return None
        items = conn.execute(
            "SELECT Item_ID, Asset_Class, Job_Type, Equipment_Type, Quantity, "
            "       Description, Gross, Tare, Net "
            "FROM Email_Item WHERE Ritm = ?",
            (ritm,),
        ).fetchall()
        return {
            "email_id": row[0], "entry_id": row[1], "store_id": row[2],
            "folder": row[3], "subject": row[4], "sender": row[5],
            "sender_email": row[6], "received": row[7],
            "location": row[8], "location_address": row[8],
            "attachment_filename": row[9],
            "items": [
                {
                    "item_id": r[0], "asset_class": r[1], "job_type": r[2],
                    "equipment_type": r[3], "quantity": r[4],
                    "description": r[5], "gross": r[6], "tare": r[7],
                    "net": r[8],
                }
                for r in items
            ],
        }
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _cli_main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="Sync DELIVERY TO emails.")
    parser.add_argument("--db", default=None, help="Path to Email_Delivery.db")
    parser.add_argument("--force-full", action="store_true")
    args = parser.parse_args()
    summary = sync_delivery_emails(args.db, force_full=args.force_full)
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    _cli_main()