import re
import os
import io
import gc
import json
import sqlite3
import html
import sys
import tempfile
import shutil
import time as _time
import functools as _functools
import logging as _logging
import threading
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, TypedDict

import pythoncom

from PySide6.QtCore import QEvent, QThread, Qt, Signal
from PySide6.QtGui import QIntValidator, QKeyEvent, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QPlainTextEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMenu, QMessageBox, QPushButton,
    QRadioButton, QScrollArea, QStyle, QSizePolicy, QTabWidget, QTableWidget,
    QTableWidgetItem, QToolButton, QVBoxLayout, QWidget,
)
import win32com.client

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from app_logging import configure_app_logging


# ---------------------------------------------------------------------------
# Config / paths
# ---------------------------------------------------------------------------

import sys as _sys
from pathlib import Path as _Path


def get_resource_path(relative: str) -> _Path:
    if getattr(_sys, "frozen", False) and hasattr(_sys, "_MEIPASS"):
        return _Path(_sys._MEIPASS) / relative
    return _Path(__file__).resolve().parent / relative


def get_user_dir() -> _Path:
    if getattr(_sys, "frozen", False):
        return _Path(_sys.executable).resolve().parent
    return _Path(__file__).resolve().parent


try:
    from version import VERSION
except Exception:
    VERSION = "dev"

_APP_LOGGER = configure_app_logging("Packing List")


def _timed(fn):
    @_functools.wraps(fn)
    def wrapper(*args, **kwargs):
        start = _time.perf_counter()
        try:
            return fn(*args, **kwargs)
        finally:
            elapsed = _time.perf_counter() - start
            _APP_LOGGER.info("%s: %.2fs", fn.__name__, elapsed)
    return wrapper


def _log_marker(text: str) -> None:
    _APP_LOGGER.info("--- %s ---", text)


CONFIG_PATH = get_user_dir() / "scanner_config.json"
DEFAULT_SUMMARIES_DIR = get_user_dir() / "Summaries"
DEFAULT_RECEIVED_FROM = "Verizon"

MAX_ITEMS_PER_LOAD = 23
MAX_ITEMS_PER_SUMMARY_TAB = 31

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105


def load_app_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "email_database_path": str(get_user_dir() / "Email_Delivery.db"),
    }


def save_app_config(data: dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


def parse_date(value: str) -> date:
    for date_format in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(value.strip(), date_format).date()
        except ValueError:
            continue
    raise ValueError("Enter the date as YYYY-MM-DD or MM/DD/YYYY.")


def normalize_item_type(value: str) -> str:
    cleaned = " ".join(value.split())
    folded = cleaned.casefold().replace("-", " ").replace("_", " ")
    folded = " ".join(folded.split())
    common_types = {
        "99 fo": "99 FO", "99fo": "99 FO", "fo99": "99 FO",
        "fo 99": "99 FO", "99 fo ": "99 FO",
        "23": "23", "type 23": "Type 23", "type23": "Type 23",
        "type 3": "Type 3", "type3": "Type 3",
        "type 4": "Type 4", "type4": "Type 4",
        "69s": "69S",
    }
    return common_types.get(folded, cleaned.upper())


def normalize_ritm(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value).upper()


def normalize_manual_ritm(value: str) -> str:
    normalized = normalize_ritm(value)
    if normalized.isdigit():
        normalized = f"RITM{normalized}"
    return normalized if re.fullmatch(r"RITM\d+", normalized) else ""


NO_REQUEST_ITEM = "No Request Item"


def _is_no_request_item(value: str | None) -> bool:
    return (
        not value
        or not value.strip()
        or value.strip().casefold() == NO_REQUEST_ITEM.casefold()
    )


def _normalize_optional_ritm(value: str | None) -> str:
    if _is_no_request_item(value):
        return NO_REQUEST_ITEM
    return str(value).strip()


def _ritm_group_key(value: str | None) -> str:
    return "" if _is_no_request_item(value) else str(value).strip().casefold()


def _match_mentions_ritm(match: "ExactMatch", ritm: str) -> bool:
    ritm_key = normalize_ritm(ritm)
    return (
        any(normalize_ritm(value) == ritm_key for value in match.ritms)
        or any(
            normalize_ritm(value) == ritm_key
            for value in match.subject_ritms
        )
        or normalize_ritm(match.term) == ritm_key
    )


# ---------------------------------------------------------------------------
# Database helpers (SQLite) — two independent DBs
# ---------------------------------------------------------------------------

_SQLITE_SCHEMA = """
CREATE TABLE IF NOT EXISTS Reel_Load_Database (
    Packing_List_Number  TEXT PRIMARY KEY,
    Receiving_Date       TEXT,
    RITM_Number          TEXT,
    Location             TEXT,
    Truck                TEXT,
    Truck_Number         TEXT,
    Total_Gross_Weight   INTEGER,
    Total_Tare           INTEGER,
    Total_Net_Weight     INTEGER,
    Number_of_Reels      INTEGER
);

CREATE TABLE IF NOT EXISTS Reel_Items (
    Item_ID              INTEGER PRIMARY KEY AUTOINCREMENT,
    Packing_List_Number  TEXT,
    RITM_Number          TEXT,
    Item_Description     TEXT,
    Item_Type            TEXT,
    Gross                INTEGER,
    Tare                 INTEGER,
    Net                  INTEGER,
    FOREIGN KEY (Packing_List_Number)
        REFERENCES Reel_Load_Database(Packing_List_Number)
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_ritm_item
    ON Reel_Items (RITM_Number, Item_Description);

CREATE INDEX IF NOT EXISTS idx_items_pl
    ON Reel_Items (Packing_List_Number);

CREATE INDEX IF NOT EXISTS idx_master_ritm
    ON Reel_Load_Database (RITM_Number);

CREATE INDEX IF NOT EXISTS idx_master_date
    ON Reel_Load_Database (Receiving_Date);

CREATE VIEW IF NOT EXISTS Reel_Load_By_Date AS
SELECT Packing_List_Number, Receiving_Date, RITM_Number, Location,
       Truck, Truck_Number,
       Total_Gross_Weight, Total_Tare, Total_Net_Weight, Number_of_Reels
FROM Reel_Load_Database
ORDER BY Receiving_Date, Packing_List_Number;

CREATE VIEW IF NOT EXISTS Reel_Items_By_Date AS
SELECT m.Receiving_Date, i.Packing_List_Number, i.RITM_Number,
       i.Item_Description, i.Item_Type, i.Gross, i.Tare, i.Net
FROM Reel_Items i
JOIN Reel_Load_Database m
  ON m.Packing_List_Number = i.Packing_List_Number
ORDER BY m.Receiving_Date, i.Packing_List_Number, i.Item_Description;
"""

_SCHEMA_LOCK = threading.Lock()
_SCHEMA_READY: dict[str, bool] = {}


def db_path() -> Path | None:
    configured = load_app_config().get("database_path")
    if not configured:
        return None
    return Path(configured)


def manual_db_path() -> Path:
    configured = load_app_config().get("manual_database_path")
    if not configured:
        return get_user_dir() / "PackingLists_Manual.db"
    return Path(configured)


def db_configured() -> bool:
    return db_path() is not None


def _ensure_schema(conn: sqlite3.Connection, path: Path) -> None:
    key = str(path.resolve()).casefold()
    if _SCHEMA_READY.get(key):
        return
    with _SCHEMA_LOCK:
        if _SCHEMA_READY.get(key):
            return
        conn.executescript(_SQLITE_SCHEMA)
        _SCHEMA_READY[key] = True


def _connect(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), timeout=10.0, isolation_level=None)
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    _ensure_schema(conn, path)
    return conn


def db_connect() -> sqlite3.Connection:
    path = db_path()
    if path is None:
        raise RuntimeError(
            "No database configured. Use the Database button."
        )
    return _connect(path)


def manual_db_connect() -> sqlite3.Connection:
    return _connect(manual_db_path())

def _connect_to_db_kind(kind: str) -> sqlite3.Connection:
    """kind is 'main' or 'manual'."""
    if kind == "manual":
        return manual_db_connect()
    return db_connect()
    
def email_db_path() -> Path:
    configured = load_app_config().get("email_database_path")
    if not configured:
        return get_user_dir() / "Email_Delivery.db"
    return Path(configured)


def email_db_connect() -> sqlite3.Connection:
    import improvedScanner
    return improvedScanner._connect(email_db_path())



def with_db_retry(fn, attempts: int = 6, base_delay: float = 0.4):
    last: Exception | None = None
    for i in range(attempts):
        try:
            return fn()
        except sqlite3.OperationalError as e:
            msg = str(e).lower()
            transient = "locked" in msg or "busy" in msg or "disk" in msg
            if not transient or i == attempts - 1:
                raise
            last = e
            _time.sleep(base_delay * (2 ** i))
    if last is not None:
        raise last


def _db_ritm_value(ritm: str | None) -> str | None:
    if _is_no_request_item(ritm):
        return None
    return str(ritm).strip()


def ritm_item_exists_in_db(ritm: str, item_descriptions: list[str]) -> bool:
    if not db_configured():
        return False
    ritm_value = _db_ritm_value(ritm)
    if ritm_value is None:
        return False
    descriptions = [d.strip() for d in item_descriptions if d and d.strip()]
    if not descriptions:
        return False

    def _query() -> bool:
        conn = db_connect()
        try:
            placeholders = ",".join(["?"] * len(descriptions))
            row = conn.execute(
                f"SELECT 1 FROM Reel_Items "
                f"WHERE RITM_Number = ? AND Item_Description IN ({placeholders}) "
                f"LIMIT 1",
                (ritm_value, *descriptions),
            ).fetchone()
            return row is not None
        finally:
            conn.close()

    try:
        return with_db_retry(_query)
    except Exception:
        _APP_LOGGER.exception("DB duplicate check failed for RITM %s.", ritm)
        return False


def record_db_duplicates(
    records: list["PackingListRecord"],
) -> list[tuple[str, str, str]]:
    if not db_configured() or not records:
        return []

    def _query() -> list[tuple[str, str, str]]:
        conn = db_connect()
        try:
            out: list[tuple[str, str, str]] = []
            seen: set[tuple[str, str, str]] = set()
            for record in records:
                ritm_value = _db_ritm_value(record.get("ritm", ""))
                if ritm_value is None:
                    continue
                current_frn = _packing_list_identifier(record)
                for item in record.get("items") or []:
                    desc = (item.get("description") or "").strip()
                    if not desc:
                        continue
                    row = conn.execute(
                        "SELECT Packing_List_Number FROM Reel_Items "
                        "WHERE RITM_Number = ? AND Item_Description = ?",
                        (ritm_value, desc),
                    ).fetchone()
                    if row is None:
                        continue
                    existing_frn = str(row[0] or "")
                    if existing_frn.casefold() == current_frn.casefold():
                        continue
                    key = (ritm_value, desc, existing_frn)
                    if key in seen:
                        continue
                    seen.add(key)
                    out.append(key)
            return out
        finally:
            conn.close()

    try:
        return with_db_retry(_query)
    except Exception:
        _APP_LOGGER.exception("DB duplicate scan failed.")
        return []


def _base_frn(packing_list_identifier: str) -> str:
    match = re.fullmatch(r"(.+)-([A-Za-z]+)", packing_list_identifier.strip())
    if match is None:
        return packing_list_identifier.strip()
    return match.group(1)


def delete_base_frn_from_db(packing_list_identifier: str) -> int:
    if not db_configured():
        return 0
    base = _base_frn(packing_list_identifier)
    if not base:
        return 0
    like = f"{base}%"

    def _delete() -> int:
        conn = db_connect()
        try:
            conn.execute("BEGIN")
            conn.execute(
                "DELETE FROM Reel_Items WHERE Packing_List_Number LIKE ?",
                (like,),
            )
            cur = conn.execute(
                "DELETE FROM Reel_Load_Database WHERE Packing_List_Number LIKE ?",
                (like,),
            )
            deleted = cur.rowcount
            conn.execute("COMMIT")
            return max(deleted, 0)
        except Exception:
            try:
                conn.execute("ROLLBACK")
            except Exception:
                pass
            raise
        finally:
            conn.close()

    try:
        return with_db_retry(_delete)
    except Exception:
        _APP_LOGGER.exception("DB delete failed for base FRN %s.", base)
        return 0


def _insert_record(conn: sqlite3.Connection, record: "PackingListRecord") -> tuple[int, int]:
    pl_number = _packing_list_identifier(record)
    ritm_value = _db_ritm_value(record.get("ritm", ""))
    items = list(record.get("items") or [])
    trucker, trucker_id = _split_trucker_and_id(record.get("driver_plate", ""))
    try:
        receiving_date = parse_date(record["date"]).isoformat()
    except Exception:
        receiving_date = None

    total_gross = sum(int(i.get("gross", 0) or 0) for i in items)
    total_tare = sum(int(i.get("tare", 0) or 0) for i in items)
    total_net = sum(int(i.get("net", 0) or 0) for i in items)

    conn.execute("BEGIN")
    try:
        conn.execute(
            "INSERT INTO Reel_Load_Database "
            "(Packing_List_Number, Receiving_Date, RITM_Number, Location, "
            " Truck, Truck_Number, Total_Gross_Weight, Total_Tare, "
            " Total_Net_Weight, Number_of_Reels) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                pl_number,
                receiving_date,
                ritm_value,
                record.get("location_address") or record.get("location") or "",
                trucker,
                trucker_id,
                total_gross,
                total_tare,
                total_net,
                len(items),
            ),
        )

        inserted = 0
        skipped = 0
        for item in items:
            desc = (item.get("description") or "").strip()
            if not desc:
                continue
            try:
                conn.execute(
                    "INSERT INTO Reel_Items "
                    "(Packing_List_Number, RITM_Number, Item_Description, "
                    " Item_Type, Gross, Tare, Net) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (
                        pl_number,
                        ritm_value,
                        desc,
                        normalize_item_type(item.get("type") or ""),
                        int(item.get("gross", 0) or 0),
                        int(item.get("tare", 0) or 0),
                        int(item.get("net", 0) or 0),
                    ),
                )
                inserted += 1
            except sqlite3.IntegrityError:
                skipped += 1
                _logging.info(f"DB skipped duplicate item: {ritm_value} + {desc}")
        conn.execute("COMMIT")
        return inserted, skipped
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        raise


def insert_record_to_db(record: "PackingListRecord") -> tuple[int, int]:
    if not db_configured():
        return (0, 0)

    def _insert() -> tuple[int, int]:
        conn = db_connect()
        try:
            return _insert_record(conn, record)
        finally:
            conn.close()

    try:
        return with_db_retry(_insert)
    except Exception:
        _APP_LOGGER.exception(
            "DB insert failed for %s.", _packing_list_identifier(record)
        )
        return (0, len(record.get("items") or []))


def insert_record_to_manual_db(record: "PackingListRecord") -> tuple[int, int]:
    def _insert() -> tuple[int, int]:
        conn = manual_db_connect()
        try:
            return _insert_record(conn, record)
        finally:
            conn.close()

    try:
        return with_db_retry(_insert)
    except Exception:
        _APP_LOGGER.exception(
            "Manual-DB insert failed for %s.",
            _packing_list_identifier(record),
        )
        return (0, len(record.get("items") or []))
def manual_db_record_exists(packing_list_identifier: str) -> bool:
    """True if a row with this exact Packing_List_Number exists in the manual DB."""
    try:
        conn = manual_db_connect()
    except Exception:
        return False
    try:
        row = conn.execute(
            "SELECT 1 FROM Reel_Load_Database WHERE Packing_List_Number = ? "
            "LIMIT 1",
            (packing_list_identifier,),
        ).fetchone()
        return row is not None
    except Exception:
        _APP_LOGGER.exception(
            "Manual-DB existence check failed for %s.", packing_list_identifier,
        )
        return False
    finally:
        try:
            conn.close()
        except Exception:
            pass


def delete_manual_db_record(packing_list_identifier: str) -> None:
    """Delete one packing list (master + items) from the manual DB."""
    try:
        conn = manual_db_connect()
    except Exception:
        return
    try:
        conn.execute("BEGIN")
        conn.execute(
            "DELETE FROM Reel_Items WHERE Packing_List_Number = ?",
            (packing_list_identifier,),
        )
        conn.execute(
            "DELETE FROM Reel_Load_Database WHERE Packing_List_Number = ?",
            (packing_list_identifier,),
        )
        conn.execute("COMMIT")
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        _APP_LOGGER.exception(
            "Manual-DB delete failed for %s.", packing_list_identifier,
        )
    finally:
        try:
            conn.close()
        except Exception:
            pass

# ---------------------------------------------------------------------------
# Week helpers
# ---------------------------------------------------------------------------
def week_range(iso_date: str) -> tuple[date, date]:
    d = parse_date(iso_date)
    weekday = d.weekday()
    if weekday >= 5:
        monday = d + timedelta(days=(7 - weekday))
    else:
        monday = d - timedelta(days=weekday)
    friday = monday + timedelta(days=4)
    return monday, friday


def summary_filename(monday: date, friday: date) -> str:
    return (
        f"Reel Summary Report WE "
        f"{monday.month}.{monday.day}.{monday.strftime('%y')} ~ "
        f"{friday.month}.{friday.day}.{friday.strftime('%y')}.xls"
    )


_FRN_DATE_RE = re.compile(r"FRN(\d{2})(\d{2})(\d{2})", re.IGNORECASE)


def _date_from_frn(packing_list_number: str) -> tuple[date | None, str]:
    if not packing_list_number:
        return None, "empty packing list number"
    match = _FRN_DATE_RE.search(packing_list_number)
    if not match:
        return None, "no FRN date found"
    yy, mm, dd = match.groups()
    try:
        return date(2000 + int(yy), int(mm), int(dd)), ""
    except ValueError as error:
        return None, str(error)


# ---------------------------------------------------------------------------
# Excel COM helpers
# ---------------------------------------------------------------------------

@_timed
def _release_excel(excel, workbook=None) -> None:
    if workbook is not None:
        try:
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
    if excel is not None:
        try:
            excel.Calculation = XL_CALC_AUTOMATIC
        except Exception:
            pass
        try:
            excel.ScreenUpdating = True
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
    del excel
    del workbook
    try:
        pythoncom.CoFreeUnusedLibraries()
    except Exception:
        pass
    gc.collect()


def _tune_excel_for_batch(excel) -> None:
    for attr, value in (
        ("ScreenUpdating", False), ("DisplayAlerts", False),
        ("EnableEvents", False), ("Calculation", XL_CALC_MANUAL),
        ("AskToUpdateLinks", False),
    ):
        try:
            setattr(excel, attr, value)
        except Exception:
            pass


def _safe_write(sheet, address: str, value) -> None:
    rng = sheet.Range(address)
    try:
        if rng.MergeCells:
            rng = rng.MergeArea.Cells(1, 1)
    except Exception:
        pass
    rng.Value = value


def _safe_clear(sheet, address: str) -> None:
    rng = sheet.Range(address)
    try:
        if rng.MergeCells:
            rng = rng.MergeArea.Cells(1, 1)
    except Exception:
        pass
    rng.ClearContents()


def _clear_item_range(sheet, first_row: int, last_row: int, first_col: int, last_col: int) -> None:
    try:
        rng = sheet.Range(
            sheet.Cells(first_row, first_col),
            sheet.Cells(last_row, last_col),
        )
        if rng.MergeCells:
            raise RuntimeError("merged block")
        rng.ClearContents()
        return
    except Exception:
        pass
    for r in range(first_row, last_row + 1):
        for c in range(first_col, last_col + 1):
            cell = sheet.Cells(r, c)
            try:
                if cell.MergeCells:
                    continue
            except Exception:
                pass
            cell.ClearContents()


def _read_block(sheet, first_row: int, first_col: int, last_row: int, last_col: int):
    return sheet.Range(
        sheet.Cells(first_row, first_col),
        sheet.Cells(last_row, last_col),
    ).Value


def _as_grid(value, rows: int, cols: int) -> list[list]:
    if value is None:
        return [[None] * cols for _ in range(rows)]
    if not isinstance(value, tuple):
        return [[value] + [None] * (cols - 1)]
    if rows == 1 and cols == 1:
        return [[value[0]]] if isinstance(value, tuple) else [[value]]
    if rows == 1:
        return [list(value)]
    if cols == 1:
        return [[v] for v in value]
    return [list(row) for row in value]


# ---------------------------------------------------------------------------
# PDF helpers
# ---------------------------------------------------------------------------

@_timed
def slice_first_pages(source_bytes: bytes, max_pages: int = 2) -> bytes | None:
    try:
        reader = PdfReader(io.BytesIO(source_bytes))
        if not reader.pages:
            return None
        writer = PdfWriter()
        for page in reader.pages[:max_pages]:
            writer.add_page(page)
        out = io.BytesIO()
        writer.write(out)
        return out.getvalue()
    except Exception:
        return None


@_timed
def annotate_pdf(source_bytes: bytes, note: str) -> bytes | None:
    try:
        reader = PdfReader(io.BytesIO(source_bytes))
        writer = PdfWriter()

        first_page = reader.pages[0]
        mb = first_page.mediabox
        page_w = float(mb.width)
        page_h = float(mb.height)

        buf = io.BytesIO()
        can = canvas.Canvas(buf, pagesize=(page_w, page_h))
        can.setFont("Helvetica-Bold", 11)
        can.setFillColorRGB(1, 0, 0)
        can.drawCentredString(page_w * 0.40, page_h - 25, note)
        can.save()
        buf.seek(0)
        overlay_page = PdfReader(buf).pages[0]

        for index, page in enumerate(reader.pages):
            if index == 0:
                page.merge_page(overlay_page)
            writer.add_page(page)

        out = io.BytesIO()
        writer.write(out)
        return out.getvalue()
    except Exception:
        return None


@_timed
def fetch_attachment_bytes(entry_id: str, store_id: str, filename: str) -> bytes | None:
    if not entry_id:
        return None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        message = None
        if store_id:
            try:
                message = namespace.GetItemFromID(entry_id, store_id)
            except Exception:
                message = None
        if message is None:
            try:
                message = namespace.GetItemFromID(entry_id)
            except Exception:
                message = None
        if message is None:
            return None

        for attachment in message.Attachments:
            if str(attachment.FileName).casefold() == filename.casefold():
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=Path(filename).suffix
                ) as tmp:
                    tmp_path = tmp.name
                try:
                    attachment.SaveAsFile(tmp_path)
                    data = Path(tmp_path).read_bytes()
                finally:
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
                return data
    except Exception:
        return None
    return None


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class NearMatch:
    term: str
    guess: str
    similarity: float
    filename: str
    folder: str
    subject: str
    received: str
    subject_ritms: list[str]
    sender: str = ""
    sender_email: str = ""
    entry_id: str = ""
    store_id: str = ""
    location: str = ""
    item_context: str = ""


@dataclass(frozen=True)
class ExactMatch:
    term: str
    filename: str
    folder: str
    subject: str
    received: str
    sender: str
    sender_email: str
    entry_id: str
    store_id: str
    ritms: list[str]
    subject_ritms: list[str]
    location: str
    location_address: str
    is_near_match: bool = False
    attachment_item_id: str = ""


@dataclass
class MatchedEmail:
    key: str
    ritm: str
    packing_list: str
    letter: str
    entry_id: str
    store_id: str
    sender: str
    sender_email: str
    subject: str
    received: str
    folder: str
    location: str = ""
    location_address: str = ""
    attachments: set[str] = field(default_factory=set)
    item_terms: set[str] = field(default_factory=set)
    attachment_item_ids: dict[str, set[str]] = field(default_factory=dict)


@dataclass
class ScanResults:
    exact_matches: list[ExactMatch]
    near_matches: list[NearMatch]


@dataclass
class ScannerGroup:
    key: str
    label: str
    terms: list[str]
    files: list[str]
    exact_files: list[str] = field(default_factory=list)
    near_matches: list[NearMatch] = field(default_factory=list)


class ItemRecord(TypedDict, total=False):
    description: str
    type: str
    gross: int
    tare: int
    net: int
    _address: str
    _ritm: str


class PackingListRecord(TypedDict, total=False):
    group: str
    ritm: str
    letter: str
    date: str
    packing_list: str
    location: str
    location_address: str
    driver_plate: str
    items: list[ItemRecord]
    source_matches: list[ExactMatch]


def _item_address(item: ItemRecord) -> str:
    return str(item.get("_address", "") or "").strip()


def _item_ritm(item: ItemRecord) -> str:
    return str(item.get("_ritm", "") or "").strip()


def _record_source_matches(record: PackingListRecord) -> list[ExactMatch]:
    return list(record.get("source_matches") or [])


def _dedupe_matches(matches: list[ExactMatch]) -> list[ExactMatch]:
    seen: set[tuple[str, str, str]] = set()
    out: list[ExactMatch] = []
    for match in matches:
        key = (match.entry_id, match.filename, match.term)
        if key in seen:
            continue
        seen.add(key)
        out.append(match)
    return out


# ---------------------------------------------------------------------------
# Assignment dialog
# ---------------------------------------------------------------------------

class AddressBatchDialog(QDialog):
    rescan_found_sources = Signal(object)

    def __init__(
        self,
        parent,
        unmatched_items: list[ItemRecord],
        staged_ritms: list[PackingListRecord],
        record_template: PackingListRecord,
        subject_matches: list[ExactMatch] | None = None,
        time_range: str = "2",
        search_terms: list[str] | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Assign location / Request Item, remove from RITMs")
        self.setMinimumWidth(820)
        self.setMinimumHeight(640)

        self._template: PackingListRecord = dict(record_template)
        self._subject_matches: list[ExactMatch] = list(subject_matches or [])
        self._time_range = time_range
        self._search_terms = list(search_terms or [])

        self.unmatched_items: list[ItemRecord] = [dict(item) for item in unmatched_items]
        self.staged_ritms: list[PackingListRecord] = []
        for r in staged_ritms:
            if not r.get("items"):
                continue
            copy_r = dict(r)
            copy_r["items"] = [dict(item) for item in r["items"]]
            copy_r.setdefault("source_matches", list(r.get("source_matches") or []))
            self.staged_ritms.append(copy_r)

        self._last_rescan_matches: list[ExactMatch] = []

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Assign location/RITM to unmatched items, or pull items out of "
            "the staged RITMs. Any RITM you assign via the Request Item or "
            "Both tab will be created if it doesn't already exist."
        ))

        self.tabs = QTabWidget()
        self.lists: list[QListWidget] = []
        self._build_tab("Location", self._build_location_tab)
        self._build_tab("Request Item", self._build_ritm_tab)
        self._build_tab("Both", self._build_both_tab)
        self._build_remove_tab()
        layout.addWidget(self.tabs)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #53645e;")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        if not self.unmatched_items:
            self.tabs.setCurrentWidget(self._remove_tab_widget)

    def _build_tab(self, title: str, builder) -> None:
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        self._populate_list(list_widget)
        self.lists.append(list_widget)
        tab_layout.addWidget(list_widget)
        builder(tab_layout, list_widget)
        self.tabs.addTab(tab, title)

    def _populate_list(self, list_widget: QListWidget) -> None:
        list_widget.clear()
        for item in self.unmatched_items:
            entry = QListWidgetItem(item["description"])
            entry.setFlags(entry.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            entry.setCheckState(Qt.CheckState.Unchecked)
            entry.setData(Qt.ItemDataRole.UserRole, id(item))
            list_widget.addItem(entry)

    def _refresh_all_assignment_lists(self) -> None:
        for list_widget in self.lists:
            self._populate_list(list_widget)

    def _build_location_tab(self, tab_layout, list_widget) -> None:
        row = QHBoxLayout()
        row.addWidget(QLabel("Location:"))
        self.location_input = QLineEdit()
        self.location_input.setPlaceholderText(
            "Street, City, State ZIP  (leave blank for No Location)"
        )
        row.addWidget(self.location_input, 1)
        tab_layout.addLayout(row)
        self._add_apply_buttons(tab_layout, "apply_location")

    def _build_ritm_tab(self, tab_layout, list_widget) -> None:
        row = QHBoxLayout()
        row.addWidget(QLabel("Request Item:"))
        self.ritm_input = QLineEdit()
        self.ritm_input.setPlaceholderText("e.g. RITM0011836256")
        row.addWidget(self.ritm_input, 1)
        tab_layout.addLayout(row)
        self._add_rescan_row(tab_layout, lambda: self.ritm_input.text())
        self._add_apply_buttons(tab_layout, "apply_ritm")

    def _build_both_tab(self, tab_layout, list_widget) -> None:
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Location:"))
        self.both_location_input = QLineEdit()
        self.both_location_input.setPlaceholderText(
            "Street, City, State ZIP  (leave blank for No Location)"
        )
        row1.addWidget(self.both_location_input, 1)
        tab_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Request Item:"))
        self.both_ritm_input = QLineEdit()
        self.both_ritm_input.setPlaceholderText("e.g. RITM0011836256")
        row2.addWidget(self.both_ritm_input, 1)
        tab_layout.addLayout(row2)
        self._add_rescan_row(tab_layout, lambda: self.both_ritm_input.text())
        self._add_apply_buttons(tab_layout, "apply_both")

    def _add_rescan_row(self, tab_layout, get_text) -> None:
        row = QHBoxLayout()
        rescan_button = QPushButton("Rescan for this RITM")
        rescan_button.setStyleSheet(
            "font-weight: bold; background: #4a6a93; color: white;"
        )
        rescan_button.clicked.connect(lambda: self._do_rescan(get_text()))
        row.addWidget(rescan_button)
        row.addStretch(1)
        tab_layout.addLayout(row)

    def _add_apply_buttons(self, tab_layout, mode: str) -> None:
        button_row = QHBoxLayout()
        apply_button = QPushButton("Apply to selected")
        apply_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        apply_button.clicked.connect(lambda: self._apply(mode))
        button_row.addWidget(apply_button)

        another_button = QPushButton("Another batch")
        another_button.clicked.connect(self._another_batch)
        button_row.addWidget(another_button)

        finished_button = QPushButton("Finished")
        finished_button.setStyleSheet("font-weight: bold;")
        finished_button.clicked.connect(self.accept)
        button_row.addWidget(finished_button)

        button_row.addStretch(1)
        tab_layout.addLayout(button_row)

    def _build_remove_tab(self) -> None:
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.addWidget(QLabel(
            "Check items to pull out of a staged RITM. They return to the "
            "unmatched pool so you can re-assign them."
        ))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(6, 6, 6, 6)
        scroll.setWidget(container)
        tab_layout.addWidget(scroll, 1)

        button_row = QHBoxLayout()
        self.apply_removals_button = QPushButton("Apply removals")
        self.apply_removals_button.setStyleSheet(
            "font-weight: bold; background: #8b2f2f; color: white;"
        )
        self.apply_removals_button.clicked.connect(self._apply_removals)
        button_row.addWidget(self.apply_removals_button)

        finished_button = QPushButton("Finished")
        finished_button.setStyleSheet("font-weight: bold;")
        finished_button.clicked.connect(self.accept)
        button_row.addWidget(finished_button)
        button_row.addStretch(1)
        tab_layout.addLayout(button_row)

        self._remove_container_layout = container_layout
        self._remove_scroll = scroll
        self._remove_tab_widget = tab
        self.tabs.addTab(tab, "Remove / Drop")
        self._rebuild_remove_tab()

    def _rebuild_remove_tab(self) -> None:
        layout = self._remove_container_layout
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.staged_ritms = [r for r in self.staged_ritms if r["items"]]

        if not self.staged_ritms:
            layout.addWidget(QLabel(
                "(No RITMs staged yet. Assign a RITM in the Request Item "
                "or Both tab to create one.)"
            ))
            layout.addStretch(1)
            self.apply_removals_button.setEnabled(False)
            self._remove_checkboxes = []
            return

        self.apply_removals_button.setEnabled(True)
        self._remove_checkboxes: list[tuple[PackingListRecord, list[QCheckBox]]] = []

        for record in self.staged_ritms:
            ritm_label = (
                "(no request item)"
                if _is_no_request_item(record.get("ritm"))
                else record["ritm"]
            )
            group = QGroupBox(
                f"RITM {ritm_label} — {len(record['items'])} item(s)"
            )
            group_layout = QVBoxLayout(group)
            boxes: list[QCheckBox] = []
            for item in record["items"]:
                box = QCheckBox(
                    f"{item['description']}  ({item['type']}, "
                    f"net {item['net']:,})"
                )
                box.setChecked(False)
                group_layout.addWidget(box)
                boxes.append(box)
            self._remove_checkboxes.append((record, boxes))
            layout.addWidget(group)

        layout.addStretch(1)

    def _checked_items(self, list_widget: QListWidget) -> list[ItemRecord]:
        wanted_ids: set[int] = set()
        for row in range(list_widget.count()):
            entry = list_widget.item(row)
            if entry.checkState() == Qt.CheckState.Checked:
                wanted_ids.add(int(entry.data(Qt.ItemDataRole.UserRole)))
        if not wanted_ids:
            return []
        return [item for item in self.unmatched_items if id(item) in wanted_ids]

    def _apply(self, mode: str) -> None:
        current_tab = self.tabs.currentIndex()
        if current_tab >= len(self.lists):
            return
        list_widget = self.lists[current_tab]
        items = self._checked_items(list_widget)

        if not items:
            QMessageBox.information(
                self, "Nothing selected",
                "Check at least one item before applying.",
            )
            return

        location = ""
        ritm = ""
        if mode == "apply_location":
            location = self.location_input.text().strip()
        elif mode == "apply_ritm":
            ritm_text = self.ritm_input.text().strip()
            ritm = normalize_manual_ritm(ritm_text)
            if not ritm_text:
                QMessageBox.information(
                    self, "No RITM entered",
                    "Type a Request Item number before applying.",
                )
                return
            if not ritm:
                QMessageBox.warning(
                    self, "Invalid RITM",
                    "Enter a Request Item in the format RITM followed by "
                    "its number, or enter the number alone.",
                )
                return
            self.ritm_input.setText(ritm)
        elif mode == "apply_both":
            location = self.both_location_input.text().strip()
            ritm_text = self.both_ritm_input.text().strip()
            ritm = normalize_manual_ritm(ritm_text)
            if ritm_text and not ritm:
                QMessageBox.warning(
                    self, "Invalid RITM",
                    "Enter a Request Item in the format RITM followed by "
                    "its number, or enter the number alone.",
                )
                return
            if not location and not ritm:
                QMessageBox.information(
                    self, "Nothing to apply",
                    "Enter a location and/or a Request Item before applying.",
                )
                return
            if ritm:
                self.both_ritm_input.setText(ritm)

        if not ritm:
            for item in items:
                item["_address"] = location
            target = None
            for r in self.staged_ritms:
                if (
                    _is_no_request_item(r.get("ritm"))
                    and r.get("location_address") == location
                ):
                    target = r
                    break
            if target is None:
                target = {
                    "group": "no request item",
                    "ritm": NO_REQUEST_ITEM,
                    "letter": "",
                    "date": self._template["date"],
                    "packing_list": self._template["packing_list"],
                    "location": self._location_from_address(location),
                    "location_address": location,
                    "driver_plate": self._template["driver_plate"],
                    "items": [],
                    "source_matches": [],
                }
                self.staged_ritms.append(target)
            target["items"].extend(items)
            moved_ids = {id(item) for item in items}
            self.unmatched_items = [
                item for item in self.unmatched_items if id(item) not in moved_ids
            ]
            self._refresh_all_assignment_lists()
            self._rebuild_remove_tab()
        elif location and ritm:
            self._move_items_to_ritm(items, ritm, location)
        elif ritm:
            self._move_items_to_ritm(items, ritm, location)

        self.location_input.clear()
        self.both_location_input.clear()

        if not self.unmatched_items:
            self.status_label.setText(
                "All items assigned. Click Finished to close."
            )
        else:
            self.status_label.setText(
                f"Applied to {len(items)} item(s). "
                f"{len(self.unmatched_items)} unmatched item(s) remain."
            )

    def _move_items_to_ritm(
        self,
        items: list[ItemRecord],
        ritm: str,
        location: str,
    ) -> None:
        moved_ids = {id(item) for item in items}
        self.unmatched_items = [
            item for item in self.unmatched_items if id(item) not in moved_ids
        ]
        self._refresh_all_assignment_lists()

        ritm_key = ritm.casefold()
        target = None
        for r in self.staged_ritms:
            if (r["ritm"] or "").casefold() == ritm_key:
                target = r
                break

        if target is None:
            sources = self._sources_for_ritm(ritm)
            matched_location_address = ""
            matched_location = ""
            for match in sources:
                if match.location_address:
                    matched_location_address = match.location_address
                    matched_location = match.location
                    break
            new_ritm: PackingListRecord = {
                "group": ritm_key,
                "ritm": ritm,
                "letter": "",
                "date": self._template["date"],
                "packing_list": self._template["packing_list"],
                "location": (
                    self._location_from_address(location) if location
                    else matched_location
                ),
                "location_address": location or matched_location_address,
                "driver_plate": self._template["driver_plate"],
                "items": [],
                "source_matches": list(sources),
            }
            self.staged_ritms.append(new_ritm)
            target = new_ritm

        for item in items:
            item["_ritm"] = ritm
            if location:
                item["_address"] = location
        target["items"].extend(items)
        if location:
            target["location_address"] = location
            target["location"] = self._location_from_address(location)

        self._rebuild_remove_tab()

    def _sources_for_ritm(self, ritm: str) -> list[ExactMatch]:
        ritm_key = ritm.casefold()
        matches = [
            match for match in self._subject_matches
            if _match_mentions_ritm(match, ritm_key)
        ]
        matches.extend(
            match for match in self._last_rescan_matches
            if _match_mentions_ritm(match, ritm_key)
        )
        return _dedupe_matches(matches)

    def _another_batch(self) -> None:
        self.location_input.clear()
        self.ritm_input.clear()
        self.both_location_input.clear()
        self.both_ritm_input.clear()
        self.status_label.setText(
            "New batch started. Check items and enter the next values."
        )

    def _do_rescan(self, ritm: str) -> None:
        ritm_text = (ritm or "").strip()
        ritm = normalize_manual_ritm(ritm_text)
        if not ritm_text:
            QMessageBox.information(
                self, "No RITM entered",
                "Type a Request Item number in the field above before "
                "rescanning.",
            )
            return
        if not ritm:
            QMessageBox.warning(
                self, "Invalid RITM",
                "Enter a Request Item in the format RITM followed by "
                "its number, or enter the number alone.",
            )
            return

        self.status_label.setText(f"Looking up {ritm} in the Email Database...")
        QApplication.processEvents()

        try:
            import improvedScanner
            rows = improvedScanner.fetch_delivery_emails(
                db_path=email_db_path(), ritm_filter=ritm,
            )
        except Exception as error:
            _APP_LOGGER.exception("Email DB lookup failed for %s.", ritm)
            QMessageBox.critical(self, "Lookup failed", str(error))
            self.status_label.setText("Lookup failed.")
            return

        if not rows:
            self.status_label.setText(
                f"{ritm} not in Email DB — syncing Outlook..."
            )
            QApplication.processEvents()
            try:
                improvedScanner.sync_delivery_emails(
                    db_path=email_db_path(), force_full=False,
                )
                rows = improvedScanner.fetch_delivery_emails(
                    db_path=email_db_path(), ritm_filter=ritm,
                )
            except Exception as error:
                _APP_LOGGER.exception("Email DB sync failed for %s.", ritm)
                QMessageBox.critical(self, "Sync failed", str(error))
                self.status_label.setText("Sync failed.")
                return

        if not rows:
            self.status_label.setText(
                f"No email in the Email DB mentions {ritm}. "
                "The RITM will be exported without a source PDF."
            )
            return

        hits = _dedupe_matches([
            ExactMatch(
                term=row["term"],
                filename=row["filename"],
                folder=row["folder"],
                subject=row["subject"],
                received=row["received"],
                sender=row["sender"],
                sender_email=row["sender_email"],
                entry_id=row["entry_id"],
                store_id=row["store_id"],
                ritms=row["ritms"],
                subject_ritms=row["subject_ritms"],
                location=row["location"],
                location_address=row["location_address"],
            )
            for row in rows
        ])

        self._last_rescan_matches.extend(hits)
        ritm_key = ritm.casefold()
        attached = 0
        for r in self.staged_ritms:
            if (r["ritm"] or "").casefold() != ritm_key:
                continue
            current = list(r.get("source_matches") or [])
            seen = {(m.entry_id, m.filename, m.term) for m in current}
            for match in hits:
                key = (match.entry_id, match.filename, match.term)
                if key in seen:
                    continue
                seen.add(key)
                current.append(match)
                attached += 1
            r["source_matches"] = current
            if not r["location_address"]:
                for match in hits:
                    if match.location_address:
                        r["location_address"] = match.location_address
                        r["location"] = match.location
                        break

        self.rescan_found_sources.emit(
            (ritm, list(hits), self._template["packing_list"])
        )
        self.status_label.setText(
            f"Found {len(hits)} email(s) in Email DB for {ritm}"
            + (f"; {attached} new source(s) attached." if attached else ".")
        )

        ritm_text = (ritm or "").strip()
        ritm = normalize_manual_ritm(ritm_text)
        if not ritm_text:
            QMessageBox.information(
                self, "No RITM entered",
                "Type a Request Item number in the field above before "
                "rescanning.",
            )
            return
        if not ritm:
            QMessageBox.warning(
                self, "Invalid RITM",
                "Enter a Request Item in the format RITM followed by "
                "its number, or enter the number alone.",
            )
            return

        term_set: list[str] = []
        seen_terms: set[str] = set()
        for item in self.unmatched_items:
            desc = item["description"].strip()
            if desc and desc.casefold() not in seen_terms:
                seen_terms.add(desc.casefold())
                term_set.append(desc)
        for r in self.staged_ritms:
            if (r["ritm"] or "").casefold() != ritm.casefold():
                continue
            for item in r["items"]:
                desc = item["description"].strip()
                if desc and desc.casefold() not in seen_terms:
                    seen_terms.add(desc.casefold())
                    term_set.append(desc)

        if ritm.casefold() not in seen_terms:
            seen_terms.add(ritm.casefold())
            term_set.append(ritm)

        if not term_set:
            QMessageBox.information(
                self, "Nothing to search",
                "Add items or a RITM with items before rescanning.",
            )
            return

        self.status_label.setText(
            f"Rescanning Outlook for {len(term_set)} item description(s) "
            f"matching {ritm}..."
        )
        QApplication.processEvents()

        try:
            import improvedScanner
            with tempfile.TemporaryDirectory(prefix="rescan-") as temp_dir:
                report_path = Path(temp_dir) / "rescan.txt"
                results_path = Path(temp_dir) / "rescan.json"
                improvedScanner.run_scan(
                    self._time_range,
                    term_set,
                    str(report_path),
                    str(results_path),
                )
                results = json.loads(results_path.read_text(encoding="utf-8"))
        except Exception as error:
            _APP_LOGGER.exception("Manual Outlook rescan failed for %s.", ritm)
            QMessageBox.critical(self, "Rescan failed", str(error))
            self.status_label.setText("Rescan failed.")
            return

        new_matches: list[ExactMatch] = [
            ExactMatch(
                term=entry["term"],
                filename=entry["filename"],
                folder=entry["folder"],
                subject=entry["subject"],
                received=entry["received"],
                sender=entry.get("sender", ""),
                sender_email=entry.get("sender_email", ""),
                entry_id=entry.get("entry_id", ""),
                store_id=entry.get("store_id", ""),
                ritms=entry.get("ritms", []),
                subject_ritms=entry.get("subject_ritms", []),
                location=entry.get("location", ""),
                location_address=entry.get("location_address", ""),
            )
            for entry in results["exact_matches"]
        ]

        ritm_key = ritm.casefold()
        hits = _dedupe_matches([
            match for match in new_matches
            if _match_mentions_ritm(match, ritm)
        ])

        if not hits:
            self.status_label.setText(
                f"Rescan found no emails mentioning {ritm}. "
                "The RITM will be exported without a source PDF."
            )
            return

        self._last_rescan_matches.extend(hits)
        attached = 0
        for r in self.staged_ritms:
            if (r["ritm"] or "").casefold() != ritm_key:
                continue
            current = list(r.get("source_matches") or [])
            seen = {(m.entry_id, m.filename, m.term) for m in current}
            for match in hits:
                key = (match.entry_id, match.filename, match.term)
                if key in seen:
                    continue
                seen.add(key)
                current.append(match)
                attached += 1
            r["source_matches"] = current
            if not r["location_address"]:
                for match in hits:
                    if match.location_address:
                        r["location_address"] = match.location_address
                        r["location"] = match.location
                        break

        self.rescan_found_sources.emit(
            (ritm, list(hits), self._template["packing_list"])
        )
        self.status_label.setText(
            f"Rescan found {len(hits)} source email(s) for {ritm}"
            + (f"; {attached} new attachment(s) attached." if attached else ".")
        )

    def _apply_removals(self) -> None:
        if not getattr(self, "_remove_checkboxes", None):
            return

        any_checked = False
        newly_freed: list[ItemRecord] = []

        for record, boxes in self._remove_checkboxes:
            drop_indexes = [i for i, box in enumerate(boxes) if box.isChecked()]
            if not drop_indexes:
                continue
            any_checked = True
            drop_set = set(drop_indexes)
            kept = [item for i, item in enumerate(record["items"]) if i not in drop_set]
            freed = [item for i, item in enumerate(record["items"]) if i in drop_set]
            record["items"] = kept
            for item in freed:
                item["_ritm"] = ""
            newly_freed.extend(freed)

        if not any_checked:
            QMessageBox.information(
                self, "Nothing selected",
                "Check at least one item to remove.",
            )
            return

        self.unmatched_items.extend(newly_freed)
        self._refresh_all_assignment_lists()
        self._rebuild_remove_tab()
        self.status_label.setText(
            f"Pulled {len(newly_freed)} item(s) back into the unmatched pool. "
            f"{len(self.unmatched_items)} unmatched item(s) available."
        )
        self.tabs.setCurrentIndex(0)

    @staticmethod
    def _location_from_address(address: str) -> str:
        parts = [p.strip() for p in address.split(",")]
        if len(parts) < 2:
            return address.strip()
        city_state = ", ".join(parts[1:])
        city_state = re.sub(r"\s+\d{5}(?:-\d{4})?\s*$", "", city_state).strip()
        return city_state.replace(",", "")


# ---------------------------------------------------------------------------
# Near-match / resend / source picker / item-ritm picker dialogs
# ---------------------------------------------------------------------------

class NearMatchConfirmDialog(QDialog):
    def __init__(self, parent, candidate: NearMatch, index: int, total: int) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Possible item match ({index} of {total})")
        self.setMinimumWidth(560)
        self.accepted: bool = False
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Your item description:"))
        mine = QLabel(candidate.term)
        mine.setStyleSheet("font-weight: bold; font-size: 12pt; padding: 4px;")
        layout.addWidget(mine)
        layout.addWidget(QLabel("Close match found in email attachment:"))
        theirs = QLabel(candidate.guess)
        theirs.setStyleSheet("font-weight: bold; font-size: 12pt; padding: 4px;")
        layout.addWidget(theirs)
        detail = QLabel(
            f"Similarity: {candidate.similarity:.0%}\n"
            f"RITM(s): {', '.join(candidate.subject_ritms) or '(none)'}\n"
            f"From: {candidate.sender} <{candidate.sender_email}>\n"
            f"Location: {candidate.location or '(not in email database)'}\n"
            f"Matched item details:\n{candidate.item_context or '(none)'}\n"
            f"File: {candidate.folder} / {candidate.filename}\n"
            f"Email: {candidate.subject}\n"
            f"Received: {candidate.received}"
        )
        detail.setStyleSheet("color: #53645e; padding-top: 6px;")
        detail.setWordWrap(True)
        layout.addWidget(detail)
        layout.addWidget(QLabel("\nAre these the same item?"))
        buttons = QDialogButtonBox()
        same_button = buttons.addButton("Same", QDialogButtonBox.ButtonRole.AcceptRole)
        not_same_button = buttons.addButton("Not the same", QDialogButtonBox.ButtonRole.RejectRole)
        same_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white; padding: 6px 16px;"
        )
        not_same_button.setStyleSheet("padding: 6px 16px;")
        buttons.accepted.connect(self._on_same)
        buttons.rejected.connect(self._on_not_same)
        layout.addWidget(buttons)

    def _on_same(self) -> None:
        self.accepted = True
        self.accept()

    def _on_not_same(self) -> None:
        self.accepted = False
        self.reject()


class ResendPickerDialog(QDialog):
    def __init__(self, parent, ritm: str, older: ExactMatch, newer: ExactMatch) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Duplicate RITM detected: {ritm}")
        self.setMinimumWidth(1000)
        self.chosen: ExactMatch | None = None
        self.older = older
        self.newer = newer
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"RITM {ritm} was found in two emails. "
            "The newer email is preselected."
        ))
        columns = QHBoxLayout()

        def make_column(title: str, match: ExactMatch, is_newer: bool) -> QWidget:
            box = QGroupBox(title)
            box_layout = QVBoxLayout(box)

            def add_field(label: str, value: str) -> None:
                row = QHBoxLayout()
                tag = QLabel(f"{label}:")
                tag.setStyleSheet("font-weight: bold;")
                tag.setFixedWidth(80)
                row.addWidget(tag)
                body = QLabel(value or "(none)")
                body.setWordWrap(True)
                row.addWidget(body, 1)
                box_layout.addLayout(row)

            add_field("Received", match.received)
            add_field("From", f"{match.sender} <{match.sender_email}>")
            add_field("Subject", match.subject)
            add_field("Folder", match.folder)
            add_field("File", match.filename)
            add_field("Terms", ", ".join(sorted({match.term})))
            if is_newer:
                box.setStyleSheet(
                    "QGroupBox { border: 2px solid #315b52; border-radius: 4px; "
                    "margin-top: 8px; padding-top: 10px; }"
                    "QGroupBox::title { subcontrol-origin: margin; "
                    "subcontrol-position: top left; padding: 0 6px; "
                    "color: #315b52; font-weight: bold; }"
                )
            return box

        self.radio_older = QRadioButton("Use OLDER")
        self.radio_newer = QRadioButton("Use NEWER")
        self.radio_newer.setChecked(True)

        older_col = QVBoxLayout()
        older_col.addWidget(self.radio_older)
        older_col.addWidget(make_column("OLDER email", older, is_newer=False))
        columns.addLayout(older_col, 1)

        newer_col = QVBoxLayout()
        newer_col.addWidget(self.radio_newer)
        newer_col.addWidget(make_column("NEWER email", newer, is_newer=True))
        columns.addLayout(newer_col, 1)

        layout.addLayout(columns)
        buttons = QDialogButtonBox()
        use_btn = buttons.addButton("Use selected", QDialogButtonBox.ButtonRole.AcceptRole)
        skip_btn = buttons.addButton("Skip this RITM", QDialogButtonBox.ButtonRole.RejectRole)
        use_btn.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white; padding: 6px 16px;"
        )
        buttons.accepted.connect(self._on_use)
        buttons.rejected.connect(self._on_skip)
        layout.addWidget(buttons)

    def _on_use(self) -> None:
        self.chosen = self.newer if self.radio_newer.isChecked() else self.older
        self.accept()

    def _on_skip(self) -> None:
        self.chosen = None
        self.accept()


class SourceMatchPickerDialog(QDialog):
    def __init__(self, parent, identifier: str, matches: list[ExactMatch]) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Choose source email for {identifier}")
        self.setMinimumWidth(600)
        self.chosen: ExactMatch | None = None
        self.matches: list[ExactMatch] = list(matches)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"This load matched multiple emails. Which one should the "
            f"'Request Item for Load {identifier}' PDF come from?"
        ))
        self.radios: list[QRadioButton] = []
        for index, match in enumerate(self.matches):
            label = (
                f"{match.subject}\n"
                f"   From:   {match.sender} <{match.sender_email}>\n"
                f"   Folder: {match.folder}\n"
                f"   File:   {match.filename}\n"
                f"   Recv:   {match.received}"
            )
            radio = QRadioButton(label)
            if index == 0:
                radio.setChecked(True)
            layout.addWidget(radio)
            self.radios.append(radio)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_ok(self) -> None:
        for index, radio in enumerate(self.radios):
            if radio.isChecked():
                self.chosen = self.matches[index]
                break
        self.accept()


class ItemRitmPickerDialog(QDialog):
    def __init__(self, parent, description: str, ritms: list[str]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Choose Request Item")
        self.setMinimumWidth(420)
        self.chosen: str | None = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f'"{description}" matched multiple Request Items. Choose one:'
        ))
        self.radios: list[QRadioButton] = []
        for index, ritm in enumerate(ritms):
            radio = QRadioButton(ritm)
            if index == 0:
                radio.setChecked(True)
            layout.addWidget(radio)
            self.radios.append(radio)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_ok(self) -> None:
        for index, radio in enumerate(self.radios):
            if radio.isChecked():
                self.chosen = radio.text()
                break
        self.accept()


class EmailSearchResultsDialog(QDialog):
    """Popup showing email search results with count and details."""
    def __init__(self, parent, results: list[dict]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Email Search Results")
        self.setMinimumSize(900, 550)
        self.setModal(True)

        self._results = results  # Store for reel popup access

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)

        # Header with count
        count = len(results)
        header = QLabel(f"Found {count} matching email{'s' if count != 1 else ''}")
        header.setStyleSheet("font-size: 14pt; font-weight: bold; color: #20332e;")
        layout.addWidget(header)

        if count == 0:
            empty = QLabel("No emails matched your criteria.")
            empty.setStyleSheet("color: #53645e; padding: 20px;")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(empty)
        else:
            # Table of results - 6 columns now (added Reels)
            table = QTableWidget(count, 6)
            table.setHorizontalHeaderLabels([
                "RITM", "Sender", "Subject", "Received", "Folder", "Reels"
            ])
            table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
            table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            table.setAlternatingRowColors(True)
            table.verticalHeader().setVisible(False)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)  # Subject wider
            table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)  # Reels button column

            for row, r in enumerate(results):
                sender = r["sender"]
                if r["sender_email"] and r["sender_email"].casefold() not in sender.casefold():
                    sender = f"{sender} <{r['sender_email']}>".strip()

                # Columns 0-4: standard data
                static_values = (
                    r["ritm"] or "—",
                    sender,
                    r["subject"] or "—",
                    r["received"] or "—",
                    r["folder"] or "—",
                )
                for col, val in enumerate(static_values):
                    item = QTableWidgetItem(str(val))
                    item.setToolTip(str(val))
                    table.setItem(row, col, item)

                # Column 5: Reels button
                reel_items = r.get("items", [])  # from fetch_delivery_emails
                reel_btn = QPushButton("Reels")
                reel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                reel_btn.setStyleSheet("""
                    QPushButton {
                        background: #315b52; color: white; border: none;
                        border-radius: 3px; padding: 4px 12px; font-weight: bold;
                    }
                    QPushButton:hover { background: #25473e; }
                    QPushButton:disabled { background: #9baea6; color: #e2e9e5; }
                """)
                if reel_items:
                    reel_btn.clicked.connect(
                        lambda checked=False, items=reel_items, ritm=r["ritm"]: self._show_reel_popup(items, ritm)
                    )
                else:
                    reel_btn.setEnabled(False)
                    reel_btn.setToolTip("No reel data available")
                table.setCellWidget(row, 5, reel_btn)

            layout.addWidget(table, 1)

        # Close button
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("font-weight: bold; padding: 8px 24px;")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _show_reel_popup(self, items: list[dict], ritm: str) -> None:
        """Show a dialog listing all reel IDs with details."""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Reels for {ritm or 'Unknown RITM'}")
        dialog.setMinimumSize(700, 400)
        dialog.setModal(True)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)

        # Header
        header = QLabel(f"{len(items)} reel(s) for {ritm}")
        header.setStyleSheet("font-size: 12pt; font-weight: bold; color: #20332e;")
        layout.addWidget(header)

        # Table with reel details - 3 columns only
        table = QTableWidget(len(items), 3)
        table.setHorizontalHeaderLabels([
            "Serial / Item ID", "Type", "Description"
        ])
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        table.setStyleSheet("""
            QTableWidget { 
                font-family: 'Consolas', 'Monospace'; 
                font-size: 10pt; 
                gridline-color: #d7dfda;
            }
            QTableWidget::item { padding: 6px; }
            QTableWidget::item:selected { background: #d5e1da; color: #20332e; }
            QHeaderView::section { background: #e2e9e5; padding: 6px; border: none; font-weight: bold; }
        """)

        for row, item in enumerate(items):
            serial = item.get("item_id") or item.get("description") or "—"
            asset_class = item.get("asset_class") or item.get("equipment_type") or "—"
            desc = item.get("description") or ""
            
            values = (serial, asset_class, desc)
            for col, val in enumerate(values):
                table_item = QTableWidgetItem(str(val))
                table_item.setToolTip(str(val))
                table.setItem(row, col, table_item)

        layout.addWidget(table, 1)

        # Buttons
        btn_row = QHBoxLayout()
        copy_btn = QPushButton("Copy Selected Serials")
        copy_btn.setToolTip("Copy selected reel serial numbers to clipboard")
        copy_btn.clicked.connect(lambda: self._copy_table_selection(table))
        btn_row.addWidget(copy_btn)
        
        copy_all_btn = QPushButton("Copy All Serials")
        copy_all_btn.setToolTip("Copy all reel serial numbers to clipboard")
        copy_all_btn.clicked.connect(lambda: self._copy_all_serials(items))
        btn_row.addWidget(copy_all_btn)
        
        btn_row.addStretch(1)
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("font-weight: bold; padding: 8px 24px;")
        close_btn.clicked.connect(dialog.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        # Enable Ctrl+C on table
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
        table.installEventFilter(self)
        dialog._reel_table = table
        
        dialog.exec()

    def _copy_table_selection(self, table: QTableWidget) -> None:
        """Copy selected serials (column 0) to clipboard."""
        selected = table.selectedItems()
        if not selected:
            return
        rows = set(item.row() for item in selected)
        serials = []
        for row in sorted(rows):
            item = table.item(row, 0)
            if item:
                serials.append(item.text())
        if serials:
            QApplication.clipboard().setText("\n".join(serials))

    def _copy_all_serials(self, items: list[dict]) -> None:
        """Copy all reel serials to clipboard."""
        serials = [item.get("item_id") or item.get("description") or "" for item in items]
        serials = [s for s in serials if s]
        if serials:
            QApplication.clipboard().setText("\n".join(serials))


# ---------------------------------------------------------------------------
# Incoming Packing List popup dialog
# ---------------------------------------------------------------------------

class IncomingPackingListDialog(QDialog):
    """The popup form that mirrors the Excel Incoming Packing List sheet.

    Gross / Tare are plain QLineEdit fields (digit-only) so they match the
    visual style of Description / Material / Net — no spin arrows.
    """

    def __init__(
        self,
        parent=None,
        existing_items: list[ItemRecord] | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Incoming Packing List")
        self.setMinimumWidth(760)
        self.setMinimumHeight(720)

        self.current_items: list[ItemRecord] = [
            dict(item) for item in (existing_items or [])
        ]
        self.result_record: PackingListRecord | None = None
        # Editing state: when not None, the top row is "held" for editing
        # and points at the index in self.current_items it came from.
        self._editing_index: int | None = None

        self.setStyleSheet(
            "QDialog { background: #f2f5f3; color: #20332e; }"
            "QLabel { font-size: 10pt; }"
            "QLineEdit { background: white; border: 1px solid #9baea6;"
            " border-radius: 3px; padding: 6px; font-size: 10pt; }"
            "QPushButton { background: #e2e9e5; border: 1px solid #9baea6;"
            " border-radius: 3px; padding: 7px 12px; font-size: 10pt; }"
            "QPushButton:hover { background: #d5e1da; }"
            "QTableWidget { background: white; alternate-background-color: #f5f8f6;"
            " gridline-color: #d7dfda; font-size: 10pt; }"
            "QHeaderView::section { background: #e2e9e5; padding: 6px;"
            " border: 0; font-weight: bold; }"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)

        title = QLabel("Incoming Packing List")
        title.setStyleSheet("font-size: 15pt; font-weight: bold;")
        layout.addWidget(title)

        # ---- Header grid ----------------------------------------------
        header_box = QGroupBox("Header")
        header_box.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 10px; padding-top: 14px; "
            "border: 1px solid #b7c7bf; border-radius: 4px; }"
            "QGroupBox::title { subcontrol-origin: margin; "
            "subcontrol-position: top left; padding: 0 6px; color: #315b52; }"
        )
        header_layout = QVBoxLayout(header_box)
        header_layout.setContentsMargins(12, 18, 12, 12)

        row1 = QHBoxLayout()

        date_col = QVBoxLayout()
        date_col.addWidget(QLabel("Date"))
        self.date_input = QLineEdit(date.today().strftime("%m/%d/%Y"))
        self.date_input.setFixedWidth(140)
        self.date_input.setPlaceholderText("MM/DD/YYYY")
        date_col.addWidget(self.date_input)
        row1.addLayout(date_col)

        trucker_col = QVBoxLayout()
        trucker_col.addWidget(QLabel("Trucker"))
        self.trucker_input = QLineEdit()
        self.trucker_input.setPlaceholderText("e.g. CEVA")
        trucker_col.addWidget(self.trucker_input)
        row1.addLayout(trucker_col, 1)

        header_layout.addLayout(row1)

        row2 = QHBoxLayout()

        received_col = QVBoxLayout()
        received_col.addWidget(QLabel("Received From"))
        self.received_from_input = QLineEdit(DEFAULT_RECEIVED_FROM)
        received_col.addWidget(self.received_from_input)
        row2.addLayout(received_col)

        trucker_id_col = QVBoxLayout()
        trucker_id_col.addWidget(QLabel("Trucker ID"))
        self.trucker_id_input = QLineEdit()
        self.trucker_id_input.setPlaceholderText("e.g. 205396 or ABC123")
        trucker_id_col.addWidget(self.trucker_id_input)
        row2.addLayout(trucker_id_col, 1)

        header_layout.addLayout(row2)

        row3 = QHBoxLayout()
        packing_col = QVBoxLayout()
        packing_col.addWidget(QLabel("Packing List #"))
        self.packing_list_input = QLineEdit()
        self.packing_list_input.setPlaceholderText("e.g. FRN261112-1R")
        packing_col.addWidget(self.packing_list_input)
        row3.addLayout(packing_col, 1)
        header_layout.addLayout(row3)

        layout.addWidget(header_box)

        # ---- Item entry row -------------------------------------------
        item_box = QGroupBox("Add Item")
        item_box.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 10px; padding-top: 14px; "
            "border: 1px solid #b7c7bf; border-radius: 4px; }"
            "QGroupBox::title { subcontrol-origin: margin; "
            "subcontrol-position: top left; padding: 0 6px; color: #315b52; }"
        )
        item_layout = QVBoxLayout(item_box)
        item_layout.setContentsMargins(12, 18, 12, 12)

        item_form = QHBoxLayout()

        desc_col = QVBoxLayout()
        desc_col.addWidget(QLabel("ITEM DESCRIPTION"))
        self.description_input = QLineEdit()
        desc_col.addWidget(self.description_input)
        item_form.addLayout(desc_col, 3)

        type_col = QVBoxLayout()
        type_col.addWidget(QLabel("Material"))
        self.type_input = QLineEdit()
        self.type_input.setPlaceholderText("e.g. 99 FO, 23, Type 3, 69S")
        type_col.addWidget(self.type_input)
        item_form.addLayout(type_col, 2)

        # Gross — plain QLineEdit, digits only, no spin arrows.
        gross_col = QVBoxLayout()
        gross_col.addWidget(QLabel("Gross"))
        self.gross_input = QLineEdit("0")
        self.gross_input.setValidator(QIntValidator(0, 1_000_000_000, self))
        self.gross_input.setAlignment(Qt.AlignmentFlag.AlignRight)
        gross_col.addWidget(self.gross_input)
        item_form.addLayout(gross_col, 1)

        # Tare — plain QLineEdit, digits only, no spin arrows.
        tare_col = QVBoxLayout()
        tare_col.addWidget(QLabel("Tare"))
        self.tare_input = QLineEdit("0")
        self.tare_input.setValidator(QIntValidator(0, 1_000_000_000, self))
        self.tare_input.setAlignment(Qt.AlignmentFlag.AlignRight)
        tare_col.addWidget(self.tare_input)
        item_form.addLayout(tare_col, 1)

        net_col = QVBoxLayout()
        net_col.addWidget(QLabel("Net"))
        self.net_preview = QLineEdit("0")
        self.net_preview.setReadOnly(True)
        net_col.addWidget(self.net_preview)
        item_form.addLayout(net_col, 1)

        item_layout.addLayout(item_form)

        self.item_input_fields = [
            self.description_input, self.type_input,
            self.gross_input, self.tare_input,
        ]
        for field in self.item_input_fields:
            field.installEventFilter(self)

        self.gross_input.textChanged.connect(self._update_totals)
        self.tare_input.textChanged.connect(self._update_totals)
        self.description_input.returnPressed.connect(self._focus_next_item_field)
        self.type_input.returnPressed.connect(self._focus_next_item_field)
        self.gross_input.returnPressed.connect(self._focus_next_item_field)

        item_buttons = QHBoxLayout()
        self.add_item_button = QPushButton("Add Item")
        self.add_item_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.add_item_button.clicked.connect(self._add_item)
        item_buttons.addWidget(self.add_item_button)

        self.cancel_edit_button = QPushButton("Cancel edit")
        self.cancel_edit_button.setVisible(False)
        self.cancel_edit_button.clicked.connect(self._cancel_edit)
        item_buttons.addWidget(self.cancel_edit_button)

        clear_row_btn = QPushButton("Clear Boxes")
        clear_row_btn.clicked.connect(self._clear_row)
        item_buttons.addWidget(clear_row_btn)

        item_buttons.addStretch(1)
        item_layout.addLayout(item_buttons)

        layout.addWidget(item_box)

        # ---- Items table ----------------------------------------------
        table_box = QGroupBox("Items so far")
        table_box.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 10px; padding-top: 14px; "
            "border: 1px solid #b7c7bf; border-radius: 4px; }"
            "QGroupBox::title { subcontrol-origin: margin; "
            "subcontrol-position: top left; padding: 0 6px; color: #315b52; }"
        )
        table_layout = QVBoxLayout(table_box)
        table_layout.setContentsMargins(12, 18, 12, 12)

        self.items_table = QTableWidget(0, 6)
        self.items_table.setHorizontalHeaderLabels(
            ["ITEM DESCRIPTION", "Material", "Gross", "Tare", "Net", ""]
        )
        self.items_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.items_table.setSelectionMode(
            QTableWidget.SelectionMode.ExtendedSelection
        )
        self.items_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.items_table.setAlternatingRowColors(True)
        self.items_table.verticalHeader().setVisible(False)
        self.items_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.items_table.horizontalHeader().setSectionResizeMode(
            5, QHeaderView.ResizeMode.ResizeToContents
        )
        self.items_table.setMinimumHeight(180)
        table_layout.addWidget(self.items_table)

        delete_btn_row = QHBoxLayout()
        delete_btn = QPushButton("Delete selected")
        delete_btn.clicked.connect(self._remove_selected_item)
        delete_btn_row.addWidget(delete_btn)
        delete_btn_row.addStretch(1)
        table_layout.addLayout(delete_btn_row)

        layout.addWidget(table_box, 1)

        # ---- Totals ---------------------------------------------------
        totals_row = QHBoxLayout()
        self.gross_total_label = QLabel("Gross total: 0")
        self.tare_total_label = QLabel("Tare total: 0")
        self.net_total_label = QLabel("Net total: 0")
        for lbl in (self.gross_total_label, self.tare_total_label, self.net_total_label):
            lbl.setStyleSheet("font-weight: bold; color: #20332e;")
            totals_row.addWidget(lbl)
        totals_row.addStretch(1)
        layout.addLayout(totals_row)

        # ---- Footer buttons -------------------------------------------
        footer = QHBoxLayout()
        footer.addStretch(1)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        footer.addWidget(cancel_btn)

        ok_btn = QPushButton("OK")
        ok_btn.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        ok_btn.clicked.connect(self._accept_form)
        footer.addWidget(ok_btn)

        layout.addLayout(footer)

        for item in self.current_items:
            self._append_item_row(item)
        self._update_totals()
        self.description_input.setFocus()

    # ---- item row helpers -------------------------------------------

    def _append_item_row(self, item: ItemRecord) -> None:
        row = self.items_table.rowCount()
        self.items_table.insertRow(row)
        values = (
            item["description"],
            item["type"],
            f"{item['gross']:,}",
            f"{item['tare']:,}",
            f"{item['net']:,}",
        )
        for col, value in enumerate(values):
            self.items_table.setItem(row, col, QTableWidgetItem(value))

        edit_button = QPushButton("Edit")
        edit_button.setFlat(True)
        edit_button.setCursor(Qt.CursorShape.PointingHandCursor)
        edit_button.setStyleSheet(
            "QPushButton { background: transparent; border: none; "
            "color: #315b52; font-weight: bold; padding: 2px 6px; }"
            "QPushButton:hover { text-decoration: underline; }"
        )
        edit_button.clicked.connect(
            lambda _checked=False, b=edit_button: self._edit_item_at_button(b)
        )
        self.items_table.setCellWidget(row, 5, edit_button)



    def _row_of_button(self, button: QPushButton) -> int | None:
        """Find the table row that owns the given Edit button."""
        for row in range(self.items_table.rowCount()):
            if self.items_table.cellWidget(row, 5) is button:
                return row
        return None

    def _edit_item_at_button(self, button: QPushButton) -> None:
        row = self._row_of_button(button)
        if row is None or row >= len(self.current_items):
            return

        # If a different item is already being edited, restore it first.
        if self._editing_index is not None:
            self._cancel_edit()

        item = self.current_items[row]
        self._editing_index = row

        # Pop the item back into the top input row.
        self.description_input.setText(item["description"])
        self.type_input.setText(item["type"])
        self.gross_input.setText(str(item["gross"]))
        self.tare_input.setText(str(item["tare"]))
        self._update_totals()

        # Remove it from the table so it's not duplicated while editing.
        self.items_table.removeRow(row)

        # Flip the buttons into "Save changes" mode.
        self.add_item_button.setText("Save Changes")
        self.cancel_edit_button.setVisible(True)
        self.description_input.setFocus()

    def _cancel_edit(self) -> None:
        if self._editing_index is None:
            return
        index = self._editing_index
        self._editing_index = None

        # Re-insert the item into the table at its original index.
        if 0 <= index <= len(self.current_items):
            self.items_table.insertRow(index)
            item = self.current_items[index]
            values = (
                item["description"],
                item["type"],
                f"{item['gross']:,}",
                f"{item['tare']:,}",
                f"{item['net']:,}",
            )
            for col, value in enumerate(values):
                self.items_table.setItem(index, col, QTableWidgetItem(value))
            edit_button = QPushButton("Edit")
            edit_button.setFlat(True)
            edit_button.setCursor(Qt.CursorShape.PointingHandCursor)
            edit_button.setStyleSheet(
                "QPushButton { background: transparent; border: none; "
                "color: #315b52; font-weight: bold; padding: 2px 6px; }"
                "QPushButton:hover { text-decoration: underline; }"
            )
            edit_button.clicked.connect(
                lambda _checked=False, b=edit_button: self._edit_item_at_button(b)
            )
            self.items_table.setCellWidget(index, 5, edit_button)

        self.add_item_button.setText("Add Item")
        self.cancel_edit_button.setVisible(False)
        self._clear_row()

    def _rebuild_items_table(self) -> None:
        self.items_table.setRowCount(0)
        for item in self.current_items:
            self._append_item_row(item)

    @staticmethod
    def _read_int(field: QLineEdit) -> int:
        text = field.text().strip().replace(",", "")
        if not text:
            return 0
        try:
            return int(text)
        except ValueError:
            return 0

    def _clear_row(self) -> None:
        self.description_input.clear()
        self.type_input.clear()
        self.gross_input.setText("0")
        self.tare_input.setText("0")
        self.description_input.setFocus()
        self._update_totals()

    def _add_item(self) -> None:
        description = self.description_input.text().strip()
        item_type = normalize_item_type(self.type_input.text())
        gross = self._read_int(self.gross_input)
        tare = self._read_int(self.tare_input)

        # Editing + empty input row = user wants to delete this item.
        if (
            self._editing_index is not None
            and not description
            and not item_type
            and gross == 0
            and tare == 0
        ):
            self._delete_held_item()
            return

        if not description or not item_type:
            QMessageBox.warning(
                self, "Missing item details",
                "Enter an item description and material.",
            )
            return
        if gross <= 0:
            QMessageBox.warning(
                self, "Check gross weight",
                "Gross weight must be greater than zero.",
            )
            return
        if tare > gross:
            QMessageBox.warning(
                self, "Check tare weight",
                "Tare cannot be greater than gross.",
            )
            return

        new_item: ItemRecord = {
            "description": description,
            "type": item_type,
            "gross": gross,
            "tare": tare,
            "net": gross - tare,
        }

        if self._editing_index is not None:
            index = self._editing_index
            self._editing_index = None
            self.current_items[index] = new_item
            self._rebuild_items_table()
            self.add_item_button.setText("Add Item")
            self.cancel_edit_button.setVisible(False)
            self._clear_row()
            return

        self.current_items.append(new_item)
        self._append_item_row(new_item)
        self._clear_row()


    def _delete_held_item(self) -> None:
        """Remove the item currently held for editing, reset to Add mode."""
        if self._editing_index is None:
            return
        index = self._editing_index
        self._editing_index = None
        if 0 <= index < len(self.current_items):
            self.current_items.pop(index)
        self._rebuild_items_table()
        self.add_item_button.setText("Add Item")
        self.cancel_edit_button.setVisible(False)
        self._clear_row()
        self._update_totals()

    def _remove_selected_item(self) -> None:
        if self._editing_index is not None:
            self._cancel_edit()
        indexes = sorted(
            {i.row() for i in self.items_table.selectionModel().selectedRows()},
            reverse=True,
        )
        if not indexes:
            return
        for index in indexes:
            self.current_items.pop(index)
        self._rebuild_items_table()
        self._update_totals()

    def _update_totals(self) -> None:
        entered_gross = self._read_int(self.gross_input)
        entered_tare = self._read_int(self.tare_input)
        entered_net = entered_gross - entered_tare
        self.net_preview.setText(f"{entered_net:,}")

        gross_total = sum(item["gross"] for item in self.current_items) + entered_gross
        tare_total = sum(item["tare"] for item in self.current_items) + entered_tare
        net_total = sum(item["net"] for item in self.current_items) + entered_net
        self.gross_total_label.setText(f"Gross total: {gross_total:,}")
        self.tare_total_label.setText(f"Tare total: {tare_total:,}")
        self.net_total_label.setText(f"Net total: {net_total:,}")

    # ---- keyboard nav -----------------------------------------------

    def _focused_item_field_index(self) -> int | None:
        focused = QApplication.focusWidget()
        for index, field in enumerate(self.item_input_fields):
            if focused is field:
                return index
        return None

    def _focus_item_field(self, index: int) -> None:
        self.item_input_fields[index % len(self.item_input_fields)].setFocus()

    def _focus_next_item_field(self) -> None:
        index = self._focused_item_field_index()
        self._focus_item_field(0 if index is None else index + 1)

    def eventFilter(self, watched: QWidget, event: QEvent) -> bool:
        if not isinstance(event, QKeyEvent) or event.type() != QEvent.Type.KeyPress:
            return super().eventFilter(watched, event)
        if event.modifiers() != Qt.KeyboardModifier.NoModifier:
            return super().eventFilter(watched, event)
        direction = 0
        if event.key() == Qt.Key.Key_Right:
            direction = 1
        elif event.key() == Qt.Key.Key_Left:
            direction = -1
        if not direction:
            return super().eventFilter(watched, event)
        index = self._focused_item_field_index()
        if index is None:
            return super().eventFilter(watched, event)
        if isinstance(watched, QLineEdit):
            if watched.hasSelectedText():
                return super().eventFilter(watched, event)
            cursor = watched.cursorPosition()
            if direction > 0 and cursor < len(watched.text()):
                return super().eventFilter(watched, event)
            if direction < 0 and cursor > 0:
                return super().eventFilter(watched, event)
        self._focus_item_field(index + direction)
        return True

    # ---- record building / DB write --------------------------------

    def _build_record(self) -> PackingListRecord | None:
        raw_date = self.date_input.text().strip()
        packing_list = self.packing_list_input.text().strip()
        if not raw_date or not packing_list:
            QMessageBox.warning(
                self, "Missing information",
                "Fill in the Date and Packing List # before continuing.",
            )
            return None
        try:
            normalized_date = parse_date(raw_date).isoformat()
        except ValueError as error:
            QMessageBox.warning(self, "Check the date", str(error))
            self.date_input.setFocus()
            return None
        if not self.current_items:
            QMessageBox.warning(
                self, "No items",
                "Add at least one item before continuing.",
            )
            return None

        trucker = self.trucker_input.text().strip()
        trucker_id = self.trucker_id_input.text().strip()

        parts: list[str] = []
        if trucker:
            parts.append(trucker)
        if trucker_id:
            parts.append(trucker_id)
        driver_plate = " / ".join(parts)

        return {
            "group": "",
            "ritm": "",
            "letter": "",
            "date": normalized_date,
            "packing_list": packing_list,
            "location": "",
            "location_address": "",
            "driver_plate": driver_plate,
            "items": [dict(item) for item in self.current_items],
        }



    def _accept_form(self) -> None:
        if self._editing_index is not None:
            description = self.description_input.text().strip()
            item_type = self.type_input.text().strip()
            gross = self._read_int(self.gross_input)
            tare = self._read_int(self.tare_input)

            row_is_empty = (
                not description
                and not item_type
                and gross == 0
                and tare == 0
            )

            if row_is_empty:
                # Cleared row + OK while editing = delete the item.
                self._delete_held_item()
            else:
                # Row still has data — require explicit Save or Cancel so
                # the user doesn't accidentally lose their edits.
                answer = QMessageBox.question(
                    self,
                    "Finish editing?",
                    "You have an item pulled into the input row.\n\n"
                    "Save it before continuing?",
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No
                    | QMessageBox.StandardButton.Cancel,
                    QMessageBox.StandardButton.Yes,
                )
                if answer == QMessageBox.StandardButton.Cancel:
                    return
                if answer == QMessageBox.StandardButton.Yes:
                    self._add_item()
                    if self._editing_index is not None:
                        # _add_item rejected the data (validation failed).
                        return
                else:
                    # "No" = discard the edit, restore the original row.
                    self._cancel_edit()

        record = self._build_record()
        if record is None:
            return
        self.result_record = record
        self.accept()


# ---------------------------------------------------------------------------
# Excel writers
# ---------------------------------------------------------------------------

@_timed
def write_packing_list_template(
    record: PackingListRecord,
    destination: Path,
    excel=None,
) -> Path:
    template = get_resource_path("Excel Bases/packing list example.xls")
    if not template.exists():
        raise FileNotFoundError(f"Packing-list template not found: {template}")
    if len(record["items"]) > MAX_ITEMS_PER_LOAD:
        raise ValueError(
            "The packing-list template has 23 item rows; split this RITM "
            "into another letter."
        )
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)

    owns_excel = excel is None
    excel_local = excel
    workbook = None
    try:
        if owns_excel:
            excel_local = win32com.client.DispatchEx("Excel.Application")
            _tune_excel_for_batch(excel_local)
        workbook = excel_local.Workbooks.Open(
            str(destination), UpdateLinks=0, ReadOnly=False
        )
        sheet = workbook.Worksheets("Packing List")
        _safe_write(
            sheet, "H8",
            datetime.combine(parse_date(record["date"]), datetime.min.time()),
        )
        sheet.Range("H8").NumberFormat = "mm/dd/yyyy"
        _safe_write(sheet, "H9", record["location"].strip() or "No Location")
        _safe_write(sheet, "H10", _packing_list_identifier(record))

        num_rows = MAX_ITEMS_PER_LOAD
        grid = [["" for _ in range(9)] for _ in range(num_rows)]
        for i, item in enumerate(record["items"]):
            if i >= num_rows:
                break
            description = item["description"]
            item_type = normalize_item_type(item["type"] or "")
            if item_type:
                description = f"{description} ({item_type})"
            grid[i][0] = description
            grid[i][6] = item["gross"]
            grid[i][7] = item["tare"]
            grid[i][8] = item["net"]

        block = sheet.Range(
            sheet.Cells(13, 2),
            sheet.Cells(13 + num_rows - 1, 10),
        )
        try:
            if block.MergeCells:
                raise RuntimeError("merged block")
            block.Value = grid
        except Exception:
            for i, item in enumerate(record["items"]):
                row = 13 + i
                description = item["description"]
                item_type = normalize_item_type(item["type"] or "")
                if item_type:
                    description = f"{description} ({item_type})"
                _safe_write(sheet, f"B{row}", description)
                _safe_write(sheet, f"H{row}", item["gross"])
                _safe_write(sheet, f"I{row}", item["tare"])
                _safe_write(sheet, f"J{row}", item["net"])
        workbook.Save()
    finally:
        if owns_excel:
            _release_excel(excel_local, workbook)
        else:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
    return destination


@_timed
def write_scrap_pickup_template(
    record: PackingListRecord,
    destination: Path,
    excel=None,
) -> Path:
    template = get_resource_path("Excel Bases/Scrap Pick Up Request for Load.xls")
    if not template.exists():
        raise FileNotFoundError(f"Scrap pickup template not found: {template}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)

    owns_excel = excel is None
    excel_local = excel
    workbook = None
    try:
        if owns_excel:
            excel_local = win32com.client.DispatchEx("Excel.Application")
            _tune_excel_for_batch(excel_local)
        workbook = excel_local.Workbooks.Open(
            str(destination), UpdateLinks=0, ReadOnly=False
        )
        sheet = workbook.Worksheets("SCRAPREM")
        request_date = datetime.combine(parse_date(record["date"]), datetime.min.time())
        for address in ("D15", "D33"):
            _safe_write(sheet, address, request_date)
            sheet.Range(address).NumberFormat = "mm/dd/yyyy"

        location_parts = [
            part.strip() for part in record["location_address"].split(",")
        ]
        if record["location_address"].strip() and location_parts[0]:
            street = location_parts[0]
            city = ", ".join(location_parts[1:]).strip()
            if not city:
                city = record["location"].strip()
        else:
            street = "No Location"
            city = ""

        _safe_write(sheet, "F16", _normalize_optional_ritm(record.get("ritm")))
        _safe_write(sheet, "D19", street)
        _safe_write(sheet, "D20", city)
        _safe_write(sheet, "D34", _packing_list_identifier(record))
        _safe_write(sheet, "D35", record["driver_plate"].replace(" / ", "# "))
        workbook.Save()
    finally:
        if owns_excel:
            _release_excel(excel_local, workbook)
        else:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
    return destination


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _packing_list_identifier(record: PackingListRecord) -> str:
    letter = (record.get("letter") or "").strip()
    if letter:
        return f"{record['packing_list']}-{letter}"
    return record["packing_list"]


_PLATE_TOKEN_RE = re.compile(r"^[A-Z0-9-]{2,10}$", re.IGNORECASE)
_NAME_TOKEN_RE = re.compile(r"[A-Za-z]{2,}", re.IGNORECASE)


def _looks_like_plate(token: str) -> bool:
    token = token.strip()
    if not token or " " in token:
        return False
    if not _PLATE_TOKEN_RE.fullmatch(token):
        return False
    if token.isalpha() and not any(ch.isdigit() for ch in token):
        return False
    return True


def _looks_like_name(token: str) -> bool:
    token = token.strip()
    if not token:
        return False
    if " " in token:
        return True
    return bool(_NAME_TOKEN_RE.fullmatch(token)) and len(token) <= 20


def _split_trucker_and_id(value: str) -> tuple[str, str]:
    return _split_driver_and_plate(value)


def _split_driver_and_plate(value: str) -> tuple[str, str]:
    raw = value.strip()
    if not raw:
        return "", ""
    parts: list[str]
    if "/" in raw:
        parts = [p.strip() for p in raw.split("/", 1)]
    elif " " in raw and not any(seg.isdigit() for seg in raw.split()):
        head, _, tail = raw.rpartition(" ")
        if _looks_like_plate(tail):
            return head.strip(), tail.strip()
        return raw, ""
    elif " " in raw:
        head, _, tail = raw.rpartition(" ")
        parts = [head.strip(), tail.strip()]
    else:
        return raw, ""

    left, right = parts[0], parts[1] if len(parts) > 1 else ""
    left_plate = _looks_like_plate(left)
    right_plate = _looks_like_plate(right)
    left_name = _looks_like_name(left)
    right_name = _looks_like_name(right)
    if left_plate and not right_plate:
        return right, left
    if right_plate and not left_plate:
        return left, right
    if left_name and right_plate:
        return left, right
    if right_name and left_plate:
        return right, left
    return left, right


def _split_reel_id(description: str) -> tuple[str, str]:
    cleaned = description.strip()
    match = re.fullmatch(r"(.*\d)([A-Za-z]+)", cleaned)
    if match is None:
        return cleaned, ""
    return match.group(1), match.group(2).upper()


def _reel_summary_sort_key(record: PackingListRecord) -> tuple[int, int, str, str]:
    frn_date, _reason = _date_from_frn(record["packing_list"])
    if frn_date is None:
        try:
            frn_date = parse_date(record["date"])
        except Exception:
            frn_date = date(2000, 1, 1)
    receiving_date = frn_date.toordinal()
    reel_number = re.search(r"-(\d+)R(?:-|$)", record["packing_list"], re.IGNORECASE)
    reel_order = int(reel_number.group(1)) if reel_number else 1_000_000
    letter = (record.get("letter") or "").strip().casefold()
    return (
        receiving_date, reel_order, letter,
        _packing_list_identifier(record).casefold(),
    )


def _read_reel_summary_record(sheet) -> PackingListRecord | None:
    identifier = str(sheet.Range("M2").Value or "").strip()
    if not identifier:
        return None
    identifier_parts = re.fullmatch(r"(.+)-([A-Za-z]+)", identifier)
    if identifier_parts is None:
        packing_list, letter = identifier, ""
    else:
        packing_list, letter = identifier_parts.groups()
    received = sheet.Range("M4").Value
    if isinstance(received, datetime):
        normalized_date = received.date().isoformat()
    elif isinstance(received, date):
        normalized_date = received.isoformat()
    else:
        normalized_date = parse_date(str(received)).isoformat()
    rows_count = 31
    cols_count = 5
    raw = _read_block(sheet, 12, 4, 42, 8)
    grid = _as_grid(raw, rows_count, cols_count)
    items: list[ItemRecord] = []
    for row_values in grid:
        reel_id = str(row_values[0] or "").strip()
        if not reel_id:
            continue
        reel_size = str(row_values[1] or "").strip()
        gross = int(float(row_values[3] or 0))
        tare = int(float(row_values[4] or 0))
        items.append({
            "description": f"{reel_id}{reel_size}",
            "type": str(row_values[2] or "").strip(),
            "gross": gross,
            "tare": tare,
            "net": gross - tare,
        })
    driver = str(sheet.Range("F8").Value or "").strip()
    plate = str(sheet.Range("F9").Value or "").strip()
    return {
        "group": "",
        "ritm": _normalize_optional_ritm(str(sheet.Range("M5").Value or "")),
        "letter": letter,
        "date": normalized_date,
        "packing_list": packing_list,
        "location": str(sheet.Range("C7").Value or "").strip(),
        "location_address": "",
        "driver_plate": f"{driver} / {plate}" if (driver or plate) else "",
        "items": items,
    }


def _reel_sheet_names_in_book(workbook) -> set[str]:
    return {sheet.Name.casefold() for sheet in workbook.Worksheets}


def _is_reel_sheet_name(name: str) -> bool:
    return re.fullmatch(r"\d+R(?:\s*\(\d+\))?", name, re.IGNORECASE) is not None


def _next_free_r_number(existing_names: set[str]) -> int:
    n = 1
    while f"{n}r" in existing_names:
        n += 1
    return n


def _copy_template_reel_tabs(workbook, template_workbook, count: int) -> list[Any]:
    source_sheet = None
    for sheet in template_workbook.Worksheets:
        if _is_reel_sheet_name(sheet.Name):
            source_sheet = sheet
            break
    if source_sheet is None:
        raise ValueError(
            "The Reel Summary template has no numbered R tabs to copy from."
        )
    created: list[Any] = []
    existing_names = _reel_sheet_names_in_book(workbook)
    after_sheet = workbook.Worksheets(workbook.Worksheets.Count)
    for _ in range(count):
        n = _next_free_r_number(existing_names)
        source_sheet.Copy(After=after_sheet)
        new_sheet = workbook.Worksheets(workbook.Worksheets.Count)
        new_sheet.Name = f"__NEW_R_{n}"
        existing_names.add(f"__new_r_{n}".casefold())
        after_sheet = new_sheet
        created.append(new_sheet)
        for address in ("M2", "M4", "M5", "C7", "I7", "C8", "F8", "C9", "F9"):
            _safe_write(new_sheet, address, "")
        _clear_item_range(new_sheet, 12, 42, 4, 8)
    return created


@_timed
def write_reel_summary_report(
    records: list[PackingListRecord],
    destination: Path,
    excel=None,
) -> Path:
    template = get_resource_path("Excel Bases/Reel Summary Report.XLS")
    if not template.exists():
        raise FileNotFoundError(f"Reel Summary template not found: {template}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    source_path = destination if destination.exists() else template
    descriptor, staged_name = tempfile.mkstemp(
        prefix=".reel-summary-", suffix=destination.suffix or ".xls", dir=destination.parent
    )
    os.close(descriptor)
    staged_path = Path(staged_name)

    owns_excel = excel is None
    excel_local = excel
    workbook = None
    template_workbook = None
    try:
        shutil.copy2(source_path, staged_path)
        if owns_excel:
            excel_local = win32com.client.DispatchEx("Excel.Application")
            _tune_excel_for_batch(excel_local)
        try:
            workbook = excel_local.Workbooks.Open(
                str(staged_path), UpdateLinks=0, ReadOnly=False
            )
            report_records: dict[str, PackingListRecord] = {}
            reel_sheets: list[Any] = []
            for sheet in workbook.Worksheets:
                is_numbered_reel_sheet = _is_reel_sheet_name(sheet.Name)
                identifier = str(sheet.Range("M2").Value or "").strip()
                is_populated_reel_sheet = bool(identifier) and sheet.Name.casefold().endswith("r")
                if not is_numbered_reel_sheet and not is_populated_reel_sheet:
                    continue
                reel_sheets.append(sheet)
                previous_record = _read_reel_summary_record(sheet)
                if previous_record is not None:
                    key = _packing_list_identifier(previous_record).casefold()
                    report_records[key] = previous_record
            if not reel_sheets:
                raise ValueError("The selected workbook does not contain any numbered R tabs.")
            for record in records:
                identifier = _packing_list_identifier(record)
                if not identifier:
                    raise ValueError("A queued packing list is missing its packing-list number.")
                report_records[identifier.casefold()] = record
            ordered_records = sorted(report_records.values(), key=_reel_summary_sort_key)
            if any(len(record["items"]) > MAX_ITEMS_PER_SUMMARY_TAB for record in ordered_records):
                record = next(
                    record for record in ordered_records
                    if len(record["items"]) > MAX_ITEMS_PER_SUMMARY_TAB
                )
                raise ValueError(
                    f"{_packing_list_identifier(record)} has more reels than a "
                    "report tab supports."
                )
            if len(ordered_records) > len(reel_sheets):
                needed = len(ordered_records) - len(reel_sheets)
                template_workbook = excel_local.Workbooks.Open(
                    str(template), UpdateLinks=0, ReadOnly=True
                )
                new_sheets = _copy_template_reel_tabs(workbook, template_workbook, needed)
                reel_sheets.extend(new_sheets)
            populated_tab_names = [
                f"{_packing_list_identifier(record)}R" for record in ordered_records
            ]
            if any(len(name) > 31 for name in populated_tab_names):
                raise ValueError("A packing-list number is too long to use as an Excel tab name.")
            reel_sheet_names = {sheet.Name.casefold() for sheet in reel_sheets}
            non_reel_sheet_names = {
                sheet.Name.casefold()
                for sheet in workbook.Worksheets
                if sheet.Name.casefold() not in reel_sheet_names
            }
            if any(name.casefold() in non_reel_sheet_names for name in populated_tab_names):
                raise ValueError("A packing-list tab name conflicts with another workbook tab.")
            existing_names = {sheet.Name.casefold() for sheet in workbook.Worksheets}
            for index, sheet in enumerate(reel_sheets, start=1):
                temporary_name = f"__R_TEMP_{index}"
                while temporary_name.casefold() in existing_names:
                    temporary_name += "_"
                sheet.Name = temporary_name
                existing_names.add(temporary_name.casefold())
            for sheet in reel_sheets:
                for address in ("M2", "M4", "M5", "C7", "I7", "C8", "F8", "C9", "F9"):
                    _safe_write(sheet, address, "")
                _clear_item_range(sheet, 12, 42, 4, 8)
            for sheet, record in zip(reel_sheets, ordered_records):
                driver, plate = _split_driver_and_plate(record["driver_plate"])
                _safe_write(sheet, "M2", _packing_list_identifier(record))
                _safe_write(
                    sheet, "M4",
                    datetime.combine(parse_date(record["date"]), datetime.min.time()),
                )
                sheet.Range("M4").NumberFormat = "mm/dd/yyyy"
                _safe_write(sheet, "M5", _normalize_optional_ritm(record.get("ritm")))
                _safe_write(sheet, "C7", record["location"].strip() or "No Location")
                _safe_write(sheet, "I7", sum(item["gross"] for item in record["items"]))
                _safe_write(sheet, "C8", "N/A")
                _safe_write(sheet, "F8", driver)
                _safe_write(sheet, "C9", "N/A")
                _safe_write(sheet, "F9", plate)
                block_rows = 31
                block_cols = 5
                grid = [["" for _ in range(block_cols)] for _ in range(block_rows)]
                for i, item in enumerate(record["items"]):
                    if i >= block_rows:
                        break
                    reel_id, reel_size = _split_reel_id(item["description"])
                    grid[i][0] = reel_id
                    grid[i][1] = reel_size
                    grid[i][2] = normalize_item_type(item["type"] or "")
                    grid[i][3] = item["gross"]
                    grid[i][4] = item["tare"]
                block = sheet.Range(
                    sheet.Cells(12, 4),
                    sheet.Cells(12 + block_rows - 1, 8),
                )
                try:
                    if block.MergeCells:
                        raise RuntimeError("merged block")
                    block.Value = grid
                except Exception:
                    for i, item in enumerate(record["items"]):
                        row = 12 + i
                        reel_id, reel_size = _split_reel_id(item["description"])
                        _safe_write(sheet, f"D{row}", reel_id)
                        _safe_write(sheet, f"E{row}", reel_size)
                        _safe_write(sheet, f"F{row}", normalize_item_type(item["type"] or ""))
                        _safe_write(sheet, f"G{row}", item["gross"])
                        _safe_write(sheet, f"H{row}", item["tare"])
            for index, sheet in enumerate(reel_sheets, start=1):
                if index <= len(ordered_records):
                    sheet.Name = populated_tab_names[index - 1]
                else:
                    n = index
                    candidate = f"{n}R"
                    while candidate.casefold() in {
                        s.Name.casefold() for s in workbook.Worksheets if s is not sheet
                    }:
                        n += 1
                        candidate = f"{n}R"
                    sheet.Name = candidate
            excel_local.CalculateFull()
            workbook.Save()
        finally:
            if template_workbook is not None:
                try:
                    template_workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if owns_excel:
                _release_excel(excel_local, None)
        os.replace(staged_path, destination)
    finally:
        if staged_path.exists():
            try:
                staged_path.unlink()
            except Exception:
                pass
    return destination


# ---------------------------------------------------------------------------
# Scanner worker
# ---------------------------------------------------------------------------

class ScannerWorker(QThread):
    """Read-only worker: queries the Email DB and emits ExactMatch rows.
    Never touches Outlook — run EmailSyncWorker first if the DB is stale."""
    scan_complete = Signal(object)
    scan_failed = Signal(str)

    def __init__(self, time_range: str = "", search_terms: list[str] | None = None) -> None:
        super().__init__()
        self.time_range = time_range
        self.search_terms = list(search_terms or [])

    def run(self) -> None:
        try:
            import improvedScanner
            rows = improvedScanner.fetch_delivery_emails(db_path=email_db_path())
            normalized_terms = {
                re.sub(r"[^a-z0-9]", "", term.casefold()): term
                for term in self.search_terms
                if term.strip()
            }
            range_days = {"1": 90, "2": 180, "3": 365}.get(self.time_range)
            cutoff = datetime.now() - timedelta(days=range_days) if range_days else None
            exact_matches: list[ExactMatch] = []
            near_by_key: dict[tuple[str, str, str, str], NearMatch] = {}
            for row in rows:
                if cutoff is not None:
                    try:
                        received_dt = datetime.fromisoformat(row["received"])
                        if received_dt.replace(tzinfo=None) < cutoff:
                            continue
                    except (TypeError, ValueError):
                        continue

                email_items = row.get("items") or []
                item_ids = [
                    (item, str(item.get("item_id") or item.get("description") or ""))
                    for item in email_items
                ]
                for term_key, term in normalized_terms.items():
                    if not term_key:
                        continue
                    exact_items = [
                        (item, item_id)
                        for item, item_id in item_ids
                        if re.sub(r"[^a-z0-9]", "", item_id.casefold()) == term_key
                    ]
                    if exact_items:
                        for _item, _item_id in exact_items:
                            exact_matches.append(ExactMatch(
                                term=term,
                                filename=row["filename"],
                                folder=row["folder"],
                                subject=row["subject"],
                                received=row["received"],
                                sender=row["sender"],
                                sender_email=row["sender_email"],
                                entry_id=row["entry_id"],
                                store_id=row["store_id"],
                                ritms=row["ritms"],
                                subject_ritms=row["subject_ritms"],
                                location=row["location"],
                                location_address=row["location_address"],
                                attachment_item_id=_item_id,
                            ))
                        continue

                    if len(term_key) < 6:
                        continue
                    best = None
                    for item, item_id in item_ids:
                        guess_key = re.sub(r"[^a-z0-9]", "", item_id.casefold())
                        if not guess_key:
                            continue
                        similarity = SequenceMatcher(None, term_key, guess_key).ratio()
                        if similarity >= 0.85 and (
                            best is None or similarity > best[0]
                        ):
                            best = (similarity, item, item_id)
                    if best is None:
                        continue

                    similarity, item, item_id = best
                    item_details = [f"Item ID: {item_id}"]
                    item_type = item.get("type") or item.get("asset_class") or ""
                    if item_type:
                        item_details.append(f"Material / type: {item_type}")
                    description = (
                        item.get("item_description") or item.get("description") or ""
                    )
                    if description and description != item_id:
                        item_details.append(f"Description: {description}")
                    item_details.append(
                        "Quantity: "
                        f"{item.get('quantity') or 1}; Gross: {item.get('gross') or 0}; "
                        f"Tare: {item.get('tare') or 0}; Net: {item.get('net') or 0}"
                    )
                    candidate = NearMatch(
                        term=term,
                        guess=item_id,
                        similarity=similarity,
                        filename=row["filename"],
                        folder=row["folder"],
                        subject=row["subject"],
                        received=row["received"],
                        subject_ritms=row["subject_ritms"],
                        sender=row["sender"],
                        sender_email=row["sender_email"],
                        entry_id=row["entry_id"],
                        store_id=row["store_id"],
                        location=row["location"],
                        item_context="\n".join(item_details),
                    )
                    source_key = row["entry_id"] or (
                        f"{row['subject']}|{row['received']}"
                    )
                    key = (
                        term_key, row["folder"].casefold(),
                        row["filename"].casefold(), source_key,
                    )
                    previous = near_by_key.get(key)
                    if previous is None or candidate.similarity > previous.similarity:
                        near_by_key[key] = candidate

            exact_matches.sort(key=lambda m: (
                m.received, m.folder.casefold(), m.subject.casefold(),
                m.filename.casefold(), m.term.casefold(),
            ))
            near_matches = sorted(near_by_key.values(), key=lambda m: (
                m.received, m.folder.casefold(), m.subject.casefold(),
                m.filename.casefold(), m.term.casefold(), -m.similarity,
            ))
            self.scan_complete.emit(ScanResults(exact_matches, near_matches))
        except Exception as error:
            _APP_LOGGER.exception("Email DB scan worker failed.")
            self.scan_failed.emit(str(error))



class EmailSyncWorker(QThread):
    """Populates the Email DB from Outlook. Separate from ScannerWorker so
    the user can sync without touching the current scan/load state."""
    sync_complete = Signal(dict)
    sync_failed = Signal(str)

    def __init__(self, force_full: bool = False) -> None:
        super().__init__()
        self.force_full = force_full

    def run(self) -> None:
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            import improvedScanner
            summary = improvedScanner.sync_delivery_emails(
                db_path=email_db_path(),
                force_full=self.force_full,
                progress_cb=lambda msg: print(msg),
            )
            self.sync_complete.emit(summary)
        except Exception as error:
            _APP_LOGGER.exception("Email DB sync worker failed.")
            self.sync_failed.emit(str(error))
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()


class EmailSearchWorker(QThread):
    results_ready = Signal(list)
    failed = Signal(str)

    def __init__(self, criteria: dict[str, str]) -> None:
        super().__init__()
        self.criteria = dict(criteria)

    def run(self) -> None:
        try:
            conn = email_db_connect()
            try:
                clauses: list[str] = []
                params: list[str] = []
                ritm = self.criteria.get("ritm", "").strip()
                address = self.criteria.get("address", "").strip()
                reel_id = self.criteria.get("reel_id", "").strip()
                if ritm:
                    clauses.append(
                        "INSTR(UPPER(REPLACE(COALESCE(r.Ritm, ''), ' ', '')), ?) > 0"
                    )
                    params.append(ritm.replace(" ", "").upper())
                if address:
                    clauses.append("INSTR(LOWER(COALESCE(d.Address, '')), ?) > 0")
                    params.append(address.casefold())
                if reel_id:
                    clauses.append(
                        "EXISTS (SELECT 1 FROM Email_Item i "
                        "WHERE i.Email_Ritm_ID = r.Email_Ritm_ID "
                        "AND INSTR(LOWER(i.Item_ID), ?) > 0)"
                    )
                    params.append(reel_id.casefold())
                if not clauses:
                    self.results_ready.emit([])
                    return

                sql = """
                    SELECT d.Email_ID, d.Entry_ID, d.Store_ID, d.Folder,
                           d.Subject, d.Sender, d.Sender_Email,
                           d.Received_Date, d.Address, r.Ritm,
                           r.Attachment_Filename,
                           (SELECT GROUP_CONCAT(DISTINCT i.Item_ID)
                              FROM Email_Item i
                             WHERE i.Email_Ritm_ID = r.Email_Ritm_ID)
                    FROM Email_Delivery d
                    LEFT JOIN Email_Ritm r ON r.Email_ID = d.Email_ID
                    WHERE """ + " AND ".join(clauses) + """
                    ORDER BY d.Received_Date DESC, d.Email_ID DESC
                    LIMIT 2000
                """
                rows = conn.execute(sql, params).fetchall()
            finally:
                conn.close()

            results: list[dict[str, str | list[str]]] = []
            for row in rows:
                (
                    email_id, entry_id, store_id, folder, subject, sender,
                    sender_email, received, address_value, ritm_value,
                    ritm_attachment, item_ids,
                ) = row
                results.append({
                    "email_id": str(email_id),
                    "entry_id": entry_id or "",
                    "store_id": store_id or "",
                    "folder": folder or "",
                    "subject": subject or "",
                    "sender": sender or "",
                    "sender_email": sender_email or "",
                    "received": received or "",
                    "address": address_value or "",
                    "ritm": ritm_value or "",
                    "reel_id": item_ids or "",
                    "attachments": [str(ritm_attachment)] if ritm_attachment else [],
                })
            self.results_ready.emit(results)
        except Exception as error:
            _APP_LOGGER.exception("Email database search failed.")
            self.failed.emit(str(error))


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class PackingListApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"Scanner Packing Lists  (v{VERSION})")
        self.resize(1180, 900)
        self.setMinimumSize(920, 720)
        # Overlay shared config so DB paths from Setup win over local config.
        self.app_config = load_app_config()
        try:
            _boot_path = (
                Path.home() / "AppData" / "Roaming" / "PackingListApp"
                / "bootstrap.json"
            )
            if _boot_path.exists():
                _boot = json.loads(_boot_path.read_text(encoding="utf-8"))
                _shared_folder = _boot.get("shared_folder", "")
                if _shared_folder:
                    _shared_cfg_path = Path(_shared_folder) / "config.json"
                    if _shared_cfg_path.exists():
                        _shared_cfg = json.loads(
                            _shared_cfg_path.read_text(encoding="utf-8")
                        )
                        self.app_config = {
                            **load_app_config(),
                            **_shared_cfg,
                        }
        except Exception:
            _APP_LOGGER.exception("Shared config overlay failed.")

        self.summaries_folder = self._resolve_summaries_folder()

        self.records: list[PackingListRecord] = []
        self.exported_record_ids_by_directory: dict[str, set[str]] = {}
        self.groups: dict[str, ScannerGroup] = {}
        self.matched_emails_by_key: dict[str, MatchedEmail] = {}
        self.email_search_results_by_key: dict[str, dict] = {}
        self.visible_email_matches: list[MatchedEmail] = []
        self.sync_worker: EmailSyncWorker | None = None
        self.scan_worker: ScannerWorker | None = None
        self.email_search_worker: EmailSearchWorker | None = None
        self.pending_record: PackingListRecord | None = None
        self.record_source_matches: dict[str, list[ExactMatch]] = {}
        # Remembers the last form the user saved via Open Input Data so the
        # popup can be pre-filled next time. Wiped by Next Load.
        self.last_form: dict | None = None
        self.setStyleSheet(
            "QMainWindow, QWidget { background: #f2f5f3; color: #20332e; }"
            "QLabel { font-size: 10pt; }"
            "QLineEdit { background: white; border: 1px solid #9baea6;"
            " border-radius: 3px; padding: 8px; font-size: 10pt; }"
            "QPushButton { background: #e2e9e5; border: 1px solid #9baea6;"
            " border-radius: 3px; padding: 8px 12px; font-size: 10pt; }"
            "QPushButton:hover { background: #d5e1da; }"
            "QTableWidget { background: white; alternate-background-color: #f5f8f6;"
            " gridline-color: #d7dfda; font-size: 10pt; }"
            "QHeaderView::section { background: #e2e9e5; padding: 7px;"
            " border: 0; font-weight: bold; }"
        )
        self._build_interface()

    def _resolve_summaries_folder(self) -> Path:
        configured = self.app_config.get("summaries_folder")
        if configured:
            p = Path(configured)
            if p.is_dir():
                return p
        DEFAULT_SUMMARIES_DIR.mkdir(parents=True, exist_ok=True)
        return DEFAULT_SUMMARIES_DIR
    def _start_email_sync(self) -> None:
        if self.sync_worker is not None and self.sync_worker.isRunning():
            QMessageBox.information(
                self, "Sync already running",
                "An email sync is already in progress.",
            )
            return
        if hasattr(self, "sync_button"):
            self.sync_button.setEnabled(False)
            self.sync_button.setText("Syncing…")
        self.sync_worker = EmailSyncWorker(force_full=False)
        self.sync_worker.sync_complete.connect(self._on_sync_complete)
        self.sync_worker.sync_failed.connect(self._on_sync_failed)
        self.sync_worker.start()

    def _on_sync_complete(self, summary: dict) -> None:
        if hasattr(self, "sync_button"):
            self.sync_button.setEnabled(True)
            self.sync_button.setText("Sync Email DB")
        QMessageBox.information(
            self, "Email DB synced",
            f"Inserted: {summary.get('inserted', 0)}\n"
            f"Skipped: {summary.get('skipped', 0)}\n"
            f"Total in DB: {summary.get('total', 0)}\n"
            f"Latest received: {summary.get('watermark', '—')}",
        )

    def _on_sync_failed(self, error: str) -> None:
        if hasattr(self, "sync_button"):
            self.sync_button.setEnabled(True)
            self.sync_button.setText("Sync Email DB")
        QMessageBox.critical(self, "Sync failed", error)

    def _open_check_database(self) -> None:
        import subprocess
        setup_path = get_user_dir() / "setup_app.py"
        if not setup_path.is_file():
            QMessageBox.warning(
                self, "Database viewer unavailable",
                f"Could not find {setup_path}.\n\n"
                "Place setup_app.py beside the application or run it directly.",
            )
            return
        try:
            subprocess.Popen(
                [sys.executable, str(setup_path)],
                cwd=str(setup_path.parent),
            )
        except OSError as error:
            _APP_LOGGER.exception("Could not launch the database viewer.")
            QMessageBox.critical(
                self, "Could not open database viewer", str(error)
            )

    # ------------------------------------------------------------------
    # Interface
    # ------------------------------------------------------------------

    def _build_interface(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(22, 22, 22, 22)

        title = QLabel("Scanner Packing Lists")
        title.setStyleSheet("font-size: 18pt; font-weight: bold;")
        root.addWidget(title)
        hint = QLabel(
            "Search emails by RITM, address, or Reel ID, or use Input Data "
            "to enter a full packing list."
        )
        hint.setStyleSheet("color: #53645e; margin-bottom: 8px;")
        root.addWidget(hint)

        folder_row = QHBoxLayout()
        folder_row.addWidget(QLabel("Summaries folder:"))
        self.summaries_label = QLabel(
            f"{self.summaries_folder}  (auto-set from the Export Folder)"
        )
        self.summaries_label.setStyleSheet("color: #53645e;")
        self.summaries_label.setWordWrap(True)
        folder_row.addWidget(self.summaries_label, 1)
        root.addLayout(folder_row)

        # ------------------------------------------------------------------
        # Search panel — dropdown mode selector
        # ------------------------------------------------------------------
        search_box = QGroupBox("Search for")
        search_box.setStyleSheet(
            "QGroupBox { font-weight: bold; margin-top: 8px; padding-top: 14px;"
            " border: 1px solid #b7c7bf; border-radius: 4px; }"
            "QGroupBox::title { subcontrol-origin: margin;"
            " subcontrol-position: top left; padding: 0 6px; color: #315b52; }"
        )
        search_layout = QVBoxLayout(search_box)
        search_layout.setContentsMargins(12, 18, 12, 12)

        search_help = QLabel(
            "Sync Outlook to update the Email Database, then search the saved "
            "messages. The first sync may take longer."
        )
        search_help.setWordWrap(True)
        search_help.setStyleSheet("color: #53645e;")
        search_layout.addWidget(search_help)

        # --- Mode dropdown -------------------------------------------------
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Mode:"))
        self.search_mode_combo = QComboBox()
        self.search_mode_combo.addItems(["RITM", "Address", "Reel ID", "Input Data"])
        self.search_mode_combo.setFixedWidth(180)
        self.search_mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        mode_row.addWidget(self.search_mode_combo)
        mode_row.addStretch(1)
        search_layout.addLayout(mode_row)

        # --- Input widgets (stacked) --------------------------------------
        self.search_ritm_edit = QLineEdit()
        self.search_ritm_edit.setPlaceholderText("e.g. RITM0011836256")

        self.search_addr_edit = QLineEdit()
        self.search_addr_edit.setPlaceholderText("e.g. 123 Main St, Dallas, TX 75201")

        self.search_reel_edit = QLineEdit()
        self.search_reel_edit.setPlaceholderText("e.g. VZ1965437LMK")

        self.search_input_data_button = QPushButton("Open Input Data")
        self.search_input_data_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white; padding: 8px 16px;"
        )
        # Keep button compact — no horizontal stretch
        self.search_input_data_button.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        self.search_input_data_button.clicked.connect(self._open_input_dialog)

        # Stack them in the layout; we’ll hide/show via _on_mode_changed
        self._mode_widgets = [
            self.search_ritm_edit,
            self.search_addr_edit,
            self.search_reel_edit,
            self.search_input_data_button,
        ]
        for w in self._mode_widgets:
            search_layout.addWidget(w)

        # --- Action buttons (Sync / Search) -------------------------------
        search_btn_row = QHBoxLayout()
        self.sync_button = QPushButton("Sync Email DB")
        self.sync_button.setToolTip(
            "Import matching Outlook Inbox messages into the Email Database. "
            "The first sync may take a while."
        )
        self.sync_button.clicked.connect(self._start_email_sync)
        search_btn_row.addWidget(self.sync_button)

        self.search_button = QPushButton("Search Email Database")
        self.search_button.setStyleSheet(
            "font-weight: bold; background: #4a6a93; color: white; padding: 8px 16px;"
        )
        self.search_button.setToolTip(
            "Search messages already saved in the Email Database."
        )
        self.search_button.clicked.connect(self._run_email_search)
        search_btn_row.addWidget(self.search_button)
        search_btn_row.addStretch(1)
        search_layout.addLayout(search_btn_row)

        # Initial visibility
        self._on_mode_changed(0)

        root.addWidget(search_box)

        # Next Load / Recall buttons
        input_row = QHBoxLayout()
        self.next_load_button = QPushButton("Next Load / Reset")
        self.next_load_button.setStyleSheet(
            "font-weight: bold; background: #4a6a93; color: white;"
            " padding: 10px 20px; font-size: 11pt;"
        )
        self.next_load_button.setToolTip(
            "Clear the saved form so the next Open Input Data starts blank."
        )
        self.next_load_button.clicked.connect(self._next_load)
        input_row.addWidget(self.next_load_button)

        self.recall_input_button = QPushButton("Recall Input Data")
        self.recall_input_button.setStyleSheet(
            "font-weight: bold; background: #8a6d3b; color: white;"
            " padding: 10px 20px; font-size: 11pt;"
        )
        self.recall_input_button.setToolTip("Browse rows from the Input database (manual DB).")
        self.recall_input_button.clicked.connect(self._recall_input_data)
        input_row.addWidget(self.recall_input_button)

        self.recall_export_button = QPushButton("Recall Exported Data")
        self.recall_export_button.setStyleSheet(
            "font-weight: bold; background: #2d5a8a; color: white;"
            " padding: 10px 20px; font-size: 11pt;"
        )
        self.recall_export_button.setToolTip("Browse rows from the Export database (main DB).")
        self.recall_export_button.clicked.connect(self._recall_export_data)
        input_row.addWidget(self.recall_export_button)

        input_row.addStretch(1)
        root.addLayout(input_row)

        db_toggle_row = QHBoxLayout()
        self.ignore_db_checkbox = QCheckBox("Ignore database entirely")
        self.ignore_db_checkbox.setToolTip(
            "When checked, no database reads or writes occur."
        )
        self.ignore_db_checkbox.toggled.connect(self._update_db_toggle_state)
        db_toggle_row.addWidget(self.ignore_db_checkbox)

        self.repeat_load_checkbox = QCheckBox(
            "Repeat load (replace this FRN in the database)"
        )
        self.repeat_load_checkbox.setToolTip(
            "When checked, all rows for this FRN's base (including any letter "
            "variants) are deleted from the database and re-inserted on export."
        )
        db_toggle_row.addWidget(self.repeat_load_checkbox)
        db_toggle_row.addStretch(1)
        root.addLayout(db_toggle_row)

        scan_row = QHBoxLayout()
        scan_row.addWidget(QLabel("Email date range:"))
        self.range_selector = QComboBox()
        for label, value in (
            ("Past 3 months", "1"),
            ("Past 6 months", "2"),
            ("Past 12 months", "3"),
            ("All time", "4"),
        ):
            self.range_selector.addItem(label, value)
        scan_row.addWidget(self.range_selector)
        self.scan_button = QPushButton("Scan Item List")
        self.scan_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.scan_button.clicked.connect(self._start_scan)
        scan_row.addWidget(self.scan_button)
        scan_row.addStretch(1)
        root.addLayout(scan_row)
        self.scan_results_label = QLabel(
            "Matching emails and attachments will appear here."
        )
        self.scan_results_label.setWordWrap(True)
        self.scan_results_label.setStyleSheet("color: #53645e;")
        root.addWidget(self.scan_results_label)

        self.email_table = QTableWidget(0, 8)
        self.email_table.setHorizontalHeaderLabels([
            "RITM", "Address", "Reel ID", "Sender", "Subject",
            "Received", "Outlook Folder", "Attachment(s)"
        ])
        # Enable custom context menu for right-click
        self.email_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.email_table.customContextMenuRequested.connect(self._show_email_context_menu)
        self.email_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.email_table.setSelectionMode(
            QTableWidget.SelectionMode.SingleSelection
        )
        self.email_table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.email_table.setAlternatingRowColors(True)
        self.email_table.verticalHeader().setVisible(False)
        self.email_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.email_table.setMaximumHeight(190)
        self.email_table.itemSelectionChanged.connect(self._update_reply_button)
        # Click handler for Reel ID column (column 2) - passes QTableWidgetItem
        self.email_table.itemClicked.connect(self._on_reel_id_clicked)
        # Double-click any column to open email in Outlook
        self.email_table.itemDoubleClicked.connect(self._on_email_double_clicked)
        root.addWidget(self.email_table)

        # Reply/Download now in right-click context menu on email table


        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels([
            "Packing List ID", "RITM", "Date", "Packing List #",
            "Trucker / ID", "Items", "Gross Total", "Tare Total", "Net Total",
        ])
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.table.setSelectionMode(
            QTableWidget.SelectionMode.ExtendedSelection
        )
        self.table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.table.setMaximumHeight(320)
        root.addWidget(self.table, 1)

        footer = QHBoxLayout()
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #53645e;")
        footer.addWidget(self.status_label, 1)

        self.export_folder_button = QToolButton()
        self.export_folder_button.setText("📁 Export Folder…")
        self.export_folder_button.setToolTip(
            "Choose where the packing list workbooks and per-load folders go"
        )
        self.export_folder_button.clicked.connect(self._save_records)
        footer.addWidget(self.export_folder_button)

        self.database_button = QToolButton()
        self.database_button.setText("🗄 Database…")
        self.database_button.setToolTip(
            "Select or create the SQLite .db file used for duplicate checks"
        )
        self.database_button.clicked.connect(self._choose_database)
        footer.addWidget(self.database_button)

        self.view_database_button = QToolButton()
        self.view_database_button.setText("🗄 View / Edit DB")
        self.view_database_button.setToolTip(
            "Open the Email, Input, and Export databases to review or edit "
            "their rows."
        )
        self.view_database_button.clicked.connect(self._open_check_database)
        footer.addWidget(self.view_database_button)

        root.addLayout(footer)
        self._update_db_toggle_state()

        # Default to Input Data mode (index 3 in the dropdown)
        self.search_mode_combo.setCurrentIndex(3)

    def _update_db_toggle_state(self, *_args) -> None:
        configured = db_configured()
        ignore = self.ignore_db_checkbox.isChecked()
        self.ignore_db_checkbox.setEnabled(configured)
        self.repeat_load_checkbox.setEnabled(configured and not ignore)
        if not configured:
            tooltip = (
                "No database configured. Use the Database button "
                "to select or create the .db file."
            )
            self.ignore_db_checkbox.setToolTip(tooltip)
            self.repeat_load_checkbox.setToolTip(tooltip)

    def _choose_database(self) -> None:
        last_dir = self.app_config.get("database_path", "")
        start_dir = str(Path(last_dir).parent) if last_dir else str(get_user_dir())
        chosen, _ = QFileDialog.getSaveFileName(
            self,
            "Select or create the SQLite database file",
            str(Path(start_dir) / "PackingLists.db"),
            "SQLite database (*.db);;All files (*)",
        )
        if not chosen:
            return
        _APP_LOGGER.info("Packing-list database selected: %s", chosen)
        self.app_config["database_path"] = chosen
        save_app_config(self.app_config)
        self._update_db_toggle_state()
        self.status_label.setText(f"Database set: {chosen}")

    # ------------------------------------------------------------------
    # Input Data → popup dialog
    # ------------------------------------------------------------------

    def _open_input_dialog(self) -> None:
        # Pre-fill the popup with the last saved form, if any.
        seed_items: list[ItemRecord] | None = None
        if self.last_form is not None:
            seed_items = [dict(i) for i in self.last_form.get("items", [])]

        dialog = IncomingPackingListDialog(self, existing_items=seed_items)
        if self.last_form is not None:
            dialog.date_input.setText(
                parse_date(self.last_form["date"]).strftime("%m/%d/%Y")
            )
            dialog.packing_list_input.setText(
                self.last_form.get("packing_list", "")
            )
            driver_plate = self.last_form.get("driver_plate", "")
            if driver_plate:
                trucker, trucker_id = _split_trucker_and_id(driver_plate)
                dialog.trucker_input.setText(trucker)
                dialog.trucker_id_input.setText(trucker_id)
            # If we stashed a Received From value, restore it too.
            received = self.last_form.get("_received_from")
            if received:
                dialog.received_from_input.setText(received)
            dialog._rebuild_items_table()
            dialog._update_totals()

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        record = dialog.result_record
        if record is None:
            return

        # Stash the form so we can re-open with the same data.
        self.last_form = {
            "date": record["date"],
            "packing_list": record["packing_list"],
            "driver_plate": record["driver_plate"],
            "items": [dict(i) for i in record["items"]],
            "_received_from": dialog.received_from_input.text().strip(),
        }

        _APP_LOGGER.info(
            "Input Data form accepted: packing_list=%s date=%s item_count=%d",
            record["packing_list"], record["date"], len(record["items"]),
        )
        self.pending_record = record
        self.status_label.setText(
            f"Loaded {record['packing_list']} ({len(record['items'])} item(s)). "
            "Click Scan Item List to match it against Outlook."
        )

    def _build_current_record(self) -> PackingListRecord | None:
        """
        Build a record from the current input form state.
        Uses pending_record if available, otherwise falls back to last_form.
        """
        # If we have a pending record from Input Data dialog, use it
        if self.pending_record is not None:
            return self.pending_record
        
        # Otherwise build from last_form (saved form data)
        if self.last_form is not None:
            items = self.last_form.get("items", [])
            if not items:
                QMessageBox.information(
                    self, "No items",
                    "No items in the saved form. Open Input Data to add items."
                )
                return None
            return {
                "items": items,
                "packing_list": self.last_form.get("packing_list", ""),
                "date": self.last_form.get("date", ""),
                "driver_plate": self.last_form.get("driver_plate", ""),
                "location": self.last_form.get("location", ""),
                "location_address": self.last_form.get("location_address", ""),
                "ritm": self.last_form.get("ritm", ""),
                "letter": self.last_form.get("letter", ""),
                "group": self.last_form.get("group", ""),
            }
        
        # Nothing available
        QMessageBox.information(
            self, "No packing list loaded",
            "Click Input Data to enter a packing list before scanning."
        )
        return None

    def _next_load(self) -> None:
        answer = QMessageBox.question(
            self,
            "Reset current work?",
            "This clears the visible queue, email results, search fields, and "
            "current input form. Saved database records and application "
            "settings will not be deleted.\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        self.last_form = None
        self.pending_record = None
        self.records.clear()
        self.exported_record_ids_by_directory.clear()
        self.groups.clear()
        self.matched_emails_by_key.clear()
        self.email_search_results_by_key.clear()
        self.visible_email_matches.clear()
        self.record_source_matches.clear()
        self.table.setRowCount(0)
        self.email_table.setRowCount(0)
        self.search_ritm_edit.clear()
        self.search_addr_edit.clear()
        self.search_reel_edit.clear()
        self.scan_results_label.setText(
            "Matching emails and attachments will appear here."
        )
        self.status_label.setText("Ready for the next load.")
        self._update_reply_button()
    def _recall_input_data(self) -> None:
        """Recall from Input Database (manual DB)."""
        self._recall_data_generic(
            db_path=manual_db_path(),
            db_connect_func=manual_db_connect,
            title="Recall Input Data — manual database",
            db_label="Input Database"
        )

    def _recall_export_data(self) -> None:
        """Recall from Export Database (main DB)."""
        if not db_configured():
            QMessageBox.warning(
                self, "No database configured",
                "Configure the export database first using the Database button."
            )
            return
        self._recall_data_generic(
            db_path=db_path(),
            db_connect_func=db_connect,
            title="Recall Exported Data — main database",
            db_label="Export Database"
        )

    def _recall_data_generic(self, db_path: Path, db_connect_func, title: str, db_label: str) -> None:
        """Generic recall dialog for any database with Reel_Load_Database/Reel_Items tables."""
        try:
            conn = db_connect_func()
        except Exception as error:
            _APP_LOGGER.exception(f"Could not open {db_label} for Recall.")
            QMessageBox.critical(self, f"Could not open {db_label}", str(error))
            return

        try:
            rows = conn.execute(
                "SELECT Packing_List_Number, Receiving_Date, RITM_Number, "
                "       Location, Truck, Truck_Number, "
                "       Total_Gross_Weight, Total_Tare, Total_Net_Weight, "
                "       Number_of_Reels "
                "FROM Reel_Load_Database"
            ).fetchall()
        except Exception as error:
            _APP_LOGGER.exception(f"Recall query failed for {db_label}.")
            QMessageBox.critical(self, f"Could not read {db_label}", str(error))
            return
        finally:
            try:
                conn.close()
            except Exception:
                pass

        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(1020)
        dialog.setMinimumHeight(620)

        layout = QVBoxLayout(dialog)

        db_label_widget = QLabel(f"Database: {db_path}")
        db_label_widget.setStyleSheet("color: #53645e;")
        db_label_widget.setWordWrap(True)
        layout.addWidget(db_label_widget)

        # ---- Search + sort controls ----------------------------------
        controls = QHBoxLayout()
        controls.addWidget(QLabel("Search:"))

        search_input = QLineEdit()
        search_input.setPlaceholderText(
            "Filter by packing list #, trucker, trucker ID, RITM, location, date…"
        )
        controls.addWidget(search_input, 1)

        controls.addWidget(QLabel("Sort by:"))
        sort_selector = QComboBox()
        sort_selector.addItem("Date (newest first)", "date_desc")
        sort_selector.addItem("Date (oldest first)", "date_asc")
        sort_selector.addItem("Packing List # (A→Z)", "pl_asc")
        sort_selector.addItem("Packing List # (Z→A)", "pl_desc")
        sort_selector.addItem("Date Created (newest first)", "created_desc")
        controls.addWidget(sort_selector)

        layout.addLayout(controls)

        created_key = lambda r: (r[1] or "")  # Receiving_Date as stand-in

        counter_label = QLabel("")
        counter_label.setStyleSheet("color: #53645e;")
        layout.addWidget(counter_label)

        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels([
            "Packing List #", "Date", "RITM", "Location",
            "Trucker", "Trucker ID", "Reels", "Net",
        ])
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(table, 1)

        layout.addWidget(QLabel("Items for selected packing list:"))
        items_table = QTableWidget(0, 5)
        items_table.setHorizontalHeaderLabels(
            ["Description", "Material", "Gross", "Tare", "Net"]
        )
        items_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        items_table.setAlternatingRowColors(True)
        items_table.verticalHeader().setVisible(False)
        items_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        items_table.setMaximumHeight(180)
        layout.addWidget(items_table)

        row_payloads: list[tuple] = []
        selected_pl: dict[str, str] = {"number": ""}

        def sort_rows(filtered: list[tuple]) -> list[tuple]:
            mode = sort_selector.currentData()
            if mode == "date_desc":
                return sorted(filtered, key=lambda row: row[1] or "", reverse=True)
            if mode == "date_asc":
                return sorted(filtered, key=lambda row: row[1] or "")
            if mode == "pl_asc":
                return sorted(filtered, key=lambda row: (row[0] or "").casefold())
            if mode == "pl_desc":
                return sorted(
                    filtered, key=lambda row: (row[0] or "").casefold(),
                    reverse=True,
                )
            if mode == "created_desc":
                return sorted(filtered, key=created_key, reverse=True)
            return list(filtered)

        def refresh() -> None:
            needle = search_input.text().strip().casefold()
            filtered = [
                row for row in rows
                if not needle or any(
                    needle in str(value).casefold()
                    for value in row if value is not None
                )
            ]
            ordered = sort_rows(filtered)
            row_payloads[:] = ordered
            table.setRowCount(0)
            for db_row in ordered:
                (
                    pl_number, recv_date, ritm, location, truck, truck_id,
                    gross, tare, net, reels,
                ) = db_row
                row_index = table.rowCount()
                table.insertRow(row_index)
                values = (
                    pl_number or "", recv_date or "", ritm or "",
                    location or "", truck or "", truck_id or "",
                    str(reels or 0), f"{net or 0:,}",
                )
                for column, value in enumerate(values):
                    table.setItem(row_index, column, QTableWidgetItem(value))
            counter_label.setText(
                f"{len(ordered)} of {len(rows)} packing list(s)."
            )
            items_table.setRowCount(0)
            selected_pl["number"] = ""
            load_btn.setEnabled(False)

        def load_items() -> None:
            row_index = table.currentRow()
            items_table.setRowCount(0)
            selected_pl["number"] = ""
            if row_index < 0 or row_index >= len(row_payloads):
                load_btn.setEnabled(False)
                return
            pl_number = row_payloads[row_index][0] or ""
            selected_pl["number"] = pl_number
            load_btn.setEnabled(bool(pl_number))
            if not pl_number:
                return
            try:
                conn2 = db_connect_func()
                try:
                    item_rows = conn2.execute(
                        "SELECT Item_Description, Item_Type, Gross, Tare, Net "
                        "FROM Reel_Items WHERE Packing_List_Number = ? "
                        "ORDER BY Item_Description",
                        (pl_number,),
                    ).fetchall()
                finally:
                    conn2.close()
            except Exception:
                _APP_LOGGER.exception(
                    f"Recall item query failed for {pl_number} in {db_label}.",
                )
                return
            for description, item_type, gross, tare, net in item_rows:
                row_i = items_table.rowCount()
                items_table.insertRow(row_i)
                for column, value in enumerate((
                    description or "", item_type or "", f"{gross or 0:,}",
                    f"{tare or 0:,}", f"{net or 0:,}",
                )):
                    items_table.setItem(row_i, column, QTableWidgetItem(value))

        search_input.textChanged.connect(lambda _text: refresh())
        sort_selector.currentIndexChanged.connect(lambda _index: refresh())
        table.itemSelectionChanged.connect(load_items)

        footer = QHBoxLayout()
        footer.addStretch(1)
        load_btn = QPushButton("Load into Input")
        load_btn.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        load_btn.setEnabled(False)
        from functools import partial
        load_btn.clicked.connect(
            partial(self._recall_load_into_input, selected_pl, dialog, db_connect_func)
        )
        footer.addWidget(load_btn)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        footer.addWidget(close_btn)
        layout.addLayout(footer)

        refresh()
        search_input.setFocus()
        dialog.exec()




    def _recall_load_into_input(
        self, selected_pl: dict, dialog: QDialog, db_connect_func,
    ) -> None:
        packing_list_number = selected_pl.get("number", "")
        if not packing_list_number:
            return
        try:
            conn = db_connect_func()
            try:
                master = conn.execute(
                    "SELECT Receiving_Date, Truck, Truck_Number "
                    "FROM Reel_Load_Database WHERE Packing_List_Number = ?",
                    (packing_list_number,),
                ).fetchone()
                item_rows = conn.execute(
                    "SELECT Item_Description, Item_Type, Gross, Tare, Net "
                    "FROM Reel_Items WHERE Packing_List_Number = ? "
                    "ORDER BY Item_Description",
                    (packing_list_number,),
                ).fetchall()
            finally:
                conn.close()
        except Exception as error:
            _APP_LOGGER.exception(
                "Recall Data load query failed for %s.", packing_list_number,
            )
            QMessageBox.critical(self, "Could not read database", str(error))
            return
        if master is None:
            QMessageBox.warning(
                self, "Not found",
                f"{packing_list_number} is not in the manual database anymore.",
            )
            return
        recv_date, truck, truck_id = master
        driver_plate = " / ".join(
            part for part in (truck or "", truck_id or "") if part
        )
        items: list[ItemRecord] = [
            {
                "description": description or "", "type": item_type or "",
                "gross": int(gross or 0), "tare": int(tare or 0),
                "net": int(net or 0),
            }
            for description, item_type, gross, tare, net in item_rows
        ]
        self.last_form = {
            "date": recv_date or date.today().isoformat(),
            "packing_list": packing_list_number,
            "driver_plate": driver_plate,
            "items": items,
            "_received_from": DEFAULT_RECEIVED_FROM,
        }
        self.pending_record = None
        _APP_LOGGER.info(
            "Recall Data loaded into form: packing_list=%s item_count=%d",
            packing_list_number, len(items),
        )
        self.status_label.setText(
            f"Recalled {packing_list_number} ({len(items)} item(s)). "
            "Click Open Input Data to view and edit it."
        )
        dialog.accept()

    def _on_mode_changed(self, index: int) -> None:
        """
        Show only the widget that matches the selected dropdown index.
        Hide the 'Search Email Database' button when Input Data (index 3) is active.
        """
        for i, w in enumerate(self._mode_widgets):
            w.setVisible(i == index)

        # Search button only for RITM / Address / Reel ID (indices 0,1,2)
        self.search_button.setVisible(index != 3)
    def _run_email_search(self) -> None:
        mode = self.search_mode_combo.currentIndex()   # 0=RITM, 1=Address, 2=Reel ID
        criteria = {}

        if mode == 0:
            val = self.search_ritm_edit.text().strip()
            if val:
                criteria["ritm"] = val
        elif mode == 1:
            val = self.search_addr_edit.text().strip()
            if val:
                criteria["address"] = val
        elif mode == 2:
            val = self.search_reel_edit.text().strip()
            if val:
                criteria["reel_id"] = val
        else:
            return

        if not criteria:
            QMessageBox.information(
                self, "Nothing to search",
                "Fill in the field before searching."
            )
            return
        self.search_button.setEnabled(False)
        self.status_label.setText("Searching Email Database...")
        self.scan_results_label.setText("Searching matching emails...")
        self.email_table.setRowCount(0)
        self.email_search_results_by_key.clear()
        self._update_reply_button()
        worker = EmailSearchWorker(criteria)
        self.email_search_worker = worker
        worker.results_ready.connect(self._display_email_search_results)
        worker.failed.connect(self._email_search_failed)
        worker.start()

    def _display_email_search_results(self, results: list[dict]) -> None:
        self.search_button.setEnabled(True)

        # Fetch reel items for each result (from Email_Item table via RITM)
        for result in results:
            ritm = result.get("ritm", "")
            if ritm:
                try:
                    conn = email_db_connect()
                    try:
                        rows = conn.execute(
                            "SELECT Item_ID, Asset_Class, Job_Type, Equipment_Type, Quantity, "
                            "       Description, Gross, Tare, Net "
                            "FROM Email_Item WHERE Ritm = ?",
                            (ritm,),
                        ).fetchall()
                        items = []
                        for r in rows:
                            items.append({
                                "item_id": r[0],
                                "asset_class": r[1],
                                "job_type": r[2],
                                "equipment_type": r[3],
                                "quantity": r[4],
                                "description": r[5],
                                "gross": r[6] or 0,
                                "tare": r[7] or 0,
                                "net": r[8] if r[8] is not None else max((r[6] or 0) - (r[7] or 0), 0),
                            })
                        result["items"] = items
                    finally:
                        conn.close()
                except Exception as e:
                    _APP_LOGGER.exception("Failed to fetch reel items for RITM %s", ritm)
                    result["items"] = []
            else:
                result["items"] = []

        # 1) Populate the main table
        self.email_table.setRowCount(0)
        self.email_search_results_by_key.clear()
        for result in results:
            key = f"search:{result['email_id']}:{result['ritm']}"
            self.email_search_results_by_key[key] = result
            row = self.email_table.rowCount()
            self.email_table.insertRow(row)
            sender = result["sender"]
            if (
                result["sender_email"]
                and result["sender_email"].casefold() not in sender.casefold()
            ):
                sender = f"{sender} <{result['sender_email']}>".strip()

            values = (
                result["ritm"], result["address"], "Reels",  # Column 2 always says "Reels"
                sender, result["subject"], result["received"], result["folder"],
                ", ".join(result["attachments"]),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, key)
                # Column 2 (Reel ID) — clickable "Reels" hyperlink with tooltip
                if column == 2:
                    reel_items = result.get("items", [])
                    if reel_items:
                        # Style as hyperlink
                        item.setForeground(Qt.GlobalColor.blue)
                        font = item.font()
                        font.setUnderline(True)
                        item.setFont(font)
                        preview = self._build_reel_preview_text(reel_items)
                        item.setToolTip(f"Click to view reels\n\n{preview}")
                        # Store reel data for click handler
                        item.setData(Qt.ItemDataRole.UserRole + 1, reel_items)
                        item.setData(Qt.ItemDataRole.UserRole + 2, result["ritm"])
                    else:
                        item.setForeground(Qt.GlobalColor.gray)
                        item.setToolTip("No reel data available")
                self.email_table.setItem(row, column, item)

        # 2) Show the popup dialog
        dialog = EmailSearchResultsDialog(self, results)
        dialog.exec()

        # 3) Update status labels
        if results:
            self.scan_results_label.setText(
                f"Found {len(results)} matching email result(s). "
                "Right Click a row to reply or download its attachments."
                "Double Click a row to Open Email in Outlook"
            )
            self.status_label.setText(f"{len(results)} matching email(s).")
        else:
            self.scan_results_label.setText("No matching emails were found.")
            self.status_label.setText("No matching emails found.")

        self._update_reply_button()

    def _email_search_failed(self, error: str) -> None:
        self.search_button.setEnabled(True)
        self.scan_results_label.setText("Email search failed.")
        self.status_label.setText("Email search failed.")
        QMessageBox.critical(self, "Email search failed", error)

    def _selected_email_key(self) -> str | None:
        row = self.email_table.currentRow()
        if row < 0:
            return None
        item = self.email_table.item(row, 0)
        key = item.data(Qt.ItemDataRole.UserRole) if item else None
        return key if isinstance(key, str) else None

    def _update_reply_button(self) -> None:
        """Kept for compatibility but buttons removed — context menu used instead."""
        pass

    # ============================================================
    # EMAIL TABLE INTERACTIONS — Complete Working Implementation
    # ============================================================

    def _on_reel_id_clicked(self, item: QTableWidgetItem) -> None:
        """Handle click on Reel ID column (column 2)."""
        if item is None or item.column() != 2:
            return
        reel_items = item.data(Qt.ItemDataRole.UserRole + 1)
        ritm = item.data(Qt.ItemDataRole.UserRole + 2)
        if reel_items:
            self._show_reel_popup(reel_items, ritm or "")
        else:
            # Fallback: find by RITM in row
            ritm_item = self.email_table.item(item.row(), 0)
            if ritm_item:
                self._show_reels_for_ritm(ritm_item.text())

    def _on_email_double_clicked(self, item: QTableWidgetItem) -> None:
        """Double-click any column to open email in Outlook."""
        row = item.row()
        ritm_item = self.email_table.item(row, 0)
        if not ritm_item:
            return
        key = ritm_item.data(Qt.ItemDataRole.UserRole)
        if not key:
            return
        
        # Try search results first
        result = self.email_search_results_by_key.get(key)
        if result and result.get("entry_id"):
            self._open_outlook_message(result["entry_id"], result.get("store_id", ""))
            return
        
        # Try matched emails
        email = self.matched_emails_by_key.get(key)
        if email and email.entry_id:
            self._open_outlook_message(email.entry_id, email.store_id)

    def _open_outlook_message(self, entry_id: str, store_id: str = "") -> None:
        """Open email in Outlook by EntryID."""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            message = None
            if store_id:
                try:
                    message = namespace.GetItemFromID(entry_id, store_id)
                except Exception:
                    pass
            if message is None:
                message = namespace.GetItemFromID(entry_id)
            if message:
                message.Display()
        except Exception as e:
            _APP_LOGGER.exception("Failed to open email in Outlook")
            QMessageBox.warning(self, "Could not open email", str(e))

    def _show_reels_for_ritm(self, ritm: str) -> None:
        """Find reel items for a RITM and show popup."""
        reel_items = []
        # Search in matched emails (scan results)
        for email in self.visible_email_matches:
            if email.ritm == ritm:
                for record in self.records:
                    if (record["packing_list"].casefold() == email.packing_list.casefold() and
                        (record["ritm"] or "").casefold() == email.ritm.casefold()):
                        for it in record.get("items", []):
                            reel_items.append({
                                "item_id": it.get("description", ""),
                                "asset_class": it.get("type", ""),
                                "job_type": "",
                                "equipment_type": it.get("type", ""),
                                "quantity": 1,
                                "description": it.get("description", ""),
                                "gross": it.get("gross", 0),
                                "tare": it.get("tare", 0),
                                "net": it.get("net", 0),
                            })
                if reel_items:
                    break
        # Search in search results
        if not reel_items:
            for key, result in self.email_search_results_by_key.items():
                if result.get("ritm") == ritm:
                    reel_items = result.get("items", [])
                    break
        if reel_items:
            self._show_reel_popup(reel_items, ritm)

    def _show_email_context_menu(self, pos) -> None:
        """Right-click context menu on email table with blue highlight."""
        from functools import partial
        
        row = self.email_table.rowAt(pos.y())
        if row < 0:
            return
        
        # Select and highlight the row
        self.email_table.clearSelection()
        self.email_table.selectRow(row)
        
        item = self.email_table.item(row, 0)
        key = item.data(Qt.ItemDataRole.UserRole) if item else None
        if not key:
            return

        # Get reel data for this row
        reel_item = self.email_table.item(row, 2)
        reel_items = reel_item.data(Qt.ItemDataRole.UserRole + 1) if reel_item else None
        ritm = reel_item.data(Qt.ItemDataRole.UserRole + 2) if reel_item else None

        # Create menu with stylesheet for blue highlight
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background: white;
                border: 1px solid #9baea6;
                border-radius: 4px;
                padding: 4px;
            }
            QMenu::item {
                padding: 8px 24px 8px 16px;
                border-radius: 3px;
                color: #20332e;
            }
            QMenu::item:selected {
                background: #315b52;
                color: white;
            }
            QMenu::item:disabled {
                color: #9baea6;
            }
            QMenu::separator {
                height: 1px;
                background: #d7dfda;
                margin: 4px 8px;
            }
        """)

        # Check capabilities
        search_result = self.email_search_results_by_key.get(key or "")
        matched_email = self.matched_emails_by_key.get(key or "")
        
        can_reply = bool(
            (search_result and search_result.get("entry_id")) or
            (matched_email and (matched_email.entry_id or (matched_email.subject and matched_email.folder)))
        )
        has_source = bool(
            (search_result and search_result.get("entry_id")) or
            (matched_email and (matched_email.entry_id or matched_email.subject))
        )

        # Reply action
        reply_action = menu.addAction("Reply to Email")
        reply_action.setEnabled(can_reply)
        reply_action.triggered.connect(partial(self._reply_to_email_by_key, key))

        # Download action
        download_action = menu.addAction("Download Attachments")
        download_action.setEnabled(has_source)
        download_action.triggered.connect(partial(self._download_attachments_by_key, key))

        menu.addSeparator()

        # View Reels action
        view_reels_action = menu.addAction("View Reels")
        view_reels_action.setEnabled(bool(reel_items))
        if reel_items:
            view_reels_action.triggered.connect(partial(self._show_reel_popup, reel_items, ritm or ""))

        # Show menu at cursor position
        menu.exec(self.email_table.viewport().mapToGlobal(pos))

    def _reply_to_email_by_key(self, key: str) -> None:
        """Reply to email using a specific key."""
        # Select the row for this key
        for row in range(self.email_table.rowCount()):
            item = self.email_table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == key:
                self.email_table.clearSelection()
                self.email_table.selectRow(row)
                break
        self._reply_to_selected_email()

    def _download_attachments_by_key(self, key: str) -> None:
        """Download attachments using a specific key."""
        for row in range(self.email_table.rowCount()):
            item = self.email_table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == key:
                self.email_table.clearSelection()
                self.email_table.selectRow(row)
                break
        self._download_selected_attachments()

    @staticmethod
    def _build_reel_preview_text(items: list[dict], max_lines: int = 5) -> str:
        """Build a short preview string for tooltip."""
        if not items:
            return "No reel data"
        lines = []
        for i, item in enumerate(items[:max_lines]):
            serial = item.get("item_id") or item.get("description") or "—"
            gross = item.get("gross", 0)
            net = item.get("net", 0)
            lines.append(f"  • {serial}  (Gross: {gross:,}, Net: {net:,})")
        if len(items) > max_lines:
            lines.append(f"  … and {len(items) - max_lines} more")
        return "\n".join(lines)

    def _show_reel_popup(self, items: list[dict], ritm: str) -> None:
        """Show a dialog listing all reel IDs with details."""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Reels for {ritm or 'Unknown RITM'}")
        dialog.setMinimumSize(700, 400)
        dialog.setModal(True)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)

        # Header
        header = QLabel(f"{len(items)} reel(s) for {ritm}")
        header.setStyleSheet("font-size: 12pt; font-weight: bold; color: #20332e;")
        layout.addWidget(header)

        has_attachment_ids = any(
            item.get("attachment_item_id") for item in items
        )
        table = QTableWidget(len(items), 3)
        if has_attachment_ids:
            table.setHorizontalHeaderLabels([
                "Reel ID entered", "Type", "Reel ID from attachment"
            ])
        else:
            table.setHorizontalHeaderLabels([
                "Serial / Item ID", "Type", "Description"
            ])
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        table.setStyleSheet("""
            QTableWidget { 
                font-family: 'Consolas', 'Monospace'; 
                font-size: 10pt; 
                gridline-color: #d7dfda;
            }
            QTableWidget::item { padding: 6px; }
            QTableWidget::item:selected { background: #d5e1da; color: #20332e; }
            QHeaderView::section { background: #e2e9e5; padding: 6px; border: none; font-weight: bold; }
        """)

        for row, item in enumerate(items):
            serial = item.get("item_id") or item.get("description") or "—"
            asset_class = item.get("asset_class") or item.get("equipment_type") or "—"
            desc = (
                item.get("attachment_item_id", "")
                if has_attachment_ids else item.get("description") or ""
            )
            
            values = (serial, asset_class, desc)
            for col, val in enumerate(values):
                table_item = QTableWidgetItem(str(val))
                table_item.setToolTip(str(val))
                table.setItem(row, col, table_item)

        layout.addWidget(table, 1)

        # Buttons
        btn_row = QHBoxLayout()
        copy_btn = QPushButton("Copy Selected Serials")
        copy_btn.setToolTip("Copy selected reel serial numbers to clipboard (Ctrl+C also works)")
        copy_btn.clicked.connect(lambda: self._copy_table_selection(table))
        btn_row.addWidget(copy_btn)
        
        copy_all_btn = QPushButton("Copy All Serials")
        copy_all_btn.setToolTip("Copy all reel serial numbers to clipboard")
        copy_all_btn.clicked.connect(lambda: self._copy_all_serials(items))
        btn_row.addWidget(copy_all_btn)
        
        btn_row.addStretch(1)
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("font-weight: bold; padding: 8px 24px;")
        close_btn.clicked.connect(dialog.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        # Enable Ctrl+C on table
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
        table.installEventFilter(self)
        dialog._reel_table = table  # Store for event filter
        
        dialog.exec()

    def _copy_table_selection(self, table: QTableWidget) -> None:
        """Copy selected serials (column 0) to clipboard."""
        selected = table.selectedItems()
        if not selected:
            return
        # Get unique rows from selection
        rows = set(item.row() for item in selected)
        serials = []
        for row in sorted(rows):
            item = table.item(row, 0)
            if item:
                serials.append(item.text())
        if serials:
            QApplication.clipboard().setText("\n".join(serials))
            self.status_label.setText(f"Copied {len(serials)} reel serial(s) to clipboard")

    def _copy_all_serials(self, items: list[dict]) -> None:
        """Copy all reel serials to clipboard."""
        serials = [item.get("item_id") or item.get("description") or "" for item in items]
        serials = [s for s in serials if s]
        if serials:
            QApplication.clipboard().setText("\n".join(serials))
            self.status_label.setText(f"Copied {len(serials)} reel serial(s) to clipboard")

    # Add event filter for Ctrl+C on reel popup table
    def eventFilter(self, watched: QWidget, event: QEvent) -> bool:
        if event.type() == QEvent.Type.KeyPress and event.matches(QKeySequence.StandardKey.Copy):
            # Check if it's our reel popup table
            if hasattr(watched, 'parent') and hasattr(watched.parent(), '_reel_table'):
                if watched is watched.parent()._reel_table:
                    self._copy_table_selection(watched)
                    return True
        return super().eventFilter(watched, event)

    
    def _copy_selected_reels(self, list_widget: QListWidget) -> None:
        """Copy selected reel serials to clipboard."""
        selected = list_widget.selectedItems()
        if not selected:
            return
        serials = [item.data(Qt.ItemDataRole.UserRole).get("serial", "") for item in selected]
        QApplication.clipboard().setText("\n".join(serials))
        self.status_label.setText(f"Copied {len(serials)} reel serial(s) to clipboard")

    def _copy_all_reels(self, items: list[dict]) -> None:
        """Copy all reel serials to clipboard."""
        serials = [item.get("item_id") or item.get("description") or "" for item in items]
        serials = [s for s in serials if s]
        QApplication.clipboard().setText("\n".join(serials))
        self.status_label.setText(f"Copied {len(serials)} reel serial(s) to clipboard")

    def _reply_to_selected_email(self) -> None:
        key = self._selected_email_key()
        if not key:
            return
        result = self.email_search_results_by_key.get(key)
        if result is not None:
            self._open_reply_popup(result)
            return
        email = self.matched_emails_by_key.get(key)
        if email is None:
            QMessageBox.warning(
                self, "Email cannot be opened",
                "The selected email could not be matched to its Outlook record.",
            )
            return
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        message = self._find_outlook_message(namespace, email)
        if message is None:
            QMessageBox.warning(
                self, "Email cannot be opened",
                "Outlook could not find the selected email by its ID or its "
                "subject, sender, date, and folder.",
            )
            return
        ritm_key = email.ritm.casefold()
        ritm_records = [
            record for record in self.records
            if (record["ritm"] or "").casefold() == ritm_key
            and record["packing_list"].casefold() == email.packing_list.casefold()
        ]
        item_descriptions = list(dict.fromkeys(
            item["description"]
            for record in ritm_records for item in record["items"]
        )) if ritm_records else sorted(email.item_terms)
        date_values = sorted({
            parse_date(record["date"]).strftime("%m/%d/%Y")
            for record in ritm_records
        })
        reply_body = (
            f"<p>Delivered {html.escape(', '.join(date_values))}</p>"
            f"<p>{'<br>'.join(html.escape(text) for text in item_descriptions)}</p>"
            "<p>Thank you</p><br>"
        )
        try:
            reply = message.ReplyAll()
            reply.HTMLBody = reply_body + reply.HTMLBody
            reply.Display()
        except Exception as error:
            _APP_LOGGER.exception("Could not create Outlook reply for %s.", email.ritm)
            QMessageBox.critical(self, "Could not prepare reply", str(error))

    def _open_reply_popup(self, result: dict) -> None:
        if not result.get("entry_id"):
            QMessageBox.warning(
                self, "No source email",
                "This row has no linked Outlook email to reply to.",
            )
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Reply - {result.get('subject') or '(no subject)'}")
        dialog.setMinimumSize(700, 500)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            f"To: {result.get('sender', '')} <{result.get('sender_email', '')}>"
        ))
        layout.addWidget(QLabel(f"Subject: Re: {result.get('subject', '')}"))
        body_edit = QPlainTextEdit()
        body_edit.setPlaceholderText("Type your reply...")
        layout.addWidget(body_edit, 1)
        buttons = QHBoxLayout()
        open_reply_button = QPushButton("Open Reply in Outlook")
        buttons.addWidget(open_reply_button)
        buttons.addStretch(1)
        cancel = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        cancel.rejected.connect(dialog.reject)
        buttons.addWidget(cancel)
        layout.addLayout(buttons)

        def open_reply() -> None:
            body = body_edit.toPlainText().strip()
            if not body:
                QMessageBox.information(dialog, "Empty reply", "Type a message first.")
                return
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                message = None
                if result.get("store_id"):
                    try:
                        message = namespace.GetItemFromID(
                            result["entry_id"], result["store_id"]
                        )
                    except Exception:
                        message = None
                if message is None:
                    message = namespace.GetItemFromID(result["entry_id"])
                reply = message.ReplyAll()
                reply.Body = body + "\n\n" + reply.Body
                reply.Display()
                dialog.accept()
            except Exception as error:
                _APP_LOGGER.exception("Could not create Outlook reply.")
                QMessageBox.critical(dialog, "Reply failed", str(error))

        open_reply_button.clicked.connect(open_reply)
        dialog.exec()

    def _download_selected_attachments(self) -> None:
        key = self._selected_email_key()
        if not key:
            return
        result = self.email_search_results_by_key.get(key)
        if result is not None:
            entry_id = result.get("entry_id", "")
            store_id = result.get("store_id", "")
            email = None
        else:
            email = self.matched_emails_by_key.get(key)
            if email is None:
                return
            entry_id = email.entry_id
            store_id = email.store_id
        destination = QFileDialog.getExistingDirectory(
            self, "Choose where to save email attachments"
        )
        if not destination:
            return
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            if email is not None:
                message = self._find_outlook_message(namespace, email)
            else:
                message = None
                if store_id:
                    try:
                        message = namespace.GetItemFromID(entry_id, store_id)
                    except Exception:
                        message = None
                if message is None:
                    message = namespace.GetItemFromID(entry_id)
            if message is None:
                QMessageBox.warning(
                    self, "Email cannot be opened",
                    "Outlook could not find the selected email.",
                )
                return
            attachments = list(message.Attachments)
        except Exception as error:
            _APP_LOGGER.exception("Could not read attachments from Outlook.")
            QMessageBox.critical(
                self, "Could not read email attachments", str(error),
            )
            return
        if not attachments:
            QMessageBox.information(
                self, "No attachments", "The selected email has no attachments."
            )
            return
        saved: list[str] = []
        failures: list[str] = []
        for attachment in attachments:
            filename = str(attachment.FileName or "")
            if not filename:
                failures.append("An attachment has no filename.")
                continue
            try:
                output_path = Path(destination) / Path(filename).name
                suffix = output_path.suffix
                stem = output_path.stem
                number = 2
                while output_path.exists():
                    output_path = output_path.with_name(
                        f"{stem} ({number}){suffix}"
                    )
                    number += 1
                attachment.SaveAsFile(str(output_path))
                saved.append(str(output_path))
            except Exception as error:
                _APP_LOGGER.exception(
                    "Could not save email attachment %s.", filename,
                )
                failures.append(f"{filename}: {error}")
        details = [f"Saved {len(saved)} attachment(s) to {destination}."]
        if failures:
            details.append("Could not save:\n" + "\n".join(failures))
        if saved:
            QMessageBox.information(self, "Attachments downloaded", "\n".join(details))
        else:
            QMessageBox.warning(self, "Attachments not downloaded", "\n".join(details))

    # ------------------------------------------------------------------
    # Queue helpers
    # ------------------------------------------------------------------

    def _append_record(self, record: PackingListRecord) -> None:
        self.records.append(record)
        row = self.table.rowCount()
        self.table.insertRow(row)
        gross_total = sum(item["gross"] for item in record["items"])
        tare_total = sum(item["tare"] for item in record["items"])
        net_total = sum(item["net"] for item in record["items"])
        values = (
            _packing_list_identifier(record),
            record["ritm"],
            record["date"],
            record["packing_list"],
            record["driver_plate"],
            str(len(record["items"])),
            f"{gross_total:,}",
            f"{tare_total:,}",
            f"{net_total:,}",
        )
        for column, value in enumerate(values):
            self.table.setItem(row, column, QTableWidgetItem(value))

    def _location_from_address(self, address: str) -> str:
        parts = [p.strip() for p in address.split(",")]
        if len(parts) < 2:
            return address.strip()
        city_state = ", ".join(parts[1:])
        city_state = re.sub(r"\s+\d{5}(?:-\d{4})?\s*$", "", city_state).strip()
        return city_state.replace(",", "")

    def _build_unmatched_records(
        self,
        record: PackingListRecord,
        items: list[ItemRecord],
        staged_ritm_records: list[tuple[PackingListRecord, list[ExactMatch]]],
    ) -> list[PackingListRecord]:
        packing_list_key = record["packing_list"].casefold()
        buckets: dict[tuple[str, str], list[ItemRecord]] = {}
        for item in items:
            addr = _item_address(item)
            r = _item_ritm(item)
            buckets.setdefault((addr, r), []).append(item)

        def existing_ritm_records() -> dict[str, PackingListRecord]:
            return {
                rec["ritm"].casefold(): rec
                for rec in self.records
                if rec["packing_list"].casefold() == packing_list_key
                and not _is_no_request_item(rec.get("ritm"))
            }

        def matches_for_ritm(ritm_value: str) -> list[ExactMatch]:
            if not ritm_value:
                return []
            ritm_key = ritm_value.casefold()
            collected: list[ExactMatch] = []
            for staged_rec, matches in staged_ritm_records:
                if (staged_rec["ritm"] or "").casefold() == ritm_key:
                    collected.extend(matches)
            for match in _record_source_matches(record):
                if any(r.casefold() == ritm_key for r in match.subject_ritms):
                    collected.append(match)
            return _dedupe_matches(collected)

        new_records: list[PackingListRecord] = []
        for (address, ritm), group_items in buckets.items():
            if ritm:
                chosen_ritm = self._resolve_ritm_typo(ritm, packing_list_key)
                target = existing_ritm_records().get(chosen_ritm.casefold())
                if target is not None:
                    self._merge_items_into_record(
                        target, group_items, packing_list_key
                    )
                    continue
                source_matches = matches_for_ritm(chosen_ritm)
                new_record: PackingListRecord = {
                    **record,
                    "group": chosen_ritm.casefold(),
                    "ritm": chosen_ritm,
                    "letter": "",
                    "location": self._location_from_address(address) if address else "",
                    "location_address": address,
                    "items": group_items,
                    "source_matches": source_matches,
                }
                self._append_record(new_record)
                new_records.append(new_record)
            else:
                new_record: PackingListRecord = {
                    **record,
                    "group": "no request item",
                    "ritm": "No Request Item",
                    "letter": "",
                    "location": self._location_from_address(address) if address else "",
                    "location_address": address,
                    "items": group_items,
                    "source_matches": [],
                }
                self._append_record(new_record)
                new_records.append(new_record)
        return new_records

    def _merge_items_into_record(
        self,
        target: PackingListRecord,
        items: list[ItemRecord],
        packing_list_key: str,
    ) -> None:
        remaining = list(items)
        current = target
        while remaining:
            space = MAX_ITEMS_PER_LOAD - len(current["items"])
            if space <= 0:
                sibling: PackingListRecord = {
                    **current,
                    "letter": "",
                    "items": [],
                    "source_matches": list(current.get("source_matches") or []),
                }
                self._append_record(sibling)
                current = sibling
                space = MAX_ITEMS_PER_LOAD
            take = min(space, len(remaining))
            current["items"].extend(remaining[:take])
            remaining = remaining[take:]
        self._refresh_queue_table()

    def _resolve_ritm_typo(self, typed_ritm: str, packing_list_key: str) -> str:
        typed_norm = normalize_ritm(typed_ritm)
        known = {
            rec["ritm"]: rec["ritm"]
            for rec in self.records
            if rec["packing_list"].casefold() == packing_list_key
            and not _is_no_request_item(rec.get("ritm"))
        }
        if not known:
            return typed_ritm
        for existing in known:
            if normalize_ritm(existing) == typed_norm:
                return existing
        candidates: list[tuple[str, float]] = []
        for existing in known:
            existing_norm = normalize_ritm(existing)
            ratio = SequenceMatcher(None, typed_norm, existing_norm).ratio()
            if ratio >= 0.90:
                candidates.append((existing, ratio))
        if not candidates:
            return typed_ritm
        candidates.sort(key=lambda x: -x[1])
        best, ratio = candidates[0]
        answer = QMessageBox.question(
            self,
            "Possible RITM typo",
            f"You typed: {typed_ritm}\n\n"
            f"A Request Item on this load is very close:\n"
            f"    {best}   (similarity {ratio:.0%})\n\n"
            f"Did you mean {best}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer == QMessageBox.StandardButton.Yes:
            return best
        return typed_ritm

    def _refresh_queue_table(self) -> None:
        self.table.setRowCount(0)
        for record in self.records:
            row = self.table.rowCount()
            self.table.insertRow(row)
            gross_total = sum(item["gross"] for item in record["items"])
            tare_total = sum(item["tare"] for item in record["items"])
            net_total = sum(item["net"] for item in record["items"])
            values = (
                _packing_list_identifier(record),
                record["ritm"],
                record["date"],
                record["packing_list"],
                record["driver_plate"],
                str(len(record["items"])),
                f"{gross_total:,}",
                f"{tare_total:,}",
                f"{net_total:,}",
            )
            for column, value in enumerate(values):
                self.table.setItem(row, column, QTableWidgetItem(value))

    def _set_processing_state(self, enabled: bool) -> None:
        for widget in (
            self.search_input_data_button,
            self.search_button,
            self.sync_button,              # Added
            self.next_load_button,
            self.recall_input_button,      # ✅ Fixed
            self.recall_export_button,     # ✅ Fixed
            self.scan_button,
            self.export_folder_button,
            self.database_button,          # Added
            self.view_database_button,     # Added
            self.range_selector,           # Added
            self.ignore_db_checkbox,       # Added
            self.repeat_load_checkbox,     # Added
            self.search_mode_combo,        # Added
            self.search_ritm_edit,         # Added
            self.search_addr_edit,         # Added
            self.search_reel_edit,         # Added
        ):
            if widget is not None:
                widget.setEnabled(enabled)

    # ------------------------------------------------------------------
    # Email merge / reply
    # ------------------------------------------------------------------

    def _merge_matched_emails(
        self,
        exact_matches: list[ExactMatch],
        ritm_letters: dict[str, str],
        packing_list: str,
        forced_ritms: set[str] | None = None,
    ) -> None:
        for match in exact_matches:
            message_key = match.entry_id or "|".join(
                (match.folder, match.subject, match.received)
            )
            email_ritms = set(match.subject_ritms or match.ritms)
            email_ritms.update(forced_ritms or set())
            for ritm in email_ritms:
                ritm_key = ritm.casefold()
                letter = ritm_letters.get(ritm_key, "")
                email_key = f"{message_key}|{ritm_key}|{packing_list.casefold()}|{letter}"
                email = self.matched_emails_by_key.get(email_key)
                if email is None:
                    email = MatchedEmail(
                        key=email_key, ritm=ritm, packing_list=packing_list,
                        letter=letter, entry_id=match.entry_id,
                        store_id=match.store_id, sender=match.sender,
                        sender_email=match.sender_email, subject=match.subject,
                        received=match.received, folder=match.folder,
                        location=match.location,
                        location_address=match.location_address,
                    )
                    self.matched_emails_by_key[email_key] = email
                else:
                    email.entry_id = match.entry_id or email.entry_id
                    email.store_id = match.store_id or email.store_id
                    email.sender = match.sender or email.sender
                    email.sender_email = match.sender_email or email.sender_email
                    email.subject = match.subject or email.subject
                    email.received = match.received or email.received
                    email.folder = match.folder or email.folder
                    email.location = match.location or email.location
                    email.location_address = (
                        match.location_address or email.location_address
                    )
                email.attachments.add(match.filename)
                email.item_terms.add(match.term)
                if match.attachment_item_id:
                    email.attachment_item_ids.setdefault(
                        match.term.casefold(), set()
                    ).add(match.attachment_item_id)
        self.visible_email_matches = sorted(
            self.matched_emails_by_key.values(),
            key=lambda email: (
                email.ritm.casefold(), email.received,
                email.subject.casefold(), email.folder.casefold(),
            ),
            reverse=True,
        )
        self.email_search_results_by_key.clear()
        self.email_table.setRowCount(0)
        for email in self.visible_email_matches:
            row = self.email_table.rowCount()
            self.email_table.insertRow(row)
            sender = email.sender
            if email.sender_email and email.sender_email.casefold() not in sender.casefold():
                sender = f"{sender} <{email.sender_email}>".strip()
            location_address = next(
                (
                    record.get("location_address", "")
                    for record in self.records
                    if record["packing_list"].casefold()
                    == email.packing_list.casefold()
                    and (record["ritm"] or "").casefold()
                    == email.ritm.casefold()
                ),
                email.location_address or email.location or "",
            )
            values = (
                email.ritm, location_address,
                ", ".join(sorted(email.item_terms)),
                sender, email.subject, email.received, email.folder,
                ", ".join(sorted(email.attachments)),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, email.key)
                # Column 2 (Reel ID) — clickable "Reels" hyperlink with tooltip
                if column == 2:
                    # Gather reel items from staged records
                    reel_items = []
                    for record in self.records:
                        if (record["packing_list"].casefold() == email.packing_list.casefold() and
                            (record["ritm"] or "").casefold() == email.ritm.casefold()):
                            for record_item in record.get("items", []):
                                input_item_id = record_item.get("description", "")
                                reel_items.append({
                                    "item_id": input_item_id,
                                    "attachment_item_id": ", ".join(sorted(
                                        email.attachment_item_ids.get(
                                            input_item_id.casefold(), set()
                                        )
                                    )),
                                    "asset_class": record_item.get("type", ""),
                                    "job_type": "",
                                    "equipment_type": record_item.get("type", ""),
                                    "quantity": 1,
                                    "description": record_item.get("description", ""),
                                    "gross": record_item.get("gross", 0),
                                    "tare": record_item.get("tare", 0),
                                    "net": record_item.get("net", 0),
                                })
                            break
                    
                    # Fallback: use email's item_terms if no staged records found
                    if not reel_items and email.item_terms:
                        for term in email.item_terms:
                            reel_items.append({
                                "item_id": term,
                                "attachment_item_id": ", ".join(sorted(
                                    email.attachment_item_ids.get(
                                        term.casefold(), set()
                                    )
                                )),
                                "asset_class": "",
                                "job_type": "",
                                "equipment_type": "",
                                "quantity": 1,
                                "description": term,
                                "gross": 0,
                                "tare": 0,
                                "net": 0,
                            })
                    if reel_items:
                        item.setForeground(Qt.GlobalColor.blue)
                        font = item.font()
                        font.setUnderline(True)
                        item.setFont(font)
                        item.setToolTip(
                            f"Click to view reels\n\n"
                            f"{self._build_reel_preview_text(reel_items)}"
                        )
                        item.setData(Qt.ItemDataRole.UserRole + 1, reel_items)
                        item.setData(Qt.ItemDataRole.UserRole + 2, email.ritm)
                    else:
                        item.setForeground(Qt.GlobalColor.gray)
                        item.setToolTip("No reel data available")
                self.email_table.setItem(row, column, item)
        self._update_reply_button()

    def _on_rescan_sources(self, payload) -> None:
        try:
            ritm, matches, packing_list = payload
        except Exception:
            return
        if not matches:
            return
        ritm_key = (ritm or "").casefold()
        packing_list_key = (packing_list or "").casefold()
        for record in self.records:
            if record["packing_list"].casefold() != packing_list_key:
                continue
            if (record["ritm"] or "").casefold() != ritm_key:
                continue
            current = list(record.get("source_matches") or [])
            seen = {(m.entry_id, m.filename, m.term) for m in current}
            for match in matches:
                key = (match.entry_id, match.filename, match.term)
                if key in seen:
                    continue
                seen.add(key)
                current.append(match)
            record["source_matches"] = current
            if not record.get("location_address"):
                for match in matches:
                    if match.location_address:
                        record["location_address"] = match.location_address
                        record["location"] = match.location
                        break
            identifier = _packing_list_identifier(record)
            self.record_source_matches[identifier] = current
        ritm_letters = {ritm_key: ""}
        for r in self.records:
            if r["packing_list"].casefold() != packing_list_key:
                continue
            if (r["ritm"] or "").casefold() == ritm_key:
                ritm_letters[ritm_key] = r["letter"]
                break
        self._merge_matched_emails(
            list(matches), ritm_letters, packing_list, forced_ritms={ritm},
        )

    def _find_outlook_message(self, namespace, email: MatchedEmail):
        if email.entry_id:
            if email.store_id:
                try:
                    message = namespace.GetItemFromID(email.entry_id, email.store_id)
                    if message is not None:
                        return message
                except Exception:
                    pass
            try:
                message = namespace.GetItemFromID(email.entry_id)
                if message is not None:
                    return message
            except Exception:
                pass
        try:
            folders = [namespace.GetDefaultFolder(6)]
        except Exception:
            return None
        while folders:
            folder = folders.pop()
            try:
                if folder.Name.casefold() == email.folder.casefold():
                    for message in folder.Items:
                        try:
                            if str(message.Subject or "") != email.subject:
                                continue
                            if email.sender_email and str(
                                message.SenderEmailAddress or ""
                            ).casefold() != email.sender_email.casefold():
                                continue
                            if not email.sender_email and email.sender and str(
                                message.SenderName or ""
                            ) != email.sender:
                                continue
                            if email.received and str(message.ReceivedTime) != email.received:
                                continue
                            return message
                        except Exception:
                            continue
                folders.extend(folder.Folders)
            except Exception:
                continue
        return None

    # ------------------------------------------------------------------
    # Scan pipeline
    # ------------------------------------------------------------------

    def _start_scan(self) -> None:
        if self.scan_worker is not None and self.scan_worker.isRunning():
            return
        record = self._build_current_record()
        if record is None:
            return
        search_terms = list(dict.fromkeys(
            item["description"] for item in record["items"]
            if item["description"].strip()
        ))
        if not search_terms:
            QMessageBox.warning(
                self, "No item descriptions",
                "Enter at least one item description to scan.",
            )
            return
        _APP_LOGGER.info(
            "Outlook scan requested: packing_list=%s date=%s item_count=%d "
            "time_range=%s database_ignored=%s",
            record.get("packing_list", ""),
            record.get("date", ""),
            len(record.get("items") or []),
            self.range_selector.currentData(),
            self.ignore_db_checkbox.isChecked(),
        )
        self.pending_record = record
        self._set_processing_state(False)
        self.status_label.setText(
            f"Searching Outlook for {len(search_terms)} item description(s)..."
        )
        self.scan_results_label.setText("Scanning attachments for exact item matches...")
        self.scan_worker = ScannerWorker(self.range_selector.currentData(), search_terms)
        self.scan_worker.scan_complete.connect(self._scan_finished)
        self.scan_worker.scan_failed.connect(self._scan_failed)
        self.scan_worker.start()


    def _confirm_near_matches(
        self, near_matches: list[NearMatch],
    ) -> set[tuple[str, str, str, str]]:
        accepted: set[tuple[str, str, str, str]] = set()
        seen: set[tuple[str, str, str, str]] = set()
        candidates: list[NearMatch] = []
        for candidate in near_matches:
            key = self._near_match_key(candidate)
            if key in seen:
                continue
            seen.add(key)
            candidates.append(candidate)
        if not candidates:
            return accepted
        total = len(candidates)
        for index, candidate in enumerate(candidates, start=1):
            dialog = NearMatchConfirmDialog(self, candidate, index, total)
            dialog.exec()
            if dialog.accepted:
                accepted.add(self._near_match_key(candidate))
        return accepted

    @staticmethod
    def _near_match_key(candidate: NearMatch) -> tuple[str, str, str, str]:
        source_key = candidate.entry_id or (
            f"{candidate.subject}|{candidate.received}"
        )
        return (
            candidate.term.casefold(), candidate.folder.casefold(),
            candidate.filename.casefold(), source_key,
        )

    @staticmethod
    def _parse_received(text: str) -> datetime:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                continue
        return datetime.min

    @staticmethod
    def _email_key(match: ExactMatch) -> str:
        return match.entry_id or f"{match.folder}|{match.subject}|{match.received}|{match.filename}"

    def _resolve_resent_emails(
        self, subject_matches: list[ExactMatch],
    ) -> tuple[list[ExactMatch], set[str]]:
        groups: dict[tuple[str, str], list[ExactMatch]] = {}
        for match in subject_matches:
            for ritm in match.subject_ritms:
                key = (match.subject.casefold(), ritm.casefold())
                groups.setdefault(key, []).append(match)
        skipped_ritms: set[str] = set()
        keep_ids: set[str] = set()
        for (subject_key, ritm_key), group in groups.items():
            unique_sources: dict[str, list[ExactMatch]] = {}
            for match in group:
                unique_sources.setdefault(self._email_key(match), []).append(match)
            if len(unique_sources) < 2:
                continue
            sorted_sources = sorted(
                unique_sources.values(),
                key=lambda ms: self._parse_received(ms[0].received),
                reverse=True,
            )
            newer_email = sorted_sources[0][0]
            older_email = sorted_sources[1][0]
            dialog = ResendPickerDialog(
                self, ritm_key.upper(), older_email, newer_email
            )
            dialog.exec()
            if dialog.chosen is None:
                skipped_ritms.add(ritm_key)
            else:
                keep_ids.add(self._email_key(dialog.chosen))
        filtered: list[ExactMatch] = []
        for match in subject_matches:
            match_ritm_keys = [r.casefold() for r in match.subject_ritms]
            if any(r in skipped_ritms for r in match_ritm_keys):
                continue
            kept = False
            for ritm_key in match_ritm_keys:
                group = groups.get((match.subject.casefold(), ritm_key), [])
                unique_srcs = {self._email_key(m) for m in group}
                if len(unique_srcs) < 2:
                    kept = True
                    break
                if self._email_key(match) in keep_ids:
                    kept = True
                    break
            if kept:
                filtered.append(match)
        return filtered, skipped_ritms

    def _scan_finished(self, results: ScanResults) -> None:
        record = self.pending_record
        if record is None:
            self._set_processing_state(True)
            return
        record_items = record.get("items") or []
        time_range = self.range_selector.currentData()
        search_terms_for_rescan = list(dict.fromkeys(
            item.get("description", "")
            for item in record_items
            if item.get("description", "").strip()
        ))
        subject_matches = [
            match for match in results.exact_matches if match.subject_ritms
        ]
        accepted_near_keys = self._confirm_near_matches(results.near_matches)
        accepted_near_matches = [
            candidate for candidate in results.near_matches
            if self._near_match_key(candidate) in accepted_near_keys
        ]
        subject_matches.extend(
            ExactMatch(
                term=candidate.term,
                filename=candidate.filename,
                folder=candidate.folder,
                subject=candidate.subject,
                received=candidate.received,
                sender=candidate.sender,
                sender_email=candidate.sender_email,
                entry_id=candidate.entry_id,
                store_id=candidate.store_id,
                ritms=candidate.subject_ritms,
                subject_ritms=candidate.subject_ritms,
                location=candidate.location,
                location_address=candidate.location,
                is_near_match=True,
                attachment_item_id=candidate.guess,
            )
            for candidate in accepted_near_matches
            if candidate.subject_ritms
        )
        email_display_matches = list(subject_matches)
        if not subject_matches:
            dialog = AddressBatchDialog(
                self, record_items, [], record,
                subject_matches=[], time_range=time_range,
                search_terms=search_terms_for_rescan,
            )
            dialog.rescan_found_sources.connect(self._on_rescan_sources)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                self.pending_record = None
                self._set_processing_state(True)
                self.status_label.setText("Scan finished. Nothing queued.")
                return
            self._commit_dialog_state(record, dialog, [])
            self.pending_record = None
            self._set_processing_state(True)
            return

        subject_matches, skipped_ritms = self._resolve_resent_emails(subject_matches)
        if not subject_matches:
            self.pending_record = None
            self._set_processing_state(True)
            self.scan_results_label.setText(
                "All matched RITMs were skipped. Nothing queued."
            )
            self.status_label.setText("Scan finished with no items queued.")
            return

        found_ritms = sorted({
            ritm for match in subject_matches for ritm in match.subject_ritms
            if ritm.casefold() not in skipped_ritms
        })
        exact_summary = sorted({
            f"{match.term}: {match.folder} / {match.filename}"
            for match in subject_matches
            if not match.is_near_match
        })
        item_ritm_candidates: dict[int, set[str]] = {}
        found_ritms_by_key = {ritm.casefold(): ritm for ritm in found_ritms}
        for index, item in enumerate(record_items):
            description_key = item.get("description", "").casefold()
            candidates = {
                ritm
                for match in subject_matches
                if match.term.casefold() == description_key
                for ritm in match.subject_ritms
            }
            candidates.update(
                ritm
                for candidate in results.near_matches
                if candidate.term.casefold() == description_key
                and self._near_match_key(candidate) in accepted_near_keys
                for ritm in candidate.subject_ritms
            )
            item_ritm_candidates[index] = {
                found_ritms_by_key[ritm.casefold()]
                for ritm in candidates
                if ritm.casefold() not in skipped_ritms
                and ritm.casefold() in found_ritms_by_key
            }

        duplicate_pairs: set[tuple[str, str]] = set()
        if db_configured() and not self.ignore_db_checkbox.isChecked():
            pending_for_dup_check: list[PackingListRecord] = []
            for ritm in found_ritms:
                probe_items = [
                    item
                    for index, item in enumerate(record_items)
                    if ritm in item_ritm_candidates[index]
                ]
                if not probe_items:
                    continue
                pending_for_dup_check.append({
                    **record, "ritm": ritm, "letter": "", "items": probe_items,
                })
            dupes = record_db_duplicates(pending_for_dup_check)
            current_base = _base_frn(_packing_list_identifier(record)).casefold()
            for ritm, description, existing_frn in dupes:
                if (
                    self.repeat_load_checkbox.isChecked()
                    and _base_frn(existing_frn).casefold() == current_base
                ):
                    continue
                duplicate_pairs.add((ritm.casefold(), description.casefold()))

        item_ritm_assignment: dict[int, str] = {}
        skipped_item_indexes: set[int] = set()
        for index, item in enumerate(record_items):
            candidates = sorted(
                (
                    ritm
                    for ritm in item_ritm_candidates[index]
                    if (
                        ritm.casefold(),
                        item.get("description", "").casefold(),
                    ) not in duplicate_pairs
                ),
                key=str.casefold,
            )
            if not candidates:
                if item_ritm_candidates[index]:
                    skipped_item_indexes.add(index)
                continue
            if len(candidates) == 1:
                item_ritm_assignment[index] = candidates[0]
                continue
            picker = ItemRitmPickerDialog(
                self, item.get("description", ""), candidates
            )
            if (
                picker.exec() == QDialog.DialogCode.Accepted
                and picker.chosen is not None
            ):
                item_ritm_assignment[index] = picker.chosen
            else:
                skipped_item_indexes.add(index)

        assigned_item_indexes: set[int] = set()
        matched_email_count = len({
            match.entry_id or (match.folder, match.subject, match.received)
            for match in subject_matches
        })
        total_accepted_files = 0
        staged_ritm_records: list[tuple[PackingListRecord, list[ExactMatch]]] = []
        staged: list[PackingListRecord] = []
        for ritm in found_ritms:
            ritm_key = ritm.casefold()
            ritm_matches = [
                match for match in subject_matches
                if ritm in match.subject_ritms
            ]
            if not ritm_matches:
                continue
            current_matched_terms = {match.term.casefold() for match in ritm_matches}
            exact_files = {
                f"{match.folder} / {match.filename}"
                for match in ritm_matches if not match.is_near_match
            }
            accepted_files: set[str] = set()
            accepted_terms = set()
            for candidate in results.near_matches:
                candidate_key = self._near_match_key(candidate)
                if candidate_key not in accepted_near_keys:
                    continue
                if ritm not in candidate.subject_ritms:
                    continue
                accepted_files.add(
                    f"{candidate.folder} / {candidate.filename} "
                    f"(accepted near match: {candidate.guess} for {candidate.term})"
                )
                accepted_terms.add(candidate.term.casefold())
                total_accepted_files += 1
            group_terms = sorted({match.term for match in ritm_matches} | {
                candidate.term for candidate in results.near_matches
                if candidate.term.casefold() in accepted_terms
            })
            matched_item_indexes = [
                index for index, item in enumerate(record_items)
                if item_ritm_assignment.get(index) == ritm
                and index not in assigned_item_indexes
                and (
                    item.get("description", "").casefold()
                    in current_matched_terms
                    or item.get("description", "").casefold() in accepted_terms
                )
            ]
            matched_items = [record_items[index] for index in matched_item_indexes]
            if not matched_items:
                continue
            assigned_item_indexes.update(matched_item_indexes)
            group = ScannerGroup(
                ritm_key,
                f"{ritm.upper()} "
                f"({len(exact_files)} exact, {len(accepted_files)} similar accepted)",
                group_terms,
                sorted(exact_files | accepted_files),
                sorted(exact_files),
                results.near_matches,
            )
            self.groups[group.key] = group
            matched_location_address = next(
                (m.location_address for m in ritm_matches if m.location_address),
                "",
            )
            matched_location = next(
                (m.location for m in ritm_matches if m.location), ""
            )
            ritm_record: PackingListRecord = {
                **record,
                "ritm": ritm, "letter": "", "group": ritm_key,
                "location": matched_location or record.get("location", ""),
                "location_address": matched_location_address or record.get("location_address", ""),
                "items": [{**item, "_ritm": ritm} for item in matched_items],
                "source_matches": list(ritm_matches),
            }
            staged_ritm_records.append((ritm_record, ritm_matches))
            staged.append(ritm_record)

        unmatched_items = [
            item for index, item in enumerate(record_items)
            if index not in assigned_item_indexes
            and index not in skipped_item_indexes
        ]
        dialog = AddressBatchDialog(
            self, unmatched_items, staged, record,
            subject_matches=subject_matches, time_range=time_range,
            search_terms=search_terms_for_rescan,
        )
        dialog.rescan_found_sources.connect(self._on_rescan_sources)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.pending_record = None
            self._set_processing_state(True)
            self.status_label.setText("Scan cancelled.")
            return

        self._commit_dialog_state(record, dialog, staged_ritm_records)
        queued_packing_lists = [
            _packing_list_identifier(r)
            for r in self.records
            if r["packing_list"].casefold() == record["packing_list"].casefold()
        ]
        self.pending_record = None
        self._set_processing_state(True)
        ritm_letters: dict[str, str] = {}
        for r in self.records:
            if r["packing_list"].casefold() == record["packing_list"].casefold():
                if not _is_no_request_item(r.get("ritm")):
                    ritm_letters[r["ritm"].casefold()] = r["letter"]
        self._merge_matched_emails(
            [
                match for match in email_display_matches
                if not any(
                    ritm.casefold() in skipped_ritms
                    for ritm in match.subject_ritms
                )
            ],
            ritm_letters,
            record["packing_list"],
        )
        self.scan_results_label.setText("Finished Scanning")
        self.status_label.setText(
            f"Processed {len(queued_packing_lists)} packing list(s)."
        )

    def _commit_dialog_state(
        self,
        template_record: PackingListRecord,
        dialog: AddressBatchDialog,
        staged_ritm_records: list[tuple[PackingListRecord, list[ExactMatch]]],
    ) -> None:
        input_identifiers = {
            _packing_list_identifier(rec) for rec, _ in staged_ritm_records
        }
        output_identifiers = {
            _packing_list_identifier(r) for r in dialog.staged_ritms
        }
        dropped_identifiers = input_identifiers - output_identifiers
        if dropped_identifiers:
            for record in list(self.records):
                identifier = _packing_list_identifier(record)
                if identifier not in dropped_identifiers:
                    continue
                if (record["packing_list"].casefold()
                        != template_record["packing_list"].casefold()):
                    continue
                try:
                    self.records.remove(record)
                except ValueError:
                    pass
                self.record_source_matches.pop(identifier, None)
                for dir_key in self.exported_record_ids_by_directory:
                    self.exported_record_ids_by_directory[dir_key].discard(
                        identifier.casefold()
                    )
        for ritm_record in dialog.staged_ritms:
            items = ritm_record["items"]
            if not items:
                continue
            packing_list_key = ritm_record["packing_list"].casefold()
            ritm_key = _ritm_group_key(ritm_record.get("ritm"))
            existing = next(
                (
                    r for r in self.records
                    if r["packing_list"].casefold() == packing_list_key
                    and _ritm_group_key(r.get("ritm")) == ritm_key
                    and (
                        ritm_key != ""
                        or r.get("location_address", "").strip() == ritm_record.get("location_address", "").strip()
                    )
                ),
                None,
            )
            if existing is not None:
                existing["items"].extend(items)
                if ritm_record.get("location_address"):
                    existing["location_address"] = ritm_record["location_address"]
                    existing["location"] = ritm_record.get("location", "")
                elif ritm_record.get("location") and not existing.get("location"):
                    existing["location"] = ritm_record["location"]
                current = list(existing.get("source_matches") or [])
                seen = {(m.entry_id, m.filename, m.term) for m in current}
                for match in _record_source_matches(ritm_record):
                    key = (match.entry_id, match.filename, match.term)
                    if key in seen:
                        continue
                    seen.add(key)
                    current.append(match)
                existing["source_matches"] = current
            else:
                self._append_record(ritm_record)
        if dialog.unmatched_items:
            self._build_unmatched_records(
                template_record, dialog.unmatched_items, staged_ritm_records,
            )
        self._assign_letters_by_priority(template_record["packing_list"].casefold())
        self._refresh_queue_table()

    def _assign_letters_by_priority(self, packing_list_key: str) -> None:
        this_list = [
            r for r in self.records
            if r["packing_list"].casefold() == packing_list_key
        ]
        if not this_list:
            return
        if len(this_list) == 1:
            this_list[0]["letter"] = ""
            return

        def key(r: PackingListRecord):
            has_ritm = not _is_no_request_item(r.get("ritm"))
            has_loc = bool(r["location"].strip() or r["location_address"].strip())
            if has_ritm and has_loc:
                tier = 0
            elif has_ritm:
                tier = 1
            elif has_loc:
                tier = 2
            else:
                tier = 3
            return (tier, -len(r["items"]), r["ritm"].casefold(), r["group"])

        this_list.sort(key=key)
        for i, r in enumerate(this_list):
            number = i + 1
            letters = ""
            while number:
                number, rem = divmod(number - 1, 26)
                letters = chr(ord("A") + rem) + letters
            r["letter"] = letters

    def _scan_failed(self, error: str) -> None:
        _APP_LOGGER.error("Outlook scan failed: %s", error)
        self.pending_record = None
        self._set_processing_state(True)
        self.status_label.setText("Scan did not complete")
        self.scan_results_label.setText(
            "The current packing list is still on the form and was not queued."
        )
        QMessageBox.critical(self, "Scanner failed", error)

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def _pick_source_match(
        self, identifier: str, matches: list[ExactMatch],
    ) -> ExactMatch | None:
        if not matches:
            return None
        by_source: dict[tuple[str, str], list[ExactMatch]] = {}
        for match in matches:
            key = (match.entry_id, match.store_id)
            by_source.setdefault(key, []).append(match)
        if len(by_source) == 1:
            return matches[0]
        dialog = SourceMatchPickerDialog(self, identifier, matches)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.chosen is not None:
            return dialog.chosen
        return None

    def _save_records(self) -> None:
        if not self.records:
            QMessageBox.information(
                self, "Nothing to save",
                "Add at least one list to the queue first.",
            )
            return

        records_missing_separator = []
        for record in self.records:
            raw = record["driver_plate"].strip()
            if not raw:
                continue
            trucker, trucker_id = _split_trucker_and_id(raw)
            if not trucker or not trucker_id:
                records_missing_separator.append(_packing_list_identifier(record))
        if records_missing_separator:
            QMessageBox.warning(
                self,
                "Trucker and ID missing",
                "Enter both a Trucker name and a Trucker ID in the Input Data "
                "form so both can be read. Check these list(s): "
                f"{', '.join(records_missing_separator)}.",
            )
            return

        last_export_folder = (
            self.app_config.get("last_export_folder")
            or str(self.summaries_folder.parent
                   if self.summaries_folder.parent.is_dir()
                   else DEFAULT_SUMMARIES_DIR.parent)
        )
        if not Path(last_export_folder).is_dir():
            last_export_folder = str(DEFAULT_SUMMARIES_DIR.parent)

        selected_directory = QFileDialog.getExistingDirectory(
            self,
            "Choose where to save packing lists, per-load folders, and the "
            "Reel Summary Report",
            last_export_folder,
        )
        if not selected_directory:
            return
        self.app_config["last_export_folder"] = selected_directory
        save_app_config(self.app_config)

        destination_directory = Path(selected_directory)
        directory_key = str(destination_directory.resolve()).casefold()
        exported_record_ids = self.exported_record_ids_by_directory.get(directory_key, set())

        self.summaries_folder = destination_directory / "Summaries"
        self.summaries_folder.mkdir(parents=True, exist_ok=True)
        self.summaries_label.setText(
            f"{self.summaries_folder}  (auto-set from the Export Folder)"
        )
        self.app_config["summaries_folder"] = str(self.summaries_folder)
        save_app_config(self.app_config)

        records_to_export = [
            record for record in self.records
            if _packing_list_identifier(record).casefold() not in exported_record_ids
        ]
        packing_list_paths: list[Path] = []
        per_load_folders: list[Path] = []
        scrap_paths: list[Path] = []
        for record in records_to_export:
            identifier = _packing_list_identifier(record)
            safe = re.sub(
                r"[^A-Za-z0-9._-]+", "_", record["packing_list"]
            ).strip("._-") or "PackingList"
            if record["letter"]:
                file_suffix = f"{safe}-{record['letter']}"
            else:
                file_suffix = safe
            packing_list_paths.append(
                destination_directory / f"Packing List {file_suffix}.xls"
            )
            load_folder = destination_directory / identifier
            per_load_folders.append(load_folder)
            scrap_paths.append(
                load_folder / f"Scrap Pick Up Request for Load {identifier}.xls"
            )
        existing_destinations: list[Path] = []
        for record, pl_path, sp_path in zip(records_to_export, packing_list_paths, scrap_paths):
            if pl_path.exists():
                existing_destinations.append(pl_path)
            if sp_path.exists():
                existing_destinations.append(sp_path)
        for record, load_folder in zip(records_to_export, per_load_folders):
            identifier = _packing_list_identifier(record)
            req_pdf = load_folder / f"Request Item for Load {identifier}.pdf"
            if req_pdf.exists():
                existing_destinations.append(req_pdf)
        if existing_destinations:
            existing_names = "\n".join(path.name for path in existing_destinations)
            answer = QMessageBox.question(
                self,
                "Replace existing files?",
                f"These files already exist:\n{existing_names}\n\nReplace them?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        records_by_week: dict[tuple[date, date], list[PackingListRecord]] = {}
        frn_date_warnings: list[str] = []
        for record in self.records:
            frn_date, reason = _date_from_frn(record["packing_list"])
            if frn_date is None:
                frn_date_warnings.append(
                    f"{_packing_list_identifier(record)}: "
                    f"could not read a valid date from the FRN ({reason}); "
                    f"using the typed date {record['date']} instead."
                )
                try:
                    frn_date = parse_date(record["date"])
                except Exception as error:
                    QMessageBox.critical(
                        self, "Could not read a record's date",
                        f"{_packing_list_identifier(record)}: {error}",
                    )
                    return
            try:
                monday, friday = week_range(frn_date.isoformat())
            except Exception as error:
                QMessageBox.critical(
                    self, "Could not read a record's date",
                    f"{_packing_list_identifier(record)}: {error}",
                )
                return
            records_by_week.setdefault((monday, friday), []).append(record)

        pdf_failures: list[str] = []
        pdf_skipped_no_source: list[str] = []
        shared_excel = None
        try:
            shared_excel = win32com.client.DispatchEx("Excel.Application")
            _tune_excel_for_batch(shared_excel)
            for (monday, friday), week_records in records_by_week.items():
                filename = summary_filename(monday, friday)
                weekly_path = self.summaries_folder / filename
                write_reel_summary_report(week_records, weekly_path, excel=shared_excel)
            for record, pl_path, load_folder, sp_path in zip(
                records_to_export, packing_list_paths, per_load_folders, scrap_paths
            ):
                identifier = _packing_list_identifier(record)
                write_packing_list_template(record, pl_path, excel=shared_excel)
                load_folder.mkdir(parents=True, exist_ok=True)
                write_scrap_pickup_template(record, sp_path, excel=shared_excel)
                matches = self.record_source_matches.get(identifier)
                if matches is None:
                    matches = _record_source_matches(record)
                chosen = self._pick_source_match(identifier, matches)
                if chosen is None:
                    pdf_skipped_no_source.append(identifier)
                    continue
                pdf_bytes = fetch_attachment_bytes(
                    chosen.entry_id, chosen.store_id, chosen.filename
                )
                if pdf_bytes is None:
                    pdf_failures.append(
                        f"{identifier}: could not re-open the source email"
                    )
                    continue
                two_page_bytes = slice_first_pages(pdf_bytes, 2)
                if two_page_bytes is None:
                    pdf_failures.append(
                        f"{identifier}: could not read the attachment as a PDF"
                    )
                    continue
                annotated = annotate_pdf(two_page_bytes, f"Delivery {identifier}")
                if annotated is None:
                    pdf_failures.append(
                        f"{identifier}: could not annotate the PDF"
                    )
                    continue
                req_pdf_path = load_folder / f"Request Item for Load {identifier}.pdf"
                req_pdf_path.write_bytes(annotated)
        except Exception as error:
            _APP_LOGGER.exception("Packing-list export failed.")
            QMessageBox.critical(self, "Could not save files", str(error))
            return
        finally:
            _release_excel(shared_excel, None)

        db_summary = ""
        if not db_configured():
            db_summary = "Database: skipped (no database configured)."
        elif self.ignore_db_checkbox.isChecked():
            db_summary = "Database: skipped (Ignore database checked)."
        else:
            if self.repeat_load_checkbox.isChecked():
                deleted_total = 0
                deleted_bases: set[str] = set()
                for record in records_to_export:
                    base = _base_frn(_packing_list_identifier(record))
                    if base in deleted_bases:
                        continue
                    deleted_bases.add(base)
                    deleted_total += delete_base_frn_from_db(
                        _packing_list_identifier(record)
                    )
            total_inserted = 0
            total_skipped = 0
            records_to_export.sort(key=lambda r: parse_date(r["date"]))
            for record in records_to_export:
                inserted, skipped = insert_record_to_db(record)
                total_inserted += inserted
                total_skipped += skipped
            if self.repeat_load_checkbox.isChecked():
                db_summary = (
                    f"Database: replaced "
                    f"{total_inserted} item row(s)"
                    + (f", skipped {total_skipped} duplicate(s)."
                       if total_skipped else ".")
                )
            else:
                db_summary = (
                    f"Database: inserted {total_inserted} item row(s)"
                    + (f", skipped {total_skipped} duplicate(s)."
                       if total_skipped else ".")
                )
        exported_record_ids.update(
            _packing_list_identifier(record).casefold() for record in records_to_export
        )
        self.exported_record_ids_by_directory[directory_key] = exported_record_ids
        count = len(records_to_export)
        week_count = len(records_by_week)
        self.status_label.setText(
            f"Saved {count} packing-list workbook(s), "
            f"{count} scrap-pickup workbook(s), "
            f"{week_count} weekly Reel Summary Report(s), and "
            f"per-load folders in {destination_directory}"
        )
        info_lines = [
            f"Saved {count} packing-list workbook(s) (top-level)",
            db_summary,
            f"Created {count} per-load folder(s) with Scrap Pickup"
            + (f" + {count - len(pdf_skipped_no_source)} Request Item PDF(s)" if pdf_skipped_no_source else " + Request Item PDF"),
            f"Wrote {week_count} weekly Reel Summary Report(s) to "
            f"{self.summaries_folder}",
        ]
        if frn_date_warnings:
            info_lines.append("")
            info_lines.append("FRN date warnings:")
            info_lines.extend(f"  - {msg}" for msg in frn_date_warnings)
        if pdf_skipped_no_source:
            info_lines.append("")
            info_lines.append(
                "Request Item PDFs skipped (no source email on file):"
            )
            info_lines.extend(f"  - {msg}" for msg in pdf_skipped_no_source)
        if pdf_failures:
            info_lines.append("")
            info_lines.append("PDF warnings:")
            info_lines.extend(f"  - {msg}" for msg in pdf_failures)
        QMessageBox.information(
            self, "Workbooks saved",
            "\n".join(info_lines) + f"\n\nLocation:\n{destination_directory}",
        )


if __name__ == "__main__":
    application = QApplication(sys.argv)
    window = PackingListApp()
    window.show()
    sys.exit(application.exec())