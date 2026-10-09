import re
import os
import io
import json
import html
import subprocess
import sys
import tempfile
import shutil
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, TypedDict

from PySide6.QtCore import QEvent, QThread, Qt, Signal
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QStyle,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
)
import win32com.client

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


# ---------------------------------------------------------------------------
# Config / paths
# ---------------------------------------------------------------------------

CONFIG_PATH = Path(__file__).resolve().parent / "scanner_config.json"
DEFAULT_SUMMARIES_DIR = Path(__file__).resolve().parent / "Summaries"

TEMPLATE_RELATIVE_PATH = Path("Excel Bases") / "Incoming Packing List.xls"
WORKING_DIR_RELATIVE_PATH = Path("Working")
INPUT_SHEET_NAME = "Packing List"

IN_DATE = "E8"
IN_PLATE = "C9"
IN_RECEIVED = "E9"
IN_CONT = "C10"
IN_PACKING_LIST = "E10"

IN_FIRST_ITEM_ROW = 13
IN_LAST_ITEM_ROW = 35
IN_DESC_COL = "B"
IN_MATERIAL_COL = "D"
IN_GROSS_COL = "E"
IN_TARE_COL = "F"

MAX_ITEMS_PER_LOAD = 23
MAX_ITEMS_PER_SUMMARY_TAB = 31

RENAME_EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
RENAME_PDF_SUFFIXES = {".pdf"}


def load_app_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


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
    common_types = {
        "99 fo": "99 FO",
        "99fo": "99 FO",
        "23": "23",
        "type 23": "Type 23",
        "type23": "Type 23",
        "type 3": "Type 3",
        "type3": "Type 3",
        "type 4": "Type 4",
        "type4": "Type 4",
        "69s": "69S",
    }
    return common_types.get(cleaned.casefold(), cleaned.upper())


def normalize_ritm(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value).upper()


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


# FRN date parsing (e.g. 'FRN260929-1R' -> date(2026, 9, 29)).
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
# Excel helpers (merge-safe)
# ---------------------------------------------------------------------------

def _safe_write(sheet, address: str, value) -> None:
    """Write to a cell, targeting the merge anchor if the cell is merged.
    Prevents Excel's 'we can't do that to a merged cell' error."""
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
    """Clear a block of cells, skipping any cell that's part of a merge
    to avoid Excel complaining."""
    for r in range(first_row, last_row + 1):
        for c in range(first_col, last_col + 1):
            cell = sheet.Cells(r, c)
            try:
                if cell.MergeCells:
                    continue
            except Exception:
                pass
            cell.ClearContents()


# ---------------------------------------------------------------------------
# PDF helpers
# ---------------------------------------------------------------------------

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


def fetch_attachment_bytes(
    entry_id: str,
    store_id: str,
    filename: str,
) -> bytes | None:
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
# File-renaming helpers
# ---------------------------------------------------------------------------

_FRN_IN_TEXT = re.compile(
    r"(FRN\d{6}-\d+[A-Z]?)(?:-([A-Za-z]+))?",
    re.IGNORECASE,
)


def _extract_frn_from_text(text: str) -> tuple[str, str]:
    if not text:
        return "", ""
    match = _FRN_IN_TEXT.search(text)
    if not match:
        return "", ""
    frn = match.group(1).upper()
    letter = (match.group(2) or "").upper()
    return frn, letter


def _extract_frn_from_excel(path: Path, excel=None) -> tuple[str, str]:
    """
    Try to pull an FRN from a workbook. If `excel` is provided, reuse that
    Excel.Application (much faster when scanning a folder); otherwise
    spin up a temporary one.
    """
    owns_excel = excel is None
    workbook = None
    try:
        if owns_excel:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=True
        )

        candidates: list[str] = []
        for sheet in workbook.Worksheets:
            for address in ("E10", "H10", "M2", "E11", "C10"):
                value = sheet.Range(address).Value
                if value is None:
                    continue
                text = str(value).strip()
                if text:
                    candidates.append(text)

        for text in candidates:
            frn, letter = _extract_frn_from_text(text)
            if frn:
                return frn, letter
        return "", ""
    except Exception:
        return "", ""
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if owns_excel and excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def _extract_frn_from_pdf(path: Path) -> tuple[str, str]:
    try:
        reader = PdfReader(str(path))
        if not reader.pages:
            return "", ""
        text = reader.pages[0].extract_text() or ""
        frn, letter = _extract_frn_from_text(text)
        return frn, letter
    except Exception:
        return "", ""


def plan_file_renames(folder: Path, progress=None) -> list[dict]:
    """
    `progress` is an optional callable(current, total, filename) used to
    report progress to a UI.
    """
    entries = [
        e for e in sorted(folder.iterdir())
        if e.is_file()
        and not e.name.startswith(".")
        and not e.name.startswith("~$")
    ]
    total = len(entries)

    plans: list[dict] = []
    excel = None
    try:
        # Only spin up Excel if there's at least one spreadsheet to look at.
        if any(e.suffix.casefold() in RENAME_EXCEL_SUFFIXES for e in entries):
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

        for index, entry in enumerate(entries, start=1):
            if progress is not None:
                progress(index, total, entry.name)

            stem = entry.stem
            suffix = entry.suffix

            frn, letter = _extract_frn_from_text(stem)
            source = "filename" if frn else ""

            if not frn and suffix.casefold() in RENAME_EXCEL_SUFFIXES:
                frn, letter = _extract_frn_from_excel(entry, excel=excel)
                if frn:
                    source = "excel"

            if not frn and suffix.casefold() in RENAME_PDF_SUFFIXES:
                frn, letter = _extract_frn_from_pdf(entry)
                if frn:
                    source = "pdf"

            if not letter:
                tail_letter = re.search(r"[-_]([A-Za-z])$", stem)
                if tail_letter:
                    letter = tail_letter.group(1).upper()

            new_name = ""
            reason = ""

            if not frn:
                reason = "No FRN found in filename or file contents"
            else:
                new_stem = f"Incoming {frn}"
                if letter:
                    new_stem += f"-{letter}"
                new_name = new_stem + suffix
                if new_name == entry.name:
                    reason = "Already named correctly"
                else:
                    reason = f"Matched via {source}"

            plans.append({
                "path": entry,
                "old_name": entry.name,
                "new_name": new_name,
                "frn": frn,
                "letter": letter,
                "reason": reason,
            })
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
    return plans


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
    attachments: set[str] = field(default_factory=set)
    item_terms: set[str] = field(default_factory=set)


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


class ItemRecord(TypedDict):
    description: str
    type: str
    gross: int
    tare: int
    net: int


class PackingListRecord(TypedDict):
    group: str
    ritm: str
    letter: str
    date: str
    packing_list: str
    location: str
    location_address: str
    driver_plate: str
    items: list[ItemRecord]


class BlankZeroSpinBox(QSpinBox):
    def textFromValue(self, value: int) -> str:
        return "" if value == self.minimum() else str(value)

    def valueFromText(self, text: str) -> int:
        cleaned = text.strip().replace(",", "")
        if not cleaned:
            return self.minimum()
        try:
            return int(cleaned)
        except ValueError:
            return self.minimum()


# ---------------------------------------------------------------------------
# Address / RITM batch dialog
# ---------------------------------------------------------------------------

class AddressBatchDialog(QDialog):
    def __init__(self, parent, items: list[ItemRecord]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Enter location / Request Item (optional)")
        self.setMinimumWidth(600)

        self._items: list[ItemRecord] = list(items)

        self.address_assignments: dict[int, str] = {}
        self.ritm_assignments: dict[int, str] = {}

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "These items were queued without a matching RITM or location.\n"
            "Check the IDs you want to assign, fill in the field(s), then click "
            "Apply. Checked items are removed from all three tabs once applied."
        ))

        self.tabs = QTabWidget()

        self.lists: list[QListWidget] = []
        self._build_tab("Location", self._build_location_tab)
        self._build_tab("Request Item", self._build_ritm_tab)
        self._build_tab("Both", self._build_both_tab)

        layout.addWidget(self.tabs)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #53645e;")
        layout.addWidget(self.status_label)

    def _build_tab(self, title: str, builder) -> None:
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)

        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        for index, item in enumerate(self._items):
            entry = QListWidgetItem(item["description"])
            entry.setFlags(entry.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            entry.setCheckState(Qt.CheckState.Unchecked)
            entry.setData(Qt.ItemDataRole.UserRole, index)
            list_widget.addItem(entry)
        self.lists.append(list_widget)
        tab_layout.addWidget(list_widget)

        builder(tab_layout, list_widget)

        self.tabs.addTab(tab, title)

    def _build_location_tab(self, tab_layout, list_widget) -> None:
        row = QHBoxLayout()
        row.addWidget(QLabel("Location:"))
        self.location_input = QLineEdit()
        self.location_input.setPlaceholderText(
            "Street, City, State ZIP  (leave blank for No Location)"
        )
        row.addWidget(self.location_input, 1)
        tab_layout.addLayout(row)
        self._add_apply_button(tab_layout, "apply_location")

    def _build_ritm_tab(self, tab_layout, list_widget) -> None:
        row = QHBoxLayout()
        row.addWidget(QLabel("Request Item:"))
        self.ritm_input = QLineEdit()
        self.ritm_input.setPlaceholderText("e.g. RITM0011836256")
        row.addWidget(self.ritm_input, 1)
        tab_layout.addLayout(row)
        self._add_apply_button(tab_layout, "apply_ritm")

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

        self._add_apply_button(tab_layout, "apply_both")

    def _add_apply_button(self, tab_layout, mode: str) -> None:
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

    def _checked_indices(self, list_widget: QListWidget) -> list[int]:
        indices: list[int] = []
        for row in range(list_widget.count()):
            entry = list_widget.item(row)
            if entry.checkState() == Qt.CheckState.Checked:
                indices.append(int(entry.data(Qt.ItemDataRole.UserRole)))
        return indices

    def _remove_from_all_lists(self, indices: set[int]) -> None:
        for list_widget in self.lists:
            row = 0
            while row < list_widget.count():
                entry = list_widget.item(row)
                if int(entry.data(Qt.ItemDataRole.UserRole)) in indices:
                    list_widget.takeItem(row)
                else:
                    row += 1

    def _remaining_count(self) -> int:
        return self.lists[0].count() if self.lists else 0

    def _apply(self, mode: str) -> None:
        current_tab = self.tabs.currentIndex()
        list_widget = self.lists[current_tab]
        indices = self._checked_indices(list_widget)

        if not indices:
            QMessageBox.information(
                self, "Nothing selected",
                "Check at least one item ID before applying.",
            )
            return

        if mode == "apply_location":
            value = self.location_input.text().strip()
            for idx in indices:
                self.address_assignments[idx] = value
            self.location_input.clear()
        elif mode == "apply_ritm":
            value = self.ritm_input.text().strip()
            if not value:
                QMessageBox.information(
                    self, "No RITM entered",
                    "Type a Request Item number before applying.",
                )
                return
            for idx in indices:
                self.ritm_assignments[idx] = value
            self.ritm_input.clear()
        elif mode == "apply_both":
            location = self.both_location_input.text().strip()
            ritm = self.both_ritm_input.text().strip()
            if not location and not ritm:
                QMessageBox.information(
                    self, "Nothing to apply",
                    "Enter a location and/or a Request Item before applying.",
                )
                return
            for idx in indices:
                if location:
                    self.address_assignments[idx] = location
                if ritm:
                    self.ritm_assignments[idx] = ritm
            self.both_location_input.clear()
            self.both_ritm_input.clear()

        self._remove_from_all_lists(set(indices))
        remaining = self._remaining_count()
        if remaining == 0:
            self.status_label.setText(
                "All items processed. Click Finished to close."
            )
        else:
            self.status_label.setText(
                f"Applied to {len(indices)} item(s). {remaining} remaining."
            )

    def _another_batch(self) -> None:
        self.location_input.clear()
        self.ritm_input.clear()
        self.both_location_input.clear()
        self.both_ritm_input.clear()
        self.status_label.setText(
            "New batch started. Check items and enter the next values."
        )


# ---------------------------------------------------------------------------
# Near-match confirmation dialog
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
            f"File: {candidate.folder} / {candidate.filename}\n"
            f"Email: {candidate.subject}\n"
            f"Received: {candidate.received}"
        )
        detail.setStyleSheet("color: #53645e; padding-top: 6px;")
        layout.addWidget(detail)

        layout.addWidget(QLabel("\nAre these the same item?"))

        buttons = QDialogButtonBox()
        same_button = buttons.addButton(
            "Same", QDialogButtonBox.ButtonRole.AcceptRole
        )
        not_same_button = buttons.addButton(
            "Not the same", QDialogButtonBox.ButtonRole.RejectRole
        )
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


# ---------------------------------------------------------------------------
# Resend picker dialog
# ---------------------------------------------------------------------------

class ResendPickerDialog(QDialog):
    def __init__(
        self,
        parent,
        ritm: str,
        older: ExactMatch,
        newer: ExactMatch,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Duplicate RITM detected: {ritm}")
        self.setMinimumWidth(1000)
        self.chosen: ExactMatch | None = None
        self.older = older
        self.newer = newer

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"RITM {ritm} was found in two emails. "
            "The newer email is preselected. Review both and pick one, "
            "or choose 'Skip this RITM' to leave it out of the queue."
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
        use_btn = buttons.addButton(
            "Use selected", QDialogButtonBox.ButtonRole.AcceptRole
        )
        skip_btn = buttons.addButton(
            "Skip this RITM", QDialogButtonBox.ButtonRole.RejectRole
        )
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


# ---------------------------------------------------------------------------
# Rename preview dialog
# ---------------------------------------------------------------------------

class RenamePreviewDialog(QDialog):
    def __init__(self, parent, plans: list[dict]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Rename Files")
        self.setMinimumWidth(820)
        self.setMinimumHeight(520)

        self.plans = plans
        self.checkboxes: list[QCheckBox] = []

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Uncheck any file you don't want renamed, then click Apply."
        ))

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(6, 6, 6, 6)

        for plan in plans:
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 2, 0, 2)

            box = QCheckBox()
            can_rename = bool(plan["new_name"]) and plan["new_name"] != plan["old_name"]
            box.setChecked(can_rename)
            box.setEnabled(can_rename)
            self.checkboxes.append(box)
            row_layout.addWidget(box)

            if can_rename:
                label_text = (
                    f"{plan['old_name']}\n"
                    f"    -> {plan['new_name']}\n"
                    f"    ({plan['reason']})"
                )
            else:
                label_text = (
                    f"{plan['old_name']}\n"
                    f"    -> (skipped: {plan['reason']})"
                )
            label = QLabel(label_text)
            if not can_rename:
                label.setStyleSheet("color: #9baea6;")
            row_layout.addWidget(label, 1)

            container_layout.addWidget(row)

        container_layout.addStretch(1)
        scroll.setWidget(container)
        layout.addWidget(scroll, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Apply")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_plans(self) -> list[dict]:
        chosen: list[dict] = []
        for box, plan in zip(self.checkboxes, self.plans):
            if box.isChecked() and plan["new_name"]:
                chosen.append(plan)
        return chosen


# ---------------------------------------------------------------------------
# Source match picker dialog
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Input-sheet parsing / template handling
# ---------------------------------------------------------------------------

def template_path() -> Path:
    return Path(__file__).resolve().parent / TEMPLATE_RELATIVE_PATH


def working_dir() -> Path:
    return Path(__file__).resolve().parent / WORKING_DIR_RELATIVE_PATH


def copy_template_to_working() -> Path:
    src = template_path()
    if not src.exists():
        raise FileNotFoundError(
            f"Input template not found: {src}\n"
            f"Put 'Incoming Packing List.xls' in the 'Excel Bases' folder."
        )
    dst_dir = working_dir()
    dst_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    dst = dst_dir / f"Incoming Packing List {timestamp}.xls"

    counter = 1
    while dst.exists():
        dst = dst_dir / f"Incoming Packing List {timestamp} ({counter}).xls"
        counter += 1

    shutil.copy2(src, dst)
    return dst


def _cell_str(sheet, address: str) -> str:
    value = sheet.Range(address).Value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _to_int(value: object) -> int:
    if not isinstance(value, (str, int, float)):
        return 0
    try:
        return int(round(float(value)))
    except (TypeError, ValueError, OverflowError):
        return 0


def _normalize_excel_date(value: object) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip() if value is not None else ""
    if not text:
        raise ValueError(f"The Date cell ({IN_DATE}) is empty.")
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    raise ValueError(
        f"Could not read the Date cell ({IN_DATE}): {text!r}. "
        "Use YYYY-MM-DD or MM/DD/YYYY."
    )


def parse_incoming_packing_list(path: str | Path) -> PackingListRecord:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input sheet not found: {path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        try:
            sheet = workbook.Worksheets(INPUT_SHEET_NAME)
        except Exception:
            sheet = workbook.Worksheets(1)

        date_iso = _normalize_excel_date(sheet.Range(IN_DATE).Value)
        packing_list = _cell_str(sheet, IN_PACKING_LIST)
        if not packing_list:
            raise ValueError(
                f"The Packing List cell ({IN_PACKING_LIST}) is empty. "
                "Fill in something like FRN260914-1R."
            )
        plate = _cell_str(sheet, IN_PLATE)
        cont = _cell_str(sheet, IN_CONT)

        items: list[ItemRecord] = []
        for row in range(IN_FIRST_ITEM_ROW, IN_LAST_ITEM_ROW + 1):
            description = str(
                sheet.Range(f"{IN_DESC_COL}{row}").Value or ""
            ).strip()
            if not description:
                continue
            material = str(
                sheet.Range(f"{IN_MATERIAL_COL}{row}").Value or ""
            ).strip()
            gross = _to_int(sheet.Range(f"{IN_GROSS_COL}{row}").Value)
            tare = _to_int(sheet.Range(f"{IN_TARE_COL}{row}").Value)
            items.append({
                "description": description,
                "type": normalize_item_type(material),
                "gross": gross,
                "tare": tare,
                "net": gross - tare,
            })

        if not items:
            raise ValueError(
                f"No item rows were found in "
                f"{IN_DESC_COL}{IN_FIRST_ITEM_ROW}:"
                f"{IN_DESC_COL}{IN_LAST_ITEM_ROW} (column {IN_DESC_COL} "
                "is blank everywhere)."
            )

        driver_plate = plate
        if cont:
            driver_plate = f"{plate} / {cont}" if plate else cont

        return {
            "group": "",
            "ritm": "",
            "letter": "",
            "date": date_iso,
            "packing_list": packing_list,
            "location": "",
            "location_address": "",
            "driver_plate": driver_plate,
            "items": items,
        }
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


# ---------------------------------------------------------------------------
# Output template writers
# ---------------------------------------------------------------------------

def write_packing_list_template(record: PackingListRecord, destination: Path) -> Path:
    template = Path(__file__).resolve().parent / "Excel Bases" / "packing list example.xls"
    if not template.exists():
        raise FileNotFoundError(f"Packing-list template not found: {template}")
    if len(record["items"]) > MAX_ITEMS_PER_LOAD:
        raise ValueError(
            "The packing-list template has 23 item rows; split this RITM "
            "into another letter."
        )

    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(destination), UpdateLinks=0, ReadOnly=False)
        sheet = workbook.Worksheets("Packing List")

        _safe_write(
            sheet, "H8",
            datetime.combine(parse_date(record["date"]), datetime.min.time()),
        )
        sheet.Range("H8").NumberFormat = "mm/dd/yyyy"
        _safe_write(sheet, "H9", record["location"].strip() or "No Location")
        _safe_write(sheet, "H10", _packing_list_identifier(record))

        for row, item in enumerate(record["items"], start=13):
            description = item["description"]
            if item["type"]:
                description = f"{description} ({item['type']})"
            _safe_write(sheet, f"B{row}", description)
            _safe_write(sheet, f"H{row}", item["gross"])
            _safe_write(sheet, f"I{row}", item["tare"])
            _safe_write(sheet, f"J{row}", item["net"])

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    return destination


def write_scrap_pickup_template(record: PackingListRecord, destination: Path) -> Path:
    template = Path(__file__).resolve().parent / "Excel Bases" / "Scrap Pick Up Request for Load.xls"
    if not template.exists():
        raise FileNotFoundError(f"Scrap pickup template not found: {template}")

    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(destination), UpdateLinks=0, ReadOnly=False)
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

        _safe_write(sheet, "F16", record["ritm"] or "No Request Item")
        _safe_write(sheet, "D19", street)
        _safe_write(sheet, "D20", city)
        _safe_write(sheet, "D34", _packing_list_identifier(record))
        _safe_write(sheet, "D35", record["driver_plate"].replace(" / ", "# "))
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    return destination


# ---------------------------------------------------------------------------
# Helpers used by the reel summary
# ---------------------------------------------------------------------------

def _packing_list_identifier(record: PackingListRecord) -> str:
    letter = (record.get("letter") or "").strip()
    if letter:
        return f"{record['packing_list']}-{letter}"
    return record["packing_list"]


def _split_driver_and_plate(value: str) -> tuple[str, str]:
    if "/" not in value:
        return value.strip(), ""
    driver, plate = value.split("/", 1)
    return driver.strip(), plate.strip()


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
        receiving_date,
        reel_order,
        letter,
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

    items: list[ItemRecord] = []
    for row in range(12, 43):
        reel_id = str(sheet.Cells(row, 4).Value or "").strip()
        if not reel_id:
            continue
        reel_size = str(sheet.Cells(row, 5).Value or "").strip()
        gross = int(float(sheet.Cells(row, 7).Value or 0))
        tare = int(float(sheet.Cells(row, 8).Value or 0))
        items.append({
            "description": f"{reel_id}{reel_size}",
            "type": str(sheet.Cells(row, 6).Value or "").strip(),
            "gross": gross,
            "tare": tare,
            "net": gross - tare,
        })

    trucker = str(sheet.Range("G8").Value or "").strip()
    trailer = str(sheet.Range("G9").Value or "").strip()
    return {
        "group": "",
        "ritm": str(sheet.Range("M5").Value or "").strip(),
        "letter": letter,
        "date": normalized_date,
        "packing_list": packing_list,
        "location": str(sheet.Range("C7").Value or "").strip(),
        "location_address": "",
        "driver_plate": f"{trucker} / {trailer}" if trucker or trailer else "",
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


def _copy_template_reel_tabs(
    workbook,
    template_workbook,
    count: int,
) -> list[Any]:
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





def write_reel_summary_report(
    records: list[PackingListRecord],
    destination: Path,
) -> Path:
    template = (
        Path(__file__).resolve().parent / "Excel Bases" / "Reel Summary Report.XLS"
    )
    if not template.exists():
        raise FileNotFoundError(f"Reel Summary template not found: {template}")

    destination.parent.mkdir(parents=True, exist_ok=True)
    source_path = destination if destination.exists() else template
    descriptor, staged_name = tempfile.mkstemp(
        prefix=".reel-summary-", suffix=destination.suffix or ".xls", dir=destination.parent
    )
    os.close(descriptor)
    staged_path = Path(staged_name)
    try:
        shutil.copy2(source_path, staged_path)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = None
        template_workbook = None
        try:
            workbook = excel.Workbooks.Open(str(staged_path), UpdateLinks=0, ReadOnly=False)

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
                template_workbook = excel.Workbooks.Open(str(template), UpdateLinks=0, ReadOnly=True)
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
                _safe_write(sheet, "M5", record["ritm"])
                _safe_write(sheet, "C7", record["location"].strip() or "No Location")
                _safe_write(sheet, "I7", sum(item["net"] for item in record["items"]))
                _safe_write(sheet, "C8", "N/A")
                _safe_write(sheet, "F8", driver)
                _safe_write(sheet, "C9", "N/A")
                _safe_write(sheet, "F9", plate)

                for row, item in enumerate(record["items"], start=12):
                    reel_id, reel_size = _split_reel_id(item["description"])
                    _safe_write(sheet, f"D{row}", reel_id)
                    _safe_write(sheet, f"E{row}", reel_size)
                    _safe_write(sheet, f"F{row}", item["type"])
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
            excel.CalculateFull()
            workbook.Save()
        finally:
            if template_workbook is not None:
                template_workbook.Close(SaveChanges=False)
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()

        os.replace(staged_path, destination)
    finally:
        if staged_path.exists():
            staged_path.unlink()
    return destination


# ---------------------------------------------------------------------------
# Scanner worker
# ---------------------------------------------------------------------------

class ScannerWorker(QThread):
    scan_complete = Signal(object)
    scan_failed = Signal(str)

    def __init__(self, time_range: str, search_terms: list[str]) -> None:
        super().__init__()
        self.time_range = time_range
        self.search_terms = list(dict.fromkeys(
            term.strip() for term in search_terms if term.strip()
        ))

    def run(self) -> None:
        scanner_path = Path(__file__).resolve().with_name("improvedScanner.py")
        try:
            with tempfile.TemporaryDirectory(prefix="packing-list-scan-") as temp_dir:
                report_path = Path(temp_dir) / "scan_results.txt"
                results_path = Path(temp_dir) / "scan_results.json"
                environment = os.environ.copy()
                environment["SCAN_RESULTS_OUTPUT"] = str(report_path)
                environment["SCAN_RESULTS_JSON"] = str(results_path)
                environment["SCAN_SEARCH_TERMS"] = "\n".join(self.search_terms)
                result = subprocess.run(
                    [sys.executable, str(scanner_path)],
                    cwd=scanner_path.parent,
                    input=f"{self.time_range}\n\n",
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=900,
                    env=environment,
                )
                if result.returncode != 0:
                    details = result.stderr.strip() or result.stdout.strip()
                    raise RuntimeError(details or "The scanner exited before completing.")
                results = json.loads(results_path.read_text(encoding="utf-8"))
                exact_matches = [
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
                near_entries: dict[tuple[str, str, str], NearMatch] = {}
                for entry in results["near_matches"]:
                    near_match = NearMatch(
                        term=entry["term"],
                        guess=entry["guess"],
                        similarity=entry["similarity"],
                        filename=entry["filename"],
                        folder=entry["folder"],
                        subject=entry["subject"],
                        received=entry["received"],
                        subject_ritms=entry.get("subject_ritms", []),
                    )
                    key = (
                        near_match.term.casefold(),
                        near_match.folder.casefold(),
                        near_match.filename.casefold(),
                    )
                    existing_match = near_entries.get(key)
                    if existing_match is None or near_match.similarity > existing_match.similarity:
                        near_entries[key] = near_match

                exact_matches.sort(key=lambda m: (
                    m.received,
                    m.folder.casefold(),
                    m.subject.casefold(),
                    m.filename.casefold(),
                    m.term.casefold(),
                ))
                near_list = list(near_entries.values())
                near_list.sort(key=lambda m: (
                    m.received,
                    m.folder.casefold(),
                    m.subject.casefold(),
                    m.filename.casefold(),
                    m.term.casefold(),
                    -m.similarity,
                ))
                self.scan_complete.emit(ScanResults(exact_matches, near_list))
        except Exception as error:
            self.scan_failed.emit(str(error))


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class PackingListApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Scanner Packing Lists")
        self.resize(1180, 900)
        self.setMinimumSize(920, 720)

        self.app_config = load_app_config()
        self.summaries_folder = self._resolve_summaries_folder()

        self.records: list[PackingListRecord] = []
        self.exported_record_ids_by_directory: dict[str, set[str]] = {}
        self.current_items: list[ItemRecord] = []
        self.groups: dict[str, ScannerGroup] = {}
        self.matched_emails_by_key: dict[str, MatchedEmail] = {}
        self.visible_email_matches: list[MatchedEmail] = []
        self.scan_worker: ScannerWorker | None = None
        self.pending_record: PackingListRecord | None = None
        self.imported_record: PackingListRecord | None = None
        self.imported_source_path: Path | None = None
        self.record_source_matches: dict[str, list[ExactMatch]] = {}
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

    # ------------------------------------------------------------------
    # Summaries folder
    # ------------------------------------------------------------------

    def _resolve_summaries_folder(self) -> Path:
        configured = self.app_config.get("summaries_folder")
        if configured:
            p = Path(configured)
            if p.is_dir():
                return p
        DEFAULT_SUMMARIES_DIR.mkdir(parents=True, exist_ok=True)
        return DEFAULT_SUMMARIES_DIR

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
            "Open the input sheet, fill it in Excel, then Import. "
            "Scan afterward to match RITMs."
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

        input_row = QHBoxLayout()
        self.open_sheet_button = QPushButton("Open Input Sheet")
        self.open_sheet_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.open_sheet_button.clicked.connect(self._open_input_sheet)
        input_row.addWidget(self.open_sheet_button)

        self.import_sheet_button = QPushButton("Import Input Sheet")
        self.import_sheet_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.import_sheet_button.clicked.connect(self._import_input_sheet)
        input_row.addWidget(self.import_sheet_button)

        self.rename_files_button = QPushButton("Rename Files")
        self.rename_files_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.rename_files_button.clicked.connect(self._rename_files)
        input_row.addWidget(self.rename_files_button)

        self.next_load_button = QPushButton("Next Load")
        self.next_load_button.clicked.connect(self._next_load)
        input_row.addWidget(self.next_load_button)

        input_row.addStretch(1)
        root.addLayout(input_row)

        self.imported_info_label = QLabel("No input sheet imported yet.")
        self.imported_info_label.setWordWrap(True)
        self.imported_info_label.setStyleSheet("color: #53645e; margin-bottom: 8px;")
        root.addWidget(self.imported_info_label)

        toggle_row = QHBoxLayout()
        self.manual_toggle_button = QPushButton("Show Manual Entry")
        self.manual_toggle_button.setCheckable(True)
        self.manual_toggle_button.toggled.connect(self._toggle_manual_entry)
        toggle_row.addWidget(self.manual_toggle_button)
        toggle_row.addStretch(1)
        root.addLayout(toggle_row)

        self.manual_panel = QWidget()
        manual_root = QVBoxLayout(self.manual_panel)
        manual_root.setContentsMargins(0, 0, 0, 0)

        form = QHBoxLayout()
        self.date_input = self._add_field(form, "Date", date.today().isoformat())
        self.packing_list_input = self._add_field(form, "Packing List #")
        self.driver_plate_input = self._add_field(form, "Driver / Plate #")
        self.driver_plate_input.setPlaceholderText("e.g. Jane Doe / ABC123")
        manual_root.addLayout(form)

        item_form = QHBoxLayout()
        self.description_input = self._add_field(item_form, "Item Description")
        self.type_input = self._add_field(item_form, "Type")
        self.type_input.setPlaceholderText("e.g. 99 FO, 23, Type 3, Type 4, 69S")
        self.gross_input = self._add_weight_field(item_form, "Gross")
        self.tare_input = self._add_weight_field(item_form, "Tare")
        self.item_input_fields = [
            self.description_input,
            self.type_input,
            self.gross_input,
            self.tare_input,
        ]
        for field in self.item_input_fields:
            field.installEventFilter(self)
            if isinstance(field, QSpinBox):
                field.lineEdit().installEventFilter(self)
        self.net_preview = self._add_field(item_form, "Net", "0")
        self.net_preview.setReadOnly(True)
        next_field_button = QToolButton()
        next_field_button.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowRight)
        )
        next_field_button.setToolTip("Move to the next item field")
        next_field_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        next_field_button.clicked.connect(self._focus_next_item_field)
        item_form.addWidget(next_field_button, 0, Qt.AlignmentFlag.AlignBottom)
        self.description_input.returnPressed.connect(self._focus_next_item_field)
        self.type_input.returnPressed.connect(self._focus_next_item_field)
        manual_root.addLayout(item_form)
        self.gross_input.valueChanged.connect(self._update_totals)
        self.tare_input.valueChanged.connect(self._update_totals)

        item_actions = QHBoxLayout()
        self.add_item_button = QPushButton("Add Item")
        self.add_item_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.add_item_button.clicked.connect(self._add_item)
        item_actions.addWidget(self.add_item_button)
        self.delete_item_button = QPushButton("Delete Item")
        self.delete_item_button.clicked.connect(self._remove_selected_item)
        item_actions.addWidget(self.delete_item_button)
        item_actions.addStretch(1)
        manual_root.addLayout(item_actions)

        self.items_table = QTableWidget(0, 5)
        self.items_table.setHorizontalHeaderLabels(
            ["Item Description", "Type", "Gross", "Tare", "Net"]
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
        self.items_table.setMaximumHeight(190)
        manual_root.addWidget(self.items_table)

        totals_row = QHBoxLayout()
        self.gross_total_label = QLabel("Gross total: 0.00")
        self.tare_total_label = QLabel("Tare total: 0.00")
        self.net_total_label = QLabel("Net total: 0.00")
        for total_label in (
            self.gross_total_label,
            self.tare_total_label,
            self.net_total_label,
        ):
            total_label.setStyleSheet("font-weight: bold; color: #20332e;")
            totals_row.addWidget(total_label)
        totals_row.addStretch(1)
        manual_root.addLayout(totals_row)

        self.manual_panel.setVisible(False)
        root.addWidget(self.manual_panel)

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

        self.email_table = QTableWidget(0, 6)
        self.email_table.setHorizontalHeaderLabels([
            "RITM", "Sender", "Subject", "Received",
            "Outlook Folder", "Matched Attachment(s)",
        ])
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
        root.addWidget(self.email_table)
        email_actions = QHBoxLayout()
        self.reply_button = QPushButton("Reply to Selected Email")
        self.reply_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.reply_button.setEnabled(False)
        self.reply_button.clicked.connect(self._reply_to_selected_email)
        email_actions.addWidget(self.reply_button)
        email_actions.addStretch(1)
        root.addLayout(email_actions)

        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels([
            "Packing List ID", "RITM", "Date", "Packing List #",
            "Driver / Plate", "Items", "Gross Total", "Tare Total", "Net Total",
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
        root.addLayout(footer)

    @staticmethod
    def _add_weight_field(layout: QHBoxLayout, label: str) -> QSpinBox:
        field = QVBoxLayout()
        field.addWidget(QLabel(label))
        value = BlankZeroSpinBox()
        value.setRange(0, 1_000_000_000)
        value.setSingleStep(1)
        field.addWidget(value)
        layout.addLayout(field, 1)
        return value

    @staticmethod
    def _add_field(layout: QHBoxLayout, label: str, value: str = "") -> QLineEdit:
        field = QVBoxLayout()
        field.addWidget(QLabel(label))
        entry = QLineEdit(value)
        field.addWidget(entry)
        layout.addLayout(field, 1)
        return entry

    # ------------------------------------------------------------------
    # Input-sheet buttons
    # ------------------------------------------------------------------

    def _open_input_sheet(self) -> None:
        try:
            working = copy_template_to_working()
        except Exception as error:
            QMessageBox.critical(self, "Could not open input sheet", str(error))
            return
        try:
            os.startfile(str(working))
        except AttributeError:
            QMessageBox.information(
                self,
                "Open input sheet",
                f"Edit this file in Excel:\n{working}",
            )
        self.status_label.setText(
            f"Opened {working.name}. Fill it in, save, close Excel, then "
            "click Import Input Sheet."
        )
        self.imported_info_label.setText(
            f"Working file: {working}\n"
            "Edit in Excel, save, close, then click Import Input Sheet."
        )

    def _import_input_sheet(self) -> None:
        default_dir = working_dir()
        if not default_dir.exists():
            default_dir = Path(__file__).resolve().parent

        chosen, _ = QFileDialog.getOpenFileName(
            self,
            "Select the filled-in Incoming Packing List",
            str(default_dir),
            "Excel 97-2003 (*.xls);;All files (*)",
        )
        if not chosen:
            return
        chosen_path = Path(chosen)

        try:
            record = parse_incoming_packing_list(chosen_path)
        except Exception as error:
            QMessageBox.critical(self, "Could not read input sheet", str(error))
            return

        if self.imported_record is not None:
            answer = QMessageBox.question(
                self,
                "Replace pending import?",
                "There is already an imported input sheet waiting to be scanned:\n"
                f"{self.imported_record['packing_list']} "
                f"({len(self.imported_record['items'])} item(s)).\n\n"
                "Replace it with the newly imported data?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        self.imported_record = record
        self.imported_source_path = chosen_path
        self.pending_record = record

        gross_total = sum(item["gross"] for item in record["items"])
        tare_total = sum(item["tare"] for item in record["items"])
        net_total = sum(item["net"] for item in record["items"])
        self.imported_info_label.setText(
            f"Imported: {chosen_path.name}\n"
            f"Packing List: {record['packing_list']}\n"
            f"Date: {record['date']}\n"
            f"Driver / Plate: {record['driver_plate'] or '(none)'}\n"
            f"Items: {len(record['items'])}  |  "
            f"Gross {gross_total:,}  |  "
            f"Tare {tare_total:,}  |  "
            f"Net {net_total:,}"
        )
        self.status_label.setText(
            f"Imported {record['packing_list']}. Click Scan Item List to "
            "match it against Outlook."
        )

    def _rename_files(self) -> None:
        default_dir = working_dir()
        if not default_dir.exists():
            default_dir = Path(__file__).resolve().parent

        chosen = QFileDialog.getExistingDirectory(
            self,
            "Choose the folder whose files should be renamed",
            str(default_dir),
        )
        if not chosen:
            return
        folder = Path(chosen)

        self.status_label.setText(f"Scanning {folder} for FRN numbers...")
        QApplication.processEvents()

        try:
            plans = plan_file_renames(folder)
        except Exception as error:
            QMessageBox.critical(self, "Could not scan folder", str(error))
            self.status_label.setText("Rename cancelled.")
            return

        if not plans:
            QMessageBox.information(
                self, "Nothing to do",
                "That folder has no files to rename.",
            )
            self.status_label.setText("Rename cancelled.")
            return

        dialog = RenamePreviewDialog(self, plans)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.status_label.setText("Rename cancelled.")
            return

        selected = dialog.selected_plans()
        if not selected:
            self.status_label.setText("No files selected for renaming.")
            return

        renamed = 0
        failures: list[str] = []
        for plan in selected:
            src: Path = plan["path"]
            dst = src.with_name(plan["new_name"])
            if dst.exists() and dst != src:
                failures.append(
                    f"{src.name}: '{plan['new_name']}' already exists"
                )
                continue
            try:
                src.rename(dst)
                renamed += 1
            except OSError as error:
                failures.append(f"{src.name}: {error}")

        message_lines = [f"Renamed {renamed} file(s)."]
        if failures:
            message_lines.append("")
            message_lines.append("Warnings:")
            message_lines.extend(f"  - {line}" for line in failures)
        QMessageBox.information(
            self, "Rename complete", "\n".join(message_lines)
        )
        self.status_label.setText(f"Renamed {renamed} file(s) in {folder}")

    def _next_load(self) -> None:
        if self.imported_record is not None:
            answer = QMessageBox.question(
                self,
                "Clear current import?",
                f"Discard the imported sheet for "
                f"{self.imported_record['packing_list']} without scanning it?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.imported_record = None
        self.imported_source_path = None
        self.pending_record = None
        self.imported_info_label.setText("No input sheet imported yet.")
        self.status_label.setText("Ready for the next load.")
        self.scan_results_label.setText(
            "Matching emails and attachments will appear here."
        )

    def _toggle_manual_entry(self, checked: bool) -> None:
        self.manual_panel.setVisible(checked)
        self.manual_toggle_button.setText(
            "Hide Manual Entry" if checked else "Show Manual Entry"
        )

    # ------------------------------------------------------------------
    # Manual entry helpers
    # ------------------------------------------------------------------

    def _focused_item_field_index(self) -> int | None:
        focused = QApplication.focusWidget()
        for index, field in enumerate(self.item_input_fields):
            if focused is field:
                return index
            if isinstance(field, QSpinBox) and focused is field.lineEdit():
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
        numeric_field = isinstance(self.item_input_fields[index], QSpinBox)
        if isinstance(watched, QLineEdit) and not numeric_field:
            if watched.hasSelectedText():
                return super().eventFilter(watched, event)
            cursor = watched.cursorPosition()
            if direction > 0 and cursor < len(watched.text()):
                return super().eventFilter(watched, event)
            if direction < 0 and cursor > 0:
                return super().eventFilter(watched, event)

        self._focus_item_field(index + direction)
        return True

    def _add_item(self) -> None:
        description = self.description_input.text().strip()
        item_type = normalize_item_type(self.type_input.text())
        gross = self.gross_input.value()
        tare = self.tare_input.value()
        if not description or not item_type:
            QMessageBox.warning(
                self, "Missing item details", "Enter an item description and type."
            )
            return
        if gross <= 0:
            QMessageBox.warning(
                self, "Check gross weight", "Gross weight must be greater than zero."
            )
            return
        if tare > gross:
            QMessageBox.warning(
                self, "Check tare weight", "Tare cannot be greater than gross."
            )
            return

        item: ItemRecord = {
            "description": description,
            "type": item_type,
            "gross": gross,
            "tare": tare,
            "net": gross - tare,
        }
        self.current_items.append(item)
        row = self.items_table.rowCount()
        self.items_table.insertRow(row)
        values = (
            description,
            item_type,
            f"{gross:,}",
            f"{tare:,}",
            f"{item['net']:,}",
        )
        for column, value in enumerate(values):
            self.items_table.setItem(row, column, QTableWidgetItem(value))
        self.description_input.clear()
        self.type_input.clear()
        self.gross_input.setValue(0)
        self.tare_input.setValue(0)
        self.description_input.setFocus()
        self._update_totals()

    def _remove_selected_item(self) -> None:
        indexes = sorted(
            {index.row() for index in self.items_table.selectionModel().selectedRows()},
            reverse=True,
        )
        if not indexes:
            self.status_label.setText("Select one or more current items to remove")
            return
        for index in indexes:
            self.current_items.pop(index)
            self.items_table.removeRow(index)
        self._update_totals()

    def _update_totals(self) -> None:
        entered_gross = self.gross_input.value()
        entered_tare = self.tare_input.value()
        entered_net = entered_gross - entered_tare
        self.net_preview.setText(f"{entered_net:,}")

        gross_total = sum(item["gross"] for item in self.current_items) + entered_gross
        tare_total = sum(item["tare"] for item in self.current_items) + entered_tare
        net_total = sum(item["net"] for item in self.current_items) + entered_net
        self.gross_total_label.setText(f"Gross total: {gross_total:,}")
        self.tare_total_label.setText(f"Tare total: {tare_total:,}")
        self.net_total_label.setText(f"Net total: {net_total:,}")

    def _build_current_record(self) -> PackingListRecord | None:
        if self.imported_record is not None:
            return self.imported_record

        raw_date = self.date_input.text().strip()
        packing_list = self.packing_list_input.text().strip()
        if not raw_date or not packing_list:
            QMessageBox.warning(
                self,
                "Missing information",
                "Fill in the date and packing-list number, or import an input sheet.",
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
                self,
                "No items",
                "Add at least one item, or import an input sheet.",
            )
            return None
        record: PackingListRecord = {
            "group": "",
            "ritm": "",
            "letter": "",
            "date": normalized_date,
            "packing_list": packing_list,
            "location": "",
            "location_address": "",
            "driver_plate": self.driver_plate_input.text().strip(),
            "items": self.current_items.copy(),
        }
        return record

    # ------------------------------------------------------------------
    # Queue / record helpers
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

    def _next_packing_list_letter(self, packing_list_key: str) -> str:
        taken: set[str] = set()
        for existing in self.records:
            if existing["packing_list"].casefold() == packing_list_key:
                if existing["letter"]:
                    taken.add(existing["letter"].upper())

        number = 1
        while True:
            candidate = ""
            n = number
            while n:
                n, rem = divmod(n - 1, 26)
                candidate = chr(ord("A") + rem) + candidate
            if candidate not in taken:
                return candidate
            number += 1

    def _location_from_address(self, address: str) -> str:
        parts = [p.strip() for p in address.split(",")]
        if len(parts) < 2:
            return ""
        city_state = ", ".join(parts[1:])
        city_state = re.sub(r"\s+\d{5}(?:-\d{4})?\s*$", "", city_state).strip()
        return city_state.replace(",", "")

    def _build_unmatched_records(
        self,
        record: PackingListRecord,
        items: list[ItemRecord],
        address_assignments: dict[int, str],
        ritm_assignments: dict[int, str],
        staged_ritm_records: list[tuple[PackingListRecord, list[ExactMatch]]],
    ) -> list[PackingListRecord]:
        packing_list_key = record["packing_list"].casefold()

        buckets: dict[tuple[str, str], list[ItemRecord]] = {}
        for idx, item in enumerate(items):
            addr = address_assignments.get(idx, "").strip()
            r = ritm_assignments.get(idx, "").strip()
            buckets.setdefault((addr, r), []).append(item)

        def existing_ritm_records() -> dict[str, PackingListRecord]:
            return {
                rec["ritm"].casefold(): rec
                for rec in self.records
                if rec["packing_list"].casefold() == packing_list_key
                and rec["ritm"]
                and rec["ritm"].casefold() != "no request item"
            }

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
                new_record: PackingListRecord = {
                    **record,
                    "group": chosen_ritm.casefold(),
                    "ritm": chosen_ritm,
                    "letter": "",
                    "location": self._location_from_address(address) if address else "",
                    "location_address": address,
                    "items": group_items,
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
            and rec["ritm"]
            and rec["ritm"].casefold() != "no request item"
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

    def _clear_current_list(self) -> None:
        self.packing_list_input.clear()
        self.driver_plate_input.clear()
        self.current_items.clear()
        self.items_table.setRowCount(0)
        self.gross_input.setValue(0)
        self.tare_input.setValue(0)
        self.description_input.setFocus()
        self._update_totals()

    def _set_processing_state(self, enabled: bool) -> None:
        for widget in (
            self.open_sheet_button,
            self.import_sheet_button,
            self.rename_files_button,
            self.next_load_button,
            self.manual_toggle_button,
            self.date_input,
            self.packing_list_input,
            self.driver_plate_input,
            self.description_input,
            self.type_input,
            self.gross_input,
            self.tare_input,
            self.add_item_button,
            self.delete_item_button,
            self.items_table,
            self.scan_button,
            self.export_folder_button,
        ):
            widget.setEnabled(enabled)

    # ------------------------------------------------------------------
    # Email matching / Outlook
    # ------------------------------------------------------------------

    def _merge_matched_emails(
        self,
        exact_matches: list[ExactMatch],
        fallback_ritms: list[str],
        ritm_letters: dict[str, str],
        packing_list: str,
    ) -> None:
        for match in exact_matches:
            message_key = match.entry_id or "|".join(
                (match.folder, match.subject, match.received)
            )
            email_ritms = match.ritms or fallback_ritms
            for ritm in email_ritms:
                ritm_key = ritm.casefold()
                letter = ritm_letters.get(ritm_key)
                if letter is None:
                    continue
                email_key = f"{message_key}|{ritm_key}|{packing_list.casefold()}|{letter}"
                email = self.matched_emails_by_key.get(email_key)
                if email is None:
                    email = MatchedEmail(
                        key=email_key,
                        ritm=ritm,
                        packing_list=packing_list,
                        letter=letter,
                        entry_id=match.entry_id,
                        store_id=match.store_id,
                        sender=match.sender,
                        sender_email=match.sender_email,
                        subject=match.subject,
                        received=match.received,
                        folder=match.folder,
                    )
                    self.matched_emails_by_key[email_key] = email
                email.attachments.add(match.filename)
                email.item_terms.add(match.term)

        self.visible_email_matches = sorted(
            self.matched_emails_by_key.values(),
            key=lambda email: (
                email.ritm.casefold(),
                email.received,
                email.subject.casefold(),
                email.folder.casefold(),
            ),
            reverse=True,
        )
        self.email_table.setRowCount(0)
        for email in self.visible_email_matches:
            row = self.email_table.rowCount()
            self.email_table.insertRow(row)
            sender = email.sender
            if email.sender_email and email.sender_email.casefold() not in sender.casefold():
                sender = f"{sender} <{email.sender_email}>".strip()
            values = (
                email.ritm,
                sender,
                email.subject,
                email.received,
                email.folder,
                ", ".join(sorted(email.attachments)),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, email.key)
                self.email_table.setItem(row, column, item)
        self._update_reply_button()

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

    def _update_reply_button(self) -> None:
        row = self.email_table.currentRow()
        if row < 0:
            self.reply_button.setEnabled(False)
            return
        ritm_item = self.email_table.item(row, 0)
        email_key = ritm_item.data(Qt.ItemDataRole.UserRole) if ritm_item else None
        if not isinstance(email_key, str):
            self.reply_button.setEnabled(False)
            return
        email = self.matched_emails_by_key.get(email_key)
        self.reply_button.setEnabled(bool(
            email and (email.entry_id or (email.subject and email.folder))
        ))

    def _reply_to_selected_email(self) -> None:
        row = self.email_table.currentRow()
        if row < 0:
            return
        ritm_item = self.email_table.item(row, 0)
        email_key = ritm_item.data(Qt.ItemDataRole.UserRole) if ritm_item else None
        if not isinstance(email_key, str):
            return
        email = self.matched_emails_by_key.get(email_key)
        if email is None:
            QMessageBox.warning(
                self,
                "Email cannot be opened",
                "The selected email could not be matched to its Outlook record.",
            )
            return

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        message = self._find_outlook_message(namespace, email)
        if message is None:
            QMessageBox.warning(
                self,
                "Email cannot be opened",
                "Outlook could not find the selected email by its ID or its "
                "subject, sender, date, and folder.",
            )
            return

        ritm_records = [
            record for record in self.records
            if record["group"] == email.ritm.casefold()
            and record["packing_list"].casefold() == email.packing_list.casefold()
            and record["letter"] == email.letter
        ]
        if ritm_records:
            item_descriptions = list(dict.fromkeys(
                item["description"]
                for record in ritm_records
                for item in record["items"]
            ))
        else:
            item_descriptions = sorted(email.item_terms)

        date_values = sorted({
            parse_date(record["date"]).strftime("%m/%d/%Y")
            for record in ritm_records
        })
        delivered_date = ", ".join(date_values) if date_values else ""

        item_lines = "<br>".join(
            html.escape(description) for description in item_descriptions
        )
        reply_body = (
            f"<p>Delivered {html.escape(delivered_date)}</p>"
            f"<p>{item_lines}</p>"
            "<p>Thank you</p><br>"
        )

        try:
            reply = message.ReplyAll()
            reply.HTMLBody = reply_body + reply.HTMLBody
            reply.Display()
        except Exception as error:
            QMessageBox.critical(self, "Could not prepare reply", str(error))

    # ------------------------------------------------------------------
    # Scan
    # ------------------------------------------------------------------

    def _start_scan(self) -> None:
        if self.scan_worker is not None and self.scan_worker.isRunning():
            return
        record = self._build_current_record()
        if record is None:
            return
        search_terms = list(dict.fromkeys(
            item["description"] for item in record["items"] if item["description"].strip()
        ))
        if not search_terms:
            QMessageBox.warning(
                self, "No item descriptions",
                "Enter at least one item description to scan.",
            )
            return
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
        self,
        near_matches: list[NearMatch],
    ) -> set[tuple[str, str, str]]:
        accepted: set[tuple[str, str, str]] = set()

        seen: set[tuple[str, str, str]] = set()
        candidates: list[NearMatch] = []
        for candidate in near_matches:
            key = (
                candidate.term.casefold(),
                candidate.folder.casefold(),
                candidate.filename.casefold(),
            )
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
                accepted.add((
                    candidate.term.casefold(),
                    candidate.folder.casefold(),
                    candidate.filename.casefold(),
                ))
        return accepted

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
        self,
        subject_matches: list[ExactMatch],
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

        subject_matches = [match for match in results.exact_matches if match.subject_ritms]
        if not subject_matches:
            dialog = AddressBatchDialog(self, record["items"])
            dialog.exec()
            new_records = self._build_unmatched_records(
                record,
                record["items"],
                dialog.address_assignments,
                dialog.ritm_assignments,
                [],
            )
            self.pending_record = None
            self.imported_record = None
            self.imported_source_path = None
            if not new_records:
                self._set_processing_state(True)
                self.scan_results_label.setText("There are no items to queue as unmatched.")
                self.status_label.setText("No items found to queue.")
                return
            # Single-load: no letter.
            self._assign_letters_by_priority(record["packing_list"].casefold())
            self._refresh_queue_table()
            queued_ids = ", ".join(
                _packing_list_identifier(r) for r in new_records
            )
            self._clear_current_list()
            self.imported_info_label.setText("No input sheet imported yet.")
            self._set_processing_state(True)
            self.scan_results_label.setText(
                "No exact item descriptions were found in the scanned attachments. "
                f"All items were queued as {queued_ids}."
            )
            self.status_label.setText("Queued all items as unmatched.")
            return

        accepted_near_keys = self._confirm_near_matches(results.near_matches)

        subject_matches, skipped_ritms = self._resolve_resent_emails(subject_matches)
        if not subject_matches:
            self.pending_record = None
            self.imported_record = None
            self.imported_source_path = None
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
        })

        assigned_item_indexes: set[int] = set()
        ritm_letters: dict[str, str] = {}
        matched_email_count = len({
            match.entry_id or (match.folder, match.subject, match.received)
            for match in subject_matches
        })
        total_accepted_files = 0
        staged_ritm_records: list[tuple[PackingListRecord, list[ExactMatch]]] = []

        for ritm in found_ritms:
            ritm_key = ritm.casefold()
            ritm_matches = [
                match for match in subject_matches
                if ritm in match.subject_ritms
            ]
            if not ritm_matches:
                continue

            current_matched_terms = {match.term.casefold() for match in ritm_matches}
            exact_files = {f"{match.folder} / {match.filename}" for match in ritm_matches}
            accepted_files: set[str] = set()
            accepted_terms = set()

            for candidate in results.near_matches:
                candidate_key = (
                    candidate.term.casefold(),
                    candidate.folder.casefold(),
                    candidate.filename.casefold(),
                )
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

            existing_group = self.groups.get(ritm_key)
            if existing_group is not None:
                exact_files.update(existing_group.exact_files)
                accepted_files.update(set(existing_group.files) - set(existing_group.exact_files))

            group_terms = sorted({match.term for match in ritm_matches} | {
                candidate.term for candidate in results.near_matches
                if candidate.term.casefold() in accepted_terms
            })
            matched_item_indexes = [
                index for index, item in enumerate(record["items"])
                if index not in assigned_item_indexes
                and (
                    item["description"].casefold() in current_matched_terms
                    or item["description"].casefold() in accepted_terms
                )
            ]
            matched_items = [record["items"][index] for index in matched_item_indexes]
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
            ritm_record: PackingListRecord = {
                **record,
                "ritm": ritm,
                "letter": "",
                "group": ritm_key,
                "location": next(
                    (match.location for match in ritm_matches if match.location), ""
                ) or record.get("location", ""),
                "location_address": next(
                    (match.location_address for match in ritm_matches if match.location_address),
                    "",
                ) or record.get("location_address", ""),
                "items": matched_items,
            }
            staged_ritm_records.append((ritm_record, ritm_matches))

        # Append records now, but defer record_source_matches keying until
        # after letters are assigned (below), since the identifier includes
        # the letter (or is the bare FRN for a single-load family).
        for ritm_record, _ in staged_ritm_records:
            self._append_record(ritm_record)

        unmatched_items = [
            item for index, item in enumerate(record["items"])
            if index not in assigned_item_indexes
        ]
        dialog_new_records: list[PackingListRecord] = []
        if unmatched_items:
            dialog = AddressBatchDialog(self, unmatched_items)
            dialog.exec()
            dialog_new_records = self._build_unmatched_records(
                record,
                unmatched_items,
                dialog.address_assignments,
                dialog.ritm_assignments,
                staged_ritm_records,
            )

        all_records_for_letters: list[PackingListRecord] = (
            [r for r, _ in staged_ritm_records] + dialog_new_records
        )
        for r in all_records_for_letters:
            if not r["letter"]:
                r["letter"] = self._next_packing_list_letter(
                    record["packing_list"].casefold()
                )
        self._assign_letters_by_priority(record["packing_list"].casefold())
        self._refresh_queue_table()

        # NOW the letters are final; key record_source_matches by identifier
        # so _save_records can find them and write the Request Item PDFs.
        for ritm_record, ritm_matches in staged_ritm_records:
            self.record_source_matches[_packing_list_identifier(ritm_record)] = list(ritm_matches)



        ritm_letters = {}
        for r in self.records:
            if r["packing_list"].casefold() == record["packing_list"].casefold():
                if r["ritm"] and r["ritm"].casefold() != "no request item":
                    ritm_letters[r["ritm"].casefold()] = r["letter"]

        self._merge_matched_emails(
            subject_matches,
            [],
            ritm_letters,
            record["packing_list"],
        )

        queued_packing_lists = [
            _packing_list_identifier(r)
            for r in self.records
            if r["packing_list"].casefold() == record["packing_list"].casefold()
        ]

        self.pending_record = None
        self.imported_record = None
        self.imported_source_path = None
        self._clear_current_list()
        self.imported_info_label.setText("No input sheet imported yet.")
        self._set_processing_state(True)
        self.scan_results_label.setText(
            f"Packing list IDs: {', '.join(queued_packing_lists)}\n"
            f"Exact item matches across {matched_email_count} email(s):\n"
            f"{chr(10).join(exact_summary)}\n"
            f"Accepted near-match attachments: {total_accepted_files}. "
            "Select an email below and choose Reply to open a draft."
        )
        self.status_label.setText(
            f"Processed packing list IDs {', '.join(queued_packing_lists)}."
        )

    def _assign_letters_by_priority(self, packing_list_key: str) -> None:
        this_list = [
            r for r in self.records
            if r["packing_list"].casefold() == packing_list_key
        ]

        # Single-load family: no letter at all — identifier stays bare FRN.
        if len(this_list) == 1:
            this_list[0]["letter"] = ""
            return

        def key(r: PackingListRecord):
            has_ritm = bool(r["ritm"]) and r["ritm"].casefold() != "no request item"
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
        self,
        identifier: str,
        matches: list[ExactMatch],
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

        records_missing_separator = [
            _packing_list_identifier(record)
            for record in self.records
            if record["driver_plate"].strip() and "/" not in record["driver_plate"]
        ]
        if records_missing_separator:
            QMessageBox.warning(
                self,
                "Separate driver and plate",
                "Enter Driver / Plate # with a slash between the two values "
                f"(for example, Jane Doe / ABC123) before exporting these list(s): "
                f"{', '.join(records_missing_separator)}.",
            )
            return

        selected_directory = QFileDialog.getExistingDirectory(
            self,
            "Choose where to save packing lists, per-load folders, and the "
            "Reel Summary Report",
            str(self.summaries_folder.parent if self.summaries_folder.parent.is_dir() else DEFAULT_SUMMARIES_DIR.parent),
        )
        if not selected_directory:
            return
        destination_directory = Path(selected_directory)
        directory_key = str(destination_directory.resolve()).casefold()
        exported_record_ids = self.exported_record_ids_by_directory.get(directory_key, set())

        # Summaries live in a 'Summaries' subfolder of the export folder.
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
            # Single-load family: file name is the bare FRN (no trailing dash,
            # no letter).
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
                        self,
                        "Could not read a record's date",
                        f"{_packing_list_identifier(record)}: {error}",
                    )
                    return
            try:
                monday, friday = week_range(frn_date.isoformat())
            except Exception as error:
                QMessageBox.critical(
                    self,
                    "Could not read a record's date",
                    f"{_packing_list_identifier(record)}: {error}",
                )
                return
            records_by_week.setdefault((monday, friday), []).append(record)

        pdf_failures: list[str] = []

        try:
            for (monday, friday), week_records in records_by_week.items():
                filename = summary_filename(monday, friday)
                weekly_path = self.summaries_folder / filename
                write_reel_summary_report(week_records, weekly_path)

            for record, pl_path, load_folder, sp_path in zip(
                records_to_export, packing_list_paths, per_load_folders, scrap_paths
            ):
                identifier = _packing_list_identifier(record)

                write_packing_list_template(record, pl_path)

                load_folder.mkdir(parents=True, exist_ok=True)
                write_scrap_pickup_template(record, sp_path)

                matches = self.record_source_matches.get(identifier, [])
                chosen = self._pick_source_match(identifier, matches)
                if chosen is None:
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
            QMessageBox.critical(self, "Could not save files", str(error))
            return

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
            f"Created {count} per-load folder(s) with Scrap Pickup + Request Item PDF",
            f"Wrote {week_count} weekly Reel Summary Report(s) to "
            f"{self.summaries_folder}",
        ]
        if frn_date_warnings:
            info_lines.append("")
            info_lines.append("FRN date warnings:")
            info_lines.extend(f"  - {msg}" for msg in frn_date_warnings)
        if pdf_failures:
            info_lines.append("")
            info_lines.append("PDF warnings:")
            info_lines.extend(f"  - {msg}" for msg in pdf_failures)
        QMessageBox.information(
            self,
            "Workbooks saved",
            "\n".join(info_lines) + f"\n\nLocation:\n{destination_directory}",
        )


if __name__ == "__main__":
    application = QApplication(sys.argv)
    window = PackingListApp()
    window.show()
    sys.exit(application.exec())