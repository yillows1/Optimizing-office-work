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
from pathlib import Path
from typing import Any, TypedDict

from PySide6.QtCore import QEvent, QThread, Qt, Signal
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
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
    QTableWidget,
    QTableWidgetItem,
    QSpinBox,
    QStyle,
    QToolButton,
    QVBoxLayout,
    QWidget,
)
import win32com.client

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


# ---------------------------------------------------------------------------
# Constants for the input template
# ---------------------------------------------------------------------------

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
# Address batch dialog
# ---------------------------------------------------------------------------

class AddressBatchDialog(QDialog):
    """Popup for assigning addresses to unmatched items."""

    def __init__(self, parent, items: list[ItemRecord]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Enter address (optional)")
        self.setMinimumWidth(520)
        self.assignments: dict[int, str] = {}

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "These items were queued as No Request Item / No Location.\n"
            "Check the IDs you want to give an address, type the address, "
            "then click Apply to selected."
        ))

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        for index, item in enumerate(items):
            entry = QListWidgetItem(item["description"])
            entry.setFlags(entry.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            entry.setCheckState(Qt.CheckState.Unchecked)
            entry.setData(Qt.ItemDataRole.UserRole, index)
            self.list_widget.addItem(entry)
        layout.addWidget(self.list_widget)

        address_row = QHBoxLayout()
        address_row.addWidget(QLabel("Address:"))
        self.address_input = QLineEdit()
        self.address_input.setPlaceholderText(
            "Street, City, State ZIP  (leave blank for No Location)"
        )
        address_row.addWidget(self.address_input, 1)
        layout.addLayout(address_row)

        button_row = QHBoxLayout()
        self.apply_button = QPushButton("Apply to selected")
        self.apply_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.apply_button.clicked.connect(self._apply_to_selected)
        button_row.addWidget(self.apply_button)

        self.another_button = QPushButton("Another batch")
        self.another_button.clicked.connect(self._another_batch)
        button_row.addWidget(self.another_button)

        self.finished_button = QPushButton("Finished")
        self.finished_button.setStyleSheet("font-weight: bold;")
        self.finished_button.clicked.connect(self.accept)
        button_row.addWidget(self.finished_button)

        button_row.addStretch(1)
        layout.addLayout(button_row)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #53645e;")
        layout.addWidget(self.status_label)

    def _selected_rows(self) -> list[int]:
        rows: list[int] = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if item.checkState() == Qt.CheckState.Checked:
                rows.append(row)
        return rows

    def _apply_to_selected(self) -> None:
        rows = self._selected_rows()
        if not rows:
            QMessageBox.information(
                self, "Nothing selected",
                "Check at least one item ID to apply the address to.",
            )
            return
        address = self.address_input.text().strip()

        for row in sorted(rows, reverse=True):
            original_index = self.list_widget.item(row).data(
                Qt.ItemDataRole.UserRole
            )
            self.assignments[int(original_index)] = address
            self.list_widget.takeItem(row)

        remaining = self.list_widget.count()
        if remaining == 0:
            self.status_label.setText(
                "All items addressed. Click Finished to close."
            )
        else:
            self.status_label.setText(
                f"Applied to {len(rows)} item(s). {remaining} remaining."
            )
        self.address_input.clear()

    def _another_batch(self) -> None:
        self.address_input.clear()
        self.status_label.setText(
            "New batch started. Check items and enter the next address."
        )


# ---------------------------------------------------------------------------
# Source match picker dialog
# ---------------------------------------------------------------------------

class SourceMatchPickerDialog(QDialog):
    """Radio-list picker for choosing which source email to pull the PDF from.

    If the user clicks Cancel, `self.chosen` stays None and the caller
    should skip the PDF for that record.
    """

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
    if len(record["items"]) > 23:
        raise ValueError("The packing-list template has 23 item rows; split this RITM into another letter.")

    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(destination), UpdateLinks=0, ReadOnly=False)
        sheet = workbook.Worksheets("Packing List")
        sheet.Range("H8").Value = datetime.combine(parse_date(record["date"]), datetime.min.time())
        sheet.Range("H8").NumberFormat = "mm/dd/yyyy"
        sheet.Range("H9").Value = record["location"].strip() or "No Location"
        sheet.Range("H10").Value = f"{record['packing_list']}-{record['letter']}"

        for row, item in enumerate(record["items"], start=13):
            description = item["description"]
            if item["type"]:
                description = f"{description} ({item['type']})"
            sheet.Cells(row, 2).Value = description
            sheet.Cells(row, 8).Value = item["gross"]
            sheet.Cells(row, 9).Value = item["tare"]
            sheet.Cells(row, 10).Value = item["net"]

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
            sheet.Range(address).Value = request_date
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
        sheet.Range("F16").Value = record["ritm"] or "No Request Item"
        sheet.Range("D19").Value = street
        sheet.Range("D20").Value = city
        sheet.Range("D34").Value = f"{record['packing_list']}-{record['letter']}"
        sheet.Range("D35").Value = record["driver_plate"]
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
    if record["letter"]:
        return f"{record['packing_list']}-{record['letter']}"
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
    receiving_date = parse_date(record["date"]).toordinal()
    reel_number = re.search(r"-(\d+)R(?:-|$)", record["packing_list"], re.IGNORECASE)
    reel_order = int(reel_number.group(1)) if reel_number else 1_000_000
    return (
        receiving_date,
        reel_order,
        record["letter"].casefold(),
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

    trucker = str(sheet.Range("F8").Value or "").strip()
    trailer = str(sheet.Range("F9").Value or "").strip()
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
            new_sheet.Range(address).Value = ""
        new_sheet.Range("D12:H42").ClearContents()

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

            if any(len(record["items"]) > 31 for record in ordered_records):
                record = next(record for record in ordered_records if len(record["items"]) > 31)
                raise ValueError(
                    f"{_packing_list_identifier(record)} has more reels than a report tab supports."
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
                    sheet.Range(address).Value = ""
                sheet.Range("D12:H42").ClearContents()

            for sheet, record in zip(reel_sheets, ordered_records):
                driver, plate = _split_driver_and_plate(record["driver_plate"])
                sheet.Range("M2").Value = _packing_list_identifier(record)
                sheet.Range("M4").Value = datetime.combine(
                    parse_date(record["date"]), datetime.min.time()
                )
                sheet.Range("M4").NumberFormat = "mm/dd/yyyy"
                sheet.Range("M5").Value = record["ritm"]
                sheet.Range("C7").Value = record["location"].strip() or "No Location"
                sheet.Range("I7").Value = sum(item["net"] for item in record["items"])
                sheet.Range("C8").Value = "N/A"
                sheet.Range("F8").Value = driver
                sheet.Range("C9").Value = "N/A"
                sheet.Range("F9").Value = plate

                for row, item in enumerate(record["items"], start=12):
                    reel_id, reel_size = _split_reel_id(item["description"])
                    sheet.Cells(row, 4).Value = reel_id
                    sheet.Cells(row, 5).Value = reel_size
                    sheet.Cells(row, 6).Value = item["type"]
                    sheet.Cells(row, 7).Value = item["gross"]
                    sheet.Cells(row, 8).Value = item["tare"]

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
                self.scan_complete.emit(ScanResults(exact_matches, list(near_entries.values())))
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
        self.save_button = QPushButton("Export Workbook")
        self.save_button.setStyleSheet("font-weight: bold;")
        self.save_button.clicked.connect(self._save_records)
        footer.addWidget(self.save_button)
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
            f"{record['packing_list']}-{record['letter']}",
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

    def _append_unmatched_record(
        self,
        record: PackingListRecord,
        items: list[ItemRecord],
        addresses: dict[int, str] | None = None,
    ) -> list[PackingListRecord]:
        if not items:
            return []
        addresses = addresses or {}

        grouped: dict[str, list[ItemRecord]] = {}
        no_address_items: list[ItemRecord] = []
        for idx, item in enumerate(items):
            key = addresses.get(idx, "").strip()
            if not key:
                no_address_items.append(item)
                continue
            grouped.setdefault(key, []).append(item)
        if no_address_items:
            grouped[""] = no_address_items

        ordered_groups = sorted(
            grouped.items(),
            key=lambda kv: (kv[0] == "", -len(kv[1])),
        )

        appended: list[PackingListRecord] = []
        for address, group_items in ordered_groups:
            if address:
                parts = [p.strip() for p in address.split(",")]
                if len(parts) >= 2:
                    city_state = ", ".join(parts[1:])
                    city_state = re.sub(
                        r"\s+\d{5}(?:-\d{4})?\s*$", "", city_state
                    ).strip()
                    location = city_state.replace(",", "")
                else:
                    location = ""
                location_address = address
            else:
                location = ""
                location_address = ""
            unmatched_record: PackingListRecord = {
                **record,
                "group": "no request item",
                "ritm": "No Request Item",
                "letter": self._next_packing_list_letter(
                    record["packing_list"].casefold()
                ),
                "location": location,
                "location_address": location_address,
                "items": group_items,
            }
            self._append_record(unmatched_record)
            appended.append(unmatched_record)
        return appended

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
            self.save_button,
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
            key=lambda email: (email.ritm.casefold(), email.received),
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

    def _scan_finished(self, results: ScanResults) -> None:
        record = self.pending_record
        if record is None:
            self._set_processing_state(True)
            return

        subject_matches = [match for match in results.exact_matches if match.subject_ritms]
        if not subject_matches:
            dialog = AddressBatchDialog(self, record["items"])
            dialog.exec()
            appended = self._append_unmatched_record(
                record, record["items"], dialog.assignments
            )
            self.pending_record = None
            self.imported_record = None
            self.imported_source_path = None
            if not appended:
                self._set_processing_state(True)
                self.scan_results_label.setText("There are no items to queue as unmatched.")
                self.status_label.setText("No items found to queue.")
                return
            queued_ids = ", ".join(
                _packing_list_identifier(r) for r in appended
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

        exact_summary = sorted({
            f"{match.term}: {match.folder} / {match.filename}"
            for match in subject_matches
        })
        found_ritms = sorted({
            ritm for match in subject_matches for ritm in match.subject_ritms
        })
        assigned_ritms = found_ritms

        queued_packing_lists = []
        assigned_item_indexes: set[int] = set()
        ritm_letters: dict[str, str] = {}
        matched_email_count = len({
            match.entry_id or (match.folder, match.subject, match.received)
            for match in subject_matches
        })
        total_accepted_files = 0
        staged_ritm_records: list[tuple[PackingListRecord, list[ExactMatch]]] = []

        for ritm in assigned_ritms:
            ritm_key = ritm.casefold()
            if found_ritms:
                ritm_matches = [
                    match for match in subject_matches if ritm in match.subject_ritms
                ]
            else:
                ritm_matches = []
            if not ritm_matches:
                continue

            current_matched_terms = {match.term.casefold() for match in ritm_matches}
            matched_terms = set(current_matched_terms)
            exact_files = {f"{match.folder} / {match.filename}" for match in ritm_matches}
            accepted_files: set[str] = set()
            accepted_terms = set()
            seen_candidates: set[tuple[str, str, str]] = set()
            for candidate in results.near_matches:
                if (
                    candidate.term.casefold() not in matched_terms
                    or ritm not in candidate.subject_ritms
                ):
                    continue
                candidate_key = (
                    candidate.term.casefold(),
                    candidate.folder.casefold(),
                    candidate.filename.casefold(),
                )
                if candidate_key in seen_candidates:
                    continue
                seen_candidates.add(candidate_key)
                answer = QMessageBox.question(
                    self,
                    "Possible item match",
                    f"This attachment contains '{candidate.guess}', similar to item "
                    f"'{candidate.term}' ({candidate.similarity:.0%}).\n\n"
                    f"File: {candidate.folder} / {candidate.filename}\n"
                    f"Email: {candidate.subject}\n"
                    f"Include this possible item in RITM {ritm.upper()}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer == QMessageBox.StandardButton.Yes:
                    accepted_files.add(
                        f"{candidate.folder} / {candidate.filename} "
                        f"(accepted near match: {candidate.guess} for {candidate.term})"
                    )
                    accepted_terms.add(candidate.term.casefold())
                    total_accepted_files += len(accepted_files)

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
                ),
                "location_address": next(
                    (match.location_address for match in ritm_matches if match.location_address),
                    "",
                ),
                "items": matched_items,
            }
            staged_ritm_records.append((ritm_record, ritm_matches))

        def _ritm_sort_key(entry: tuple[PackingListRecord, list[ExactMatch]]):
            rec, _ = entry
            has_location = bool(rec["location"].strip())
            tier = 0 if has_location else 1
            return (tier, -len(rec["items"]), rec["ritm"].casefold())

        staged_ritm_records.sort(key=_ritm_sort_key)

        for ritm_record, ritm_matches in staged_ritm_records:
            letter = self._next_packing_list_letter(record["packing_list"].casefold())
            ritm_record["letter"] = letter
            ritm_letters[ritm_record["group"]] = letter
            self._append_record(ritm_record)
            self.record_source_matches[_packing_list_identifier(ritm_record)] = list(ritm_matches)
            queued_packing_lists.append(f"{record['packing_list']}-{letter}")

        unmatched_items = [
            item for index, item in enumerate(record["items"])
            if index not in assigned_item_indexes
        ]
        if unmatched_items:
            dialog = AddressBatchDialog(self, unmatched_items)
            dialog.exec()
            appended = self._append_unmatched_record(
                record, unmatched_items, dialog.assignments
            )
            for unmatched_record in appended:
                queued_packing_lists.append(
                    f"{_packing_list_identifier(unmatched_record)} (unmatched)"
                )

        self._merge_matched_emails(
            subject_matches,
            [],
            ritm_letters,
            record["packing_list"],
        )

        if not queued_packing_lists:
            self.pending_record = None
            self.imported_record = None
            self.imported_source_path = None
            self._set_processing_state(True)
            self.status_label.setText("No matched items could be assigned to an RITM group.")
            return
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

        default_directory = Path(__file__).resolve().parent / "packing_list_exports"
        selected_directory = QFileDialog.getExistingDirectory(
            self,
            "Choose where to save packing lists and the Reel Summary Report",
            str(default_directory),
        )
        if not selected_directory:
            return
        destination_directory = Path(selected_directory)
        directory_key = str(destination_directory.resolve()).casefold()
        exported_record_ids = self.exported_record_ids_by_directory.get(directory_key, set())
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
            packing_list_paths.append(
                destination_directory / f"Packing List {safe}-{record['letter']}.xls"
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
        for record in self.records:
            try:
                monday, friday = week_range(record["date"])
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
                weekly_path = destination_directory / filename
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
            f"Updated {week_count} weekly Reel Summary Report(s)",
        ]
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