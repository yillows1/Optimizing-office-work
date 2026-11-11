"""
PDF Merge Tool

For each scanned PDF in a "scans" folder:
  1. Read the FRN identifier (e.g. FRN260929-1R-A) from the scan's filename.
  2. Look for <export root>/<identifier>/ in the export root the user picked.
  3. If found, grab <that folder>/Request Item for Load <identifier>.pdf
     and append the scan's pages AFTER it (in place, atomically).
  4. File the scan into that folder as "Scanned Delivery <identifier>.pdf".

Standalone by design: no imports from the packing-list app, no Excel,
no Outlook.  Only pypdf (already a dependency) and PySide6.
"""

from __future__ import annotations

import io
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
try:
    from version import VERSION
except Exception:
    VERSION = "dev"

from app_logging import configure_app_logging


def get_user_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


log = configure_app_logging("PDF Merge")

from pypdf import PdfReader, PdfWriter


# ---------------------------------------------------------------------------
# FRN helpers
# ---------------------------------------------------------------------------
#
# Matches "FRN260929-1R" optionally followed by "-A" / "-B" / ...
# The load folders and the "Request Item for Load ..." PDFs use exactly
# this identifier form, so a scan named "scan FRN260929-1R-A.pdf" resolves
# to the folder "FRN260929-1R-A" and the PDF
# "Request Item for Load FRN260929-1R-A.pdf".
#
_FRN_RE = re.compile(
    r"(FRN\d{6}-\d+R)(?:-([A-Za-z]+))?",
    re.IGNORECASE,
)


def extract_frn(text: str) -> tuple[str, str]:
    """Return (frn, letter) from arbitrary text, or ('', '')."""
    if not text:
        return "", ""
    match = _FRN_RE.search(text)
    if not match:
        return "", ""
    frn = match.group(1).upper()
    letter = (match.group(2) or "").upper()
    return frn, letter


def frn_identifier(frn: str, letter: str) -> str:
    """Build the load-folder identifier exactly like _packing_list_identifier."""
    return f"{frn}-{letter}" if letter else frn


def request_item_pdf_name(identifier: str) -> str:
    """The PDF the packing-list app writes into each load folder."""
    return f"Request Item for Load {identifier}.pdf"


# ---------------------------------------------------------------------------
# Scanning / matching
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ScanMatch:
    scan_path: Path
    frn: str
    letter: str
    identifier: str
    target_folder: Path | None = None
    target_pdf: Path | None = None
    status: str = ""


def plan_scan_matches(
    scan_folder: Path,
    export_root: Path,
    progress=None,
) -> list[ScanMatch]:

    scans = sorted(
        p for p in scan_folder.iterdir()
        if p.is_file()
        and p.suffix.casefold() == ".pdf"
        and not p.name.startswith(".")
    )
    total = len(scans)
    plans: list[ScanMatch] = []

    for i, scan in enumerate(scans, start=1):
        if progress:
            progress(i, total, scan.name)

        frn, letter = extract_frn(scan.stem)
        if not frn:
            plans.append(ScanMatch(
                scan, "", "", "",
                status="No FRN found in filename",
            ))
            continue

        identifier = frn_identifier(frn, letter)
        folder = export_root / identifier

        if not folder.is_dir():
            plans.append(ScanMatch(
                scan, frn, letter, identifier,
                status=f"No folder {identifier} in export root",
            ))
            continue

        target_pdf = folder / request_item_pdf_name(identifier)
        if not target_pdf.exists():
            plans.append(ScanMatch(
                scan, frn, letter, identifier,
                target_folder=folder,
                target_pdf=target_pdf,
                status=f"Missing {target_pdf.name}",
            ))
            continue

        plans.append(ScanMatch(
            scan, frn, letter, identifier,
            target_folder=folder,
            target_pdf=target_pdf,
            status="Ready",
        ))

    return plans


# ---------------------------------------------------------------------------
# Merge
# ---------------------------------------------------------------------------

def _append_pdf(target: Path, source_bytes: bytes) -> None:

    writer = PdfWriter()
    for page in PdfReader(str(target)).pages:
        writer.add_page(page)
    for page in PdfReader(io.BytesIO(source_bytes)).pages:
        writer.add_page(page)

    target.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=".merge-", suffix=".pdf", dir=str(target.parent)
    )
    Path(tmp_name).close()
    try:
        with open(tmp_name, "wb") as fh:
            writer.write(fh)
        Path(tmp_name).replace(target)
    finally:
        try:
            Path(tmp_name).unlink()
        except OSError:
            pass


def _archive_scan(scan: ScanMatch, overwrite: bool) -> Path:
    """Move the scan into its load folder, renamed to a stable name."""
    assert scan.target_folder is not None
    destination = scan.target_folder / f"Scanned Delivery {scan.identifier}.pdf"

    if destination.exists() and not overwrite:
        n = 2
        while True:
            candidate = scan.target_folder / (
                f"Scanned Delivery {scan.identifier} ({n}).pdf"
            )
            if not candidate.exists():
                destination = candidate
                break
            n += 1

    # Copy-then-delete is safer than rename across network shares.
    shutil.copy2(scan.scan_path, destination)
    try:
        scan.scan_path.unlink()
    except OSError:
        pass
    return destination


def apply_scan_matches(
    plans: list[ScanMatch],
    overwrite_scans: bool = False,
) -> tuple[int, int, list[str]]:
    """Run the merges. Returns (merged_count, failed_count, messages)."""
    merged = 0
    failed = 0
    messages: list[str] = []

    for plan in plans:
        if plan.target_pdf is None or plan.target_folder is None:
            continue
        try:
            source_bytes = plan.scan_path.read_bytes()
            _append_pdf(plan.target_pdf, source_bytes)
            archived = _archive_scan(plan, overwrite=overwrite_scans)
            log.info(f"merge {plan.scan_path} -> {plan.target_pdf}")
            merged += 1
            messages.append(
                f"{plan.scan_path.name} -> {plan.target_pdf.name} "
                f"(scan filed as {archived.name})"
            )
        except Exception as error:
            failed += 1
            log.exception(
                "Failed merging scan %s into %s.",
                plan.scan_path,
                plan.target_pdf,
            )
            messages.append(f"{plan.scan_path.name}: {error}")

    return merged, failed, messages


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class PdfMergeApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(
            f"PDF Merge — Append Scanned Deliveries  (v{VERSION})"
        )
        self.resize(1040, 660)
        self.setMinimumSize(840, 540)

        self.scan_folder: Path | None = None
        self.export_root: Path | None = None
        self.plans: list[ScanMatch] = []

        self._build_ui()

    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 20, 20, 20)

        title = QLabel("Append Scanned Deliveries")
        title.setStyleSheet("font-size: 18pt; font-weight: bold;")
        root.addWidget(title)

        hint = QLabel(
            "Point at a folder of scanned PDFs and at the export root that "
            "holds the FRN… load folders. Each scan is matched by the FRN "
            "in its filename to <export root>/<FRN identifier>/, appended "
            "after 'Request Item for Load <FRN identifier>.pdf' inside "
            "that folder, and then filed there."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #53645e; margin-bottom: 8px;")
        root.addWidget(hint)

        # --- folder pickers ---
        scan_row = QHBoxLayout()
        scan_btn = QPushButton("Choose Scans Folder…")
        scan_btn.clicked.connect(self._choose_scan_folder)
        self.scan_folder_label = QLabel("Scans folder: (not set)")
        self.scan_folder_label.setStyleSheet("color: #53645e;")
        self.scan_folder_label.setWordWrap(True)
        scan_row.addWidget(scan_btn)
        scan_row.addWidget(self.scan_folder_label, 1)
        root.addLayout(scan_row)

        root_row = QHBoxLayout()
        root_btn = QPushButton("Choose Export Root…")
        root_btn.clicked.connect(self._choose_export_root)
        self.export_root_label = QLabel("Export root: (not set)")
        self.export_root_label.setStyleSheet("color: #53645e;")
        self.export_root_label.setWordWrap(True)
        root_row.addWidget(root_btn)
        root_row.addWidget(self.export_root_label, 1)
        root.addLayout(root_row)

        # --- options ---
        options = QGroupBox("Options")
        opt_layout = QVBoxLayout(options)
        self.overwrite_checkbox = QCheckBox(
            "Overwrite an existing 'Scanned Delivery <FRN>.pdf' instead of "
            "numbering it"
        )
        opt_layout.addWidget(self.overwrite_checkbox)
        root.addWidget(options)

        # --- preview table ---
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Scan file", "FRN identifier", "Target PDF", "Status"]
        )
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        root.addWidget(self.table, 1)

        # --- buttons ---
        button_row = QHBoxLayout()
        self.preview_button = QPushButton("Preview Matches")
        self.preview_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.preview_button.clicked.connect(self._preview)
        button_row.addWidget(self.preview_button)

        self.merge_button = QPushButton("Merge && File Scans")
        self.merge_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white;"
        )
        self.merge_button.setEnabled(False)
        self.merge_button.clicked.connect(self._merge)
        button_row.addWidget(self.merge_button)
        button_row.addStretch(1)
        root.addLayout(button_row)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #53645e;")
        root.addWidget(self.status_label)

    # ------------------------------------------------------------------

    def _choose_scan_folder(self) -> None:
        chosen = QFileDialog.getExistingDirectory(self, "Choose scans folder")
        if chosen:
            self.scan_folder = Path(chosen)
            log.info("Scans folder selected: %s", self.scan_folder)
            self.scan_folder_label.setText(f"Scans folder: {chosen}")
            self._invalidate()

    def _choose_export_root(self) -> None:
        chosen = QFileDialog.getExistingDirectory(
            self, "Choose export root (parent of the FRN… folders)"
        )
        if chosen:
            self.export_root = Path(chosen)
            log.info("Export root selected: %s", self.export_root)
            self.export_root_label.setText(f"Export root: {chosen}")
            self._invalidate()

    def _invalidate(self) -> None:
        self.plans = []
        self.table.setRowCount(0)
        self.merge_button.setEnabled(False)

    def _preview(self) -> None:
        log.info(
            "Preview requested: scans=%s; export_root=%s",
            self.scan_folder,
            self.export_root,
        )
        if self.scan_folder is None or self.export_root is None:
            log.warning("Preview rejected: one or both folders were not selected.")
            QMessageBox.information(
                self, "Folders missing",
                "Choose both a scans folder and an export root first.",
            )
            return
        if not self.scan_folder.is_dir():
            log.error("Scans folder does not exist: %s", self.scan_folder)
            QMessageBox.warning(self, "Scans folder missing", str(self.scan_folder))
            return
        if not self.export_root.is_dir():
            log.error("Export root does not exist: %s", self.export_root)
            QMessageBox.warning(self, "Export root missing", str(self.export_root))
            return

        self.status_label.setText("Scanning for matches…")
        QApplication.processEvents()

        try:
            self.plans = plan_scan_matches(self.scan_folder, self.export_root)
        except Exception:
            log.exception(
                "Preview scan failed for scans=%s; export_root=%s",
                self.scan_folder,
                self.export_root,
            )
            QMessageBox.critical(
                self, "Preview failed", "Could not scan for matching PDFs. "
                "See the shared log for details."
            )
            return

        self.table.setRowCount(0)
        ready = 0
        for plan in self.plans:
            row = self.table.rowCount()
            self.table.insertRow(row)
            target_text = (
                plan.target_pdf.name if plan.target_pdf
                else "—"
            )
            values = (
                plan.scan_path.name,
                plan.identifier or "(none)",
                target_text,
                plan.status,
            )
            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))
            if plan.status == "Ready":
                ready += 1

        self.merge_button.setEnabled(ready > 0)
        skipped = len(self.plans) - ready
        log.info(
            "Preview completed: total=%d ready=%d skipped=%d",
            len(self.plans), ready, skipped,
        )
        for plan in self.plans:
            log.info(
                "Preview item: scan=%s identifier=%s target=%s status=%s",
                plan.scan_path,
                plan.identifier or "(none)",
                plan.target_pdf,
                plan.status,
            )
        self.status_label.setText(
            f"{len(self.plans)} scan(s); {ready} ready"
            + (f"; {skipped} skipped." if skipped else ".")
        )

    def _merge(self) -> None:
        ready = [p for p in self.plans if p.status == "Ready"]
        if not ready:
            log.warning("Merge requested with no ready scans.")
            return
        log.info(
            "Merge confirmation requested: ready=%d overwrite_existing=%s",
            len(ready),
            self.overwrite_checkbox.isChecked(),
        )

        confirm = QMessageBox.question(
            self,
            "Merge scans?",
            f"Append {len(ready)} scan(s) after their "
            "'Request Item for Load …' PDFs and file the scans into the "
            "load folders?\n\n"
            "This modifies the existing PDFs. Back them up first if you "
            "need to keep the originals.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            log.info("User cancelled PDF merge.")
            return

        merged, failed, messages = apply_scan_matches(
            ready, overwrite_scans=self.overwrite_checkbox.isChecked()
        )
        log.info("Merge finished: merged=%d failed=%d", merged, failed)

        self.status_label.setText(
            f"Merged {merged} scan(s); {failed} failure(s)."
        )
        self.merge_button.setEnabled(False)

        lines = [f"Merged {merged} scan(s)."]
        if failed:
            lines.append(f"{failed} failure(s):")
        lines.extend(messages)
        QMessageBox.information(self, "Merge complete", "\n".join(lines))

        # Refresh so the user sees the new state.
        self._preview()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfMergeApp()
    window.show()
    sys.exit(app.exec())