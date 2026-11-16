"""
setup_app.py — One-time setup for Packing List App.

Points this account at a shared folder containing config.json + the three
databases, creates any missing databases, and lets you browse/edit rows
in each one. Edit mode is read-only by default; check the box to allow
cell edits (double-click → type → Enter), then click Save changes.
"""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPushButton, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)
import logging as _logging
_APP_LOGGER_local = _logging.getLogger("packing_list_setup")


APP_DIR = Path(__file__).resolve().parent
LOCAL_BOOTSTRAP = Path.home() / "AppData" / "Roaming" / "PackingListApp" / "bootstrap.json"

SCHEMA_EMAIL = """
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

SCHEMA_MANUAL = """
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

SCHEMA_EXPORT = SCHEMA_MANUAL  # identical table shapes


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------
def read_bootstrap() -> dict:
    if LOCAL_BOOTSTRAP.exists():
        try:
            return json.loads(LOCAL_BOOTSTRAP.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def write_bootstrap(data: dict) -> None:
    LOCAL_BOOTSTRAP.parent.mkdir(parents=True, exist_ok=True)
    LOCAL_BOOTSTRAP.write_text(json.dumps(data, indent=2), encoding="utf-8")


def read_shared_config(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def write_shared_config(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def ensure_db(path: Path, schema: str) -> bool:
    if path.exists():
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript(schema)
    finally:
        conn.close()
    return True


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class SetupWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Packing List — Setup")
        self.resize(1000, 760)
        self.shared_config_path: Path | None = None
        self.shared_config: dict = {}

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        layout.addWidget(QLabel(
            "<b>Step 1.</b> Choose the shared folder where all databases "
            "and summaries live. Both accounts should open the same folder."
        ))
        folder_row = QHBoxLayout()
        self.folder_input = QLineEdit()
        self.folder_input.setReadOnly(True)
        self.folder_input.setPlaceholderText("No folder selected")
        folder_row.addWidget(self.folder_input, 1)
        pick_btn = QPushButton("Choose folder…")
        pick_btn.clicked.connect(self._pick_folder)
        folder_row.addWidget(pick_btn)
        layout.addLayout(folder_row)

        layout.addWidget(QLabel(
            "<b>Step 2.</b> Create the databases in that folder if missing."
        ))
        self.ensure_btn = QPushButton("Create / verify databases")
        self.ensure_btn.clicked.connect(self._ensure_databases)
        self.ensure_btn.setEnabled(False)
        layout.addWidget(self.ensure_btn)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        layout.addWidget(QLabel(
            "<b>Step 3.</b> Browse and edit rows. Check 'Allow editing' to "
            "enable cell edits, then click 'Save changes' to commit."
        ))
        self.tabs = QTabWidget()
        self.tabs.setEnabled(False)
        layout.addWidget(self.tabs, 1)

        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        close_row.addWidget(close_btn)
        layout.addLayout(close_row)

        boot = read_bootstrap()
        existing = boot.get("shared_folder")
        if existing and Path(existing).is_dir():
            self.folder_input.setText(existing)
            self._load_folder(Path(existing))

    def _pick_folder(self) -> None:
        chosen = QFileDialog.getExistingDirectory(
            self, "Choose shared folder", str(Path.home())
        )
        if not chosen:
            return
        self.folder_input.setText(chosen)
        self._load_folder(Path(chosen))

    def _load_folder(self, folder: Path) -> None:
        write_bootstrap({"shared_folder": str(folder)})
        self.shared_config_path = folder / "config.json"
        self.shared_config = read_shared_config(self.shared_config_path)
        for key, default in (
            ("email_database_path", str(folder / "Email_Delivery.db")),
            ("manual_database_path", str(folder / "PackingLists_Manual.db")),
            ("database_path", str(folder / "PackingLists.db")),
            ("summaries_folder", str(folder / "Summaries")),
        ):
            self.shared_config.setdefault(key, default)
        write_shared_config(self.shared_config_path, self.shared_config)
        
        # Also write these paths into the local scanner_config.json so
        # packing_list_gui.py reads the same folder without needing the
        # bootstrap overlay.
        try:
            local_cfg_path = APP_DIR / "scanner_config.json"
            local_cfg = {}
            if local_cfg_path.exists():
                try:
                    local_cfg = json.loads(
                        local_cfg_path.read_text(encoding="utf-8")
                    )
                except Exception:
                    local_cfg = {}
            local_cfg["email_database_path"] = self.shared_config["email_database_path"]
            local_cfg["manual_database_path"] = self.shared_config["manual_database_path"]
            local_cfg["database_path"] = self.shared_config["database_path"]
            local_cfg["summaries_folder"] = self.shared_config["summaries_folder"]
            local_cfg_path.write_text(
                json.dumps(local_cfg, indent=2), encoding="utf-8"
            )
        except Exception:
            _APP_LOGGER_local.exception("Could not write local scanner_config.json.")
        self.ensure_btn.setEnabled(True)
        self.status_label.setText(
            f"Shared config: {self.shared_config_path}\n"
            "Click 'Create / verify databases' to set up the files."
        )

    def _ensure_databases(self) -> None:
        if not self.shared_config.get("email_database_path"):
            QMessageBox.warning(
                self, "No folder selected",
                "Pick a folder at the top of the window first, then click "
                "Create / verify databases again.",
            )
            return

        lines = []
        for key, schema, label in (
            ("email_database_path", SCHEMA_EMAIL, "Email DB"),
            ("manual_database_path", SCHEMA_MANUAL, "Input DB"),
            ("database_path", SCHEMA_EXPORT, "Export DB"),
        ):
            path = Path(self.shared_config[key])
            if ensure_db(path, schema):
                lines.append(f"  • created {label}: {path}")
            else:
                lines.append(f"  • verified {label}: {path}")
        Path(self.shared_config["summaries_folder"]).mkdir(
            parents=True, exist_ok=True
        )
        self.status_label.setText("\n".join(lines))
        self._build_tabs()
        self.tabs.setEnabled(True)

    def _build_tabs(self) -> None:
        self.tabs.clear()
        for title, key, tables in (
            ("Email DB", "email_database_path",
             ["Email_Delivery", "Email_Ritm", "Email_Item"]),
            ("Input DB", "manual_database_path",
             ["Reel_Load_Database", "Reel_Items"]),
            ("Export DB", "database_path",
             ["Reel_Load_Database", "Reel_Items"]),
        ):
            self.tabs.addTab(
                self._make_db_tab(Path(self.shared_config[key]), tables),
                title,
            )

    def _make_db_tab(self, db_path: Path, tables: list[str]) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        controls = QHBoxLayout()
        refresh_btn = QPushButton("Refresh")
        controls.addWidget(refresh_btn)
        edit_check = QCheckBox("Allow editing")
        controls.addWidget(edit_check)
        save_btn = QPushButton("Save changes")
        save_btn.setStyleSheet("font-weight: bold;")
        controls.addWidget(save_btn)
        controls.addStretch(1)
        layout.addLayout(controls)

        table_selector = QTabWidget()
        layout.addWidget(table_selector, 1)

        views: dict[str, QTableWidget] = {}
        for table_name in tables:
            widget = QTableWidget()
            widget.setAlternatingRowColors(True)
            widget.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            table_selector.addTab(widget, table_name)
            views[table_name] = widget

        def refresh():
            conn = sqlite3.connect(str(db_path), timeout=5.0)
            try:
                for table_name, widget in views.items():
                    self._fill_table(conn, table_name, widget)
            finally:
                conn.close()

        def toggle_edit(state):
            trigger = (
                QTableWidget.EditTrigger.DoubleClicked
                | QTableWidget.EditTrigger.EditKeyPressed
                if state
                else QTableWidget.EditTrigger.NoEditTriggers
            )
            for widget in views.values():
                widget.setEditTriggers(trigger)

        def save_all():
            conn = sqlite3.connect(str(db_path), timeout=10.0)
            saved = 0
            try:
                for table_name, widget in views.items():
                    pk_col = self._primary_key_column(conn, table_name)
                    if pk_col is None:
                        continue
                    headers = [
                        widget.horizontalHeaderItem(c).text()
                        for c in range(widget.columnCount())
                    ]
                    if pk_col not in headers:
                        continue
                    pk_index = headers.index(pk_col)
                    conn.execute("BEGIN")
                    try:
                        for row in range(widget.rowCount()):
                            pk_item = widget.item(row, pk_index)
                            if pk_item is None:
                                continue
                            pk_value = pk_item.text()
                            for col, header in enumerate(headers):
                                if header == pk_col:
                                    continue
                                cell = widget.item(row, col)
                                if cell is None:
                                    continue
                                conn.execute(
                                    f"UPDATE {table_name} SET {header} = ? "
                                    f"WHERE {pk_col} = ?",
                                    (cell.text(), pk_value),
                                )
                                saved += 1
                        conn.execute("COMMIT")
                    except Exception as exc:
                        conn.execute("ROLLBACK")
                        QMessageBox.critical(
                            self, "Save failed", f"{table_name}: {exc}"
                        )
                        return
                QMessageBox.information(
                    self, "Saved", f"{saved} cell(s) written."
                )
            finally:
                conn.close()
            refresh()

        refresh_btn.clicked.connect(refresh)
        edit_check.stateChanged.connect(toggle_edit)
        save_btn.clicked.connect(save_all)

        refresh()
        return tab

    @staticmethod
    def _fill_table(conn, table_name: str, widget: QTableWidget) -> None:
        try:
            cursor = conn.execute(f"SELECT * FROM {table_name} LIMIT 500")
            columns = [d[0] for d in cursor.description]
            rows = cursor.fetchall()
        except Exception as exc:
            widget.setRowCount(1)
            widget.setColumnCount(1)
            widget.setHorizontalHeaderLabels(["Error"])
            widget.setItem(0, 0, QTableWidgetItem(str(exc)))
            return
        widget.setRowCount(len(rows))
        widget.setColumnCount(len(columns))
        widget.setHorizontalHeaderLabels(columns)
        widget.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                item = QTableWidgetItem("" if value is None else str(value))
                widget.setItem(r, c, item)

    @staticmethod
    def _primary_key_column(conn, table_name: str) -> str | None:
        try:
            rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
            for _cid, name, _typ, _notnull, _dflt, pk in rows:
                if pk:
                    return name
        except Exception:
            return None
        return None


if __name__ == "__main__":
    app = QApplication([])
    window = SetupWindow()
    window.show()
    app.exec()