"""
Launcher menu for the Verizon packing tools.
"""

import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QGridLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from app_logging import configure_app_logging, get_shared_log_path

log = configure_app_logging("Launcher")

try:
    from version import VERSION, APP_NAME
except Exception:
    APP_NAME = "Verizon Packing Tools"
    VERSION = "dev"


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class LauncherWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  —  v{VERSION}")
        self.setFixedSize(560, 420)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 28, 28, 28)
        root.setSpacing(12)

        title = QLabel(APP_NAME)
        title.setStyleSheet("font-size: 18pt; font-weight: bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(title)

        version_label = QLabel(f"Version {VERSION}")
        version_label.setStyleSheet("color: #53645e;")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(version_label)
        root.addSpacing(6)

        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)

        buttons = [
            ("Packing List",      "PackingList",          "PackingList.exe"),
            ("Split Material",    "SplitMaterial",        "SplitMaterial.exe"),
            ("Fill Payment Stmt", "FillPaymentStatement", "FillPaymentStatement.exe"),
            ("PDF Merge",         "PDFMerge",             "PDFMerge.exe"),
        ]

        for i, (label, subfolder, exe) in enumerate(buttons):
            btn = QPushButton(label)
            btn.setMinimumHeight(64)
            btn.setStyleSheet(
                "font-size: 12pt; font-weight: bold; "
                "background: #315b52; color: white; border-radius: 5px;"
            )
            btn.setToolTip(f"Open {label}")
            btn.clicked.connect(
                lambda _=False, s=subfolder, e=exe: self._launch(s, e)
            )
            grid.addWidget(btn, i // 2, i % 2)

        root.addLayout(grid)
        root.addSpacing(6)

        logs_btn = QPushButton("Open Shared Log")
        logs_btn.setStyleSheet(
            "font-size: 10pt; background: #e2e9e5; color: #20332e; "
            "border: 1px solid #9baea6; border-radius: 4px; padding: 8px;"
        )
        logs_btn.clicked.connect(self._open_logs)
        root.addWidget(logs_btn)
        root.addStretch(1)

    def _launch(self, subfolder: str, exe_name: str) -> None:
        exe_path = app_dir() / subfolder / exe_name
        log.info("Launch requested: app=%s path=%s", subfolder, exe_path)
        if not exe_path.exists():
            log.error("Application executable is missing: %s", exe_path)
            QMessageBox.critical(
                self,
                "Missing file",
                f"Could not find:\n{exe_path}\n\n"
                f"Make sure the '{subfolder}' folder sits next to this launcher "
                f"and contains {exe_name}.",
            )
            return
        try:
            subprocess.Popen(
                [str(exe_path)],
                cwd=str(exe_path.parent),
                close_fds=True,
            )
        except Exception as error:
            log.exception("Could not launch %s.", exe_path)
            QMessageBox.critical(self, "Could not launch", f"{exe_name}: {error}")

    def _open_logs(self) -> None:
        log_path = get_shared_log_path()
        log.info("User opened shared log: %s", log_path)
        try:
            subprocess.Popen(["notepad.exe", str(log_path)])
        except Exception:
            log.exception("Could not open shared log file %s.", log_path)
            QMessageBox.critical(
                self, "Could not open log",
                f"Could not open the shared log file:\n{log_path}",
            )


def main() -> None:
    log.info("Launcher UI starting.")
    app = QApplication(sys.argv)
    window = LauncherWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()