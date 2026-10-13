"""
Launcher menu for the Verizon packing tools.
"""

import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class LauncherWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Verizon Packing Tools")
        self.setFixedSize(440, 280)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 28, 28, 28)
        root.setSpacing(14)

        title = QLabel("Verizon Packing Tools")
        title.setStyleSheet("font-size: 16pt; font-weight: bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(title)

        subtitle = QLabel("Choose a tool to open.")
        subtitle.setStyleSheet("color: #53645e;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(subtitle)
        root.addSpacing(8)

        packing_btn = QPushButton("Packing List")
        packing_btn.setMinimumHeight(52)
        packing_btn.setStyleSheet(
            "font-size: 13pt; font-weight: bold; "
            "background: #315b52; color: white; border-radius: 5px;"
        )
        packing_btn.clicked.connect(
            lambda: self._launch("PackingList", "PackingList.exe")
        )
        root.addWidget(packing_btn)

        split_btn = QPushButton("Split Material")
        split_btn.setMinimumHeight(52)
        split_btn.setStyleSheet(
            "font-size: 13pt; font-weight: bold; "
            "background: #315b52; color: white; border-radius: 5px;"
        )
        split_btn.clicked.connect(
            lambda: self._launch("SplitMaterial", "SplitMaterial.exe")
        )
        root.addWidget(split_btn)
        root.addStretch(1)

    def _launch(self, subfolder: str, exe_name: str) -> None:
        exe_path = app_dir() / subfolder / exe_name
        if not exe_path.exists():
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
            QMessageBox.critical(self, "Could not launch", f"{exe_name}: {error}")


def main() -> None:
    app = QApplication(sys.argv)
    window = LauncherWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()