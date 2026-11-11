"""Shared logging setup for the Verizon Packing Tools applications."""

import logging
import sys
from pathlib import Path
from types import TracebackType
from typing import Type


_APP_FOLDERS = {
    "packinglist",
    "splitmaterial",
    "fillpaymentstatement",
    "pdfmerge",
}


class _ApplicationFilter(logging.Filter):
    def __init__(self, app_name: str) -> None:
        super().__init__()
        self.app_name = app_name

    def filter(self, record: logging.LogRecord) -> bool:
        record.__dict__["app_name"] = self.app_name
        return True


class _SharedFileHandler(logging.FileHandler):
    def __init__(self, path: Path, app_name: str) -> None:
        super().__init__(path, encoding="utf-8")
        self.addFilter(_ApplicationFilter(app_name))
        self.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)s [%(app_name)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))


def get_shared_log_path() -> Path:
    if getattr(sys, "frozen", False):
        executable_dir = Path(sys.executable).resolve().parent
        if executable_dir.name.casefold() in _APP_FOLDERS:
            app_root = executable_dir.parent
        else:
            app_root = executable_dir
    else:
        app_root = Path(__file__).resolve().parent
    return app_root / "logs" / "verizon_tools.log"


def configure_app_logging(app_name: str) -> logging.Logger:
    log_path = get_shared_log_path()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    resolved_log_path = str(log_path.resolve())
    if not any(
        isinstance(handler, _SharedFileHandler)
        and handler.baseFilename == resolved_log_path
        for handler in root_logger.handlers
    ):
        root_logger.addHandler(_SharedFileHandler(log_path, app_name))

    logger = logging.getLogger(app_name)
    logger.info("Application started; shared log: %s", log_path)

    previous_excepthook = sys.excepthook

    def log_uncaught_exception(
        exc_type: Type[BaseException],
        exc_value: BaseException,
        exc_traceback: TracebackType | None,
    ) -> None:
        logging.getLogger(app_name).critical(
            "Uncaught exception",
            exc_info=(exc_type, exc_value, exc_traceback),
        )
        previous_excepthook(exc_type, exc_value, exc_traceback)

    sys.excepthook = log_uncaught_exception
    return logger
