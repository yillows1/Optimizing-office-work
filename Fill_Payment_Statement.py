# Fill_Payment_Statement.py
"""
Collect loads for a week, drag-to-reorder, and fill the Verizon BAU Payment
Statement.

Inputs:
  1. JC Payment Statement            (.xls)  — received from Jersey City
  2. Receiving root (summaries)      (dir)   — contains YYYY/"Month YYYY"/Reel Summary/
  3. Receiving root (loose)          (dir)   — contains YYYY/"Month YYYY"/<loose files>
  4. Blank template                  (.xls)  — for the Rahway-only statement
  5. Scrap Pickup root               (dir)   — contains Scrap Pick Up Request files
  6. Week Ending output root         (dir)   — outputs go under YYYY/"Month YYYY"/

Outputs:
  A) JC-merged   = JC file + our reels + our loose   (email this)
  B) Rahway-only = blank template + our reels + our loose  (print this)
"""

from __future__ import annotations

import json
import re
import sys
import shutil
import logging as _logging
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import win32com.client

from PySide6.QtCore import Qt
from PySide6.QtGui import QShortcut, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QHBoxLayout, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QHeaderView, QAbstractItemView,
)


# =============================================================================
# Config / paths
# =============================================================================

def get_user_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_PATH = get_user_dir() / "scanner_config.json"
_LOG = get_user_dir() / "fill_statement.log"
_logging.basicConfig(
    filename=str(_LOG), level=_logging.INFO,
    format="%(asctime)s %(message)s", datefmt="%H:%M:%S",
)


def log(msg: str) -> None:
    _logging.info(msg)


def load_cfg() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_cfg(cfg: dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass


# =============================================================================
# Constants
# =============================================================================

DATE_TAB_FMT = "%m_%d_%y"

FOR_C_FRNS = [f"B{r}" for r in range(7, 12)]
VZ_FRNS = [f"D{r}" for r in range(7, 12)]

CLASS_TO_ROW = {
    "17": 14, "2": 15, "3": 16, "10": 17, "53": 18, "20": 19, "23": 20,
    "REG Wire": 21, "4": 22, "49": 23,
    "PICS A": 26, "PICS B": 27, "PICS C": 28,
    "5.1": 29, "Plastic SC": 30, "AL SC": 31, "LEAD SC": 32, "Irony AL": 33,
    "16": 34, "5.3": 35, "69": 36, "MTL SC": 37, "3.99": 38, "69S": 39,
    "19": 40, "DPWIRL": 41, "21": 42, "183": 43, "183W": 44,
    "99 FO": 45, "CDBD": 46, "WOOD": 47, "99X": 48,
    "Loose Plastic": 49, "Keyboard/Mouse": 50, "Remote": 51,
    "STB": 52, "RETURN": 53, "Batteries": 54,
    "Ice/Water/Snow": 55, "PKG (Bin,Skid,Box)": 56,
}

CLASS_ALIASES = {
    "L": "99X", "LOOSE": "99X", "LOOSE SCRAP": "99X",
    "WOOD REELS": "WOOD", "PLASTIC": "Loose Plastic",
    "KBD/MOUSE": "Keyboard/Mouse", "KBD": "Keyboard/Mouse",
    "MOUSE": "Keyboard/Mouse", "REMOTE": "Remote", "STB": "STB",
    "RETURN": "RETURN", "EMPTY REELS RETURN": "RETURN",
    "OTHER": "RETURN",
    "99FO": "99 FO", "FO99": "99 FO",
}


def normalize_class(raw: str) -> str:
    if not raw:
        return "99X"
    key = str(raw).strip()
    if re.fullmatch(r"\d+\.0+", key):
        key = key.split(".")[0]
    upper = key.upper()
    if upper in CLASS_ALIASES:
        return CLASS_ALIASES[upper]
    if key in CLASS_TO_ROW:
        return key
    for c in CLASS_TO_ROW:
        if c.upper() == upper:
            return c
    return "99X"


_MATERIAL_TOKENS = sorted(
    {k for k in CLASS_TO_ROW if k.upper() != "L"} |
    {k for k in CLASS_ALIASES if k.upper() != "L"},
    key=len,
    reverse=True,
)

_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def month_folder_name(d: date) -> str:
    return f"{_MONTH_NAMES[d.month - 1]} {d.year}"


def month_folders_for_week(monday: date) -> list[tuple[int, int]]:
    """Return (year, month) pairs touched by Mon..Sat of this week."""
    saturday = monday + timedelta(days=5)
    out: list[tuple[int, int]] = []
    cur = monday
    while cur <= saturday:
        key = (cur.year, cur.month)
        if key not in out:
            out.append(key)
        cur = (cur.replace(day=1) + timedelta(days=32)).replace(day=1)
    return out


# =============================================================================
# Data model
# =============================================================================

@dataclass
class LoadRow:
    source: str
    frn: str
    ritm: str = ""
    origin: str = ""
    trucker: str = ""
    trailer: str = ""
    date: str = ""
    gross: int = 0
    tare: int = 0
    net: int = 0
    class_nets: dict = field(default_factory=dict)
    class_gross: dict = field(default_factory=dict)
    class_tare: dict = field(default_factory=dict)
    folder: str = ""
    note: str = ""


# =============================================================================
# COM helpers
# =============================================================================

def _read_cell(ws, address: str):
    try:
        return ws.Range(address).Value
    except Exception:
        return None


def _try_write(ws, address: str, value) -> bool:
    rng = ws.Range(address)
    if ws.ProtectContents and rng.Locked:
        log(f"SKIP protected: {ws.Name}!{address}")
        return False
    try:
        rng.Value = value
        return True
    except Exception as e:
        msg = str(e).lower()
        if "protected" in msg or "read-only" in msg:
            log(f"SKIP protected: {ws.Name}!{address}")
            return False
        raise


def _additive_formula(existing_formula, existing_value, additions: list[int]) -> str | None:
    """Build a formula that adds contributions without discarding existing data."""
    nz = [int(value) for value in additions if value]
    if not nz:
        return None

    base = ""
    if isinstance(existing_formula, str):
        existing_formula = existing_formula.strip()
        if existing_formula.startswith("="):
            try:
                float(existing_value)
            except (TypeError, ValueError):
                raise ValueError(
                    f"Cannot safely add to non-numeric cell value {existing_value!r}"
                )
            base = existing_formula[1:]
        elif existing_formula:
            try:
                number = float(existing_formula)
            except ValueError as e:
                raise ValueError(
                    f"Cannot safely add to cell contents {existing_formula!r}"
                ) from e
            if number:
                base = existing_formula
        elif existing_value not in (None, ""):
            raise ValueError(
                f"Cannot safely add to cell value {existing_value!r}"
            )
    elif existing_formula is None and existing_value not in (None, ""):
        raise ValueError(f"Cannot safely add to cell value {existing_value!r}")
    elif isinstance(existing_formula, (int, float)) and existing_formula:
        base = (str(int(existing_formula)) if float(existing_formula).is_integer()
                else str(existing_formula))

    terms = ([f"({base})"] if base else []) + [str(value) for value in nz]
    return "=" + "+".join(terms)


class _ReadOnlyWorkbook:
    def __init__(self, path: Path):
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False
        self.excel.EnableEvents = False
        self.wb = self.excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)

    def sheets(self) -> list[str]:
        return [s.Name for s in self.wb.Worksheets]

    def sheet(self, name: str):
        return self.wb.Worksheets(name)

    def close(self) -> None:
        try: self.wb.Close(SaveChanges=False)
        except Exception: pass
        try: self.excel.Quit()
        except Exception: pass
        try:
            import pythoncom
            pythoncom.CoFreeUnusedLibraries()
        except Exception: pass


class _ExcelHandle:
    def __init__(self):
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False
        self.excel.EnableEvents = False
        try: self.excel.Calculation = -4135
        except Exception: pass
        self.wb = None

    def open(self, path: Path):
        self.wb = self.excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        return self.wb

    def close(self) -> None:
        try:
            if self.wb is not None: self.wb.Save()
        except Exception: pass
        try:
            if self.wb is not None: self.wb.Close(SaveChanges=False)
        except Exception: pass
        try: self.excel.Calculation = -4105
        except Exception: pass
        try: self.excel.ScreenUpdating = True
        except Exception: pass
        try: self.excel.Quit()
        except Exception: pass
        del self.excel; del self.wb
        try:
            import pythoncom
            pythoncom.CoFreeUnusedLibraries()
        except Exception: pass


# =============================================================================
# Date helpers
# =============================================================================

def week_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())


def week_from_range_str(label: str) -> tuple[date, date] | None:
    """Parse 'WE MM-DD-YYYY' from a filename and return (monday, friday)."""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})", label)
    if not m:
        return None
    mm, dd, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    end = date(yyyy, mm, dd)
    monday = week_monday(end)
    return monday, monday + timedelta(days=4)


def _frn_date(frn: str) -> str:
    m = re.match(r"FRN(\d{2})(\d{2})(\d{2})-", frn or "", re.IGNORECASE)
    if not m:
        return ""
    yy, mm, dd = m.groups()
    try:
        return date(2000 + int(yy), int(mm), int(dd)).isoformat()
    except ValueError:
        return ""


def find_reel_summary(summary_root: Path, monday: date) -> Path | None:
    """Locate the Reel Summary for this week under <root>/YYYY/Month YYYY/Reel Summary."""
    for (y, m) in month_folders_for_week(monday):
        folder = summary_root / str(y) / month_folder_name(date(y, m, 1)) / "Reel Summary"
        if not folder.exists():
            continue
        candidates = [
            p for p in folder.glob("*.xls*")
            if not p.name.startswith("~$")
        ]
        if not candidates:
            continue
        m_str = f"{monday.month}.{monday.day}.{monday.strftime('%y')}"
        for p in candidates:
            if m_str in p.name:
                return p
        for p in candidates:
            if f"{monday.month:02d}" in p.name and monday.strftime("%y") in p.name:
                return p
        return max(candidates, key=lambda p: p.stat().st_mtime)
    return None


# =============================================================================
# Collectors
# =============================================================================

def collect_from_reel_summary(path: Path) -> list[LoadRow]:
    """Read the 'Summary' tab of a Reel Summary Report. Tare comes from the
    per-FRN tab (e.g. FRN261020-1R-AR)."""
    rw = _ReadOnlyWorkbook(path)
    try:
        if "Summary" not in rw.sheets():
            raise ValueError("Reel Summary has no 'Summary' tab.")
        all_sheets = set(rw.sheets())
        log(f"REEL summary tabs: {len(all_sheets)}")
        ws = rw.sheet("Summary")
        max_row = ws.UsedRange.Rows.Count

        def cell(row, col):
            try: return ws.Cells(row, col).Value
            except Exception: return None

        def as_iso(v) -> str:
            if isinstance(v, datetime): return v.date().isoformat()
            if isinstance(v, date): return v.isoformat()
            if isinstance(v, str) and v.strip():
                try: return datetime.strptime(v.strip()[:10], "%Y-%m-%d").date().isoformat()
                except Exception: pass
            return ""

        def is_class_header_row(row: int) -> bool:
            c3 = cell(row, 3)
            return isinstance(c3, str) and c3.strip().upper() == "CLASS"

        def is_class_values_row(row: int, header_row: int) -> bool:
            if row == header_row:
                return False
            numeric_hits = 0
            for c in range(4, 19):
                v = cell(row, c)
                try:
                    float(v)
                    numeric_hits += 1
                except (TypeError, ValueError):
                    pass
            return numeric_hits >= 2

        def find_label_value_in_sheet(sheet, want_upper: str) -> int | None:
            def fcell(r, c):
                try: return sheet.Cells(r, c).Value
                except Exception: return None
            for rr in range(1, 20):
                for cc in range(1, 20):
                    lbl = fcell(rr, cc)
                    if not isinstance(lbl, str):
                        continue
                    if want_upper not in lbl.upper():
                        continue
                    for off in range(1, 6):
                        v = fcell(rr, cc + off)
                        if v in (None, ""):
                            continue
                        try:
                            return int(float(v))
                        except Exception:
                            continue
            return None

        loads: list[LoadRow] = []
        r = 1
        while r <= max_row:
            a = cell(r, 2)
            if not (isinstance(a, str) and a.strip().upper() == "PL #"):
                r += 1
                continue

            frn_row = r + 1
            frn_val = cell(frn_row, 2)
            frn_str = str(frn_val).strip() if frn_val is not None else ""
            if not frn_str.startswith("FRN"):
                r += 1
                continue

            origin = cell(r, 5)
            origin = origin.strip().upper() if isinstance(origin, str) else ""
            if not origin:
                origin = "NO LOCATION"

            trucker = cell(r, 14)
            trucker = str(trucker).strip() if trucker is not None else ""
            trv = cell(r, 17)
            if trv is None:
                trailer = ""
            elif isinstance(trv, float) and trv.is_integer():
                trailer = str(int(trv))
            else:
                trailer = str(trv).strip()

            ritm = ""
            date_iso = ""
            gross = 0
            tare = 0
            scan = frn_row + 1
            class_header_row = None
            while scan <= max_row and scan < frn_row + 10:
                if is_class_header_row(scan):
                    class_header_row = scan
                    break
                rv = cell(scan, 2)
                if isinstance(rv, str):
                    if rv.strip().upper().startswith("RITM") and not ritm:
                        ritm = rv.strip()
                dv = cell(scan, 19)
                if dv is not None and not date_iso:
                    iso = as_iso(dv)
                    if iso:
                        date_iso = iso
                gv = cell(scan, 20)
                if gv is not None and not gross:
                    try: gross = int(float(gv))
                    except Exception: pass
                scan += 1

            if class_header_row is None:
                log(f"Summary: no class header near row {r} for {frn_str}; skipping")
                r += 1
                continue

            # Tare/gross live on the per-FRN tab (FRN + "R").
            frn_tab_name = frn_str + "R"
            if frn_tab_name in all_sheets:
                fws = rw.sheet(frn_tab_name)
                if not tare:
                    tare = find_label_value_in_sheet(fws, "TOTAL TARE") or 0
                if not gross:
                    gross = find_label_value_in_sheet(fws, "TOTAL GROSS") or 0
                log(f"Per-FRN {frn_tab_name}: gross={gross} tare={tare}")
            else:
                log(f"Per-FRN tab {frn_tab_name} not found; tare stays 0")

            class_values_row = class_header_row + 1
            if not is_class_values_row(class_values_row, class_header_row):
                class_values_row = class_header_row + 2

            class_nets: dict[str, int] = {}
            if class_values_row <= max_row:
                for c in range(4, 19):
                    lbl = cell(class_header_row, c)
                    val = cell(class_values_row, c)
                    if lbl is None:
                        continue
                    try:
                        w = int(float(val or 0))
                    except Exception:
                        w = 0
                    if w:
                        k = normalize_class(str(lbl))
                        if k == "99X" and "99" not in str(lbl) and "X" not in str(lbl).upper():
                            log(f"Unknown class label {lbl!r} on {frn_str}, mapped to 99X")
                        class_nets[k] = class_nets.get(k, 0) + w

            net = sum(class_nets.values())
            if not gross and net:
                gross = net + tare

            loads.append(LoadRow(
                source="reel", frn=frn_str, ritm=ritm, origin=origin,
                trucker=trucker, trailer=trailer, date=date_iso,
                gross=gross, tare=tare, net=net, class_nets=class_nets,
            ))
            r = class_values_row + 1
        return loads
    finally:
        rw.close()


def collect_from_jc_payment_statement(jc_file: Path) -> list[LoadRow]:
    """Read FRNs from every date tab of the JC Payment Statement."""
    loads: list[LoadRow] = []
    rw = _ReadOnlyWorkbook(jc_file)
    try:
        for sheet_name in rw.sheets():
            if not re.fullmatch(r"\d{2}_\d{2}_\d{2}", sheet_name):
                continue
            ws = rw.sheet(sheet_name)
            try: tab_date = datetime.strptime(sheet_name, DATE_TAB_FMT).date()
            except Exception: continue
            for cell in FOR_C_FRNS + VZ_FRNS:
                v = _read_cell(ws, cell)
                if v is None: continue
                frn = str(v).strip()
                if not frn.upper().startswith("FRN"): continue
                ritm_cell = ("C" if cell[0] == "B" else "E") + cell[1:]
                ritm_val = _read_cell(ws, ritm_cell)
                ritm = ""
                if ritm_val and str(ritm_val).strip().lower() != "no request item":
                    ritm = str(ritm_val).strip()
                loads.append(LoadRow(
                    source="jc", frn=frn, ritm=ritm, date=tab_date.isoformat(),
                ))
    finally:
        rw.close()
    return loads


def read_jc_prices(jc_file: Path) -> dict:
    """Read the four metal price cells from JC's 'Pmt Sum-Internal' tab."""
    rw = _ReadOnlyWorkbook(jc_file)
    try:
        if "Pmt Sum-Internal" not in rw.sheets(): return {}
        ws = rw.sheet("Pmt Sum-Internal")
        return {
            "copper": _read_cell(ws, "C7") or 0,
            "lead": _read_cell(ws, "C8") or 0,
            "gold": _read_cell(ws, "C9") or 0,
            "aluminum": _read_cell(ws, "C10") or 0,
        }
    finally:
        rw.close()


# =============================================================================
# Loose file parser
# =============================================================================

def _materials_in_squashed_block(text: str) -> list[str]:
    if not text:
        return []
    upper = text.upper()
    found: list[str] = []
    pos = 0
    while pos < len(upper):
        matched = False
        for token in _MATERIAL_TOKENS:
            tok_upper = token.upper()
            if tok_upper == "L":
                continue
            if upper.startswith(tok_upper, pos):
                after = pos + len(tok_upper)
                if after < len(upper) and upper[after].isalnum():
                    continue
                before_ok = pos == 0 or not upper[pos - 1].isalnum()
                if not before_ok:
                    continue
                found.append(normalize_class(token))
                pos = after
                matched = True
                break
        if not matched:
            pos += 1
    return found


def parse_loose_file(path: Path, excel=None) -> dict:
    own_excel = excel is None
    if own_excel:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
    wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
    try:
        ws = wb.Worksheets(1)

        def cell(row: int, col: int):
            try: return ws.Cells(row, col).Value
            except Exception: return None

        def as_str(v) -> str:
            if v is None: return ""
            if isinstance(v, (datetime, date)): return v.strftime("%Y-%m-%d")
            return str(v).strip()

        raw_date = cell(3, 7)
        if isinstance(raw_date, (datetime, date)):
            date_iso = raw_date.strftime("%Y-%m-%d")
        elif raw_date is not None:
            try:
                date_iso = datetime.strptime(str(raw_date).strip()[:10], "%Y-%m-%d").date().isoformat()
            except Exception:
                date_iso = ""
        else:
            date_iso = ""

        packing_list = as_str(cell(4, 7))
        ritm = as_str(cell(5, 7))
        if ritm.casefold() == "no request item":
            ritm = ""
        location = as_str(cell(9, 1)).upper()
        if not location:
            location = "NO LOCATION"

        def total(row: int) -> int:
            try: return int(float(cell(row, 7) or 0))
            except Exception: return 0
        total_gross = total(13)
        total_tare = total(14)
        total_net = total(15)

        items: list[dict] = []
        for row in range(10, 31):
            line_no = cell(row, 1)
            if line_no is None: continue
            try: int(line_no)
            except (TypeError, ValueError): continue
            code = as_str(cell(row, 2))
            material = as_str(cell(row, 3))
            item_type = as_str(cell(row, 4))
            gross = cell(row, 5)
            tare = cell(row, 6)
            if gross is None and tare is None: continue
            try:
                g = int(float(gross or 0)); t = int(float(tare or 0))
            except Exception:
                continue
            description = f"{code} {material}".strip() if material else code
            items.append({
                "description": description or "Loose Scrap",
                "type": material or item_type or "L",
                "gross": g, "tare": t, "net": g - t,
            })

        if not items and total_net > 0:
            squashed = ""
            for row in range(9, 12):
                v = as_str(cell(row, 1))
                if v:
                    squashed = v
                    break
            materials = _materials_in_squashed_block(squashed)
            if materials:
                per_mat = total_net // len(materials)
                remainder = total_net - per_mat * len(materials)
                for idx, mat in enumerate(materials):
                    net_here = per_mat + (remainder if idx == 0 else 0)
                    items.append({
                        "description": f"Loose {mat}",
                        "type": mat,
                        "gross": net_here,
                        "tare": 0,
                        "net": net_here,
                    })
            else:
                items.append({
                    "description": "Loose Scrap", "type": "L",
                    "gross": total_gross, "tare": total_tare, "net": total_net,
                })

        class_gross: dict[str, int] = {}
        class_tare: dict[str, int] = {}
        for it in items:
            cls = normalize_class(it.get("type") or "L")
            class_gross[cls] = class_gross.get(cls, 0) + int(it.get("gross", 0))
            class_tare[cls] = class_tare.get(cls, 0) + int(it.get("tare", 0))

        return {
            "date": date_iso,
            "packing_list": packing_list,
            "ritm": ritm,
            "location": location,
            "items": items,
            "class_gross": class_gross,
            "class_tare": class_tare,
            "total_gross": total_gross,
            "total_tare": total_tare,
            "total_net": total_net,
        }
    finally:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
        if own_excel:
            try: excel.Quit()
            except Exception: pass
            try:
                import pythoncom
                pythoncom.CoFreeUnusedLibraries()
            except Exception: pass


_LOOSE_FILE_FRN_RE = re.compile(r"FRN\d{6}-\d+L(?:-[A-Za-z]+)?", re.IGNORECASE)
_REEL_FILE_FRN_RE = re.compile(r"FRN\d{6}-\d+R(?:-[A-Za-z]+)*", re.IGNORECASE)


def collect_from_loose_folders(loose_root: Path, monday: date, excel=None) -> list[LoadRow]:
    week_prefixes = {
        f"FRN{(monday + timedelta(days=i)).strftime('%y%m%d')}-" for i in range(5)
    }
    loads: list[LoadRow] = []
    folders = [
        loose_root / str(y) / month_folder_name(date(y, m, 1))
        for (y, m) in month_folders_for_week(monday)
    ]
    folders = [f for f in folders if f.exists()]
    if not folders:
        log(f"Loose: no month folder found under {loose_root} for week of {monday}")
        return loads
    own_excel = excel is None
    if own_excel:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
    try:
        all_paths: list[Path] = []
        for folder in folders:
            all_paths.extend(folder.glob("*"))
        for path in all_paths:
            if not path.is_file(): continue
            if path.suffix.lower() not in (".xls", ".xlsx", ".xlsm"): continue
            if path.name.startswith("~$"): continue
            if _REEL_FILE_FRN_RE.search(path.name):
                log(f"Loose scan: skipping reel file {path.name}")
                continue
            m = _LOOSE_FILE_FRN_RE.search(path.name)
            if not m:
                log(f"Loose scan: skipping non-load file {path.name}")
                continue
            frn = m.group(0).upper()
            if not any(frn.startswith(p) for p in week_prefixes):
                log(f"Loose scan: skipping out-of-week file {path.name}")
                continue
            try:
                data = parse_loose_file(path, excel=excel)
                class_nets: dict[str, int] = {}
                for it in data.get("items", []):
                    cls = normalize_class(it.get("type") or "L")
                    class_nets[cls] = class_nets.get(cls, 0) + int(it.get("net", 0))
                date_iso = data.get("date", "") or _frn_date(frn)
                raw_origin = (data.get("location") or "").strip().upper()
                if not raw_origin:
                    raw_origin = "NO LOCATION"
                net = sum(class_nets.values())
                gross = data.get("total_gross", 0) or (net + data.get("total_tare", 0))
                tare = data.get("total_tare", 0)
                loads.append(LoadRow(
                    source="loose", frn=frn,
                    ritm=data.get("ritm", ""),
                    origin=raw_origin,
                    date=date_iso,
                    gross=gross,
                    tare=tare,
                    net=net,
                    class_nets=class_nets,
                    class_gross=dict(data.get("class_gross", {})),
                    class_tare=dict(data.get("class_tare", {})),
                    folder=str(path.parent),
                ))
            except Exception as e:
                log(f"Loose parse failed for {frn}: {e}")
    finally:
        if own_excel:
            try: excel.Quit()
            except Exception: pass
            try:
                import pythoncom
                pythoncom.CoFreeUnusedLibraries()
            except Exception: pass
    return loads


_SCRAP_PICKUP_FRN_RE = re.compile(r"FRN\d{6}-\d+R(?:-[A-Za-z]+)*", re.IGNORECASE)


def parse_scrap_pickup_file(path: Path, excel=None) -> dict:
    """Read the location + request-item status from a Scrap Pick Up Request .xls."""
    own_excel = excel is None
    if own_excel:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
    wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
    try:
        ws = wb.Worksheets(1)

        def cell(row: int, col: int):
            try: return ws.Cells(row, col).Value
            except Exception: return None

        def as_str(v) -> str:
            if v is None: return ""
            if isinstance(v, (datetime, date)): return v.strftime("%Y-%m-%d")
            return str(v).strip()

        frn = ""
        for r in range(1, 60):
            if as_str(cell(r, 1)).upper().startswith("LOAD #"):
                frn = as_str(cell(r, 4)).upper()
                break

        street = city = ""
        for r in range(1, 60):
            label = as_str(cell(r, 1)).strip().upper().rstrip(":")
            if label == "STREET" and not street:
                street = as_str(cell(r, 4))
            elif label == "CITY" and not city:
                city = as_str(cell(r, 4))
            if street and city:
                break

        log(f"Scrap pickup {path.name}: frn={frn!r} street={street!r} city={city!r}")

        parts = []
        for v in (street, city):
            if v and v.upper() != "NO LOCATION":
                parts.append(v.rstrip(", ").strip())
        location = ", ".join(parts).strip().upper()
        if not location:
            location = "NO LOCATION"

        ritm = ""
        req_label = as_str(cell(12, 6)).upper()
        no_req = as_str(cell(12, 7)).upper()
        x_in_f = as_str(cell(11, 6)).upper() == "X"
        x_in_e = as_str(cell(11, 5)).upper() == "X"
        if req_label.startswith("REQUEST ITEM") and x_in_f:
            ritm = "Request Item"
        elif no_req.startswith("NO REQUEST") and x_in_e:
            ritm = "No Request Item"

        return {"frn": frn, "location": location, "ritm": ritm}
    finally:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
        if own_excel:
            try: excel.Quit()
            except Exception: pass
            try:
                import pythoncom
                pythoncom.CoFreeUnusedLibraries()
            except Exception: pass


def collect_scrap_pickup_locations(root: Path, monday: date, excel=None) -> dict[str, dict]:
    """Scan the week's month folders for Scrap Pick Up Request files."""
    out: dict[str, dict] = {}
    if not root.exists():
        log(f"Scrap pickup folder does not exist: {root}")
        return out
    own_excel = excel is None
    if own_excel:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
    try:
        # Only scan the 1 or 2 month folders the week touches.
        month_dirs = [
            root / str(y) / month_folder_name(date(y, m, 1))
            for (y, m) in month_folders_for_week(monday)
        ]
        month_dirs = [d for d in month_dirs if d.exists()]
        if not month_dirs:
            log(f"Scrap pickup: no month folders found under {root} for week of {monday}")
            return out

        candidates: list[Path] = []
        for folder in month_dirs:
            candidates.extend(folder.rglob("*"))

        log(f"Scrap pickup scan: {len(month_dirs)} month folder(s), {len(candidates)} entries")
        for path in candidates:
            if not path.is_file(): continue
            if path.suffix.lower() not in (".xls", ".xlsx", ".xlsm"): continue
            if path.name.startswith("~$"): continue
            # Name filter: only Scrap Pick Up Request files hold the address.
            if "scrap pick up request" not in path.name.lower():
                continue
            log(f"Scrap pickup: parsing {path.name}")
            try:
                data = parse_scrap_pickup_file(path, excel=excel)
            except Exception as e:
                log(f"Scrap pickup parse failed for {path.name}: {e}")
                continue
            frn = (data.get("frn") or "").strip().upper()
            if not frn:
                m = _SCRAP_PICKUP_FRN_RE.search(path.name)
                if m: frn = m.group(0).upper()
            if not frn:
                log(f"Scrap pickup: no FRN in {path.name}; skipping")
                continue
            out[frn] = {"location": data.get("location", "NO LOCATION"),
                        "ritm": data.get("ritm", "")}
    finally:
        if own_excel:
            try: excel.Quit()
            except Exception: pass
            try:
                import pythoncom
                pythoncom.CoFreeUnusedLibraries()
            except Exception: pass
    return out


# =============================================================================
# Write into the workbook
# =============================================================================

_AUDIT_PATH: Path | None = None


def _rename_date_tabs(wb, monday: date) -> dict[str, str]:
    mapping = {}
    if wb.ProtectStructure:
        log("SKIP protected workbook structure: date tabs not renamed")
        return mapping
    existing = [s.Name for s in wb.Worksheets if re.fullmatch(r"\d{2}_\d{2}_\d{2}", s.Name)]
    ed = []
    for name in existing:
        try: ed.append((datetime.strptime(name, DATE_TAB_FMT).date(), name))
        except Exception: continue
    ed.sort()

    wk = [(d, n) for d, n in ed if d.weekday() < 5]
    we = [(d, n) for d, n in ed if d.weekday() >= 5]

    wanted_wk = [monday + timedelta(days=i) for i in range(5)]
    for (_, old), new_d in zip(wk, wanted_wk):
        new_name = new_d.strftime(DATE_TAB_FMT)
        if old != new_name:
            try:
                wb.Worksheets(old).Name = new_name
                mapping[old] = new_name
            except Exception as e:
                log(f"Rename failed {old}→{new_name}: {e}")

    # Saturday tab closes the prior week.
    saturday = monday - timedelta(days=2)
    if we:
        _, old = we[-1]
        new_name = saturday.strftime(DATE_TAB_FMT)
        if old != new_name:
            try:
                wb.Worksheets(old).Name = new_name
                mapping[old] = new_name
                log(f"Renamed Saturday tab {old}→{new_name}")
            except Exception as e:
                log(f"Rename failed {old}→{new_name}: {e}")
    else:
        log("No Saturday tab found to rename")

    return mapping


def _append_frn(ws, frn, ritm) -> int | None:
    """Write FRN/RITM to the first empty slot. Returns the row written, or None."""
    for cell in FOR_C_FRNS + VZ_FRNS:
        if _read_cell(ws, cell) not in (None, ""):
            continue
        ritm_cell = ("C" if cell[0] == "B" else "E") + cell[1:]
        if not _try_write(ws, cell, frn):
            continue
        _try_write(ws, ritm_cell, ritm or "No Request Item")
        try:
            return int(cell[1:])
        except Exception:
            return None
    return None


def _find_frn_addr(ws, frn: str) -> str | None:
    for addr in FOR_C_FRNS + VZ_FRNS:
        v = _read_cell(ws, addr)
        if v is not None and str(v).strip() == frn:
            return addr
    return None


def _audit_comment(ws, target: str, lr: LoadRow, class_nets: dict,
                   gross: int, tare: int, net: int):
    parts = " + ".join(f"{cls}:{w}" for cls, w in class_nets.items() if w)
    lines = [f"{lr.frn}  ({lr.date})"]
    if parts:
        lines.append(parts)
    if gross or tare or net:
        lines.append(f"gross {gross} | tare {tare} | net {net}")
    comment = "\n".join(lines)
    try:
        rng = ws.Range(target)
        try:
            if rng.Comment is not None:
                rng.Comment.Delete()
        except Exception:
            pass
        rng.AddComment(comment)
    except Exception as e:
        log(f"comment failed on {ws.Name}!{target}: {e}")


def _audit_log_line(lr: LoadRow, gross: int, tare: int, net: int,
                    class_nets: dict):
    if _AUDIT_PATH is None:
        return
    visible = {k: v for k, v in class_nets.items() if v}
    line = (
        f"{lr.date}\t{lr.frn}\t{lr.source}\t"
        f"gross+={gross}\ttare+={tare}\tnet+={net}\t"
        f"classes={visible}"
    )
    try:
        with open(_AUDIT_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        log(f"audit log write failed: {e}")


def _update_statistic_tab(wb, loads):
    if "Statistic" not in [s.Name for s in wb.Worksheets]: return
    ws = wb.Worksheets("Statistic")
    start = 5
    while _read_cell(ws, f"B{start}") not in (None, ""): start += 1
    r = start
    for load in loads:
        wrote = False
        if _try_write(ws, f"A{r}", load.date): wrote = True
        if _try_write(ws, f"B{r}", load.frn): wrote = True
        if _try_write(ws, f"C{r}", load.ritm or "DELIVERY"): wrote = True
        if _try_write(ws, f"G{r}", load.net): wrote = True
        landfill_wt = 0
        cn = load.class_nets or {}
        for cls in ("99X", "WOOD", "Loose Plastic", "Keyboard/Mouse", "Remote"):
            landfill_wt += int(cn.get(cls, 0) or 0)
        if _try_write(ws, f"H{r}", landfill_wt): wrote = True
        loc = (load.origin or "").strip().upper() or "NO LOCATION"
        if _try_write(ws, f"K{r}", loc): wrote = True
        if wrote: r += 1


def _fill_b5_first_tab_only(wb, monday: date) -> None:
    """Write B5 on the Monday tab only; every other tab's B5 chains from it."""
    epoch = date(1904, 1, 1) if wb.Date1904 else date(1899, 12, 30)
    tab_name = monday.strftime(DATE_TAB_FMT)
    try:
        ws = wb.Worksheets(tab_name)
    except Exception:
        log(f"B5: Monday tab {tab_name} not found")
        return
    try:
        _try_write(ws, "B5", (monday - epoch).days)
    except Exception as e:
        log(f"B5 write failed on {tab_name}: {e}")


def _verify_audit_totals():
    if _AUDIT_PATH is None or not _AUDIT_PATH.exists():
        return
    totals = {"gross": 0, "tare": 0, "net": 0}
    try:
        for line in _AUDIT_PATH.read_text(encoding="utf-8").splitlines():
            for k in totals:
                m = re.search(rf"{k}\+=(\d+)", line)
                if m:
                    totals[k] += int(m.group(1))
    except Exception as e:
        log(f"audit verify read failed: {e}")
        return
    log(f"AUDIT totals: {totals}")


def fill_payment_statement(source_template, loads, out_path, *, monday,
                           add_to_existing, jc_prices=None):
    global _AUDIT_PATH
    out_path.parent.mkdir(parents=True, exist_ok=True)
    _AUDIT_PATH = out_path.with_name(out_path.stem + "_loads.log")
    try:
        _AUDIT_PATH.write_text("", encoding="utf-8")
    except Exception as e:
        log(f"could not create audit log {_AUDIT_PATH}: {e}")
        _AUDIT_PATH = None

    tmp = out_path.with_suffix(".working.xls")
    shutil.copy2(source_template, tmp)
    excel = _ExcelHandle()
    try:
        wb = excel.open(tmp)

        _rename_date_tabs(wb, monday)
        _fill_b5_first_tab_only(wb, monday)

        by_date: dict[str, list[LoadRow]] = {}
        for lr in loads:
            by_date.setdefault(lr.date, []).append(lr)

        for iso, day in by_date.items():
            try:
                d = datetime.strptime(iso, "%Y-%m-%d").date()
            except Exception:
                log(f"Skipping load with bad date {iso}")
                continue
            try:
                ws = wb.Worksheets(d.strftime(DATE_TAB_FMT))
            except Exception:
                log(f"Date tab {d.strftime(DATE_TAB_FMT)} missing — skipping {iso}")
                continue

            day_c: dict[str, list[int]] = {}
            day_d: dict[str, list[int]] = {}
            return_tare_vals: list[int] = []

            for lr in day:
                class_nets = dict(lr.class_nets or {})
                if not class_nets and lr.net:
                    class_nets = {"99X": lr.net}

                # Reel loads carry a per-reel tare from the per-FRN tab.
                # Accumulate for the RETURN row's Gross column (C53).
                if lr.source == "reel" and lr.tare:
                    return_tare_vals.append(int(lr.tare))

                real_classes = [c for c in class_nets if c != "RETURN"]
                if not real_classes:
                    continue

                if lr.source == "loose" and (lr.class_gross or lr.class_tare):
                    for cls in real_classes:
                        g = int((lr.class_gross or {}).get(cls, 0) or 0)
                        t = int((lr.class_tare or {}).get(cls, 0) or 0)
                        if g:
                            day_c.setdefault(cls, []).append(g)
                        if t:
                            day_d.setdefault(cls, []).append(t)
                else:
                    for cls in real_classes:
                        n = int(class_nets.get(cls, 0) or 0)
                        if n:
                            day_c.setdefault(cls, []).append(n)

            # 1) Write FRNs + RITMs.
            for lr in day:
                row = _append_frn(ws, lr.frn, lr.ritm)
                if row is None:
                    log(f"{d.strftime(DATE_TAB_FMT)}: FRN slots full; {lr.frn} skipped")

            # 2) Per-class C/D.
            for cls, vals in day_c.items():
                row = CLASS_TO_ROW.get(cls) or CLASS_TO_ROW["99X"]
                address = f"C{row}"
                rng = ws.Range(address)
                formula = _additive_formula(rng.Formula, rng.Value, vals)
                if formula is not None:
                    _try_write(ws, address, formula)

            for cls, vals in day_d.items():
                row = CLASS_TO_ROW.get(cls) or CLASS_TO_ROW["99X"]
                address = f"D{row}"
                rng = ws.Range(address)
                formula = _additive_formula(rng.Formula, rng.Value, vals)
                if formula is not None:
                    _try_write(ws, address, formula)

            # 3) RETURN row: only C (Gross) gets reel tare, additive.
            ret_row = CLASS_TO_ROW["RETURN"]
            log(f"RETURN for {d.strftime(DATE_TAB_FMT)}: tares={return_tare_vals}")
            if return_tare_vals:
                rng = ws.Range(f"C{ret_row}")
                formula = _additive_formula(rng.Formula, rng.Value, return_tare_vals)
                if formula is not None:
                    _try_write(ws, f"C{ret_row}", formula)

            # 4) Comments + sidecar audit.
            for lr in day:
                frn_addr = _find_frn_addr(ws, lr.frn)
                class_nets = dict(lr.class_nets or {})
                if not class_nets and lr.net:
                    class_nets = {"99X": lr.net}
                if frn_addr:
                    _audit_comment(ws, frn_addr, lr, class_nets,
                                   lr.gross, lr.tare, lr.net)
                _audit_log_line(lr, lr.gross, lr.tare, lr.net, class_nets)

        _update_statistic_tab(wb, loads)

        if not add_to_existing and jc_prices:
            try:
                ws = wb.Worksheets("Pmt Sum-Internal")
                _try_write(ws, "C7", jc_prices.get("copper", 0))
                _try_write(ws, "C8", jc_prices.get("lead", 0))
                _try_write(ws, "C9", jc_prices.get("gold", 0))
                _try_write(ws, "C10", jc_prices.get("aluminum", 0))
            except Exception as e:
                log(f"Price copy failed: {e}")

        excel.excel.Calculation = -4105
        excel.excel.CalculateFullRebuild()
        excel.close()
        if tmp.exists(): tmp.replace(out_path)
        _verify_audit_totals()
        log(f"Wrote {out_path} ({len(loads)} loads)")
    except Exception:
        try: excel.close()
        finally: pass
        if tmp.exists():
            try: tmp.unlink()
            except Exception: pass
        raise


# =============================================================================
# Sort GUI
# =============================================================================

class _DragDropTable(QTableWidget):
    from PySide6.QtCore import Signal
    rowDropped = Signal(int, int)

    def __init__(self, rows: int, cols: int, parent=None):
        super().__init__(rows, cols, parent)
        self._drag_source_row = -1
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropOverwriteMode(False)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        # Reordering is handled by SortDialog's LoadRow list. Returning a
        # MoveAction makes Qt also remove the original table row afterward.
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.CopyAction)

    def startDrag(self, supported_actions):
        self._drag_source_row = self.currentRow()
        try:
            super().startDrag(supported_actions)
        finally:
            self._drag_source_row = -1

    def dropEvent(self, event):
        src = self._drag_source_row
        if src < 0:
            event.ignore()
            return

        pos = event.position().toPoint()
        row = self.rowAt(pos.y())
        if row < 0:
            insertion = self.rowCount()
        else:
            row_rect = self.visualRect(self.model().index(row, 0))
            insertion = row if pos.y() < row_rect.center().y() else row + 1

        destination = insertion - (1 if insertion > src else 0)
        if destination == src:
            event.ignore()
            return

        self.rowDropped.emit(src, destination)
        event.setDropAction(Qt.DropAction.CopyAction)
        event.accept()


class SortDialog(QDialog):
    def __init__(self, loads: list[LoadRow], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Order the loads, then write")
        self.resize(1100, 720)
        self.loads = list(loads)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Drag rows to reorder. Date is editable. "
            "Ctrl+R re-sorts reels + loose by date."
        ))

        self.table = _DragDropTable(0, 8)
        self.table.setHorizontalHeaderLabels([
            "Order", "Source", "FRN", "RITM", "Date (editable)",
            "Origin", "Trucker", "Net",
        ])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.rowDropped.connect(self._on_row_dropped)
        layout.addWidget(self.table, 1)

        self.status = QLabel("Ready.")
        self.status.setStyleSheet("color: #53645e; font-style: italic;")
        layout.addWidget(self.status)

        btn_row = QHBoxLayout()
        resort = QPushButton("Sort reels + loose by date (Ctrl+R)")
        resort.setMinimumWidth(260)
        resort.clicked.connect(self._resort_by_date)
        btn_row.addWidget(resort)
        btn_row.addStretch(1)
        done = QPushButton("Write both files")
        done.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white; padding: 8px 16px;"
        )
        done.clicked.connect(self._on_write)
        btn_row.addWidget(done)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        btn_row.addWidget(cancel)
        layout.addLayout(btn_row)

        self._shortcut = QShortcut(QKeySequence("Ctrl+R"), self)
        self._shortcut.activated.connect(self._resort_by_date)

        self._refresh()
        self.status.setText(f"Loaded {len(self.loads)} rows.")

    def _refresh(self):
        self.table.setUpdatesEnabled(False)
        try:
            self.table.setRowCount(0)
            for i, lr in enumerate(self.loads):
                self.table.insertRow(i)
                vals = [
                    str(i + 1), lr.source, lr.frn, lr.ritm, lr.date,
                    lr.origin, f"{lr.trucker} / {lr.trailer}".strip(" /"), str(lr.net),
                ]
                for c, v in enumerate(vals):
                    item = QTableWidgetItem(v)
                    if c == 4:
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                    else:
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.table.setItem(i, c, item)
        finally:
            self.table.setUpdatesEnabled(True)
            self.table.viewport().update()

    def _capture_date_edits(self):
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 4)
            if item is not None and i < len(self.loads):
                self.loads[i].date = item.text().strip()

    def _on_row_dropped(self, src: int, dst: int):
        if not (0 <= src < len(self.loads) and 0 <= dst < len(self.loads)):
            self.status.setText(f"Drop ignored (src={src}, dst={dst}).")
            return
        self._capture_date_edits()
        row = self.loads.pop(src)
        self.loads.insert(dst, row)
        self._refresh()
        self.table.selectRow(dst)
        self.status.setText(f"Moved row {src + 1} → position {dst + 1}.")

    def _resort_by_date(self):
        try:
            self._capture_date_edits()
            jc = [l for l in self.loads if l.source == "jc"]
            ours = [l for l in self.loads if l.source != "jc"]
            if not ours:
                self.status.setText("No reel/loose rows to sort.")
                return
            before = [(l.source, l.date, l.frn) for l in ours]
            ours.sort(key=lambda lr: (
                lr.date or _frn_date(lr.frn) or "9999-12-31",
                lr.source,
                lr.frn,
            ))
            after = [(l.source, l.date, l.frn) for l in ours]
            self.loads = jc + ours
            self._refresh()
            changed = "yes" if before != after else "no (already sorted)"
            self.status.setText(
                f"Sorted {len(ours)} reel/loose rows by date (interleaved). "
                f"Order changed: {changed}. JC kept on top."
            )
            log(f"RESORT: {len(ours)} rows, changed={changed}")
        except Exception as e:
            self.status.setText(f"Sort failed: {e}")
            log(f"RESORT ERROR: {e}")

    def _on_write(self):
        self._capture_date_edits()
        self.accept()

    def ordered(self):
        return self.loads


# =============================================================================
# Main window
# =============================================================================

class FillWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fill Verizon Payment Statement")
        self.resize(1000, 640)
        self.setStyleSheet(
            "QMainWindow, QWidget { background: #f2f5f3; color: #20332e; }"
            "QLabel { font-size: 10pt; }"
            "QLineEdit { background: white; border: 1px solid #9baea6; padding: 6px; }"
            "QPushButton { background: #e2e9e5; border: 1px solid #9baea6; padding: 6px 12px; }"
            "QPushButton:hover { background: #d5e1da; }"
        )
        cfg = load_cfg()
        central = QWidget(); self.setCentralWidget(central)
        v = QVBoxLayout(central)

        # 1. JC file
        v.addWidget(QLabel("1. JC Payment Statement (its filename's 'WE MM-DD-YYYY' drives everything):"))
        row = QHBoxLayout()
        self.jc_input = QLineEdit(cfg.get("jc_file", "")); self.jc_input.setReadOnly(True)
        row.addWidget(self.jc_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_jc); row.addWidget(b)
        v.addLayout(row)

        # 2. Summary root
        v.addWidget(QLabel("2. Receiving root for Reel Summaries (expects YYYY/Month YYYY/Reel Summary/):"))
        row = QHBoxLayout()
        self.summary_input = QLineEdit(cfg.get("summary_root", "")); self.summary_input.setReadOnly(True)
        row.addWidget(self.summary_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_summary); row.addWidget(b)
        v.addLayout(row)

        # 3. Loose root
        v.addWidget(QLabel("3. Receiving root for loose files (expects YYYY/Month YYYY/):"))
        row = QHBoxLayout()
        self.loose_input = QLineEdit(cfg.get("loose_root", "")); self.loose_input.setReadOnly(True)
        row.addWidget(self.loose_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_loose); row.addWidget(b)
        v.addLayout(row)

        # 4. Blank template
        v.addWidget(QLabel("4. Blank template (for the Rahway-only statement):"))
        row = QHBoxLayout()
        self.template_input = QLineEdit(cfg.get("blank_template", "")); self.template_input.setReadOnly(True)
        row.addWidget(self.template_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_template); row.addWidget(b)
        v.addLayout(row)

        # 5. Scrap pickup root
        v.addWidget(QLabel("5. Scrap Pickup root (contains Scrap Pick Up Request files):"))
        row = QHBoxLayout()
        self.pickup_input = QLineEdit(cfg.get("pickup_root", "")); self.pickup_input.setReadOnly(True)
        row.addWidget(self.pickup_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_pickup); row.addWidget(b)
        v.addLayout(row)

        # 6. Output root (Merged)
        v.addWidget(QLabel("6. Export folder for MERGED file:"))
        row = QHBoxLayout()
        self.output_input = QLineEdit(cfg.get("output_root", "")); self.output_input.setReadOnly(True)
        row.addWidget(self.output_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_output); row.addWidget(b)
        v.addLayout(row)

        # 7. Output root (Just Us / Rahway)
        v.addWidget(QLabel("7. Export folder for RAHWAY-ONLY file:"))
        row = QHBoxLayout()
        self.output_rahway_input = QLineEdit(cfg.get("output_rahway", "")); self.output_rahway_input.setReadOnly(True)
        row.addWidget(self.output_rahway_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_output_rahway); row.addWidget(b)
        v.addLayout(row)

        self.go_button = QPushButton("Collect & sort…")
        self.go_button.setStyleSheet(
            "font-weight: bold; background: #315b52; color: white; padding: 10px 20px;"
        )
        self.go_button.clicked.connect(self._go)
        v.addWidget(self.go_button)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.status.setStyleSheet("color: #53645e;")
        v.addWidget(self.status)
        v.addStretch(1)

    def _browse_jc(self):
        f, _ = QFileDialog.getOpenFileName(self, "JC Payment Statement", "", "Excel (*.xls *.xlsx *.xlsm)")
        if f: self.jc_input.setText(f); cfg = load_cfg(); cfg["jc_file"] = f; save_cfg(cfg)

    def _browse_summary(self):
        f = QFileDialog.getExistingDirectory(self, "Receiving root (summaries)")
        if f: self.summary_input.setText(f); cfg = load_cfg(); cfg["summary_root"] = f; save_cfg(cfg)

    def _browse_loose(self):
        f = QFileDialog.getExistingDirectory(self, "Receiving root (loose)")
        if f: self.loose_input.setText(f); cfg = load_cfg(); cfg["loose_root"] = f; save_cfg(cfg)

    def _browse_template(self):
        f, _ = QFileDialog.getOpenFileName(self, "Blank template", "", "Excel (*.xls *.xlsx *.xlsm)")
        if f: self.template_input.setText(f); cfg = load_cfg(); cfg["blank_template"] = f; save_cfg(cfg)

    def _browse_pickup(self):
        f = QFileDialog.getExistingDirectory(self, "Scrap Pickup root")
        if f: self.pickup_input.setText(f); cfg = load_cfg(); cfg["pickup_root"] = f; save_cfg(cfg)

    def _browse_output(self):
        f = QFileDialog.getExistingDirectory(self, "Export folder for MERGED file")
        if f: self.output_input.setText(f); cfg = load_cfg(); cfg["output_root"] = f; save_cfg(cfg)

    def _browse_output_rahway(self):
        f = QFileDialog.getExistingDirectory(self, "Export folder for RAHWAY file")
        if f: self.output_rahway_input.setText(f); cfg = load_cfg(); cfg["output_rahway"] = f; save_cfg(cfg)

    def _go(self):
        jc = Path(self.jc_input.text())
        tpl = Path(self.template_input.text())
        summary_root = Path(self.summary_input.text()) if self.summary_input.text() else None
        loose_root = Path(self.loose_input.text()) if self.loose_input.text() else None
        output_root = Path(self.output_input.text()) if self.output_input.text() else None
        output_rahway = Path(self.output_rahway_input.text()) if self.output_rahway_input.text() else None
        pickup_root = Path(self.pickup_input.text()) if self.pickup_input.text() else None

        if not (jc.exists() and tpl.exists()):
            QMessageBox.warning(self, "Missing files", "Pick the JC file and blank template."); return
        if summary_root is None or not summary_root.exists():
            QMessageBox.warning(self, "Missing folder", "Pick the Receiving root for Reel Summaries."); return
        if loose_root is None or not loose_root.exists():
            QMessageBox.warning(self, "Missing folder", "Pick the Receiving root for loose files."); return
        if output_root is None or not output_root.exists():
            QMessageBox.warning(self, "Missing folder", "Pick the export folder for the Merged file."); return
        if output_rahway is None or not output_rahway.exists():
            QMessageBox.warning(self, "Missing folder", "Pick the export folder for the Rahway-only file."); return

        week = week_from_range_str(jc.name)
        if week is None:
            QMessageBox.warning(self, "Bad file name",
                "JC file name must contain WE MM-DD-YYYY."); return
        monday, friday = week

        rr = find_reel_summary(summary_root, monday)
        if rr is None:
            QMessageBox.warning(self, "Reel Summary not found",
                f"No Reel Summary found under {summary_root} for week of {monday:%m/%d/%Y}."); return
        log(f"Derived Reel Summary: {rr}")

        try:
            shared = win32com.client.DispatchEx("Excel.Application")
            shared.ScreenUpdating = False
            shared.DisplayAlerts = False
            shared.EnableEvents = False
            import time
            try:
                t0 = time.perf_counter()
                jc_loads = collect_from_jc_payment_statement(jc)
                log(f"TIMING jc: {time.perf_counter()-t0:.2f}s")
                t0 = time.perf_counter()
                reel_loads = collect_from_reel_summary(rr)
                log(f"TIMING reel: {time.perf_counter()-t0:.2f}s")
                t0 = time.perf_counter()
                loose_loads = collect_from_loose_folders(loose_root, monday, excel=shared)
                log(f"TIMING loose: {time.perf_counter()-t0:.2f}s")
            finally:
                try: shared.Quit()
                except Exception: pass
                try:
                    import pythoncom
                    pythoncom.CoFreeUnusedLibraries()
                except Exception: pass
        except Exception as e:
            QMessageBox.critical(self, "Collect failed", str(e)); return

        pickup_map: dict[str, dict] = {}
        if pickup_root is not None and pickup_root.exists():
            shared = win32com.client.DispatchEx("Excel.Application")
            shared.ScreenUpdating = False
            shared.DisplayAlerts = False
            shared.EnableEvents = False
            try:
                pickup_map = collect_scrap_pickup_locations(pickup_root, monday, excel=shared)
            finally:
                try: shared.Quit()
                except Exception: pass
                try:
                    import pythoncom
                    pythoncom.CoFreeUnusedLibraries()
                except Exception: pass

        for lr in reel_loads + loose_loads:
            info = pickup_map.get(lr.frn.strip().upper())
            if info:
                pickup_location = (info.get("location") or "").strip().upper()
                if pickup_location and pickup_location != "NO LOCATION":
                    lr.origin = pickup_location
                if info.get("ritm"):
                    lr.ritm = info["ritm"]
            elif not lr.origin:
                lr.origin = "NO LOCATION"

        def _date_key(lr: LoadRow):
            return (lr.date or _frn_date(lr.frn) or "9999-12-31", lr.source, lr.frn)
        ours_loads = sorted(reel_loads + loose_loads, key=_date_key)
        all_loads = jc_loads + ours_loads

        self.status.setText(
            f"Collected: {len(jc_loads)} JC (file order), "
            f"{len(reel_loads)} reel, {len(loose_loads)} loose "
            f"(interleaved by date). Opening sort dialog…"
        )
        dlg = SortDialog(all_loads, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            self.status.setText("Cancelled."); return
        ordered = dlg.ordered()
        ours_only = [l for l in ordered if l.source != "jc"]

        try:
            # Write directly to the chosen folders; no date subfolders created.
            out_a = output_root / f"Verizon BAU Payment Statement WE {friday:%m-%d-%Y} (merged).xls"
            out_b = output_rahway / f"Verizon BAU NJ Payment Statement WE {friday:%m-%d-%Y}.xls"
            
            jc_prices = read_jc_prices(jc)
            
            # Write Merged
            fill_payment_statement(jc, ours_only, out_a, monday=monday,
                                   add_to_existing=True)
            # Write Rahway-only
            fill_payment_statement(tpl, ours_only, out_b, monday=monday,
                                   add_to_existing=False, jc_prices=jc_prices)
        except Exception as e:
            QMessageBox.critical(self, "Write failed", str(e)); return
        self.status.setText(
            f"Wrote merged ({len(ordered)} loads, includes JC): {out_a.name}\n"
            f"Wrote Rahway-only ({len(ours_only)} loads, excludes JC): {out_b.name}\n"
            f"Merged folder: {output_root}\n"
            f"Rahway folder: {output_rahway}"
        )
        QMessageBox.information(self, "Done",
            f"Merged (JC + ours, {len(ordered)} loads):\n  {out_a}\n\n"
            f"Rahway-only (ours only, {len(ours_only)} loads):\n  {out_b}\n\n"
            f"Per-load audit logs:\n  {out_a.stem}_loads.log\n  {out_b.stem}_loads.log")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = FillWindow()
    w.show()
    sys.exit(app.exec())