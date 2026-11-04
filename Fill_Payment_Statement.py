# Fill_Payment_Statement.py
"""
Collect loads for a week, drag-to-reorder, and fill the Verizon BAU Payment
Statement.

Inputs:
  1. JC Payment Statement            (.xls)  — received from Jersey City
  2. Reel Summary Report             (.xls)  — our reel loads
  3. Receiving root folder           — contains loose .xls files named
                                       like "1012-FRN261012-5L.xls"
  4. Blank template                  (.xls)  — for the Rahway-only statement

Outputs:
  A) JC-merged   = JC file + our reels + our loose   (email this)
  B) Rahway-only = blank template + our reels + our loose  (print this)

Reads/writes .xls and .xlsx via Excel COM. Skips protected cells silently.
JC keeps file order; reels and loose interleave by date.

Audit trail:
  - Each FRN cell on a date tab gets a comment showing what was added.
  - A sidecar "<output>_loads.log" records one line per load.
  - The final totals are checked against the audit sum in fill_statement.log.
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
FIRST_CLASS_ROW = 14
LAST_CLASS_ROW = 56
TOTAL_ROW = 57

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
    try:
        rng.Value = value
        return True
    except Exception as e:
        msg = str(e).lower()
        if "protected" in msg or "read-only" in msg:
            log(f"SKIP protected: {ws.Name}!{address}")
            return False
        raise


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


# =============================================================================
# Collectors
# =============================================================================

def collect_from_reel_summary(path: Path) -> list[LoadRow]:
    """Read the 'Summary' tab of a Reel Summary Report."""
    rw = _ReadOnlyWorkbook(path)
    try:
        if "Summary" not in rw.sheets():
            raise ValueError("Reel Summary has no 'Summary' tab.")
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
                tv = cell(scan, 21)
                if tv is not None and not tare:
                    try: tare = int(float(tv))
                    except Exception: pass
                scan += 1

            if class_header_row is None:
                log(f"Summary: no class header near row {r} for {frn_str}; skipping")
                r += 1
                continue

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


def parse_loose_file(path: Path) -> dict:
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

        return {
            "date": date_iso,
            "packing_list": packing_list,
            "ritm": ritm,
            "location": location,
            "items": items,
            "total_gross": total_gross,
            "total_tare": total_tare,
            "total_net": total_net,
        }
    finally:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
        try: excel.Quit()
        except Exception: pass
        try:
            import pythoncom
            pythoncom.CoFreeUnusedLibraries()
        except Exception: pass


_LOOSE_FILE_FRN_RE = re.compile(r"FRN\d{6}-\d+L(?:-[A-Za-z]+)?", re.IGNORECASE)
_REEL_FILE_FRN_RE = re.compile(r"FRN\d{6}-\d+R(?:-[A-Za-z]+)*", re.IGNORECASE)


def collect_from_loose_folders(root: Path, monday: date) -> list[LoadRow]:
    week_prefixes = {
        f"FRN{(monday + timedelta(days=i)).strftime('%y%m%d')}-" for i in range(5)
    }
    loads: list[LoadRow] = []
    if not root.exists():
        return loads
    for path in root.rglob("*"):
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
            data = parse_loose_file(path)
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
                folder=str(path.parent),
            ))
        except Exception as e:
            log(f"Loose parse failed for {path.name}: {e}")
    return loads


# =============================================================================
# Write into the workbook
# =============================================================================

_AUDIT_PATH: Path | None = None  # CHANGED: set in fill_payment_statement


def _rename_date_tabs(wb, monday: date) -> dict[str, str]:
    mapping = {}
    existing = [s.Name for s in wb.Worksheets if re.fullmatch(r"\d{2}_\d{2}_\d{2}", s.Name)]
    wanted = [monday + timedelta(days=i) for i in range(5)]
    ed = []
    for name in existing:
        try: ed.append((datetime.strptime(name, DATE_TAB_FMT).date(), name))
        except Exception: continue
    ed.sort()
    wk = [(d, n) for d, n in ed if d.weekday() < 5]
    for (_, old), new_d in zip(wk, wanted):
        new_name = new_d.strftime(DATE_TAB_FMT)
        if old != new_name:
            try:
                wb.Worksheets(old).Name = new_name
                mapping[old] = new_name
            except Exception as e:
                log(f"Rename failed {old}→{new_name}: {e}")
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


def _add_class_weight(ws, class_key, net_weight,
                     gross_weight=0, tare_weight=0,
                     *, is_first_class_for_load=False):
    """Add weight to a class row. Returns (gross_delta, tare_delta, net_delta)
    actually written, so the caller can log the per-load contribution.

    Gross/Tare are only written on the *first* class row for a load so the
    column totals reconcile to one gross/tare per scale ticket.
    """
    row = CLASS_TO_ROW.get(class_key) or CLASS_TO_ROW["99X"]

    def bump(col, amount):
        if amount == 0:
            return 0
        addr = f"{col}{row}"
        cur = _read_cell(ws, addr)
        try: cur = int(float(cur or 0))
        except Exception: cur = 0
        _try_write(ws, addr, cur + int(amount))
        return int(amount)

    net_d = bump("E", net_weight)
    gross_d = bump("C", gross_weight) if is_first_class_for_load else 0
    tare_d = bump("D", tare_weight) if is_first_class_for_load else 0
    return gross_d, tare_d, net_d


def _recompute_total(ws):
    for col in ("C", "D", "E"):
        total = 0
        for r in range(FIRST_CLASS_ROW, LAST_CLASS_ROW + 1):
            v = _read_cell(ws, f"{col}{r}")
            try: total += int(float(v or 0))
            except Exception: pass
        _try_write(ws, f"{col}{TOTAL_ROW}", total)


def _find_frn_addr(ws, frn: str) -> str | None:
    """Locate which of B7:B11 / D7:D11 holds this FRN."""
    for addr in FOR_C_FRNS + VZ_FRNS:
        v = _read_cell(ws, addr)
        if v is not None and str(v).strip() == frn:
            return addr
    return None


def _audit_comment(ws, target: str, lr: LoadRow, class_nets: dict,
                   gross: int, tare: int, net: int):
    """Drop a cell comment on the FRN cell showing what was added."""
    parts = []
    if gross: parts.append(f"Gross {gross}")
    if tare: parts.append(f"Tare {tare}")
    if net: parts.append(f"Net {net}")
    detail = " + ".join(f"{cls}:{w}" for cls, w in class_nets.items() if w)
    comment = f"{lr.frn}  ({lr.date})"
    if parts:
        comment += "\n" + " | ".join(parts)
    if detail:
        comment += f"\nclasses: {detail}"
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
    """Append one tab-separated line to the sidecar log (if configured)."""
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
        if _try_write(ws, f"C{r}", load.ritm or "No Request Item"): wrote = True
        if _try_write(ws, f"G{r}", load.net): wrote = True
        landfill_wt = int((load.class_nets or {}).get("99X", 0) or 0)
        if _try_write(ws, f"H{r}", landfill_wt): wrote = True
        loc = (load.origin or "").strip().upper() or "NO LOCATION"
        if _try_write(ws, f"K{r}", loc): wrote = True
        if wrote: r += 1


def _fill_b5_per_date_tab(wb, monday: date) -> None:
    tab_dates = [monday + timedelta(days=i) for i in range(5)]
    for d in tab_dates:
        tab_name = d.strftime(DATE_TAB_FMT)
        try:
            ws = wb.Worksheets(tab_name)
        except Exception:
            continue
        try:
            _try_write(
                ws, "B5",
                datetime.combine(d, datetime.min.time()),
            )
            ws.Range("B5").NumberFormat = "mm/dd/yyyy"
        except Exception as e:
            log(f"B5 write failed on {tab_name}: {e}")


def _verify_audit_totals():
    """Read the sidecar log and log its grand totals for cross-checking."""
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
        _AUDIT_PATH.write_text("", encoding="utf-8")  # truncate per run
    except Exception as e:
        log(f"could not create audit log {_AUDIT_PATH}: {e}")
        _AUDIT_PATH = None

    tmp = out_path.with_suffix(".working.xls")
    shutil.copy2(source_template, tmp)
    excel = _ExcelHandle()
    try:
        wb = excel.open(tmp)
        _rename_date_tabs(wb, monday)
        _fill_b5_per_date_tab(wb, monday)
        by_date: dict[str, list[LoadRow]] = {}
        for lr in loads: by_date.setdefault(lr.date, []).append(lr)
        for iso, day in by_date.items():
            try: d = datetime.strptime(iso, "%Y-%m-%d").date()
            except Exception:
                log(f"Skipping load with bad date {iso}"); continue
            try: ws = wb.Worksheets(d.strftime(DATE_TAB_FMT))
            except Exception:
                log(f"Date tab {d.strftime(DATE_TAB_FMT)} missing — skipping {iso}"); continue
            for lr in day:
                row = _append_frn(ws, lr.frn, lr.ritm)
                if row is None:
                    log(f"{d.strftime(DATE_TAB_FMT)}: FRN slots full; {lr.frn} skipped")
                    continue

                class_nets = lr.class_nets or {}
                if not class_nets and lr.net:
                    class_nets = {"99X": lr.net}

                # CHANGED: accumulate what this load actually wrote, and
                # record it on the FRN cell + sidecar log.
                load_gross = load_tare = load_net = 0
                first = True
                for cls, w in class_nets.items():
                    g, t, n = _add_class_weight(
                        ws, cls, w,
                        gross_weight=lr.gross if first else 0,
                        tare_weight=lr.tare if first else 0,
                        is_first_class_for_load=first,
                    )
                    load_gross += g
                    load_tare += t
                    load_net += n
                    first = False

                frn_addr = _find_frn_addr(ws, lr.frn)
                if frn_addr:
                    _audit_comment(ws, frn_addr, lr, class_nets,
                                   load_gross, load_tare, load_net)
                _audit_log_line(lr, load_gross, load_tare, load_net, class_nets)
            _recompute_total(ws)

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
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropOverwriteMode(False)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)

    def dropEvent(self, event):
        src = self.currentRow()
        if src < 0:
            event.ignore(); return
        pos = event.position().toPoint()
        dst = self.rowAt(pos.y())
        if dst < 0:
            dst = self.rowCount() - 1
        if dst == src:
            event.ignore(); return
        self.rowDropped.emit(src, dst)
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
        if dst > src:
            dst -= 1
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
        self.resize(900, 480)
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

        v.addWidget(QLabel("1. JC Payment Statement (received):"))
        row = QHBoxLayout()
        self.jc_input = QLineEdit(cfg.get("jc_file", "")); self.jc_input.setReadOnly(True)
        row.addWidget(self.jc_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_jc); row.addWidget(b)
        v.addLayout(row)

        v.addWidget(QLabel("2. Reel Summary Report (our reels):"))
        row = QHBoxLayout()
        self.rr_input = QLineEdit(cfg.get("rr_file", "")); self.rr_input.setReadOnly(True)
        row.addWidget(self.rr_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_rr); row.addWidget(b)
        v.addLayout(row)

        v.addWidget(QLabel("3. Loose folder (contains 1012-FRN261012-5L.xls style files):"))
        row = QHBoxLayout()
        self.loose_input = QLineEdit(cfg.get("loose_root", "")); self.loose_input.setReadOnly(True)
        row.addWidget(self.loose_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_loose); row.addWidget(b)
        v.addLayout(row)

        v.addWidget(QLabel("4. Blank template (for the Rahway-only statement):"))
        row = QHBoxLayout()
        self.template_input = QLineEdit(cfg.get("blank_template", "")); self.template_input.setReadOnly(True)
        row.addWidget(self.template_input, 1)
        b = QPushButton("Browse…"); b.clicked.connect(self._browse_template); row.addWidget(b)
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

    def _browse_rr(self):
        f, _ = QFileDialog.getOpenFileName(self, "Reel Summary Report", "", "Excel (*.xls *.xlsx *.xlsm)")
        if f: self.rr_input.setText(f); cfg = load_cfg(); cfg["rr_file"] = f; save_cfg(cfg)

    def _browse_loose(self):
        f = QFileDialog.getExistingDirectory(self, "Loose folder")
        if f: self.loose_input.setText(f); cfg = load_cfg(); cfg["loose_root"] = f; save_cfg(cfg)

    def _browse_template(self):
        f, _ = QFileDialog.getOpenFileName(self, "Blank template", "", "Excel (*.xls *.xlsx *.xlsm)")
        if f: self.template_input.setText(f); cfg = load_cfg(); cfg["blank_template"] = f; save_cfg(cfg)

    def _go(self):
        jc = Path(self.jc_input.text()); rr = Path(self.rr_input.text())
        root = Path(self.loose_input.text()); tpl = Path(self.template_input.text())
        if not (jc.exists() and rr.exists() and root.exists() and tpl.exists()):
            QMessageBox.warning(self, "Missing files", "Fill in all four paths."); return
        try:
            week = week_from_range_str(jc.name)
            if week is None:
                QMessageBox.warning(self, "Bad file name",
                    "JC file name must contain WE MM-DD-YYYY."); return
            monday, friday = week
            jc_loads = collect_from_jc_payment_statement(jc)
            reel_loads = collect_from_reel_summary(rr)
            loose_loads = collect_from_loose_folders(root, monday)
        except Exception as e:
            QMessageBox.critical(self, "Collect failed", str(e)); return

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
            out_a = jc.with_name(f"Verizon BAU Payment Statement WE {friday:%m-%d-%Y} (merged).xls")
            out_b = jc.with_name(f"Verizon BAU NJ Payment Statement WE {friday:%m-%d-%Y}.xls")
            jc_prices = read_jc_prices(jc)
            fill_payment_statement(jc, ours_only, out_a, monday=monday,
                                   add_to_existing=True)
            fill_payment_statement(tpl, ours_only, out_b, monday=monday,
                                   add_to_existing=False, jc_prices=jc_prices)
        except Exception as e:
            QMessageBox.critical(self, "Write failed", str(e)); return
        self.status.setText(
            f"Wrote merged ({len(ordered)} loads, includes JC): {out_a.name}\n"
            f"Wrote Rahway-only ({len(ours_only)} loads, excludes JC): {out_b.name}\n"
            f"Audit logs: {out_a.stem}_loads.log, {out_b.stem}_loads.log"
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