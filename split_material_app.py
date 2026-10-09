"""
Standalone app for splitting Verizon material weights across a family of
packing-list loads.

Usage:
  python split_material_app.py
"""

import json
import os
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtGui import QTextCursor

import win32com.client


# ---------------------------------------------------------------------------
# Config / paths
# ---------------------------------------------------------------------------

CONFIG_PATH = Path(__file__).resolve().parent / "splitter_config.json"
EXCEL_BASES = Path(__file__).resolve().parent / "Excel Bases"
SPLIT_MATERIAL_TEMPLATE = EXCEL_BASES / "Split Material Weighted.xls"
DEFAULT_SUMMARIES_DIR = Path(__file__).resolve().parent / "Summaries"
DEFAULT_SPLIT_DIR = Path(__file__).resolve().parent / "Split Material"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_config(data: dict) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

def parse_frn_date(frn: str) -> date:
    match = re.search(r"FRN(\d{2})(\d{2})(\d{2})", frn, re.IGNORECASE)
    if not match:
        raise ValueError(
            f"Could not parse a date from '{frn}'. Expected something like "
            "FRN260929-1R."
        )
    yy, mm, dd = match.groups()
    return date(2000 + int(yy), int(mm), int(dd))


def _frn_order_key(identifier: str) -> tuple[date, int, str, str]:
    match = re.fullmatch(
        r"FRN(\d{6})(?:-(\d+)R)?(?:-([A-Za-z]+))?(?:[RP])?",
        identifier.strip(),
        re.IGNORECASE,
    )
    if not match:
        return date.max, 0, "", identifier.casefold()

    try:
        frn_date = parse_frn_date(f"FRN{match.group(1)}")
    except ValueError:
        return date.max, 0, "", identifier.casefold()
    sequence = int(match.group(2) or 0)
    family_letter = (match.group(3) or "").upper()
    return frn_date, sequence, family_letter, identifier.casefold()


def week_range(d: date) -> tuple[date, date]:
    weekday = d.weekday()
    if weekday >= 5:
        monday = d + timedelta(days=(7 - weekday))
    else:
        monday = d - timedelta(days=weekday)
    friday = monday + timedelta(days=4)
    return monday, friday


def summary_filename_for_frn(frn: str) -> str:
    d = parse_frn_date(frn)
    mon, fri = week_range(d)
    return (
        f"Reel Summary Report WE "
        f"{mon.month}.{mon.day}.{mon.strftime('%y')} ~ "
        f"{fri.month}.{fri.day}.{fri.strftime('%y')}.xls"
    )


# ---------------------------------------------------------------------------
# Material name normalization
# ---------------------------------------------------------------------------

# Material comparison: spaces and case never matter. Everything is
# uppercased and whitespace-stripped before comparing. Two extra equivalences
# are applied because these pairs mean the same thing in practice:
#   99FO == FO99
#   (any other reorderings of letters/digits are NOT merged — only this one.)
ALIASES = {
    "99FO": "99FO",
    "FO99": "99FO",
}


def norm(s: str) -> str:
    """Canonical key for comparing material names. Spaces and case are
    ignored; the 99FO/FO99 pair is folded together."""
    if s is None:
        return ""
    text = str(s)
    key = re.sub(r"\s+", "", text).upper()
    return ALIASES.get(key, key)


def normalize_material(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        f = float(text)
        if f.is_integer():
            return str(int(f))
        return str(f)
    except (ValueError, TypeError):
        return text


def split_material_cell(value) -> list[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    parts = []
    for part in text.split(","):
        nm = normalize_material(part)
        if nm:
            parts.append(nm)
    return parts


def canon_material_cell(value) -> str:
    return ", ".join(split_material_cell(value))


# ---------------------------------------------------------------------------
# Solver
# ---------------------------------------------------------------------------

MIN_WEIGHT = 10

try:
    from ortools.sat.python import cp_model
    HAS_ORTOOLS = True
except ImportError:
    HAS_ORTOOLS = False


def solve_cp(bundles, remaining, step, totals, time_limit=15.0):
    model = cp_model.CpModel()
    min_units = -(-MIN_WEIGHT // step)
    x = {}
    for bi, b in enumerate(bundles):
        for it in b["items"]:
            max_units = min(remaining[it] // step, b["total"] // step)
            if max_units < min_units:
                return None
            x[(bi, it)] = model.new_int_var(min_units, max_units, f"x_{bi}_{it}")
    for bi, b in enumerate(bundles):
        model.add(sum(x[(bi, it)] * step for it in b["items"]) == b["total"])
    for item, need in remaining.items():
        contributors = [
            x[(bi, it)] * step
            for bi, b in enumerate(bundles)
            for it in b["items"]
            if it == item
        ]
        if contributors:
            model.add(sum(contributors) == need)
        else:
            if need != 0:
                return None
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None
    result = []
    for bi, b in enumerate(bundles):
        weights = {it: solver.Value(x[(bi, it)]) * step for it in b["items"]}
        result.append({"load": b["load"], "items": b["items"], "weights": weights})
    return result


ITER_LIMIT = 500_000
_iter_count = 0
_partition_cache = {}


def partitions_capped_uncached(total, caps, step, min_weight=MIN_WEIGHT):
    k = len(caps)
    floor = ((min_weight + step - 1) // step) * step
    if k == 1:
        if floor <= total <= caps[0] and total % step == 0:
            yield (total,)
        return
    lo = floor
    hi = min(caps[0], total - floor * (k - 1))
    if hi < lo:
        return
    for first in range(lo, hi + 1, step):
        for rest in partitions_capped_uncached(total - first, caps[1:], step, min_weight):
            yield (first,) + rest


def partitions_capped_cached(total, caps, step, min_weight=MIN_WEIGHT):
    key = (total, tuple(caps), step, min_weight)
    cached = _partition_cache.get(key)
    if cached is not None:
        return cached
    result = tuple(partitions_capped_uncached(total, caps, step, min_weight))
    _partition_cache[key] = result
    return result


def solve_recursive(bundles, remaining, step, totals):
    global _iter_count
    assignment = []
    remaining = dict(remaining)

    def recurse(idx):
        global _iter_count
        if idx == len(bundles):
            return all(v == 0 for v in remaining.values())
        _iter_count += 1
        if _iter_count > ITER_LIMIT:
            raise RuntimeError("iteration limit reached")
        b = bundles[idx]
        items = b["items"]
        caps = tuple(remaining[it] for it in items)
        remaining_sum = sum(remaining.values())
        needed = sum(b2["total"] for b2 in bundles[idx:])
        if remaining_sum < needed:
            return False
        full_caps = tuple(totals[it] for it in items)
        for combo in partitions_capped_cached(b["total"], full_caps, step):
            if any(w > c for w, c in zip(combo, caps)):
                continue
            for it, w in zip(items, combo):
                remaining[it] -= w
            assignment.append(
                {"load": b["load"], "items": items, "weights": dict(zip(items, combo))}
            )
            if recurse(idx + 1):
                return True
            assignment.pop()
            for it, w in zip(items, combo):
                remaining[it] += w
        return False

    return assignment if recurse(0) else None


# ---------------------------------------------------------------------------
# Reel Summary reading
# ---------------------------------------------------------------------------

def _is_p_tab_name(name: str) -> bool:
    return re.fullmatch(r"\d+P(?:\s*\(\d+\))?", name, re.IGNORECASE) is not None


def _family_letter_from_tab(name: str, frn: str) -> str | None:
    if not name.upper().startswith(frn.upper()):
        return None
    m = re.search(r"-([A-Za-z]+)[RP]$", name)
    return m.group(1).upper() if m else None


def _sort_family_tabs(workbook) -> None:
    sheets = [sheet for sheet in workbook.Worksheets]
    ordered_sheets = list(sheets)
    positions_by_kind: dict[str, list[int]] = {"R": [], "P": []}
    family_sheets_by_kind: dict[str, list[tuple[str, object]]] = {
        "R": [],
        "P": [],
    }

    for position, sheet in enumerate(sheets):
        match = re.fullmatch(
            r"(FRN\d{6})(?:-(\d+)R)?-([A-Za-z]+)([RP])",
            sheet.Name.strip(),
            re.IGNORECASE,
        )
        if not match:
            continue
        kind = match.group(4).upper()
        positions_by_kind[kind].append(position)
        sequence = match.group(2)
        identifier = f"{match.group(1)}"
        if sequence:
            identifier += f"-{sequence}R"
        identifier += f"-{match.group(3)}"
        family_sheets_by_kind[kind].append((identifier, sheet))

    for kind, positions in positions_by_kind.items():
        sorted_sheets = [
            sheet
            for _, sheet in sorted(
                family_sheets_by_kind[kind],
                key=lambda entry: _frn_order_key(entry[0]),
            )
        ]
        for position, sheet in zip(positions, sorted_sheets):
            ordered_sheets[position] = sheet

    for position, expected_sheet in enumerate(ordered_sheets, start=1):
        current_sheet = workbook.Worksheets(position)
        if current_sheet.Name != expected_sheet.Name:
            expected_sheet.Move(Before=current_sheet)


@dataclass
class _ItemRow:
    reel_id: str
    reel_size: str
    material: str
    net: int
    original_material: str = ""


@dataclass
class _LoadRow:
    letter: str
    identifier: str
    items: list[_ItemRow] = field(default_factory=list)
    bundle_total: int = 0    # Total Net, summed from column I
    total_tare: int = 0      # Total Tare, summed from column H
    ritm: str = ""           # RITM # (R tab M5)
    driver: str = ""         # Trucker value (R tab G8)
    plate: str = ""          # Trailer value (R tab G9)


def read_family_r_loads(path: Path, frn: str) -> list[_LoadRow]:
    loads: list[_LoadRow] = []
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        for sheet in workbook.Worksheets:
            name = sheet.Name.strip()
            if not name.upper().endswith("R"):
                continue
            letter = _family_letter_from_tab(name, frn)
            if letter is None:
                continue

            items: list[_ItemRow] = []
            sum_net = 0
            sum_tare = 0
            for row in range(12, 43):
                reel_id_raw = sheet.Cells(row, 4).Value   # D: Reel ID
                reel_id = str(reel_id_raw or "").strip()
                if not reel_id:
                    continue
                reel_size = str(sheet.Cells(row, 5).Value or "").strip()   # E

                raw_material = sheet.Cells(row, 6).Value                   # F: Class
                parts = split_material_cell(raw_material)
                material = ", ".join(parts)

                gross_raw = sheet.Cells(row, 7).Value or 0                  # G: Gross
                tare_raw = sheet.Cells(row, 8).Value or 0                   # H: Tare
                try:
                    gross = int(round(float(gross_raw)))
                except (TypeError, ValueError):
                    gross = 0
                try:
                    tare = int(round(float(tare_raw)))
                except (TypeError, ValueError):
                    tare = 0
                net = gross - tare

                sum_tare += tare
                sum_net += net

                items.append(_ItemRow(
                    reel_id=reel_id,
                    reel_size=reel_size,
                    material=material,
                    net=net,
                    original_material=material,
                ))

            total_net = sum_net
            total_tare = sum_tare

            ritm = str(sheet.Range("M5").Value or "").strip()
            driver = str(sheet.Range("G8").Value or "").strip()
            plate = str(sheet.Range("G9").Value or "").strip()

            loads.append(_LoadRow(
                letter=letter,
                identifier=f"{frn}-{letter}",
                items=items,
                bundle_total=total_net,
                total_tare=total_tare,
                ritm=ritm,
                driver=driver,
                plate=plate,
            ))
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    loads.sort(key=lambda l: l.letter)
    return loads


# ---------------------------------------------------------------------------
# Split Material Weighted writer
# ---------------------------------------------------------------------------

def write_split_material_workbook(
    load: _LoadRow,
    load_date: date,
    weights: dict[str, int],
    destination: Path,
) -> Path:
    template = SPLIT_MATERIAL_TEMPLATE
    if not template.exists():
        raise FileNotFoundError(
            f"Split Material Weighted template not found: {template}"
        )
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, destination)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(destination), UpdateLinks=0, ReadOnly=False)
        try:
            sheet = workbook.Worksheets("Packing List")
        except Exception:
            sheet = workbook.Worksheets(1)

        sheet.Range("H8").Value = datetime.combine(load_date, datetime.min.time())
        sheet.Range("H8").NumberFormat = "mm/dd/yyyy"
        sheet.Range("C9").Value = "VERIZON"
        sheet.Range("C10").Value = load.ritm or ""
        cont_number = ""
        if load.driver or load.plate:
            cont_number = f"{load.driver} / {load.plate}".strip(" /")
        sheet.Range("H9").Value = cont_number
        sheet.Range("H10").Value = load.identifier

        for i in range(23):
            row = 13 + i
            sheet.Cells(row, 3).Value = ""
            sheet.Cells(row, 8).Value = ""
            sheet.Cells(row, 9).Value = ""
            sheet.Cells(row, 10).Value = ""

        sorted_items = sorted(weights.items(), key=lambda kv: kv[0])
        for i, (material, weight) in enumerate(sorted_items):
            if i >= 23:
                raise ValueError(f"{load.identifier} has more than 23 material rows.")
            row = 13 + i
            sheet.Cells(row, 3).Value = normalize_material(material)
            sheet.Cells(row, 8).Value = weight
            sheet.Cells(row, 9).Value = 0
            sheet.Cells(row, 10).Value = weight

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    return destination


# ---------------------------------------------------------------------------
# R tab write-back (only changed Class cells)
# ---------------------------------------------------------------------------

def write_back_r_tabs(
    summary_path: Path,
    frn: str,
    loads: list[_LoadRow],
) -> int:
    updated = 0
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(summary_path), UpdateLinks=0, ReadOnly=False)
        for load in loads:
            target_sheet = None
            for sheet in workbook.Worksheets:
                name = sheet.Name.strip()
                if not name.upper().endswith("R"):
                    continue
                letter = _family_letter_from_tab(name, frn)
                if letter == load.letter:
                    target_sheet = sheet
                    break
            if target_sheet is None:
                continue

            row = 12
            for item in load.items:
                while row <= 42 and not str(target_sheet.Cells(row, 4).Value or "").strip():
                    row += 1
                if row > 42:
                    break
                new_canon = canon_material_cell(item.material)
                old_canon = canon_material_cell(item.original_material)
                if new_canon != old_canon:
                    target_sheet.Cells(row, 6).Value = new_canon
                    updated += 1
                row += 1

        if updated:
            workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    return updated


# ---------------------------------------------------------------------------
# P tab filler
# ---------------------------------------------------------------------------

P_FIRST_ITEM_ROW = 22
P_LAST_ITEM_ROW = 41
P_MAX_ITEMS = P_LAST_ITEM_ROW - P_FIRST_ITEM_ROW + 1   # 20


def fill_p_tabs(
    summary_path: Path,
    frn: str,
    per_load_splits: dict[str, dict[str, int]],
    loads: list[_LoadRow],
    load_date: date,
) -> None:
    load_by_letter = {ld.letter: ld for ld in loads}

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(summary_path), UpdateLinks=0, ReadOnly=False)

        family_p_tabs: dict[str, object] = {}
        unused_p_tabs: list[object] = []
        for sheet in workbook.Worksheets:
            name = sheet.Name.strip()
            letter = _family_letter_from_tab(name, frn)
            if letter is not None and name.upper().endswith("P"):
                family_p_tabs[letter] = sheet
            elif _is_p_tab_name(name):
                unused_p_tabs.append(sheet)

        needed_letters = sorted(per_load_splits.keys())
        missing = [L for L in needed_letters if L not in family_p_tabs]
        if len(missing) > len(unused_p_tabs):
            raise ValueError(
                f"Need {len(missing)} new P tab(s) but only "
                f"{len(unused_p_tabs)} unused P tabs available."
            )
        for i, letter in enumerate(missing):
            family_p_tabs[letter] = unused_p_tabs[i]

        for letter in needed_letters:
            sheet = family_p_tabs[letter]
            weights = per_load_splits[letter]
            load = load_by_letter.get(letter)
            total_tare = load.total_tare if load else 0

            print(f"[fill_p_tabs] {frn}-{letter}P  weights={weights!r}  "
                  f"total_tare={total_tare}")

            for row in range(P_FIRST_ITEM_ROW, P_LAST_ITEM_ROW + 1):
                for col in (3, 4, 5, 6):
                    sheet.Cells(row, col).Value = ""

            new_name = f"{frn}-{letter}P"
            if len(new_name) > 31:
                raise ValueError(f"Tab name too long: {new_name}")
            sheet.Name = new_name

            sheet.Range("G2").Value = f"{frn}-{letter}"
            sheet.Range("G4").Value = datetime.combine(load_date, datetime.min.time())
            sheet.Range("G4").NumberFormat = "mm/dd/yyyy"

            sorted_items = sorted(weights.items(), key=lambda kv: kv[0])
            if len(sorted_items) > P_MAX_ITEMS - 1:
                raise ValueError(
                    f"{new_name} has {len(sorted_items)} materials; "
                    f"at most {P_MAX_ITEMS - 1} are allowed."
                )

            row = P_FIRST_ITEM_ROW
            for material, weight in sorted_items:
                sheet.Cells(row, 3).Value = normalize_material(material)
                sheet.Cells(row, 4).Value = ""
                sheet.Cells(row, 5).Value = weight
                sheet.Cells(row, 6).Value = 0
                row += 1

            if row <= P_LAST_ITEM_ROW:
                sheet.Cells(row, 3).Value = "Other"
                sheet.Cells(row, 4).Value = "Empty Reels Return"
                sheet.Cells(row, 5).Value = total_tare
                sheet.Cells(row, 6).Value = total_tare

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


# ---------------------------------------------------------------------------
# Summary tab fill
# ---------------------------------------------------------------------------

def fill_summary_tab(
    summary_path: Path,
    frn: str,
    loads: list[_LoadRow],
) -> int:
    MAX_SLOTS = 20
    FIRST_ROW = 8
    STRIDE = 6

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    workbook = None
    written = 0
    try:
        workbook = excel.Workbooks.Open(str(summary_path), UpdateLinks=0, ReadOnly=False)
        try:
            sheet = workbook.Worksheets("Summary")
        except Exception:
            raise ValueError("No 'Summary' tab found in the workbook.")

        identifiers: list[str] = []
        seen_identifiers: set[str] = set()
        for i in range(MAX_SLOTS):
            row = FIRST_ROW + STRIDE * i
            value = sheet.Range(f"B{row}").Value
            text = str(value).strip() if value not in (None, "") else ""
            if text:
                identifiers.append(text)
                seen_identifiers.add(text.casefold())

        for ld in loads:
            key = ld.identifier.casefold()
            if key not in seen_identifiers:
                identifiers.append(ld.identifier)
                seen_identifiers.add(key)

        identifiers.sort(key=_frn_order_key)
        if len(identifiers) > MAX_SLOTS:
            raise ValueError(
                f"Summary tab has {MAX_SLOTS} PL # slots but "
                f"{len(identifiers)} load entries are present."
            )

        for slot, identifier in enumerate(identifiers):
            row = FIRST_ROW + STRIDE * slot
            sheet.Range(f"B{row}").Value = identifier

        _sort_family_tabs(workbook)
        written = len(loads)

        if written:
            workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    return written


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class SolverWorker(QThread):
    done = Signal(object, str)

    def __init__(self, totals, raw_bundles, start_step, min_weight):
        super().__init__()
        self.totals = totals
        self.raw_bundles = raw_bundles
        self.start_step = start_step
        self.min_weight = min_weight

    def run(self):
        global MIN_WEIGHT
        try:
            old = MIN_WEIGHT
            MIN_WEIGHT = self.min_weight
            result = self._do_solve()
            MIN_WEIGHT = old
            self.done.emit(result, "")
        except Exception as e:
            self.done.emit(None, str(e))

    def _do_solve(self):
        remaining = dict(self.totals)
        bundles = []
        locked = []
        for load, items, total in self.raw_bundles:
            if not items:
                raise ValueError(f"Load {load} has no materials.")
            if len(items) == 1:
                item = items[0]
                if item not in remaining:
                    raise ValueError(f"Unknown material '{item}' in load {load}.")
                remaining[item] -= total
                locked.append({"load": load, "items": items, "weights": {item: total}})
            else:
                bundles.append({"load": load, "items": items, "total": total})
        for k, v in remaining.items():
            if v < 0:
                raise ValueError(f"'{k}' is over budget by {-v}.")
        bundles.sort(key=lambda b: (len(b["items"]), -b["total"]))
        steps = [s for s in (1, 5, 10) if s <= self.start_step]

        result = None
        used_step = None
        if HAS_ORTOOLS:
            for step in steps:
                try:
                    result = solve_cp(bundles, remaining, step, self.totals, time_limit=15.0)
                except Exception:
                    result = None
                if result is not None:
                    used_step = step
                    break
        else:
            for step in (5, 1):
                if step > self.start_step:
                    continue
                _partition_cache.clear()
                global _iter_count
                _iter_count = 0
                try:
                    result = solve_recursive(bundles, remaining, step, self.totals)
                except RuntimeError:
                    result = None
                if result is not None:
                    used_step = step
                    break

        if result is None:
            raise ValueError("No solution found.")

        all_bundles = locked + result
        per_load: dict[str, dict[str, int]] = {}
        for b in all_bundles:
            load = b["load"]
            per_load.setdefault(load, {})
            for it, w in b["weights"].items():
                per_load[load][it] = per_load[load].get(it, 0) + w
        return per_load


class LoadItemsWidget(QWidget):
    """One load's items: Reel ID, Size, Class (editable), Net."""

    def __init__(self, load: _LoadRow, parent=None):
        super().__init__(parent)
        self.load = load

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        header = QHBoxLayout()
        header.addWidget(QLabel(f"Load: {load.identifier}"))
        if load.ritm:
            header.addWidget(QLabel(f"  RITM: {load.ritm}"))
        header.addStretch(1)
        header.addWidget(QLabel("Total Net (from R tab):"))
        self.total_label = QLabel(str(load.bundle_total))
        self.total_label.setStyleSheet("font-weight: bold;")
        header.addWidget(self.total_label)
        header.addWidget(QLabel("  Total Tare:"))
        self.tare_label = QLabel(str(load.total_tare))
        self.tare_label.setStyleSheet("font-weight: bold;")
        header.addWidget(self.tare_label)
        layout.addLayout(header)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Reel ID", "Reel Size", "Class", "Net"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table)
        self._populate()

    def _populate(self):
        self.table.setRowCount(0)
        for item in self.load.items:
            r = self.table.rowCount()
            self.table.insertRow(r)

            reel_id_item = QTableWidgetItem(item.reel_id)
            reel_id_item.setFlags(reel_id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 0, reel_id_item)

            size_item = QTableWidgetItem(item.reel_size)
            size_item.setFlags(size_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 1, size_item)

            self.table.setItem(r, 2, QTableWidgetItem(item.material))

            net_item = QTableWidgetItem(str(item.net))
            net_item.setFlags(net_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(r, 3, net_item)

    def collect(self) -> list[_ItemRow]:
        result: list[_ItemRow] = []
        for r in range(self.table.rowCount()):
            reel_id = self.table.item(r, 0).text() if self.table.item(r, 0) else ""
            reel_size = self.table.item(r, 1).text() if self.table.item(r, 1) else ""
            material_raw = self.table.item(r, 2).text() if self.table.item(r, 2) else ""
            net_text = self.table.item(r, 3).text() if self.table.item(r, 3) else "0"
            try:
                net = int(net_text) if net_text else 0
            except ValueError:
                net = 0
            material = material_raw.strip()
            original = self.load.items[r].original_material if r < len(self.load.items) else material
            result.append(_ItemRow(
                reel_id=reel_id,
                reel_size=reel_size,
                material=material,
                net=net,
                original_material=original,
            ))
        return result


class SplitMaterialWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Split Material Weights")
        self.resize(1150, 900)

        self.config = load_config()
        self.summary_path: Path | None = None
        self.loads: list[_LoadRow] = []
        self.load_widgets: list[LoadItemsWidget] = []

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 18, 18, 18)

        # --- Summaries folder row ---
        top = QHBoxLayout()
        top.addWidget(QLabel("Summary folder:"))
        self.folder_label = QLabel(
            self.config.get("summary_folder") or str(DEFAULT_SUMMARIES_DIR)
        )
        self.folder_label.setStyleSheet("color: #53645e;")
        self.folder_label.setWordWrap(True)
        top.addWidget(self.folder_label, 1)
        pick_btn = QToolButton()
        pick_btn.setText("📁 Choose…")
        pick_btn.setToolTip("Choose the folder where Reel Summary Reports live")
        pick_btn.setStyleSheet("font-weight: bold; padding: 6px 10px;")
        pick_btn.clicked.connect(self._pick_folder)
        top.addWidget(pick_btn)
        root.addLayout(top)

        # --- Split Material folder row (Option B) ---
        split_row = QHBoxLayout()
        self.split_same_checkbox = QCheckBox(
            "Split Material files go next to the Summary Report"
        )
        self.split_same_checkbox.setChecked(
            not bool(self.config.get("split_folder"))
        )
        self.split_same_checkbox.toggled.connect(self._toggle_split_same)
        split_row.addWidget(self.split_same_checkbox)

        self.split_folder_label = QLabel(
            self.config.get("split_folder") or "(uses Summary folder)"
        )
        self.split_folder_label.setStyleSheet("color: #53645e;")
        self.split_folder_label.setWordWrap(True)
        split_row.addWidget(self.split_folder_label, 1)

        self.split_pick_btn = QToolButton()
        self.split_pick_btn.setText("📁 Choose…")
        self.split_pick_btn.setToolTip(
            "Choose where the Split Material Weighted workbooks go"
        )
        self.split_pick_btn.setStyleSheet("font-weight: bold; padding: 6px 10px;")
        self.split_pick_btn.clicked.connect(self._pick_split_folder)
        self.split_pick_btn.setEnabled(not self.split_same_checkbox.isChecked())
        split_row.addWidget(self.split_pick_btn)
        root.addLayout(split_row)

        # --- FRN row ---
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("FRN:"))
        self.frn_input = QLineEdit()
        self.frn_input.setPlaceholderText("FRN260929-1R")
        self.frn_input.returnPressed.connect(self._load_family)
        row2.addWidget(self.frn_input, 1)
        load_btn = QPushButton("Load Family")
        load_btn.clicked.connect(self._load_family)
        row2.addWidget(load_btn)
        root.addLayout(row2)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)

        opt_row = QHBoxLayout()
        opt_row.addWidget(QLabel("Step:"))
        self.step_combo = QComboBox()
        self.step_combo.addItems(["1", "5", "10"])
        self.step_combo.setCurrentText("1")
        opt_row.addWidget(self.step_combo)
        opt_row.addWidget(QLabel("Min weight:"))
        self.min_weight_spin = QSpinBox()
        self.min_weight_spin.setRange(1, 1000)
        self.min_weight_spin.setValue(3)
        opt_row.addWidget(self.min_weight_spin)
        opt_row.addStretch(1)
        self.run_button = QPushButton("▶  Run Solver && Write Files")
        self.run_button.setStyleSheet("font-weight: bold; padding: 6px 20px;")
        self.run_button.clicked.connect(self._run)
        opt_row.addWidget(self.run_button)
        root.addLayout(opt_row)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setStyleSheet("font-family: Consolas, monospace;")
        self.output.setMaximumHeight(180)
        root.addWidget(self.output)

        mats_tab = QWidget()
        mats_layout = QVBoxLayout(mats_tab)
        self.mats_table = QTableWidget(0, 2)
        self.mats_table.setHorizontalHeaderLabels(["Material", "Total Weight"])
        self.mats_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.mats_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        mats_layout.addWidget(self.mats_table)
        mats_btns = QHBoxLayout()
        add_mat = QPushButton("+ Add Material")
        add_mat.clicked.connect(lambda: self._add_material_row("", ""))
        del_mat = QPushButton("− Remove Selected")
        del_mat.clicked.connect(self._remove_material_row)
        refresh_mat = QPushButton("↻ Rebuild from Loads")
        refresh_mat.clicked.connect(self._rebuild_materials_from_loads)
        mats_btns.addWidget(add_mat)
        mats_btns.addWidget(del_mat)
        mats_btns.addWidget(refresh_mat)
        mats_btns.addStretch(1)
        mats_layout.addLayout(mats_btns)
        self.materials_tab = mats_tab
        self.tabs.addTab(mats_tab, "Materials")

    # ------------------------------------------------------------------
    # Folder pickers
    # ------------------------------------------------------------------

    def _pick_folder(self) -> None:
        start = self.config.get("summary_folder") or str(DEFAULT_SUMMARIES_DIR)
        DEFAULT_SUMMARIES_DIR.mkdir(parents=True, exist_ok=True)
        chosen = QFileDialog.getExistingDirectory(
            self,
            "Choose the folder containing Reel Summary Reports",
            start,
        )
        if not chosen:
            return
        self.config["summary_folder"] = chosen
        save_config(self.config)
        self.folder_label.setText(chosen)

    def _toggle_split_same(self, checked: bool) -> None:
        # Checked -> Split Material files go next to the loaded summary.
        # Unchecked -> use the picked folder.
        self.split_pick_btn.setEnabled(not checked)
        if checked:
            self.config.pop("split_folder", None)
            save_config(self.config)
            self.split_folder_label.setText("(uses Summary folder)")
        else:
            saved = self.config.get("split_folder")
            if saved and Path(saved).is_dir():
                self.split_folder_label.setText(saved)
            else:
                self._pick_split_folder()

    def _pick_split_folder(self) -> None:
        start = self.config.get("split_folder") or str(DEFAULT_SPLIT_DIR)
        DEFAULT_SPLIT_DIR.mkdir(parents=True, exist_ok=True)
        chosen = QFileDialog.getExistingDirectory(
            self,
            "Choose where the Split Material Weighted workbooks go",
            start,
        )
        if not chosen:
            # User cancelled; if nothing is saved, put the checkbox back on
            # so we don't leave a half-set state.
            if not self.config.get("split_folder"):
                self.split_same_checkbox.setChecked(True)
            return
        self.config["split_folder"] = chosen
        save_config(self.config)
        self.split_folder_label.setText(chosen)
        self.split_same_checkbox.setChecked(False)

    # ------------------------------------------------------------------
    # Summary lookup
    # ------------------------------------------------------------------

    def _resolve_summary(self, frn: str) -> Path | None:
        target_name = summary_filename_for_frn(frn)
        target_cf = target_name.casefold()

        folder = self.config.get("summary_folder")
        if folder and Path(folder).is_dir():
            candidate = Path(folder) / target_name
            if candidate.exists():
                return candidate
            for f in Path(folder).iterdir():
                if f.name.casefold() == target_cf:
                    return f

        DEFAULT_SUMMARIES_DIR.mkdir(parents=True, exist_ok=True)
        candidate = DEFAULT_SUMMARIES_DIR / target_name
        if candidate.exists():
            return candidate
        for f in DEFAULT_SUMMARIES_DIR.iterdir():
            if f.name.casefold() == target_cf:
                return f

        script_dir = Path(__file__).resolve().parent
        for f in script_dir.iterdir():
            if f.is_file() and f.name.casefold() == target_cf:
                return f

        return None

    def _load_family(self) -> None:
        frn = self.frn_input.text().strip()
        if not frn:
            QMessageBox.warning(self, "Missing FRN", "Type the FRN first.")
            return
        path = self._resolve_summary(frn)
        if path is None:
            QMessageBox.warning(
                self, "Summary not found",
                f"Could not find:\n{summary_filename_for_frn(frn)}\n\n"
                f"In: {self.config.get('summary_folder') or DEFAULT_SUMMARIES_DIR}"
            )
            return
        self.summary_path = path
        try:
            loads = read_family_r_loads(path, frn)
        except Exception as e:
            QMessageBox.critical(self, "Read failed", str(e))
            return

        if not loads:
            QMessageBox.warning(
                self, "No loads found",
                f"No R tabs starting with '{frn}' were found in:\n{path}",
            )
            return

        self.loads = loads
        self._rebuild_tabs()
        self._rebuild_materials_from_loads()
        self._log(f"Loaded {len(loads)} load(s) from {path.name}.")
        for ld in loads:
            self._log(
                f"  {ld.identifier}: {len(ld.items)} item(s)  "
                f"net={ld.bundle_total}  tare={ld.total_tare}  "
                f"ritm={ld.ritm or '(none)'}"
            )

    def _rebuild_tabs(self) -> None:
        while self.tabs.count() > 1:
            widget = self.tabs.widget(0)
            self.tabs.removeTab(0)
            if widget is not None:
                widget.deleteLater()

        self.load_widgets = []
        for ld in self.loads:
            w = LoadItemsWidget(ld)
            self.load_widgets.append(w)
            self.tabs.insertTab(self.tabs.count() - 1, w, ld.identifier)

    def _rebuild_materials_from_loads(self) -> None:
        # Preserve whatever the user already had in the Materials tab,
        # keyed by the canonical form so spaces/case don't matter.
        existing: dict[str, str] = {}
        for r in range(self.mats_table.rowCount()):
            name_item = self.mats_table.item(r, 0)
            weight_item = self.mats_table.item(r, 1)
            if not name_item or not weight_item:
                continue
            raw_name = name_item.text().strip()
            if not raw_name:
                continue
            existing[norm(raw_name)] = weight_item.text().strip()

        display_name: dict[str, str] = {}
        seen_order: list[str] = []
        for w in self.load_widgets:
            for item in w.collect():
                for part in split_material_cell(item.material):
                    if not part:
                        continue
                    key = norm(part)
                    if key not in display_name:
                        display_name[key] = normalize_material(part)
                        seen_order.append(key)

        for key in existing:
            if key not in display_name:
                display_name[key] = key
                seen_order.append(key)

        self.mats_table.setRowCount(0)
        for key in seen_order:
            self._add_material_row(display_name[key], existing.get(key, ""))

    def _add_material_row(self, name="", weight="") -> None:
        r = self.mats_table.rowCount()
        self.mats_table.insertRow(r)
        self.mats_table.setItem(r, 0, QTableWidgetItem(str(name)))
        self.mats_table.setItem(r, 1, QTableWidgetItem(str(weight)))

    def _remove_material_row(self) -> None:
        rows = sorted({idx.row() for idx in self.mats_table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.mats_table.removeRow(r)

    def _log(self, text: str) -> None:
        self.output.insertPlainText(text + "\n")
        self.output.moveCursor(QTextCursor.MoveOperation.End)

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------

    def _run(self) -> None:
        self.output.clear()

        if not self.loads or self.summary_path is None:
            QMessageBox.warning(self, "Nothing loaded", "Load the family first.")
            return

        frn = self.frn_input.text().strip()
        try:
            load_date = parse_frn_date(frn)
        except ValueError as e:
            QMessageBox.critical(self, "Bad FRN", str(e))
            return

        updated_loads: list[_LoadRow] = []
        for w in self.load_widgets:
            items = w.collect()
            ld = _LoadRow(
                letter=w.load.letter,
                identifier=w.load.identifier,
                items=items,
                bundle_total=w.load.bundle_total,
                total_tare=w.load.total_tare,
                ritm=w.load.ritm,
                driver=w.load.driver,
                plate=w.load.plate,
            )
            updated_loads.append(ld)
        self.loads = updated_loads

        totals: dict[str, int] = {}
        for r in range(self.mats_table.rowCount()):
            name_item = self.mats_table.item(r, 0)
            weight_item = self.mats_table.item(r, 1)
            if not name_item or not weight_item:
                continue
            name = normalize_material(name_item.text())
            weight_text = weight_item.text().strip()
            if not name or not weight_text:
                continue
            try:
                w_val = int(weight_text)
            except ValueError:
                QMessageBox.critical(self, "Bad material weight",
                                     f"Material row {r+1}: '{weight_text}' is not an integer.")
                return
            if w_val <= 0:
                QMessageBox.critical(self, "Bad material weight",
                                     f"Material row {r+1}: weight must be positive.")
                return
            totals[norm(name)] = w_val

        if not totals:
            QMessageBox.warning(self, "No materials", "Add at least one material.")
            return

        raw_bundles: list[tuple[str, list[str], int]] = []
        for ld in self.loads:
            if ld.bundle_total <= 0:
                continue
            mats: list[str] = []
            for item in ld.items:
                for part in split_material_cell(item.material):
                    m = norm(part)
                    if m and m not in mats:
                        mats.append(m)
            for it in mats:
                if it not in totals:
                    QMessageBox.critical(
                        self, "Unknown material",
                        f"Load {ld.identifier} uses material '{it}' "
                        f"but no total weight was entered for it."
                    )
                    return
            raw_bundles.append((ld.identifier, mats, ld.bundle_total))

        if not raw_bundles:
            QMessageBox.warning(self, "No bundle totals", "No load has a total to solve.")
            return

        self.run_button.setEnabled(False)
        self._log(f"Running solver on {len(raw_bundles)} load(s)…")
        self.worker = SolverWorker(
            totals, raw_bundles,
            int(self.step_combo.currentText()),
            self.min_weight_spin.value(),
        )
        self.worker.done.connect(self._solver_done)
        self.worker.start()

    def _solver_done(self, per_load, error: str) -> None:
        self.run_button.setEnabled(True)
        if error:
            self._log(f"ERROR: {error}")
            QMessageBox.critical(self, "Solver failed", error)
            return

        frn = self.frn_input.text().strip()
        if self.summary_path is not None:
            try:
                updated = write_back_r_tabs(self.summary_path, frn, self.loads)
                if updated:
                    self._log(f"Updated {updated} Class cell(s) in R tabs.")
            except Exception as e:
                self._log(f"R tab write-back failed: {e}")

        self._log("")
        self._log("=== SOLUTION ===")
        for load_id in sorted(per_load.keys()):
            weights = per_load[load_id]
            total = sum(weights.values())
            self._log(f"{load_id}:")
            for mat in sorted(weights.keys()):
                self._log(f"    {mat}: {weights[mat]}")
            self._log(f"    Total: {total}")
            self._log("")

        material_totals: dict[str, int] = {}
        for load_id, weights in per_load.items():
            for mat, w in weights.items():
                material_totals[mat] = material_totals.get(mat, 0) + w

        self._log("=== CHECK (material totals) ===")
        for r in range(self.mats_table.rowCount()):
            name_item = self.mats_table.item(r, 0)
            weight_item = self.mats_table.item(r, 1)
            if not name_item or not weight_item:
                continue
            name = normalize_material(name_item.text())
            if not name:
                continue
            key = norm(name)
            if key not in material_totals:
                continue
            try:
                expected = int(weight_item.text().strip())
            except ValueError:
                continue
            got = material_totals[key]
            flag = "OK" if got == expected else "MISMATCH"
            self._log(f"  {name}: {got} / {expected}  [{flag}]")

        try:
            self._write_outputs(per_load)
        except Exception as e:
            self._log(f"\nWRITE ERROR: {e}")
            QMessageBox.critical(self, "Write failed", str(e))
            return
        self._log("\nAll files written.")

    def _write_outputs(self, per_load: dict[str, dict[str, int]]) -> None:
        frn = self.frn_input.text().strip()
        load_date = parse_frn_date(frn)
        if self.summary_path is None:
            raise RuntimeError("No summary report loaded.")
        summary_path = self.summary_path

        # R/P tabs and the Summary tab are written into the loaded report
        # itself, so they always live next to it. Split Material Weighted
        # workbooks can go somewhere else if the user asked for it.
        summary_dir = summary_path.parent
        summary_dir.mkdir(parents=True, exist_ok=True)

        configured_split = self.config.get("split_folder")
        if configured_split and Path(configured_split).is_dir():
            split_out_dir = Path(configured_split)
        else:
            split_out_dir = summary_dir
        split_out_dir.mkdir(parents=True, exist_ok=True)

        load_by_letter = {ld.letter: ld for ld in self.loads}

        per_letter: dict[str, dict[str, int]] = {}
        for ident, weights in per_load.items():
            m = re.search(r"-([A-Za-z]+)$", ident)
            if not m:
                continue
            per_letter[m.group(1).upper()] = weights

        for letter in sorted(per_letter.keys()):
            ld = load_by_letter.get(letter)
            if ld is None:
                continue
            dest = split_out_dir / f"Split Material Weighted {frn}-{letter}.xls"
            write_split_material_workbook(
                load=ld,
                load_date=load_date,
                weights=per_letter[letter],
                destination=dest,
            )
            self._log(f"Wrote {dest.name}")

        fill_p_tabs(summary_path, frn, per_letter, self.loads, load_date)
        self._log(f"Updated P tabs in {summary_path.name}")

        try:
            written = fill_summary_tab(summary_path, frn, self.loads)
            self._log(f"Filled {written} PL # cell(s) in Summary tab.")
        except Exception as e:
            self._log(f"Summary tab fill failed: {e}")
            raise


def main() -> None:
    app = QApplication(sys.argv)
    window = SplitMaterialWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()