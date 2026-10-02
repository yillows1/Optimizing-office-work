import re
import sys

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTextEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QSpinBox, QAbstractItemView, QGroupBox,
    QListWidget, QListWidgetItem, QInputDialog, QComboBox
)
from PySide6.QtGui import QTextCursor

# ============================================================
# ALIASES / NORMALIZATION
# ============================================================
ALIASES = {
    "99FO": "99FO", "FO99": "99FO",
}

def norm(s):
    key = re.sub(r"\s+", "", s).upper()
    return ALIASES.get(key, key)

# ============================================================
# SETTINGS
# ============================================================
MIN_WEIGHT = 3

# ============================================================
# SOLVER
# ============================================================
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
        weights = {}
        for it in b["items"]:
            weights[it] = solver.Value(x[(bi, it)]) * step
        result.append({
            "load": b["load"],
            "items": b["items"],
            "weights": weights,
        })
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
            assignment.append({"load": b["load"], "items": items,
                               "weights": dict(zip(items, combo))})
            if recurse(idx + 1):
                return True
            assignment.pop()
            for it, w in zip(items, combo):
                remaining[it] += w
        return False

    if recurse(0):
        return assignment
    return None

# ============================================================
# GUI
# ============================================================
class SolverWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bundle Solver")
        self.resize(1300, 850)

        self.loads = {}

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # ---------------- ITEMS TABLE ----------------
        items_box = QGroupBox("Items (name + total weight)")
        items_layout = QVBoxLayout(items_box)

        self.items_table = QTableWidget(0, 2)
        self.items_table.setHorizontalHeaderLabels(["Item Name", "Total Weight"])
        self.items_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.items_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.items_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.items_table.setMaximumHeight(180)
        items_layout.addWidget(self.items_table)

        items_btns = QHBoxLayout()
        add_item_btn = QPushButton("+ Add Item")
        add_item_btn.clicked.connect(self.add_item_row)
        del_item_btn = QPushButton("− Remove Selected Item")
        del_item_btn.clicked.connect(self.remove_item_row)
        items_btns.addWidget(add_item_btn)
        items_btns.addWidget(del_item_btn)
        items_btns.addStretch()
        items_layout.addLayout(items_btns)

        root.addWidget(items_box)

        # ---------------- LOADS + BUNDLES SPLIT ----------------
        loads_and_bundles = QGroupBox("Loads & Bundles")
        lb_layout = QHBoxLayout(loads_and_bundles)

        left = QVBoxLayout()
        left.addWidget(QLabel("Loads"))
        self.loads_list = QListWidget()
        self.loads_list.currentItemChanged.connect(self.on_load_selected)
        left.addWidget(self.loads_list, stretch=1)

        load_btns = QHBoxLayout()
        add_load_btn = QPushButton("+ Add Load")
        add_load_btn.clicked.connect(self.add_load)
        del_load_btn = QPushButton("− Remove Load")
        del_load_btn.clicked.connect(self.remove_load)
        load_btns.addWidget(add_load_btn)
        load_btns.addWidget(del_load_btn)
        left.addLayout(load_btns)

        left_widget = QWidget()
        left_widget.setLayout(left)
        left_widget.setMaximumWidth(220)

        lb_layout.addWidget(left_widget)

        right = QVBoxLayout()
        self.bundles_label = QLabel("Select a load on the left")
        self.bundles_label.setStyleSheet("font-weight: bold;")
        right.addWidget(self.bundles_label)

        self.bundles_table = QTableWidget(0, 2)
        self.bundles_table.setHorizontalHeaderLabels(["Items (comma-sep)", "Total Weight"])
        self.bundles_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.bundles_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.bundles_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.bundles_table.itemChanged.connect(self.on_bundle_cell_changed)
        right.addWidget(self.bundles_table, stretch=1)

        bundle_btns = QHBoxLayout()
        add_bundle_btn = QPushButton("+ Add Bundle to This Load")
        add_bundle_btn.clicked.connect(self.add_bundle_row)
        del_bundle_btn = QPushButton("− Remove Selected Bundle")
        del_bundle_btn.clicked.connect(self.remove_bundle_row)
        dup_bundle_btn = QPushButton("Duplicate Selected")
        dup_bundle_btn.clicked.connect(self.duplicate_bundle_row)
        bundle_btns.addWidget(add_bundle_btn)
        bundle_btns.addWidget(del_bundle_btn)
        bundle_btns.addWidget(dup_bundle_btn)
        bundle_btns.addStretch()
        right.addLayout(bundle_btns)

        right_widget = QWidget()
        right_widget.setLayout(right)
        lb_layout.addWidget(right_widget, stretch=1)

        root.addWidget(loads_and_bundles, stretch=1)

        # ---------------- OPTIONS ----------------
        opts_box = QGroupBox("Options")
        opts_layout = QHBoxLayout(opts_box)

        opts_layout.addWidget(QLabel("Start Step:"))
        self.step_combo = QComboBox()
        self.step_combo.addItems(["1", "5", "10"])
        self.step_combo.setCurrentText("1")
        opts_layout.addWidget(self.step_combo)

        opts_layout.addWidget(QLabel("Min Weight:"))
        self.min_weight_spin = QSpinBox()
        self.min_weight_spin.setRange(1, 1000)
        self.min_weight_spin.setValue(MIN_WEIGHT)
        opts_layout.addWidget(self.min_weight_spin)

        opts_layout.addStretch()

        run_btn = QPushButton("▶  RUN SOLVER")
        run_btn.setStyleSheet("font-weight: bold; padding: 6px 20px;")
        run_btn.clicked.connect(self.run_solver)
        opts_layout.addWidget(run_btn)

        clear_results_btn = QPushButton("Clear Results")
        clear_results_btn.clicked.connect(lambda: self.output.clear())
        opts_layout.addWidget(clear_results_btn)

        root.addWidget(opts_box)

        # ---------------- OUTPUT ----------------
        output_box = QGroupBox("Results")
        output_layout = QVBoxLayout(output_box)
        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setStyleSheet("font-family: Consolas, monospace;")
        output_layout.addWidget(self.output)
        root.addWidget(output_box, stretch=1)

        self._loading_bundle_table = False

    # --------------------------------------------------------
    # ITEMS
    # --------------------------------------------------------
    def add_item_row(self, name="", weight=""):
        r = self.items_table.rowCount()
        self.items_table.insertRow(r)
        self.items_table.setItem(r, 0, QTableWidgetItem(str(name)))
        self.items_table.setItem(r, 1, QTableWidgetItem(str(weight)))

    def remove_item_row(self):
        rows = sorted({idx.row() for idx in self.items_table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "No selection", "Select a row in the Items table to remove.")
            return
        for r in rows:
            self.items_table.removeRow(r)

    # --------------------------------------------------------
    # LOADS
    # --------------------------------------------------------
    def next_load_letter(self):
        used = set(self.loads.keys())
        for i in range(26):
            letter = chr(ord("A") + i)
            if letter not in used:
                return letter
        i = 26
        while True:
            letter = ""
            n = i
            while True:
                letter = chr(ord("A") + n % 26) + letter
                n = n // 26 - 1
                if n < 0:
                    break
            if letter not in used:
                return letter
            i += 1

    def add_load(self):
        n, ok = QInputDialog.getInt(
            self, "Add Loads", "How many loads to add?", 1, 1, 26
        )
        if not ok:
            return
        for _ in range(n):
            letter = self.next_load_letter()
            self.loads[letter] = []
            self.loads_list.addItem(QListWidgetItem(f"Load {letter}"))

    def remove_load(self):
        item = self.loads_list.currentItem()
        if item is None:
            QMessageBox.information(self, "No selection", "Select a load to remove.")
            return
        letter = item.text().replace("Load ", "").strip()
        if letter in self.loads:
            del self.loads[letter]
        self.loads_list.takeItem(self.loads_list.row(item))
        self.bundles_table.setRowCount(0)
        self.bundles_label.setText("Select a load on the left")

    def on_load_selected(self, current, previous):
        if previous is not None:
            self._save_current_bundles(previous)

        self._loading_bundle_table = True
        self.bundles_table.setRowCount(0)

        if current is None:
            self.bundles_label.setText("Select a load on the left")
            self._loading_bundle_table = False
            return

        letter = current.text().replace("Load ", "").strip()
        self.bundles_label.setText(f"Bundles in Load {letter}")
        for items_str, total_str in self.loads.get(letter, []):
            r = self.bundles_table.rowCount()
            self.bundles_table.insertRow(r)
            self.bundles_table.setItem(r, 0, QTableWidgetItem(items_str))
            self.bundles_table.setItem(r, 1, QTableWidgetItem(total_str))

        self._loading_bundle_table = False

    def _save_current_bundles(self, item):
        if item is None:
            return
        letter = item.text().replace("Load ", "").strip()
        rows = []
        for r in range(self.bundles_table.rowCount()):
            items_item = self.bundles_table.item(r, 0)
            total_item = self.bundles_table.item(r, 1)
            items_str = items_item.text() if items_item else ""
            total_str = total_item.text() if total_item else ""
            rows.append((items_str, total_str))
        self.loads[letter] = rows

    def on_bundle_cell_changed(self, item):
        if self._loading_bundle_table:
            return
        current = self.loads_list.currentItem()
        if current is not None:
            self._save_current_bundles(current)

    # --------------------------------------------------------
    # BUNDLES
    # --------------------------------------------------------
    def add_bundle_row(self, items="", total=""):
        current = self.loads_list.currentItem()
        if current is None:
            QMessageBox.information(self, "No load", "Add or select a load first.")
            return
        r = self.bundles_table.rowCount()
        self.bundles_table.insertRow(r)
        self.bundles_table.setItem(r, 0, QTableWidgetItem(str(items)))
        self.bundles_table.setItem(r, 1, QTableWidgetItem(str(total)))
        self._save_current_bundles(current)

    def remove_bundle_row(self):
        current = self.loads_list.currentItem()
        if current is None:
            return
        rows = sorted({idx.row() for idx in self.bundles_table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "No selection", "Select a row to remove.")
            return
        for r in rows:
            self.bundles_table.removeRow(r)
        self._save_current_bundles(current)

    def duplicate_bundle_row(self):
        current = self.loads_list.currentItem()
        if current is None:
            return
        rows = sorted({idx.row() for idx in self.bundles_table.selectedIndexes()})
        if not rows:
            QMessageBox.information(self, "No selection", "Select a row to duplicate.")
            return
        for r in rows:
            items_item = self.bundles_table.item(r, 0)
            total_item = self.bundles_table.item(r, 1)
            items_str = items_item.text() if items_item else ""
            total_str = total_item.text() if total_item else ""
            rr = self.bundles_table.rowCount()
            self.bundles_table.insertRow(rr)
            self.bundles_table.setItem(rr, 0, QTableWidgetItem(items_str))
            self.bundles_table.setItem(rr, 1, QTableWidgetItem(total_str))
        self._save_current_bundles(current)

    # --------------------------------------------------------
    # RUN
    # --------------------------------------------------------
    def log(self, text):
        self.output.insertPlainText(text + "\n")
        self.output.moveCursor(QTextCursor.MoveOperation.End)

    def read_items(self):
        totals = {}
        display = {}
        for r in range(self.items_table.rowCount()):
            name_item = self.items_table.item(r, 0)
            weight_item = self.items_table.item(r, 1)
            if not name_item or not weight_item:
                continue
            name = name_item.text().strip()
            w = weight_item.text().strip()
            if not name or not w:
                continue
            try:
                wv = int(w)
            except ValueError:
                raise ValueError(f"Items row {r+1}: weight '{w}' is not a whole number.")
            if wv <= 0:
                raise ValueError(f"Items row {r+1}: weight must be positive.")
            key = norm(name)
            totals[key] = wv
            display[key] = name
        return totals, display

    def read_bundles(self):
        current = self.loads_list.currentItem()
        if current is not None:
            self._save_current_bundles(current)

        raw = []
        for letter, rows in self.loads.items():
            for idx, (items_str, total_str) in enumerate(rows):
                items_str = items_str.strip()
                total_str = total_str.strip()
                if not items_str and not total_str:
                    continue
                if not items_str or not total_str:
                    raise ValueError(
                        f"Load {letter}, bundle {idx+1}: both items and total are required."
                    )
                items = [norm(s) for s in items_str.split(",") if s.strip()]
                try:
                    total = int(total_str)
                except ValueError:
                    raise ValueError(
                        f"Load {letter}, bundle {idx+1}: total '{total_str}' is not a whole number."
                    )
                if total <= 0:
                    raise ValueError(
                        f"Load {letter}, bundle {idx+1}: total must be positive."
                    )
                raw.append((letter, items, total))
        return raw

    def run_solver(self):
        self.output.clear()

        try:
            totals, display = self.read_items()
            raw_bundles = self.read_bundles()
        except ValueError as e:
            QMessageBox.critical(self, "Input error", str(e))
            return

        if not totals:
            QMessageBox.warning(self, "Missing items", "Add at least one item.")
            return
        if not raw_bundles:
            QMessageBox.warning(self, "Missing bundles", "Add at least one bundle.")
            return

        for load, items, total in raw_bundles:
            unknown = [it for it in items if it not in totals]
            if unknown:
                QMessageBox.critical(
                    self, "Unknown item",
                    f"Load {load} references unknown items: {unknown}\n"
                    f"Known items: {list(display.values())}"
                )
                return

        min_w = self.min_weight_spin.value()
        start_step = int(self.step_combo.currentText())

        global MIN_WEIGHT
        old_min = MIN_WEIGHT
        MIN_WEIGHT = min_w
        try:
            self._do_solve(totals, display, raw_bundles, start_step)
        finally:
            MIN_WEIGHT = old_min

    def _do_solve(self, totals, display, raw_bundles, start_step):
        remaining = dict(totals)
        bundles = []
        locked = []
        for load, items, total in raw_bundles:
            if len(items) == 1:
                item = items[0]
                if item not in remaining:
                    self.log(f"ERROR: unknown item '{item}'")
                    return
                remaining[item] -= total
                locked.append({"load": load, "items": items, "weights": {item: total}})
            else:
                bundles.append({"load": load, "items": items, "total": total})

        for k, v in remaining.items():
            if v < 0:
                self.log(f"ERROR: {display.get(k,k)} is already over budget by {-v}")
                return

        bundles.sort(key=lambda b: (len(b["items"]), -b["total"]))

        steps = [s for s in (1, 5, 10) if s <= start_step]

        result = None
        used_step = None

        if HAS_ORTOOLS:
            self.log("[using OR-Tools CP-SAT solver]")
            for step in steps:
                self.log(f"[trying step={step} ...]")
                QApplication.processEvents()
                try:
                    result = solve_cp(bundles, remaining, step, totals, time_limit=15.0)
                except Exception as e:
                    self.log(f"  error: {e}")
                    result = None
                if result is not None:
                    used_step = step
                    self.log(f"  found a solution with step={step}")
                    break
                else:
                    self.log(f"  no solution at step={step}")
        else:
            self.log("[OR-Tools not found — using slower recursive solver]")
            for step in (5, 1):
                if step > start_step:
                    continue
                self.log(f"[trying step={step} ...]")
                QApplication.processEvents()
                _partition_cache.clear()
                global _iter_count
                _iter_count = 0
                try:
                    result = solve_recursive(bundles, remaining, step, totals)
                except RuntimeError as e:
                    self.log(f"  gave up: {e}")
                    result = None
                if result is not None:
                    used_step = step
                    self.log(f"  found a solution with step={step}")
                    break
                else:
                    self.log(f"  no solution at step={step}")

        if result is None:
            self.log("\nNo solution found.")
            return

        if used_step is None:
            self.log("\nSolver returned a solution without a step size.")
            return

        effective_min = ((MIN_WEIGHT + used_step - 1) // used_step) * used_step
        bad = []
        for b in result:
            for it, w in b["weights"].items():
                if w < effective_min:
                    bad.append((b["load"], it, w))
        if bad:
            self.log(f"\n[!] Solver returned weights below the minimum of {effective_min}:")
            for load, it, w in bad:
                self.log(f"    Load {load}: {display.get(it, it)} = {w}")
            self.log("    Rejecting solution.")
            return

        all_bundles = locked + result

        self.log(f"\n=== SOLUTION (minimum weight {MIN_WEIGHT}, multiples of {used_step}) ===\n")

        per_load = {}
        item_totals = {k: 0 for k in totals}
        for b in all_bundles:
            load = b["load"]
            per_load.setdefault(load, {})
            for it, w in b["weights"].items():
                per_load[load][it] = per_load[load].get(it, 0) + w
                item_totals[it] += w

        for load in sorted(per_load):
            self.log(f"Load {load} ->")
            for it, w in sorted(per_load[load].items()):
                self.log(f"    {display.get(it, it)}: {w}")
            self.log("")

        self.log("=== LOAD TOTALS ===")
        for load in sorted(per_load):
            self.log(f"  Load {load}: {sum(per_load[load].values())}")

        self.log("\n=== CHECK: item totals ===")
        for k in totals:
            flag = "OK" if item_totals[k] == totals[k] else "MISMATCH"
            self.log(f"  {display.get(k,k)}: {item_totals[k]} / {totals[k]}  [{flag}]")

# ============================================================
# MAIN
# ============================================================
def main():
    app = QApplication(sys.argv)
    win = SolverWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()