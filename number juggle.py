import re
import sys

# Uses: py -m pip install --user ortools make sure to install that before trying to run this code

# ============================================================
# ALIASES
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
MIN_WEIGHT = 3   # No item in any bundle may be assigned less than this

# ============================================================
# SAFE INPUT HELPERS
# ============================================================
def ask_int(prompt, allow_back=False, min_value=None, max_value=None):
    while True:
        raw = input(prompt).strip()
        if allow_back and raw.lower() == "back":
            return "back"
        try:
            val = int(raw)
        except ValueError:
            print("  [!] Wrong input — please enter a whole number.")
            continue
        if min_value is not None and val < min_value:
            print(f"  [!] Must be at least {min_value}.")
            continue
        if max_value is not None and val > max_value:
            print(f"  [!] Must be at most {max_value}.")
            continue
        return val

def ask_str(prompt, allow_back=False, validator=None):
    while True:
        raw = input(prompt).strip()
        if allow_back and raw.lower() == "back":
            return "back"
        if not raw:
            print("  [!] Can't be empty. Try again.")
            continue
        if validator is not None and not validator(raw):
            print("  [!] Wrong input — try again (or type 'back' to revise).")
            continue
        return raw

def ask_yes_no(prompt, default="n"):
    while True:
        raw = input(prompt).strip().lower()
        if not raw:
            raw = default
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("  [!] Please answer y or n.")

# ============================================================
# SOLVER 1: OR-Tools CP-SAT (preferred — fast + complete)
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

# ============================================================
# SOLVER 2: Recursive fallback
# ============================================================
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
# INPUT: items and totals (with back/undo)
# ============================================================
print("Enter item types and their TOTAL weights.")
print("Format: name,weight   (blank line to finish, or 'back' to undo last item)")

totals = {}
display = {}
entry_order = []

while True:
    line = input("  > ").strip()
    if not line:
        if not totals:
            print("  [!] You need at least one item before finishing.")
            continue
        break

    if line.lower() == "back":
        if not entry_order:
            print("  [!] Nothing to undo.")
            continue
        last = entry_order.pop()
        totals.pop(last, None)
        display.pop(last, None)
        print(f"  [removed last item: {last}]")
        continue

    if "," not in line:
        print("  [!] Wrong input — use the format: name,weight")
        continue

    name_raw, w = line.split(",", 1)
    name_raw = name_raw.strip()
    w = w.strip()

    if not name_raw:
        print("  [!] Item name can't be empty.")
        continue

    try:
        weight_val = int(w)
    except ValueError:
        print("  [!] Weight must be a whole number.")
        continue

    if weight_val <= 0:
        print("  [!] Weight must be positive.")
        continue

    key = norm(name_raw)
    totals[key] = weight_val
    display[key] = name_raw
    if key not in entry_order:
        entry_order.append(key)
    print(f"  [set {name_raw} = {weight_val}]")

# ============================================================
# INPUT: loads (with back support)
# ============================================================
while True:
    num_loads_raw = input("\nHow many loads? (or 'back' to redo items): ").strip()
    if num_loads_raw.lower() == "back":
        print("\n--- Redo item totals ---")
        totals.clear()
        display.clear()
        entry_order.clear()
        print("Enter item types and their TOTAL weights.")
        print("Format: name,weight   (blank line to finish, or 'back' to undo last item)")
        while True:
            line = input("  > ").strip()
            if not line:
                if not totals:
                    print("  [!] You need at least one item.")
                    continue
                break
            if line.lower() == "back":
                if not entry_order:
                    print("  [!] Nothing to undo.")
                    continue
                last = entry_order.pop()
                totals.pop(last, None)
                display.pop(last, None)
                print(f"  [removed last item: {last}]")
                continue
            if "," not in line:
                print("  [!] Wrong input — use: name,weight")
                continue
            name_raw, w = line.split(",", 1)
            name_raw = name_raw.strip()
            w = w.strip()
            if not name_raw:
                print("  [!] Item name can't be empty.")
                continue
            try:
                weight_val = int(w)
            except ValueError:
                print("  [!] Weight must be a whole number.")
                continue
            if weight_val <= 0:
                print("  [!] Weight must be positive.")
                continue
            key = norm(name_raw)
            totals[key] = weight_val
            display[key] = name_raw
            if key not in entry_order:
                entry_order.append(key)
            print(f"  [set {name_raw} = {weight_val}]")
        continue

    try:
        num_loads = int(num_loads_raw)
    except ValueError:
        print("  [!] Wrong input — enter a whole number.")
        continue
    if num_loads <= 0:
        print("  [!] Must be at least 1 load.")
        continue
    break

raw_bundles = []
i = 0
while i < num_loads:
    load = chr(ord("A") + i)

    # --- Ask how many bundles for THIS load ---
    n = None
    while n is None:
        n_raw = input(f"\nLoad {load}: how many bundles? "
                      f"(or 'back' to redo this load, 'prev' to redo previous load) ").strip()

        if n_raw.lower() == "back":
            raw_bundles[:] = [b for b in raw_bundles if b[0] != load]
            print(f"  [cleared bundles for Load {load}, re-asking count]")
            continue

        if n_raw.lower() == "prev":
            if i == 0:
                print("  [!] No previous load to redo.")
                continue
            prev_load = chr(ord("A") + i - 1)
            raw_bundles[:] = [b for b in raw_bundles if b[0] != prev_load]
            print(f"  [cleared Load {prev_load}, going back]")
            i -= 1
            n = "break_outer"
            break

        try:
            n = int(n_raw)
        except ValueError:
            print("  [!] Wrong input — enter a whole number.")
            continue
        if n <= 0:
            print("  [!] Must be at least 1 bundle.")
            n = None
            continue

    if n == "break_outer":
        continue

    # --- Gather bundles for this load ---
    load_bundles = []
    j = 0
    while j < n:
        items_str = input(f"  Bundle {j+1} item types "
                          f"(or 'back' to redo previous bundle): ").strip()

        if items_str.lower() == "back":
            if not load_bundles:
                print("  [!] Nothing to undo in this load.")
                continue
            load_bundles.pop()
            j -= 1
            continue

        items = [norm(s) for s in items_str.split(",") if s.strip()]
        if not items:
            print("  [!] Wrong input — list at least one item.")
            continue
        unknown = [it for it in items if it not in totals]
        if unknown:
            print(f"  [!] Unknown items: {unknown}")
            print(f"      Known items: {list(display.values())}")
            continue

        total_raw = input(f"  Bundle {j+1} total weight "
                          f"(or 'back' to redo this bundle): ").strip()
        if total_raw.lower() == "back":
            continue
        try:
            total = int(total_raw)
        except ValueError:
            print("  [!] Weight must be a whole number. Redoing this bundle.")
            continue
        if total <= 0:
            print("  [!] Weight must be positive. Redoing this bundle.")
            continue

        load_bundles.append((load, items, total))
        j += 1

    raw_bundles.extend(load_bundles)
    i += 1

# ============================================================
# Lock single-item bundles
# ============================================================
remaining = dict(totals)
bundles = []
locked = []
for load, items, total in raw_bundles:
    if len(items) == 1:
        item = items[0]
        if item not in remaining:
            print(f"ERROR: unknown item '{item}'. Known keys: {list(totals)}")
            sys.exit()
        remaining[item] -= total
        locked.append({"load": load, "items": items,
                       "weights": {item: total}})
    else:
        for it in items:
            if it not in remaining:
                print(f"ERROR: unknown item '{it}'. Known keys: {list(totals)}")
                sys.exit()
        bundles.append({"load": load, "items": items, "total": total})

for k, v in remaining.items():
    if v < 0:
        print(f"ERROR: {k} is already over budget by {-v}")
        sys.exit()

bundles.sort(key=lambda b: (len(b["items"]), -b["total"]))

# ============================================================
# Solve
# ============================================================
result = None
used_step = None

if HAS_ORTOOLS:
    print("\n[using OR-Tools CP-SAT solver]")
    for step in (1, 5, 10):
        print(f"[trying step={step} ...]", flush=True)
        try:
            result = solve_cp(bundles, remaining, step, totals, time_limit=15.0)
        except Exception as e:
            print(f"  error: {e}")
            result = None
        if result is not None:
            used_step = step
            print(f"  found a solution with step={step}", flush=True)
            break
        else:
            print(f"  no solution at step={step}", flush=True)
else:
    print("\n[OR-Tools not found — using slower recursive solver]")
    print("[install with:  py -m pip install --user ortools]")
    for step in (5, 1):
        print(f"[trying step={step} ...]", flush=True)
        _partition_cache.clear()
        _iter_count = 0
        try:
            result = solve_recursive(bundles, remaining, step, totals)
        except RuntimeError as e:
            print(f"  gave up: {e}", flush=True)
            result = None
        if result is not None:
            used_step = step
            print(f"  found a solution with step={step}", flush=True)
            break
        else:
            print(f"  no solution at step={step}", flush=True)

if result is None:
    print("\nNo solution found.")
    sys.exit()

if used_step is None:
    print("\nSolver returned a solution without a step size.")
    sys.exit()

# --- Reject any solution with a weight below the effective minimum ---
effective_min = ((MIN_WEIGHT + used_step - 1) // used_step) * used_step
bad = []
for b in result:
    for it, w in b["weights"].items():
        if w < effective_min:
            bad.append((b["load"], it, w))
if bad:
    print(f"\n[!] Solver returned weights below the minimum of {effective_min}:")
    for load, it, w in bad:
        print(f"    Load {load}: {display.get(it, it)} = {w}")
    print("    Rejecting solution.")
    sys.exit()

all_bundles = locked + result

# ============================================================
# Print
# ============================================================
print(f"\n=== SOLUTION (minimum weight {MIN_WEIGHT}, multiples of {used_step}) ===\n")

per_load = {}
item_totals = {k: 0 for k in totals}
for b in all_bundles:
    load = b["load"]
    per_load.setdefault(load, {})
    for it, w in b["weights"].items():
        per_load[load][it] = per_load[load].get(it, 0) + w
        item_totals[it] += w

for load in sorted(per_load):
    print(f"Load {load} ->")
    for it, w in sorted(per_load[load].items()):
        print(f"    {display.get(it, it)}: {w}")
    print()

print("=== LOAD TOTALS ===")
for load in sorted(per_load):
    print(f"  Load {load}: {sum(per_load[load].values())}")

print("\n=== CHECK: item totals ===")
for k in totals:
    flag = "OK" if item_totals[k] == totals[k] else "MISMATCH"
    print(f"  {display.get(k,k)}: {item_totals[k]} / {totals[k]}  [{flag}]")