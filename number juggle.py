import re
import sys

# Uses: py -m pip install --user ortools make sure to install that before trying to run this code

# ============================================================
# ALIASES
# ============================================================
ALIASES = {
    "99FO": "99FO", "FO99": "99FO",
    "69S":  "69S",  "69":   "69S",
}

def norm(s):
    key = re.sub(r"\s+", "", s).upper()
    return ALIASES.get(key, key)

# ============================================================
# SETTINGS
# ============================================================
MIN_WEIGHT = 4   # No item in any bundle may be assigned less than this

# ============================================================
# SOLVER 1: OR-Tools CP-SAT (preferred — fast + complete)
# ============================================================
try:
    from ortools.sat.python import cp_model
    HAS_ORTOOLS = True
except ImportError:
    HAS_ORTOOLS = False

def solve_cp(bundles, remaining, step, totals, time_limit=15.0):
    """
    Solve using Google OR-Tools CP-SAT.
    Returns a list of bundle assignments in the same format as the recursive solver,
    or None if no solution exists.
    Every item in every bundle must be >= MIN_WEIGHT.
    """
    model = cp_model.CpModel()

    # Smallest allowed value that is both >= MIN_WEIGHT and a multiple of step
    min_units = -(-MIN_WEIGHT // step)   # ceiling division

    # Decision variable for each (bundle, item) pair: integer count of `step` units.
    x = {}
    for bi, b in enumerate(bundles):
        for it in b["items"]:
            max_units = min(remaining[it] // step, b["total"] // step)
            if max_units < min_units:
                # This item can't reach the minimum in this bundle.
                # The whole bundle+step combination is infeasible.
                return None
            x[(bi, it)] = model.NewIntVar(min_units, max_units, f"x_{bi}_{it}")

    # Each bundle's assigned weights must sum exactly to its total
    for bi, b in enumerate(bundles):
        model.Add(
            sum(x[(bi, it)] * step for it in b["items"]) == b["total"]
        )

    # Each item's total across all bundles must equal its remaining amount
    for item, need in remaining.items():
        contributors = [
            x[(bi, it)] * step
            for bi, b in enumerate(bundles)
            for it in b["items"]
            if it == item
        ]
        if contributors:
            model.Add(sum(contributors) == need)
        else:
            if need != 0:
                return None  # leftover item that no bundle can consume

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    # Rebuild assignment in the same format as the recursive solver
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
# SOLVER 2: Recursive fallback (used only if OR-Tools isn't installed)
# ============================================================
ITER_LIMIT = 500_000
_iter_count = 0
_partition_cache = {}

def partitions_capped_uncached(total, caps, step, min_weight=MIN_WEIGHT):
    k = len(caps)
    # Smallest multiple of step that is >= min_weight
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
# INPUT: items and totals
# ============================================================
print("Enter item types and their TOTAL weights.")
print("Format: name,weight  (blank line to finish)")
totals = {}
display = {}
while True:
    line = input("  > ").strip()
    if not line:
        break
    name_raw, w = line.split(",")
    key = norm(name_raw)
    totals[key] = int(w.strip())
    display[key] = name_raw.strip()

# ============================================================
# INPUT: loads
# ============================================================
num_loads = int(input("\nHow many loads? "))

raw_bundles = []
for i in range(num_loads):
    load = chr(ord("A") + i)
    n = int(input(f"\nLoad {load}: how many bundles? "))
    for j in range(n):
        items_str = input(f"  Bundle {j+1} item types (comma-sep): ")
        items = [norm(s) for s in items_str.split(",")]
        total = int(input(f"  Bundle {j+1} total weight: "))
        raw_bundles.append((load, items, total))

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
    # Try finer steps first — coarser steps are more likely to be infeasible
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
