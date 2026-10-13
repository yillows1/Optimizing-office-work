
import re
import random
import sys

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
# SOLVER
# ============================================================
ITER_LIMIT = 5_000_000          # hard safety cap
_iter_count = 0

def partitions_capped(total, caps, step):
    """Yield tuples w0..wk-1 with step <= wi <= caps[i], sum == total."""
    k = len(caps)
    if k == 1:
        if step <= total <= caps[0] and total % step == 0:
            yield (total,)
        return
    hi = min(caps[0], total - step * (k - 1))
    choices = list(range(step, hi + 1, step))
    random.shuffle(choices)
    for first in choices:
        for rest in partitions_capped(total - first, caps[1:], step):
            yield (first,) + rest

def solve(bundles, remaining, step, idx=0, assignment=None):
    global _iter_count
    if assignment is None:
        assignment = []
    if idx == len(bundles):
        return assignment if all(v == 0 for v in remaining.values()) else None

    _iter_count += 1
    if _iter_count > ITER_LIMIT:
        raise RuntimeError("iteration limit reached")

    b = bundles[idx]
    items = b["items"]
    caps = [remaining[it] for it in items]

    for combo in partitions_capped(b["total"], caps, step):
        if any(w > remaining[it] for it, w in zip(items, combo)):
            continue
        for it, w in zip(items, combo):
            remaining[it] -= w
        assignment.append({"load": b["load"], "items": items,
                           "weights": dict(zip(items, combo))})
        result = solve(bundles, remaining, step, idx + 1, assignment)
        if result is not None:
            return result
        assignment.pop()
        for it, w in zip(items, combo):
            remaining[it] += w
    return None

# ============================================================
# INPUT: items and totals
# ============================================================
print("Enter item types and their TOTAL weights.")
print("Format: name,weight  (blank line to finish)")
print("(case and spaces are ignored, so '99 fo' == '99FO' == 'fo99')")
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

# Most-constrained first: fewest items, then largest total
bundles.sort(key=lambda b: (len(b["items"]), -b["total"]))

# ============================================================
# Solve
# ============================================================
result = None
used_step = None
for step in (10, 5, 1):
    print(f"\n[trying step={step} ...]", flush=True)
    remaining_copy = dict(remaining)
    _iter_count = 0
    try:
        result = solve(bundles, remaining_copy, step)
    except RuntimeError as e:
        print(f"  gave up: {e}", flush=True)
        result = None
    if result is not None:
        used_step = step
        print(f"  found a solution with step={step}", flush=True)
        break

if result is None:
    print("\nNo solution found.")
    sys.exit()

all_bundles = locked + result

# ============================================================
# Print: grouped by load
# ============================================================
print(f"\n=== SOLUTION (multiples of {used_step}) ===\n")

# Group: load -> item -> total weight
per_load = {}
item_totals = {k: 0 for k in totals}

for b in all_bundles:
    load = b["load"]
    per_load.setdefault(load, {})
    for it, w in b["weights"].items():
        per_load[load][it] = per_load[load].get(it, 0) + w
        item_totals[it] += w

# Print per load
for load in sorted(per_load):
    print(f"Load {load} ->")
    for it, w in sorted(per_load[load].items()):
        print(f"    {display.get(it, it)}: {w}")
    print()

# Optional: also show each load's grand total
print("=== LOAD TOTALS ===")
for load in sorted(per_load):
    print(f"  Load {load}: {sum(per_load[load].values())}")

print("\n=== CHECK: item totals ===")
for k in totals:
    flag = "OK" if item_totals[k] == totals[k] else "MISMATCH"
    print(f"  {display.get(k,k)}: {item_totals[k]} / {totals[k]}  [{flag}]")



