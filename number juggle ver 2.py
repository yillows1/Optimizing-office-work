# Solver that ignores the last bundle and tries to satisfy the rest. Use if number juggle.py can't find a solution.
import re
import sys

# Uses: py -m pip install --user ortools

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
MIN_WEIGHT = 8

# ============================================================
# OR-TOOLS
# ============================================================
try:
    from ortools.sat.python import cp_model
    HAS_ORTOOLS = True
except ImportError:
    HAS_ORTOOLS = False

def solve(ag_bundles, h_bundles, remaining, totals, min_weight=MIN_WEIGHT):
    """
    Fixed-total bundles (ag_bundles): each bundle's item weights sum exactly to its total.
    Flexible-total bundles (h_bundles): each load's grand total is fixed, but per-bundle
    totals are free, subject to grand total.
    Soft objective: push weights up, balance within bundles, prefer multiples of 5.
    """
    model = cp_model.CpModel()

    # ---- Variables ----
    x = {}
    for bi, b in enumerate(ag_bundles):
        for it in b["items"]:
            max_units = min(remaining.get(it, 0), b["total"])
            if max_units < min_weight:
                return None, (f"Fixed bundle in Load {b['load']} with items {b['items']} "
                              f"total {b['total']} cannot give '{it}' >= {min_weight}.")
            x[(bi, it)] = model.new_int_var(min_weight, max_units, f"x_{bi}_{it}")

    y = {}
    for bi, b in enumerate(h_bundles):
        for it in b["items"]:
            max_units = remaining.get(it, 0)
            if max_units < min_weight:
                return None, (f"Flexible bundle in Load {b['load']} with items {b['items']} "
                              f"cannot give '{it}' >= {min_weight} (only {remaining.get(it,0)} left).")
            y[(bi, it)] = model.new_int_var(min_weight, max_units, f"y_{bi}_{it}")

    # ---- Hard constraints ----
    for bi, b in enumerate(ag_bundles):
        model.add(sum(x[(bi, it)] for it in b["items"]) == b["total"])

    loads_in_flex = sorted({b["load"] for b in h_bundles})
    for load in loads_in_flex:
        grand = next(b["grand"] for b in h_bundles if b["load"] == load)
        model.add(
            sum(y[(bi, it)]
                for bi, b in enumerate(h_bundles)
                if b["load"] == load
                for it in b["items"]) == grand
        )

    for item, need in remaining.items():
        contribs = []
        for bi, b in enumerate(ag_bundles):
            if item in b["items"]:
                contribs.append(x[(bi, item)])
        for bi, b in enumerate(h_bundles):
            if item in b["items"]:
                contribs.append(y[(bi, item)])
        if not contribs:
            if need != 0:
                return None, f"Item '{item}' has {need} left but appears in no bundle."
            continue
        model.add(sum(contribs) == need)

    # ---- Soft objective ----
    soft_terms = []

    # (1) Push weights up
    all_vars = list(x.values()) + list(y.values())
    if all_vars:
        soft_terms.append((1, sum(all_vars)))

    # (2) Balance within each multi-item bundle
    def add_spread_penalty(vs, ub, tag):
        if len(vs) < 2:
            return
        mn = model.new_int_var(0, ub, f"mn_{tag}")
        mx = model.new_int_var(0, ub, f"mx_{tag}")
        model.add_min_equality(mn, vs)
        model.add_max_equality(mx, vs)
        spread = model.new_int_var(0, ub, f"spread_{tag}")
        model.add(spread == mx - mn)
        soft_terms.append((-3, spread))

    for bi, b in enumerate(ag_bundles):
        add_spread_penalty([x[(bi, it)] for it in b["items"]],
                           b["total"], f"ag_{bi}")
    for bi, b in enumerate(h_bundles):
        ub = max(remaining.get(it, 0) for it in b["items"])
        add_spread_penalty([y[(bi, it)] for it in b["items"]],
                           ub, f"h_{bi}")

    # (3) Prefer multiples of 5
    def add_round_penalty(v, tag):
        r = model.new_int_var(0, 4, f"r_{tag}")
        model.add_modulo_equality(r, v, 5)
        d = model.new_int_var(0, 2, f"d_{tag}")
        is_high = model.new_bool_var(f"hi_{tag}")
        model.add(r >= 3).only_enforce_if(is_high)
        model.add(r <= 2).only_enforce_if(is_high.Not())
        model.add(d == 5 - r).only_enforce_if(is_high)
        model.add(d == r).only_enforce_if(is_high.Not())
        soft_terms.append((-1, d))

    for bi, b in enumerate(ag_bundles):
        for it in b["items"]:
            add_round_penalty(x[(bi, it)], f"ag_{bi}_{it}")
    for bi, b in enumerate(h_bundles):
        for it in b["items"]:
            add_round_penalty(y[(bi, it)], f"h_{bi}_{it}")

    if soft_terms:
        model.maximize(sum(w * t for w, t in soft_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, "CP-SAT could not find a feasible assignment."

    result = []
    for bi, b in enumerate(ag_bundles):
        weights = {it: solver.Value(x[(bi, it)]) for it in b["items"]}
        result.append({"load": b["load"], "items": b["items"],
                       "weights": weights, "total": b["total"], "fixed": True})
    for bi, b in enumerate(h_bundles):
        weights = {it: solver.Value(y[(bi, it)]) for it in b["items"]}
        result.append({"load": b["load"], "items": b["items"],
                       "weights": weights,
                       "total": sum(weights.values()), "fixed": False})
    return result, None

# ============================================================
# INPUT
# ============================================================
def input_items():
    print("Enter item types and their TOTAL weights.")
    print("Format: name,weight   (blank line to finish, or 'back' to undo)")
    totals, display, order = {}, {}, []
    while True:
        line = input("  > ").strip()
        if not line:
            if not totals:
                print("  [!] Need at least one item.")
                continue
            break
        if line.lower() == "back":
            if not order:
                print("  [!] Nothing to undo.")
                continue
            last = order.pop()
            totals.pop(last, None); display.pop(last, None)
            print(f"  [removed {last}]")
            continue
        if "," not in line:
            print("  [!] Use: name,weight")
            continue
        name_raw, w = line.split(",", 1)
        name_raw = name_raw.strip(); w = w.strip()
        if not name_raw:
            print("  [!] Name can't be empty.")
            continue
        try:
            wt = int(w)
        except ValueError:
            print("  [!] Weight must be a whole number.")
            continue
        if wt <= 0:
            print("  [!] Weight must be positive.")
            continue
        key = norm(name_raw)
        totals[key] = wt
        display[key] = name_raw
        if key not in order:
            order.append(key)
        print(f"  [set {name_raw} = {wt}]")
    return totals, display

def input_loads(totals, display):
    while True:
        n_raw = input("\nHow many loads? ").strip()
        try:
            num_loads = int(n_raw)
        except ValueError:
            print("  [!] Enter a whole number.")
            continue
        if num_loads <= 0:
            print("  [!] Must be at least 1.")
            continue
        break

    fixed_bundles = []
    flexible_bundles = []
    flexible_grand = {}

    i = 0
    while i < num_loads:
        load = chr(ord("A") + i)

        while True:
            mode = input(f"\nLoad {load}: (F)ixed per-bundle totals, "
                         f"or (X) flexible per-bundle totals with one grand total? ").strip().lower()
            if mode in ("f", "fixed"):
                is_fixed = True
                break
            if mode in ("x", "flex", "flexible", "n", ""):
                is_fixed = False
                break
            print("  [!] Enter F or X.")

        n = None
        while n is None:
            raw = input(f"Load {load}: how many bundles? "
                        f"(or 'back' to redo this load, 'prev' to redo previous) ").strip()
            if raw.lower() == "back":
                fixed_bundles[:] = [b for b in fixed_bundles if b["load"] != load]
                flexible_bundles[:] = [b for b in flexible_bundles if b["load"] != load]
                flexible_grand.pop(load, None)
                print(f"  [cleared Load {load}]")
                continue
            if raw.lower() == "prev":
                if i == 0:
                    print("  [!] No previous load.")
                    continue
                prev = chr(ord("A") + i - 1)
                fixed_bundles[:] = [b for b in fixed_bundles if b["load"] != prev]
                flexible_bundles[:] = [b for b in flexible_bundles if b["load"] != prev]
                flexible_grand.pop(prev, None)
                i -= 1
                break
            try:
                n = int(raw)
            except ValueError:
                print("  [!] Enter a whole number.")
                continue
            if n <= 0:
                print("  [!] Must be at least 1.")
                n = None
        if n is None:
            continue

        this_load_bundles = []
        j = 0
        while j < n:
            items_str = input(f"  Bundle {j+1} items (or 'back'): ").strip()
            if items_str.lower() == "back":
                if not this_load_bundles:
                    print("  [!] Nothing to undo.")
                    continue
                this_load_bundles.pop()
                j -= 1
                continue
            items = [norm(s) for s in items_str.split(",") if s.strip()]
            if not items:
                print("  [!] List at least one item.")
                continue
            unknown = [it for it in items if it not in totals]
            if unknown:
                print(f"  [!] Unknown items: {unknown}")
                print(f"      Known: {list(display.values())}")
                continue

            if is_fixed:
                total_raw = input(f"  Bundle {j+1} total (or 'back'): ").strip()
                if total_raw.lower() == "back":
                    continue
                try:
                    total = int(total_raw)
                except ValueError:
                    print("  [!] Whole number please.")
                    continue
                if total <= 0:
                    print("  [!] Must be positive.")
                    continue
                this_load_bundles.append({"load": load, "items": items, "total": total})
            else:
                this_load_bundles.append({"load": load, "items": items})
            j += 1

        if is_fixed:
            fixed_bundles.extend(this_load_bundles)
        else:
            flexible_bundles.extend(this_load_bundles)
            while True:
                g_raw = input(f"Load {load} GRAND total (sum of all its bundles): ").strip()
                try:
                    g = int(g_raw)
                except ValueError:
                    print("  [!] Whole number please.")
                    continue
                if g <= 0:
                    print("  [!] Must be positive.")
                    continue
                flexible_grand[load] = g
                break

        i += 1

    return fixed_bundles, flexible_bundles, flexible_grand

# ============================================================
# MAIN
# ============================================================
def main():
    if not HAS_ORTOOLS:
        print("[!] OR-Tools not installed. Run: py -m pip install --user ortools")
        return

    totals, display = input_items()
    fixed_bundles, flexible_bundles, flexible_grand = input_loads(totals, display)

    for b in flexible_bundles:
        b["grand"] = flexible_grand[b["load"]]

    # Lock single-item bundles ONLY in fixed loads
    remaining = dict(totals)
    locked = []
    fixed_multi = []
    for b in fixed_bundles:
        if len(b["items"]) == 1:
            it = b["items"][0]
            remaining[it] -= b["total"]
            locked.append({"load": b["load"], "items": b["items"],
                           "weights": {it: b["total"]}, "total": b["total"],
                           "fixed": True})
        else:
            fixed_multi.append(b)

    over = [(k, v) for k, v in remaining.items() if v < 0]
    if over:
        print("\n[!] Over budget after locking fixed single-item bundles:")
        for k, v in over:
            print(f"    {display.get(k,k)}: over by {-v}")
        return

    print(f"\n[after locking fixed single-item bundles] remaining:")
    for k, v in remaining.items():
        print(f"    {display.get(k,k)}: {v}")

    rem_sum = sum(remaining.values())
    flex_sum = sum(flexible_grand.values())
    print(f"\n[check] remaining sum = {rem_sum}, flexible grand totals = {flex_sum}")
    if rem_sum != flex_sum:
        print(f"[!] Mismatch: diff = {rem_sum - flex_sum}")
    print()

    print("[using OR-Tools CP-SAT solver]")
    result, err = solve(fixed_multi, flexible_bundles, remaining, totals)
    if result is None:
        print(f"\n[no solution] {err}")
        return

    all_bundles = locked + result
    print(f"\n=== SOLUTION ===\n")

    per_load = {}
    item_totals = {k: 0 for k in totals}
    for b in all_bundles:
        per_load.setdefault(b["load"], {})
        for it, w in b["weights"].items():
            per_load[b["load"]][it] = per_load[b["load"]].get(it, 0) + w
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

    print("\n=== PER-BUNDLE DETAIL (flexible loads) ===")
    for b in result:
        if not b["fixed"]:
            print(f"  {b['load']}: {b['items']} = {b['total']}")
            for it, w in b["weights"].items():
                print(f"      {display.get(it, it)}: {w}")

if __name__ == "__main__":
    main()