import math
import sys

import pandas as pd
import streamlit as st


from itertools import combinations_with_replacement
from collections import Counter, defaultdict

CARRIER_MAX_GAP = 1000
CARRIER_END_GAP = 150
STANDARD_PRODUCT_LENGTHS = [3050, 3250, 3650, 4550, 5550, 4350, 4250, 2550]
CARRIER_TYPES = {
    "Grille 2550": "Carrier",
    "Aerofoil": "Carrier",
    "Cottal": "Aluminum Pipe",
    "Fluted": "Carrier",
    "S_Louver_1": "Carrier",
    "S_Louver_2": "Carrier",
    "C-Louvers": "Carrier",
    "Rectangular Louvers": "Carrier",
}
CARRIER_LENGTHS = {
    "Grille 2550": 3050,
    "Aerofoil": 3000,
    "Cottal": 3650,
    "Fluted": 3650,
    "S_Louver_1": 55,
    "S_Louver_2": 115.3,
    "C-Louvers": 3000,
    "Rectangular Louvers": 3000,
}
WASTAGE = 10000


def calculate_divisions(width, pitch):
    return math.ceil(width / pitch)


def filter_inv(df, product):  # , divisions):
    inv = df[(df["name"] == product)]  # & (df['stock'] >= divisions)]
    return inv


def find_counts(df):
    counts = {}
    for _, row in df.iterrows():
        if int(row["length"]) not in counts:
            counts[row["length"]] = 0
        counts[row["length"]] += row["stock"]

    return counts


def calculate_endcaps(window, orientation, divisions):

    first_ort = ["top", "left"]
    second_ort = ["bottom", "right"]
    idx = ["Vertical", "Horizontal"].index(orientation)
    first = first_ort[idx]
    second = second_ort[idx]

    st.subheader("End Caps")

    col1, col2, col3 = st.columns([1, 1, 2])

    with col1:
        first_req = st.radio(
            f"{first.capitalize()} end caps required?",
            ("Yes", "No"),
            key=f"first_{window}",
        )

    with col2:
        second_req = st.radio(
            f"{second.capitalize()} end caps required?",
            ("Yes", "No"),
            key=f"second_{window}",
        )

    endcap_cnt = 0
    if first_req == "Yes":
        endcap_cnt += divisions
    if second_req == "Yes":
        endcap_cnt += divisions

    return endcap_cnt


def find_combinations(target, lengths, counts, req_pieces):
    output = set()

    def dfs(total, arr, i):
        if target <= total <= target + WASTAGE:
            stock_avail = True
            for length in set(arr):
                req = arr.count(length) * req_pieces
                if counts[length] < req:
                    stock_avail = False
                    break
            if stock_avail:
                output.add(tuple(sorted(arr)))
            return

        if (total > target + WASTAGE) or (i >= len(lengths)):
            return

        arr.append(lengths[i])
        dfs(total + lengths[i], arr, i)
        arr.pop()
        dfs(total, arr, i + 1)

    dfs(0, [], 0)
    return [list(x) for x in output]


def calculate_carrier_distances(length_combination):

    centre_gaps_per_piece = []
    no_carriers_per_piece = []

    for length in length_combination:
        carrier_lengths = 1  # for end piece
        centre_width = length - (2 * CARRIER_END_GAP)
        if centre_width >= 900:
            carrier_lengths += math.ceil(centre_width / CARRIER_MAX_GAP)
            centre_gap = centre_width / (carrier_lengths - 1)
        else:
            carrier_lengths += 1
            centre_gap = centre_width

        no_carriers_per_piece.append(carrier_lengths)
        centre_gaps_per_piece.append(centre_gap)

    carrier_distances = []
    for j in range(len(no_carriers_per_piece)):
        curr = 0
        curr_distances = []
        for i in range(no_carriers_per_piece[j]):
            if i == 0:
                curr += CARRIER_END_GAP
            else:
                curr += centre_gaps_per_piece[j]
            curr_distances.append(math.ceil(curr))
        carrier_distances.append(curr_distances)

    return carrier_distances


def build_cut_plan(parsed_data, divisions, qty_windows):
    """
    Builds structured cut plan with:
    - single_window_qty (divisions)
    - total_window_qty (divisions * qty_windows)
    - grouped scrap with reusable tracking
    """

    window_plan = {"window_id": parsed_data["window_id"], "pieces": []}

    single_window_qty = divisions
    total_window_qty = int(divisions * qty_windows)

    for piece in parsed_data["pieces"]:

        stock_length = piece["stock_length"]
        cuts = piece["cuts"]
        used_sum = sum(cuts)
        scrap_length = stock_length - used_sum

        piece_obj = {
            "stock_length": stock_length,
            "cuts": cuts.copy(),
            "single_window_qty": single_window_qty,
            "total_window_qty": total_window_qty,
            "scrap": [],
        }

        if scrap_length > 0:
            piece_obj["scrap"].append(
                {
                    "length": scrap_length,
                    "reuse": [{"reused_in": None, "qty": single_window_qty}],
                }
            )

        window_plan["pieces"].append(piece_obj)

    return window_plan


def calc_length_combination(cut_plan):
    """
    Returns list of cut lengths used in the window
    (not multiplied by qty).
    """

    length_combination = []

    for piece in cut_plan["pieces"]:
        length_combination.extend(piece["cuts"])

    return length_combination


def display_product_pieces(length_combination, divisions):

    st.subheader("Number of pieces required for each length:")
    total_product_pieces = []
    for combo in set(length_combination):
        length_cnt_per_division = length_combination.count(combo)
        product_cnt_per_length = divisions * (length_cnt_per_division)
        total_product_pieces.append(product_cnt_per_length)
        st.write(
            "{len} mm --> {length_cnt_per_division} piece(s) \
            per division * {divisions} divisions = {number} pieces".format(
                len=combo,
                length_cnt_per_division=length_cnt_per_division,
                divisions=divisions,
                number=product_cnt_per_length,
            )
        )


def display_carrier_calc(
    product,
    height,
    length_combination,
    total_carrier_divisions,
    total_carrier_length,
    carrier_distances_per_piece,
):

    st.subheader("Carrier Calculations")

    st.write("Length option chosen: {}".format(length_combination))
    st.write(
        "{} divisions for each piece in the chosen option:".format(
            CARRIER_TYPES[product]
        )
    )

    carrier_combinations = {}
    for i in range(len(length_combination)):
        carrier_combinations["{}mm".format(length_combination[i])] = [
            carrier_distances_per_piece[i]
        ]

    carrier_table = pd.DataFrame.from_dict(carrier_combinations)

    carrier_table["Total {} divisions".format(CARRIER_TYPES[product])] = (
        total_carrier_divisions
    )

    carrier_table[
        "Total length of {x} required".format(x=CARRIER_TYPES[product].lower())
    ] = "{divs} * {length} = {tcl} mm".format(
        divs=total_carrier_divisions, length=height, tcl=total_carrier_length
    )

    carrier_table[
        "Number of {}mm {} required".format(
            CARRIER_LENGTHS[product], CARRIER_TYPES[product].lower() + "s"
        )
    ] = round(total_carrier_length / CARRIER_LENGTHS[product], 2)

    # st.write(carrier_table.T.rename_axis('Item'))
    st.write(carrier_table)


def build_req_plan(row):
    """
    Builds required piece summary.

    Args:
        cut_summary (list): List of piece lengths, e.g. [1500,1500]
        divisions (int): Number of divisions per window
        qty_areas (int): Number of windows

    Returns:
        dict: {length: total_pieces_required}
    """
    cut_summary = row["cut_summary"]
    divisions = row["divisions"]
    qty_areas = row["qty_areas"]

    req_plan = {}
    multiplier = divisions * qty_areas

    for piece in cut_summary:
        if piece not in req_plan:
            req_plan[piece] = 0

        req_plan[int(piece)] += int(multiplier)

    return req_plan


def combine_req_plan(req_plans):
    """
    Combines multiple req_plan dictionaries into one master plan.

    Args:
        req_plans (iterable): Column/array of dictionaries like {length: qty}

    Returns:
        dict: Combined requirement plan
    """

    master_plan = {}

    for plan in req_plans:
        if not plan:
            continue

        for length, qty in plan.items():
            master_plan[int(length)] = int(master_plan.get(length, 0) + qty)

    return dict(sorted(master_plan.items()))


"""
Cutting Stock Optimizer
-----------------------
Standard stock lengths : 3050, 3650, 4550 mm
Kerf (blade loss)      : 10 mm per cut
Algorithm              : Column-generation + ILP (PuLP / CBC)
Fallback               : Fast greedy if PuLP not installed
"""

KERF = 5  # mm lost per cut
STANDARD_LENGTHS = [2450, 3050, 3250, 3650, 4550]


# ──────────────────────────────────────────────
# Pattern generation (kerf-aware)
# ──────────────────────────────────────────────


def generate_patterns(piece_lengths, stock_lengths, max_cuts=10):
    """
    Return all feasible cutting patterns across all stock lengths.
    Each pattern accounts for kerf on every cut (n pieces → n kerfs).
    """
    patterns = []
    for stock in stock_lengths:
        for n_cuts in range(1, max_cuts + 1):
            for combo in combinations_with_replacement(piece_lengths, n_cuts):
                # n pieces require n saw cuts
                used = sum(combo) + n_cuts * KERF
                if used > stock:
                    continue  # combinations_with_replacement is ordered, so we can break
                waste = stock - used
                patterns.append(
                    {
                        "stock_length": stock,
                        "cuts": list(combo),
                        "waste": waste,
                        "n_cuts": n_cuts,
                    }
                )
    return patterns


# ──────────────────────────────────────────────
# ILP solver  (requires: pip install pulp)
# ──────────────────────────────────────────────


def solve_ilp(req_plan, stock_lengths):
    """
    Minimise total stock used (weighted by length) subject to demand constraints.
    Uses PuLP with the bundled CBC solver — typically solves in milliseconds.
    """
    try:
        import pulp
    except ImportError:
        return None, "PuLP not installed"
    # print("PuLP version:", pulp.__version__)
    piece_lengths = sorted(req_plan.keys(), reverse=True)
    # print("Normalized req:", req_plan)
    patterns = generate_patterns(piece_lengths, stock_lengths)

    # print(f"Total patterns generated: {len(patterns)}")
    # print("First 10 patterns:")
    # for p in patterns[:10]:
    #     print(" ", p)

    if not patterns:
        return [], "No feasible patterns"

    prob = pulp.LpProblem("CuttingStock", pulp.LpMinimize)

    # Decision variable: how many times to use each pattern
    x = [
        pulp.LpVariable(f"x{i}", lowBound=0, cat="Integer")
        for i in range(len(patterns))
    ]

    # Objective: minimise total stock length consumed
    prob += pulp.lpSum(x[i] * patterns[i]["waste"] for i in range(len(patterns)))

    # Demand constraints: each piece length must be covered
    for length, qty in req_plan.items():
        piece_count = pulp.lpSum(
            x[i] * patterns[i]["cuts"].count(length) for i in range(len(patterns))
        )
        prob += piece_count >= qty, f"demand_min_{length}"

    prob.solve(pulp.COIN_CMD(path="/opt/homebrew/bin/cbc", msg=0))

    if pulp.LpStatus[prob.status] != "Optimal":
        return None, f"Solver status: {pulp.LpStatus[prob.status]}"

    # Expand solution into one dict per stock bar used
    result = []
    for i, var in enumerate(x):
        count = int(round(var.varValue or 0))
        for _ in range(count):
            result.append(patterns[i].copy())

    return result, "optimal"


# ──────────────────────────────────────────────
# Greedy fallback  (no extra dependencies)
# ──────────────────────────────────────────────


def solve_greedy(req_plan, stock_lengths):
    """
    Fast greedy: always pick the pattern that satisfies the most remaining
    demand with the least waste.  O(patterns × iterations).
    """
    piece_lengths = sorted(req_plan.keys(), reverse=True)
    patterns = generate_patterns(piece_lengths, stock_lengths)

    # Pre-sort: least waste first, prefer larger stock as tiebreak
    patterns.sort(key=lambda p: (p["waste"], -p["stock_length"]))

    req = req_plan.copy()
    result = []

    while any(q > 0 for q in req.values()):
        best = None
        best_score = (-1, float("inf"))  # (pieces_satisfied, waste)

        for p in patterns:
            # Check feasibility
            needed = defaultdict(int)
            for c in p["cuts"]:
                needed[c] += 1
            if any(req.get(c, 0) < cnt for c, cnt in needed.items()):
                continue

            # Score: how many units of demand does this pattern consume?
            satisfied = sum(min(req.get(c, 0), cnt) for c, cnt in needed.items())
            score = (satisfied, -p["waste"])
            if score > best_score:
                best_score = score
                best = p

        if best is None:
            # Shouldn't happen with valid inputs
            raise ValueError("No feasible pattern found — check your req_plan.")

        for c in best["cuts"]:
            req[c] -= 1
        result.append(best.copy())

    return result, "greedy"


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────


def trim_result(result, req_plan):
    remaining = {int(k): int(v) for k, v in req_plan.items()}
    trimmed = []

    for bar in result:
        new_cuts = []
        for cut in bar["cuts"]:
            cut = int(cut)
            if remaining.get(cut, 0) > 0:
                new_cuts.append(cut)
                remaining[cut] -= 1

        if new_cuts:
            # Find smallest standard stock that fits these cuts + kerf
            min_length_needed = sum(new_cuts) + len(new_cuts) * KERF
            best_stock = None
            for stock in sorted(STANDARD_LENGTHS):
                if stock >= min_length_needed:
                    best_stock = stock
                    break
            if best_stock is None:
                best_stock = bar["stock_length"]  # fallback to original

            new_waste = best_stock - sum(new_cuts) - len(new_cuts) * KERF
            trimmed.append(
                {
                    "stock_length": best_stock,
                    "cuts": new_cuts,
                    "waste": new_waste,
                    "n_cuts": len(new_cuts),
                }
            )

    return trimmed


def solve_ilp_exact(req_plan, stock_lengths):
    """Second pass: exact quantities, minimise waste globally."""
    try:
        import pulp
        import shutil
    except ImportError:
        return None, "PuLP not installed"

    req_plan = {int(k): int(v) for k, v in req_plan.items()}
    piece_lengths = sorted(req_plan.keys(), reverse=True)
    patterns = generate_patterns(piece_lengths, stock_lengths)

    prob = pulp.LpProblem("CuttingStockExact", pulp.LpMinimize)
    x = [
        pulp.LpVariable(f"x{i}", lowBound=0, cat="Integer")
        for i in range(len(patterns))
    ]

    # Minimise total waste
    prob += pulp.lpSum(x[i] * patterns[i]["waste"] for i in range(len(patterns)))

    # EXACT constraints — no over or under
    for length, qty in req_plan.items():
        piece_count = pulp.lpSum(
            x[i] * patterns[i]["cuts"].count(length) for i in range(len(patterns))
        )
        prob += piece_count == qty, f"exact_{length}"

    cbc_path = shutil.which("cbc") or "/opt/homebrew/bin/cbc"
    prob.solve(pulp.COIN_CMD(path=cbc_path, msg=0))

    if pulp.LpStatus[prob.status] != "Optimal":
        return None, f"Solver status: {pulp.LpStatus[prob.status]}"

    result = []
    for i, var in enumerate(x):
        count = int(round(var.varValue or 0))
        for _ in range(count):
            result.append(patterns[i].copy())

    return result, "optimal-exact"


def optimize_stock(req_plan, stock_lengths=None, prefer="ilp"):
    if stock_lengths is None:
        stock_lengths = STANDARD_LENGTHS

    stock_lengths = sorted(stock_lengths)
    req_plan = {int(k): int(v) for k, v in req_plan.items()}

    if prefer == "ilp":
        # Pass 1: get correct piece counts
        result, method = solve_ilp(req_plan, stock_lengths)
        if result is None:
            print(f"[ILP fallback] {method} — switching to greedy.", file=sys.stderr)
            result, method = solve_greedy(req_plan, stock_lengths)
            result = trim_result(result, req_plan)
        else:
            # Pass 2: re-run ILP with tighter objective now that we know exact counts
            # Count what trim would produce, then re-solve cleanly
            trimmed = trim_result(result, req_plan)

            # Build exact req from trimmed (no over/under)
            exact_req = Counter()
            for bar in trimmed:
                for cut in bar["cuts"]:
                    exact_req[cut] += 1

            # Re-solve with exact quantities and == constraints
            result2, method2 = solve_ilp_exact(dict(exact_req), stock_lengths)
            if result2:
                result = result2
                method = method2 + " (2-pass)"
            else:
                result = trimmed
    else:
        result, method = solve_greedy(req_plan, stock_lengths)
        result = trim_result(result, req_plan)

    # ... rest of summary building

    # ── Summary stats ──
    total_bars = len(result)
    total_material = sum(b["stock_length"] for b in result)
    total_waste = sum(b["waste"] for b in result)
    total_kerf_loss = sum(b["n_cuts"] * KERF for b in result)
    efficiency = (1 - total_waste / total_material) * 100 if total_material else 0

    bar_counts = defaultdict(int)
    for b in result:
        bar_counts[b["stock_length"]] += 1

    summary = {
        "total_bars": total_bars,
        "bar_breakdown": dict(bar_counts),
        "total_material_mm": total_material,
        "total_waste_mm": total_waste,
        "total_kerf_loss_mm": total_kerf_loss,
        "efficiency_pct": round(efficiency, 2),
        "method": method,
    }

    return {"bars": result, "summary": summary, "method": method}


def cross_window_optimize(per_window_bars, stock_lengths, min_saving=200):

    good_bars = [b for b in per_window_bars if b["waste"] < 200]
    candidates = [b for b in per_window_bars if b["waste"] >= 200]

    if not candidates:
        return per_window_bars

    # Pool from high-waste bars
    pool_req = {}
    for bar in candidates:
        for cut in bar["cuts"]:
            pool_req[cut] = pool_req.get(cut, 0) + 1

    # ── NEW: also pull in small pieces from good bars that have room ──
    # Any piece that is <= smallest stock length / 3 is a candidate filler
    MAX_FILLER = min(stock_lengths) // 3  # e.g. 2450//3 = 816mm
    filler_req = {}
    filler_bars_used = []

    for bar in good_bars:
        # If this bar only has small pieces, it might combine better with high-waste bars
        if all(c <= MAX_FILLER for c in bar["cuts"]):
            for cut in bar["cuts"]:
                filler_req[cut] = filler_req.get(cut, 0) + 1
            filler_bars_used.append(bar)

    # Combined pool: high-waste pieces + small filler pieces
    combined_req = {**pool_req}
    for k, v in filler_req.items():
        combined_req[k] = combined_req.get(k, 0) + v

    # print(f"[Cross-window] Re-optimizing {len(candidates)} high-waste bars "
    #       f"+ {len(filler_bars_used)} small-piece bars...")

    re_result, method = solve_ilp_exact(combined_req, stock_lengths)
    if re_result is None:
        return per_window_bars

    # Compare total waste
    original_waste = sum(b["waste"] for b in candidates) + sum(
        b["waste"] for b in filler_bars_used
    )
    new_waste = sum(b["waste"] for b in re_result)
    saving = original_waste - new_waste

    if saving >= min_saving:
        # print(f"[Cross-window] Saving {saving}mm by recombining "
        #       f"({original_waste}mm → {new_waste}mm)")
        # good_bars minus the filler bars we pulled out, plus re-solved result
        remaining_good = [b for b in good_bars if b not in filler_bars_used]
        return remaining_good + re_result
    else:
        # print(f"[Cross-window] Only {saving}mm saving — keeping per-window plan")
        return per_window_bars


# def optimize_stock_v2(data, stock_lengths=None):
#     """
#     Full 2-stage optimizer.
#     data: DataFrame with req_plan column (per window)
#     """
#     if stock_lengths is None:
#         stock_lengths = STANDARD_LENGTHS

#     all_bars = []

#     # ── Stage 1: Per-window solve ──
#     print("=== Stage 1: Per-window optimization ===")
#     for idx, row in data.iterrows():
#         window_req = {int(k): int(v) for k, v in row["req_plan"].items()}
#         if not window_req:
#             continue

#         result, method = solve_ilp(window_req, stock_lengths)
#         if result is None:
#             result, _ = solve_greedy(window_req, stock_lengths)
#         result = trim_result(result, window_req)

#         for bar in result:
#             bar["window_id"] = idx
#         all_bars.extend(result)

#     stage1_waste = sum(b["waste"] for b in all_bars)
#     print(f"Stage 1 total waste: {stage1_waste}mm across {len(all_bars)} bars")

#     # ── Stage 2: Cross-window re-optimization of high-waste bars ──
#     print("\n=== Stage 2: Cross-window gap optimization ===")
#     all_bars = cross_window_optimize(all_bars, stock_lengths, min_saving=200)

#     stage2_waste = sum(b["waste"] for b in all_bars)
#     print(f"Stage 2 total waste: {stage2_waste}mm  "
#           f"(saved {stage1_waste - stage2_waste}mm)")

#     # ── Build summary ──
#     total_material = sum(b["stock_length"] for b in all_bars)
#     total_waste = sum(b["waste"] for b in all_bars)
#     total_kerf_loss = sum(b["n_cuts"] * KERF for b in all_bars)
#     efficiency = (1 - total_waste / total_material) * 100 if total_material else 0

#     bar_counts = defaultdict(int)
#     for b in all_bars:
#         bar_counts[b["stock_length"]] += 1

#     summary = {
#         "total_bars"        : len(all_bars),
#         "bar_breakdown"     : dict(bar_counts),
#         "total_material_mm" : total_material,
#         "total_waste_mm"    : total_waste,
#         "total_kerf_loss_mm": total_kerf_loss,
#         "efficiency_pct"    : round(efficiency, 2),
#         "method"            : "2-stage",
#         "stage1_waste_mm"   : stage1_waste,
#         "stage2_waste_mm"   : stage2_waste,
#         "cross_window_saving_mm": stage1_waste - stage2_waste,
#     }

#     return {"bars": all_bars, "summary": summary}


def build_cut_pool(all_bars):
    """Build a pool of {(length, window_id): qty} from Stage 1 bars."""
    pool = defaultdict(int)
    for bar in all_bars:
        for cut in bar["cuts_detail"]:
            pool[(cut["length"], cut["window_id"])] += 1
    return pool


# def assign_window_ids(cuts, cut_pool):
#     """
#     Given a list of cut lengths, assign window_ids from the pool.
#     Greedy: for each length, pick the window_id with the most remaining qty.
#     """
#     cuts_detail = []
#     for length in cuts:
#         # Find all window_ids that have this length available
#         candidates = {
#             wid: qty for (l, wid), qty in cut_pool.items()
#             if l == length and qty > 0
#         }
#         if candidates:
#             # Pick window with most remaining cuts of this length (greedy)
#             best_wid = max(candidates, key=candidates.get)
#             cut_pool[(length, best_wid)] -= 1
#             cuts_detail.append({"length": length, "window_id": best_wid})
#         else:
#             # Shouldn't happen, but fallback
#             cuts_detail.append({"length": length, "window_id": "unknown"})
#     return cuts_detail


def assign_window_ids(cuts, cut_pool, req_plans=None):
    """
    Assign window_ids to cuts from the pool.
    If req_plans provided ({window_id: {length: qty}}), use demand-aware assignment
    to ensure every window gets exactly what it needs.
    Falls back to greedy if no demand info.
    """
    cuts_detail = []
    for length in cuts:
        # Find all window_ids that still have demand for this length
        if req_plans:
            # Prefer windows that still have unfulfilled demand
            candidates = {
                wid: qty
                for (l, wid), qty in cut_pool.items()
                if l == length and qty > 0
            }

            # Score: windows with remaining demand get priority
            def demand_score(wid):
                remaining_demand = req_plans.get(wid, {}).get(length, 0)
                pool_qty = cut_pool.get((length, wid), 0)
                return (remaining_demand > 0, pool_qty)

            if candidates:
                best_wid = max(candidates, key=demand_score)
                cut_pool[(length, best_wid)] -= 1
                if req_plans.get(best_wid, {}).get(length, 0) > 0:
                    req_plans[best_wid][length] -= 1
                cuts_detail.append({"length": length, "window_id": best_wid})
            else:
                cuts_detail.append({"length": length, "window_id": "unknown"})
        else:
            # Original greedy fallback
            candidates = {
                wid: qty
                for (l, wid), qty in cut_pool.items()
                if l == length and qty > 0
            }
            if candidates:
                best_wid = max(candidates, key=candidates.get)
                cut_pool[(length, best_wid)] -= 1
                cuts_detail.append({"length": length, "window_id": best_wid})
            else:
                cuts_detail.append({"length": length, "window_id": "unknown"})

    return cuts_detail


def absorb_singleton_bars(all_bars, stock_lengths):
    """
    Find bars with unique patterns (qty=1) where the same cuts exist
    on a different stock length with higher qty. Re-assign to dominant stock.
    """
    from collections import Counter

    # Count how many times each (stock_length, cuts) pattern appears
    pattern_counts = Counter()
    for bar in all_bars:
        key = (bar["stock_length"], tuple(sorted(bar["cuts"])))
        pattern_counts[key] += 1

    absorbed = []
    for bar in all_bars:
        key = (bar["stock_length"], tuple(sorted(bar["cuts"])))
        if pattern_counts[key] > 1:
            absorbed.append(bar)
            continue

        # singleton — check if same cuts exist on another stock length with qty > 1
        cuts_key = tuple(sorted(bar["cuts"]))
        dominant = None
        best_qty = 1
        for (stock, ck), qty in pattern_counts.items():
            if ck == cuts_key and stock != bar["stock_length"] and qty > best_qty:
                # verify cuts still fit in this stock length
                min_needed = sum(bar["cuts"]) + len(bar["cuts"]) * KERF
                if stock >= min_needed:
                    dominant = stock
                    best_qty = qty

        if dominant:
            # re-assign this bar to dominant stock length
            new_waste = dominant - sum(bar["cuts"]) - len(bar["cuts"]) * KERF
            bar = bar.copy()
            bar["stock_length"] = dominant
            bar["waste"] = new_waste
            # update pattern_counts
            pattern_counts[key] -= 1
            new_key = (dominant, cuts_key)
            pattern_counts[new_key] += 1

        absorbed.append(bar)

    return absorbed


def optimize_stock_v2(data, stock_lengths=None):
    if stock_lengths is None:
        stock_lengths = STANDARD_LENGTHS

    stock_lengths = sorted(stock_lengths)
    all_bars = []

    # ── Stage 1: Per-window solve ──
    # print("=== Stage 1: Per-window optimization ===")
    for idx, row in data.iterrows():
        window_req = {int(k): int(v) for k, v in row["req_plan"].items()}
        if not window_req:
            continue

        result, method = solve_ilp(window_req, stock_lengths)
        if result is None:
            result, _ = solve_greedy(window_req, stock_lengths)
        result = trim_result(result, window_req)

        for bar in result:
            bar["cuts_detail"] = [{"length": c, "window_id": idx} for c in bar["cuts"]]
            bar["window_id"] = idx
        all_bars.extend(result)

    stage1_waste = sum(b["waste"] for b in all_bars)
    # print(f"Stage 1 total waste: {stage1_waste}mm across {len(all_bars)} bars")

    # ── Build cut pool BEFORE Stage 2 ──
    cut_pool = build_cut_pool(all_bars)

    # # ── Stage 2: Cross-window re-optimization ──
    # # print("\n=== Stage 2: Cross-window gap optimization ===")
    # all_bars = cross_window_optimize(all_bars, stock_lengths, min_saving=200)

    # # ── Reassign window_ids to all bars using the pool ──
    # for bar in all_bars:
    #     bar["cuts_detail"] = assign_window_ids(bar["cuts"], cut_pool)
    #     bar["window_id"] = list(set(c["window_id"] for c in bar["cuts_detail"]))
    #     if len(bar["window_id"]) == 1:
    #         bar["window_id"] = bar["window_id"][0]  # unwrap if single window

    demand_map = {}
    for idx, row in data.iterrows():
        if row["req_plan"]:
            demand_map[idx] = {int(k): int(v) for k, v in row["req_plan"].items()}

    # pool_4900 = {k: v for k, v in cut_pool.items() if k[0] == 4900}

    # Reassign window_ids using demand-aware assignment
    for bar in all_bars:
        bar["cuts_detail"] = assign_window_ids(
            bar["cuts"], cut_pool, req_plans=demand_map
        )
        bar["window_id"] = list(set(c["window_id"] for c in bar["cuts_detail"]))
        if len(bar["window_id"]) == 1:
            bar["window_id"] = bar["window_id"][0]

    stage2_waste = sum(b["waste"] for b in all_bars)
    # print(f"Stage 2 total waste: {stage2_waste}mm  (saved {stage1_waste - stage2_waste}mm)")

    # ── Build structured stock_plan ──
    grouped = {}
    for bar in all_bars:
        key = (bar["stock_length"], tuple(sorted(bar["cuts"])))
        if key not in grouped:
            grouped[key] = {
                "length": bar["stock_length"],
                "qty": 0,
                "cuts": bar["cuts_detail"].copy(),
                "waste": bar["waste"],
                "window_ids": [],
            }
        grouped[key]["qty"] += 1
        grouped[key]["window_ids"] = list(
            set(
                grouped[key]["window_ids"]
                + [c["window_id"] for c in bar["cuts_detail"]]
            )
        )

    stock_plan = sorted(grouped.values(), key=lambda x: -x["qty"])

    # ── Summary ──
    total_material = sum(b["stock_length"] for b in all_bars)
    total_waste = sum(b["waste"] for b in all_bars)
    total_kerf_loss = sum(b["n_cuts"] * KERF for b in all_bars)
    efficiency = (1 - total_waste / total_material) * 100 if total_material else 0

    bar_counts = defaultdict(int)
    for b in all_bars:
        bar_counts[b["stock_length"]] += 1

    summary = {
        "total_bars": len(all_bars),
        "bar_breakdown": dict(bar_counts),
        "total_material_mm": total_material,
        "total_waste_mm": total_waste,
        "total_kerf_loss_mm": total_kerf_loss,
        "efficiency_pct": round(efficiency, 2),
        "method": "2-stage",
        "stage1_waste_mm": stage1_waste,
        "stage2_waste_mm": stage2_waste,
        "cross_window_saving_mm": stage1_waste - stage2_waste,
    }

    return {
        "bars": all_bars,
        "stock_plan": stock_plan,
        "summary": summary,
    }


def build_window_cut_plan(row, bars):
    """Pass out["bars"] not out["stock_plan"]"""
    window_id = row.name
    cut_summary = [int(c) for c in row["cut_summary"]]

    window_bars = []
    for bar in bars:
        window_cuts = [
            c["length"] for c in bar["cuts_detail"] if c["window_id"] == window_id
        ]
        if not window_cuts:
            continue
        window_bars.append(
            {
                "stock_length": bar["stock_length"],
                "cuts": window_cuts,
            }
        )

    # Group identical patterns and count qty correctly
    from collections import Counter

    grouped = Counter()
    pattern_cuts = {}
    for entry in window_bars:
        key = (entry["stock_length"], tuple(sorted(entry["cuts"])))
        grouped[key] += 1
        pattern_cuts[key] = entry["cuts"]

    result = [
        {
            "stock_length": stock,
            "qty": qty,
            "cuts": pattern_cuts[(stock, cuts_key)],
        }
        for (stock, cuts_key), qty in grouped.items()
    ]

    def sort_key(entry):
        positions = [cut_summary.index(c) for c in entry["cuts"] if c in cut_summary]
        return min(positions) if positions else 999

    result.sort(key=sort_key)
    return result
