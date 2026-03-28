import math
import pandas as pd
import streamlit as st

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, COMMON_ACCESSORIES

# Grille specific codes from excel_utils.py
GRILLE_PRODUCTS = {
    "PROFILE": ("GR-2550-PR-01", "GRILLE 2550 PROFILE"),
    "CARRIER": ("GR-CA-01", "GRILLE CARRIER"),
    "COVERING": ("GR-CP-01", "GRILLE COVERING PROFILE"),
    "L_ENDCAP": ("AC-GR-25EL", "GRILLE 2550 L-ENDCAP"),
    "INVERSE_L_ENDCAP": ("AC-GR-25ER", "GRILLE 2550 INVERSE-L ENDCAP"),
    "JOINING_PIECES": ("AC-GR-25JP", "GRILLE 2550 JOINING PIECES"),
}
ENDCAP_OPTIONS = [
    "No Endcaps",
    "Both sides",
    "Single Side - L",
    "Single Side - Inverse L",
]


def calculate_endcaps(row):

    endcap_opt = row["endcaps"]
    divisions = row["divisions"]
    num_L, num_inverse_L = 0, 0

    if endcap_opt in ["Single Side - L", "Both sides"]:
        num_L += divisions

    if endcap_opt in ["Single Side - Inverse L", "Both sides"]:
        num_inverse_L += divisions

    return num_L, num_inverse_L


def generate_offer_df(data):

    offer_df_cols = {
        "s_no": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "area_name": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "orientation": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "height": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "width": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "qty_areas": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "area_sqft": {
            "type": "formula",
            "hide_if_zero": False,
        },
        "divisions": {
            "type": "formula",
            "hide_if_zero": False,
        },
        "total_product_length": {
            "type": "formula",
            "hide_if_zero": False,
        },
        "joining_pieces": {
            "type": "formula",
            "hide_if_zero": True,
        },
        "endcap_cnt": {
            "type": "formula",
            "hide_if_zero": True,
        },
    }
    offer_df = data[offer_df_cols.keys()].copy()

    return offer_df_cols, offer_df


def generate_inventory_df(data, pitch, stock_plan):

    all_inventory_rows = []

    for _, row in data.iterrows():

        sequence = 0

        profile_rows = [
            {
                "Product Code": GRILLE_PRODUCTS["PROFILE"][0],
                "Product Name": GRILLE_PRODUCTS["PROFILE"][1],
                "Length": item["length"],
                "Quantity": item["qty"],
                "UOM": "m",
            }
            for item in stock_plan
        ]

        additional_items = [
            {
                "Product Code": GRILLE_PRODUCTS["CARRIER"][0],
                "Product Name": GRILLE_PRODUCTS["CARRIER"][1],
                "Length": 3050,
                "Quantity": int(math.ceil(row["total_carrier_length"] / 3050)),
                "UOM": "m",
                "CNC Hole Distance": pitch,
            },
            {
                "Product Code": GRILLE_PRODUCTS["COVERING"][0],
                "Product Name": GRILLE_PRODUCTS["COVERING"][1],
                "Length": 3050,
                "Quantity": int(math.ceil(row["total_carrier_length"] / 3050)),
                "UOM": "m",
            },
        ]

        conditional_items = []

        if row["num_L"] > 0:
            conditional_items.append(
                {
                    "Product Code": GRILLE_PRODUCTS["L_ENDCAP"][0],
                    "Product Name": GRILLE_PRODUCTS["L_ENDCAP"][1],
                    "Quantity": row["num_L"],
                    "UOM": "pcs",
                }
            )

        if row["num_inverse_L"] > 0:
            conditional_items.append(
                {
                    "Product Code": GRILLE_PRODUCTS["INVERSE_L_ENDCAP"][0],
                    "Product Name": GRILLE_PRODUCTS["INVERSE_L_ENDCAP"][1],
                    "Quantity": row["num_inverse_L"],
                    "UOM": "pcs",
                }
            )

        if row["joining_pieces"] > 0:
            conditional_items.append(
                {
                    "Product Code": GRILLE_PRODUCTS["JOINING_PIECES"][0],
                    "Product Name": GRILLE_PRODUCTS["JOINING_PIECES"][1],
                    "Quantity": row["joining_pieces"],
                    "UOM": "pcs",
                }
            )

        accs = [
            {
                "Product Code": COMMON_ACCESSORIES["NUT_3_5X25"][0],
                "Product Name": COMMON_ACCESSORIES["NUT_3_5X25"][1],
                "Quantity": row["divisions"] * row["total_carrier_divisions"],
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["BOLT_3_5X25"][0],
                "Product Name": COMMON_ACCESSORIES["BOLT_3_5X25"][1],
                "Quantity": row["divisions"] * row["total_carrier_divisions"],
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["WASHER"][0],
                "Product Name": COMMON_ACCESSORIES["WASHER"][1],
                "Quantity": row["divisions"] * row["total_carrier_divisions"],
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_25MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_25MM"][1],
                "Quantity": int(math.ceil(row["total_carrier_length"] / 300)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PAINT"][0],
                "Product Name": COMMON_ACCESSORIES["PAINT"][1],
                "Quantity": int(math.ceil(((row["divisions"] / 100 * 2) / 2))),
                "UOM": "l",
            },
        ]

        # combine everything
        row_items = profile_rows + additional_items + conditional_items + accs

        # multiply by qty_areas BEFORE appending
        for item in row_items:
            item["Quantity"] = int(item["Quantity"] * row["qty_areas"])
            item["item_order"] = sequence
            sequence += 1

        all_inventory_rows.extend(row_items)

    inv_data = (
        pd.DataFrame(all_inventory_rows)
        .reindex(columns=INV_COLUMNS + ["item_order"])
        .fillna("")
    )

    inv_data = (
        inv_data.groupby(
            [
                "Product Code",
                "Product Name",
                "Length",
                "UOM",
                "Colour",
                "Finish",
                "CNC Hole Distance",
                "Remarks",
            ],
            as_index=False,
            sort=False,
        )
        .agg({"Quantity": "sum", "item_order": "min"})
        .sort_values(["item_order"])
        .drop(columns=["item_order"])
    )

    return inv_data


class GrilleCalculator:
    def __init__(self, vars):
        self.project_title = vars["project_title"]
        self.pitch = vars["pitch"]
        self.areas = vars["areas"]

    def validate_input(row, idx):

        VERTICAL_LOUVER = {"Top L", "Top Inverse L"}
        HORIZONTAL_LOUVER = {"Right L"}

        louver = str(row["louver_direction"]).strip()
        orientation = str(row["orientation"]).strip()

        if orientation == "Horizontal":
            if louver not in HORIZONTAL_LOUVER:
                st.warning(
                    f"Row {idx + 1}: Louver direction '{louver}' not valid for Horizontal."
                )
                return False
        else:
            if louver not in VERTICAL_LOUVER:
                st.warning(
                    f"Row {idx + 1}: Louver direction '{louver}' not valid for Vertical."
                )
                return False

        return True

    def get_data_input():

        empty_df = pd.DataFrame(
            {
                "s_no": pd.Series(dtype="str"),
                "area_name": pd.Series(dtype="str"),
                "height": pd.Series(dtype="int"),
                "width": pd.Series(dtype="int"),
                "orientation": pd.Series(dtype="str"),
                "louver_direction": pd.Series(dtype="str"),
                "endcaps": pd.Series(dtype="str"),
                "qty_areas": pd.Series(dtype="int"),
                "cut_summary": pd.Series(dtype="str"),
            }
        )

        required_cols = [
            "width",
            "height",
            "orientation",
            "louver_direction",
            "endcaps",
            "qty_areas",
            "cut_summary",
        ]

        input_data = st.data_editor(
            data=empty_df,
            column_config={
                "s_no": st.column_config.TextColumn("S. No", required=False),
                "area_name": st.column_config.TextColumn("Area Name", required=False),
                "width": st.column_config.NumberColumn(
                    "Width (mm)", min_value=1, step=1, required=True
                ),
                "height": st.column_config.NumberColumn(
                    "Height (mm)", min_value=1, step=1, required=True
                ),
                "pitch": st.column_config.NumberColumn(
                    "Pitch (mm)", min_value=1, step=1, required=True
                ),
                "orientation": st.column_config.SelectboxColumn(
                    "Orientation", options=["Horizontal", "Vertical"], required=True
                ),
                "louver_direction": st.column_config.SelectboxColumn(
                    "Grille Direction",
                    options=["Top L", "Top Inverse L", "Right L"],
                    required=True,
                ),
                "endcaps": st.column_config.SelectboxColumn(
                    "Endcaps", options=ENDCAP_OPTIONS, required=True
                ),
                "qty_areas": st.column_config.NumberColumn(
                    "Similar Areas", min_value=1, step=1, required=True
                ),
                "cut_summary": st.column_config.TextColumn(
                    "Cut Summary", required=True
                ),
            },
            num_rows="dynamic",
        )

        return input_data, required_cols

    # def print_stuff(out, master_req):

    #     from collections import Counter

    #     # Check raw ILP output BEFORE trimming
    #     print("Total bars:", len(out["bars"]))

    #     piece_totals_raw = Counter()
    #     for bar in out["bars"]:
    #         for cut in bar["cuts"]:
    #             piece_totals_raw[cut] += 1

    #     print("Raw pieces from ILP:")
    #     for length, qty in sorted(piece_totals_raw.items(), reverse=True):
    #         required = int(master_req.get(length, master_req.get(float(length), 0)))
    #         print(f"  {length}mm → {qty} pcs (required: {required})")

    #     print("\n===== CUT PLAN =====")
    #     groups = Counter()
    #     pattern_lookup = {}
    #     for bar in out["bars"]:
    #         key = (bar["stock_length"], tuple(sorted(bar["cuts"])))
    #         groups[key] += 1
    #         pattern_lookup[key] = bar

    #     for (stock, cuts), qty in sorted(groups.items(), key=lambda x: -x[1]):
    #         bar = pattern_lookup[(stock, cuts)]
    #         print(
    #             f"  {qty:>4}x  {stock}mm  →  cuts: {list(cuts)}  |  waste/bar: {bar['waste']}mm  |  total waste: {bar['waste']*qty}mm"
    #         )

    #     # Piece tally
    #     piece_totals = Counter()
    #     for bar in out["bars"]:
    #         for cut in bar["cuts"]:
    #             piece_totals[cut] += 1

    #     s = out["summary"]
    #     print("\n===== SUMMARY =====")
    #     print(f"Total bars used : {s['total_bars']}")
    #     print(f"Bar breakdown   : {s['bar_breakdown']}")
    #     print(f"Total waste     : {s['total_waste_mm']} mm")
    #     print(f"Kerf loss       : {s['total_kerf_loss_mm']} mm")
    #     print(f"Efficiency      : {s['efficiency_pct']} %")

    #     print("\n===== PIECES CUT =====")
    #     for length, qty in sorted(piece_totals.items(), reverse=True):
    #         required = int(master_req.get(length, master_req.get(float(length), 0)))
    #         diff = qty - required
    #         if diff == 0:
    #             status = "✓"
    #         elif diff > 0:
    #             status = f"↑ {diff} extra"
    #         else:
    #             status = f"↓ {abs(diff)} short"
    #         print(f"  {length}mm  →  {qty} pcs  (required: {required})  {status}")

    #     print(master_req)

    #     s = out["summary"]
    #     waste_pct = round(s["total_waste_mm"] / s["total_material_mm"] * 100, 2)
    #     kerf_pct = round(s["total_kerf_loss_mm"] / s["total_material_mm"] * 100, 2)
    #     used_pct = round(100 - waste_pct - kerf_pct, 2)

    #     print("\n===== MATERIAL BREAKDOWN =====")
    #     print(f"  Useful cuts : {used_pct} %")
    #     print(f"  Kerf loss   : {kerf_pct} %")
    #     print(f"  Waste       : {waste_pct} %")

    #     st.badge(f"Cut plan generated with {used_pct}% efficiency", color="green")

    def generate_image(row, common_vars):

        pitch = common_vars["pitch"]
        divisions = row["divisions"]
        cut_summary = row["cut_summary"]
        louver = str(row["louver_direction"]).strip()
        endcap_choice = str(row["endcaps"]).strip()
        num_L = row["num_L"]
        num_inverse_L = row["num_inverse_L"]
        orient = row["orientation"]

        if len(cut_summary) > 1:
            cut_summary_str = "+".join(str(c) for c in cut_summary)
        else:
            cut_summary_str = f"Single Piece {cut_summary[0]} mm"

        info_lines = [
            (f"{divisions} Divisions", "#333", 12, True),
            (f"@ {pitch} mm Pitch", "#333", 12, True),
            ("", "#333", 12, True),
            ("Breakdown", "#333", 12, False),
            (f"{cut_summary_str}", "#333", 12, True),
            ("Grille Direction", "#333", 12, False),
            (f"{louver}", "#333", 12, True),
        ]

        # Compute endcap sides
        draw_top = draw_bottom = draw_left = draw_right = False

        if orient == "Vertical":
            if louver == "Top L":
                draw_top = (
                    endcap_choice in ("Both sides", "Single Side - L") and num_L > 0
                )
                draw_bottom = (
                    endcap_choice in ("Both sides", "Single Side - Inverse L")
                    and num_inverse_L > 0
                )
            elif louver == "Top Inverse L":
                draw_top = (
                    endcap_choice in ("Both sides", "Single Side - Inverse L")
                    and num_inverse_L > 0
                )
                draw_bottom = (
                    endcap_choice in ("Both sides", "Single Side - L") and num_L > 0
                )
        else:
            if louver == "Right L":
                draw_right = (
                    endcap_choice in ("Both sides", "Single Side - L") and num_L > 0
                )
                draw_left = (
                    endcap_choice in ("Both sides", "Single Side - Inverse L")
                    and num_inverse_L > 0
                )

        endcap_sides = {
            "top": draw_top,
            "bottom": draw_bottom,
            "left": draw_left,
            "right": draw_right,
        }

        config = {
            "show_carriers": True,
            "show_endcaps": True,
            "show_joints": True,
            "bar_color": "#888",
            "info_lines": info_lines,
            "endcap_sides": endcap_sides,
        }
        return config

    def run(self):

        data = self.areas.copy()

        # PROFILE
        data["req_plan"] = data.apply(profile_utils.build_req_plan, axis=1)
        # master_req = profile_utils.combine_req_plan(data["req_plan"])
        out = profile_utils.optimize_stock_v2(data)
        data["cut_plan"] = data.apply(
            lambda row: profile_utils.build_window_cut_plan(row, out["bars"]), axis=1
        )

        # CARRIER
        data["carrier_distances"] = data.apply(
            lambda row: profile_utils.calculate_carrier_distances(row["cut_summary"]),
            axis=1,
        )
        data["total_carrier_divisions"] = data["carrier_distances"].apply(
            lambda x: sum(len(sublist) for sublist in x)
        )
        data["total_carrier_length"] = data.apply(
            lambda row: row["total_carrier_divisions"] * row["perpendicular_length"],
            axis=1,
        )

        # ACCS
        data[["num_L", "num_inverse_L"]] = data.apply(
            lambda r: pd.Series(calculate_endcaps(r)), axis=1
        )
        data["endcap_cnt"] = data["num_L"] + data["num_inverse_L"]
        data[["nuts_bolts_cnt", "joining_pieces"]] = data.apply(
            lambda r: pd.Series(
                {
                    "nuts_bolts_cnt": r["divisions"] * r["total_carrier_divisions"],
                    "joining_pieces": (len(r["cut_summary"]) - 1) * r["divisions"],
                }
            ),
            axis=1,
        )

        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_df = generate_inventory_df(data, self.pitch, out["stock_plan"])

        results = [data, offer_df_cols, offer_df, inventory_df]

        return results
