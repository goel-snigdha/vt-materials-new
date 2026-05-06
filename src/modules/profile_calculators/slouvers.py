import math
import streamlit as st
import pandas as pd
from matplotlib.lines import Line2D

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, COMMON_ACCESSORIES

S_LOUVER_PRODUCTS = {
    "PROFILE": ("SL-PR-02", "S-LOUVER PROFILE"),
    "CARRIER": ("SL-CA-02", "S-LOUVER CARRIER"),
    "FIXTURE": ("", "SELECT S-LOUVER FIXTURE"),
}


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
    }
    offer_df = data[offer_df_cols.keys()].copy()

    return offer_df_cols, offer_df


def generate_inventory_df(self, data, stock_plan):

    all_inventory_rows = []

    sorted_plan = sorted(stock_plan, key=lambda x: x["length"], reverse=True)
    profile_rows = [
        {
            "Product Code": S_LOUVER_PRODUCTS["PROFILE"][0],
            "Product Name": S_LOUVER_PRODUCTS["PROFILE"][1],
            "Length": item["length"],
            "Quantity": item["qty"],
            "UOM": "m",
            "item_order": i,
        }
        for i, item in enumerate(sorted_plan)
    ]
    l_angle_rows = [
        {
            "Product Code": COMMON_ACCESSORIES["L_ANGLE_19X19"][0],
            "Product Name": COMMON_ACCESSORIES["L_ANGLE_19X19"][1],
            "Length": 3050,
            "Quantity": int(math.ceil(item["length"] / 3050)),
            "UOM": "m",
            "item_order": len(profile_rows) + i,
        }
        for i, item in enumerate(sorted_plan)
    ]
    all_inventory_rows.extend(profile_rows)
    all_inventory_rows.extend(l_angle_rows)

    sequence = len(profile_rows) + len(l_angle_rows)

    for _, row in data.iterrows():

        additional_items = [
            {
                "Product Code": S_LOUVER_PRODUCTS["CARRIER"][0],
                "Product Name": S_LOUVER_PRODUCTS["CARRIER"][1],
                "Length": 3050,
                "Quantity": int(math.ceil((row["num_carrier_pieces"] * 55) / 3050)),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": row["self_drill_screws"],
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["RIVET_6MM"][0],
                "Product Name": COMMON_ACCESSORIES["RIVET_6MM"][1],
                "Quantity": row["rivets"],
                "UOM": "pcs",
            },
        ]

        row_items = additional_items

        for item in row_items:
            item["Quantity"] = int(item["Quantity"] * row["qty_areas"])
            item["item_order"] = sequence
            sequence += 1

        all_inventory_rows.extend(row_items)

    all_inventory_rows += [
        {"Product Code": COMMON_ACCESSORIES["PAINT"][0], "Product Name": COMMON_ACCESSORIES["PAINT"][1], "Quantity": 1, "UOM": "l", "item_order": sequence},
        {"Product Code": COMMON_ACCESSORIES["PAINT_BRUSH"][0], "Product Name": COMMON_ACCESSORIES["PAINT_BRUSH"][1], "Quantity": 1, "UOM": "pcs", "item_order": sequence + 1},
    ]

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
        .reindex(columns=INV_COLUMNS)
        .fillna("")
    )

    return inv_data


class SLouverCalculator:

    def __init__(self, vars):
        self.project_title = vars["project_title"]
        self.louver_size = vars["louver_size"]
        self.pitch = vars["pitch"]
        self.areas = vars["areas"]

    @staticmethod
    def get_data_input(**kwargs):

        empty_df = pd.DataFrame(
            {
                "s_no": pd.Series(dtype="str"),
                "area_name": pd.Series(dtype="str"),
                "height": pd.Series(dtype="int"),
                "width": pd.Series(dtype="int"),
                "orientation": pd.Series(dtype="str"),
                "qty_areas": pd.Series(dtype="int"),
                "cut_summary": pd.Series(dtype="str"),
            }
        )

        required_cols = [
            "width",
            "height",
            "orientation",
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
                "orientation": st.column_config.SelectboxColumn(
                    "Orientation", options=["Horizontal", "Vertical"], required=True
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

    @staticmethod
    def generate_image(row, common_vars):

        divisions = row["divisions"]
        cut_summary = row["cut_summary"]

        if len(cut_summary) > 1:
            cut_summary_str = "+".join(str(c) for c in cut_summary)
        else:
            cut_summary_str = f"Single Piece {cut_summary[0]} mm"

        orientation = str(row.get("orientation", "Vertical")).strip()
        L_ANGLE_COLOR = "#9B30FF"

        def draw_l_angle(ax, _row, W, H, _total):
            lw = 4
            if orientation == "Horizontal":
                ax.plot([0, W], [H, H], color=L_ANGLE_COLOR, linewidth=lw, zorder=6)
            else:
                ax.plot([0, 0], [0, H], color=L_ANGLE_COLOR, linewidth=lw, zorder=6)

        info_lines = [
            (f"{divisions} Divisions", "#333", 12, True),
            ("Breakdown", "#333", 12, False),
            (f"{cut_summary_str}", "#333", 12, True),
        ]

        config = {
            "show_carriers": True,
            "show_endcaps": False,
            "show_joints": True,
            "bar_color": "#888",
            "info_lines": info_lines,
            "extras": draw_l_angle,
            "legend_extras": [
                Line2D([0], [0], color=L_ANGLE_COLOR, linewidth=3, label="L-Angle"),
            ],
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
        data["num_carrier_pieces"] = data.apply(
            lambda row: math.ceil(row["total_carrier_length"] / 133.7), axis=1
        )

        # ACCS
        data["self_drill_screws"] = data.apply(
            lambda row: row["num_carrier_pieces"] * 2,
            axis=1,
        )

        data["rivets"] = data.apply(
            lambda row: row["num_carrier_pieces"] * 3,
            axis=1,
        )

        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_df = generate_inventory_df(self, data, out["stock_plan"])

        results = [data, offer_df_cols, offer_df, inventory_df]

        return results
