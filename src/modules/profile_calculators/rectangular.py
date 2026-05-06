import streamlit as st
import pandas as pd
import math

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, COMMON_ACCESSORIES

RECTANGULAR_SECTION_MAPPER = {
    "30x60": ("RS-03060-PR-02", "RECTANGULAR SECTION 30x60"),
    "50x75": ("RS-05075-PR-02", "RECTANGULAR SECTION 50x75"),
    "50x100": ("RS-50100-PR-02", "RECTANGULAR SECTION 50x100"),
    "50x125": ("RS-50125-PR-02", "RECTANGULAR SECTION 50x125"),
}

RECTANGULAR_ENDCAP_MAPPER = {
    "30x60": ("RS-03060-EC-02", "RECTANGULAR ENDCAP 30x60"),
    "50x75": ("RS-05075-EC-02", "RECTANGULAR ENDCAP 50x75"),
    "50x100": ("RS-50100-EC-02", "RECTANGULAR ENDCAP 50x100"),
    "50x125": ("RS-50125-EC-02", "RECTANGULAR ENDCAP 50x125"),
}

RECTANGULAR_CARRIER_CODES = {
    "CARRIER_50X35": ("RS-05030-C1-02", "RECTANGULAR SECTION CARRIER 50x35"),
    "CARRIER_48X08": ("RS-04808-C2-02", "RECTANGULAR SECTION CARRIER 48x08"),
    "CARRIER_28X10": ("RS-02810-C3-02", "RECTANGULAR SECTION CARRIER 28x10"),
}

ENDCAP_MULTIPLIER = {
    "No Endcaps": 0,
    "Both sides": 2,
    "Top Side": 1,
    "Bottom Side": 1,
    "Right Side": 1,
    "Left Side": 1,
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
        "endcap_cnt": {
            "type": "formula",
            "hide_if_zero": True,
        },
    }
    offer_df = data[offer_df_cols.keys()].copy()

    return offer_df_cols, offer_df


def generate_inventory_df(self, data, stock_plan):

    all_inventory_rows = []

    section_code, section_name = RECTANGULAR_SECTION_MAPPER[self.louver_size]
    profile_rows = [
        {
            "Product Code": section_code,
            "Product Name": section_name,
            "Length": item["length"],
            "Quantity": item["qty"],
            "UOM": "m",
            "item_order": i,
        }
        for i, item in enumerate(sorted(stock_plan, key=lambda x: x["length"], reverse=True))
    ]
    all_inventory_rows.extend(profile_rows)

    sequence = len(profile_rows)

    for _, row in data.iterrows():

        carrier_item = [
            {
                "Product Code": RECTANGULAR_CARRIER_CODES["CARRIER_50X35"][0],
                "Product Name": RECTANGULAR_CARRIER_CODES["CARRIER_50X35"][1],
                "Length": 3650,
                "Quantity": int(math.ceil(row["total_carrier_length"] / 3650)),
                "UOM": "m",
            }
        ]

        endcap_code, endcap_name = RECTANGULAR_ENDCAP_MAPPER[self.louver_size]

        additional_items = [
            {
                "Product Code": endcap_code,
                "Product Name": endcap_name,
                "Quantity": row["endcap_cnt"],
                "UOM": "pcs",
            },
            {
                "Product Name": "SELECT ENDCAP FIXTURE",
                "Quantity": row["endcap_cnt"] * 3,
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": row["rivet_cnt"],
                "UOM": "pcs",
            },
        ]

        row_items = carrier_item + additional_items

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


class RectangularCalculator:

    def __init__(self, vars):
        self.project_title = vars["project_title"]
        self.louver_size = vars["louver_size"]
        self.pitch = vars["pitch"]
        self.areas = vars["areas"]

    def validate_input(row, idx):

        VALID_ENDCAPS_BOTH = {"No Endcaps", "Both sides"}
        VERTICAL_ENDCAPS = {"Top Side", "Bottom Side"}
        HORIZONTAL_ENDCAPS = {"Right Side", "Left Side"}

        orientation = str(row.get("orientation", "")).strip()
        endcaps = str(row.get("endcaps", "")).strip()

        valid = VALID_ENDCAPS_BOTH | (
            HORIZONTAL_ENDCAPS if orientation == "Horizontal" else VERTICAL_ENDCAPS
        )

        if endcaps not in valid:
            st.warning(
                f"Row {idx + 1}: Endcap '{endcaps}' is not valid for {orientation} orientation. "
                f"Valid options: {', '.join(sorted(valid))}"
            )
            return False

        return True

    def get_data_input(**kwargs):

        empty_df = pd.DataFrame(
            {
                "s_no": pd.Series(dtype="str"),
                "area_name": pd.Series(dtype="str"),
                "height": pd.Series(dtype="int"),
                "width": pd.Series(dtype="int"),
                "orientation": pd.Series(dtype="str"),
                "endcaps": pd.Series(dtype="str"),
                "qty_areas": pd.Series(dtype="int"),
                "cut_summary": pd.Series(dtype="str"),
            }
        )

        required_cols = [
            "width",
            "height",
            "orientation",
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
                "orientation": st.column_config.SelectboxColumn(
                    "Orientation", options=["Horizontal", "Vertical"], required=True
                ),
                "endcaps": st.column_config.SelectboxColumn(
                    "Endcaps", options=ENDCAP_MULTIPLIER.keys(), required=True
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

    def generate_image(row, common_vars):

        pitch = common_vars["pitch"]
        divisions = row["divisions"]
        cut_summary = row["cut_summary"]
        endcap_choice = str(row["endcaps"]).strip()
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
        ]

        # Compute endcap sides
        draw_top = draw_bottom = draw_left = draw_right = False

        if orient == "Vertical":
            draw_top = endcap_choice in ("Both sides", "Top Side")
            draw_bottom = endcap_choice in ("Both sides", "Bottom Side")
        else:
            draw_right = endcap_choice in ("Both sides", "Right Side")
            draw_left = endcap_choice in ("Both sides", "Left Side")
        endcap_sides = {
            "top": draw_top,
            "bottom": draw_bottom,
            "left": draw_left,
            "right": draw_right,
        }

        config = {
            "show_carriers": False,
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
        data["total_carrier_divisions"] = data["divisions"]
        data["total_carrier_length"] = data.apply(
            lambda row: row["total_carrier_divisions"] * row["single_division_length"],
            axis=1,
        )

        # ACCS
        data["endcap_cnt"] = data.apply(
            lambda row: row["divisions"] * ENDCAP_MULTIPLIER[row["endcaps"]], axis=1
        )
        data["rivet_distance"] = data.apply(
            lambda row: profile_utils.calculate_carrier_distances(row["cut_summary"]),
            axis=1,
        )
        data["rivets_per_division"] = data["rivet_distance"].apply(
            lambda x: sum(len(sublist) for sublist in x)
        )
        data["rivet_cnt"] = data.apply(
            lambda row: row["rivets_per_division"] * row["divisions"], axis=1
        )
        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_df = generate_inventory_df(self, data, out["stock_plan"])

        results = [data, offer_df_cols, offer_df, inventory_df]

        return results
