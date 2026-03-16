
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


class RectangularCalculator:

    def __init__(self, vars):
        self.project_title = vars["project_title"]
        self.window_title = vars["window_title"]
        self.s_no = vars["s_no"]
        self.qty_windows = vars["qty_windows"]
        self.orientation = vars["orientation"]
        self.width = vars["width"]
        self.height = vars["height"]
        self.pitch = vars["pitch"]
        self.window = vars["window"]
        self.louver_size = vars["louver_size"]
        self.divisions = profile_utils.calculate_divisions(self.height, self.pitch)

    def run(self):

        endcap_cnt = profile_utils.calculate_endcaps(
            self.window, self.orientation, self.divisions
        )

        st.write("Number of total divisions: ", self.divisions)

        vars, success = profile_utils.run(
            self.width, self.height, self.divisions, self.window, self.qty_windows
        )

        total_product_length = vars["total_product_length"]
        total_carrier_length = vars["total_carrier_length"]
        total_carrier_divisions = vars["total_carrier_divisions"]
        carrier_distances_per_piece = vars["carrier_distances_per_piece"]
        used_table = vars["used_table"]
        waste_table = vars["waste_table"]

        rivet_df = pd.DataFrame()
        rivet_df["Rivet Distance"] = [carrier_distances_per_piece[0]]
        rivet_pcs = int(math.ceil(len(carrier_distances_per_piece[0]) * self.divisions))
        rivet_df["Total Rivets Required"] = rivet_pcs
        st.subheader("Rivet Calculations")
        st.write(rivet_df.T.rename_axis("Item"))

        profile_rows = []
        for item in used_table:
            section_code, section_name = RECTANGULAR_SECTION_MAPPER[self.louver_size]
            profile_rows.append(
                {
                    "Product Code": section_code,
                    "Product Name": section_name,
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

        carrier_item = [
            {
                "Product Code": RECTANGULAR_CARRIER_CODES["CARRIER_50X35"][0],
                "Product Name": RECTANGULAR_CARRIER_CODES["CARRIER_50X35"][1],
                "Length": 3650,
                "Quantity": int(math.ceil((total_carrier_length / 3650))),
                "UOM": "m",
            }
        ]

        paint_qty = round(self.divisions / 50 * 2) / 2

        endcap_code, endcap_name = RECTANGULAR_ENDCAP_MAPPER[self.louver_size]

        additional_items = [
            {
                "Product Code": endcap_code,
                "Product Name": endcap_name,
                "Quantity": endcap_cnt,
                "UOM": "pcs",
            },
            {
                "Product Name": "SELECT ENDCAP FIXTURE",
                "Quantity": endcap_cnt * 3,
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": rivet_pcs,
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PAINT"][0],
                "Product Name": COMMON_ACCESSORIES["PAINT"][1],
                "Quantity": paint_qty,
                "UOM": "l",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PAINT_BRUSH"][0],
                "Product Name": COMMON_ACCESSORIES["PAINT_BRUSH"][1],
                "Quantity": 1,
                "UOM": "pcs",
            },
        ]

        all_rows = []
        for block in [
            profile_rows,
            carrier_item,
            additional_items
        ]:
            if block:
                all_rows.extend(block)

        inventory_out = (
            pd.DataFrame(all_rows)
            .reindex(columns=INV_COLUMNS)
            .fillna("")
        )
        inventory_out["Quantity"] = inventory_out["Quantity"] * self.qty_windows

        results = pd.DataFrame(
            {
                "Project Title": [self.project_title],
                "Window": [self.window + 1],
                "Area Name": [self.window_title],
                "S. No": [self.s_no],
                "Louver Size": [self.louver_size],
                "Width (mm)": [
                    self.width if self.orientation == "Horizontal" else self.height
                ],
                "Height (mm)": [
                    self.height if self.orientation == "Horizontal" else self.width
                ],
                "Orientation": [self.orientation],
                "No. of Pieces": [self.divisions],
                "Area (ft2)": [round(((self.width * self.height) / (304.8**2)), 2)],
                "Area Qty (nos)": [self.qty_windows],
                "Pitch (mm)": [self.pitch],
                "Product Divisions": [self.divisions],
                "Total Product Length (m)": [total_product_length],
                "Total Carrier Divisions": [total_carrier_divisions],
                "Total Carrier Length (m)": [total_carrier_length / 1000],
                "Total Rivets/Screws (pcs)": [rivet_pcs],
                "End Caps (pcs)": [endcap_cnt],
                "Used Table": [used_table],
                "Waste Table": [waste_table],
            }
        )

        return results, inventory_out, success
