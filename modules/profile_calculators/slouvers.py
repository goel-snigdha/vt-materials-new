import math

import pandas as pd

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, COMMON_ACCESSORIES

S_LOUVER_PRODUCTS = {
    "PROFILE": ("SL-PR-02", "S-LOUVER PROFILE"),
    "CARRIER": ("SL-CA-02", "S-LOUVER CARRIER"),
    "FIXTURE": ("", "SELECT S-LOUVER FIXTURE"),
}


class SLouverCalculator:

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
        vars, success = profile_utils.run(
            self.width, self.height, self.divisions, self.window, self.qty_windows
        )

        total_product_length = vars["total_product_length"]
        total_carrier_length = vars["total_carrier_length"]
        total_carrier_divisions = vars["total_carrier_divisions"]
        num_carrier_pieces = math.ceil(total_carrier_length / 133.7)
        self_drill_screws = int(
            math.ceil(
                (self.divisions * total_carrier_divisions) + (num_carrier_pieces * 2)
            )
        )
        used_table = vars["used_table"]
        waste_table = vars["waste_table"]

        profile_rows = []
        l_angle_rows = []
        for item in used_table:
            profile_rows.append(
                {
                    "Product Code": S_LOUVER_PRODUCTS["PROFILE"][0],
                    "Product Name": S_LOUVER_PRODUCTS["PROFILE"][1],
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

            l_angle_rows.append(
                {
                    "Product Code": COMMON_ACCESSORIES["L_ANGLE_19X19"][0],
                    "Product Name": COMMON_ACCESSORIES["L_ANGLE_19X19"][1],
                    "Length": 3050,
                    "Quantity": int(math.ceil(item / 3050)),
                    "UOM": "m",
                }
            )

        additional_items = [
            {
                "Product Code": S_LOUVER_PRODUCTS["CARRIER"][0],
                "Product Name": S_LOUVER_PRODUCTS["CARRIER"][1],
                "Length": 3050,
                "Quantity": int(math.ceil((num_carrier_pieces * 55) / 3050)),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": self_drill_screws,
                "UOM": "pcs",
            },
        ]

        all_rows = []
        for block in [
            profile_rows,
            l_angle_rows,
            additional_items,
        ]:
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
                "Width (mm)": [
                    self.width if self.orientation == "Horizontal" else self.height
                ],
                "Height (mm)": [
                    self.height if self.orientation == "Horizontal" else self.width
                ],
                "Orientation": [self.orientation],
                "Area (ft2)": [round(((self.width * self.height) / (304.8**2)), 2)],
                "Area Qty (nos)": [self.qty_windows],
                "No. of Pieces": [self.divisions],
                "Pitch (mm)": [self.pitch],
                "Louver Size": [self.louver_size],
                "Product Divisions": [self.divisions],
                "Total Product Length (m)": [total_product_length],
                "Total Carrier Divisions": [total_carrier_divisions],
                "Total Carrier Length (m)": [total_carrier_length / 1000],
                "133.7 mm Carrier Pieces (nos)": [num_carrier_pieces],
                "Self-Drilling 3/4 Inch Screws (pcs)": [self_drill_screws],
                "Rivet": [self.divisions * total_carrier_divisions],
                "Used Table": [used_table],
                "Waste Table": [waste_table],
            }
        )

        return results, inventory_out, success
