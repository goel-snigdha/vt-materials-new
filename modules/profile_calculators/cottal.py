import math

import pandas as pd

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, PIPE_MAPPER, COMMON_ACCESSORIES

# Cottal specific codes from your provided list
COTTAL_PRODUCTS = {
    "85 mm": {
        "PROFILE": ("CT-PR-02", "COTTAL PROFILE"),
        "START_PIECE": ("CT-SP-02", "COTTAL START PIECE"),
        "COVERING_PIECE": ("CT-CP-02", "COTTAL COVERING PIECE"),
    },
    "130 mm": {
        "PROFILE": ("CT-PR-02", "COTTAL PROFILE"),
        "START_PIECE": ("CT-SP-02", "COTTAL START PIECE"),
        "COVERING_PIECE": ("CT-CP-02", "COTTAL COVERING PIECE"),
    },
    "230 mm": {
        "PROFILE": ("CT-PR-02", "COTTAL PROFILE"),
        "START_PIECE": ("CT-SP-02", "COTTAL START PIECE"),
        "COVERING_PIECE": ("CT-CP-02", "COTTAL COVERING PIECE"),
    }
}

# Cottal specific accessories with empty codes for dropdown selection
COTTAL_ACCESSORIES = {
    "L_ANGLES": ("", "SELECT COTTAL L-ANGLES"),
    "CORNER_PIECES": ("", "SELECT COTTAL CORNER PIECES"),
}


class CottalCalculator:
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
        self.pipe_grade = vars["pipe_grade"]
        self.louver_size = vars["louver_size"]
        self.divisions = profile_utils.calculate_divisions(self.height, self.pitch)

    def run(self):
        vars, success = profile_utils.run(
            self.width, self.height, self.divisions, self.window, self.qty_windows
        )

        total_product_length = vars["total_product_length"]
        total_carrier_length = vars["total_carrier_length"]
        total_carrier_divisions = vars["total_carrier_divisions"]
        used_table = vars["used_table"]
        waste_table = vars["waste_table"]

        # Add cottal profiles, start pieces, and covering pieces for each length
        cottal_type = COTTAL_PRODUCTS[self.louver_size]
        for item in used_table:
            profile_rows = [
                {
                    "Product Code": cottal_type["PROFILE"][0],
                    "Product Name": cottal_type["PROFILE"][1],
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                    "Colour": "COLOURED",
                },
                {
                    "Product Code": cottal_type["START_PIECE"][0],
                    "Product Name": cottal_type["START_PIECE"][1],
                    "Length": item,
                    "Quantity": 1,
                    "UOM": "m",
                    "Colour": "COLOURED",
                },
                {
                    "Product Code": cottal_type["COVERING_PIECE"][0],
                    "Product Name": cottal_type["COVERING_PIECE"][1],
                    "Length": item,
                    "Quantity": 2,
                    "UOM": "m",
                    "Colour": "COLOURED",
                },
            ]

        paint_qty = round(self.divisions / 50 * 2) / 2
        brush_qty = math.ceil(paint_qty)

        # Get pipe code and name from PIPE_MAPPER
        pipe_code, pipe_name = PIPE_MAPPER[self.pipe_grade]

        additional_items = [
            {
                "Product Code": pipe_code,
                "Product Name": pipe_name,
                "Length": 3650,
                "Quantity": int(math.ceil(total_carrier_length / 3650)),
                "UOM": "m",
                "Colour": "MILL",
            },
            {
                "Product Code": COTTAL_ACCESSORIES["L_ANGLES"][0],
                "Product Name": COTTAL_ACCESSORIES["L_ANGLES"][1],
                "Quantity": 0,
                "UOM": "m",
                "Colour": "COLOURED",
            },
            {
                "Product Code": COTTAL_ACCESSORIES["CORNER_PIECES"][0],
                "Product Name": COTTAL_ACCESSORIES["CORNER_PIECES"][1],
                "Quantity": 0,
                "UOM": "pcs",
                "Colour": "COLOURED",
            },
            {
                "Product Code": COMMON_ACCESSORIES["EPDM_GASKET"][0],
                "Product Name": COMMON_ACCESSORIES["EPDM_GASKET"][1],
                "Quantity": int(math.ceil(total_product_length)),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": int(math.ceil(self.divisions * total_carrier_divisions)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                "Quantity": int(math.ceil(total_carrier_length / 500)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                "Quantity": int(math.ceil(total_carrier_length / 500)),
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
                "Quantity": brush_qty,
                "UOM": "pcs",
            },
        ]

        # Append the additional items
        all_rows = []
        for block in [
            profile_rows,
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
                "Pitch (mm)": [self.pitch],
                "Orientation": [self.orientation],
                "Area (ft2)": [round(((self.width * self.height) / (304.8**2)), 2)],
                "Area Qty (nos)": [self.qty_windows],
                "No. of Pieces": [self.divisions],
                "Total Product Length (m)": [total_product_length],
                "Louver Size": [self.louver_size],
                "Aluminum Pipe Grade": [self.pipe_grade],
                "Aluminum Pipe Length (m)": [total_carrier_length / 1000],
                "Total Pipe Divisions": [total_carrier_divisions],
                "Total 3650mm Pipe Pieces": [round(total_carrier_length / 3650, 1)],
                "Start Piece (29-3002-00) Piece Length (mm)": [self.width],
                "Start Piece (29-3002-00) Qty (Pcs)": [1],
                "End Pieces (29-4001-00) Qty (pcs)": [2],
                "End Pieces (29-4001-00) Piece Length (mm)": [self.width],
                "EPDM Gasket Length (m)": [total_product_length],
                "Self-Drilling 3/4 Inch Screws (pcs)": [
                    self.divisions * total_carrier_divisions
                ],
                "Full Threaded 75mm Screws (pcs)": [total_carrier_length / 500],
                "PVC Gitty (pcs)": [total_carrier_length / 500],
                "Used Table": [used_table],
                "Waste Table": [waste_table],
            }
        )
        return results, inventory_out, success
