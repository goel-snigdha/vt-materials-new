import math
import pandas as pd
from modules.excel_utils import INV_COLUMNS, PIPE_MAPPER, COMMON_ACCESSORIES

# standard sheet dimesions are 8ft x 4ft
STANDARD_BEAM_C_LENGTH = 2400
STANDARD_BEAM_C_WIDTH = 1200
STANDARD_PROFILE_LENGTH = 2400


class BeamCCalculator:

    def __init__(self, vars):
        self.project_title = vars["project_title"]
        self.window_title = vars["window_title"]
        self.s_no = vars["s_no"]
        self.qty_windows = vars["qty_windows"]
        self.width = vars["width"]
        self.length = vars["length"]
        self.qty_windows = vars["qty_windows"]
        self.window = vars["window"]
        self.pipe_grade = vars["pipe_grade"]

    def run(self):

        buffer_width = self.width + 50
        total_length = self.length * self.qty_windows
        divisions_per_sheet = STANDARD_BEAM_C_WIDTH // buffer_width
        length_per_sheet = STANDARD_BEAM_C_LENGTH * divisions_per_sheet
        # waste_per_sheet = STANDARD_BEAM_C_WIDTH % divisions
        no_sheets = math.ceil((total_length) / length_per_sheet)
        # total_waste = waste_per_sheet * no_sheets
        profile_qty = math.ceil((total_length * 2) / STANDARD_PROFILE_LENGTH)
        used_table = {STANDARD_BEAM_C_LENGTH: divisions_per_sheet}
        sheet_code, sheet_desc = "BC-SHT-CS", "BEAM C CHANNEL SHEET CUT TO SIZE"

        # Append rows correctly
        first_row = [
            {
                "Product Code": "BC-CHN-01",
                "Product Name": "BEAM C CHANNEL",
                "Length": 2500,
                "Quantity": profile_qty,
                "UOM": "m",
            }
        ]

        additional_items = [
            {
                "Product Code": sheet_code,
                "Product Name": sheet_desc,
                "Quantity": divisions_per_sheet * no_sheets,
                "UOM": "pcs",
                "Remarks": f"{buffer_width} x {STANDARD_BEAM_C_LENGTH} mm each",
            },
            {
                "Product Code": COMMON_ACCESSORIES["EPDM_GASKET"][0],
                "Product Name": COMMON_ACCESSORIES["EPDM_GASKET"][1],
                "Quantity": int(math.ceil((total_length * 2) / 1000)),
                "UOM": "m",
            },
            {
                "Product Code": PIPE_MAPPER[self.pipe_grade][0],
                "Product Name": PIPE_MAPPER[self.pipe_grade][1],
                "Length": 3650,
                "Quantity": int(math.ceil((total_length * 2) / 3650)),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["RIVET_6MM"][0],
                "Product Name": COMMON_ACCESSORIES["RIVET_6MM"][1],
                "Quantity": int(math.ceil(total_length / 300)),
                "UOM": "pcs",
            },
            {
                "Product Code": "AC-GN-SB",
                "Product Name": "SILICON BOTTLE BLACK",
                "Quantity": int(math.ceil(total_length / 3000)),
                "UOM": "pcs",
            },
        ]

        if self.pipe_grade == "50x25":
            screw_rows = [
                {
                    "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                    "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                    "Quantity": int(math.ceil(total_length / 300)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                    "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                    "Quantity": int(math.ceil(total_length / 300)),
                    "UOM": "pcs",
                },
            ]
        elif self.pipe_grade == "25x12":
            screw_rows = [
                {
                    "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_25MM"][0],
                    "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_25MM"][1],
                    "Quantity": int(math.ceil(total_length / 300)),
                    "UOM": "pcs",
                }
            ]

        l_angle = [{"Product Name": "SELECT L-ANGLES", "UOM": "m"}]

        all_rows = []
        for block in [
            first_row,
            additional_items,
            screw_rows,
            l_angle,
        ]:
            all_rows.extend(block)

        inventory_out = pd.DataFrame(all_rows).reindex(columns=INV_COLUMNS).fillna("")

        results = pd.DataFrame(
            {
                "Project Title": [self.project_title],
                "Window": [self.window + 1],
                "Area Name": [self.window_title],
                "S. No": [self.s_no],
                "Width (mm)": [self.width],
                "Length (mm)": [self.length],
                "Area Qty (nos)": [self.qty_windows],
                "Total Product Length (m)": [(self.length) / 1000],
                "EPDM Rubber Length (m)": [(self.length * 2) / 1000],
                "Profile Length (m)": [(self.length * 2) / 1000],
                "Aluminium Pipe (m)": [(self.length * 2) / 1000],
                "Full Threaded 75mm Screws (pcs)": [math.ceil(self.length / 400)],
                "PVC Gitty (pcs)": [math.ceil(self.length / 400)],
                "Total Rivets (pcs)": [math.ceil(self.length / 400)],
                "Used Table": [used_table],
                "Waste Table": [{}],
            }
        )

        return results, inventory_out, True
