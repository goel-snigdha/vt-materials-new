import streamlit as st
import pandas as pd

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS


class CLouverCalculator:

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
        self.divisions = profile_utils.calculate_divisions(self.height, self.pitch)

    def run(self):

        st.write("Number of total divisions: ", self.divisions)

        total_product_length = (self.width * self.divisions) / 1000

        carrier_lengths = 0
        no_carriers_per_piece = []

        _, no_carriers_per_piece = profile_utils.carrier_calculation(
            [self.width], carrier_lengths, no_carriers_per_piece
        )

        no_carriers = self.divisions * no_carriers_per_piece
        total_carrier_length = no_carriers_per_piece[0] * self.height

        # CLouverCalculator doesn't return inventory_out,
        # so no changes needed for that
        results = pd.DataFrame(
            {
                "Width (mm)": [
                    self.width if self.orientation == "Horizontal" else self.height
                ],
                "Height (mm)": [
                    self.height if self.orientation == "Horizontal" else self.width
                ],
                "Orientation": [self.orientation],
                "Area (ft2)": [round(((self.width * self.height) / (304.8**2)), 2)],
                "Product Divisions": [self.divisions],
                "Total Product Length (m)": [total_product_length],
                "Total Carrier Length (m)": [total_carrier_length / 1000],
                "Self-Drilling 3/4 Inch Screws (pcs)": [no_carriers * 2],
            }
        )

        inventory_out = pd.DataFrame(columns=INV_COLUMNS)
        success = -1

        return results, inventory_out, success
