import math
import pandas as pd

from modules.excel_utils import COMMON_ACCESSORIES, INV_COLUMNS
from modules.profile_calculators.aerofoil_fixing.aerofoil_common import (
    AEROFOIL_SECTION_MAPPER,
    AEROFOIL_FRINGE_ENDCAPS,
)


class AerofoilFringe:

    def __init__(self, common_vars):
        self.af_type = common_vars["af_type"]

    def generate_image(row, common_vars):

        pitch = common_vars.get("pitch", "")
        divisions = row["divisions"]
        cut_summary = row["cut_summary"]

        if len(cut_summary) > 1:
            cut_summary_str = "+".join(str(c) for c in cut_summary)
        else:
            cut_summary_str = f"Single Piece {cut_summary[0]} mm"

        info_lines = [
            (f"{divisions} Divisions", "#333", 12, True),
            (f"@ {pitch} mm Pitch",    "#333", 12, True),
            ("",                       "#333", 12, False),
            ("Breakdown",              "#333", 12, False),
            (f"{cut_summary_str}",     "#333", 12, True),
        ]

        orientation = row.get("orientation", "Vertical")
        endcap_sides = (
            {"top": True,  "bottom": True,  "left": False, "right": False}
            if orientation == "Vertical"
            else {"top": False, "bottom": False, "left": True, "right": True}
        )

        return {
            "show_carriers": False,
            "show_endcaps":  True,
            "endcap_sides":  endcap_sides,
            "show_joints":   False,
            "bar_color":     "#aaa",
            "info_lines":    info_lines,
        }

    def run(self, data, stock_plan):

        data = data.copy()
        data["fringe_endcap_cnt"] = data["divisions"] * 2

        profile_code, profile_name = AEROFOIL_SECTION_MAPPER[self.af_type]
        fringe_code, fringe_name = AEROFOIL_FRINGE_ENDCAPS[self.af_type]

        all_rows = [
            {
                "Product Code": profile_code,
                "Product Name": profile_name,
                "Length": item["length"],
                "Quantity": item["qty"],
                "UOM": "m",
                "item_order": i,
            }
            for i, item in enumerate(sorted(stock_plan, key=lambda x: x["length"], reverse=True))
        ]
        sequence = len(all_rows)

        for _, row in data.iterrows():
            divisions = int(row["divisions"])
            qty_areas = int(row["qty_areas"])

            items = [
                {
                    "Product Code": fringe_code,
                    "Product Name": fringe_name,
                    "Quantity": int(row['fringe_endcap_cnt'] * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                    "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                    "Quantity": int(divisions * 4 * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 1,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                    "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                    "Quantity": int(math.ceil(divisions * 4) * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 2,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                    "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                    "Quantity": int(math.ceil(divisions * 4) * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 3,
                },
            ]
            sequence += len(items)
            all_rows.extend(items)

        inv_data = (
            pd.DataFrame(all_rows)
            .reindex(columns=INV_COLUMNS + ["item_order"])
            .fillna("")
        )
        inv_data = (
            inv_data.groupby(
                ["Product Code", "Product Name", "Length", "UOM",
                 "Colour", "Finish", "CNC Hole Distance", "Remarks"],
                as_index=False, sort=False,
            )
            .agg({"Quantity": "sum", "item_order": "min"})
            .sort_values("item_order")
            .drop(columns=["item_order"])
            .reindex(columns=INV_COLUMNS)
            .fillna("")
        )

        offer_df_cols = {
            "s_no": {"type": "desc",    "hide_if_zero": False},
            "area_name": {"type": "desc",    "hide_if_zero": False},
            "orientation": {"type": "desc",    "hide_if_zero": False},
            "height": {"type": "desc",    "hide_if_zero": False},
            "width": {"type": "desc",    "hide_if_zero": False},
            "qty_areas": {"type": "desc",    "hide_if_zero": False},
            "area_sqft": {"type": "formula", "hide_if_zero": False},
            "divisions": {"type": "formula", "hide_if_zero": False},
            "total_product_length": {"type": "formula", "hide_if_zero": False},
            "fringe_endcap_cnt": {"type": "formula", "hide_if_zero": False},
        }
        offer_df = data[offer_df_cols.keys()].copy()

        return data, offer_df_cols, offer_df, inv_data
