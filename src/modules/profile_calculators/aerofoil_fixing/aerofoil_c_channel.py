import math
import pandas as pd
from matplotlib.lines import Line2D

from modules.excel_utils import COMMON_ACCESSORIES, INV_COLUMNS
from modules.profile_calculators.aerofoil_fixing.aerofoil_common import (
    AEROFOIL_SECTION_MAPPER,
    C_PLATE_CODES,
)


class AerofoilCChannel:

    def __init__(self, common_vars):
        self.af_type = common_vars["af_type"]

    @staticmethod
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

        C_CHANNEL_COLOR = "#9B30FF"
        orientation = row.get("orientation", "Vertical")

        def draw_c_channels(ax, _row, W, H, _total):
            lw = 4
            if orientation == "Vertical":
                ax.plot([0, W], [0, 0], color=C_CHANNEL_COLOR, linewidth=lw, zorder=5)
                ax.plot([0, W], [H, H], color=C_CHANNEL_COLOR, linewidth=lw, zorder=5)
            else:
                ax.plot([0, 0], [0, H], color=C_CHANNEL_COLOR, linewidth=lw, zorder=5)
                ax.plot([W, W], [0, H], color=C_CHANNEL_COLOR, linewidth=lw, zorder=5)

        return {
            "show_carriers": False,
            "show_endcaps":  False,
            "show_joints":   False,
            "bar_color":     "#aaa",
            "info_lines":    info_lines,
            "extras":        draw_c_channels,
            "legend_extras": [
                Line2D([0], [0], color=C_CHANNEL_COLOR, linewidth=3, label="C-Channel"),
            ],
        }

    def run(self, data, stock_plan):

        data = data.copy()
        data["plate_length"] = (data["perpendicular_length"] * 2)
        data["plate_length_m"] = (data["perpendicular_length"] * 2) / 1000

        profile_code, profile_name = AEROFOIL_SECTION_MAPPER[self.af_type]

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
            plate_width = int(row["plate_width"])
            plate_length = int(row["plate_length"])

            c_plate_code, c_plate_name = C_PLATE_CODES[plate_width]

            items = [
                {
                    "Product Code": c_plate_code,
                    "Product Name": c_plate_name,
                    "Length": 3650,
                    "Quantity": int(math.ceil(plate_length / 3650) * qty_areas),
                    "UOM": "m",
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
                    "Quantity": int(math.ceil(plate_length / 300) * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 2,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                    "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                    "Quantity": int(math.ceil(plate_length / 300) * qty_areas),
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
            "s_no": {"type": "desc", "hide_if_zero": False},
            "area_name": {"type": "desc", "hide_if_zero": False},
            "orientation": {"type": "desc", "hide_if_zero": False},
            "height": {"type": "desc", "hide_if_zero": False},
            "width": {"type": "desc", "hide_if_zero": False},
            "qty_areas": {"type": "desc", "hide_if_zero": False},
            "area_sqft": {"type": "formula", "hide_if_zero": False},
            "divisions": {"type": "formula", "hide_if_zero": False},
            "total_product_length": {"type": "formula", "hide_if_zero": False},
            "plate_length_m": {"type": "formula", "hide_if_zero": False},
        }
        offer_df = data[offer_df_cols.keys()].copy()

        return data, offer_df_cols, offer_df, inv_data
