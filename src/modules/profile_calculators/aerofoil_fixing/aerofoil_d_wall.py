import math
import pandas as pd

from modules.excel_utils import COMMON_ACCESSORIES, INV_COLUMNS
from modules.profile_calculators.aerofoil_fixing.aerofoil_common import (
    AEROFOIL_SECTION_MAPPER,
    AEROFOIL_ENDCAPS,
    AEROFOIL_ACCESSORIES,
)


class AerofoilDWall:

    def __init__(self, common_vars):
        self.af_type = common_vars["af_type"]

    def generate_image(row, common_vars):
        from matplotlib.patches import Rectangle as MplRect

        pitch = common_vars.get("pitch", "")
        divisions = row["divisions"]
        cut_summary = row["cut_summary"]
        orientation = row.get("orientation", "Vertical")
        division_length = int(row.get("single_division_length", 1) or 1)

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

        BRACKET_COLOR = "#1ABC9C"
        BRACKET_GAP_MM = 150
        BOX_SIZE = 8  # visual size of the bracket square in canvas units

        endcap_choice = str(row.get("endcaps", "") or "").strip()
        endcap_sides = {
            "top":    endcap_choice in ("Both Sides", "Top"),
            "bottom": endcap_choice in ("Both Sides", "Bottom"),
            "left":   endcap_choice in ("Both Sides", "Left"),
            "right":  endcap_choice in ("Both Sides", "Right"),
        }

        def draw_brackets(ax, _row, W, H, _total):
            N_BARS = 20
            gap_px = (BRACKET_GAP_MM / division_length)

            if orientation == "Vertical":
                gap_top = H * (1 - gap_px)       # 150mm from top
                gap_bot = H * gap_px              # 150mm from bottom
                for i in range(N_BARS):
                    x = (i + 0.5) * W / N_BARS
                    # bracket at top end
                    ax.add_patch(MplRect(
                        (x - BOX_SIZE / 2, gap_top - BOX_SIZE / 2),
                        BOX_SIZE, BOX_SIZE,
                        linewidth=1, edgecolor=BRACKET_COLOR, facecolor=BRACKET_COLOR, zorder=6
                    ))
                    # bracket at bottom end
                    ax.add_patch(MplRect(
                        (x - BOX_SIZE / 2, gap_bot - BOX_SIZE / 2),
                        BOX_SIZE, BOX_SIZE,
                        linewidth=1, edgecolor=BRACKET_COLOR, facecolor=BRACKET_COLOR, zorder=6
                    ))
            else:
                gap_left = W * gap_px
                gap_right = W * (1 - gap_px)
                for i in range(N_BARS):
                    y = (i + 0.5) * H / N_BARS
                    ax.add_patch(MplRect(
                        (gap_left - BOX_SIZE / 2, y - BOX_SIZE / 2),
                        BOX_SIZE, BOX_SIZE,
                        linewidth=1, edgecolor=BRACKET_COLOR, facecolor=BRACKET_COLOR, zorder=6
                    ))
                    ax.add_patch(MplRect(
                        (gap_right - BOX_SIZE / 2, y - BOX_SIZE / 2),
                        BOX_SIZE, BOX_SIZE,
                        linewidth=1, edgecolor=BRACKET_COLOR, facecolor=BRACKET_COLOR, zorder=6
                    ))

        return {
            "show_carriers": False,
            "show_endcaps":  any(endcap_sides.values()),
            "endcap_sides":  endcap_sides,
            "show_joints":   False,
            "bar_color":     "#aaa",
            "info_lines":    info_lines,
            "extras":        draw_brackets,
            "legend_extras": [
                MplRect((0, 0), 1, 1, facecolor=BRACKET_COLOR, edgecolor="#0e8c6e", label="D+V Bracket"),
            ],
        }

    def run(self, data, stock_plan):

        data = data.copy()

        def calc_endcap_cnt(row):
            choice = str(row.get("endcaps", "") or "").strip()
            divisions = int(row["divisions"])
            if choice == "Both Sides":
                return divisions * 2
            elif choice in ("Top", "Bottom", "Left", "Right"):
                return divisions
            return 0

        data["endcap_cnt"] = data.apply(calc_endcap_cnt, axis=1)

        profile_code, profile_name = AEROFOIL_SECTION_MAPPER[self.af_type]
        endcap_code, endcap_name = AEROFOIL_ENDCAPS[self.af_type]

        v_bracket_key = (
            "AEROFOIL_AF100_V_BRACKET" if self.af_type == "AF100"
            else "AEROFOIL_AF150_V_BRACKET"
        )
        v_bracket_code, v_bracket_name = AEROFOIL_ACCESSORIES[v_bracket_key]

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
            total_carrier_divs = int(row["total_carrier_divisions"])

            brackets = divisions * total_carrier_divs
            endcap_cnt = int(row["endcap_cnt"])

            items = [
                {
                    "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_D_BRACKET"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_D_BRACKET"][1],
                    "Length": 3050,
                    "Quantity": int(math.ceil(brackets * 55 / 3050) * qty_areas),
                    "UOM": "m",
                    "item_order": sequence,
                },
                {
                    "Product Code": v_bracket_code,
                    "Product Name": v_bracket_name,
                    "Length": 3050,
                    "Quantity": int(math.ceil(brackets * 55 / 3050) * qty_areas),
                    "UOM": "m",
                    "item_order": sequence + 1,
                },
                {
                    "Product Code": endcap_code,
                    "Product Name": endcap_name,
                    "Quantity": int(endcap_cnt * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 2,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                    "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                    "Quantity": int(brackets * 5 * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 3,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["TRUSS_HEAD_16MM"][0],
                    "Product Name": COMMON_ACCESSORIES["TRUSS_HEAD_16MM"][1],
                    "Quantity": int(brackets * 2 * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 4,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                    "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                    "Quantity": int((
                        endcap_cnt * 2 if self.af_type in ["AF60", "AF100"]
                        else endcap_cnt * 4
                    ) * qty_areas),
                    "UOM": "pcs",
                    "item_order": sequence + 5,
                },
            ]
            sequence += len(items)
            all_rows.extend(items)

        all_rows += [
            {"Product Code": COMMON_ACCESSORIES["PAINT"][0], "Product Name": COMMON_ACCESSORIES["PAINT"][1], "Quantity": 1, "UOM": "l", "item_order": sequence},
            {"Product Code": COMMON_ACCESSORIES["PAINT_BRUSH"][0], "Product Name": COMMON_ACCESSORIES["PAINT_BRUSH"][1], "Quantity": 1, "UOM": "pcs", "item_order": sequence + 1},
        ]

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
            "endcap_cnt": {"type": "formula", "hide_if_zero": True},
        }
        offer_df = data[offer_df_cols.keys()].copy()

        return data, offer_df_cols, offer_df, inv_data
