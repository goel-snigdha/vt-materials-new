import math
import pandas as pd

from modules.excel_utils import COMMON_ACCESSORIES, INV_COLUMNS
from modules.profile_calculators.aerofoil_fixing.aerofoil_common import (
    AEROFOIL_SECTION_MAPPER,
    AEROFOIL_ENDCAPS,
    AEROFOIL_MOVEABLE_ENDCAPS,
    AEROFOIL_TOP_PIVOTS,
    AEROFOIL_BOTTOM_PIVOTS,
    AEROFOIL_PACKING_WASHERS,
    AEROFOIL_PIVOT_T_WASHERS,
    AEROFOIL_ACCESSORIES,
    L_ANGLE_LENGTH,
)

BLACK_GYPSUM_PER_DIVISION = {"AF60": 4, "AF100": 4, "AF150": 8, "AF200": 8}


class AerofoilMoveable:

    def __init__(self, common_vars):
        self.af_type = common_vars["af_type"]
        self.pitch = common_vars["pitch"]

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

        return {
            "show_carriers": False,
            "show_endcaps":  False,
            "show_joints":   False,
            "bar_color":     "#aaa",
            "info_lines":    info_lines,
        }

    def run(self, data, stock_plan):

        data = data.copy()
        profile_code, profile_name = AEROFOIL_SECTION_MAPPER[self.af_type]
        endcap_code, endcap_name = AEROFOIL_ENDCAPS[self.af_type]
        mov_endcap_code, mov_endcap_name = AEROFOIL_MOVEABLE_ENDCAPS[self.af_type]
        top_pivot_code, top_pivot_name = AEROFOIL_TOP_PIVOTS[self.af_type]
        bot_pivot_code, bot_pivot_name = AEROFOIL_BOTTOM_PIVOTS[self.af_type]
        pack_washer_code, pack_washer_name = AEROFOIL_PACKING_WASHERS[self.af_type]
        pivot_tw_code, pivot_tw_name = AEROFOIL_PIVOT_T_WASHERS[self.af_type]

        bg_qty = BLACK_GYPSUM_PER_DIVISION.get(self.af_type, 4)
        pcs_per_l_angle = L_ANGLE_LENGTH / self.pitch

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
            height = int(row["height"])
            width = int(row["width"])

            tcl = height * 2
            no_l_angle = math.ceil(divisions / pcs_per_l_angle)

            items = [
                {
                    "Product Code": COMMON_ACCESSORIES["L_ANGLE_25X25"][0],
                    "Product Name": COMMON_ACCESSORIES["L_ANGLE_25X25"][1],
                    "Length": L_ANGLE_LENGTH,
                    "Quantity": no_l_angle,
                    "UOM": "m",
                    "CNC Hole Distance": self.pitch,
                },
                {
                    "Product Code": COMMON_ACCESSORIES["GRILLE_CARRIER"][0],
                    "Product Name": COMMON_ACCESSORIES["GRILLE_CARRIER"][1],
                    "Length": 3050,
                    "Quantity": int(math.ceil(tcl / 3050)),
                    "UOM": "m",
                    "CNC Hole Distance": self.pitch,
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_MOVEABLE_U_CHANNEL"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_MOVEABLE_U_CHANNEL"][1],
                    "Length": 3050,
                    "Quantity": int(math.ceil(height / 3050)),
                    "UOM": "m",
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_MOVEABLE_KNOB"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_MOVEABLE_KNOB"][1],
                    "Quantity": no_l_angle,
                    "UOM": "pcs",
                },
                {
                    "Product Code": endcap_code,
                    "Product Name": endcap_name,
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": mov_endcap_code,
                    "Product Name": mov_endcap_name,
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": top_pivot_code,
                    "Product Name": top_pivot_name,
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": bot_pivot_code,
                    "Product Name": bot_pivot_name,
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": pivot_tw_code,
                    "Product Name": pivot_tw_name,
                    "Quantity": divisions * 2,
                    "UOM": "pcs",
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_T_WASHER"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_T_WASHER"][1],
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_WASHER"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_WASHER"][1],
                    "Quantity": divisions * 2,
                    "UOM": "pcs",
                },
                {
                    "Product Code": pack_washer_code,
                    "Product Name": pack_washer_name,
                    "Quantity": int(math.ceil(divisions * 0.20)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                    "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                    "Quantity": int(math.ceil(divisions * bg_qty + width / 600)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": COMMON_ACCESSORIES["RIVET_6MM"][0],
                    "Product Name": COMMON_ACCESSORIES["RIVET_6MM"][1],
                    "Quantity": int(math.ceil(width * 2 / 300)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                    "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                    "Quantity": int(math.ceil(tcl / 500)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                    "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                    "Quantity": int(math.ceil(tcl / 500)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["HEX_NYLON_NUT_4MM"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["HEX_NYLON_NUT_4MM"][1],
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["CSK_PHILIPS_SCREW_4X20MM"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["CSK_PHILIPS_SCREW_4X20MM"][1],
                    "Quantity": divisions,
                    "UOM": "pcs",
                },
                {
                    "Product Code": AEROFOIL_ACCESSORIES["C_CLAMP"][0],
                    "Product Name": AEROFOIL_ACCESSORIES["C_CLAMP"][1],
                    "Quantity": int(math.ceil(divisions * 0.1)),
                    "UOM": "pcs",
                },
            ]

            for j, r in enumerate(items):
                r["Quantity"] = int(r["Quantity"] * qty_areas)
                r["item_order"] = sequence + j
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
        }
        offer_df = data[offer_df_cols.keys()].copy()

        return data, offer_df_cols, offer_df, inv_data
