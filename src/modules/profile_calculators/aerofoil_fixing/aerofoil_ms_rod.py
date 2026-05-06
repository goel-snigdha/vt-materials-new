import pandas as pd

from modules.excel_utils import INV_COLUMNS
from modules.profile_calculators.aerofoil_fixing.aerofoil_common import AEROFOIL_SECTION_MAPPER


class AerofoilMSRod:

    def __init__(self, common_vars):
        self.af_type = common_vars["af_type"]

    @staticmethod
    def generate_image(row, common_vars):

        pitch = common_vars.get("pitch", "")
        divisions = row["divisions"]
        cut_summary = row["cut_summary"]
        top_suspension = int(row.get("top_suspension", 0) or 0)
        bot_suspension = int(row.get("bottom_suspension", 0) or 0)
        orientation = row.get("orientation", "Vertical")
        division_length = int(row.get("single_division_length", 1) or 1)

        if len(cut_summary) > 1:
            cut_summary_str = "+".join(str(c) for c in cut_summary)
        else:
            cut_summary_str = f"Single Piece {cut_summary[0]} mm"

        info_lines = [
            (f"{divisions} Divisions",          "#333", 12, True),
            (f"@ {pitch} mm Pitch",             "#333", 12, True),
            ("",                                "#333", 12, False),
            ("Breakdown",                       "#333", 12, False),
            (f"{cut_summary_str}",              "#333", 12, True),
            ("",                                "#333", 12, False),
            (f"Top Suspension: {top_suspension} mm",  "#333", 12, False),
            (f"Bot Suspension: {bot_suspension} mm",  "#333", 12, False),
        ]

        MS_ROD_COLOR = "#E67E22"

        def draw_ms_rods(ax, _row, W, H, _total):
            N_BARS = 20  # match the hardcoded visual bar count in excel_utils.generate_window_image
            MAX_STUB = 0.35  # max fraction of H (or W) a stub can occupy

            ref = division_length if division_length > 0 else 1
            top_stub = min((top_suspension / ref) * H, H * MAX_STUB)
            bot_stub = min((bot_suspension / ref) * H, H * MAX_STUB)

            if orientation == "Vertical":
                # Bars run vertically — rod enters from top and bottom ends
                for i in range(N_BARS):
                    x = (i + 0.5) * W / N_BARS
                    if top_suspension > 0:
                        ax.plot([x, x], [H, H - top_stub], color=MS_ROD_COLOR, linewidth=3, zorder=6)
                    if bot_suspension > 0:
                        ax.plot([x, x], [0, bot_stub], color=MS_ROD_COLOR, linewidth=3, zorder=6)
                # Labels outside the frame above and below
                if top_suspension > 0:
                    ax.text(W / 2, H - top_stub - 4, f"{top_suspension} mm Suspension",
                            ha="center", va="top", fontsize=8, color=MS_ROD_COLOR, fontweight="bold")
                if bot_suspension > 0:
                    ax.text(W / 2, bot_stub + 4, f"{bot_suspension} mm Suspension",
                            ha="center", va="bottom", fontsize=8, color=MS_ROD_COLOR, fontweight="bold")
            else:
                top_stub_h = min((top_suspension / ref) * W, W * MAX_STUB)
                bot_stub_h = min((bot_suspension / ref) * W, W * MAX_STUB)
                # Bars run horizontally — rod enters from left and right ends
                for i in range(N_BARS):
                    y = (i + 0.5) * H / N_BARS
                    if top_suspension > 0:
                        ax.plot([0, top_stub_h], [y, y], color=MS_ROD_COLOR, linewidth=3, zorder=6)
                    if bot_suspension > 0:
                        ax.plot([W - bot_stub_h, W], [y, y], color=MS_ROD_COLOR, linewidth=3, zorder=6)
                # Labels outside the frame on left and right
                if top_suspension > 0:
                    ax.text(-8, H / 2, f"{top_suspension} mm Suspension",
                            ha="right", va="center", fontsize=8, color=MS_ROD_COLOR, fontweight="bold", rotation=90)
                if bot_suspension > 0:
                    ax.text(W + 8, H / 2, f"{bot_suspension} mm Suspension",
                            ha="left", va="center", fontsize=8, color=MS_ROD_COLOR, fontweight="bold", rotation=90)

        from matplotlib.lines import Line2D
        return {
            "show_carriers": False,
            "show_endcaps":  False,
            "show_joints":   False,
            "bar_color":     "#aaa",
            "info_lines":    info_lines,
            "extras":        draw_ms_rods,
            "legend_extras": [
                Line2D([0], [0], color=MS_ROD_COLOR, linewidth=2, label="MS Rod"),
            ],
        }

    def run(self, data, stock_plan):

        data = data.copy()
        data["ms_rod_length"] = (
            (data["top_suspension"].fillna(0).astype(int) + 50 + data["bottom_suspension"].fillna(0).astype(int) + 50)
            * data["divisions"]
        ) / 1000

        profile_code, profile_name = AEROFOIL_SECTION_MAPPER[self.af_type]

        inventory_df = pd.DataFrame([
            {
                "Product Code": profile_code,
                "Product Name": profile_name,
                "Length": item["length"],
                "Quantity": item["qty"],
                "UOM": "m",
            }
            for item in sorted(stock_plan, key=lambda x: x["length"], reverse=True)
        ]).reindex(columns=INV_COLUMNS).fillna("")

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
            "ms_rod_length": {"type": "formula", "hide_if_zero": False},
        }
        offer_df = data[offer_df_cols.keys()].copy()

        return data, offer_df_cols, offer_df, inventory_df
