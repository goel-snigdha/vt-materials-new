import pandas as pd
import numpy as np
import streamlit as st

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS
from modules.profile_calculators.aerofoil_fixing.aerofoil_common import C_PLATE_OPTIONS
from modules.profile_calculators.aerofoil_fixing.aerofoil_fringe import AerofoilFringe
from modules.profile_calculators.aerofoil_fixing.aerofoil_c_channel import AerofoilCChannel
from modules.profile_calculators.aerofoil_fixing.aerofoil_ms_rod import AerofoilMSRod
from modules.profile_calculators.aerofoil_fixing.aerofoil_d_wall import AerofoilDWall
from modules.profile_calculators.aerofoil_fixing.aerofoil_moveable import AerofoilMoveable


def generate_offer_df(data):

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
    return offer_df_cols, offer_df


class AerofoilCalculator:

    def __init__(self, vars):
        self.vars = vars
        self.areas = vars["areas"]
        self.pitch = vars["pitch"]
        self.af_type = vars.get("af_type", "AF100")
        self.fixing_method = vars.get("fixing_method", None)

    def get_data_input(**kwargs):

        af_type = kwargs.get("af_type", "AF100")
        fixing_method = kwargs.get("fixing_method", None)

        base_cols = {
            "s_no":        pd.Series(dtype="str"),
            "area_name":   pd.Series(dtype="str"),
            "height":      pd.Series(dtype="int"),
            "width":       pd.Series(dtype="int"),
            "orientation": pd.Series(dtype="str"),
            "qty_areas":   pd.Series(dtype="int"),
        }

        extra_cols = {}
        if fixing_method == "C-Channel":
            extra_cols["plate_width"] = pd.Series(dtype="int")
        elif fixing_method == "MS Rod/Slot Cut Pipe":
            extra_cols["top_suspension"] = pd.Series(dtype="int")
            extra_cols["bottom_suspension"] = pd.Series(dtype="int")
        elif fixing_method == "D-Wall Bracket":
            extra_cols["endcaps"] = pd.Series(dtype="str")

        empty_df = pd.DataFrame({**base_cols, **extra_cols, "cut_summary": pd.Series(dtype="str")})

        required_cols = ["width", "height", "orientation", "qty_areas", "cut_summary"]

        col_config = {
            "s_no":        st.column_config.TextColumn("S. No",          required=False),
            "area_name":   st.column_config.TextColumn("Area Name",      required=False),
            "width":       st.column_config.NumberColumn("Width (mm)",   min_value=1, step=1, required=True),
            "height":      st.column_config.NumberColumn("Height (mm)",  min_value=1, step=1, required=True),
            "orientation": st.column_config.SelectboxColumn("Orientation", options=["Horizontal", "Vertical"], required=True),
            "qty_areas":   st.column_config.NumberColumn("Similar Areas", min_value=1, step=1, required=True),
            "cut_summary": st.column_config.TextColumn("Cut Summary",    required=True),
        }

        if fixing_method == "C-Channel":
            plate_options = [str(w) for w in C_PLATE_OPTIONS.get(af_type, [75, 100, 112])]
            col_config["plate_width"] = st.column_config.SelectboxColumn(
                "C-Plate Width (mm)", options=plate_options, required=True
            )
            required_cols.append("plate_width")
        elif fixing_method == "MS Rod/Slot Cut Pipe":
            col_config["top_suspension"] = st.column_config.NumberColumn(
                "Top Suspension (mm)", min_value=0, step=1, required=True
            )
            col_config["bottom_suspension"] = st.column_config.NumberColumn(
                "Bottom Suspension (mm)", min_value=0, step=1, required=True
            )
            required_cols += ["top_suspension", "bottom_suspension"]
        elif fixing_method == "D-Wall Bracket":
            col_config["endcaps"] = st.column_config.SelectboxColumn(
                "Endcaps",
                options=["No Endcaps", "Both Sides", "Top", "Bottom", "Left", "Right"],
                required=True,
            )
            required_cols.append("endcaps")

        input_data = st.data_editor(
            data=empty_df,
            column_config=col_config,
            num_rows="dynamic",
        )

        return input_data, required_cols

    def validate_input(row, idx):
        endcaps = str(row.get("endcaps", "") or "").strip()
        orientation = str(row.get("orientation", "") or "").strip()

        VERTICAL_ONLY = {"Top", "Bottom"}
        HORIZONTAL_ONLY = {"Left", "Right"}

        if endcaps in VERTICAL_ONLY and orientation != "Vertical":
            st.warning(f"Row {idx + 1}: Endcap side '{endcaps}' is only valid for Vertical orientation.")
            return False
        if endcaps in HORIZONTAL_ONLY and orientation != "Horizontal":
            st.warning(f"Row {idx + 1}: Endcap side '{endcaps}' is only valid for Horizontal orientation.")
            return False
        return True

    def get_validator(df, corner_df=None):
        def validator(row, idx):
            return AerofoilCalculator.validate_input(row, idx)
        return validator

    def generate_image(row, common_vars):
        fixing_method = common_vars.get("fixing_method", None)

        if fixing_method == "Moveable (Manual)":
            return AerofoilMoveable.generate_image(row, common_vars)
        if fixing_method == "Fringe End Caps":
            return AerofoilFringe.generate_image(row, common_vars)
        if fixing_method == "C-Channel":
            return AerofoilCChannel.generate_image(row, common_vars)
        if fixing_method == "MS Rod/Slot Cut Pipe":
            return AerofoilMSRod.generate_image(row, common_vars)
        if fixing_method == "D-Wall Bracket":
            return AerofoilDWall.generate_image(row, common_vars)
        return {}

    def _get_subclass(self):
        if self.fixing_method == "Moveable (Manual)":
            return AerofoilMoveable(self.vars)
        if self.fixing_method == "Fringe End Caps":
            return AerofoilFringe(self.vars)
        if self.fixing_method == "C-Channel":
            return AerofoilCChannel(self.vars)
        if self.fixing_method == "MS Rod/Slot Cut Pipe":
            return AerofoilMSRod(self.vars)
        if self.fixing_method == "D-Wall Bracket":
            return AerofoilDWall(self.vars)
        return None

    def run(self):

        data = self.areas.copy()
        data["width"] = data["width"].astype(int)
        data["height"] = data["height"].astype(int)
        data["qty_areas"] = data["qty_areas"].astype(int)

        # main.py computes generic divisions but aerofoil needs product-specific overrides
        # (moveable end gaps, MS Rod suspension reductions) — recalculate here intentionally.

        # For moveable, reduce the division width by the suspension end gaps
        if self.fixing_method == "Moveable (Manual)":
            af_width = int(self.af_type.strip("AF"))
            ends_gap = (af_width / 2) + 10
            data["single_division_length"] = data.apply(
                lambda r: (
                    r["height"] - (ends_gap * 2)
                    if r["orientation"] == "Vertical"
                    else r["width"] - (ends_gap * 2)
                ),
                axis=1,
            )

            data["divisions"] = np.ceil(
                data["perpendicular_length"] / self.pitch
            ).astype(int)

        elif self.fixing_method == "MS Rod/Slot Cut Pipe":
            data["single_division_length"] = data.apply(
                lambda r: (
                    r["width"] - (
                        int(r.get("top_suspension", 0) or 0) / 2 +
                        int(r.get("bottom_suspension", 0) or 0) / 2
                    )
                    if r["orientation"] == "Horizontal"
                    else r["height"] - (
                        int(r.get("top_suspension", 0) or 0) / 2 +
                        int(r.get("bottom_suspension", 0) or 0) / 2
                    )
                ),
                axis=1,
            )

        if self.fixing_method in ["Moveable (Manual)", "MS Rod/Slot Cut Pipe"]:

            data["divisions"] = np.ceil(
                data["perpendicular_length"] / self.pitch
            ).astype(int)
            data["total_product_length"] = (
                data["single_division_length"] * data["divisions"]
            ) / 1000
            data["area_sqft"] = (data["width"] * data["height"]) / 92903.04

        # Carrier distances needed for D-Wall
        if self.fixing_method == "D-Wall Bracket":
            data["carrier_distances"] = data.apply(
                lambda r: profile_utils.calculate_carrier_distances(r["cut_summary"]),
                axis=1,
            )
            data["total_carrier_divisions"] = data["carrier_distances"].apply(
                lambda x: sum(len(s) for s in x)
            )

        # Cut plan
        data["req_plan"] = data.apply(profile_utils.build_req_plan, axis=1)
        out = profile_utils.optimize_stock_v2(data)
        data["cut_plan"] = data.apply(
            lambda r: profile_utils.build_window_cut_plan(r, out["bars"]), axis=1
        )

        subclass = self._get_subclass()
        if subclass is None:
            offer_df_cols, offer_df = generate_offer_df(data)
            inventory_df = pd.DataFrame(columns=INV_COLUMNS)
            return [data, offer_df_cols, offer_df, inventory_df]

        data, offer_df_cols, offer_df, inventory_df = subclass.run(data, out["stock_plan"])
        return [data, offer_df_cols, offer_df, inventory_df]
