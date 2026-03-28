import math
import streamlit as st
import pandas as pd
from matplotlib.lines import Line2D

import modules.profile_utils as profile_utils
from modules.excel_utils import (
    INV_COLUMNS,
    PIPE_MAPPER,
    COMMON_ACCESSORIES,
    COVERING_OPTIONS,
    COVERING_CODE_MAP,
)

# Fluted specific codes from excel_utils.py
FLUTED_PRODUCTS = {
    "PROFILE": ("FL-PR-01", "FLUTED PROFILE"),
    "START_PIECE": ("FL-SP-02", "FLUTED START PIECE"),
    "CORNER_PIECE": ("FL-CR-02", "FLUTED CORNER PIECE"),
}


def generate_offer_df(data):

    offer_df_cols = {
        "s_no": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "area_name": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "orientation": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "height": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "width": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "qty_areas": {
            "type": "desc",
            "hide_if_zero": False,
        },
        "area_sqft": {
            "type": "formula",
            "hide_if_zero": False,
        },
        "divisions": {
            "type": "formula",
            "hide_if_zero": False,
        },
        "total_product_length": {
            "type": "formula",
            "hide_if_zero": False,
        },
        "epdm_length": {
            "type": "formula",
            "hide_if_zero": False,
        },
    }
    offer_df = data[offer_df_cols.keys()].copy()

    return offer_df_cols, offer_df


def generate_inventory_df(data, pipe_grade, stock_plan):

    all_inventory_rows = []

    for _, row in data.iterrows():

        sequence = 0

        profile_rows = [
            {
                "Product Code": FLUTED_PRODUCTS["PROFILE"][0],
                "Product Name": FLUTED_PRODUCTS["PROFILE"][1],
                "Length": item["length"],
                "Quantity": item["qty"],
                "UOM": "m",
            }
            for item in stock_plan
        ]

        pipe_code, pipe_name = PIPE_MAPPER[pipe_grade]

        additional_items = [
            {
                "Product Code": FLUTED_PRODUCTS["START_PIECE"][0],
                "Product Name": FLUTED_PRODUCTS["START_PIECE"][1],
                "Length": row["single_division_length"],
                "Quantity": 1,
                "UOM": "m",
            },
            {
                "Product Code": pipe_code,
                "Product Name": pipe_name,
                "Length": 3650,
                "Quantity": int(math.ceil(row["total_carrier_length"] / 3650)),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["EPDM_GASKET"][0],
                "Product Name": COMMON_ACCESSORIES["EPDM_GASKET"][1],
                "Quantity": int(math.ceil(row["total_product_length"])),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": row["divisions"] * row["total_carrier_divisions"],
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                "Quantity": int(math.ceil(row["total_carrier_length"] / 500)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                "Quantity": int(math.ceil(row["total_carrier_length"] / 500)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PAINT"][0],
                "Product Name": COMMON_ACCESSORIES["PAINT"][1],
                "Quantity": 0,
                "UOM": "l",
            },
        ]

        if row.get("total_frame_covering_length", 0) > 0:
            additional_items.append(
                {
                    "Product Code": row["frame_covering_type"],
                    "Product Name": COVERING_CODE_MAP[row["frame_covering_type"]],
                    "Length": 3050,
                    "Quantity": int(
                        math.ceil(row["total_frame_covering_length"] / 3050)
                    ),
                    "UOM": "m",
                }
            )

        if row.get("has_corner_piece", False):
            additional_items.insert(
                1,
                {
                    "Product Code": FLUTED_PRODUCTS["CORNER_PIECE"][0],
                    "Product Name": FLUTED_PRODUCTS["CORNER_PIECE"][1],
                    "Length": row["single_division_length"],
                    "Quantity": 1,
                    "UOM": "m",
                },
            )

        row_items = profile_rows + additional_items

        for item in row_items:
            item["Quantity"] = int(item.get("Quantity", 0) * row["qty_areas"])
            item["item_order"] = sequence
            sequence += 1

        all_inventory_rows.extend(row_items)

    inv_data = (
        pd.DataFrame(all_inventory_rows)
        .reindex(columns=INV_COLUMNS + ["item_order"])
        .fillna("")
    )

    inv_data = (
        inv_data.groupby(
            [
                "Product Code",
                "Product Name",
                "Length",
                "UOM",
                "Colour",
                "Finish",
                "CNC Hole Distance",
                "Remarks",
            ],
            as_index=False,
            sort=False,
        )
        .agg({"Quantity": "sum", "item_order": "min"})
        .sort_values(["item_order"])
        .drop(columns=["item_order"])
    )

    return inv_data


def parse_covering(val):
    if val is None:
        return []
    if isinstance(val, list):
        return [str(s).strip() for s in val]
    if isinstance(val, str):
        val = val.strip().strip("[]").replace("'", "").replace('"', "")
        return [s.strip() for s in val.split(",") if s.strip()]
    return []


def calc_frame_row(row):

    covering = row["frame_covering"]
    width = row["width"]
    height = row["height"]
    total = 0

    for side in covering:
        if side in ("Top", "Bottom"):
            total += width
        elif side in ("Left", "Right"):
            total += height

    return total


class FlutedCalculator:

    def __init__(self, vars):

        self.vars = vars
        self.project_title = vars["project_title"]
        self.pipe_grade = vars["pipe_grade"]
        self.pitch = vars["pitch"]
        self.areas = vars["areas"]

    def validate_input(row, idx, df):

        frame_covering = parse_covering(row.get("frame_covering"))
        covering_type = str(row.get("frame_covering_type", "")).strip()

        has_none = "None" in frame_covering
        has_other = any(s != "None" for s in frame_covering)

        if has_none and has_other:
            st.warning(
                f"Row {idx + 1}: 'None' cannot be selected together with other frame covering options."
            )
            return False

        if has_none and covering_type not in ("", "nan", "None"):
            st.warning(
                f"Row {idx + 1}: Frame Covering Type must be empty when 'None' is selected."
            )
            return False

        has_covering = has_other
        if has_covering and (not covering_type or covering_type in ("", "nan", "None")):
            st.warning(
                f"Row {idx + 1}: Frame Covering Type is required when Frame Covering is selected."
            )
            return False

        return True

    def validate_corner_joints(corner_df, area_df):

        if corner_df is None or len(corner_df.dropna(how="all")) == 0:
            return True

        INVERSE_SIDE = {
            "Top": "Bottom",
            "Bottom": "Top",
            "Left": "Right",
            "Right": "Left",
        }
        VALID_CORNER_VERTICAL = {"Left", "Right"}
        VALID_CORNER_HORIZONTAL = {"Top", "Bottom"}

        all_snos = area_df["s_no"].dropna().astype(str).tolist()
        used_sides = {}

        for idx, row in corner_df.dropna(how="all").iterrows():
            w1 = str(row.get("window_1", "")).strip()
            w2 = str(row.get("window_2", "")).strip()
            dir1 = str(row.get("direction_on_window_1", "")).strip()

            # Both windows must exist
            if w1 not in all_snos:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w1}' does not exist.")
                return False
            if w2 not in all_snos:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w2}' does not exist.")
                return False

            # Window 1 and 2 cannot be the same
            if w1 == w2:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window 1 and Window 2 cannot be the same."
                )
                return False

            # Direction must be filled
            if not dir1 or dir1 in ("", "nan", "None"):
                st.warning(
                    f"Corner Joint Row {idx + 1}: Direction on Window 1 is required."
                )
                return False

            w1_row = area_df[area_df["s_no"].astype(str) == w1].iloc[0]
            w2_row = area_df[area_df["s_no"].astype(str) == w2].iloc[0]
            w1_orient = str(w1_row.get("orientation", "")).strip()
            w2_orient = str(w2_row.get("orientation", "")).strip()

            # Same orientation check
            if w1_orient != w2_orient:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Windows '{w1}' and '{w2}' have different orientations."
                )
                return False

            # Direction valid for orientation
            if w1_orient == "Vertical" and dir1 not in VALID_CORNER_VERTICAL:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Direction '{dir1}' not valid for Vertical. Use Left or Right."
                )
                return False
            if w1_orient == "Horizontal" and dir1 not in VALID_CORNER_HORIZONTAL:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Direction '{dir1}' not valid for Horizontal. Use Top or Bottom."
                )
                return False

            # Corner cannot be in the same direction as frame covering for window 1
            w1_covering = w1_row.get("frame_covering") or []
            if isinstance(w1_covering, str):
                w1_covering = [s.strip() for s in w1_covering.split(",") if s.strip()]
            if dir1 in w1_covering:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window '{w1}' already has a frame covering on '{dir1}' — "
                    f"cannot place a corner joint on the same side."
                )
                return False

            # Corner cannot be in the same direction as frame covering for window 2
            dir2 = INVERSE_SIDE[dir1]
            w2_covering = w2_row.get("frame_covering") or []
            if isinstance(w2_covering, str):
                w2_covering = [s.strip() for s in w2_covering.split(",") if s.strip()]
            if dir2 in w2_covering:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window '{w2}' has a frame covering on '{dir2}' — "
                    f"cannot place a corner joint on the same side."
                )
                return False

            # Same single_division_length
            w1_sdl = (
                w1_row.get("width")
                if w1_orient == "Horizontal"
                else w1_row.get("height")
            )
            w2_sdl = (
                w2_row.get("width")
                if w2_orient == "Horizontal"
                else w2_row.get("height")
            )
            if int(w1_sdl) != int(w2_sdl):
                st.warning(
                    f"Corner Joint Row {idx + 1}: Windows '{w1}' and '{w2}' have different division lengths ({w1_sdl}mm vs {w2_sdl}mm)."
                )
                return False

            # No duplicate sides per window
            if w1 not in used_sides:
                used_sides[w1] = set()
            if w2 not in used_sides:
                used_sides[w2] = set()

            if dir1 in used_sides[w1]:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window '{w1}' already has a corner on '{dir1}'."
                )
                return False
            if dir2 in used_sides[w2]:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window '{w2}' already has a corner on '{dir2}'."
                )
                return False

            used_sides[w1].add(dir1)
            used_sides[w2].add(dir2)

            # qty_areas must be 1 for both
            if int(w1_row.get("qty_areas", 1)) != 1:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window '{w1}' cannot have Similar Areas > 1."
                )
                return False
            if int(w2_row.get("qty_areas", 1)) != 1:
                st.warning(
                    f"Corner Joint Row {idx + 1}: Window '{w2}' cannot have Similar Areas > 1."
                )
                return False

        return True

    def get_validator(df, corner_df=None):

        # Normalize frame_covering column in df upfront
        if "frame_covering" in df.columns:
            df = df.copy()
            df["frame_covering"] = df["frame_covering"].apply(parse_covering)

        # Run corner validation once upfront
        if corner_df is not None and len(corner_df.dropna(how="all")) > 0:
            corner_valid = FlutedCalculator.validate_corner_joints(corner_df, df)
            if not corner_valid:
                return lambda row, idx: False

        def validator(row, idx):
            return FlutedCalculator.validate_input(row, idx, df)

        return validator

    def get_data_input():

        empty_df = pd.DataFrame(
            {
                "s_no": pd.Series(dtype="str"),
                "area_name": pd.Series(dtype="str"),
                "height": pd.Series(dtype="int"),
                "width": pd.Series(dtype="int"),
                "orientation": pd.Series(dtype="str"),
                "qty_areas": pd.Series(dtype="int"),
                "cut_summary": pd.Series(dtype="str"),
                "frame_covering": pd.Series(["None"], dtype="object"),
                "frame_covering_type": pd.Series(dtype="str"),
            }
        )

        required_cols = [
            "width",
            "height",
            "orientation",
            "qty_areas",
            "cut_summary",
            "frame_covering",
        ]

        input_data = st.data_editor(
            data=empty_df,
            column_config={
                "s_no": st.column_config.TextColumn("S. No", required=False),
                "area_name": st.column_config.TextColumn("Area Name", required=False),
                "width": st.column_config.NumberColumn(
                    "Width (mm)", min_value=1, step=1, required=True
                ),
                "height": st.column_config.NumberColumn(
                    "Height (mm)", min_value=1, step=1, required=True
                ),
                "orientation": st.column_config.SelectboxColumn(
                    "Orientation", options=["Horizontal", "Vertical"], required=True
                ),
                "qty_areas": st.column_config.NumberColumn(
                    "Similar Areas", min_value=1, step=1, required=True
                ),
                "cut_summary": st.column_config.TextColumn(
                    "Cut Summary", required=True
                ),
                "frame_covering": st.column_config.MultiselectColumn(
                    "Frame Covering",
                    options=["None", "Top", "Bottom", "Left", "Right"],
                    required=True,
                ),
                "frame_covering_type": st.column_config.SelectboxColumn(
                    "Frame Covering Type", options=COVERING_OPTIONS, required=False
                ),
            },
            num_rows="dynamic",
        )

        st.markdown("**Corner Joints** (optional)")
        corner_df = pd.DataFrame(
            {
                "window_1": pd.Series(dtype="str"),
                "window_2": pd.Series(dtype="str"),
                "direction_on_window_1": pd.Series(dtype="str"),
            }
        )
        corner_input = st.data_editor(
            data=corner_df,
            column_config={
                "window_1": st.column_config.TextColumn("Window 1 S.No"),
                "window_2": st.column_config.TextColumn("Window 2 S.No"),
                "direction_on_window_1": st.column_config.SelectboxColumn(
                    "Direction on Window 1", options=["Top", "Bottom", "Left", "Right"]
                ),
            },
            num_rows="dynamic",
        )

        return input_data, required_cols, corner_input

    def generate_image(row, common_vars):

        pitch = common_vars.get("pitch", "")
        divisions = row["divisions"]
        cut_summary = row["cut_summary"]
        frame_covering = row["frame_covering"]
        covering_type = str(row.get("frame_covering_type", "")).strip()
        corner_joints = common_vars.get("corner_joints")
        this_sno = str(row.get("s_no", "")).strip()

        INVERSE_SIDE = {
            "Top": "Bottom",
            "Bottom": "Top",
            "Left": "Right",
            "Right": "Left",
        }

        # Find all corner directions for this window from the corner_joints table
        corner_dirs = []
        if corner_joints is not None and len(corner_joints.dropna(how="all")) > 0:
            for _, cj in corner_joints.iterrows():
                w1 = str(cj.get("window_1", "")).strip()
                w2 = str(cj.get("window_2", "")).strip()
                dir1 = str(cj.get("direction_on_window_1", "")).strip()
                if w1 == this_sno and dir1:
                    corner_dirs.append(dir1)
                elif w2 == this_sno and dir1:
                    corner_dirs.append(INVERSE_SIDE.get(dir1, ""))

        has_corner = len(corner_dirs) > 0
        has_covering = len(frame_covering) > 0 and frame_covering != ["None"]

        if len(cut_summary) > 1:
            cut_summary_str = "+".join(str(c) for c in cut_summary)
        else:
            cut_summary_str = f"Single Piece {cut_summary[0]} mm"

        covering_label = covering_type if covering_type else "Frame Covering"

        info_lines = [
            (f"{divisions} Divisions", "#333", 12, True),
            (f"@ {pitch} mm Pitch", "#333", 12, True),
            ("", "#333", 12, False),
            ("Breakdown", "#333", 12, False),
            (f"{cut_summary_str}", "#333", 12, True),
        ]

        if has_covering:
            info_lines += [
                ("", "#333", 12, False),
                ("Frame Covering", "#333", 12, False),
                (f"{covering_label}", "#333", 12, True),
                (", ".join(frame_covering), "#333", 12, True),
            ]

        if has_corner:
            info_lines += [
                ("", "#333", 12, False),
                ("Corner", "#333", 12, False),
                (", ".join(corner_dirs), "#333", 12, True),
            ]

        def extras(ax, row, W, H, total):

            COVERING_COLOR = "#9B30FF"
            CORNER_COLOR = "#00C851"
            LINE_WIDTH = 4

            SIDE_MAP = {
                "Top": [(0, H), (W, H)],
                "Bottom": [(0, 0), (W, 0)],
                "Left": [(0, 0), (0, H)],
                "Right": [(W, 0), (W, H)],
            }

            # Frame covering lines — one per selected side
            for side in frame_covering:
                if side in SIDE_MAP:
                    (x1, y1), (x2, y2) = SIDE_MAP[side]
                    ax.plot(
                        [x1, x2],
                        [y1, y2],
                        color=COVERING_COLOR,
                        linewidth=LINE_WIDTH,
                        zorder=6,
                    )

            # Corner joint lines — one per direction
            for cd in corner_dirs:
                if cd in SIDE_MAP:
                    (x1, y1), (x2, y2) = SIDE_MAP[cd]
                    ax.plot(
                        [x1, x2],
                        [y1, y2],
                        color=CORNER_COLOR,
                        linewidth=LINE_WIDTH,
                        zorder=6,
                    )

        legend_extras = []
        if has_covering:
            legend_extras.append(
                Line2D([0], [0], color="#9B30FF", linewidth=3, label=covering_label)
            )
        if has_corner:
            legend_extras.append(
                Line2D([0], [0], color="#00C851", linewidth=3, label="Corner")
            )

        config = {
            "show_carriers": True,
            "show_endcaps": False,
            "show_joints": True,
            "bar_color": "#888",
            "info_lines": info_lines,
            "extras": extras,
            "legend_extras": legend_extras,
        }
        return config

    def run(self):

        data = self.areas.copy()
        corner_joints = self.vars["corner_joints"]

        # PROFILE
        data["req_plan"] = data.apply(profile_utils.build_req_plan, axis=1)
        # master_req = profile_utils.combine_req_plan(data["req_plan"])
        out = profile_utils.optimize_stock_v2(data)
        data["cut_plan"] = data.apply(
            lambda row: profile_utils.build_window_cut_plan(row, out["bars"]), axis=1
        )

        # CARRIER
        data["carrier_distances"] = data.apply(
            lambda row: profile_utils.calculate_carrier_distances(row["cut_summary"]),
            axis=1,
        )
        data["total_carrier_divisions"] = data["carrier_distances"].apply(
            lambda x: sum(len(sublist) for sublist in x)
        )
        data["total_carrier_length"] = data.apply(
            lambda row: row["total_carrier_divisions"] * row["perpendicular_length"],
            axis=1,
        )

        # FRAME COVERING
        data["frame_covering"] = data["frame_covering"].apply(parse_covering)
        data["total_frame_covering_length"] = data.apply(calc_frame_row, axis=1)

        # ACCS
        data["epdm_length"] = data["total_product_length"]

        corner_window_1s = set()
        if corner_joints is not None and len(corner_joints.dropna(how="all")) > 0:
            for _, cj in corner_joints.iterrows():
                w1 = str(cj.get("window_1", "")).strip()
                if w1:
                    corner_window_1s.add(w1)
        (corner_window_1s)

        data["has_corner_piece"] = data["s_no"].astype(str).isin(corner_window_1s)

        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_df = generate_inventory_df(data, self.pipe_grade, out["stock_plan"])

        results = [data, offer_df_cols, offer_df, inventory_df]

        return results
