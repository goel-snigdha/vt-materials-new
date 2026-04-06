import math
from pathlib import Path
import streamlit as st
import pandas as pd

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, PIPE_MAPPER, COMMON_ACCESSORIES

COTTAL_PRODUCTS = {
    "85 mm":  {
        "PROFILE": ("CT-PR-02", "COTTAL PROFILE"),
        "START_PIECE": ("CT-SP-02", "COTTAL START PIECE"),
        "COVERING_PIECE": ("CT-CP-02", "COTTAL COVERING PIECE")
    },
    "130 mm": {
        "PROFILE": ("CT-PR-02", "COTTAL PROFILE"),
        "START_PIECE": ("CT-SP-02", "COTTAL START PIECE"),
        "COVERING_PIECE": ("CT-CP-02", "COTTAL COVERING PIECE")
    },
    "230 mm": {
        "PROFILE": ("CT-PR-02", "COTTAL PROFILE"),
        "START_PIECE": ("CT-SP-02", "COTTAL START PIECE"),
        "COVERING_PIECE": ("CT-CP-02", "COTTAL COVERING PIECE")
    },
}

COTTAL_ACCESSORIES = {
    "L_ANGLES":      ("", "SELECT COTTAL L-ANGLES"),
    "CORNER_PIECES": ("", "SELECT COTTAL CORNER PIECES"),
}

CORNER_JOINT_TYPES = [
    "COTTAL CORNER 1 VTJ-21/24",
    "COTTAL CORNER 2 VTJ-3A/24",
    "COTTAL CORNER 3 VTJ-3B/24",
    "COTTAL CORNER 4 VTJ-20/24",
]


def generate_offer_df(data):
    offer_df_cols = {
        "s_no":                 {"type": "desc",    "hide_if_zero": False},
        "area_name":            {"type": "desc",    "hide_if_zero": False},
        "orientation":          {"type": "desc",    "hide_if_zero": False},
        "height":               {"type": "desc",    "hide_if_zero": False},
        "width":                {"type": "desc",    "hide_if_zero": False},
        "qty_areas":            {"type": "desc",    "hide_if_zero": False},
        "area_sqft":            {"type": "formula", "hide_if_zero": False},
        "divisions":            {"type": "formula", "hide_if_zero": False},
        "total_product_length": {"type": "formula", "hide_if_zero": False},
        "epdm_gasket_length":   {"type": "formula", "hide_if_zero": False},
    }
    offer_df = data[offer_df_cols.keys()].copy()
    return offer_df_cols, offer_df


def generate_inventory_df(data, pipe_grade, louver_size, stock_plan, corner_joints=None):

    all_inventory_rows = []
    sequence = 0

    cottal_type = COTTAL_PRODUCTS[louver_size]

    # Build set of window_1 s_nos that have corners and their joint types
    corner_map = {}  # {s_no: joint_type}
    if corner_joints is not None and len(corner_joints.dropna(how="all")) > 0:
        for _, cj in corner_joints.iterrows():
            w1 = str(cj.get("window_1", "")).strip()
            joint_type = str(cj.get("joint_type", "")).strip()
            if w1:
                corner_map[w1] = joint_type

    for _, row in data.iterrows():

        profile_rows = []
        for item in stock_plan:
            profile_rows += [
                {
                    "Product Code": cottal_type["PROFILE"][0],
                    "Product Name": cottal_type["PROFILE"][1],
                    "Length": item["length"],
                    "Quantity": item["qty"],
                    "UOM": "m",
                    "Colour": "COLOURED",
                },
                {
                    "Product Code": cottal_type["START_PIECE"][0],
                    "Product Name": cottal_type["START_PIECE"][1],
                    "Length": item["length"],
                    "Quantity": 1,
                    "UOM": "m",
                    "Colour": "COLOURED",
                },
                {
                    "Product Code": cottal_type["COVERING_PIECE"][0],
                    "Product Name": cottal_type["COVERING_PIECE"][1],
                    "Length": item["length"],
                    "Quantity": 2,
                    "UOM": "m",
                    "Colour": "COLOURED",
                },
            ]

        pipe_code, pipe_name = PIPE_MAPPER[pipe_grade]
        paint_qty = round(row["divisions"] / 50 * 2) / 2
        brush_qty = math.ceil(paint_qty)

        additional_items = [
            {
                "Product Code": pipe_code,
                "Product Name": pipe_name,
                "Length": 3650,
                "Quantity": int(math.ceil(row["total_carrier_length"] / 3650)),
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
                "Product Code": COMMON_ACCESSORIES["EPDM_GASKET"][0],
                "Product Name": COMMON_ACCESSORIES["EPDM_GASKET"][1],
                "Quantity": int(math.ceil(row["total_product_length"])),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": int(math.ceil(row["divisions"] * row["total_carrier_divisions"])),
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

        # Add corner piece only for window_1 of each corner pair
        sno = str(row.get("s_no", "")).strip()
        if sno in corner_map:
            joint_type = corner_map[sno]
            additional_items.insert(1, {
                "Product Code": "",
                "Product Name": joint_type,
                "Quantity": 1,
                "UOM": "pcs",
                "Colour": "COLOURED",
            })

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
        inv_data.sort_values("item_order")
        .groupby(
            ["Product Code", "Product Name", "Length", "UOM",
             "Colour", "Finish", "CNC Hole Distance", "Remarks"],
            as_index=False,
            sort=False
        )
        .agg({"Quantity": "sum", "item_order": "min"})
        .sort_values("item_order")
        .drop(columns=["item_order"])
    )

    return inv_data


class CottalCalculator:

    def __init__(self, vars):
        self.project_title = vars["project_title"]
        self.pitch = vars["pitch"]
        self.pipe_grade = vars["pipe_grade"]
        self.louver_size = vars["louver_size"]
        self.areas = vars["areas"]
        self.vars = vars

    def validate_corner_joints(corner_df, area_df):
        if corner_df is None or len(corner_df.dropna(how="all")) == 0:
            return True

        INVERSE_SIDE = {"Top": "Bottom", "Bottom": "Top", "Left": "Right", "Right": "Left"}
        VALID_CORNER_VERTICAL = {"Left", "Right"}
        VALID_CORNER_HORIZONTAL = {"Top", "Bottom"}

        all_snos = area_df["s_no"].dropna().astype(str).tolist()
        used_sides = {}

        for idx, row in corner_df.dropna(how="all").iterrows():
            w1 = str(row.get("window_1", "")).strip()
            w2 = str(row.get("window_2", "")).strip()
            dir1 = str(row.get("direction_on_window_1", "")).strip()
            jt = str(row.get("joint_type", "")).strip()

            # Both windows must exist
            if w1 not in all_snos:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w1}' does not exist.")
                return False
            if w2 not in all_snos:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w2}' does not exist.")
                return False

            # Window 1 and 2 cannot be the same
            if w1 == w2:
                st.warning(f"Corner Joint Row {idx + 1}: Window 1 and Window 2 cannot be the same window.")
                return False

            # Direction must be filled
            if not dir1 or dir1 in ("", "nan", "None"):
                st.warning(f"Corner Joint Row {idx + 1}: Direction on Window 1 is required.")
                return False

            # Joint type must be filled
            if not jt or jt in ("", "nan", "None"):
                st.warning(f"Corner Joint Row {idx + 1}: Joint Type is required.")
                return False

            w1_row = area_df[area_df["s_no"].astype(str) == w1].iloc[0]
            w2_row = area_df[area_df["s_no"].astype(str) == w2].iloc[0]
            w1_orient = str(w1_row.get("orientation", "")).strip()

            # Direction valid for window 1 orientation only
            if w1_orient == "Vertical" and dir1 not in VALID_CORNER_VERTICAL:
                st.warning(f"Corner Joint Row {idx + 1}: Direction '{dir1}' not valid for Vertical Window 1. Use Left or Right.")
                return False
            if w1_orient == "Horizontal" and dir1 not in VALID_CORNER_HORIZONTAL:
                st.warning(f"Corner Joint Row {idx + 1}: Direction '{dir1}' not valid for Horizontal Window 1. Use Top or Bottom.")
                return False

            # Same single_division_length check: Top/Bottom joint → compare widths, Left/Right joint → compare heights
            if dir1 in VALID_CORNER_HORIZONTAL:
                w1_sdl = w1_row.get("width")
                w2_sdl = w2_row.get("width")
            else:
                w1_sdl = w1_row.get("height")
                w2_sdl = w2_row.get("height")
            if jt != "COTTAL CORNER 1 VTJ-21/24" and int(w1_sdl) != int(w2_sdl):
                st.warning(f"Corner Joint Row {idx + 1}: Windows '{w1}' and '{w2}' have different lengths at the corner: ({w1_sdl}mm vs {w2_sdl}mm).")
                return False

            # No duplicate sides per window
            dir2 = INVERSE_SIDE[dir1]

            if w1 not in used_sides:
                used_sides[w1] = set()
            if w2 not in used_sides:
                used_sides[w2] = set()

            if dir1 in used_sides[w1]:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w1}' already has a corner on '{dir1}'.")
                return False
            if dir2 in used_sides[w2]:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w2}' already has a corner on '{dir2}'.")
                return False

            used_sides[w1].add(dir1)
            used_sides[w2].add(dir2)

            # qty_areas must be 1 for both
            if int(w1_row.get("qty_areas", 1)) != 1:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w1}' cannot have Similar Areas > 1.")
                return False
            if int(w2_row.get("qty_areas", 1)) != 1:
                st.warning(f"Corner Joint Row {idx + 1}: Window '{w2}' cannot have Similar Areas > 1.")
                return False

        return True

    def validate_input(row, idx):
        return True

    def get_validator(df, corner_df=None):
        if corner_df is not None and len(corner_df.dropna(how="all")) > 0:
            corner_valid = CottalCalculator.validate_corner_joints(corner_df, df)
            if not corner_valid:
                return lambda row, idx: False

        def validator(row, idx):
            return CottalCalculator.validate_input(row, idx)

        return validator

    def get_data_input():

        empty_df = pd.DataFrame({
            "s_no":        pd.Series(dtype="str"),
            "area_name":   pd.Series(dtype="str"),
            "height":      pd.Series(dtype="int"),
            "width":       pd.Series(dtype="int"),
            "orientation": pd.Series(dtype="str"),
            "qty_areas":   pd.Series(dtype="int"),
            "cut_summary": pd.Series(dtype="str"),
        })

        required_cols = [
            "width", "height", "orientation", "qty_areas", "cut_summary",
        ]

        input_data = st.data_editor(
            data=empty_df,
            column_config={
                "s_no":        st.column_config.TextColumn("S. No",         required=False),
                "area_name":   st.column_config.TextColumn("Area Name",     required=False),
                "width":       st.column_config.NumberColumn("Width (mm)",   min_value=1, step=1, required=True),
                "height":      st.column_config.NumberColumn("Height (mm)",  min_value=1, step=1, required=True),
                "orientation": st.column_config.SelectboxColumn("Orientation", options=["Horizontal", "Vertical"], required=True),
                "qty_areas":   st.column_config.NumberColumn("Similar Areas", min_value=1, step=1, required=True),
                "cut_summary": st.column_config.TextColumn("Cut Summary",   required=True),
            },
            num_rows="dynamic",
        )

        # Corner joint reference images
        st.markdown("")
        st.markdown("")

        st.markdown("**Corner Joints** (optional)")
        corner_df = pd.DataFrame({
            "window_1":              pd.Series(dtype="str"),
            "window_2":              pd.Series(dtype="str"),
            "direction_on_window_1": pd.Series(dtype="str"),
            "joint_type":            pd.Series(dtype="str"),
        })
        corner_input = st.data_editor(
            data=corner_df,
            column_config={
                "window_1": st.column_config.TextColumn("Window 1 S.No"),
                "window_2": st.column_config.TextColumn("Window 2 S.No"),
                "direction_on_window_1": st.column_config.SelectboxColumn(
                    "Direction on Window 1",
                    options=["Top", "Bottom", "Left", "Right"]
                ),
                "joint_type": st.column_config.SelectboxColumn(
                    "Joint Type",
                    options=CORNER_JOINT_TYPES,
                ),
            },
            num_rows="dynamic",
        )

        st.markdown("")
        st.markdown("")

        st.markdown("**Corner Joint Types Reference:**")

        base_dir = Path(__file__).parent.parent.parent
        img_path_base = base_dir / "cottal_corner_references"

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.image(str(img_path_base) + "/corner1.png", width=150)
            # st.caption("COTTAL CORNER 1 VTJ-21/24")
        with col2:
            st.image(str(img_path_base) + "/corner2.png", width=150)
            # st.caption("COTTAL CORNER 2 VTJ-3A/24")
        with col3:
            st.image(str(img_path_base) + "/corner3.png", width=225)
            # st.caption("COTTAL CORNER 3 VTJ-3B/24")
        with col4:
            st.image(str(img_path_base) + "/corner4.png", width=300)
            # st.caption("COTTAL CORNER 4 VTJ-20/24")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            # st.image(str(img_path_base) + "/corner1.png", width=150)
            st.caption("COTTAL CORNER 2 VTJ-3A/24")

        with col2:
            # st.image(str(img_path_base) + "/corner2.png", width=150)
            st.caption("COTTAL CORNER 3 VTJ-3B/24")
        with col3:
            # st.image(str(img_path_base) + "/corner3.png", width=225)
            st.caption("COTTAL CORNER 4 VTJ-20/24")
        with col4:
            # st.image(str(img_path_base) + "/corner4.png", width=350)
            st.caption("COTTAL CORNER 1 VTJ-21/24")

        return input_data, required_cols, corner_input

    def generate_offer_df(data):
        return generate_offer_df(data)

    def generate_image(row, common_vars):
        pitch = common_vars["pitch"]
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

        config = {
            "show_carriers": True,
            "show_endcaps": False,
            "show_joints": True,
            "bar_color": "#888",
            "info_lines": info_lines,
        }
        return config

    def run(self):
        data = self.areas.copy()
        corner_joints = self.vars.get("corner_joints")

        # PROFILE
        data["req_plan"] = data.apply(profile_utils.build_req_plan, axis=1)
        # master_req = profile_utils.combine_req_plan(data["req_plan"])
        out = profile_utils.optimize_stock_v2(data)
        data["cut_plan"] = data.apply(
            lambda row: profile_utils.build_window_cut_plan(row, out["bars"]),
            axis=1
        )

        # CARRIER
        data["carrier_distances"] = data.apply(
            lambda row: profile_utils.calculate_carrier_distances(row["cut_summary"]),
            axis=1
        )
        data["total_carrier_divisions"] = data["carrier_distances"].apply(
            lambda x: sum(len(sublist) for sublist in x)
        )
        data["total_carrier_length"] = data.apply(
            lambda row: row["total_carrier_divisions"] * row["perpendicular_length"],
            axis=1
        )

        # ACCS
        data["epdm_gasket_length"] = data.apply(
            lambda row: int(math.ceil(row["total_product_length"])),
            axis=1
        )

        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_df = generate_inventory_df(
            data, self.pipe_grade, self.louver_size, out["stock_plan"], corner_joints
        )

        results = [
            data,
            offer_df_cols,
            offer_df,
            inventory_df,
        ]

        return results
