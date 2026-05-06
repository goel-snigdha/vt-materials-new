import math
import pandas as pd

import streamlit as st

STANDARD_PITCH = {
    "130 mm": 135,
    "85 mm": 90,
    "230 mm": 235,
    "Fluted": 148,
    "54.5x31.3": 43.7,
    "84.2x31.3": 57.5,  # TODO
}
MAX_STOCK_LENGTH = 4550
# MIN_MAX_PITCH = {"Rectangular Louvers": (50, 200), "C-Louvers": (80, 100)}


def get_num_windows():
    num_windows = st.number_input("Window Count:", min_value=1, value=1, step=1)
    return num_windows


def get_pitch(product, kwargs):
    if product in ["Rectangular Louvers", "C-Louvers"]:
        pitch = st.number_input(
            "Pitch:",
            value=100,
        )
    elif product in ["Cottal", "Fluted", "S-Louvers"]:
        pitch_lookup = ""
        if product in ["S-Louvers", "Cottal"]:
            pitch_lookup = kwargs["louver_size"]
        else:
            pitch_lookup = product
        pitch = STANDARD_PITCH[pitch_lookup]
    else:
        pitch = st.number_input("Pitch:", min_value=0, max_value=1000, value=50)

    return pitch


def get_dimensions(dimension, window, unit):

    if unit == "ft":
        col1, col2 = st.columns([1, 1])
        with col1:
            feet = st.number_input(
                f"{dimension} (ft)",
                min_value=0,
                value=0,
                key=f"{dimension}_ft_{window}",
            )
        with col2:
            inches = st.number_input(
                "in", min_value=0, max_value=11, value=0, key=f"{dimension}_in_{window}"
            )
        dimension = math.ceil((((feet * 12) + inches) * 25.4) / 5) * 5
    else:
        dimension = st.number_input(
            f"{dimension} (mm)", min_value=0, value=0, key=f"{dimension}_mm_{window}"
        )
        dimension = math.ceil(dimension / 5) * 5

    if dimension > 0:
        return dimension


def get_params(window, prev):
    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
    with col1:
        window_title_key = f"window_title_{window}"
        window_title = st.text_input(
            "Area Name",
            value=prev["window_title"] if "window_title" in prev else "",
            key=window_title_key,
        )
        prev["window_title"] = window_title

    with col2:
        s_no_key = f"s_no_{window}"
        s_no = st.text_input(
            "S. No", value=prev["s_no"] if "s_no" in prev else "", key=s_no_key
        )
        prev["s_no"] = s_no

    with col3:
        qty_windows_key = f"qty_windows_{window}"
        qty_windows = st.number_input(
            "Similar Windows", min_value=1, value=1, key=qty_windows_key
        )

    with col4:
        orts = ["Horizontal", "Vertical"]
        orientation_key = f"orientation_{window}"
        orientation = st.selectbox(
            "Orientation",
            orts,
            index=prev["orientation"] if "orientation" in prev else 0,
            key=orientation_key,
        )
        prev["orientation"] = orts.index(orientation)

    col5, col6, col7 = st.columns([2, 2, 1])

    with col7:
        units = ["mm", "ft"]
        unit = st.selectbox(
            "Unit",
            units,
            index=prev["unit_idx"] if "unit_idx" in prev else 0,
            key=f"unit_{window}",
        )
        prev["unit_idx"] = units.index(unit)

    with col5:
        height = get_dimensions("Height", window, unit)

    with col6:
        width = get_dimensions("Width", window, unit)

    if width is None or height is None:
        st.warning("Please enter non-zero dimensions.")
        st.stop()  # Stop further execution
    else:
        st.write(f"Window Dimensions: {width}x{height} mm")

    if orientation == "Vertical":
        width, height = height, width

    vars = {
        "window_title": window_title,
        "s_no": s_no,
        "qty_windows": qty_windows,
        "orientation": orientation,
        "width": width,
        "height": height,
        "window": window,
    }

    return vars, prev


def get_params_beamc(window, prev):

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        window_title_key = f"window_title_{window}"
        window_title = st.text_input(
            "Area Name",
            value=prev["window_title"] if "window_title" in prev else "",
            key=window_title_key,
        )
        prev["window_title"] = window_title

    with col2:
        s_no_key = f"s_no_{window}"
        s_no = st.text_input(
            "S. No", value=prev["s_no"] if "s_no" in prev else "", key=s_no_key
        )
        prev["s_no"] = s_no

    with col3:
        qty_windows_key = f"qty_windows_{window}"
        qty_windows = st.number_input(
            "Similar Windows", min_value=1, value=1, key=qty_windows_key
        )

    col1, col2, col3 = st.columns([2, 2, 1])

    with col3:
        units = ["mm", "ft"]
        unit = st.selectbox("Unit", units, key=f"unit_{window}")

    with col1:
        width = get_dimensions("Width", window, unit)

    with col2:
        length = get_dimensions("Length", window, unit)

    if width is None or length is None:
        st.warning("Please enter non-zero dimensions.")
        st.stop()  # Stop further execution
    else:
        st.write(f"Beam C Length: {length/1000} mm")

    vars = {
        "window_title": window_title,
        "s_no": s_no,
        "qty_windows": qty_windows,
        "width": width,
        "length": length,
        "window": window,
        "qty_windows": qty_windows,
    }

    return vars


def arrow_safe(value):

    if isinstance(value, dict):
        return {str(k): arrow_safe(v) for k, v in value.items()}
    if isinstance(value, list):
        return [arrow_safe(v) for v in value]
    return value


def parse_cuts(df, division_length_col=None):
    """
    Validates cut_summary for each row.
    Expected format: 1500,1500
    Returns list of cut arrays per row if valid, else False.
    If division_length_col is provided, reads expected length from that column
    instead of deriving it from orientation + width/height.
    """

    all_rows_cuts = []

    for idx, row in df.iterrows():

        orientation = row["orientation"]
        width = row["width"]
        height = row["height"]
        cut_summary = row["cut_summary"]

        # 🚨 cut_summary must not be empty
        if not cut_summary or str(cut_summary).strip() == "":
            st.warning(f"Row {idx + 1}: Cut Summary cannot be empty.")
            return False

        # Determine required division length
        if division_length_col and division_length_col in row.index:
            division_length = int(row[division_length_col])
        else:
            division_length = width if orientation == "Horizontal" else height

        # Remove spaces
        cut_summary = str(cut_summary).replace(" ", "")

        pieces = cut_summary.split(",")

        row_cuts = []
        total_used = 0

        for piece in pieces:
            try:
                piece_length = int(piece)
            except ValueError:
                st.warning(f"Row {idx + 1}: Invalid piece length '{piece}'.")
                st.stop()

            if piece_length <= 0:
                st.warning(f"Row {idx + 1}: Piece length must be greater than 0.")
                st.stop()

            if piece_length > MAX_STOCK_LENGTH:
                st.warning(
                    f"Row {idx + 1}: Piece length {piece_length}mm exceeds maximum "
                    f"stock length of {MAX_STOCK_LENGTH}mm. Please revise your cut summary."
                )
                st.stop()

            row_cuts.append(piece_length)
            total_used += piece_length

        # Validate total equals required division length
        if total_used != division_length:
            st.warning(
                f"Row {idx + 1}: Total cut length ({total_used}) "
                f"does not equal required division length ({division_length})."
            )
            st.stop()

        all_rows_cuts.append(row_cuts)

    return True, all_rows_cuts


# def parse_cut_input(df):
#     """
#     Parses cut input into structured format.
#     Performs only format validation.
#     Returns:
#         valid (bool),
#         parsed_data (list)
#     """

#     parsed_data = []

#     for idx, row in df.iterrows():

#         orientation = row["orientation"]
#         width = int(row["width"])
#         height = int(row["height"])
#         cut_summary = row["cut_summary"]
#         qty = int(row.get("qty", 1))

#         if not cut_summary or str(cut_summary).strip() == "":
#             st.warning(f"Row {idx + 1}: Cut Summary cannot be empty.")
#             return False, []

#         division_length = width if orientation == "Horizontal" else height

#         cut_summary = str(cut_summary).replace(" ", "")
#         raw_pieces = cut_summary.split(",")

#         pieces = []

#         for piece in raw_pieces:

#             if "/" not in piece:
#                 st.warning(
#                     f"Row {idx + 1}: Invalid format '{piece}'. Expected 'piece/used'."
#                 )
#                 return False, []

#             stock_str, used_str = piece.split("/")

#             try:
#                 stock_length = int(stock_str)
#             except ValueError:
#                 st.warning(
#                     f"Row {idx + 1}: Invalid stock length '{stock_str}'."
#                 )
#                 return False, []

#             cuts = []
#             used_parts = used_str.split("+")

#             for part in used_parts:
#                 try:
#                     cut_length = int(part)
#                 except ValueError:
#                     st.warning(
#                         f"Row {idx + 1}: Invalid cut length '{part}'."
#                     )
#                     return False, []

#                 if cut_length > stock_length:
#                     st.warning(
#                         f"Row {idx + 1}: Cut {cut_length} "
#                         f"cannot exceed stock length {stock_length}."
#                     )
#                     return False, []

#                 cuts.append(cut_length)

#             pieces.append({
#                 "stock_length": stock_length,
#                 "cuts": cuts
#             })

#         parsed_data.append({
#             "window_id": idx,
#             "qty": qty,
#             "division_length": division_length,
#             "pieces": pieces
#         })

#     return True, parsed_data


def validate_cut_logic(parsed_data):
    """
    Validates business rules on parsed data.
    Returns:
        valid (bool)
    """

    for window in parsed_data:

        division_length = window["division_length"]

        total_used_per_window = 0

        for piece in window["pieces"]:
            total_used_per_window += sum(piece["cuts"])

        # Multiply by qty (because pieces repeat physically)
        if total_used_per_window != division_length:
            st.warning(
                f"Window {window['window_id'] + 1}: "
                f"Total used ({total_used_per_window}) "
                f"does not equal required division length ({division_length})."
            )
            return False

    return True


def validate_required_fields(df, required_cols, extra_validator=None):

    df = df.dropna(how="all")

    if len(df) == 0:
        st.warning("No data found. Please add at least one row.")
        return False

    for idx, row in df.iterrows():
        for col in required_cols:
            value = row[col]
            if pd.isna(value):
                st.warning(f"Row {idx + 1}: '{col}' is required.")
                return False
            if isinstance(value, str) and value.strip() == "":
                st.warning(f"Row {idx + 1}: '{col}' cannot be empty.")
                return False

        if extra_validator:
            result = extra_validator(row, idx)
            if result is not True:
                return False

    return True
