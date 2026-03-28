import math

import pandas as pd
import streamlit as st

import modules.profile_utils as profile_utils
from modules.excel_utils import INV_COLUMNS, COMMON_ACCESSORIES

L_ANGLE_LENGTH = 3050

AEROFOIL_SECTION_MAPPER = {
    "AF60": ("AF-0060-PR-02", "AEROFOIL AF60 PROFILE"),
    "AF100": ("AF-0100-PR-01", "AEROFOIL AF100 PROFILE"),
    "AF150": ("AF-0150-PR-01", "AEROFOIL AF150 PROFILE"),
    "AF200": ("AF-0200-PR-04", "AEROFOIL AF200 PROFILE"),
    "AF250": ("AF-0250-PR-09", "AEROFOIL AF250 PROFILE"),
}

AEROFOIL_ENDCAPS = {
    "AF60": ("AC-AF-0060-FE", "AEROFOIL 60 ENDCAP"),
    "AF100": ("AC-AF-0100-FE", "AEROFOIL 100 ENDCAP"),
    "AF150": ("AC-AF-0150-FE", "AEROFOIL 150 ENDCAP"),
    "AF200": ("AC-AF-0200-FE", "AEROFOIL 200 ENDCAP"),
    "AF250": ("AC-AF-0250-FE", "AEROFOIL 250 ENDCAP"),
}

AEROFOIL_FRINGE_ENDCAPS = {
    "AF60": ("AC-AF-0060-GE", "AEROFOIL 60 FRINGE ENDCAP"),
    "AF100": ("AC-AF-0100-GE", "AEROFOIL 100 FRINGE ENDCAP"),
    "AF150": ("AC-AF-0150-GE", "AEROFOIL 150 FRINGE ENDCAP"),
    "AF200": ("AC-AF-0200-GE", "AEROFOIL 200 FRINGE ENDCAP"),
    "AF250": ("AC-AF-0250-GE", "AEROFOIL 250 FRINGE ENDCAP"),
}

AEROFOIL_MOVEABLE_ENDCAPS = {
    "AF60": ("AC-AF-0060-ME", "AEROFOIL 60 MOVEABLE ENDCAP"),
    "AF100": ("AC-AF-0100-ME", "AEROFOIL 100 MOVEABLE ENDCAP"),
    "AF150": ("AC-AF-0150-ME", "AEROFOIL 150 MOVEABLE ENDCAP"),
    "AF200": ("AC-AF-0200-ME", "AEROFOIL 200 MOVEABLE ENDCAP"),
}

AEROFOIL_TOP_PIVOTS = {
    "AF60": ("AC-AF-6010-PT", "AEROFOIL 60/100 PIVOT TOP"),
    "AF100": ("AC-AF-6010-PT", "AEROFOIL 60/100 PIVOT TOP"),
    "AF150": ("AC-AF-1520-PT", "AEROFOIL 150/200 PIVOT TOP"),
    "AF200": ("AC-AF-1520-PT", "AEROFOIL 150/200 PIVOT TOP"),
}

AEROFOIL_BOTTOM_PIVOTS = {
    "AF60": ("AC-AF-6010-PB", "AEROFOIL 60/100 PIVOT BOTTOM"),
    "AF100": ("AC-AF-6010-PB", "AEROFOIL 60/100 PIVOT BOTTOM"),
    "AF150": ("AC-AF-1520-PB", "AEROFOIL 150/200 PIVOT BOTTOM"),
    "AF200": ("AC-AF-1520-PB", "AEROFOIL 150/200 PIVOT BOTTOM"),
}

AEROFOIL_PACKING_WASHERS = {
    "AF60": ("AC-AF-6010-WS", "AEROFOIL 60/100 PACKING WASHER"),
    "AF100": ("AC-AF-6010-WS", "AEROFOIL 60/100 PACKING WASHER"),
    "AF150": ("AC-AF-1520-WS", "AEROFOIL 150/200 PACKING WASHER"),
    "AF200": ("AC-AF-1520-WS", "AEROFOIL 150/200 PACKING WASHER"),
}

AEROFOIL_PIVOT_T_WASHERS = {
    "AF60": ("AC-AF-6010-TW", "AEROFOIL 60/100 PIVOT T-WASHER"),
    "AF100": ("AC-AF-6010-TW", "AEROFOIL 60/100 PIVOT T-WASHER"),
    "AF150": ("AC-AF-1520-TW", "AEROFOIL 150/200 PIVOT T-WASHER"),
    "AF200": ("AC-AF-1520-TW", "AEROFOIL 150/200 PIVOT T-WASHER"),
}

C_PLATE_CODES = {
    50: ("AC-AF-50-CH", "AEROFOIL C-CHANNEL 50MM"),
    75: ("AC-AF-75-CH", "AEROFOIL C-CHANNEL 75MM"),
    100: ("AC-AF-100-CH", "AEROFOIL C-CHANNEL 100MM"),
    112: ("AC-AF-112-CH", "AEROFOIL C-CHANNEL 112MM"),
}

# Additional aerofoil-specific accessories
AEROFOIL_ACCESSORIES = {
    "AEROFOIL_MOVEABLE_U_CHANNEL": ("AF-GNRL-UC-00", "AEROFOIL MOVEABLE U-CHANNEL"),
    "AEROFOIL_MOVEABLE_KNOB": ("AC-AF-KNOB-GN", "AEROFOIL MOVEABLE KNOB"),
    "AEROFOIL_LINKING_T_WASHER": ("AC-AF-GNRL-LT", "AEROFOIL LINKING T-WASHER"),
    "AEROFOIL_LINKING_WASHER": ("AC-AF-GNRL-LW", "AEROFOIL LINKING WASHER"),
    "AEROFOIL_D_BRACKET": ("AF-0000-DB-09", "AEROFOIL D-BRACKET"),
    "AEROFOIL_AF100_V_BRACKET": ("AF-0100-VB-09", "AEROFOIL AF100 V-BRACKET"),
    "AEROFOIL_AF150_V_BRACKET": ("AF-GNRL-VB-09", "AEROFOIL AF150 V-BRACKET"),
    "HEX_NYLON_NUT_4MM": ("AC-NT-HN04", "HEX NYLON NUT 4MM"),
    "CSK_PHILIPS_SCREW_4X20MM": ("AC-SC-CP20", "CSK PHILIPS SCREW 4x20MM"),
    "C_CLAMP": ("AC-GN-CC", "C-CLAMP"),
}


class AerofoilCalculator:

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
        self.af_type = vars["af_type"]
        self.installation = vars["installation"]
        self.fixing_method = (
            vars["fixing_method"] if self.installation == "Fixed" else None
        )

        if self.installation == "Moveable (Manual)":
            af_width = int(self.af_type.strip("AF"))
            ends_gap = (af_width / 2) + 10
            centre_width = self.height - (ends_gap * 2)
        else:
            centre_width = self.height
        self.divisions = profile_utils.calculate_divisions(centre_width, self.pitch)

        self.section_id_mapper = {
            "AF60": "AEROFOIL AF60 PROFILE",
            "AF100": "AEROFOIL AF100 PROFILE",
            "AF150": "AEROFOIL AF150 PROFILE",
            "AF200": "AEROFOIL AF200 PROFILE",
            "AF250": "AEROFOIL AF250 PROFILE",
        }

        self.endcaps = {
            "AF60": "AEROFOIL 60 ENDCAP",
            "AF100": "AEROFOIL 100 ENDCAP",
            "AF150": "AEROFOIL 150 ENDCAP",
            "AF200": "AEROFOIL 200 ENDCAP",
            "AF250": "AEROFOIL 250 ENDCAP",
        }

    def generate_output_df(self):
        return pd.DataFrame(
            {
                "Project Title": [self.project_title],
                "Area (ft2)": [round(((self.width * self.height) / (304.8**2)), 2)],
                "Window": [self.window + 1],
                "Area Name": [self.window_title],
                "S. No": [self.s_no],
                "Width (mm)": [
                    self.width if self.orientation == "Horizontal" else self.height
                ],
                "Height (mm)": [
                    self.height if self.orientation == "Horizontal" else self.width
                ],
                "Orientation": [self.orientation],
                "Area Qty (nos)": [self.qty_windows],
                "Pitch (mm)": [self.pitch],
                "No. of Pieces": [self.divisions],
                "Total Length (m)": [(self.divisions * self.width) / 1000],
                "Aerofoil Type": [self.af_type],
                "Installation Method": [self.installation],
            }
        )

    def run_fringe(self):

        vars, success = profile_utils.run(
            self.width,
            self.height,
            self.divisions,
            self.window,
            self.qty_windows,
            combination=False,
        )
        used_table = vars["used_table"]

        # Add aerofoil profiles
        profile_rows = []
        for item in used_table:
            product_code, product_name = AEROFOIL_SECTION_MAPPER[self.af_type]
            profile_rows.append(
                {
                    "Product Code": product_code,
                    "Product Name": product_name,
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

        # Add fringe endcaps and accessories
        fringe_endcap_code, fringe_endcap_name = AEROFOIL_FRINGE_ENDCAPS[self.af_type]

        additional_items = [
            {
                "Product Code": fringe_endcap_code,
                "Product Name": fringe_endcap_name,
                "Quantity": self.divisions * 2,
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                "Quantity": round(self.divisions * 4, 0),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                "Quantity": math.ceil(self.divisions * 4),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                "Quantity": math.ceil(self.divisions * 4),
                "UOM": "pcs",
            },
        ]

        all_rows = []
        for block in [
            profile_rows,
            additional_items,
        ]:
            all_rows.extend(block)

        inventory_out = pd.DataFrame(all_rows).reindex(columns=INV_COLUMNS).fillna("")
        inventory_out["Quantity"] = inventory_out["Quantity"] * self.qty_windows

        output_df = self.generate_output_df()
        results = pd.DataFrame(
            {
                "Fixing Method": [self.fixing_method],
                "Fringe End Caps (pcs)": [self.divisions * 2],
                "19mm Black Gypsum Screws (pcs)": [self.divisions * 4],
                "Full Threaded Screws (pcs)": [self.divisions * 4],
                "PVC Gitty (pcs)": [self.divisions * 4],
                "Used Table": [vars["used_table"]],
                "Waste Table": [vars["waste_table"]],
            }
        )
        output_df = pd.concat([output_df, results], axis=1)

        return output_df, inventory_out, success

    def run_c_channel(self):

        plate_key = f"option_{self.window}_{self.af_type}"
        if self.af_type == "AF60":
            plate_width = st.selectbox(
                "Select C-Plate Width (mm):", [50, 75], key=plate_key
            )
        elif self.af_type == "AF100":
            plate_width = st.selectbox(
                "Select C-Plate Width (mm):", [50, 75, 100, 112], key=plate_key
            )
        else:
            plate_width = st.selectbox(
                "Select C-Plate Width (mm):", [75, 100, 112], key=plate_key
            )
        plate_length = self.height * 2

        vars, success = profile_utils.run(
            self.width,
            self.height,
            self.divisions,
            self.window,
            self.qty_windows,
            combination=False,
        )
        used_table = vars["used_table"]

        # Add aerofoil profiles
        profile_rows = []
        for item in used_table:
            product_code, product_name = AEROFOIL_SECTION_MAPPER[self.af_type]
            profile_rows.append(
                {
                    "Product Code": product_code,
                    "Product Name": product_name,
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

        # Add C-channel and accessories
        c_plate_code, c_plate_name = C_PLATE_CODES[plate_width]

        additional_items = [
            {
                "Product Code": c_plate_code,
                "Product Name": c_plate_name,
                "Length": 3650,
                "Quantity": math.ceil(plate_length / 3650),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                "Quantity": round(self.divisions * 4, 0),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                "Quantity": math.ceil(plate_length / 300),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                "Quantity": math.ceil(plate_length / 300),
                "UOM": "pcs",
            },
        ]

        all_rows = []
        for block in [
            profile_rows,
            additional_items,
        ]:
            all_rows.extend(block)

        inventory_out = pd.DataFrame(all_rows).reindex(columns=INV_COLUMNS).fillna("")
        inventory_out["Quantity"] = inventory_out["Quantity"] * self.qty_windows

        output_df = self.generate_output_df()
        results = pd.DataFrame(
            {
                "Fixing Method": [self.fixing_method],
                "Total Product Length (m)": [(self.width * self.divisions) / 1000],
                "C-Plate Width": [plate_width],
                "Total C-Channel Length (m)": [plate_length / 1000],
                "19mm Black Gypsum Screws (pcs)": [self.divisions * 4],
                "Full Threaded Screws (pcs)": [math.ceil(plate_length / 300)],
                "PVC Gitty (pcs)": [math.ceil(plate_length / 300)],
                "Used Table": [vars["used_table"]],
                "Waste Table": [vars["waste_table"]],
            }
        )
        output_df = pd.concat([output_df, results], axis=1)

        return (
            output_df,
            inventory_out,
            success,
        )

    def run_ms_rod(self):
        rod_lengths = []
        top_key = f"top_{self.window}"
        top_suspension = st.number_input(
            "Top Air Suspension (mm):", min_value=0, value=100, key=top_key
        )
        if top_suspension > 0:
            rod_lengths.append(top_suspension)

        bottom_key = f"bottom_{self.window}"
        bottom_suspension = st.number_input(
            "Bottom Air Suspension (mm):", min_value=0, value=100, key=bottom_key
        )
        if bottom_suspension > 0:
            rod_lengths.append(bottom_suspension)

        actual_width = self.width - ((top_suspension / 2) + (bottom_suspension / 2))
        suspension_msg = (
            "Actual Aerofoil Width after 50% of top and "
            "bottom suspensions are applied: {}".format(actual_width)
        )
        st.write(suspension_msg)

        vars, success = profile_utils.run(
            actual_width,
            self.height,
            self.divisions,
            self.window,
            self.qty_windows,
            combination=False,
        )
        used_table = vars["used_table"]

        # Add aerofoil profiles
        profile_rows = []
        for item in used_table:
            product_code, product_name = AEROFOIL_SECTION_MAPPER[self.af_type]
            profile_rows.append(
                {
                    "Product Code": product_code,
                    "Product Name": product_name,
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

        all_rows = []
        for block in [
            profile_rows,
        ]:
            all_rows.extend(block)

        inventory_out = pd.DataFrame(all_rows).reindex(columns=INV_COLUMNS).fillna("")

        output_df = self.generate_output_df()
        results = pd.DataFrame(
            {
                "Fixing Method": [self.fixing_method],
                "MS Rod Lengths (mm)": [str(rod_lengths)],
                "Total MS Rods for each Length (pcs)": [self.divisions],
                "Total MS Rod Length (m)": [
                    sum([rod_length * self.divisions for rod_length in rod_lengths])
                    / 1000
                ],
                "Piece Length (mm)": [actual_width],
                "End Caps with Center Hole (pcs)": [self.divisions * 2],
                "19mm Black Gypsum Screws (pcs)": [self.divisions * 4],
                "Used Table": [vars["used_table"]],
                "Waste Table": [vars["waste_table"]],
            }
        )
        output_df = pd.concat([output_df, results], axis=1)

        return output_df, inventory_out, success

    def run_d_wall(self):

        endcap_cnt = profile_utils.calculate_endcaps(
            self.window, self.orientation, self.divisions
        )

        vars, success = profile_utils.run(
            self.width, self.height, self.divisions, self.window, self.qty_windows
        )

        length_combination = vars["length_combination"]
        total_product_length = vars["total_product_length"]
        total_carrier_divisions = vars["total_carrier_divisions"]
        used_table = vars["used_table"]
        waste_table = vars["waste_table"]
        joining_pieces = (len(length_combination) - 1) * self.divisions
        brackets = self.divisions * total_carrier_divisions

        # Add aerofoil profiles
        profile_rows = []
        for item in used_table:
            product_code, product_name = AEROFOIL_SECTION_MAPPER[self.af_type]
            profile_rows.append(
                {
                    "Product Code": product_code,
                    "Product Name": product_name,
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

        # Add brackets and accessories
        v_bracket_str = (
            "AEROFOIL AF100 V-BRACKET"
            if self.af_type == "AF100"
            else "AEROFOIL AF150 V-BRACKET"
        )
        v_bracket_code = (
            AEROFOIL_ACCESSORIES["AEROFOIL_AF100_V_BRACKET"][0]
            if self.af_type == "AF100"
            else AEROFOIL_ACCESSORIES["AEROFOIL_AF150_V_BRACKET"][0]
        )
        endcap_code, endcap_name = AEROFOIL_ENDCAPS[self.af_type]
        paint_qty = round(self.divisions / 50 * 2) / 2

        additional_items = [
            {
                "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_D_BRACKET"][0],
                "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_D_BRACKET"][1],
                "Length": 3050,
                "Quantity": math.ceil(brackets * 55 / 3050),
                "UOM": "m",
            },
            {
                "Product Code": v_bracket_code,
                "Product Name": v_bracket_str,
                "Length": 3050,
                "Quantity": math.ceil(brackets * 55 / 3050),
                "UOM": "m",
            },
            {
                "Product Code": endcap_code,
                "Product Name": endcap_name,
                "Quantity": endcap_cnt,
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_19MM"][1],
                "Quantity": brackets * 5,
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["TRUSS_HEAD_16MM"][0],
                "Product Name": COMMON_ACCESSORIES["TRUSS_HEAD_16MM"][1],
                "Quantity": brackets * 2,
                "UOM": "pcs",
            },
        ]

        if endcap_cnt > 0:
            endcap_screw = [
                {
                    "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                    "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                    "Quantity": (
                        endcap_cnt * 2
                        if self.af_type in ["AF60", "AF100"]
                        else endcap_cnt * 4
                    ),
                    "UOM": "pcs",
                }
            ]

        paint = [
            {
                "Product Code": COMMON_ACCESSORIES["PAINT"][0],
                "Product Name": COMMON_ACCESSORIES["PAINT"][1],
                "Quantity": paint_qty,
                "UOM": "l",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PAINT_BRUSH"][0],
                "Product Name": COMMON_ACCESSORIES["PAINT_BRUSH"][1],
                "Quantity": 1,
                "UOM": "pcs",
            },
        ]

        all_rows = []
        for block in [
            profile_rows,
            additional_items,
            endcap_screw,
            paint,
        ]:
            if block:
                all_rows.extend(block)

        inventory_out = pd.DataFrame(all_rows).reindex(columns=INV_COLUMNS).fillna("")
        inventory_out["Quantity"] = inventory_out["Quantity"] * self.qty_windows

        output_df = self.generate_output_df()
        results = pd.DataFrame(
            {
                "Fixing Method": [self.fixing_method],
                "Total Product Length (m)": [total_product_length],
                "Total D-Brackets (pcs)": [self.divisions * total_carrier_divisions],
                "Total D-Bracket Length (m)": [
                    (self.divisions * total_carrier_divisions * 55) / 1000
                ],
                "Total V-Brackets (pcs)": [self.divisions * total_carrier_divisions],
                "Total V-Bracket Length (m)": [
                    (self.divisions * total_carrier_divisions * 55) / 1000
                ],
                "Self-Drilling Screws (pcs)": [
                    self.divisions * total_carrier_divisions * 4
                ],
                "Rivets/Washers With Screw (pcs)": [
                    self.divisions * total_carrier_divisions
                ],
                "Joining Pieces (pcs)": [joining_pieces],
                "Total End Caps (pcs)": [endcap_cnt],
                "Used Table": [used_table],
                "Waste Table": [waste_table],
            }
        )
        output_df = pd.concat([output_df, results], axis=1)

        return output_df, inventory_out, success

    def run_manual_moveable(self):

        total_carrier_length = self.height * 2
        pcs_per_l_angle = L_ANGLE_LENGTH / self.pitch
        no_l_angle = math.ceil(self.divisions / pcs_per_l_angle)
        length_per_l_angle = round(self.height / no_l_angle, 1)

        vars, success = profile_utils.run(
            self.width,
            self.height,
            self.divisions,
            self.window,
            self.qty_windows,
            combination=False,
        )
        used_table = vars["used_table"]

        # Add aerofoil profiles
        profile_rows = []
        for item in used_table:
            product_code, product_name = AEROFOIL_SECTION_MAPPER[self.af_type]
            profile_rows.append(
                {
                    "Product Code": product_code,
                    "Product Name": product_name,
                    "Length": item,
                    "Quantity": used_table[item],
                    "UOM": "m",
                }
            )

        # Get product codes from mappers
        endcap_code, endcap_name = AEROFOIL_ENDCAPS[self.af_type]
        moveable_endcap_code, moveable_endcap_name = AEROFOIL_MOVEABLE_ENDCAPS[
            self.af_type
        ]
        top_pivot_code, top_pivot_name = AEROFOIL_TOP_PIVOTS[self.af_type]
        bottom_pivot_code, bottom_pivot_name = AEROFOIL_BOTTOM_PIVOTS[self.af_type]
        packing_washer_code, packing_washer_name = AEROFOIL_PACKING_WASHERS[
            self.af_type
        ]
        pivot_t_washer_code, pivot_t_washer_name = AEROFOIL_PIVOT_T_WASHERS[
            self.af_type
        ]

        black_gypsum = {"AF60": 4, "AF100": 4, "AF150": 8, "AF200": 8}

        additional_items = [
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
                "Quantity": int(math.ceil(total_carrier_length / 3050)),
                "UOM": "m",
                "CNC Hole Distance": self.pitch,
            },
            {
                "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_MOVEABLE_U_CHANNEL"][0],
                "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_MOVEABLE_U_CHANNEL"][1],
                "Length": 3050,
                "Quantity": int(math.ceil(self.height / 3050)),
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
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": moveable_endcap_code,
                "Product Name": moveable_endcap_name,
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": top_pivot_code,
                "Product Name": top_pivot_name,
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": bottom_pivot_code,
                "Product Name": bottom_pivot_name,
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": pivot_t_washer_code,
                "Product Name": pivot_t_washer_name,
                "Quantity": self.divisions * 2,
                "UOM": "pcs",
            },
            {
                "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_T_WASHER"][0],
                "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_T_WASHER"][1],
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_WASHER"][0],
                "Product Name": AEROFOIL_ACCESSORIES["AEROFOIL_LINKING_WASHER"][1],
                "Quantity": self.divisions * 2,
                "UOM": "pcs",
            },
            {
                "Product Code": packing_washer_code,
                "Product Name": packing_washer_name,
                "Quantity": int(math.ceil(self.divisions * 0.20)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][0],
                "Product Name": COMMON_ACCESSORIES["BLACK_GYPSUM_19MM"][1],
                "Quantity": int(
                    math.ceil(
                        (self.divisions * black_gypsum[self.af_type])
                        + (self.width / 600)
                    )
                ),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["RIVET_6MM"][0],
                "Product Name": COMMON_ACCESSORIES["RIVET_6MM"][1],
                "Quantity": int(math.ceil(self.width * 2 / 300)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                "Quantity": int(math.ceil(total_carrier_length / 500)),
                "UOM": "pcs",
            },
            {
                "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                "Quantity": int(math.ceil(total_carrier_length / 500)),
                "UOM": "pcs",
            },
            {
                "Product Code": AEROFOIL_ACCESSORIES["HEX_NYLON_NUT_4MM"][0],
                "Product Name": AEROFOIL_ACCESSORIES["HEX_NYLON_NUT_4MM"][1],
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": AEROFOIL_ACCESSORIES["CSK_PHILIPS_SCREW_4X20MM"][0],
                "Product Name": AEROFOIL_ACCESSORIES["CSK_PHILIPS_SCREW_4X20MM"][1],
                "Quantity": self.divisions,
                "UOM": "pcs",
            },
            {
                "Product Code": AEROFOIL_ACCESSORIES["C_CLAMP"][0],
                "Product Name": AEROFOIL_ACCESSORIES["C_CLAMP"][1],
                "Quantity": int(math.ceil(self.divisions * 0.1)),
                "UOM": "pcs",
            },
        ]

        all_rows = []
        for block in [
            profile_rows,
            additional_items,
        ]:
            all_rows.extend(block)

        inventory_out = pd.DataFrame(all_rows).reindex(columns=INV_COLUMNS).fillna("")
        inventory_out["Quantity"] = inventory_out["Quantity"] * self.qty_windows

        output_df = self.generate_output_df()
        results = pd.DataFrame(
            {
                "Total Carrier Length (m)": [total_carrier_length / 1000],
                "Carrier Hole Gap (pcs)": [self.pitch],
                "Top End Caps (pcs)": [self.divisions],
                "Bottom End Caps (pcs)": [self.divisions],
                "Pivot (pcs)": [self.divisions * 2],
                "Total Length of L-Angle (m)": [self.height / 1000],
                "Number of L-Angles (pcs)": [no_l_angle],
                "Length per L-Angle (mm)": [length_per_l_angle],
                "Knobs (pcs)": [no_l_angle],
                "Black Gypsum Screws (pcs)": [self.divisions * 4],
                "3/4 Inch SS Screws (pcs)": [self.divisions * 1],
                "75mm Full Threaded Screws (pcs)": [
                    round(total_carrier_length / 300, 2)
                ],
                "PVC Gitty (pcs)": [round(total_carrier_length / 300, 2)],
                "6mm Interlocking Screws (pcs)": [self.divisions],
                "Used Table": [vars["used_table"]],
                "Waste Table": [vars["waste_table"]],
            }
        )
        output_df = pd.concat([output_df, results], axis=1)

        return output_df, inventory_out, success

    def run_motorized_moveable(self):
        success = -1
        return pd.DataFrame(), success

    def run_fixed(self):
        if self.fixing_method == "Fringe End Caps":
            return self.run_fringe()
        elif self.fixing_method == "C-Channel":
            return self.run_c_channel()
        elif self.fixing_method == "MS Rod/Slot Cut Pipe":
            return self.run_ms_rod()
        elif self.fixing_method == "D-Wall Bracket":
            return self.run_d_wall()

    def run(self):
        if self.installation == "Fixed":
            return self.run_fixed()
        elif self.installation == "Moveable (Manual)":
            return self.run_manual_moveable()
        elif self.installation == "Moveable (Motorized)":
            return self.run_motorized_moveable()
