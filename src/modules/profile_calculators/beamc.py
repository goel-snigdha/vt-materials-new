import math
import pandas as pd
import streamlit as st
from modules.excel_utils import INV_COLUMNS, PIPE_MAPPER, COMMON_ACCESSORIES

# standard sheet dimesions are 8ft x 4ft
STANDARD_BEAM_C_LENGTH = 2400
STANDARD_BEAM_C_WIDTH = 1200
STANDARD_PROFILE_LENGTH = 2400


def generate_offer_df(data):

    offer_df_cols = {
        "s_no": {"type": "desc",    "hide_if_zero": False},
        "area_name": {"type": "desc",    "hide_if_zero": False},
        "width": {"type": "desc",    "hide_if_zero": False},
        "length": {"type": "desc",    "hide_if_zero": False},
        "qty_areas": {"type": "desc",    "hide_if_zero": False},
        "total_length_m": {"type": "formula", "hide_if_zero": False},
    }

    offer_df = data[offer_df_cols.keys()].copy()
    return offer_df_cols, offer_df


def generate_inventory_df(data, pipe_grade):

    sheet_code, sheet_desc = "BC-SHT-CS", "BEAM C CHANNEL SHEET CUT TO SIZE"
    all_inventory_rows = []
    items = []

    for _, row in data.iterrows():
        tl = row["total_length"]
        items += [
            {
                "Product Code": "BC-CHN-01",
                "Product Name": "BEAM C CHANNEL",
                "Length": 2500,
                "Quantity": int(row["profile_qty"]),
                "UOM": "m",
            },
            {
                "Product Code": sheet_code,
                "Product Name": sheet_desc,
                "Quantity": int(row["divisions_per_sheet"] * row["no_sheets"]),
                "UOM": "pcs",
                "Remarks": f"{int(row['buffer_width'])} x {STANDARD_BEAM_C_LENGTH} mm each",
            },
            {
                "Product Code": COMMON_ACCESSORIES["EPDM_GASKET"][0],
                "Product Name": COMMON_ACCESSORIES["EPDM_GASKET"][1],
                "Quantity": int(math.ceil((tl * 2) / 1000)),
                "UOM": "m",
            },
            {
                "Product Code": PIPE_MAPPER[pipe_grade][0],
                "Product Name": PIPE_MAPPER[pipe_grade][1],
                "Length": 3650,
                "Quantity": int(math.ceil((tl * 2) / 3650)),
                "UOM": "m",
            },
            {
                "Product Code": COMMON_ACCESSORIES["RIVET_6MM"][0],
                "Product Name": COMMON_ACCESSORIES["RIVET_6MM"][1],
                "Quantity": int(math.ceil(tl / 300)),
                "UOM": "pcs",
            },
            {
                "Product Code": "AC-GN-SB",
                "Product Name": "SILICON BOTTLE BLACK",
                "Quantity": int(math.ceil(tl / 3000)),
                "UOM": "pcs",
            },
        ]

        if pipe_grade == "50x25":
            items += [
                {
                    "Product Code": COMMON_ACCESSORIES["FULL_THREADED_75MM"][0],
                    "Product Name": COMMON_ACCESSORIES["FULL_THREADED_75MM"][1],
                    "Quantity": int(math.ceil(tl / 300)),
                    "UOM": "pcs",
                },
                {
                    "Product Code": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][0],
                    "Product Name": COMMON_ACCESSORIES["PVC_GITTY_50X10MM"][1],
                    "Quantity": int(math.ceil(tl / 300)),
                    "UOM": "pcs",
                },
            ]
        elif pipe_grade == "25x12":
            items.append({
                "Product Code": COMMON_ACCESSORIES["SELF_DRILLING_25MM"][0],
                "Product Name": COMMON_ACCESSORIES["SELF_DRILLING_25MM"][1],
                "Quantity": int(math.ceil(tl / 300)),
                "UOM": "pcs",
            })

        items.append({"Product Name": "SELECT L-ANGLES", "UOM": "m"})
        all_inventory_rows.extend(items)

    return pd.DataFrame(all_inventory_rows).reindex(columns=INV_COLUMNS).fillna("")


class BeamCCalculator:

    def __init__(self, vars):
        self.vars = vars
        self.project_title = vars["project_title"]
        self.areas = vars["areas"]
        self.pipe_grade = vars["pipe_grade"]

    def get_data_input(**kwargs):

        empty_df = pd.DataFrame({
            "s_no":      pd.Series(dtype="str"),
            "area_name": pd.Series(dtype="str"),
            "width":     pd.Series(dtype="int"),
            "length":    pd.Series(dtype="int"),
            "qty_areas": pd.Series(dtype="int"),
        })

        required_cols = [
            "width", "length", "qty_areas",
        ]

        input_data = st.data_editor(
            data=empty_df,
            column_config={
                "s_no":      st.column_config.TextColumn("S. No",         required=False),
                "area_name": st.column_config.TextColumn("Area Name",     required=False),
                "width":     st.column_config.NumberColumn("Width (mm)",  min_value=1, step=1, required=True),
                "length":    st.column_config.NumberColumn("Length (mm)", min_value=1, step=1, required=True),
                "qty_areas": st.column_config.NumberColumn("Similar Areas", min_value=1, step=1, required=True),
            },
            num_rows="dynamic",
        )

        return input_data, required_cols

    def run(self):

        data = self.areas.copy()

        data["buffer_width"] = data["width"] + 50
        data["divisions_per_sheet"] = data["buffer_width"].apply(
            lambda bw: 1 if bw > STANDARD_BEAM_C_WIDTH else STANDARD_BEAM_C_WIDTH // bw
        )
        data["length_per_sheet"] = STANDARD_BEAM_C_LENGTH * data["divisions_per_sheet"]
        data["total_length"] = data["length"] * data["qty_areas"]
        data["total_length_m"] = data["length"] * data["qty_areas"] / 1000
        data["no_sheets"] = data.apply(lambda r: math.ceil(r["total_length"] / r["length_per_sheet"]), axis=1)
        data["profile_qty"] = data["total_length"].apply(lambda tl: math.ceil((tl * 2) / STANDARD_PROFILE_LENGTH))

        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_out = generate_inventory_df(data, self.pipe_grade)

        return [data, offer_df_cols, offer_df, inventory_out]
