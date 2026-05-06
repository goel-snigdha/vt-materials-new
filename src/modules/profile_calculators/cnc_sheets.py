import streamlit as st
import pandas as pd

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
    }
    offer_df = data[offer_df_cols.keys()].copy()

    return offer_df_cols, offer_df


def generate_inventory_df(data):

    return pd.DataFrame()


class CNCSheetCalculator:

    def __init__(self, vars):

        self.vars = vars
        self.project_title = vars["project_title"]
        self.areas = vars["areas"]

    def validate_input(row, idx):

        return True

    def get_data_input(**kwargs):

        empty_df = pd.DataFrame(
            {
                "s_no": pd.Series(dtype="str"),
                "area_name": pd.Series(dtype="str"),
                "height": pd.Series(dtype="int"),
                "width": pd.Series(dtype="int"),
                "qty_areas": pd.Series(dtype="int"),
            }
        )

        required_cols = [
            "width",
            "height",
            "qty_areas",
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
                "qty_areas": st.column_config.NumberColumn(
                    "Similar Areas", min_value=1, step=1, required=True
                ),
            },
            num_rows="dynamic",
        )

        return input_data, required_cols

    def generate_image(row, common_vars):

        return {}

    def run(self):

        data = self.areas.copy()

        offer_df_cols, offer_df = generate_offer_df(data)
        inventory_df = generate_inventory_df(data)

        results = [data, offer_df_cols, offer_df, inventory_df]

        return results
