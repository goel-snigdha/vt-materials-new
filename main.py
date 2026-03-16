import numpy as np
import pandas as pd
import streamlit as st

import modules.excel_processor as to_excel
from modules.profile_calculators.aerofoil import AerofoilCalculator
from modules.profile_calculators.beamc import BeamCCalculator
from modules.profile_calculators.cottal import CottalCalculator
from modules.profile_calculators.fluted import FlutedCalculator
from modules.profile_calculators.grille import GrilleCalculator
from modules.profile_calculators.slouvers import SLouverCalculator
from modules.profile_calculators.rectangular import RectangularCalculator
from utils import (
    # get_num_windows, get_params, get_params_beamc, get_pitch, arrow_safe, parse_cuts, validate_required_fields
    get_pitch,
    parse_cuts,
    validate_cut_logic,
    validate_required_fields
)

PRODUCTS = [
    "Grille 2550",
    "Aerofoil",
    "Cottal",
    "Fluted",
    "S-Louvers",
    # 'C-Louvers',
    "Rectangular Louvers",
    "Beam C-Channel",
]
CALCULATOR_MAPPING = {
    "Grille 2550": GrilleCalculator,
    "Aerofoil": AerofoilCalculator,
    "Cottal": CottalCalculator,
    "Fluted": FlutedCalculator,
    "S-Louvers": SLouverCalculator,
    # "C-Louvers": CLouverCalculator,
    "Rectangular Louvers": RectangularCalculator,
    "Beam C-Channel": BeamCCalculator,
}
CONFIGURE_MAPPING = {

}


def handle_conversion(product, results, common_vars):

    # profile_output_obj, offer_output_obj, inv_output_obj = to_excel.convert(
    #     product, results, common_vars
    # )
    offer_output_obj, inv_output_obj, installer_output_obj = to_excel.convert(
        product, results, common_vars
    )
    # st.session_state["profile_output_obj"] = profile_output_obj
    st.session_state["offer_output_obj"] = offer_output_obj
    st.session_state["inv_output_obj"] = inv_output_obj
    st.session_state["installer_output_obj"] = installer_output_obj
    st.session_state["conversion_done"] = True


def post_process(product, results, common_vars):

    if "conversion_done" in st.session_state:
        # profile_output_obj = st.session_state["profile_output_obj"]
        offer_output_obj = st.session_state["offer_output_obj"]
        inv_output_obj = st.session_state["inv_output_obj"]
        installer_output_obj = st.session_state["installer_output_obj"]

        title = common_vars["project_title"].replace("/", "-").replace("\\", "-")

        # btn1 = st.download_button(
        #     label="📥 Download Profile Excel",
        #     data=profile_output_obj,
        #     file_name=f"Profile Sheet - {title}.xlsx",
        # )
        btn2 = st.download_button(
            label="📥 Download Offer Excel",
            data=offer_output_obj,
            file_name=f"Offer Sheet - {title}.xlsx",
        )
        btn3 = st.download_button(
            label="📥 Download Inventory Excel",
            data=inv_output_obj,
            file_name=f"Inventory Sheet - {title}.xlsx",
        )
        btn4 = st.download_button(
            label="📥 Download Installer PDFs",
            data=installer_output_obj,
            file_name=f"Installer Files.pdf",
        )
        # if btn1:
        #     st.write("Profile Excel Downloaded Successfully")
        if btn2:
            st.write("Offer Excel Downloaded Successfully")
        if btn3:
            st.write("Inventory Excel Downloaded Successfully")
        if btn4:
            st.write("Installer Sheet Downloaded Successfully")


def run(product, project_title, **kwargs):

    if product == "Beam C-Channel":
        pitch = pd.NA
    else:
        pitch = get_pitch(product, kwargs)
    output = pd.DataFrame()
    inv_data = pd.DataFrame()
    product_class = CALCULATOR_MAPPING[product]
    vars = {}

    vars["project_title"] = project_title
    vars["pitch"] = pitch
    vars["areas"], required_cols = product_class.get_data_input()
    submit = st.button("Submit")

    if submit:
        required = validate_required_fields(vars["areas"], required_cols)
        if not required:
            st.error("Some required fields are empty. Please re-submit")
            return

        valid, cut_summary = parse_cuts(vars["areas"])

        if valid:
            vars["areas"]["width"] = vars["areas"]["width"].astype(int)
            vars["areas"]["height"] = vars["areas"]["height"].astype(int)
            vars["areas"]["cut_summary"] = cut_summary
            # valid_logic = validate_cut_logic(cut_summary)

            if required and valid:
                st.badge(
                    f'Successfully submitted. Processing {len(vars["areas"])} row(s).',
                    color="yellow"
                )
            
            vars["areas"]["single_division_length"] = np.where(
                vars["areas"]["orientation"] == "Horizontal",
                vars["areas"]["width"],
                vars["areas"]["height"]
            )
            vars["areas"]["perpendicular_length"] = np.where(
                vars["areas"]["orientation"] == "Horizontal",
                vars["areas"]["height"],
                vars["areas"]["width"]
            )
            vars["areas"]["divisions"] = np.ceil(
                vars["areas"]["perpendicular_length"] / vars["pitch"]
            ).astype(int)
            vars["areas"]["total_product_length"] = (
                vars["areas"]["single_division_length"] * vars["areas"]["divisions"]
            ) / 1000
            vars["areas"]["area_sqft"] = (
                vars["areas"]["width"] * vars["areas"]["height"]
            ) / 92903.04

            calculator = product_class(vars)
            results = calculator.run()

            common_vars = {
                "product": product,
                "pitch": vars["pitch"],
                "project_title": vars["project_title"],
            }

            handle_conversion(product, results, common_vars)
            post_process(product, results, common_vars)

    st.stop()

    calculator_class = CALCULATOR_MAPPING[product]
    calculator = calculator_class(vars)
    df, inv_df, success = calculator.run()

    # for window in range(num_windows):
    #     window_expander = st.expander(f"Window {window + 1}", expanded=False)
    #     if window == 0:
    #         prev = {}
    #     with window_expander:
    #         # Only Beam C has a completely different input
    #         if product == "Beam C-Channel":
    #             vars = get_params_beamc(window, prev)
    #         else:
    #             vars, prev = get_params(window, prev)
    #         vars["project_title"] = project_title
    #         vars["pitch"] = pitch

    #         calculator_class = CALCULATOR_MAPPING[product]

    #         # Additional arguments specific to each calculator
    #         if calculator_class == AerofoilCalculator:
    #             vars["af_type"] = kwargs["af_type"]
    #             vars["installation"] = kwargs["installation"]
    #             if kwargs["installation"] == "Fixed":
    #                 vars["fixing_method"] = kwargs["fixing_method"]
    #         elif calculator_class in [
    #             FlutedCalculator,
    #             BeamCCalculator,
    #         ]:
    #             vars["pipe_grade"] = kwargs["pipe_grade"]
    #         elif calculator_class in [CottalCalculator]:
    #             vars["pipe_grade"] = kwargs["pipe_grade"]
    #             vars["louver_size"] = kwargs["louver_size"]
    #         elif calculator_class in [SLouverCalculator, RectangularCalculator]:
    #             vars["louver_size"] = kwargs["louver_size"]

    #         calculator = calculator_class(vars)
    #         df, inv_df, success = calculator.run()
    #         st.subheader("Results")

    #         display_df = df.copy()

    #         for col in display_df.columns:
    #             if display_df[col].dtype == "object":
    #                 display_df[col] = display_df[col].apply(arrow_safe)

    #         st.dataframe(display_df)

    #         output = pd.concat([output, df], axis=0)
    #         inv_data = pd.concat([inv_data, inv_df], axis=0)

    #         if success:
    #             success_cnt += 1

    # if success_cnt == num_windows:
    #     post_process(product, output, inv_data)


def main():

    st.title("Vibrant Technik Material Calculator")

    product = st.selectbox("Select a product:", sorted(PRODUCTS))
    project_title = st.text_input("Project Details", placeholder="Customer Name & City")

    # num_windows = get_num_windows()
    if product == "Grille 2550":
        run(product, project_title)
    # elif product == "Aerofoil":
    #     af_type = st.selectbox(
    #         "Aerofoil type:",
    #         [
    #             "AF60",
    #             "AF100",
    #             "AF150",
    #             "AF200",
    #             "AF250",
    #             "AF400",
    #         ],
    #     )
    #     if af_type in ["AF250", "AF400"]:
    #         installation = st.selectbox(
    #             "Installation method:",
    #             [
    #                 "Fixed",
    #                 # 'Moveable (Manual)',
    #                 # 'Moveable (Motorized)'
    #             ],
    #         )
    #     else:
    #         installation = st.selectbox(
    #             "Installation method:",
    #             [
    #                 "Fixed",
    #                 "Moveable (Manual)",
    #                 # 'Moveable (Motorized)'
    #             ],
    #         )
    #     if installation == "Fixed":
    #         fixing_method = st.selectbox(
    #             "Fixing Method:",
    #             [
    #                 "Fringe End Caps",
    #                 "C-Channel",
    #                 "MS Rod/Slot Cut Pipe",
    #                 "D-Wall Bracket",
    #             ],
    #         )
    #         run(
    #             product,
    #             project_title,
    #             num_windows,
    #             af_type=af_type,
    #             installation=installation,
    #             fixing_method=fixing_method,
    #         )
    #     else:
    #         run(
    #             product,
    #             project_title,
    #             num_windows,
    #             af_type=af_type,
    #             installation=installation,
    #         )
    # elif product == "Cottal":
    #     pipe_grade = st.selectbox("Pipe Grade:", ["50x25", "38x25"], key="pipe_cottal")
    #     louver_size = st.selectbox("Cottal Size:", ["85 mm", "130 mm", "230 mm"], key="size_cottal")
    #     run(product, project_title, num_windows, pipe_grade=pipe_grade, louver_size=louver_size)
    # elif product == "Fluted":
    #     pipe_grade = st.selectbox("Pipe Grade:", ["50x25", "38x25"], key="pipe_fluted")
    #     run(product, project_title, num_windows, pipe_grade=pipe_grade)
    # elif product == "S-Louvers":
    #     louver_size = st.selectbox(
    #         "S-Louvers Size:",
    #         [
    #             "54.5x31.3"
    #             # ,'84.2x31.3'
    #         ],
    #     )
    #     run(product, project_title, num_windows, louver_size=louver_size)
    # elif product == "C-Louvers":
    #     run(product, project_title, num_windows)
    # elif product == "Rectangular Louvers":
    #     louver_size = st.selectbox(
    #         "Rectangular Louvers Size:", ["30x60", "50x75", "50x100", "50x125"]
    #     )
    #     run(product, project_title, num_windows, louver_size=louver_size)
    # elif product == "Beam C-Channel":
    #     pipe_grade = st.selectbox("Pipe Grade:", ["50x25", "25x12"], key="pipe_beamc")
    #     run(product, project_title, num_windows, pipe_grade=pipe_grade)

st.set_page_config(layout="wide")
main()
