
import streamlit as st
import pandas as pd
from datetime import date
import re


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Asset Management System",
    layout="wide"
)

st.title("Asset Management System")


# =========================================================
# EXPECTED 24 COLUMNS
# =========================================================

EXPECTED_COLUMNS = [
    "Umoja Equipment Number",
    "Barcode",
    "Asset Number",
    "Description",
    "Serial Number (Mandatory Field)",
    "Umoja Equipment Status",
    "Umoja Equipment Status Description",
    "Equipment Status (Mandatory Field)",
    "Equipment Condition (Mandatory Field)",
    "Cost Center",
    "Acquisition date",
    "Acquisition Value",
    "Umoja Notification Number",
    "Disposal Case Number",
    "User Accountable (Mandatory Field)",
    "User Accountable Index Number (Mandatory Field)",
    "Is the Equipment Used By a UNV/Consultant/Intern",
    "UNV/Consultant/Intern Full Names",
    "Staff Member To Assign Equipment Used By UNV/Consultants/Interns",
    "Staff Member Index Number",
    "Functional Location",
    "Verification Date (Mandatory Field)",
    "Verified By (Mandatory Field)",
    "Additional Comments (Optional)"
]


# =========================================================
# HEADER CLEANING
# =========================================================

def normalize_column_name(name):

    name = str(name)

    name = name.replace("\xa0", " ")
    name = name.replace("\n", " ")
    name = name.replace("\r", " ")
    name = name.replace("\t", " ")

    name = name.strip()

    name = re.sub(
        r"\s+",
        " ",
        name
    )

    return name


def column_key(name):

    name = normalize_column_name(name)

    return re.sub(
        r"[^a-z0-9]",
        "",
        name.lower()
    )


# =========================================================
# LOAD DATA
# =========================================================

@st.cache_data
def load_data():

    data = pd.read_excel(
        "assets.xlsx",
        sheet_name="Sheet1"
    )

    data.columns = [
        normalize_column_name(col)
        for col in data.columns
    ]

    actual_columns = {
        column_key(col): col
        for col in data.columns
    }

    rename_map = {}

    for expected_column in EXPECTED_COLUMNS:

        expected_key = column_key(
            expected_column
        )

        if expected_key in actual_columns:

            actual_column = actual_columns[
                expected_key
            ]

            rename_map[
                actual_column
            ] = expected_column

    data = data.rename(
        columns=rename_map
    )

    missing_columns = [
        column
        for column in EXPECTED_COLUMNS
        if column not in data.columns
    ]

    if missing_columns:

        st.error(
            "Some expected columns could not be found in the Excel file."
        )

        st.write("Missing columns:")

        for column in missing_columns:

            st.write(
                f"- {column}"
            )

        st.write(
            "Actual Excel columns detected:"
        )

        for column in data.columns:

            st.code(
                repr(column)
            )

        st.stop()

    return data


df = load_data()


# =========================================================
# SAVE DATA
# =========================================================

def save_data(data):

    with pd.ExcelWriter(
        "assets.xlsx",
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:

        data.to_excel(
            writer,
            sheet_name="Sheet1",
            index=False
        )


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


def display_value(value):

    text = clean_text(value)

    if text == "":
        return "—"

    return text


def safe_date(value):

    if pd.isna(value):
        return date.today()

    try:

        return pd.to_datetime(
            value
        ).date()

    except:

        return date.today()


def display_date(value):

    if pd.isna(value):
        return "—"

    try:

        return pd.to_datetime(
            value
        ).strftime("%d/%m/%Y")

    except:

        return clean_text(value)


def safe_number(value):

    number = pd.to_numeric(
        value,
        errors="coerce"
    )

    if pd.isna(number):
        return 0.0

    return float(number)


# =========================================================
# SESSION STATE
# =========================================================

if "edit_mode" not in st.session_state:

    st.session_state.edit_mode = False


if "edit_index" not in st.session_state:

    st.session_state.edit_index = None


if "asset_search" not in st.session_state:

    st.session_state.asset_search = ""


# =========================================================
# CLEAR SEARCH
# =========================================================

def clear_search():

    st.session_state.asset_search = ""

    st.session_state.edit_mode = False

    st.session_state.edit_index = None


# =========================================================
# DASHBOARD
# =========================================================

st.subheader(
    "Asset Overview"
)


total_assets = len(df)


status_series = (
    df[
        "Equipment Status (Mandatory Field)"
    ]
    .astype(str)
    .str.strip()
    .str.upper()
)


in_use = (
    status_series == "IN USE"
).sum()


in_store = (
    status_series == "IN STORE"
).sum()


faulty = (
    status_series == "FAULTY"
).sum()


disposed = (
    status_series == "DISPOSED"
).sum()


lost = (
    status_series == "LOST"
).sum()


col1, col2, col3 = st.columns(3)


with col1:

    st.metric(
        "Total Assets",
        total_assets
    )


with col2:

    st.metric(
        "In Use",
        in_use
    )


with col3:

    st.metric(
        "In Store",
        in_store
    )


col4, col5, col6 = st.columns(3)


with col4:

    st.metric(
        "Faulty",
        faulty
    )


with col5:

    st.metric(
        "Disposed",
        disposed
    )


with col6:

    st.metric(
        "Lost",
        lost
    )


st.divider()


# =========================================================
# SEARCH ASSET
# =========================================================

st.subheader(
    "Search Asset"
)


search_col, clear_col = st.columns(
    [6, 1]
)


with search_col:

    search_value = st.text_input(
        "Search",
        placeholder=(
            "Enter Serial Number, Barcode, "
            "User, Status, or Description"
        ),
        key="asset_search",
        label_visibility="collapsed"
    )


with clear_col:

    st.button(
        "Clear",
        use_container_width=True,
        on_click=clear_search
    )


# =========================================================
# SEARCH RESULT
# =========================================================

if search_value:

    search_clean = (
        search_value
        .strip()
        .lower()
    )


    result = df[

        df[
            "Serial Number (Mandatory Field)"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.contains(
            search_clean,
            na=False,
            regex=False
        )

        |

        df[
            "Barcode"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.contains(
            search_clean,
            na=False,
            regex=False
        )

        |

        df[
            "User Accountable (Mandatory Field)"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.contains(
            search_clean,
            na=False,
            regex=False
        )

        |

        df[
            "Equipment Status (Mandatory Field)"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.contains(
            search_clean,
            na=False,
            regex=False
        )

        |

        df[
            "Description"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.contains(
            search_clean,
            na=False,
            regex=False
        )

    ]


    if result.empty:

        st.warning(
            "Asset not found."
        )


    else:

        serial_series = (
            df[
                "Serial Number (Mandatory Field)"
            ]
            .astype(str)
            .str.strip()
            .str.lower()
        )


        exact_serial_result = df[
            serial_series == search_clean
        ]


        # =================================================
        # EXACT SERIAL -> ASSET DETAIL CARD
        # =================================================

        if len(exact_serial_result) == 1:

            exact_index = (
                exact_serial_result
                .index[0]
            )

            exact_asset = (
                exact_serial_result
                .iloc[0]
            )


            st.success(
                "Asset found."
            )


            st.markdown(
                "### Asset Details"
            )


            # -------------------------------------------------
            # MAIN INFORMATION
            # -------------------------------------------------

            info1, info2, info3 = st.columns(3)


            with info1:

                st.markdown(
                    "**Description**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Description"
                        ]
                    )
                )


                st.markdown(
                    "**Serial Number**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Serial Number (Mandatory Field)"
                        ]
                    )
                )


                st.markdown(
                    "**Barcode**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Barcode"
                        ]
                    )
                )


            with info2:

                st.markdown(
                    "**Equipment Status**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Equipment Status (Mandatory Field)"
                        ]
                    )
                )


                st.markdown(
                    "**Equipment Condition**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Equipment Condition (Mandatory Field)"
                        ]
                    )
                )


                st.markdown(
                    "**Functional Location**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Functional Location"
                        ]
                    )
                )


            with info3:

                st.markdown(
                    "**User Accountable**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "User Accountable (Mandatory Field)"
                        ]
                    )
                )


                st.markdown(
                    "**Verification Date**"
                )

                st.write(
                    display_date(
                        exact_asset[
                            "Verification Date (Mandatory Field)"
                        ]
                    )
                )


                st.markdown(
                    "**Verified By**"
                )

                st.write(
                    display_value(
                        exact_asset[
                            "Verified By (Mandatory Field)"
                        ]
                    )
                )


            # -------------------------------------------------
            # SHOW ALL 24 FIELDS
            # -------------------------------------------------

            with st.expander(
                "Show All Details"
            ):

                detail_data = pd.DataFrame({
                    "Field": EXPECTED_COLUMNS,

                    "Value": [
                        display_value(
                            exact_asset[column]
                        )
                        for column in EXPECTED_COLUMNS
                    ]
                })


                st.dataframe(
                    detail_data,
                    use_container_width=True,
                    hide_index=True
                )


            # -------------------------------------------------
            # UPDATE BUTTON
            # -------------------------------------------------

            update_col1, update_col2 = st.columns(
                [1, 5]
            )


            with update_col1:

                if st.button(
                    "✏️ Update Asset",
                    use_container_width=True,
                    key=f"update_{exact_index}"
                ):

                    st.session_state.edit_mode = True

                    st.session_state.edit_index = (
                        exact_index
                    )

                    st.rerun()


        # =================================================
        # GENERAL SEARCH -> TABLE
        # =================================================

        else:

            st.success(
                f"{len(result)} asset(s) found."
            )


            st.dataframe(
                result,
                use_container_width=True
            )


            csv = (
                result
                .to_csv(
                    index=False
                )
                .encode(
                    "utf-8"
                )
            )


            st.download_button(
                label="Download Search Results",
                data=csv,
                file_name="asset_search_results.csv",
                mime="text/csv"
            )


st.divider()


# =========================================================
# FORM MODE
# =========================================================

if (
    st.session_state.edit_mode
    and
    st.session_state.edit_index is not None
):

    form_mode = "UPDATE"

    edit_index = (
        st.session_state.edit_index
    )

    asset = df.loc[
        edit_index
    ]


else:

    form_mode = "ADD"

    asset = None


# =========================================================
# FORM TITLE
# =========================================================

if form_mode == "UPDATE":

    st.subheader(
        "Update Asset"
    )


    st.info(
        "The selected asset has been loaded into the form."
    )


    if st.button(
        "Cancel Update"
    ):

        st.session_state.edit_mode = False

        st.session_state.edit_index = None

        st.rerun()


else:

    st.subheader(
        "Add New Asset"
    )


st.caption(
    "* Fields marked with * are mandatory."
)


# =========================================================
# DEFAULT VALUES
# =========================================================

if form_mode == "UPDATE":

    default_umoja_equipment_number = clean_text(
        asset[
            "Umoja Equipment Number"
        ]
    )


    default_barcode = clean_text(
        asset[
            "Barcode"
        ]
    )


    default_asset_number = clean_text(
        asset[
            "Asset Number"
        ]
    )


    default_description = clean_text(
        asset[
            "Description"
        ]
    )


    default_serial_number = clean_text(
        asset[
            "Serial Number (Mandatory Field)"
        ]
    )


    default_umoja_equipment_status = clean_text(
        asset[
            "Umoja Equipment Status"
        ]
    )


    default_umoja_status_description = clean_text(
        asset[
            "Umoja Equipment Status Description"
        ]
    )


    default_equipment_status = clean_text(
        asset[
            "Equipment Status (Mandatory Field)"
        ]
    ).upper()


    default_equipment_condition = clean_text(
        asset[
            "Equipment Condition (Mandatory Field)"
        ]
    )


    default_cost_center = clean_text(
        asset[
            "Cost Center"
        ]
    )


    default_acquisition_date = safe_date(
        asset[
            "Acquisition date"
        ]
    )


    default_acquisition_value = safe_number(
        asset[
            "Acquisition Value"
        ]
    )


    default_notification_number = clean_text(
        asset[
            "Umoja Notification Number"
        ]
    )


    default_disposal_case = clean_text(
        asset[
            "Disposal Case Number"
        ]
    )


    default_user_accountable = clean_text(
        asset[
            "User Accountable (Mandatory Field)"
        ]
    )


    default_user_index = clean_text(
        asset[
            "User Accountable Index Number (Mandatory Field)"
        ]
    )


    default_used_by_unv = clean_text(
        asset[
            "Is the Equipment Used By a UNV/Consultant/Intern"
        ]
    )


    default_unv_names = clean_text(
        asset[
            "UNV/Consultant/Intern Full Names"
        ]
    )


    default_staff_assign = clean_text(
        asset[
            "Staff Member To Assign Equipment Used By UNV/Consultants/Interns"
        ]
    )


    default_staff_index = clean_text(
        asset[
            "Staff Member Index Number"
        ]
    )


    default_location = clean_text(
        asset[
            "Functional Location"
        ]
    )


    default_verification_date = safe_date(
        asset[
            "Verification Date (Mandatory Field)"
        ]
    )


    default_verified_by = clean_text(
        asset[
            "Verified By (Mandatory Field)"
        ]
    )


    default_comments = clean_text(
        asset[
            "Additional Comments (Optional)"
        ]
    )


else:

    default_umoja_equipment_number = ""
    default_barcode = ""
    default_asset_number = ""
    default_description = ""
    default_serial_number = ""
    default_umoja_equipment_status = ""
    default_umoja_status_description = ""
    default_equipment_status = "Select Status"
    default_equipment_condition = ""
    default_cost_center = ""
    default_acquisition_date = date.today()
    default_acquisition_value = 0.0
    default_notification_number = ""
    default_disposal_case = ""
    default_user_accountable = ""
    default_user_index = ""
    default_used_by_unv = ""
    default_unv_names = ""
    default_staff_assign = ""
    default_staff_index = ""
    default_location = ""
    default_verification_date = date.today()
    default_verified_by = ""
    default_comments = ""


# =========================================================
# STATUS OPTIONS
# =========================================================

status_options = [
    "Select Status",
    "IN USE",
    "IN STORE",
    "FAULTY",
    "DISPOSED",
    "LOST"
]


if default_equipment_status in status_options:

    status_index = status_options.index(
        default_equipment_status
    )


else:

    if default_equipment_status:

        status_options.append(
            default_equipment_status
        )

        status_index = (
            len(status_options) - 1
        )

    else:

        status_index = 0


# =========================================================
# ADD / UPDATE FORM
# =========================================================

form_key = (
    f"asset_form_{form_mode}_"
    f"{st.session_state.edit_index}"
)


with st.form(
    form_key
):

    col1, col2 = st.columns(2)


    # =====================================================
    # LEFT COLUMN
    # =====================================================

    with col1:

        umoja_equipment_number = st.text_input(
            "Umoja Equipment Number",
            value=default_umoja_equipment_number
        )


        barcode = st.text_input(
            "Barcode",
            value=default_barcode
        )


        asset_number = st.text_input(
            "Asset Number",
            value=default_asset_number
        )


        description = st.text_input(
            "Description",
            value=default_description
        )


        serial_number = st.text_input(
            "Serial Number *",
            value=default_serial_number
        )


        umoja_equipment_status = st.text_input(
            "Umoja Equipment Status",
            value=default_umoja_equipment_status
        )


        umoja_equipment_status_description = st.text_input(
            "Umoja Equipment Status Description",
            value=default_umoja_status_description
        )


        equipment_status = st.selectbox(
            "Equipment Status *",
            status_options,
            index=status_index
        )


        equipment_condition = st.text_input(
            "Equipment Condition *",
            value=default_equipment_condition
        )


        cost_center = st.text_input(
            "Cost Center",
            value=default_cost_center
        )


        acquisition_date = st.date_input(
            "Acquisition Date",
            value=default_acquisition_date
        )


        acquisition_value = st.number_input(
            "Acquisition Value",
            min_value=0.0,
            value=default_acquisition_value,
            step=1.0
        )


    # =====================================================
    # RIGHT COLUMN
    # =====================================================

    with col2:

        umoja_notification_number = st.text_input(
            "Umoja Notification Number",
            value=default_notification_number
        )


        disposal_case_number = st.text_input(
            "Disposal Case Number",
            value=default_disposal_case
        )


        user_accountable = st.text_input(
            "User Accountable *",
            value=default_user_accountable
        )


        user_accountable_index = st.text_input(
            "User Accountable Index Number",
            value=default_user_index
        )


        used_by_unv = st.text_input(
            "Is the Equipment Used By a UNV/Consultant/Intern",
            value=default_used_by_unv
        )


        unv_full_names = st.text_input(
            "UNV/Consultant/Intern Full Names",
            value=default_unv_names
        )


        staff_member_assign = st.text_input(
            "Staff Member To Assign Equipment Used By UNV/Consultants/Interns",
            value=default_staff_assign
        )


        staff_member_index = st.text_input(
            "Staff Member Index Number",
            value=default_staff_index
        )


        functional_location = st.text_input(
            "Functional Location",
            value=default_location
        )


        verification_date = st.date_input(
            "Verification Date *",
            value=default_verification_date
        )


        verified_by = st.text_input(
            "Verified By *",
            value=default_verified_by
        )


        additional_comments = st.text_area(
            "Additional Comments",
            value=default_comments
        )


    # =====================================================
    # FORM BUTTON
    # =====================================================

    if form_mode == "UPDATE":

        submitted = st.form_submit_button(
            "Save Changes"
        )


    else:

        submitted = st.form_submit_button(
            "Save Asset"
        )


# =========================================================
# VALIDATION
# =========================================================

if submitted:

    missing_fields = []


    if not serial_number.strip():

        missing_fields.append(
            "Serial Number"
        )


    if equipment_status == "Select Status":

        missing_fields.append(
            "Equipment Status"
        )


    if not equipment_condition.strip():

        missing_fields.append(
            "Equipment Condition"
        )


    if not user_accountable.strip():

        missing_fields.append(
            "User Accountable"
        )


    if verification_date is None:

        missing_fields.append(
            "Verification Date"
        )


    if not verified_by.strip():

        missing_fields.append(
            "Verified By"
        )


    if missing_fields:

        st.error(
            "Please complete all mandatory fields."
        )


        for field in missing_fields:

            st.write(
                f"- {field}"
            )


    else:

        # =================================================
        # DUPLICATE CHECK
        # =================================================

        serial_clean = (
            serial_number
            .strip()
            .lower()
        )


        barcode_clean = (
            barcode
            .strip()
            .lower()
        )


        if form_mode == "UPDATE":

            check_df = df.drop(
                index=st.session_state.edit_index
            )


        else:

            check_df = df


        existing_serials = (
            check_df[
                "Serial Number (Mandatory Field)"
            ]
            .astype(str)
            .str.strip()
            .str.lower()
        )


        duplicate_serial = (
            serial_clean
            in
            existing_serials.values
        )


        duplicate_barcode = False


        if barcode_clean:

            existing_barcodes = (
                check_df[
                    "Barcode"
                ]
                .astype(str)
                .str.strip()
                .str.lower()
            )


            duplicate_barcode = (
                barcode_clean
                in
                existing_barcodes.values
            )


        if duplicate_serial:

            st.error(
                "Another asset already uses this Serial Number."
            )


        elif duplicate_barcode:

            st.error(
                "Another asset already uses this Barcode."
            )


        else:

            # =================================================
            # ASSET DATA
            # =================================================

            asset_data = {

                "Umoja Equipment Number":
                    umoja_equipment_number,

                "Barcode":
                    barcode,

                "Asset Number":
                    asset_number,

                "Description":
                    description,

                "Serial Number (Mandatory Field)":
                    serial_number,

                "Umoja Equipment Status":
                    umoja_equipment_status,

                "Umoja Equipment Status Description":
                    umoja_equipment_status_description,

                "Equipment Status (Mandatory Field)":
                    equipment_status,

                "Equipment Condition (Mandatory Field)":
                    equipment_condition,

                "Cost Center":
                    cost_center,

                "Acquisition date":
                    acquisition_date,

                "Acquisition Value":
                    acquisition_value,

                "Umoja Notification Number":
                    umoja_notification_number,

                "Disposal Case Number":
                    disposal_case_number,

                "User Accountable (Mandatory Field)":
                    user_accountable,

                "User Accountable Index Number (Mandatory Field)":
                    user_accountable_index,

                "Is the Equipment Used By a UNV/Consultant/Intern":
                    used_by_unv,

                "UNV/Consultant/Intern Full Names":
                    unv_full_names,

                "Staff Member To Assign Equipment Used By UNV/Consultants/Interns":
                    staff_member_assign,

                "Staff Member Index Number":
                    staff_member_index,

                "Functional Location":
                    functional_location,

                "Verification Date (Mandatory Field)":
                    verification_date,

                "Verified By (Mandatory Field)":
                    verified_by,

                "Additional Comments (Optional)":
                    additional_comments
            }


            # =================================================
            # UPDATE EXISTING ASSET
            # =================================================

            if form_mode == "UPDATE":

                row_index = (
                    st.session_state.edit_index
                )


                for column, value in asset_data.items():

                    df.at[
                        row_index,
                        column
                    ] = value


                df = df[
                    EXPECTED_COLUMNS
                ]


                save_data(
                    df
                )


                st.cache_data.clear()


                st.session_state.edit_mode = False

                st.session_state.edit_index = None


                st.success(
                    "Asset updated successfully."
                )


                st.rerun()


            # =================================================
            # ADD NEW ASSET
            # =================================================

            else:

                new_asset_df = pd.DataFrame(
                    [asset_data]
                )


                updated_df = pd.concat(
                    [
                        df,
                        new_asset_df
                    ],
                    ignore_index=True
                )


                updated_df = updated_df[
                    EXPECTED_COLUMNS
                ]


                save_data(
                    updated_df
                )


                st.cache_data.clear()


                st.success(
                    "Asset saved successfully."
                )


                st.rerun()
