import streamlit as st
import pandas as pd
import datetime
import os
from openpyxl import load_workbook

# ==============================
# BASIC SETTINGS
# ==============================

FILE_NAME = "/tmp/Energy Sheet.xlsx"
SOURCE_FILE = "Energy Sheet.xlsx"

import shutil

if not os.path.exists(FILE_NAME):
    shutil.copy(SOURCE_FILE, FILE_NAME)

current_month = datetime.datetime.now().strftime("%B")
today = datetime.datetime.now()
today_str = today.strftime("%d-%m-%Y")

st.title("Energy Monitoring System")

# ==============================
# USER INPUTS
# ==============================

st.header("Enter Meter Readings")

def input_grid(labels):
    values = {}
    for i in range(0, len(labels), 3):
        cols = st.columns(3)
        for j, label in enumerate(labels[i:i+3]):
            with cols[j]:
                values[label] = st.number_input(label, step=1.0, key=label)
    return values


tr_labels = [
"TR-1 (31.5 MVA)", "TR-2 (31.5 MVA)", "TR-3 (31.5 MVA)",
"TR-4 (31.5 MVA)", "TR-5 (31.5 MVA)"
]

lhf_labels = [
"LHF-1 (44 MVA)", "LHF-2 (44 MVA)"
]

lcp_labels = [
"LCP FDR-1", "LCP FDR-3"
]

lcss9_labels = [
"LCSS-9 FDR-1", "LCSS-9 FDR-2", "LCSS-9 FDR-3"
]

lcss8_labels = [
"LCSS-8 FDR-1", "LCSS-8 FDR-2", "LCSS-8 FDR-3"
]

ccm_labels = [
"CCM-1 EMS-1", "CCM-1 EMS-2"
]

fan_labels = [
"Primary ID Fan #1", "Primary ID Fan #2",
"Secondary ID Fan #1", "Secondary ID Fan #2", "Secondary ID Fan #3"
]

rcph_labels = [
"RCPH I/C-1", "RCPH I/C-2"
]

other_labels = [
"Grinder I/C Caster"
]


tr_values = input_grid(tr_labels)
lhf_values = input_grid(lhf_labels)
lcp_values = input_grid(lcp_labels)
lcss9_values = input_grid(lcss9_labels)
lcss8_values = input_grid(lcss8_labels)
ccm_values = input_grid(ccm_labels)
fan_values = input_grid(fan_labels)
rcph_values = input_grid(rcph_labels)
other_values = input_grid(other_labels)

# ==============================
# SUBMIT BUTTON
# ==============================

if st.button("Submit"):

    if not os.path.exists(FILE_NAME):
        st.error("Excel file not found!")
        st.stop()

    df = pd.read_excel(FILE_NAME, sheet_name=current_month)

    # Create today's column if missing
    if today_str not in df.columns:
        df[today_str] = 0

    # ==============================
    # CALCULATIONS
    # ==============================

    total_consumption = sum(tr_values.values())
    total_lf = sum(lhf_values.values())
    total_lcp = sum(lcp_values.values())
    total_lcss9 = sum(lcss9_values.values())
    total_lcss8 = sum(lcss8_values.values())
    total_caster = total_lcss8 + total_lcss9
    total_rcph = sum(rcph_values.values())
    total_id_fan = sum(fan_values.values())
    grinder = other_values["Grinder I/C Caster"]
    ccm1 = ccm_values["CCM-1 EMS-1"]
    ccm2 = ccm_values["CCM-1 EMS-2"]
    total_bof = total_consumption - total_caster

    # ==============================
    # FUNCTION TO UPDATE VALUE
    # ==============================

    def update_value(name, value):

        mask = df.iloc[:, 1] == name

        if mask.any():
            df.loc[mask, today_str] = int(value)

        else:
            new_row = {col: 0 for col in df.columns}
            new_row[df.columns[1]] = name
            new_row[today_str] = int(value)

            df.loc[len(df)] = new_row

    # ==============================
    # UPDATE EQUIPMENT VALUES
    # ==============================

   

    # ==============================
    # UPDATE TOTALS
    # ==============================

    update_value("TOTAL CONSUMPTION", total_consumption)
    update_value("TOTAL LF CONSUMPTION", total_lf)
    update_value("TOTAL LCP CONSUMPTION", total_lcp)
    update_value("TOTAL LCSS9 CONSUMPTION", total_lcss9)
    update_value("TOTAL LCSS8 CONSUMPTION", total_lcss8)
    update_value("TOTAL CASTER CONSUMPTION", total_caster)
    update_value("TOTAL BOF CONSUMPTION", total_bof)
    update_value("TOTAL RCPH CONSUMPTION", total_rcph)
    update_value("TOTAL ID FAN CONSUMPTION", total_id_fan)

    # ==============================
    # SAVE TO EXCEL
    # ==============================

    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=current_month, index=False)

    st.success("Data Saved Successfully ✅")

    st.subheader("Updated Data Preview")
    st.dataframe(df)

