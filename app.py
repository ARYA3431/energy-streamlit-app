import streamlit as st
import pandas as pd
import datetime
import os
from openpyxl import load_workbook

# ==============================
# BASIC SETTINGS
# ==============================

FILE_NAME = "Energy Sheet.xlsx"

current_month = datetime.datetime.now().strftime("%B")
today = datetime.datetime.now()
today_str = today.strftime("%d-%m-%Y")

st.title("Energy Monitoring System")

# ==============================
# USER INPUTS
# ==============================

st.header("Enter Meter Readings")

tr1 = st.number_input("TR-1 (31.5 MVA)", step=1.0)
tr2 = st.number_input("TR-2 (31.5 MVA)", step=1.0)
tr3 = st.number_input("TR-3 (31.5 MVA)", step=1.0)
tr4 = st.number_input("TR-4 (31.5 MVA)", step=1.0)
tr5 = st.number_input("TR-5 (31.5 MVA)", step=1.0)

lhf1 = st.number_input("LHF-1 (44 MVA)", step=1.0)
lhf2 = st.number_input("LHF-2 (44 MVA)", step=1.0)

lcp1 = st.number_input("LCP FDR-1", step=1.0)
lcp3 = st.number_input("LCP FDR-3", step=1.0)

lcss9_1 = st.number_input("LCSS-9 FDR-1", step=1.0)
lcss9_2 = st.number_input("LCSS-9 FDR-2", step=1.0)
lcss9_3 = st.number_input("LCSS-9 FDR-3", step=1.0)

lcss8_1 = st.number_input("LCSS-8 FDR-1", step=1.0)
lcss8_2 = st.number_input("LCSS-8 FDR-2", step=1.0)
lcss8_3 = st.number_input("LCSS-8 FDR-3", step=1.0)

ccm1 = st.number_input("CCM-1 EMS-1", step=1.0)
ccm2 = st.number_input("CCM-1 EMS-2", step=1.0)

pid1 = st.number_input("Primary ID Fan #1", step=1.0)
pid2 = st.number_input("Primary ID Fan #2", step=1.0)
sid1 = st.number_input("Secondary ID Fan #1", step=1.0)
sid2 = st.number_input("Secondary ID Fan #2", step=1.0)
sid3 = st.number_input("Secondary ID Fan #3", step=1.0)

rcph1 = st.number_input("RCPH I/C-1", step=1.0)
rcph2 = st.number_input("RCPH I/C-2", step=1.0)

grinder = st.number_input("Grinder I/C Caster", step=1.0)

# ==============================
# SUBMIT BUTTON
# ==============================

if st.button("Submit"):

    if not os.path.exists(FILE_NAME):
        st.error("Excel file not found!")
        st.stop()

    df = pd.read_excel(FILE_NAME, sheet_name=current_month, header=1)

    # Create today's column if missing
    if today_str not in df.columns:
        df[today_str] = 0

    # ==============================
    # CALCULATIONS
    # ==============================

    total_consumption = tr1 + tr2 + tr3 + tr4 + tr5
    total_lf = lhf1 + lhf2
    total_lcp = lcp1 + lcp3
    total_lcss9 = lcss9_1 + lcss9_2 + lcss9_3
    total_lcss8 = lcss8_1 + lcss8_2 + lcss8_3
    total_caster = total_lcss8 + total_lcss9 + ccm1 + ccm2 + grinder
    total_rcph = rcph1 + rcph2
    total_bof = total_consumption - (total_lcp + total_caster)
    total_id_fan = pid1 + pid2 + sid1 + sid2 + sid3

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

    update_value("TR-1 (31.5 MVA)", tr1)
    update_value("TR-2 (31.5 MVA)", tr2)
    update_value("TR-3 (31.5 MVA)", tr3)
    update_value("TR-4 (31.5 MVA)", tr4)
    update_value("TR-5 (31.5 MVA)", tr5)

    update_value("LHF#1 - TR (44 MVA)", lhf1)
    update_value("LHF#2 - TR (44 MVA)", lhf2)

    update_value("LCP FDR-1 (FDR33)", lcp1)
    update_value("LCP FDR-3 (FDR12)", lcp3)

    update_value("LCSS-9 FDR-1", lcss9_1)
    update_value("LCSS-9 FDR-2", lcss9_2)
    update_value("LCSS-9 FDR-3", lcss9_3)

    update_value("LCSS-8 FDR-1", lcss8_1)
    update_value("LCSS-8 FDR-2", lcss8_2)
    update_value("LCSS-8 FDR-3", lcss8_3)

    update_value("CCM-1 EMS-1", ccm1)
    update_value("CCM-1 EMS-2", ccm2)

    update_value("Primary ID Fan #1", pid1)
    update_value("Primary ID Fan #2", pid2)
    update_value("Secondary ID Fan #1", sid1)
    update_value("Secondary ID Fan #2", sid2)
    update_value("Secondary ID Fan #3", sid3)

    update_value("RCPH I/C-1 (FDR14)", rcph1)
    update_value("RCPH I/C-2 (FDR27)", rcph2)

    update_value("Grinder I/C Caster (FDR36)", grinder)

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

    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
         df.to_excel(writer, sheet_name=current_month, index=False, startrow=1)
st.success("Data Saved Successfully ✅")
st.subheader("Updated Data Preview")
st.dataframe(df)


