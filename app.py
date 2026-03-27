import streamlit as st
import pandas as pd
import datetime
import os
from openpyxl import load_workbook
import shutil

# ==============================
# BASIC SETTINGS
# ==============================

FILE_NAME = "Energy Sheet.xlsx"

if not os.path.exists(FILE_NAME):
    st.error("Excel file not found!")
    st.stop()

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

lhf_labels = ["LHF-1 (44 MVA)", "LHF-2 (44 MVA)"]
lcp_labels = ["LCP FDR-1", "LCP FDR-3"]

lcss9_labels = ["LCSS-9 FDR-1", "LCSS-9 FDR-2", "LCSS-9 FDR-3"]
lcss8_labels = ["LCSS-8 FDR-1", "LCSS-8 FDR-2", "LCSS-8 FDR-3"]

ccm_labels = ["CCM-1 EMS-1", "CCM-1 EMS-2"]

fan_labels = [
    "Primary ID Fan #1", "Primary ID Fan #2",
    "Secondary ID Fan #1", "Secondary ID Fan #2", "Secondary ID Fan #3"
]

rcph_labels = ["RCPH I/C-1", "RCPH I/C-2"]

other_labels = ["Grinder I/C Caster"]


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

    wb = load_workbook(FILE_NAME)
    ws = wb[current_month]

    # ==============================
    # FIND TODAY COLUMN (STRING MATCH)
    # ==============================

    col_index = None

    for col in range(3, ws.max_column + 1):
        cell_value = ws.cell(row=2, column=col).value

        if str(cell_value) == today_str:
            col_index = col
            break

    # If today's column not found → create new
    if col_index is None:
        col_index = ws.max_column + 1
        ws.cell(row=2, column=col_index).value = today_str

    # ==============================
    # UPDATE FUNCTION
    # ==============================

    def update_excel(name, value):
        for row in range(4, ws.max_row + 1):
            if ws.cell(row=row, column=2).value == name:
                ws.cell(row=row, column=col_index).value = int(value)
                return

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
    total_bof = total_consumption - total_caster

    # ==============================
    # UPDATE VALUES
    # ==============================

    update_excel("TR-1 (31.5 MVA)", tr_values["TR-1 (31.5 MVA)"])
    update_excel("TR-2 (31.5 MVA)", tr_values["TR-2 (31.5 MVA)"])
    update_excel("TR-3 (31.5 MVA)", tr_values["TR-3 (31.5 MVA)"])
    update_excel("TR-4 (31.5 MVA)", tr_values["TR-4 (31.5 MVA)"])
    update_excel("TR-5 (31.5 MVA)", tr_values["TR-5 (31.5 MVA)"])

    update_excel("TOTAL CONSUMPTION", total_consumption)

    # ==============================
    # SAVE FILE
    # ==============================

    wb.save(FILE_NAME)

    # ==============================
    # READ + CLEAN DATA FOR DISPLAY
    # ==============================

    df = pd.read_excel(FILE_NAME, sheet_name=current_month, header=1, dtype=object)

    # Fix column names (dates)
    df = pd.read_excel(FILE_NAME, sheet_name=current_month, header=1, dtype=object)

# Fix column names (dates)
new_cols = list(df.columns[:2])

for col in df.columns[2:]:
    try:
        new_col = pd.to_datetime(col).strftime("%d-%m-%Y")
    except:
        new_col = col
    new_cols.append(new_col)

df.columns = new_cols

# Convert values properly
for col in df.columns[2:]:
    df[col] = pd.to_numeric(df[col], errors='coerce').astype("Int64")

# Only fill blanks (do NOT delete rows)
df = df.fillna("")

st.subheader("📊 Full Energy Data (Live)")
st.dataframe(df, use_container_width=True)

    # ==============================
    # DOWNLOAD BUTTON
    # ==============================

    with open(FILE_NAME, "rb") as file:
        st.download_button(
        label="📥 Download Updated Excel",
        data=file,
        file_name="Energy Sheet.xlsx"
        )
