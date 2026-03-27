import streamlit as st
import pandas as pd
import datetime
import os
from openpyxl import load_workbook

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

st.title("⚡ Energy Monitoring System")

# ==============================
# INPUT GRID FUNCTION
# ==============================

def input_grid(labels):
    values = {}
    for i in range(0, len(labels), 3):
        cols = st.columns(3)
        for j, label in enumerate(labels[i:i+3]):
            with cols[j]:
                values[label] = st.number_input(label, step=1.0, key=label)
    return values

# ==============================
# LABELS
# ==============================

tr_labels = [
    "TR-1 (31.5 MVA)", "TR-2 (31.5 MVA)", "TR-3 (31.5 MVA)",
    "TR-4 (31.5 MVA)", "TR-5 (31.5 MVA)"
]

lhf_labels = ["LHF-1 (44 MVA)", "LHF-2 (44 MVA)"]
lcp_labels = ["LCP FDR-1", "LCP FDR-3"]

lcss9_labels = ["LCSS-9 FDR-1", "LCSS-9 FDR-2", "LCSS-9 FDR-3"]
lcss8_labels = ["LCSS-8 FDR-1", "LCSS-8 FDR-2", "LCSS-8 FDR-3"]

# ==============================
# INPUTS
# ==============================

st.header("Enter Meter Readings")

tr_values = input_grid(tr_labels)
lhf_values = input_grid(lhf_labels)
lcp_values = input_grid(lcp_labels)
lcss9_values = input_grid(lcss9_labels)
lcss8_values = input_grid(lcss8_labels)

# ==============================
# SUBMIT BUTTON
# ==============================

if st.button("Submit"):

    wb = load_workbook(FILE_NAME)
    ws = wb[current_month]

    # ------------------------------
    # FIND TODAY COLUMN
    # ------------------------------
    col_index = None

    for col in range(3, ws.max_column + 1):
        if str(ws.cell(row=2, column=col).value) == today_str:
            col_index = col
            break

    if col_index is None:
        col_index = ws.max_column + 1
        ws.cell(row=2, column=col_index).value = today_str

    # ------------------------------
    # UPDATE FUNCTION
    # ------------------------------
    def update_excel(name, value):
        for row in range(4, ws.max_row + 1):
            if ws.cell(row=row, column=2).value == name:
                ws.cell(row=row, column=col_index).value = int(value)
                return

    # ------------------------------
    # CALCULATIONS (PYTHON BASED)
    # ------------------------------
    total_tr = sum(tr_values.values())
    total_lhf = sum(lhf_values.values())
    total_lcp = sum(lcp_values.values())
    total_lcss9 = sum(lcss9_values.values())
    total_lcss8 = sum(lcss8_values.values())

    total_caster = total_lcss8 + total_lcss9
    total_bof = total_tr - total_caster

    # ------------------------------
    # UPDATE RAW VALUES
    # ------------------------------
    for key, val in tr_values.items():
        update_excel(key, val)

    for key, val in lhf_values.items():
        update_excel(key, val)

    for key, val in lcp_values.items():
        update_excel(key, val)

    # ------------------------------
    # UPDATE TOTALS (NOW SAFE)
    # ------------------------------
    update_excel("TOTAL", total_tr)
    update_excel("TOTAL CASTER", total_caster)
    update_excel("TOTAL BOF", total_bof)

    # ------------------------------
    # SAVE FILE
    # ------------------------------
    wb.save(FILE_NAME)

    st.success("✅ Data Saved Successfully")

# ==============================
# DISPLAY DATA
# ==============================

if os.path.exists(FILE_NAME):

    df = pd.read_excel(FILE_NAME, sheet_name=current_month, header=1, dtype=object)

    # Fix date columns
    new_cols = list(df.columns[:2])
    for col in df.columns[2:]:
        try:
            new_col = pd.to_datetime(col).strftime("%d-%m-%Y")
        except:
            new_col = col
        new_cols.append(new_col)

    df.columns = new_cols

    # Convert numbers
    for col in df.columns[2:]:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Clean display
    df_display = df.copy()

    for col in df_display.columns[2:]:
        df_display[col] = df_display[col].apply(
            lambda x: int(x) if pd.notnull(x) else ""
        )

    df_display = df_display.fillna("")

    st.subheader("📊 Energy Data (Live)")
    st.dataframe(df_display, use_container_width=True)

    # ==============================
    # DOWNLOAD BUTTON
    # ==============================

    with open(FILE_NAME, "rb") as file:
        st.download_button(
            label="📥 Download Excel",
            data=file,
            file_name="Energy Sheet.xlsx"
        )
