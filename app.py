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
# ALL LABELS (MATCHING YOUR SHEET)
# ==============================

tr_labels = [
    "TR-1 (31.5 MVA)", "TR-2 (31.5 MVA)", "TR-3 (31.5 MVA)",
    "TR-4 (31.5 MVA)", "TR-5 (31.5 MVA)"
]

lhf_labels = ["LHF#1", "LHF#2"]

lcss9_labels = [
    "LCSS-9 FDR-1", "LCSS-9 FDR-3", "LCSS-9 FDR-2"
]

lcss8_labels = [
    "LCSS-8 FDR-1", "LCSS-8 FDR-3", "LCSS-8 FDR-2"
]

ccm_labels = ["CCM-1 EMS-1", "CCM-1 EMS-2"]

fan_labels = [
    "PRIMARY ID FAN #1", "PRIMARY ID FAN #2",
    "SECONDARY ID FAN#1", "SECONDARY ID FAN#2", "SECONDARY ID FAN#3"
]

rcph_labels = ["RCPH I/C-1", "RCPH I/C-2"]

lcp_labels = ["LCP FDR-1", "LCP FDR-3"]

other_labels = ["Grinder I/C Caster"]

# ✅ NEW
heat_labels = ["No. of Heat Tap", "No. of Heat Cast"]

# ==============================
# INPUT SECTION
# ==============================

st.header("Enter Meter Readings")

tr_values = input_grid(tr_labels)
lhf_values = input_grid(lhf_labels)
lcss9_values = input_grid(lcss9_labels)
lcss8_values = input_grid(lcss8_labels)
ccm_values = input_grid(ccm_labels)
fan_values = input_grid(fan_labels)
rcph_values = input_grid(rcph_labels)
lcp_values = input_grid(lcp_labels)
other_values = input_grid(other_labels)

# ✅ HEAT INPUT
heat_values = input_grid(heat_labels)
# ==============================
# CALCULATIONS
# ==============================

# ==============================
# LIVE CALCULATIONS (BEFORE SUBMIT)
# ==============================

st.subheader("⚡ Live Energy Calculation")

total_tr = sum(tr_values.values())
total_lf = sum(lhf_values.values())
total_lcp = sum(lcp_values.values())

total_caster = (
    sum(lcss8_values.values()) +
    sum(lcss9_values.values()) +
    sum(ccm_values.values()) +
    other_values["Grinder I/C Caster"]
)

total_bof = total_tr - total_lcp - total_caster

# TOTAL RCPH
total_rcph = sum(rcph_values.values())

# HEAT BASED CALCULATION
heat_tap = heat_values["No. of Heat Tap"]
heat_cast = heat_values["No. of Heat Cast"]

per_ton = 0
if heat_cast > 0:
    per_ton = total_tr / heat_cast

# ==============================
# DISPLAY (DASHBOARD STYLE)
# ==============================

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("⚡ TOTAL TR", int(total_tr))

with col2:
    st.metric("🔥 TOTAL LF", int(total_lf))

with col3:
    st.metric("🏭 TOTAL LCP", int(total_lcp))

col4, col5, col6 = st.columns(3)

with col4:
    st.metric("🏗 TOTAL CASTER", int(total_caster))

with col5:
    st.metric("🏭 TOTAL BOF", int(total_bof))

with col6:
    st.metric("🌡 TOTAL RCPH", int(total_rcph))

col7, col8 = st.columns(2)

with col7:
    st.metric("🔢 HEAT TAP", int(heat_tap))

with col8:
    st.metric("⚙ HEAT CAST", int(heat_cast))

st.metric("📊 PER TON CONSUMPTION", round(per_ton, 2))

# ==============================
# SUBMIT BUTTON
# ==============================

if st.button("Submit"):
    
        wb = load_workbook(FILE_NAME, data_only=False)
        ws = wb[current_month]

        # FIND TODAY COLUMN
        col_index = None
        for col in range(3, ws.max_column + 1):
            if str(ws.cell(row=2, column=col).value) == today_str:
                col_index = col
                break

        if col_index is None:
            col_index = ws.max_column + 1
            ws.cell(row=2, column=col_index).value = today_str

    
        def get_previous_value(name):

            for row in range(4, ws.max_row + 1):

                col1 = ws.cell(row=row, column=1).value
                col2 = ws.cell(row=row, column=2).value

                combined = f"{col1} {col2}".upper()

                if name.upper() in combined:

                    # Previous column
                    prev_col = col_index - 1

                    if prev_col >= 3:
                        prev_value = ws.cell(row=row, column=prev_col).value
                        return prev_value if prev_value else 0

        return 0

        # FUNCTION (INSIDE)
        def clean_text(text):
            return str(text).upper().replace("-", "").replace("#", "").replace(" ", "")
    
        def update_excel(name, value):
    
            clean_name = clean_text(name)
    
            for row in range(4, ws.max_row + 1):

                col1 = ws.cell(row=row, column=1).value
                col2 = ws.cell(row=row, column=2).value

                combined = f"{col1} {col2}"
                clean_combined = clean_text(combined)

                if clean_name in clean_combined:
                    ws.cell(row=row, column=col_index).value = int(value)
                    return

            st.write(f"❌ NOT FOUND: {name}")

    # ==============================
    # UPDATE INPUT VALUES
    # ==============================

        for group in [
            tr_values, lhf_values, lcss9_values, lcss8_values,
            ccm_values, fan_values, rcph_values, lcp_values, other_values
        ]:
            for key, val in group.items():
                update_excel(key, val)

        # HEAT
        update_excel("No. of Heat Tap", heat_values["No. of Heat Tap"])
        update_excel("No. of Heat Cast", heat_values["No. of Heat Cast"])

    # ==============================
        # ✅ TOTALS (INSIDE SUBMIT ONLY)
    # ==============================

        update_excel("Total", total_tr)
        update_excel("TOTAL LF CONSUMPTION", total_lf)
        update_excel("TOTAL LCP CONSUMPTION", total_lcp)
        update_excel("TOTAL CASTER CONSUMPTION", total_caster)
        update_excel("TOTAL BOF CONSUMPTION", total_bof)
        update_excel("TOTAL RCPH CONSUMPTION", total_rcph)

    # SAVE
        wb.calculation.fullCalcOnLoad = True
        wb.save(FILE_NAME)

        st.success("✅ Data Saved Successfully")

# ==============================
# DISPLAY DATA (LIVE VIEW)
# ==============================

# ==============================
# DISPLAY DATA (READ FORMULA VALUES)
# ==============================

if os.path.exists(FILE_NAME):

    # ✅ Read calculated values (IMPORTANT)
    wb_data = load_workbook(FILE_NAME, data_only=True)
    ws_data = wb_data[current_month]

    data = list(ws_data.values)

    # Convert to DataFrame
    df = pd.DataFrame(data[2:], columns=data[1])

    # ==============================
    # FIX COLUMN NAMES (DATES)
    # ==============================

    new_cols = list(df.columns[:2])

    for col in df.columns[2:]:
        try:
            new_col = pd.to_datetime(col).strftime("%d-%m-%Y")
        except:
            new_col = col
        new_cols.append(new_col)

    df.columns = new_cols

    # ==============================
    # CLEAN DATA
    # ==============================

    for col in df.columns[2:]:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df_display = df.copy()

    for col in df_display.columns[2:]:
        df_display[col] = df_display[col].apply(
            lambda x: int(x) if pd.notnull(x) else ""
        )

    df_display = df_display.fillna("")

    # ==============================
    # SHOW DATA
    # ==============================

    st.subheader("📊 Energy Data (Live)")
    st.dataframe(df_display, use_container_width=True)

    # ==============================
    # DOWNLOAD BUTTON
    # ==============================

    with open(FILE_NAME, "rb") as file:
        st.download_button(
            label="📥 Download Updated Excel",
            data=file,
            file_name="Energy Sheet.xlsx"
        )
