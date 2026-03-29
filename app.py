import streamlit as st
import pandas as pd
import datetime
import os
from openpyxl import load_workbook

def get_previous_total(ws, col_index, name):

    prev_col = col_index - 1

    if prev_col < 3:
        return 0  # No previous data

    def clean_text(text):
        return str(text).upper().replace("-", "").replace("#", "").replace(" ", "")

    clean_name = clean_text(name)

    for row in range(4, ws.max_row + 1):

        col1 = ws.cell(row=row, column=1).value
        col2 = ws.cell(row=row, column=2).value

        combined = f"{col1} {col2}"
        clean_combined = clean_text(combined)

        if clean_name in clean_combined:
            value = ws.cell(row=row, column=prev_col).value

            if value is None:
                return 0

            return float(value)

    return 0

# ==============================
# BASIC SETTINGS
# ==============================

FILE_NAME = "Energy Sheet.xlsx"

if not os.path.exists(FILE_NAME):
    st.error("Excel file not found!")
    st.stop()

current_month = datetime.datetime.now().strftime("%B")
today_str = datetime.datetime.now().strftime("%d-%m-%Y")

st.title("⚡ Energy Monitoring System")

# ==============================
# HELPER FUNCTIONS (TOP ONLY)
# ==============================

def clean_text(text):
    return str(text).upper().replace("-", "").replace("#", "").replace(" ", "")

def update_excel(ws, col_index, name, value):
    clean_name = clean_text(name)

    for row in range(4, ws.max_row + 1):
        col1 = ws.cell(row=row, column=1).value
        col2 = ws.cell(row=row, column=2).value

        combined = f"{col1} {col2}"
        clean_combined = clean_text(combined)

        if clean_name in clean_combined:
            ws.cell(row=row, column=col_index).value = int(value)
            return

def get_previous_value(ws, name, col_index):
    clean_name = clean_text(name)

    for row in range(4, ws.max_row + 1):
        col1 = ws.cell(row=row, column=1).value
        col2 = ws.cell(row=row, column=2).value

        combined = f"{col1} {col2}"
        clean_combined = clean_text(combined)

        if clean_name in clean_combined:
            val = ws.cell(row=row, column=col_index - 1).value
            return val if val else 0

    return 0

def per_day(ws, col_index, name, today_val):
    yesterday = get_previous_value(ws, name, col_index)
    return today_val - yesterday

# ==============================
# INPUT FUNCTION
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
# INPUT LABELS
# ==============================

tr_labels = [
    "TR-1 (31.5 MVA)", "TR-2 (31.5 MVA)", "TR-3 (31.5 MVA)",
    "TR-4 (31.5 MVA)", "TR-5 (31.5 MVA)"
]

lhf_labels = ["LHF#1", "LHF#2"]

lcss9_labels = ["LCSS-9 FDR-1", "LCSS-9 FDR-3", "LCSS-9 FDR-2"]
lcss8_labels = ["LCSS-8 FDR-1", "LCSS-8 FDR-3", "LCSS-8 FDR-2"]

ccm_labels = ["CCM-1 EMS-1", "CCM-1 EMS-2"]

fan_labels = [
    "PRIMARY ID FAN #1", "PRIMARY ID FAN #2",
    "SECONDARY ID FAN#1", "SECONDARY ID FAN#2", "SECONDARY ID FAN#3"
]

rcph_labels = ["RCPH I/C-1", "RCPH I/C-2"]

lcp_labels = ["LCP FDR-1", "LCP FDR-3"]

other_labels = ["Grinder I/C Caster"]

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
heat_values = input_grid(heat_labels)

# ==============================
# LIVE CALCULATION
# ==============================

st.subheader("⚡ Live Calculation")

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
total_rcph = sum(rcph_values.values())

heat_tap = heat_values["No. of Heat Tap"]
heat_cast = heat_values["No. of Heat Cast"]
# ==============================
# STEP 3: PER DAY CALCULATION
# ==============================
lcp_per_day = total_lcp - lcp_yesterday

per_ton = total_tr / heat_cast if heat_cast > 0 else 0

# ==============================
# DISPLAY METRICS
# ==============================

col1, col2, col3 = st.columns(3)
col1.metric("TOTAL TR", int(total_tr))
col2.metric("TOTAL LF", int(total_lf))
col3.metric("TOTAL LCP", int(total_lcp))

col4, col5, col6 = st.columns(3)
col4.metric("TOTAL CASTER", int(total_caster))
col5.metric("TOTAL BOF", int(total_bof))
col6.metric("TOTAL RCPH", int(total_rcph))

col7, col8 = st.columns(2)
col7.metric("HEAT TAP", int(heat_tap))
col8.metric("HEAT CAST", int(heat_cast))

st.metric("PER TON", round(per_ton, 2))

# ==============================
# SUBMIT
# ==============================

if st.button("Submit"):

    wb = load_workbook(FILE_NAME, data_only=False)
    ws = wb[current_month]

    # Find / Create column
    col_index = None
    for col in range(3, ws.max_column + 1):
        if str(ws.cell(row=2, column=col).value) == today_str:
            col_index = col
            break

    if col_index is None:
        col_index = ws.max_column + 1
        ws.cell(row=2, column=col_index).value = today_str

    # ✅ STEP 2
    lcp_yesterday = get_previous_total(ws, col_index, "TOTAL LCP CONSUMPTION")

    # ✅ STEP 3 (THIS WAS MISSING)
    lcp_per_day = total_lcp - lcp_yesterday

    # UPDATE INPUT VALUES
    for group in [
        tr_values, lhf_values, lcss9_values, lcss8_values,
        ccm_values, fan_values, rcph_values, lcp_values, other_values
    ]:
        for key, val in group.items():
            update_excel(ws, col_index, key, val)

    # HEAT
    update_excel(ws, col_index, "No. of Heat Tap", heat_tap)
    update_excel(ws, col_index, "No. of Heat Cast", heat_cast)

    # TOTALS
    update_excel(ws, col_index, "Total", total_tr)
    update_excel(ws, col_index, "TOTAL LF CONSUMPTION", total_lf)
    update_excel(ws, col_index, "TOTAL LCP CONSUMPTION", total_lcp)
    update_excel(ws, col_index, "TOTAL CASTER CONSUMPTION", total_caster)
    update_excel(ws, col_index, "TOTAL BOF CONSUMPTION", total_bof)
    update_excel(ws, col_index, "TOTAL RCPH CONSUMPTION", total_rcph)

    # ✅ STEP 4
    update_excel(ws, col_index, "LCP PER DAY CONSUMPTION", lcp_per_day)

    # SAVE
    wb.calculation.fullCalcOnLoad = True
    wb.save(FILE_NAME)

    st.success("✅ Data Saved Successfully")
# ==============================
# DISPLAY TABLE
# ==============================

wb_data = load_workbook(FILE_NAME, data_only=True)
ws_data = wb_data[current_month]

data = list(ws_data.values)
df = pd.DataFrame(data[2:], columns=data[1])

# Fix date columns
new_cols = list(df.columns[:2])
for col in df.columns[2:]:
    try:
        new_cols.append(pd.to_datetime(col).strftime("%d-%m-%Y"))
    except:
        new_cols.append(col)

df.columns = new_cols

# Clean display
for col in df.columns[2:]:
    df[col] = pd.to_numeric(df[col], errors='coerce')

df = df.fillna("")

for col in df.columns[2:]:
    df[col] = df[col].apply(lambda x: int(x) if x != "" else "")

st.subheader("📊 Energy Data")
st.dataframe(df, use_container_width=True)

# DOWNLOAD
with open(FILE_NAME, "rb") as file:
    st.download_button("📥 Download Excel", file, "Energy Sheet.xlsx")
