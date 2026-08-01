import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="HQA E&M Open Receipt",
    layout="wide",
    initial_sidebar_state="expanded"
)
st.markdown("""
<style>
[data-testid="stSidebarCollapsedControl"] {
    display: hide;
}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<style>
.block-container {
    max-width: 100%;
    padding-top: 1rem;
    padding-left: 1rem;
    padding-right: 1rem;
}

# -----------------------------
# Page Configuration
# -----------------------------

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
[data-testid="stSidebar"] {
    min-width: 150px;
    max-width: 150px;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Title
# -----------------------------
col1, col2 = st.columns([2, 1])

with col1:
    st.title("📊 HQA E&M Open Recepite Dashboard")
with col2:
    st.markdown(
        f"""
        <div style="margin-top:90px; text-align:right; font-size:15px;">
            📅 {datetime.now().strftime('%d-%m-%Y')}
        </div>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------
# Upload Excel
# -----------------------------
uploaded_file = st.file_uploader(
    "📂 Upload MB52 Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is None:
    st.info("Please upload an MB52 Excel file.")
    st.stop()

# -----------------------------
# Read Excel
# -----------------------------
try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Unable to read Excel file.\n\n{e}")
    st.stop()

if df.empty:
    st.error("Uploaded file is empty.")
    st.stop()

# -----------------------------
# Required Column Positions
# -----------------------------
MATERIAL = 0
PLANT = 2
GRN = 8
GRN_DATE = 9
VALUE = 19

# Validate column count
if len(df.columns) <= VALUE:
    st.error(
        "Excel format is incorrect.\n"
        "Pending Value column (T) not found."
    )
    st.stop()

# -----------------------------
# Prepare Data
# -----------------------------
df["Material"] = (
  df.iloc[:, MATERIAL]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

df["Plant"] = (
    df.iloc[:, PLANT]
    .fillna("")
    .apply(lambda x: str(x).split(".")[0])
    .str.strip()
)

df["GRN"] = (
    df.iloc[:, GRN]
    .fillna("")
    .apply(lambda x: str(x).split(".")[0])
    .str.strip()
)

df["GRN DATE"] = pd.to_datetime(
    df.iloc[:, GRN_DATE],
    errors="coerce"
)

df["Value"] = pd.to_numeric(
    df.iloc[:, VALUE],
    errors="coerce"
).fillna(0)

# -----------------------------
# Load Electrical Material Master
# -----------------------------
master = pd.read_excel("material_master.xlsx", header=None)

electrical_items = set(
    master[0]
    .fillna(0)
    .astype(str)
    .str.strip()
)

df["Department"] = df["Material"].apply(
    lambda x: "Electrical"
    if str(x).strip() in electrical_items
    else "Mechanical"
)

# -----------------------------
# Department Logic
# -----------------------------
def get_department(material):

    material = str(material)

    if (
        material.startswith("1000")
        or material.startswith("385")
        or material.startswith("44")
        or material.startswith("45")
        or material.startswith("46")
        or material.startswith("485")
        or material.startswith("63")
    ):
        return "Electrical"

    return "Mechanical"

df["Department"] = df["Material"].apply(get_department)

# Convert Plant codes to names
plant_map = {
    "1201": "Ecity",
    "1202": "Vemgal"
}

df["Plant"] = (
    df["Plant"]
      .astype(str)
      .str.strip()
      .map(plant_map)
      .fillna(df["Plant"].astype(str))
)

# -----------------------------
# Sidebar Filters
# -----------------------------
st.sidebar.header("Filters")

plant_list = ["All Plants"] + sorted(
    df["Plant"].dropna().unique().tolist()
)

dept_list = [
    "All",
    "Electrical",
    "Mechanical"
]

selected_plant = st.sidebar.selectbox(
    "🏭 Plant",
    plant_list
)

selected_department = st.sidebar.selectbox(
    "⚙️ Department",
    dept_list
)
#Tasl Logo
st.sidebar.markdown("---")

st.sidebar.image("./tasl_logo.png", width=200)

# -----------------------------
# Apply Filters
# -----------------------------
filtered = df.copy()

if selected_plant != "All Plants":
    filtered = filtered[
        filtered["Plant"] == selected_plant
    ]

if selected_department != "All":
    filtered = filtered[
        filtered["Department"] == selected_department
    ]

# =====================================================
# KPI CARDS
# =====================================================
st.markdown("---")
st.subheader("📌 Dashboard")

kpi1, kpi2, kpi3 = st.columns(3)

total_value = filtered["Value"].sum()
grn_count = filtered["GRN"].nunique()
lot_count = len(filtered)

kpi1.metric(
    "💰 Pending Value",
    f"{total_value:,.2f}"
)

kpi2.metric(
    "📄 GRN Count",
    grn_count
)

kpi3.metric(
    "📦 Lot Count",
    lot_count
)

# =====================================================
# SUMMARY TABLES
# =====================================================

# Department Summary
dept_summary = (
    filtered
    .groupby("Department", as_index=False)
    .agg(Pending_Value=("Value", "sum"))
)

# Plant Summary
summary = (
    filtered[
        filtered["Plant"].notna() &
        (filtered["Plant"].astype(str).str.strip() != "")
    ]
    .groupby(["Plant", "Department"], as_index=False)
    .agg(
        GRN_Count=("GRN", "nunique"),
        Lot_Count=("GRN", "count"),
        Pending_Value=("Value", "sum")
    )
)

st.markdown("---")

col1, col2 = st.columns([3, 3])

with col1:
    st.subheader("⚙️ Department Pending Value")
    st.dataframe(
        dept_summary,
        hide_index=True,
        use_container_width=False
    )
with col1:
    st.subheader("📊 Plant Summary")
    st.dataframe(
        summary,
        hide_index=True,
        use_container_width=True,
        height=180
    )

display_df = filtered.copy()

# =====================================================
# DETAILED DATA
# =====================================================

st.markdown("---")
st.subheader("📋 Detailed Pending Data")

# Create detail_df FIRST
detail_df = (
    display_df
    .groupby(["Plant", "Department", "GRN", "GRN DATE"], as_index=False)
    .agg(
        Value=("Value", "sum")
    )
)

# Filters
col1, col2, col3 = st.columns(3)

with col1:
    detail_plant = st.selectbox(
        "🏭 Plant",
        ["All"] + sorted(detail_df["Plant"].unique().tolist())
    )

with col2:
    detail_department = st.selectbox(
        "⚙️ Department",
        ["All"] + sorted(detail_df["Department"].unique().tolist())
    )

with col3:
    detail_grn = st.selectbox(
        "📄 GRN",
        ["All"] + sorted(detail_df["GRN"].astype(str).unique().tolist())
    )

filtered_detail = detail_df.copy()

if detail_plant != "All":
    filtered_detail = filtered_detail[
        filtered_detail["Plant"] == detail_plant
    ]

if detail_department != "All":
    filtered_detail = filtered_detail[
        filtered_detail["Department"] == detail_department
    ]

if detail_grn != "All":
    filtered_detail = filtered_detail[
        filtered_detail["GRN"].astype(str) == detail_grn
    ]
filtered_detail = filtered_detail[
    filtered_detail["Plant"].fillna("").astype(str).str.strip() != ""
]

detail = (
    filtered_detail.groupby(
        [
            "Plant",
            "Department",
            "GRN",
            "GRN DATE",
        ],
        as_index=False
    )
    .agg(
        value=("Value", "sum")
    )
)

from pandas.tseries.offsets import BDay
import pandas as pd

# Convert GRN DATE to datetime
detail["GRN DATE"] = pd.to_datetime(
    detail["GRN DATE"],
    errors="coerce"
)
# Today's date
today = pd.Timestamp.today().normalize()

# Calculate 5 working day due date (Saturday & Sunday excluded)
detail["Due Date"] = detail["GRN DATE"] + BDay(5)

# Calculate 5 working day due date (Saturday & Sunday excluded)
detail["Due Date"] = detail["GRN DATE"] + BDay(5)

# Quarter-end override
quarter_end = detail["GRN DATE"] + pd.offsets.QuarterEnd(0)

mask = (
    detail["GRN DATE"].dt.month.isin([3, 6, 9, 12]) &
    (detail["Due Date"] > quarter_end)
)

detail.loc[mask, "Due Date"] = quarter_end[mask]

# Calculate Ageing
today = pd.Timestamp.today().normalize()

detail["Closing 5 Days"] = (
    detail["Due Date"] - today
).dt.days

detail["GRN DATE"] = detail["GRN DATE"].dt.strftime("%d-%m-%Y")

# Show Due Date without time
detail["Due day"] = detail["Due Date"].dt.strftime("%d-%m-%Y")

# Remove .000000 from GRN

detail["GRN"] = (

    pd.to_numeric(detail["GRN"], errors="coerce")

    .astype("Int64")

    .astype(str)

)

 # Remove  .00000 from value
detail["value"] = detail["value"].apply(lambda x: format(float(x), ".2f").rstrip("0").rstrip("."))

detail.rename(columns={
    "Plant": "Plant",
    "Department": "Department",
    "GRN": "GRN",
    "GRN DATE": "GRN Date",
    "Due Date": "Closing Date",
    "Closing 5 Days": "Days Left",
    "Value": "Pending Value (₹)"
}, inplace=True)

# Highlight overdue rows
def highlight_overdue(row):
    if row["Days Left"] < 0:
        return ["background-color: #ffcccc"] * len(row)
    return [""] * len(row)
    
# Remove Closing Date column from display
display_detail = detail.drop(columns=["Closing Date"])

# Display table
st.dataframe(
    detail.style.apply(highlight_overdue, axis=1),
    use_container_width=True,
    hide_index=True
)
#Excel Download
output = io.BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:

    electrical = detail[detail["Department"] == "Electrical"]
    mechanical = detail[detail["Department"] == "Mechanical"]

    electrical.to_excel(
        writer,
        sheet_name="Electrical",
        index=False
    )

    mechanical.to_excel(
        writer,
        sheet_name="Mechanical",
        index=False
    )

    workbook = writer.book

    header_fill = PatternFill(
        start_color="4F81BD",
        end_color="4F81BD",
        fill_type="solid"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    for ws in workbook.worksheets:

        last_row = ws.max_row + 1

        ws.cell(last_row, 1).value = "TOTAL"

        value_col = ws.max_column

        ws.cell(
            last_row,
            value_col
        ).value = f"=SUM({get_column_letter(value_col)}2:{get_column_letter(value_col)}{last_row-1})"

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        ws.auto_filter.ref = ws.dimensions

        # Indian number format
for row in ws.iter_rows(min_row=2):
    for cell in row:
        if isinstance(cell.value, (int, float)):
            if float(cell.value).is_integer():
                cell.value = int(cell.value)
            cell.number_format = '#,##,##0.00'

        for col in ws.columns:
            length = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = length + 3

output.seek(0)

st.download_button(
    "📥 Download Excel",
    output,
    file_name="Pending_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
# =====================================================
# DOWNLOAD SUMMARY
# =====================================================

st.markdown("---")

download_df = summary.copy()

excel_data = download_df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="📥 Download Summary",
    data=excel_data,
    file_name=f"MB52_Pending_{datetime.now().strftime('%d%m%Y')}.csv",
    mime="text/csv"
)

# =====================================================
# DOWNLOAD DETAIL (EXCEL)
# =====================================================

import io
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

output = io.BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:

    # Write data
    display_df.to_excel(writer, sheet_name="Detailed Data", index=False)

    ws = writer.sheets["Detailed Data"]

    # -----------------------------
    # TOTAL ROW
    # -----------------------------
    last_row = ws.max_row + 1

    ws.cell(row=last_row, column=1).value = "TOTAL"
    ws.cell(row=last_row, column=1).font = Font(bold=True)

    # Column T = 20 (Value in QualInsp.)
    value_col = 20
    col_letter = get_column_letter(value_col)

    ws.cell(
        row=last_row,
        column=value_col
    ).value = f"=SUM({col_letter}2:{col_letter}{last_row-1})"

    ws.cell(row=last_row, column=value_col).font = Font(bold=True)
    ws.cell(row=last_row, column=value_col).number_format = "#,##0.00"

    # -----------------------------
    # AUTO WIDTH
    # -----------------------------
    for column in ws.columns:
        max_length = 0
        letter = column[0].column_letter

        for cell in column:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[letter].width = max_length + 3

    # -----------------------------
    # NUMBER FORMAT
    # -----------------------------
    for row in ws.iter_rows(min_row=2, max_row=last_row):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                if float(cell.value).is_integer():
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "#,##0.00"

output.seek(0)

st.download_button(
    label="📥 Download Detailed Data",
    data=output,
    file_name=f"HQA_EM_Open_Receipt_{datetime.now().strftime('%d%m%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# =====================================================
# Pending Status Mail 
# =====================================================

# Calculate values
value_num = pd.to_numeric(detail["value"], errors="coerce")

mech_total = value_num[detail["Department"] == "Mechanical"].sum()
elec_total = value_num[detail["Department"] == "Electrical"].sum()
total_value = mech_total + elec_total

today = pd.Timestamp.today()

mail_text = f"""
Dear Sir,

Please find below the HQA Open Receipt pending value as of {today.strftime('%d-%b-%Y')}.

Mechanical Department: ₹ {mech_total:,.0f}
Electrical Department: ₹ {elec_total:,.0f}

Total Pending Value: ₹ {total_value:,.0f}

Regards,
HQA Team.
"""

st.subheader("📧 Mail Content")
st.markdown("""
<style>
textarea {
    font-size: 17px !important;
    font-family: Calibri, Arial, sans-serif !important;
}
</style>
""", unsafe_allow_html=True)

st.text_area(
    "📧 Copy and paste into Outlook",
    mail_text,
    height=400,
)

# Remove this line
# st.download_button("📥 Download Detailed Data", ...)

# Or restore your original download button
st.download_button(
    "📥 Download Detailed Data",
    data=excel_data,          # your Excel bytes variable
    file_name="Detailed_Data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
# =====================================================
# FOOTER
# =====================================================

st.markdown("---")

st.caption(
    f"""
HQA E&M Open Receipt Pending Dashboard

Records : {len(display_df):,}

Generated : {datetime.now().strftime('%d-%m-%Y %H:%M')}
"""
)
