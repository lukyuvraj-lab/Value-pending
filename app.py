import streamlit as st
import pandas as pd
from datetime import datetime

# -----------------------------
# Page Configuration
# -----------------------------
st.set_page_config(
    page_title="MB52 Pending Dashboard",
    page_icon="📊",
    layout="wide"
)

# -----------------------------
# Title
# -----------------------------
col1, col2 = st.columns([5, 1])

with col1:
    st.title("📊 MB52 Pending Dashboard")

with col2:
    st.markdown(
        f"### 📅 {datetime.now().strftime('%d-%m-%Y')}"
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

material_list = ["All"] + sorted(
    df["Material"].dropna().unique().tolist()
)

grn_list = ["All"] + sorted(
    df["GRN"].dropna().unique().tolist()
)

selected_plant = st.sidebar.selectbox(
    "🏭 Plant",
    plant_list
)

selected_department = st.sidebar.selectbox(
    "⚙️ Department",
    dept_list
)

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
st.subheader("📌 Dashboard Overview")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_value = filtered["Value"].sum()
grn_count = filtered["GRN"].nunique()
lot_count = len(filtered)
plant_count = filtered["Plant"].nunique() - 1

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

kpi4.metric(
    "🏭 Plants",
    plant_count)

# =====================================================
# PLANT WISE SUMMARY
# =====================================================
st.markdown("---")
st.subheader("🏭 Plant-wise Pending Value")

plant_summary = (
    filtered
    .groupby("Plant", as_index=False)["Value"]
    .sum()
    .sort_values("Value", ascending=False)
)

total_row = pd.DataFrame({
    "Plant": ["TOTAL"],
    "Value": [plant_summary["Value"].sum()]
})

plant_summary = pd.concat(
    [plant_summary, total_row],
    ignore_index=True
)

st.dataframe(
    plant_summary,
    use_container_width=True
)

# =====================================================
# DEPARTMENT SUMMARY
# =====================================================
st.markdown("---")
st.subheader("⚙️ Department-wise Pending Value")

dept_summary = (
    filtered
    .groupby("Department", as_index=False)["Value"]
    .sum()
)

st.dataframe(
    dept_summary,
    use_container_width=True
)

# =====================================================
# PLANT + DEPARTMENT SUMMARY
# =====================================================
st.markdown("---")
st.subheader("📊 Plant & Department Summary")

summary = (
    filtered
    .groupby(["Plant", "Department"])
    .agg(
        GRN_Count=("GRN", "nunique"),
        Lot_Count=("GRN", "count"),
        Pending_Value=("Value", "sum")
    )
    .reset_index()
)

total = pd.DataFrame({
    "Plant": ["TOTAL"],
    "Department": [""],
    "GRN_Count": [summary["GRN_Count"].sum()],
    "Lot_Count": [summary["Lot_Count"].sum()],
    "Pending_Value": [summary["Pending_Value"].sum()]
})

summary = pd.concat(
    [summary, total],
    ignore_index=True
)

st.dataframe(
    summary,
    use_container_width=True
)

# =====================================================
# SEARCH
# =====================================================

display_df = filtered.copy()

if search:
    display_df = display_df[
        display_df["Material"].str.contains(search, case=False, na=False)
        |
        display_df["GRN"].str.contains(search, case=False, na=False)
    ]

# =====================================================
# DETAILED DATA
# =====================================================

st.markdown("---")
st.subheader("📋 Detailed Pending Data")

# Create detail_df FIRST
detail_df = (
    display_df
    .groupby(["Plant", "Department", "GRN"], as_index=False)
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

st.dataframe(
    filtered_detail,
    use_container_width=True,
    height=500
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
# DOWNLOAD DETAIL
# =====================================================

detail_csv = display_df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="📥 Download Detailed Data",
    data=detail_csv,
    file_name=f"MB52_Detail_{datetime.now().strftime('%d%m%Y')}.csv",
    mime="text/csv"
)

# =====================================================
# FOOTER
# =====================================================

st.markdown("---")

st.caption(
    f"""
MB52 Pending Dashboard

Records : {len(display_df):,}

Generated : {datetime.now().strftime('%d-%m-%Y %H:%M')}
"""
)
