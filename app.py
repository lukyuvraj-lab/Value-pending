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
CURRENCY = 13
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
    .fillna("")
    .astype(str)
    .str.strip()
)

df["Plant"] = (
    df.iloc[:, PLANT]
    .fillna("")
    .astype(str)
    .str.strip()
)

df["GRN"] = (
    df.iloc[:, GRN]
    .fillna("")
    .astype(str)
    .str.strip()
)

df["Currency"] = (
    df.iloc[:, CURRENCY]
    .fillna("")
    .astype(str)
    .str.strip()
)

df["Value"] = pd.to_numeric(
    df.iloc[:, VALUE],
    errors="coerce"
).fillna(0)

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

selected_material = st.sidebar.selectbox(
    "📦 Material",
    material_list
)

selected_grn = st.sidebar.selectbox(
    "📄 GRN",
    grn_list
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

if selected_material != "All":
    filtered = filtered[
        filtered["Material"] == selected_material
    ]

if selected_grn != "All":
    filtered = filtered[
        filtered["GRN"] == selected_grn
    ]
