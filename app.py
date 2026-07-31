import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="MB52 Pending Dashboard", layout="wide")

st.title("📊 MB52 Pending Dashboard")
st.write(f"📅 Date: {datetime.now().strftime('%d-%m-%Y')}")

uploaded_file = st.file_uploader(
    "Upload MB52 Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    if df.empty:
        st.error("Uploaded file is empty.")
        st.stop()

    # Column Positions
    MATERIAL = 0      # Column A
    PLANT = 2         # Column C
    GRN = 8           # Column I
    VALUE = 19        # Column T

    # Read columns
    material = df.iloc[:, MATERIAL].fillna("").astype(str).str.strip()
    plant = df.iloc[:, PLANT].fillna("").astype(str).str.strip()
    value = pd.to_numeric(
        df.iloc[:, VALUE],
        errors="coerce"
    ).fillna(0)

    # Store back in dataframe
    df["Plant"] = plant
    df["Value"] = value

    # Department Function
    def get_department(mat):
        mat = str(mat)

        if (
            mat.startswith("1000")
            or mat.startswith("385")
            or mat.startswith("44")
            or mat.startswith("45")
            or mat.startswith("46")
            or mat.startswith("485")
            or mat.startswith("63")
        ):
            return "Electrical"
        else:
            return "Mechanical"

    df["Department"] = material.apply(get_department)

    # Filters
    plants = ["All Plants"] + sorted(df["Plant"].dropna().unique().tolist())

    selected_plant = st.selectbox(
        "🏭 Select Plant",
        plants
    )

    selected_dept = st.selectbox(
        "Department",
        ["All", "Electrical", "Mechanical"]
    )

    filtered = df.copy()

    if selected_plant != "All Plants":
        filtered = filtered[
            filtered["Plant"] == selected_plant
        ]

    if selected_dept != "All":
        filtered = filtered[
            filtered["Department"] == selected_dept
        ]

    # Summary
    summary = (
        filtered.groupby(["Plant", "Department"])["Value"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    if "Electrical" not in summary.columns:
        summary["Electrical"] = 0

    if "Mechanical" not in summary.columns:
        summary["Mechanical"] = 0

    summary.rename(
        columns={
            "Electrical": "⚡ Electrical Value",
            "Mechanical": "🔧 Mechanical Value",
        },
        inplace=True,
    )

    total = pd.DataFrame({
        "Plant": ["TOTAL"],
        "⚡ Electrical Value": [summary["⚡ Electrical Value"].sum()],
        "🔧 Mechanical Value": [summary["🔧 Mechanical Value"].sum()],
    })

    summary = pd.concat([summary, total], ignore_index=True)

    st.subheader("Plant-wise Pending Value")
    st.dataframe(summary, use_container_width=True)

else:
    st.info("Please upload the MB52 Excel file.")
