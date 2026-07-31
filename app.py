import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="MB52 Pending Dashboard", layout="wide")

st.title("📊 MB52 Pending Dashboard")
st.write("📅 Date :", date.today().strftime("%d-%m-%Y"))

uploaded_file = st.file_uploader("Upload MB52 Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)

    # Convert required columns
    df["Material"] = df["Material"].astype(str)
    df["Value in QualInsp."] = pd.to_numeric(
        df["Value in QualInsp."], errors="coerce"
    ).fillna(0)

    # Function to identify department
    def department(material):
        material = str(material).strip()

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
        else:
            return "Mechanical"

    # Create Department column
    df["Department"] = df["Material"].apply(department)

    # Filters
    plants = ["All Plants"] + sorted(df["Plant"].astype(str).unique())
    selected_plant = st.selectbox("🏭 Select Plant", plants)

    departments = ["All", "Electrical", "Mechanical"]
    selected_dept = st.selectbox("Department", departments)

    filtered = df.copy()

    if selected_plant != "All Plants":
        filtered = filtered[filtered["Plant"].astype(str) == selected_plant]

    if selected_dept != "All":
        filtered = filtered[filtered["Department"] == selected_dept]

    # Plant-wise Summary
    summary = (
        filtered.pivot_table(
            index="Plant",
            columns="Department",
            values="Value in QualInsp.",
            aggfunc="sum",
            fill_value=0,
        )
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

    st.subheader("Plant Wise Pending Value")
    st.dataframe(summary, use_container_width=True)

else:
    st.info("Upload the MB52 Excel file.")
