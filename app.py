import streamlit as st
import pandas as pd

st.set_page_config(page_title="Pending Dashboard", layout="wide")

st.title("📊 Plant Wise Pending Dashboard")

uploaded_file = st.file_uploader(
    "Upload MB52 Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Convert required columns
    df["Material Group"] = pd.to_numeric(df["Material Group"], errors="coerce")
    df["Value in QualInsp."] = pd.to_numeric(
        df["Value in QualInsp."], errors="coerce"
    ).fillna(0)

    # Electrical Material Groups
    electrical_groups = [10096, 10097, 10098, 10103, 10104, 10113]

    # Separate Electrical & Mechanical
    electrical_df = df[df["Material Group"].isin(electrical_groups)]
    mechanical_df = df[~df["Material Group"].isin(electrical_groups)]

    # Plant Selection
    plants = ["All Plants"] + sorted(df["Plant"].dropna().astype(str).unique())
    selected_plant = st.selectbox("🏭 Select Plant", plants)

    if selected_plant != "All Plants":
        electrical_df = electrical_df[
            electrical_df["Plant"].astype(str) == selected_plant
        ]
        mechanical_df = mechanical_df[
            mechanical_df["Plant"].astype(str) == selected_plant
        ]

    # Totals
    electrical_value = electrical_df["Value in QualInsp."].sum()
    electrical_grn = electrical_df["GRN NO"].nunique()

    mechanical_value = mechanical_df["Value in QualInsp."].sum()
    mechanical_grn = mechanical_df["GRN NO"].nunique()

    # Top Metrics
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("⚡ Electrical Value", f"₹ {electrical_value:,.2f}")
    col2.metric("⚡ Electrical GRNs", electrical_grn)
    col3.metric("🔧 Mechanical Value",
