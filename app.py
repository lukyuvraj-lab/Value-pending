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

    # Plant Selection
plants = sorted(df["Plant"].dropna().unique())
selected_plant = st.selectbox("Select Plant", plants)

# Filter selected plant
elec = electrical_df[electrical_df["Plant"] == selected_plant]
mech = mechanical_df[mechanical_df["Plant"] == selected_plant]

# Calculate values
elec_value = elec["Value in QualInsp."].sum()
elec_grn = elec["GRN NO"].nunique()

mech_value = mech["Value in QualInsp."].sum()
mech_grn = mech["GRN NO"].nunique()

# Show metrics
col1, col2, col3, col4 = st.columns(4)

col1.metric("⚡ Electrical Value", f"₹ {elec_value:,.2f}")
col2.metric("⚡ Electrical GRNs", elec_grn)

col3.metric("🔧 Mechanical Value", f"₹ {mech_value:,.2f}")
col4.metric("🔧 Mechanical GRNs", mech_grn)

    # Convert columns
    df["Material Group"] = pd.to_numeric(df["Material Group"], errors="coerce")
    df["Value in QualInsp."] = pd.to_numeric(
        df["Value in QualInsp."], errors="coerce"
    ).fillna(0)

    # Electrical Material Groups
    electrical_groups = [10096, 10097, 10098, 10103, 10104, 10113]

    # Separate Departments
    electrical_df = df[df["Material Group"].isin(electrical_groups)]
    mechanical_df = df[~df["Material Group"].isin(electrical_groups)]

    # Totals
    electrical_value = electrical_df["Value in QualInsp."].sum()
    electrical_grn = electrical_df["GRN NO"].nunique()

    mechanical_value = mechanical_df["Value in QualInsp."].sum()
    mechanical_grn = mechanical_df["GRN NO"].nunique()

    # Top Metrics
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("⚡ Electrical Value", f"₹ {electrical_value:,.2f}")
    col2.metric("⚡ Electrical GRNs", electrical_grn)

    col3.metric("🔧 Mechanical Value", f"₹ {mechanical_value:,.2f}")
    col4.metric("🔧 Mechanical GRNs", mechanical_grn)

    st.divider()

    # Electrical Plant Wise Summary
    st.subheader("⚡ Electrical Plant Wise Summary")

    electrical_summary = (
        electrical_df.groupby("Plant")
        .agg(
            Pending_GRN_Count=("GRN NO", "nunique"),
            Pending_Value=("Value in QualInsp.", "sum")
        )
        .reset_index()
    )

    electrical_summary["Pending_Value"] = electrical_summary["Pending_Value"].map(
        lambda x: f"₹ {x:,.2f}"
    )

    st.dataframe(electrical_summary, use_container_width=True)

    st.divider()

    # Mechanical Plant Wise Summary
    st.subheader("🔧 Mechanical Plant Wise Summary")

    mechanical_summary = (
        mechanical_df.groupby("Plant")
        .agg(
            Pending_GRN_Count=("GRN NO", "nunique"),
            Pending_Value=("Value in QualInsp.", "sum")
        )
        .reset_index()
    )

    mechanical_summary["Pending_Value"] = mechanical_summary["Pending_Value"].map(
        lambda x: f"₹ {x:,.2f}"
    )

    st.dataframe(mechanical_summary, use_container_width=True)

else:
    st.info("Please upload the MB52 Excel file.")
