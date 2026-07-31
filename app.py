import streamlit as st
import pandas as pd

st.set_page_config(page_title="Pending Dashboard", layout="wide")

st.title("Plant Wise Pending Dashboard")

uploaded_file = st.file_uploader("Upload MB52 Excel File", type=["xlsx"])

if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)

    # Electrical Material Groups
    electrical_groups = [10096, 10097, 10098, 10103, 10104, 10113]

    # Convert columns
    df["Material Group"] = pd.to_numeric(df["Material Group"], errors="coerce")
    df["Value in QualInsp."] = pd.to_numeric(df["Value in QualInsp."], errors="coerce").fillna(0)

    # Electrical
    electrical_df = df[df["Material Group"].isin(electrical_groups)]

    # Mechanical (all remaining material groups)
    mechanical_df = df[~df["Material Group"].isin(electrical_groups)]

    # Electrical Summary
    electrical_summary = electrical_df.groupby("Plant").agg(
        Pending_GRN_Count=("GRN NO", "nunique"),
        Pending_Value=("Value in QualInsp.", "sum")
    ).reset_index()

    # Mechanical Summary
    mechanical_summary = mechanical_df.groupby("Plant").agg(
        Pending_GRN_Count=("GRN NO", "nunique"),
        Pending_Value=("Value in QualInsp.", "sum")
    ).reset_index()

    st.subheader("⚡ Electrical Pending")
    st.dataframe(electrical_summary, use_container_width=True)

    st.subheader("🔧 Mechanical Pending")
    st.dataframe(mechanical_summary, use_container_width=True)
