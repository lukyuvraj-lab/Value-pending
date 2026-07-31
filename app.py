import streamlit as st
import pandas as pd

st.set_page_config(page_title="Electrical Pending Dashboard", layout="wide")

st.title("📊 Electrical Pending Dashboard")

uploaded_file = st.file_uploader(
    "Upload MB52 Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Electrical Material Groups
    electrical_groups = [10096, 10097, 10098, 10103, 10104, 10113]

    # D = Material Group
    df.iloc[:, 3] = pd.to_numeric(df.iloc[:, 3], errors="coerce")

    # T = Pending Value
    df.iloc[:, 19] = pd.to_numeric(df.iloc[:, 19], errors="coerce").fillna(0)

    # Filter Electrical Material Groups
    electrical_df = df[df.iloc[:, 3].isin(electrical_groups)]

    # Total Pending Value
    total_value = electrical_df.iloc[:, 19].sum()

    st.metric(
        "Total Electrical Pending Value",
        f"₹ {total_value:,.2f}"
    )

    # Plant-wise Summary
    summary = (
        electrical_df
        .groupby(df.columns[2])  # Column C = Plant
        .agg(
            Pending_GRN_Count=(df.columns[8], "count"),   # Column I = GRN No
            Pending_Value=(df.columns[19], "sum")         # Column T = Value
        )
        .reset_index()
    )

    st.subheader("🏭 Plant-wise Pending Summary")
    st.dataframe(summary, use_container_width=True)

else:
    st.info("Please upload the MB52 Excel file.")
