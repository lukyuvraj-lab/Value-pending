import streamlit as st
import pandas as pd

st.set_page_config(page_title="Pending Value Dashboard", layout="wide")

st.title("📊 Electrical Pending Value Dashboard")

uploaded_file = st.file_uploader(
    "Upload MB52 Excel File",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Electrical Material Groups
    electrical_groups = [
        10096,
        10097,
        10098,
        10103,
        10104,
        10113
    ]

    # Convert Material Group column (D) to numeric
    df.iloc[:, 3] = pd.to_numeric(df.iloc[:, 3], errors="coerce")

    # Convert Value column (T) to numeric
    df.iloc[:, 19] = pd.to_numeric(df.iloc[:, 19], errors="coerce").fillna(0)

    # Filter Electrical Material Groups
    electrical_df = df[df.iloc[:, 3].isin(electrical_groups)]

    total_value = electrical_df.iloc[:, 19].sum()

    st.metric(
        label="Electrical Pending Value",
        value=f"₹ {total_value:,.2f}"
    )

    st.subheader("Electrical Pending Records")

    st.dataframe(
        electrical_df.iloc[:, [2, 3, 8, 19]],
        use_container_width=True
    )

    csv = electrical_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Electrical Data",
        data=csv,
        file_name="electrical_pending.csv",
        mime="text/csv"
    )

else:
    st.info("Please upload the MB52 Excel file.")
