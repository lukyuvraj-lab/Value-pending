import streamlit as st
import pandas as pd

st.title("📊 MB52 Open Receipts Dashboard")

uploaded_file = st.file_uploader("Upload MB52 Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    electrical_groups = [10096, 10097, 10098, 10103, 10104, 10113]

    electrical = df[df.iloc[:, 3].isin(electrical_groups)]["Value in QualInsp"].sum()
    mechanical = df[~df.iloc[:, 3].isin(electrical_groups)]["Value in QualInsp"].sum()

    col1, col2 = st.columns(2)

    with col1:
        st.metric("🏭 Mechanical Pending", f"₹{mechanical:,.2f}")

    with col2:
        st.metric("⚡ Electrical Pending", f"₹{electrical:,.2f}")

    st.subheader("Uploaded Data")
    st.dataframe(df)
