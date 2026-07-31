import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="MB52 Pending Dashboard", layout="wide")

st.title("📊 MB52 Pending Dashboard")
st.write(f"📅 Date: {datetime.now().strftime('%d-%m-%Y')}")

uploaded_file = st.file_uploader("Upload MB52 Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Columns (based on your file)
    MATERIAL = 0   # Column A
    PLANT = 2      # Column C
    GRN = 8        # Column I
    VALUE = 19     # Column T

    # Convert values
    df.iloc[:, MATERIAL] = df.iloc[:, MATERIAL].astype(str).str.strip()
    df.iloc[:, PLANT] = df.iloc[:, PLANT].astype(str).str.strip()
    df.iloc[:, VALUE] = pd.to_numeric(df.iloc[:, VALUE], errors="coerce").fillna(0)

    # Department Identification
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
        else:
            return "Mechanical"

    df["Department"] = df.iloc[:, MATERIAL].apply(get_department)

    # ---------------- Filters ----------------
    plants = sorted(df.iloc[:, PLANT].dropna().unique().tolist())
    plants.insert(0, "All Plants")

    selected_plant = st.selectbox("🏭 Select Plant", plants)

    departments = ["All", "Electrical", "Mechanical"]
    selected_department = st.selectbox("Department", departments)

    filtered = df.copy()

    if selected_plant != "All Plants":
        filtered = filtered[filtered.iloc[:, PLANT] == selected_plant]

    if selected_department != "All":
        filtered = filtered[filtered["Department"] == selected_department]

    # ---------------- Summary ----------------
    summary = (
        filtered.groupby([filtered.columns[PLANT], "Department"])[filtered.columns[VALUE]]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    if "Electrical" not in summary.columns:
        summary["Electrical"] = 0

    if "Mechanical" not in summary.columns:
        summary["Mechanical"] = 0

    summary.rename(columns={
        summary.columns[0]: "Plant",
        "Electrical": "⚡ Electrical Value",
        "Mechanical": "🔧 Mechanical Value"
    }, inplace=True)

    # Total Row
    total = pd.DataFrame({
        "Plant": ["TOTAL"],
        "⚡ Electrical Value": [summary["⚡ Electrical Value"].sum()],
        "🔧 Mechanical Value": [summary["🔧 Mechanical Value"].sum()]
    })

    summary = pd.concat([summary, total], ignore_index=True)

    st.subheader("Plant-wise Pending Value")
    st.dataframe(summary, use_container_width=True)

else:
    st.info("
