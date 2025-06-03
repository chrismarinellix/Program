# app.py  – Streamlit explorer for the IFS Activity report
# Default data file:
#   C:\Reporting\Data Downloaded from IFS\reportX.xlsx

import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

DEFAULT_PATH = Path(r"C:\Reporting\Data Downloaded from IFS\reportX.xlsx")

st.set_page_config(page_title="IFS Activity Budget Explorer", layout="wide")
st.title("📊 IFS Activity Budget / Actual Explorer")

# ------------------------------------------------------------------
# 1.  Load data
# ------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_default_file(path: Path) -> pd.DataFrame:
    if path.exists():
        st.sidebar.success(f"Loaded default file:\n{path}")
        return pd.read_excel(path, engine="openpyxl")
    st.sidebar.warning(f"Default file not found:\n{path}")
    return pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_uploaded(upload):
    if upload is None:
        return pd.DataFrame()
    if upload.name.lower().endswith((".xls", ".xlsx")):
        return pd.read_excel(upload, engine="openpyxl")
    return pd.read_csv(upload)

df = load_default_file(DEFAULT_PATH)          # load default first

# ------------------------------------------------------------------
# 2.  Sidebar – optional override
# ------------------------------------------------------------------
with st.sidebar:
    st.header("📂 Data source")
    st.markdown(
        "• The app loads the default file above.  \n"
        "• Upload another file here to *replace* it in this session."
    )
    uploaded = st.file_uploader(
        "Upload Excel/CSV", type=["xlsx", "xls", "csv"], key="uploader"
    )

if uploaded:
    df = load_uploaded(uploaded)

if df.empty:
    st.stop()

# ------------------------------------------------------------------
# 3.  Dynamic filters
# ------------------------------------------------------------------
st.sidebar.header("🔎 Filter data")
filtered = df.copy()

for col in df.columns:
    col_data = df[col]

    if pd.api.types.is_numeric_dtype(col_data):
        low, high = st.sidebar.slider(
            col,
            float(col_data.min()), float(col_data.max()),
            (float(col_data.min()), float(col_data.max())),
            format="%.2f",
        )
        filtered = filtered[filtered[col].between(low, high)]

    elif pd.api.types.is_datetime64_any_dtype(col_data):
        start, end = st.sidebar.date_input(
            col,
            (col_data.min(), col_data.max()),
            min_value=col_data.min(), max_value=col_data.max(),
        )
        filtered = filtered[(col_data >= pd.to_datetime(start)) &
                            (col_data <= pd.to_datetime(end))]

    else:
        opts = st.sidebar.multiselect(col, sorted(col_data.dropna().unique()))
        if opts:
            filtered = filtered[col_data.isin(opts)]

# ------------------------------------------------------------------
# 4.  Display & download
# ------------------------------------------------------------------
st.subheader(f"Filtered results  •  {len(filtered):,} rows")

cols = st.columns(4)
if {"Estimated Cost", "Actual Cost"}.issubset(filtered.columns):
    cols[0].metric("Σ Estimated Cost",
                   f"${filtered['Estimated Cost'].sum():,.0f}")
    cols[1].metric("Σ Actual Cost",
                   f"${filtered['Actual Cost'].sum():,.0f}")
if "Budget Remaining" in filtered:
    cols[2].metric("Σ Budget Remaining",
                   f"${filtered['Budget Remaining'].sum():,.0f}")

st.dataframe(filtered, use_container_width=True, height=450)

if {"Project", "Budget Remaining"}.issubset(filtered.columns):
    st.subheader("Budget Remaining by Project (top 30)")
    chart_data = (filtered.groupby("Project", dropna=False)
                           ["Budget Remaining"]
                           .sum()
                           .sort_values(ascending=False)
                           .head(30))
    st.bar_chart(chart_data)

def to_excel_bytes(df_):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_.to_excel(writer, index=False, sheet_name="Filtered")
    return buffer.getvalue()

st.sidebar.download_button(
    "⬇️ Download filtered data (Excel)",
    data=to_excel_bytes(filtered),
    file_name="filtered_report.xlsx",
    mime=("application/vnd.openxmlformats-officedocument."
          "spreadsheetml.sheet"),
)

st.sidebar.caption("© YourCompany 2025")
