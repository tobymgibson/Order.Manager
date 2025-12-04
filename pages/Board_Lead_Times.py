import streamlit as st
import pandas as pd
from utils import load_lead_sheet

st.set_page_config(page_title="Board Lead Times", layout="wide")

st.title("📦 Board Grade Lead Times")

# Load data
try:
    lead_df = load_lead_sheet()
except Exception as e:
    st.error(f"Failed to load Board Grade Lead Times sheet: {e}")
    st.stop()

# If empty, stop
if lead_df.empty:
    st.warning("No data found in the Board Grade Lead Times sheet.")
    st.stop()

# Clean column names
lead_df.columns = (
    lead_df.columns.str.strip()
    .str.replace(" ", "_")
    .str.replace("-", "_")
)

# ================================================
# ⛔ REMOVE invalid / blank lead time rows (date version)
# ================================================
LEAD_COL = "Lead_Time"  # update this name if needed

if LEAD_COL not in lead_df.columns:
    st.error(f"Column '{LEAD_COL}' not found. Columns: {list(lead_df.columns)}")
    st.stop()

# Normalise to strings first
lead_df[LEAD_COL] = lead_df[LEAD_COL].astype(str).str.strip()

# Remove non-usable entries
invalid_values = ["", " ", "-", "N/A", "#N/A", "nan", "None"]
lead_df = lead_df[~lead_df[LEAD_COL].isin(invalid_values)]

# Convert to datetime (UK format DD/MM/YYYY)
lead_df[LEAD_COL] = pd.to_datetime(
    lead_df[LEAD_COL],
    errors="coerce",
    dayfirst=True   # 👈 UK DATE FORMAT (DD/MM/YYYY)
)

# Drop anything that couldn't convert
lead_df = lead_df.dropna(subset=[LEAD_COL])

# Drop rows where conversion failed
lead_df = lead_df.dropna(subset=[LEAD_COL])

# ================================================
# 🔍 Searchable Board Grade dropdown
# ================================================
board_column = "Board_Grade"  # Change if your sheet uses another name

if board_column not in lead_df.columns:
    st.error(f"Column '{board_column}' not found. Columns: {list(lead_df.columns)}")
    st.stop()

selected_grade = st.selectbox(
    "Search Board Grade:",
    ["All Grades"] + sorted(lead_df[board_column].dropna().unique().tolist())
)

# Apply filter
if selected_grade != "All Grades":
    lead_df = lead_df[lead_df[board_column] == selected_grade]

# Display filtered table
st.dataframe(lead_df, use_container_width=True)