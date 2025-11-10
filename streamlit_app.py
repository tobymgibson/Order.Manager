import streamlit as st
import pandas as pd
import gspread
from google.oauth2 import service_account
from datetime import date

# --------------------------
# PAGE SETUP
# --------------------------
st.set_page_config(page_title="CTI Production Dashboard", layout="wide")

# --------------------------
# GOOGLE SHEETS CONNECTION
# --------------------------
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"], scopes=scope
)
client = gspread.authorize(creds)

# Google Sheet URL and tab
sheet_url = "https://docs.google.com/spreadsheets/d/1dLHPyIkZfJby-N-EMXTraJv25BxkauBHS2ycdndC1PY"
spreadsheet = client.open_by_url(sheet_url)
worksheet = spreadsheet.worksheet("CTI")

# --------------------------
# LOAD AND CLEAN DATA
# --------------------------
data = worksheet.get_all_records()
df = pd.DataFrame(data)
df.columns = df.columns.str.strip().str.replace(" ", "_").str.replace("-", "_")

def find_column(possible_names):
    for name in possible_names:
        if name in df.columns:
            return name
    return None

machine_col = find_column(["Machine", "machine"])
customer_col = find_column(["Customer", "Customer_Name"])
feeds_col = find_column(["Feeds", "Feed"])
quantity_col = find_column(["Quantity", "Qty"])
finish_col = find_column(["Finish", "Estimated_Finish"])
next_order_col = find_column(["Next_Uncovered_Order", "Next_Shortage"])
order_value_col = find_column(["Order_Value", "Value"])

rename_map = {
    machine_col: "Machine",
    customer_col: "Customer",
    feeds_col: "Feeds",
    quantity_col: "Quantity",
    finish_col: "Finish",
    next_order_col: "Next_Uncovered_Order",
    order_value_col: "Order_Value",
}
rename_map = {k: v for k, v in rename_map.items() if k}
df.rename(columns=rename_map, inplace=True)

# Ensure numeric conversion
for col in ["Feeds", "Quantity", "Order_Value"]:
    if col in df.columns:
        df[col] = (
            pd.to_numeric(df[col].astype(str).str.replace("[^0-9.-]", "", regex=True), errors="coerce")
            .fillna(0)
        )

# Convert Finish column
if "Finish" in df.columns:
    df["Finish"] = pd.to_datetime(df["Finish"], errors="coerce")

# Drop rows with no machine
if "Machine" in df.columns:
    df = df.dropna(subset=["Machine"], how="all")

# --------------------------
# DASHBOARD METRICS
# --------------------------
st.title("🏭 CTI Production Overview")

total_feeds = int(df["Feeds"].sum()) if "Feeds" in df.columns else 0
total_value = df["Order_Value"].sum() if "Order_Value" in df.columns else 0
machines = df["Machine"].nunique() if "Machine" in df.columns else 0

col1, col2, col3 = st.columns(3)
col1.metric("Total Feeds Planned", f"{total_feeds:,}")
col2.metric("Total Order Value (£)", f"£{total_value:,.2f}")
col3.metric("Machines Active", machines)

st.divider()

# --------------------------
# MACHINE UTILISATION
# --------------------------
machine_capacity = {
    "BO1": 24000, "KO1": 50000, "KO3": 15000, "JC1": 48000, "JC": 48000, "TCY": 9000,
}
avg_util = []
if "Machine" in df.columns and "Feeds" in df.columns:
    for m, cap in machine_capacity.items():
        m_df = df[df["Machine"] == m]
        if not m_df.empty:
            avg_feeds = m_df["Feeds"].mean()
            util = (avg_feeds / cap) * 100 if cap > 0 else 0
            avg_util.append({
                "Machine": m,
                "Avg Feeds/Day": round(avg_feeds, 0),
                "Capacity (Feeds/Day)": cap,
                "Utilisation (%)": round(util, 1)
            })
avg_util_df = pd.DataFrame(avg_util)

def highlight_util(val):
    colour = "green" if val > 95 else "orange" if val > 80 else "red"
    return f"background-color: {colour}; color: white;"

if not avg_util_df.empty:
    st.subheader("⚙️ Overall Machine Utilisation Across All Planned Orders")
    st.dataframe(
        avg_util_df.style.map(highlight_util, subset=["Utilisation (%)"]),
        use_container_width=True
    )

st.divider()

# --------------------------
# FILTERS AND SORTING
# --------------------------
st.subheader("📅 Orders Scheduled to Finish in Selected Range")

if "Finish" in df.columns:
    min_date = df["Finish"].min().date() if not df["Finish"].isna().all() else date.today()
    max_date = df["Finish"].max().date() if not df["Finish"].isna().all() else date.today()

    date_selection = st.date_input(
        "Select date range:",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    if isinstance(date_selection, (list, tuple)):
        if len(date_selection) == 2:
            start_date, end_date = date_selection
        else:
            start_date = end_date = date_selection[0]
    else:
        start_date = end_date = date_selection

    date_filtered_df = df[
        (df["Finish"].dt.date >= start_date) & (df["Finish"].dt.date <= end_date)
    ].copy()
else:
    date_filtered_df = df.copy()

machines_available = sorted(df["Machine"].dropna().unique().tolist())
selected_machine = st.selectbox("Select Machine:", ["All Machines"] + machines_available)

if selected_machine != "All Machines":
    filtered_day_df = date_filtered_df[date_filtered_df["Machine"] == selected_machine]
else:
    filtered_day_df = date_filtered_df.copy()

sort_order = st.radio("Sort by Finish Date:", ["Earliest first", "Latest first"], horizontal=True)
if "Finish" in filtered_day_df.columns:
    filtered_day_df = filtered_day_df.sort_values(
        by="Finish",
        ascending=True if sort_order == "Earliest first" else False
    )

# --------------------------
# LOCAL DELETE FEATURE
# --------------------------
if "deleted_rows" not in st.session_state:
    st.session_state.deleted_rows = []

if not filtered_day_df.empty:
    filtered_day_df["Delete?"] = False  # Add checkbox column

    preferred_cols = [
        "Delete?", "Machine", "Customer", "ROW", "Feeds", "Quantity",
        "Finish", "Next_Uncovered_Order", "Order_Value"
    ]
    show_cols = [c for c in preferred_cols if c in filtered_day_df.columns]

    # Remove any rows that were previously deleted locally
    display_df = filtered_day_df[~filtered_day_df.index.isin(st.session_state.deleted_rows)]

    total_visible_value = display_df["Order_Value"].sum()
    st.markdown(f"**💰 Total Value of Orders in View:** £{total_visible_value:,.2f}")

    edited_df = st.data_editor(
        display_df[show_cols],
        use_container_width=True,
        num_rows="fixed",
        key="editable_table"
    )

    if st.button("🗑️ Delete Selected Rows (Local Only)"):
        to_delete = edited_df[edited_df["Delete?"] == True]
        if not to_delete.empty:
            st.session_state.deleted_rows.extend(to_delete.index)
            st.success("✅ Selected rows removed locally.")
            st.rerun()
        else:
            st.info("No rows selected for deletion.")

    if st.button("🔄 Reset Deleted Rows"):
        st.session_state.deleted_rows = []
        st.success("✅ All deleted rows have been restored.")
        st.rerun()
else:
    st.info("No orders found for the selected filters.")