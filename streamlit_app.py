import streamlit as st
import pandas as pd
import re
from datetime import date, datetime
import gspread
from google.oauth2 import service_account

# =========================================================
# Page config
# =========================================================
st.set_page_config(page_title="CTI Production Dashboard", layout="wide")

# =========================================================
# Google Sheets connection
# =========================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dLHPyIkZfJby-N-EMXTraJv25BxkauBHS2ycdndC1PY"
TAB_NAME = "CTI"

@st.cache_data(show_spinner=False, ttl=120)
def load_sheet(_sheet_url: str, _tab: str) -> pd.DataFrame:
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    client = gspread.authorize(creds)
    ss = client.open_by_url(_sheet_url)
    ws = ss.worksheet(_tab)
    data = ws.get_all_records()
    return pd.DataFrame(data)

reload = st.sidebar.button("🔄 Refresh data")
if reload:
    st.cache_data.clear()

try:
    df = load_sheet(SHEET_URL, TAB_NAME)
    st.sidebar.success("Connected to Google Sheets ✅")
except Exception as e:
    st.sidebar.error(f"⚠️ Could not load data from Google Sheets.\n{e}")
    st.stop()

# =========================================================
# Clean and normalise data
# =========================================================
df.columns = (
    df.columns.str.strip()
              .str.replace(" ", "_")
              .str.replace("-", "_")
              .str.replace(r"__+", "_", regex=True)
)

def pick_col(dframe: pd.DataFrame, options):
    for name in options:
        if name in dframe.columns:
            return name
    return None

col_machine = pick_col(df, ["Machine", "machine"])
col_customer = pick_col(df, ["Customer", "Customer_Name", "Cust_name"])
col_row = pick_col(df, ["ROW", "Row", "Spec_Number"])
col_feeds = pick_col(df, ["Feeds", "Feed", "feeds"])
col_qty = pick_col(df, ["Quantity", "Qty", "quantity"])
col_finish = pick_col(df, ["Finish", "Estimated_Finish", "finish"])
col_next = pick_col(df, ["Next_Uncovered_Order", "Next_Shortage", "NextUncoveredOrder"])
col_value = pick_col(df, ["Order_Value", "OrderValue", "Order_val", "Value", "value"])

rename_map = {
    col_machine: "Machine",
    col_customer: "Customer",
    col_row: "ROW",
    col_feeds: "Feeds",
    col_qty: "Quantity",
    col_finish: "Finish",
    col_next: "Next_Uncovered_Order",
    col_value: "Order_Value",
}
rename_map = {k: v for k, v in rename_map.items() if k}
df = df.rename(columns=rename_map)

def to_number(s):
    return pd.to_numeric(
        pd.Series(s, dtype="object").astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

for col in ["Feeds", "Quantity", "Order_Value"]:
    if col in df.columns:
        df[col] = to_number(df[col]).fillna(0)

# --- Force UK-style (day-first) date parsing ---
if "Finish" in df.columns:
    df["Finish"] = pd.to_datetime(df["Finish"], errors="coerce", dayfirst=True)

if "Machine" in df.columns:
    df = df.dropna(subset=["Machine"], how="all")

# =========================================================
# Session state (local-only deletions)
# =========================================================
def build_key_col(dframe: pd.DataFrame) -> pd.Series:
    if "ROW" in dframe.columns:
        return dframe["ROW"].astype(str)
    parts = []
    for p in ["Machine", "Customer", "Finish", "Feeds", "Quantity", "Order_Value"]:
        parts.append(dframe[p].astype(str) if p in dframe.columns else "")
    return (parts[0] + "|" + parts[1] + "|" + parts[2] + "|" + parts[3] + "|" + parts[4] + "|" + (parts[5] if len(parts) > 5 else ""))

df["_RowKey"] = build_key_col(df)
if "locally_deleted_keys" not in st.session_state:
    st.session_state.locally_deleted_keys = set()

# =========================================================
# Top metrics
# =========================================================
st.title("🏭 CTI Production Dashboard")

total_orders = len(df)
total_value = df["Order_Value"].sum() if "Order_Value" in df.columns else 0
machine_count = df["Machine"].nunique() if "Machine" in df.columns else 0

m1, m2, m3 = st.columns(3)
m1.metric("Total Orders", f"{total_orders:,}")
m2.metric("Total Order Value", f"£{total_value:,.2f}")
m3.metric("Machines Active", machine_count)

st.divider()

# =========================================================
# Overall Machine Utilisation
# =========================================================
st.subheader("⚙️ Overall Machine Utilisation Across All Planned Orders")

machine_capacity = {
    "BO1": 24000,
    "KO1": 50000,
    "KO3": 15000,
    "JC1": 48000,
    "JC": 48000,
    "TCY": 9000,
}

util_df = pd.DataFrame(columns=["Machine", "Avg_Feeds_per_Day", "Capacity_Feeds_per_Day", "Utilisation_%"])

if {"Machine", "Feeds", "Finish"}.issubset(df.columns):
    day_df = df.dropna(subset=["Finish"]).copy()
    day_df["Finish_Date"] = day_df["Finish"].dt.date
    group = day_df.groupby(["Machine", "Finish_Date"], as_index=False)["Feeds"].sum()

    rows = []
    for m, cap in machine_capacity.items():
        if m == "JC1":
            m_slice = group[group["Machine"].isin(["JC1", "JC"])]
            label = "JC1"
        elif m == "JC":
            continue
        else:
            m_slice = group[group["Machine"] == m]
            label = m

        if not m_slice.empty:
            avg_daily = m_slice["Feeds"].mean()
            util = (avg_daily / cap * 100.0) if cap > 0 else 0
            rows.append({
                "Machine": label,
                "Avg_Feeds_per_Day": round(avg_daily, 0),
                "Capacity_Feeds_per_Day": cap,
                "Utilisation_%": round(util, 1)
            })
    if rows:
        util_df = pd.DataFrame(rows)

def colour_util(val):
    if pd.isna(val):
        return ""
    if val >= 95:
        return "background-color: green; color: white;"
    if val >= 80:
        return "background-color: orange; color: white;"
    return "background-color: red; color: white;"

if not util_df.empty:
    st.dataframe(util_df.style.map(colour_util, subset=["Utilisation_%"]), use_container_width=True)
else:
    st.info("No utilisation data available (check Machine / Feeds / Finish columns).")

st.divider()

# =========================================================
# Filters (Date + Machine)
# =========================================================
st.subheader("📅 Orders Scheduled to Finish in Selected Range")

if "Finish" in df.columns:
    min_date = df["Finish"].min().date() if not df["Finish"].isna().all() else date.today()
    max_date = df["Finish"].max().date() if not df["Finish"].isna().all() else date.today()
else:
    min_date = max_date = date.today()

date_selection = st.date_input(
    "Select date range:",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
    format="DD/MM/YYYY"
)

if isinstance(date_selection, (list, tuple)):
    if len(date_selection) == 2:
        start_date, end_date = date_selection
    elif len(date_selection) == 1:
        start_date = end_date = date_selection[0]
    else:
        start_date = end_date = min_date
else:
    start_date = end_date = date_selection

# Filter by selected date range first
filtered_pre_machine = df.copy()
if "Finish" in filtered_pre_machine.columns and not filtered_pre_machine["Finish"].isna().all():
    filtered_pre_machine = filtered_pre_machine[
        (filtered_pre_machine["Finish"].dt.date >= start_date) &
        (filtered_pre_machine["Finish"].dt.date <= end_date)
    ]

# Only show machines that have data in this range
machines_available = sorted(filtered_pre_machine["Machine"].dropna().unique().tolist()) if "Machine" in filtered_pre_machine.columns else []
selected_machine = st.selectbox("Filter by Machine", ["All Machines"] + machines_available)

sort_order = st.radio("Sort by Finish Date:", ["Earliest first", "Latest first"], horizontal=True)

# =========================================================
# Apply filters + local deletions
# =========================================================
filtered = filtered_pre_machine.copy()

if selected_machine != "All Machines" and "Machine" in filtered.columns:
    filtered = filtered[filtered["Machine"] == selected_machine]

if "Finish" in filtered.columns:
    filtered = filtered.sort_values(by="Finish", ascending=(sort_order == "Earliest first"))

if "_RowKey" in filtered.columns and st.session_state.locally_deleted_keys:
    filtered = filtered[~filtered["_RowKey"].isin(st.session_state.locally_deleted_keys)]

# =========================================================
# Risk column logic
# =========================================================
def parse_next_shortage(text: str):
    if not isinstance(text, str) or text.strip() == "":
        return ("Unknown", None, None)
    lower = text.lower()
    if "all covered" in lower:
        return ("Low", None, None)
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", text)
    if m:
        try:
            d = datetime.strptime(m.group(0), "%d/%m/%Y").date()
            days = (d - date.today()).days
            if days <= 3:
                return ("High", d, days)
            else:
                return ("Medium", d, days)
        except Exception:
            return ("Unknown", None, None)
    return ("Unknown", None, None)

def risk_badge(risk_tuple):
    level, d, days = risk_tuple
    if level == "Low":
        return "🟢 All covered"
    if level == "High":
        return f"🔴 Next shortage ≤ 3 days ({days}d)"
    if level == "Medium":
        return f"🟠 Shortage in {days}d" if days is not None else "🟠 Shortage upcoming"
    return "⚪ Unknown"

if "Next_Uncovered_Order" in filtered.columns:
    risks = filtered["Next_Uncovered_Order"].apply(parse_next_shortage)
    filtered = filtered.assign(Risk=risks.apply(risk_badge))

# =========================================================
# Display + Local Delete
# =========================================================
preferred_cols = ["Machine", "Customer", "ROW", "Feeds", "Quantity", "Finish", "Next_Uncovered_Order", "Risk", "Order_Value"]
show_cols = [c for c in preferred_cols if c in filtered.columns]

visible_total_value = filtered["Order_Value"].sum() if "Order_Value" in filtered.columns else 0.0
cA, cB = st.columns([3, 1])
with cA:
    st.markdown(f"**💰 Total Value of Orders in View:** £{visible_total_value:,.2f}")
with cB:
    st.download_button(
        label="⬇️ Download current view (CSV)",
        data=filtered[show_cols].to_csv(index=False).encode("utf-8"),
        file_name="orders_filtered_view.csv",
        mime="text/csv",
        use_container_width=True,
    )

if not filtered.empty:
    display_df = filtered.copy()
    display_df["Delete?"] = False
    cols_order = (["Delete?"] + show_cols)
    cols_order = [c for c in cols_order if c in display_df.columns]

    st.caption("Tick rows to hide locally, then click **Delete Selected (Local Only)**.")
    edited = st.data_editor(
        display_df[cols_order],
        use_container_width=True,
        num_rows="fixed",
        key="orders_editor",
        hide_index=True
    )

    to_delete_keys = set()
    if "Delete?" in edited.columns:
        edited_key = build_key_col(edited.rename(columns={"_RowKey": "_RowKey"}))
        mask = (edited["Delete?"] == True)
        if mask.any():
            to_delete_keys = set(edited_key[mask].astype(str).tolist())

    b1, b2 = st.columns(2)
    with b1:
        if st.button("🗑️ Delete Selected (Local Only)"):
            if to_delete_keys:
                st.session_state.locally_deleted_keys.update(to_delete_keys)
                st.success(f"Removed {len(to_delete_keys)} row(s) locally.")
                st.rerun()
            else:
                st.info("No rows selected for deletion.")
    with b2:
        if st.button("🔄 Reset Deleted Rows"):
            st.session_state.locally_deleted_keys = set()
            st.success("All locally deleted rows have been restored.")
            st.rerun()
else:
    st.info("No orders match the current filters.")

st.markdown(
    "<div style='color:#888;margin-top:12px;'>Note: Deletions are local only. Refresh or click “Reset Deleted Rows” to restore.</div>",
    unsafe_allow_html=True
)