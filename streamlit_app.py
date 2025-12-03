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

PO_SHEET_URL = "https://docs.google.com/spreadsheets/d/1Ct5jXkpG6x13AnQeJ12589jj2ZCe7f3EERTU77KEYhQ"
PO_TAB_NAME = "Progroup_POs"

# =========================================================
# RENAME MAPS — MUST COME *BEFORE* LOADING SHEETS
# =========================================================

cti_rename_map = {
    "Machine": "Machine",
    "Customer Name ": "Customer",
    "ROW": "ROW",
    "Feeds": "Feeds",
    "Quantity": "Quantity",
    "Finish": "Finish",
    "Next Uncovered Order": "Next_Uncovered_Order",
    "Order Value": "Order_Value",
}

po_rename_map = {
    "Supplier_Name": "Supplier",
    "PO_Number": "PO_Number",
    "Product_Code": "Product_Code",
    "Qty_Ordered": "Qty_Ordered",
    "Qty_Delivered": "Qty_Delivered",
    "Qty_Outstanding": "Qty_Outstanding",
    "Free_Stock": "Free_Stock",
    "Orig_Due_Date": "Orig_Due_Date",
    "Current_Due_Date": "Current_Due_Date",
    "Difference": "Difference",
    "Active Works Orders": "Active_Works_Orders",   # important fix
    "Customer ": "Customer",
    "W/O_Due_Date": "WO_Due_Date",
    "Supplier_Trip_No": "Supplier_Trip_No",
}

# =========================================================
# Credentials Loader
# =========================================================
def _get_creds():
    info = dict(st.secrets["google"])
    return service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )

# =========================================================
# CTI Sheet Loader
# =========================================================
@st.cache_data(show_spinner=False, ttl=120)
def load_sheet(sheet_url: str, tab: str) -> pd.DataFrame:
    creds = _get_creds()
    client = gspread.authorize(creds)
    ws = client.open_by_url(sheet_url).worksheet(tab)
    return pd.DataFrame(ws.get_all_records())

# =========================================================
# PO Sheet Loader (detects header row)
# =========================================================
@st.cache_data(show_spinner=False, ttl=120)
def load_po_sheet(sheet_url: str, tab: str) -> pd.DataFrame:

    creds = _get_creds()
    client = gspread.authorize(creds)
    ws = client.open_by_url(sheet_url).worksheet(tab)

    rows = ws.get_all_values()
    if not rows:
        return pd.DataFrame()

    # detect header row
    header_idx = None
    for i, row in enumerate(rows):
        if "Supplier_Name" in row and "PO_Number" in row:
            header_idx = i
            break

    if header_idx is None:
        header = rows[0]
        data = rows[1:]
    else:
        header = rows[header_idx]
        data = rows[header_idx + 1:]

    return pd.DataFrame(data, columns=header)

# =========================================================
# Reload button and initial loads
# =========================================================
reload = st.sidebar.button("🔄 Refresh data")
if reload:
    st.cache_data.clear()

# Production sheet
try:
    df = load_sheet(SHEET_URL, TAB_NAME)
    st.sidebar.success("Connected to CTI Production Sheet ✅")
except Exception as e:
    st.sidebar.error(f"⚠️ Could not load data from Production Google Sheet.\n{e}")
    st.stop()

# Purchase orders sheet
po_df = None
try:
    po_df = load_po_sheet(PO_SHEET_URL, PO_TAB_NAME)
    po_df = po_df.rename(columns=po_rename_map)

    st.sidebar.success("Loaded Progroup POs successfully")
except Exception as e:
    st.sidebar.error(f"⚠️ Could not load data from Progroup POs sheet.\n{e}")
    po_df = pd.DataFrame()

    # =========================================================
# Clean and normalise CTI production data
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
# Session state (local-only deletions) for production table
# =========================================================
def build_key_col(dframe: pd.DataFrame) -> pd.Series:
    if "ROW" in dframe.columns:
        return dframe["ROW"].astype(str)
    parts = []
    for p in ["Machine", "Customer", "Finish", "Feeds", "Quantity", "Order_Value"]:
        parts.append(dframe[p].astype(str) if p in dframe.columns else "")
    return (
        parts[0]
        + "|"
        + parts[1]
        + "|"
        + parts[2]
        + "|"
        + parts[3]
        + "|"
        + parts[4]
        + "|"
        + (parts[5] if len(parts) > 5 else "")
    )

df["_RowKey"] = build_key_col(df)
if "locally_deleted_keys" not in st.session_state:
    st.session_state.locally_deleted_keys = set()

# =========================================================
# Top metrics (production)
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
# Filters (Date + Machine) - production
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
# Customer Filter (NEW)
# =========================================================

customer_col = "Customer"  # adjust if your column is named differently

if customer_col in filtered_pre_machine.columns:
    customers_available = (
        filtered_pre_machine[customer_col]
        .dropna()
        .unique()
        .tolist()
    )
    customers_available = sorted(customers_available)

    selected_customer = st.selectbox(
        "Filter by Customer",
        ["All Customers"] + customers_available
    )
else:
    selected_customer = "All Customers"

# =========================================================
# Apply filters + local deletions (production)
# =========================================================
filtered = filtered_pre_machine.copy()

if selected_machine != "All Machines" and "Machine" in filtered.columns:
    filtered = filtered[filtered["Machine"] == selected_machine]

# Apply Customer filter
if selected_customer != "All Customers" and customer_col in filtered.columns:
    filtered = filtered[filtered[customer_col] == selected_customer]    

if "Finish" in filtered.columns:
    filtered = filtered.sort_values(by="Finish", ascending=(sort_order == "Earliest first"))

if "_RowKey" in filtered.columns and st.session_state.locally_deleted_keys:
    filtered = filtered[~filtered["_RowKey"].isin(st.session_state.locally_deleted_keys)]

# =========================================================
# Risk column logic (production)
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
# Display + Local Delete (production)
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

# =========================================================
# 📦 Incoming Purchase Orders – Progroup POs
# =========================================================
st.divider()
st.header("📦 Incoming Purchase Orders – Progroup POs")

if po_df is None or po_df.empty:
    st.info("No purchase order data available (could not load PO sheet or sheet is empty).")
else:
    # -----------------------------
    # Normalise column names
    # -----------------------------
    po_df = po_df.copy()
    po_df.columns = (
        po_df.columns.str.strip()
                     .str.replace(" ", "_")
                     .str.replace("-", "_")
                     .str.replace(r"__+", "_", regex=True)
    )

    # Expected headers from your PO sheet:
    # Supplier_Name, PO_Number, Acknowledge_Date, Supplier_Trip_No,
    # Product_Code, Qty_Ordered, Qty_Delivered, Qty_Outstanding,
    # Free_Stock, Orig_Due_Date, Current_Due_Date, Difference,
    # Active Works Orders

    po_rename_map = {
    "Supplier_Name": "Supplier",
    "PO_Number": "PO_Number",
    "Product_Code": "Product_Code",
    "Qty_Ordered": "Qty_Ordered",
    "Qty_Delivered": "Qty_Delivered",
    "Qty_Outstanding": "Qty_Outstanding",
    "Free_Stock": "Free_Stock",
    "Orig_Due_Date": "Orig_Due_Date",
    "Current_Due_Date": "Current_Due_Date",
    "Difference": "Difference",
    "Active Works Orders": "Active_Works_Orders",
    "Customer ": "Customer",
    "W/O_Due_Date": "WO_Due_Date",
    "Supplier_Trip_No": "Supplier_Trip_No",
}
    po_df = po_df.rename(columns={k: v for k, v in po_rename_map.items() if k in po_df.columns})

    # Create Customer alias from Supplier if no Customer column
    if "Customer" not in po_df.columns and "Supplier" in po_df.columns:
        po_df["Customer"] = po_df["Supplier"]

    # -----------------------------
    # Clean numeric & date fields
    # -----------------------------
    def po_to_number(s):
        return pd.to_numeric(
            pd.Series(s, dtype="object").astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
            errors="coerce"
        )

    for num_col in ["Qty_Ordered", "Qty_Delivered", "Qty_Outstanding", "Free_Stock", "Difference"]:
        if num_col in po_df.columns:
            po_df[num_col] = po_to_number(po_df[num_col]).fillna(0)

    for dcol in ["Orig_Due_Date", "Current_Due_Date", "Acknowledge_Date"]:
        if dcol in po_df.columns:
            po_df[dcol] = pd.to_datetime(po_df[dcol], errors="coerce", dayfirst=True)

    # Keep rows with a PO number
    if "PO_Number" in po_df.columns:
        po_df = po_df.dropna(subset=["PO_Number"], how="all")

    # -----------------------------
    # Top-level metrics
    # -----------------------------
    total_po_lines = len(po_df)
    total_qty_outstanding = po_df["Qty_Outstanding"].sum() if "Qty_Outstanding" in po_df.columns else 0

    today = date.today()
    overdue_lines = 0
    if "Current_Due_Date" in po_df.columns:
        overdue_mask = po_df["Current_Due_Date"].notna() & (po_df["Current_Due_Date"].dt.date < today)
        overdue_lines = int(overdue_mask.sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("PO Lines", f"{total_po_lines:,}")
    c2.metric("Qty Outstanding (all lines)", f"{int(total_qty_outstanding):,}")
    c3.metric("Overdue Lines (Current Due < today)", f"{overdue_lines:,}")

    st.subheader("🗓️ Filter Purchase Orders")

    # -----------------------------
    # Date range filter (Current_Due_Date)
    # -----------------------------
    if "Current_Due_Date" in po_df.columns and not po_df["Current_Due_Date"].isna().all():
        po_min_date = po_df["Current_Due_Date"].min().date()
        po_max_date = po_df["Current_Due_Date"].max().date()
    else:
        po_min_date = po_max_date = today

    po_date_range = st.date_input(
        "Select Current Due Date range:",
        value=(po_min_date, po_max_date),
        min_value=po_min_date,
        max_value=po_max_date,
        format="DD/MM/YYYY",
        key="po_date_range",
    )

    if isinstance(po_date_range, (list, tuple)):
        if len(po_date_range) == 2:
            po_start, po_end = po_date_range
        elif len(po_date_range) == 1:
            po_start = po_end = po_date_range[0]
        else:
            po_start = po_end = po_min_date
    else:
        po_start = po_end = po_date_range

    po_filtered = po_df.copy()
    if "Current_Due_Date" in po_filtered.columns and not po_filtered["Current_Due_Date"].isna().all():
        po_filtered = po_filtered[
            (po_filtered["Current_Due_Date"].dt.date >= po_start) &
            (po_filtered["Current_Due_Date"].dt.date <= po_end)
        ]

# -----------------------------
# Works Order dropdown (searchable)
# -----------------------------
wo_col = "Active_Works_Orders"   # normalised column name

if wo_col in po_filtered.columns:
    works_orders_available = (
        po_filtered[wo_col]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    works_orders_available = sorted(works_orders_available)

    selected_wo = st.selectbox(
        "Select Works Order (searchable):",
        ["All Works Orders"] + works_orders_available,
        key="po_wo_select"
    )

    if selected_wo != "All Works Orders":
        po_filtered = po_filtered[
            po_filtered[wo_col].astype(str) == selected_wo
        ]
else:
    st.warning(f"Column '{wo_col}' not found in PO sheet.")

# -----------------------------
# Product dropdown (unchanged)
# -----------------------------
if "Product_Code" in po_filtered.columns:
    products_available = sorted(
        po_filtered["Product_Code"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected_product = st.selectbox(
        "Or select a Product Code:",
        ["All Products"] + products_available,
        key="po_product_select"
    )

    if selected_product != "All Products":
        po_filtered = po_filtered[
            po_filtered["Product_Code"].astype(str) == str(selected_product)
        ]
else:
    st.warning("Column 'Product_Code' not found in Purchase Orders sheet.")

# -----------------------------
# Trip number filter
# -----------------------------
if "Supplier_Trip_No" in po_filtered.columns:
    trip_options = sorted(po_filtered["Supplier_Trip_No"].dropna().unique().tolist())
    selected_trip = st.selectbox(
        "Filter by Trip Number:",
        ["All Trips"] + trip_options,
        key="po_trip_select"
    )
    if selected_trip != "All Trips":
        po_filtered = po_filtered[po_filtered["Supplier_Trip_No"] == selected_trip]

# -----------------------------
# Build column list for display
# -----------------------------
preferred_cols_po = [
    "PO_Number",
    "Supplier_Trip_No",
    "Product_Code",
    "Qty_Outstanding",
    "Orig_Due_Date",
    "Current_Due_Date",
    "Difference",
    "Active_Works_Orders",
    "Customer",
    "WO_Due_Date"
]
po_show_cols = [c for c in preferred_cols_po if c in po_filtered.columns]
if not po_show_cols:
    po_show_cols = po_filtered.columns.tolist()

# -----------------------------
# Date formatting helper (UK style)
# -----------------------------
def format_po_dates(dframe: pd.DataFrame) -> pd.DataFrame:
    dframe = dframe.copy()
    for dcol in ["Orig_Due_Date", "Current_Due_Date", "Acknowledge_Date"]:
        if dcol in dframe.columns and pd.api.types.is_datetime64_any_dtype(dframe[dcol]):
            dframe[dcol] = dframe[dcol].dt.strftime("%d/%m/%Y")
    return dframe

# -----------------------------
# Summary + download
# -----------------------------
po_cA, po_cB = st.columns([3, 1])
with po_cA:
    visible_outstanding = po_filtered["Qty_Outstanding"].sum() if "Qty_Outstanding" in po_filtered.columns else 0
    st.markdown(f"**📊 Outstanding Qty in View:** {int(visible_outstanding):,}")

with po_cB:
    st.download_button(
        label="⬇️ Download PO view (CSV)",
        data=po_filtered[po_show_cols].to_csv(index=False).encode("utf-8"),
        file_name="purchase_orders_filtered_view.csv",
        mime="text/csv",
        use_container_width=True,
        key="po_download_btn",
    )

# -----------------------------
# Final display – always show table
# -----------------------------
if not po_filtered.empty:
    display_po = format_po_dates(po_filtered[po_show_cols])
    st.dataframe(display_po, use_container_width=True)
else:
    st.warning("No purchase orders match the current filters – showing all POs instead.")
    all_display_po = format_po_dates(po_df[po_show_cols])
    st.dataframe(all_display_po, use_container_width=True)

    # =========================================================
    # 🚨 Orders with Due Date Changes (Early / Late)
    # =========================================================
    st.subheader("🚨 Orders with Due Date Changes")

    if "Difference" not in po_df.columns:
        st.info("No 'Difference' column found in the PO data.")
    else:
        diff_df = po_df.copy()
        diff_df = diff_df[diff_df["Difference"] != 0]

        if diff_df.empty:
            st.info("No orders have a due date change (Difference = 0 for all rows).")
        else:
            def describe_diff(d):
                try:
                    d_int = int(d)
                except Exception:
                    return ""
                if d_int > 0:
                    return f"Late by {d_int} day(s)"
                elif d_int < 0:
                    return f"Early by {abs(d_int)} day(s)"
                return ""

            diff_df = diff_df.copy()
            diff_df["Due_Date_Change"] = diff_df["Difference"].apply(describe_diff)

            diff_cols_preferred = [
                "PO_Number",
                "Product_Code",
                "Qty_Outstanding",
                "Orig_Due_Date",
                "Current_Due_Date",
                "Difference",
                "Due_Date_Change",
                "Active Works Orders",
                "Customer",
                "W/O_Due_Date"
            ]
            diff_show_cols = [c for c in diff_cols_preferred if c in diff_df.columns]
            if not diff_show_cols:
                diff_show_cols = diff_df.columns.tolist()

            diff_display = format_po_dates(diff_df[diff_show_cols])

            def colour_row_by_difference(row):
                diff_val = row.get("Difference", None)
                try:
                    diff_val = float(diff_val)
                except Exception:
                    return [""] * len(row)
                if diff_val > 0:
                    return ["background-color: #ffcccc"] * len(row)  # late: light red
                elif diff_val < 0:
                    return ["background-color: #ccffcc"] * len(row)  # early: light green
                else:
                    return [""] * len(row)

            styled_diff = diff_display.style.apply(colour_row_by_difference, axis=1)
            st.dataframe(styled_diff, use_container_width=True)