import streamlit as st
import pandas as pd
from datetime import date, datetime
from utils import load_po_sheet
from io import BytesIO

st.set_page_config(page_title="Purchase Orders", layout="wide")
st.title("📦 Incoming Purchase Orders – Progroup POs")

# ---------------------------------------------------------
# Load PO data
# ---------------------------------------------------------
try:
    po_df = load_po_sheet()  # ✅ no URL / tab passed in
    if po_df is None or po_df.empty:
        st.warning("Could not load PO sheet or sheet is empty.")
        st.stop()
except Exception as e:
    st.error(f"Failed to load Progroup POs sheet: {e}")
    st.stop()

# ---------------------------------------------------------
# NORMALISE COLUMNS (same as before)
# ---------------------------------------------------------
po_df = po_df.copy()
po_df.columns = (
    po_df.columns.str.strip()
        .str.replace(" ", "_")
        .str.replace("-", "_")
        .str.replace(r"__+", "_", regex=True)
)

rename_map = {
    "Supplier_Name": "Supplier",
    "PO_Number": "PO_Number",
    "Acknowledge_Date": "Acknowledge_Date",
    "Supplier_Trip_No": "Supplier_Trip_No",
    "Product_Code": "Product_Code",
    "Qty_Ordered": "Qty_Ordered",
    "Qty_Delivered": "Qty_Delivered",
    "Qty_Outstanding": "Qty_Outstanding",
    "Free_Stock": "Free_Stock",
    "Orig_Due_Date": "Orig_Due_Date",
    "Current_Due_Date": "Current_Due_Date",
    "Difference": "Difference",
    "Active_Works_Orders": "Active_Works_Orders",
    "Customer_": "Customer",
    "W_O_Due_Date": "WO_Due_Date",
}

po_df = po_df.rename(columns={k: v for k, v in rename_map.items() if k in po_df.columns})

if "Customer" not in po_df.columns and "Supplier" in po_df.columns:
    po_df["Customer"] = po_df["Supplier"]

# ---------------------------------------------------------
# CLEAN NUMBERS + DATES
# ---------------------------------------------------------
def to_num(x):
    return pd.to_numeric(
        pd.Series(x, dtype="object").astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

for c in ["Qty_Ordered", "Qty_Delivered", "Qty_Outstanding", "Free_Stock", "Difference"]:
    if c in po_df.columns:
        po_df[c] = to_num(po_df[c]).fillna(0)

for d in ["Orig_Due_Date", "Current_Due_Date", "Acknowledge_Date", "WO_Due_Date"]:
    if d in po_df.columns:
        po_df[d] = pd.to_datetime(po_df[d], errors="coerce", dayfirst=True)

if "PO_Number" in po_df.columns:
    po_df = po_df.dropna(subset=["PO_Number"], how="all")

# ---------------------------------------------------------
# TOP SUMMARY METRICS
# ---------------------------------------------------------
today = date.today()

total_lines = len(po_df)
total_outstanding = po_df["Qty_Outstanding"].sum() if "Qty_Outstanding" in po_df.columns else 0
overdue = (po_df["Current_Due_Date"].dt.date < today).sum() if "Current_Due_Date" in po_df.columns else 0

c1, c2, c3 = st.columns(3)
c1.metric("PO Lines", f"{total_lines:,}")
c2.metric("Qty Outstanding", f"{int(total_outstanding):,}")
c3.metric("Overdue Lines", f"{overdue:,}")

st.subheader("🗓️ Filter Purchase Orders")

# ---------------------------------------------------------
# DATE FILTER
# ---------------------------------------------------------
if "Current_Due_Date" in po_df.columns and not po_df["Current_Due_Date"].isna().all():
    min_d, max_d = po_df["Current_Due_Date"].min().date(), po_df["Current_Due_Date"].max().date()
else:
    min_d = max_d = today

sel = st.date_input("Select Current Due Date range:", (min_d, max_d), format="DD/MM/YYYY")

s, e = sel if isinstance(sel, tuple) else (sel, sel)

po_filtered = po_df[
    (po_df["Current_Due_Date"].dt.date >= s) &
    (po_df["Current_Due_Date"].dt.date <= e)
]

# ---------------------------------------------------------
# WORKS ORDER DROPDOWN (searchable)
# ---------------------------------------------------------
if "Active_Works_Orders" in po_filtered.columns:
    options = sorted(po_filtered["Active_Works_Orders"].dropna().astype(str).unique())
    wo = st.selectbox("Filter by Works Order:", ["All Works Orders"] + options)
    if wo != "All Works Orders":
        po_filtered = po_filtered[po_filtered["Active_Works_Orders"].astype(str) == wo]

# ---------------------------------------------------------
# PRODUCT CODE DROPDOWN
# ---------------------------------------------------------
if "Product_Code" in po_filtered.columns:
    prod_list = sorted(po_filtered["Product_Code"].dropna().astype(str).unique())
    prod = st.selectbox("Filter by Product Code:", ["All Products"] + prod_list)
    if prod != "All Products":
        po_filtered = po_filtered[po_filtered["Product_Code"].astype(str) == prod]

# ---------------------------------------------------------
# TRIP NUMBER FILTER
# ---------------------------------------------------------
if "Supplier_Trip_No" in po_filtered.columns:
    trips = sorted(
    po_filtered["Supplier_Trip_No"]
    .dropna()
    .astype(str)     # convert all to strings
    .unique()
)
    trip = st.selectbox("Filter by Trip Number:", ["All Trips"] + trips)
    if trip != "All Trips":
        po_filtered = po_filtered[po_filtered["Supplier_Trip_No"] == trip]

# ---------------------------------------------------------
# TABLE COLUMNS
# ---------------------------------------------------------
show_cols = [
    "PO_Number",
    "Supplier_Trip_No",
    "Product_Code",
    "Qty_Outstanding",
    "Orig_Due_Date",
    "Current_Due_Date",
    "Difference",
    "Active_Works_Orders",
    "Customer",
    "WO_Due_Date",
]
show_cols = [c for c in show_cols if c in po_filtered.columns]

# ---------------------------------------------------------
# DATE FORMAT + COLOUR CODING
# ---------------------------------------------------------
def format_dates(df):
    df = df.copy()
    for c in ["Orig_Due_Date", "Current_Due_Date", "Acknowledge_Date", "WO_Due_Date"]:
        if c in df.columns and pd.api.types.is_datetime64_any_dtype(df[c]):
            df[c] = df[c].dt.strftime("%d/%m/%Y")
    return df

def row_colour(r):
    d = r.get("Difference", 0)
    try:
        d = float(d)
    except:
        return [""] * len(r)
    if d > 0:
        return ["background-color:#ffcccc"] * len(r)
    if d < 0:
        return ["background-color:#ccffcc"] * len(r)
    return [""] * len(r)

st.subheader("📋 Filtered Purchase Orders")
df_display = format_dates(po_filtered[show_cols])
styled = df_display.style.apply(row_colour, axis=1)
st.dataframe(styled, use_container_width=True)

# ---------------------------------------------------------
# ORDERS WITH DUE DATE CHANGES
# ---------------------------------------------------------
st.subheader("🚨 Orders with Due Date Changes")

# Filter for rows where the due date changed
diff_df = po_df[po_df["Difference"] != 0].copy()

if diff_df.empty:
    st.info("No early or late orders.")
else:
    # ----------------------------------------
    # 1️⃣ Customer filter (ONLY for diff table)
    # ----------------------------------------
    if "Customer" in diff_df.columns:
        customer_options = sorted(
            diff_df["Customer"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        selected_customers = st.multiselect(
            "Filter by Customer(s) for Due Date Changes:",
            options=customer_options,
            default=customer_options,       # show all by default
            key="diff_customer_filter"
        )

        diff_df = diff_df[
            diff_df["Customer"].astype(str).isin(selected_customers)
        ]

    # If no rows after customer filtering
    if diff_df.empty:
        st.warning("No due date changes for the selected customer(s).")
    else:

        # ----------------------------------------
        # 2️⃣ Format dates (UK format)
        # ----------------------------------------
        diff_df = format_dates(diff_df[show_cols])

        # ----------------------------------------
        # 3️⃣ Export to Excel button
        # ----------------------------------------
        from io import BytesIO
        buffer = BytesIO()

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            diff_df.to_excel(writer, index=False, sheet_name="Due_Date_Changes")
        buffer.seek(0)

        st.download_button(
            label="⬇️ Download Due Date Change Report (Excel)",
            data=buffer,
            file_name="due_date_changes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="diff_export_button"
        )

        # ----------------------------------------
        # 4️⃣ Colour-coded table output
        # ----------------------------------------
        styled_diff = diff_df.style.apply(row_colour, axis=1)
        st.dataframe(styled_diff, use_container_width=True)