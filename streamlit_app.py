import streamlit as st
import pandas as pd
import re

# --- Machine capacities (feeds per day) ---
MACHINE_CAPACITY = {
    "BO1": 24000,   # Century
    "KO1": 50000,   # Klett1
    "JC1": 48000,   # JinChang
    "TCY": 9000,    # TCY
    "KO3": 15000,   # Klett3
}

# --- Optional chart support ---
try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# --- Streamlit setup ---
st.set_page_config(page_title="🏭 Production Dashboard", layout="wide")
st.title("🏭 Production Planning Dashboard")
st.caption("Upload your Excel (.xlsx) file to view KPIs, utilization, and planned orders by machine.")

# --- File upload ---
uploaded_file = st.file_uploader("📁 Upload Excel (.xlsx) or CSV", type=["xlsx", "csv"])

if uploaded_file:
    # --- Read file ---
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            # Automatically load the first sheet in the workbook
            excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
            first_sheet = excel_file.sheet_names[0]
            df = pd.read_excel(excel_file, sheet_name=first_sheet)
    except Exception as e:
        st.error(f"❌ Could not read file: {e}")
        st.stop()

    # --- Clean headers ---
    df.columns = df.columns.astype(str).str.strip()

    # --- Helper to find columns by keyword ---
    def find_col(keywords):
        for col in df.columns:
            for key in keywords:
                if key.lower() in col.lower():
                    return col
        return None

    # --- Identify columns ---
    col_machine = find_col(["machine"])
    col_customer = find_col(["customer"])
    col_row = find_col(["row"])
    col_qty = find_col(["quantity"])
    col_wo = find_col(["works order", "w/o"])
    col_overall_qty = find_col(["overall quantity", "quantity.1"])
    col_finish = find_col(["finish"])
    col_earliest = find_col(["earliest delivery"])
    col_next = find_col(["next uncovered order"])
    col_status = find_col(["status"])
    col_decision = find_col(["run decision"])
    col_value = find_col(["order value", "value"])

    # --- Rename consistently ---
    rename_map = {
        col_machine: "Machine",
        col_customer: "Customer",
        col_row: "ROW",
        col_qty: "Quantity",
        col_wo: "Works_Order",
        col_overall_qty: "Overall_Quantity",
        col_finish: "Estimated_Finish",
        col_earliest: "Earliest_Delivery",
        col_next: "Next_Uncovered_Order",
        col_status: "Status",
        col_decision: "Run_Decision",
        col_value: "Order_Value",
    }
    rename_map = {k: v for k, v in rename_map.items() if k is not None}
    df.rename(columns=rename_map, inplace=True)

    # --- Fill down machine names ---
    if "Machine" in df.columns:
        df["Machine"] = df["Machine"].ffill()

    # --- Drop filler rows (if empty) ---
    subset = [c for c in ["Customer", "Quantity"] if c in df.columns]
    if subset:
        df = df.dropna(subset=subset, how="all")

    # --- Convert numeric columns ---
    for col in ["Quantity", "Overall_Quantity", "Order_Value"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- Parse date columns ---
    for col in ["Estimated_Finish", "Earliest_Delivery"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # --- Extract sales order dates from text ---
    def extract_date(text):
        if not isinstance(text, str):
            return pd.NaT
        m = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", text)
        return pd.to_datetime(m.group(1), dayfirst=True, errors="coerce") if m else pd.NaT

    if "Next_Uncovered_Order" in df.columns:
        df["Sales_Order_Date"] = df["Next_Uncovered_Order"].apply(extract_date)
    if "Earliest_Delivery" in df.columns:
        df["Sales_Order_Date"] = df["Sales_Order_Date"].fillna(df["Earliest_Delivery"])

    machine_col = "Machine" if "Machine" in df.columns else None

    # --- Sidebar filters ---
    st.sidebar.header("🔍 Filters")
    if "Status" in df.columns:
        statuses = sorted(df["Status"].dropna().unique())
        sel_status = st.sidebar.multiselect("Status", statuses, default=statuses)
        df = df[df["Status"].isin(sel_status)]
    if "Run_Decision" in df.columns:
        decisions = sorted(df["Run_Decision"].dropna().unique())
        sel_dec = st.sidebar.multiselect("Run Decision", decisions, default=decisions)
        df = df[df["Run_Decision"].isin(sel_dec)]
    if machine_col:
        machines = sorted(df[machine_col].dropna().unique())
        sel_mach = st.sidebar.multiselect("Machine(s)", machines, default=machines)
        df = df[df[machine_col].isin(sel_mach)]

    # --- KPIs ---
    st.subheader("📊 Key Metrics")
    total_orders = len(df)
    total_value = df["Order_Value"].sum(skipna=True) if "Order_Value" in df.columns else 0
    total_qty = df["Quantity"].sum(skipna=True) if "Quantity" in df.columns else 0
    total_overall_qty = df["Overall_Quantity"].sum(skipna=True) if "Overall_Quantity" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Orders", total_orders)
    c2.metric("Total Value", f"£{total_value:,.0f}")
    c3.metric("Line Qty", f"{int(total_qty):,}")
    c4.metric("Overall Qty", f"{int(total_overall_qty):,}")

    # --- Utilization ---
    st.subheader("⚙️ Estimated Machine Utilization")

    if machine_col:
        util_data = []

        # Detect correct quantity column
        qty_col = None
        for candidate in ["Overall_Quantity", "Quantity.1", "Overall quantity", "Quantity"]:
            if candidate in df.columns:
                qty_col = candidate
                break

        for m in sorted(df[machine_col].dropna().unique()):
            m_code = str(m).strip().upper()
            cap = MACHINE_CAPACITY.get(m_code, None)
            planned = df.loc[df[machine_col].str.upper() == m_code, qty_col].sum() if qty_col else 0

            if cap:
                util = (planned / cap) * 100
                util_data.append((m_code, cap, planned, round(util, 1)))
            else:
                util_data.append((m_code, "N/A", planned, None))

        util_df = pd.DataFrame(util_data, columns=["Machine", "Feeds/Day", "Planned Qty", "Utilization (%)"])

        def color_util(val):
            if pd.isna(val):
                return ""
            if val > 100:
                return "background-color:#ff4d4d;color:white"
            elif val >= 70:
                return "background-color:#ffd633"
            else:
                return "background-color:#85e085"

        st.dataframe(util_df.style.applymap(color_util, subset=["Utilization (%)"]),
                     use_container_width=True)

        if PLOTLY_AVAILABLE:
            fig = px.bar(util_df, x="Machine", y="Utilization (%)", text_auto=True,
                         color="Utilization (%)", color_continuous_scale=["green", "yellow", "red"],
                         range_y=[0, 120])
            st.plotly_chart(fig, use_container_width=True)

        # --- Machine-Specific Planned Orders ---
        st.subheader("🧾 Machine-Specific Planned Orders")
        selected_machine = st.selectbox("Select a Machine", ["All Machines"] + sorted(df[machine_col].dropna().unique()))
        subset = df if selected_machine == "All Machines" else df[df[machine_col] == selected_machine]
        if not subset.empty:
            show_cols = [c for c in ["Machine", "Customer", "ROW", "Works_Order", "Quantity", "Overall_Quantity",
                                     "Status", "Run_Decision", "Order_Value", "Estimated_Finish",
                                     "Earliest_Delivery", "Next_Uncovered_Order"] if c in subset.columns]
            st.dataframe(subset[show_cols], use_container_width=True)
        else:
            st.info(f"No orders found for {selected_machine}.")
    else:
        st.warning("⚠️ No 'Machine' column detected in your file. Ensure Column A contains machine codes (BO1, KO1, etc.).")

    # --- Orders over time ---
    if "Sales_Order_Date" in df.columns:
        st.subheader("📅 Orders Over Time")
        df_sorted = df.sort_values("Sales_Order_Date")

        value_col = None
        for candidate in ["Order_Value", "Value", "Order Value"]:
            if candidate in df_sorted.columns:
                value_col = candidate
                break

        if value_col:
            if PLOTLY_AVAILABLE:
                fig2 = px.line(
                    df_sorted,
                    x="Sales_Order_Date",
                    y=value_col,
                    color=machine_col if machine_col in df.columns else None,
                    markers=True,
                    title="Orders Over Time"
                )
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.line_chart(df_sorted.set_index("Sales_Order_Date")[value_col])
        else:
            st.info("No value column found to plot orders over time.")

    # --- Download button ---
    st.download_button("⬇️ Download Filtered Data", df.to_csv(index=False).encode("utf-8"),
                       "filtered_orders.csv", "text/csv")

else:
    st.info("👆 Upload your Excel or CSV file to begin.")