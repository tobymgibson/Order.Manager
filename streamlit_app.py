import streamlit as st
import pandas as pd
import re
from datetime import datetime

# --- Machine capacities (feeds per day) ---
MACHINE_CAPACITY = {
    "BO1": 24000,   # Century
    "KO1": 50000,   # Klett 1
    "JC1": 48000,   # Jin Chang
    "TCY": 9000,    # TCY
    "KO3": 15000,   # Klett 3
}

# --- Alternate machine name mappings ---
MACHINE_ALIASES = {
    "JC": "JC1",
    "BO": "BO1",
    "KLETT": "KO1",
    "KLETT3": "KO3",
}

# --- Optional chart support ---
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# --- Streamlit setup ---
st.set_page_config(page_title="🏭 Production Dashboard", layout="wide")
st.title("🏭 Production Planning Dashboard")
st.caption(f"🕒 Last refreshed: {datetime.now():%d %b %Y, %H:%M}")
st.caption("Upload your Excel (.xlsx) file to view KPIs, utilisation, and planned orders by machine.")

# --- File upload ---
uploaded_file = st.file_uploader("📁 Upload Excel (.xlsx) or CSV", type=["xlsx", "csv"])

if uploaded_file:
    # --- Read file ---
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
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

    # --- Identify key columns ---
    col_machine = find_col(["machine"])
    col_customer = find_col(["customer"])
    col_row = find_col(["row"])
    col_feeds = find_col(["feeds"])
    col_overall_qty = find_col(["overall quantity", "quantity.1"])
    col_wo = find_col(["works order", "w/o"])
    col_finish = find_col(["finish"])
    col_earliest = find_col(["earliest delivery"])
    col_next = find_col(["next uncovered order"])
    col_status = find_col(["status"])
    col_decision = find_col(["run decision"])
    col_value = find_col(["order value", "value"])

    # --- Rename columns consistently ---
    rename_map = {
        col_machine: "Machine",
        col_customer: "Customer",
        col_row: "ROW",
        col_feeds: "Feeds",
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

    # --- Remove filler rows ---
    subset = [c for c in ["Customer", "Feeds"] if c in df.columns]
    if subset:
        df = df.dropna(subset=subset, how="all")

    # --- Force Feeds to numeric ---
    if "Feeds" in df.columns:
        df["Feeds"] = pd.to_numeric(df["Feeds"], errors="coerce").fillna(0).astype(float)
    else:
        st.error("❌ 'Feeds' column not found. Please ensure column D is labelled 'Feeds'.")
        st.stop()

    # --- Convert numeric columns ---
    for col in ["Overall_Quantity", "Order_Value"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # --- Parse date columns ---
    for col in ["Estimated_Finish", "Earliest_Delivery"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # --- Extract sales order dates ---
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

    # --- KPIs + Gauge ---
    st.subheader("📊 Key Metrics")
    total_orders = len(df)
    total_value = df["Order_Value"].sum(skipna=True)
    total_feeds = df["Feeds"].sum(skipna=True)
    total_overall_qty = df["Overall_Quantity"].sum(skipna=True) if "Overall_Quantity" in df.columns else 0

    c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 2])
    c1.metric("Orders", total_orders)
    c2.metric("Total Value", f"£{total_value:,.0f}")
    c3.metric("Total Feeds", f"{int(total_feeds):,}")
    c4.metric("Overall Qty", f"{int(total_overall_qty):,}")

    # --- Utilisation calculation (Feeds only) ---
    util_data = []
    if machine_col:
        df[machine_col] = df[machine_col].astype(str).str.strip().str.upper()

        for m in sorted(df[machine_col].dropna().unique()):
            m_code = str(m).strip().upper()
            m_lookup = MACHINE_ALIASES.get(m_code, m_code)
            cap = MACHINE_CAPACITY.get(m_lookup, None)
            planned_feeds = df.loc[df[machine_col] == m, "Feeds"].sum()

            if cap:
                utilisation = (planned_feeds / cap) * 100
                util_data.append((m_lookup, cap, planned_feeds, round(utilisation, 1)))
            else:
                util_data.append((m_lookup, "N/A", planned_feeds, None))

    util_df = pd.DataFrame(util_data, columns=["Machine", "Feeds/Day", "Planned Feeds", "Utilisation (%)"])

    # --- Average utilisation gauge ---
    if not util_df.empty and "Utilisation (%)" in util_df.columns:
        avg_util = util_df["Utilisation (%)"].mean(skipna=True)
        if PLOTLY_AVAILABLE and not pd.isna(avg_util):
            gauge_fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=avg_util,
                title={'text': "Average Utilisation (%)"},
                gauge={
                    'axis': {'range': [0, 120]},
                    'bar': {'color': "green"},
                    'steps': [
                        {'range': [0, 80], 'color': "red"},
                        {'range': [80, 100], 'color': "yellow"},
                        {'range': [100, 120], 'color': "limegreen"},
                    ],
                }
            ))
            c5.plotly_chart(gauge_fig, use_container_width=True)

    # --- Utilisation table and chart ---
    st.subheader("⚙️ Estimated Machine Utilisation")

    if not util_df.empty:
        def colour_util(val):
            if pd.isna(val):
                return ""
            if val > 100:
                return "background-colour:#66ff66;colour:black"     # Green = over capacity
            elif val >= 80:
                return "background-colour:#ffff66;colour:black"     # Yellow = near capacity
            else:
                return "background-colour:#ff6666;colour:white"     # Red = under capacity

        st.dataframe(util_df.style.applymap(colour_util, subset=["Utilisation (%)"]),
                     use_container_width=True)

        if PLOTLY_AVAILABLE:
            fig = px.bar(
                util_df,
                x="Machine",
                y="Utilisation (%)",
                text_auto=True,
                color="Utilisation (%)",
                color_continuous_scale=["red", "yellow", "green"],
                range_y=[0, 120]
            )
            fig.update_traces(textfont_size=12)
            st.plotly_chart(fig, use_container_width=True)

    # --- Machine-specific orders ---
    st.subheader("🧾 Machine-Specific Planned Orders")
    if machine_col:
        selected_machine = st.selectbox("Select a Machine", ["All Machines"] + sorted(df[machine_col].dropna().unique()))
        subset = df if selected_machine == "All Machines" else df[df[machine_col] == selected_machine]
        if not subset.empty:
            show_cols = [c for c in ["Machine", "Customer", "ROW", "Works_Order", "Feeds", "Overall_Quantity",
                                     "Status", "Run_Decision", "Order_Value", "Estimated_Finish",
                                     "Earliest_Delivery", "Next_Uncovered_Order"] if c in subset.columns]
            st.dataframe(subset[show_cols], use_container_width=True)
        else:
            st.info(f"No orders found for {selected_machine}.")
    else:
        st.warning("⚠️ No 'Machine' column detected. Ensure Column A contains machine codes (BO1, KO1, JC1, TCY, KO3).")

    # --- Orders over time ---
    if "Sales_Order_Date" in df.columns:
        st.subheader("📅 Orders Over Time")
        df_sorted = df.sort_values("Sales_Order_Date")

        if "Order_Value" in df_sorted.columns:
            if PLOTLY_AVAILABLE:
                fig2 = px.line(
                    df_sorted,
                    x="Sales_Order_Date",
                    y="Order_Value",
                    color=machine_col,
                    markers=True,
                    title="Orders Over Time"
                )
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.line_chart(df_sorted.set_index("Sales_Order_Date")["Order_Value"])

    # --- Download filtered data ---
    st.download_button("⬇️ Download Filtered Data",
                       df.to_csv(index=False).encode("utf-8"),
                       "filtered_orders.csv",
                       "text/csv")

else:
    st.info("👆 Upload your Excel or CSV file to begin.")