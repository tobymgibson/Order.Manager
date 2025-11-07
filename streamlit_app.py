import streamlit as st
import pandas as pd
import re
from datetime import datetime

# --- Machine capacities (feeds per day) ---
MACHINE_CAPACITY = {
    "BO1": 24000,
    "KO1": 50000,
    "JC1": 48000,
    "TCY": 9000,
    "KO3": 15000,
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
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# --- Streamlit setup ---
st.set_page_config(page_title="🏭 Production Dashboard", layout="wide")
st.title("🏭 Production Planning Dashboard")
st.caption(f"🕒 Last refreshed: {datetime.now():%d %b %Y, %H:%M}")
st.caption("Upload your Excel (.xlsx) file to view KPIs, utilisation, and planned orders by machine.")

# --- Load uploaded file ---
def load_data(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        first_sheet = excel_file.sheet_names[0]
        df = pd.read_excel(excel_file, sheet_name=first_sheet)
    return df


uploaded_file = st.file_uploader("📁 Upload Excel (.xlsx) or CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        df = load_data(uploaded_file)
    except Exception as e:
        st.error(f"❌ Could not read file: {e}")
        st.stop()

    # Save original for reset
    if "original_df" not in st.session_state:
        st.session_state.original_df = df.copy()

    # Reset button
    if st.button("🔄 Reset Data to Original Upload"):
        st.session_state.editable_df = st.session_state.original_df.copy()
        st.success("✅ Data reset to original upload.")
        st.rerun()

    # Work with editable dataframe
    if "editable_df" not in st.session_state:
        st.session_state.editable_df = df.copy()

    df = st.session_state.editable_df.copy()

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
    col_value = find_col(["order value", "value", "£"])

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

    # --- Fill machine names down ---
    if "Machine" in df.columns:
        df["Machine"] = df["Machine"].ffill()

    # --- Clean numeric columns ---
    for col in ["Feeds", "Overall_Quantity", "Order_Value"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0  # ensure columns always exist

    # --- Parse dates ---
    for col in ["Estimated_Finish", "Earliest_Delivery"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # --- Extract sales order date ---
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
    total_value = df["Order_Value"].sum(skipna=True)
    total_feeds = df["Feeds"].sum(skipna=True)
    total_overall_qty = df["Overall_Quantity"].sum(skipna=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Orders", total_orders)
    c2.metric("Total Value", f"£{total_value:,.0f}")
    c3.metric("Total Feeds", f"{int(total_feeds):,}")
    c4.metric("Overall Qty", f"{int(total_overall_qty):,}")

    # --- Utilisation ---
    st.subheader("⚙️ Estimated Machine Utilisation")
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

    def colour_util(val):
        if pd.isna(val):
            return ""
        if val > 100:
            return "background-color:#66ff66; color:black"
        elif val >= 80:
            return "background-color:#ffff66; color:black"
        else:
            return "background-color:#ff6666; color:white"

    st.dataframe(util_df.style.applymap(colour_util, subset=["Utilisation (%)"]), use_container_width=True)

    # --- Gauge ---
    if not util_df.empty and "Utilisation (%)" in util_df.columns:
        avg_util = util_df["Utilisation (%)"].mean(skipna=True)
        st.markdown("### ⚖️ Average Utilisation Across All Machines")

        if PLOTLY_AVAILABLE and not pd.isna(avg_util):
            gauge_colour = (
                "red" if avg_util < 80 else
                "yellow" if avg_util <= 100 else
                "limegreen"
            )

            gauge_fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=avg_util,
                title={'text': "Average Utilisation (%)"},
                gauge={
                    'axis': {'range': [0, 120]},
                    'bar': {'color': gauge_colour},
                    'steps': [
                        {'range': [0, 80], 'color': "red"},
                        {'range': [80, 100], 'color': "yellow"},
                        {'range': [100, 120], 'color': "limegreen"},
                    ],
                }
            ))

            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.plotly_chart(gauge_fig, use_container_width=True)

    # --- Editable Machine Orders ---
    st.subheader("🧾 Machine-Specific Planned Orders")

    if machine_col:
        selected_machine = st.selectbox(
            "Select a Machine",
            ["All Machines"] + sorted(df[machine_col].dropna().unique())
        )

        subset = df if selected_machine == "All Machines" else df[df[machine_col] == selected_machine]

        if not subset.empty:
            show_cols = [c for c in ["Machine", "Customer", "ROW", "Works_Order", "Feeds", "Overall_Quantity",
                                     "Status", "Run_Decision", "Order_Value", "Estimated_Finish",
                                     "Earliest_Delivery", "Next_Uncovered_Order"] if c in subset.columns]

            st.caption("✏️ Edit cells directly or tick boxes to select rows for deletion.")

            # Add checkbox column
            subset = subset.reset_index(drop=True)
            subset["Delete?"] = False

            edited_df = st.data_editor(subset[["Delete?"] + show_cols], num_rows="dynamic", use_container_width=True,
                                       key=f"editor_{selected_machine}")

            delete_rows = edited_df.index[edited_df["Delete?"] == True].tolist()

            if st.button("🗑️ Delete Selected Rows"):
                if delete_rows:
                    st.session_state.editable_df = st.session_state.editable_df.drop(subset.index[delete_rows])
                    st.success(f"Deleted {len(delete_rows)} row(s) from {selected_machine}.")
                    st.rerun()
                else:
                    st.warning("Please tick at least one row to delete.")

            # Save edits (ignoring Delete? column)
            edited_df = edited_df.drop(columns=["Delete?"], errors="ignore")
            if selected_machine == "All Machines":
                st.session_state.editable_df = edited_df
            else:
                st.session_state.editable_df.loc[
                    st.session_state.editable_df[machine_col] == selected_machine, show_cols
                ] = edited_df.values

            st.download_button(
                "⬇️ Download Updated Data",
                st.session_state.editable_df.to_csv(index=False).encode("utf-8"),
                f"{selected_machine}_updated.csv",
                "text/csv"
            )
else:
    st.info("👆 Upload your Excel or CSV file to begin.")