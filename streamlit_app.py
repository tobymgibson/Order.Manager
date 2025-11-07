import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# =========================
# MACHINE SETTINGS
# =========================
MACHINE_CAPACITY = {
    "BO1": 24000,
    "KO1": 50000,
    "JC1": 48000,
    "TCY": 9000,
    "KO3": 15000,
}
MACHINE_ALIASES = {"JC": "JC1", "BO": "BO1", "KLETT": "KO1", "KLETT3": "KO3"}

# =========================
# STREAMLIT SETUP
# =========================
st.set_page_config(page_title="Production Dashboard", layout="wide")
st.title("🏭 Production Planning Dashboard")
st.caption(f"Last refreshed: {datetime.now():%d %b %Y, %H:%M}")

# =========================
# HELPERS
# =========================
def load_data(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    excel = pd.ExcelFile(uploaded_file, engine="openpyxl")
    first_sheet = excel.sheet_names[0]
    return pd.read_excel(excel, sheet_name=first_sheet)

def clean_numeric(series):
    series = series.astype(str).fillna("")
    series = (
        series.str.replace(r"[£$,]", "", regex=True)
              .str.replace(r"[^\d.\-]", "", regex=True)
              .str.replace(" ", "")
              .replace("", "0")
    )
    return pd.to_numeric(series, errors="coerce").fillna(0)

def extract_date_from_text(text):
    if not isinstance(text, str):
        return None
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d/%m/%Y")
        except Exception:
            return None
    return None

def risk_level(val):
    if pd.isna(val) or not str(val).strip():
        return "Unknown"
    t = str(val).strip().lower()
    if "all covered" in t:
        return "All Covered"
    d = extract_date_from_text(t)
    if d:
        return "Urgent" if d.date() <= (datetime.now().date() + timedelta(days=3)) else "Future"
    return "Future"

def risk_flag(val):
    r = risk_level(val)
    return {"Urgent": "🔴 Urgent", "Future": "🟡 Future", "All Covered": "✅ Covered", "Unknown": "❔"}[r]

def apply_util_colour(val):
    if pd.isna(val): return ""
    if val > 100:   return "background-color: #00cc00; colour: black;"
    if val >= 80:   return "background-color: #ffff66; colour: black;"
    return "background-color: #ff6666; colour: white;"

def find_col(df, possible_names):
    for col in df.columns:
        for name in possible_names:
            if name.lower() in col.lower():
                return col
    return None

def ensure_rowid(df):
    if "RowID" not in df.columns:
        df.insert(0, "RowID", range(len(df)))
    return df

# =========================
# FILE UPLOAD
# =========================
uploaded_file = st.file_uploader("Upload Excel (.xlsx) or CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        raw_df = load_data(uploaded_file)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()

    raw_df.columns = raw_df.columns.astype(str).str.strip()

    # Detect key columns
    next_col = find_col(raw_df, ["next uncovered", "uncovered order"])
    value_col = find_col(raw_df, ["value", "order value"]) or "Order_Value"
    quantity_col = find_col(raw_df, ["quantity", "overall quantity"]) or "Quantity"
    customer_col = find_col(raw_df, ["customer"]) or "Customer"
    machine_col = find_col(raw_df, ["machine"]) or "Machine"
    feeds_col = find_col(raw_df, ["feeds"]) or "Feeds"
    finish_col = find_col(raw_df, ["finish"]) or "Finish"
    works_order_col = find_col(raw_df, ["works order", "works_order", "wo"]) or "Works_Order"
    row_spec_col = find_col(raw_df, ["row", "spec"]) or "ROW"

    df = raw_df.copy()
    if next_col and next_col != "Next_Uncovered_Order":
        df.rename(columns={next_col: "Next_Uncovered_Order"}, inplace=True)
    for old, new in [
        (customer_col, "Customer"),
        (machine_col, "Machine"),
        (feeds_col, "Feeds"),
        (finish_col, "Finish"),
        (works_order_col, "Works_Order"),
        (row_spec_col, "ROW"),
    ]:
        if old != new and old in df.columns:
            df.rename(columns={old: new}, inplace=True)

    for col in ["Machine", "Customer"]:
        if col in df.columns:
            df[col] = df[col].ffill().astype(str).str.upper()

    for col in ["Feeds", quantity_col, value_col]:
        if col in df.columns:
            df[col] = clean_numeric(df[col])

    if "Finish" in df.columns:
        df["Finish"] = pd.to_datetime(df["Finish"], errors="coerce")

    if "Next_Uncovered_Order" in df.columns:
        df["Next_Shortage_Date"] = df["Next_Uncovered_Order"].apply(extract_date_from_text)

    # Session state
    if "original_df" not in st.session_state:
        st.session_state.original_df = df.copy()
    if "editable_df" not in st.session_state:
        st.session_state.editable_df = ensure_rowid(df.copy())

    if st.button("Reset Data to Original Upload"):
        st.session_state.editable_df = ensure_rowid(st.session_state.original_df.copy())
        st.success("Data reset to original upload.")
        st.rerun()

    df = st.session_state.editable_df.copy()

    # =========================
    # DATE RANGE SELECTION
    # =========================
    st.header("Production Overview — Select Multiple Finish Dates")
    if "Finish" not in df.columns or df["Finish"].isna().all():
        st.warning("No valid Finish dates found.")
        st.stop()

    min_date, max_date = df["Finish"].min().date(), df["Finish"].max().date()

    date_selection = st.date_input(
        "Select date range:",
        value=(min_date, min_date),
        min_value=min_date,
        max_value=max_date,
    )

    # Handle all return types
    if isinstance(date_selection, tuple):
        if len(date_selection) == 1 and isinstance(date_selection[0], tuple):
            start_date, end_date = date_selection[0]
        elif len(date_selection) == 2:
            start_date, end_date = date_selection
        else:
            start_date = end_date = date_selection[0]
    else:
        start_date = end_date = date_selection

    date_filtered_df = df[
        (df["Finish"].dt.date >= start_date) &
        (df["Finish"].dt.date <= end_date)
    ].copy()

    if date_filtered_df.empty:
        st.info(f"No orders between {start_date:%d %b %Y} and {end_date:%d %b %Y}.")
        st.stop()

    # =========================
    # SUMMARY
    # =========================
    total_orders = len(date_filtered_df)
    total_feeds = date_filtered_df["Feeds"].sum(skipna=True)
    total_value = date_filtered_df[value_col].sum(skipna=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Orders Planned", total_orders)
    c2.metric("Total Feeds", f"{int(total_feeds):,}")
    c3.metric("Total Value", f"£{total_value:,.0f}")

    # =========================
    # OVERALL MACHINE UTILISATION (All Planned Orders)
    # =========================
    st.subheader("Overall Machine Utilisation Across All Planned Orders")

    overall_grouped = (
        df.groupby(["Machine", df["Finish"].dt.date])["Feeds"]
          .sum()
          .reset_index()
          .groupby("Machine")["Feeds"]
          .mean()
          .reset_index(name="Avg_Feeds_Per_Day")
    )

    overall_util_rows = []
    for _, row in overall_grouped.iterrows():
        m = row["Machine"]
        avg_feeds = row["Avg_Feeds_Per_Day"]
        cap = MACHINE_CAPACITY.get(MACHINE_ALIASES.get(m, m), None)
        util = (avg_feeds / cap) * 100 if cap else None
        overall_util_rows.append((m, cap or "N/A", round(avg_feeds, 1), round(util, 1) if util else None))

    overall_util_df = pd.DataFrame(
        overall_util_rows,
        columns=["Machine", "Feeds/Day Capacity", "Average Planned Feeds", "Utilisation (%)"]
    )

    st.dataframe(overall_util_df.style.applymap(apply_util_colour, subset=["Utilisation (%)"]), use_container_width=True)

    # =========================
    # MACHINE UTILISATION (Selected Range)
    # =========================
    st.subheader("Estimated Machine Utilisation (Average Feeds/Day Capacity vs Planned — Selected Range)")

    util_source = date_filtered_df.copy()
    grouped = (
        util_source.groupby(["Machine", util_source["Finish"].dt.date])["Feeds"]
          .sum()
          .reset_index()
          .groupby("Machine")["Feeds"]
          .mean()
          .reset_index(name="Avg_Feeds_Per_Day")
    )

    util_rows = []
    for _, row in grouped.iterrows():
        m = row["Machine"]
        avg_feeds = row["Avg_Feeds_Per_Day"]
        cap = MACHINE_CAPACITY.get(MACHINE_ALIASES.get(m, m), None)
        util = (avg_feeds / cap) * 100 if cap else None
        util_rows.append((m, cap or "N/A", round(avg_feeds, 1), round(util, 1) if util else None))
    util_df = pd.DataFrame(util_rows, columns=["Machine", "Feeds/Day Capacity", "Average Planned Feeds", "Utilisation (%)"])
    st.dataframe(util_df.style.applymap(apply_util_colour, subset=["Utilisation (%)"]), use_container_width=True)

    # =========================
    # ORDERS TABLE
    # =========================
    st.subheader("Orders Scheduled to Finish in Selected Range")

    available_machines = sorted(date_filtered_df["Machine"].dropna().unique())
    machine_choice = st.selectbox("Select Machine to View", options=["All Machines"] + available_machines, index=0)

    if machine_choice == "All Machines":
        filtered_day_df = date_filtered_df.copy()
    else:
        filtered_day_df = date_filtered_df[date_filtered_df["Machine"] == machine_choice].copy()

    sort_order = st.radio("Sort Finish Time", ["Earliest First", "Latest First"], horizontal=True)
    ascending = True if sort_order == "Earliest First" else False
    filtered_day_df = filtered_day_df.sort_values(by="Finish", ascending=ascending)

    filtered_day_df["Risk"] = filtered_day_df["Next_Uncovered_Order"].apply(risk_level)
    filtered_day_df["Risk_Flag"] = filtered_day_df["Next_Uncovered_Order"].apply(risk_flag)
    filtered_day_df["Delete?"] = False

    show_cols = [c for c in [
        "Delete?",
        "Machine",
        "Customer",
        "ROW",
        "Works_Order",
        "Feeds",
        quantity_col,
        "Next_Uncovered_Order",
        "Next_Shortage_Date",
        "Risk_Flag",
        "Run_Decision",
        value_col,
        "Finish",
    ] if c in filtered_day_df.columns or c in ["Delete?", "Risk_Flag"]]

    editor_df = filtered_day_df[["RowID"] + show_cols].reset_index(drop=True)
    edited = st.data_editor(editor_df.drop(columns=["RowID"]),
                            num_rows="dynamic",
                            use_container_width=True,
                            key=f"editor_{start_date}_{end_date}")

    edited_with_id = editor_df.copy()
    edited_with_id["Delete?"] = edited["Delete?"].values
    delete_rids = edited_with_id.loc[edited_with_id["Delete?"] == True, "RowID"].tolist()

    if st.button("Delete Selected Rows"):
        if delete_rids:
            st.session_state.editable_df = st.session_state.editable_df[
                ~st.session_state.editable_df["RowID"].isin(delete_rids)
            ]
            st.success(f"Deleted {len(delete_rids)} row(s) between {start_date:%d %b %Y} and {end_date:%d %b %Y}.")
            st.rerun()
        else:
            st.warning("Tick at least one row using the ‘Delete?’ checkbox.")
else:
    st.info("Upload your Excel or CSV file to begin.")