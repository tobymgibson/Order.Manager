import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import re

# =====================================
# GOOGLE SHEETS CONNECTION
# =====================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dLHPyIkZfJby-N-EMXTraJv25BxkauBHS2ycdndC1PY/edit?gid=542956255#gid=542956255"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
client = gspread.authorize(creds)

@st.cache_data(ttl=120)
def load_gsheet_data():
    """Load and clean Google Sheet data"""
    try:
        sh = client.open_by_url(SHEET_URL)
        worksheet = sh.sheet1
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        df = df[df.astype(str).apply(lambda x: "".join(x).strip() != "", axis=1)]
        return df
    except Exception as e:
        st.error(f"⚠️ Could not load data from Google Sheets: {e}")
        return pd.DataFrame()

df = load_gsheet_data()
if df.empty:
    st.stop()

# =====================================
# PAGE SETUP
# =====================================
st.set_page_config(page_title="Production Dashboard", layout="wide")
st.title("🏭 Live Production Planning Dashboard")
st.caption(f"Connected to Google Sheets • Updated {datetime.now():%d %b %Y, %H:%M}")

# =====================================
# MACHINE SETTINGS
# =====================================
MACHINE_CAPACITY = {"BO1": 24000, "KO1": 50000, "JC1": 48000, "TCY": 9000, "KO3": 15000}

# =====================================
# HELPER FUNCTIONS
# =====================================
def clean_numeric(series):
    s = series.astype(str).fillna("")
    s = s.str.replace(r"[£$,]", "", regex=True).str.replace(r"[^\d.\-]", "", regex=True)
    s = s.replace("", "0")
    return pd.to_numeric(s, errors="coerce").fillna(0)

def extract_date_from_text(text):
    if not isinstance(text, str):
        return None
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", text)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d/%m/%Y")
        except:
            return None
    return None

def risk_level(val):
    if pd.isna(val) or not str(val).strip():
        return "Unknown"
    text = str(val).lower()
    if "all covered" in text:
        return "All Covered"
    date = extract_date_from_text(val)
    if date:
        if date.date() <= (datetime.now().date() + timedelta(days=3)):
            return "Urgent"
        else:
            return "Future"
    return "Future"

def risk_flag(val):
    r = risk_level(val)
    return {
        "Urgent": "🔴 Urgent",
        "Future": "🟡 Future",
        "All Covered": "✅ Covered",
        "Unknown": "❔",
    }[r]

def apply_util_colour(val):
    if pd.isna(val): return ""
    if val > 100: return "background-color: #00cc00; color: black;"
    if val >= 80: return "background-color: #ffff66; color: black;"
    return "background-color: #ff6666; color: white;"

# =====================================
# CLEANUP AND NORMALISE DATA
# =====================================
df.columns = df.columns.astype(str).str.strip()

rename_map = {
    "Machine": "Machine",
    "Customer": "Customer",
    "Feeds": "Feeds",
    "Finish": "Finish",
    "Works Order": "Works_Order",
    "Value": "Order_Value",
    "Next Uncovered Order": "Next_Uncovered_Order",
}
df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

# Fill down machine names (fixing the blank rows)
if "Machine" in df.columns:
    df["Machine"] = df["Machine"].ffill().astype(str).str.upper()

if "Customer" in df.columns:
    df["Customer"] = df["Customer"].astype(str).str.upper()

for col in ["Feeds", "Order_Value"]:
    if col in df.columns:
        df[col] = clean_numeric(df[col])

if "Finish" in df.columns:
    df["Finish"] = pd.to_datetime(df["Finish"], errors="coerce")

if "Next_Uncovered_Order" in df.columns:
    df["Next_Shortage_Date"] = df["Next_Uncovered_Order"].apply(extract_date_from_text)

# Filter out blanks
df = df[df["Machine"].notna() & (df["Feeds"] > 0)]

# =====================================
# OVERALL MACHINE UTILISATION
# =====================================
st.subheader("Overall Machine Utilisation Across All Planned Orders")

utilisation_data = (
    df.groupby(["Machine", df["Finish"].dt.date])["Feeds"]
    .sum()
    .reset_index()
    .groupby("Machine")["Feeds"]
    .mean()
    .reset_index(name="Avg_Feeds_Per_Day")
)

util_rows = []
for _, row in utilisation_data.iterrows():
    machine = row["Machine"]
    avg = row["Avg_Feeds_Per_Day"]
    capacity = MACHINE_CAPACITY.get(machine)
    if capacity:
        util = (avg / capacity) * 100
        util_rows.append((machine, capacity, round(avg, 1), round(util, 1)))

util_df = pd.DataFrame(util_rows, columns=["Machine", "Feeds/Day Capacity", "Average Planned Feeds", "Utilisation (%)"])
util_df = util_df.sort_values("Machine").reset_index(drop=True)

st.dataframe(util_df.style.applymap(apply_util_colour, subset=["Utilisation (%)"]), use_container_width=True)

# =====================================
# ORDERS TABLE SECTION
# =====================================
st.header("Orders Scheduled to Finish in Selected Range")

min_date, max_date = df["Finish"].min().date(), df["Finish"].max().date()
date_selection = st.date_input(
    "Select date range:",
    value=(min_date, min_date),
    min_value=min_date,
    max_value=max_date,
)

# Parse selected range
if isinstance(date_selection, tuple):
    if len(date_selection) == 2:
        start_date, end_date = date_selection
    else:
        start_date = end_date = date_selection[0]
else:
    start_date = end_date = date_selection

filtered_df = df[(df["Finish"].dt.date >= start_date) & (df["Finish"].dt.date <= end_date)].copy()

if filtered_df.empty:
    st.info(f"No orders found between {start_date:%d %b %Y} and {end_date:%d %b %Y}.")
    st.stop()

# Dropdown machine selection
available_machines = sorted(filtered_df["Machine"].dropna().unique())
selected_machine = st.selectbox("Select machine to view:", options=["All Machines"] + available_machines, index=0)

if selected_machine != "All Machines":
    filtered_day_df = filtered_df[filtered_df["Machine"] == selected_machine].copy()
else:
    filtered_day_df = filtered_df.copy()

# Calculate metrics
total_orders = len(filtered_day_df)
total_feeds = filtered_day_df["Feeds"].sum(skipna=True)
total_value = filtered_day_df["Order_Value"].sum(skipna=True)

c1, c2, c3 = st.columns(3)
c1.metric("Orders Planned", total_orders)
c2.metric("Total Feeds", f"{int(total_feeds):,}")
c3.metric("Total Value", f"£{total_value:,.0f}")

# Add colour-coded risk
filtered_day_df["Risk_Flag"] = filtered_day_df["Next_Uncovered_Order"].apply(risk_flag)

# Editable table
st.markdown("### Orders List (Editable)")
editable_df = filtered_day_df.copy()
editable_df["Delete?"] = False

edited = st.data_editor(
    editable_df,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
)

if st.button("🗑️ Delete Selected Rows"):
    delete_idx = edited[edited["Delete?"] == True].index
    if len(delete_idx) > 0:
        df = df.drop(delete_idx)
        st.success(f"Deleted {len(delete_idx)} rows.")
    else:
        st.info("No rows selected for deletion.")