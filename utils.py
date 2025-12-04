import streamlit as st
import pandas as pd
import gspread
from google.oauth2 import service_account

# =========================================================
# Google Sheet URLs / Tab Names
# =========================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dLHPyIkZfJby-N-EMXTraJv25BxkauBHS2ycdndC1PY"
TAB_NAME = "CTI"

PO_SHEET_URL = "https://docs.google.com/spreadsheets/d/1Ct5jXkpG6x13AnQeJ12589jj2ZCe7f3EERTU77KEYhQ"
PO_TAB_NAME = "Progroup_POs"

LEAD_SHEET_URL = "https://docs.google.com/spreadsheets/d/1CwC7rpoMU9wCwVh7qYHhn93Ld5C0jBXLX2SYVQ1c1o8"
LEAD_TAB_NAME = "Progroup_Lead_Times"

# =========================================================
# Credentials from Streamlit secrets
# =========================================================
def _get_creds():
    info = dict(st.secrets["google"])
    return service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )

# =========================================================
# Load CTI sheet
# =========================================================
@st.cache_data(show_spinner=False, ttl=120)
def load_cti_sheet() -> pd.DataFrame:
    creds = _get_creds()
    client = gspread.authorize(creds)
    ws = client.open_by_url(SHEET_URL).worksheet(TAB_NAME)
    return pd.DataFrame(ws.get_all_records())

# =========================================================
# Load Progroup POs sheet (auto-detect header)
# =========================================================
@st.cache_data(show_spinner=False, ttl=120)
def load_po_sheet():
    creds = _get_creds()
    client = gspread.authorize(creds)

    ss = client.open_by_url(PO_SHEET_URL)
    ws = ss.worksheet(PO_TAB_NAME)

    rows = ws.get_all_records()
    if not rows:
        return pd.DataFrame()

    return pd.DataFrame(rows)

# =========================================================
# Load Board Grade Lead Times sheet
# =========================================================
@st.cache_data(show_spinner=False, ttl=120)
def load_lead_sheet() -> pd.DataFrame:
    creds = _get_creds()
    client = gspread.authorize(creds)

    ss = client.open_by_url(LEAD_SHEET_URL)
    ws = ss.worksheet(LEAD_TAB_NAME)

    rows = ws.get_all_values()
    if not rows:
        return pd.DataFrame()

    # assume first row is header
    header = rows[0]
    data = rows[1:]

    return pd.DataFrame(data, columns=header)