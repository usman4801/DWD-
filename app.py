"""
DWD Tool
========
A professional Streamlit app that automates daily attendance ("DWD") report
generation for warehouse operations with a clean background design.
"""

import io
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook

# =========================================================================
# CONFIGURATION
# =========================================================================

TEMPLATE_GITHUB_URL = st.secrets.get(
    "TEMPLATE_GITHUB_URL",
    "https://raw.githubusercontent.com/<org>/<repo>/<branch>/Blank%20file.xlsx",
)
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")

SHEET_ROSTER = "Roster"

# Exact raw CSV column mappings matching your actual export files
COL_EMP_ID = "EmpID"
COL_DATE = "Date"
COL_PRESENT_STATUS = "present_status"
COL_GB_LEAVE = "gb_leave_type"
COL_TIME_OFF = "time_off_type"
COL_IS_OVERTIME = "is_overtime"

# Attendance codes as expected by the Dashboard formulas
CODE_PRESENT = "P"
CODE_OT = "OT"
CODE_PL = "PL"
CODE_ABWI = "ABWI"
CODE_SL = "SL"
CODE_ZERO = "0"

# Roster sheet layout based on Blank file.xlsx
ROSTER_HEADER_DATE_CELL = "B1"     
ROSTER_HEADER_DAY_CELL = "C1"      
ROSTER_ID_COLUMN = "B"             # PsoftNo column in Roster sheet
ROSTER_FIRST_DATA_ROW = 7          # First employee row in Roster sheet
ROSTER_ATTENDANCE_COLUMN = "T"     # Attendance column in Roster sheet


# =========================================================================
# CORE LOGIC
# =========================================================================

@st.cache_data(show_spinner=False)
def fetch_template_bytes(url: str, token: str) -> bytes:
    """Download the blank template from GitHub as raw bytes."""
    headers = {"Accept": "application/vnd.github.v3.raw"}
    if token:
        headers["Authorization"] = f"token {token}"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.content


def compute_attendance_code(row: pd.Series) -> str:
    """
    Business rules for attendance:
    - Handle PL, SL, ABWI
    - If present on scheduled shift -> P
    - If present on week off -> OT
    - Otherwise -> 0
    """
    present_status = str(row.get(COL_PRESENT_STATUS, "")).strip().lower()
    gb_leave = str(row.get(COL_GB_LEAVE, "")).strip().lower()
    time_off = str(row.get(COL_TIME_OFF, "")).strip().lower()
    is_ot = row.get(COL_IS_OVERTIME, 0)

    # Check leaves / exclusions
    if "annual leave" in gb_leave or "pl" in time_off or "planned" in time_off:
        return CODE_PL
    if "sick" in gb_leave:
        return CODE_SL
    if "loss of pay" in gb_leave:
        return CODE_ABWI

    is_present = present_status == "present"
    
    if is_present and (is_ot == 1 or is_ot == "1"):
        return CODE_OT
    elif is_present:
        return CODE_PRESENT
    
    return CODE_ZERO


def parse_report_date(raw_df: pd.DataFrame) -> datetime:
    """Extract the single report date from the raw file."""
    dates = pd.to_datetime(raw_df[COL_DATE], errors="coerce").dropna()
    if dates.empty:
        raise ValueError(f"Could not find any valid dates in the '{COL_DATE}' column.")
    return dates.iloc[0].to_pydatetime()


def build_report(raw_df: pd.DataFrame, template_bytes: bytes) -> tuple:
    """
    Populate the Roster sheet of the template while keeping Dashboard formulas
    and all formatting/colors 100% intact.
    """
    report_date = parse_report_date(raw_df)

    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    if SHEET_ROSTER not in wb.sheetnames:
        raise ValueError(f"Template is missing the '{SHEET_ROSTER}' sheet.")
    roster = wb[SHEET_ROSTER]

    # --- Update Date & Day Headers ---
    roster[ROSTER_HEADER_DATE_CELL] = report_date.strftime("%Y-%m-%d")
    roster[ROSTER_HEADER_DAY_CELL] = report_date.strftime("%A")

    # --- Map employee data from raw file ---
    codes_by_id = {}
    for _, row in raw_df.iterrows():
        emp_id = str(row.get(COL_EMP_ID, "")).strip()
        if not emp_id or emp_id == "nan":
            continue
        code = compute_attendance_code(row)
        codes_by_id[emp_id] = code

    # --- Fill Roster Attendance Column ---
    row_idx = ROSTER_FIRST_DATA_ROW
    while True:
        id_cell = roster[f"{ROSTER_ID_COLUMN}{row_idx}"]
        if id_cell.value in (None, ""):
            break
        emp_id = str(id_cell.value).strip()
        code = codes_by_id.get(emp_id)
        if code is not None:
            roster[f"{ROSTER_ATTENDANCE_COLUMN}{row_idx}"] = code
        row_idx += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), report_date


# =========================================================================
# UI — Professional Design with Styled Background & 2 Buttons
# =========================================================================

st.set_page_config(page_title="DWD Tool", page_icon="📋", layout="centered")

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .main-card {
        background: #ffffff;
        padding: 2.5rem;
        border-radius: 16px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.08);
        max-width: 650px;
        margin: auto;
        margin-top: 3rem;
    }
    h1 {
        text-align: center;
        color: #1f2937;
        font-weight: 700;
        font-size: 2.5rem;
        margin-bottom: 0.2rem;
    }
    p.subtitle {
        text-align: center;
        color: #4b5563;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    div.stButton, div.stDownloadButton {
        display: flex;
        justify-content: center;
        margin-top: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="main-card">
        <h1>DWD Tool</h1>
        <p class="subtitle">Daily Workforce Dashboard Report Generator</p>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader("Upload Raw File (CSV)", type=["csv"])

if uploaded_file is not None:
    try:
        raw_df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Raw CSV file parhne mein masla aaya: {e}")
        st.stop()

    with st.spinner("Template fetch ho raha hai aur DWD report ban rahi hai..."):
        try:
            template_bytes = fetch_template_bytes(TEMPLATE_GITHUB_URL, GITHUB_TOKEN)
            report_bytes, report_date = build_report(raw_df, template_bytes)
        except Exception as e:
            st.error(f"Report banane mein error aya: {e}")
            st.stop()

    st.success("DWD Report kamyabi ke sath tayar ho gayi hai!")
    st.download_button(
        "Download File",
        data=report_bytes,
        file_name=f"DWD-AUH1-{report_date.strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("</div>", unsafe_allow_html=True)
