"""
DWD Tool
========
A professional Streamlit app that automates daily attendance ("DWD") report
generation for warehouse operations with Amazon theme background and date selection.
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

TEMPLATE_GITHUB_URL = "https://raw.githubusercontent.com/usman4801/DWD-/main/Blank%20file.xlsx"
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")

SHEET_ROSTER = "Roster"

# Exact raw CSV column mappings
COL_EMP_ID = "EmpID"
COL_DATE = "Date"
COL_PRESENT_STATUS = "present_status"
COL_IS_OVERTIME = "is_overtime"
COL_GB_LEAVE = "gb_leave_type"
COL_TIME_OFF = "time_off_type"

# Attendance codes based on template legend
CODE_PRESENT = "P"
CODE_PL = "PL"
CODE_OFF = "OFF"
CODE_ZERO = "0"

# Roster sheet layout based on Blank file.xlsx
ROSTER_HEADER_DATE_CELL = "B1"     
ROSTER_HEADER_DAY_CELL = "C1"      
ROSTER_ID_COLUMN = "B"             # PsoftNo column in Roster sheet
ROSTER_FIRST_DATA_ROW = 7          # First employee row in Roster sheet
ROSTER_ATTENDANCE_COLUMN = "T"     # Attendance column in Roster sheet
ROSTER_REMARKS_COLUMN = "U"        # Remarks column in Roster sheet
ROSTER_OFF1_COLUMN = "L"           # OFF1 column in Roster sheet
ROSTER_OFF2_COLUMN = "M"           # OFF2 column in Roster sheet


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


def build_report(raw_df: pd.DataFrame, template_bytes: bytes, selected_date: datetime.date) -> tuple:
    """
    Populate the Roster sheet of the template while keeping Dashboard formulas
    and all formatting/colors 100% intact.
    """
    day_name = selected_date.strftime("%A").strip().lower()

    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    if SHEET_ROSTER not in wb.sheetnames:
        raise ValueError(f"Template is missing the '{SHEET_ROSTER}' sheet.")
    roster = wb[SHEET_ROSTER]

    # --- Update Date & Day Headers ---
    roster[ROSTER_HEADER_DATE_CELL] = selected_date.strftime("%Y-%m-%d")
    roster[ROSTER_HEADER_DAY_CELL] = selected_date.strftime("%A")

    # --- Map employee data from raw file ---
    emp_data = {}
    for _, row in raw_df.iterrows():
        emp_id = str(row.get(COL_EMP_ID, "")).strip()
        if not emp_id or emp_id == "nan":
            continue
        present_status = str(row.get(COL_PRESENT_STATUS, "")).strip().lower()
        gb_leave = str(row.get(COL_GB_LEAVE, "")).strip().lower()
        time_off = str(row.get(COL_TIME_OFF, "")).strip().lower()
        is_ot = row.get(COL_IS_OVERTIME, 0)
        
        is_present = present_status == "present"
        is_pl = "annual leave" in gb_leave or "annualleave" in time_off
        
        emp_data[emp_id] = {
            "is_present": is_present,
            "is_pl": is_pl,
            "is_ot": is_ot == 1 or is_ot == "1"
        }

    # --- Fill Roster Attendance, OFF replacement, and Remarks Columns ---
    row_idx = ROSTER_FIRST_DATA_ROW
    while True:
        id_cell = roster[f"{ROSTER_ID_COLUMN}{row_idx}"]
        if id_cell.value in (None, ""):
            break
        emp_id = str(id_cell.value).strip()
        
        data = emp_data.get(emp_id)
        if data:
            off1_cell = roster[f"{ROSTER_OFF1_COLUMN}{row_idx}"]
            off2_cell = roster[f"{ROSTER_OFF2_COLUMN}{row_idx}"]
            
            off1_val = str(off1_cell.value or "").strip().lower()
            off2_val = str(off2_cell.value or "").strip().lower()
            
            is_scheduled_off = (day_name == off1_val or day_name == off2_val)

            if data["is_pl"]:
                roster[f"{ROSTER_ATTENDANCE_COLUMN}{row_idx}"] = CODE_PL
                roster[f"{ROSTER_REMARKS_COLUMN}{row_idx}"] = "Annual Leave"
            elif data["is_present"]:
                roster[f"{ROSTER_ATTENDANCE_COLUMN}{row_idx}"] = CODE_PRESENT
                
                # Check if present on week off -> OT conversion
                if day_name == off1_val:
                    off1_cell.value = "OT"
                    roster[f"{ROSTER_REMARKS_COLUMN}{row_idx}"] = "6th day OT"
                elif day_name == off2_val:
                    off2_cell.value = "OT"
                    roster[f"{ROSTER_REMARKS_COLUMN}{row_idx}"] = "7th day OT"
            elif is_scheduled_off:
                roster[f"{ROSTER_ATTENDANCE_COLUMN}{row_idx}"] = CODE_OFF
            else:
                roster[f"{ROSTER_ATTENDANCE_COLUMN}{row_idx}"] = CODE_ZERO
            
        row_idx += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), selected_date


# =========================================================================
# UI — Amazon Professional Theme & 2 Buttons
# =========================================================================

st.set_page_config(page_title="DWD Tool - Amazon Operations", page_icon="📦", layout="centered")

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #131921 0%, #232f3e 50%, #37475a 100%);
    }
    .main-card {
        background: #ffffff;
        padding: 2.5rem;
        border-radius: 12px;
        box-shadow: 0 12px 30px rgba(0,0,0,0.3);
        max-width: 650px;
        margin: auto;
        margin-top: 3rem;
        border-top: 5px solid #ff9900;
    }
    h1 {
        text-align: center;
        color: #232f3e;
        font-weight: 700;
        font-size: 2.5rem;
        margin-bottom: 0.2rem;
    }
    p.subtitle {
        text-align: center;
        color: #555555;
        font-size: 1.1rem;
        margin-bottom: 2rem;
        font-weight: 500;
    }
    div.stButton, div.stDownloadButton {
        display: flex;
        justify-content: center;
        margin-top: 1.5rem;
    }
    .stButton>button, .stDownloadButton>button {
        background-color: #ff9900;
        color: #131921;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        border: none;
        box-shadow: 0 4px 10px rgba(255, 153, 0, 0.3);
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #e88b00;
        color: #ffffff;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="main-card">
        <h1>DWD Tool</h1>
        <p class="subtitle">Amazon Daily Workforce Dashboard Generator</p>
    """,
    unsafe_allow_html=True,
)

selected_date = st.date_input("Select Report Date", value=datetime.today())

uploaded_file = st.file_uploader("Upload Raw File (CSV)", type=["csv"])

if uploaded_file is not None:
    try:
        raw_df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading raw CSV file: {e}")
        st.stop()

    with st.spinner("Fetching template from GitHub and generating DWD report..."):
        try:
            template_bytes = fetch_template_bytes(TEMPLATE_GITHUB_URL, GITHUB_TOKEN)
            report_bytes, report_date = build_report(raw_df, template_bytes, selected_date)
        except Exception as e:
            st.error(f"Failed to generate report: {e}")
            st.stop()

    st.success("DWD Report generated successfully!")
    st.download_button(
        "Download File",
        data=report_bytes,
        file_name=f"DWD-AUH1-{report_date.strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("</div>", unsafe_allow_html=True)
