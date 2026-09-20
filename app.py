"""
DWD Tool
========
A minimal Streamlit app that automates daily attendance ("DWD") report
generation for warehouse operations.

Workflow
--------
1. User uploads the raw daily roster CSV.
2. The app pulls the master "Blank file.xlsx" template (Dashboard + Roster
   sheets, with all formulas/formatting) from GitHub.
3. Attendance is computed per the PL / ABWI / SL / Present / OT rules below
   and written into the Roster sheet, cell-by-cell, so every existing
   formula, style, border and color in the template is left untouched.
4. User downloads the finished report.

CONFIGURE ME
------------
Everything you are likely to need to change for your real template/CSV is
grouped under the "CONFIGURATION" section below. The raw CSV column names
and the exact Roster-sheet cell layout are assumptions based on your spec;
update the constants and the two `# ADAPT:` functions to match your actual
files once you can share a sample of each.
"""

import io
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook

# =========================================================================
# CONFIGURATION — adjust these to match your real repo / template / CSV
# =========================================================================

# Raw GitHub URL to the master template. Use the "raw.githubusercontent.com"
# form, e.g.:
# https://raw.githubusercontent.com/<org>/<repo>/<branch>/Blank%20file.xlsx
TEMPLATE_GITHUB_URL = st.secrets.get(
    "TEMPLATE_GITHUB_URL",
    "https://raw.githubusercontent.com/<org>/<repo>/<branch>/Blank%20file.xlsx",
)

# If the repo is private, add a "GITHUB_TOKEN" entry to .streamlit/secrets.toml
# and it will be sent as a bearer token automatically.
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")

SHEET_ROSTER = "Roster"
SHEET_DASHBOARD = "Dashboard"

# Raw CSV column names (rename these to match your actual export)
COL_EMP_ID = "Employee ID"
COL_EMP_NAME = "Employee Name"
COL_SHIFT_STATUS = "Shift Status"      # e.g. "WO" (Week Off) or "Scheduled"
COL_ATTENDANCE = "Attendance Status"   # e.g. "Present", "Absent", "PL", "SL", "ABWI"
COL_DATE = "Date"                      # e.g. 2026-09-20

# Attendance codes as they should appear in the Roster sheet
CODE_PRESENT = "P"
CODE_OT = "OT"
CODE_PL = "PL"
CODE_ABWI = "ABWI"
CODE_SL = "SL"
CODE_ABSENT = "A"

# Roster sheet layout (1-indexed, openpyxl style).
# ADAPT: point these at the real header/data rows and columns of your template.
ROSTER_HEADER_DATE_CELL = "B1"     # cell that shows the report date
ROSTER_HEADER_DAY_CELL = "B2"      # cell that shows the day name
ROSTER_ID_COLUMN = "A"             # column that holds employee IDs, starting at:
ROSTER_FIRST_DATA_ROW = 4          # first row containing an employee record
ROSTER_ATTENDANCE_COLUMN = "C"     # column to write the day's attendance code into


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


def compute_attendance_code(shift_status: str, attendance_status: str) -> str:
    """
    ADAPT: business rules for turning a raw (shift, attendance) pair into
    the single code that goes on the Roster sheet.
    """
    shift = (shift_status or "").strip().upper()
    status = (attendance_status or "").strip().upper()

    is_present = status in {"PRESENT", "P"}
    is_week_off = shift in {"WO", "WEEK OFF", "WEEKOFF"}

    if status in {"PL", "PLANNED LEAVE"}:
        return CODE_PL
    if status in {"ABWI", "ABSENT WITHOUT INFORMATION"}:
        return CODE_ABWI
    if status in {"SL", "SICK LEAVE"}:
        return CODE_SL

    if is_present and is_week_off:
        # Present on a scheduled week off -> counts as overtime
        return CODE_OT
    if is_present:
        return CODE_PRESENT

    return CODE_ABSENT


def parse_report_date(raw_df: pd.DataFrame) -> datetime:
    """Extract the single report date from the raw file."""
    dates = pd.to_datetime(raw_df[COL_DATE], errors="coerce").dropna()
    if dates.empty:
        raise ValueError(
            f"Could not find any valid dates in the '{COL_DATE}' column of the raw file."
        )
    return dates.iloc[0].to_pydatetime()


def build_report(raw_df: pd.DataFrame, template_bytes: bytes) -> bytes:
    """
    Populate the Roster sheet of the template with the day's attendance and
    return the finished workbook as bytes. Formulas/styles are preserved
    because we only ever set `.value` on individual cells — the workbook
    (including the Dashboard sheet) is otherwise untouched.
    """
    report_date = parse_report_date(raw_df)

    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    if SHEET_ROSTER not in wb.sheetnames:
        raise ValueError(f"Template is missing the '{SHEET_ROSTER}' sheet.")
    roster = wb[SHEET_ROSTER]

    # --- Headers -------------------------------------------------------
    roster[ROSTER_HEADER_DATE_CELL] = report_date.strftime("%d-%b-%Y")
    roster[ROSTER_HEADER_DAY_CELL] = report_date.strftime("%A")

    # --- Attendance codes per employee ---------------------------------
    codes_by_id = {}
    for _, row in raw_df.iterrows():
        emp_id = str(row.get(COL_EMP_ID, "")).strip()
        if not emp_id:
            continue
        code = compute_attendance_code(
            row.get(COL_SHIFT_STATUS, ""), row.get(COL_ATTENDANCE, "")
        )
        codes_by_id[emp_id] = code

    row_idx = ROSTER_FIRST_DATA_ROW
    unmatched = []
    while True:
        id_cell = roster[f"{ROSTER_ID_COLUMN}{row_idx}"]
        if id_cell.value in (None, ""):
            break
        emp_id = str(id_cell.value).strip()
        code = codes_by_id.get(emp_id)
        if code is None:
            unmatched.append(emp_id)
        else:
            roster[f"{ROSTER_ATTENDANCE_COLUMN}{row_idx}"] = code
        row_idx += 1

    if unmatched:
        st.warning(
            f"{len(unmatched)} employee ID(s) on the Roster sheet had no "
            f"matching row in the raw file and were left unchanged: "
            f"{', '.join(unmatched[:10])}{'…' if len(unmatched) > 10 else ''}"
        )

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), report_date


# =========================================================================
# UI — deliberately just two controls
# =========================================================================

st.set_page_config(page_title="DWD Tool", page_icon="📋", layout="centered")

st.markdown(
    """
    <style>
    .block-container {max-width: 640px; padding-top: 4rem;}
    div.stButton, div.stDownloadButton {display: flex; justify-content: center;}
    h1 {text-align: center; font-weight: 600;}
    p.subtitle {text-align: center; color: #6b7280; margin-top: -0.5rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<h1>DWD Tool</h1>", unsafe_allow_html=True)
st.markdown(
    "<p class='subtitle'>Upload the raw roster. Download the finished report.</p>",
    unsafe_allow_html=True,
)
st.write("")

uploaded_file = st.file_uploader("Upload Raw File", type=["csv"], label_visibility="visible")

if uploaded_file is not None:
    try:
        raw_df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Could not read the CSV file: {e}")
        st.stop()

    missing_cols = [
        c
        for c in [COL_EMP_ID, COL_EMP_NAME, COL_SHIFT_STATUS, COL_ATTENDANCE, COL_DATE]
        if c not in raw_df.columns
    ]
    if missing_cols:
        st.error(
            "The uploaded CSV is missing expected column(s): "
            f"{', '.join(missing_cols)}. Update the COL_* constants in app.py "
            "to match your actual export, or check the file."
        )
        st.stop()

    with st.spinner("Fetching template and building report…"):
        try:
            template_bytes = fetch_template_bytes(TEMPLATE_GITHUB_URL, GITHUB_TOKEN)
            report_bytes, report_date = build_report(raw_df, template_bytes)
        except Exception as e:
            st.error(f"Failed to build the report: {e}")
            st.stop()

    st.success("Report generated.")
    st.download_button(
        "Download File",
        data=report_bytes,
        file_name=f"DWD_Report_{report_date.strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
