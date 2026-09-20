"""
DWD Tool
========
A Streamlit app that automates daily attendance ("DWD") report generation
for warehouse operations, styled to match the Amazon Seller Tools design.
Laid out to fit a single viewport with no page scrolling.
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


def build_report(raw_df: pd.DataFrame, template_bytes: bytes, selected_date) -> tuple:
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
            "is_ot": is_ot == 1 or is_ot == "1",
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
# UI — Amazon Seller Tools theme, single-viewport (no scroll)
# =========================================================================

st.set_page_config(page_title="DWD Tool - Amazon Operations", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    #MainMenu, header, footer {visibility: hidden;}
    html, body {height: 100%; overflow: hidden;}
    .stApp {background: #eef1f5; height: 100vh; overflow: hidden;}
    .block-container {
        padding: 0 !important;
        max-width: 100% !important;
        height: 100vh;
        overflow: hidden;
    }

    /* ---------- Top navbar ---------- */
    .amz-navbar {
        background: #131921;
        padding: 8px 40px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .amz-logo {
        color: #ffffff;
        font-size: 1.25rem;
        font-weight: 800;
        font-style: italic;
        letter-spacing: -0.5px;
    }
    .amz-logo::after {
        content: "⌣";
        color: #ff9900;
        font-size: 1rem;
        margin-left: 2px;
    }
    .amz-nav-icons {
        display: flex;
        align-items: center;
        gap: 18px;
        color: #ffffff;
        font-size: 1rem;
    }
    .amz-nav-icons .avatar {
        width: 26px; height: 26px;
        border-radius: 50%;
        background: #37475a;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.85rem;
    }

    /* ---------- Hero ---------- */
    .amz-hero {
        background: linear-gradient(120deg, #0f1a2b 0%, #16324f 45%, #2c4a6e 100%);
        padding: 18px 40px 16px 40px;
    }
    .amz-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255,153,0,0.18);
        border: 1px solid rgba(255,153,0,0.5);
        color: #ff9900;
        font-weight: 600;
        font-size: 0.72rem;
        padding: 3px 12px;
        border-radius: 20px;
        margin-bottom: 8px;
    }
    .amz-hero h1 {
        color: #ffffff;
        font-size: 2rem;
        font-weight: 800;
        margin: 0 0 2px 0;
        line-height: 1.1;
    }
    .amz-hero h1 span {color: #ff9900;}
    .amz-hero h2 {
        color: #ffffff;
        font-size: 1.02rem;
        font-weight: 600;
        margin: 0 0 6px 0;
    }
    .amz-hero p.desc {
        color: #c9d3e0;
        font-size: 0.84rem;
        max-width: 560px;
        line-height: 1.4;
        margin-bottom: 12px;
    }
    .amz-features {display: flex; gap: 26px;}
    .amz-feature {display: flex; align-items: center; gap: 8px;}
    .amz-feature .icon {
        width: 28px; height: 28px;
        border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.85rem;
        flex-shrink: 0;
    }
    .amz-feature .icon.orange {background: #ff9900;}
    .amz-feature .icon.blue {background: #2f80ed;}
    .amz-feature .icon.green {background: #17a672;}
    .amz-feature .label {color: #ffffff; font-weight: 700; font-size: 0.82rem; line-height: 1.1;}
    .amz-feature .sub {color: #a9b7c9; font-size: 0.7rem; line-height: 1.1;}

    /* ---------- Content cards ---------- */
    .amz-content {padding: 14px 40px 8px 40px;}
    .amz-panel {
        background: #ffffff;
        border-radius: 12px;
        box-shadow: 0 8px 20px rgba(20,30,50,0.08);
        padding: 4px;
    }
    .left-card {
        background: #f3f6fb;
        border-radius: 10px;
        padding: 16px 16px;
        height: 100%;
    }
    .left-card .icon-circle {
        width: 34px; height: 34px;
        border-radius: 50%;
        background: #dbe6fb;
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem;
        margin-bottom: 8px;
    }
    .left-card h3 {
        color: #131921;
        font-size: 1.02rem;
        font-weight: 800;
        margin: 0 0 8px 0;
        line-height: 1.2;
    }
    .left-card p {
        color: #5b6673;
        font-size: 0.78rem;
        line-height: 1.4;
        margin-bottom: 10px;
    }
    .left-card .tagline {
        font-style: italic;
        font-weight: 700;
        color: #16324f;
        font-size: 0.86rem;
        border-bottom: 2px solid #ff9900;
        display: inline-block;
        padding-bottom: 1px;
    }

    .step-row {display: flex; gap: 10px; padding: 8px 16px 2px 16px;}
    .step-num {
        width: 24px; height: 24px;
        border-radius: 50%;
        color: #fff; font-weight: 700; font-size: 0.75rem;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0;
    }
    .step-num.blue {background: #2f80ed;}
    .step-num.purple {background: #7c4fe0;}
    .step-title {color: #131921; font-weight: 700; font-size: 0.86rem; margin-bottom: 1px;}
    .step-sub {color: #6b7280; font-size: 0.72rem; line-height: 1.25;}

    div[data-testid="stDateInput"] {padding: 0 16px 2px 50px;}
    div[data-testid="stDateInput"] input {
        border-radius: 7px !important;
        border: 1px solid #d7dde5 !important;
        padding: 6px 10px !important;
        font-size: 0.82rem !important;
    }
    div[data-testid="stDateInput"] label {display: none;}

    div[data-testid="stFileUploaderDropzone"] {
        background: #f5f4ff !important;
        border: 2px dashed #a78bfa !important;
        border-radius: 8px !important;
        padding: 4px !important;
    }
    div[data-testid="stFileUploader"] {padding: 2px 16px 4px 50px;}
    div[data-testid="stFileUploader"] section {padding: 6px !important;}
    div[data-testid="stFileUploaderDropzoneInstructions"] span {font-size: 0.78rem !important;}
    div[data-testid="stFileUploaderDropzoneInstructions"] small {font-size: 0.68rem !important;}

    div.stButton {display: flex; justify-content: flex-end; padding: 2px 16px 8px 0;}
    div.stButton>button {
        background: #ff9900;
        color: #131921;
        font-weight: 700;
        border-radius: 7px;
        padding: 0.35rem 1.3rem;
        font-size: 0.85rem;
        border: none;
        box-shadow: 0 4px 10px rgba(255,153,0,0.35);
    }
    div.stButton>button:hover {background: #e88b00; color: #ffffff;}

    div.stDownloadButton {display: flex; justify-content: flex-end; padding: 0 16px 8px 0;}
    div.stDownloadButton>button {
        background: #17a672;
        color: #ffffff;
        font-weight: 700;
        border-radius: 7px;
        padding: 0.35rem 1.3rem;
        font-size: 0.85rem;
        border: none;
        box-shadow: 0 4px 10px rgba(23,166,114,0.35);
    }

    div[data-testid="stAlert"] {padding: 6px 10px; margin: 0 16px 6px 16px; font-size: 0.8rem;}

    /* ---------- Footer ---------- */
    .amz-footer {
        display: flex; justify-content: center; gap: 40px;
        padding: 6px 20px 10px 20px;
        color: #6b7280; font-size: 0.72rem;
    }
    .amz-footer span {display: flex; align-items: center; gap: 5px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Navbar ---------------------------------------------------------------
st.markdown(
    """
    <div class="amz-navbar">
        <div class="amz-logo">amazon</div>
        <div class="amz-nav-icons">
            <span>⠿</span>
            <span>🔔</span>
            <div class="avatar">👤</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Hero -------------------------------------------------------------------
st.markdown(
    """
    <div class="amz-hero">
        <div class="amz-badge">⚡ Amazon Seller Tools</div>
        <h1>DWD <span>Tool</span></h1>
        <h2>Amazon Daily Workforce Dashboard Generator</h2>
        <p class="desc">
            Generate your daily workforce dashboard report quickly and easily.
            Upload your data and get actionable insights in just a few clicks.
        </p>
        <div class="amz-features">
            <div class="amz-feature">
                <div class="icon orange">⚡</div>
                <div><div class="label">Fast</div><div class="sub">Get results in seconds</div></div>
            </div>
            <div class="amz-feature">
                <div class="icon blue">🛡️</div>
                <div><div class="label">Secure</div><div class="sub">Your data stays safe</div></div>
            </div>
            <div class="amz-feature">
                <div class="icon green">📈</div>
                <div><div class="label">Accurate</div><div class="sub">Reliable insights</div></div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Main content -------------------------------------------------------
st.markdown('<div class="amz-content">', unsafe_allow_html=True)

left_col, right_col = st.columns([1, 2.1], gap="medium")

with left_col:
    st.markdown(
        """
        <div class="left-card">
            <div class="icon-circle">📄</div>
            <h3>Create Your Dashboard</h3>
            <p>
                Select the report date and upload your raw data file in CSV
                format to generate your Amazon workforce dashboard.
            </p>
            <div class="tagline">Simple. Fast. Powerful.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with right_col:
    st.markdown('<div class="amz-panel">', unsafe_allow_html=True)

    st.markdown(
        """
        <div class="step-row">
            <div class="step-num blue">1</div>
            <div>
                <div class="step-title">Select Report Date</div>
                <div class="step-sub">Choose the date for which you want to generate the workforce dashboard.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    selected_date = st.date_input(
        "Select Report Date", value=datetime.today(), label_visibility="collapsed"
    )

    st.markdown(
        """
        <div class="step-row">
            <div class="step-num purple">2</div>
            <div>
                <div class="step-title">Upload Raw File (CSV)</div>
                <div class="step-sub">Upload your workforce data file in CSV format. Maximum file size: 200MB per file.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "Upload Raw File (CSV)", type=["csv"], label_visibility="collapsed"
    )

    generate_clicked = st.button("📊  Generate Dashboard  →", use_container_width=False)

    if generate_clicked:
        if uploaded_file is None:
            st.error("Please upload a raw CSV file first.")
        else:
            try:
                raw_df = pd.read_csv(uploaded_file)
            except Exception as e:
                st.error(f"Error reading raw CSV file: {e}")
                raw_df = None

            if raw_df is not None:
                with st.spinner("Generating DWD report..."):
                    try:
                        template_bytes = fetch_template_bytes(TEMPLATE_GITHUB_URL, GITHUB_TOKEN)
                        report_bytes, report_date = build_report(raw_df, template_bytes, selected_date)
                        st.session_state["report_bytes"] = report_bytes
                        st.session_state["report_date"] = report_date
                    except Exception as e:
                        st.error(f"Failed to generate report: {e}")

    if "report_bytes" in st.session_state:
        st.download_button(
            "⬇  Download File",
            data=st.session_state["report_bytes"],
            file_name=f"DWD-AUH1-{st.session_state['report_date'].strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("</div>", unsafe_allow_html=True)  # close amz-panel

st.markdown("</div>", unsafe_allow_html=True)  # close amz-content

# --- Footer -----------------------------------------------------------------
st.markdown(
    """
    <div class="amz-footer">
        <span>🛡️ Your data is always protected</span>
        <span>〰️ Built for Amazon Sellers</span>
        <span>🕒 Save time, make better decisions</span>
    </div>
    """,
    unsafe_allow_html=True,
)
