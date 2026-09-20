"""
DWD Tool
========
A Streamlit app that automates daily attendance ("DWD") report generation
for warehouse operations, styled as a polished Amazon Seller Tools page.
Fits a single viewport — no page scrolling.
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
# UI — Amazon Seller Tools theme, polished, single viewport
# =========================================================================

st.set_page_config(page_title="DWD Tool - Amazon Operations", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    #MainMenu, header, footer {visibility: hidden;}
    html, body {height: 100%; margin: 0; padding: 0;}
    .stApp {background: #eef1f5; min-height: 100vh; overflow-y: auto; overflow-x: hidden;}
    .block-container {
        padding: 0 !important;
        max-width: 100% !important;
    }
    div[data-testid="stVerticalBlock"] {gap: 0.2rem !important;}
    div[data-testid="element-container"] {margin-bottom: 0.1rem !important;}
    div[data-testid="stAppViewBlockContainer"] {padding: 0 !important;}

    /* ---------- Top navbar ---------- */
    .amz-navbar {
        background: #0e1420;
        padding: 9px 40px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #232f3e;
    }
    .amz-logo {
        color: #ffffff;
        font-size: 1.3rem;
        font-weight: 800;
        font-style: italic;
        letter-spacing: -0.5px;
    }
    .amz-logo::after {
        content: "⌣";
        color: #ff9900;
        font-size: 1.05rem;
        margin-left: 2px;
    }
    .amz-nav-icons {
        display: flex;
        align-items: center;
        gap: 18px;
        color: #d9dee6;
        font-size: 1rem;
    }
    .amz-nav-icons .avatar {
        width: 27px; height: 27px;
        border-radius: 50%;
        background: linear-gradient(135deg,#ff9900,#7c4fe0);
        display: flex; align-items: center; justify-content: center;
        font-size: 0.8rem;
        color: #fff;
    }

    /* ---------- Hero ---------- */
    .amz-hero {
        position: relative;
        background:
            radial-gradient(circle at 88% 15%, rgba(124,79,224,0.35) 0%, rgba(124,79,224,0) 45%),
            radial-gradient(circle at 8% 100%, rgba(255,153,0,0.25) 0%, rgba(255,153,0,0) 40%),
            linear-gradient(120deg, #0b1524 0%, #14273e 45%, #223f5f 100%);
        padding: 14px 40px 12px 40px;
        overflow: hidden;
    }
    .hero-flex {display: flex; justify-content: space-between; align-items: center; gap: 20px;}
    .hero-left {flex: 1; min-width: 0;}
    .amz-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255,153,0,0.16);
        border: 1px solid rgba(255,153,0,0.55);
        color: #ffb54d;
        font-weight: 600;
        font-size: 0.7rem;
        padding: 3px 12px;
        border-radius: 20px;
        margin-bottom: 8px;
        letter-spacing: 0.2px;
    }
    .amz-hero h1 {
        color: #ffffff;
        font-size: 2.15rem;
        font-weight: 800;
        margin: 0 0 2px 0;
        line-height: 1.1;
        letter-spacing: -0.5px;
    }
    .amz-hero h1 span {
        background: linear-gradient(90deg, #ffb230, #ff8a00);
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent;
    }
    .amz-hero h2 {
        color: #eef2f7;
        font-size: 1.02rem;
        font-weight: 600;
        margin: 0 0 6px 0;
    }
    .amz-hero p.desc {
        color: #aebdd1;
        font-size: 0.82rem;
        max-width: 460px;
        line-height: 1.4;
        margin-bottom: 12px;
    }
    .amz-features {display: flex; gap: 24px;}
    .amz-feature {display: flex; align-items: center; gap: 8px;}
    .amz-feature .icon {
        width: 28px; height: 28px;
        border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 0.82rem;
        flex-shrink: 0;
        box-shadow: 0 3px 8px rgba(0,0,0,0.3);
    }
    .amz-feature .icon.orange {background: linear-gradient(135deg,#ffb230,#ff8a00);}
    .amz-feature .icon.blue {background: linear-gradient(135deg,#5fa3f7,#2f6fed);}
    .amz-feature .icon.green {background: linear-gradient(135deg,#33cb95,#0f9d68);}
    .amz-feature .label {color: #ffffff; font-weight: 700; font-size: 0.8rem; line-height: 1.15;}
    .amz-feature .sub {color: #93a2b8; font-size: 0.68rem; line-height: 1.1;}

    /* ---------- Hero decorative visual ---------- */
    .hero-visual {
        position: relative;
        width: 220px; height: 150px;
        flex-shrink: 0;
        display: flex; align-items: center; justify-content: center;
    }
    .hero-visual .blob {
        position: absolute; border-radius: 50%; filter: blur(2px); opacity: 0.55;
    }
    .hero-visual .blob.b1 {width: 110px; height: 110px; background: #ff9900; top: -10px; right: 10px; opacity: 0.35;}
    .hero-visual .blob.b2 {width: 70px; height: 70px; background: #2f6fed; bottom: -14px; left: 0; opacity: 0.35;}
    .mock-card {
        position: relative; z-index: 2;
        width: 165px; height: 108px;
        background: #ffffff;
        border-radius: 12px;
        box-shadow: 0 18px 34px rgba(0,0,0,0.4);
        padding: 8px 10px;
        transform: rotate(-3deg);
    }
    .mock-card .mc-logo {
        font-style: italic; font-weight: 800; color: #131921; font-size: 0.75rem; margin-bottom: 6px;
    }
    .mock-card .mc-logo::after {content:"⌣"; color:#ff9900; margin-left:1px;}
    .mock-bars {display: flex; align-items: flex-end; gap: 3px; height: 32px; margin-bottom: 5px;}
    .mock-bars div {width: 8px; border-radius: 2px 2px 0 0; background: linear-gradient(180deg,#5fa3f7,#2f6fed);}
    .mock-donut {
        position: absolute; right: 8px; top: 24px; width: 28px; height: 28px; border-radius: 50%;
        background: conic-gradient(#ff9900 0deg 130deg, #2f6fed 130deg 260deg, #33cb95 260deg 360deg);
    }
    .mock-donut::after {
        content: ""; position: absolute; inset: 5px; border-radius: 50%; background: #ffffff;
    }
    .mock-lines div {height: 4px; border-radius: 3px; background: #e3e8ef; margin-bottom: 3px;}
    .mock-lines div:nth-child(1) {width: 70%;}
    .mock-lines div:nth-child(2) {width: 45%;}
    .mock-xlsx {
        position: absolute; z-index: 3; right: -12px; bottom: -8px;
        width: 34px; height: 34px; border-radius: 8px;
        background: linear-gradient(160deg,#1f8f4e,#0f6b37);
        display: flex; align-items: center; justify-content: center;
        color: #fff; font-weight: 800; font-size: 0.85rem;
        box-shadow: 0 10px 18px rgba(0,0,0,0.35);
        transform: rotate(6deg);
    }
    .hero-visual .quip {
        position: absolute; top: -14px; right: -6px;
        font-style: italic; font-weight: 600; font-size: 0.62rem;
        color: #cdd8e8; line-height: 1.2; text-align: right; width: 120px;
    }
    .hero-visual .quip b {color: #ffb54d;}

    /* ---------- Content section ---------- */
    .amz-content {
        position: relative;
        padding: 10px 40px 4px 40px;
        overflow: hidden;
    }
    .amz-content .cblob {position: absolute; border-radius: 50%; filter: blur(30px); z-index: 0;}
    .amz-content .cblob.c1 {width: 160px; height: 160px; background: rgba(255,153,0,0.16); left: -40px; bottom: -50px;}
    .amz-content .cblob.c2 {width: 140px; height: 140px; background: rgba(47,111,237,0.14); right: 60px; top: -30px;}

    .amz-panel-wrap {position: relative; z-index: 1;}
    .amz-panel {
        background: #ffffff;
        border-radius: 14px;
        box-shadow: 0 10px 26px rgba(20,30,50,0.1);
        border: 1px solid #eef0f4;
        padding-bottom: 4px;
    }
    .left-card {
        background: linear-gradient(160deg,#f5f8fd,#eef2fa);
        border-radius: 12px;
        padding: 18px 16px;
        height: 100%;
        border: 1px solid #e6ecf6;
    }
    .left-card .icon-circle {
        width: 36px; height: 36px;
        border-radius: 10px;
        background: linear-gradient(135deg,#dbe6fb,#c8d9f8);
        display: flex; align-items: center; justify-content: center;
        font-size: 1rem;
        margin-bottom: 10px;
    }
    .left-card h3 {
        color: #131921;
        font-size: 1.05rem;
        font-weight: 800;
        margin: 0 0 8px 0;
        line-height: 1.2;
    }
    .left-card p {
        color: #5b6673;
        font-size: 0.79rem;
        line-height: 1.42;
        margin-bottom: 10px;
    }
    .left-card .tagline {
        font-style: italic;
        font-weight: 700;
        color: #16324f;
        font-size: 0.88rem;
        border-bottom: 2px solid #ff9900;
        display: inline-block;
        padding-bottom: 1px;
    }

    .step-row {display: flex; gap: 10px; padding: 14px 18px 2px 18px;}
    .step-row.first {padding-top: 16px;}
    .step-num {
        width: 25px; height: 25px;
        border-radius: 50%;
        color: #fff; font-weight: 700; font-size: 0.76rem;
        display: flex; align-items: center; justify-content: center;
        flex-shrink: 0;
        box-shadow: 0 3px 7px rgba(0,0,0,0.15);
    }
    .step-num.blue {background: linear-gradient(135deg,#5fa3f7,#2f6fed);}
    .step-num.purple {background: linear-gradient(135deg,#a685f2,#7c4fe0);}
    .step-title {color: #131921; font-weight: 700; font-size: 0.87rem; margin-bottom: 1px;}
    .step-sub {color: #6b7280; font-size: 0.72rem; line-height: 1.25;}

    div[data-testid="stDateInput"] {padding: 6px 18px 6px 51px; margin-top: 2px !important;}
    div[data-testid="stDateInput"] input {
        border-radius: 8px !important;
        border: 1px solid #d7dde5 !important;
        padding: 6px 10px !important;
        font-size: 0.82rem !important;
        background: #fafbfd !important;
    }

    div[data-testid="stFileUploaderDropzone"] {
        background: linear-gradient(160deg,#f6f4ff,#efeaff) !important;
        border: 2px dashed #b39ff5 !important;
        border-radius: 9px !important;
        padding: 3px !important;
    }
    div[data-testid="stFileUploader"] {padding: 6px 18px 6px 51px; margin-top: 2px !important;}
    div[data-testid="stFileUploader"] section {padding: 5px !important;}
    div[data-testid="stFileUploaderDropzoneInstructions"] span {font-size: 0.78rem !important;}
    div[data-testid="stFileUploaderDropzoneInstructions"] small {font-size: 0.67rem !important;}

    div.stButton {display: flex; justify-content: flex-end; padding: 4px 18px 10px 0;}
    div.stButton>button {
        background: linear-gradient(135deg,#ffb230,#ff8a00);
        color: #14202e;
        font-weight: 700;
        border-radius: 8px;
        padding: 0.36rem 1.35rem;
        font-size: 0.85rem;
        border: none;
        box-shadow: 0 6px 14px rgba(255,153,0,0.4);
        transition: transform 0.12s ease;
    }
    div.stButton>button:hover {transform: translateY(-1px); box-shadow: 0 8px 18px rgba(255,153,0,0.5);}

    div.stDownloadButton {display: flex; justify-content: flex-end; padding: 0 18px 10px 0;}
    div.stDownloadButton>button {
        background: linear-gradient(135deg,#33cb95,#0f9d68);
        color: #ffffff;
        font-weight: 700;
        border-radius: 8px;
        padding: 0.36rem 1.35rem;
        font-size: 0.85rem;
        border: none;
        box-shadow: 0 6px 14px rgba(15,157,104,0.4);
    }

    div[data-testid="stAlert"] {padding: 6px 10px; margin: 2px 18px 6px 18px !important; font-size: 0.8rem; border-radius: 8px;}

    /* ---------- Footer ---------- */
    .amz-footer {
        display: flex; justify-content: center; gap: 40px;
        padding: 8px 20px 12px 20px;
        color: #6b7280; font-size: 0.72rem;
    }
    .amz-footer span {display: flex; align-items: center; gap: 5px;}
    @media (max-height: 760px) {
        .amz-footer {display: none;}
    }

    /* ---------- Lock screen ---------- */
    .lock-screen {
        position: fixed;
        inset: 0;
        z-index: 999999;
        background:
            radial-gradient(circle at 85% 20%, rgba(124,79,224,0.35) 0%, rgba(124,79,224,0) 45%),
            radial-gradient(circle at 10% 90%, rgba(255,153,0,0.25) 0%, rgba(255,153,0,0) 40%),
            linear-gradient(120deg, #0b1524 0%, #14273e 45%, #223f5f 100%);
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .lock-card-outer {
        position: relative; z-index: 2;
        width: 320px;
        background: rgba(255,255,255,0.97);
        border-radius: 16px;
        box-shadow: 0 24px 60px rgba(0,0,0,0.45);
        padding: 30px 26px 18px 26px;
        text-align: center;
    }
    .lock-icon {
        width: 46px; height: 46px;
        margin: 0 auto 12px auto;
        border-radius: 50%;
        background: linear-gradient(135deg,#ffb230,#ff8a00);
        display: flex; align-items: center; justify-content: center;
        font-size: 1.3rem;
        box-shadow: 0 6px 14px rgba(255,153,0,0.4);
    }
    .lock-title {color: #131921; font-weight: 800; font-size: 1.15rem; margin-bottom: 4px;}
    .lock-sub {color: #6b7280; font-size: 0.8rem; margin-bottom: 16px;}
    .lock-card-outer div[data-testid="stTextInput"] input {
        border-radius: 8px !important;
        border: 1px solid #d7dde5 !important;
        padding: 8px 12px !important;
        font-size: 0.9rem !important;
        text-align: center;
        background: #fafbfd !important;
    }
    .lock-error {color: #e0384c; font-size: 0.78rem; margin-top: 8px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================================
# ACCESS GATE — simple shared-code lock, blurred background screen
# =========================================================================
if "unlocked" not in st.session_state:
    st.session_state.unlocked = False

if not st.session_state.unlocked:
    st.markdown('<div class="lock-screen"><div class="lock-card-outer">', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="lock-icon">🔒</div>
        <div class="lock-title">DWD Tool</div>
        <div class="lock-sub">Enter the access code to continue</div>
        """,
        unsafe_allow_html=True,
    )
    entered_code = st.text_input(
        "Access code", placeholder="Access code", label_visibility="collapsed", key="access_code_input", type="password"
    )
    if entered_code:
        if entered_code.strip().lower() == "javmuhak":
            st.session_state.unlocked = True
            st.rerun()
        else:
            st.markdown('<div class="lock-error">Incorrect code</div>', unsafe_allow_html=True)
    st.markdown("</div></div>", unsafe_allow_html=True)
    st.stop()


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
        <div class="hero-flex">
            <div class="hero-left">
                <div class="amz-badge">⚡ Amazon Tool</div>
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
            <div class="hero-visual">
                <div class="quip">Turn your data<br/>into <b>insights</b></div>
                <div class="blob b1"></div>
                <div class="blob b2"></div>
                <div class="mock-card">
                    <div class="mc-logo">amazon</div>
                    <div class="mock-lines"><div></div><div></div></div>
                    <div class="mock-bars">
                        <div style="height:14px;"></div>
                        <div style="height:22px;"></div>
                        <div style="height:18px;"></div>
                        <div style="height:30px;"></div>
                        <div style="height:26px;"></div>
                    </div>
                    <div class="mock-donut"></div>
                </div>
                <div class="mock-xlsx">X</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Main content -------------------------------------------------------
st.markdown(
    '<div class="amz-content"><div class="cblob c1"></div><div class="cblob c2"></div>'
    '<div class="amz-panel-wrap">',
    unsafe_allow_html=True,
)

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
    st.markdown(
        """
        <div class="amz-panel">
        <div class="step-row first">
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

    generate_clicked = st.button("📊  Download DWD  →", use_container_width=False)

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
            "⬇  Save DWD Report",
            data=st.session_state["report_bytes"],
            file_name=f"DWD-AUH1-{st.session_state['report_date'].strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("</div>", unsafe_allow_html=True)  # close amz-panel

st.markdown("</div></div>", unsafe_allow_html=True)  # close amz-panel-wrap + amz-content

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
