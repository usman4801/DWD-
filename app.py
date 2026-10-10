import html
import sqlite3
from pathlib import Path

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# QuessCorp | Voice of Associate (VoA)
# Run: streamlit run app.py
# Keep DMD-DXB5-07102026.xlsx beside this file (sheet: Roster).
# ---------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
ROSTER_PATH = APP_DIR / "DMD-DXB5-07102026.xlsx"
DB_PATH = APP_DIR / "quess_voa.db"
SITE = "DXB5"
MANAGEMENT_IDS = {"quesscorp", "javmuhak"}
CATEGORIES = [
    "Payroll / Salary",
    "Transport / Attendance",
    "Workplace Management",
    "Medical / Insurance",
    "Appreciation / Feedback",
]
STATUSES = ["Open", "In Progress", "Resolved", "Closed"]

st.set_page_config(
    page_title="QuessCorp | Voice of Associate",
    page_icon="🗣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
    :root {
      --quess-blue: #004b87; --quess-deep: #07345b; --ink: #172b3a;
      --muted: #687989; --line: #e3eaf0; --canvas: #f4f7fa;
    }
    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
    .stApp { background: var(--canvas); color: var(--ink); }
    [data-testid="stHeader"] { background: rgba(244,247,250,.92); }
    .block-container { max-width: 1450px; padding-top: 1.35rem; padding-bottom: 3rem; }
    [data-testid="stSidebar"] { display: none; }
    h1, h2, h3, h4 { font-family: 'Manrope', sans-serif; color: var(--ink); letter-spacing: -.025em; }
    .brandbar {
      background: linear-gradient(115deg, #06345b 0%, #004b87 62%, #0876b9 100%);
      border-radius: 18px; padding: 20px 24px; color: white;
      box-shadow: 0 12px 28px rgba(0,75,135,.13); margin-bottom: 20px;
    }
    .brand-row { display:flex; align-items:center; justify-content:space-between; gap:20px; }
    .brand-lockup { display:flex; align-items:center; gap:16px; }
    .brand-name { font-family:'Manrope',sans-serif; font-size:21px; line-height:1.1; font-weight:800; }
    .brand-title { font-size:16px; font-weight:700; margin-top:3px; }
    .brand-sub { font-size:12px; color:rgba(255,255,255,.76); margin-top:2px; }
    .site-pill { border:1px solid rgba(255,255,255,.35); border-radius:999px; padding:6px 12px; font-size:12px; font-weight:700; white-space:nowrap; }
    .page-title { font-family:'Manrope',sans-serif; font-size:24px; font-weight:800; margin:0 0 4px; }
    .page-subtitle { color:var(--muted); font-size:13px; margin-bottom:15px; }
    .post-card {
      background:#fff; border:1px solid var(--line); border-radius:16px;
      padding:21px 23px; margin:12px 0 8px; box-shadow:0 3px 12px rgba(17,46,70,.035);
    }
    .post-top { display:flex; justify-content:space-between; align-items:flex-start; gap:14px; }
    .author { font-weight:700; font-size:14px; color:var(--ink); }
    .post-date { color:#8997a3; font-size:11px; margin-top:4px; }
    .category-chip { display:inline-block; border-radius:6px; background:#edf5fb; color:#07548c; padding:5px 9px; font-size:11px; font-weight:700; margin:13px 0 10px; }
    .status-chip { display:inline-block; border-radius:999px; padding:5px 10px; font-size:11px; font-weight:800; white-space:nowrap; }
    .status-open { background:#fff2d9; color:#8a5b00; }
    .status-progress { background:#e5f1ff; color:#07599a; }
    .status-resolved { background:#e4f7ec; color:#147442; }
    .status-closed { background:#edf0f3; color:#526170; }
    .post-message { font-size:14px; line-height:1.72; color:#34495a; white-space:pre-wrap; overflow-wrap:anywhere; }
    .anon-chip { display:inline-block; margin-left:8px; border-radius:5px; background:#fff0dc; color:#8d5600; padding:3px 6px; font-size:10px; font-weight:800; vertical-align:2px; }
    .response-box { background:#f2f7fb; border:1px solid #dceaf5; border-left:3px solid var(--quess-blue); border-radius:9px; padding:14px 16px; margin:12px 0; }
    .response-title { color:var(--quess-blue); font-size:12px; font-weight:800; margin-bottom:5px; }
    .response-text { color:#435565; font-size:13px; line-height:1.6; white-space:pre-wrap; }
    .login-wrap { max-width:450px; margin:4rem auto 0; }
    .login-panel { background:#fff; border:1px solid var(--line); border-radius:18px; padding:30px 30px 20px; box-shadow:0 12px 32px rgba(18,47,72,.06); }
    div.stButton > button, div.stFormSubmitButton > button {
      border-radius:9px; font-weight:700; min-height:42px; border:1px solid #d5e0e9;
      transition:all .15s ease;
    }
    div.stButton > button[kind="primary"], div.stFormSubmitButton > button[kind="primary"] {
      background:var(--quess-blue); border-color:var(--quess-blue); color:white;
    }
    div.stButton > button[kind="primary"]:hover, div.stFormSubmitButton > button[kind="primary"]:hover {
      background:#003761; border-color:#003761; color:white;
    }
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div {
      border-radius:9px; border-color:#dce4eb;
    }
    [data-testid="stMetric"] { background:white; border:1px solid var(--line); border-radius:13px; padding:14px 16px; }
    [data-testid="stMetricLabel"] { color:var(--muted); }
    [data-testid="stMetricValue"] { color:var(--quess-blue); font-family:'Manrope',sans-serif; }
    div[data-testid="stExpander"] { background:#fff; border:1px solid var(--line); border-radius:11px; }
    hr { border-color:var(--line); }
    .footer { color:#93a0ab; text-align:center; font-size:11px; padding:24px 0 5px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def db_connect():
  conn = sqlite3.connect(DB_PATH, timeout=15)
  conn.row_factory = sqlite3.Row
  return conn


def init_db():
  with db_connect() as conn:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            amz_id TEXT NOT NULL,
            employee_name TEXT NOT NULL,
            site TEXT NOT NULL DEFAULT 'DXB5',
            is_anonymous INTEGER NOT NULL DEFAULT 0,
            category TEXT NOT NULL,
            message TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Open',
            reply TEXT DEFAULT '',
            upvotes INTEGER NOT NULL DEFAULT 0,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )


@st
