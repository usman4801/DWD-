import html
import sqlite3
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# QuessCorp | Voice of Associate (MyVoice Clone)
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
    page_title="MyVoice",
    page_icon="🗣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Amazon+Ember:wght@400;500;700&display=swap');
    
    :root {
      --orange: #ff9900; 
      --ink: #0f1111; 
      --bg: #f2f4f8; 
      --line: #d5d9d9;
      --muted: #565959;
    }
    html, body, [class*="css"] { font-family: 'Amazon Ember', Arial, sans-serif; }
    .stApp { background: var(--bg); color: var(--ink); }
    [data-testid="stHeader"] { display: none; }
    [data-testid="stSidebar"] { display: none; }
    .block-container { max-width: 1400px; padding-top: 0 !important; padding-left: 0 !important; padding-right: 0 !important; }
    
    /* Top Header */
    .myvoice-header {
      background: white; border-bottom: 1px solid var(--line);
      padding: 18px 40px 0px 40px; margin-bottom: 30px;
      display: flex; align-items: flex-end; gap: 30px;
    }
    .myvoice-logo { font-size: 22px; font-weight: 800; color: var(--ink); padding-bottom: 15px; }
    .myvoice-tab { font-size: 15px; font-weight: 700; color: var(--ink); border-bottom: 3px solid var(--ink); padding-bottom: 12px; }
    
    /* Login */
    .login-box { max-width: 450px; margin: 60px auto; background: white; padding: 40px; border: 1px solid var(--line); border-radius: 8px; }
    .login-title { font-size: 28px; font-weight: 700; margin-bottom: 10px; }
    
    /* Controls row */
    .controls-wrapper { padding: 0 40px; margin-bottom: 20px; }
    
    /* Buttons */
    div.stButton > button, div.stFormSubmitButton > button {
      background: var(--orange); color: white; border: none; border-radius:
