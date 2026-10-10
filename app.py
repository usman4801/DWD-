import html
import sqlite3
import base64
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# QuessCorp | QuessConnect Portal
# Run: streamlit run app.py
# Keep DMD-DXB5-07102026.xlsx and image_2c8c24.png beside this file.
# ---------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
ROSTER_PATH = APP_DIR / "DMD-DXB5-07102026.xlsx"
DB_PATH = APP_DIR / "quessconnect.db"
LOGO_PATH = APP_DIR / "image_2c8c24.png"  # Quess Logo File
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
    page_title="QuessConnect",
    page_icon="🗣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Function to encode local image to display in HTML
def get_base64_image(path):
  try:
    with open(path, "rb") as f:
      return base64.b64encode(f.read()).decode()
  except Exception:
    return ""

logo_b64 = get_base64_image(LOGO_PATH)
if logo_b64:
  login_logo = f'<img src="data:image/png;base64,{logo_b64}" style="max-height: 55px; margin-bottom: 15px;">'
  header_logo = f'<img src="data:image/png;base64,{logo_b64}" style="max-height: 28px; margin-bottom: 12px; margin-right: 15px;">'
else:
  login_logo = '<h2 style="color: #004b87; margin: 0; font-weight: 800; margin-bottom: 15px;">QUESS</h2>'
  header_logo = '<div style="font-size: 22px; font-weight: 800; color: #0f1111; padding-bottom: 15px; margin-right: 15px;">QuessConnect</div>'


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
    html, body, [class*="css"] { font-family: 'Amazon Ember', Arial, sans-serif; overflow-x: hidden; }
    .stApp { background: var(--bg); color: var(--ink); }
    [data-testid="stHeader"] { display: none; }
    [data-testid="stSidebar"] { display: none; }
    
    /* Removed padding so it fits screen perfectly without scrolling */
    .block-container { max-width: 1400px; padding-top: 2rem !important; padding-bottom: 0 !important; padding-left: 0 !important; padding-right: 0 !important; }
    
    /* Streamlit Form Styling */
    [data-testid="stForm"] { 
        background: white; 
        border: 1px solid var(--line); 
        border-radius: 8px; 
        padding: 40px; 
        box-shadow: 0 1px 2px rgba(0,0,0,0.05); 
    }
    
    /* Top Header */
    .myvoice-header {
      background: white; border-bottom: 1px solid var(--line);
      padding: 10px 40px 0px 40px; margin-bottom: 30px; margin-top: -2rem;
      display: flex; align-items: flex-end; gap: 20px;
    }
    .myvoice-tab { font-size: 15px; font-weight: 700; color: var(--ink); border-bottom: 3px solid var(--ink); padding-bottom: 12px; }
    
    /* Controls row */
    .controls-wrapper { padding: 0 40px; margin-bottom: 20px; }
    
    /* Buttons */
    div.stButton > button, div.stFormSubmitButton > button {
      background: var(--orange); color: white; border: none; border-radius: 20px;
      font-weight: 700; font-size: 14px; min-height:
