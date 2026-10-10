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
      font-weight: 700; font-size: 14px; min-height: 38px;
    }
    div.stButton > button:hover, div.stFormSubmitButton > button:hover { background: #e38800; color: white; }
    
    /* Form Dropdowns & Inputs */
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div { 
        border-radius: 4px; border-color: #888c8c; 
    }
    
    /* Cards */
    .card-wrapper { padding: 0 40px; }
    .post-card { background: white; border: 1px solid var(--line); border-radius: 8px; margin-bottom: 20px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
    .post-header { display: flex; justify-content: space-between; padding: 20px 24px 10px; }
    .post-author { display: flex; align-items: center; gap: 12px; font-weight: 700; font-size: 14px; }
    .avatar { width: 36px; height: 36px; background: #e7e9e9; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: #565959; font-size: 18px; }
    .post-status { font-size: 13px; font-weight: 600; color: #565959; display: flex; align-items: center; gap: 6px; }
    .post-body { padding: 10px 24px 20px; font-size: 14px; line-height: 1.6; color: var(--ink); white-space: pre-wrap; }
    .post-actions { padding: 0 24px 20px; display: flex; gap: 15px; }
    
    /* Upvote text */
    .upvote-text { display: inline-flex; align-items: center; gap: 8px; color: var(--ink); font-weight: 600; font-size: 14px; cursor: pointer; }
    .upvote-text:hover { text-decoration: underline; }
    
    /* Response */
    .response-section { padding: 20px 24px; border-top: 1px solid #eaeded; }
    .response-header { font-weight: 700; font-size: 15px; margin-bottom: 12px; color: var(--ink); }
    .response-box-inner { background: #f8f9fa; border-radius: 8px; padding: 16px; border: 1px solid #eaeded; }
    .responder-info { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; }
    .responder-name { font-weight: 700; font-size: 14px; color: var(--ink); }
    .responder-title { font-size: 12px; color: var(--muted); }
    .response-text { font-size: 14px; line-height: 1.6; color: var(--ink); }
    
    /* Ghost button for filters */
    .ghost-btn { background: white !important; color: var(--ink) !important; border: 1px solid #888c8c !important; border-radius: 20px !important; }
    .ghost-btn:hover { background: #f7fafa !important; }
    
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


@st.cache_data(show_spinner=False)
def load_roster():
  if not ROSTER_PATH.exists():
    return None, f"Roster workbook not found: {ROSTER_PATH.name}"
  try:
    df = pd.read_excel(ROSTER_PATH, sheet_name="Roster", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    normalized = {str(c).strip().casefold(): c for c in df.columns}
    id_col = (
        normalized.get("amz id")
        or normalized.get("amz_id")
        or normalized.get("amzid")
    )
    name_col = (
        normalized.get("employee name")
        or normalized.get("employee_name")
        or normalized.get("name")
    )
    if not id_col or not name_col:
      return None, "Missing AMZ ID or Name columns."
    df = df[[id_col, name_col]].copy()
    df.columns = ["AMZ ID", "Employee Name"]
    df = df.dropna(subset=["AMZ ID", "Employee Name"])
    df["AMZ ID"] = df["AMZ ID"].astype(str).str.strip()
    df["Employee Name"] = df["Employee Name"].astype(str).str.strip()
    df = df[(df["AMZ ID"] != "") & (df["Employee Name"] != "")]
    return df, None
  except Exception as exc:
    return None, f"Could not read the Roster sheet: {exc}"


def safe(value):
  return html.escape("" if value is None else str(value))


def is_management_user():
  return str(st.session_state.get("amz_id", "")).strip().lower() in MANAGEMENT_IDS


def format_time_ago(date_str):
  try:
    dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
    hours = int((datetime.now() - dt).total_seconds() / 3600)
    if hours == 0:
      return "Just now"
    return f"{hours}h"
  except:
    return "8h"


init_db()
roster_df, roster_error = load_roster()

if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.amz_id = ""
  st.session_state.emp_name = ""
  st.session_state.site = SITE
  st.session_state.show_form = False


# ----------------------------- Login ---------------------------------
if not st.session_state.logged_in:
  
  st.write("")
  st.write("")
  st.write("")

  col1, col2, col3 = st.columns([1, 1.2, 1])
  with col2:
    with st.form("login_form", clear_on_submit=False):
      st.markdown(
          f"""<div style="text-align: left; margin-bottom: 20px;">
{login_logo}
<div style="font-size: 11px; font-weight: 700; color: #007185; margin-bottom: 5px; letter-spacing: 1px; text-transform: uppercase;">Your voice matters</div>
<div style="font-size: 26px; font-weight: 700; margin-bottom: 10px; color: #0f1111;">Welcome to QuessConnect</div>
<div style="color: #565959; font-size: 14px; margin-bottom: 25px;">Share feedback, raise a concern, or recognise something that is working well.</div>
<h3 style="margin-bottom: 5px; font-size: 16px; color: #0f1111;">Associate login</h3>
<p style="color: #565959; font-size: 13px; margin-bottom: 10px;">Sign in with your roster AMZ ID. Site is set to DXB5.</p>
</div>""",
          unsafe_allow_html=True,
      )
      
      amz_input = st.text_input(
          "AMZ ID", placeholder="Enter your AMZ ID", label_visibility="collapsed", autocomplete="username"
      )
      login_btn = st.form_submit_button(
          "Continue", type="primary", use_container_width=True
      )
      
    if login_btn:
      amz_clean = str(amz_input).strip().lower()
      if not amz_clean:
        st.error("Please enter your AMZ ID.")
      elif amz_clean in MANAGEMENT_IDS:
        st.session_state.logged_in = True
        st.session_state.amz_id = amz_clean
        st.session_state.emp_name = (
            "QuessCorp Management" if amz_clean == "quesscorp" else "Javmuhak"
        )
        st.session_state.site = SITE
        st.rerun()
      elif roster_df is None:
        st.error(roster_error or "Roster could not be loaded.")
      else:
        matches = roster_df[roster_df["AMZ ID"].str.lower() == amz_clean]
        if matches.empty:
          st.error("AMZ ID not found in the roster.")
        else:
          st.session_state.logged_in = True
          st.session_state.amz_id = amz_clean
          st.session_state.emp_name = str(matches.iloc[0]["Employee Name"])
          st.session_state.site = SITE
          st.rerun()

# ----------------------------- Portal / Feed --------------------------------
else:
  management = is_management_user()

  # Top Header inside portal
  st.markdown(
      f"""
      <div class="myvoice-header">
          {header_logo}
          <div class="myvoice-tab">Home</div>
      </div>
      """,
      unsafe_allow_html=True,
  )

  # Action Bar (Site Dropdown, Search, Filters, New Comment)
  st.markdown('<div class="controls-wrapper">', unsafe_allow_html=True)
  c1, c2, c3, c4 = st.columns([1.5, 3.5, 1, 1.2])
  
  with c1:
    st.selectbox("Site", ["DXB5 QuessConnect"], label_visibility="collapsed")
  with c2:
    search_query = st.text_input("Search", placeholder="🔍 Search comments by keyword, author, etc.", label_visibility="collapsed")
  with c3:
    if st.button("Log out", use_container_width=True):
      for key in ["logged_in", "amz_id", "emp_name", "site", "show_form"]:
        st.session_state[key] = False if key == "logged_in" or key == "show_form" else ""
      st.rerun()
  with c4:
    if st.button("+ New Comment", type="primary", use_container_width=True):
      st.session_state.show_form = not st.session_state.show_form
  st.markdown('</div>', unsafe_allow_html=True)


  if st.session_state.show_form:
    st.markdown('<div class="card-wrapper">', unsafe_allow_html=True)
    with st.form("new_comment_form", clear_on_submit=True):
      st.markdown("### Create New Comment")
      category = st.selectbox("Category", CATEGORIES)
      message = st.text_area("Your comment", placeholder="Type your message here...", height=120)
      post_anonymous = st.checkbox("Post as Anonymous (Hide identity on feed)", value=True)
      
      f1, f2 = st.columns([1, 1])
      with f1:
          submit_comment = st.form_submit_button("Submit", type="primary", use_container_width=True)
      with f2:
          cancel_comment = st.form_submit_button("Cancel", use_container_width=True)
          
    if cancel_comment:
      st.session_state.show_form = False
      st.rerun()
    if submit_comment:
      if not message.strip():
        st.error("Comment cannot be empty.")
      else:
        with db_connect() as conn:
          conn.execute(
              """INSERT INTO complaints
                          (amz_id, employee_name, site, is_anonymous, category, message, status, reply, upvotes)
                          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
              (
                  st.session_state.amz_id,
                  st.session_state.emp_name,
                  SITE,
                  1 if post_anonymous else 0,
                  category,
                  message.strip(),
                  "Open",
                  "",
                  0,
              ),
          )
        st.session_state.show_form = False
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


  # Fetch Feed
  with db_connect() as conn:
    all_posts = pd.read_sql_query("SELECT * FROM complaints ORDER BY id DESC", conn)

  st.markdown('<div class="card-wrapper">', unsafe_allow_html=True)
  
  if all_posts.empty:
    st.info("No comments yet.")
  else:
    feed_df = all_posts.copy()
    if search_query.strip():
      q = search_query.strip()
      mask = (
          feed_df["message"].astype(str).str.contains(q, case=False, na=False)
          | feed_df["category"].astype(str).str.contains(q, case=False, na=False)
      )
      feed_df = feed_df[mask]

    for _, post in feed_df.iterrows():
      anonymous = int(post["is_anonymous"] or 0) == 1
      time_ago = format_time_ago(post["date"])
      status = str(post["status"] or "Open")
      
      if anonymous:
          author_name = "Anonymous"
          avatar = "👤"
          if management:
              author_name = f"Anonymous (Real: {safe(post['employee_name'])} @{safe(post['amz_id'])})"
      else:
          author_name = f"{safe(post['employee_name'])}@"
          avatar = "🧑🏾‍🦱"

      message_html = safe(post["message"])
      upvotes = int(post["upvotes"] or 0)

      # Main Card Layout
      st.markdown(
          f"""
          <div class="post-card">
              <div class="post-header">
                  <div class="post-author">
                      <div class="avatar">{avatar}</div>
                      <div>{author_name} &nbsp;·&nbsp; {time_ago}</div>
                  </div>
                  <div class="post-status">
                      🕑 {safe(status)}
                  </div>
              </div>
              
              <div class="post-body">
                  {message_html}
              </div>
              
              <div class="post-actions">
                  <div class="upvote-text">＋ {upvotes}</div>
              </div>
          """, unsafe_allow_html=True
      )

      # Upvote Button (Streamlit overlay)
      c_up, c_space = st.columns([1, 10])
      with c_up:
         if st.button("Upvote", key=f"up_{post['id']}", help="Like this post"):
             with db_connect() as conn:
                 conn.execute("UPDATE complaints SET upvotes = upvotes + 1 WHERE id = ?", (int(post["id"]),))
             st.rerun()

      # Response Section
      reply = str(post["reply"] or "").strip()
      if reply:
          st.markdown(
              f"""
              <div class="response-section">
                  <div class="response-header">Response</div>
                  <div class="response-box-inner">
                      <div class="responder-info">
                          <div class="avatar">👩🏻‍💼</div>
                          <div>
                              <div class="responder-name">QuessCorp HR Partner@ &nbsp;·&nbsp; Just now</div>
                              <div class="responder-title">HR Associate Partner</div>
                          </div>
                      </div>
                      <div class="response-text">{safe(reply)}</div>
                  </div>
              </div>
              """, unsafe_allow_html=True
          )

      st.markdown("</div>", unsafe_allow_html=True) # close post-card

      # Management Edit
      if management:
        with st.expander(f"⚙️ Manage post #{int(post['id'])}"):
          with st.form(f"manage_{int(post['id'])}"):
            selected_status = st.selectbox("Status", STATUSES, index=STATUSES.index(status) if status in STATUSES else 0)
            response_text = st.text_area("HR Response", value=reply)
            if st.form_submit_button("Save"):
              with db_connect() as conn:
                conn.execute("UPDATE complaints SET status = ?, reply = ? WHERE id = ?", (selected_status, response_text.strip(), int(post["id"])))
              st.rerun()

  st.markdown('</div>', unsafe_allow_html=True) # close wrapper
