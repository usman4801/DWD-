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
      border-radius: 18px; padding: 23px 28px; color: white;
      box-shadow: 0 12px 28px rgba(0,75,135,.13); margin-bottom: 20px;
    }
    .brand-row { display:flex; align-items:center; justify-content:space-between; gap:20px; }
    .brand-lockup { display:flex; align-items:center; gap:16px; }
    .brand-mark {
      display:flex; align-items:center; justify-content:center; width:54px; height:54px;
      border:1px solid rgba(255,255,255,.35); border-radius:15px;
      background:rgba(255,255,255,.12); font-family:'Manrope',sans-serif;
      font-size:17px; font-weight:800; letter-spacing:-.05em;
    }
    .brand-name { font-family:'Manrope',sans-serif; font-size:23px; line-height:1.1; font-weight:800; }
    .brand-tag { font-size:10px; letter-spacing:2px; opacity:.8; margin-top:5px; font-weight:700; }
    .brand-title { font-size:19px; font-weight:700; margin-top:4px; }
    .brand-sub { font-size:12px; color:rgba(255,255,255,.76); margin-top:4px; }
    .site-pill { border:1px solid rgba(255,255,255,.35); border-radius:999px; padding:8px 13px; font-size:12px; font-weight:700; white-space:nowrap; }
    .eyebrow { color:var(--quess-blue); text-transform:uppercase; letter-spacing:1.6px; font-size:10px; font-weight:800; }
    .page-title { font-family:'Manrope',sans-serif; font-size:27px; font-weight:800; letter-spacing:-.04em; margin:3px 0 4px; }
    .page-subtitle { color:var(--muted); font-size:14px; margin-bottom:18px; }
    .metric-card {
      background:#fff; border:1px solid var(--line); border-radius:14px; padding:16px 18px;
      min-height:94px; box-shadow:0 2px 8px rgba(20,50,75,.025);
    }
    .metric-label { color:var(--muted); font-size:12px; font-weight:600; }
    .metric-value { color:var(--ink); font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; margin-top:4px; }
    .metric-note { color:#8796a4; font-size:11px; margin-top:1px; }
    .section-head { display:flex; align-items:center; justify-content:space-between; margin:20px 0 12px; }
    .section-title { font-family:'Manrope',sans-serif; font-size:18px; font-weight:800; }
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
    .login-wrap { max-width:610px; margin:3.2rem auto 0; }
    .login-panel { background:#fff; border:1px solid var(--line); border-radius:18px; padding:26px 28px 12px; box-shadow:0 12px 32px rgba(18,47,72,.06); }
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
    @media (max-width: 700px) {
      .block-container { padding-left:1rem; padding-right:1rem; }
      .brandbar { padding:18px; }
      .brand-name { font-size:19px; }
      .brand-title { font-size:16px; }
      .post-card { padding:17px; }
    }
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
      return (
          None,
          "The Roster sheet must contain 'AMZ ID' and 'Employee Name' columns.",
      )
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


def status_class(status):
  return {
      "Open": "status-open",
      "In Progress": "status-progress",
      "Resolved": "status-resolved",
      "Closed": "status-closed",
  }.get(status, "status-open")


def brand_header(logged_in=True):
  if logged_in:
    identity = f"Welcome back, <b>{safe(st.session_state.get('emp_name', 'Associate'))}</b> · @{safe(st.session_state.get('amz_id', ''))}"
    role = "MANAGEMENT ACCESS" if is_management_user() else "ASSOCIATE PORTAL"
  else:
    identity = "A better workplace starts with your voice."
    role = "ASSOCIATE PORTAL"
  st.markdown(
      f"""
        <div class="brandbar">
          <div class="brand-row">
            <div class="brand-lockup">
              <div class="brand-mark">QUESS</div>
              <div>
                <div class="brand-name">QuessCorp</div>
                <div class="brand-tag">WINNING TOGETHER</div>
                <div class="brand-title">Voice of Associate <span style="opacity:.65;font-weight:500">/ VoA</span></div>
                <div class="brand-sub">{identity}</div>
              </div>
            </div>
            <div class="site-pill">DXB5 &nbsp;·&nbsp; {role}</div>
          </div>
        </div>
        """,
      unsafe_allow_html=True,
  )


def is_management_user():
  return str(st.session_state.get("amz_id", "")).strip().lower() in MANAGEMENT_IDS


def metric_card(label, value, note):
  st.markdown(
      f'<div class="metric-card"><div class="metric-label">{safe(label)}</div>'
      f'<div class="metric-value">{safe(value)}</div><div class="metric-note">{safe(note)}</div></div>',
      unsafe_allow_html=True,
  )


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
  brand_header(logged_in=False)
  st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
  st.markdown(
      """
        <div class="eyebrow">Your voice matters</div>
        <div class="page-title">Welcome to VoA</div>
        <div class="page-subtitle">Share feedback, raise a concern, or recognise something that is working well.</div>
        """,
      unsafe_allow_html=True,
  )
  st.markdown('<div class="login-panel">', unsafe_allow_html=True)
  st.markdown("### Associate login")
  st.caption("Sign in with your roster AMZ ID. Your site is set to DXB5.")
  with st.form("login_form", clear_on_submit=False):
    amz_input = st.text_input(
        "AMZ ID", placeholder="Enter your AMZ ID", autocomplete="username"
    )
    login_btn = st.form_submit_button(
        "Continue to portal", type="primary", use_container_width=True
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
        st.error(
            "That AMZ ID was not found in the roster. Please check and try"
            " again."
        )
      else:
        st.session_state.logged_in = True
        st.session_state.amz_id = amz_clean
        st.session_state.emp_name = str(matches.iloc[0]["Employee Name"])
        st.session_state.site = SITE
        st.rerun()
  st.markdown("</div>", unsafe_allow_html=True)
  if roster_error:
    st.warning(roster_error)
  st.markdown(
      '<div class="footer">QUESSCORP · DXB5 · VOICE OF ASSOCIATE<br>Feedback'
      " is a shared step toward a better workplace.</div>",
      unsafe_allow_html=True,
  )

# ----------------------------- Portal --------------------------------
else:
  management = is_management_user()
  brand_header()

  top_left, top_right = st.columns([6, 1])
  with top_left:
    st.markdown(
        '<div class="eyebrow">Associate feedback hub</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-title">Your voice. Our workplace.</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-subtitle">One place to raise concerns, share ideas,'
        " and track progress.</div>",
        unsafe_allow_html=True,
    )
  with top_right:
    st.write("")
    st.write("")
    if st.button("Log out", use_container_width=True):
      for key in ["logged_in", "amz_id", "emp_name", "site", "show_form"]:
        st.session_state[key] = (
            False if key == "logged_in" or key == "show_form" else ""
        )
      st.rerun()

  with db_connect() as conn:
    all_posts = pd.read_sql_query("SELECT * FROM complaints ORDER BY id DESC", conn)

  total = len(all_posts)
  open_count = (
      int(all_posts["status"].isin(["Open", "In Progress"]).sum())
      if total
      else 0
  )
  resolved_count = (
      int(all_posts["status"].isin(["Resolved", "Closed"]).sum()) if total else 0
  )
  votes_count = int(all_posts["upvotes"].fillna(0).sum()) if total else 0
  m1, m2, m3, m4 = st.columns(4)
  with m1:
    metric_card("TOTAL VOICES", total, "All posts submitted")
  with m2:
    metric_card("NEEDS ATTENTION", open_count, "Open or in progress")
  with m3:
    metric_card("RESOLVED / CLOSED", resolved_count, "Updates from management")
  with m4:
    metric_card("COMMUNITY SUPPORT", votes_count, "Total upvotes")

  st.markdown(
      '<div class="section-head"><div class="section-title">Community'
      " feed</div></div>",
      unsafe_allow_html=True,
  )
  filter_col, search_col, status_col, button_col = st.columns(
      [1.55, 2.4, 1.35, 1.35]
  )
  with filter_col:
    category_filter = st.selectbox(
        "Category", ["All categories"] + CATEGORIES, label_visibility="collapsed"
    )
  with search_col:
    search_query = st.text_input(
        "Search posts",
        placeholder="Search feedback...",
        label_visibility="collapsed",
    )
  with status_col:
    status_filter = st.selectbox(
        "Status", ["All statuses"] + STATUSES, label_visibility="collapsed"
    )
  with button_col:
    if st.button("+ New Comment", type="primary", use_container_width=True):
      st.session_state.show_form = not st.session_state.show_form

  if st.session_state.show_form:
    with st.container(border=True):
      form_head, form_hint = st.columns([3, 2])
      with form_head:
        st.markdown("### Share your voice")
        st.caption(
            "Please be respectful and include enough detail for the team to"
            " understand your feedback."
        )
      with form_hint:
        st.info("Your post will be visible on the DXB5 community feed.")
      with st.form("new_comment_form", clear_on_submit=True):
        category = st.selectbox("Feedback category", CATEGORIES)
        message = st.text_area(
            "Your comment",
            placeholder=(
                "What happened? What would you like the team to know?"
            ),
            height=150,
            max_chars=5000,
        )
        post_anonymous = st.checkbox(
            "Post as Anonymous (Your name will be hidden on the public feed)",
            value=False,
        )
        left, right = st.columns([1, 1])
        with left:
          submit_comment = st.form_submit_button(
              "Publish comment", type="primary", use_container_width=True
          )
        with right:
          cancel_comment = st.form_submit_button("Cancel", use_container_width=True)
      if cancel_comment:
        st.session_state.show_form = False
        st.rerun()
      if submit_comment:
        cleaned_message = message.strip()
        if not cleaned_message:
          st.error("Please add a comment before publishing.")
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
                    cleaned_message,
                    "Open",
                    "",
                    0,
                ),
            )
          st.session_state.show_form = False
          st.success("Your comment has been published.")
          st.rerun()

  st.markdown("---")
  if all_posts.empty:
    st.markdown("### It starts with one voice")
    st.info(
        "There are no posts yet. Use **+ New Comment** to share the first"
        " piece of feedback."
    )
  else:
    feed_df = all_posts.copy()
    if category_filter != "All categories":
      feed_df = feed_df[feed_df["category"] == category_filter]
    if status_filter != "All statuses":
      feed_df = feed_df[feed_df["status"] == status_filter]
    if search_query.strip():
      q = search_query.strip()
      mask = (
          feed_df["message"].astype(str).str.contains(q, case=False, na=False)
          | feed_df["category"]
          .astype(str)
          .str.contains(q, case=False, na=False)
          | feed_df["status"].astype(str).str.contains(q, case=False, na=False)
      )
      feed_df = feed_df[mask]

    st.caption(f"Showing {len(feed_df)} of {total} posts")
    if feed_df.empty:
      st.info(
          "No posts match those filters. Try a different search or category."
      )
    for _, post in feed_df.iterrows():
      anonymous = int(post["is_anonymous"] or 0) == 1
      if management:
        author = f"{safe(post['employee_name'])} <span style='color:#718293;font-weight:500'>@{safe(post['amz_id'])}</span>"
        anon_badge = (
            '<span class="anon-chip">ANONYMOUS REPORT</span>'
            if anonymous
            else ""
        )
      else:
        author = (
            "Anonymous"
            if anonymous
            else f"{safe(post['employee_name'])} <span style='color:#718293;font-weight:500'>@{safe(post['amz_id'])}</span>"
        )
        anon_badge = ""

      status = str(post["status"] or "Open")
      category_html = safe(post["category"])
      date_html = safe(post["date"])
      message_html = safe(post["message"])
      st.markdown(
          f"""
                <div class="post-card">
                  <div class="post-top">
                    <div>
                      <div class="author">◉ &nbsp;{author}{anon_badge}</div>
                      <div class="post-date">{date_html} &nbsp;·&nbsp; Site DXB5</div>
                    </div>
                    <span class="status-chip {status_class(status)}">{safe(status)}</span>
                  </div>
                  <div class="category-chip">{category_html}</div>
                  <div class="post-message">{message_html}</div>
                </div>
                """,
          unsafe_allow_html=True,
      )

      actions_col, spacer_col = st.columns([1.25, 6])
      with actions_col:
        if st.button(
            f"♡  Support · {int(post['upvotes'] or 0)}",
            key=f"up_{post['id']}",
            use_container_width=True,
        ):
          with db_connect() as conn:
            conn.execute(
                "UPDATE complaints SET upvotes = upvotes + 1 WHERE id = ?",
                (int(post["id"]),),
            )
          st.rerun()

      reply = str(post["reply"] or "").standard_trim() if hasattr(str(post["reply"] or ""), "standard_trim") else str(post["reply"] or "").strip()
      if reply:
        st.markdown(
            f"""<div class="response-box">
                    <div class="response-title">✦ &nbsp;QUESSCORP MANAGEMENT RESPONSE</div>
                    <div class="response-text">{safe(reply)}</div>
                    </div>""",
            unsafe_allow_html=True,
        )

      if management:
        with st.expander(
            f"Manage post #{int(post['id'])} · Update status or publish a"
            " response"
        ):
          with st.form(f"manage_{int(post['id'])}"):
            selected_status = st.selectbox(
                "Post status",
                STATUSES,
                index=(
                    STATUSES.index(status) if status in STATUSES else 0
                ),
                key=f"status_{int(post['id'])}",
            )
            response_text = st.text_area(
                "Official management response",
                value=reply,
                placeholder="Write an update for the associate community...",
                height=110,
                key=f"reply_{int(post['id'])}",
            )
            save = st.form_submit_button(
                "Save management update", type="primary"
            )
          if save:
            with db_connect() as conn:
              conn.execute(
                  "UPDATE complaints SET status = ?, reply = ? WHERE id = ?",
                  (selected_status, response_text.strip(), int(post["id"])),
              )
            st.success("Management update saved.")
            st.rerun()

  st.markdown(
      '<div class="footer">QUESSCORP · DXB5 · VOICE OF ASSOCIATE<br>Listen with'
      " care. Respond with action.</div>",
      unsafe_allow_html=True,
  )
