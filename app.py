import html
import sqlite3
import base64
import hashlib
from pathlib import Path
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# QuessCorp | QuessConnect Portal
# Run: streamlit run app.py
# Keep DMD-DXB5-07102026.xlsx and image_2c8c24.png beside this file.
# Needs Streamlit 1.39 or newer (keyed containers are used for the card styling).
# ---------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
ROSTER_PATH = APP_DIR / "DMD-DXB5-07102026.xlsx"
DB_PATH = APP_DIR / "quessconnect.db"
LOGO_PATH = APP_DIR / "image_2c8c24.png"  # Quess Logo File
SITE = "DXB5"
MANAGEMENT_IDS = {"quesscorp", "javmuhak"}   # admins: can reply to and delete posts
CATEGORIES = [
    "Payroll / Salary",
    "Transport / Attendance",
    "Workplace Management",
    "Medical / Insurance",
    "Appreciation / Feedback",
]
STATUSES = ["Open", "In Progress", "Resolved", "Closed"]
STATUS_CLASS = {"Open": "s-open", "In Progress": "s-prog", "Resolved": "s-done", "Closed": "s-closed"}

# Same avatar for the same person every time, no picker. Colours are neutral (suit anyone).
AVATAR_COLORS = ["#6366f1", "#0ea5a4", "#e11d48", "#d97706", "#7c3aed", "#2563eb", "#16a34a", "#db2777"]

st.set_page_config(
    page_title="QuessConnect",
    page_icon="🗣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def get_base64_image(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""


logo_b64 = get_base64_image(LOGO_PATH)
if logo_b64:
    login_logo = f'<img src="data:image/png;base64,{logo_b64}" style="max-height:55px;margin-bottom:15px;">'
    header_logo = f'<img src="data:image/png;base64,{logo_b64}" class="qc-logo">'
else:
    login_logo = '<h2 style="color:#004b87;margin:0;font-weight:800;margin-bottom:15px;">QUESS</h2>'
    header_logo = ""


st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    :root {
      --orange:#ff9900; --quess:#004b87; --ink:#0f172a; --bg:#f3f5f9;
      --line:#e2e8f0; --muted:#64748b;
    }
    html, body, [class*="css"] { font-family:'Inter','Segoe UI',Arial,sans-serif; }
    .stApp { background:var(--bg); color:var(--ink); }
    [data-testid="stHeader"], [data-testid="stSidebar"] { display:none; }
    .block-container { max-width:1180px; padding:1.4rem 1.5rem 3rem !important; }

    /* ---------- header ---------- */
    .qc-header { display:flex; align-items:center; gap:16px; background:#fff; border:1px solid var(--line);
      border-radius:14px; padding:12px 22px; margin-bottom:18px; box-shadow:0 1px 3px rgba(15,23,42,.05); }
    .qc-logo { max-height:28px; }
    .qc-title { font-size:21px; font-weight:800; color:var(--ink); letter-spacing:-.3px; }
    .qc-tab { font-size:13px; font-weight:600; color:var(--quess); background:#e6f0fa; padding:5px 14px; border-radius:999px; }
    .qc-user { margin-left:auto; display:flex; align-items:center; gap:10px; }
    .qc-user b { display:block; font-size:13px; color:var(--ink); line-height:1.2; }
    .qc-user small { display:block; font-size:11px; color:var(--muted); }

    /* ---------- avatars ---------- */
    .qc-av { width:40px; height:40px; border-radius:50%; display:flex; align-items:center; justify-content:center;
      color:#fff; font-weight:700; font-size:14px; flex:none; letter-spacing:.3px; }
    .qc-av.q { background:linear-gradient(135deg,#004b87,#0a6cc4); font-size:21px; font-weight:800;
      box-shadow:0 0 0 3px #dbeafe; }
    .qc-av.anon { background:#cbd5e1; }
    .qc-av.sm { width:34px; height:34px; font-size:13px; }
    .qc-av.sm.q { font-size:18px; }

    /* ---------- inputs / buttons ---------- */
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div {
      background:#fff; border:1px solid var(--line); border-radius:10px; min-height:42px; }
    div.stButton > button, div.stFormSubmitButton > button {
      background:var(--orange); color:#fff; border:none; border-radius:999px; font-weight:700; font-size:14px; min-height:42px; }
    div.stButton > button:hover, div.stFormSubmitButton > button:hover { background:#e38800; color:#fff; }
    .st-key-logout_btn button { background:#fff !important; color:#334155 !important; border:1px solid var(--line) !important; }
    .st-key-logout_btn button:hover { background:#f8fafc !important; }
    [data-testid="stForm"] { background:#fff; border:1px solid var(--line); border-radius:14px; padding:26px;
      box-shadow:0 1px 3px rgba(15,23,42,.05); }

    /* ---------- post card ---------- */
    [class*="st-key-post_"] { background:#fff; border:1px solid var(--line) !important; border-radius:14px !important;
      box-shadow:0 1px 3px rgba(15,23,42,.05); padding:6px 8px 8px; margin-bottom:6px; }
    .qc-ph { display:flex; justify-content:space-between; align-items:flex-start; gap:12px; padding:10px 8px 4px; }
    .qc-author { display:flex; align-items:center; gap:12px; min-width:0; }
    .qc-line { display:flex; flex-wrap:wrap; align-items:center; gap:7px; font-size:14px; font-weight:700; color:var(--ink); }
    .qc-time { font-weight:500; color:var(--muted); font-size:13px; }
    .qc-dot { color:#cbd5e1; }
    .qc-sub { margin-top:3px; }
    .qc-chip { display:inline-block; font-size:11px; font-weight:600; color:#475569; background:#f1f5f9;
      padding:2px 9px; border-radius:999px; }
    .qc-admin { font-size:10px; font-weight:800; letter-spacing:.6px; color:#fff; background:var(--quess);
      padding:2px 7px; border-radius:5px; }
    .qc-real { font-size:11px; font-weight:600; color:#9a3412; background:#ffedd5; padding:2px 8px; border-radius:6px; }
    .qc-pill { font-size:12px; font-weight:700; padding:4px 12px; border-radius:999px; white-space:nowrap; }
    .s-open { background:#fff4e0; color:#b45309; }
    .s-prog { background:#e0edff; color:#1d4ed8; }
    .s-done { background:#dcfce7; color:#15803d; }
    .s-closed { background:#e5e7eb; color:#475569; }
    .qc-body { padding:8px 8px 14px; font-size:14.5px; line-height:1.65; color:#1e293b; word-wrap:break-word; }

    [class*="st-key-post_"] div.stButton > button { background:#fff; color:#1e293b; border:1px solid var(--line);
      min-height:34px; font-size:13px; font-weight:600; padding:0 14px; }
    [class*="st-key-post_"] div.stButton > button:hover { background:#f1f5f9; color:#0f172a; }
    [class*="st-key-up_"] button { color:var(--quess) !important; }
    [class*="st-key-del_"] button { color:#b91c1c !important; border-color:#fecaca !important; }
    [class*="st-key-del_"] button:hover { background:#fef2f2 !important; }
    [class*="st-key-delyes_"] button { background:#dc2626 !important; color:#fff !important; border-color:#dc2626 !important; }
    [class*="st-key-delyes_"] button:hover { background:#b91c1c !important; }
    .qc-warn { font-size:13px; font-weight:600; color:#991b1b; background:#fef2f2; border:1px solid #fecaca;
      border-radius:10px; padding:9px 14px; margin:6px 0; }

    /* ---------- response ---------- */
    .qc-resp { background:#f4f8fd; border:1px solid #d6e4f5; border-left:4px solid var(--quess);
      border-radius:10px; padding:14px 16px; margin:8px 4px 4px; }
    .qc-resp-h { display:flex; align-items:center; gap:10px; margin-bottom:8px; }
    .qc-resp-name { font-weight:700; font-size:13.5px; color:var(--quess); }
    .qc-resp-sub { font-size:11.5px; color:var(--muted); }
    .qc-resp-t { font-size:14px; line-height:1.6; color:#1e293b; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------- database --------------------------------
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
        # older databases: add the column that remembers when a response was written
        cols = {r["name"] for r in conn.execute("PRAGMA table_info(complaints)")}
        if "reply_date" not in cols:
            conn.execute("ALTER TABLE complaints ADD COLUMN reply_date TIMESTAMP")


@st.cache_data(show_spinner=False)
def load_roster():
    if not ROSTER_PATH.exists():
        return None, f"Roster workbook not found: {ROSTER_PATH.name}"
    try:
        df = pd.read_excel(ROSTER_PATH, sheet_name="Roster", dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        normalized = {str(c).strip().casefold(): c for c in df.columns}
        id_col = normalized.get("amz id") or normalized.get("amz_id") or normalized.get("amzid")
        name_col = normalized.get("employee name") or normalized.get("employee_name") or normalized.get("name")
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


# ----------------------------- helpers ---------------------------------
def safe(value):
    return html.escape("" if value is None else str(value))


def is_management_user():
    return str(st.session_state.get("amz_id", "")).strip().lower() in MANAGEMENT_IDS


def utc_now():
    return datetime.now(timezone.utc).replace(tzinfo=None)


def format_time_ago(date_str):
    """Timestamps are stored by SQLite in UTC, so compare against UTC."""
    try:
        dt = datetime.strptime(str(date_str)[:19], "%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""
    secs = max(0, int((utc_now() - dt).total_seconds()))
    if secs < 60:
        return "Just now"
    if secs < 3600:
        return f"{secs // 60}m ago"
    if secs < 86400:
        return f"{secs // 3600}h ago"
    return f"{secs // 86400}d ago"


def initials(name):
    parts = [p for p in str(name).replace("@", " ").split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def avatar_html(kind, name="", key="", small=False):
    """kind: 'admin' (blue Q), 'anon' (neutral silhouette) or 'user' (initials on a colour picked from the ID)."""
    sm = " sm" if small else ""
    if kind == "admin":
        return f'<div class="qc-av q{sm}">Q</div>'
    if kind == "anon":
        return (f'<div class="qc-av anon{sm}"><svg width="20" height="20" viewBox="0 0 24 24" fill="#fff">'
                '<circle cx="12" cy="8.5" r="4"/><path d="M4 21c0-4.2 3.6-6.5 8-6.5s8 2.3 8 6.5z"/></svg></div>')
    idx = int(hashlib.md5(str(key).lower().encode()).hexdigest()[:8], 16) % len(AVATAR_COLORS)
    return f'<div class="qc-av{sm}" style="background:{AVATAR_COLORS[idx]}">{safe(initials(name))}</div>'


def card(key):
    try:
        return st.container(border=True, key=key)
    except TypeError:  # very old Streamlit without keyed containers
        return st.container(border=True)


# ----------------------------- actions ---------------------------------
def do_upvote(pid):
    with db_connect() as conn:
        conn.execute("UPDATE complaints SET upvotes = upvotes + 1 WHERE id = ?", (int(pid),))


def open_reply(pid):
    if is_management_user():
        st.session_state.reply_open = pid
        st.session_state.confirm_del = None


def ask_delete(pid):
    if is_management_user():
        st.session_state.confirm_del = pid
        st.session_state.reply_open = None


def cancel_delete():
    st.session_state.confirm_del = None


def do_delete(pid):
    if not is_management_user():      # only admins can ever delete
        return
    with db_connect() as conn:
        conn.execute("DELETE FROM complaints WHERE id = ?", (int(pid),))
    st.session_state.confirm_del = None


def save_reply(pid, text, status):
    if not is_management_user():      # only admins can ever reply
        return
    text = (text or "").strip()
    with db_connect() as conn:
        conn.execute(
            "UPDATE complaints SET status = ?, reply = ?, reply_date = ? WHERE id = ?",
            (status, text, utc_now().strftime("%Y-%m-%d %H:%M:%S") if text else None, int(pid)),
        )
    st.session_state.reply_open = None


init_db()
roster_df, roster_error = load_roster()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.amz_id = ""
    st.session_state.emp_name = ""
    st.session_state.site = SITE
    st.session_state.show_form = False
st.session_state.setdefault("reply_open", None)
st.session_state.setdefault("confirm_del", None)


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
                <div style="font-size: 11px; font-weight: 700; color: #004b87; margin-bottom: 5px; letter-spacing: 1px; text-transform: uppercase;">Your voice matters</div>
                <div style="font-size: 26px; font-weight: 800; margin-bottom: 10px; color: #0f172a;">Welcome to QuessConnect</div>
                <div style="color: #64748b; font-size: 14px; margin-bottom: 25px;">Share feedback, raise a concern, or recognise something that is working well.</div>
                <h3 style="margin-bottom: 5px; font-size: 16px; color: #0f172a;">Associate login</h3>
                <p style="color: #64748b; font-size: 13px; margin-bottom: 10px;">Sign in with your roster AMZ ID. Site is set to DXB5.</p>
                </div>""",
                unsafe_allow_html=True,
            )

            amz_input = st.text_input(
                "AMZ ID", placeholder="Enter your AMZ ID", label_visibility="collapsed", autocomplete="username"
            )
            login_btn = st.form_submit_button("Continue", type="primary", use_container_width=True)

        if login_btn:
            amz_clean = str(amz_input).strip().lower()
            if not amz_clean:
                st.error("Please enter your AMZ ID.")
            elif amz_clean in MANAGEMENT_IDS:
                st.session_state.logged_in = True
                st.session_state.amz_id = amz_clean
                st.session_state.emp_name = "QuessCorp Management" if amz_clean == "quesscorp" else "Javmuhak"
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

    # ----- header -----
    me_av = avatar_html("admin" if management else "user", st.session_state.emp_name, st.session_state.amz_id, small=True)
    me_role = "Admin · Quess" if management else f"{SITE} associate"
    st.markdown(
        f"<div class='qc-header'>{header_logo}<span class='qc-title'>QuessConnect</span>"
        f"<span class='qc-tab'>Home</span>"
        f"<div class='qc-user'>{me_av}<div><b>{safe(st.session_state.emp_name)}</b><small>{safe(me_role)}</small></div></div></div>",
        unsafe_allow_html=True,
    )

    # ----- action bar -----
    c1, c2, c3, c4 = st.columns([1.5, 3.6, 1.2, 1], vertical_alignment="center")
    with c1:
        st.selectbox("Site", ["DXB5 QuessConnect"], label_visibility="collapsed")
    with c2:
        search_query = st.text_input("Search", placeholder="Search comments by keyword, author, category...",
                                     label_visibility="collapsed")
    with c3:
        if st.button("+ New Comment", type="primary", use_container_width=True, key="new_comment_btn"):
            st.session_state.show_form = not st.session_state.show_form
    with c4:
        if st.button("Log out", use_container_width=True, key="logout_btn"):
            for key in ["logged_in", "amz_id", "emp_name", "site", "show_form"]:
                st.session_state[key] = False if key in ("logged_in", "show_form") else ""
            st.session_state.reply_open = None
            st.session_state.confirm_del = None
            st.rerun()

    if st.session_state.show_form:
        with st.form("new_comment_form", clear_on_submit=True):
            st.markdown("### Create New Comment")
            category = st.selectbox("Category", CATEGORIES)
            message = st.text_area("Your comment", placeholder="Type your message here...", height=120)

            post_anonymous = st.toggle("Post as Anonymous (Hide identity on feed)", value=True)

            st.write("")
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
                        """
                        INSERT INTO complaints
                        (amz_id, employee_name, site, is_anonymous, category, message, status, reply, upvotes)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
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

    # ----- feed -----
    with db_connect() as conn:
        all_posts = pd.read_sql_query("SELECT * FROM complaints ORDER BY id DESC", conn)

    st.write("")
    if all_posts.empty:
        st.info("No comments yet. Be the first to share something.")
    else:
        feed_df = all_posts.copy()
        if search_query.strip():
            q = search_query.strip()
            anon_mask = feed_df["is_anonymous"].fillna(0).astype(int) == 1
            name_hit = feed_df["employee_name"].astype(str).str.contains(q, case=False, na=False, regex=False)
            mask = (
                feed_df["message"].astype(str).str.contains(q, case=False, na=False, regex=False)
                | feed_df["category"].astype(str).str.contains(q, case=False, na=False, regex=False)
                | (name_hit & (~anon_mask | management))   # anonymous authors are only searchable by admins
            )
            feed_df = feed_df[mask]
            if feed_df.empty:
                st.info("No comments match your search.")

        for _, post in feed_df.iterrows():
            pid = int(post["id"])
            anonymous = int(post["is_anonymous"] or 0) == 1
            status = str(post["status"] or "Open")
            is_post_mgmt = str(post["amz_id"]).lower() in MANAGEMENT_IDS
            reply = str(post["reply"] or "").strip()
            upvotes = int(post["upvotes"] or 0)

            # ----- who wrote it + their avatar -----
            if anonymous:
                av = avatar_html("anon")
                name_html = "Anonymous"
                if management:   # admins can see who is behind an anonymous post
                    name_html += (f" <span class='qc-real'>Real: {safe(post['employee_name'])} "
                                  f"@{safe(post['amz_id'])}</span>")
            elif is_post_mgmt:
                av = avatar_html("admin")
                name_html = f"<span style='color:#004b87'>{safe(post['employee_name'])}</span> <span class='qc-admin'>QUESS ADMIN</span>"
            else:
                av = avatar_html("user", post["employee_name"], post["amz_id"])
                name_html = safe(post["employee_name"])

            message_html = safe(post["message"]).replace("\r", "").replace("\n", "<br>")
            header = (
                f"<div class='qc-ph'><div class='qc-author'>{av}<div>"
                f"<div class='qc-line'>{name_html}<span class='qc-dot'>·</span>"
                f"<span class='qc-time'>{format_time_ago(post['date'])}</span></div>"
                f"<div class='qc-sub'><span class='qc-chip'>{safe(post['category'])}</span></div></div></div>"
                f"<div class='qc-pill {STATUS_CLASS.get(status, 's-open')}'>{safe(status)}</div></div>"
                f"<div class='qc-body'>{message_html}</div>"
            )

            with card(f"post_{pid}"):
                st.markdown(header, unsafe_allow_html=True)

                # ----- actions inside the card: upvote for everyone, reply + delete for admins only -----
                if management:
                    a = st.columns([1.3, 1, 1, 7])
                else:
                    a = st.columns([1.3, 10])
                a[0].button(f"▲ Upvote · {upvotes}", key=f"up_{pid}", on_click=do_upvote, args=(pid,),
                            use_container_width=True)
                if management:
                    a[1].button("Edit reply" if reply else "Reply", key=f"reply_{pid}", on_click=open_reply,
                                args=(pid,), use_container_width=True)
                    a[2].button("Delete", key=f"del_{pid}", on_click=ask_delete, args=(pid,),
                                use_container_width=True)

                # ----- delete confirmation (admins only) -----
                if management and st.session_state.confirm_del == pid:
                    st.markdown("<div class='qc-warn'>Delete this post permanently? This cannot be undone.</div>",
                                unsafe_allow_html=True)
                    d = st.columns([1.3, 1, 8])
                    d[0].button("Yes, delete", key=f"delyes_{pid}", on_click=do_delete, args=(pid,),
                                use_container_width=True)
                    d[1].button("Cancel", key=f"delno_{pid}", on_click=cancel_delete, use_container_width=True)

                # ----- reply form (admins only) -----
                if management and st.session_state.reply_open == pid:
                    with st.form(f"rf_{pid}"):
                        txt = st.text_area("Response", value=reply, height=100, label_visibility="collapsed",
                                           placeholder="Write a response the employee and team will see...")
                        s1, s2, s3 = st.columns([2, 1, 1])
                        new_status = s1.selectbox("Status", STATUSES, label_visibility="collapsed",
                                                  index=STATUSES.index(status) if status in STATUSES else 0,
                                                  format_func=lambda s: f"Status: {s}")
                        send = s2.form_submit_button("Send", type="primary", use_container_width=True)
                        cancel = s3.form_submit_button("Cancel", use_container_width=True)
                    if send:
                        save_reply(pid, txt, new_status)
                        st.rerun()
                    if cancel:
                        st.session_state.reply_open = None
                        st.rerun()

                # ----- response everyone can read -----
                if reply:
                    when = format_time_ago(post.get("reply_date")) if "reply_date" in post.index else ""
                    when_html = f" &nbsp;·&nbsp; {when}" if when else ""
                    st.markdown(
                        f"<div class='qc-resp'><div class='qc-resp-h'>{avatar_html('admin', small=True)}"
                        f"<div><div class='qc-resp-name'>QuessCorp Management{when_html}</div>"
                        f"<div class='qc-resp-sub'>HR Operations</div></div></div>"
                        f"<div class='qc-resp-t'>{safe(reply).replace(chr(13), '').replace(chr(10), '<br>')}</div></div>",
                        unsafe_allow_html=True,
                    )
