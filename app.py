import html
import sqlite3
import base64
import hashlib
import json
import urllib.error
import urllib.request
from pathlib import Path
from datetime import datetime, timezone

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------
# QuessCorp | QuessConnect Portal
# Run: streamlit run app.py
# Keep DMD-DXB5-07102026.xlsx, image_2c8c24.png and login_bg.jpg beside this file.
# Needs Streamlit 1.39 or newer (keyed containers are used for the card styling).
# ---------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
ROSTER_PATH = APP_DIR / "DMD-DXB5-07102026.xlsx"
DB_PATH = APP_DIR / "quessconnect.db"
LOGO_PATH = APP_DIR / "image_2c8c24.png"  # Quess Logo File
LOGIN_BG_PATH = APP_DIR / "login_bg.jpg"  # Dubai skyline photo used on the login page
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


# ----------------------- persistent GitHub backup -----------------------
# Configure these values in Streamlit Cloud -> App settings -> Secrets:
# [github]
# token = "YOUR_FINE_GRAINED_TOKEN_WITH_CONTENTS_READ_WRITE"
# repo = "YOUR_GITHUB_USERNAME/YOUR_PRIVATE_REPOSITORY"
# branch = "main"
# path = "data/quessconnect-backup.json"
#
# The private repository stores a JSON snapshot of comments, replies, statuses
# and upvotes. Never hard-code the token in this file.
def github_settings():
    try:
        cfg = st.secrets.get("github", {})
        token = str(cfg.get("token", "")).strip()
        repo = str(cfg.get("repo", "")).strip()
        branch = str(cfg.get("branch", "main")).strip() or "main"
        path = str(cfg.get("path", "data/quessconnect-backup.json")).strip().lstrip("/")
        if token and repo and path:
            return {"token": token, "repo": repo, "branch": branch, "path": path}
    except Exception:
        pass
    return None


def github_api_request(url, token, method="GET", payload=None):
    body = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        method=method,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "QuessConnect-Backup",
            **({"Content-Type": "application/json"} if body is not None else {}),
        },
    )
    with urllib.request.urlopen(req, timeout=15) as response:
        raw = response.read().decode("utf-8")
        return json.loads(raw) if raw else {}


def github_backup_url(cfg):
    from urllib.parse import quote
    encoded_path = "/".join(quote(part, safe="") for part in cfg["path"].split("/"))
    return (
        f"https://api.github.com/repos/{cfg['repo']}/contents/{encoded_path}"
        f"?ref={quote(cfg['branch'], safe='')}"
    )


def read_github_backup():
    cfg = github_settings()
    if not cfg:
        return None
    try:
        remote = github_api_request(github_backup_url(cfg), cfg["token"])
        encoded = "".join(str(remote.get("content", "")).split())
        if not encoded:
            return None
        return json.loads(base64.b64decode(encoded).decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None
        st.session_state["github_backup_notice"] = (
            f"GitHub backup could not be read (HTTP {exc.code}). Check the GitHub token, repository, branch and file permissions."
        )
    except Exception as exc:
        st.session_state["github_backup_notice"] = f"GitHub backup could not be read: {exc}"
    return None


def save_github_backup():
    cfg = github_settings()
    if not cfg:
        st.session_state["github_backup_notice"] = (
            "GitHub backup is not configured. Comments are currently saved only in this app's local database."
        )
        return False
    try:
        with db_connect() as conn:
            rows = conn.execute("SELECT * FROM complaints ORDER BY id").fetchall()
            columns = [item["name"] for item in conn.execute("PRAGMA table_info(complaints)").fetchall()]
        payload_data = {
            "format_version": 1,
            "updated_at_utc": datetime.now(timezone.utc).isoformat(),
            "complaints": [{column: row[column] for column in columns} for row in rows],
        }
        url = github_backup_url(cfg).split("?ref=", 1)[0]
        sha = None
        try:
            current = github_api_request(github_backup_url(cfg), cfg["token"])
            sha = current.get("sha")
        except urllib.error.HTTPError as exc:
            if exc.code != 404:
                raise

        encoded_content = base64.b64encode(
            json.dumps(payload_data, ensure_ascii=False, indent=2).encode("utf-8")
        ).decode("ascii")
        put_payload = {
            "message": "Update QuessConnect comments backup",
            "content": encoded_content,
            "branch": cfg["branch"],
        }
        if sha:
            put_payload["sha"] = sha
        github_api_request(url, cfg["token"], method="PUT", payload=put_payload)
        st.session_state["github_backup_notice"] = ""
        st.session_state["github_backup_saved"] = True
        return True
    except Exception as exc:
        st.session_state["github_backup_notice"] = (
            f"GitHub backup failed: {exc}. The change may still be saved locally."
        )
        st.session_state["github_backup_saved"] = False
        return False


def restore_github_backup_if_needed():
    """Restore the GitHub snapshot only when the local complaints table is empty."""
    try:
        with db_connect() as conn:
            existing = conn.execute("SELECT COUNT(*) FROM complaints").fetchone()[0]
        if existing:
            return
        backup = read_github_backup()
        if not isinstance(backup, dict) or not isinstance(backup.get("complaints"), list):
            return
        with db_connect() as conn:
            columns = {row["name"] for row in conn.execute("PRAGMA table_info(complaints)")}
            allowed = [
                "id", "amz_id", "employee_name", "site", "is_anonymous", "category",
                "message", "status", "reply", "upvotes", "date", "reply_date"
            ]
            allowed = [column for column in allowed if column in columns]
            for item in backup["complaints"]:
                if not isinstance(item, dict):
                    continue
                values = [item.get(column) for column in allowed]
                placeholders = ", ".join("?" for _ in allowed)
                names = ", ".join(allowed)
                conn.execute(
                    f"INSERT OR REPLACE INTO complaints ({names}) VALUES ({placeholders})",
                    values,
                )
        st.session_state["github_backup_notice"] = ""
    except Exception as exc:
        st.session_state["github_backup_notice"] = f"Could not restore comments from GitHub: {exc}"


def initialize_github_backup_if_missing():
    """Seed GitHub with existing local comments if no remote backup exists yet.

    This prevents older local comments from being lost when GitHub backup is
    configured for the first time. Never overwrites a remote file that exists.
    """
    cfg = github_settings()
    if not cfg:
        return
    try:
        with db_connect() as conn:
            local_count = conn.execute("SELECT COUNT(*) FROM complaints").fetchone()[0]
        if not local_count:
            return
        try:
            github_api_request(github_backup_url(cfg), cfg["token"])
            # A remote snapshot already exists: do not overwrite it on startup.
            return
        except urllib.error.HTTPError as exc:
            if exc.code != 404:
                st.session_state["github_backup_notice"] = (
                    f"Could not check the GitHub backup (HTTP {exc.code}); no data was overwritten."
                )
                return
        save_github_backup()
    except Exception as exc:
        st.session_state["github_backup_notice"] = f"Could not initialize GitHub backup: {exc}"


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
    save_github_backup()


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
    save_github_backup()
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
    save_github_backup()
    st.session_state.reply_open = None


init_db()
if not st.session_state.get("github_restore_checked", False):
    restore_github_backup_if_needed()
    initialize_github_backup_if_missing()
    st.session_state["github_restore_checked"] = True
roster_df, roster_error = load_roster()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.amz_id = ""
    st.session_state.emp_name = ""
    st.session_state.site = SITE
    st.session_state.show_form = False
st.session_state.setdefault("reply_open", None)
st.session_state.setdefault("confirm_del", None)

# Show backup configuration status inside the portal, not above the login page.
# This keeps the login screen clean and prevents unnecessary vertical scrolling.
if st.session_state.logged_in:
    if st.session_state.get("github_backup_notice"):
        st.warning(st.session_state["github_backup_notice"])
    elif not github_settings():
        st.warning(
            "GitHub backup is not configured. Comments currently rely on this app's local database; "
            "configure Streamlit Secrets to enable persistent GitHub backups."
        )


# ----------------------------- Login ---------------------------------
# Dubai skyline photo, embedded so the login page works even without login_bg.jpg
LOGIN_BG_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAoHCAkIBgoJCAkMCwoMDxoRDw4ODx8WGBMaJSEnJiQhJCMpLjsyKSw4LCMkM0Y0OD0/QkNCKDFITUhATTtBQj//2wBDAQsMDA8NDx4RER4/KiQqPz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz8/Pz//wAARCANmA4QDASIAAhEBAxEB/8QAHAABAQADAQEBAQAAAAAAAAAAAAECAwQFBgcI/8QAThAAAgIBAwIEAwUDCAcFBwQDAAECAxEEEiEFMQYTQVEiYXEUMoGRsSNCoQcVJDNScsHRFkNTYnOCkiU0Y6LhVGR0g7LC8CZEk/EXNTb/xAAaAQEBAAMBAQAAAAAAAAAAAAAAAQIDBAUG/8QANhEAAgIBBAAEAwgCAQQDAQAAAAECEQMEEiExEyJBUQVhcRQygZGhscHwI9HhJDNCUhVi8TT/2gAMAwEAAhEDEQA/APzYAHsHmgAEAKQAFBAACkKAQAAAAAAAAFBCggABQAAAUEKAAAAAAAAAUAAAhQAAAAAAAAAAAAAAAAAAAACkBQUAAAAAgKQAFAAAABQAAAUEKCAAAAAAAAFAAAAAABQAAAACAAAAAAAAAAAAApClAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABzgA1G0AAgAAAKQAAADIAAAABCgAAAAAFBQQoIAAAAAAUEABQAAAAUAAAFAAIAAAAAAAAAAAAAAAAAAAAUFBCggAAAABQCkKAAAAAAACkKAAACAAFAAAAAAAKQoAAAAAAAAAAAAAAAICgFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABzgENJtKQAAoIACggAKCFAIAAAAAAUgAKCFAAABAUgKUoIUEAAABSAAoAAAAAAAKCggBCgAAAAAAAAAAAAAAAAoBSFAAABAAAAUAAAAoAAABQAAAAAAAQAoAICgoIUAAAAAAAAAAAAAEBQAAACgAFAICgAgKACAoAICgAgKACAoAICgAgKACAoAICgAgKACAoAICgAgKADlABpNoAAAABCgAFIAAQoDAAAAAAAKQAAAFIACghQAAACggKCgAEAAAKCFAAAAAAABSApCgAAAAAAAAAAAFIUAAAAAAAAFKAAAQAAAFAAAAKAAUAAAAAAEAAAAAKAAAAUhQAAAAAAACgEIUAAAFAICgoICgAgKACAoAICgAgKACAoAICgAgKACAoAICgAgKACAoAICgA5AAaDcAQAFIAACkABQQAAAAAAAFBCgAAAAAFIAAAUEKAAAACkABQACApAUFAAAAAABQAQoAAABSAAAAAoBAUAAAAAAoAABQAAACgAAAAgBQAAAUAAAAAoBCgAgAAABQAQFBQQoAAAKAQFABCgAAAoIQFAAABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAcQANBuAAAAAAAAAAAAAAAAAAAAABSAAoAQAABQAAAUAAgAABQEAAAACgAEAAKCgAAAAAAAAAAAFIUpAAAACgAhQAAAACgAAAAAFAKAAAQAAAAoAAAKACgAAAAAoBAAAAAAACgoIUAAAAAAoAICgAgKAQgKACAoAICgAgKACAoAICgAgKACAoAICgAgKACAoAICgA4QAaTcAAAAQEAAAAKQpQAAAAAAAAAAAAUAAAAAAAFAKACAAAFQAAAAAKAAQAFKAAAAAAAAAACgAAAAFIUpAAAAAUAhQAAAACgAAAAAAAoBQAAAAQFAAABSghQAAAUAhQAQAAAAAoAAABQAACgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4AAaTaCFABAUhAAAACgAAAFAAAAAAAAKAAAAAAAAClIAAACkKAAAACgAAAAgKAAAAUAAAAAAFAAAAKUgAAAAABQAAAAACgAAAFAAKAAACAAoAAAABQUAAAAFAIAAAAAAAAUAFAAAABQAAAAAAUAhAUAEBQAQFABAUAEBQAQFABAUAEBQAQFABAUAEBQAQFABAUAHngpDSbgAAAAAAACAAAoAAAAAAAAABSFAAAAAAABQCkAAAKAAAUhQAAAQFAAAAKAAAAAUAhQAAAUpAAAAAAAUAAAFAAAAAAKACgAAAEBQAAAUAhQCgAFAAABAAAAACgAFAABQCAoBAAAACgAAAAAoKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADzyFIaDcAAAAAAAAAAAAAAAAAAAAAUAAAAAAAoIAAUApCgAAAFAAAAKCAAFAAAAKAAAAAAClAAAICgAAAAAAoAAAABQAQoBQAUAAAAgKAAAClBCgAAAoBAUAgAKAQFBQQFAABQAQoAAAKCEKAAACg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"

LOGIN_CSS = """
<style>
.block-container, [data-testid="stMainBlockContainer"] {
    max-width:100% !important; padding:0 !important;
}
.stApp {
    background:
        radial-gradient(circle at 96% 6%, rgba(147,180,255,.45), transparent 22%),
        radial-gradient(circle at 100% 60%, rgba(186,210,255,.35), transparent 28%),
        linear-gradient(160deg,#f7faff 0%,#eaf1fc 100%);
}
div[data-testid="stHorizontalBlock"] { gap:0 !important; }
[data-testid="stVerticalBlock"] { gap:0 !important; }
@media (min-width: 901px) {      /* whole login fits in one screen, no scrolling */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] { height:100vh !important; overflow:hidden !important; }
}

/* ---------- left: skyline panel ---------- */
.lg-panel {
    position:relative; box-sizing:border-box; height:100vh; overflow:hidden; color:#fff;
    display:flex; flex-direction:column; justify-content:space-between;
    padding:clamp(18px,4.6vh,48px) clamp(24px,4.2vw,72px) clamp(18px,4.6vh,44px);
    background-image: __BG__, linear-gradient(180deg,#012a5b 0%,#0b4390 62%,#e48a47 100%);
    background-size:cover; background-position:center bottom; background-repeat:no-repeat;
    border-radius:0 48px 48px 0;
}
.lg-brand { display:flex; align-items:center; gap:14px; }
.lg-q { width:clamp(40px,6.4vh,56px); height:clamp(40px,6.4vh,56px); border-radius:50%; background:#fff; color:#0b2a63;
    font-weight:800; font-size:clamp(24px,3.9vh,34px); display:flex; align-items:center; justify-content:center; flex:none;
    box-shadow:0 8px 22px rgba(0,0,0,.25); }
.lg-brand b { font-size:clamp(20px,3.5vh,30px); font-weight:700; letter-spacing:-.4px; }
.lg-tagline { margin-top:1.4vh; font-size:clamp(9px,1.45vh,12.5px); font-weight:500; letter-spacing:.32em; color:#d6e4ff; }
.lg-h { margin:0 0 1.6vh; font-size:clamp(32px,min(7.4vh,4.6vw),70px); font-weight:800; line-height:1.0;
    letter-spacing:-2px; color:#fff; }
.lg-h span { display:block; background:linear-gradient(90deg,#38c0fa,#2f7df6);
    -webkit-background-clip:text; background-clip:text; color:transparent; }
.lg-p { max-width:500px; margin:0; font-size:clamp(12px,2vh,17px); line-height:1.5; color:#eaf1ff; }
.lg-cards { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; max-width:620px; }
.lg-c { background:rgba(255,255,255,.09); border:1px solid rgba(255,255,255,.2); border-radius:16px;
    padding:clamp(10px,1.9vh,18px) clamp(10px,1.1vw,16px); backdrop-filter:blur(8px); -webkit-backdrop-filter:blur(8px); }
.lg-ic { width:clamp(30px,5vh,42px); height:clamp(30px,5vh,42px); border-radius:50%; background:#2f6fe8; display:flex;
    align-items:center; justify-content:center; margin-bottom:1.3vh; box-shadow:0 6px 16px rgba(20,60,160,.45); }
.lg-ic svg { width:52%; height:52%; }
.lg-c b { display:block; font-size:clamp(12px,2vh,16px); font-weight:700; margin-bottom:.5vh; }
.lg-c span { display:block; font-size:clamp(10px,1.6vh,13px); line-height:1.4; color:#d5e3ff; }

/* ---------- right: login card ---------- */
.lg-top { position:fixed; top:2.8vh; right:40px; z-index:5; display:flex; align-items:center; gap:16px;
    font-size:12px; font-weight:600; letter-spacing:.2em; color:#25344f; }
.lg-top i { width:1px; height:13px; background:#a9b7d0; display:block; }
.st-key-login_card {
    background:#fff !important; border:1px solid #e4ecf9 !important; border-radius:26px !important;
    padding:clamp(18px,3.6vh,38px) clamp(20px,2.6vw,40px) clamp(16px,3vh,30px) !important;
    box-shadow:0 26px 70px rgba(30,64,140,.15) !important;
    max-width:540px; width:calc(100% - 48px); margin:0 auto; box-sizing:border-box;
}
.lg-tag { font-size:clamp(11px,1.8vh,14px); font-weight:800; letter-spacing:.08em; text-transform:uppercase; color:#1d5fe0;
    margin-bottom:1vh; }
.lg-title { font-size:clamp(28px,5.6vh,46px); font-weight:800; color:#0b1b3f; letter-spacing:-1.4px;
    line-height:1.08; margin-bottom:1.6vh; }
.lg-sub { font-size:clamp(12px,1.95vh,15.5px); color:#5b6b8c; line-height:1.5; margin-bottom:clamp(12px,2.8vh,26px); }
.lg-label { font-size:clamp(13px,2vh,16px); font-weight:700; color:#0b1b3f; margin-bottom:1vh; }
.st-key-login_card [data-testid="stForm"] {
    background:transparent !important; border:0 !important; border-radius:0 !important;
    padding:0 !important; box-shadow:none !important;
}
.st-key-login_card div[data-baseweb="input"],
.st-key-login_card div[data-baseweb="base-input"],
.st-key-login_card div[data-baseweb="input"] > div {
    background:transparent !important; border:0 !important; box-shadow:none !important; min-height:0 !important;
}
.st-key-login_card input {
    min-height:clamp(42px,6.6vh,56px) !important; border-radius:14px !important; font-size:15px !important; color:#0b1b3f !important;
    border:1px solid #d3e0f4 !important; padding-left:64px !important;
    background-color:#f6f9fe !important;
    background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%236b7fa6' stroke-width='1.8' stroke-linecap='round'><circle cx='12' cy='8' r='4'/><path d='M4 21c0-4 3.6-6.5 8-6.5s8 2.5 8 6.5'/></svg>"),
                     linear-gradient(#d3e0f4,#d3e0f4) !important;
    background-repeat:no-repeat !important; background-position:20px center, 54px center !important;
    background-size:22px 22px, 1px 26px !important;
}
.st-key-login_card input:focus { border-color:#3b82f6 !important; box-shadow:0 0 0 3px rgba(59,130,246,.16) !important; }
.st-key-login_card input::placeholder { color:#8a9bb8 !important; }
.st-key-login_card [data-testid="stForm"] button {
    min-height:clamp(44px,6.8vh,58px) !important; border-radius:999px !important; border:0 !important; margin-top:1.8vh;
    background:linear-gradient(90deg,#0b4fd1,#3b8bff) !important; color:#fff !important;
    box-shadow:0 12px 26px rgba(37,99,235,.33) !important; transition:transform .15s ease, box-shadow .15s ease;
}
.st-key-login_card [data-testid="stForm"] button p { color:#fff !important; font-size:16px !important; font-weight:700 !important; }
.st-key-login_card [data-testid="stForm"] button:hover { transform:translateY(-1px); box-shadow:0 16px 30px rgba(37,99,235,.4) !important; }
.lg-help { display:flex; align-items:center; gap:14px; margin:clamp(10px,2.4vh,22px) 0 clamp(8px,1.6vh,12px); color:#5b6b8c; font-size:clamp(11px,1.7vh,13.5px); }
.lg-help:before, .lg-help:after { content:""; height:1px; background:#e3eaf6; flex:1; }
.lg-hr { display:flex; align-items:center; gap:14px; background:#f1f6fd; border-radius:16px; padding:clamp(10px,1.9vh,16px) 18px;
    font-size:clamp(13px,2vh,15.5px); font-weight:600; color:#0b1b3f; }
.lg-hr svg { flex:none; }
.lg-hr i { width:1px; height:22px; background:#d3e0f4; display:block; }
.lg-hr span { flex:1; }
.lg-foot3 { display:flex; justify-content:center; align-items:center; gap:clamp(10px,1.4vw,22px);
    max-width:540px; margin:clamp(10px,2.2vh,20px) auto 0; width:calc(100% - 48px); }
.lg-f { display:flex; align-items:center; gap:8px; font-size:clamp(10px,1.55vh,12.5px); line-height:1.3; color:#4a5a78; }
.lg-f svg { flex:none; width:clamp(20px,3.4vh,28px); height:clamp(20px,3.4vh,28px); }
.lg-fl { width:1px; height:26px; background:#cbd6ea; display:block; }

@media (max-width: 900px) {
    .lg-panel { height:auto; gap:22px; border-radius:0 0 32px 32px; padding-bottom:30px; }
    .lg-cards { grid-template-columns:1fr; }
    .lg-top { position:static; justify-content:center; padding:14px 0; }
    .st-key-login_card { margin:20px auto; width:calc(100% - 32px); }
    .lg-foot3 { flex-wrap:wrap; padding-bottom:20px; }
}
</style>
"""

ICON_CHAT = ("<svg viewBox='0 0 24 24' fill='#fff'><path d='M4 3h16a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H9l-5 4v-4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z'/>"
             "<circle cx='8' cy='10' r='1.3' fill='#2f6fe8'/><circle cx='12' cy='10' r='1.3' fill='#2f6fe8'/>"
             "<circle cx='16' cy='10' r='1.3' fill='#2f6fe8'/></svg>")
ICON_PEOPLE = ("<svg viewBox='0 0 24 24' fill='#fff'><circle cx='9' cy='8' r='3.3'/><path d='M2.5 20c0-3.8 2.8-6 6.5-6s6.5 2.2 6.5 6z'/>"
               "<circle cx='17' cy='9' r='2.5'/><path d='M16.5 14.2c3.2 0 5 1.8 5 4.8h-4.2c0-2-.4-3.6-1.5-4.6z'/></svg>")
ICON_STAR = "<svg viewBox='0 0 24 24' fill='#fff'><polygon points='12,2 15,9 22,9.3 16.5,14 18.3,21 12,17 5.7,21 7.5,14 2,9.3 9,9'/></svg>"
ICON_HEADSET = ("<svg width='26' height='26' viewBox='0 0 24 24' fill='none' stroke='#0b1b3f' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'>"
                "<path d='M4 14v-2a8 8 0 0 1 16 0v2'/><rect x='3' y='13' width='4' height='6' rx='1.5'/>"
                "<rect x='17' y='13' width='4' height='6' rx='1.5'/><path d='M19 19c0 1.5-1.5 2.5-4 2.5h-2'/></svg>")
ICON_CHEVRON = ("<svg width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='#0b1b3f' stroke-width='2' "
                "stroke-linecap='round' stroke-linejoin='round'><path d='M9 5l7 7-7 7'/></svg>")
ICON_SHIELD = ("<svg width='28' height='28' viewBox='0 0 24 24' fill='#1e3a8a'><path d='M12 2l8 3v6c0 5-3.4 9.3-8 11-4.6-1.7-8-6-8-11V5z'/>"
               "<path d='M8.5 12l2.5 2.5 4.5-5' fill='none' stroke='#fff' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'/></svg>")
ICON_LOCK = ("<svg width='28' height='28' viewBox='0 0 24 24' fill='#1e3a8a'><rect x='4' y='10' width='16' height='11' rx='2.5'/>"
             "<path d='M8 10V7a4 4 0 0 1 8 0v3' fill='none' stroke='#1e3a8a' stroke-width='2.2' stroke-linecap='round'/>"
             "<circle cx='12' cy='15.5' r='1.6' fill='#fff'/></svg>")
ICON_TEAM = ("<svg width='30' height='30' viewBox='0 0 24 24' fill='#1e3a8a'><circle cx='12' cy='7.5' r='3'/><path d='M6 19c0-3.4 2.6-5.3 6-5.3s6 1.9 6 5.3z'/>"
             "<circle cx='5' cy='9.5' r='2.2'/><path d='M0.8 18c0-2.6 1.7-4 4.2-4 .7 0 1.3.1 1.8.3-1.2 1-1.8 2.3-1.9 3.7z'/>"
             "<circle cx='19' cy='9.5' r='2.2'/><path d='M23.2 18c0-2.6-1.7-4-4.2-4-.7 0-1.3.1-1.8.3 1.2 1 1.8 2.3 1.9 3.7z'/></svg>")

if not st.session_state.logged_in:

    bg_b64 = get_base64_image(LOGIN_BG_PATH) or LOGIN_BG_B64
    bg_layer = f"url(data:image/jpeg;base64,{bg_b64})" if bg_b64 else "none"
    st.markdown(LOGIN_CSS.replace("__BG__", bg_layer), unsafe_allow_html=True)

    left, right = st.columns([1.5, 1], gap="small", vertical_alignment="center")
    with left:
        st.markdown(
            "<div class='lg-panel'>"
            "<div><div class='lg-brand'><div class='lg-q'>Q</div><b>QuessConnect</b></div>"
            "<div class='lg-tagline'>PEOPLE &middot; FEEDBACK &middot; BETTER TOMORROW</div></div>"
            "<div><div class='lg-h'>Your voice,<span>heard.</span></div>"
            f"<p class='lg-p'>A safe space for the {SITE} team to share feedback, raise concerns and recognise great work.</p></div>"
            "<div class='lg-cards'>"
            f"<div class='lg-c'><div class='lg-ic'>{ICON_CHAT}</div><b>Speak Freely</b><span>Share your thoughts openly or anonymously.</span></div>"
            f"<div class='lg-c'><div class='lg-ic'>{ICON_PEOPLE}</div><b>We Listen</b><span>Management reads and responds.</span></div>"
            f"<div class='lg-c'><div class='lg-ic'>{ICON_STAR}</div><b>Make a Difference</b><span>Upvote what matters to you.</span></div>"
            "</div></div>",
            unsafe_allow_html=True,
        )

    with right:
        st.markdown(
            f"<div class='lg-top'><span>{SITE}</span><i></i><span>QuessConnect</span></div>",
            unsafe_allow_html=True,
        )
        with st.container(key="login_card"):
            st.markdown(
                "<div class='lg-tag'>Associate login</div>"
                "<div class='lg-title'>Welcome back</div>"
                f"<div class='lg-sub'>Sign in with your roster AMZ ID to continue.<br>Site is set to {SITE}.</div>"
                "<div class='lg-label'>AMZ ID</div>",
                unsafe_allow_html=True,
            )
            with st.form("login_form", clear_on_submit=False):
                amz_input = st.text_input(
                    "AMZ ID", placeholder="Enter your AMZ ID",
                    label_visibility="collapsed", autocomplete="username"
                )
                login_btn = st.form_submit_button("Continue  →", type="primary", use_container_width=True)
            st.markdown(
                "<div class='lg-help'>Need help signing in?</div>"
                f"<div class='lg-hr'>{ICON_HEADSET}<i></i><span>Contact HR Operations</span>{ICON_CHEVRON}</div>",
                unsafe_allow_html=True,
            )
        st.markdown(
            "<div class='lg-foot3'>"
            f"<div class='lg-f'>{ICON_SHIELD}<span>Your feedback<br>is safe</span></div><i class='lg-fl'></i>"
            f"<div class='lg-f'>{ICON_LOCK}<span>Confidential<br>and secure</span></div><i class='lg-fl'></i>"
            f"<div class='lg-f'>{ICON_TEAM}<span>For a better<br>tomorrow</span></div>"
            "</div>",
            unsafe_allow_html=True,
        )

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

            post_anonymous = st.toggle("Post as Anonymous (Hide identity on feed)", value=False)

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
                save_github_backup()
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
