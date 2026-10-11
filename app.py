import html
import sqlite3
import base64
import hashlib
import json
import time
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
MANAGEMENT_IDS = {"quesscorp48", "javmuhak"}   # admins: can reply to and delete posts
LEGACY_ADMIN_IDS = {"quesscorp"}   # old admin ID: not a login any more, only keeps its old posts shown as admin posts
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
# Streamlit Cloud wipes the local database file whenever the app restarts or is
# redeployed, so every comment is also stored as ONE json file in a private GitHub
# repository. On start-up the file is read back into the local database.
#
# Configure these values in Streamlit Cloud -> App settings -> Secrets:
# [github]
# token = "YOUR_FINE_GRAINED_TOKEN_WITH_CONTENTS_READ_WRITE"
# repo = "YOUR_GITHUB_USERNAME/YOUR_PRIVATE_DATA_REPOSITORY"   # NOT the repo this app is deployed from
# branch = "main"
# path = "data/quessconnect-backup.json"
#
# The JSON snapshot holds comments, replies, statuses and upvotes.
# Never hard-code the token in this file.
BACKUP_COLUMNS = [
    "id", "amz_id", "employee_name", "site", "is_anonymous", "category",
    "message", "status", "reply", "upvotes", "date", "reply_date",
]


@st.cache_resource
def sync_state():
    """One shared dict per running app process (survives reruns and new sessions)."""
    return {"ready": False, "retry_at": 0.0, "last_saved": None}


def github_settings():
    try:
        cfg = st.secrets.get("github", {})
        token = str(cfg.get("token", "")).strip()
        repo = str(cfg.get("repo", "")).strip().strip("/")
        branch = str(cfg.get("branch", "main")).strip() or "main"
        path = str(cfg.get("path", "data/quessconnect-backup.json")).strip().lstrip("/")
        if token and repo and path:
            return {"token": token, "repo": repo, "branch": branch, "path": path}
    except Exception:
        pass
    return None


def github_api_request(url, token, method="GET", payload=None, accept="application/vnd.github+json", as_text=False):
    body = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        method=method,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": accept,
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "QuessConnect-Backup",
            **({"Content-Type": "application/json"} if body is not None else {}),
        },
    )
    with urllib.request.urlopen(req, timeout=20) as response:
        raw = response.read().decode("utf-8")
        if as_text:
            return raw
        return json.loads(raw) if raw else {}


def github_backup_url(cfg):
    from urllib.parse import quote
    encoded_path = "/".join(quote(part, safe="") for part in cfg["path"].split("/"))
    return (
        f"https://api.github.com/repos/{cfg['repo']}/contents/{encoded_path}"
        f"?ref={quote(cfg['branch'], safe='')}"
    )


def github_error_text(exc):
    """Short human readable reason, so the warning tells you what to fix."""
    if isinstance(exc, urllib.error.HTTPError):
        hint = {
            401: "the token is invalid or expired",
            403: "the token has no permission - it needs Contents: Read and write on that repository",
            404: "repository / branch not found, or the token cannot access that repository",
            409: "the file was changed at the same moment, please retry",
            422: "the file was changed at the same moment, please retry",
        }.get(exc.code, "")
        detail = ""
        try:
            detail = str(json.loads(exc.read().decode("utf-8")).get("message", ""))[:120]
        except Exception:
            pass
        text = f"HTTP {exc.code}"
        if hint:
            text += f" - {hint}"
        if detail:
            text += f" ({detail})"
        return text
    return str(exc)


def github_fetch_backup(cfg):
    """Return (sha, backup_dict_or_None). A missing file gives (None, None); any other problem raises."""
    try:
        meta = github_api_request(github_backup_url(cfg), cfg["token"])
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None, None
        raise
    sha = meta.get("sha")
    encoded = "".join(str(meta.get("content", "")).split())
    if encoded:
        raw = base64.b64decode(encoded).decode("utf-8")
    else:
        # files over 1 MB come back without inline content - ask for the raw bytes instead
        raw = github_api_request(
            github_backup_url(cfg), cfg["token"], accept="application/vnd.github.raw+json", as_text=True
        )
    raw = (raw or "").strip()
    return sha, (json.loads(raw) if raw else None)


def ensure_github_sync():
    """Bring the local database and the GitHub snapshot together once per app start.

    - GitHub rows are loaded into the local database (so nothing vanishes after a restart).
    - Rows that only exist locally are kept and pushed to GitHub.
    - Until this has succeeded, the GitHub file is never overwritten (see save_github_backup).
    """
    cfg = github_settings()
    if not cfg:
        return False
    state = sync_state()
    if state["ready"]:
        return True
    if time.time() < state["retry_at"]:
        return False
    try:
        sha, backup = github_fetch_backup(cfg)
        remote_rows = backup.get("complaints") if isinstance(backup, dict) else None
        if not isinstance(remote_rows, list):
            remote_rows = []
        remote_ids = set()
        with db_connect() as conn:
            columns = {row["name"] for row in conn.execute("PRAGMA table_info(complaints)")}
            allowed = [column for column in BACKUP_COLUMNS if column in columns]
            names = ", ".join(allowed)
            placeholders = ", ".join("?" for _ in allowed)
            for item in remote_rows:
                if not isinstance(item, dict) or item.get("id") is None:
                    continue
                try:
                    conn.execute(
                        f"INSERT OR REPLACE INTO complaints ({names}) VALUES ({placeholders})",
                        [item.get(column) for column in allowed],
                    )
                    remote_ids.add(int(item["id"]))
                except Exception:
                    continue   # skip one damaged row, keep the rest
            local_ids = {int(row[0]) for row in conn.execute("SELECT id FROM complaints")}
        state["ready"] = True
        state["retry_at"] = 0.0
        st.session_state["github_backup_notice"] = ""
        if sha is None or not local_ids <= remote_ids:
            save_github_backup()   # first run, or local rows GitHub does not have yet
        return True
    except Exception as exc:
        state["retry_at"] = time.time() + 60
        st.session_state["github_backup_notice"] = (
            f"GitHub backup could not be read: {github_error_text(exc)}. "
            "New comments are saved locally only until this is fixed."
        )
        return False


def save_github_backup():
    cfg = github_settings()
    if not cfg:
        st.session_state["github_backup_notice"] = (
            "GitHub backup is not configured. Comments are currently saved only in this app's local database."
        )
        return False
    state = sync_state()
    if not state["ready"] and not ensure_github_sync():
        # We could not confirm what GitHub holds, so never overwrite it with a partial copy.
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
        encoded_content = base64.b64encode(
            json.dumps(payload_data, ensure_ascii=False, indent=2).encode("utf-8")
        ).decode("ascii")
        put_url = github_backup_url(cfg).split("?ref=", 1)[0]
        last_error = None
        for attempt in range(3):
            try:
                sha, _ = github_fetch_backup(cfg)
                put_payload = {
                    "message": "Update QuessConnect comments backup",
                    "content": encoded_content,
                    "branch": cfg["branch"],
                }
                if sha:
                    put_payload["sha"] = sha
                github_api_request(put_url, cfg["token"], method="PUT", payload=put_payload)
                last_error = None
                break
            except urllib.error.HTTPError as exc:
                last_error = exc
                if exc.code in (409, 422) and attempt < 2:
                    time.sleep(0.6 * (attempt + 1))   # file changed at the same moment - fetch the new sha and retry
                    continue
                break
        if last_error is not None:
            raise last_error
        state["last_saved"] = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")
        st.session_state["github_backup_notice"] = ""
        st.session_state["github_backup_saved"] = True
        return True
    except Exception as exc:
        st.session_state["github_backup_notice"] = (
            f"GitHub backup failed: {github_error_text(exc)}. The change is saved locally but not yet on GitHub."
        )
        st.session_state["github_backup_saved"] = False
        return False


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
ensure_github_sync()     # reload comments from GitHub first (cheap once it has succeeded)
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
    elif is_management_user():
        _saved = sync_state().get("last_saved")
        st.caption("✅ GitHub backup is connected" + (f" · last saved {_saved} UTC" if _saved else ""))


# ----------------------------- Login ---------------------------------
# Optional: paste a base64 copy of the Dubai skyline photo here. login_bg.jpg beside this file is used first.
LOGIN_BG_B64 = ""

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
/* the "AMZ ID" label is now the input's own label, so it can never slide under the box */
.st-key-login_card [data-testid="stWidgetLabel"] { min-height:0 !important; margin-bottom:clamp(6px,1.1vh,10px) !important; }
.st-key-login_card [data-testid="stWidgetLabel"] p { font-size:clamp(13px,2vh,16px) !important; font-weight:700 !important;
    color:#0b1b3f !important; line-height:1.4 !important; margin:0 !important; }
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
                f"<div class='lg-sub'>Sign in with your roster AMZ ID to continue.<br>Site is set to {SITE}.</div>",
                unsafe_allow_html=True,
            )
            with st.form("login_form", clear_on_submit=False):
                amz_input = st.text_input(
                    "AMZ ID", placeholder="Enter your AMZ ID", autocomplete="username"
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
            st.session_state.emp_name = "QuessCorp Management" if amz_clean == "quesscorp48" else "Javmuhak"
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
            is_post_mgmt = str(post["amz_id"]).lower() in (MANAGEMENT_IDS | LEGACY_ADMIN_IDS)
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
