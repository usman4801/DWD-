import sqlite3
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="QuessCorp - Voice of Associate", page_icon="🗣️", layout="wide"
)

# Custom CSS Styling
st.markdown(
    """
    <style>
        .main { background-color: #f8f9fa; }
        .stButton>button {
            background-color: #004b87;
            color: white;
            font-weight: bold;
            border-radius: 4px;
            border: none;
        }
        .stButton>button:hover {
            background-color: #003366;
            color: white;
        }
        .card {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            border: 1px solid #e0e0e0;
            margin-bottom: 15px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .response-box {
            background-color: #f1f3f5;
            padding: 15px;
            border-radius: 6px;
            margin-top: 10px;
            border-left: 4px solid #004b87;
        }
        .management-badge {
            background-color: #d97706;
            color: white;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: bold;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# Database Setup
def init_db():
  conn = sqlite3.connect("quess_voa.db")
  cursor = conn.cursor()
  cursor.execute(
      """
        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            amz_id TEXT,
            employee_name TEXT,
            site TEXT,
            is_anonymous INTEGER,
            category TEXT,
            message TEXT,
            status TEXT,
            reply TEXT,
            upvotes INTEGER DEFAULT 0,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """
  )
  conn.commit()
  conn.close()


init_db()


# Load Employee Roster from Excel
@st.cache_data
def load_roster():
  try:
    df = pd.read_excel("DMD-DXB5-07102026.xlsx", sheet_name="Roster")
    df.columns = [str(c).strip() for c in df.columns]
    return df
  except Exception as e:
    return None


roster_df = load_roster()

# Session State Initialization for Login
if "logged_in" not in st.session_state:
  st.session_state["logged_in"] = False
  st.session_state["amz_id"] = ""
  st.session_state["emp_name"] = ""
  st.session_state["site"] = "DXB5"

# -------------------------------------------------------------
# LOGIN SCREEN (AMZ ID Only & Site fixed to DXB5)
# -------------------------------------------------------------
if not st.session_state["logged_in"]:
  st.markdown(
      """
        <div style="text-align: center; padding: 30px 0;">
            <h2 style="color: #004b87; margin-bottom: 5px;">QUESS</h2>
            <p style="color: #666; font-size: 14px; font-weight: bold; letter-spacing: 1px;">WINNING TOGETHER</p>
            <h3 style="color: #333; margin-top: 10px;">Voice of Associate (VoA) Portal</h3>
        </div>
    """,
      unsafe_allow_html=True,
  )

  col1, col2, col3 = st.columns([1, 1.5, 1])
  with col2:
    with st.form("login_form"):
      st.subheader("Associate Login")
      st.info("Site is fixed to: **DXB5**")
      amz_input = st.text_input(
          "Enter your AMZ ID", placeholder="e.g. brooabra"
      )

      login_btn = st.form_submit_button("Login to Portal", use_container_width=True)

      if login_btn:
        amz_clean = str(amz_input).strip().lower()

        # Management accounts check
        if amz_clean in ["quesscorp", "javmuhak"]:
          st.session_state["logged_in"] = True
          st.session_state["amz_id"] = amz_clean
          st.session_state["emp_name"] = (
              "QuessCorp Management"
              if amz_clean == "quesscorp"
              else "Javmuhak"
          )
          st.session_state["site"] = "DXB5"
          st.success("Management Login Successful!")
          st.rerun()

        elif roster_df is not None:
          matched = roster_df[
              roster_df["AMZ ID"]
              .astype(str)
              .str.strip()
              .str.lower()
              == amz_clean
          ]

          if not matched.empty:
            st.session_state["logged_in"] = True
            st.session_state["amz_id"] = amz_clean
            st.session_state["emp_name"] = matched.iloc[0]["Employee Name"]
            st.session_state["site"] = "DXB5"
            st.success("Login Successful!")
            st.rerun()
          else:
            st.error("Invalid AMZ ID! Please check your roster ID.")
        else:
          st.error("Roster file could not be loaded.")

# -------------------------------------------------------------
# MAIN PORTAL FEED
# -------------------------------------------------------------
else:
  is_management = st.session_state["amz_id"].lower() in ["quesscorp", "javmuhak"]

  # Header with Quess Branding
  st.markdown(
      f"""
        <div style="display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #004b87; padding-bottom: 10px; margin-bottom: 20px;">
            <div style="display: flex; align-items: center; gap: 15px;">
                <div>
                    <h3 style="margin: 0; color: #004b87; font-weight: 800; letter-spacing: 0.5px;">QUESS</h3>
                    <span style="font-size: 10px; color: #333; font-weight: bold;">WINNING TOGETHER</span>
                </div>
                <div style="border-left: 1px solid #ccc; padding-left: 15px;">
                    <h4 style="margin: 0; color: #1e293b;">Voice of Associate (VoA)</h4>
                    <span style="font-size: 12px; color: #666;">Welcome, <b>{st.session_state['emp_name']}</b> (@{st.session_state['amz_id']})</span>
                </div>
            </div>
            <div>
                <span style="background-color: #004b87; color: white; padding: 5px 10px; border-radius: 4px; font-size: 12px; font-weight: bold;">Site: DXB5</span>
            </div>
        </div>
    """,
      unsafe_allow_html=True,
  )

  # Logout button row
  col_logout1, col_logout2 = st.columns([6, 1])
  with col_logout2:
    if st.button("Logout", use_container_width=True):
      st.session_state["logged_in"] = False
      st.rerun()

  # Action Bar: Filter, Search, and New Comment
  col_filter, col_search, col_btn = st.columns([1.5, 3, 1])

  with col_filter:
    category_filter = st.selectbox(
        "Filter Category",
        [
            "All Categories",
            "Payroll / Salary",
            "Transport / Attendance",
            "Workplace Management",
            "Medical / Insurance",
            "Appreciation / Feedback",
        ],
        label_visibility="collapsed",
    )

  with col_search:
    search_query = st.text_input(
        "Search",
        placeholder="Search complaints by keyword...",
        label_visibility="collapsed",
    )

  with col_btn:
    open_new_comment = st.button("+ New Comment", use_container_width=True)

  if open_new_comment:
    st.session_state["show_form"] = True

  # New Comment Form
  if st.session_state.get("show_form", False):
    with st.form("new_comment_form"):
      st.subheader("Register New Complaint / Feedback")

      post_anonymous = st.checkbox(
          "Post as Anonymous (Your name will be hidden on the public feed)",
          value=True,
      )
      category = st.selectbox(
          "Select Category",
          [
              "Payroll / Salary",
              "Transport / Attendance",
              "Workplace Management",
              "Medical / Insurance",
              "Appreciation / Feedback",
          ],
      )
      message = st.text_area(
          "Complaint Details",
          placeholder="Describe your concern in detail here...",
      )

      col_sub1, col_sub2 = st.columns(2)
      with col_sub1:
        submit_comment = st.form_submit_button("Submit Post")
      with col_sub2:
        cancel_comment = st.form_submit_button("Cancel")

      if cancel_comment:
        st.session_state["show_form"] = False
        st.rerun()

      if submit_comment:
        if not message:
          st.error("Please enter your message details.")
        else:
          conn = sqlite3.connect("quess_voa.db")
          cursor = conn.cursor()
          cursor.execute(
              """
                        INSERT INTO complaints (amz_id, employee_name, site, is_anonymous, category, message, status, reply, upvotes)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
              (
                  st.session_state["amz_id"],
                  st.session_state["emp_name"],
                  "DXB5",
                  1 if post_anonymous else 0,
                  category,
                  message,
                  "Open",
                  "Thank you for your feedback. Management is reviewing your case.",
                  0,
              ),
          )
          conn.commit()
          conn.close()
          st.success("Your complaint has been successfully published.")
          st.session_state["show_form"] = False
          st.rerun()

  st.markdown("---")

  # Feed Retrieval
  conn = sqlite3.connect("quess_voa.db")
  query_str = "SELECT * FROM complaints ORDER BY id DESC"
  feed_df = pd.read_sql(query_str, conn)
  conn.close()

  if feed_df.empty:
    st.info("No complaints or feedback found.")
  else:
    if category_filter != "All Categories":
      feed_df = feed_df[feed_df["category"] == category_filter]

    if search_query:
      feed_df = feed_df[
          feed_df["message"].str.contains(search_query, case=False, na=False)
          | feed_df["category"].str.contains(
              search_query, case=False, na=False
          )
      ]

    for index, row in feed_df.iterrows():
      # Determine author presentation based on role
      if is_management:
        if row["is_anonymous"] == 1:
          author_display = f"{row['employee_name']} (@{row['amz_id']}) <span class='management-badge'>[Anonymous Report]</span>"
        else:
          author_display = f"{row['employee_name']} (@{row['amz_id']})"
      else:
        if row["is_anonymous"] == 1:
          author_display = "Anonymous"
        else:
          author_display = f"{row['employee_name']} (@{row['amz_id']})"

      status_badge = row["status"]

      # Render Post Card
      st.markdown(
          f"""
            <div class="card">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                    <div>
                        <span style="font-weight: bold; font-size: 15px; color: #1e293b;">👤 {author_display}</span>
                        <span style="color: #666; font-size: 13px; margin-left: 10px;">• {row['date']}</span>
                    </div>
                    <div>
                        <span style="background-color: #e2e8f0; color: #1e293b; padding: 3px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">Status: {status_badge}</span>
                    </div>
                </div>
                <div style="font-size: 12px; color: #004b87; font-weight: bold; margin-bottom: 8px;">Category: {row['category']} | Site: DXB5</div>
                <div style="font-size: 14px; color: #333; line-height: 1.5; margin-bottom: 15px;">
                    {row['message']}
                </div>
        """,
          unsafe_allow_html=True,
      )

      # Upvote Button
      col_up, col_dummy = st.columns([1, 6])
      with col_up:
        if st.button(f"👍 {row['upvotes']}", key=f"up_{row['id']}"):
          conn = sqlite3.connect("quess_voa.db")
          cursor = conn.cursor()
          cursor.execute(
              "UPDATE complaints SET upvotes = upvotes + 1 WHERE id = ?",
              (row["id"],),
          )
          conn.commit()
          conn.close()
          st.rerun()

      # Official Response Box display
      if row["reply"]:
        st.markdown(
            f"""
                <div class="response-box">
                    <div style="font-weight: bold; font-size: 13px; color: #004b87; margin-bottom: 4px;">
                        QuessCorp Management Response
                    </div>
                    <div style="font-size: 13px; color: #444;">
                        {row['reply']}
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

      # INLINE REPLY OPTION EXCLUSIVE TO `quesscorp` and `javmuhak`
      if is_management:
        with st.expander(f"⚙️ Reply / Manage Post #{row['id']}"):
          with st.form(f"mgmt_reply_{row['id']}"):
            new_status = st.selectbox(
                "Update Status",
                ["Open", "In Progress", "Resolved", "Closed"],
                index=(
                    ["Open", "In Progress", "Resolved", "Closed"].index(
                        row["status"]
                    )
                    if row["status"]
                    in ["Open", "In Progress", "Resolved", "Closed"]
                    else 0
                ),
            )
            new_response = st.text_area(
                "Write Management Response",
                value=(
                    row["reply"]
                    if row["reply"]
                    != "Thank you for your feedback. Management is reviewing your case."
                    else ""
                ),
            )
            submit_reply = st.form_submit_button("Publish Response")

            if submit_reply:
              conn = sqlite3.connect("quess_voa.db")
              cursor = conn.cursor()
              cursor.execute(
                  """
                                UPDATE complaints 
                                SET status = ?, reply = ? 
                                WHERE id = ?
                            """,
                  (new_status, new_response, row["id"]),
              )
              conn.commit()
              conn.close()
              st.success("Response published successfully!")
              st.rerun()

      st.markdown("</div>", unsafe_allow_html=True)
