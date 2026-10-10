import sqlite3
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Workforce A to Z - MyVoice", page_icon="🗣️", layout="wide"
)

# Custom CSS for Amazon A to Z MyVoice Look & Feel
st.markdown(
    """
    <style>
        .main { background-color: #f8f9fa; }
        .stButton>button {
            background-color: #ff9900;
            color: white;
            font-weight: bold;
            border-radius: 4px;
            border: none;
        }
        .stButton>button:hover {
            background-color: #e88b00;
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
            border-left: 4px solid #ff9900;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# Database Setup
def init_db():
  conn = sqlite3.connect("myvoice_portal.db")
  cursor = conn.cursor()
  cursor.execute(
      """
        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            psoft_no TEXT,
            amz_id TEXT,
            employee_name TEXT,
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


# Load Employee Roster
@st.cache_data
def load_roster():
  try:
    df = pd.read_excel("DMD-DXB5-07102026.xlsx", sheet_name="Roster")
    df.columns = [str(c).strip() for c in df.columns]
    return df
  except Exception as e:
    return None


roster_df = load_roster()

# Top Header Navigation Bar (A to Z Style)
st.markdown(
    """
    <div style="display: flex; align-items: center; border-bottom: 1px solid #e0e0e0; padding-bottom: 10px; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #232f3e; font-size: 22px;">workforce <span style="color: #ff9900;">a to z</span></h2>
        <span style="margin-left: 30px; font-weight: bold; font-size: 18px; color: #0073bb; border-bottom: 3px solid #0073bb; padding-bottom: 8px;">MyVoice</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# App Mode
app_mode = st.sidebar.selectbox(
    "Navigation Mode", ["MyVoice Feed", "QuessCorp Admin Panel"]
)

# -------------------------------------------------------------
# 1. MYVOICE FEED (ASSOCIATE PORTAL)
# -------------------------------------------------------------
if app_mode == "MyVoice Feed":
  col_filter, col_search, col_btn = st.columns([1.5, 3, 1])

  with col_filter:
    site_filter = st.selectbox(
        "Site", ["DXB5 VOA", "DXB8 VOA", "DWC3 VOA"], label_visibility="collapsed"
    )

  with col_search:
    search_query = st.text_input(
        "Search",
        placeholder="Search comments by keyword, author, etc.",
        label_visibility="collapsed",
    )

  with col_btn:
    open_new_comment = st.button("+ New Comment", use_container_width=True)

  if open_new_comment:
    st.session_state["show_form"] = True

  if st.session_state.get("show_form", False):
    with st.form("new_comment_form"):
      st.subheader("💬 Create New MyVoice Comment")
      col1, col2 = st.columns(2)
      with col1:
        psoft_input = st.text_input(
            "Enter Psoft No:", placeholder="e.g. 204228003"
        )
      with col2:
        amz_input = st.text_input("Enter AMZ ID:", placeholder="e.g. brooabra")

      post_anonymous = st.checkbox(
          "Post as Anonymous (Aapka naam public feed par nahi dikhega)",
          value=True,
      )
      category = st.selectbox(
          "Category",
          [
              "Payroll / Salary",
              "Transport / Attendance",
              "Workplace Management",
              "Medical / Insurance",
              "Appreciation / Feedback",
          ],
      )
      message = st.text_area(
          "Your Message", placeholder="Apni baat yahan tafseel se likhein..."
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
        if not psoft_input or not amz_input or not message:
          st.error("⚠️ Kripya sabhi fields mukammal bharein!")
        elif roster_df is None:
          st.error("⚠️ Roster data load nahi ho saka.")
        else:
          psoft_clean = str(psoft_input).strip()
          amz_clean = str(amz_input).strip().lower()

          matched = roster_df[
              (roster_df["Psoft No"].astype(str).str.strip() == psoft_clean)
              & (
                  roster_df["AMZ ID"]
                  .astype(str)
                  .str.strip()
                  .str.lower()
                  == amz_clean
              )
          ]

          if not matched.empty:
            emp_name = matched.iloc[0]["Employee Name"]
            conn = sqlite3.connect("myvoice_portal.db")
            cursor = conn.cursor()
            cursor.execute(
                """
                            INSERT INTO complaints (psoft_no, amz_id, employee_name, is_anonymous, category, message, status, reply, upvotes)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                (
                    psoft_clean,
                    amz_clean,
                    emp_name,
                    1 if post_anonymous else 0,
                    category,
                    message,
                    "Open",
                    (
                        "Thank you for taking the time to share this. Your"
                        " feedback has been received — we're reviewing it with"
                        " the relevant team."
                    ),
                    0,
                ),
            )
            conn.commit()
            conn.close()
            st.success("🎉 Post safalpurvak publish ho gayi hai!")
            st.session_state["show_form"] = False
            st.rerun()
          else:
            st.error(
                "❌ Galat Psoft No ya AMZ ID! Roster ke mutabiq sahi details"
                " dalein."
            )

  st.markdown("---")

  conn = sqlite3.connect("myvoice_portal.db")
  query_str = "SELECT * FROM complaints ORDER BY id DESC"
  feed_df = pd.read_sql(query_str, conn)
  conn.close()

  if feed_df.empty:
    st.info(
        "Abhi koi comment ya post nahi hai. Upar '+ New Comment' par click"
        " karke pehli post dalein!"
    )
  else:
    if search_query:
      feed_df = feed_df[
          feed_df["message"].str.contains(search_query, case=False, na=False)
          | feed_df["category"].str.contains(
              search_query, case=False, na=False
          )
      ]

    for index, row in feed_df.iterrows():
      if row["is_anonymous"] == 1:
        author_display = "Anonymous"
      else:
        author_display = f"{row['employee_name']} (@{row['amz_id']})"

      status_badge = row["status"]

      st.markdown(
          f"""
            <div class="card">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                    <div>
                        <span style="font-weight: bold; font-size: 15px; color: #232f3e;">👤 {author_display}</span>
                        <span style="color: #666; font-size: 13px; margin-left: 10px;">• {row['date']}</span>
                    </div>
                    <div>
                        <span style="background-color: #e9ecef; color: #495057; padding: 3px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">📌 {status_badge}</span>
                    </div>
                </div>
                <div style="font-size: 12px; color: #0073bb; font-weight: bold; margin-bottom: 8px;">Category: {row['category']}</div>
                <div style="font-size: 14px; color: #333; line-height: 1.5; margin-bottom: 15px;">
                    {row['message']}
                </div>
        """,
          unsafe_allow_html=True,
      )

      col_up, col_dummy = st.columns([1, 6])
      with col_up:
        if st.button(f"👍 {row['upvotes']}", key=f"up_{row['id']}"):
          conn = sqlite3.connect("myvoice_portal.db")
          cursor = conn.cursor()
          cursor.execute(
              "UPDATE complaints SET upvotes = upvotes + 1 WHERE id = ?",
              (row["id"],),
          )
          conn.commit()
          conn.close()
          st.rerun()

      if row["reply"]:
        st.markdown(
            f"""
                <div class="response-box">
                    <div style="font-weight: bold; font-size: 13px; color: #232f3e; margin-bottom: 4px;">
                        🏢 QuessCorp HR Partner & Management Team
                    </div>
                    <div style="font-size: 13px; color: #444;">
                        {row['reply']}
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

      st.markdown("</div>", unsafe_allow_html=True)


# -------------------------------------------------------------
# 2. QUESSCORP ADMIN PANEL
# -------------------------------------------------------------
elif app_mode == "QuessCorp Admin Panel":
  st.title("🔐 QuessCorp Management Portal")
  st.markdown(
      "Yahan agency admin login karke saari confidential employee identities"
      " dekh sakti hai aur official responses post kar sakti hai."
  )

  admin_pass = st.text_input(
      "Enter Admin Password", type="password", placeholder="Password dalein"
  )

  if admin_pass == "quess123":
    st.success("🔓 Access Granted!")

    conn = sqlite3.connect("myvoice_portal.db")
    admin_df = pd.read_sql("SELECT * FROM complaints ORDER BY id DESC", conn)
    conn.close()

    if admin_df.empty:
      st.info("Abhi koi complaint ya feedback darj nahi hua hai.")
    else:
      st.subheader("📊 All Submitted Complaints (Confidential Backend)")
      st.dataframe(
          admin_df[[
              "id",
              "date",
              "psoft_no",
              "amz_id",
              "employee_name",
              "is_anonymous",
              "category",
              "status",
              "upvotes",
          ]],
          use_container_width=True,
      )

      st.markdown("---")
      st.subheader("✍️ Respond / Update Status")

      selected_id = st.selectbox(
          "Select Post ID to Manage", admin_df["id"].tolist()
      )

      if selected_id:
        row_data = admin_df[admin_df["id"] == selected_id].iloc[0]
        st.markdown(
            f"**Real Employee:** {row_data['employee_name']} (Psoft:"
            f" {row_data['psoft_no']} | AMZ ID: {row_data['amz_id']})"
        )
        st.markdown(f"**Message:** {row_data['message']}")

        with st.form("admin_reply"):
          new_stat = st.selectbox(
              "Update Status",
              ["Open", "In Progress", "Resolved", "Closed"],
              index=(
                  ["Open", "In Progress", "Resolved", "Closed"].index(
                      row_data["status"]
                  )
                  if row_data["status"]
                  in ["Open", "In Progress", "Resolved", "Closed"]
                  else 0
              ),
          )
          new_resp = st.text_area(
              "Official QuessCorp Response", value=row_data["reply"]
          )

          save_btn = st.form_submit_button("Publish Response & Update Status")

          if save_btn:
            conn = sqlite3.connect("myvoice_portal.db")
            cursor = conn.cursor()
            cursor.execute(
                """
                            UPDATE complaints 
                            SET status = ?, reply = ? 
                            WHERE id = ?
                        """,
                (new_stat, new_resp, selected_id),
            )
            conn.commit()
            conn.close()
            st.success("✅ Response update ho gaya hai!")
            st.rerun()

  elif admin_pass != "":
    st.error("❌ Galat password! (Default password: `quess123`)")
