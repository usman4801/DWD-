import sqlite3
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="QuessCorp - Voice of Associate (VoA)",
    page_icon="🗣️",
    layout="wide",
)

# Database Setup
def init_db():
  conn = sqlite3.connect("voa_portal.db")
  cursor = conn.cursor()
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS complaints (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            psoft_no TEXT,
            amz_id TEXT,
            employee_name TEXT,
            category TEXT,
            message TEXT,
            status TEXT,
            reply TEXT,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
  conn.commit()
  conn.close()


init_db()

# Load Employee Roster from Excel
@st.cache_data
def load_roster():
  try:
    df = pd.read_excel("DMD-DXB5-07102026.xlsx", sheet_name="Roster")
    # Clean up column names just in case
    df.columns = [str(c).strip() for c in df.columns]
    return df
  except Exception as e:
    return None


roster_df = load_roster()

# Sidebar Navigation
st.sidebar.title("🗣️ QuessCorp VoA Portal")
st.sidebar.markdown(
    "**Voice of Associate** - Aapki awaz, hamari zimmedari! [cite:"
    "DMD-DXB5-07102026.xlsx]"
)
app_mode = st.sidebar.selectbox(
    "Choose Portal Mode", ["Associate Portal", "QuessCorp Admin Panel"]
)

# -------------------------------------------------------------
# 1. ASSOCIATE PORTAL
# -------------------------------------------------------------
if app_mode == "Associate Portal":
  st.title("🛡️ QuessCorp Voice of Associate (VoA)")
  st.markdown(
      "Yahan aap apni complaint ya feedback bilkul **Anonymous** tarike se"
      " darj kar sakte hain. Public feed par aapki pehchan chupayi jayegi!"
  )

  with st.form("complaint_form"):
    st.subheader("📝 Register New Complaint")

    col1, col2 = st.columns(2)
    with col1:
      psoft_input = st.text_input(
          "Enter your Psoft No (e.g., 204228003):",
          placeholder="Apna Psoft No dalein",
      )
    with col2:
      amz_input = st.text_input(
          "Enter your AMZ ID (e.g., brooabra):",
          placeholder="Apni AMZ ID dalein",
      )

    category = st.selectbox(
        "Complaint Category",
        [
            "Payroll / Salary",
            "Transport / Attendance",
            "Harrassment / Workplace Issue",
            "Medical / Insurance",
            "General Feedback",
            "Other",
        ],
    )

    message = st.text_area(
        "Your Message / Complaint Details",
        placeholder="Apni poori baat yahan likhein...",
    )

    submit_btn = st.form_submit_button("Submit Complaint Anonymously")

    if submit_btn:
      if not psoft_input or not amz_input or not message:
        st.error(
            "⚠️ Kripya Psoft No, AMZ ID aur Message sabhi fields bharein!"
        )
      elif roster_df is None:
        st.error("⚠️ System error: Roster file nahi mil saki!")
      else:
        # Verify employee in roster
        psoft_clean = str(psoft_input).strip()
        amz_clean = str(amz_input).strip().lower()

        matched_emp = roster_df[
            (roster_df["Psoft No"].astype(str).str.strip() == psoft_clean)
            & (roster_df["AMZ ID"].astype(str).str.strip().str.lower() == amz_clean)
        ]

        if not matched_emp.empty:
          emp_name = matched_emp.iloc[0]["Employee Name"]

          # Save to DB
          conn = sqlite3.connect("voa_portal.db")
          cursor = conn.cursor()
          cursor.execute(
              """
                        INSERT INTO complaints (psoft_no, amz_id, employee_name, category, message, status, reply)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
              (
                  psoft_clean,
                  amz_clean,
                  emp_name,
                  category,
                  message,
                  "Open",
                  "Awaiting response...",
              ),
          )
          conn.commit()
          conn.close()

          st.success(
              "🎉 Aapki complaint safalpurvak register ho gayi hai! Public"
              " board par yeh Anonymous dikhegi."
          )
        else:
          st.error(
              "❌ Galat Psoft No ya AMZ ID! Kripya apne sahi credentials"
              " dalein jo roster mein maujood hain."
          )

  st.markdown("---")
  st.subheader("📢 Public Complaint Feed (Anonymous View)")
  st.markdown(
      "Yahan sabhi associates ki complaints anonymous roop mein dikhti hain."
  )

  # Fetch complaints for public view (hiding identity)
  conn = sqlite3.connect("voa_portal.db")
  public_df = pd.read_sql(
      "SELECT id, category, message, status, reply, date FROM complaints ORDER"
      " BY id DESC",
      conn,
  )
  conn.close()

  if public_df.empty:
    st.info("Abhi tak koi complaint darj nahi ki gayi hai.")
  else:
    for index, row in public_df.iterrows():
      with st.container():
        st.markdown(f"### 🏷️ Category: {row['category']}")
        st.markdown(f"> **Complaint:** {row['message']}")
        st.markdown(
            f"👤 **Author:** `Anonymous Associate` | 📅 **Date:**"
            f" {row['date']} | 📌 **Status:** **{row['status']}**"
        )
        if row["reply"] and row["reply"] != "Awaiting response...":
          st.success(f"💬 **QuessCorp Reply:** {row['reply']}")
        else:
          st.info("💬 **QuessCorp Reply:** Review mein hai...")
        st.markdown("---")


# -------------------------------------------------------------
# 2. QUESSCORP ADMIN PANEL
# -------------------------------------------------------------
elif app_mode == "QuessCorp Admin Panel":
  st.title("🔐 QuessCorp Management Dashboard")
  st.markdown("Yahan agency admin login karke saari asal identities dekh sakti hai aur reply kar sakti hai.")

  admin_pass = st.text_input("Enter Admin Password", type="password", placeholder="Password dalein")

  # Simple hardcoded admin check (Change password as needed)
  if admin_pass == "quess123":
    st.success("🔓 Admin Access Granted!")

    conn = sqlite3.connect("voa_portal.db")
    admin_df = pd.read_sql("SELECT * FROM complaints ORDER BY id DESC", conn)
    conn.close()

    if admin_df.empty:
      st.info("Koi active complaint maujood nahi hai.")
    else:
      st.subheader("📋 All Employee Complaints (Confidential View)")
      
      # Display raw data table for quick overview
      st.dataframe(admin_df[['id', 'date', 'psoft_no', 'amz_id', 'employee_name', 'category', 'status']], use_container_width=True)

      st.markdown("---")
      st.subheader("✍️ Update Complaint / Send Reply")

      selected_id = st.selectbox("Select Complaint ID to Manage", admin_df['id'].tolist())
      
      if selected_id:
        current_row = admin_df[admin_df['id'] == selected_id].iloc[0]
        
        st.markdown(f"**Employee Name:** {current_row['employee_name']} (Psoft: {current_row['psoft_no']} | AMZ: {current_row['amz_id']})")
        st.markdown(f"**Category:** {current_row['category']}")
        st.markdown(f"**Message:** {current_row['message']}")

        with st.form("admin_reply_form"):
          new_status = st.selectbox("Update Status", ["Open", "In Progress", "Resolved", "Closed"], index=["Open", "In Progress", "Resolved", "Closed"].index(current_row['status']) if current_row['status'] in ["Open", "In Progress", "Resolved", "Closed"] else 0)
          new_reply = st.text_area("Write QuessCorp Reply", value=current_row['reply'] if current_row['reply'] != "Awaiting response..." else "")
          
          update_btn = st.form_submit_button("Update Complaint")

          if update_btn:
            conn = sqlite3.connect("voa_portal.db")
            cursor = conn.cursor()
            cursor.execute("""
                        UPDATE complaints 
                        SET status = ?,
