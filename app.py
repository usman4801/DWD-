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
  cursor.execute(
      """
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
        psoft_clean = str(psoft_input).strip()
        amz_clean = str(amz_input).strip().lower()

        matched_emp = roster_df[
            (roster_df["Psoft No"].astype(str).str.strip() == psoft_clean)
            & (
                roster_df["AMZ ID"]
                .astype(str)
                .str.strip()
                .str.lower()
                == amz_clean
            )
        ]

        if not matched_emp.empty:
          emp_name = matched_emp.iloc[0]["Employee Name"]

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

  conn = sqlite3.connect("voa_portal.db")
  public_df = pd.read
