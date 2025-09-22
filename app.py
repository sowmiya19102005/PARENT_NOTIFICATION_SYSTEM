import streamlit as st
import pandas as pd
import pywhatkit
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from datetime import datetime
import pyttsx3  # For voice notification

# ---------- CONFIG ----------
# Note: For security, move these to Streamlit secrets (.streamlit/secrets.toml)
# Example: SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
SENDER_EMAIL = "abc@gmail.com"
SENDER_PASSWORD = "app password"
EMAIL_SUBJECT = "Student Notification"

# ---------- FUNCTIONS ----------
def send_email(receiver_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())
        server.quit()
        return "Sent"
    except Exception as e:
        return f"Failed: {e}"

def send_whatsapp(phone, message, wait=20):
    try:
        if not str(phone).startswith('+'):
            phone = f"+{phone}"
        pywhatkit.sendwhatmsg_instantly(phone, message, wait_time=wait, tab_close=True)
        return "Sent"
    except Exception as e:
        return f"Failed: {e}"

def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, dtype=str)
    elif file.name.endswith((".xls", ".xlsx")):
        return pd.read_excel(file, dtype=str)
    else:
        st.error("Unsupported file type.")
        return None

def download_merged_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="MergedData")
    output.seek(0)
    return output

def speak(text):
    engine = pyttsx3.init()
    engine.setProperty('rate', 150)
    engine.setProperty('volume', 1.0)
    engine.say(text)
    engine.runAndWait()

# ---------- STREAMLIT UI ----------
st.markdown("""
<style>
.reportview-container, .main { background-color: #121212; color: #e0e0e0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
.title { font-size: 2.8rem; font-weight: 700; color: #bb86fc; margin-bottom: 0.2rem; }
.subtitle { font-size: 1.2rem; color: #a1a1a1; margin-top: 0; margin-bottom: 1.5rem; }
.section-header { color: #bb86fc; margin-top: 2rem; margin-bottom: 0.6rem; font-weight: 600; border-bottom: 2px solid #3700b3; padding-bottom: 6px; }
.footer { font-size: 0.85rem; color: #888888; margin-top: 3rem; text-align: center; }
div.stButton > button:first-child { background-color: #bb86fc; color: #121212; font-weight: 600; border-radius: 8px; padding: 10px 24px; transition: background-color 0.3s ease; box-shadow: 0 4px 6px rgba(187, 134, 252, 0.6); }
div.stButton > button:first-child:hover { background-color: #9a4dff; color: #fff; box-shadow: 0 6px 8px rgba(154, 77, 255, 0.8); }
div[role="radiogroup"] > label { font-size: 1rem; color: #e0e0e0; padding: 6px 12px; border-radius: 6px; border: 1.5px solid transparent; transition: all 0.3s ease; cursor: pointer; margin-right: 10px; user-select: none; background-color: #1e1e1e; }
div[role="radiogroup"] > label:hover { background-color: #3700b3; border-color: #bb86fc; color: #bb86fc; }
div[role="radiogroup"] > label[data-baseweb="radio"] > input:checked + span { background-color: #bb86fc !important; border-color: #bb86fc !important; color: #121212 !important; }
.stDataFrame div[data-testid="stDataFrameContainer"] { max-width: 100% !important; }
div[data-testid="stDataFrameContainer"]::-webkit-scrollbar { height: 8px; }
div[data-testid="stDataFrameContainer"]::-webkit-scrollbar-thumb { background-color: #bb86fc; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# Header with styled title and new emoji
st.markdown('<h1 class="title">📚 Parent Notification System</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Send Attendance or Marks notifications via Email & WhatsApp</p>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Settings")
    notify_type = st.radio("Select Notification Type:", ("Attendance", "Marks"))
    st.markdown("---")
    st.subheader("Upload Files")
    if notify_type == "Attendance":
        student_file = st.file_uploader("Student Attendance Details (with 'status')", type=["csv", "xls", "xlsx"])
    else:
        student_file = st.file_uploader("Student Marks Details (with 'marks')", type=["csv", "xls", "xlsx"])
    parent_file = st.file_uploader("Parent Details", type=["csv", "xls", "xlsx"])
    delay = st.slider("Delay between WhatsApp messages (seconds)", 2, 10, 4)
    send_button = st.button("Send Notifications")

if student_file and parent_file:
    students_df = read_file(student_file)
    parents_df = read_file(parent_file)

    if students_df is not None and parents_df is not None:
        students_df = students_df.drop_duplicates(subset=['student_id'], keep='first')
        parents_df = parents_df.drop_duplicates(subset=['student_id'], keep='first')

        # Validate required columns
        student_required_cols = ["student_id", "name"]
        student_required_cols.append("status" if notify_type == "Attendance" else "marks")
        parent_required_cols = ["student_id", "parent_name", "email", "phone"]

        missing_student_cols = [col for col in student_required_cols if col not in students_df.columns]
        missing_parent_cols = [col for col in parent_required_cols if col not in parents_df.columns]

        if missing_student_cols:
            st.error(f"Student file missing columns: {', '.join(missing_student_cols)}")
        elif missing_parent_cols:
            st.error(f"Parent file missing columns: {', '.join(missing_parent_cols)}")
        else:
            st.markdown('<div class="section-header">Student Data</div>', unsafe_allow_html=True)
            st.dataframe(students_df, width='stretch')

            st.markdown('<div class="section-header">Parent Data</div>', unsafe_allow_html=True)
            st.dataframe(parents_df, width='stretch')

            merged_df = pd.merge(students_df, parents_df, on="student_id")
            merged_df['phone'] = merged_df['phone'].apply(lambda x: f"+{x}" if not str(x).startswith('+') else str(x))
            merged_df = merged_df.drop_duplicates(subset=['student_id', 'parent_name'])

            st.markdown('<div class="section-header">Merged Data</div>', unsafe_allow_html=True)
            st.dataframe(merged_df, width='stretch')

            excel_bytes = download_merged_excel(merged_df)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.download_button(
                label="Download Merged Data",
                data=excel_bytes,
                file_name=f"merged_data_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if send_button:
                st.info("Sending messages... please wait ⏳")
                progress_bar = st.progress(0)
                total = len(merged_df)
                sent_count = 0
                failed_notifications = []

                for i, row in merged_df.iterrows():
                    student = row['name']
                    parent = row['parent_name']
                    email = row['email']
                    phone = row['phone']

                    if notify_type == "Attendance":
                        status = row.get('status')
                        if status and str(status).lower() == "absent":
                            message = f"Hello {parent}, This is to inform you that your child {student} was absent today. Kindly update us with the reason."
                        else:
                            st.info(f"Skipped {student}, status: {status}")
                            progress_bar.progress((i + 1) / total)
                            continue
                    else:
                        marks = row.get('marks')
                        if marks and str(marks).strip() != "":
                            message =  f"Dear {parent}, this is to inform you that your child {student} has scored {marks} marks. Kindly encourage them to maintain consistency."
                        else:
                            st.info(f"Skipped {student}, marks not available")
                            progress_bar.progress((i + 1) / total)
                            continue

                    email_status = send_email(email, EMAIL_SUBJECT, message)
                    whatsapp_status = send_whatsapp(phone, message)
                    if email_status.startswith("Failed") or whatsapp_status.startswith("Failed"):
                        failed_notifications.append({"student": student, "email_status": email_status, "whatsapp_status": whatsapp_status})

                    st.success(f"Sent to {parent} | Email: {email_status} | WhatsApp: {whatsapp_status}")
                    sent_count += 1
                    time.sleep(delay)
                    progress_bar.progress((i + 1) / total)

                if failed_notifications:
                    st.error(f"Failed notifications: {len(failed_notifications)}")
                    st.dataframe(pd.DataFrame(failed_notifications), width='stretch')

                st.balloons()
                st.success(f"All messages processed! Total sent: {sent_count}")

                # Voice notification
                speak("All messages have been sent successfully")
else:
    st.warning("Please upload both Student and Parent files to proceed.")