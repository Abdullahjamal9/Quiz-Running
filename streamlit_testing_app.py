import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import time
import os
import gspread
from google.oauth2.service_account import Credentials

# =====================
# Paths / Files (local Excel for reading only)
# =====================
BASE_DIR = os.path.dirname(__file__)   # absolute path (safe for Streamlit Cloud)
DB_FOLDER = os.path.join(BASE_DIR, "db")
QUESTIONS_FOLDER = os.path.join(DB_FOLDER, "Questions")
EMP_STD_FILE = os.path.join(DB_FOLDER, "Result 2.xlsx")
INFO_FILE = os.path.join(DB_FOLDER, "info.xlsx")

# =====================
# Google Sheets Setup (for saving results)
# =====================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
client = gspread.authorize(creds)
GSHEET_URL = st.secrets["connections"]["gsheets"]["spreadsheet"]

# =====================
# Cached Loaders
# =====================
@st.cache_data
def load_employees_and_standards():
    try:
        employees = pd.read_excel(EMP_STD_FILE, sheet_name="Emloyees Data")
    except Exception:
        employees = pd.DataFrame(columns=["ID","Name"])
    try:
        standards = pd.read_excel(EMP_STD_FILE, sheet_name="Standard")
        if len(standards.columns) < 2:
            while len(standards.columns) < 2:
                standards[standards.columns[-1] + "_dup" + str(len(standards.columns))] = ""
            standards.columns = ["Standard","ShortName"]
        else:
            standards.columns = ["Standard","ShortName"]
    except Exception:
        standards = pd.DataFrame(columns=["Standard","ShortName"])
    standards["Standard"] = standards["Standard"].astype(str).str.strip()
    standards["ShortName"] = standards["ShortName"].astype(str).str.strip()
    return employees, standards


@st.cache_data
def load_questions():
    expected = ["Qno","Question","A","B","C","D","Answer","Standard"]
    rename_map = {
        "NO.": "Qno",
        "Opt A": "A",
        "Opt B": "B",
        "Opt C": "C",
        "Opt D": "D"
    }
    all_q = []

    single_file = os.path.join(DB_FOLDER, "Questions.xlsx")
    if os.path.exists(single_file):
        try:
            q = pd.read_excel(single_file)
            q = q.rename(columns=rename_map)
            all_q.append(q)
        except Exception:
            pass

    if os.path.isdir(QUESTIONS_FOLDER):
        for fname in os.listdir(QUESTIONS_FOLDER):
            if fname.lower().endswith((".xlsx", ".xls")):
                try:
                    q = pd.read_excel(os.path.join(QUESTIONS_FOLDER, fname))
                    q = q.rename(columns=rename_map)
                    all_q.append(q)
                except Exception:
                    pass

    if all_q:
        q = pd.concat(all_q, ignore_index=True)
    else:
        q = pd.DataFrame(columns=expected)

    for col in expected:
        if col not in q.columns:
            q[col] = np.nan

    q["Standard"] = q["Standard"].astype(str).str.strip()
    return q[expected]


@st.cache_data
def get_info_for_standard(standards_df, selected_standard):
    if standards_df.empty or selected_standard == "":
        return 0, 0, "00", "00", "00"
    try:
        short_name = standards_df.loc[
            standards_df["Standard"].str.upper() == str(selected_standard).strip().upper(),
            "ShortName"
        ].values[0]
    except Exception:
        short_name = selected_standard
    sheet_name = str(short_name).strip() if str(short_name).strip() else selected_standard
    try:
        vals = pd.read_excel(INFO_FILE, sheet_name=sheet_name, header=None)[1].values
        total = int(vals[0])
        criteria = float(vals[1])
        h = f"{int(vals[2]):02d}"
        m = f"{int(vals[3]):02d}"
        s = f"{int(vals[4]):02d}"
        return total, criteria, h, m, s
    except Exception:
        return 0, 0, "00", "00", "00"

# =====================
# Helpers
# =====================
def start_quiz_session(emp_id, emp_name, standard, questions_df, total):
    if standard == "Cummulative":
        cand = questions_df.copy()
    else:
        cand = questions_df[
            questions_df["Standard"].astype(str).str.strip().str.upper()
            == str(standard).strip().upper()
        ]
    cand = cand.dropna(subset=["Question","A","B","C","D","Answer"])
    if total <= 0 or cand.empty:
        return False, "Questions not defined for this standard."
    if len(cand) < total:
        total = len(cand)
    sampled = cand.sample(total, random_state=None).reset_index(drop=True)

    st.session_state.quiz = {
        "emp_id": str(emp_id),
        "emp_name": str(emp_name),
        "standard": str(standard),
        "total": int(total),
        "rows": sampled,
        "queue": list(range(int(total))),
        "right": 0,
        "wrong": 0,
        "answers": {},
        "start_ts": time.time(),
    }
    return True, ""


def format_timer(h, m, s):
    try:
        hh = int(h); mm = int(m); ss = int(s)
        return hh*3600 + mm*60 + ss
    except Exception:
        return 0

# =====================
# Save to Google Sheets
# =====================
def append_result(emp_id, emp_name, total, right, wrong, criteria_pct, status, test_type):
    try:
        sheet = client.open_by_url(GSHEET_URL)
        worksheet = sheet.worksheet("Result")

        now = dt.datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")
        pct = (right/total)*100 if total else 0.0

        new_row = [
            str(emp_id), str(emp_name), int(total), int(right), int(wrong),
            f"{pct:.2f}%", f"{criteria_pct:.0f}%", str(status), str(test_type), now
        ]

        worksheet.append_row(new_row)
        return True, ""
    except Exception as e:
        return False, str(e)

# =====================
# UI
# =====================
st.set_page_config(page_title="PTIS Online Testing", page_icon="📝", layout="centered")
st.title("PTIS Online Testing Module")

employees, standards = load_employees_and_standards()
questions = load_questions()

# Counter for reset
if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

if "quiz" not in st.session_state:
    st.subheader("Login")

    emp_id = st.text_input("Employee ID", value="", key=f"id_{st.session_state.reset_counter}")

    fetched_name = ""
    if emp_id and not employees.empty:
        try:
            fetched = employees[employees.iloc[:,0].astype(str).str.strip() == str(emp_id).strip()]
            if not fetched.empty:
                fetched_name = str(fetched.iloc[0,1])
        except Exception:
            pass
    name = st.text_input("Name (auto-fills if ID found)", value=fetched_name, key=f"name_{st.session_state.reset_counter}")

    options = standards["Standard"].dropna().unique().tolist()
    options = sorted(options)
    if "Cummulative" not in options:
        options = ["Cummulative"] + options
    selected_standard = st.selectbox("Select Standard", options, index=0 if options else None, key=f"std_{st.session_state.reset_counter}")

    total, criteria, h, m, s = get_info_for_standard(standards, selected_standard)

    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Total Questions", total)
    with c2: st.metric("Passing Criteria (%)", criteria)
    with c3: st.metric("Timer (HH:MM:SS)", f"{h}:{m}:{s}")

    with st.form("start_form"):
        submitted = st.form_submit_button("Start Test")

    if submitted:
        if not emp_id or not name or not selected_standard:
            st.error("Please enter ID, Name and select a Standard.")
        else:
            ok, msg = start_quiz_session(emp_id, name, selected_standard, questions, total)
            if not ok:
                st.error(msg)
            else:
                st.rerun()

else:
    qstate = st.session_state.quiz
    total, criteria, h, m, s = get_info_for_standard(standards, qstate["standard"])
    total_secs = format_timer(h, m, s)

    elapsed = int(time.time() - qstate["start_ts"])
    remaining = max(0, total_secs - elapsed)
    if total_secs > 0:
        rem_h, rem_m, rem_s = remaining // 3600, (remaining % 3600) // 60, remaining % 60
        st.info(f"⏰ Time Left: {rem_h:02d}:{rem_m:02d}:{rem_s:02d}")

        if remaining <= 0 and len(qstate["queue"]) > 0:
            qstate["wrong"] += len(qstate["queue"])
            qstate["queue"] = []
            st.session_state.quiz = qstate
            st.rerun()

    answered_count = qstate["total"] - len(qstate["queue"])

    st.markdown(
        f"""
        <div style="padding:10px; border-radius:10px; background-color:#1E3A8A; color:white; text-align:center; font-size:16px; margin-bottom:15px;">
            <b>ID:</b> {qstate['emp_id']} &nbsp; • &nbsp;
            <b>Name:</b> {qstate['emp_name']} &nbsp; • &nbsp;
            <b>Standard:</b> {qstate['standard']} &nbsp; • &nbsp;
            <b>Answered:</b> {answered_count}/{qstate['total']}
        </div>
        """,
        unsafe_allow_html=True
    )

    if len(qstate["queue"]) > 0:
        current_qid = qstate["queue"][0]
        row = qstate["rows"].iloc[current_qid]
        qno, question, A, B, C, D, correct = row["Qno"], row["Question"], row["A"], row["B"], row["C"], row["D"], row["Answer"]

        st.subheader(f"Q{current_qid+1}. {question}")
        choice = st.radio("Choose your answer:", [A, B, C, D], index=None, key=f"q_{current_qid}")

        col1, col2 = st.columns([1,1])

        with col1:
            if st.button("Next", use_container_width=True):
                if choice is None:
                    st.warning("⚠️ Please select an option before moving on.")
                else:
                    mapping = {"A": A, "B": B, "C": C, "D": D}
                    correct_text = mapping.get(str(correct).strip(), str(correct).strip())
                    is_correct = str(choice).strip() == str(correct_text).strip()
                    qstate["answers"][current_qid] = {
                        "choice": choice,
                        "correct": correct_text,
                        "is_correct": is_correct
                    }
                    if is_correct:
                        qstate["right"] += 1
                    else:
                        qstate["wrong"] += 1
                    qstate["queue"].pop(0)
                    st.session_state.quiz = qstate
                    st.rerun()

        with col2:
            if len(qstate["queue"]) > 1:
                if st.button("Skip", use_container_width=True):
                    qstate["queue"].append(qstate["queue"].pop(0))
                    st.session_state.quiz = qstate
                    st.rerun()

    if len(qstate["queue"]) == 0:
        right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
        pct = (right/total_q)*100 if total_q else 0.0
        status = "Pass" if pct >= float(criteria) else "Fail"

        if "submitted" not in st.session_state:
            st.success("All questions attempted. You can now submit your test.")

            submit_clicked = st.button("Submit", use_container_width=True)
            if submit_clicked:
                ok, msg = append_result(
                    qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
                )
                st.session_state["submitted"] = True
                st.session_state["submit_result"] = (ok, msg, right, total_q, pct, criteria, status)
                st.rerun()
        else:
            if "submit_result" in st.session_state:
                ok, msg, right, total_q, pct, criteria, status = st.session_state["submit_result"]

                color = "#16A34A" if status == "Pass" else "#DC2626"
                st.markdown(
                    f"""
                    <div style="padding:20px; border-radius:12px; background-color:#111827; color:white; text-align:center; margin-top:20px;">
                        <h3 style="color:{color};">Final Result : {status}</h3>
                        <p style="font-size:18px;">
                            <b>Score:</b> {right}/{total_q}<br>
                            <b>Percentage:</b> {pct:.2f}%<br>
                            <b>Passing Criteria:</b> {criteria:.0f}%
                        </p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
