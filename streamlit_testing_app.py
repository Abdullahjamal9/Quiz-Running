import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import time
import os

st.write("Working directory:", os.getcwd())
st.write("Files in root:", os.listdir())
if os.path.exists("db"):
    st.write("Files in db/:", os.listdir("db"))
else:
    st.error("❌ db folder not found!")
# =====================
# Paths / Files
# =====================
DB_FOLDER = "db"
QUESTIONS_FOLDER = os.path.join(DB_FOLDER, "Questions")
EMP_STD_FILE = os.path.join(DB_FOLDER, "Result 2.xlsx")
INFO_FILE = os.path.join(DB_FOLDER, "info.xlsx")

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

    # Queue-based navigation
    st.session_state.quiz = {
        "emp_id": str(emp_id),
        "emp_name": str(emp_name),
        "standard": str(standard),
        "total": int(total),
        "rows": sampled,
        "queue": list(range(int(total))),  # indices of questions
        "right": 0,
        "wrong": 0,
        "answers": {},  # qid -> {"choice": text, "correct": text, "is_correct": bool}
        "start_ts": time.time(),
    }
    return True, ""

def format_timer(h, m, s):
    try:
        hh = int(h); mm = int(m); ss = int(s)
        return hh*3600 + mm*60 + ss
    except Exception:
        return 0

def append_result(emp_id, emp_name, total, right, wrong, criteria_pct, status, test_type):
    try:
        now = dt.datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")
        pct = (right/total)*100 if total else 0.0
        row = [int(emp_id), emp_name, int(total), int(right), int(wrong),
               f"{pct:.2f}%", f"{criteria_pct:.0f}%", status, test_type, now]
        try:
            df_old = pd.read_excel(EMP_STD_FILE, sheet_name="Result")
        except Exception:
            df_old = pd.DataFrame(columns=[
                "ID","Name","Total","Right","Wrong","%","Criteria%",
                "Status","Type","DateTime"
            ])
        cols = df_old.columns.tolist() if len(df_old.columns) else [
            "ID","Name","Total","Right","Wrong","%","Criteria%",
            "Status","Type","DateTime"
        ]
        df_new = pd.DataFrame([row], columns=cols)
        out = pd.concat([df_old, df_new], ignore_index=True)

        with pd.ExcelWriter(EMP_STD_FILE, mode="a", if_sheet_exists="replace") as xw:
            try:
                emp = pd.read_excel(EMP_STD_FILE, sheet_name="Emloyees Data")
                emp.to_excel(xw, sheet_name="Emloyees Data", index=False)
            except Exception:
                pass
            try:
                std = pd.read_excel(EMP_STD_FILE, sheet_name="Standard")
                std.to_excel(xw, sheet_name="Standard", index=False)
            except Exception:
                pass
            out.to_excel(xw, sheet_name="Result", index=False)
        return True, ""
    except Exception as e:
        return False, str(e)

# =====================
# UI
# =====================
st.set_page_config(page_title="PTIS Online Testing (Streamlit)", page_icon="📝", layout="centered")
st.title("PTIS Online Testing Module (Streamlit)")

employees, standards = load_employees_and_standards()
questions = load_questions()

if "quiz" not in st.session_state:
    with st.form("login_form", clear_on_submit=False):
        st.subheader("Login")
        emp_id = st.text_input("Employee ID", value="", help="Enter your numeric ID")
        fetched_name = ""
        if emp_id and not employees.empty:
            try:
                fetched = employees[employees.iloc[:,0].astype(str).str.strip() == str(emp_id).strip()]
                if not fetched.empty:
                    fetched_name = str(fetched.iloc[0,1])
            except Exception:
                pass
        name = st.text_input("Name (auto-fills if ID found)", value=fetched_name)
        options = standards["Standard"].dropna().unique().tolist()
        options = sorted(options)
        if "Cummulative" not in options:
            options = ["Cummulative"] + options
        selected_standard = st.selectbox("Select Standard", options, index=0 if options else None)

        total, criteria, h, m, s = get_info_for_standard(standards, selected_standard)

        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Total Questions", total)
        with c2: st.metric("Passing Criteria (%)", criteria)
        with c3: st.metric("Timer (HH:MM:SS)", f"{h}:{m}:{s}")

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
    rem_h, rem_m, rem_s = remaining // 3600, (remaining % 3600) // 60, remaining % 60

    answered_count = qstate["total"] - len(qstate["queue"])
    st.caption(
        f"ID: {qstate['emp_id']} • Name: {qstate['emp_name']} • "
        f"Standard: {qstate['standard']} • Answered: {answered_count}/{qstate['total']}"
    )
    st.info(f"Time Left: {rem_h:02d}:{rem_m:02d}:{rem_s:02d}")

    # Handle time up
    if total_secs > 0 and remaining <= 0 and len(qstate["queue"]) > 0:
        qstate["wrong"] += len(qstate["queue"])
        qstate["queue"] = []
        st.session_state.quiz = qstate
        st.rerun()

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

        st.success("All questions attempted. You can now submit your test.")
        if st.button("Submit", use_container_width=True):
            ok, msg = append_result(
                qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
            )
            if ok:
                st.info("✅ Result saved to 'Result' sheet in 'Result 2.xlsx'.")
            else:
                st.warning("⚠️ Could not write result to Excel. You can download your results below.")
                out = pd.DataFrame([{
                    "ID": qstate["emp_id"],
                    "Name": qstate["emp_name"],
                    "Total": total_q,
                    "Right": right,
                    "Wrong": wrong,
                    "%": f"{pct:.2f}%",
                    "Criteria%": f"{criteria:.0f}%",
                    "Status": status,
                    "Type": qstate["standard"],
                    "DateTime": dt.datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")
                }])
                st.download_button(
                    "📥 Download Result CSV",
                    out.to_csv(index=False).encode("utf-8"),
                    file_name=f"result_{qstate['emp_id']}.csv",
                    mime="text/csv"
                )
            st.write(f"**Score:** {right}/{total_q} • **Percentage:** {pct:.2f}% • Passing Criteria: {criteria:.0f}% • **Status:** {status}")

            if st.button("Restart / New Attempt"):
                for k in list(st.session_state.keys()):
                    if k.startswith("q_"): del st.session_state[k]
                del st.session_state["quiz"]
                st.rerun()
