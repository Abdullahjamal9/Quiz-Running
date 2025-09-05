import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import time
import os

# =====================
# Paths / Files
# =====================
DB_FOLDER = "db"
QUESTIONS_FOLDER = os.path.join(DB_FOLDER, "Questions")
EMP_STD_FILE = os.path.join(DB_FOLDER, "Result 2.xlsx")
INFO_FILE = os.path.join(DB_FOLDER, "info.xlsx")

# =====================
# Custom CSS Styling
# =====================
def apply_custom_styles():
    st.markdown("""
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    
    /* Global Styling */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Poppins', sans-serif;
    }
    
    /* Main container styling */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 900px;
    }
    
    /* Title styling */
    h1 {
        color: white !important;
        text-align: center;
        font-weight: 700;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        margin-bottom: 2rem;
        font-size: 2.5rem !important;
    }
    
    /* Subheader styling */
    h2, h3 {
        color: white !important;
        font-weight: 600;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
    }
    
    /* Login form container */
    .login-container {
        background: rgba(255, 255, 255, 0.95);
        padding: 2.5rem;
        border-radius: 20px;
        box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255,255,255,0.2);
        margin-bottom: 2rem;
    }
    
    /* Input styling */
    .stTextInput input {
        border-radius: 12px !important;
        border: 2px solid #e1e5e9 !important;
        padding: 12px 16px !important;
        font-size: 16px !important;
        transition: all 0.3s ease !important;
        background: white !important;
    }
    
    .stTextInput input:focus {
        border-color: #667eea !important;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1) !important;
    }
    
    /* Selectbox styling */
    .stSelectbox select {
        border-radius: 12px !important;
        border: 2px solid #e1e5e9 !important;
        padding: 12px 16px !important;
        font-size: 16px !important;
        background: white !important;
    }
    
    /* Button styling */
    .stButton button {
        background: linear-gradient(135deg, #667eea, #764ba2) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 12px 24px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3) !important;
    }
    
    .stButton button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4) !important;
    }
    
    /* Form submit button */
    .stForm button {
        background: linear-gradient(135deg, #4CAF50, #45a049) !important;
        width: 100% !important;
        padding: 16px !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        margin-top: 1rem !important;
    }
    
    /* Metrics styling */
    .metric-card {
        background: linear-gradient(135deg, #fff, #f8f9ff);
        padding: 1.5rem;
        border-radius: 16px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        text-align: center;
        border: 1px solid rgba(102, 126, 234, 0.1);
        transition: transform 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 40px rgba(0,0,0,0.15);
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #667eea;
        margin-bottom: 0.5rem;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #64748b;
        font-weight: 500;
    }
    
    /* Question container */
    .question-container {
        background: rgba(255, 255, 255, 0.95);
        padding: 2rem;
        border-radius: 20px;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        backdrop-filter: blur(10px);
    }
    
    /* Question title */
    .question-title {
        color: #2d3748 !important;
        font-size: 1.3rem !important;
        font-weight: 600 !important;
        margin-bottom: 1.5rem !important;
        line-height: 1.6 !important;
    }
    
    /* Radio button styling */
    .stRadio > div {
        background: white;
        padding: 1rem;
        border-radius: 12px;
        margin-top: 1rem;
    }
    
    .stRadio label {
        padding: 12px 16px !important;
        margin: 8px 0 !important;
        border-radius: 10px !important;
        background: #f8f9fa !important;
        border: 2px solid #e9ecef !important;
        transition: all 0.3s ease !important;
        cursor: pointer !important;
        display: block !important;
    }
    
    .stRadio label:hover {
        background: #e3f2fd !important;
        border-color: #667eea !important;
    }
    
    /* Success/Error messages */
    .stSuccess, .stError, .stWarning, .stInfo {
        border-radius: 12px !important;
        border: none !important;
    }
    
    /* Progress info bar */
    .progress-bar {
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 16px;
        text-align: center;
        font-weight: 500;
        margin-bottom: 1.5rem;
        box-shadow: 0 8px 32px rgba(102, 126, 234, 0.3);
    }
    
    /* Final result styling */
    .result-container {
        background: linear-gradient(135deg, #1a202c, #2d3748);
        color: white;
        padding: 2.5rem;
        border-radius: 20px;
        text-align: center;
        margin-top: 2rem;
        box-shadow: 0 20px 40px rgba(0,0,0,0.2);
    }
    
    .result-title {
        font-size: 2rem !important;
        font-weight: 700 !important;
        margin-bottom: 1.5rem !important;
    }
    
    .result-details {
        font-size: 1.1rem !important;
        line-height: 2 !important;
    }
    
    /* Animation keyframes */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    @keyframes pulse {
        0% { transform: scale(1); opacity: 1; }
        50% { transform: scale(1.05); opacity: 0.8; }
        100% { transform: scale(1); opacity: 1; }
    }
    
    .timer-pulse {
        animation: pulse 1s infinite;
    }
    
    /* Fade in animation for containers */
    .login-container, .question-container {
        animation: fadeIn 0.8s ease-out;
    }
    
    /* Hide Streamlit elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

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

def show_live_timer(standards, qstate):
    """Live timer with auto-refresh and enhanced styling"""
    # Auto-refresh every second only during quiz
    if len(qstate.get("queue", [])) > 0:
        st_autorefresh(interval=1000, limit=None, key="timer_refresh")
    
    total, criteria, h, m, s = get_info_for_standard(standards, qstate["standard"])
    total_secs = format_timer(h, m, s)
    
    if total_secs > 0:
        elapsed = int(time.time() - qstate["start_ts"])
        remaining = max(0, total_secs - elapsed)
        
        # Auto-submit if time is up
        if remaining <= 0 and len(qstate["queue"]) > 0:
            st.error("⏰ Time is up! Auto-submitting your test...")
            qstate["wrong"] += len(qstate["queue"])
            qstate["queue"] = []
            st.session_state.quiz = qstate
            st.rerun()
            return
        
        rem_h, rem_m, rem_s = remaining // 3600, (remaining % 3600) // 60, remaining % 60
        
        # Enhanced color coding with gradients
        if remaining <= 300:  # Last 5 minutes - critical red
            bg_gradient = "linear-gradient(135deg, #DC2626, #B91C1C, #991B1B)"
            text_color = "white"
            icon = "🚨"
            pulse_class = "timer-pulse"
            border_color = "#DC2626"
        elif remaining <= 900:  # Last 15 minutes - warning red
            bg_gradient = "linear-gradient(135deg, #EA580C, #DC2626)"
            text_color = "white"
            icon = "⚠️"
            pulse_class = ""
            border_color = "#EA580C"
        elif remaining <= 1800:  # Last 30 minutes - caution orange
            bg_gradient = "linear-gradient(135deg, #D97706, #F59E0B)"
            text_color = "white"
            icon = "⏰"
            pulse_class = ""
            border_color = "#D97706"
        else:  # Normal - gradient blue
            bg_gradient = "linear-gradient(135deg, #667eea, #764ba2)"
            text_color = "white"
            icon = "⏱️"
            pulse_class = ""
            border_color = "#667eea"
        
        # Progress bar percentage
        progress_percent = (remaining / total_secs) * 100
        
        timer_html = f"""
        <div class="timer-container {pulse_class}" style="
            background: {bg_gradient};
            color: {text_color};
            padding: 25px;
            border-radius: 20px;
            text-align: center;
            font-size: 22px;
            font-weight: 600;
            margin-bottom: 25px;
            box-shadow: 0 15px 35px rgba(0, 0, 0, 0.2), 0 5px 15px rgba(0, 0, 0, 0.1);
            border: 3px solid {border_color};
            backdrop-filter: blur(10px);
            position: relative;
            overflow: hidden;
        ">
            <!-- Decorative background pattern -->
            <div style="
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: radial-gradient(circle at 20% 80%, rgba(255,255,255,0.1) 0%, transparent 50%),
                           radial-gradient(circle at 80% 20%, rgba(255,255,255,0.1) 0%, transparent 50%);
                pointer-events: none;
            "></div>
            
            <div style="
                display: flex; 
                align-items: center; 
                justify-content: center; 
                gap: 20px;
                position: relative;
                z-index: 2;
            ">
                <span style="font-size: 32px; filter: drop-shadow(2px 2px 4px rgba(0,0,0,0.3));">{icon}</span>
                <span style="font-weight: 500; text-shadow: 1px 1px 2px rgba(0,0,0,0.3);">Time Remaining:</span>
                <span style="
                    font-family: 'Courier New', monospace; 
                    font-size: 32px; 
                    background: rgba(0,0,0,0.3); 
                    padding: 8px 20px; 
                    border-radius: 12px;
                    box-shadow: inset 0 2px 4px rgba(0,0,0,0.3);
                    border: 1px solid rgba(255,255,255,0.2);
                    text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
                ">
                    {rem_h:02d}:{rem_m:02d}:{rem_s:02d}
                </span>
            </div>
            
            <!-- Enhanced progress bar -->
            <div style="
                width: 100%;
                height: 8px;
                background: rgba(0,0,0,0.3);
                border-radius: 4px;
                overflow: hidden;
                margin-top: 20px;
                position: relative;
                z-index: 2;
                box-shadow: inset 0 2px 4px rgba(0,0,0,0.3);
            ">
                <div style="
                    height: 100%;
                    background: linear-gradient(90deg, #10B981, #34D399, #6EE7B7);
                    width: {progress_percent:.1f}%;
                    border-radius: 4px;
                    transition: width 1s ease-in-out;
                    box-shadow: 0 0 10px rgba(16, 185, 129, 0.5);
                    position: relative;
                ">
                    <div style="
                        position: absolute;
                        top: 0;
                        left: 0;
                        right: 0;
                        bottom: 0;
                        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
                        animation: shimmer 2s infinite;
                    "></div>
                </div>
            </div>
        </div>
        
        <style>
        @keyframes shimmer {{
            0% {{ transform: translateX(-100%); }}
            100% {{ transform: translateX(100%); }}
        }}
        </style>
        """
        
        st.markdown(timer_html, unsafe_allow_html=True)
        
        # Enhanced warning messages with better styling
        if remaining <= 300:
            st.error("🚨 **CRITICAL**: Less than 5 minutes remaining!")
        elif remaining <= 900:
            st.warning("⚠️ **WARNING**: Less than 15 minutes remaining!")
        elif remaining <= 1800:
            st.info("⏰ **NOTICE**: Less than 30 minutes remaining!")

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
st.set_page_config(
    page_title="PTIS Online Testing", 
    page_icon="🎓", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Apply custom styling
apply_custom_styles()

st.title("🎓 PTIS Online Testing Module")

employees, standards = load_employees_and_standards()
questions = load_questions()

# Counter for reset
if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

if "quiz" not in st.session_state:
    # Login Section with enhanced styling
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    
    st.markdown("### 📋 Student Login")
    st.markdown("Please enter your details to begin the test")
    
    emp_id = st.text_input(
        "🆔 Employee ID", 
        value="", 
        key=f"id_{st.session_state.reset_counter}",
        placeholder="Enter your employee ID"
    )

    fetched_name = ""
    if emp_id and not employees.empty:
        try:
            fetched = employees[employees.iloc[:,0].astype(str).str.strip() == str(emp_id).strip()]
            if not fetched.empty:
                fetched_name = str(fetched.iloc[0,1])
        except Exception:
            pass
    
    name = st.text_input(
        "👤 Name (auto-fills if ID found)", 
        value=fetched_name, 
        key=f"name_{st.session_state.reset_counter}",
        placeholder="Enter your full name"
    )

    options = standards["Standard"].dropna().unique().tolist()
    options = sorted(options)
    if "Cummulative" not in options:
        options = ["Cummulative"] + options
    
    selected_standard = st.selectbox(
        "📚 Select Standard", 
        options, 
        index=0 if options else None, 
        key=f"std_{st.session_state.reset_counter}"
    )

    total, criteria, h, m, s = get_info_for_standard(standards, selected_standard)

    # Enhanced metrics display
    st.markdown("### 📊 Test Information")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total}</div>
            <div class="metric-label">📝 Total Questions</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{criteria}%</div>
            <div class="metric-label">🎯 Passing Criteria</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{h}:{m}:{s}</div>
            <div class="metric-label">⏱️ Time Limit</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # Start test form
    with st.form("start_form"):
        st.markdown("### 🚀 Ready to Start?")
        st.markdown("Make sure all your details are correct before starting the test.")
        submitted = st.form_submit_button("🎯 Start Test", use_container_width=True)

    if submitted:
        if not emp_id or not name or not selected_standard:
            st.error("❌ Please enter ID, Name and select a Standard.")
        else:
            ok, msg = start_quiz_session(emp_id, name, selected_standard, questions, total)
            if not ok:
                st.error(f"❌ {msg}")
            else:
                st.success("✅ Test started successfully!")
                time.sleep(1)
                st.rerun()

else:
    qstate = st.session_state.quiz
    
    # Display live timer with auto-refresh
    show_live_timer(standards, qstate)

    answered_count = qstate["total"] - len(qstate["queue"])

    # Enhanced progress info bar
    progress_percentage = (answered_count / qstate["total"]) * 100
    st.markdown(
        f"""
        <div class="progress-bar">
            <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 15px;">
                <div><strong>🆔 ID:</strong> {qstate['emp_id']}</div>
                <div><strong>👤 Name:</strong> {qstate['emp_name']}</div>
                <div><strong>📚 Standard:</strong> {qstate['standard']}</div>
                <div><strong>📊 Progress:</strong> {answered_count}/{qstate['total']} ({progress_percentage:.1f}%)</div>
            </div>
            <div style="
                width: 100%;
                height: 6px;
                background: rgba(255,255,255,0.3);
                border-radius: 3px;
                margin-top: 15px;
                overflow: hidden;
            ">
                <div style="
                    height: 100%;
                    background: linear-gradient(90deg, #10B981, #34D399);
                    width: {progress_percentage:.1f}%;
                    border-radius: 3px;
                    transition: width 0.5s ease;
                "></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if len(qstate["queue"]) > 0:
        current_qid = qstate["queue"][0]
        row = qstate["rows"].iloc[current_qid]
        qno, question, A, B, C, D, correct = row["Qno"], row["Question"], row["A"], row["B"], row["C"], row["D"], row["Answer"]

        # Question container with enhanced styling
        st.markdown('<div class="question-container">', unsafe_allow_html=True)
        
        st.markdown(f'<h3 class="question-title">❓ Question {current_qid+1}: {question}</h3>', unsafe_allow_html=True)
        
        choice = st.radio(
            "Select your answer:", 
            [A, B, C, D], 
            index=None, 
            key=f"q_{current_qid}"
        )
        
        st.markdown('</div>', unsafe_allow_html=True)

        # Action buttons with enhanced styling
        col1, col2 = st.columns([1,1])

        with col1:
            if st.button("➡️ Next Question", use_container_width=True):
                if choice is None:
                    st.warning("⚠️ Please select an option before moving on.")
                else:
