import streamlit as st
import pandas as pd
import numpy as np
import datetime as dt
import time
import os
import gspread
from google.oauth2.service_account import Credentials
import streamlit.components.v1 as components
import pytz

# =====================
# Paths / Files (local Excel for reading only)
# =====================
BASE_DIR = os.path.dirname(__file__)   # absolute path (safe for Streamlit Cloud)
DB_FOLDER = os.path.join(BASE_DIR, "db")
QUESTIONS_FOLDER = os.path.join(DB_FOLDER, "Questions")
INFO_FILE = os.path.join(DB_FOLDER, "info.xlsx")

# =====================
# Google Sheets Setup (for saving and reading results)
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
        sheet = client.open_by_url(GSHEET_URL)
        # Load Employees Data
        try:
            employees_data = sheet.worksheet("Emloyees Data").get_all_records()
            employees = pd.DataFrame(employees_data)
            if len(qstate["queue"]) == 0 and "submitted" not in st.session_state:
                right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
                pct = (right/total_q)*100 if total_q else 0.0
                status = "Pass" if pct >= float(criteria) else "Fail"
    
                st.success("All questions attempted. You can now submit your test.")
    
                submit_clicked = st.button("Submit", use_container_width=True)
            if submit_clicked:
                ok, msg = append_result(
                    qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
                )
                st.session_state["submitted"] = True
                st.session_state["submit_result"] = (ok, msg, right, total_q, pct, criteria, status)
                st.rerun()

        
            if "submit_result" in st.session_state:
                ok, msg, right, total_q, pct, criteria, status = st.session_state["submit_result"]
                if not ok:
                    st.error(f"Failed to save results to Google Sheets: {msg}")

                color = "#043006" if status == "Pass" else "#DC2626"
                st.markdown(
                    f"""
                    <div style="padding:20px; border-radius:12px; background: linear-gradient(135deg, #3B82F6, #2563EB, #1E3A8A); color:white; text-align:center; margin-top:20px;">
                        <h3 style="color:{color}; font-weight:700;">Final Result : <span style="font-weight:700;">{status}</span></h3>
                        <p style="font-size:18px;"><b>Score :</b> {right}/{total_q}<br><b>Percentage :</b> {pct:.2f}%<br><b>Passing Criteria :</b> {criteria:.0f}%</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                employees.empty or not any(col.lower() in ["id", "name"] for col in employees.columns)
                employees = pd.DataFrame(columns=["ID", "Name"])
            else:
                id_col = next((col for col in employees.columns if "id" in col.lower()), "ID")
                name_col = next((col for col in employees.columns if "name" in col.lower()), "Name")
                employees = employees[[id_col, name_col]].rename(columns={id_col: "ID", name_col: "Name"})
        except Exception:
            employees = pd.DataFrame(columns=["ID", "Name"])
        # Load Standards
        try:
            standards_data = sheet.worksheet("Standard").get_all_records()
            standards = pd.DataFrame(standards_data)
            if standards.empty or len(standards.columns) < 2:
                while len(standards.columns) < 2:
                    standards[standards.columns[-1] + "_dup" + str(len(standards.columns))] = ""
                standards.columns = ["Standard", "ShortName"]
            else:
                standards.columns = ["Standard", "ShortName"]
        except Exception:
            standards = pd.DataFrame(columns=["Standard", "ShortName"])
        standards["Standard"] = standards["Standard"].astype(str).str.strip()
        standards["ShortName"] = standards["ShortName"].astype(str).str.strip()
        return employees, standards
    except Exception:
        employees = pd.DataFrame(columns=["ID", "Name"])
        standards = pd.DataFrame(columns=["Standard", "ShortName"])
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


@st.cache_data
def load_all_results():
    try:
        sheet = client.open_by_url(GSHEET_URL)
        
        # Try different possible worksheet names
        worksheet_names = ["Result 2", "Result2", "Result", "Results"]
        worksheet = None
        
        # Try to find the correct worksheet
        for name in worksheet_names:
            try:
                worksheet = sheet.worksheet(name)
                break
            except Exception:
                continue
        
        if worksheet is None:
            # If no worksheet found, try to find any worksheet with "result" in the name
            try:
                all_worksheets = sheet.worksheets()
                for ws in all_worksheets:
                    if "result" in ws.title.lower():
                        worksheet = ws
                        break
            except:
                pass
        
        if worksheet is None:
            st.error("Could not find any results worksheet. Please ensure there's a worksheet named 'Result 2'")
            return pd.DataFrame(columns=["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Timestamp"])
        
        records = worksheet.get_all_records()
        if not records:
            return pd.DataFrame(columns=["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Timestamp"])
        
        df = pd.DataFrame(records)
        
        # Map column names to handle variations
        column_mapping = {
            'ID': ['ID', 'id', 'Id', 'Employee ID', 'EMP ID'],
            'Name': ['NAME', 'Name', 'name', 'Employee Name', 'EMP NAME'],
            'Total': ['TOTAL QUESTION', 'Total Question', 'Total', 'total', 'Total Questions'],
            'Right': ['CORRECT ANSWER', 'Correct Answer', 'Right', 'right', 'Correct'],
            'Wrong': ['WRONG ANSWER', 'Wrong Answer', 'Wrong', 'wrong', 'Incorrect'],
            'Percentage': ['PERCENTAGE', 'Percentage', 'percentage', 'Score', 'score'],
            'Criteria': ['PASSING CRITERIA %', 'Passing Criteria', 'criteria', 'Criteria'],
            'Status': ['STATUS', 'Status', 'status', 'Result'],
            'Test Type': ['STANDARD', 'Standard', 'Test Type', 'test_type'],
            'Timestamp': ['DATE', 'Date', 'date', 'Timestamp', 'timestamp', 'Time']
        }
        
        # Rename columns to standard names
        for standard_name, possible_names in column_mapping.items():
            for col in df.columns:
                if col in possible_names:
                    df = df.rename(columns={col: standard_name})
                    break
        
        # Ensure all required columns exist
        required_columns = ["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Timestamp"]
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""
        
        # Handle numeric columns more robustly
        numeric_cols = ["Total", "Right", "Wrong"]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        # Handle percentage columns more robustly
        if "Percentage" in df.columns:
            # Remove % sign and convert to numeric
            df["Percentage"] = df["Percentage"].astype(str).str.replace("%", "").str.replace(" ", "")
            df["Percentage"] = pd.to_numeric(df["Percentage"], errors='coerce').fillna(0).astype(float)
        
        if "Criteria" in df.columns:
            # Remove % sign and convert to numeric
            df["Criteria"] = df["Criteria"].astype(str).str.replace("%", "").str.replace(" ", "")
            df["Criteria"] = pd.to_numeric(df["Criteria"], errors='coerce').fillna(0).astype(float)
        
        # Sort by timestamp if available (most recent first)
        if "Timestamp" in df.columns:
            try:
                df = df.sort_values("Timestamp", ascending=False)
            except:
                pass
        
        return df[required_columns]
        
    except Exception as e:
        st.error(f"Error loading results: {str(e)}")
        # Show more detailed error information
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return pd.DataFrame(columns=["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Timestamp"])

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
    sampled = cand.sample(total, random_state=int(time.time())).reset_index(drop=True)

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


def append_result(emp_id, emp_name, total, right, wrong, criteria_pct, status, test_type):
    try:
        sheet = client.open_by_url(GSHEET_URL)
        
        # Try different possible worksheet names for saving results
        worksheet_names = ["Result 2", "Result2", "Result", "Results"]
        worksheet = None
        
        # Try to find the correct worksheet
        for name in worksheet_names:
            try:
                worksheet = sheet.worksheet(name)
                break
            except Exception:
                continue
        
        if worksheet is None:
            # If no worksheet found, try to find any worksheet with "result" in the name
            try:
                all_worksheets = sheet.worksheets()
                for ws in all_worksheets:
                    if "result" in ws.title.lower():
                        worksheet = ws
                        break
            except:
                pass
        
        if worksheet is None:
            return False, "Could not find results worksheet to save data"

        pkt_tz = pytz.timezone('Asia/Karachi')
        now = dt.datetime.now(pkt_tz).strftime("%d-%m-%Y %I:%M:%S %p")
        pct = (right/total)*100 if total else 0.0

        # Get headers to ensure we're appending in the right format
        try:
            headers = worksheet.row_values(1)
        except:
            headers = []
        
        # Create row based on existing headers or default format
        if headers:
            # Map data to existing headers
            data_mapping = {
                'ID': str(emp_id),
                'NAME': str(emp_name),
                'TOTAL QUESTION': int(total),
                'CORRECT ANSWER': int(right),
                'WRONG ANSWER': int(wrong),
                'PERCENTAGE': f"{pct:.2f}%",
                'PASSING CRITERIA %': f"{criteria_pct:.0f}%",
                'STATUS': str(status),
                'STANDARD': str(test_type),
                'DATE': now
            }
            
            new_row = []
            for header in headers:
                header_upper = header.upper()
                if header_upper in data_mapping:
                    new_row.append(data_mapping[header_upper])
                elif 'ID' in header_upper:
                    new_row.append(str(emp_id))
                elif 'NAME' in header_upper:
                    new_row.append(str(emp_name))
                elif 'TOTAL' in header_upper and 'QUESTION' in header_upper:
                    new_row.append(int(total))
                elif 'CORRECT' in header_upper:
                    new_row.append(int(right))
                elif 'WRONG' in header_upper:
                    new_row.append(int(wrong))
                elif 'PERCENTAGE' in header_upper:
                    new_row.append(f"{pct:.2f}%")
                elif 'CRITERIA' in header_upper:
                    new_row.append(f"{criteria_pct:.0f}%")
                elif 'STATUS' in header_upper:
                    new_row.append(str(status))
                elif 'STANDARD' in header_upper:
                    new_row.append(str(test_type))
                elif 'DATE' in header_upper or 'TIME' in header_upper:
                    new_row.append(now)
                else:
                    new_row.append('')  # Empty for unknown columns
        else:
            # Default format if no headers found
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

# Admin login state
if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False

# Admin login page
if not st.session_state.admin_logged_in and "quiz" not in st.session_state:
    st.subheader("Admin Login")
    admin_username = st.text_input("Username", key="admin_username")
    admin_password = st.text_input("Password", type="password", key="admin_password")
    if st.button("Login"):
        if admin_username == "admin" and admin_password == "AdminPtis-3692":  # Hardcoded for simplicity
            st.session_state.admin_logged_in = True
            st.rerun()
        else:
            st.error("Invalid username or password")

# Enhanced Admin dashboard with filters
if st.session_state.admin_logged_in:
    st.subheader("Admin Dashboard - Employee Results")
    
    # Add a refresh button to reload the data
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()
    
    results_df = load_all_results()
    if not results_df.empty:
        # Display some summary statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Tests", len(results_df))
        with col2:
            pass_count = len(results_df[results_df["Status"] == "Pass"]) if "Status" in results_df.columns else 0
            st.metric("Passed", pass_count)
        with col3:
            fail_count = len(results_df[results_df["Status"] == "Fail"]) if "Status" in results_df.columns else 0
            st.metric("Failed", fail_count)
        with col4:
            if "Percentage" in results_df.columns and len(results_df) > 0:
                avg_score = results_df["Percentage"].mean()
                st.metric("Avg Score", f"{avg_score:.1f}%")
            else:
                st.metric("Avg Score", "N/A")
        
        st.markdown("---")
        
        # Add filters section
        st.subheader("🔍 Filters")
        
        # Create filter columns
        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        
        with filter_col1:
            # Employee ID filter
            employee_ids = ["All"] + sorted(results_df["ID"].astype(str).unique().tolist())
            selected_emp_id = st.selectbox("Filter by Employee ID", employee_ids, key="emp_id_filter")
        
        with filter_col2:
            # Employee Name filter
            employee_names = ["All"] + sorted(results_df["Name"].unique().tolist())
            selected_emp_name = st.selectbox("Filter by Employee Name", employee_names, key="emp_name_filter")
        
        with filter_col3:
            # Status filter
            statuses = ["All"] + sorted(results_df["Status"].unique().tolist())
            selected_status = st.selectbox("Filter by Status", statuses, key="status_filter")
        
        with filter_col4:
            # Test Type/Standard filter
            test_types = ["All"] + sorted(results_df["Test Type"].unique().tolist())
            selected_test_type = st.selectbox("Filter by Test Type", test_types, key="test_type_filter")
        
        # Additional filters row
        filter_col5, filter_col6, filter_col7, filter_col8 = st.columns(4)
        
        with filter_col5:
            # Percentage range filter
            if "Percentage" in results_df.columns:
                min_percentage = st.number_input("Min Percentage (%)", 
                                               min_value=0.0, 
                                               max_value=100.0, 
                                               value=0.0, 
                                               step=1.0,
                                               key="min_percentage_filter")
        
        with filter_col6:
            if "Percentage" in results_df.columns:
                max_percentage = st.number_input("Max Percentage (%)", 
                                               min_value=0.0, 
                                               max_value=100.0, 
                                               value=100.0, 
                                               step=1.0,
                                               key="max_percentage_filter")
        
        with filter_col7:
            st.write("")  # Empty column for spacing
        
        with filter_col8:
            # Clear filters button aligned with input boxes
            if st.button("🗑️ Clear All Filters", use_container_width=True):
                for key in ["emp_id_filter", "emp_name_filter", "status_filter", "test_type_filter", 
                           "min_percentage_filter", "max_percentage_filter"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
        
        # Apply filters
        filtered_df = results_df.copy()
        
        if selected_emp_id != "All":
            filtered_df = filtered_df[filtered_df["ID"].astype(str) == selected_emp_id]
        
        if selected_emp_name != "All":
            filtered_df = filtered_df[filtered_df["Name"] == selected_emp_name]
        
        if selected_status != "All":
            filtered_df = filtered_df[filtered_df["Status"] == selected_status]
        
        if selected_test_type != "All":
            filtered_df = filtered_df[filtered_df["Test Type"] == selected_test_type]
        
        if "Percentage" in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df["Percentage"] >= min_percentage) & 
                (filtered_df["Percentage"] <= max_percentage)
            ]
        
        # Show filtered results count
        if len(filtered_df) != len(results_df):
            st.info(f"Showing {len(filtered_df)} of {len(results_df)} total records")
        
        st.markdown("---")
        
        # Display the filtered results table
        if not filtered_df.empty:
            # Add export functionality
            export_col1, export_col2, export_col3 = st.columns([1, 1, 2])
            
            with export_col1:
                # Download as CSV
                csv = filtered_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download as CSV",
                    data=csv,
                    file_name=f"test_results_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with export_col2:
                # Show/Hide columns
                if st.button("⚙️ Column Settings"):
                    st.session_state.show_column_settings = not st.session_state.get("show_column_settings", False)
            
            # Column visibility settings
            if st.session_state.get("show_column_settings", False):
                st.subheader("Column Visibility")
                cols_to_show = []
                col_settings = st.columns(5)
                for i, col in enumerate(filtered_df.columns):
                    with col_settings[i % 5]:
                        if st.checkbox(col, value=True, key=f"show_{col}"):
                            cols_to_show.append(col)
                filtered_df = filtered_df[cols_to_show] if cols_to_show else filtered_df
            
            # Display the dataframe with enhanced formatting
            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Percentage": st.column_config.ProgressColumn(
                        "Percentage",
                        help="Test Score Percentage",
                        format="%.1f%%",
                        min_value=0,
                        max_value=100,
                    ),
                    "Status": st.column_config.TextColumn(
                        "Status",
                        help="Pass/Fail Status",
                    ),
                    "Total": st.column_config.NumberColumn(
                        "Total Questions",
                        help="Total number of questions in the test",
                        format="%d"
                    ),
                    "Right": st.column_config.NumberColumn(
                        "Correct Answers",
                        help="Number of correct answers",
                        format="%d"
                    ),
                    "Wrong": st.column_config.NumberColumn(
                        "Wrong Answers", 
                        help="Number of wrong answers",
                        format="%d"
                    )
                }
            )
        else:
            st.warning("No results found matching the current filters.")
    else:
        st.info("No results available yet in the Result 2 sheet.")
    
    if st.button("Logout"):
        st.session_state.admin_logged_in = False
        st.session_state.pop("quiz", None)
        st.rerun()

# Employee login and quiz
if not st.session_state.admin_logged_in:
    if "reset_counter" not in st.session_state:
        st.session_state.reset_counter = 0

    if "quiz" not in st.session_state:
        st.subheader("👤 Employee Login")

        col1, col2 = st.columns(2)
        
        with col1:
            emp_id = st.text_input(
                "Employee ID", 
                value="", 
                key=f"id_{st.session_state.reset_counter}",
                help="Enter your employee identification number",
                on_change=lambda: st.session_state.update({"name": fetch_name(employees, st.session_state[f"id_{st.session_state.reset_counter}"])})
            )

        def fetch_name(employees_df, emp_id_input):
            if emp_id_input and not employees_df.empty:
                try:
                    fetched = employees_df[employees_df["ID"].astype(str).str.strip() == str(emp_id_input).strip()]
                    if not fetched.empty:
                        return str(fetched.iloc[0]["Name"])
                except Exception:
                    pass
            return ""

        fetched_name = fetch_name(employees, emp_id) if "name" not in st.session_state else st.session_state["name"]
        
        with col2:
            name = st.text_input(
                "Name", 
                value=fetched_name, 
                key=f"name_{st.session_state.reset_counter}",
                help="This will auto-fill if your Employee ID is found"
            )

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

        st.markdown("---")
        with st.form("start_form"):
            st.markdown("### Ready to start your test?")
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                submitted = st.form_submit_button("🚀 Start Test", use_container_width=True)

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

        if total_secs > 0 and len(qstate["queue"]) > 0 and "submitted" not in st.session_state:
            if remaining <= 0:
                st.error("Time is up! Auto-submitting your test...")
                qstate["wrong"] += len(qstate["queue"])
                qstate["queue"] = []
                st.session_state.quiz = qstate
                
                right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
                pct = (right/total_q)*100 if total_q else 0.0
                status = "Pass" if pct >= float(criteria) else "Fail"
                ok, msg = append_result(
                    qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
                )
                st.session_state["submitted"] = True
                st.session_state["submit_result"] = (ok, msg, right, total_q, pct, criteria, status)
                st.query_params.clear()
                st.rerun()

            rem_h = remaining // 3600
            rem_m = (remaining % 3600) // 60
            rem_s = remaining % 60

            if remaining <= 300:
                bg_color = "#DC2626"
                text_color = "white"
                icon = "🚨"
                pulse_class = "timer-pulse"
            elif remaining <= 900:
                bg_color = "#DC2626"
                text_color = "white"
                icon = "⚠️"
                pulse_class = ""
            elif remaining <= 1800:
                bg_color = "#D97706"
                text_color = "white"
                icon = "⏰"
                pulse_class = ""
            else:
                bg_color = "#1E3A8A"
                text_color = "white"
                icon = "⏰"
                pulse_class = ""

            progress_percent = (remaining / total_secs) * 100 if total_secs > 0 else 0

            timer_html = f"""
            <style>
            @keyframes pulse {{
                0% {{ transform: scale(1); opacity: 1; }}
                50% {{ transform: scale(1.05); opacity: 0.8; }}
                100% {{ transform: scale(1); opacity: 1; }}
            }}
            .timer-pulse {{
                animation: pulse 1s infinite;
            }}
            .timer-container {{
                padding: 20px;
                border-radius: 15px;
                text-align: center;
                font-size: 22px;
                font-weight: bold;
                margin-bottom: 20px;
                box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
                border: 3px solid rgba(255, 255, 255, 0.1);
            }}
            </style>
            <div id="timer_container" class="timer-container {pulse_class}" style="background: linear-gradient(135deg, {bg_color}, {bg_color}CC); color: {text_color};">
                <div style="display: flex; align-items: center; justify-content: center; gap: 15px;">
                    <span id="timer_icon" style="font-size: 28px;">{icon}</span>
                    <span>Time Remaining :</span>
                    <span id="timer_display" style="font-family: 'Courier New', monospace; font-size: 28px; background: rgba(0,0,0,0.2); padding: 5px 15px; border-radius: 8px;">
                        {rem_h:02d}:{rem_m:02d}:{rem_s:02d}
                    </span>
                </div>
                <div style="width: 100%; height: 6px; background-color: rgba(255,255,255,0.3); border-radius: 3px; overflow: hidden; margin-top: 15px;">
                    <div id="progress_bar" style="height: 100%; background: linear-gradient(90deg, #10B981, #34D399); width: {progress_percent:.1f}%; border-radius: 3px; transition: width 0.5s ease-in-out;"></div>
                </div>
            </div>
            <script>
            (function() {{
                var remaining = {remaining};
                var total_secs = {total_secs};
                var interval = null;

                function updateTimer() {{
                    if (remaining <= 0) {{
                        document.getElementById('timer_display').innerText = '00:00:00';
                        document.getElementById('progress_bar').style.width = '0%';
                        clearInterval(interval);
                        var form = document.createElement('form');
                        form.method = 'POST';
                        form.action = window.location.href;
                        var input = document.createElement('input');
                        input.type = 'hidden';
                        input.name = 'timeout';
                        input.value = 'true';
                        form.appendChild(input);
                        document.body.appendChild(form);
                        form.submit();
                        return;
                    }}
                    var h = Math.floor(remaining / 3600);
                    var m = Math.floor((remaining % 3600) / 60);
                    var s = remaining % 60;
                    document.getElementById('timer_display').innerText = `${{h.toString().padStart(2, '0')}}:${{m.toString().padStart(2, '0')}}:${{s.toString().padStart(2, '0')}}`;
                    var progress = (remaining / total_secs) * 100;
                    document.getElementById('progress_bar').style.width = progress + '%';
                    var container = document.getElementById('timer_container');
                    var iconElem = document.getElementById('timer_icon');
                    var bg_color, text_color, icon, pulse_class = '';
                    if (remaining <= 300) {{
                        bg_color = '#DC2626';
                        text_color = 'white';
                        icon = '🚨';
                        pulse_class = 'timer-pulse';
                    }} else if (remaining <= 900) {{
                        bg_color = '#DC2626';
                        text_color = 'white';
                        icon = '⚠️';
                    }} else if (remaining <= 1800) {{
                        bg_color = '#D97706';
                        text_color = 'white';
                        icon = '⏰';
                    }} else {{
                        bg_color = '#1E3A8A';
                        text_color = 'white';
                        icon = '⏰';
                    }}
                    container.style.background = `linear-gradient(135deg, ${bg_color}, ${bg_color}CC)`;
                    container.style.color = text_color;
                    iconElem.innerText = icon;
                    if (pulse_class) {{
                        container.classList.add(pulse_class);
                    }} else {{
                        container.classList.remove('timer-pulse');
                    }}
                    remaining--;
                }}

                if (interval) {{
                    clearInterval(interval);
                }}
                updateTimer();
                interval = setInterval(updateTimer, 1000);
            }})();
            </script>
            """
            components.html(timer_html, height=150)

            if st.query_params.get("timeout", ["false"])[0] == "true":
                if len(qstate["queue"]) > 0:
                    st.error("Time is up! Auto-submitting your test...")
                    qstate["wrong"] += len(qstate["queue"])
                    qstate["queue"] = []
                    st.session_state.quiz = qstate
                    
                    right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
                    pct = (right/total_q)*100 if total_q else 0.0
                    status = "Pass" if pct >= float(criteria) else "Fail"
                    ok, msg = append_result(
                        qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
                    )
                    st.session_state["submitted"] = True
                    st.session_state["submit_result"] = (ok, msg, right, total_q, pct, criteria, status)
                    st.query_params.clear()
                    st.rerun()

            if remaining <= 300:
                st.warning("🚨 URGENT: Less than 5 minutes remaining!")
            elif remaining <= 900:
                st.warning("⚠️ WARNING: Less than 15 minutes remaining!")
            elif remaining <= 1800:
                st.info("⏰ NOTICE: Less than 30 minutes remaining!")

        elif total_secs > 0 and "submitted" in st.session_state:
            rem_h = remaining // 3600
            rem_m = (remaining % 3600) // 60
            rem_s = remaining % 60

            if remaining <= 300:
                bg_color = "#DC2626"
                text_color = "white"
                icon = "🚨"
                pulse_class = "timer-pulse"
            elif remaining <= 900:
                bg_color = "#DC2626"
                text_color = "white"
                icon = "⚠️"
                pulse_class = ""
            elif remaining <= 1800:
                bg_color = "#D97706"
                text_color = "white"
                icon = "⏰"
                pulse_class = ""
            else:
                bg_color = "#1E3A8A"
                text_color = "white"
                icon = "⏰"
                pulse_class = ""

            progress_percent = (remaining / total_secs) * 100 if total_secs > 0 else 0

            stopped_timer_html = f"""
            <style>
            @keyframes pulse {{
                0% {{ transform: scale(1); opacity: 1; }}
                50% {{ transform: scale(1.05); opacity: 0.8; }}
                100% {{ transform: scale(1); opacity: 1; }}
            }}
            .timer-pulse {{
                animation: pulse 1s infinite;
            }}
            .timer-container {{
                padding: 20px;
                border-radius: 15px;
                text-align: center;
                font-size: 22px;
                font-weight: bold;
                margin-bottom: 20px;
                box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);
                border: 3px solid rgba(255, 255, 255, 0.1);
            }}
            </style>
            <div id="timer_container" class="timer-container {pulse_class}" style="background: linear-gradient(135deg, {bg_color}, {bg_color}CC); color: {text_color};">
                <div style="display: flex; align-items: center; justify-content: center; gap: 15px;">
                    <span id="timer_icon" style="font-size: 28px;">{icon}</span>
                    <span>Test Submitted - Time Remaining :</span>
                    <span id="timer_display" style="font-family: 'Courier New', monospace; font-size: 28px; background: rgba(0,0,0,0.2); padding: 5px 15px; border-radius: 8px;">
                        {rem_h:02d}:{rem_m:02d}:{rem_s:02d}
                    </span>
                </div>
                <div style="width: 100%; height: 6px; background-color: rgba(255,255,255,0.3); border-radius: 3px; overflow: hidden; margin-top: 15px;">
                    <div id="progress_bar" style="height: 100%; background: linear-gradient(90deg, #10B981, #34D399); width: {progress_percent:.1f}%; border-radius: 3px;"></div>
                </div>
            </div>
            """
            components.html(stopped_timer_html, height=150)

        answered_count = qstate["total"] - len(qstate["queue"])

        st.markdown(
            f"""
            <div style="padding: 12px 15px; border-radius: 8px; background: linear-gradient(135deg, #1E3A8A, #3B82F6); color: white; text-align: center; font-size: 17px; margin-bottom: 20px; white-space: nowrap; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <b>ID :</b> {qstate['emp_id']} &nbsp;•&nbsp; <b>Name :</b> {qstate['emp_name']} &nbsp;•&nbsp; <b>Standard :</b> {qstate['standard']} &nbsp;•&nbsp; <b>Progress :</b> {answered_count}/{qstate['total']}
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



        if len(qstate["queue"]) == 0 and "submitted" not in st.session_state:
            right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
            pct = (right/total_q)*100 if total_q else 0.0
            status = "Pass" if pct >= float(criteria) else "Fail"

            st.success("All questions attempted. You can now submit your test.")

            submit_clicked = st.button("Submit", use_container_width=True)
            if submit_clicked:
                ok, msg = append_result(
                    qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
                )
                st.session_state["submitted"] = True
                st.session_state["submit_result"] = (ok, msg, right, total_q, pct, criteria, status)
                st.rerun()

        if "submitted" in st.session_state:
            if "submit_result" in st.session_state:
                ok, msg, right, total_q, pct, criteria, status = st.session_state["submit_result"]
                if not ok:
                    st.error(f"Failed to save results to Google Sheets: {msg}")

                color = "#043006" if status == "Pass" else "#DC2626"
                st.markdown(
                    f"""
                    <div style="padding:20px; border-radius:12px; background: linear-gradient(135deg, #3B82F6, #2563EB, #1E3A8A); color:white; text-align:center; margin-top:20px;">
                        <h3 style="color:{color}; font-weight:700;">Final Result : <span style="font-weight:700;">{status}</span></h3>
                        <p style="font-size:18px;"><b>Score :</b> {right}/{total_q}<br><b>Percentage :</b> {pct:.2f}%<br><b>Passing Criteria :</b> {criteria:.0f}%</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
