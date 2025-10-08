import streamlit as st
import pandas as pd
import numpy as np
import datetime
import time
import os
import gspread
from google.oauth2.service_account import Credentials
import streamlit.components.v1 as components
import pytz
import io
import requests
import zipfile
import traceback
import fitz  # PyMuPDF
import re
from dateutil.relativedelta import relativedelta

# =====================
# Paths / Files
# =====================
BASE_DIR = os.path.dirname(__file__)
DB_FOLDER = os.path.join(BASE_DIR, "db")
QUESTIONS_FOLDER = os.path.join(DB_FOLDER, "Questions")

# =====================
# Google Sheets Setup
# =====================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
client = gspread.authorize(creds)
GSHEET_URL = st.secrets["connections"]["gsheets"]["spreadsheet"]

# =====================
# Helpers: Date parsing & font fit for PDF
# =====================
def _to_dt_general(v):
    """
    Robustly parse a variety of date strings from your sheet (e.g., '17-07-2025 03:45 PM', '17-07-2025', etc.)
    Returns a Python datetime, or NaT if it cannot parse.
    """
    if pd.isna(v):
        return pd.NaT
    s = str(v).strip()
    # try common formats quickly
    fmts = [
        "%d-%m-%Y %I:%M:%S %p",
        "%d-%m-%Y %I:%M %p",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ]
    for f in fmts:
        try:
            return datetime.datetime.strptime(s, f)
        except Exception:
            pass
    # fallback to pandas
    try:
        return pd.to_datetime(s, dayfirst=True).to_pydatetime()
    except Exception:
        return pd.NaT

# --- font-fit helpers (used by table writing) ---
def _fit_fontsize(page, text, max_width, fontname="helv", start_size=18, min_size=8):
    fs = start_size
    while fs >= min_size:
        width = fitz.get_text_length(str(text), fontname=fontname, fontsize=fs)
        if width <= max_width:
            return fs
        fs -= 1
    return min_size

def _draw_cell(page, x0, y0, width, height, text, fontname="helv", start_size=16, align=fitz.TEXT_ALIGN_LEFT):
    fs = _fit_fontsize(page, text, width - 2, fontname=fontname, start_size=start_size)
    rect = fitz.Rect(x0+1, y0+1, x0 + width - 1, y0 + height - 1)
    page.insert_textbox(rect, str(text), fontname=fontname, fontsize=fs, color=(0,0,0), align=align)

# =====================
# Cached Loaders
# =====================
@st.cache_data
def load_employees_and_standards():
    try:
        sheet = client.open_by_url(GSHEET_URL)
        # Employees
        try:
            employees_data = sheet.worksheet("Emloyees Data").get_all_records()
            employees = pd.DataFrame(employees_data)
            if employees.empty or not any(col.lower() in ["id", "name"] for col in employees.columns):
                employees = pd.DataFrame(columns=["ID", "Name"])
            else:
                id_col = next((col for col in employees.columns if "id" in col.lower()), "ID")
                name_col = next((col for col in employees.columns if "name" in col.lower()), "Name")
                employees = employees[[id_col, name_col]].rename(columns={id_col: "ID", name_col: "Name"})
        except Exception as e:
            st.warning(f"Error loading employees: {str(e)}")
            employees = pd.DataFrame(columns=["ID", "Name"])
        # Info / Standards
        try:
            standards_data = sheet.worksheet("Info").get_all_records()
            standards = pd.DataFrame(standards_data)
            if standards.empty:
                st.warning("Info sheet is empty. Using default standards.")
                standards = pd.DataFrame(columns=["ID", "Standard", "Total Questions", "Passing Criteria", "Hours", "Minutes", "Seconds"])
            else:
                required_cols = ["ID", "Standard", "Total Questions", "Passing Criteria", "Hours", "Minutes", "Seconds"]
                for col in required_cols:
                    if col not in standards.columns:
                        standards[col] = ""
                standards = standards[required_cols]
                standards["Standard"] = standards["Standard"].astype(str).str.strip()
        except Exception as e:
            st.error(f"Error loading Info sheet: {str(e)}")
            standards = pd.DataFrame(columns=["ID", "Standard", "Total Questions", "Passing Criteria", "Hours", "Minutes", "Seconds"])
        return employees, standards
    except Exception as e:
        st.error(f"Error in load_employees_and_standards: {str(e)}")
        return pd.DataFrame(columns=["ID", "Name"]), pd.DataFrame(columns=["ID", "Standard", "Total Questions", "Passing Criteria", "Hours", "Minutes", "Seconds"])

@st.cache_data
def load_all_results():
    try:
        sheet = client.open_by_url(GSHEET_URL)
        worksheet_names = ["Result 2", "Result2", "Result", "Results"]
        worksheet = None
        for name in worksheet_names:
            try:
                worksheet = sheet.worksheet(name)
                break
            except Exception:
                continue
        if worksheet is None:
            st.error("Could not find any results worksheet.")
            return pd.DataFrame(columns=["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Date / Time"])
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            return pd.DataFrame(columns=["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Date / Time"])
        headers = all_values[0]
        data_rows = all_values[1:]
        df = pd.DataFrame(data_rows, columns=headers)
        df['_original_order'] = range(len(df))
        df = df[~df.apply(lambda x: all(str(val).strip() == '' for val in x[:-1]), axis=1)]
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
            'Date / Time': ['DATE', 'Date', 'date', 'Timestamp', 'timestamp', 'Time', 'Date / Time']
        }
        for standard_name, possible_names in column_mapping.items():
            for col in df.columns:
                if col in possible_names and col != '_original_order':
                    df = df.rename(columns={col: standard_name})
                    break
        required_columns = ["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Date / Time"]
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""
        # numeric casts
        for col in ["Total", "Right", "Wrong"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        if "Percentage" in df.columns:
            df["Percentage"] = df["Percentage"].astype(str).str.replace("%", "", regex=False).str.replace(" ", "", regex=False)
            df["Percentage"] = pd.to_numeric(df["Percentage"], errors='coerce').fillna(0).astype(float)
        df = df.sort_values('_original_order').drop('_original_order', axis=1).reset_index(drop=True)
        return df[required_columns]
    except Exception as e:
        st.error(f"Error loading results: {str(e)}")
        st.error(f"Detailed error: {traceback.format_exc()}")
        return pd.DataFrame(columns=["ID", "Name", "Total", "Right", "Wrong", "Percentage", "Criteria", "Status", "Test Type", "Date / Time"])

@st.cache_data
def load_questions():
    try:
        sheet = client.open_by_url(GSHEET_URL)
        question_worksheet_names = ["Questions", "Question Bank", "Quiz Questions", "QuestionData"]
        questions_data = None
        for name in question_worksheet_names:
            try:
                worksheet = sheet.worksheet(name)
                questions_data = worksheet.get_all_records()
                break
            except Exception as ws_error:
                st.warning(f"Worksheet '{name}' not found or inaccessible: {str(ws_error)}")
                continue
        if questions_data is None or not questions_data:
            raise Exception("No valid questions worksheet found.")
        questions = pd.DataFrame(questions_data)
        if questions.empty:
            raise Exception("Questions worksheet is empty.")
        required_columns = ["Qno", "Standard", "Question", "A", "B", "C", "D", "Answer"]
        for col in required_columns:
            if col not in questions.columns:
                questions[col] = ""
        for c in ["Standard","Question","A","B","C","D","Answer"]:
            questions[c] = questions[c].astype(str).str.strip()
        return questions[required_columns]
    except Exception as e:
        st.error(f"Error loading questions from Google Sheet: {str(e)}")
        st.info("Generating sample questions for testing...")
        sample_questions = pd.DataFrame({
            "Qno": [1, 2, 3, 4, 5],
            "Standard": ["Basic", "Basic", "Advanced", "Advanced", "Cummulative"],
            "Question": [
                "What is 2 + 2?",
                "Capital of France?",
                "What is Python?",
                "Boiling point of water?",
                "Who wrote Romeo and Juliet?"
            ],
            "A": ["3", "Berlin", "A language", "50°C", "Dickens"],
            "B": ["4", "Paris", "A snake", "100°C", "Shakespeare"],
            "C": ["5", "London", "A fruit", "0°C", "Twain"],
            "D": ["6", "Madrid", "A bird", "212°F", "Hemingway"],
            "Answer": ["B", "B", "A", "B", "B"]
        })
        st.warning("Using sample questions. Add a 'Questions' worksheet to your Google Sheet for real data.")
        return sample_questions

def get_info_for_standard(standards, selected_standard):
    try:
        if selected_standard == "Cummulative":
            return 50, 80, 0, 50, 0
        row = standards[standards["Standard"].str.strip().str.upper() == str(selected_standard).strip().upper()]
        if not row.empty:
            total = int(row.iloc[0].get("Total Questions", 50))
            criteria = int(row.iloc[0].get("Passing Criteria", 70))
            h = int(row.iloc[0].get("Hours", 1))
            m = int(row.iloc[0].get("Minutes", 0))
            s = int(row.iloc[0].get("Seconds", 0))
            return total, criteria, h, m, s
        else:
            st.warning(f"No info found for standard: {selected_standard}. Using defaults.")
            return 50, 70, 1, 0, 0
    except Exception as e:
        st.error(f"Error getting standard info: {str(e)}")
        return 50, 70, 1, 0, 0

# =====================
# Certificate: template path
# =====================
def get_template_path(template_type):
    template_path = os.path.join(DB_FOLDER, f"{template_type}.pdf") if not template_type.endswith("_template") else os.path.join(DB_FOLDER, f"{template_type}.pdf")
    # allow names with _template already passed; else accept bare names like "PT"
    if not os.path.exists(template_path):
        # try simple name -> "<name>_template.pdf"
        guess = os.path.join(DB_FOLDER, f"{template_type}_template.pdf")
        if os.path.exists(guess):
            template_path = guess
    if os.path.exists(template_path):
        return template_path
    # fallback: GitHub raw (user must set correct repo or put files into /db)
    github_url = f"https://raw.githubusercontent.com/your_username/your_repo_name/main/db/{template_type}.pdf"
    try:
        response = requests.get(github_url, timeout=10)
        if response.status_code == 200:
            temp_path = f"/tmp/{template_type}.pdf"
            with open(temp_path, "wb") as file:
                file.write(response.content)
            return temp_path
        else:
            raise Exception(f"HTTP {response.status_code}")
    except Exception as e:
        st.error(f"Failed to get template '{template_type}': {str(e)}. Please add the pdf into db/ directory.")
        return None

# =====================
# Certificate Generation (tight redactions + overlay table)
# =====================
def generate_certificate(emp_id, emp_name, test_date, status, template_type,
                         table_rows: dict | None = None,
                         has_validity: bool = False):
    template_path = get_template_path(template_type)
    if not template_path:
        st.error(f"No {template_type} template available. Cannot generate certificate.")
        return None, None
    try:
        doc = fitz.open(template_path)
        page = doc[0]

        # date
        dt = _to_dt_general(test_date)
        if pd.isna(dt):
            st.error(f"Invalid date: {test_date}")
            doc.close()
            return None, None
        if not isinstance(dt, datetime.datetime):
            dt = pd.to_datetime(dt).to_pydatetime()
        date_str_disp = dt.strftime("%d-%B-%Y")

        # fonts
        corsiva_font = "times-italic"
        arial_font   = "helv"
        try:
            corsiva_fontfile = os.path.join(DB_FOLDER, "monotype_corsiva.ttf")
            if os.path.exists(corsiva_fontfile):
                f = fitz.Font(fontfile=corsiva_fontfile)
                if f.valid:
                    doc.insert_font(fontname="MonotypeCorsiva", fontfile=corsiva_fontfile)
                    corsiva_font = "MonotypeCorsiva"
        except Exception:
            pass
        try:
            arial_fontfile = os.path.join(DB_FOLDER, "arial.ttf")
            if os.path.exists(arial_fontfile):
                f = fitz.Font(fontfile=arial_fontfile)
                if f.valid:
                    doc.insert_font(fontname="Arial", fontfile=arial_fontfile)
                    arial_font = "Arial"
        except Exception:
            pass

        # certificate type + number
        tmap = {
            "MT_template":"MT","PT_template":"PT","UT_template":"UT","VT_template":"VT",
            "MT":"MT","PT":"PT","UT":"UT","VT":"VT",
            "DS-1_template":"DS-1","Cumulative_template":"Cumulative",
            "API RP 7G-2_template":"API RP 7G-2","API SPEC 5CT & 5A5_template":"API SPEC 5CT & 5A5",
            "DS-1":"DS-1","Cumulative":"Cumulative","API RP 7G-2":"API RP 7G-2","API SPEC 5CT & 5A5":"API SPEC 5CT & 5A5",
        }
        cert_type = tmap.get(template_type, template_type)
        cert_no = f"{emp_id}/PTIS/{cert_type}/{dt.year}"
        status_text = "Pass" if str(status).strip().lower()=="pass" else "Fail"

        # ---- tight replace helper ----
        def _tight_replace(find_text, new_text, fontname, fontsize, align=fitz.TEXT_ALIGN_LEFT):
            hits = page.search_for(find_text)
            if not hits:
                return False
            r = hits[0]
            pad = 1.5
            tight = fitz.Rect(r.x0, r.y0 - pad, r.x1, r.y1 + pad)
            page.add_redact_annot(tight, text=str(new_text), fontname=fontname, fontsize=fontsize,
                                  align=align, text_color=(0,0,0), fill=(1,1,1))
            return True

        # Name (placeholder or overlay)
        name_hit = page.search_for("Usman Waheed")
        if name_hit:
            _tight_replace("Usman Waheed", emp_name, corsiva_font, 40, align=fitz.TEXT_ALIGN_CENTER)
        else:
            _draw_cell(page, 120, 330, 380, 40, emp_name, fontname=corsiva_font, start_size=40, align=fitz.TEXT_ALIGN_CENTER)

        # Cert No (smaller font = 16) to the right of label
        for lbl in ("CERTIFICATE NO:", "Certificate No:", "CERTIFICATE NO"):
            lab = page.search_for(lbl)
            if lab:
                L = lab[0]
                w = 360
                box = fitz.Rect(L.x1 + 6, L.y0 - 1.5, L.x1 + 6 + w, L.y1 + 1.5)
                page.add_redact_annot(box, text=str(cert_no), fontname=arial_font, fontsize=16,
                                      align=fitz.TEXT_ALIGN_LEFT, text_color=(0,0,0), fill=(1,1,1))
                break

        # Date (smaller font = 16) to the right of label
        for lbl in ("DATE:", "Date:"):
            lab = page.search_for(lbl)
            if lab:
                L = lab[0]
                w = 260
                box = fitz.Rect(L.x1 + 6, L.y0 - 1.5, L.x1 + 6 + w, L.y1 + 1.5)
                page.add_redact_annot(box, text=str(date_str_disp), fontname=arial_font, fontsize=16,
                                      align=fitz.TEXT_ALIGN_LEFT, text_color=(0,0,0), fill=(1,1,1))
                break

        # Status if present
        for pat in ('Status: Pass','Status: Fail','Status:','Pass','Fail'):
            hits = page.search_for(pat)
            if hits:
                r = hits[0]
                page.add_redact_annot(r, text=f"Status: {status_text}", fontname=arial_font, fontsize=18,
                                      align=fitz.TEXT_ALIGN_LEFT, text_color=(0,0,0), fill=(1,1,1))
                break

        # Validity (rare for these templates)
        if has_validity:
            vhit = page.search_for("Validity:")
            if vhit:
                v = vhit[0]
                validity_date = (dt + relativedelta(years=5)).strftime('%d-%B-%Y')
                box = fitz.Rect(v.x1 + 6, v.y0 - 1.5, v.x1 + 6 + 260, v.y1 + 1.5)
                page.add_redact_annot(box, text=f"Validity: {validity_date}", fontname=arial_font, fontsize=21,
                                      align=fitz.TEXT_ALIGN_LEFT, text_color=(0,0,0), fill=(1,1,1))

        # Burn tight redactions only
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # ---- Table overlay (no redactions) ----
        if table_rows:
            hs = page.search_for("Standard") or page.search_for("STANDARD") or page.search_for("Standard:")
            hp = page.search_for("Percentage") or page.search_for("PERCENTAGE") or page.search_for("Percentage:")
            hc = page.search_for("Criteria") or page.search_for("CRITERIA") or page.search_for("Criteria:")
            h_std = hs[0] if hs else None
            h_pct = hp[0] if hp else None
            h_crt = hc[0] if hc else None
            if h_std and h_pct and h_crt:
                col_gap_y  = 6
                row_height = 20
                max_rows   = 12
                x_std = h_std.x0 + 4;  w_std = max(140, h_std.width)
                x_pct = h_pct.x0 + 4;  w_pct = max(80,  h_pct.width)
                x_crt = h_crt.x0 + 4;  w_crt = max(80,  h_crt.width)
                base_y0 = h_std.y1 + col_gap_y
                preferred = ["DS-1","Cumulative","API SPEC 5CT & 5A5","API RP 7G-2"]
                rest = [k for k in table_rows.keys() if k not in preferred]
                ordered = [k for k in preferred if k in table_rows] + rest
                rcount = 0
                for label in ordered:
                    if rcount >= max_rows: break
                    vals = table_rows.get(label)
                    if not vals: continue
                    y0 = base_y0 + rcount * row_height
                    _draw_cell(page, x_std, y0, w_std, row_height, label, fontname=arial_font, start_size=16, align=fitz.TEXT_ALIGN_LEFT)
                    pct = float(vals.get("Percentage", 0.0))
                    _draw_cell(page, x_pct, y0, w_pct, row_height, f"{pct:.2f}%", fontname=arial_font, start_size=16, align=fitz.TEXT_ALIGN_CENTER)
                    crit_raw = vals.get("Criteria", "")
                    crit = float(str(crit_raw).replace("%","").strip() or 0)
                    _draw_cell(page, x_crt, y0, w_crt, row_height, f"{crit:.0f}%", fontname=arial_font, start_size=16, align=fitz.TEXT_ALIGN_CENTER)
                    rcount += 1
            else:
                st.warning("Table headers not found in template; skipped table fill.")

        # save
        safe_name = "".join(c for c in str(emp_name) if c.isalnum() or c in (" ","-","_")).rstrip()
        out_name = f"{template_type}_Certificate_{emp_id}_{safe_name}_{dt.strftime('%d-%m-%Y')}.pdf"
        out_path = f"/tmp/{out_name}"
        doc.save(out_path, garbage=3, deflate=True)
        doc.close()
        st.success(f"Generated certificate: {out_name}")
        return out_path, out_name
    except Exception as e:
        try:
            doc.close()
        except Exception:
            pass
        st.error(f"Error generating {template_type} certificate: {e}")
        return None, None

# =====================
# Individual Test Report (unchanged)
# =====================
def create_individual_test_report(emp_id, emp_name, test_date, test_type, total, right, wrong, pct, criteria, status):
    report_data = {
        'Test Information': [
            ['Employee ID', emp_id],
            ['Employee Name', emp_name],
            ['Test Date & Time', test_date],
            ['Test Type/Standard', test_type],
            ['Total Questions', total],
            ['Correct Answers', right],
            ['Wrong Answers', wrong],
            ['Final Score', f"{right - (wrong * 0.25):.2f}/{total}"],
            ['Percentage', f"{pct:.2f}%"],
            ['Passing Criteria', f"{criteria}%"],
            ['Status', status]
        ]
    }
    return pd.DataFrame(report_data['Test Information'], columns=['Field', 'Value'])

def download_individual_test(emp_id, emp_name, test_data):
    report_df = create_individual_test_report(
        emp_id, emp_name, test_data['Date / Time'], test_data['Test Type'],
        test_data['Total'], test_data['Right'], test_data['Wrong'],
        test_data['Percentage'], test_data['Criteria'], test_data['Status']
    )
    csv_buffer = io.StringIO()
    report_df.to_csv(csv_buffer, index=False)
    csv_data = csv_buffer.getvalue()
    timestamp = test_data['Date / Time'].replace('/', '_').replace(' ', '_').replace(':', '-')
    safe_name = "".join(c for c in emp_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
    filename = f"Test_Report_{emp_id}_{safe_name}_{test_data['Test Type']}_{timestamp}.csv"
    return csv_data, filename

# =====================
# Quiz helpers
# =====================
def start_quiz_session(emp_id, emp_name, standard, questions_df, total):
    if standard == "Cummulative":
        cand = questions_df.copy()
    else:
        cand = questions_df[questions_df["Standard"].astype(str).str.strip().str.upper() == str(standard).strip().upper()]
    cand = cand.dropna(subset=["Question", "A", "B", "C", "D", "Answer"])
    if total <= 0 or cand.empty:
        return False, f"Questions not defined for standard: {standard}."
    if len(cand) < total:
        total = len(cand)
    sampled = cand.sample(n=min(total, len(cand)), random_state=int(time.time())).reset_index(drop=True)
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
        "attempted": set(),
        "skipped_questions": set(),
    }
    return True, ""

def format_timer(h, m, s):
    try:
        return int(h) * 3600 + int(m) * 60 + int(s)
    except Exception:
        return 0

def append_result(emp_id, emp_name, total, right, wrong, criteria_pct, status, test_type):
    try:
        sheet = client.open_by_url(GSHEET_URL)
        worksheet_names = ["Result 2", "Result2", "Result", "Results"]
        worksheet = None
        for name in worksheet_names:
            try:
                worksheet = sheet.worksheet(name)
                st.info(f"Saving results to worksheet: '{name}'")
                break
            except Exception:
                continue
        if worksheet is None:
            try:
                for ws in sheet.worksheets():
                    if "result" in ws.title.lower():
                        worksheet = ws
                        st.info(f"Saving results to worksheet: '{ws.title}'")
                        break
            except:
                pass
        if worksheet is None:
            return False, "Could not find results worksheet to save data"

        pkt_tz = pytz.timezone('Asia/Karachi')
        now = datetime.datetime.now(pkt_tz).strftime("%d-%m-%Y %I:%M:%S %p")
        raw_score = right - (wrong * 0.25)
        final_score = max(0, raw_score)
        pct = (final_score / total) * 100 if total else 0.0

        try:
            headers = worksheet.row_values(1)
        except:
            headers = []
        if headers:
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
                'DATE': now,
                'DATE / TIME': now
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
                elif 'DATE' in header_upper or 'TIME' in header_upper or 'TIMESTAMP' in header_upper:
                    new_row.append(now)
                else:
                    new_row.append('')
        else:
            new_row = [
                str(emp_id), str(emp_name), int(total), int(right), int(wrong),
                f"{pct:.2f}%", f"{criteria_pct:.0f}%", str(status), str(test_type), now
            ]
        worksheet.append_row(new_row)
        st.success("Results saved to Google Sheet.")
        return True, ""
    except Exception as e:
        st.error(f"Error saving results: {str(e)}")
        return False, str(e)

# =====================
# Result -> Certificate table utilities
# =====================
def normalize_test_type(s):
    return str(s).strip()

def has_core4(emp_df, relaxed=True):
    needed = {"DS-1","Cumulative","API SPEC 5CT & 5A5","API RP 7G-2"}
    got = set(normalize_test_type(x) for x in emp_df.loc[emp_df["Status"]=="Pass","Test Type"])
    # relaxed: allow case/spaces variations already normalized above
    return needed.issubset(got)

def has_ndt_requirements(emp_df, which, relaxed=True):
    tt = emp_df[emp_df["Status"]=="Pass"]["Test Type"].map(normalize_test_type)
    s = set(tt)
    if which == "MT":
        return "MPT (General)" in s and "MPT (Specific)" in s
    if which == "PT":
        return "Penetrant Testing (General)" in s and "Penetrant Testing (Specific)" in s
    if which == "UT":
        return "Ultrasonic" in s
    if which == "VT":
        return "Visual Testing" in s
    return False

def get_latest_scores_for(emp_df, labels):
    """
    Build {label: {'Percentage': float, 'Criteria': float, 'Date': str}} using the
    latest row per label for the employee. Only includes labels present.
    """
    out = {}
    # ensure date sorting works
    tmp = emp_df.copy()
    tmp["__dt"] = tmp["Date / Time"].apply(_to_dt_general)
    tmp = tmp.sort_values("__dt")  # ascending by time
    for label in labels:
        dfl = tmp[tmp["Test Type"].astype(str).str.strip() == label]
        if dfl.empty:
            continue
        last = dfl.iloc[-1]
        crit = str(last["Criteria"])
        try:
            crit_f = float(str(crit).replace("%","").strip())
        except:
            crit_f = 0.0
        out[label] = {
            "Percentage": float(last["Percentage"]),
            "Criteria": crit_f,
            "Date": last["Date / Time"],
        }
    return out

# =====================
# UI
# =====================
st.set_page_config(page_title="PTIS Online Testing Module", page_icon="📝", layout="centered")
st.title("PTIS Online Testing Module")

employees, standards = load_employees_and_standards()
questions = load_questions()

if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False
if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0
if "filter_reset_counter" not in st.session_state:
    st.session_state.filter_reset_counter = 0

# Admin login
if not st.session_state.admin_logged_in and "quiz" not in st.session_state:
    st.subheader("Admin Login")
    username = st.text_input("Username", key="admin_username")
    password = st.text_input("Password", type="password", key="admin_password")
    if st.button("Login", key="admin_login_btn"):
        if username == "admin" and password == "AdminPtis-3692":
            st.session_state.admin_logged_in = True
            st.success("Admin login successful!")
            st.rerun()
        else:
            st.error("Invalid username or password")

# Admin dashboard
if st.session_state.admin_logged_in:
    st.subheader("Admin Dashboard - Employee Results")
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()

    results_df = load_all_results()
    if not results_df.empty:
        st.markdown("---")
        st.subheader("🔍 Filters")

        id_name_mapping = dict(zip(results_df["ID"].astype(str), results_df["Name"]))
        name_id_mapping = dict(zip(results_df["Name"], results_df["ID"].astype(str)))

        id_key = f"emp_id_filter_{st.session_state.filter_reset_counter}"
        name_key = f"emp_name_filter_{st.session_state.filter_reset_counter}"
        if id_key not in st.session_state: st.session_state[id_key] = "All"
        if name_key not in st.session_state: st.session_state[name_key] = "All"

        def sync_id_to_name():
            sid = st.session_state[id_key]
            if sid != "All" and sid in id_name_mapping:
                st.session_state[name_key] = id_name_mapping[sid]
            elif sid == "All":
                st.session_state[name_key] = "All"

        def sync_name_to_id():
            sname = st.session_state[name_key]
            if sname != "All" and sname in name_id_mapping:
                st.session_state[id_key] = name_id_mapping[sname]
            elif sname == "All":
                st.session_state[id_key] = "All"

        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        with filter_col1:
            employee_ids = ["All"] + sorted(results_df["ID"].astype(str).unique().tolist())
            selected_emp_id = st.selectbox("Filter by Employee ID", employee_ids,
                                           index=employee_ids.index(st.session_state[id_key]) if st.session_state[id_key] in employee_ids else 0,
                                           key=id_key, on_change=sync_id_to_name)
        with filter_col2:
            employee_names = ["All"] + sorted(results_df["Name"].unique().tolist())
            selected_emp_name = st.selectbox("Filter by Employee Name", employee_names,
                                             index=employee_names.index(st.session_state[name_key]) if st.session_state[name_key] in employee_names else 0,
                                             key=name_key, on_change=sync_name_to_id)
        with filter_col3:
            statuses = ["All"] + sorted(results_df["Status"].unique().tolist())
            selected_status = st.selectbox("Filter by Status", statuses, index=0, key=f"status_filter_{st.session_state.filter_reset_counter}")
        with filter_col4:
            test_types = ["All"] + sorted(results_df["Test Type"].unique().tolist())
            selected_test_type = st.selectbox("Filter by Test Type", test_types, index=0, key=f"test_type_filter_{st.session_state.filter_reset_counter}")

        filter_col5, filter_col6, filter_col7, filter_col8 = st.columns(4)
        with filter_col5:
            st.write("")
            if st.button("🗑️ Clear All Filters"):
                st.session_state.filter_reset_counter += 1
                keys_to_remove = [key for key in st.session_state.keys() if key.startswith(('emp_id_filter_', 'emp_name_filter_', 'status_filter_', 'test_type_filter_', 'prev_emp_id_filter_', 'prev_emp_name_filter_'))]
                for key in keys_to_remove:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()

        filtered_df = results_df.copy()
        if selected_emp_id != "All":
            filtered_df = filtered_df[filtered_df["ID"].astype(str) == selected_emp_id]
        elif selected_emp_name != "All":
            filtered_df = filtered_df[filtered_df["Name"] == selected_emp_name]
        if selected_status != "All":
            filtered_df = filtered_df[filtered_df["Status"] == selected_status]
        if selected_test_type != "All":
            filtered_df = filtered_df[filtered_df["Test Type"] == selected_test_type]

        if selected_emp_id != "All" or selected_emp_name != "All":
            display_name = selected_emp_name if selected_emp_name != "All" else id_name_mapping.get(selected_emp_id, "Unknown")
            display_id = selected_emp_id if selected_emp_id != "All" else name_id_mapping.get(selected_emp_name, "Unknown")
            st.info(f"🔗 **Selected Employee**: ID: {display_id} | Name: {display_name}")

        st.markdown("---")
        st.subheader("📥 Individual Test Download")
        if selected_emp_id != "All" or selected_emp_name != "All":
            if selected_emp_id != "All":
                emp_filtered = filtered_df[filtered_df["ID"].astype(str) == selected_emp_id]
                emp_name_display = id_name_mapping.get(selected_emp_id, selected_emp_id)
                emp_id_display = selected_emp_id
            else:
                emp_filtered = filtered_df[filtered_df["Name"] == selected_emp_name]
                emp_name_display = selected_emp_name
                emp_id_display = name_id_mapping.get(selected_emp_name, "Unknown")

            if not emp_filtered.empty:
                st.info(f"Showing {len(emp_filtered)} test(s) for employee: **{emp_name_display}** (ID: {emp_id_display})")
                emp_filtered = emp_filtered.sort_values("Date / Time", ascending=False).reset_index(drop=True)
                for idx, test_row in emp_filtered.iterrows():
                    with st.expander(f"Test {idx+1}: {test_row['Test Type']} - {test_row['Date / Time']} ({test_row['Status']})", expanded=False):
                        col1, col2, col3 = st.columns([2, 1, 1])
                        with col1:
                            st.metric("Score", f"{test_row['Right']}/{test_row['Total']}")
                            st.metric("Percentage", f"{test_row['Percentage']:.1f}%")
                        with col2:
                            st.metric("Status", test_row['Status'])
                        with col3:
                            csv_data, filename = download_individual_test(test_row['ID'], test_row['Name'], test_row)
                            st.download_button(label=f"📄 Download Test Report", data=csv_data, file_name=filename, mime="text/csv", use_container_width=True)
                        st.write("**Test Details:**")
                        st.json({
                            "Employee ID": test_row['ID'],
                            "Employee Name": test_row['Name'],
                            "Standard": test_row['Test Type'],
                            "Total Questions": test_row['Total'],
                            "Correct": test_row['Right'],
                            "Wrong": test_row['Wrong'],
                            "Passing Criteria": f"{test_row['Criteria']}%",
                            "Completed": test_row['Date / Time']
                        })
            else:
                st.warning("No test results found for the selected employee.")
        else:
            st.info("👆 **Select an Employee ID or Name** to view and download individual test reports")

        st.markdown("---")
        st.subheader("📊 Test Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Tests", len(filtered_df))
        with col2:
            pass_count = len(filtered_df[filtered_df["Status"] == "Pass"]) if "Status" in filtered_df.columns else 0
            st.metric("Passed", pass_count)
        with col3:
            fail_count = len(filtered_df[filtered_df["Status"] == "Fail"]) if "Status" in filtered_df.columns else 0
            st.metric("Failed", fail_count)
        with col4:
            if "Percentage" in filtered_df.columns and len(filtered_df) > 0:
                avg_score = filtered_df["Percentage"].mean()
                st.metric("Avg Score", f"{avg_score:.1f}%")
            else:
                st.metric("Avg Score", "N/A")

        if len(filtered_df) != len(results_df):
            st.info(f"Showing {len(filtered_df)} of {len(results_df)} total records")

        st.markdown("---")
        if not filtered_df.empty:
            display_df = filtered_df.copy()
            display_df.insert(0, 'S.No.', range(1, len(display_df) + 1))
            export_col1, export_col2, export_col3 = st.columns([1, 1, 2])
            with export_col1:
                csv = display_df.to_csv(index=False)
                st.download_button(label="📄 Download CSV", data=csv, file_name=f"all_test_results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", mime="text/csv")
            with export_col2:
                if st.button("⚙️ Column Settings"):
                    st.session_state.show_column_settings = not st.session_state.get("show_column_settings", False)
            if st.session_state.get("show_column_settings", False):
                st.subheader("Column Visibility")
                cols_to_show = []
                col_settings = st.columns(5)
                for i, col in enumerate(filtered_df.columns):
                    with col_settings[i % 5]:
                        if st.checkbox(col, value=True, key=f"show_{col}"):
                            cols_to_show.append(col)
                filtered_df = filtered_df[cols_to_show] if cols_to_show else filtered_df
                display_df = filtered_df.copy()
                display_df.insert(0, 'S.No.', range(1, len(display_df) + 1))
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "S.No.": st.column_config.NumberColumn("S.No.", help="Serial Number", format="%d", width="small"),
                    "Percentage": st.column_config.ProgressColumn("Percentage", help="Test Score Percentage", format="%.1f%%", min_value=0, max_value=100),
                    "Status": st.column_config.TextColumn("Status", help="Pass/Fail Status"),
                    "Total": st.column_config.NumberColumn("Total Questions", help="Total number of questions in the test", format="%d"),
                    "Right": st.column_config.NumberColumn("Correct Answers", help="Number of correct answers", format="%d"),
                    "Wrong": st.column_config.NumberColumn("Wrong Answers", help="Number of wrong answers", format="%d"),
                    "Date / Time": st.column_config.TextColumn("Date / Time", help="Test completion date and time"),
                }
            )

        # =====================
        # Certificate Generation (updated)
        # =====================
        st.markdown("---")
        st.subheader("📜 Generate Certificates")

        passed_results = results_df[results_df["Status"] == "Pass"].copy()
        cert_employee_names = ["All"] + sorted(passed_results["Name"].unique().tolist())
        selected_cert_name = st.selectbox(
            "Filter Certificates by Employee Name",
            cert_employee_names, index=0,
            key=f"cert_name_filter_{st.session_state.filter_reset_counter}"
        )

        if st.button("Generate Certificates for Qualifying Employees"):
            to_process = []
            if selected_cert_name == "All":
                for name, grp in passed_results.groupby("Name"):
                    to_process.append((name, grp.copy()))
            else:
                grp = passed_results[passed_results["Name"] == selected_cert_name].copy()
                if grp.empty:
                    st.warning("No passed results for the selected employee.")
                else:
                    to_process.append((selected_cert_name, grp))

            certificate_files = []
            core_labels = ["DS-1", "Cumulative", "API SPEC 5CT & 5A5", "API RP 7G-2"]

            for emp_name, emp_df in to_process:
                emp_id = str(emp_df.iloc[0]["ID"])

                # ---- Core 4 (only if all 4 passed) ----
                if has_core4(emp_df, relaxed=True):
                    rows_core = get_latest_scores_for(emp_df, core_labels)
                    dates = [v["Date"] for v in rows_core.values()]
                    cert_date = max(dates, key=_to_dt_general) if dates else emp_df.iloc[0]["Date / Time"]
                    for core_template in [
                        "DS-1_template",
                        "Cumulative_template",
                        "API RP 7G-2_template",
                        "API SPEC 5CT & 5A5_template",
                    ]:
                        pth, fn = generate_certificate(
                            emp_id, emp_name, cert_date, status="Pass",
                            template_type=core_template,
                            table_rows=rows_core,    # << fills table
                            has_validity=False
                        )
                        if pth:
                            certificate_files.append((pth, fn))

                # ---- NDT individual templates (independent) ----
                ndt_templates = []
                if has_ndt_requirements(emp_df, "MT", relaxed=True): ndt_templates.append("MT_template")
                if has_ndt_requirements(emp_df, "PT", relaxed=True): ndt_templates.append("PT_template")
                if has_ndt_requirements(emp_df, "UT", relaxed=True): ndt_templates.append("UT_template")
                if has_ndt_requirements(emp_df, "VT", relaxed=True): ndt_templates.append("VT_template")

                if ndt_templates:
                    ndt_labels = core_labels + [
                        "MPT (General)", "MPT (Specific)",
                        "Penetrant Testing (General)", "Penetrant Testing (Specific)",
                        "Ultrasonic", "Visual Testing",
                    ]
                    rows_ndt = get_latest_scores_for(emp_df, ndt_labels)
                    dates_ndt = [v["Date"] for v in rows_ndt.values()]
                    cert_date_ndt = max(dates_ndt, key=_to_dt_general) if dates_ndt else emp_df.iloc[0]["Date / Time"]
                    for nt in ndt_templates:
                        pth, fn = generate_certificate(
                            emp_id, emp_name, cert_date_ndt, status="Pass",
                            template_type=nt,
                            table_rows=rows_ndt,     # << fills table
                            has_validity=False
                        )
                        if pth:
                            certificate_files.append((pth, fn))

            if certificate_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for cert_path, cert_filename in certificate_files:
                        zipf.write(cert_path, cert_filename)
                zip_buffer.seek(0)
                filename_suffix = selected_cert_name if selected_cert_name != "All" else "all_qualifying"
                st.download_button(
                    label=f"Download Certificates (ZIP) for {filename_suffix}",
                    data=zip_buffer,
                    file_name=f"certificates_{filename_suffix}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip"
                )
            else:
                st.info("No qualifying certificates to generate with the current filter and data.")

        st.markdown("---")
        if st.button("Logout"):
            st.session_state.admin_logged_in = False
            st.session_state.pop("quiz", None)
            st.rerun()
    else:
        st.info("No results available yet in the Result 2 sheet.")
        if st.button("Logout"):
            st.session_state.admin_logged_in = False
            st.session_state.pop("quiz", None)
            st.rerun()

# =====================
# Employee login and quiz
# =====================
if not st.session_state.admin_logged_in and "quiz" not in st.session_state:
    st.subheader("Employee Login")

    def auto_populate_name():
        emp_id_input = st.session_state[f"id_{st.session_state.reset_counter}"]
        if emp_id_input and not employees.empty:
            try:
                fetched = employees[employees["ID"].astype(str).str.strip() == str(emp_id_input).strip()]
                if not fetched.empty:
                    st.session_state[f"name_{st.session_state.reset_counter}"] = str(fetched.iloc[0]["Name"])
                else:
                    st.session_state[f"name_{st.session_state.reset_counter}"] = ""
            except Exception:
                st.session_state[f"name_{st.session_state.reset_counter}"] = ""

    col1, col2 = st.columns(2)
    with col1:
        emp_id = st.text_input("Employee ID", value="", key=f"id_{st.session_state.reset_counter}",
                               help="Enter your employee identification number and press Enter", on_change=auto_populate_name)
    name_key = f"name_{st.session_state.reset_counter}"
    if name_key not in st.session_state:
        st.session_state[name_key] = ""
    with col2:
        name = st.text_input("Name", key=name_key, help="This will auto-fill when you enter a valid Employee ID")

    options = standards["Standard"].dropna().unique().tolist()
    options = sorted(options)
    if "Cummulative" not in options:
        options = ["Cummulative"] + options
    selected_standard = st.selectbox("Select Standard", options, index=0 if options else None, key=f"std_{st.session_state.reset_counter}")
    total, criteria, h, m, s = get_info_for_standard(standards, selected_standard)
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("Total Questions", total)
    with c2: st.metric("Passing Criteria (%)", criteria)
    with c3: st.metric("Timer (HH:MM:SS)", f"{h:02d}:{m:02d}:{s:02d}")
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

# =====================
# Quiz interface
# =====================
elif "quiz" in st.session_state:
    qstate = st.session_state.quiz
    total, criteria, h, m, s = get_info_for_standard(standards, qstate["standard"])
    total_secs = format_timer(h, m, s)
    elapsed = int(time.time() - qstate["start_ts"])
    remaining = max(0, total_secs - elapsed)

    # (Timer & quiz UI — unchanged except for tight logic)
    # --- auto-submit on timeout ---
    if total_secs > 0 and len(qstate["queue"]) > 0 and "submitted" not in st.session_state:
        if remaining <= 0:
            st.error("Time is up! Auto-submitting your test...")
            for qid in qstate["queue"]:
                if qid not in qstate.get("attempted", set()):
                    qstate["wrong"] += 1
            qstate["queue"] = []
            st.session_state.quiz = qstate
            right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
            raw_score = right - (wrong * 0.25)
            final_score = max(0, raw_score)
            pct = (final_score/total_q)*100 if total_q else 0.0
            status = "Pass" if pct >= float(criteria) else "Fail"
            ok, msg = append_result(qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"])
            st.session_state["submitted"] = True
            st.session_state["submit_result"] = (ok, msg, right, wrong, total_q, pct, criteria, status, final_score)
            st.query_params.clear()
            st.rerun()

        # timer widget
        rem_h = remaining // 3600
        rem_m = (remaining % 3600) // 60
        rem_s = remaining % 60
        if remaining <= 300:
            bg_color = "#DC2626"; text_color = "white"; icon = "🚨"; pulse_class = "timer-pulse"
        elif remaining <= 900:
            bg_color = "#DC2626"; text_color = "white"; icon = "⚠️"; pulse_class = ""
        elif remaining <= 1200:
            bg_color = "#D97706"; text_color = "white"; icon = "⏰"; pulse_class = ""
        else:
            bg_color = "#1E3A8A"; text_color = "white"; icon = "⏰"; pulse_class = ""
        progress_percent = (remaining / total_secs) * 100 if total_secs > 0 else 0
        timer_html = f"""
        <style>
        @keyframes pulse {{0% {{ transform: scale(1); opacity: 1; }}50% {{ transform: scale(1.05); opacity: 0.8; }}100% {{ transform: scale(1); opacity: 1; }}}}
        .timer-pulse {{animation: pulse 1s infinite;}}
        .timer-container {{padding: 20px;border-radius: 15px;text-align: center;font-size: 22px;font-weight: bold;margin-bottom: 20px;box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2);border: 3px solid rgba(255, 255, 255, 0.1);}}
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
        """
        components.html(timer_html, height=150)

        if remaining <= 300:
            st.warning("🚨 URGENT: Less than 5 minutes remaining!")
        elif remaining <= 900:
            st.warning("⚠️ WARNING: Less than 15 minutes remaining!")
        elif remaining <= 1200:
            st.info("⏰ NOTICE: Less than 20 minutes remaining!")

    elif total_secs > 0 and "submitted" in st.session_state:
        pass  # show the stopped timer (omitted for brevity, same visuals)

    if "attempted" not in st.session_state.quiz:
        st.session_state.quiz["attempted"] = set()
    if "skipped_questions" not in st.session_state.quiz:
        st.session_state.quiz["skipped_questions"] = set()

    answered_count = qstate["total"] - len(qstate["queue"])
    st.markdown(
        f"""
        <div style="padding: 12px 15px; border-radius: 8px; background: linear-gradient(135deg, #1E3A8A, #3B82F6); color: white; text-align: center; font-size: 17px; margin-bottom: 20px; white-space: nowrap; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <b>ID :</b> {qstate['emp_id']} &nbsp;•&nbsp; <b>Name :</b> {qstate['emp_name']} &nbsp;•&nbsp; <b>Standard :</b> {qstate['standard']} &nbsp;•&nbsp; <b>Progress :</b> {answered_count}/{qstate['total']}
        </div>
        """,
        unsafe_allow_html=True
    )

    st.info("📌 **Scoring System**: +1 mark for correct answer, -0.25 marks for wrong answer, 0 marks for unattempted questions")

    if len(qstate["queue"]) > 0:
        current_qid = qstate["queue"][0]
        row = qstate["rows"].iloc[current_qid]
        qno, question, A, B, C, D, correct = row["Qno"], row["Question"], row["A"], row["B"], row["C"], row["D"], row["Answer"]
        is_previously_skipped = current_qid in qstate["skipped_questions"]
        st.subheader(f"Q{current_qid+1}. {question}")
        choice = st.radio("Choose your answer:", [A, B, C, D], index=None, key=f"q_{current_qid}")
        col1, col2 = st.columns([1,1])
        with col1:
            if st.button("Next", use_container_width=True):
                if choice is None:
                    st.warning("⚠️ Please select an option before moving on.")
                else:
                    qstate["attempted"].add(current_qid)
                    mapping = {"A": A, "B": B, "C": C, "D": D}
                    correct_text = mapping.get(str(correct).strip(), str(correct).strip())
                    is_correct = str(choice).strip() == str(correct_text).strip()
                    qstate["answers"][current_qid] = {"choice": choice, "correct": correct_text, "is_correct": is_correct}
                    if is_correct:
                        qstate["right"] += 1
                    else:
                        qstate["wrong"] += 1
                    qstate["queue"].pop(0)
                    st.session_state.quiz = qstate
                    st.rerun()
        with col2:
            if len(qstate["queue"]) > 1 and not is_previously_skipped:
                if st.button("Skip", use_container_width=True):
                    qstate["skipped_questions"].add(current_qid)
                    qstate["queue"].append(qstate["queue"].pop(0))
                    st.session_state.quiz = qstate
                    st.rerun()

    if len(qstate["queue"]) == 0 and "submitted" not in st.session_state:
        right, wrong, total_q = qstate["right"], qstate["wrong"], qstate["total"]
        raw_score = right - (wrong * 0.25)
        final_score = max(0, raw_score)
        pct = (final_score/total_q)*100 if total_q else 0.0
        status = "Pass" if pct >= float(criteria) else "Fail"
        st.success("All questions attempted. You can now submit your test.")
        submit_clicked = st.button("Submit", use_container_width=True)
        if submit_clicked:
            ok, msg = append_result(qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"])
            st.session_state["submitted"] = True
            st.session_state["submit_result"] = (ok, msg, right, wrong, total_q, pct, criteria, status, final_score)
            st.rerun()

    if "submitted" in st.session_state and "submit_result" in st.session_state:
        ok, msg, right, wrong, total_q, pct, criteria, status, final_score = st.session_state["submit_result"]
        if not ok:
            st.error(f"Failed to save results to Google Sheets: {msg}")
        color = "#10B981" if status == "Pass" else "#DC2626"
        st.markdown(
            f"""
            <div style="padding:20px; border-radius:12px; background: linear-gradient(135deg, #1E3A8A, #2563EB); color:white; text-align:center; margin-top:20px;">
                <h3 style="color:{color}; font-weight:700;">Final Result : <span style="font-weight:700;">{status}</span></h3>
                <p style="font-size:18px;">
                    <b>Correct Answers :</b> {right}<br>
                    <b>Wrong Answers :</b> {wrong}<br>
                    <b>Final Score :</b> {final_score:.2f}/{total_q}<br>
                    <b>Percentage :</b> {pct:.2f}%<br>
                    <b>Passing Criteria :</b> {criteria:.0f}%
                </p>
                <small style="opacity: 0.8;">Negative marking: -0.25 marks per wrong answer</small>
            </div>
            """,
            unsafe_allow_html=True
        )
