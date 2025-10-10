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
import fitz
import re

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
            if employees.empty or not any(col.lower() in ["id", "name"] for col in employees.columns):
                employees = pd.DataFrame(columns=["ID", "Name"])
            else:
                id_col = next((col for col in employees.columns if "id" in col.lower()), "ID")
                name_col = next((col for col in employees.columns if "name" in col.lower()), "Name")
                employees = employees[[id_col, name_col]].rename(columns={id_col: "ID", name_col: "Name"})
        except Exception as e:
            st.warning(f"Error loading employees: {str(e)}")
            employees = pd.DataFrame(columns=["ID", "Name"])
        
        # Load Standards from Info sheet
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
        employees = pd.DataFrame(columns=["ID", "Name"])
        standards = pd.DataFrame(columns=["ID", "Standard", "Total Questions", "Passing Criteria", "Hours", "Minutes", "Seconds"])
        return employees, standards

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
        
        for col in ["Total", "Right", "Wrong"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        if "Percentage" in df.columns:
            df["Percentage"] = df["Percentage"].astype(str).str.replace("%", "").str.replace(" ", "")
            df["Percentage"] = pd.to_numeric(df["Percentage"], errors='coerce').fillna(0).astype(float)
        
        df = df.sort_values('_original_order').drop('_original_order', axis=1)
        df = df.reset_index(drop=True)
        
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
        
        for col in ["Standard", "Question", "A", "B", "C", "D", "Answer"]:
            questions[col] = questions[col].astype(str).str.strip()
        
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
# Certificate Generation
# =====================

def get_template_path(template_type):
    """
    Returns path to template pdf in /db or downloads a fallback.
    Expected 'template_type' already includes the exact filename stem like:
      "Ds-1_template", "Cumulative_template", "API RP 7G-2_template", "API SPEC 5CT & 5A5_template"
    """
    template_path = os.path.join(DB_FOLDER, f"{template_type}.pdf")
    if os.path.exists(template_path):
        return template_path
    
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
        st.error(f"Failed to load template '{template_type}': {str(e)}. Please add 'db/{template_type}.pdf' to your repo.")
        return None

def generate_certificate(
    emp_id,
    emp_name,
    test_date,
    status,
    template_type,
    standard_text=None,
    percentage_text=None,
    criteria_text=None,
    skip_dates=True,
):
    """
    Fixed certificate generation with guaranteed rendering of name, cert no, and date
    """
    template_path = get_template_path(template_type)
    if not template_path:
        st.error(f"No {template_type} template available. Cannot generate certificate.")
        return None, None

    def _nice_date(dt_str):
        try:
            d = pd.to_datetime(str(dt_str), errors="coerce", dayfirst=True)
            if pd.isna(d):
                d = pd.to_datetime(str(dt_str), errors="coerce", utc=True)
            if pd.isna(d):
                return str(dt_str).split(" ")[0]
            return d.strftime("%d-%B-%Y")
        except Exception:
            return str(dt_str).split(" ")[0]

    try:
        doc = fitz.open(template_path)
        page = doc[0]
        pw, ph = page.rect.width, page.rect.height

        # ---------- Load fonts ----------
        arial_font = "helv"
        name_font = "times-italic"

        try:
            arial_fontfile = os.path.join(DB_FOLDER, "arial.ttf")
            if os.path.exists(arial_fontfile):
                doc.insert_font(fontname="Arial", fontfile=arial_fontfile)
                arial_font = "Arial"
        except:
            pass

        try:
            corsiva_fontfile = os.path.join(DB_FOLDER, "monotype_corsiva.ttf")
            if os.path.exists(corsiva_fontfile):
                doc.insert_font(fontname="MonotypeCorsiva", fontfile=corsiva_fontfile)
                name_font = "MonotypeCorsiva"
        except:
            pass

        # ---------- FIND NAME POSITION ----------
        award_texts = [
            "Certificate of Accomplishment Awarded to",
            "Certificate of Accomplishment  Awarded to",
            "Awarded to",
        ]

        award_rect = None
        for text in award_texts:
            hits = page.search_for(text)
            if hits:
                award_rect = hits[0]
                break

        if award_rect:
            name_y = award_rect.y1 + 15
            name_rect = fitz.Rect(pw * 0.25, name_y, pw * 0.75, name_y + 30)
        else:
            name_rect = fitz.Rect(pw * 0.25, ph * 0.28, pw * 0.75, ph * 0.32)

        page.insert_textbox(
            name_rect,
            str(emp_name),
            fontname=name_font,
            fontsize=22,
            align=fitz.TEXT_ALIGN_CENTER,
            color=(0, 0, 0),
        )

        # ---------- FIND EXAMINATION RESULT POSITION ----------
        exam_hits = page.search_for("EXAMINATION RESULT")
        if not exam_hits:
            exam_hits = page.search_for("Examination Result")

        exam_rect = exam_hits[0] if exam_hits else fitz.Rect(pw * 0.1, ph * 0.42, pw * 0.9, ph * 0.45)

        # ---------- CREATE TABLE ----------
        table_top = exam_rect.y1 + 18
        table_left = pw * 0.07
        table_width = pw * 0.86
        table_height = 50
        header_h = 22
        data_h = table_height - header_h
        col_w = table_width / 3

        table_bbox = fitz.Rect(table_left, table_top, table_left + table_width, table_top + table_height)
        page.draw_rect(table_bbox, fill=(1, 1, 1), color=(0, 0, 0), width=0.7)

        page.draw_line(
            fitz.Point(table_left, table_top + header_h),
            fitz.Point(table_left + table_width, table_top + header_h),
            color=(0, 0, 0),
            width=0.7,
        )

        for i in range(1, 3):
            x = table_left + i * col_w
            page.draw_line(
                fitz.Point(x, table_top),
                fitz.Point(x, table_top + table_height),
                color=(0, 0, 0),
                width=0.7,
            )

        headers = ["Standard", "Achieved Percentage", "Passing Criteria"]
        for i, title in enumerate(headers):
            cell = fitz.Rect(table_left + i * col_w, table_top, table_left + (i + 1) * col_w, table_top + header_h)
            page.insert_textbox(cell, title, fontname=arial_font, fontsize=11, align=fitz.TEXT_ALIGN_CENTER, color=(0, 0, 0))

        v_standard = (standard_text or "").strip()
        v_pct = str(percentage_text or "").strip()
        v_crit = str(criteria_text or "").strip()

        if v_pct and not v_pct.endswith("%"):
            v_pct += "%"
        if v_crit and not v_crit.endswith("%"):
            v_crit += "%"

        values = [v_standard, v_pct, v_crit]
        for i, val in enumerate(values):
            cell = fitz.Rect(table_left + i * col_w, table_top + header_h, table_left + (i + 1) * col_w, table_top + header_h + data_h)
            page.insert_textbox(cell, str(val), fontname=arial_font, fontsize=11, align=fitz.TEXT_ALIGN_CENTER, color=(0, 0, 0))

        # ---------- CERTIFICATE NO ----------
        cert_tag = {
            "Ds-1_template": "DS-1",
            "Cumulative_template": "CUMULATIVE",
            "API RP 7G-2_template": "API RP 7G-2",
            "API SPEC 5CT & 5A5_template": "API SPEC 5CT & 5A5",
        }.get(template_type, template_type)
        cert_value = f"{emp_id}/PTIS/{cert_tag}/2025"

        cert_label_texts = ["CERTIFICATE NO:", "CERTIFICATE NO :", "Certificate No:", "Certificate No :"]
        cert_inline_text = f"CERTIFICATE NO: {cert_value}"

        cert_replaced = False
        for text in cert_label_texts:
            hits = page.search_for(text)
            if hits:
                cert_label = hits[0]
                cert_rect = fitz.Rect(
                    cert_label.x0,
                    cert_label.y0,
                    cert_label.x0 + cert_label.width + 300,
                    cert_label.y1
                )

                fs = 14  # Increased font size
                if cert_rect.height < fs:
                    cy = (cert_rect.y0 + cert_rect.y1) / 2
                    cert_rect.y0 = cy - fs / 1.5
                    cert_rect.y1 = cy + fs / 1.5

                page.add_redact_annot(
                    cert_rect,
                    text=cert_inline_text,
                    fontname=arial_font,
                    fontsize=fs,
                    align=fitz.TEXT_ALIGN_LEFT,
                    text_color=(0, 0, 0),
                    fill=(1, 1, 1)
                )
                cert_replaced = True
                break

        if not cert_replaced:
            st.warning("Could not find CERTIFICATE NO label for replacement")

        # ---------- DATE ----------
        nice_date = _nice_date(test_date)
        date_label_texts = ["DATE:", "DATE :", "Date:", "Date :"]
        date_inline_text = f"{nice_date}"

        date_replaced = False
        for text in date_label_texts:
            hits = page.search_for(text)
            if hits:
                date_label = hits[0]
                date_rect = fitz.Rect(
                    date_label.x0 + 40,  # Moved 40px right
                    date_label.y0,
                    date_label.x0 + 240,
                    date_label.y1
                )

                fs = 14  # Increased font size
                if date_rect.height < fs:
                    cy = (date_rect.y0 + date_rect.y1) / 2
                    date_rect.y0 = cy - fs / 1.5
                    date_rect.y1 = cy + fs / 1.5

                page.add_redact_annot(
                    date_rect,
                    text=date_inline_text,
                    fontname=arial_font,
                    fontsize=fs,
                    align=fitz.TEXT_ALIGN_LEFT,
                    text_color=(0, 0, 0),
                    fill=(1, 1, 1)
                )
                date_replaced = True
                break

        if not date_replaced:
            st.warning("Could not find DATE label for replacement")

        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        safe_name = "".join(c for c in emp_name if c.isalnum() or c in (" ", "-", "_")).rstrip()
        stamp = nice_date.replace("/", "-").replace(":", "-").replace(" ", "_")
        certificate_filename = f"{template_type}_Certificate_{emp_id}_{safe_name}_{stamp}.pdf"
        output_path = f"/tmp/{certificate_filename}"

        doc.save(output_path, garbage=3, deflate=True)
        doc.close()

        st.success(f"✅ Generated certificate: {certificate_filename}")
        return output_path, certificate_filename

    except Exception as e:
        try:
            doc.close()
        except:
            pass
        st.error(f"❌ Error generating {template_type} certificate: {e}")
        st.error(f"Traceback: {traceback.format_exc()}")
        return None, None

        
# =====================
# Individual Test Downloads
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
    
    report_df = pd.DataFrame(report_data['Test Information'], columns=['Field', 'Value'])
    return report_df

def download_individual_test(emp_id, emp_name, test_data):
    report_df = create_individual_test_report(
        emp_id, 
        emp_name, 
        test_data['Date / Time'], 
        test_data['Test Type'], 
        test_data['Total'], 
        test_data['Right'], 
        test_data['Wrong'], 
        test_data['Percentage'], 
        test_data['Criteria'], 
        test_data['Status']
    )
    
    csv_buffer = io.StringIO()
    report_df.to_csv(csv_buffer, index=False)
    csv_data = csv_buffer.getvalue()
    
    timestamp = str(test_data['Date / Time']).replace('/', '_').replace(' ', '_').replace(':', '-')
    safe_name = "".join(c for c in emp_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
    filename = f"Test_Report_{emp_id}_{safe_name}_{test_data['Test Type']}_{timestamp}.csv"
    
    return csv_data, filename

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
        hh = int(h); mm = int(m); ss = int(s)
        return hh * 3600 + mm * 60 + ss
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
                all_worksheets = sheet.worksheets()
                for ws in all_worksheets:
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

# Admin login section
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
        
        # Create mappings for ID-Name relationship
        id_name_mapping = dict(zip(results_df["ID"].astype(str), results_df["Name"]))
        name_id_mapping = dict(zip(results_df["Name"], results_df["ID"].astype(str)))
        
        # Initialize session state keys if they don't exist
        id_key = f"emp_id_filter_{st.session_state.filter_reset_counter}"
        name_key = f"emp_name_filter_{st.session_state.filter_reset_counter}"
        
        if id_key not in st.session_state:
            st.session_state[id_key] = "All"
        if name_key not in st.session_state:
            st.session_state[name_key] = "All"
        
        # Callback functions for synchronization
        def sync_id_to_name():
            selected_id = st.session_state[id_key]
            if selected_id != "All" and selected_id in id_name_mapping:
                st.session_state[name_key] = id_name_mapping[selected_id]
            elif selected_id == "All":
                st.session_state[name_key] = "All"
        
        def sync_name_to_id():
            selected_name = st.session_state[name_key]
            if selected_name != "All" and selected_name in name_id_mapping:
                st.session_state[id_key] = name_id_mapping[selected_name]
            elif selected_name == "All":
                st.session_state[id_key] = "All"
        
        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        
        with filter_col1:
            employee_ids = ["All"] + sorted(results_df["ID"].astype(str).unique().tolist())
            selected_emp_id = st.selectbox(
                "Filter by Employee ID", 
                employee_ids, 
                index=employee_ids.index(st.session_state[id_key]) if st.session_state[id_key] in employee_ids else 0,
                key=id_key,
                on_change=sync_id_to_name
            )
        
        with filter_col2:
            employee_names = ["All"] + sorted(results_df["Name"].unique().tolist())
            selected_emp_name = st.selectbox(
                "Filter by Employee Name", 
                employee_names, 
                index=employee_names.index(st.session_state[name_key]) if st.session_state[name_key] in employee_names else 0,
                key=name_key,
                on_change=sync_name_to_id
            )
        
        with filter_col3:
            statuses = ["All"] + sorted(results_df["Status"].unique().tolist())
            selected_status = st.selectbox(
                "Filter by Status", 
                statuses, 
                index=0,
                key=f"status_filter_{st.session_state.filter_reset_counter}"
            )
        
        with filter_col4:
            test_types = ["All"] + sorted(results_df["Test Type"].unique().tolist())
            selected_test_type = st.selectbox(
                "Filter by Test Type", 
                test_types, 
                index=0,
                key=f"test_type_filter_{st.session_state.filter_reset_counter}"
            )
        
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
                            csv_data, filename = download_individual_test(
                                test_row['ID'], 
                                test_row['Name'], 
                                test_row
                            )
                            st.download_button(
                                label=f"📄 Download Test Report",
                                data=csv_data,
                                file_name=filename,
                                mime="text/csv",
                                use_container_width=True
                            )
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
                st.download_button(
                    label="📄 Download CSV",
                    data=csv,
                    file_name=f"all_test_results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
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
        
        # ======================
        # Certificate Generation
        # ======================
        st.markdown("---")
        st.subheader("📜 Generate Certificates")
        
        passed_results = results_df[results_df["Status"].astype(str).str.upper().eq("PASS")].copy()

        # Name filter list based on passed results
        cert_employee_names = ["All"] + sorted(passed_results["Name"].dropna().unique().tolist())
        selected_cert_name = st.selectbox(
            "Filter Certificates by Employee Name",
            cert_employee_names,
            index=0,
            key=f"cert_name_filter_{st.session_state.filter_reset_counter}"
        )

        # ---- Normalizer for standard names ----
        def _norm(s: str) -> str:
            s = str(s or "").upper().strip()
            s = s.replace("&", " ")
            s = s.replace("-", " ")
            s = re.sub(r"\s+", " ", s)
            s = s.replace("CUMMULATIVE", "CUMULATIVE")
            return s

        required_norm = {"DS 1", "CUMULATIVE", "API RP 7G 2", "API SPEC 5CT 5A5"}
        template_map = {
            "DS 1": "Ds-1_template",
            "CUMULATIVE": "Cumulative_template",
            "API RP 7G 2": "API RP 7G-2_template",
            "API SPEC 5CT 5A5": "API SPEC 5CT & 5A5_template",
        }
        required_in_order = ["DS 1", "CUMULATIVE", "API RP 7G 2", "API SPEC 5CT 5A5"]

        if st.button("Generate Certificates for Qualifying Employees"):
            # Normalize helper column
            passed_results["Test Type (norm)"] = passed_results["Test Type"].map(_norm)
            # Parse date to pick sensible rows when needed
            passed_results["_parsed_dt"] = pd.to_datetime(
                passed_results["Date / Time"], errors="coerce", dayfirst=True
            )

            grouped = passed_results.groupby("Name", dropna=False)
            qualifying_rows = []
            for name, group in grouped:
                passed_set = set(group["Test Type (norm)"].dropna().tolist())
                if required_norm.issubset(passed_set):
                    # Prefer CUMULATIVE row; else latest among required
                    cum_row = group[group["Test Type (norm)"] == "CUMULATIVE"]
                    if not cum_row.empty:
                        pick = cum_row.iloc[0]
                    else:
                        req_rows = group[group["Test Type (norm)"].isin(required_norm)].copy()
                        req_rows = req_rows.sort_values("_parsed_dt", ascending=False)
                        if req_rows.empty:
                            continue
                        pick = req_rows.iloc[0]
                    qualifying_rows.append(pick)

            qualifying_df = pd.DataFrame(qualifying_rows)

            # Guard before indexing
            if qualifying_df.empty or "Name" not in qualifying_df.columns:
                st.warning("No qualifying candidates found yet for the selected filters.")
            else:
                # Optional filter by selected name
                if selected_cert_name != "All":
                    qualifying_df = qualifying_df[qualifying_df["Name"].astype(str) == str(selected_cert_name)]

                if qualifying_df.empty:
                    st.warning("Candidate is ineligible as not all required standards are passed.")
                else:
                    certificate_files = []
                    # Generate four certificates per candidate, filling the row values from each standard's own result
                    for _, row_pick in qualifying_df.iterrows():
                        emp_id = row_pick["ID"]
                        emp_name = row_pick["Name"]
                        person_all = passed_results[passed_results["Name"] == emp_name].copy()

                        # For each required standard, find that row and render with that standard's Percentage/Criteria
                        for norm_std in required_in_order:
                            std_row = person_all[person_all["Test Type (norm)"] == norm_std]
                            if std_row.empty:
                                continue  # safety
                            r = std_row.iloc[0]

                            standard_text  = str(r["Test Type"]).strip()
                            # Normalize % fields
                            pct_val = r["Percentage"]
                            try:
                                pct_val_num = float(str(pct_val).replace("%","").strip())
                                percentage_text = f"{pct_val_num:.0f}%"
                            except:
                                percentage_text = str(pct_val) if str(pct_val).strip().endswith("%") else f"{str(pct_val).strip()}%"

                            crit_val = r["Criteria"]
                            try:
                                crit_val_num = float(str(crit_val).replace("%","").strip())
                                criteria_text = f"{crit_val_num:.0f}%"
                            except:
                                criteria_text = str(crit_val) if str(crit_val).strip().endswith("%") else f"{str(crit_val).strip()}%"

                            template_type = template_map.get(norm_std)
                            if not template_type:
                                continue

                            certificate_path, certificate_filename = generate_certificate(
                                emp_id=emp_id,
                                emp_name=emp_name,
                                test_date=str(r["Date / Time"]),   # used for filename; also used if "Date of Certification" exists
                                status=str(r["Status"]),
                                template_type=template_type,
                                standard_text=standard_text,
                                percentage_text=percentage_text,
                                criteria_text=criteria_text,
                                skip_dates=True  # your templates don't have validity/date; but date/cert no (if present) are replaced at fs=18
                            )
                            if certificate_path:
                                certificate_files.append((certificate_path, certificate_filename))

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
                        st.error("Failed to generate any certificates. Check templates and permissions.")
        
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

# Employee login and quiz
if not st.session_state.admin_logged_in and "quiz" not in st.session_state:
    st.subheader("Employee Login")
    
    def auto_populate_name():
        emp_id_input = st.session_state[f"id_{st.session_state.reset_counter}"]
        if emp_id_input and not employees.empty:
            try:
                fetched = employees[employees["ID"].astype(str).str.strip() == str(emp_id_input).strip()]
                if not fetched.empty:
                    fetched_name = str(fetched.iloc[0]["Name"])
                    st.session_state[f"name_{st.session_state.reset_counter}"] = fetched_name
                else:
                    st.session_state[f"name_{st.session_state.reset_counter}"] = ""
            except Exception:
                st.session_state[f"name_{st.session_state.reset_counter}"] = ""
    
    col1, col2 = st.columns(2)
    with col1:
        emp_id = st.text_input(
            "Employee ID", 
            value="", 
            key=f"id_{st.session_state.reset_counter}",
            help="Enter your employee identification number and press Enter",
            on_change=auto_populate_name
        )
    
    name_key = f"name_{st.session_state.reset_counter}"
    if name_key not in st.session_state:
        st.session_state[name_key] = ""
    
    with col2:
        name = st.text_input(
            "Name", 
            key=name_key,
            help="This will auto-fill when you enter a valid Employee ID"
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

# Quiz interface
elif "quiz" in st.session_state:
    qstate = st.session_state.quiz
    total, criteria, h, m, s = get_info_for_standard(standards, qstate["standard"])
    total_secs = format_timer(h, m, s)

    elapsed = int(time.time() - qstate["start_ts"])
    remaining = max(0, total_secs - elapsed)

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
            ok, msg = append_result(
                qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
            )
            st.session_state["submitted"] = True
            st.session_state["submit_result"] = (ok, msg, right, wrong, total_q, pct, criteria, status, final_score)
            st.query_params.clear()
            st.rerun()

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
        @keyframes pulse {{
            0% {{ transform: scale(1); opacity: 1; }}
            50% {{ transform: scale(1.05); opacity: 0.8; }}
            100% {{ transform: scale(1); opacity: 1; }}
        }}
        .timer-pulse {{ animation: pulse 1s infinite; }}
        .timer-container {{
            padding: 20px; border-radius: 15px; text-align: center; font-size: 22px; font-weight: bold;
            margin-bottom: 20px; box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2); border: 3px solid rgba(255, 255, 255, 0.1);
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
                    form.method = 'POST'; form.action = window.location.href;
                    var input = document.createElement('input');
                    input.type = 'hidden'; input.name = 'timeout'; input.value = 'true';
                    form.appendChild(input); document.body.appendChild(form); form.submit(); return;
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
                if (remaining <= 300) {{ bg_color = '#DC2626'; text_color = 'white'; icon = '🚨'; pulse_class = 'timer-pulse'; }}
                else if (remaining <= 900) {{ bg_color = '#DC2626'; text_color = 'white'; icon = '⚠️'; }}
                else if (remaining <= 1200) {{ bg_color = '#D97706'; text_color = 'white'; icon = '⏰'; }}
                else {{ bg_color = '#1E3A8A'; text_color = 'white'; icon = '⏰'; }}
                container.style.background = `linear-gradient(135deg, ${bg_color}, ${bg_color}CC)`;
                container.style.color = text_color; iconElem.innerText = icon;
                if (pulse_class) {{ container.classList.add(pulse_class); }} else {{ container.classList.remove('timer-pulse'); }}
                remaining--;
            }}
            if (interval) {{ clearInterval(interval); }}
            updateTimer(); interval = setInterval(updateTimer, 1000);
        }})();
        </script>
        """
        components.html(timer_html, height=150)

        if st.query_params.get("timeout", ["false"])[0] == "true":
            if len(qstate["queue"]) > 0:
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
                ok, msg = append_result(
                    qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
                )
                st.session_state["submitted"] = True
                st.session_state["submit_result"] = (ok, msg, right, wrong, total_q, pct, criteria, status, final_score)
                st.query_params.clear()
                st.rerun()

        if remaining <= 300:
            st.markdown('<div style="margin-bottom: 8px;"></div>', unsafe_allow_html=True)
            st.warning("🚨 URGENT: Less than 5 minutes remaining!")
        elif remaining <= 900:
            st.markdown('<div style="margin-bottom: 8px;"></div>', unsafe_allow_html=True)
            st.warning("⚠️ WARNING: Less than 15 minutes remaining!")
        elif remaining <= 1200:
            st.markdown('<div style="margin-bottom: 8px;"></div>', unsafe_allow_html=True)
            st.info("⏰ NOTICE: Less than 20 minutes remaining!")

    elif total_secs > 0 and "submitted" in st.session_state:
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

        stopped_timer_html = f"""
        <style>
        @keyframes pulse {{
            0% {{ transform: scale(1); opacity: 1; }}
            50% {{ transform: scale(1.05); opacity: 0.8; }}
            100% {{ transform: scale(1); opacity: 1; }}
        }}
        .timer-pulse {{ animation: pulse 1s infinite; }}
        .timer-container {{
            padding: 20px; border-radius: 15px; text-align: center; font-size: 22px; font-weight: bold;
            margin-bottom: 20px; box-shadow: 0 8px 16px rgba(0, 0, 0, 0.2); border: 3px solid rgba(255, 255, 255, 0.1);
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
        
        if is_previously_skipped:
            st.markdown("🔄 **This question was skipped earlier**")
            st.subheader(f"Q{current_qid+1}. {question}")
        else:
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
            ok, msg = append_result(
                qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, criteria, status, qstate["standard"]
            )
            st.session_state["submitted"] = True
            st.session_state["submit_result"] = (ok, msg, right, wrong, total_q, pct, criteria, status, final_score)
            st.rerun()

    if "submitted" in st.session_state:
        if "submit_result" in st.session_state:
            result_data = st.session_state["submit_result"]
            ok, msg, right, wrong, total_q, pct, criteria, status, final_score = result_data
            
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
