import streamlit as st
import pandas as pd
import numpy as np
import streamlit as st
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
import tempfile
import json
import datetime

# =====================
# Paths / Files
# =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Templates are in root folder, not in db subfolder
DB_FOLDER = os.path.join(BASE_DIR, "db")
QUESTIONS_FOLDER = os.path.join(BASE_DIR, "Questions")

# =====================
# Google Sheets Setup
# =====================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load Google Sheets credentials from Streamlit secrets
try:
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scope
    )
    client = gspread.authorize(creds)
    GSHEET_URL = st.secrets["connections"]["gsheets"]["spreadsheet"]
    GSHEETS_AVAILABLE = True
except Exception as e:
    st.error(f"⚠️ Google Sheets credentials error: {str(e)}")
    st.info("Please configure Google Sheets credentials in Streamlit secrets")
    client = None
    GSHEET_URL = None
    GSHEETS_AVAILABLE = False

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
            'Date / Time': ['DATE', 'Date', 'date', 'Timestamp', 'timestamp', 'Time', 'Date / Time'],
            'Answers': ['ANSWERS', 'Answers', 'answers', 'Answer Data', 'ANSWER DATA']
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
        
        # Add Answers column if it exists (optional column)
        if "Answers" not in df.columns:
            df["Answers"] = ""
        
        for col in ["Total", "Right", "Wrong"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        if "Percentage" in df.columns:
            df["Percentage"] = df["Percentage"].astype(str).str.replace("%", "").str.replace(" ", "")
            df["Percentage"] = pd.to_numeric(df["Percentage"], errors='coerce').fillna(0).astype(float)
        
        df = df.sort_values('_original_order').drop('_original_order', axis=1)
        df = df.reset_index(drop=True)
        
        # Return all columns including Answers
        return_columns = required_columns + ["Answers"]
        return df[return_columns]
        
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
            "Standard": ["Basic", "Basic", "Advanced", "Advanced", "Cumulative"],
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
    Returns path to template pdf in root folder.
    Expected 'template_type' already includes the exact filename stem like:
      "Ds-1_template", "Cumulative_template", "API RP 7G-2_template", "API SPEC 5CT & 5A5_template"
    """
    # Templates are directly in the root folder
    template_path = os.path.join(DB_FOLDER, f"{template_type}.pdf")
    if os.path.exists(template_path):
        print(f"✅ Found template: {template_path}")
        return template_path
    else:
        st.error(f"❌ Template not found: {template_path}")
        st.info(f"Please ensure '{template_type}.pdf' exists in the project folder")
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
    Supports fallback to generic template if specific template not found
    """
    template_path = get_template_path(template_type)
    
    # If specific template not found, try fallback templates
    if not template_path:
        fallback_templates = ["Ds-1_template", "Cumulative_template"]
        for fallback in fallback_templates:
            template_path = get_template_path(fallback)
            if template_path:
                st.warning(f"⚠️ Using fallback template '{fallback}' for '{template_type}'")
                break
    
    if not template_path:
        st.error(f"No template available for '{template_type}'. Cannot generate certificate.")
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
        name_font = "times-bolditalic"  # Bold + Italic combined

        try:
            arial_fontfile = os.path.join(DB_FOLDER, "arial.ttf")
            if os.path.exists(arial_fontfile):
                doc.insert_font(fontname="Arial", fontfile=arial_fontfile)
                arial_font = "Arial"
        except:
            pass

        try:
            # Try to load a bold-italic custom font if available
            corsiva_fontfile = os.path.join(DB_FOLDER, "monotype_corsiva.ttf")
            if os.path.exists(corsiva_fontfile):
                doc.insert_font(fontname="MonotypeCorsiva", fontfile=corsiva_fontfile)
                name_font = "MonotypeCorsiva"  # Corsiva is already italic and decorative
            else:
                name_font = "times-bolditalic"  # Fallback to Times Bold-Italic
        except:
            name_font = "times-bolditalic"

        # ---------- REPLACE TEMPLATE NAME "Israr Hussain" ----------
        template_name_hits = page.search_for("Israr Hussain")
        if template_name_hits:
            for hit in template_name_hits:
                # Define desired font size
                replace_fontsize = 22  # Adjust this as needed
                
                # Calculate text width for the new name to minimize white space
                # Estimate: each character is roughly 60% of font size in width
                estimated_width = len(str(emp_name)) * replace_fontsize * 0.6
                
                # Use minimal height - just enough for the text (1.2x font size)
                rect_height = replace_fontsize * 1.2
                
                # Center the rect around the original hit position
                center_x = (hit.x0 + hit.x1) / 2
                center_y = (hit.y0 + hit.y1) / 2
                
                # Create tight-fitting rectangle
                name_replace_rect = fitz.Rect(
                    center_x - estimated_width / 2,
                    center_y - rect_height / 2,
                    center_x + estimated_width / 2,
                    center_y + rect_height / 2
                )
                
                # Apply redaction with actual employee name (bold-italic)
                page.add_redact_annot(
                    name_replace_rect,
                    text=str(emp_name),
                    fontname="times-bolditalic",  # Bold + Italic
                    fontsize=replace_fontsize,
                    align=fitz.TEXT_ALIGN_CENTER,
                    text_color=(0, 0, 0),
                    fill=(1, 1, 1)
                )

        # ---------- FIND NAME POSITION ----------
        # Search for the "Certificate of Accomplishment Awarded to" text
        award_texts = [
            "Certificate of Accomplishment Awarded to",
            "Certificate of Accomplishment Awarded to",
            "Awarded to"
        ]
        award_rect = None
        for text in award_texts:
            hits = page.search_for(text)
            if hits:
                award_rect = hits[0]
                break

        # Position name below the award text with REDUCED gap
        name_fontsize = 28  # Increased font size
        if award_rect:
            # Place name centered, 5px below the award text (reduced from 15px)
            name_y = award_rect.y1 + 5
            # Use minimal height (1.2x font size) to avoid covering other content
            name_rect = fitz.Rect(pw * 0.25, name_y, pw * 0.75, name_y + (name_fontsize * 1.2))
        else:
            # Fallback: center of upper third with minimal height
            name_rect = fitz.Rect(pw * 0.25, ph * 0.28, pw * 0.75, ph * 0.28 + (name_fontsize * 1.2))

        # Insert name (bold-italic, DO NOT draw white rectangle first)
        page.insert_textbox(
            name_rect,
            str(emp_name),
            fontname="times-bolditalic",  # Bold + Italic
            fontsize=name_fontsize,  # Now uses variable for easy adjustment
            align=fitz.TEXT_ALIGN_CENTER,
            color=(0, 0, 0),
        )

        # ---------- FIX "FOR" TEXT GAP ----------
        # Search for "For" text and move it closer to the name
        for_hits = page.search_for("For")
        if for_hits:
            for hit in for_hits:
                # Check if this "For" is below the name (y coordinate is larger)
                if hit.y0 > name_rect.y1:
                    # Reduce gap between name and "For" - move "For" up by covering with white and redrawing
                    for_fontsize = 16
                    # Cover the old "For" text
                    cover_rect = fitz.Rect(hit.x0 - 5, hit.y0 - 2, hit.x1 + 5, hit.y1 + 2)
                    page.draw_rect(cover_rect, fill=(1, 1, 1), color=(1, 1, 1))
                    
                    # Redraw "For" closer to name (reduced gap from ~20px to ~8px)
                    new_for_y = name_rect.y1 + 8
                    new_for_rect = fitz.Rect(pw * 0.35, new_for_y, pw * 0.65, new_for_y + for_fontsize * 1.3)
                    page.insert_textbox(
                        new_for_rect,
                        "For",
                        fontname=arial_font,
                        fontsize=for_fontsize,
                        align=fitz.TEXT_ALIGN_CENTER,
                        color=(0, 0, 0),
                    )
                    break  # Only process first match

        # ---------- FIND EXAMINATION RESULT POSITION ----------
        exam_hits = page.search_for("EXAMINATION RESULT")
        if not exam_hits:
            exam_hits = page.search_for("Examination Result")
        exam_rect = exam_hits[0] if exam_hits else fitz.Rect(pw * 0.1, ph * 0.42, pw * 0.9, ph * 0.45)

        # ---------- CREATE IMPROVED TABLE ----------
        table_top = exam_rect.y1 + 18
        table_left = pw * 0.07  # Wider table
        table_width = pw * 0.86  # Increased width (from 0.8 to 0.86)
        table_height = 45  # Reduced height (from 55 to 45)
        header_h = 25  # Taller header
        data_h = table_height - header_h  # 2nd row height will be 20px (reduced)
        
        # Equal width columns (divide by 2 for Standard and Achieved/Passing together)
        col1_w = table_width * 0.33  # Standard column
        col2_w = table_width * 0.33  # Achieved Percentage column
        col3_w = table_width * 0.34  # Passing Criteria column (slightly wider to use remaining space)

        # Clear table area with white rectangle
        table_bbox = fitz.Rect(table_left, table_top, table_left + table_width, table_top + table_height)
        page.draw_rect(table_bbox, fill=(1, 1, 1), color=(0, 0, 0), width=1.0)

        # Draw horizontal line (thicker)
        page.draw_line(
            fitz.Point(table_left, table_top + header_h),
            fitz.Point(table_left + table_width, table_top + header_h),
            color=(0, 0, 0),
            width=1.0,
        )

        # Draw vertical lines (thicker)
        x1 = table_left + col1_w
        page.draw_line(
            fitz.Point(x1, table_top),
            fitz.Point(x1, table_top + table_height),
            color=(0, 0, 0),
            width=1.0,
        )
        
        x2 = table_left + col1_w + col2_w
        page.draw_line(
            fitz.Point(x2, table_top),
            fitz.Point(x2, table_top + table_height),
            color=(0, 0, 0),
            width=1.0,
        )

        # Add BOLD headers with better font
        headers = ["Standard", "Achieved Percentage", "Passing Criteria"]
        header_positions = [
            (table_left, col1_w),
            (table_left + col1_w, col2_w),
            (table_left + col1_w + col2_w, col3_w)
        ]
        
        for i, (title, (x_pos, width)) in enumerate(zip(headers, header_positions)):
            cell = fitz.Rect(x_pos, table_top, x_pos + width, table_top + header_h)
            # Use bold font for headers
            page.insert_textbox(
                cell, 
                title, 
                fontname="times-bold",  # BOLD font for headers
                fontsize=12,  # Slightly larger
                align=fitz.TEXT_ALIGN_CENTER, 
                color=(0, 0, 0)
            )

        # Add values with proper formatting and CENTERED alignment
        v_standard = (standard_text or "").strip()
        v_pct = str(percentage_text or "").strip()
        v_crit = str(criteria_text or "").strip()

        if v_pct and not v_pct.endswith("%"):
            v_pct += "%"
        if v_crit and not v_crit.endswith("%"):
            v_crit += "%"

        values = [v_standard, v_pct, v_crit]
        for i, (val, (x_pos, width)) in enumerate(zip(values, header_positions)):
            cell = fitz.Rect(x_pos, table_top + header_h, x_pos + width, table_top + table_height)
            page.insert_textbox(
                cell, 
                str(val), 
                fontname=arial_font, 
                fontsize=11, 
                align=fitz.TEXT_ALIGN_CENTER,  # Explicitly centered
                color=(0, 0, 0)
            )

        # ---------- CERTIFICATE NO ----------
        # Extract short code for certificate number (e.g., "DS-1" from "DS-1 3rd Volume 5th Edition")
        cert_tag = {
            "Ds-1_template": "DS-1",
            "Cumulative_template": "Cumulative",
            "API RP 7G-2_template": "API RP 7G-2",
            "API SPEC 5CT & 5A5_template": "API SPEC 5CT & 5A5",
        }.get(template_type, template_type)
        
        # If cert_tag is still template_type, try to extract from standard_text
        # This handles cases like "DS-1 3rd Volume 5th Edition" -> "DS-1"
        if cert_tag == template_type and standard_text:
            # Extract first meaningful part before space/volume/edition keywords
            parts = str(standard_text).strip().split()
            if parts:
                # Look for patterns like "DS-1", "API", etc.
                first_part = parts[0]
                # If it's a compound like "DS-1", use it directly
                if "-" in first_part or len(parts) == 1:
                    cert_tag = first_part
                # If multiple parts, check if first 2-3 form a standard code
                elif len(parts) >= 2 and parts[1].replace("-", "").replace(".", "").isalnum():
                    cert_tag = f"{parts[0]}-{parts[1]}".replace("--", "-")
                else:
                    cert_tag = first_part

        cert_value = f"{emp_id}/PTIS/{cert_tag}/2025"

        # Search for CERTIFICATE NO label and replace the entire line
        cert_label_texts = ["CERTIFICATE NO:", "CERTIFICATE NO :", "Certificate No:", "Certificate No :"]
        cert_inline_text = f"CERTIFICATE NO: {cert_value}"
        cert_replaced = False

        for text in cert_label_texts:
            hits = page.search_for(text)
            if hits:
                cert_label = hits[0]
                # Create rect that covers the label + value area
                cert_rect = fitz.Rect(
                    cert_label.x0,
                    cert_label.y0,
                    cert_label.x0 + cert_label.width + 400,
                    cert_label.y1
                )

                # Adjust height if needed
                fs = 14
                if cert_rect.height < fs * 1.1:
                    cy = (cert_rect.y0 + cert_rect.y1) / 2
                    cert_rect.y0 = cy - fs
                    cert_rect.y1 = cy + fs

                # Apply redaction with new text
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

        # Search for DATE label and replace the entire line
        date_label_texts = ["DATE:", "DATE :", "Date:", "Date :"]
        date_inline_text = f"Date: {nice_date}"
        date_replaced = False

        for text in date_label_texts:
            hits = page.search_for(text)
            if hits:
                date_label = hits[0]
                # Create rect that covers the label + value area
                date_rect = fitz.Rect(
                    date_label.x0,
                    date_label.y0,
                    date_label.x0 + date_label.width + 110,
                    date_label.y1
                )

                # Adjust height if needed
                fs = 14
                if date_rect.height < fs * 1.1:
                    cy = (date_rect.y0 + date_rect.y1) / 2
                    date_rect.y0 = cy - fs
                    date_rect.y1 = cy + fs

                # Apply redaction with new text
                page.add_redact_annot(
                    date_rect,
                    text=date_inline_text,
                    fontname=arial_font,
                    fontsize=fs,
                    align=fitz.TEXT_ALIGN_CENTER,
                    text_color=(0, 0, 0),
                    fill=(1, 1, 1)
                )
                date_replaced = True
                break

        if not date_replaced:
            st.warning("Could not find DATE label for replacement")

        # ---------- APPLY REDACTIONS ----------
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # ---------- SAVE ----------
        safe_name = "".join(c for c in emp_name if c.isalnum() or c in (" ", "-", "_")).rstrip()
        
        # Clean template_type: remove "_template" suffix
        cert_name = template_type.replace("_template", "").replace("_Template", "")
        
        # Certificate filename WITHOUT date: Standard_Certificate_ID_Name.pdf
        certificate_filename = f"{cert_name}_Certificate_{emp_id}_{safe_name}.pdf"
        
        # Use tempfile module for cross-platform temp directory
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, certificate_filename)

        doc.save(output_path, garbage=3, deflate=True)
        doc.close()

        st.success(f"Generated certificate: {certificate_filename}")
        return output_path, certificate_filename

    except Exception as e:
        try:
            doc.close()
        except:
            pass
        st.error(f"Error generating {template_type} certificate: {e}")
        st.error(f"Traceback: {traceback.format_exc()}")
        return None, None

def generate_mpt_pt_certificate(
    emp_id,
    emp_name,
    test_date,
    template_type,
    general_standard,
    general_percentage,
    general_criteria,
    specific_standard,
    specific_percentage,
    specific_criteria,
):
    """
    Generate MPT/PT certificates with 2-row table (General + Specific)
    Certificate generated only when both tests are passed
    """
    template_path = get_template_path(template_type)
    
    if not template_path:
        fallback_templates = ["MT_template", "PT_template"]
        for fallback in fallback_templates:
            template_path = get_template_path(fallback)
            if template_path:
                st.warning(f"⚠️ Using fallback template '{fallback}' for '{template_type}'")
                break
    
    if not template_path:
        st.error(f"No template available for '{template_type}'. Cannot generate certificate.")
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
        name_font = "times-bolditalic"

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
            else:
                name_font = "times-bolditalic"
        except:
            name_font = "times-bolditalic"

        # ---------- REPLACE TEMPLATE NAME "Israr Hussain" ----------
        template_name_hits = page.search_for("Israr Hussain")
        if template_name_hits:
            for hit in template_name_hits:
                replace_fontsize = 22
                estimated_width = len(str(emp_name)) * replace_fontsize * 0.6
                rect_height = replace_fontsize * 1.2
                center_x = (hit.x0 + hit.x1) / 2
                center_y = (hit.y0 + hit.y1) / 2
                
                name_replace_rect = fitz.Rect(
                    center_x - estimated_width / 2,
                    center_y - rect_height / 2,
                    center_x + estimated_width / 2,
                    center_y + rect_height / 2
                )
                
                page.add_redact_annot(
                    name_replace_rect,
                    text=str(emp_name),
                    fontname="times-bolditalic",
                    fontsize=replace_fontsize,
                    align=fitz.TEXT_ALIGN_CENTER,
                    text_color=(0, 0, 0),
                    fill=(1, 1, 1)
                )

        # ---------- FIND EXAMINATION RESULT POSITION ----------
        exam_hits = page.search_for("EXAMINATION RESULT")
        if not exam_hits:
            exam_hits = page.search_for("Examination Result")
        exam_rect = exam_hits[0] if exam_hits else fitz.Rect(pw * 0.1, ph * 0.42, pw * 0.9, ph * 0.45)

        # ---------- CREATE 2-ROW TABLE ----------
        table_top = exam_rect.y1 + 18
        table_left = pw * 0.07
        table_width = pw * 0.86
        table_height = 65  # Taller for 2 data rows
        header_h = 20
        row_h = 22.5  # Each data row height
        
        col1_w = table_width * 0.33
        col2_w = table_width * 0.33
        col3_w = table_width * 0.34

        # Clear table area
        table_bbox = fitz.Rect(table_left, table_top, table_left + table_width, table_top + table_height)
        page.draw_rect(table_bbox, fill=(1, 1, 1), color=(0, 0, 0), width=1.0)

        # Draw horizontal lines
        page.draw_line(
            fitz.Point(table_left, table_top + header_h),
            fitz.Point(table_left + table_width, table_top + header_h),
            color=(0, 0, 0), width=1.0
        )
        page.draw_line(
            fitz.Point(table_left, table_top + header_h + row_h),
            fitz.Point(table_left + table_width, table_top + header_h + row_h),
            color=(0, 0, 0), width=1.0
        )

        # Draw vertical lines
        x1 = table_left + col1_w
        x2 = table_left + col1_w + col2_w
        for x in [x1, x2]:
            page.draw_line(
                fitz.Point(x, table_top),
                fitz.Point(x, table_top + table_height),
                color=(0, 0, 0), width=1.0
            )

        # Headers
        headers = ["Standard", "Percentage", "Criteria"]
        header_positions = [
            (table_left, col1_w),
            (table_left + col1_w, col2_w),
            (table_left + col1_w + col2_w, col3_w)
        ]
        
        for title, (x_pos, width) in zip(headers, header_positions):
            cell = fitz.Rect(x_pos, table_top, x_pos + width, table_top + header_h)
            page.insert_textbox(
                cell, title, 
                fontname="times-bold",
                fontsize=11,
                align=fitz.TEXT_ALIGN_CENTER,
                color=(0, 0, 0)
            )

        # Row 1: General
        general_values = [general_standard, general_percentage, general_criteria]
        for val, (x_pos, width) in zip(general_values, header_positions):
            cell = fitz.Rect(x_pos, table_top + header_h, x_pos + width, table_top + header_h + row_h)
            page.insert_textbox(
                cell, str(val),
                fontname=arial_font,
                fontsize=10,
                align=fitz.TEXT_ALIGN_CENTER,
                color=(0, 0, 0)
            )

        # Row 2: Specific
        specific_values = [specific_standard, specific_percentage, specific_criteria]
        for val, (x_pos, width) in zip(specific_values, header_positions):
            cell = fitz.Rect(x_pos, table_top + header_h + row_h, x_pos + width, table_top + table_height)
            page.insert_textbox(
                cell, str(val),
                fontname=arial_font,
                fontsize=10,
                align=fitz.TEXT_ALIGN_CENTER,
                color=(0, 0, 0)
            )

        # ---------- CERTIFICATE NO ----------
        cert_tag = "MT" if "MT" in template_type else "PT"
        cert_value = f"{emp_id}/PTIS/{cert_tag}/2025"
        
        cert_label_texts = ["CERTIFICATE NO:", "CERTIFICATE NO :", "Certificate No:", "Certificate No :"]
        cert_inline_text = f"CERTIFICATE NO: {cert_value}"
        
        for text in cert_label_texts:
            hits = page.search_for(text)
            if hits:
                cert_label = hits[0]
                cert_rect = fitz.Rect(
                    cert_label.x0, cert_label.y0,
                    cert_label.x0 + cert_label.width + 400,
                    cert_label.y1
                )
                fs = 14
                if cert_rect.height < fs * 1.1:
                    cy = (cert_rect.y0 + cert_rect.y1) / 2
                    cert_rect.y0 = cy - fs
                    cert_rect.y1 = cy + fs
                
                page.add_redact_annot(
                    cert_rect, text=cert_inline_text,
                    fontname=arial_font, fontsize=fs,
                    align=fitz.TEXT_ALIGN_LEFT,
                    text_color=(0, 0, 0), fill=(1, 1, 1)
                )
                break

        # ---------- DATE OF CERTIFICATION ----------
        nice_date = _nice_date(test_date)
        date_label_texts = ["Date of Certification:", "Date of Certification :", "DATE OF CERTIFICATION:"]
        date_inline_text = f"Date of Certification: {nice_date}"
        
        for text in date_label_texts:
            hits = page.search_for(text)
            if hits:
                date_label = hits[0]
                date_rect = fitz.Rect(
                    date_label.x0, date_label.y0,
                    date_label.x0 + date_label.width + 150,
                    date_label.y1
                )
                fs = 10
                if date_rect.height < fs * 1.1:
                    cy = (date_rect.y0 + date_rect.y1) / 2
                    date_rect.y0 = cy - fs
                    date_rect.y1 = cy + fs
                
                page.add_redact_annot(
                    date_rect, text=date_inline_text,
                    fontname=arial_font, fontsize=fs,
                    align=fitz.TEXT_ALIGN_LEFT,
                    text_color=(0, 0, 0), fill=(1, 1, 1)
                )
                break

        # ---------- VALIDITY (5 years) ----------
        try:
            cert_date = pd.to_datetime(test_date, dayfirst=True)
            validity_date = cert_date + pd.DateOffset(years=5)
            validity_str = validity_date.strftime("%d-%B-%Y")
        except:
            validity_str = "N/A"
        
        validity_label_texts = ["Validity:", "Validity :", "VALIDITY:"]
        validity_inline_text = f"Validity: {validity_str}"
        
        for text in validity_label_texts:
            hits = page.search_for(text)
            if hits:
                validity_label = hits[0]
                validity_rect = fitz.Rect(
                    validity_label.x0, validity_label.y0,
                    validity_label.x0 + validity_label.width + 150,
                    validity_label.y1
                )
                fs = 10
                if validity_rect.height < fs * 1.1:
                    cy = (validity_rect.y0 + validity_rect.y1) / 2
                    validity_rect.y0 = cy - fs
                    validity_rect.y1 = cy + fs
                
                page.add_redact_annot(
                    validity_rect, text=validity_inline_text,
                    fontname=arial_font, fontsize=fs,
                    align=fitz.TEXT_ALIGN_LEFT,
                    text_color=(0, 0, 0), fill=(1, 1, 1)
                )
                break

        # ---------- APPLY REDACTIONS ----------
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # ---------- SAVE ----------
        safe_name = "".join(c for c in emp_name if c.isalnum() or c in (" ", "-", "_")).rstrip()
        cert_name = template_type.replace("_template", "").replace("_Template", "")
        certificate_filename = f"{cert_name}_Certificate_{emp_id}_{safe_name}.pdf"
        
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, certificate_filename)

        doc.save(output_path, garbage=3, deflate=True)
        doc.close()

        st.success(f"Generated certificate: {certificate_filename}")
        return output_path, certificate_filename

    except Exception as e:
        try:
            doc.close()
        except:
            pass
        st.error(f"Error generating {template_type} certificate: {e}")
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

def generate_test_sheet_pdf(emp_id, emp_name, standard, test_date, questions_data, answers_data, right, wrong, total, pct, criteria, status):
    """
    Generate a detailed PDF with all test questions, candidate answers, and correct answers
    """
    try:
        # Create PDF
        doc = fitz.open()
        page_width = 595  # A4 width
        page_height = 842  # A4 height
        
        # Add first page
        page = doc.new_page(width=page_width, height=page_height)
        
        y_position = 50
        margin_left = 50
        margin_right = page_width - 50
        
        # Header - using insert_text which is more reliable
        header_text = "TEST ANSWER SHEET"
        page.insert_text(
            fitz.Point(page_width / 2 - 100, y_position),
            header_text,
            fontsize=18,
            color=(0, 0, 0.5)
        )
        y_position += 40
        
        # Employee Info
        info_lines = [
            f"Employee ID: {emp_id}",
            f"Employee Name: {emp_name}",
            f"Test Standard: {standard}",
            f"Test Date: {test_date}",
            "",
            f"Total Questions: {total}",
            f"Correct Answers: {right}",
            f"Wrong Answers: {wrong}",
            f"Score: {right - (wrong * 0.25):.2f}/{total}",
            f"Percentage: {pct:.2f}%",
            f"Passing Criteria: {criteria}%",
            f"Status: {status}",
        ]
        
        for line in info_lines:
            if line:
                page.insert_text(
                    fitz.Point(margin_left, y_position),
                    line,
                    fontsize=11,
                    color=(0, 0, 0)
                )
            y_position += 18
        
        # Draw separator line
        page.draw_line(
            fitz.Point(margin_left, y_position),
            fitz.Point(margin_right, y_position),
            color=(0, 0, 0),
            width=1
        )
        y_position += 20
        
        # Questions and Answers
        for idx, (q_data, ans_info) in enumerate(zip(questions_data, answers_data), 1):
            # Check if we need a new page
            if y_position > page_height - 150:
                page = doc.new_page(width=page_width, height=page_height)
                y_position = 50
            
            # Question number and text
            question_text = f"Q{idx}. {q_data['Question']}"
            
            # Simple text wrapping
            max_chars = 80
            question_lines = []
            while len(question_text) > max_chars:
                # Find last space before max_chars
                split_pos = question_text[:max_chars].rfind(' ')
                if split_pos == -1:
                    split_pos = max_chars
                question_lines.append(question_text[:split_pos])
                question_text = question_text[split_pos:].lstrip()
            if question_text:
                question_lines.append(question_text)
            
            # Draw question
            for line in question_lines:
                page.insert_text(
                    fitz.Point(margin_left, y_position),
                    line,
                    fontsize=10,
                    color=(0, 0, 0)
                )
                y_position += 15
            
            y_position += 5
            
            # Options
            options = [
                ('A', q_data['A']),
                ('B', q_data['B']),
                ('C', q_data['C']),
                ('D', q_data['D'])
            ]
            
            correct_answer = str(q_data['Answer']).strip().upper()
            candidate_choice = ans_info.get('choice', 'Not Attempted')
            
            # Map correct answer to full text
            correct_text = ""
            for opt_letter, opt_text in options:
                if opt_letter == correct_answer:
                    correct_text = opt_text
                    break
            
            # Also map candidate choice to letter for proper comparison
            candidate_letter = ""
            for opt_letter, opt_text in options:
                if opt_text == candidate_choice:
                    candidate_letter = opt_letter
                    break
            
            for opt_letter, opt_text in options:
                # Determine color based on answer - compare letters, not text
                is_correct = (opt_letter == correct_answer)
                is_selected = (opt_letter == candidate_letter)
                
                if is_selected and is_correct:
                    color = (0, 0.5, 0)  # Green - correct selection
                    prefix = "[OK] "
                elif is_selected and not is_correct:
                    color = (0.8, 0, 0)  # Red - wrong selection
                    prefix = "[X] "
                elif is_correct:
                    color = (0, 0.4, 0)  # Dark green - correct answer
                    prefix = "[->] "
                else:
                    color = (0, 0, 0)  # Black - normal
                    prefix = "[ ] "
                
                option_text = f"{prefix}{opt_letter}) {opt_text}"
                
                max_opt_chars = 75
                if len(option_text) > max_opt_chars:
                    option_text = option_text[:max_opt_chars] + "..."
                
                page.insert_text(
                    fitz.Point(margin_left + 20, y_position),
                    option_text,
                    fontsize=9,
                    color=color
                )
                y_position += 14
            
            # Show result for this question
            if ans_info.get('is_correct'):
                result_text = "[OK] CORRECT"
                result_color = (0, 0.5, 0)
            elif candidate_choice == 'Not Attempted' or candidate_choice == 'Not Stored':
                if candidate_choice == 'Not Stored':
                    result_text = f"[i] Candidate Answer Not Recorded | Correct Answer: {correct_answer}) {correct_text}"
                else:
                    result_text = "[  ] NOT ATTEMPTED"
                result_color = (0.5, 0.5, 0.5)
            elif candidate_letter == correct_answer:
                result_text = "[OK] CORRECT"
                result_color = (0, 0.5, 0)
            else:
                result_text = f"[X] Wrong | Correct Answer: {correct_answer}) {correct_text}"
                result_color = (0.8, 0, 0)
            
            if len(result_text) > 75:
                result_text = result_text[:75] + "..."
            
            page.insert_text(
                fitz.Point(margin_left + 20, y_position),
                result_text,
                fontsize=9,
                color=result_color
            )
            y_position += 25
        
        # Save PDF
        safe_name = "".join(c for c in emp_name if c.isalnum() or c in (" ", "-", "_")).rstrip()
        
        # Answer sheet filename WITHOUT date: AnswerSheet_ID_Name_Standard.pdf
        pdf_filename = f"AnswerSheet_{emp_id}_{safe_name}_{standard}.pdf"
        
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, pdf_filename)
        
        doc.save(output_path, garbage=3, deflate=True)
        doc.close()
        
        return output_path, pdf_filename
        
    except Exception as e:
        st.error(f"Error generating test sheet PDF: {str(e)}")
        st.error(f"Traceback: {traceback.format_exc()}")
        return None, None

# =====================
# Helpers
# =====================
def start_quiz_session(emp_id, emp_name, standard, questions_df, total):
    if standard == "Cumulative":
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

    # Convert sampled questions to dict format for PDF generation
    sampled_questions = sampled.to_dict('records')
    
    st.session_state.quiz = {
        "emp_id": str(emp_id),
        "emp_name": str(emp_name),
        "standard": str(standard),
        "total": int(total),
        "rows": sampled,
        "sampled_questions": sampled_questions,  # Add this for PDF generation
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

def append_result(emp_id, emp_name, total, right, wrong, criteria_pct, status, test_type, answers_json=None):
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
                'DATE / TIME': now,
                'ANSWERS': str(answers_json) if answers_json else ''
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
                elif 'ANSWER' in header_upper and header_upper != 'CORRECT ANSWER' and header_upper != 'WRONG ANSWER':
                    new_row.append(str(answers_json) if answers_json else '')
                else:
                    new_row.append('')
        else:
            new_row = [
                str(emp_id), str(emp_name), int(total), int(right), int(wrong),
                f"{pct:.2f}%", f"{criteria_pct:.0f}%", str(status), str(test_type), now,
                str(answers_json) if answers_json else ''
            ]

        worksheet.append_row(new_row)
        st.success("Results saved to Google Sheet.")
        return True, "", now
        
    except Exception as e:
        st.error(f"Error saving results: {str(e)}")
        return False, str(e), None

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
    
    # Use form to enable Enter key submission
    with st.form("admin_login_form", clear_on_submit=False):
        username = st.text_input("Username", key="admin_username")
        password = st.text_input("Password", type="password", key="admin_password")
        submitted = st.form_submit_button("Login", use_container_width=True)
        
        if submitted:
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
        
        # Create unique keys for filters
        id_key = f"emp_id_filter_{st.session_state.filter_reset_counter}"
        name_key = f"emp_name_filter_{st.session_state.filter_reset_counter}"
        
        # Initialize session state if not exists
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
            # Sort IDs numerically (not alphabetically) - convert to int for sorting, back to str for display
            unique_ids = results_df["ID"].astype(str).unique().tolist()
            try:
                # Try numeric sort first (works if all IDs are numeric)
                sorted_ids = sorted(unique_ids, key=lambda x: int(x) if x.isdigit() else float('inf'))
            except:
                # Fallback to alphabetic sort if some IDs are not numeric
                sorted_ids = sorted(unique_ids)
            employee_ids = ["All"] + sorted_ids
            selected_emp_id = st.selectbox(
                "Filter by Employee ID", 
                employee_ids, 
                key=id_key,
                on_change=sync_id_to_name
            )
        
        with filter_col2:
            employee_names = ["All"] + sorted(results_df["Name"].unique().tolist())
            selected_emp_name = st.selectbox(
                "Filter by Employee Name", 
                employee_names, 
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
        
        # Keep original data for internal use (with Answers column)
        filtered_df_original = results_df.copy()
        if selected_emp_id != "All":
            filtered_df_original = filtered_df_original[filtered_df_original["ID"].astype(str) == selected_emp_id]
        elif selected_emp_name != "All":
            filtered_df_original = filtered_df_original[filtered_df_original["Name"] == selected_emp_name]
        
        if selected_status != "All":
            filtered_df_original = filtered_df_original[filtered_df_original["Status"] == selected_status]
        if selected_test_type != "All":
            filtered_df_original = filtered_df_original[filtered_df_original["Test Type"] == selected_test_type]
        
        # Create display version without internal columns
        filtered_df = filtered_df_original.copy()
        columns_to_hide = ['ANSWERS', 'Answers', 'answers']
        for col in columns_to_hide:
            if col in filtered_df.columns:
                filtered_df = filtered_df.drop(columns=[col])

        if selected_emp_id != "All" or selected_emp_name != "All":
            display_name = selected_emp_name if selected_emp_name != "All" else id_name_mapping.get(selected_emp_id, "Unknown")
            display_id = selected_emp_id if selected_emp_id != "All" else name_id_mapping.get(selected_emp_name, "Unknown")
            st.info(f"🔗 **Selected Employee**: ID: {display_id} | Name: {display_name}")

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
        
        # Individual Test Answer Sheets
        st.markdown("---")
        st.subheader("📋 Download Test Answer Sheets")
        
        if selected_emp_id != "All" or selected_emp_name != "All":
            if selected_emp_id != "All":
                emp_filtered = filtered_df_original[filtered_df_original["ID"].astype(str) == selected_emp_id]
                emp_name_display = id_name_mapping.get(selected_emp_id, selected_emp_id)
                emp_id_display = selected_emp_id
            else:
                emp_filtered = filtered_df_original[filtered_df_original["Name"] == selected_emp_name]
                emp_name_display = selected_emp_name
                emp_id_display = name_id_mapping.get(selected_emp_name, "Unknown")
            
            if not emp_filtered.empty:
                st.info(f"Showing {len(emp_filtered)} test(s) for employee: **{emp_name_display}** (ID: {emp_id_display})")
                emp_filtered = emp_filtered.sort_values("Date / Time", ascending=False).reset_index(drop=True)
                
                for idx, test_row in emp_filtered.iterrows():
                    with st.expander(f"**Test {idx+1}:** {test_row['Test Type']} - {test_row['Date / Time']} ({test_row['Status']})"):
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            st.json({
                                "Employee ID": test_row['ID'],
                                "Employee Name": test_row['Name'],
                                "Standard": test_row['Test Type'],
                                "Total Questions": test_row['Total'],
                                "Correct": test_row['Right'],
                                "Wrong": test_row['Wrong'],
                                "Percentage": f"{test_row['Percentage']:.1f}%",
                                "Passing Criteria": f"{test_row['Criteria']}%",
                                "Status": test_row['Status'],
                                "Completed": test_row['Date / Time']
                            })
                        
                        with col2:
                            # Generate answer sheet button
                            if st.button(f"📥 Download Answer Sheet", key=f"dl_sheet_{idx}_{test_row['ID']}", use_container_width=True):
                                with st.spinner("Generating answer sheet..."):
                                    # Get questions for this standard
                                    test_standard = test_row['Test Type']
                                    total_questions = int(test_row['Total'])
                                    
                                    # Load questions for this standard
                                    if test_standard == "Cumulative":
                                        test_questions = questions.copy()
                                    else:
                                        test_questions = questions[
                                            questions["Standard"].astype(str).str.strip().str.upper() 
                                            == str(test_standard).strip().upper()
                                        ]
                                    
                                    test_questions = test_questions.dropna(subset=["Question", "A", "B", "C", "D", "Answer"])
                                    
                                    if not test_questions.empty and len(test_questions) >= total_questions:
                                        # Try to load candidate answers from Google Sheet first
                                        answers_data = []
                                        questions_data = []
                                        try:
                                            import json
                                            # Check for Answers column (case insensitive)
                                            answers_col = None
                                            for col in test_row.index:
                                                if col.upper() == 'ANSWERS':
                                                    answers_col = col
                                                    break
                                            
                                            if answers_col and test_row[answers_col] and test_row[answers_col] != '[]':
                                                stored_answers = json.loads(test_row[answers_col])
                                                
                                                # Load questions in the same order as stored answers
                                                for ans in stored_answers:
                                                    qno = str(ans.get('qno', ''))
                                                    if qno:
                                                        # Find question with this Qno
                                                        matching_q = test_questions[test_questions['Qno'].astype(str) == qno]
                                                        if not matching_q.empty:
                                                            questions_data.append(matching_q.iloc[0].to_dict())
                                                            answers_data.append({
                                                                "choice": ans.get("choice", "Not Stored"),
                                                                "correct": ans.get("correct", ""),
                                                                "is_correct": ans.get("is_correct", False)
                                                            })
                                            else:
                                                # No answers stored - old test, load first N questions
                                                sampled_questions = test_questions.head(total_questions)
                                                questions_data = sampled_questions.to_dict('records')
                                                for q in questions_data:
                                                    answers_data.append({
                                                        "choice": "Not Stored",
                                                        "correct": q.get("Answer", ""),
                                                        "is_correct": False
                                                    })
                                        except Exception as e:
                                            # Fallback if answers parsing fails
                                            for q in questions_data:
                                                answers_data.append({
                                                    "choice": "Not Stored",
                                                    "correct": q.get("Answer", ""),
                                                    "is_correct": False
                                                })
                                        
                                        # Clean percentage values
                                        pct_value = str(test_row['Percentage']).replace('%', '').strip()
                                        criteria_value = str(test_row['Criteria']).replace('%', '').strip()
                                        
                                        pdf_path, pdf_filename = generate_test_sheet_pdf(
                                            emp_id=test_row['ID'],
                                            emp_name=test_row['Name'],
                                            standard=test_row['Test Type'],
                                            test_date=test_row['Date / Time'],
                                            questions_data=questions_data,
                                            answers_data=answers_data,
                                            right=int(test_row['Right']),
                                            wrong=int(test_row['Wrong']),
                                            total=int(test_row['Total']),
                                            pct=float(pct_value),
                                            criteria=float(criteria_value),
                                            status=test_row['Status']
                                        )
                                        
                                        if pdf_path and os.path.exists(pdf_path):
                                            with open(pdf_path, "rb") as pdf_file:
                                                pdf_data = pdf_file.read()
                                            
                                            st.success("✅ Answer sheet generated!")
                                            st.download_button(
                                                label="📥 Download PDF",
                                                data=pdf_data,
                                                file_name=pdf_filename,
                                                mime="application/pdf",
                                                key=f"final_dl_{idx}",
                                                use_container_width=True
                                            )
                                        else:
                                            st.error("Failed to generate PDF.")
                                    else:
                                        st.error(f"Not enough questions for {test_standard}.")
            else:
                st.warning("No test results found for the selected employee.")
        else:
            st.info("👆 **Select an Employee ID or Name** to view and download test answer sheets")
        
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
        # Certificate Generation - OLD (COMMENTED OUT)
        # ======================
        # st.markdown("---")
        # st.subheader("📜 Generate Certificates")
        # 
        # passed_results = results_df[results_df["Status"].astype(str).str.upper().eq("PASS")].copy()
        #
        # # Name filter list based on passed results
        # cert_employee_names = ["All"] + sorted(passed_results["Name"].dropna().unique().tolist())
        # selected_cert_name = st.selectbox(
        #     "Filter Certificates by Employee Name",
        #     cert_employee_names,
        #     index=0,
        #     key=f"cert_name_filter_{st.session_state.filter_reset_counter}"
        # )
        #
        # # ---- Normalizer for standard names ----
        # def _norm(s: str) -> str:
        #     s = str(s or "").upper().strip()
        #     # Remove edition/volume information to match core standard name
        #     # e.g., "DS-1 3rd Volume 5th Edition" -> "DS 1"
        #     s = re.sub(r'\d+(ST|ND|RD|TH)\s+(VOLUME|EDITION)', '', s, flags=re.IGNORECASE)
        #     s = re.sub(r'(VOLUME|EDITION)\s+\d+', '', s, flags=re.IGNORECASE)
        #     s = s.replace("&", " ")
        #     s = s.replace("-", " ")
        #     s = re.sub(r"\s+", " ", s)
        #     s = s.replace("CUMMULATIVE", "CUMULATIVE")
        #     return s.strip()
        #
        # required_norm = {"DS 1", "CUMULATIVE", "API RP 7G 2", "API SPEC 5CT 5A5"}
        # template_map = {
        #     "DS 1": "Ds-1_template",
        #     "CUMULATIVE": "Cumulative_template",
        #     "API RP 7G 2": "API RP 7G-2_template",
        #     "API SPEC 5CT 5A5": "API SPEC 5CT & 5A5_template",
        # }
        # required_in_order = ["DS 1", "CUMULATIVE", "API RP 7G 2", "API SPEC 5CT 5A5"]
        #
        # ===== COMMENTED OUT: Old Certificate Generation (4 required standards) =====
        # if st.button("Generate Certificates for Qualifying Employees", use_container_width=True):
        #     # Normalize helper column
        #     passed_results["Test Type (norm)"] = passed_results["Test Type"].map(_norm)
        #     # Parse date to pick sensible rows when needed
        #     passed_results["_parsed_dt"] = pd.to_datetime(
        #         passed_results["Date / Time"], errors="coerce", dayfirst=True
        #     )
        #
        #     grouped = passed_results.groupby("Name", dropna=False)
        #     qualifying_rows = []
        #     for name, group in grouped:
        #         passed_set = set(group["Test Type (norm)"].dropna().tolist())
        #         if required_norm.issubset(passed_set):
        #             # Prefer CUMULATIVE row; else latest among required
        #             cum_row = group[group["Test Type (norm)"] == "CUMULATIVE"]
        #             if not cum_row.empty:
        #                 pick = cum_row.iloc[0]
        #             else:
        #                 req_rows = group[group["Test Type (norm)"].isin(required_norm)].copy()
        #                 req_rows = req_rows.sort_values("_parsed_dt", ascending=False)
        #                 if req_rows.empty:
        #                     continue
        #                 pick = req_rows.iloc[0]
        #             qualifying_rows.append(pick)
        #
        #     qualifying_df = pd.DataFrame(qualifying_rows)
        #
        #     # Guard before indexing
        #     if qualifying_df.empty or "Name" not in qualifying_df.columns:
        #         st.warning("No qualifying candidates found yet for the selected filters.")
        #     else:
        #         # Optional filter by selected name
        #         if selected_cert_name != "All":
        #             qualifying_df = qualifying_df[qualifying_df["Name"].astype(str) == str(selected_cert_name)]
        #
        #         if qualifying_df.empty:
        #             st.warning("Candidate is ineligible as not all required standards are passed.")
        #         else:
        #             certificate_files = []
        #             # Generate four certificates per candidate, filling the row values from each standard's own result
        #             for _, row_pick in qualifying_df.iterrows():
        #                 emp_id = row_pick["ID"]
        #                 emp_name = row_pick["Name"]
        #                 person_all = passed_results[passed_results["Name"] == emp_name].copy()
        #
        #                 # For each required standard, find that row and render with that standard's Percentage/Criteria
        #                 for norm_std in required_in_order:
        #                     std_row = person_all[person_all["Test Type (norm)"] == norm_std]
        #                     if std_row.empty:
        #                         continue  # safety
        #                     r = std_row.iloc[0]
        #
        #                     standard_text  = str(r["Test Type"]).strip()
        #                     # Normalize % fields
        #                     pct_val = r["Percentage"]
        #                     try:
        #                         pct_val_num = float(str(pct_val).replace("%","").strip())
        #                         percentage_text = f"{pct_val_num:.0f}%"
        #                     except:
        #                         percentage_text = str(pct_val) if str(pct_val).strip().endswith("%") else f"{str(pct_val).strip()}%"
        #
        #                     crit_val = r["Criteria"]
        #                     try:
        #                         crit_val_num = float(str(crit_val).replace("%","").strip())
        #                         criteria_text = f"{crit_val_num:.0f}%"
        #                     except:
        #                         criteria_text = str(crit_val) if str(crit_val).strip().endswith("%") else f"{str(crit_val).strip()}%"
        #
        #                     template_type = template_map.get(norm_std)
        #                     if not template_type:
        #                         continue
        #
        #                     certificate_path, certificate_filename = generate_certificate(
        #                         emp_id=emp_id,
        #                         emp_name=emp_name,
        #                         test_date=str(r["Date / Time"]),   # used for filename; also used if "Date of Certification" exists
        #                         status=str(r["Status"]),
        #                         template_type=template_type,
        #                         standard_text=standard_text,
        #                         percentage_text=percentage_text,
        #                         criteria_text=criteria_text,
        #                         skip_dates=True  # your templates don't have validity/date; but date/cert no (if present) are replaced at fs=18
        #                     )
        #                     if certificate_path:
        #                         certificate_files.append((certificate_path, certificate_filename))
        #
        #             if certificate_files:
        #                 zip_buffer = io.BytesIO()
        #                 with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        #                     for cert_path, cert_filename in certificate_files:
        #                         zipf.write(cert_path, cert_filename)
        #
        #                 zip_buffer.seek(0)
        #                 filename_suffix = selected_cert_name if selected_cert_name != "All" else "all_qualifying"
        #                 st.download_button(
        #                     label=f"Download Certificates (ZIP) for {filename_suffix}",
        #                     data=zip_buffer,
        #                     file_name=f"certificates_{filename_suffix}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        #                     mime="application/zip",
        #                     use_container_width=True
        #                 )
        #             else:
        #                 st.error("Failed to generate any certificates. Check templates and permissions.")
        
        # ======================
        # Individual Certificate Generation (for any passed test)
        # ======================
        st.markdown("---")
        st.subheader("📜 Generate Certificate")
        # st.info("Generate certificate for any single passed test, even if the employee hasn't completed all 4 required standards.")
        
        # Required data for certificate generation
        passed_results = results_df[results_df["Status"].astype(str).str.upper().eq("PASS")].copy()
        
        # Normalizer for standard names
        def _norm(s: str) -> str:
            s = str(s or "").upper().strip()
            s = re.sub(r'\d+(ST|ND|RD|TH)\s+(VOLUME|EDITION)', '', s, flags=re.IGNORECASE)
            s = re.sub(r'(VOLUME|EDITION)\s+\d+', '', s, flags=re.IGNORECASE)
            s = s.replace("&", " ")
            s = s.replace("-", " ")
            s = re.sub(r"\s+", " ", s)
            s = s.replace("CUMMULATIVE", "CUMULATIVE")
            return s.strip()
        
        template_map = {
            "DS 1": "Ds-1_template",
            "CUMULATIVE": "Cumulative_template",
            "API RP 7G 2": "API RP 7G-2_template",
            "API SPEC 5CT 5A5": "API SPEC 5CT & 5A5_template",
            "MPT GENERAL": "MT_template",
            "MPT SPECIFIC": "MT_template",
            "PENETRANT TESTING GENERAL": "PT_template",
            "PENETRANT TESTING SPECIFIC": "PT_template",
        }
        
        # MPT/PT pairs for combined certificate validation
        mpt_pt_pairs = {
            "MT": ["MPT (GENERAL)", "MPT (SPECIFIC)"],
            "PT": ["PENETRANT TESTING (GENERAL)", "PENETRANT TESTING (SPECIFIC)"]
        }
        
        # Filter for individual certificate
        ind_cert_col1, ind_cert_col2 = st.columns(2)
        
        with ind_cert_col1:
            ind_cert_names = ["Select Employee"] + sorted(passed_results["Name"].dropna().unique().tolist())
            selected_ind_name = st.selectbox(
                "Select Employee Name",
                ind_cert_names,
                index=0,
                key=f"ind_cert_name_{st.session_state.filter_reset_counter}"
            )
        
        with ind_cert_col2:
            if selected_ind_name != "Select Employee":
                # Get all passed tests for this employee
                emp_passed_tests = passed_results[passed_results["Name"] == selected_ind_name].copy()
                emp_passed_tests = emp_passed_tests.sort_values("Date / Time", ascending=False)
                
                # Create options with MPT/PT combined certificates
                def _norm(txt):
                    return str(txt).upper().strip()
                
                # Get unique test types
                unique_tests = emp_passed_tests.drop_duplicates(subset=["Test Type"], keep="first")
                individual_tests = unique_tests["Test Type"].tolist()
                
                # Check for MPT/PT eligibility (need both General + Specific)
                passed_norm_set = set(_norm(t) for t in individual_tests)
                test_options = ["Select Test"]
                
                # Check if eligible for combined MPT certificate
                if "MPT GENERAL" in passed_norm_set and "MPT SPECIFIC" in passed_norm_set:
                    test_options.append("MPT Certificate")
                
                # Check if eligible for combined PT certificate
                if "PENETRANT TESTING GENERAL" in passed_norm_set and "PENETRANT TESTING SPECIFIC" in passed_norm_set:
                    test_options.append("PT Certificate")
                
                # Add individual tests (skip General/Specific if combined cert available)
                for test in individual_tests:
                    norm_test = _norm(test)
                    # Skip individual MPT tests if combined cert available
                    if norm_test in ["MPT GENERAL", "MPT SPECIFIC"] and "MPT Certificate" in test_options:
                        continue
                    # Skip individual PT tests if combined cert available
                    if norm_test in ["PENETRANT TESTING GENERAL", "PENETRANT TESTING SPECIFIC"] and "PT Certificate" in test_options:
                        continue
                    test_options.append(test)
                
                # Add "All" option if multiple certificates available (excluding "Select Test")
                if len(test_options) > 2:  # More than just "Select Test" and one certificate
                    test_options.insert(1, "All")  # Insert after "Select Test"
                
                selected_test_option = st.selectbox(
                    "Select Test",
                    test_options,
                    index=0,
                    key=f"ind_cert_test_{st.session_state.filter_reset_counter}"
                )
            else:
                selected_test_option = "Select Test"
                st.selectbox("Select Test", ["Select Employee First"], index=0, disabled=True)
        
        # Generate individual certificate button
        if selected_ind_name != "Select Employee" and selected_test_option not in ["Select Test", "Select Employee First"]:
            # Change button text based on selection
            button_text = "Generate All Certificates (ZIP)" if selected_test_option == "All" else "Generate Certificate for Selected Test"
            
            if st.button(button_text, use_container_width=True):
                if selected_test_option == "All":
                    # Generate certificates for all passed tests of this employee
                    certificate_files = []
                    
                    for _, r in emp_passed_tests.iterrows():
                        # Normalize test type to find template
                        norm_test_type = _norm(r["Test Type"])
                        template_type = template_map.get(norm_test_type)
                        
                        if not template_type:
                            template_type = "Generic_template"
                        
                        # Prepare data
                        emp_id = r["ID"]
                        emp_name = r["Name"]
                        standard_text = str(r["Test Type"]).strip()
                        
                        # Normalize percentage
                        pct_val = r["Percentage"]
                        try:
                            pct_val_num = float(str(pct_val).replace("%","").strip())
                            percentage_text = f"{pct_val_num:.0f}%"
                        except:
                            percentage_text = str(pct_val) if str(pct_val).strip().endswith("%") else f"{str(pct_val).strip()}%"
                        
                        # Normalize criteria
                        crit_val = r["Criteria"]
                        try:
                            crit_val_num = float(str(crit_val).replace("%","").strip())
                            criteria_text = f"{crit_val_num:.0f}%"
                        except:
                            criteria_text = str(crit_val) if str(crit_val).strip().endswith("%") else f"{str(crit_val).strip()}%"
                        
                        # Generate certificate
                        certificate_path, certificate_filename = generate_certificate(
                            emp_id=emp_id,
                            emp_name=emp_name,
                            test_date=str(r["Date / Time"]),
                            status=str(r["Status"]),
                            template_type=template_type,
                            standard_text=standard_text,
                            percentage_text=percentage_text,
                            criteria_text=criteria_text,
                            skip_dates=True
                        )
                        
                        if certificate_path:
                            certificate_files.append((certificate_path, certificate_filename))
                    
                    if certificate_files:
                        # Create ZIP file
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                            for cert_path, cert_filename in certificate_files:
                                zipf.write(cert_path, cert_filename)
                        
                        zip_buffer.seek(0)
                        st.success(f"✅ Generated {len(certificate_files)} certificate(s) for {selected_ind_name}!")
                        st.download_button(
                            label=f"Download All Certificates (ZIP) - {selected_ind_name}",
                            data=zip_buffer,
                            file_name=f"certificates_{selected_ind_name}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                    else:
                        st.error("❌ Failed to generate any certificates. Please check template availability.")
                
                else:
                    # Single certificate generation (including MPT/PT combined)
                    if selected_test_option in ["MPT Certificate", "PT Certificate"]:
                        # Combined MPT/PT certificate (2-row table)
                        cert_type = "MT" if selected_test_option == "MPT Certificate" else "PT"
                        template_type = "MT_template" if cert_type == "MT" else "PT_template"
                        
                        # Get General and Specific test rows
                        if cert_type == "MT":
                            general_rows = emp_passed_tests[_norm(emp_passed_tests["Test Type"]) == "MPT GENERAL"]
                            specific_rows = emp_passed_tests[_norm(emp_passed_tests["Test Type"]) == "MPT SPECIFIC"]
                        else:  # PT
                            general_rows = emp_passed_tests[_norm(emp_passed_tests["Test Type"]) == "PENETRANT TESTING GENERAL"]
                            specific_rows = emp_passed_tests[_norm(emp_passed_tests["Test Type"]) == "PENETRANT TESTING SPECIFIC"]
                        
                        if general_rows.empty or specific_rows.empty:
                            st.error(f"❌ Both General and Specific tests must be passed for {selected_test_option}")
                        else:
                            # Get latest of each
                            general_row = general_rows.sort_values("Date / Time", ascending=False).iloc[0]
                            specific_row = specific_rows.sort_values("Date / Time", ascending=False).iloc[0]
                            
                            # Use latest date between the two tests
                            general_date = pd.to_datetime(general_row["Date / Time"], dayfirst=True)
                            specific_date = pd.to_datetime(specific_row["Date / Time"], dayfirst=True)
                            latest_date = max(general_date, specific_date)
                            
                            # Extract data
                            emp_id = general_row["ID"]
                            emp_name = general_row["Name"]
                            
                            # General test data
                            general_standard = str(general_row["Test Type"]).strip()
                            general_pct = general_row["Percentage"]
                            try:
                                general_pct_num = float(str(general_pct).replace("%","").strip())
                                general_percentage = f"{general_pct_num:.0f}%"
                            except:
                                general_percentage = str(general_pct) if str(general_pct).strip().endswith("%") else f"{str(general_pct).strip()}%"
                            
                            general_crit = general_row["Criteria"]
                            try:
                                general_crit_num = float(str(general_crit).replace("%","").strip())
                                general_criteria = f"{general_crit_num:.0f}%"
                            except:
                                general_criteria = str(general_crit) if str(general_crit).strip().endswith("%") else f"{str(general_crit).strip()}%"
                            
                            # Specific test data
                            specific_standard = str(specific_row["Test Type"]).strip()
                            specific_pct = specific_row["Percentage"]
                            try:
                                specific_pct_num = float(str(specific_pct).replace("%","").strip())
                                specific_percentage = f"{specific_pct_num:.0f}%"
                            except:
                                specific_percentage = str(specific_pct) if str(specific_pct).strip().endswith("%") else f"{str(specific_pct).strip()}%"
                            
                            specific_crit = specific_row["Criteria"]
                            try:
                                specific_crit_num = float(str(specific_crit).replace("%","").strip())
                                specific_criteria = f"{specific_crit_num:.0f}%"
                            except:
                                specific_criteria = str(specific_crit) if str(specific_crit).strip().endswith("%") else f"{str(specific_crit).strip()}%"
                            
                            # Generate combined certificate
                            certificate_path, certificate_filename = generate_mpt_pt_certificate(
                                emp_id=emp_id,
                                emp_name=emp_name,
                                test_date=str(latest_date),
                                template_type=template_type,
                                general_standard=general_standard,
                                general_percentage=general_percentage,
                                general_criteria=general_criteria,
                                specific_standard=specific_standard,
                                specific_percentage=specific_percentage,
                                specific_criteria=specific_criteria
                            )
                            
                            if certificate_path:
                                st.success(f"✅ Generated {selected_test_option} for {emp_name}!")
                                with open(certificate_path, "rb") as f:
                                    st.download_button(
                                        label=f"Download {selected_test_option}",
                                        data=f,
                                        file_name=certificate_filename,
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                            else:
                                st.error(f"❌ Failed to generate {selected_test_option}. Check template availability.")
                    
                    else:
                        # Regular single certificate
                        # Find the selected test row (latest one for this test type)
                        selected_test_row = emp_passed_tests[emp_passed_tests["Test Type"] == selected_test_option]
                        selected_test_row = selected_test_row.sort_values("Date / Time", ascending=False)
                        
                        if not selected_test_row.empty:
                            r = selected_test_row.iloc[0]
                            
                            # Normalize test type to find template
                            norm_test_type = _norm(r["Test Type"])
                            template_type = template_map.get(norm_test_type)
                            
                            # If no template in map, use a generic template name
                            # generate_certificate function will handle fallback automatically
                            if not template_type:
                                template_type = "Generic_template"  # Will fallback to available template
                            
                            # Prepare data
                            emp_id = r["ID"]
                            emp_name = r["Name"]
                            standard_text = str(r["Test Type"]).strip()
                            
                            # Normalize percentage
                            pct_val = r["Percentage"]
                            try:
                                pct_val_num = float(str(pct_val).replace("%","").strip())
                                percentage_text = f"{pct_val_num:.0f}%"
                            except:
                                percentage_text = str(pct_val) if str(pct_val).strip().endswith("%") else f"{str(pct_val).strip()}%"
                            
                            # Normalize criteria
                            crit_val = r["Criteria"]
                            try:
                                crit_val_num = float(str(crit_val).replace("%","").strip())
                                criteria_text = f"{crit_val_num:.0f}%"
                            except:
                                criteria_text = str(crit_val) if str(crit_val).strip().endswith("%") else f"{str(crit_val).strip()}%"
                            
                            # Generate certificate
                            certificate_path, certificate_filename = generate_certificate(
                                emp_id=emp_id,
                                emp_name=emp_name,
                                test_date=str(r["Date / Time"]),
                                status=str(r["Status"]),
                                template_type=template_type,
                                standard_text=standard_text,
                                percentage_text=percentage_text,
                                criteria_text=criteria_text,
                                skip_dates=True
                            )
                        
                        if certificate_path:
                            # Read the file and provide download
                            with open(certificate_path, "rb") as f:
                                cert_data = f.read()
                            
                            st.success(f"✅ Certificate generated successfully for {emp_name}!")
                            st.download_button(
                                label=f"Download Certificate {emp_name}",
                                data=cert_data,
                                file_name=certificate_filename,
                                mime="application/pdf",
                                use_container_width=True
                            )
                        else:
                            st.error("❌ Failed to generate certificate. Please check template availability.")
        
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
    if "Cumulative" not in options:
        options = ["Cumulative"] + options
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
                    # Debug: Show answer saved
                    # st.success(f"✓ Answer saved for Q{current_qid+1}: {choice} ({'Correct' if is_correct else 'Wrong'})")
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
            # Get questions data
            questions_data = qstate.get("sampled_questions", [])
            
            # Prepare answers data from quiz state
            answers_data = []
            answers_json_list = []  # For storing in Google Sheet
            
            for idx, q in enumerate(questions_data):
                # Get question ID from Qno field
                qno = str(q.get("Qno", idx))  # Fallback to index if Qno not found
                # Use index, not ID, to retrieve from qstate["answers"]
                answer_dict = qstate.get("answers", {}).get(idx, None)
                
                if answer_dict and isinstance(answer_dict, dict):
                    # User answered this question
                    user_choice = answer_dict.get("choice", "Not Attempted")
                    correct_ans = answer_dict.get("correct", "")
                    is_correct = answer_dict.get("is_correct", False)
                    
                    answers_data.append({
                        "choice": user_choice,
                        "correct": correct_ans,
                        "is_correct": is_correct
                    })
                    
                    # Store as compact format for Google Sheet (use Qno as identifier)
                    answers_json_list.append({
                        "qno": qno,
                        "choice": user_choice,
                        "correct": correct_ans,
                        "is_correct": is_correct
                    })
                else:
                    # Question not attempted
                    answers_data.append({
                        "choice": "Not Attempted",
                        "correct": q.get("Answer", ""),
                        "is_correct": False
                    })
                    
                    answers_json_list.append({
                        "qno": qno,
                        "choice": "Not Attempted",
                        "correct": q.get("Answer", ""),
                        "is_correct": False
                    })
            
            # Convert answers to JSON string for storage
            import json
            answers_json = json.dumps(answers_json_list)
            
            # Save results to Google Sheet WITH answers
            ok, msg, test_date = append_result(
                qstate["emp_id"], qstate["emp_name"], total_q, right, wrong, 
                criteria, status, qstate["standard"], answers_json
            )
            
            # Generate answer sheet PDF automatically
            pdf_path = None
            pdf_filename = None
            try:
                
                # Generate PDF
                pdf_path, pdf_filename = generate_test_sheet_pdf(
                    emp_id=qstate["emp_id"],
                    emp_name=qstate["emp_name"],
                    standard=qstate["standard"],
                    test_date=test_date if test_date else datetime.datetime.now().strftime("%d-%m-%Y %I:%M:%S %p"),
                    questions_data=questions_data,
                    answers_data=answers_data,
                    right=right,
                    wrong=wrong,
                    total=total_q,
                    pct=pct,
                    criteria=criteria,
                    status=status
                )
                
                if pdf_path and os.path.exists(pdf_path):
                    st.success(f"✅ Test answer sheet generated: {pdf_filename}")
            except Exception as e:
                st.warning(f"⚠️ Could not generate answer sheet: {str(e)}")
            
            st.session_state["submitted"] = True
            st.session_state["submit_result"] = (ok, msg, right, wrong, total_q, pct, criteria, status, final_score, pdf_path, pdf_filename)
            st.rerun()

    if "submitted" in st.session_state:
        if "submit_result" in st.session_state:
            result_data = st.session_state["submit_result"]
            
            # Handle both old and new format
            if len(result_data) == 9:
                ok, msg, right, wrong, total_q, pct, criteria, status, final_score = result_data
                pdf_path = None
                pdf_filename = None
            else:
                ok, msg, right, wrong, total_q, pct, criteria, status, final_score, pdf_path, pdf_filename = result_data
            
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
            
            # Answer sheet saved for admin to download
            # if pdf_path and pdf_filename and os.path.exists(pdf_path):
            #     st.info("📋 Your test answer sheet has been saved. Contact admin to download.")

