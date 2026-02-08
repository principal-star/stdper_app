import json
from flask import Flask, render_template, request, jsonify, send_file
import gspread
import pandas as pd
import re
#from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import threading
import time
from gspread.exceptions import APIError
from playwright.sync_api import sync_playwright

from flask import Flask, render_template
from flask import send_file

import pdfkit
from google.oauth2.service_account import Credentials
import tempfile
import os
from flask import request, send_file
from playwright.sync_api import sync_playwright
import tempfile
import os
app = Flask(__name__)
"""
@app.route("/")
def home():
    #return render_template("home_final1print.html")
    return render_template("home_final1print.html", sheets=sheets, analytics_fields=ANALYTICS_FIELDS)
"""
if __name__ == "__main__":
    app.run()

# ---------------- GOOGLE SHEET CONFIG ---------------- #

SHEET_URL = "https://docs.google.com/spreadsheets/d/1QbsOeLrqJd6w6Wn-3ycnjyEtiXHONWFnhnpU_NIIhp4/edit?gid=1755838128#gid=1755838128"
#SHEET_URL = "https://docs.google.com/spreadsheets/d/1Gk0W2XfdCAZ3MmsZcPqhsa0h6AktkF71SSNEfThRjGM/edit?gid=0#gid=0"
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

#creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)




scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)

#client = gspread.authorize(creds)


client = gspread.authorize(creds)
workbook = client.open_by_url(SHEET_URL)

# ---------------- In-memory cache ---------------- #
# cache structure: { sheet_title: { 'df': pandas.DataFrame, 'loaded_at': timestamp } }
SHEET_CACHE = {}
CACHE_LOCK = threading.Lock()
CACHE_TTL = 60 * 60 * 6  # optional TTL: 6 hours (not auto-enforced here, endpoint to refresh provided)

SEM_COLORS = {
    "SEM1": "#e3f2fd",
    "SEM2": "#e8f5e9",
    "SEM3": "#fff3e0",
    "SEM4": "#fce4ec",
    "SEM5": "#ede7f6",
    "SEM6": "#e0f7fa",
    "SEM7": "#f1f8e9",
    "SEM8": "#fbe9e7"
}

# Analytics dropdown fields (same as in frontend)
ANALYTICS_FIELDS = [
    "Gender",
    "HSC Marks",
    "Cut off",
    "First Graduate (Y/N)",
    "GQ/MQ",
    "Hosteller/Dayscholar/Outsiders",
    "Native district",
    "Belongs to 7.5"
]

# ---------------- Utility functions ---------------- #

GRADE_POINT_MAP = {
    "O": 10, "A+": 9, "A": 8, "B+": 7, "B": 6,
    "RA": 5, "U": 5, "FAIL": 5, "F": 5, "ABSENT": 5
}
def convert_drive_link(url):
    if not url or not isinstance(url, str):
        return None

    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        return None

    file_id = match.group(1)
    return f"https://drive.google.com/uc?export=view&id={file_id}"


"""
def convert_drive_link(url: str):
    
    if not url:
        return None
    s = str(url).strip()
    # Remove angle brackets and whitespace
    s = s.strip("<>").strip()
    # If it's a HYPERLINK formula, extract first argument
    m = re.search(r'HYPERLINK\("([^"]+)"', s, re.IGNORECASE)
    if m:
        s = m.group(1).strip("<>").strip()
    # If /d/FILEID/ pattern
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", s)
    if m:
        file_id = m.group(1)
        return f"https://drive.google.com/uc?export=view&id={file_id}"
    # If id=FILEID parameter
    m = re.search(r"id=([a-zA-Z0-9_-]+)", s)
    if m:
        file_id = m.group(1)
        return f"https://drive.google.com/uc?export=view&id={file_id}"
    # Already uc?export=view
    if "uc?export=view" in s:
        return s
    return s
"""
"""

def load_sheet_to_cache(sheet_title: str):
        with CACHE_LOCK:
        try:
            ws = workbook.worksheet(sheet_title)
            values = ws.get_all_values()

            if not values or len(values) < 2:
                df = pd.DataFrame()
            else:
                raw_headers = values[0]

                # Fix empty / duplicate headers
                headers = []
                seen = {}

                for i, h in enumerate(raw_headers):
                    h = h.strip() if h else f"Unnamed_{i}"
                    if h in seen:
                        seen[h] += 1
                        h = f"{h}_{seen[h]}"
                    else:
                        seen[h] = 0
                    headers.append(h)

                df = pd.DataFrame(values[1:], columns=headers)

            
            records = ws.get_all_records()
            df = pd.DataFrame(records)
            
            
            # Normalize column names to str
            df.columns = [str(c) for c in df.columns]
            SHEET_CACHE[sheet_title] = {"df": df, "loaded_at": time.time()}
            return SHEET_CACHE[sheet_title]
        except APIError as e:
            # re-raise to caller
            raise
"""
def load_sheet_to_cache(sheet_title):
    global SHEET_CACHE

    if sheet_title in SHEET_CACHE:
        return SHEET_CACHE[sheet_title]

    ws = workbook.worksheet(sheet_title)

    values = ws.get_all_values()

    if not values or len(values) < 2:
        df = pd.DataFrame()
    else:
        raw_headers = values[0]

        headers = []
        seen = {}

        for i, h in enumerate(raw_headers):
            h = h.strip() if h else f"Unnamed_{i}"
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 0
            headers.append(h)

        df = pd.DataFrame(values[1:], columns=headers)

    SHEET_CACHE[sheet_title] = {
        "df": df
    }

    return SHEET_CACHE[sheet_title]


def get_sheet_df(sheet_title: str):
    """Return cached DataFrame for a sheet, loading it if necessary."""
    with CACHE_LOCK:
        entry = SHEET_CACHE.get(sheet_title)
    if entry is None:
        # load it
        try:
            return load_sheet_to_cache(sheet_title)["df"]
        except APIError as e:
            raise
    else:
        return entry["df"]


def list_sheets():
    """Return list of sheet titles (workbook worksheets)."""
    return [ws.title for ws in workbook.worksheets()]


def normalize_colname(text: str):
    return ''.join(ch.lower() for ch in str(text) if ch.isalnum())


def compute_weighted_cgpa_from_row(row: dict):
    total_points = 0
    total_credits = 0
    for key, val in row.items():
        m = re.search(r"(sem\d+)_.*_(\d+)$", key.lower())
        if not m:
            continue
        credits = int(m.group(2))
        grade = str(val).strip().upper()
        if grade in GRADE_POINT_MAP:
            total_points += GRADE_POINT_MAP[grade] * credits
            total_credits += credits
    if total_credits == 0:
        return None
    return round(total_points / total_credits, 2)


def compute_sem_gpa_from_row(row: dict):
    sem_totals = {}
    for key, val in row.items():
        m = re.search(r"(sem\d+)_.*_(\d+)$", key.lower())
        if not m:
            continue
        sem = m.group(1).capitalize()
        credits = int(m.group(2))
        grade = str(val).strip().upper()
        sem_totals.setdefault(sem, {"points": 0, "credits": 0})
        if grade in GRADE_POINT_MAP:
            sem_totals[sem]["points"] += GRADE_POINT_MAP[grade] * credits
            sem_totals[sem]["credits"] += credits
    labels = []
    values = []
    for sem in sorted(sem_totals.keys(), key=lambda s: int(s[3:])):
        pts = sem_totals[sem]["points"]
        cr = sem_totals[sem]["credits"]
        gpa = round(pts / cr, 2) if cr > 0 else None
        labels.append(sem)
        values.append(gpa)
    return labels, values


def compute_semester_arrears_from_row(row: dict):
    arrear_keywords = {"RA", "U", "UA", "F", "FAIL", "ABSENT"}
    sem_arrears = {}
    for key, val in row.items():
        m = re.search(r"(sem\d+)_.*_(\d+)$", key.lower())
        if not m:
            continue
        sem = m.group(1).capitalize()
        grade = str(val).strip().upper()
        sem_arrears[sem] = sem_arrears.get(sem, 0) + (1 if grade in arrear_keywords else 0)
    items = sorted(sem_arrears.items(), key=lambda x: int(x[0][3:]))
    labels = [k for k, _ in items]
    values = [v for _, v in items]
    return labels, values


# ---------------- ROUTES ---------------- #
@app.route("/test")
def test():
    return "Flask is working"
    
@app.route("/")
def home():
    sheets = list_sheets()
    dashboard = {}
    return render_template("home_final1print.html", sheets=sheets, analytics_fields=ANALYTICS_FIELDS,
        dashboard=dashboard)


@app.route("/refresh_cache", methods=["POST", "GET"])
def refresh_cache():
    """Refresh cache for one sheet or all sheets. Admin endpoint."""
    sheet = request.args.get("sheet")
    try:
        if sheet:
            load_sheet_to_cache(sheet)
            return jsonify({"status": "ok", "sheet": sheet}), 200
        else:
            # load all sheets
            sheets = list_sheets()
            for s in sheets:
                load_sheet_to_cache(s)
            return jsonify({"status": "ok", "sheets_loaded": sheets}), 200
    except APIError as e:
        return jsonify({"error": "Google API error", "detail": str(e)}), 500


@app.route("/students", methods=["POST"])
def students():
    """Return list of student names from cached sheet (load sheet if missing)."""
    sheet = request.form.get("sheet")
    if not sheet:
        return jsonify({"error": "sheet parameter required"}), 400
    try:
        df = get_sheet_df(sheet)
    except APIError as e:
        return jsonify({"error": "Google API error", "detail": str(e)}), 500

    # find name column
    name_col = next((c for c in df.columns if "name" in str(c).lower()), None)
    if name_col is None:
        return jsonify([])

    names = df[name_col].astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
    names = sorted(names)
    return jsonify(names)


@app.route("/student_details", methods=["POST"])
def student_details():
    """Return detailed data for a single student using cached DataFrame."""
    sheet = request.form.get("sheet")
    student = request.form.get("student")
    if not sheet or not student:
        return jsonify({"error": "sheet and student parameters required"}), 400

    try:
        df = get_sheet_df(sheet)
    except APIError as e:
        return jsonify({"error": "Google API error", "detail": str(e)}), 500

    name_col = next((c for c in df.columns if "name" in str(c).lower()), None)
    if name_col is None:
        return jsonify({"error": "Name column not found"}), 400

    df_copy = df.copy()
    df_copy["__match__"] = df_copy[name_col].astype(str).str.strip().str.lower()
    mask = df_copy["__match__"] == student.strip().lower()
    if not mask.any():
        return jsonify({"error": f"Student '{student}' not found"}), 404

    row = df_copy[mask].iloc[0].drop(labels="__match__")
    student_data = row.to_dict()
    # -------------------------------
    # Semester-wise subject grouping
    # -------------------------------
    semester_subjects = {}
    
    for col, val in student_data.items():
        m = re.search(r"(sem\d+)_([A-Za-z0-9]+)_\d+", str(col).lower())
        if not m:
            continue
    
        sem = m.group(1).upper()       # SEM1, SEM2, ...
        subject = m.group(2).upper()   # MA101, CS204, etc
        grade = str(val).strip().upper()
    
        semester_subjects.setdefault(sem, []).append({
            "subject": subject,
            "grade": grade
        })

    
    # photo detection & conversion
    photo_url = None
    for col in df.columns:
        if "photo" in str(col).lower() or "image" in str(col).lower():
            raw = row.get(col)
            photo_url = convert_drive_link(raw)
            break

    # overall arrears
    arrear_keywords = {"RA", "U", "UA", "F", "FAIL", "ABSENT"}
    arrears_count = sum(1 for v in student_data.values() if str(v).strip().upper() in arrear_keywords)

    # grade distribution (only semester subject columns)
    """
    grades = {}
    for k, v in student_data.items():
        if re.search(r"(sem\d+)_.*_(\d+)$", str(k).lower()):
            g = str(v).strip().upper()
            if g in GRADE_POINT_MAP:
                grades[g] = grades.get(g, 0) + 1
    """
    # semester-wise subjects and grades
    semester_results = {}
    
    for col, val in student_data.items():
        m = re.search(r"(sem\d+)_([A-Za-z0-9]+)", str(col), re.IGNORECASE)
        if not m:
            continue
    
        sem = m.group(1).upper()   # SEM1, SEM2...
        subject = m.group(2).upper()
        grades = str(val).strip().upper()
    
        semester_results.setdefault(sem, []).append({
            "subject": subject,
            "grade": grades
        })

    # semester arrears
    sem_arrears_labels, sem_arrears_values = compute_semester_arrears_from_row(student_data)

    # categorized details
    personal = {}
    parent = {}
    academic = {}
    other = {}
    for k, v in student_data.items():
        kl = str(k).lower()
        if any(t in kl for t in ["father", "mother", "guardian", "parent"]):
            parent[k] = v
        elif any(t in kl for t in ["sem", "gpa", "grade", "subject", "marks", "hsc", "cut off", "cutoff"]):
            academic[k] = v
        elif any(t in kl for t in ["name", "gender", "dob", "reg", "roll", "mobile", "email", "address", "native", "district"]):
            personal[k] = v
        else:
            other[k] = v

    # cgpa
    cgpa_value = compute_weighted_cgpa_from_row(student_data)

    # class CGPA list and rank
    cgpa_list = []
    for _, rec in df_copy.iterrows():
        rdict = rec.drop(labels="__match__").to_dict()
        cg = compute_weighted_cgpa_from_row(rdict)
        if cg is not None:
            cgpa_list.append(cg)
    class_size = len(cgpa_list)
    rank_value = None
    if cgpa_value is not None and class_size > 0:
        sorted_cg = sorted(cgpa_list, reverse=True)
        rank_value = sorted_cg.index(cgpa_value) + 1

    # semester GPA & rank per semester
    sem_gpa_labels, sem_gpa_values = compute_sem_gpa_from_row(student_data)

    # prepare per-student sem gpa mapping across class for ranking
    all_sem_gpa_dicts = []
    for _, rec in df_copy.iterrows():
        rdict = rec.drop(labels="__match__").to_dict()
        labs, vals = compute_sem_gpa_from_row(rdict)
        all_sem_gpa_dicts.append(dict(zip(labs, vals)))

    sem_rank_values = []
    for sem, my_gpa in zip(sem_gpa_labels, sem_gpa_values):
        if my_gpa is None:
            sem_rank_values.append(None)
            continue
        gpas_this_sem = [d.get(sem) for d in all_sem_gpa_dicts if d.get(sem) is not None]
        if not gpas_this_sem:
            sem_rank_values.append(None)
        else:
            sorted_list = sorted(gpas_this_sem, reverse=True)
            # protect against missing my_gpa in list (shouldn't happen)
            try:
                sem_rank_values.append(sorted_list.index(my_gpa) + 1)
            except ValueError:
                sem_rank_values.append(None)

    response = {
        "details": student_data,
        "arrears": arrears_count,
        "grades": grades,
        "sem_arrears_labels": sem_arrears_labels,
        "sem_arrears_values": sem_arrears_values,
        "sections": {
            "personal": personal,
            "parent": parent,
            "academic": academic,
            "other": other
        },
        "photo_url": photo_url,
        "cgpa": cgpa_value,
        "rank": rank_value,
        "class_size": class_size,
        "sem_gpa_labels": sem_gpa_labels,
        "sem_gpa_values": sem_gpa_values,
        "semester_subjects": semester_subjects,
        "semester_colors": SEM_COLORS,
        "sem_rank_values": sem_rank_values
    }

    return jsonify(response)


@app.route("/analytics_data", methods=["POST"])
def analytics_data():
    """Return analytics (bins or categories) for a given sheet and column using cached df."""
    sheet = request.form.get("sheet")
    column_label = request.form.get("column")
    if not sheet or not column_label:
        return jsonify({"error": "sheet and column parameters required"}), 400

    try:
        df = get_sheet_df(sheet)
    except APIError as e:
        return jsonify({"error": "Google API error", "detail": str(e)}), 500

    # find column that best matches (normalize both)
    target = normalize_colname(column_label)
    col_match = next((c for c in df.columns if target in normalize_colname(c)), None)
    if col_match is None:
        return jsonify({"error": f"No matching column for {column_label}"}), 400

    series = df[col_match].dropna()
    numeric = pd.to_numeric(series, errors="coerce")

    # HSC bins
    if "hsc" in target and numeric.notna().sum() > 0:
        values = numeric.dropna()
        bins = [
            ("<=300", (values <= 300).sum()),
            (">300–350", ((values > 300) & (values <= 350)).sum()),
            (">350–400", ((values > 350) & (values <= 400)).sum()),
            (">400–450", ((values > 400) & (values <= 450)).sum()),
            (">450", (values > 450).sum())
        ]
        labels, counts = zip(*bins)
        return jsonify({"labels": list(labels), "values": [int(c) for c in counts], "chart_type": "bar"})

    # Cutoff bins (rename detection by 'cut' prefix)
    if "cut" in target and numeric.notna().sum() > 0:
        values = numeric.dropna()
        bins = [
            ("<=80", (values <= 80).sum()),
            (">80–100", ((values > 80) & (values <= 100)).sum()),
            (">100–120", ((values > 100) & (values <= 120)).sum()),
            (">120–140", ((values > 120) & (values <= 140)).sum()),
            (">140–160", ((values > 140) & (values <= 160)).sum()),
            (">160–180", ((values > 160) & (values <= 180)).sum()),
            (">180", (values > 180).sum())
        ]
        labels, counts = zip(*bins)
        return jsonify({"labels": list(labels), "values": [int(c) for c in counts], "chart_type": "bar"})

    # Generic numeric -> 5 bins
    if numeric.notna().sum() > 0:
        values = numeric.dropna()
        binned = pd.cut(values, bins=5)
        counts = binned.value_counts().sort_index()
        return jsonify({"labels": [str(i) for i in counts.index], "values": counts.tolist(), "chart_type": "bar"})

    # Categorical -> value counts
    counts = series.astype(str).str.strip().value_counts()
    return jsonify({"labels": counts.index.tolist(), "values": counts.values.tolist(), "chart_type": "bar"})


@app.route("/batch_arrear_status", methods=["POST"])
def batch_arrear_status():
    """
    Compute arrear buckets for the whole sheet and return as labels/values.
    Buckets: 0,1,2,3,4,5,6-10,>10
    """
    sheet = request.form.get("sheet")
    if not sheet:
        return jsonify({"error": "sheet parameter required"}), 400

    try:
        df = get_sheet_df(sheet)
    except APIError as e:
        return jsonify({"error": "Google API error", "detail": str(e)}), 500

    # compute arrears per row efficiently by scanning sem columns
    sem_cols = [c for c in df.columns if re.search(r"(sem\d+)_.*_(\d+)$", str(c).lower())]
    if not sem_cols:
        return jsonify({"labels": [], "values": []})

    def row_arrears_count(row):
        count = 0
        for c in sem_cols:
            val = str(row.get(c, "")).strip().upper()
            if val in {"RA", "U", "UA", "F", "FAIL", "ABSENT"}:
                count += 1
        return count

    counts = df.apply(lambda r: row_arrears_count(r), axis=1).fillna(0).astype(int)

    buckets = {
        "0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0,
        "6–10": 0, ">10": 0
    }
    for v in counts:
        if v <= 5:
            buckets[str(v)] += 1
        elif v <= 10:
            buckets["6–10"] += 1
        else:
            buckets[">10"] += 1

    labels = list(buckets.keys())
    values = list(buckets.values())
    return jsonify({"labels": labels, "values": values, "chart_type": "bar"})


@app.route("/batch_top_bottom", methods=["POST"])
def batch_top_bottom():
    """Return top 5 students by CGPA and bottom 5 by arrears (most arrears)."""
    sheet = request.form.get("sheet")
    if not sheet:
        return jsonify({"error": "sheet parameter required"}), 400

    try:
        df = get_sheet_df(sheet)
    except APIError as e:
        return jsonify({"error": "Google API error", "detail": str(e)}), 500

    name_col = next((c for c in df.columns if "name" in str(c).lower()), None)
    if name_col is None:
        return jsonify({"error": "Name column not found"}), 400

    # compute CGPA for each row
    def cgpa_for_row(row):
        rdict = row.to_dict()
        return compute_weighted_cgpa_from_row(rdict)

    df2 = df.copy()
    df2["__cgpa__"] = df2.apply(lambda r: cgpa_for_row(r), axis=1)

    # compute arrears for each student
    sem_cols = [c for c in df.columns if re.search(r"(sem\d+)_.*_(\d+)$", str(c).lower())]

    def arrears_for_row(row):
        cnt = 0
        for c in sem_cols:
            val = str(row.get(c, "")).strip().upper()
            if val in {"RA", "U", "UA", "F", "FAIL", "ABSENT"}:
                cnt += 1
        return cnt

    df2["__arrears__"] = df2.apply(lambda r: arrears_for_row(r), axis=1)

    # top 5 by CGPA (desc). Keep name and cgpa
    top_df = df2.dropna(subset=["__cgpa__"]).sort_values("__cgpa__", ascending=False).head(5)
    top_list = [{"name": str(row[name_col]), "cgpa": round(row["__cgpa__"], 2)} for _, row in top_df.iterrows()]

    # bottom 5 by arrears (highest arrears first)
    bottom_df = df2.sort_values("__arrears__", ascending=False).head(5)
    bottom_list = [{"name": str(row[name_col]), "arrears": int(row["__arrears__"])} for _, row in bottom_df.iterrows()]

    return jsonify({"top": top_list, "bottom": bottom_list})


@app.route("/student_report")
def student_report():
    """Generate a simple PDF report for a student using cached data (ReportLab)."""
    sheet = request.args.get("sheet")
    student_name_raw = request.args.get("student")
    if not sheet or not student_name_raw:
        return "Missing parameters", 400

    try:
        df = get_sheet_df(sheet)
    except APIError as e:
        return f"Google API error: {e}", 500

    name_col = next((c for c in df.columns if "name" in str(c).lower()), None)
    if not name_col:
        return "Name column not found", 400

    df_copy = df.copy()
    df_copy["__match__"] = df_copy[name_col].astype(str).str.strip().str.lower()
    mask = df_copy["__match__"] == student_name_raw.strip().lower()
    if not mask.any():
        return "Student not found", 404

    row = df_copy[mask].iloc[0].drop(labels="__match__")
    student_data = row.to_dict()

    cgpa_value = compute_weighted_cgpa_from_row(student_data)
    sem_gpa_labels, sem_gpa_values = compute_sem_gpa_from_row(student_data)

    # build PDF
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 50

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, f"Student Report: {student_name_raw}")
    y -= 30

    c.setFont("Helvetica", 12)
    if cgpa_value is not None:
        c.drawString(50, y, f"CGPA: {cgpa_value}")
        y -= 20

    c.drawString(50, y, f"Sheet: {sheet}")
    y -= 30

    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, y, "Personal Details:")
    y -= 20
    c.setFont("Helvetica", 11)
    for k, v in student_data.items():
        if any(t in str(k).lower() for t in ["name", "gender", "dob", "reg", "roll", "native", "district"]):
            c.drawString(60, y, f"{k}: {v}")
            y -= 15
            if y < 80:
                c.showPage()
                y = height - 50

    # Semester GPA
    y -= 10
    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, y, "Semester GPA:")
    y -= 20
    c.setFont("Helvetica", 11)
    for sem, gpa in zip(sem_gpa_labels, sem_gpa_values):
        c.drawString(60, y, f"{sem}: {gpa}")
        y -= 15
        if y < 80:
            c.showPage()
            y = height - 50

    c.showPage()
    c.save()
    buffer.seek(0)
    filename = f"{student_name_raw.replace(' ', '_')}_report.pdf"
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="application/pdf")

@app.route("/subject_codes", methods=["POST"])
def subject_codes():
    sheet = request.form.get("sheet")
    df = get_sheet_df(sheet)

    subject_codes = set()

    # Example column: Sem1_BS3171_2
    for col in df.columns:
        parts = col.split("_")
        if len(parts) >= 3 and parts[0].lower().startswith("sem"):
            subject_codes.add(parts[1].upper())

    return jsonify(sorted(subject_codes))

@app.route("/subject_analytics", methods=["POST"])
def subject_analytics():
    subject = request.form.get("subject")
    scope = request.form.get("scope")  # batch / dept
    sheet = request.form.get("sheet")
    department = request.form.get("department")

    def is_fail(v):
        return str(v).strip().upper() in ["RA", "U", "F", "FAIL", "ABSENT"]

    def grade_bucket(v):
        v = str(v).strip().upper()
        return v if v not in ["", "NAN"] else "NA"

    #sheets = [sheet] if scope == "batch" else list(SHEET_CACHE.keys())
    if scope == "batch":
        sheets = [sheet]
    else:
        # ONLY sheets belonging to selected department
        sheets = [
            s for s in SHEET_CACHE.keys()
            if department and department in s
        ]
    if not sheets:
        return jsonify({
            "scope": scope,
            "sheets": [],
            "data": {},
            "pass_percent": {}
        })

    failures = []
    grade_count = {}

    for sh in sheets:
        df = get_sheet_df(sh)
        name_col = next(c for c in df.columns if "name" in c.lower())

        for col in df.columns:
            if f"_{subject}_" not in col:
                continue

            for _, row in df.iterrows():
                name = row[name_col]
                val = row[col]

                if is_fail(val):
                    failures.append({
                        "name": name,
                        "batch": sh
                    })

                grade = grade_bucket(val)
                grade_count[grade] = grade_count.get(grade, 0) + 1

    return jsonify({
        "failures": failures,
        "grades": grade_count
    })

@app.route("/download_student_pdf")
def download_student_pdf():
    sheet = request.args.get("sheet")
    student = request.args.get("student")

    # URL of already rendered page
    url = f"http://127.0.0.1:5000/?sheet={sheet}&student={student}&pdf=1"

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        page.goto(url, wait_until="networkidle")

        pdf_path = os.path.join(tempfile.gettempdir(), f"{student}_report.pdf")

        page.pdf(
            path=pdf_path,
            format="A4",
            print_background=True
        )

        browser.close()

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=f"{student}_performance.pdf"
    )
#app routes for new features are added from here
@app.route("/students_by_arrears", methods=["POST"])
def students_by_arrears():
    scope = request.form.get("scope")       # batch / dept
    sheet = request.form.get("sheet")
    dept = request.form.get("department")
    bucket = request.form.get("bucket")

    def in_bucket(a):
        if bucket == "6-10":
            return 6 <= a <= 10
        if bucket == ">10":
            return a > 10
        return a == int(bucket)

    result = []

    sheets = (
        [sheet] if scope == "batch"
        else [s for s in SHEET_CACHE if dept in s]
    )

    for sh in sheets:
        if dept not in sh:
            continue

        df = get_sheet_df(sh)
        if df.empty:
            continue

        # robust name detection
        name_col = next(
            (c for c in df.columns
             if "name" in c.lower() and "father" not in c.lower()),
            None
        )
        if not name_col:
            continue

        for _, row in df.iterrows():
            arrears = sum(
                str(v).strip().upper() in ["RA", "U", "F", "FAIL"]
                for v in row.values
            )

            if in_bucket(arrears):

                result.append({
                    "name": str(row[name_col]).strip(),
                    "arrears": arrears,
                    "batch": sh
                })

    return jsonify(result)

from flask import send_file
from playwright.sync_api import sync_playwright
import tempfile
import os

@app.route("/print_page")
def print_page():
    # URL of the page to print (same app)
    page_url = "http://127.0.0.1:5000/"

    # Create temporary PDF file
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    temp_pdf.close()

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # Load page fully (important for charts)
        page.goto(page_url, wait_until="networkidle")

        page.pdf(
            path=temp_pdf.name,
            format="A4",
            print_background=True,
            margin={
                "top": "20mm",
                "bottom": "20mm",
                "left": "15mm",
                "right": "15mm"
            }
        )

        browser.close()

    return send_file(
        temp_pdf.name,
        as_attachment=True,
        download_name="Student_Performance_Report.pdf",
        mimetype="application/pdf"
    )

@app.route("/autosave_pdf", methods=["POST"])
def autosave_pdf():
    data = request.json

    html_content = data.get("html", "")
    batch = data.get("batch", "BATCH")
    student = data.get("student", "ALL")

    safe_batch = batch.replace(" ", "_")
    safe_student = student.replace(" ", "_")
    filename = f"{safe_batch}_{safe_student}_Performance_Report.pdf"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf_path = tmp.name

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        # Load HTML exactly as sent from browser
        page.set_content(html_content, wait_until="networkidle")

        page.pdf(
            path=pdf_path,
            format="A4",
            print_background=True,
            margin={
                "top": "20mm",
                "bottom": "20mm",
                "left": "15mm",
                "right": "15mm"
            }
        )

        browser.close()

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf"
    )

@app.route("/subject_pass_percentage", methods=["POST"])
def subject_pass_percentage():
    sheet = request.form.get("sheet")
    df = get_sheet_df(sheet)

    sem_cols = [c for c in df.columns if re.search(r"Sem\d+_.*_\d+", c, re.I)]
    results = []

    fail_grades = {"U", "RA", "AU", "ABSENT", "F", "FAIL"}

    for col in sem_cols:
        subject = col.split("_")[1].upper()
        total = df[col].notna().sum()
        passed = df[col].apply(
            lambda x: str(x).strip().upper() not in fail_grades
        ).sum()

        if total > 0:
            pass_percent = round((passed / total) * 100, 2)
            results.append({
                "subject": subject,
                "pass_percent": pass_percent
            })

    return jsonify(results)

@app.route("/subject_pass_percentage_dept", methods=["POST"])
def subject_pass_percentage_dept():
    dept = request.form["department"]
    result = {}

    for batch, df in load_all_batches():
        d = df[df["Department"] == dept]

        for col in subject_columns(d):
            passed = d[col].isin(["S","A","B","C","D","E"]).sum()
            total = d[col].notna().sum()
            result[col] = round(passed * 100 / total, 2) if total else 0

    return jsonify([
        {"subject": k, "pass_percent": v}
        for k, v in result.items()
    ])

"""
@app.route("/dept_subject_pass_percentage", methods=["POST"])
def dept_subject_pass_percentage():
    department = request.form.get("department")

    sheets = [s for s in list_sheets() if department in s]
    output = {}

    fail_grades = {"U", "RA", "AU", "ABSENT", "F", "FAIL"}

    for sh in sheets:
        df = get_sheet_df(sh)
        sem_cols = [c for c in df.columns if re.search(r"Sem\d+_.*_\d+", c, re.I)]

        batch_result = {}
        for col in sem_cols:
            subject = col.split("_")[1].upper()
            total = df[col].notna().sum()
            passed = df[col].apply(
                lambda x: str(x).strip().upper() not in fail_grades
            ).sum()

            if total > 0:
                batch_result[subject] = round((passed / total) * 100, 2)

        output[sh] = batch_result

    return jsonify(output)
"""

@app.route("/cutoff_vs_arrears_batch", methods=["POST"])
def cutoff_vs_arrears_batch():
    sheet = request.form["sheet"]
    df = load_sheet(sheet)

    buckets = {"<150":0, "150-175":0, "175-200":0, ">200":0}

    for _, r in df.iterrows():
        arrears = r["Arrears"]
        cutoff = r["Cutoff"]

        if arrears > 0:
            if cutoff < 150: buckets["<150"] += 1
            elif cutoff <= 175: buckets["150-175"] += 1
            elif cutoff <= 200: buckets["175-200"] += 1
            else: buckets[">200"] += 1

    return jsonify([
        {"cutoff": k, "count": v} for k,v in buckets.items()
    ])

"""
@app.route("/cutoff_arrear_batch", methods=["POST"])
def cutoff_arrear_batch():
    sheet = request.form.get("sheet")
    df = get_sheet_df(sheet)

    cutoff_col = next(c for c in df.columns if "cut" in c.lower())

    buckets = {
        "80-100": (80,100),
        "100-120": (100,120),
        "120-140": (120,140),
        "140-160": (140,160),
        "160-180": (160,180),
        ">180": (181,1000)
    }

    result = {}

    for label,(lo,hi) in buckets.items():
        subset = df[pd.to_numeric(df[cutoff_col], errors="coerce").between(lo,hi)]
        counts = {"0":0,"1":0,"2":0,"3":0,"4":0,"5":0,"6-10":0,">10":0}

        for _,row in subset.iterrows():
            arrears = sum(
                str(v).strip().upper() in ["RA","U","F","FAIL","ABSENT"]
                for v in row.values
            )
            if arrears <= 5:
                counts[str(arrears)] += 1
            elif arrears <= 10:
                counts["6-10"] += 1
            else:
                counts[">10"] += 1

        result[label] = counts

    return jsonify(result)
"""
@app.route("/cutoff_vs_arrears_dept", methods=["POST"])
def cutoff_vs_arrears_dept():
    dept = request.form["department"]
    buckets = {"<150":0, "150-175":0, "175-200":0, ">200":0}

    for _, df in load_all_batches():
        d = df[df["Department"] == dept]

        for _, r in d.iterrows():
            if r["Arrears"] > 0:
                c = r["Cutoff"]
                if c < 150: buckets["<150"] += 1
                elif c <= 175: buckets["150-175"] += 1
                elif c <= 200: buckets["175-200"] += 1
                else: buckets[">200"] += 1

    return jsonify([
        {"cutoff": k, "count": v} for k,v in buckets.items()
    ])

"""
@app.route("/cutoff_arrear_dept", methods=["POST"])
def cutoff_arrear_dept():
    department = request.form.get("department")
    sheets = [s for s in list_sheets() if department in s]

    output = {}
    for sh in sheets:
        output[sh] = json.loads(
            cutoff_arrear_batch().get_data(as_text=True)
        )

    return jsonify(output)
"""
@app.route("/dept_dashboard", methods=["POST"])
def dept_dashboard():
    dept = request.form.get("department")

    dashboard = []

    for batch, df in load_all_batches():  # your existing loader
        d = df[df["Department"] == dept]

        dashboard.append({
            "batch": batch,
            "strength": len(d),
            "boys": int((d["Gender"] == "M").sum()),
            "girls": int((d["Gender"] == "F").sum()),
            "hostellers": int((d["Hostel"] == "Yes").sum()),
            "day_scholars": int((d["Hostel"] == "No").sum()),
            "fg": int((d["Quota"] == "FG").sum()),
            "gq": int((d["Quota"] == "GQ").sum()),
            "mq": int((d["Quota"] == "MQ").sum())
        })

    return jsonify(dashboard)

"""
@app.route("/dept_dashboard", methods=["POST"])
def dept_dashboard():
    department = request.form.get("department")
    sheets = [s for s in list_sheets() if department in s]

    dashboard = []

    for sh in sheets:
        df = get_sheet_df(sh)

        def count_col(keyword):
            col = next((c for c in df.columns if keyword in c.lower()), None)
            return df[col].astype(str).str.upper().value_counts().to_dict() if col else {}

        dashboard.append({
            "batch": sh,
            "strength": len(df),
            "gender": count_col("gender"),
            "hostel": count_col("hostel"),
            "fg": count_col("first"),
            "quota": count_col("gq")
        })

    return jsonify(dashboard)
"""

if __name__ == "__main__":
    app.run(debug=True)











