import io
import os
import re
import secrets
import sqlite3
import sys
import traceback
import uuid
import smtplib
import ssl
from email.message import EmailMessage
from urllib import parse as urlparse, request as urlrequest
from datetime import date, timedelta

from flask import Flask, jsonify, redirect, render_template, request, send_file, url_for, session
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename


def _load_local_env(path):
    """Load simple KEY=VALUE settings without overwriting real environment variables."""
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip("\"").strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


_load_local_env(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))


app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY") or secrets.token_hex(32)

UPLOAD_FOLDER = os.path.join(app.root_path, "static", "uploads")
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ─── SQLite Database ────────────────────────────────────────────────────
DB_PATH = os.path.join(app.root_path, "student_records.db")


def get_db():
    """Return a new SQLite connection with row factory enabled."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create all required tables if they do not exist."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            register_number TEXT NOT NULL UNIQUE,
            student_phone TEXT NOT NULL,
            course TEXT NOT NULL,
            year TEXT NOT NULL,
            parent_name TEXT DEFAULT '',
            parent_phone TEXT DEFAULT '',
            photo_filename TEXT DEFAULT '',
            semester TEXT DEFAULT '',
            section TEXT DEFAULT ''
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            course TEXT NOT NULL,
            year TEXT NOT NULL,
            semester TEXT NOT NULL,
            subject TEXT NOT NULL,
            register_number TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Present',
            UNIQUE(date, course, year, semester, subject, register_number)
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS internal_marks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course TEXT NOT NULL,
            year TEXT NOT NULL,
            semester TEXT NOT NULL,
            subject TEXT NOT NULL,
            test_name TEXT NOT NULL,
            register_number TEXT NOT NULL,
            marks_obtained INTEGER NOT NULL,
            max_marks INTEGER NOT NULL,
            test_date TEXT DEFAULT '',
            remarks TEXT DEFAULT 'Good',
            UNIQUE(course, year, semester, subject, test_name, register_number)
        )
"""
    )
    cur.execute("""CREATE TABLE IF NOT EXISTS subjects (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, semester TEXT NOT NULL, subject_name TEXT NOT NULL, subject_code TEXT NOT NULL DEFAULT '', UNIQUE(course, semester, subject_code), UNIQUE(course, semester, subject_name))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS timetable_periods (id INTEGER PRIMARY KEY AUTOINCREMENT, label TEXT NOT NULL, start_time TEXT NOT NULL, end_time TEXT NOT NULL, is_saturday INTEGER NOT NULL DEFAULT 0, UNIQUE(label, start_time, end_time, is_saturday))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS timetable_entries (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, year TEXT NOT NULL, semester TEXT NOT NULL, section TEXT NOT NULL, day TEXT NOT NULL, period_id INTEGER NOT NULL, subject_id INTEGER, faculty_name TEXT NOT NULL DEFAULT '', lab_batch TEXT NOT NULL DEFAULT '', room TEXT NOT NULL DEFAULT '', activity_type TEXT NOT NULL DEFAULT '', UNIQUE(course, year, semester, section, day, period_id, lab_batch))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, event_date TEXT NOT NULL, start_time TEXT DEFAULT '', end_time TEXT DEFAULT '', course TEXT NOT NULL, year TEXT NOT NULL, semester TEXT NOT NULL, section TEXT NOT NULL, created_at TEXT NOT NULL, created_by TEXT DEFAULT '')""")
    cur.execute("""CREATE TABLE IF NOT EXISTS event_students (event_id INTEGER NOT NULL, register_number TEXT NOT NULL, PRIMARY KEY(event_id, register_number))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS event_claims (id INTEGER PRIMARY KEY AUTOINCREMENT, event_id INTEGER NOT NULL, subject TEXT NOT NULL, advisor_name TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'PENDING', message_text TEXT DEFAULT '', created_on TEXT DEFAULT '', message_sent_on TEXT DEFAULT '', claimed_on TEXT DEFAULT '', reminder_sent_on TEXT DEFAULT '', UNIQUE(event_id, subject, advisor_name))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS event_notifications (id INTEGER PRIMARY KEY AUTOINCREMENT, claim_id INTEGER NOT NULL, recipient_role TEXT NOT NULL, recipient_name TEXT DEFAULT '', notification_type TEXT NOT NULL, message TEXT NOT NULL, created_on TEXT NOT NULL, read_on TEXT DEFAULT '', UNIQUE(claim_id, recipient_role, notification_type))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS faculty (id INTEGER PRIMARY KEY AUTOINCREMENT, faculty_id TEXT NOT NULL UNIQUE, name TEXT NOT NULL, email TEXT NOT NULL UNIQUE, phone TEXT NOT NULL, department TEXT NOT NULL, designation TEXT NOT NULL, password_hash TEXT NOT NULL, created_at TEXT NOT NULL)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS parent_alerts (id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT NOT NULL, alert_type TEXT NOT NULL, condition_key TEXT NOT NULL UNIQUE, message TEXT NOT NULL, attendance_percentage REAL, failed_subjects TEXT DEFAULT '', created_date TEXT NOT NULL, email_status TEXT NOT NULL DEFAULT 'Not configured', sms_status TEXT NOT NULL DEFAULT 'Not configured', status TEXT NOT NULL DEFAULT 'New', last_error TEXT DEFAULT '')""")
    cur.executemany("INSERT OR IGNORE INTO timetable_periods (label,start_time,end_time,is_saturday) VALUES (?,?,?,?)", [("1st Period","09:30","10:20",0),("2nd Period","10:25","11:15",0),("3rd Period","11:20","12:10",0),("4th Period","12:10","13:05",0),("5th Period","14:00","14:50",0),("6th Period","14:55","15:45",0),("1st Period","09:30","10:20",1),("2nd Period","10:25","11:15",1),("3rd Period","11:20","12:10",1),("4th Period","12:10","13:05",1)])
    cur.executemany("INSERT OR IGNORE INTO subjects (course, semester, subject_name, subject_code) VALUES (?, ?, ?, ?)", [("BCA", "I Semester", "Python Programming", "BCA101"), ("BCA", "I Semester", "Database Management", "BCA102"), ("BCA", "I Semester", "Web Technology", "BCA103"), ("BCA", "I Semester", "Mathematics", "BCA104"), ("BCA", "III Semester", "Data Structures", "BCS302")])
    conn.commit()
    conn.close()


def _column_exists(cur, table, column):
    """Return True if a column exists on the given table."""
    cols = [r[1] for r in cur.execute(f"PRAGMA table_info({table})").fetchall()]
    return column in cols


def migrate_db():
    """Add missing columns to existing databases (safe migration)."""
    conn = get_db()
    cur = conn.cursor()
    if not _column_exists(cur, "students", "parent_email"):
        cur.execute("ALTER TABLE students ADD COLUMN parent_email TEXT DEFAULT ''")
    for column in ("semester", "section", "language_group", "lab_batch"):
        if not _column_exists(cur, "students", column):
            cur.execute(f"ALTER TABLE students ADD COLUMN {column} TEXT DEFAULT ''")
    if not _column_exists(cur, "subjects", "subject_type"):
        cur.execute("ALTER TABLE subjects ADD COLUMN subject_type TEXT DEFAULT 'Theory'")
    for table, column, definition in (("events", "created_by", "TEXT DEFAULT ''"), ("event_claims", "message_text", "TEXT DEFAULT ''"), ("event_claims", "created_on", "TEXT DEFAULT ''")):
        if not _column_exists(cur, table, column):
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
    if not _column_exists(cur, "timetable_entries", "faculty_id"):
        cur.execute("ALTER TABLE timetable_entries ADD COLUMN faculty_id TEXT DEFAULT ''")
    # Map legacy timetable rows that stored a faculty name but not a faculty ID.
    # Rows without either value remain unassigned and are never shown to a faculty member.
    cur.execute("""UPDATE timetable_entries
                   SET faculty_id = (
                       SELECT f.faculty_id FROM faculty f
                       WHERE lower(trim(f.name)) = lower(trim(timetable_entries.faculty_name))
                       LIMIT 1
                   )
                   WHERE trim(COALESCE(faculty_id, '')) = ''
                     AND trim(COALESCE(faculty_name, '')) <> ''
                     AND EXISTS (
                       SELECT 1 FROM faculty f
                       WHERE lower(trim(f.name)) = lower(trim(timetable_entries.faculty_name))
                   )""")
    conn.commit()
    conn.close()


init_db()
migrate_db()


# ─── Helpers ────────────────────────────────────────────────────────────

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def is_valid_phone(phone):
    return phone.isdigit() and len(phone) == 10


def _digits_only(value):
    return re.sub(r"\D", "", value or "")


def _format_sms_number(phone_number):
    """Return an E.164 phone number for Twilio, or an empty string if invalid."""
    digits = _digits_only(phone_number)
    country_code = os.getenv("SMS_COUNTRY_CODE", "+91").strip() or "+91"
    country_digits = _digits_only(country_code)

    if len(digits) == 10:
        return f"+{country_digits}{digits}"
    if country_digits and digits.startswith(country_digits) and len(digits) == len(country_digits) + 10:
        return f"+{digits}"
    return ""


def attendance_mark_from_percentage(percentage):
    """Convert an attendance percentage to the Semester Results mark out of 5."""
    if percentage <= 20:
        return 1
    if percentage <= 40:
        return 2
    if percentage <= 60:
        return 3
    if percentage < 90:
        return 4
    return 5

def _scaled_component(marks_value, max_marks, target_marks):
    """Convert a raw assessment score into the required out-of-target marks."""
    try:
        marks = float(marks_value)
        maximum = float(max_marks or 0)
    except (TypeError, ValueError):
        return 0.0
    if maximum <= 0:
        return 0.0
    return (marks / maximum) * float(target_marks)


def _grade_from_percentage(percentage):
    if percentage >= 90:
        return "A+"
    if percentage >= 80:
        return "A"
    if percentage >= 70:
        return "B+"
    if percentage >= 60:
        return "B"
    if percentage >= 50:
        return "C"
    if percentage >= 40:
        return "D"
    return "F"


def _grade_point_for_grade(grade):
    return {"A+": 10, "A": 9, "B+": 8, "B": 7, "C": 6, "D": 5, "F": 0}.get(grade, 0)


def get_current_student():
    register_number = session.get("student_id")
    if not register_number:
        return None
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM students WHERE register_number = ?", (register_number,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def get_current_parent_student():
    register_number = session.get("parent_student_id")
    if not register_number:
        return None
    conn = get_db()
    row = conn.execute("SELECT * FROM students WHERE register_number = ?", (register_number,)).fetchone()
    conn.close()
    return dict(row) if row else None


@app.context_processor
def inject_student_session():
    student = get_current_student()
    return {
        "student_logged_in": bool(student),
        "current_student": student,
    }


# ─── Public Pages ───────────────────────────────────────────────────────

@app.get("/")
def home():
    return render_template("index.html")


@app.get("/courses")
def courses():
    return render_template("courses.html")


@app.get("/about")
def about():
    return render_template("about.html")


@app.get("/contact")
def contact():
    return render_template("about.html")


# ─── Student Management ─────────────────────────────────────────────────

@app.route("/students/new", methods=["GET", "POST"])
def new_student():
    if request.method == "POST":
        student = {
            "name": request.form.get("name", "").strip(),
            "register_number": request.form.get("register_number", "").strip(),
            "student_phone": request.form.get("student_phone", "").strip(),
            "course": request.form.get("course", "").strip(),
            "year": request.form.get("year", "").strip(),
            "parent_name": request.form.get("parent_name", "").strip(),
            "parent_phone": request.form.get("parent_phone", "").strip(),
            "parent_email": request.form.get("parent_email", "").strip(),
        }
        required_fields = ("name", "register_number", "student_phone", "course", "year", "parent_name", "parent_phone")
        if not all(student[field] for field in required_fields):
            return render_template("new_student.html", error="Please complete every field.", student=student)
        if not is_valid_phone(student["student_phone"]):
            return render_template("new_student.html", error="Student phone must be exactly 10 digits.", student=student)
        if student["parent_email"] and not _valid_parent_email(student["parent_email"]):
            return render_template("new_student.html", error="Enter a valid parent email address.", student=student)
        if not is_valid_phone(student["parent_phone"]):
            return render_template("new_student.html", error="Parent phone must be exactly 10 digits.", student=student)
        conn = get_db()
        existing = conn.execute(
            "SELECT id FROM students WHERE register_number = ?", (student["register_number"],)
        ).fetchone()
        if existing:
            conn.close()
            return render_template("new_student.html", error="That register number already exists.", student=student, recent_students=_recent_students())
        conn.execute(
            """INSERT INTO students
               (name, register_number, student_phone, course, year, parent_name, parent_phone, parent_email, photo_filename)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                student["name"],
                student["register_number"],
                student["student_phone"],
                student["course"],
                student["year"],
                student["parent_name"],
                student["parent_phone"],
                student["parent_email"],
                "",
            ),
        )
        conn.commit()
        conn.close()
        session["student_id"] = student["register_number"]
        return redirect(url_for("student_dashboard"))

    student = {
        "course": request.args.get("course", "").strip(),
        "year": request.args.get("year", "").strip(),
    }
    imported_register_numbers = [value.strip() for value in request.args.get("imported", "").split(",") if value.strip()]
    return render_template("new_student.html", student=student, recent_students=_recent_students(imported_register_numbers), message=request.args.get("message", ""), message_type=request.args.get("message_type", "success"))


def _recent_students(register_numbers=None):
    conn = get_db()
    register_numbers = [value for value in (register_numbers or []) if value]
    if register_numbers:
        placeholders = ",".join("?" for _ in register_numbers)
        rows = conn.execute(
            f"SELECT * FROM students ORDER BY CASE WHEN register_number IN ({placeholders}) THEN 0 ELSE 1 END, id DESC LIMIT 10",
            register_numbers,
        ).fetchall()
    else:
        rows = conn.execute("SELECT * FROM students ORDER BY id DESC LIMIT 10").fetchall()
    conn.close()
    return [dict(row) for row in rows]


@app.get("/students")
def students():
    course = request.args.get("course", "").strip()
    sort_order = request.args.get("sort", "asc").strip().lower()
    message = request.args.get("message", "")
    if sort_order not in {"asc", "desc"}:
        sort_order = "asc"
    direction = "ASC" if sort_order == "asc" else "DESC"

    conn = get_db()
    if course:
        rows = conn.execute(
            f"SELECT * FROM students WHERE course = ? ORDER BY name {direction}", (course,)
        ).fetchall()
    else:
        rows = conn.execute(f"SELECT * FROM students ORDER BY name {direction}").fetchall()
    conn.close()
    student_list = [dict(row) for row in rows]
    return render_template("students.html", students=student_list, course=course, sort_order=sort_order, message=message)


@app.route("/students/delete-all", methods=["POST"])
def delete_all_students():
    conn = get_db()
    conn.execute("DELETE FROM students")
    conn.commit()
    conn.close()
    return redirect(url_for("students", message="All student records have been deleted."))


@app.get("/students/export.xlsx")
def export_students():
    conn = get_db()
    rows = conn.execute("SELECT * FROM students ORDER BY name ASC").fetchall()
    conn.close()
    student_list = [dict(row) for row in rows]

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Students"
    headers = ["Student Name", "Register Number", "Student Phone", "Course", "Year", "Semester", "Section", "Parent Name", "Parent Phone", "Parent Email"]
    worksheet.append(headers)
    for student in student_list:
        worksheet.append(
            [
                student.get("name", ""),
                student.get("register_number", ""),
                student.get("student_phone", ""),
                student.get("course", ""),
                student.get("year", ""),
                student.get("semester", ""),
                student.get("section", ""),
                student.get("parent_name", ""),
                student.get("parent_phone", ""),
                student.get("parent_email", ""),
            ]
        )
    for column in worksheet.columns:
        width = max(len(str(cell.value or "")) for cell in column) + 2
        worksheet.column_dimensions[column[0].column_letter].width = min(width, 30)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="student_records.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─── Excel Upload (Faculty Add Student module) ──────────────────────────



def _normalise_header(value):
    """Turn a header cell into a lowercase key for matching."""
    return re.sub(r"[^a-z0-9]+", "_", str(value or "").strip().lower()).strip("_")


def _as_text(value):
    """Convert Excel cell values to a safe text string without changing real content."""
    return "" if value is None else str(value).strip()


def _sorted_text_values(values):
    """Sort values safely even when Excel cells were parsed as ints or strings."""
    seen = set()
    ordered = []
    for value in values:
        text = _as_text(value)
        if not text or text in seen:
            continue
        seen.add(text)
        ordered.append(text)
    return sorted(ordered, key=lambda item: item.lower())


def _normalise_semester(value):
    value = _as_text(value)
    aliases = {}
    for number, roman in enumerate(("I", "II", "III", "IV", "V", "VI"), start=1):
        aliases[str(number)] = f"{roman} Semester"
        aliases[f"semester {number}"] = f"{roman} Semester"
        aliases[f"{roman.lower()} semester"] = f"{roman} Semester"
    return aliases.get(value.lower(), value)


def _normalise_course(value):
    value = _as_text(value)
    return {
        "bca": "BCA", "bba": "BBA", "ba": "BA",
        "b.com": "B.Com", "bcom": "B.Com",
        "b.sc": "B.Sc", "bsc": "B.Sc",
    }.get(value.lower(), value)


def _normalise_year(value):
    value = _as_text(value)
    return {
        "1": "1st Year", "1st": "1st Year", "first": "1st Year", "first year": "1st Year", "1st year": "1st Year",
        "2": "2nd Year", "2nd": "2nd Year", "second": "2nd Year", "second year": "2nd Year", "2nd year": "2nd Year",
        "3": "3rd Year", "3rd": "3rd Year", "third": "3rd Year", "third year": "3rd Year", "3rd year": "3rd Year",
    }.get(value.lower(), value)


def _normalise_excel_phone(value):
    """Convert common Excel and Indian phone formats to the stored 10 digits."""
    text = _as_text(value)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    if not re.fullmatch(r"[+()\d\s.-]+", text):
        return text
    digits = _digits_only(text)
    if len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    elif len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    return digits




@app.post("/students/import-excel")
def import_students_excel():
    """Excel import lives within Add Student; it is not an Admin Dashboard feature."""
    file = request.files.get("file")
    if not file or not file.filename.lower().endswith(".xlsx"):
        return redirect(url_for("new_student", message="Choose a valid .xlsx student file.", message_type="error"))
    try:
        workbook = load_workbook(file, read_only=True, data_only=True)
        rows = list(workbook.active.iter_rows(values_only=True)); workbook.close()
        if not rows:
            raise ValueError("The Excel file is empty.")
        aliases = {
            "reg_no": "register_number",
            "register_no": "register_number",
            "registration_no": "register_number",
            "registration_number": "register_number",
            "student_id": "register_number",
            "name": "student_name",
            "phone": "student_phone",
            "phone_no": "student_phone",
            "phone_number": "student_phone",
            "student_phone_no": "student_phone",
            "student_phone_number": "student_phone",
            "parent_phone_no": "parent_phone",
            "parent_phone_number": "parent_phone",
            "parent_mobile": "parent_phone",
            "parent_mobile_number": "parent_phone",
            "parents_email": "parent_email",
            "parent_email_id": "parent_email",
            "parents_email_id": "parent_email",
        }
        index_map = {aliases.get(_normalise_header(value), _normalise_header(value)): index for index, value in enumerate(rows[0])}
        required = {"register_number", "student_name", "course", "year", "semester", "section", "parent_name", "parent_phone"}
        if required - set(index_map):
            raise ValueError("Missing required columns: " + ", ".join(sorted(required - set(index_map))))
        conn = get_db(); added = updated = failed = 0; errors = []; seen = set(); imported_register_numbers = []
        try:
            for row_number, row in enumerate(rows[1:], start=2):
                if not row or all(not _as_text(cell) for cell in row):
                    continue
                get = lambda key: _as_text(row[index_map[key]]) if key in index_map and index_map[key] < len(row) else ""
                record = {key: get(key) for key in ("student_name", "register_number", "student_phone", "course", "year", "semester", "section", "parent_name", "parent_phone", "parent_email")}
                record["semester"] = _normalise_semester(record["semester"])
                record["course"] = _normalise_course(record["course"])
                record["year"] = _normalise_year(record["year"])
                record["section"] = record["section"].upper()
                record["student_phone"] = _normalise_excel_phone(record["student_phone"])
                record["parent_phone"] = _normalise_excel_phone(record["parent_phone"])
                if not all(record[key] for key in ("student_name", "register_number", "course", "year", "semester", "section", "parent_name", "parent_phone")):
                    failed += 1; errors.append(f"Row {row_number}: required student details are missing."); continue
                if record["course"] not in {"BCA", "BBA", "BA", "B.Com", "B.Sc"} or record["year"] not in {"1st Year", "2nd Year", "3rd Year"} or record["semester"] not in {"I Semester", "II Semester", "III Semester", "IV Semester", "V Semester", "VI Semester"}:
                    failed += 1; errors.append(f"Row {row_number}: invalid course, year, or semester."); continue
                if not record["student_phone"]:
                    record["student_phone"] = record["parent_phone"]
                if not is_valid_phone(record["student_phone"]):
                    failed += 1; errors.append(f"Row {row_number}: student phone must contain 10 digits."); continue
                if not is_valid_phone(record["parent_phone"]):
                    failed += 1; errors.append(f"Row {row_number}: parent phone must contain 10 digits (an optional +91 prefix is accepted)."); continue
                if record["parent_email"] and not _valid_parent_email(record["parent_email"]):
                    failed += 1; errors.append(f"Row {row_number}: parent email is invalid."); continue
                rn = record["register_number"]
                if rn in seen:
                    failed += 1; errors.append(f"Row {row_number}: registration number {rn} is duplicated in the file."); continue
                seen.add(rn)
                try:
                    existing = conn.execute("SELECT 1 FROM students WHERE register_number=?", (rn,)).fetchone()
                    if existing:
                        result = conn.execute("""UPDATE students SET name=?,student_phone=?,course=?,year=?,semester=?,section=?,parent_name=?,parent_phone=?,parent_email=?
                            WHERE register_number=?""", (record["student_name"],record["student_phone"],record["course"],record["year"],record["semester"],record["section"],record["parent_name"],record["parent_phone"],record["parent_email"],rn))
                        if result.rowcount != 1:
                            raise RuntimeError("the existing database row could not be updated")
                        updated += 1
                    else:
                        result = conn.execute("""INSERT INTO students (name,register_number,student_phone,course,year,semester,section,parent_name,parent_phone,parent_email,photo_filename)
                            VALUES (?,?,?,?,?,?,?,?,?,?,'')""", (record["student_name"],rn,record["student_phone"],record["course"],record["year"],record["semester"],record["section"],record["parent_name"],record["parent_phone"],record["parent_email"]))
                        if result.rowcount != 1:
                            raise RuntimeError("the database did not insert the row")
                        added += 1
                    imported_register_numbers.append(rn)
                except Exception as row_error:
                    failed += 1
                    errors.append(f"Row {row_number} ({rn}): {row_error}")
                    app.logger.exception("Student Excel import row %s failed in %s", row_number, DB_PATH)
            conn.commit()
        finally:
            conn.close()
        if added or updated:
            verify_conn = get_db()
            try:
                placeholders = ",".join("?" for _ in imported_register_numbers)
                saved = {item["register_number"] for item in verify_conn.execute(
                    f"SELECT register_number FROM students WHERE register_number IN ({placeholders})",
                    imported_register_numbers,
                ).fetchall()}
            finally:
                verify_conn.close()
            missing = [rn for rn in imported_register_numbers if rn not in saved]
            if missing:
                raise RuntimeError("Database verification failed for: " + ", ".join(missing))
            parts = []
            if added:
                parts.append(f"added {added} new student(s)")
            if updated:
                parts.append(f"updated {updated} existing student(s)")
            note = "Successfully " + " and ".join(parts) + f"; {failed} row(s) failed."
            if errors:
                note += " " + " ".join(errors[:5])
            app.logger.info("Student Excel import committed to %s: added=%s updated=%s failed=%s", DB_PATH, added, updated, failed)
            return redirect(url_for("students", message=note))
        raise ValueError(" ".join(errors[:6]) or "No valid students were found.")
    except Exception as exc:
        app.logger.exception("Student Excel import failed using database %s", DB_PATH)
        return redirect(url_for("new_student", message=f"Student Excel import failed: {exc}", message_type="error"))


# ─── Student Records (searchable) ───────────────────────────────────────

@app.route("/faculty/students-records")
def student_records():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))

    message = request.args.get("message", "")
    message_type = request.args.get("message_type", "success")

    conn = get_db()
    rows = conn.execute("SELECT * FROM students ORDER BY name ASC").fetchall()
    total_students = conn.execute("SELECT COUNT(*) AS c FROM students").fetchone()["c"]
    courses = _sorted_text_values(r["course"] for r in rows if r["course"])
    years = _sorted_text_values(r["year"] for r in rows if r["year"])
    semesters = _sorted_text_values(r["semester"] for r in rows if r["semester"])
    sections = _sorted_text_values(r["section"] for r in rows if r["section"])
    conn.close()
    student_list = [dict(row) for row in rows]

    return render_template(
        "student_records.html",
        students=student_list,
        total_students=total_students,
        courses=courses,
        years=years,
        semesters=semesters,
        sections=sections,
        message=message,
        message_type=message_type,
    )


@app.route("/faculty/students/delete", methods=["POST"])
def student_delete():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))

    register_numbers = request.form.getlist("register_number")
    register_numbers = [rn.strip() for rn in register_numbers if rn.strip()]
    if not register_numbers:
        return redirect(url_for("student_records", message="Please select at least one student to delete.", message_type="error"))

    conn = get_db()
    placeholders = ",".join("?" for _ in register_numbers)
    cursor = conn.execute(
        f"DELETE FROM students WHERE register_number IN ({placeholders})", register_numbers
    )
    conn.commit()
    deleted = cursor.rowcount
    conn.close()

    if deleted:
        return redirect(
            url_for(
                "student_records",
                message=f"{deleted} student(s) deleted successfully.",
                message_type="success",
            )
        )
    return redirect(url_for("student_records", message="No matching students found.", message_type="error"))


@app.route("/faculty/students/<register_number>")
def student_profile(register_number):
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))

    conn = get_db()
    row = conn.execute(
        "SELECT * FROM students WHERE register_number = ?", (register_number,)
    ).fetchone()
    conn.close()
    if not row:
        return redirect(url_for("student_records", message="Student not found.", message_type="error"))
    return render_template("student_profile.html", student=dict(row))


@app.route("/faculty/students/<register_number>/edit", methods=["GET", "POST"])
def student_edit(register_number):
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))

    conn = get_db()
    row = conn.execute(
        "SELECT * FROM students WHERE register_number = ?", (register_number,)
    ).fetchone()
    if not row:
        conn.close()
        return redirect(url_for("student_records", message="Student not found.", message_type="error"))
    student = dict(row)
    conn.close()

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        student_phone = request.form.get("student_phone", "").strip()
        course = request.form.get("course", "").strip()
        year = request.form.get("year", "").strip()
        semester = request.form.get("semester", "").strip()
        section = request.form.get("section", "").strip()
        parent_name = request.form.get("parent_name", "").strip()
        parent_phone = request.form.get("parent_phone", "").strip()
        parent_email = request.form.get("parent_email", "").strip()

        if not all([name, student_phone, course, year, parent_name, parent_phone]):
            return render_template(
                "student_edit.html",
                student={**student, "name": name, "student_phone": student_phone, "course": course,
                         "year": year, "semester": semester, "section": section,
                         "parent_name": parent_name, "parent_phone": parent_phone},
                error="Please complete every required field.",
            )
        if not is_valid_phone(student_phone):
            return render_template("student_edit.html", student={**student, "student_phone": student_phone, "parent_name": parent_name, "parent_phone": parent_phone, "name": name, "course": course, "year": year, "semester": semester, "section": section}, error="Student phone must be exactly 10 digits.")
        if not is_valid_phone(parent_phone):
            return render_template("student_edit.html", student={**student, "student_phone": student_phone, "parent_name": parent_name, "parent_phone": parent_phone, "name": name, "course": course, "year": year, "semester": semester, "section": section}, error="Parent phone must be exactly 10 digits.")

        conn = get_db()
        conn.execute(
            """UPDATE students SET
               name = ?, student_phone = ?, course = ?, year = ?,
               semester = ?, section = ?, parent_name = ?, parent_phone = ?, parent_email = ?
               WHERE register_number = ?""",
            (name, student_phone, course, year, semester, section, parent_name, parent_phone, parent_email, register_number),
        )
        conn.commit()
        conn.close()
        return redirect(url_for("student_profile", register_number=register_number, message="Student details updated.", message_type="success"))

    return render_template("student_edit.html", student=student)


# ─── Faculty Management ─────────────────────────────────────────────────

def _valid_email(email):
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email))


def _faculty_rows(conn):
    return [dict(row) for row in conn.execute("SELECT faculty_id,name,email,phone,department,designation FROM faculty ORDER BY name").fetchall()]


@app.route("/admin/faculty", methods=["GET", "POST"])
def manage_faculty():
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    message = request.args.get("message", ""); message_type = request.args.get("message_type", "success")
    conn = get_db()
    if request.method == "POST" and request.form.get("action") == "add":
        fields = {key: request.form.get(key, "").strip() for key in ("faculty_id", "name", "email", "phone", "department", "designation", "password")}
        if not all(fields.values()) or not _valid_email(fields["email"]) or not is_valid_phone(fields["phone"]) or len(fields["password"]) < 8:
            message, message_type = "Enter all details, a valid email/10-digit phone, and a password of at least 8 characters.", "error"
        else:
            try:
                conn.execute("""INSERT INTO faculty (faculty_id,name,email,phone,department,designation,password_hash,created_at)
                    VALUES (?,?,?,?,?,?,?,?)""", (fields["faculty_id"], fields["name"], fields["email"], fields["phone"], fields["department"], fields["designation"], generate_password_hash(fields["password"]), date.today().isoformat()))
                conn.commit(); message = "Faculty member added successfully."
            except sqlite3.IntegrityError:
                message, message_type = "Faculty ID or email already exists.", "error"
    faculty = _faculty_rows(conn); conn.close()
    temporary_credentials = session.pop("faculty_upload_credentials", [])
    return render_template("manage_faculty.html", faculty=faculty, message=message, message_type=message_type, temporary_credentials=temporary_credentials)


@app.post("/admin/faculty/upload")
def upload_faculty():
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    file = request.files.get("file")
    if not file or not file.filename.lower().endswith(".xlsx"):
        return redirect(url_for("manage_faculty", message="Faculty upload failed. Please correct the errors and upload again.", message_type="error"))
    try:
        workbook = load_workbook(file, read_only=True, data_only=True)
        rows = list(workbook.active.iter_rows(values_only=True)); workbook.close()
        if not rows:
            raise ValueError("The Excel file is empty.")
        headers = [_normalise_header(value) for value in rows[0]]
        index_map = {header: index for index, header in enumerate(headers)}
        required = {"faculty_id", "faculty_name", "email", "phone", "department", "designation"}
        if required - set(index_map):
            raise ValueError("Missing required columns: " + ", ".join(sorted(required - set(index_map))))
        conn = get_db(); imported = skipped = 0; errors = []; credentials = []; seen = set()
        try:
            for row_number, row in enumerate(rows[1:], start=2):
                if not row or all(not _as_text(cell) for cell in row):
                    continue
                get = lambda key: _as_text(row[index_map[key]]) if key in index_map and index_map[key] < len(row) else ""
                faculty_id, name, email, phone, department, designation = (get(key) for key in ("faculty_id", "faculty_name", "email", "phone", "department", "designation"))
                password = get("password")
                if not all((faculty_id, name, email, phone, department, designation)) or not _valid_email(email) or not is_valid_phone(phone):
                    skipped += 1; errors.append(f"Row {row_number}: required details, email, or phone are invalid."); continue
                if faculty_id in seen or conn.execute("SELECT 1 FROM faculty WHERE faculty_id=? OR email=?", (faculty_id, email)).fetchone():
                    skipped += 1; errors.append(f"Row {row_number}: faculty already exists."); continue
                seen.add(faculty_id)
                if not password:
                    password = secrets.token_urlsafe(9)
                    credentials.append(f"{faculty_id}: {password}")
                elif len(password) < 8:
                    skipped += 1; errors.append(f"Row {row_number}: password must be at least 8 characters."); continue
                conn.execute("INSERT INTO faculty (faculty_id,name,email,phone,department,designation,password_hash,created_at) VALUES (?,?,?,?,?,?,?,?)", (faculty_id,name,email,phone,department,designation,generate_password_hash(password),date.today().isoformat()))
                imported += 1
            conn.commit()
        finally:
            conn.close()
        if imported:
            message = "✓ Faculty details uploaded successfully!"
            if credentials:
                session["faculty_upload_credentials"] = credentials
            if skipped: message += f" {skipped} row(s) skipped."
            return redirect(url_for("manage_faculty", message=message, message_type="success"))
        raise ValueError(" ".join(errors[:6]) or "No valid faculty records found.")
    except Exception as exc:
        return redirect(url_for("manage_faculty", message=f"❌ Faculty upload failed. Please correct the errors and upload again. {exc}", message_type="error"))


@app.route("/admin/faculty/<faculty_id>/edit", methods=["GET", "POST"])
def edit_faculty(faculty_id):
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    conn = get_db(); row = conn.execute("SELECT * FROM faculty WHERE faculty_id=?", (faculty_id,)).fetchone()
    if not row:
        conn.close(); return redirect(url_for("manage_faculty", message="Faculty member not found.", message_type="error"))
    error = None
    if request.method == "POST":
        fields = {key: request.form.get(key, "").strip() for key in ("name", "email", "phone", "department", "designation", "password")}
        if not all(fields[key] for key in ("name", "email", "phone", "department", "designation")) or not _valid_email(fields["email"]) or not is_valid_phone(fields["phone"]) or (fields["password"] and len(fields["password"]) < 8):
            error = "Enter valid details; if changing it, the password must be at least 8 characters."
        else:
            values = [fields["name"], fields["email"], fields["phone"], fields["department"], fields["designation"]]
            query = "UPDATE faculty SET name=?,email=?,phone=?,department=?,designation=?"
            if fields["password"]:
                query += ",password_hash=?"; values.append(generate_password_hash(fields["password"]))
            try:
                conn.execute(query + " WHERE faculty_id=?", (*values, faculty_id)); conn.commit(); conn.close()
                return redirect(url_for("manage_faculty", message="Faculty member updated."))
            except sqlite3.IntegrityError:
                error = "That email is already used by another faculty member."
        row = {**dict(row), **fields}
    conn.close()
    return render_template("faculty_form.html", faculty=dict(row), error=error)


@app.post("/admin/faculty/<faculty_id>/delete")
def delete_faculty(faculty_id):
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    conn = get_db(); conn.execute("DELETE FROM faculty WHERE faculty_id=?", (faculty_id,)); conn.commit(); conn.close()
    return redirect(url_for("manage_faculty", message="Faculty member deleted."))


# ─── Parent alerts ─────────────────────────────────────────────────────

def _valid_parent_email(value):
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", (value or "").strip()))


def send_parent_email(parent_email, subject, message):
    """Send email only when all SMTP environment settings are present."""
    if not _valid_parent_email(parent_email):
        return "Not sent: no valid parent email", ""
    host = os.getenv("SMTP_HOST", "").strip()
    sender = os.getenv("SMTP_SENDER_EMAIL", "").strip()
    password = os.getenv("SMTP_PASSWORD", "")
    if not (host and sender and password):
        return "Not configured", "SMTP_HOST, SMTP_SENDER_EMAIL, or SMTP_PASSWORD is missing"
    try:
        port = int(os.getenv("SMTP_PORT", "587"))
        mail = EmailMessage(); mail["From"] = sender; mail["To"] = parent_email; mail["Subject"] = subject
        mail.set_content(message)
        with smtplib.SMTP(host, port, timeout=12) as server:
            server.starttls(context=ssl.create_default_context())
            server.login(sender, password); server.send_message(mail)
        return "Sent", ""
    except Exception as exc:
        return "Failed", str(exc)[:240]


def send_parent_sms(phone_number, message):
    """Send through Twilio when configured; never reports a send without a provider."""
    destination = _format_sms_number(phone_number)
    if not destination:
        return "Not sent: no valid parent phone", ""
    sid = os.getenv("TWILIO_ACCOUNT_SID", "").strip()
    token = os.getenv("TWILIO_AUTH_TOKEN", "").strip()
    sender = os.getenv("TWILIO_FROM_NUMBER", "").strip()
    if not (sid and token and sender):
        return "Not configured", "Twilio credentials are not configured"
    try:
        data = urlparse.urlencode({"To": destination, "From": sender, "Body": message}).encode()
        auth = __import__("base64").b64encode(f"{sid}:{token}".encode()).decode()
        req = urlrequest.Request(
            f"https://api.twilio.com/2010-04-01/Accounts/{sid}/Messages.json",
            data=data,
            headers={
                "Authorization": f"Basic {auth}",
                "Content-Type": "application/x-www-form-urlencoded",
            },
        )
        with urlrequest.urlopen(req, timeout=12) as response:
            if response.status not in (200, 201): raise RuntimeError(f"SMS provider returned {response.status}")
        return "Sent", ""
    except urlrequest.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        return "Failed", f"Twilio returned {exc.code}: {body[:200]}"
    except Exception as exc:
        return "Failed", str(exc)[:240]


def _create_parent_alert(conn, student, alert_type, condition_key, message, attendance_percentage=None, failed_subjects=""):
    """Create once per changed condition and record real delivery outcomes."""
    if conn.execute("SELECT 1 FROM parent_alerts WHERE condition_key=?", (condition_key,)).fetchone():
        return False
    email_subject = ("Attendance Alert – " if alert_type == "Attendance" else "Academic Result Alert – ") + student["name"]
    email_status, email_error = send_parent_email(student["parent_email"], email_subject, message)
    sms_status = "Disabled"
    errors = email_error
    conn.execute("""INSERT INTO parent_alerts
        (student_id,alert_type,condition_key,message,attendance_percentage,failed_subjects,created_date,email_status,sms_status,status,last_error)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (student["register_number"], alert_type, condition_key, message, attendance_percentage, failed_subjects,
         date.today().isoformat(), email_status, sms_status, "New", errors))
    return True


def check_attendance_alerts(conn, register_numbers, alert_month=None):
    alert_month = (alert_month or date.today().strftime("%Y-%m"))[:7]
    for rn in set(register_numbers):
        student = conn.execute("SELECT * FROM students WHERE register_number=?", (rn,)).fetchone()
        if not student: continue
        row = conn.execute("""SELECT SUM(CASE WHEN status='Present' THEN 1 ELSE 0 END) AS present_count,
            SUM(CASE WHEN status IN ('Present','Absent') THEN 1 ELSE 0 END) AS conducted
            FROM attendance WHERE register_number=? AND substr(date,1,7)=?""", (rn, alert_month)).fetchone()
        conducted = row["conducted"] or 0
        if not conducted: continue
        percentage = round((row["present_count"] or 0) / conducted * 100, 1)
        if percentage < 75:
            key = f"attendance:{rn}:{alert_month}"
            message = f"Dear Parent, your ward {student['name']}'s attendance for {alert_month} is {percentage}%, which is below the required 75%. Please take necessary action."
            _create_parent_alert(conn, student, "Attendance", key, message, attendance_percentage=percentage)


def check_failed_alerts(conn, register_numbers, alert_month=None):
    alert_month = (alert_month or date.today().strftime("%Y-%m"))[:7]
    thresholds = {"Internal 1": 9, "Internal 2": 18, "Semester Examination": 35}
    for rn in set(register_numbers):
        student = conn.execute("SELECT * FROM students WHERE register_number=?", (rn,)).fetchone()
        if not student: continue
        failed = sorted({row["subject"] for row in conn.execute("SELECT subject,test_name,marks_obtained FROM internal_marks WHERE register_number=? AND substr(test_date,1,7)=?", (rn, alert_month)) if row["test_name"] in thresholds and row["marks_obtained"] < thresholds[row["test_name"]]})
        if failed:
            subjects = ", ".join(failed); key = f"failed:{rn}:{alert_month}"
            message = f"Dear Parent, your ward {student['name']} has failed in {subjects} for {alert_month}. Please check the academic results in the Parent Portal."
            _create_parent_alert(conn, student, "Failed", key, message, failed_subjects=subjects)
# ─── Attendance ─────────────────────────────────────────────────────────

def _write_attendance_excel():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Attendance"
    headers = ["Date", "Course", "Year", "Semester", "Subject", "Register Number", "Status"]
    worksheet.append(headers)
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM attendance ORDER BY date ASC, course ASC, register_number ASC"
    ).fetchall()
    conn.close()
    for record in rows:
        worksheet.append(
            [
                record["date"],
                record["course"],
                record["year"],
                record["semester"],
                record["subject"],
                record["register_number"],
                record["status"],
            ]
        )
    for column in worksheet.columns:
        width = max(len(str(cell.value or "")) for cell in column) + 2
        worksheet.column_dimensions[column[0].column_letter].width = min(width, 30)

    workbook.save(os.getenv("ATTENDANCE_EXCEL_FILE", os.path.join(app.root_path, "attendance_records.xlsx")))


@app.route("/attendance", methods=["GET", "POST"])
def attendance():
    course = request.values.get("course", "BCA").strip()
    year = request.values.get("year", "1st Year").strip()
    semester = request.values.get("semester", "I Semester").strip()
    subject = request.values.get("subject", "Problem Solving using C").strip()
    language_group = request.values.get("language_group", "").strip()
    section = request.values.get("section", "").strip()
    attendance_date = request.values.get("attendance_date", date.today().isoformat()).strip()
    error = None
    saved = False

    try:
        if request.method == "POST":
            conn = get_db()
            for register_number, status in request.form.items():
                if not register_number.startswith("status_"):
                    continue
                rn = register_number.removeprefix("status_")
                conn.execute(
                    """INSERT INTO attendance
                       (date, course, year, semester, subject, register_number, status)
                       VALUES (?, ?, ?, ?, ?, ?, ?)
                       ON CONFLICT(date, course, year, semester, subject, register_number)
                       DO UPDATE SET status = excluded.status""",
                    (attendance_date, course, year, semester, subject, rn, status),
                )
            check_attendance_alerts(conn, [key.removeprefix("status_") for key in request.form if key.startswith("status_")], attendance_date[:7])
            conn.commit()
            conn.close()
            _write_attendance_excel()
            saved = True

        conn = get_db()
        legacy_semesters = {
            "I Semester": "1st Semester", "II Semester": "2nd Semester",
            "III Semester": "3rd Semester", "IV Semester": "4th Semester",
            "V Semester": "5th Semester", "VI Semester": "6th Semester",
        }
        # Existing records use both B.Com/B.com and Roman/ordinal semester labels.
        query = """SELECT * FROM students
                   WHERE lower(replace(course, '.', '')) = lower(replace(?, '.', ''))
                     AND year = ?
                     AND (semester = ? OR semester = ? OR COALESCE(semester, '') = '')"""
        params = [course, year, semester, legacy_semesters.get(semester, semester)]
        if section:
            query += " AND section = ?"; params.append(section)
        if language_group:
            query += " AND language_group = ?"; params.append(language_group)
        rows = conn.execute(query + " ORDER BY name ASC", params).fetchall()
        student_rows = [dict(row) for row in rows]
        status_rows = conn.execute(
            """SELECT register_number, status FROM attendance
               WHERE course = ? AND year = ? AND semester = ? AND subject = ? AND date = ?""",
            (course, year, semester, subject, attendance_date),
        ).fetchall()
        conn.close()
        statuses = {row["register_number"]: row["status"] for row in status_rows}
    except (OSError, sqlite3.Error):
        student_rows = []
        statuses = {}
        error = "Attendance could not be saved. Check the database and the Excel file location."

    counts = {"Present": 0, "Absent": 0, "Event": 0}
    for student in student_rows:
        status = statuses.get(student.get("register_number"), "Present")
        counts[status] = counts.get(status, 0) + 1
    return render_template(
        "attendance.html",
        students=student_rows,
        statuses=statuses,
        counts=counts,
        course=course,
        year=year,
        semester=semester,
        subject=subject,
        language_group=language_group,
        section=section,
        attendance_date=attendance_date,
        error=error,
        saved=saved,
    )


@app.post("/attendance/save")
def save_attendance_status():
    payload = request.get_json(silent=True) or request.form
    fields = {key: str(payload.get(key, "")).strip() for key in ("course", "year", "semester", "subject", "attendance_date", "register_number", "status")}
    if not all(fields.values()) or fields["status"] not in {"Present", "Absent", "Event"}:
        return jsonify(ok=False, message="Invalid attendance details."), 400
    try:
        conn = get_db()
        legacy_semesters = {
            "I Semester": "1st Semester", "II Semester": "2nd Semester",
            "III Semester": "3rd Semester", "IV Semester": "4th Semester",
            "V Semester": "5th Semester", "VI Semester": "6th Semester",
        }
        found = conn.execute("""SELECT 1 FROM students
            WHERE register_number = ?
              AND lower(replace(course, '.', '')) = lower(replace(?, '.', ''))
              AND year = ?
              AND (semester = ? OR semester = ? OR COALESCE(semester, '') = '')""",
            (fields["register_number"], fields["course"], fields["year"], fields["semester"],
             legacy_semesters.get(fields["semester"], fields["semester"]))).fetchone()
        if not found:
            conn.close(); return jsonify(ok=False, message="Student not found."), 404
        conn.execute("""INSERT INTO attendance (date, course, year, semester, subject, register_number, status) VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(date, course, year, semester, subject, register_number) DO UPDATE SET status = excluded.status""", (fields["attendance_date"], fields["course"], fields["year"], fields["semester"], fields["subject"], fields["register_number"], fields["status"]))
        check_attendance_alerts(conn, [fields["register_number"]], fields["attendance_date"][:7])
        conn.commit(); conn.close(); _write_attendance_excel()
        return jsonify(ok=True, message="Attendance saved automatically.")
    except (OSError, sqlite3.Error):
        return jsonify(ok=False, message="Attendance could not be saved."), 500


@app.get("/attendance/view")
def view_attendance():
    course = request.args.get("course", "BCA").strip(); year = request.args.get("year", "1st Year").strip(); semester = request.args.get("semester", "I Semester").strip(); subject = request.args.get("subject", "Problem Solving using C").strip(); month = request.args.get("month", date.today().strftime("%Y-%m")).strip()
    conn = get_db()
    students = [dict(row) for row in conn.execute("SELECT * FROM students WHERE course = ? AND year = ? ORDER BY name", (course, year)).fetchall()]
    rows = conn.execute("SELECT date, register_number, status FROM attendance WHERE course = ? AND year = ? AND semester = ? AND subject = ? AND substr(date, 1, 7) = ? ORDER BY date", (course, year, semester, subject, month)).fetchall(); conn.close()
    grouped = {student["register_number"]: [] for student in students}
    for row in rows: grouped.setdefault(row["register_number"], []).append(dict(row))
    report = []
    for student in students:
        records = grouped.get(student["register_number"], []); present = sum(row["status"] == "Present" for row in records); total = sum(row["status"] in {"Present", "Absent"} for row in records); percentage = round(present / total * 100, 1) if total else 0
        report.append({**student, "present": present, "total": total, "percentage": percentage})
    return render_template("attendance_view.html", report=report, course=course, year=year, semester=semester, subject=subject, month=month)


@app.get("/attendance/modify")
def modify_attendance():
    course = request.args.get("course", "BCA").strip(); year = request.args.get("year", "1st Year").strip(); semester = request.args.get("semester", "I Semester").strip(); subject = request.args.get("subject", "Problem Solving using C").strip(); attendance_date = request.args.get("attendance_date", date.today().isoformat()).strip()
    conn = get_db(); students = [dict(row) for row in conn.execute("SELECT * FROM students WHERE course = ? AND year = ? ORDER BY name", (course, year)).fetchall()]; rows = conn.execute("SELECT register_number, status FROM attendance WHERE course = ? AND year = ? AND semester = ? AND subject = ? AND date = ?", (course, year, semester, subject, attendance_date)).fetchall(); conn.close()
    return render_template("attendance_modify.html", students=students, statuses={row["register_number"]: row["status"] for row in rows}, course=course, year=year, semester=semester, subject=subject, attendance_date=attendance_date)

# ─── Internal Marks ─────────────────────────────────────────────────────

COURSES = ["BCA", "BBA", "BCom", "BA", "BSc"]
SEMESTERS = ["I Semester", "II Semester", "III Semester", "IV Semester", "V Semester", "VI Semester"]


def subject_label(subject):
    return f"{subject['subject_name']} ({subject['subject_code']})" if subject["subject_code"] else subject["subject_name"]


@app.route("/faculty/manage-subjects", methods=["GET", "POST"])
def manage_subjects():
    is_admin = bool(session.get("admin_username"))
    if not (is_admin or session.get("faculty_username")):
        return redirect(url_for("admin_login"))
    course = request.values.get("course", "BCA").strip()
    semester = request.values.get("semester", "I Semester").strip()
    error = None
    if request.method == "POST":
        action = request.form.get("action", "")
        name = request.form.get("subject_name", "").strip()
        code = request.form.get("subject_code", "").strip().upper()
        subject_id = request.form.get("subject_id", "").strip()
        try:
            conn = get_db()
            if action == "add" and name:
                conn.execute("INSERT INTO subjects (course, semester, subject_name, subject_code) VALUES (?, ?, ?, ?)", (course, semester, name, code))
            elif action == "edit" and name and subject_id:
                conn.execute("UPDATE subjects SET subject_name = ?, subject_code = ? WHERE id = ? AND course = ? AND semester = ?", (name, code, subject_id, course, semester))
            elif action == "delete" and subject_id:
                conn.execute("DELETE FROM subjects WHERE id = ? AND course = ? AND semester = ?", (subject_id, course, semester))
            else:
                raise ValueError
            conn.commit()
            conn.close()
            return redirect(url_for("manage_subjects", course=course, semester=semester, message="Subject saved successfully."))
        except (ValueError, sqlite3.IntegrityError):
            error = "Enter a unique subject name and code."
            if "conn" in locals(): conn.close()
    conn = get_db()
    subjects = [dict(row) for row in conn.execute("SELECT * FROM subjects WHERE course = ? AND semester = ? ORDER BY subject_name", (course, semester)).fetchall()]
    conn.close()
    return render_template("manage_subjects.html", courses=COURSES, semesters=SEMESTERS, course=course, semester=semester, subjects=subjects, message=request.args.get("message"), error=error, is_admin=is_admin)


@app.route("/internal-marks", methods=["GET", "POST"])
def internal_marks():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    course = request.values.get("course", "BCA").strip()
    year = request.values.get("year", "1st Year").strip()
    semester = request.values.get("semester", "I Semester").strip()
    semesters_by_year = {
        "1st Year": ["I Semester", "II Semester"],
        "2nd Year": ["III Semester", "IV Semester"],
        "3rd Year": ["V Semester", "VI Semester"],
    }
    available_semesters = semesters_by_year.get(year, SEMESTERS)
    if semester not in available_semesters:
        semester = available_semesters[0]
    assessment = request.values.get("assessment", "internal_1").strip()
    test_date = date.today().strftime("%d/%m/%Y")
    error = None
    message = None
    assessment_config = {
        "internal_1": {"title": "Internal 1", "maximum": 25, "internal": "Internal 1", "assignment": "Assignment 1", "attendance": "Attendance Internal 1", "total": 35},
        "internal_2": {"title": "Internal 2", "maximum": 50, "internal": "Internal 2", "assignment": "Assignment 2", "attendance": "Attendance Internal 2", "total": 60},
    }
    if assessment not in assessment_config:
        assessment = "internal_1"
    config = assessment_config[assessment]
    student_rows = []
    subjects = []
    marks = {}
    attendance = {}
    attendance_marks = {}
    selected_subject = ""
    try:
        conn = get_db()
        subjects = [dict(row) for row in conn.execute("SELECT * FROM subjects WHERE course = ? AND semester = ? ORDER BY subject_name", (course, semester)).fetchall()]
        configured_names = {item["subject_name"].strip().lower() for item in subjects}
        attendance_subjects = conn.execute(
            "SELECT DISTINCT subject FROM attendance WHERE course = ? AND year = ? AND semester = ? ORDER BY subject",
            (course, year, semester),
        ).fetchall()
        subjects.extend(
            {"subject_name": row["subject"], "subject_code": ""}
            for row in attendance_subjects
            if row["subject"].strip().lower() not in configured_names
        )
        valid_subject_labels = [subject_label(item) for item in subjects]
        requested_subject = request.values.get("subject", "").strip()
        selected_subject = requested_subject if requested_subject in valid_subject_labels else (valid_subject_labels[0] if valid_subject_labels else "")
        student_rows = [dict(row) for row in conn.execute(
            "SELECT * FROM students WHERE course = ? AND year = ? AND semester = ? ORDER BY name ASC",
            (course, year, semester),
        ).fetchall()]
        student_map = {student["register_number"]: student for student in student_rows}
        if request.method == "POST":
            uploaded_file = request.files.get("excel_file")
            if not uploaded_file or not uploaded_file.filename.lower().endswith(".xlsx"):
                raise ValueError("Choose a valid .xlsx Excel file.")
            workbook = load_workbook(uploaded_file, data_only=True)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                raise ValueError("The Excel file is empty.")
            headers = {str(value).strip().lower(): index for index, value in enumerate(rows[0]) if value is not None}
            required = ["reg no", "student name", f"{config['title'].lower()} /{config['maximum']}", f"assignment {1 if assessment == 'internal_1' else 2} /5"]
            missing = [column for column in required if column not in headers]
            if missing:
                raise ValueError("Missing Excel column(s): " + ", ".join(missing))
            validated = []
            seen = set()
            errors = []
            for excel_row, row in enumerate(rows[1:], start=2):
                if not any(value is not None and str(value).strip() for value in row):
                    continue
                def cell(column):
                    index = headers[column]
                    return row[index] if index < len(row) else None
                reg_no = str(cell("reg no") or "").strip()
                if not reg_no:
                    errors.append(f"Row {excel_row}: Registration Number is required.")
                    continue
                if reg_no in seen:
                    errors.append(f"Row {excel_row}: Duplicate Registration Number {reg_no}.")
                    continue
                seen.add(reg_no)
                if reg_no not in student_map:
                    errors.append(f"Row {excel_row}: Student not found for Registration Number {reg_no}.")
                    continue
                values = []
                for column, maximum in [(required[2], config["maximum"]), (required[3], 5)]:
                    raw = cell(column)
                    try:
                        score = int(raw)
                        if isinstance(raw, float) and raw != score:
                            raise ValueError
                        if score < 0 or score > maximum:
                            raise ValueError
                        values.append(score)
                    except (TypeError, ValueError):
                        errors.append(f"Row {excel_row}: {column} must be a whole number from 0 to {maximum}.")
                        break
                else:
                    validated.append((reg_no, values))
            if errors:
                raise ValueError("Excel upload failed. Please correct the errors: " + " ".join(errors[:6]))
            if not validated:
                raise ValueError("Excel upload failed. No valid marks were found.")
            for reg_no, values in validated:
                for component_name, score, maximum in [(config["internal"], values[0], config["maximum"]), (config["assignment"], values[1], 5)]:
                    conn.execute("""INSERT INTO internal_marks (course, year, semester, subject, test_name, register_number, marks_obtained, max_marks, test_date, remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(course, year, semester, subject, test_name, register_number) DO UPDATE SET marks_obtained=excluded.marks_obtained, max_marks=excluded.max_marks, test_date=excluded.test_date""", (course, year, semester, selected_subject, component_name, reg_no, score, maximum, test_date, "Imported from Excel"))
            check_failed_alerts(conn, [reg_no for reg_no, _ in validated], test_date[:7])
            conn.commit()
            message = f"{config['title']} marks uploaded successfully!"
        mark_rows = conn.execute("SELECT * FROM internal_marks WHERE course = ? AND year = ? AND semester = ? AND subject = ?", (course, year, semester, selected_subject)).fetchall()
        marks = {(row["register_number"], row["test_name"]): dict(row) for row in mark_rows}
        attendance_rows = conn.execute("SELECT register_number, status FROM attendance WHERE course = ? AND year = ? AND semester = ?", (course, year, semester)).fetchall()
        conn.close()
        attendance_by_student = {}
        for row in attendance_rows:
            if row["status"] not in {"Present", "Absent"}:
                continue
            item = attendance_by_student.setdefault(row["register_number"], {"total": 0, "present": 0})
            item["total"] += 1
            item["present"] += row["status"] == "Present"
        attendance = {rn: round(item["present"] / item["total"] * 100) if item["total"] else 0 for rn, item in attendance_by_student.items()}
        attendance_marks = {student["register_number"]: attendance_mark_from_percentage(attendance.get(student["register_number"], 0)) for student in student_rows}
    except (ValueError, sqlite3.Error) as exc:
        if "conn" in locals(): conn.close()
        marks = {}
        attendance = {}
        attendance_marks = {}
        error = str(exc) or "Excel upload failed. Please correct the highlighted errors and upload again."
    return render_template("internal_marks.html", students=student_rows, marks=marks, attendance=attendance, attendance_marks=attendance_marks, subjects=subjects, subject_label=subject_label, courses=COURSES, semesters=available_semesters, course=course, year=year, semester=semester, subject=selected_subject, assessment=assessment, config=config, error=error, message=message)
@app.route("/faculty/semester-results", methods=["GET"])
def semester_results():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))

    course = request.args.get("course", "BCA").strip() or "BCA"
    year = request.args.get("year", "1st Year").strip() or "1st Year"
    semester = request.args.get("semester", "I Semester").strip() or "I Semester"
    courses = ["BCA", "BBA", "BA", "BS", "B.Com", "B.Sc"]
    years = ["1st Year", "2nd Year", "3rd Year"]
    semesters = ["I Semester", "II Semester", "III Semester", "IV Semester", "V Semester", "VI Semester"]

    conn = get_db()
    students = conn.execute(
        "SELECT * FROM students WHERE course = ? AND year = ? ORDER BY name ASC",
        (course, year),
    ).fetchall()

    subject_rows = conn.execute(
        """
        SELECT subject FROM (
            SELECT DISTINCT subject AS subject FROM attendance WHERE course = ? AND year = ? AND semester = ?
            UNION
            SELECT DISTINCT subject AS subject FROM internal_marks WHERE course = ? AND year = ? AND semester = ?
        )
        ORDER BY subject ASC
        """,
        (course, year, semester, course, year, semester),
    ).fetchall()
    subject_names = [row["subject"] for row in subject_rows] or [
        "Data Structures (BCS302)", "FOC", "Programming in C", "Mathematics", "Kannada", "Hindi", "English"
    ]

    result_rows = []
    for student in students:
        student_result = {
            "register_number": student["register_number"],
            "name": student["name"],
            "course": student["course"],
            "semester": semester,
            "section": student["section"] or "—",
            "subjects": [],
            "total_marks": 0.0,
            "maximum_marks": 0,
            "sgpa": 0.0,
            "cgpa": 0.0,
            "overall_percentage": 0.0,
            "overall_grade": "F",
            "result_status": "FAIL",
        }

        grade_points = []
        total_percent = 0.0
        for subject in subject_names:
            component_row = {}
            for name, target_marks in {
                "Internal 1": 25,
                "Internal 2": 50,
                "Assignment 1": 5,
                "Attendance": 5,
            }.items():
                if name == "Attendance":
                    attendance_rows = conn.execute(
                        "SELECT status FROM attendance WHERE course = ? AND year = ? AND semester = ? AND register_number = ?",
                        (course, year, semester, student["register_number"]),
                    ).fetchall()
                    total_attendance = len(attendance_rows)
                    present_count = sum(1 for row in attendance_rows if row["status"] == "Present")
                    attendance_percentage = round((present_count / total_attendance) * 100, 2) if total_attendance else 0.0
                    attendance_marks = attendance_mark_from_percentage(attendance_percentage)
                    component_row["attendance_percentage"] = attendance_percentage
                    component_row["attendance_marks"] = round(attendance_marks, 2)
                    continue

                mark_record = conn.execute(
                    """
                    SELECT marks_obtained, max_marks
                    FROM internal_marks
                    WHERE course = ? AND year = ? AND semester = ? AND subject = ? AND register_number = ? AND test_name = ?
                    """,
                    (course, year, semester, subject, student["register_number"], name),
                ).fetchone()
                if mark_record:
                    component_row[name] = round(_scaled_component(mark_record["marks_obtained"], mark_record["max_marks"], target_marks), 2)
                else:
                    component_row[name] = 0.0

            subject_total = (
                component_row.get("Internal 1", 0.0)
                + component_row.get("Internal 2", 0.0)
                + component_row.get("Assignment 1", 0.0)
                + component_row.get("attendance_marks", 0.0)
            )
            percentage = round(subject_total, 2)
            grade = _grade_from_percentage(percentage)
            status = "PASS" if percentage >= 50 else "FAIL"
            grade_point = _grade_point_for_grade(grade)
            grade_points.append(grade_point)
            total_percent += percentage
            student_result["subjects"].append({
                "subject": subject,
                "first_internal": round(component_row.get("Internal 1", 0.0), 2),
                "second_internal": round(component_row.get("Internal 2", 0.0), 2),
                "assignment_1": round(component_row.get("Assignment 1", 0.0), 2),
                "attendance_percentage": round(component_row.get("attendance_percentage", 0.0), 2),
                "attendance_marks": round(component_row.get("attendance_marks", 0.0), 2),
                "total_marks": round(subject_total, 2),
                "percentage": percentage,
                "grade": grade,
                "status": status,
            })

        if student_result["subjects"]:
            student_result["total_marks"] = round(total_percent, 2)
            student_result["maximum_marks"] = len(student_result["subjects"]) * 85
            student_result["overall_percentage"] = round(total_percent / len(student_result["subjects"]), 2)
            student_result["overall_grade"] = _grade_from_percentage(student_result["overall_percentage"])
            student_result["result_status"] = "PASS" if student_result["overall_percentage"] >= 50 else "FAIL"
            student_result["sgpa"] = round(sum(grade_points) / len(grade_points), 2) if grade_points else 0.0
            student_result["cgpa"] = student_result["sgpa"]

        result_rows.append(student_result)

    conn.close()

    return render_template(
        "semester_results.html",
        course=course,
        year=year,
        semester=semester,
        courses=courses,
        years=years,
        semesters=semesters,
        results=result_rows,
        subjects=subject_names,
    )


@app.route("/admin/timetable", methods=["GET", "POST"])
def manage_timetable():
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    course = request.values.get("course", "BCA").strip()
    year = request.values.get("year", "1st Year").strip()
    semester = request.values.get("semester", "I Semester").strip()
    section = request.values.get("section", "A").strip()
    message = None
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add_period":
            conn.execute("INSERT INTO timetable_periods (label,start_time,end_time,is_saturday) VALUES (?,?,?,?)", (request.form["label"].strip(), request.form["start_time"], request.form["end_time"], 1 if request.form.get("saturday") else 0))
        elif action == "save_entry":
            faculty_id = request.form.get("faculty_id", "").strip()
            faculty_row = conn.execute(
                "SELECT name FROM faculty WHERE faculty_id=?", (faculty_id,)
            ).fetchone() if faculty_id else None
            if not faculty_row:
                message = "Select a valid faculty member before saving the timetable entry."
            else:
                faculty_name = faculty_row["name"]
                conn.execute("""INSERT INTO timetable_entries
                    (course,year,semester,section,day,period_id,subject_id,faculty_name,faculty_id,lab_batch,room,activity_type)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(course,year,semester,section,day,period_id,lab_batch)
                    DO UPDATE SET subject_id=excluded.subject_id, faculty_name=excluded.faculty_name,
                        faculty_id=excluded.faculty_id, room=excluded.room, activity_type=excluded.activity_type""",
                    (course, year, semester, section, request.form["day"], request.form["period_id"],
                     request.form.get("subject_id") or None, faculty_name, faculty_id,
                     request.form.get("lab_batch", "").strip(), request.form.get("room", "").strip(),
                     request.form.get("activity_type", "").strip()))
        elif action == "delete_entry":
            conn.execute("DELETE FROM timetable_entries WHERE id = ?", (request.form["entry_id"],))
        conn.commit(); message = message or "Timetable updated successfully."
    subjects = conn.execute("SELECT * FROM subjects WHERE course=? AND semester=? ORDER BY subject_name", (course, semester)).fetchall()
    periods = conn.execute("SELECT * FROM timetable_periods ORDER BY start_time").fetchall()
    rows = conn.execute("SELECT t.*, p.label,p.start_time,p.end_time,s.subject_name FROM timetable_entries t JOIN timetable_periods p ON p.id=t.period_id LEFT JOIN subjects s ON s.id=t.subject_id WHERE t.course=? AND t.year=? AND t.semester=? AND t.section=? ORDER BY p.start_time", (course,year,semester,section)).fetchall()
    faculty = _faculty_rows(conn)
    conn.close()
    return render_template("manage_timetable.html", course=course, year=year, semester=semester, section=section, subjects=subjects, periods=periods, faculty=faculty, entries=[dict(r) for r in rows], message=message)


@app.get("/faculty/my-timetable")
def faculty_timetable():
    faculty_id = session.get("faculty_username")
    if not faculty_id:
        return redirect(url_for("faculty_login"))

    conn = get_db()
    faculty = conn.execute(
        "SELECT faculty_id, name, department FROM faculty WHERE faculty_id = ?",
        (faculty_id,),
    ).fetchone()
    faculty_name = faculty["name"] if faculty else ""
    entries = conn.execute(
        """SELECT t.*, p.start_time, p.end_time, s.subject_name
           FROM timetable_entries t
           JOIN timetable_periods p ON p.id = t.period_id
           LEFT JOIN subjects s ON s.id = t.subject_id
           WHERE lower(trim(COALESCE(t.faculty_id, ''))) = lower(trim(?))
              OR (
                  trim(COALESCE(t.faculty_id, '')) = ''
                  AND lower(trim(COALESCE(t.faculty_name, ''))) = lower(trim(?))
              )
           ORDER BY CASE lower(trim(t.day))
               WHEN 'monday' THEN 1 WHEN 'tuesday' THEN 2
               WHEN 'wednesday' THEN 3 WHEN 'thursday' THEN 4
               WHEN 'friday' THEN 5 WHEN 'saturday' THEN 6 ELSE 7 END,
               p.start_time""",
        (faculty_id, faculty_name),
    ).fetchall()
    conn.close()

    return render_template(
        "faculty_timetable.html",
        entries=[dict(entry) for entry in entries],
        faculty=dict(faculty) if faculty else None,
    )

def _event_message(event, students, subject, advisor_name):
    """Build the stored in-app message shown to the selected recipient."""
    student_lines = "; ".join(f"{student['name']} ({student['register_number']})" for student in students)
    return (
        f"Attendance claim for {event['name']} | Date: {event['event_date']} | "
        f"Time: {event['start_time'] or '—'} to {event['end_time'] or '—'} | "
        f"Class: {event['course']} / {event['year']} / {event['semester']} / Section {event['section'] or '—'} | "
        f"Subject: {subject} | Recipient: {advisor_name} | Students: {student_lines}"
    )


def _add_event_notification(conn, claim_id, recipient_role, recipient_name, notification_type, message):
    conn.execute(
        """INSERT OR IGNORE INTO event_notifications
           (claim_id,recipient_role,recipient_name,notification_type,message,created_on)
           VALUES (?,?,?,?,?,?)""",
        (claim_id, recipient_role, recipient_name, notification_type, message, date.today().isoformat()),
    )


def _event_claims_with_reminders(conn):
    """Return open claims and create one persistent day-three reminder per stage.

    A reminder never changes a claim's state; a request remains actionable until
    an advisor/teacher claims it.
    """
    claims = [dict(row) for row in conn.execute("""SELECT c.*, e.name, e.event_date, e.start_time,
        e.end_time, e.course, e.year, e.semester, e.section, e.created_at, e.created_by
        FROM event_claims c JOIN events e ON e.id=c.event_id ORDER BY e.event_date DESC, c.id DESC""").fetchall()]
    today = date.today()
    for claim in claims:
        if claim["status"] == "ATTENDANCE CLAIMED":
            claim["reminder"] = ""
            continue
        anchor = claim["message_sent_on"] or claim["created_on"] or claim["created_at"] or claim["event_date"]
        try:
            is_due = today >= date.fromisoformat(anchor) + timedelta(days=3)
        except (TypeError, ValueError):
            is_due = False
        claim["reminder"] = ""
        if not is_due:
            continue
        if claim["message_sent_on"]:
            recipient_role, recipient_name, notification_type = "Advisor / Subject Teacher", claim["advisor_name"], "CLAIM_OVERDUE"
            reminder = f"Attendance claim pending: please claim E (College Event) attendance for {claim['name']} on {claim['event_date']}."
        else:
            recipient_role, recipient_name, notification_type = "Event Coordinator", claim["created_by"], "SEND_OVERDUE"
            reminder = f"Attendance claim for {claim['name']} on {claim['event_date']} has not been sent to {claim['advisor_name']}."
        _add_event_notification(conn, claim["id"], recipient_role, recipient_name, notification_type, reminder)
        claim["reminder"] = reminder
        claim["reminder_for"] = recipient_role
    conn.commit()
    return claims


def _faculty_display_name(conn, faculty_key):
    """Return a readable faculty name from either a faculty_id or a stored name."""
    if not faculty_key:
        return ""
    faculty_key = str(faculty_key).strip()
    if not faculty_key:
        return ""
    row = conn.execute(
        "SELECT name, faculty_id FROM faculty WHERE lower(trim(faculty_id)) = lower(trim(?)) OR lower(trim(name)) = lower(trim(?)) LIMIT 1",
        (faculty_key, faculty_key),
    ).fetchone()
    if row:
        return row["name"]
    return faculty_key


def _faculty_matches_user(conn, advisor_name, faculty_username):
    """Return True when the stored advisor value refers to the logged-in faculty user."""
    advisor_value = str(advisor_name or "").strip()
    username = str(faculty_username or "").strip()
    if not advisor_value or not username:
        return False

    row = conn.execute(
        "SELECT faculty_id, name FROM faculty WHERE lower(trim(faculty_id)) = lower(trim(?)) OR lower(trim(name)) = lower(trim(?)) LIMIT 1",
        (advisor_value, advisor_value),
    ).fetchone()
    user_row = conn.execute(
        "SELECT faculty_id, name FROM faculty WHERE lower(trim(faculty_id)) = lower(trim(?)) OR lower(trim(name)) = lower(trim(?)) LIMIT 1",
        (username, username),
    ).fetchone()

    candidates = {
        advisor_value.casefold(),
        _faculty_display_name(conn, advisor_value).casefold(),
    }
    if row:
        candidates.add(str(row["faculty_id"]).casefold())
        candidates.add(str(row["name"]).casefold())
    user_candidates = {
        username.casefold(),
        _faculty_display_name(conn, username).casefold(),
    }
    if user_row:
        user_candidates.add(str(user_row["faculty_id"]).casefold())
        user_candidates.add(str(user_row["name"]).casefold())

    return bool(candidates & user_candidates)


def _event_recipients(conn, course, year, semester, section):
    """Use existing timetable faculty assignments as event-claim recipients."""
    rows = conn.execute("""SELECT DISTINCT
        COALESCE(s.subject_name, t.activity_type) AS subject,
        COALESCE(NULLIF(trim(t.faculty_id), ''), trim(t.faculty_name)) AS faculty_id,
        COALESCE(
            (SELECT f.name FROM faculty f WHERE lower(trim(f.faculty_id)) = lower(trim(COALESCE(NULLIF(trim(t.faculty_id), ''), trim(t.faculty_name)))) LIMIT 1),
            NULLIF(trim(t.faculty_name), ''),
            NULLIF(trim(t.faculty_id), '')
        ) AS faculty_name
        FROM timetable_entries t
        LEFT JOIN subjects s ON s.id=t.subject_id
        WHERE t.course=? AND t.year=? AND t.semester=? AND t.section=?
          AND (trim(COALESCE(t.faculty_name, '')) <> '' OR trim(COALESCE(t.faculty_id, '')) <> '')
        ORDER BY subject, faculty_name""",
        (course, year, semester, section)).fetchall()
    return [dict(row) for row in rows if row["subject"]]


def _event_subjects(conn, course, year, semester, section):
    """Return class subjects even before a timetable teacher is assigned."""
    rows = conn.execute("""SELECT subject FROM (
        SELECT DISTINCT COALESCE(s.subject_name, t.activity_type) AS subject
          FROM timetable_entries t LEFT JOIN subjects s ON s.id=t.subject_id
         WHERE t.course=? AND t.year=? AND t.semester=? AND t.section=?
        UNION
        SELECT DISTINCT subject FROM attendance
         WHERE course=? AND year=? AND semester=?
        UNION
        SELECT DISTINCT subject_name AS subject FROM subjects
         WHERE course=? AND semester=?
    ) WHERE trim(COALESCE(subject, '')) <> '' ORDER BY subject""",
        (course, year, semester, section, course, year, semester, course, semester)).fetchall()
    return [row["subject"] for row in rows]


def _event_students(conn, course, year, semester, section):
    query = "SELECT * FROM students WHERE course=? AND year=?"
    params = [course, year]
    if semester:
        query += " AND semester=?"; params.append(semester)
    if section:
        query += " AND section=?"; params.append(section)
    return [dict(row) for row in conn.execute(query + " ORDER BY name", params).fetchall()]


@app.route("/event-coordinator", methods=["GET", "POST"])
def event_coordinator():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    course = request.values.get("course", "BCA").strip()
    year = request.values.get("year", "1st Year").strip()
    semester = request.values.get("semester", "I Semester").strip()
    section = request.values.get("section", "A").strip()
    message = request.args.get("message", "")
    conn = get_db()
    if request.method == "POST" and request.form.get("action") == "create_event":
        fields = {key: request.form.get(key, "").strip() for key in ("name", "event_date", "start_time", "end_time", "course", "year", "semester", "section")}
        selected_students = request.form.getlist("students")
        valid_students = _event_students(conn, fields["course"], fields["year"], fields["semester"], fields["section"])
        valid_numbers = {student["register_number"] for student in valid_students}
        selected_students = [number for number in selected_students if number in valid_numbers]
        if not all(fields.values()) or not selected_students:
            message = "Complete all event details and select at least one student."
        else:
            cur = conn.execute("""INSERT INTO events
                (name,event_date,start_time,end_time,course,year,semester,section,created_at,created_by)
                VALUES (?,?,?,?,?,?,?,?,?,?)""", (*fields.values(), date.today().isoformat(), session["faculty_username"]))
            conn.executemany("INSERT INTO event_students (event_id,register_number) VALUES (?,?)", [(cur.lastrowid, number) for number in selected_students])
            conn.commit()
            message = "Event saved. Prepare a claim below, then send it when ready."
    students = _event_students(conn, course, year, semester, section)
    events = [dict(row) for row in conn.execute("SELECT * FROM events ORDER BY event_date DESC, id DESC").fetchall()]
    selected_event_id = request.args.get("event_id", type=int)
    selected_event = next((event for event in events if event["id"] == selected_event_id), None)
    claim_course = selected_event["course"] if selected_event else course
    claim_year = selected_event["year"] if selected_event else year
    claim_semester = selected_event["semester"] if selected_event else semester
    claim_section = selected_event["section"] if selected_event else section
    recipients = _event_recipients(conn, claim_course, claim_year, claim_semester, claim_section)
    claim_subjects = _event_subjects(conn, claim_course, claim_year, claim_semester, claim_section)
    claims = _event_claims_with_reminders(conn)
    notifications = [dict(row) for row in conn.execute("""SELECT n.* FROM event_notifications n
        WHERE n.recipient_role='Event Coordinator' AND (n.recipient_name='' OR n.recipient_name=?)
        ORDER BY n.created_on DESC, n.id DESC""", (session["faculty_username"],)).fetchall()]
    conn.close()
    return render_template("event_coordinator.html", students=students, events=events, recipients=recipients, claim_subjects=claim_subjects,
        claims=claims, notifications=notifications, course=course, year=year, semester=semester, section=section,
        selected_event_id=selected_event_id, claim_course=claim_course, claim_year=claim_year,
        claim_semester=claim_semester, claim_section=claim_section, message=message)


@app.post("/event-coordinator/prepare-claim")
def prepare_event_claim():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    event_id = request.form.get("event_id", type=int)
    subject = request.form.get("subject", "").strip()
    advisor = request.form.get("advisor_name", "").strip()
    conn = get_db()
    event = conn.execute("SELECT * FROM events WHERE id=?", (event_id,)).fetchone()
    if not event or not subject or not advisor:
        conn.close(); return redirect(url_for("event_coordinator", message="Choose an event, subject, and recipient first."))
    recipients = _event_recipients(conn, event["course"], event["year"], event["semester"], event["section"])
    valid_faculty = {item["faculty_name"] for item in recipients}
    valid_faculty_ids = {item["faculty_id"] for item in recipients if item.get("faculty_id")}
    recipient_pairs = {(item["subject"], item["faculty_name"]) for item in recipients} | {(item["subject"], item["faculty_id"]) for item in recipients if item.get("faculty_id")}
    if (subject, advisor) not in recipient_pairs and advisor not in valid_faculty and advisor not in valid_faculty_ids:
        conn.close(); return redirect(url_for("event_coordinator", message="Choose the teacher assigned to that subject in the timetable."))
    conn.execute("""INSERT INTO event_claims (event_id,subject,advisor_name,status,created_on)
        VALUES (?,?,?,'PENDING',?) ON CONFLICT(event_id,subject,advisor_name) DO NOTHING""",
        (event_id, subject, advisor, date.today().isoformat()))
    conn.commit(); conn.close()
    return redirect(url_for("event_coordinator", message="Claim draft prepared. Send it when ready."))


@app.post("/event-coordinator/send-claim/<int:claim_id>")
def send_event_claim(claim_id):
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    conn = get_db()
    claim = conn.execute("""SELECT c.*, e.name, e.event_date, e.start_time, e.end_time, e.course, e.year,
        e.semester, e.section FROM event_claims c JOIN events e ON e.id=c.event_id WHERE c.id=?""", (claim_id,)).fetchone()
    if not claim or claim["status"] == "ATTENDANCE CLAIMED":
        conn.close(); return redirect(url_for("event_coordinator", message="That claim is unavailable."))
    students = [dict(row) for row in conn.execute("""SELECT s.name, s.register_number FROM event_students es
        JOIN students s ON s.register_number=es.register_number WHERE es.event_id=? ORDER BY s.name""", (claim["event_id"],)).fetchall()]
    advisor_display = _faculty_display_name(conn, claim["advisor_name"])
    text = _event_message(claim, students, claim["subject"], advisor_display)
    conn.execute("UPDATE event_claims SET status='MESSAGE SENT', message_text=?, message_sent_on=? WHERE id=?",
        (text, date.today().isoformat(), claim_id))
    _add_event_notification(conn, claim_id, "Advisor / Subject Teacher", claim["advisor_name"], "CLAIM_MESSAGE", text)
    conn.commit(); conn.close()
    return redirect(url_for("event_coordinator", message="Attendance claim message sent."))


@app.get("/faculty/event-claims")
def faculty_event_claims():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    conn = get_db()
    claims = _event_claims_with_reminders(conn)
    username = session["faculty_username"]
    faculty = conn.execute(
        "SELECT faculty_id, name, department FROM faculty WHERE faculty_id=?",
        (username,),
    ).fetchone()
    assigned_claims = []
    for claim in claims:
        if _faculty_matches_user(conn, claim.get("advisor_name", ""), username):
            assigned_claims.append(claim)
    notifications = [dict(row) for row in conn.execute("""SELECT * FROM event_notifications
        WHERE recipient_role='Advisor / Subject Teacher' AND recipient_name=?
        ORDER BY created_on DESC, id DESC""", (username,)).fetchall()]
    conn.close()
    return render_template("faculty_event_claims.html", claims=assigned_claims, notifications=notifications,
        faculty=dict(faculty) if faculty else {"faculty_id": username, "name": username, "department": ""})


@app.post("/faculty/event-claims/<int:claim_id>/claim")
def claim_event_attendance(claim_id):
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    conn = get_db()
    claim = conn.execute("""SELECT c.*, e.event_date, e.course, e.year, e.semester FROM event_claims c
        JOIN events e ON e.id=c.event_id WHERE c.id=?""", (claim_id,)).fetchone()
    if not claim or not _faculty_matches_user(conn, claim["advisor_name"], session["faculty_username"]) or not claim["message_sent_on"]:
        conn.close(); return redirect(url_for("faculty_event_claims"))
    students = conn.execute("SELECT register_number FROM event_students WHERE event_id=?", (claim["event_id"],)).fetchall()
    for student in students:
        # Existing attendance persistence uses the Event value; its UI labels it
        # as E (College Event), alongside Present and Absent.
        conn.execute("""INSERT INTO attendance (date,course,year,semester,subject,register_number,status)
            VALUES (?,?,?,?,?,?, 'Event') ON CONFLICT(date,course,year,semester,subject,register_number)
            DO UPDATE SET status='Event'""", (claim["event_date"], claim["course"], claim["year"], claim["semester"], claim["subject"], student["register_number"]))
    conn.execute("UPDATE event_claims SET status='ATTENDANCE CLAIMED', claimed_on=? WHERE id=?", (date.today().isoformat(), claim_id))
    conn.commit(); conn.close(); _write_attendance_excel()
    return redirect(url_for("faculty_event_claims"))


# ─── Auth ───────────────────────────────────────────────────────────────

@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        admin_username = os.getenv("ADMIN_USERNAME", "")
        admin_password = os.getenv("ADMIN_PASSWORD", "")
        if admin_username and admin_password and username == admin_username and password == admin_password:
            session["admin_username"] = username
            return redirect(url_for("admin_dashboard"))
        error = "Invalid username or password."
    return render_template("admin_login.html", error=error)


@app.route("/parent-alerts", methods=["GET", "POST"])
def parent_alerts():
    if not (session.get("admin_username") or session.get("faculty_username")):
        return redirect(url_for("admin_login"))
    conn = get_db()
    if request.method == "POST":
        alert_id = request.form.get("alert_id", type=int)
        status = request.form.get("status", "").strip()
        if alert_id and status in {"New", "Read", "Resolved"}:
            conn.execute("UPDATE parent_alerts SET status=? WHERE id=?", (status, alert_id)); conn.commit()
    alert_type = request.args.get("type", "").strip()
    status_filter = request.args.get("status", "").strip()
    clauses, params = [], []
    if alert_type in {"Attendance", "Failed"}: clauses.append("a.alert_type=?"); params.append(alert_type)
    if status_filter in {"New", "Read", "Resolved"}: clauses.append("a.status=?"); params.append(status_filter)
    where = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    rows = conn.execute("""SELECT a.*, s.name, s.register_number, s.parent_name, s.parent_phone, s.parent_email,
        s.course, s.year FROM parent_alerts a
        JOIN students s ON s.register_number=a.student_id""" + where + " ORDER BY a.id DESC", params).fetchall()
    current_month = date.today().strftime("%Y-%m")
    monthly_counts = conn.execute("""SELECT COUNT(*) AS total,
        SUM(CASE WHEN alert_type='Attendance' THEN 1 ELSE 0 END) AS attendance,
        SUM(CASE WHEN alert_type='Failed' THEN 1 ELSE 0 END) AS academic
        FROM parent_alerts WHERE substr(created_date,1,7)=?""", (current_month,)).fetchone()
    conn.close()
    return render_template("parent_alerts.html", alerts=[dict(row) for row in rows], alert_type=alert_type,
        status_filter=status_filter, current_month=current_month,
        alert_counts={"total": monthly_counts["total"] or 0, "attendance": monthly_counts["attendance"] or 0,
                      "academic": monthly_counts["academic"] or 0, "general": 0})


@app.post("/parent-alerts/<int:alert_id>/retry")
def retry_parent_alert(alert_id):
    if not (session.get("admin_username") or session.get("faculty_username")):
        return redirect(url_for("admin_login"))
    conn = get_db()
    row = conn.execute("""SELECT a.*,s.parent_phone,s.parent_email FROM parent_alerts a
        JOIN students s ON s.register_number=a.student_id WHERE a.id=?""", (alert_id,)).fetchone()
    if row:
        email_status, email_error = send_parent_email(row["parent_email"], "Parent Alert", row["message"])
        sms_status = "Disabled"
        errors = email_error
        conn.execute("UPDATE parent_alerts SET email_status=?,sms_status=?,last_error=? WHERE id=?",
                     (email_status, sms_status, errors, alert_id))
        conn.commit()
    conn.close()
    return redirect(url_for("parent_alerts"))

@app.get("/admin-dashboard")
def admin_dashboard():
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    low_rows = _low_attendance_rows({})
    conn = get_db()
    rows = conn.execute("SELECT register_number, SUM(CASE WHEN status='Present' THEN 1 ELSE 0 END) AS present_count, SUM(CASE WHEN status IN ('Present','Absent') THEN 1 ELSE 0 END) AS conducted FROM attendance GROUP BY register_number").fetchall()
    failed_rows = _failed_subject_rows(conn)
    timetable_count = conn.execute("SELECT COUNT(*) AS count FROM timetable_entries").fetchone()["count"]
    timetable = conn.execute("""SELECT t.day,p.label,p.start_time,p.end_time,COALESCE(s.subject_name,t.activity_type,'—') AS title
        FROM timetable_entries t JOIN timetable_periods p ON p.id=t.period_id LEFT JOIN subjects s ON s.id=t.subject_id
        ORDER BY CASE t.day WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3 WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 ELSE 6 END,p.start_time LIMIT 30""").fetchall()
    conn.close()
    low_count = len(low_rows)
    return render_template("admin_dashboard.html", low_count=low_count, low_rows=low_rows[:5], failed_count=len({row['register_number'] for row in failed_rows}), failed_rows=failed_rows[:5], timetable_count=timetable_count, timetable=[dict(r) for r in timetable])


def _failed_subject_rows(conn):
    """Return failed subjects using the same combined /85 rule as Semester Results."""
    rows = conn.execute("""SELECT DISTINCT m.register_number,m.course,m.year,m.semester,m.subject,
        s.name,s.section FROM internal_marks m
        JOIN students s ON s.register_number=m.register_number
        ORDER BY s.name,m.subject""").fetchall()
    failed = []
    for row in rows:
        item = dict(row)
        mark_rows = conn.execute("""SELECT test_name,marks_obtained,max_marks FROM internal_marks
            WHERE register_number=? AND course=? AND year=? AND semester=? AND subject=?""",
            (item["register_number"], item["course"], item["year"], item["semester"], item["subject"])).fetchall()
        components = {mark["test_name"]: mark for mark in mark_rows}
        first = components.get("Internal 1")
        second = components.get("Internal 2")
        assignment = components.get("Assignment 1")
        first_marks = _scaled_component(first["marks_obtained"], first["max_marks"], 25) if first else 0.0
        second_marks = _scaled_component(second["marks_obtained"], second["max_marks"], 50) if second else 0.0
        assignment_marks = _scaled_component(assignment["marks_obtained"], assignment["max_marks"], 5) if assignment else 0.0
        attendance_rows = conn.execute("""SELECT status FROM attendance
            WHERE register_number=? AND course=? AND year=? AND semester=?""",
            (item["register_number"], item["course"], item["year"], item["semester"])).fetchall()
        total_attendance = len(attendance_rows)
        present_count = sum(1 for attendance in attendance_rows if attendance["status"] == "Present")
        attendance_percentage = (present_count / total_attendance * 100) if total_attendance else 0.0
        attendance_marks = attendance_mark_from_percentage(attendance_percentage)
        subject_total = round(first_marks + second_marks + assignment_marks + attendance_marks, 2)
        if subject_total < 50:
            item["exam_type"] = "Overall Subject Result"
            item["marks_obtained"] = subject_total
            item["maximum_marks"] = 85
            item["overall_result"] = "FAILED"
            failed.append(item)
    return failed


@app.get("/admin/failed-subjects")
def admin_failed_subjects():
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    conn = get_db()
    rows = _failed_subject_rows(conn)
    conn.close()
    return render_template("failed_subjects_report.html", rows=rows)


@app.get("/admin-logout")
def admin_logout():
    session.pop("admin_username", None)
    return redirect(url_for("home"))


def _low_attendance_rows(filters):
    """Return students below 75% using the same attendance records as faculty."""
    clauses, params = ["1=1"], []
    for column in ("course", "year", "semester"):
        if filters.get(column):
            clauses.append(f"a.{column} = ?"); params.append(filters[column])
    if filters.get("section"):
        clauses.append("s.section = ?"); params.append(filters["section"])
    if filters.get("month"):
        clauses.append("substr(a.date, 1, 7) = ?"); params.append(filters["month"])
    conn = get_db()
    query = f"""SELECT s.*, COUNT(CASE WHEN a.status IN ('Present','Absent') THEN 1 END) AS conducted,
        SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) AS present_count
        FROM students s JOIN attendance a ON a.register_number=s.register_number
        WHERE {' AND '.join(clauses)} GROUP BY s.register_number"""
    rows = []
    for row in conn.execute(query, params).fetchall():
        item = dict(row); conducted = item["conducted"] or 0
        item["attendance_percentage"] = round(item["present_count"] / conducted * 100, 1) if conducted else 0
        if conducted and item["attendance_percentage"] < 75:
            item["alert_status"] = "Sent" if item.get("parent_phone") else "Pending"
            rows.append(item)
    conn.close()
    return sorted(rows, key=lambda item: item["attendance_percentage"])


@app.get("/admin/low-attendance")
def admin_low_attendance():
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    filters = {key: request.args.get(key, "").strip() for key in ("course", "year", "semester", "section", "month", "advisor")}
    return render_template("low_attendance_report.html", rows=_low_attendance_rows(filters), filters=filters)


@app.get("/admin/low-attendance/<register_number>")
def admin_student_attendance_details(register_number):
    if not session.get("admin_username"):
        return redirect(url_for("admin_login"))
    conn = get_db()
    student_row = conn.execute("SELECT * FROM students WHERE register_number=?", (register_number,)).fetchone()
    if not student_row:
        conn.close(); return redirect(url_for("admin_low_attendance"))
    subject_rows = conn.execute("""SELECT subject, COUNT(CASE WHEN status IN ('Present','Absent') THEN 1 END) AS conducted,
        SUM(CASE WHEN status='Present' THEN 1 ELSE 0 END) AS present_count,
        SUM(CASE WHEN status='Absent' THEN 1 ELSE 0 END) AS absent_count,
        SUM(CASE WHEN status='Event' THEN 1 ELSE 0 END) AS event_count
        FROM attendance WHERE register_number=? GROUP BY subject ORDER BY subject""", (register_number,)).fetchall()
    conn.close()
    subjects = []
    total_present = total_conducted = 0
    for row in subject_rows:
        item = dict(row); total_present += item["present_count"] or 0; total_conducted += item["conducted"] or 0
        item["percentage"] = round((item["present_count"] or 0) / item["conducted"] * 100, 1) if item["conducted"] else 0
        subjects.append(item)
    overall = round(total_present / total_conducted * 100, 1) if total_conducted else 0
    return render_template("student_attendance_details.html", student=dict(student_row), subjects=subjects, overall=overall)


@app.route("/faculty-login", methods=["GET", "POST"])
def faculty_login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        conn = get_db()
        faculty = conn.execute("SELECT * FROM faculty WHERE faculty_id=?", (username,)).fetchone()
        conn.close()
        if faculty and check_password_hash(faculty["password_hash"], password):
            session["faculty_username"] = faculty["faculty_id"]
            return redirect(url_for("faculty_dashboard"))
        error = "Invalid username or password."
    return render_template("faculty_login.html", error=error)


@app.get("/faculty-logout")
def faculty_logout():
    session.pop("faculty_username", None)
    return redirect(url_for("faculty_login"))


@app.get("/faculty-dashboard")
def faculty_dashboard():
    if not session.get("faculty_username"):
        return redirect(url_for("faculty_login"))
    attendance_course = request.args.get("course", "All Courses").strip()
    attendance_year = request.args.get("year", "All Years").strip()
    attendance_filter = {}
    if attendance_course != "All Courses":
        attendance_filter["course"] = attendance_course
    if attendance_year != "All Years":
        attendance_filter["year"] = attendance_year

    attendance_counts = {"Present": 0, "Absent": 0, "Event": 0}
    course_attendance = []
    weekly_attendance = []
    try:
        conn = get_db()
        total_students = conn.execute("SELECT COUNT(*) AS c FROM students").fetchone()["c"]

        today = date.today().isoformat()
        where = ["date = ?"]
        params = [today]
        if "course" in attendance_filter:
            where.append("course = ?")
            params.append(attendance_filter["course"])
        if "year" in attendance_filter:
            where.append("year = ?")
            params.append(attendance_filter["year"])
        status_rows = conn.execute(
            f"SELECT status FROM attendance WHERE {' AND '.join(where)}", params
        ).fetchall()
        for row in status_rows:
            status = row["status"]
            attendance_counts[status] = attendance_counts.get(status, 0) + 1

        for course in ["BCA", "BBA", "BA", "BS", "B.Com", "B.Sc"]:
            course_params = [today, course]
            course_where = "date = ? AND course = ?"
            if attendance_year != "All Years":
                course_where += " AND year = ?"
                course_params.append(attendance_year)
            counts = {"Present": 0, "Absent": 0, "Event": 0}
            for row in conn.execute(
                f"SELECT status FROM attendance WHERE {course_where}", course_params
            ).fetchall():
                status = row["status"]
                counts[status] = counts.get(status, 0) + 1
            course_attendance.append({"course": course, "counts": counts})

        for days_ago in range(5, -1, -1):
            graph_date = date.today() - timedelta(days=days_ago)
            day_params = [graph_date.isoformat()]
            day_where = "date = ?"
            if attendance_course != "All Courses":
                day_where += " AND course = ?"
                day_params.append(attendance_course)
            if attendance_year != "All Years":
                day_where += " AND year = ?"
                day_params.append(attendance_year)
            day_records = conn.execute(
                f"SELECT status FROM attendance WHERE {day_where}", day_params
            ).fetchall()
            present_count = sum(1 for r in day_records if r["status"] == "Present")
            weekly_attendance.append(
                {
                    "label": graph_date.strftime("%a"),
                    "percentage": round(present_count / len(day_records) * 100) if day_records else 0,
                }
            )
        conn.close()
    except sqlite3.Error:
        return render_template(
            "faculty_dashboard.html",
            total_students=0,
            attendance_counts=attendance_counts,
            course_attendance=course_attendance,
            weekly_attendance=weekly_attendance,
            error="Database is unavailable.",
        )
    return render_template(
        "faculty_dashboard.html",
        total_students=total_students,
        dashboard_iso_date=date.today().isoformat(),
        dashboard_date=date.today().strftime("%d %B %Y, %A"),
        attendance_course=attendance_course,
        attendance_year=attendance_year,
        attendance_counts=attendance_counts,
        course_attendance=course_attendance,
        weekly_attendance=weekly_attendance,
    )


@app.route("/student-login", methods=["GET", "POST"])
def student_login():
    error = None
    student = None
    if request.method == "POST":
        register_number = request.form.get("register_number", "").strip()
        password = request.form.get("password", "").strip()
        conn = get_db()
        row = conn.execute(
            "SELECT * FROM students WHERE register_number = ?", (register_number,)
        ).fetchone()
        conn.close()
        student = dict(row) if row else None
        if student and student.get("student_phone") == password:
            session["student_id"] = student["register_number"]
            return redirect(url_for("student_dashboard"))
        error = "Invalid register number or password."
    return render_template("student_login.html", error=error, logged_in=False, student=None)


@app.route("/parent-login", methods=["GET", "POST"])
def parent_login():
    error = None
    if request.method == "POST":
        register_number = request.form.get("register_number", "").strip()
        parent_phone = request.form.get("parent_phone", "").strip()
        parent_email = request.form.get("parent_email", "").strip()
        conn = get_db()
        row = conn.execute("SELECT * FROM students WHERE register_number = ?", (register_number,)).fetchone()
        conn.close()
        student = dict(row) if row else None
        if student and student.get("parent_phone") == parent_phone:
            session["parent_student_id"] = student["register_number"]
            return redirect(url_for("parent_dashboard"))
        error = "Invalid register number or registered parent phone number."
    return render_template("student_login.html", error=error, logged_in=False, student=None, parent_login=True)


@app.get("/parent-dashboard")
def parent_dashboard():
    student = get_current_parent_student()
    if not student:
        return redirect(url_for("parent_login"))
    # Reuse the existing read-only student view: its session identifier is always
    # the child verified by the parent phone number above.
    session["student_id"] = student["register_number"]
    return redirect(url_for("student_dashboard", parent_view="1"))


@app.get("/parent-logout")
def parent_logout():
    session.pop("parent_student_id", None)
    session.pop("student_id", None)
    return redirect(url_for("parent_login"))


@app.route("/student-dashboard")
def student_dashboard():
    student = get_current_student()
    if not student:
        return redirect(url_for("student_login"))

    conn = get_db()
    attendance_rows = conn.execute(
        "SELECT * FROM attendance WHERE register_number = ?", (student["register_number"],)
    ).fetchall()
    test_rows = conn.execute(
        "SELECT * FROM internal_marks WHERE register_number = ?", (student["register_number"],)
    ).fetchall()
    conn.close()

    attendance_records = [dict(r) for r in attendance_rows]
    test_records = [dict(r) for r in test_rows]

    attendance_summary = {
        "Present": sum(1 for r in attendance_records if r.get("status") == "Present"),
        "Absent": sum(1 for r in attendance_records if r.get("status") == "Absent"),
        "Event": sum(1 for r in attendance_records if r.get("status") == "Event"),
    }
    total_attendance = len(attendance_records)
    attendance_rate = round(attendance_summary["Present"] / total_attendance * 100) if total_attendance else 0
    attendance_history = [
        {"label": record.get("date", ""), "status": record.get("status", "Present")}
        for record in sorted(attendance_records, key=lambda r: r.get("date", ""))[-6:]
    ]

    marks_by_subject = []
    if test_records:
        subject_groups = {}
        for record in test_records:
            subject = record.get("subject", "Unknown")
            marks = record.get("marks_obtained", 0)
            max_marks = record.get("max_marks", 1) or 1
            subject_groups.setdefault(subject, []).append(marks / max_marks * 100)
        for subject, values in subject_groups.items():
            marks_by_subject.append(
                {
                    "subject": subject,
                    "average_percent": round(sum(values) / len(values)),
                    "tests": len(values),
                }
            )
        marks_by_subject.sort(key=lambda item: item["average_percent"], reverse=True)

    total_tests = len(test_records)
    average_score = 0
    if total_tests:
        total_percent = sum(
            (r.get("marks_obtained", 0) / r.get("max_marks", 1) * 100) if r.get("max_marks") else 0
            for r in test_records
        )
        average_score = round(total_percent / total_tests)

    base_sgpa = (average_score / 100) * 9 + 1
    semester_sgpa = [round(base_sgpa - (6 - sem) * 0.2, 2) for sem in range(1, 7)]

    return render_template(
        "student_dashboard.html",
        parent_view=request.args.get("parent_view") == "1" and bool(session.get("parent_student_id")),
        student=student,
        attendance_records=attendance_records,
        attendance_summary=attendance_summary,
        test_records=test_records,
        today=date.today().strftime("%d %B %Y"),
        attendance_rate=attendance_rate,
        attendance_history=attendance_history,
        marks_by_subject=marks_by_subject,
        total_tests=total_tests,
        average_score=average_score,
        semester_sgpa=semester_sgpa,
    )


@app.route("/student-logout")
def student_logout():
    session.pop("student_id", None)
    return redirect(url_for("student_login"))


if __name__ == "__main__":
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print("=" * 60)
    print("  Student Record & Parent Alert System")
    print("=" * 60)
    print(f"  Local:   http://127.0.0.1:5000")
    print(f"  Network: http://{local_ip}:5000")
    print("=" * 60)
    # use_reloader=False prevents the server restarting mid-request when
    # files (e.g. the database or attendance Excel) change during an upload,
    # which previously caused the connection to drop (ERR_FAILED).
    app.run(host="0.0.0.0", port=5000, debug=True, use_reloader=False)
