from datetime import datetime
import os
import re
import sqlite3
from functools import wraps
from typing import Any, Dict, List
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

from flask import Flask, flash, redirect, render_template, request, session, url_for

app = Flask(__name__)
app.secret_key = "school-report-secret"
app.config["UPLOAD_FOLDER"] = os.path.join(app.root_path, "static", "uploads")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
DB_PATH = os.path.join(app.root_path, "school.db")
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS schools (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                motto TEXT,
                logo_url TEXT,
                signature_url TEXT,
                report_template TEXT DEFAULT 'default'
            );

            CREATE TABLE IF NOT EXISTS teachers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                full_name TEXT NOT NULL,
                school_id INTEGER DEFAULT 1,
                role TEXT DEFAULT 'admin',
                is_super_admin INTEGER DEFAULT 0,
                FOREIGN KEY(school_id) REFERENCES schools(id)
            );

            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                admission_number TEXT NOT NULL,
                payment_code TEXT,
                full_name TEXT NOT NULL,
                gender TEXT,
                grade TEXT NOT NULL,
                class_name TEXT,
                dob TEXT,
                parent_name TEXT,
                parent_contact TEXT,
                school_id INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(school_id) REFERENCES schools(id)
            );

            CREATE TABLE IF NOT EXISTS subjects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                short_name TEXT,
                grade_level TEXT,
                school_id INTEGER DEFAULT 1,
                FOREIGN KEY(school_id) REFERENCES schools(id)
            );

            CREATE TABLE IF NOT EXISTS exams (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                exam_type TEXT NOT NULL,
                term TEXT,
                grade_level TEXT,
                school_id INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(school_id) REFERENCES schools(id)
            );

            CREATE TABLE IF NOT EXISTS marks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                exam_id INTEGER NOT NULL,
                subject_id INTEGER NOT NULL,
                marks_obtained REAL NOT NULL,
                out_of REAL DEFAULT 100,
                school_id INTEGER DEFAULT 1,
                FOREIGN KEY(student_id) REFERENCES students(id),
                FOREIGN KEY(exam_id) REFERENCES exams(id),
                FOREIGN KEY(subject_id) REFERENCES subjects(id),
                FOREIGN KEY(school_id) REFERENCES schools(id)
            );

            CREATE TABLE IF NOT EXISTS fee_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                school_id INTEGER NOT NULL,
                grade TEXT NOT NULL,
                academic_year TEXT NOT NULL,
                term TEXT NOT NULL,
                name TEXT NOT NULL,
                amount INTEGER NOT NULL CHECK(amount > 0),
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(school_id, grade, academic_year, term, name),
                FOREIGN KEY(school_id) REFERENCES schools(id)
            );

            CREATE TABLE IF NOT EXISTS fee_charges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                school_id INTEGER NOT NULL,
                student_id INTEGER NOT NULL,
                fee_item_id INTEGER NOT NULL,
                item_name TEXT NOT NULL,
                academic_year TEXT NOT NULL,
                term TEXT NOT NULL,
                amount INTEGER NOT NULL CHECK(amount > 0),
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(student_id, fee_item_id),
                FOREIGN KEY(school_id) REFERENCES schools(id),
                FOREIGN KEY(student_id) REFERENCES students(id),
                FOREIGN KEY(fee_item_id) REFERENCES fee_items(id)
            );

            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                school_id INTEGER NOT NULL,
                amount INTEGER NOT NULL CHECK(amount > 0),
                channel TEXT NOT NULL,
                transaction_code TEXT,
                payer_phone TEXT,
                reference TEXT NOT NULL,
                student_id INTEGER,
                suggested_student_id INTEGER,
                status TEXT NOT NULL CHECK(status IN ('matched', 'pending')),
                received_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                verified_by INTEGER,
                verified_at TEXT,
                verification_note TEXT,
                FOREIGN KEY(school_id) REFERENCES schools(id),
                FOREIGN KEY(student_id) REFERENCES students(id),
                FOREIGN KEY(suggested_student_id) REFERENCES students(id),
                FOREIGN KEY(verified_by) REFERENCES teachers(id)
            );

            CREATE UNIQUE INDEX IF NOT EXISTS idx_payments_transaction_code
            ON payments(school_id, transaction_code)
            WHERE transaction_code IS NOT NULL AND transaction_code != '';

            CREATE TABLE IF NOT EXISTS receipts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                school_id INTEGER NOT NULL,
                payment_id INTEGER NOT NULL UNIQUE,
                receipt_number TEXT NOT NULL UNIQUE,
                previous_balance INTEGER NOT NULL DEFAULT 0,
                new_balance INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(school_id) REFERENCES schools(id),
                FOREIGN KEY(payment_id) REFERENCES payments(id)
            );

            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                school_id INTEGER NOT NULL,
                actor_id INTEGER,
                action TEXT NOT NULL,
                entity_type TEXT NOT NULL,
                entity_id INTEGER NOT NULL,
                details TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(school_id) REFERENCES schools(id),
                FOREIGN KEY(actor_id) REFERENCES teachers(id)
            );
            """
        )

        def table_has_column(table_name: str, column_name: str) -> bool:
            columns = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
            return any(column[1] == column_name for column in columns)

        if not table_has_column("schools", "signature_url"):
            conn.execute("ALTER TABLE schools ADD COLUMN signature_url TEXT")
        if not table_has_column("schools", "report_template"):
            conn.execute("ALTER TABLE schools ADD COLUMN report_template TEXT DEFAULT 'default'")
        if not table_has_column("teachers", "school_id"):
            conn.execute("ALTER TABLE teachers ADD COLUMN school_id INTEGER DEFAULT 1")
        if not table_has_column("teachers", "role"):
            conn.execute("ALTER TABLE teachers ADD COLUMN role TEXT DEFAULT 'admin'")
        if not table_has_column("teachers", "is_super_admin"):
            conn.execute("ALTER TABLE teachers ADD COLUMN is_super_admin INTEGER DEFAULT 0")
        if not table_has_column("students", "school_id"):
            conn.execute("ALTER TABLE students ADD COLUMN school_id INTEGER DEFAULT 1")
        if not table_has_column("students", "payment_code"):
            conn.execute("ALTER TABLE students ADD COLUMN payment_code TEXT")
        if not table_has_column("receipts", "previous_balance"):
            conn.execute("ALTER TABLE receipts ADD COLUMN previous_balance INTEGER NOT NULL DEFAULT 0")
        if not table_has_column("receipts", "new_balance"):
            conn.execute("ALTER TABLE receipts ADD COLUMN new_balance INTEGER NOT NULL DEFAULT 0")
        if not table_has_column("subjects", "school_id"):
            conn.execute("ALTER TABLE subjects ADD COLUMN school_id INTEGER DEFAULT 1")
        if not table_has_column("exams", "school_id"):
            conn.execute("ALTER TABLE exams ADD COLUMN school_id INTEGER DEFAULT 1")
        if not table_has_column("marks", "school_id"):
            conn.execute("ALTER TABLE marks ADD COLUMN school_id INTEGER DEFAULT 1")

        conn.execute("UPDATE teachers SET school_id = 1 WHERE school_id IS NULL")
        conn.execute("UPDATE students SET school_id = 1 WHERE school_id IS NULL")
        conn.execute("UPDATE subjects SET school_id = 1 WHERE school_id IS NULL")
        conn.execute("UPDATE exams SET school_id = 1 WHERE school_id IS NULL")
        conn.execute("UPDATE marks SET school_id = 1 WHERE school_id IS NULL")

        conn.execute(
            "INSERT OR IGNORE INTO schools (id, name, motto, logo_url, signature_url, report_template) VALUES (1, ?, ?, ?, ?, ?)",
            (
                "Sharleen School",
                "Nurturing Excellence, Character and Purpose",
                "/static/sample-report-form.png",
                "",
                "default",
            ),
        )
        conn.execute(
            "INSERT OR IGNORE INTO schools (id, name, motto, logo_url, signature_url, report_template) VALUES (2, ?, ?, ?, ?, ?)",
            (
                "Bright Stars Academy",
                "Learning with Purpose",
                "/static/sample-report-form.png",
                "",
                "default",
            ),
        )
        conn.execute(
            "INSERT OR IGNORE INTO teachers (id, username, password, full_name, school_id, role, is_super_admin) VALUES (1, ?, ?, ?, ?, ?, ?)",
            ("teacher", "password", "Teacher Admin", 1, "super_admin", 1),
        )
        conn.execute(
            "INSERT OR IGNORE INTO teachers (id, username, password, full_name, school_id, role, is_super_admin) VALUES (2, ?, ?, ?, ?, ?, ?)",
            ("teacher2", "password", "Teacher Admin 2", 2, "admin", 0),
        )

        default_subjects = [
            ("English", "ENG", "All"),
            ("Kiswahili", "KIS", "All"),
            ("Mathematics", "MATH", "All"),
            ("Science", "SCI", "All"),
            ("Social Studies", "SST", "All"),
            ("Creative Arts", "CA", "All"),
            ("Religious Education", "RE", "All"),
            ("Physical Education", "PE", "All"),
        ]
        for subject in default_subjects:
            conn.execute(
                "INSERT OR IGNORE INTO subjects (name, short_name, grade_level, school_id) VALUES (?, ?, ?, ?)",
                (*subject, 1),
            )
            conn.execute(
                "INSERT OR IGNORE INTO subjects (name, short_name, grade_level, school_id) VALUES (?, ?, ?, ?)",
                (*subject, 2),
            )

        default_exams = [
            ("Opening Assessment", "Opener", "Term 1", "All"),
            ("Midterm Assessment", "Midterm", "Term 1", "All"),
            ("End Term Assessment", "End Term", "Term 1", "All"),
        ]
        for exam in default_exams:
            conn.execute(
                "INSERT OR IGNORE INTO exams (name, exam_type, term, grade_level, school_id) VALUES (?, ?, ?, ?, ?)",
                (*exam, 1),
            )
            conn.execute(
                "INSERT OR IGNORE INTO exams (name, exam_type, term, grade_level, school_id) VALUES (?, ?, ?, ?, ?)",
                (*exam, 2),
            )

        default_students = [
            (
                "ADM001",
                "Amina Muli",
                "Female",
                "Grade 7",
                "JSS 1",
                "2014-05-10",
                "Muli Mwangi",
                "0721000111",
            ),
            (
                "ADM002",
                "Daniel Otieno",
                "Male",
                "Grade 8",
                "JSS 2",
                "2013-03-22",
                "Grace Otieno",
                "0721000222",
            ),
            (
                "ADM003",
                "Faith Wanjiku",
                "Female",
                "Grade 9",
                "JSS 3",
                "2012-08-04",
                "Joseph Wanjiku",
                "0721000333",
            ),
            (
                "ADM004",
                "Brian Kariuki",
                "Male",
                "Grade 6",
                "Primary 6",
                "2015-01-18",
                "Jane Kariuki",
                "0721000444",
            ),
            (
                "ADM005",
                "Lilian Njeri",
                "Female",
                "Grade 5",
                "Primary 5",
                "2016-10-02",
                "Charles Njeri",
                "0721000555",
            ),
        ]
        for student in default_students:
            conn.execute(
                "INSERT OR IGNORE INTO students (admission_number, full_name, gender, grade, class_name, dob, parent_name, parent_contact, school_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (*student, 1),
            )

        students_without_payment_codes = conn.execute(
            "SELECT id, school_id FROM students WHERE payment_code IS NULL OR payment_code = ''"
        ).fetchall()
        for student in students_without_payment_codes:
            conn.execute(
                "UPDATE students SET payment_code = ? WHERE id = ?",
                (f"S{student['school_id']}-{student['id']:06d}", student["id"]),
            )
        conn.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_students_school_payment_code "
            "ON students(school_id, payment_code)"
        )

        if conn.execute("SELECT COUNT(*) AS count FROM marks").fetchone()["count"] == 0:
            sample_students = conn.execute("SELECT id FROM students WHERE school_id = 1 ORDER BY id").fetchall()
            sample_subjects = conn.execute("SELECT id FROM subjects WHERE school_id = 1 ORDER BY id").fetchall()
            sample_exam = conn.execute("SELECT id FROM exams WHERE school_id = 1 ORDER BY id LIMIT 1").fetchone()
            if sample_students and sample_subjects and sample_exam:
                for student in sample_students:
                    for subject in sample_subjects[:5]:
                        marks = 72 + (student["id"] % 3) * 5 + (subject["id"] % 2) * 3
                        conn.execute(
                            "INSERT INTO marks (student_id, exam_id, subject_id, marks_obtained, out_of, school_id) VALUES (?, ?, ?, ?, ?, ?)",
                            (student["id"], sample_exam["id"], subject["id"], marks, 100, 1),
                        )


init_db()


def parse_ksh_amount(value: str | None) -> int | None:
    try:
        amount = int(value or "")
    except (TypeError, ValueError):
        return None
    return amount if amount > 0 else None


def normalize_phone(value: str | None) -> str:
    digits = re.sub(r"\D", "", value or "")
    if digits.startswith("254") and len(digits) == 12:
        return "0" + digits[3:]
    return digits


def add_audit_log(
    conn: sqlite3.Connection,
    school_id: int,
    actor_id: int | None,
    action: str,
    entity_type: str,
    entity_id: int,
    details: str,
) -> None:
    conn.execute(
        """
        INSERT INTO audit_logs (school_id, actor_id, action, entity_type, entity_id, details)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (school_id, actor_id, action, entity_type, entity_id, details),
    )


def create_payment_receipt(
    conn: sqlite3.Connection,
    payment_id: int,
    school_id: int,
    actor_id: int | None,
) -> int:
    receipt = conn.execute(
        "SELECT id FROM receipts WHERE payment_id = ?", (payment_id,)
    ).fetchone()
    if receipt:
        return int(receipt["id"])

    payment = conn.execute(
        "SELECT student_id, amount FROM payments WHERE id = ? AND school_id = ? AND status = 'matched'",
        (payment_id, school_id),
    ).fetchone()
    if not payment:
        raise ValueError("Only a matched payment can receive a receipt.")
    charges = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM fee_charges WHERE student_id = ? AND school_id = ?",
        (payment["student_id"], school_id),
    ).fetchone()[0]
    other_payments = conn.execute(
        """
        SELECT COALESCE(SUM(amount), 0) FROM payments
        WHERE student_id = ? AND school_id = ? AND status = 'matched' AND id != ?
        """,
        (payment["student_id"], school_id, payment_id),
    ).fetchone()[0]
    previous_balance = int(charges) - int(other_payments)
    new_balance = previous_balance - int(payment["amount"])
    cursor = conn.execute(
        """
        INSERT INTO receipts
            (school_id, payment_id, receipt_number, previous_balance, new_balance)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            school_id,
            payment_id,
            f"PENDING-{payment_id}",
            previous_balance,
            new_balance,
        ),
    )
    receipt_id = int(cursor.lastrowid)
    receipt_number = f"RC-{datetime.now().year}-{receipt_id:06d}"
    conn.execute(
        "UPDATE receipts SET receipt_number = ? WHERE id = ?",
        (receipt_number, receipt_id),
    )
    add_audit_log(
        conn,
        school_id,
        actor_id,
        "Receipt generated",
        "receipt",
        receipt_id,
        receipt_number,
    )
    return receipt_id


def teacher_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("teacher_logged_in"):
            flash("Please sign in before using the school records system.")
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped


def teacher_can_manage_school(teacher_id: int | None, school_id: int) -> bool:
    if teacher_id is None:
        return False
    with get_db() as conn:
        teacher = conn.execute("SELECT school_id, is_super_admin FROM teachers WHERE id = ?", (teacher_id,)).fetchone()
    if not teacher:
        return False
    return bool(teacher["is_super_admin"]) or int(teacher["school_id"]) == int(school_id)


def get_active_school_id() -> int:
    try:
        selected_school_id = session.get("school_id")
    except RuntimeError:
        selected_school_id = None

    if selected_school_id is None:
        with get_db() as conn:
            teacher_id = None
            try:
                teacher_id = session.get("teacher_id")
            except RuntimeError:
                teacher_id = None
            teacher = conn.execute("SELECT school_id FROM teachers WHERE id = ?", (teacher_id,)).fetchone() if teacher_id else None
            selected_school_id = teacher["school_id"] if teacher else 1
            try:
                session["school_id"] = selected_school_id
            except RuntimeError:
                pass
    return int(selected_school_id)


def get_current_school() -> Dict[str, Any]:
    school_id = get_active_school_id()
    with get_db() as conn:
        row = conn.execute("SELECT * FROM schools WHERE id = ?", (school_id,)).fetchone()
    return dict(row) if row else {"id": school_id, "name": "School", "motto": "", "logo_url": "/static/sample-report-form.png", "signature_url": "", "report_template": "default"}


def get_report_template_name(grade_label: str) -> str:
    cleaned = str(grade_label).strip().lower().replace(" ", "")
    if cleaned in {"grade6"}:
        return "report_card_grade_6.html"
    if cleaned in {"grade8", "grade9"}:
        return "report_card_grade_8.html" if cleaned == "grade8" else "report_card_grade_9.html"
    return "report_card_default.html"


def get_subject_grade_rubric(grade_level: str, percentage: float) -> Dict[str, Any]:
    if normalize_grade_level(grade_level) == "kjsea":
        if percentage >= 61:
            return {"grade": "EE1", "descriptor": "Excellent", "band": "7", "band_name": "Exceeding Expectation 1"}
        if percentage >= 55:
            return {"grade": "EE2", "descriptor": "Very Good", "band": "6", "band_name": "Exceeding Expectation 2"}
        if percentage >= 49:
            return {"grade": "ME1", "descriptor": "Good", "band": "5", "band_name": "Meeting Expectation 1"}
        if percentage >= 43:
            return {"grade": "ME2", "descriptor": "Satisfactory", "band": "4", "band_name": "Meeting Expectation 2"}
        if percentage >= 37:
            return {"grade": "AE1", "descriptor": "Approaching", "band": "3", "band_name": "Approaching Expectation 1"}
        if percentage >= 31:
            return {"grade": "AE2", "descriptor": "Approaching", "band": "2", "band_name": "Approaching Expectation 2"}
        return {"grade": "BE", "descriptor": "Below Expectation", "band": "1", "band_name": "Below Expectation"}
    return get_grade_info(grade_level, percentage)


def normalize_grade_level(grade_label: str) -> str:
    cleaned = str(grade_label).strip().lower().replace(" ", "")
    if cleaned in {"playgroup", "pp1", "pp2", "grade1", "grade2", "grade3", "grade4", "grade5"}:
        return "primary"
    if cleaned in {"grade6"}:
        return "kpsea"
    if cleaned in {"grade7", "grade8", "grade9"}:
        return "kjsea"
    return "primary"


def get_grade_info(grade_level: str, percentage: float) -> Dict[str, Any]:
    system = normalize_grade_level(grade_level)
    if system == "kjsea":
        if percentage >= 80:
            return {"grade": "A", "descriptor": "Exceeding Expectation", "band": "4", "band_name": "Exceeding Expectation"}
        if percentage >= 60:
            return {"grade": "B", "descriptor": "Meeting Expectation", "band": "3", "band_name": "Meeting Expectation"}
        if percentage >= 40:
            return {"grade": "C", "descriptor": "Approaching Expectation", "band": "2", "band_name": "Approaching Expectation"}
        return {"grade": "D", "descriptor": "Below Expectation", "band": "1", "band_name": "Below Expectation"}

    if system == "kpsea":
        if percentage >= 80:
            return {"grade": "A", "descriptor": "Exceeding Expectations", "band": "4", "band_name": "Exceeding Expectations"}
        if percentage >= 60:
            return {"grade": "B", "descriptor": "Meeting Expectations", "band": "3", "band_name": "Meeting Expectations"}
        if percentage >= 40:
            return {"grade": "C", "descriptor": "Approaching Expectations", "band": "2", "band_name": "Approaching Expectations"}
        return {"grade": "D", "descriptor": "Below Expectations", "band": "1", "band_name": "Below Expectations"}

    if percentage >= 80:
        return {"grade": "A", "descriptor": "Excellent", "band": "", "band_name": ""}
    if percentage >= 70:
        return {"grade": "B", "descriptor": "Very Good", "band": "", "band_name": ""}
    if percentage >= 60:
        return {"grade": "C", "descriptor": "Good", "band": "", "band_name": ""}
    if percentage >= 50:
        return {"grade": "D", "descriptor": "Satisfactory", "band": "", "band_name": ""}
    return {"grade": "E", "descriptor": "Needs Support", "band": "", "band_name": ""}


def generate_personalized_remark(student_name: str, overall_average: float, strength_subjects: List[str], weak_subjects: List[str]) -> str:
    strengths = ", ".join(strength_subjects[:2]) if strength_subjects else "core learning areas"
    growth = ", ".join(weak_subjects[:2]) if weak_subjects else "key concepts"
    if overall_average >= 80:
        return (
            f"{student_name} is a highly capable and dependable learner who consistently demonstrates excellent understanding. "
            f"The learner shows strength in {strengths} and is encouraged to keep building confidence by extending learning in {growth}."
        )
    if overall_average >= 65:
        return (
            f"{student_name} is a focused and improving learner who is making steady progress. "
            f"The learner performs strongly in {strengths} and can strengthen further by giving extra attention to {growth}."
        )
    if overall_average >= 50:
        return (
            f"{student_name} is a determined learner who shows willingness to grow and improve. "
            f"With continued effort in {growth}, the learner can build on strengths in {strengths} and gain greater confidence."
        )
    return (
        f"{student_name} is a resilient learner who is developing confidence and independence. "
        f"With regular practice and support in {growth}, the learner can make strong progress and continue to grow in {strengths}."
    )


def build_student_report(student_id: int) -> Dict[str, Any]:
    school_id = get_active_school_id()
    with get_db() as conn:
        student = conn.execute("SELECT * FROM students WHERE id = ? AND school_id = ?", (student_id, school_id)).fetchone()
        if not student:
            raise ValueError("Student not found")
        marks_rows = conn.execute(
            """
            SELECT m.*, s.name AS subject_name, s.short_name AS short_name
            FROM marks m
            JOIN subjects s ON s.id = m.subject_id
            WHERE m.student_id = ?
            ORDER BY s.name
            """,
            (student_id,),
        ).fetchall()

    subject_scores: Dict[str, List[float]] = {}
    for row in marks_rows:
        subject_scores.setdefault(row["subject_name"], []).append(float(row["marks_obtained"]))

    subject_summary: List[Dict[str, Any]] = []
    for name, values in sorted(subject_scores.items()):
        average = round(sum(values) / len(values), 1)
        subject_summary.append(
            {
                "name": name,
                "average": average,
                "grade": get_subject_grade_rubric(student["grade"], average),
            }
        )

    overall_average = round(sum(item["average"] for item in subject_summary) / len(subject_summary), 1) if subject_summary else 0.0
    overall_grade = get_subject_grade_rubric(student["grade"], overall_average)

    with get_db() as conn:
        class_students = conn.execute(
            "SELECT id, grade, class_name FROM students WHERE grade = ? AND class_name = ? AND school_id = ?",
            (student["grade"], student["class_name"], school_id),
        ).fetchall()

    positions: List[Dict[str, Any]] = []
    for roster_student in class_students:
        roster_report = build_student_report(roster_student["id"]) if roster_student["id"] != student_id else None
        if roster_report is None:
            positions.append({"id": roster_student["id"], "average": overall_average})
        else:
            positions.append({"id": roster_student["id"], "average": roster_report["overall_average"]})

    positions_sorted = sorted(positions, key=lambda item: item["average"], reverse=True)
    current_position = 1
    for index, item in enumerate(positions_sorted):
        if item["id"] == student_id:
            current_position = index + 1
            break

    strength_subjects = [item["name"] for item in subject_summary if item["average"] >= 70]
    weak_subjects = [item["name"] for item in subject_summary if item["average"] < 50]
    if not weak_subjects:
        weak_subjects = ["core concepts and revision habits"]

    remark = generate_personalized_remark(student["full_name"], overall_average, strength_subjects, weak_subjects)

    return {
        "student": dict(student),
        "subject_summary": subject_summary,
        "overall_average": overall_average,
        "overall_grade": overall_grade,
        "position": current_position,
        "strength_subjects": strength_subjects,
        "weak_subjects": weak_subjects,
        "remark": remark,
        "performance_band": overall_grade["band_name"] or overall_grade["descriptor"],
        "report_template": get_report_template_name(student["grade"]),
    }


@app.context_processor
def inject_school_settings():
    return {
        "school_settings": get_current_school(),
        "available_schools": get_available_schools(),
    }


def get_available_schools() -> List[Dict[str, Any]]:
    teacher_id = session.get("teacher_id")
    if not teacher_id:
        with get_db() as conn:
            rows = conn.execute("SELECT id, name, motto, logo_url FROM schools ORDER BY name").fetchall()
        return [dict(row) for row in rows]

    with get_db() as conn:
        teacher = conn.execute("SELECT school_id, is_super_admin FROM teachers WHERE id = ?", (teacher_id,)).fetchone()
    if not teacher:
        return []
    if teacher["is_super_admin"]:
        with get_db() as conn:
            rows = conn.execute("SELECT id, name, motto, logo_url FROM schools ORDER BY name").fetchall()
        return [dict(row) for row in rows]
    with get_db() as conn:
        rows = conn.execute("SELECT id, name, motto, logo_url FROM schools WHERE id = ? ORDER BY name", (teacher["school_id"],)).fetchall()
    return [dict(row) for row in rows]


@app.route("/")
@teacher_required
def index():
    school_id = get_active_school_id()
    with get_db() as conn:
        student_count = conn.execute("SELECT COUNT(*) AS count FROM students WHERE school_id = ?", (school_id,)).fetchone()["count"]
        subject_count = conn.execute("SELECT COUNT(*) AS count FROM subjects WHERE school_id = ?", (school_id,)).fetchone()["count"]
        exam_count = conn.execute("SELECT COUNT(*) AS count FROM exams WHERE school_id = ?", (school_id,)).fetchone()["count"]
        mark_count = conn.execute("SELECT COUNT(*) AS count FROM marks WHERE school_id = ?", (school_id,)).fetchone()["count"]
        finance = conn.execute(
            """
            SELECT
                COALESCE((SELECT SUM(amount) FROM payments
                          WHERE school_id = ? AND status = 'matched'
                            AND date(received_at) = date('now', 'localtime')), 0) AS collected_today,
                MAX(
                    COALESCE((SELECT SUM(amount) FROM fee_charges WHERE school_id = ?), 0)
                    - COALESCE((SELECT SUM(amount) FROM payments
                                WHERE school_id = ? AND status = 'matched'), 0),
                    0
                ) AS outstanding,
                (SELECT COUNT(*) FROM payments WHERE school_id = ?) AS payment_count,
                (SELECT COUNT(*) FROM payments WHERE school_id = ? AND status = 'pending') AS unmatched_count,
                (SELECT COUNT(*) FROM receipts WHERE school_id = ?) AS receipt_count
            """,
            (school_id, school_id, school_id, school_id, school_id, school_id),
        ).fetchone()
    return render_template(
        "index.html",
        student_count=student_count,
        subject_count=subject_count,
        exam_count=exam_count,
        mark_count=mark_count,
        finance=finance,
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        with get_db() as conn:
            teacher = conn.execute(
                "SELECT * FROM teachers WHERE username = ? AND password = ?",
                (request.form.get("username"), request.form.get("password")),
            ).fetchone()
        if teacher:
            session["teacher_logged_in"] = True
            session["teacher_id"] = teacher["id"]
            session["teacher_name"] = teacher["full_name"]
            session["school_id"] = teacher["school_id"]
            flash("Welcome back.")
            return redirect(url_for("index"))
        flash("The teacher credentials were not accepted.")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have signed out.")
    return redirect(url_for("login"))


@app.route("/switch-school/<int:school_id>")
@teacher_required
def switch_school(school_id: int):
    teacher_id = session.get("teacher_id")
    if not teacher_can_manage_school(teacher_id, school_id):
        flash("You are not allowed to switch to that school.")
        return redirect(url_for("index"))
    with get_db() as conn:
        school = conn.execute("SELECT id FROM schools WHERE id = ?", (school_id,)).fetchone()
    if school:
        session["school_id"] = school_id
        flash("School switched successfully.")
    else:
        flash("The selected school could not be found.")
    return redirect(url_for("index"))


@app.route("/students", methods=["GET", "POST"])
@teacher_required
def students():
    school_id = get_active_school_id()
    if request.method == "POST":
        with get_db() as conn:
            cursor = conn.execute(
                """
                INSERT INTO students (admission_number, full_name, gender, grade, class_name, dob, parent_name, parent_contact, school_id, payment_code)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '')
                """,
                (
                    request.form.get("admission_number"),
                    request.form.get("full_name"),
                    request.form.get("gender"),
                    request.form.get("grade"),
                    request.form.get("class_name"),
                    request.form.get("dob"),
                    request.form.get("parent_name"),
                    request.form.get("parent_contact"),
                    school_id,
                ),
            )
            student_id = int(cursor.lastrowid)
            payment_code = f"S{school_id}-{student_id:06d}"
            conn.execute(
                "UPDATE students SET payment_code = ? WHERE id = ?",
                (payment_code, student_id),
            )
        flash(f"Student added successfully. Payment code: {payment_code}")
        return redirect(url_for("students"))

    with get_db() as conn:
        student_rows = conn.execute("SELECT * FROM students WHERE school_id = ? ORDER BY grade, class_name, full_name", (school_id,)).fetchall()
    return render_template("students.html", students=student_rows)


@app.route("/students/<int:student_id>/edit", methods=["GET", "POST"])
@teacher_required
def edit_student(student_id: int):
    school_id = get_active_school_id()
    with get_db() as conn:
        student = conn.execute("SELECT * FROM students WHERE id = ? AND school_id = ?", (student_id, school_id)).fetchone()
    if not student:
        flash("Student not found.")
        return redirect(url_for("students"))

    if request.method == "POST":
        with get_db() as conn:
            conn.execute(
                "UPDATE students SET admission_number = ?, full_name = ?, gender = ?, grade = ?, class_name = ?, dob = ?, parent_name = ?, parent_contact = ? WHERE id = ?",
                (
                    request.form.get("admission_number"),
                    request.form.get("full_name"),
                    request.form.get("gender"),
                    request.form.get("grade"),
                    request.form.get("class_name"),
                    request.form.get("dob"),
                    request.form.get("parent_name"),
                    request.form.get("parent_contact"),
                    student_id,
                ),
            )
        flash("Student updated successfully.")
        return redirect(url_for("students"))

    return render_template("edit_student.html", student=student)


@app.route("/fees", methods=["GET", "POST"])
@teacher_required
def fees():
    school_id = get_active_school_id()
    if request.method == "POST":
        grade = (request.form.get("grade") or "").strip()
        academic_year = (request.form.get("academic_year") or "").strip()
        term = (request.form.get("term") or "").strip()
        name = (request.form.get("name") or "").strip()
        amount = parse_ksh_amount(request.form.get("amount"))
        if not all((grade, academic_year, term, name)) or amount is None:
            flash("Enter a grade, year, term, fee name, and a positive whole-KSh amount.")
            return redirect(url_for("fees"))

        with get_db() as conn:
            exists = conn.execute(
                """
                SELECT id FROM fee_items
                WHERE school_id = ? AND grade = ? AND academic_year = ? AND term = ? AND name = ?
                """,
                (school_id, grade, academic_year, term, name),
            ).fetchone()
            if exists:
                flash("That fee item already exists for this grade, year, and term.")
                return redirect(url_for("fees"))

            cursor = conn.execute(
                """
                INSERT INTO fee_items (school_id, grade, academic_year, term, name, amount)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (school_id, grade, academic_year, term, name, amount),
            )
            fee_item_id = int(cursor.lastrowid)
            roster = conn.execute(
                "SELECT id FROM students WHERE school_id = ? AND grade = ?",
                (school_id, grade),
            ).fetchall()
            for student in roster:
                conn.execute(
                    """
                    INSERT OR IGNORE INTO fee_charges
                        (school_id, student_id, fee_item_id, item_name, academic_year, term, amount)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (school_id, student["id"], fee_item_id, name, academic_year, term, amount),
                )
            add_audit_log(
                conn,
                school_id,
                session.get("teacher_id"),
                "Fee item created and charged",
                "fee_item",
                fee_item_id,
                f"{grade} {academic_year} {term}: {name}, KSh {amount}",
            )
        flash(f"Fee item saved and charged to {len(roster)} {grade} learner(s).")
        return redirect(url_for("fees"))

    with get_db() as conn:
        fee_items = conn.execute(
            """
            SELECT f.*, COUNT(c.id) AS learners_charged
            FROM fee_items f LEFT JOIN fee_charges c ON c.fee_item_id = f.id
            WHERE f.school_id = ?
            GROUP BY f.id
            ORDER BY f.academic_year DESC, f.term, f.grade, f.name
            """,
            (school_id,),
        ).fetchall()
        grades = conn.execute(
            "SELECT DISTINCT grade FROM students WHERE school_id = ? ORDER BY grade",
            (school_id,),
        ).fetchall()
    return render_template("fees.html", fee_items=fee_items, grades=grades)


@app.route("/fees/<int:fee_item_id>/apply", methods=["POST"])
@teacher_required
def apply_fee_item(fee_item_id: int):
    school_id = get_active_school_id()
    with get_db() as conn:
        fee_item = conn.execute(
            "SELECT * FROM fee_items WHERE id = ? AND school_id = ?",
            (fee_item_id, school_id),
        ).fetchone()
        if not fee_item:
            flash("Fee item not found.")
            return redirect(url_for("fees"))
        cursor = conn.execute(
            """
            INSERT OR IGNORE INTO fee_charges
                (school_id, student_id, fee_item_id, item_name, academic_year, term, amount)
            SELECT ?, id, ?, ?, ?, ?, ?
            FROM students WHERE school_id = ? AND grade = ?
            """,
            (
                school_id,
                fee_item_id,
                fee_item["name"],
                fee_item["academic_year"],
                fee_item["term"],
                fee_item["amount"],
                school_id,
                fee_item["grade"],
            ),
        )
        add_audit_log(
            conn,
            school_id,
            session.get("teacher_id"),
            "Fee item applied to uncharged learners",
            "fee_item",
            fee_item_id,
            f"Added {cursor.rowcount} learner charge(s).",
        )
    flash(f"Fee item applied to {cursor.rowcount} newly eligible learner(s).")
    return redirect(url_for("fees"))


@app.route("/payments", methods=["GET", "POST"])
@teacher_required
def payments():
    school_id = get_active_school_id()
    if request.method == "POST":
        amount = parse_ksh_amount(request.form.get("amount"))
        channel = (request.form.get("channel") or "").strip()
        reference = (request.form.get("reference") or "").strip()
        transaction_code = (request.form.get("transaction_code") or "").strip()
        payer_phone = (request.form.get("payer_phone") or "").strip()
        if amount is None or channel not in {"M-Pesa", "Bank", "Cash", "Other"} or not reference:
            flash("Enter a positive whole-KSh amount, payment channel, and learner payment code or admission number.")
            return redirect(url_for("payments"))

        with get_db() as conn:
            if transaction_code and conn.execute(
                "SELECT id FROM payments WHERE school_id = ? AND transaction_code = ?",
                (school_id, transaction_code),
            ).fetchone():
                flash("That transaction code has already been recorded.")
                return redirect(url_for("payments"))

            candidate = conn.execute(
                """
                SELECT * FROM students
                WHERE school_id = ? AND payment_code = ?
                """,
                (school_id, reference),
            ).fetchone()
            if not candidate:
                candidates = conn.execute(
                    """
                    SELECT * FROM students
                    WHERE school_id = ? AND admission_number = ?
                    ORDER BY id LIMIT 2
                    """,
                    (school_id, reference),
                ).fetchall()
                if len(candidates) == 1:
                    candidate = candidates[0]

            phone_matches = bool(
                candidate
                and payer_phone
                and candidate["parent_contact"]
                and normalize_phone(payer_phone) == normalize_phone(candidate["parent_contact"])
            )
            auto_match = bool(candidate and (not payer_phone or phone_matches))
            status = "matched" if auto_match else "pending"
            timestamp = datetime.now().isoformat(sep=" ", timespec="seconds")
            cursor = conn.execute(
                """
                INSERT INTO payments
                    (school_id, amount, channel, transaction_code, payer_phone, reference,
                     student_id, suggested_student_id, status, received_at, verified_by,
                     verified_at, verification_note)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    school_id,
                    amount,
                    channel,
                    transaction_code or None,
                    payer_phone or None,
                    reference,
                    candidate["id"] if auto_match else None,
                    candidate["id"] if candidate else None,
                    status,
                    timestamp,
                    session.get("teacher_id") if auto_match else None,
                    timestamp if auto_match else None,
                    "Matched using learner reference and payer details." if auto_match else None,
                ),
            )
            payment_id = int(cursor.lastrowid)
            if auto_match:
                receipt_id = create_payment_receipt(
                    conn, payment_id, school_id, session.get("teacher_id")
                )
                add_audit_log(
                    conn,
                    school_id,
                    session.get("teacher_id"),
                    "Payment matched",
                    "payment",
                    payment_id,
                    f"KSh {amount} assigned to {candidate['full_name']}",
                )
            else:
                receipt_id = None
                add_audit_log(
                    conn,
                    school_id,
                    session.get("teacher_id"),
                    "Payment requires verification",
                    "payment",
                    payment_id,
                    f"KSh {amount}; reference {reference}",
                )
        if receipt_id:
            flash("Payment matched and receipt generated.")
            return redirect(url_for("payment_receipt", receipt_id=receipt_id))
        flash("Payment recorded in the unmatched queue for bursar verification.")
        return redirect(url_for("payments", queue="unmatched"))

    with get_db() as conn:
        payment_rows = conn.execute(
            """
            SELECT p.*, s.full_name AS student_name, s.admission_number,
                   suggested.full_name AS suggested_name,
                   r.id AS receipt_id, r.receipt_number
            FROM payments p
            LEFT JOIN students s ON s.id = p.student_id
            LEFT JOIN students suggested ON suggested.id = p.suggested_student_id
            LEFT JOIN receipts r ON r.payment_id = p.id
            WHERE p.school_id = ?
            ORDER BY p.received_at DESC, p.id DESC
            """,
            (school_id,),
        ).fetchall()
        student_rows = conn.execute(
            "SELECT id, full_name, admission_number FROM students WHERE school_id = ? ORDER BY full_name",
            (school_id,),
        ).fetchall()
    return render_template(
        "payments.html",
        payments=payment_rows,
        students=student_rows,
        unmatched_only=request.args.get("queue") == "unmatched",
    )


@app.route("/payments/<int:payment_id>/verify", methods=["POST"])
@teacher_required
def verify_payment(payment_id: int):
    school_id = get_active_school_id()
    student_id = request.form.get("student_id", type=int)
    note = (request.form.get("verification_note") or "").strip()
    if not student_id or not note:
        flash("Select the learner and record a verification note.")
        return redirect(url_for("payments", queue="unmatched"))

    with get_db() as conn:
        payment = conn.execute(
            "SELECT * FROM payments WHERE id = ? AND school_id = ? AND status = 'pending'",
            (payment_id, school_id),
        ).fetchone()
        student = conn.execute(
            "SELECT id FROM students WHERE id = ? AND school_id = ?",
            (student_id, school_id),
        ).fetchone()
        if not payment or not student:
            flash("The pending payment or selected learner could not be found.")
            return redirect(url_for("payments", queue="unmatched"))
        timestamp = datetime.now().isoformat(sep=" ", timespec="seconds")
        conn.execute(
            """
            UPDATE payments
            SET student_id = ?, status = 'matched', verified_by = ?, verified_at = ?,
                verification_note = ?
            WHERE id = ?
            """,
            (student_id, session.get("teacher_id"), timestamp, note, payment_id),
        )
        receipt_id = create_payment_receipt(
            conn, payment_id, school_id, session.get("teacher_id")
        )
        add_audit_log(
            conn,
            school_id,
            session.get("teacher_id"),
            "Payment verified",
            "payment",
            payment_id,
            note,
        )
    flash("Payment verified and receipt generated.")
    return redirect(url_for("payment_receipt", receipt_id=receipt_id))


@app.route("/receipts/<int:receipt_id>")
@teacher_required
def payment_receipt(receipt_id: int):
    school_id = get_active_school_id()
    with get_db() as conn:
        receipt = conn.execute(
            """
            SELECT r.*, p.amount, p.channel, p.transaction_code, p.payer_phone,
                   p.received_at, p.reference, r.previous_balance, r.new_balance,
                   s.full_name, s.admission_number,
                   s.id AS student_id, s.grade, s.payment_code, school.name AS school_name
            FROM receipts r
            JOIN payments p ON p.id = r.payment_id
            JOIN students s ON s.id = p.student_id
            JOIN schools school ON school.id = r.school_id
            WHERE r.id = ? AND r.school_id = ? AND p.status = 'matched'
            """,
            (receipt_id, school_id),
        ).fetchone()
        if not receipt:
            flash("Receipt not found.")
            return redirect(url_for("payments"))
    return render_template("payment_receipt.html", receipt=receipt)


@app.route("/students/<int:student_id>/statement")
@teacher_required
def fee_statement(student_id: int):
    school_id = get_active_school_id()
    with get_db() as conn:
        student = conn.execute(
            "SELECT * FROM students WHERE id = ? AND school_id = ?",
            (student_id, school_id),
        ).fetchone()
        if not student:
            flash("Student not found.")
            return redirect(url_for("students"))
        charges = conn.execute(
            """
            SELECT id, item_name AS description, amount, created_at, academic_year, term
            FROM fee_charges WHERE student_id = ? AND school_id = ?
            """,
            (student_id, school_id),
        ).fetchall()
        credits = conn.execute(
            """
            SELECT p.id, p.amount, p.received_at AS created_at, p.channel,
                   p.transaction_code, r.id AS receipt_id, r.receipt_number
            FROM payments p LEFT JOIN receipts r ON r.payment_id = p.id
            WHERE p.student_id = ? AND p.school_id = ? AND p.status = 'matched'
            """,
            (student_id, school_id),
        ).fetchall()

    entries: List[Dict[str, Any]] = []
    for charge in charges:
        entries.append(
            {
                "id": charge["id"],
                "created_at": charge["created_at"],
                "description": f"{charge['description']} ({charge['academic_year']} {charge['term']})",
                "debit": int(charge["amount"]),
                "credit": 0,
                "receipt_id": None,
            }
        )
    for credit in credits:
        entries.append(
            {
                "id": credit["id"],
                "created_at": credit["created_at"],
                "description": f"{credit['channel']} payment"
                + (f" · {credit['transaction_code']}" if credit["transaction_code"] else ""),
                "debit": 0,
                "credit": int(credit["amount"]),
                "receipt_id": credit["receipt_id"],
            }
        )
    entries.sort(key=lambda entry: (entry["created_at"], entry["id"]))
    balance = 0
    for entry in entries:
        balance += entry["debit"] - entry["credit"]
        entry["balance"] = balance
    total_charged = sum(entry["debit"] for entry in entries)
    total_paid = sum(entry["credit"] for entry in entries)
    return render_template(
        "fee_statement.html",
        student=student,
        entries=entries,
        total_charged=total_charged,
        total_paid=total_paid,
        balance=balance,
    )


@app.route("/teachers", methods=["GET", "POST"])
@teacher_required
def teachers():
    school_id = get_active_school_id()
    if request.method == "POST":
        with get_db() as conn:
            conn.execute(
                "INSERT INTO teachers (username, password, full_name, school_id) VALUES (?, ?, ?, ?)",
                (
                    request.form.get("username"),
                    request.form.get("password"),
                    request.form.get("full_name"),
                    school_id,
                ),
            )
        flash("Teacher account added successfully.")
        return redirect(url_for("teachers"))

    with get_db() as conn:
        teacher_rows = conn.execute("SELECT * FROM teachers WHERE school_id = ? ORDER BY full_name", (school_id,)).fetchall()
    return render_template("teachers.html", teachers=teacher_rows)


@app.route("/subjects", methods=["GET", "POST"])
@teacher_required
def subjects():
    school_id = get_active_school_id()
    if request.method == "POST":
        with get_db() as conn:
            conn.execute(
                "INSERT INTO subjects (name, short_name, grade_level, school_id) VALUES (?, ?, ?, ?)",
                (request.form.get("name"), request.form.get("short_name"), request.form.get("grade_level"), school_id),
            )
        flash("Subject added successfully.")
        return redirect(url_for("subjects"))

    with get_db() as conn:
        subject_rows = conn.execute("SELECT * FROM subjects WHERE school_id = ? ORDER BY name", (school_id,)).fetchall()
    return render_template("subjects.html", subjects=subject_rows)


@app.route("/exams", methods=["GET", "POST"])
@teacher_required
def exams():
    school_id = get_active_school_id()
    if request.method == "POST":
        with get_db() as conn:
            conn.execute(
                "INSERT INTO exams (name, exam_type, term, grade_level, school_id) VALUES (?, ?, ?, ?, ?)",
                (
                    request.form.get("name"),
                    request.form.get("exam_type"),
                    request.form.get("term"),
                    request.form.get("grade_level"),
                    school_id,
                ),
            )
        flash("Exam created successfully.")
        return redirect(url_for("exams"))

    with get_db() as conn:
        exam_rows = conn.execute("SELECT * FROM exams WHERE school_id = ? ORDER BY created_at DESC", (school_id,)).fetchall()
    return render_template("exams.html", exams=exam_rows)


@app.route("/marks", methods=["GET", "POST"])
@teacher_required
def marks():
    school_id = get_active_school_id()
    with get_db() as conn:
        students = conn.execute("SELECT * FROM students WHERE school_id = ? ORDER BY grade, class_name, full_name", (school_id,)).fetchall()
        subjects = conn.execute("SELECT * FROM subjects WHERE school_id = ? ORDER BY name", (school_id,)).fetchall()
        exams = conn.execute("SELECT * FROM exams WHERE school_id = ? ORDER BY created_at DESC", (school_id,)).fetchall()

    if request.method == "POST":
        if "marks_file" in request.files and request.files["marks_file"].filename:
            file = request.files["marks_file"]
            path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(file.filename))
            file.save(path)
            wb = load_workbook(path, data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            if rows:
                header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
                for row in rows[1:]:
                    if not any(cell not in (None, "") for cell in row):
                        continue
                    record = dict(zip(header, row))
                    admission = record.get("admission") or record.get("admission_number") or record.get("Admission Number")
                    subject_name = record.get("subject") or record.get("subject_name") or record.get("Subject")
                    marks_value = record.get("marks") or record.get("score") or record.get("Marks")
                    if not admission or not subject_name or marks_value in (None, ""):
                        continue
                    student = next((item for item in students if str(item["admission_number"]).lower() == str(admission).strip().lower()), None)
                    subject = next((item for item in subjects if str(item["name"]).lower() == str(subject_name).strip().lower()), None)
                    if student and subject:
                        with get_db() as conn:
                            conn.execute(
                                "INSERT INTO marks (student_id, exam_id, subject_id, marks_obtained, out_of, school_id) VALUES (?, ?, ?, ?, ?, ?)",
                                (student["id"], request.form.get("exam_id"), subject["id"], float(marks_value), 100, school_id),
                            )
            flash("Marks imported from Excel successfully.")
            return redirect(url_for("marks"))

        exam_id = request.form.get("exam_id")
        with get_db() as conn:
            for student in students:
                for subject in subjects:
                    raw_value = request.form.get(f"mark_{student['id']}_{subject['id']}")
                    if raw_value is None or raw_value == "":
                        continue
                    conn.execute(
                        "INSERT INTO marks (student_id, exam_id, subject_id, marks_obtained, out_of, school_id) VALUES (?, ?, ?, ?, ?, ?)",
                        (student["id"], exam_id, subject["id"], float(raw_value), 100, school_id),
                    )
        flash("Marks were saved.")
        return redirect(url_for("marks"))

    return render_template("marks.html", students=students, subjects=subjects, exams=exams)


@app.route("/reports")
@teacher_required
def reports():
    school_id = get_active_school_id()
    with get_db() as conn:
        student_rows = conn.execute("SELECT * FROM students WHERE school_id = ? ORDER BY grade, class_name, full_name", (school_id,)).fetchall()
    student_reports = [build_student_report(student["id"]) for student in student_rows]
    return render_template("reports.html", student_reports=student_reports)


@app.route("/report/<int:student_id>")
@teacher_required
def report_card(student_id: int):
    report = build_student_report(student_id)
    return render_template(report["report_template"], report=report)


@app.route("/print-report/<int:student_id>")
@teacher_required
def print_report(student_id: int):
    report = build_student_report(student_id)
    return render_template(report["report_template"], report=report, print_mode=True)


@app.route("/class-analysis")
@teacher_required
def class_analysis():
    school_id = get_active_school_id()
    with get_db() as conn:
        student_rows = conn.execute("SELECT * FROM students WHERE school_id = ? ORDER BY grade, class_name, full_name", (school_id,)).fetchall()
        subjects = conn.execute("SELECT * FROM subjects WHERE school_id = ? ORDER BY name", (school_id,)).fetchall()

    analyses: List[Dict[str, Any]] = []
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for student in student_rows:
        group = f"{student['grade']} | {student['class_name']}"
        grouped.setdefault(group, []).append(student)

    for group_name, students in sorted(grouped.items()):
        students_reports = [build_student_report(student["id"]) for student in students]
        averages = []
        for subject in subjects:
            values = []
            for report in students_reports:
                for item in report["subject_summary"]:
                    if item["name"] == subject["name"]:
                        values.append(item["average"])
            if values:
                averages.append((subject["name"], round(sum(values) / len(values), 1)))
        analyses.append(
            {
                "group_name": group_name,
                "student_count": len(students_reports),
                "average": round(sum(report["overall_average"] for report in students_reports) / len(students_reports), 1) if students_reports else 0,
                "subject_averages": averages,
            }
        )

    return render_template("class_analysis.html", analyses=analyses)


@app.route("/settings", methods=["GET", "POST"])
@teacher_required
def settings():
    school_id = get_active_school_id()
    with get_db() as conn:
        current_settings = conn.execute("SELECT * FROM schools WHERE id = ?", (school_id,)).fetchone()

    if request.method == "POST":
        logo_url = request.form.get("logo_url")
        signature_url = request.form.get("signature_url")
        if "logo_file" in request.files:
            file = request.files["logo_file"]
            if file and file.filename:
                filename = secure_filename(file.filename)
                path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(path)
                logo_url = f"/static/uploads/{filename}"
        if "signature_file" in request.files:
            file = request.files["signature_file"]
            if file and file.filename:
                filename = secure_filename(file.filename)
                path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(path)
                signature_url = f"/static/uploads/{filename}"
        with get_db() as conn:
            conn.execute(
                "UPDATE schools SET name = ?, motto = ?, logo_url = ?, signature_url = ?, report_template = ? WHERE id = ?",
                (
                    request.form.get("name"),
                    request.form.get("motto"),
                    logo_url,
                    signature_url,
                    request.form.get("report_template"),
                    school_id,
                ),
            )
        flash("School branding updated.")
        return redirect(url_for("settings"))

    return render_template("settings.html", current_settings=current_settings)


if __name__ == "__main__":
    app.run(debug=True)
