import base64
from io import BytesIO
from pathlib import Path
import zipfile

import app as app_module
import pytest
from openpyxl import Workbook


@pytest.fixture
def school_app(tmp_path, monkeypatch):
    monkeypatch.setattr(app_module, "DB_PATH", str(tmp_path / "school.db"))
    app_module.init_db()

    def login_as(role):
        username = f"{role}_user"
        with app_module.get_db() as conn:
            conn.execute(
                """
                INSERT INTO teachers (username, password, full_name, school_id, role)
                VALUES (?, ?, ?, 1, ?)
                ON CONFLICT(username) DO UPDATE SET
                    password = excluded.password,
                    full_name = excluded.full_name,
                    school_id = excluded.school_id,
                    role = excluded.role
                """,
                (username, "test-password", f"{role.title()} User", role),
            )
        client = app_module.app.test_client()
        response = client.post(
            "/login",
            data={"username": username, "password": "test-password"},
        )
        assert response.status_code == 302
        return client

    return login_as


def test_role_dashboards_enforce_finance_scope(school_app):
    client = school_app("finance_officer")

    dashboard = client.get("/")
    assert dashboard.status_code == 200
    assert b"Fees collected today" in dashboard.data
    assert b"Student Management" not in dashboard.data
    assert client.get("/fees").status_code == 200
    assert client.get("/teachers").status_code == 403
    assert client.get("/students").status_code == 403


def test_learner_stage_and_parent_relationship_are_saved(school_app):
    client = school_app("registrar")
    response = client.post(
        "/students",
        data={
            "admission_number": "SEN001",
            "full_name": "Senior Learner",
            "gender": "Female",
            "school_stage": "Senior School",
            "grade": "Grade 10",
            "class_name": "10A",
            "parent_name": "Parent Example",
            "parent_relationship": "Mother",
        },
        follow_redirects=True,
    )

    assert response.status_code == 200
    with app_module.get_db() as conn:
        learner = conn.execute(
            "SELECT school_stage, grade, parent_relationship FROM students WHERE admission_number = 'SEN001'"
        ).fetchone()
    assert tuple(learner) == ("Senior School", "Grade 10", "Mother")

    invalid = client.post(
        "/students",
        data={
            "admission_number": "SEN002",
            "full_name": "Invalid Stage",
            "school_stage": "Senior School",
            "grade": "Grade 8",
            "parent_relationship": "Father",
        },
        follow_redirects=True,
    )
    assert b"Choose a grade that belongs to the selected school stage" in invalid.data


def test_teacher_can_only_enter_assigned_learning_area_marks(school_app):
    with app_module.get_db() as conn:
        teacher = conn.execute(
            "SELECT id FROM teachers WHERE username = 'teacher_user'"
        ).fetchone()
        area = conn.execute(
            "SELECT id FROM subjects WHERE school_id = 1 AND name = 'Mathematics'"
        ).fetchone()
        if not teacher:
            conn.execute(
                """
                INSERT INTO teachers (username, password, full_name, school_id, role)
                VALUES ('teacher_user', 'test-password', 'Teacher User', 1, 'teacher')
                """
            )
            teacher = conn.execute(
                "SELECT id FROM teachers WHERE username = 'teacher_user'"
            ).fetchone()
        exam = conn.execute(
            """
            INSERT INTO exams (name, exam_type, term, grade_level, school_id)
            VALUES ('CBE Test', 'End Term', 'Term 1', 'Grade 7', 1)
            """
        )
        exam_id = exam.lastrowid
        conn.execute(
            """
            INSERT INTO teacher_learning_areas (school_id, teacher_id, subject_id, grade)
            VALUES (1, ?, ?, 'Grade 7')
            """,
            (teacher["id"], area["id"]),
        )

    client = school_app("teacher")
    response = client.post(
        "/marks",
        data={
            "exam_id": str(exam_id),
            "mark_1_" + str(area["id"]): "91",
        },
        follow_redirects=True,
    )
    assert response.status_code == 200
    with app_module.get_db() as conn:
        mark = conn.execute(
            """
            SELECT marks_obtained FROM marks
            WHERE student_id = 1 AND exam_id = ? AND subject_id = ?
            """,
            (exam_id, area["id"]),
        ).fetchone()
        conn.execute(
            """
            INSERT INTO students
                (admission_number, payment_code, full_name, grade, school_stage, class_name, school_id)
            VALUES ('RANK001', 'TEST-RANK001', 'Second Learner', 'Grade 7', 'Junior School', 'JSS 1', 1)
            """
        )
        peer_id = conn.execute(
            "SELECT id FROM students WHERE admission_number = 'RANK001'"
        ).fetchone()["id"]
        conn.execute(
            """
            INSERT INTO marks
                (student_id, exam_id, subject_id, marks_obtained, out_of, school_id)
            VALUES (?, ?, ?, 50, 100, 1)
            """,
            (peer_id, exam_id, area["id"]),
        )
    assert mark["marks_obtained"] == 91
    result = client.get(f"/report/1?exam_id={exam_id}&view=points")
    assert result.status_code == 200
    assert b"EE1" in result.data
    assert b"8 points" in result.data
    assert b"Class position:</strong> 1" in result.data
    printed_marks = client.get(f"/print-report/1?exam_id={exam_id}&view=marks")
    assert b"91.0 / 100.0" in printed_marks.data
    downloaded_points = client.get(
        f"/report/1/download?exam_id={exam_id}&view=points"
    )
    assert downloaded_points.status_code == 200
    assert downloaded_points.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert client.get("/report/2").status_code == 404
    assert client.get("/teachers").status_code == 403


def test_exam_template_is_branded_and_importable(school_app, tmp_path, monkeypatch):
    client = school_app("exam_officer")
    project_root = Path(app_module.app.root_path)
    monkeypatch.setattr(app_module.app, "root_path", str(tmp_path))
    monkeypatch.setattr(
        app_module.app, "template_folder", str(project_root / "templates")
    )
    logo_path = tmp_path / "static" / "uploads" / "school.png"
    logo_path.parent.mkdir(parents=True)
    logo_path.write_bytes(
        base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC"
        )
    )
    with app_module.get_db() as conn:
        conn.execute(
            "UPDATE schools SET name = 'Test Academy', logo_url = '/static/uploads/school.png' WHERE id = 1"
        )
        exam = conn.execute(
            """
            INSERT INTO exams (name, exam_type, term, grade_level, school_id)
            VALUES ('Template Exam', 'End Term', 'Term 2', 'Grade 7', 1)
            """
        )
        exam_id = exam.lastrowid

    template_response = client.get(f"/exams/{exam_id}/marks-template")
    assert template_response.status_code == 200
    with zipfile.ZipFile(BytesIO(template_response.data)) as package:
        assert "xl/media/school-logo.png" in package.namelist()
        assert b"Test Academy" in package.read("xl/worksheets/sheet1.xml")

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Test Academy"])
    sheet.append(["Template Exam"])
    sheet.append(["Term 2"])
    sheet.append([])
    sheet.append([])
    sheet.append(["Learner Name", "Admission Number", "Learning Area", "Score (0-100)", "CBE Grade", "Points"])
    sheet.append(["Amina Muli", "ADM001", "Mathematics", 90, "", ""])
    upload = BytesIO()
    workbook.save(upload)
    upload.seek(0)

    import_response = client.post(
        "/marks",
        data={
            "exam_id": str(exam_id),
            "marks_file": (upload, "marks.xlsx"),
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert import_response.status_code == 200
    with app_module.get_db() as conn:
        mark = conn.execute(
            "SELECT marks_obtained FROM marks WHERE exam_id = ? AND student_id = 1",
            (exam_id,),
        ).fetchone()
    assert mark["marks_obtained"] == 90
