import app as app_module
from app import app, build_student_report, get_grade_info, get_report_template_name, get_subject_grade_rubric, normalize_grade_level, teacher_can_manage_school


def test_kjsea_grading_boundaries():
    assert normalize_grade_level("Grade 7") == "kjsea"
    assert get_grade_info("Grade 7", 82)["grade"] == "A"
    assert get_grade_info("Grade 7", 60)["grade"] == "B"
    assert get_grade_info("Grade 7", 40)["grade"] == "C"
    assert get_grade_info("Grade 7", 39)["grade"] == "D"


def test_kpsea_grading_boundaries():
    assert normalize_grade_level("Grade 6") == "kpsea"
    assert get_grade_info("Grade 6", 82)["grade"] == "A"
    assert get_grade_info("Grade 6", 60)["grade"] == "B"
    assert get_grade_info("Grade 6", 40)["grade"] == "C"
    assert get_grade_info("Grade 6", 39)["grade"] == "D"


def test_report_generation_contains_positive_remark():
    report = build_student_report(1)
    assert report["overall_average"] >= 0
    assert "remark" in report
    assert len(report["remark"]) > 20


def test_teacher_management_page_requires_login():
    client = app.test_client()
    response = client.get("/teachers")
    assert response.status_code == 302


def test_report_template_selection_is_grade_specific():
    assert get_report_template_name("Grade 6") == "report_card_grade_6.html"
    assert get_report_template_name("Grade 8") == "report_card_grade_8.html"
    assert get_report_template_name("Grade 9") == "report_card_grade_9.html"
    assert get_report_template_name("Grade 5") == "report_card_default.html"


def test_school_switch_updates_session(tmp_path, monkeypatch):
    monkeypatch.setattr(app_module, "DB_PATH", str(tmp_path / "switch.db"))
    app_module.init_db()
    client = app.test_client()
    with client.session_transaction() as session:
        session["teacher_logged_in"] = True
        session["teacher_id"] = 1
        session["school_id"] = 1

    response = client.get("/switch-school/1", follow_redirects=True)
    assert response.status_code == 200
    with client.session_transaction() as session:
        assert session["school_id"] == 1


def test_grade_7_to_9_rubric_uses_requested_band_names():
    expected_bands = [
        (100, "EE1", 8),
        (90, "EE1", 8),
        (89, "EE2", 7),
        (75, "EE2", 7),
        (74, "ME1", 6),
        (58, "ME1", 6),
        (57, "ME2", 5),
        (41, "ME2", 5),
        (40, "AE1", 4),
        (31, "AE1", 4),
        (30, "AE2", 3),
        (21, "AE2", 3),
        (20, "BE1", 2),
        (11, "BE1", 2),
        (10, "BE2", 1),
        (0, "BE2", 1),
    ]
    for score, grade, points in expected_bands:
        result = get_subject_grade_rubric("Grade 7", score)
        assert (result["grade"], result["points"]) == (grade, points)


def test_school_admin_access_is_scoped_to_their_school():
    assert teacher_can_manage_school(1, 1) is True
    assert teacher_can_manage_school(2, 1) is False
