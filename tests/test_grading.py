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


def test_school_switch_updates_session():
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
    assert get_subject_grade_rubric("Grade 7", 62)["grade"] == "EE1"
    assert get_subject_grade_rubric("Grade 7", 57)["grade"] == "EE2"
    assert get_subject_grade_rubric("Grade 7", 50)["grade"] == "ME1"
    assert get_subject_grade_rubric("Grade 7", 44)["grade"] == "ME2"
    assert get_subject_grade_rubric("Grade 7", 38)["grade"] == "AE1"
    assert get_subject_grade_rubric("Grade 7", 32)["grade"] == "AE2"
    assert get_subject_grade_rubric("Grade 7", 10)["grade"] == "BE"


def test_school_admin_access_is_scoped_to_their_school():
    assert teacher_can_manage_school(1, 1) is True
    assert teacher_can_manage_school(2, 1) is False
