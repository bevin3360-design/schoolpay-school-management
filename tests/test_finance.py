import app as app_module
import pytest


@pytest.fixture
def finance_client(tmp_path, monkeypatch):
    monkeypatch.setattr(app_module, "DB_PATH", str(tmp_path / "school.db"))
    app_module.init_db()
    client = app_module.app.test_client()
    with client.session_transaction() as session:
        session["teacher_logged_in"] = True
        session["teacher_id"] = 1
        session["school_id"] = 1
    return client


def add_term_fee(client):
    client.post(
        "/fees",
        data={
            "grade": "Grade 7",
            "academic_year": "2026",
            "term": "Term 1",
            "name": "Tuition",
            "amount": "5000",
        },
    )


def test_fee_item_charges_grade_and_statement_tracks_payment(finance_client):
    response = finance_client.post(
        "/fees",
        data={
            "grade": "Grade 7",
            "academic_year": "2026",
            "term": "Term 1",
            "name": "Tuition",
            "amount": "5000",
        },
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert b"charged to 1 Grade 7 learner" in response.data

    response = finance_client.post(
        "/payments",
        data={
            "amount": "2000",
            "channel": "M-Pesa",
            "transaction_code": "QWE123",
            "payer_phone": "0721000111",
            "student_id": "1",
            "academic_year": "2026",
            "term": "Term 1",
        },
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert b"RC-" in response.data
    assert b"Amina Muli" in response.data
    assert b"Previous balance:</strong> KSh 5,000" in response.data
    assert b"New balance:</strong> KSh 3,000" in response.data

    response = finance_client.get("/students/1/statement")
    assert response.status_code == 200
    assert b"KSh 5,000" in response.data
    assert b"KSh 2,000" in response.data
    assert b"KSh 3,000" in response.data


def test_phone_mismatch_requires_note_and_bursar_verification(finance_client):
    add_term_fee(finance_client)
    finance_client.post(
        "/payments",
        data={
            "amount": "1500",
            "channel": "M-Pesa",
            "transaction_code": "QWE124",
            "payer_phone": "0799999999",
            "student_id": "1",
            "academic_year": "2026",
            "term": "Term 1",
        },
    )

    response = finance_client.get("/payments?queue=unmatched")
    assert response.status_code == 200
    assert b"Needs verification" in response.data
    assert b"Possible learner: Amina Muli" in response.data

    response = finance_client.post(
        "/payments/1/verify",
        data={"student_id": "1", "verification_note": "Confirmed with guardian"},
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert b"RC-" in response.data
    with app_module.get_db() as conn:
        payment = conn.execute("SELECT status FROM payments WHERE id = 1").fetchone()
        assert payment["status"] == "matched"
        assert conn.execute(
            "SELECT COUNT(*) FROM audit_logs WHERE entity_type = 'payment'"
        ).fetchone()[0] == 2


def test_duplicate_transaction_code_is_not_recorded_twice(finance_client):
    add_term_fee(finance_client)
    payment_data = {
        "amount": "1000",
        "channel": "Bank",
        "transaction_code": "BANK-001",
        "payer_phone": "0721000111",
        "student_id": "1",
        "academic_year": "2026",
        "term": "Term 1",
    }
    finance_client.post("/payments", data=payment_data)
    response = finance_client.post(
        "/payments", data=payment_data, follow_redirects=True
    )

    assert response.status_code == 200
    assert b"already been recorded" in response.data
    with app_module.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM payments").fetchone()[0] == 1


def test_payments_require_ten_digit_phone_and_selected_term(finance_client):
    add_term_fee(finance_client)
    invalid_phone = finance_client.post(
        "/payments",
        data={
            "amount": "1000",
            "channel": "M-Pesa",
            "payer_phone": "721000111",
            "student_id": "1",
            "academic_year": "2026",
            "term": "Term 1",
        },
        follow_redirects=True,
    )
    assert b"10-digit Kenyan phone number" in invalid_phone.data

    invalid_term = finance_client.post(
        "/payments",
        data={
            "amount": "1000",
            "channel": "M-Pesa",
            "payer_phone": "0721000111",
            "student_id": "1",
            "academic_year": "2026",
            "term": "Term 2",
        },
        follow_redirects=True,
    )
    assert b"no fee charges for that year and term" in invalid_term.data
    with app_module.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM payments").fetchone()[0] == 0


def test_fee_items_reject_terms_outside_school_calendar(finance_client):
    response = finance_client.post(
        "/fees",
        data={
            "grade": "Grade 7",
            "academic_year": "2026",
            "term": "Summer",
            "name": "Tuition",
            "amount": "5000",
        },
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert b"valid term" in response.data
    with app_module.get_db() as conn:
        assert conn.execute("SELECT COUNT(*) FROM fee_items").fetchone()[0] == 0


def test_new_students_can_be_charged_existing_fee_items(finance_client):
    finance_client.post(
        "/fees",
        data={
            "grade": "Grade 7",
            "academic_year": "2026",
            "term": "Term 1",
            "name": "Tuition",
            "amount": "5000",
        },
    )
    finance_client.post(
        "/students",
        data={
            "admission_number": "ADM006",
            "full_name": "New Learner",
            "gender": "Female",
            "grade": "Grade 7",
            "parent_relationship": "Mother",
        },
    )
    response = finance_client.post("/fees/1/apply", follow_redirects=True)

    assert response.status_code == 200
    with app_module.get_db() as conn:
        charges = conn.execute(
            "SELECT COUNT(*) FROM fee_charges WHERE fee_item_id = 1"
        ).fetchone()[0]
        student = conn.execute(
            "SELECT payment_code FROM students WHERE admission_number = 'ADM006'"
        ).fetchone()
    assert charges == 2
    assert student["payment_code"] == "S1-000006"
