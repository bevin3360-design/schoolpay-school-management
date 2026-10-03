# SchoolPay school management app

A Flask application for school records and financial workflows, including learner
profiles, grade-based fee charges, payment recording and review, receipts, and
learner fee statements.

## Run locally

Use Python 3.10 or newer:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

Open <http://127.0.0.1:5000>. The current development seed account is
`teacher` / `password`. Change the authentication and secret-key configuration
before any deployment. Do not use real learner or financial data with this
development application.

## Run tests

```powershell
python -m pytest tests -q
```

## Scope and data

Payment entry is currently manual; no M-Pesa or bank integration is configured.
Offline synchronization and parent accounts are not implemented. The SQLite
database and school-specific documents, spreadsheets, and uploaded files are
local-only and excluded from Git.
