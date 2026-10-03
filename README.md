# SchoolPay school management app

A Flask application for CBC/CBE learning records and financial workflows. It
includes stage-aware learner profiles, term-specific fee charges, payment review
and receipts, staff-role dashboards, teaching assignments, CBE marks and points,
and printable/downloadable learner results.

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

## School workspaces

School staff roles are assigned by an administrator: finance officer, registrar,
exam officer, or teacher. The platform super-administrator remains a separate
developer account. Each role is restricted to its routes and dashboard. Teachers
can enter marks only for their assigned learning areas and grades.

Learner records select a school stage and a grade, and record the primary contact
as father, mother, or guardian. Fee payments require a learner, a charged
academic year and term, and a ten-digit Kenyan phone number. Payments with
unmatched payer details remain in the verification queue.

Learning-area names follow the [KICD CBC/CBE curriculum-design resources](https://kicd.ac.ke/cbc-materials/curriculum-designs/).
The CBE scale used for entered marks is:

| Mark | Grade | Points |
|---|---|---:|
| 90–100 | EE1 | 8 |
| 75–89 | EE2 | 7 |
| 58–74 | ME1 | 6 |
| 41–57 | ME2 | 5 |
| 31–40 | AE1 | 4 |
| 21–30 | AE2 | 3 |
| 11–20 | BE1 | 2 |
| 0–10 | BE2 | 1 |

Exam officers can download a school-branded Excel marks template, enter scores,
and upload the workbook in Marks Entry. The template can also be uploaded into
Google Sheets and exported back as `.xlsx`; direct Google Drive/Sheets API
integration is not configured.

## Scope and data

Payment entry is currently manual; no M-Pesa or bank integration is configured.
Offline synchronization, parent logins, multiple guardian accounts, and online
Google Sheets integration are not implemented. The SQLite database and
school-specific documents, spreadsheets, and uploaded files are local-only and
excluded from Git.
