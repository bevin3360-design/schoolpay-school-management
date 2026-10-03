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

Open <http://127.0.0.1:5000>. The development-only seed account is
`teacher` / `password`. Do not use it or development-seeded data in production.

## Production requirements

The app fails to start in production unless `APP_ENV=production` and a unique
`SECRET_KEY` are configured. The first production startup also requires
`INITIAL_SCHOOL_NAME`, `INITIAL_ADMIN_USERNAME`, and
`INITIAL_ADMIN_PASSWORD` (at least 12 characters); the initial admin password
is stored as a hash. Production startup does not create demo learners or the
development accounts. It refuses to use a database containing the default
development usernames.

Set `DATABASE_PATH` and `UPLOAD_FOLDER` to locations on persistent storage.
The default SQLite database and uploads are local files and will be lost on
hosts with ephemeral filesystems. Configure HTTPS and persistent backups before
storing real learner or payment records. Real deployments also need an
appropriate production WSGI server and a host plan that supports persistent
storage.

## Free public demo on PythonAnywhere

The checked-in `deploy/pythonanywhere_wsgi.example.py` is a **demo-only**
PythonAnywhere WSGI configuration. It stores its empty demo database and
uploads in the account home directory rather than the application checkout.
The free plan has limited capacity and no production SLA or managed backup.
Keep this deployment on fake data; do not use it for learner, parent, or
payment records. Arrange and test off-site backups before considering any real
data deployment.

1. Create a free account at <https://www.pythonanywhere.com/> and open a Bash
   console.
2. Clone the public demo branch and install the app dependencies:

   ```bash
   git clone --branch deploy/pythonanywhere-demo https://github.com/bevin3360-design/schoolpay-school-management.git
   mkvirtualenv --python=/usr/bin/python3.13 schoolpay-demo
   pip install -r ~/schoolpay-school-management/requirements.txt
   ```

3. In the **Web** tab, add a **Manual configuration** web app using Python 3.13,
   set its virtualenv to `schoolpay-demo`, and open the WSGI configuration file.
4. Copy the contents of `deploy/pythonanywhere_wsgi.example.py` into that WSGI
   file. Replace the username and demo login placeholders. Generate a unique
   secret key in a Bash console with
   `python -c 'import secrets; print(secrets.token_hex(32))'`; use that output
   as `secret_key`. Do not put real credentials in the repository.
5. Save the WSGI file, reload the web app in the Web tab, and open the
   `your-username.pythonanywhere.com` address shown there. Sign in with the
   demo account you configured and change its password immediately.
6. Once the initial deployment succeeds, change `initial_admin_password` to
   an empty string in the WSGI file and reload. The app only needs the initial
   admin password when creating a new, empty database.

To publish later code changes, pull the `deploy/pythonanywhere-demo` branch in
the Bash console and reload the web app. Keep the `.local/share/schoolpay-demo`
directory and its database outside the cloned source directory.

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
