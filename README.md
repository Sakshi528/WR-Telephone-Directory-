# Telephone Directory Management System

A Flask + PostgreSQL web app for managing the WRLDC employee telephone directory,
common directory numbers (hospitals, emergency, control rooms), administrative
heads, user registrations, and employee update requests.

## Features

- User registration and login
- Admin login and dashboard
- Employee directory with keyword, organization, and department filters
- Employee detail pages and user-submitted update requests
- Admin approval/rejection of update requests, with audit logging
- Admin CRUD for employees, directory numbers, and administrative heads
- Email distribution group management
- Directory versioning/archival and PDF export (reportlab)
- Import/reconciliation tooling for employee data from Word/Excel sources

## Tech Stack

- Flask 3, Flask-Login, Flask-SQLAlchemy
- PostgreSQL (via `psycopg`)
- python-docx / openpyxl for Word and Excel import
- reportlab for PDF generation

## Setup

1. Create and activate a virtual environment.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies.

```powershell
pip install -r requirements.txt
```

3. Create the PostgreSQL database and base schema.

```powershell
psql -U postgres -d telephone_directory -f database/postgres_schema.sql
```

Then apply the incremental migrations in order:

```powershell
Get-ChildItem migrations\*.sql | Sort-Object Name | ForEach-Object {
    psql -U postgres -d telephone_directory -f $_.FullName
}
```

4. Configure the database connection if it differs from the local default.

```powershell
$env:DATABASE_URL = "postgresql+psycopg://postgres:your_password@localhost:5432/telephone_directory"
$env:SECRET_KEY = "change-this-secret"
```

If `DATABASE_URL` is not set, the app uses the default connection string in `config.py`.

5. Create the default admin account.

```powershell
python create_admin.py
```

6. Run the app.

```powershell
python app.py
```

The server listens on `http://127.0.0.1:8000`. Open `http://127.0.0.1:8000/admin-login`
for the admin dashboard. `FLASK_DEBUG=1` enables the Werkzeug debugger for local
development only — leave it unset (default) on any machine reachable by others.

## Project Layout

- `app.py` — application factory/entry point, blueprint registration
- `config.py` — configuration (database URL, upload/storage paths, session lifetime)
- `models/` — SQLAlchemy models (employees, directory numbers, admin heads, audit log, etc.)
- `routes/` — Flask blueprints (`auth_routes`, `user_routes`, `admin_routes`)
- `database/` — base schema SQL
- `migrations/` — incremental, numbered SQL migrations applied after the base schema
- `templates/`, `static/` — Jinja templates and static assets
- `services/`, `utils/` — shared business logic and helpers
- `scripts/` — one-off data import, reconciliation, and data-quality audit scripts
- `reports/` — generated CSV/Excel audit and reconciliation reports
- `docs/PROJECT_DOCUMENTATION.md` — extended project documentation

## Data Import & Reconciliation

Import employees from the Word document in `uploads/`:

```powershell
python import_employee.py
python import_employee.py --replace   # rebuild with corrected column mapping
```

Import hospitals, emergency numbers, and control-room numbers:

```powershell
python import_directory_numbers.py --replace
```

Other import/reconciliation entry points (`import_employees_excel.py`,
`import_from_excel_db.py`, `import_organizations.py`, `reconcile_word_database.py`,
and the rest of `scripts/`) support one-off data-quality audits and backfills against
the Word and Excel source documents; see each script's `--help` and
`docs/PROJECT_DOCUMENTATION.md` for details.

## Tests

```powershell
python test_import.py
python test_import_employee.py
```
