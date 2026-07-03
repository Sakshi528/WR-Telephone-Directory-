# Telephone Directory Management System

A Flask and MySQL web app for managing employee contact details, common directory numbers, user registrations, and employee update requests.

## Features

- User registration and login
- Admin login and dashboard
- Employee directory with keyword, organization, and department filters
- Employee detail pages and user update requests
- Admin approval/rejection of update requests
- Admin CRUD for employees and directory numbers
- Import helper for employee data from a Word document

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

3. Create the MySQL database and tables.

```powershell
mysql -u root -p < database/schema.sql
```

4. Configure the database connection if your MySQL password or host differs.

```powershell
$env:DATABASE_URL = "mysql+pymysql://root:your_password@localhost/telephone_directory"
$env:SECRET_KEY = "change-this-secret"
```

If `DATABASE_URL` is not set, the app uses the local default from `config.py`.

5. Create the default admin account.

```powershell
python create_admin.py
```

Default admin credentials:

- Email: `admin@gmail.com`
- Password: `admin123`

6. Run the app.

```powershell
python app.py
```

Open `http://127.0.0.1:5000/login`.

## Optional Import

To import employees from the Word document in `uploads/`, run:

```powershell
python import_employee.py
```

To rebuild employee data with corrected column mapping:

```powershell
python import_employee.py --replace
```

To import hospitals, emergency numbers, and control-room numbers:

```powershell
python import_directory_numbers.py --replace
```
