"""
Generate a project summary Word document.
Run: python generate_summary.py
Output: Project_Summary_Telephone_Directory.docx
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

NAVY  = RGBColor(0x1A, 0x3A, 0x6B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY  = RGBColor(0xF2, 0xF2, 0xF2)
DARK  = RGBColor(0x22, 0x22, 0x22)


def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def add_heading(doc, text, level=1, color=None):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.color.rgb = color or NAVY
        run.bold = True
    return p


def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix + " ")
        r.bold = True
        r.font.color.rgb = NAVY
    p.add_run(text)
    return p


def add_table_row(table, cells, header=False):
    row = table.add_row()
    row_idx = len(table.rows) - 1
    for i, val in enumerate(cells):
        cell = row.cells[i]
        cell.text = val
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in para.runs:
                run.font.size = Pt(9.5)
                if header:
                    run.bold = True
                    run.font.color.rgb = WHITE
        if header:
            set_cell_bg(cell, "1A3A6B")
        elif row_idx % 2 == 0:
            set_cell_bg(cell, "F2F2F2")
    return row


def build():
    doc = Document()

    # ── Page margins ─────────────────────────────────────────────────────────────
    for section in doc.sections:
        section.page_width  = Inches(8.27)   # A4
        section.page_height = Inches(11.69)
        section.left_margin = section.right_margin   = Inches(1.0)
        section.top_margin  = section.bottom_margin  = Inches(0.9)

    # ── Cover heading ─────────────────────────────────────────────────────────────
    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cover.add_run("Western Region Telephone Directory")
    r.font.size  = Pt(22)
    r.font.bold  = True
    r.font.color.rgb = NAVY

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run("Grid India – WRLDC  |  Project Summary Document")
    r2.font.size  = Pt(12)
    r2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run(f"Prepared: {datetime.date.today().strftime('%d %B %Y')}").font.size = Pt(10)

    doc.add_paragraph()  # spacer

    # ── 1. Project Overview ───────────────────────────────────────────────────────
    add_heading(doc, "1.  Project Overview", level=1)
    doc.add_paragraph(
        "The Western Region Telephone Directory is a web-based internal directory system "
        "developed for Grid India – Western Region Load Despatch Centre (WRLDC). "
        "It provides a centralised, searchable repository of employee contact information, "
        "utility organisation heads, emergency contacts, and important telephone numbers "
        "across all power utilities in the Western Region of India."
    )
    doc.add_paragraph(
        "The application is built with Python (Flask) and PostgreSQL, deployed as a local "
        "intranet web application. All directory pages are fully public (no login required). "
        "A protected Admin Panel allows authorised administrators to manage all data."
    )

    # ── 2. Technology Stack ───────────────────────────────────────────────────────
    add_heading(doc, "2.  Technology Stack", level=1)

    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.style     = "Table Grid"
    add_table_row(tbl, ["Component", "Technology / Version"], header=True)

    stack = [
        ("Web Framework",        "Flask 3.1.1 (Python)"),
        ("Database",             "PostgreSQL (via psycopg2-binary 2.9.10)"),
        ("ORM",                  "Flask-SQLAlchemy 3.1.1"),
        ("Authentication",       "Flask-Login 0.6.3"),
        ("Excel Export",         "openpyxl 3.1.5"),
        ("Word Export",          "python-docx 1.1.2"),
        ("Frontend Framework",   "Bootstrap 5.3 (CDN)"),
        ("Icons",                "Bootstrap Icons (CDN)"),
        ("Colour Scheme",        "Navy Blue (#1A3A6B) on white / Bootstrap neutrals"),
        ("Session Management",   "Flask server-side sessions with 20-minute idle timeout"),
    ]
    for row in stack:
        add_table_row(tbl, list(row))

    doc.add_paragraph()

    # ── 3. Database Structure ─────────────────────────────────────────────────────
    add_heading(doc, "3.  Database Structure", level=1)
    doc.add_paragraph(
        "The PostgreSQL database contains 7 tables. Below is a summary of each table "
        "and its purpose."
    )

    tbl2 = doc.add_table(rows=1, cols=3)
    tbl2.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl2.style     = "Table Grid"
    add_table_row(tbl2, ["Table", "Key Columns", "Purpose"], header=True)

    tables_info = [
        ("organizations",
         "id, organization_name, region, address",
         "Stores every power utility, load despatch centre, and government body."),
        ("departments",
         "id, department_name",
         "Lookup table for internal departments (e.g. Operations, IT, HR)."),
        ("employees",
         "id, employee_name, designation, organization_id (FK), department_id (FK), "
         "office_phone, mobile_phone, residence_phone, email, is_utility_head",
         "Main table of all personnel imported from the Word directory document. "
         "is_utility_head flag marks the head/senior officer of each organisation."),
        ("directory_numbers",
         "id, name, phone_number, email, organization, category",
         "Control room numbers, ALDC lines, and other important numbers not tied to a person."),
        ("emergency_contacts",
         "id, employee_id (FK), contact_name, relation, phone",
         "Personal/family emergency contacts of Grid India employees."),
        ("users",
         "id, email, password (hashed), role, status",
         "Admin accounts only. Passwords stored as Werkzeug pbkdf2 hashes."),
        ("update_requests",
         "id, employee_id / directory_number_id (FK), field_name, new_value, "
         "requested_by, department, contact_number, status, request_date",
         "Anonymous correction requests submitted by any user. Admin approves or rejects."),
    ]
    for row in tables_info:
        add_table_row(tbl2, list(row))

    doc.add_paragraph()

    # ── 4. Public Features ────────────────────────────────────────────────────────
    add_heading(doc, "4.  Public Features  (No Login Required)", level=1)
    doc.add_paragraph(
        "All pages listed below are accessible by any user on the network without "
        "creating an account or logging in."
    )

    # 4.1 Home
    add_heading(doc, "4.1  Home Page  (/)", level=2)
    items = [
        ("Hero banner",     "Organisation title and tagline."),
        ("Live statistics", "Total number of employees and organisations in the database."),
        ("Last updated",    "Month and year the data was last refreshed."),
        ("Universal search","A large search bar with real-time autocomplete that queries "
                            "employees, organisations, designations, and directory numbers."),
        ("Quick-access cards", "Direct links to Employee Directory, Telephone Directory, "
                               "Utility Heads, and Emergency Contacts."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.2 Employee Directory
    add_heading(doc, "4.2  Employee Directory  (/directory)", level=2)
    doc.add_paragraph(
        "Displays all employees belonging to WRLDC / Grid India "
        "(Western Region Load Despatch Centre)."
    )
    items = [
        ("Search & filter",   "Filter by keyword (name, designation, phone), organisation, "
                              "and department simultaneously."),
        ("Autocomplete",      "Live search suggestions while typing."),
        ("Utility Head badge","Green star badge on employees marked as organisation head."),
        ("View details",      "Click any employee to see full contact information."),
        ("Export – Excel",    "Download the current filtered view as a formatted .xlsx file."),
        ("Export – Word",     "Download the current filtered view as a formatted .docx file."),
        ("Admin controls",    "When logged in as admin: Edit, Delete, and toggle Utility Head "
                              "buttons appear inline in the table (no need to go to admin panel)."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.3 Telephone Directory
    add_heading(doc, "4.3  Telephone Directory  (/telephone-directory)", level=2)
    doc.add_paragraph(
        "Lists employees and control room numbers from ALL organisations (not just WRLDC), "
        "grouped by organisation."
    )
    items = [
        ("Organisation cards",  "Each organisation is shown as a card with its name and "
                                "physical address in the header."),
        ("Employee rows",       "Name, designation, office phone, mobile, email, utility head "
                                "indicator per employee."),
        ("Control room rows",   "Control room / ALDC numbers displayed with a red badge, "
                                "highlighted in a distinct row colour."),
        ("Search",              "Search by organisation, employee name, or phone number across "
                                "the entire directory."),
        ("Export – Word (per org)", "Download a single organisation card as a Word document."),
        ("Export – Excel/Word (all)", "Download the entire directory (or search results) as "
                                      "Excel or Word."),
        ("Admin controls",      "When logged in: Edit, Delete, Toggle Utility Head buttons on "
                                "each employee row; Edit/Delete on control room rows."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.4 Utility Heads
    add_heading(doc, "4.4  Utility Heads  (/utility-heads)", level=2)
    doc.add_paragraph(
        "A consolidated single-page view of the head or most senior officer of every "
        "organisation in the directory. Automatically determined by designation-priority "
        "scoring (Chairman/CMD = highest, Manager = lowest) and confirmed by admin."
    )
    items = [
        ("Columns",           "Sr No., Name, Designation, Organisation, Office Phone, "
                              "Mobile, Email, Address."),
        ("Organisation address", "The registered address of each organisation is shown in "
                                 "the Address column."),
        ("Export – Excel",    "Download the complete utility heads list as .xlsx."),
        ("Export – Word",     "Download the complete utility heads list as .docx."),
        ("Admin controls",    "Edit, Remove from Utility Heads, and Delete buttons appear "
                              "for admin users."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.5 Emergency Contacts
    add_heading(doc, "4.5  Emergency Contacts  (/emergency-contacts)", level=2)
    doc.add_paragraph(
        "Lists personal / family emergency contacts for Grid India employees."
    )
    items = [
        ("Columns",        "Sr No., Employee (with designation), Contact Person, "
                           "Relation, Phone Number."),
        ("Employee link",  "Each employee name is clickable and leads to their full profile."),
        ("Admin controls", "Edit and Delete buttons appear inline for admin users."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.6 Employee Detail
    add_heading(doc, "4.6  Employee Detail Page  (/employee/<id>)", level=2)
    items = [
        ("Full contact card",    "Name, designation, organisation, department, office phone, "
                                 "mobile phone, residence phone, email, region, location."),
        ("Utility Head badge",   "Gold star badge if the employee is the organisation head."),
        ("Organisation address", "Address of the employee's organisation displayed below the card."),
        ("Request Correction",   "Any visitor can submit an anonymous update request if they "
                                 "believe the information is incorrect."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.7 Correction Request
    add_heading(doc, "4.7  Correction Request  (/employee/request-update/<id>)", level=2)
    doc.add_paragraph(
        "A fully anonymous form that any visitor can submit to flag incorrect data."
    )
    items = [
        ("Fields",        "Name of submitter, Department, Contact Number, Field to update "
                          "(Mobile / Office Phone / Residence Phone / Email / Designation / "
                          "Address), Correct value, Reason."),
        ("No account needed", "No registration or login required to submit a request."),
        ("Admin review",  "Submitted requests go to the Admin Panel for Approve / Reject."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 4.8 Universal Search
    add_heading(doc, "4.8  Universal Search  (/search)", level=2)
    items = [
        ("Scope",          "Searches employees, designations, organisations, departments, "
                           "and directory numbers simultaneously."),
        ("Autocomplete",   "Suggestions appear after 2 characters are typed; ranked by "
                           "relevance (exact match first, then prefix, then contains)."),
        ("Result grouping","Results are grouped by type: Employees, Organisations, "
                           "Directory Numbers."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    doc.add_paragraph()

    # ── 5. Admin Features ─────────────────────────────────────────────────────────
    add_heading(doc, "5.  Admin Features  (Login Required)", level=1)
    doc.add_paragraph(
        "The Admin Panel is accessible at /admin-login. Only accounts with "
        "role = 'admin' and status = 'active' can log in. All admin routes are "
        "protected by the admin_required decorator."
    )

    # 5.1 Dashboard
    add_heading(doc, "5.1  Admin Dashboard  (/admin/dashboard)", level=2)
    items = [
        ("Statistics cards",  "Total employees, total organisations, total departments, "
                              "and pending update requests."),
        ("Pending badge",     "A red badge on the Update Requests button shows the count "
                              "of unreviewed submissions."),
        ("Quick links",       "One-click navigation to Employees, Organisations, "
                              "Directory Numbers, Emergency Contacts, Update Requests."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.2 Manage Employees
    add_heading(doc, "5.2  Manage Employees  (/admin/employees)", level=2)
    items = [
        ("Search",          "Search across name, designation, email, mobile, organisation, "
                            "and department."),
        ("Table view",      "Name, Designation, Organisation, Mobile, Email, Utility Head "
                            "status per row."),
        ("Add employee",    "Create a new employee with all contact fields, organisation, "
                            "department, and utility head flag."),
        ("Edit employee",   "Update any field including the is_utility_head checkbox."),
        ("Delete employee", "Permanently remove an employee record."),
        ("Utility Head toggle", "Click the star button in the table to instantly mark or unmark "
                                "an employee as Utility Head – works from the admin page AND "
                                "from public directory pages."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.3 Manage Organisations
    add_heading(doc, "5.3  Manage Organisations  (/admin/organizations)", level=2)
    items = [
        ("Search",         "Search by name, region, or address."),
        ("Add / Edit",     "Organisation name, region, and physical address."),
        ("Delete",         "Remove an organisation (cascades to employees)."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.4 Directory Numbers
    add_heading(doc, "5.4  Directory Numbers  (/admin/directory-numbers)", level=2)
    items = [
        ("Search",        "Search by name, phone, organisation, or category."),
        ("Add / Edit",    "Name, phone number, email, organisation, category (e.g. Control Room)."),
        ("Delete",        "Remove a directory number entry."),
        ("Dashboard link","Back to dashboard button in the page header."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.5 Emergency Contacts
    add_heading(doc, "5.5  Emergency Contacts  (/admin/emergency-contacts)", level=2)
    items = [
        ("Search",        "Search by employee name, contact name, or relation."),
        ("Add / Edit",    "Link to employee, contact name, relation, phone."),
        ("Delete",        "Remove a contact record."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    # 5.6 Update Requests
    add_heading(doc, "5.6  Update Requests  (/admin/requests)", level=2)
    doc.add_paragraph(
        "Central inbox for all anonymous correction requests submitted by visitors."
    )
    items = [
        ("View all requests", "Shows submitter name, department, contact, field to update, "
                              "proposed value, reason, and submission date."),
        ("Approve",           "Automatically applies the requested change to the database "
                              "(employee field or directory number field)."),
        ("Reject",            "Marks the request as Rejected without changing the data."),
        ("Status tracking",   "Each request shows Pending / Approved / Rejected status."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    doc.add_paragraph()

    # ── 6. Security & Session Management ─────────────────────────────────────────
    add_heading(doc, "6.  Security & Session Management", level=1)
    items = [
        ("Public access",          "All five directory pages require no authentication. "
                                   "Any user on the intranet can view employee data."),
        ("Admin-only access",      "All /admin/* routes are protected by admin_required "
                                   "decorator which verifies session + role."),
        ("Password hashing",       "Admin passwords are stored as Werkzeug pbkdf2:sha256 "
                                   "hashes – never plain text."),
        ("Idle timeout – Admin",   "Admin session automatically logs out after 20 minutes "
                                   "of inactivity (PERMANENT_SESSION_LIFETIME = 20 min)."),
        ("Idle timeout – Public",  "An inactivity timer (20 min) redirects public users to "
                                   "the session-expired page with a notification message."),
        ("Unauthorised handler",   "Flask-Login is configured with a custom handler that "
                                   "always redirects to the admin login page instead of "
                                   "returning a 401 error."),
        ("No user registration",   "There is no self-registration. Admin accounts are "
                                   "created manually via the create_admin.py script."),
    ]
    for bold, desc in items:
        add_bullet(doc, desc, bold)

    doc.add_paragraph()

    # ── 7. Export Capabilities ────────────────────────────────────────────────────
    add_heading(doc, "7.  Export Capabilities", level=1)

    tbl3 = doc.add_table(rows=1, cols=3)
    tbl3.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl3.style     = "Table Grid"
    add_table_row(tbl3, ["Export Type", "Format", "Description"], header=True)

    exports = [
        ("Employee Directory (WRLDC)", "Excel (.xlsx) / Word (.docx)",
         "Current search/filter results; preserves navy header styling."),
        ("Telephone Directory (all)",  "Excel (.xlsx) / Word (.docx)",
         "All organisations or search results with employee and control room data."),
        ("Single Organisation Card",   "Word (.docx)",
         "One organisation with address, employees, and control room numbers."),
        ("Utility Heads List",         "Excel (.xlsx) / Word (.docx)",
         "All utility heads with organisation address column."),
    ]
    for row in exports:
        add_table_row(tbl3, list(row))

    doc.add_paragraph()

    # ── 8. URL Reference ─────────────────────────────────────────────────────────
    add_heading(doc, "8.  URL Reference", level=1)

    tbl4 = doc.add_table(rows=1, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl4.style     = "Table Grid"
    add_table_row(tbl4, ["URL", "Access", "Description"], header=True)

    urls = [
        ("/",                                  "Public", "Home page"),
        ("/directory",                         "Public", "Employee Directory (WRLDC only)"),
        ("/telephone-directory",               "Public", "Telephone Directory (all orgs)"),
        ("/utility-heads",                     "Public", "Utility Heads list"),
        ("/emergency-contacts",                "Public", "Emergency Contacts"),
        ("/employee/<id>",                     "Public", "Employee detail page"),
        ("/employee/request-update/<id>",      "Public", "Submit correction request"),
        ("/directory-number/<id>",             "Public", "Directory number detail"),
        ("/search",                            "Public", "Universal search results"),
        ("/export/employees",                  "Public", "Export Employee Directory"),
        ("/export/telephone",                  "Public", "Export Telephone Directory"),
        ("/export/org-word/<name>",            "Public", "Export single org as Word"),
        ("/export/utility-heads",              "Public", "Export Utility Heads"),
        ("/admin-login",                       "Public", "Admin login form"),
        ("/logout",                            "Admin",  "Log out admin"),
        ("/admin/dashboard",                   "Admin",  "Admin dashboard"),
        ("/admin/employees",                   "Admin",  "Manage employees"),
        ("/admin/employees/add",               "Admin",  "Add new employee"),
        ("/admin/employees/edit/<id>",         "Admin",  "Edit employee"),
        ("/admin/employees/delete/<id>",       "Admin",  "Delete employee"),
        ("/admin/employees/<id>/toggle-utility-head", "Admin", "Toggle utility head flag"),
        ("/admin/organizations",               "Admin",  "Manage organisations"),
        ("/admin/organizations/add",           "Admin",  "Add organisation"),
        ("/admin/organizations/edit/<id>",     "Admin",  "Edit organisation"),
        ("/admin/organizations/delete/<id>",   "Admin",  "Delete organisation"),
        ("/admin/directory-numbers",           "Admin",  "Manage directory numbers"),
        ("/admin/directory-numbers/add",       "Admin",  "Add directory number"),
        ("/admin/directory-numbers/edit/<id>", "Admin",  "Edit directory number"),
        ("/admin/directory-numbers/delete/<id>","Admin", "Delete directory number"),
        ("/admin/emergency-contacts",          "Admin",  "Manage emergency contacts"),
        ("/admin/emergency-contacts/add",      "Admin",  "Add emergency contact"),
        ("/admin/emergency-contacts/edit/<id>","Admin",  "Edit emergency contact"),
        ("/admin/emergency-contacts/delete/<id>","Admin","Delete emergency contact"),
        ("/admin/requests",                    "Admin",  "View update requests"),
        ("/admin/requests/approve/<id>",       "Admin",  "Approve correction request"),
        ("/admin/requests/reject/<id>",        "Admin",  "Reject correction request"),
    ]
    for row in urls:
        add_table_row(tbl4, list(row))

    doc.add_paragraph()

    # ── 9. Project Files ──────────────────────────────────────────────────────────
    add_heading(doc, "9.  Project File Structure", level=1)

    tbl5 = doc.add_table(rows=1, cols=2)
    tbl5.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl5.style     = "Table Grid"
    add_table_row(tbl5, ["File / Folder", "Purpose"], header=True)

    files = [
        ("app.py",                    "Flask application factory; registers blueprints, login manager, DB."),
        ("config.py",                 "Config class: DB URI, secret key, session lifetime."),
        ("requirements.txt",          "Python package dependencies."),
        ("models/__init__.py",        "SQLAlchemy db instance."),
        ("models/organization.py",    "Organization model."),
        ("models/department.py",      "Department model."),
        ("models/employee.py",        "Employee model (includes is_utility_head)."),
        ("models/directory_number.py","DirectoryNumber model (control rooms, etc.)."),
        ("models/emergency_contact.py","EmergencyContact model."),
        ("models/update_request.py",  "UpdateRequest model."),
        ("models/user.py",            "Admin User model."),
        ("routes/auth_routes.py",     "Login and logout routes."),
        ("routes/user_routes.py",     "All public-facing routes + export logic."),
        ("routes/admin_routes.py",    "All admin CRUD routes."),
        ("templates/base.html",       "Base layout: navbar, Bootstrap, idle-timer JS."),
        ("templates/index.html",      "Home page template."),
        ("templates/directory.html",  "Employee Directory template."),
        ("templates/telephone_directory.html", "Telephone Directory template."),
        ("templates/utility_heads.html",       "Utility Heads template."),
        ("templates/emergency_contacts.html",  "Emergency Contacts template."),
        ("templates/admin_dashboard.html",     "Admin Dashboard template."),
        ("templates/manage_employees.html",    "Admin – employee list + utility head toggle."),
        ("templates/manage_organizations.html","Admin – organisation list."),
        ("templates/directory_numbers.html",   "Admin – directory numbers list."),
        ("templates/manage_emergency_contacts.html", "Admin – emergency contacts list."),
        ("templates/manage_requests.html",     "Admin – update requests inbox."),
        ("static/css/style.css",      "Custom styles: navy colour scheme, org card, badges."),
        ("create_admin.py",           "One-time script to create the admin user in the DB."),
        ("import_employee.py",        "Script to import employees from the Word document."),
        ("import_directory_numbers.py","Script to import control room numbers."),
        ("update_addresses_and_heads.py","Script to extract org addresses and mark utility heads."),
        ("fix_data_quality.py",       "Script to correct org names, spacing, and Unicode errors."),
        ("database/postgres_schema.sql","Full PostgreSQL DDL for all 7 tables."),
    ]
    for row in files:
        add_table_row(tbl5, list(row))

    doc.add_paragraph()

    # ── Footer note ───────────────────────────────────────────────────────────────
    add_heading(doc, "10.  Admin Login Credentials", level=1)
    doc.add_paragraph(
        "The admin account was created using create_admin.py. Default credentials:"
    )
    add_bullet(doc, "admin@gmail.com",   "Email:")
    add_bullet(doc, "admin123",           "Password:")
    doc.add_paragraph(
        "It is strongly recommended to change the password after the first login "
        "by updating the hashed password in the database."
    )

    out = "Project_Summary_Telephone_Directory.docx"
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    build()
